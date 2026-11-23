import os
import re
import glob
import warnings
import numpy as np
import winreg

import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ==================== Configuration ====================

def get_desktop_path():
    """Read desktop path from registry"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        )
        desktop, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        desktop = os.path.expandvars(desktop)
        return desktop
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


PO_SOURCE_COL = "Purchasing Document (Aging Accrual Table)"

LEG_COLUMN_ALIASES = {
    "last": ["PO# (Last Leg)", "PO# (last leg)"],
    "2nd":  ["PO# (2nd Leg)"],
}
IR_KEYWORDS = ["Import Customs Declaration #", "Import Customs Declaration"]


# ==================== Utility Functions ====================

def get_base_path():
    base = os.path.join(get_desktop_path(), "TCD AGING REPORT")
    if not os.path.exists(base):
        base = os.path.join(get_desktop_path(), "TCD AGING")
    return base


def find_main_report(base_path):
    files = glob.glob(os.path.join(base_path, "TCD Aging Accrual Report*.xlsx"))
    files = [f for f in files if "TEMPLATE" not in f.upper()]
    return files[0] if files else None


def to_iso_date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        v = v.strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(v, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        return v
    return v


def extract_year(v):
    if hasattr(v, "year"):
        return v.year
    if isinstance(v, str):
        m = re.search(r"(\d{4})", v)
        return int(m.group(1)) if m else None
    return None


def norm_col(c):
    return re.sub(r"\s+", " ", str(c)).strip().lower()


def find_col(columns, *keywords):
    normed = {norm_col(c): c for c in columns}
    for kw in keywords:
        k = norm_col(kw)
        if k in normed:
            return normed[k]
    for k, orig in normed.items():
        for kw in keywords:
            if norm_col(kw) in k:
                return orig
    return None


def po_key(v):
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    s = str(v).strip()
    s = re.sub(r"\s+", "", s)
    try:
        f = float(s)
        if np.isnan(f):
            return ""
        s = str(int(f)) if f.is_integer() else f"{f:.0f}"
    except (ValueError, OverflowError):
        pass
    s = re.sub(r"\.0+$", "", s)
    return s.lower()


def is_valid_ir(v):
    if v is None:
        return False
    if isinstance(v, float) and np.isnan(v):
        return False
    s = po_key(v)
    if not s:
        return False
    if s in ("n/a", "tba", "nan"):
        return False
    return True


def safe_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    s = str(v)
    if s in ("None", "nan", ""):
        return ""
    # ==== New: uniform formatting ====
    s = s.replace(",", "")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


def ssr_year_from_filename(fname):
    m = re.search(r'[ _](\d{2})\d{4}', os.path.basename(fname))
    if m:
        yy = int(m.group(1))
        return 2000 + yy if yy <= 50 else 1900 + yy
    return None


# ==================== Load all SSR into memory ====================

def load_all_ssr(base_path):
    """Load all SSR files at once, build PO -> (year, IR) mapping"""
    ssr_dir = os.path.join(base_path, "SSR")
    if not os.path.exists(ssr_dir):
        print("[ERROR] Cannot find SSR folder:", ssr_dir)
        return {}, []

    all_files = glob.glob(os.path.join(ssr_dir, "*.xlsx"))
    print(f"  -> Found {len(all_files)} SSR files")

    ssr_by_year = {}
    file_info = []

    for f in all_files:
        year = ssr_year_from_filename(f)
        if year is None:
            print(f"  [SKIP] Year cannot be determined: {os.path.basename(f)}")
            continue

        try:
            dfs = pd.read_excel(f, sheet_name="SSR", dtype=str)
        except Exception as e:
            print(f"  [WARN] Failed to read {os.path.basename(f)}: {e}")
            continue

        ir_col = find_col(dfs.columns, *IR_KEYWORDS)
        if not ir_col:
            print(f"  [WARN] {os.path.basename(f)}: IR column not found")
            continue

        leg_cols = []
        norm_map = {norm_col(c): c for c in dfs.columns}
        for leg_name, aliases in LEG_COLUMN_ALIASES.items():
            for alias in aliases:
                if norm_col(alias) in norm_map and norm_map[norm_col(alias)] not in leg_cols:
                    leg_cols.append(norm_map[norm_col(alias)])
                    break
        if not leg_cols:
            print(f"  [WARN] {os.path.basename(f)}: No PO leg columns found")
            continue

        # Filter rows with valid IR
        valid_mask = dfs[ir_col].apply(is_valid_ir)
        dfs_valid = dfs[valid_mask].copy()

        # Collect PO -> IR mappings from all leg columns
        po_to_ir = {}
        for pc in leg_cols:
            for _, row in dfs_valid.iterrows():
                k = po_key(row[pc])
                if k and k not in po_to_ir:
                    po_to_ir[k] = row[ir_col]

        ssr_by_year.setdefault(year, {})
        added = 0
        for k, v in po_to_ir.items():
            if k not in ssr_by_year[year]:
                ssr_by_year[year][k] = v
                added += 1

        file_info.append((os.path.basename(f), year, leg_cols, added))
        print(f"  -> {os.path.basename(f)} ({year}): leg cols={leg_cols}, added {added} entries")

    return ssr_by_year, file_info


# ==================== Step 1 ====================

def step1_filter(df_ph):
    print("[Step 1] Filtering PH sheet...")
    mask_ao = df_ph["Action Owner"].astype(str).str.upper().str.contains("4PL", na=False)
    mask_rm = (
        df_ph["Remarks and Next Step"].astype(str).str.contains("No Broker", na=False) |
        df_ph["Remarks and Next Step"].astype(str).str.contains("Amount mismatch", na=False)
    )
    mask = mask_ao & mask_rm
    print(f"  -> Filtered {mask.sum()} rows")
    return mask


# ==================== Step 2 (in-memory) ====================

def step2_fill_ir(df_ph_filtered, ssr_by_year):
    """Fill Import Declaration from IR doc directly on the DataFrame"""
    print("[Step 2] Filling Import Declaration from IR doc...")

    years_desc = sorted(ssr_by_year.keys(), reverse=True)
    print(f"  -> SSR covers years: {sorted(ssr_by_year.keys())}")

    # Find rows that need filling
    need_mask = (
        df_ph_filtered[PO_SOURCE_COL].notna() &
        (~df_ph_filtered["Import Declaration from IR doc"].apply(is_valid_ir))
    )
    need_indices = df_ph_filtered.index[need_mask]
    print(f"  -> Need to fill {len(need_indices)} PO(s)")

    if len(need_indices) == 0:
        return

    # Batch match
    filled_count = 0
    for idx in need_indices:
        po_val = df_ph_filtered.at[idx, PO_SOURCE_COL]
        date_val = df_ph_filtered.at[idx, "Date"]
        yr = extract_year(date_val)

        k = po_key(po_val)
        if not k or yr is None:
            continue

        # Prefer same year
        val = ssr_by_year.get(yr, {}).get(k)
        if not val:
            # Fallback descending order
            for y in years_desc:
                if y == yr:
                    continue
                if k in ssr_by_year.get(y, {}):
                    val = ssr_by_year[y][k]
                    break

        if val:
            df_ph_filtered.at[idx, "Import Declaration from IR doc"] = val
            filled_count += 1

    print(f"  -> Successfully matched {filled_count} PO(s)")


# ==================== Step 3 (vectorized) ====================

PH_BT_HEADERS = [
    "Purchasing Document Number",
    "Company Code",
    "Material",
    "Import Declaration Number",
    "Duty Coding",
    "TMSLoad#",
    "Import Date",
    "Qty",
    "UOM",
    "Duty Rate",
    "Duty Currency",
    "Basic Duty Amount",
    "Other Taxes Amount",
    "VAT Amount",
    "Country of Origin",
    "Dutiable Amount",
    "HTS Code",
    "Aging bucket",
    "Remarks and Next Step",
    "Action Owner",
    "MUSD",
]

PH_SOURCE_OVERRIDE = {
    "Purchasing Document Number": "Purchasing Document (Aging Accrual Table)",
    "Company Code":              "Company Code",
    "Material":                   "Material (Aging Accrual Table)",
    "Import Declaration Number":  "Import Declaration from IR doc",
    "Duty Coding":               "Duty Coding",
    "TMSLoad#":                  "TMS LOAD ID",
    "Import Date":                "Date",
    "Qty":                        "Quantity",
    "UOM":                        "Base Unit of Measure",
    "Duty Rate":                 "Duty Rate",
    "Duty Currency":             "Duty Currency",
    "Basic Duty Amount":          "Basic Duty Amount (Broker Table)",
    "Other Taxes Amount":         "Other Taxes Amount (Broker Table)",
    "VAT Amount":                 "VAT Amount (Broker Table)",
    "Country of Origin":          "Country of origin",
    "Dutiable Amount":           "Dutiable Amount",
    "HTS Code":                  "HTS Code",
    "Aging bucket":              "Aging bucket",
    "Remarks and Next Step":      "Remarks and Next Step",
    "Action Owner":               "Action Owner",
    "MUSD":                      "MUSD",
}


def step3_build_bt(df_ph_filtered):
    """Build PH-BT DataFrame vectorized"""
    print("[Step 3] Building PH-BT...")

    bt_data = pd.DataFrame()
    for bt_hdr in PH_BT_HEADERS:
        src_col = PH_SOURCE_OVERRIDE.get(bt_hdr, bt_hdr)
        if src_col in df_ph_filtered.columns:
            if bt_hdr == "Import Date":
                bt_data[bt_hdr] = df_ph_filtered[src_col].apply(to_iso_date)
            else:
                bt_data[bt_hdr] = df_ph_filtered[src_col]
        else:
            bt_data[bt_hdr] = None

    # References for Amount mismatch = A + C + D + J + L + M + N + P + Q
    cols_amount_mismatch = [
        "Purchasing Document Number",  # A
        "Material",                     # C
        "Import Declaration Number",    # D
        "Duty Rate",                    # J
        "Basic Duty Amount",            # L
        "Other Taxes Amount",           # M
        "VAT Amount",                   # N
        "Dutiable Amount",              # P
        "HTS Code",                     # Q
    ]

    bt_data["References for Amount mismatch"] = bt_data.apply(
        lambda r: "".join(safe_str(r[hdr]) for hdr in cols_amount_mismatch), axis=1
    )

    bt_data["COMPARE TO SAP"] = ""

    # References for NO BT = A + C + D
    cols_no_bt = [
        "Purchasing Document Number",  # A
        "Material",                     # C
        "Import Declaration Number",    # D
    ]

    bt_data["References for NO BT"] = bt_data.apply(
        lambda r: "".join(safe_str(r[hdr]) for hdr in cols_no_bt), axis=1
    )

    bt_data["COMPARE TO SAP2"] = ""
    bt_data["4PL REMARKS"] = ""
    bt_data["Next Action"] = ""
    bt_data["Assigned to"] = ""

    print(f"  -> PH-BT has {len(bt_data)} rows")
    return bt_data


# ==================== Step 3.5 ====================

BUCKETS = {
    "A. 0-30":     ["A"],
    "B. 31-60":    ["B"],
    "D. 91-120":   ["D"],
    "G. OVER 180": ["G"],
}


def step35_buckets(bt_data):
    """Split into buckets, return dict of DataFrames per sheet"""
    print("[Step 3.5] Splitting into buckets A/B/D/G...")

    def shorten(x):
        x = str(x)
        if "Amount mismatch" in x:
            return "Amount mismatch"
        if "No Broker" in x:
            return "No Broker"
        return x

    sheets = {}
    for sheet_name, keys in BUCKETS.items():
        mask = bt_data["Aging bucket"].astype(str).apply(lambda x: any(k in x for k in keys))
        subset = bt_data.loc[mask, ["Import Declaration Number", "Import Date", "Remarks and Next Step"]].copy()
        before = len(subset)
        subset = subset.drop_duplicates(subset=["Import Declaration Number"], keep="first").reset_index(drop=True)
        if before != len(subset):
            print(f"  -> {sheet_name}: dedup {before} -> {len(subset)} rows")
        subset["Remarks and Next Step"] = subset["Remarks and Next Step"].apply(shorten)
        sheets[sheet_name] = subset
        print(f"  -> {sheet_name}: {len(subset)} rows")

    return sheets


# ==================== Step 4 ====================

def step4_sap_check(bt_data):
    """SAP CHECK dedup, return DataFrame"""
    print("[Step 4] SAP CHECK dedup...")
    unique_irs = bt_data["Import Declaration Number"].dropna().unique()
    df_sc = pd.DataFrame({"Import Declaration Number": unique_irs})
    print(f"  -> SAP CHECK has {len(df_sc)} unique values")
    return df_sc


# ==================== Write all sheets at once ====================

def write_all_sheets(template, df_ph, bt_data, bucket_sheets, sap_check_df):
    """Write all sheets at once to reduce Excel I/O"""
    print("[Write] Writing all sheets at once...")

    with pd.ExcelWriter(template, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_ph.to_excel(writer, sheet_name="PH", index=False)
        bt_data.to_excel(writer, sheet_name="PH-BT", index=False)
        for sheet_name, df in bucket_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        sap_check_df.to_excel(writer, sheet_name="SAP CHECK", index=False)

    print("  -> All sheets written successfully")


# ==================== Main Flow ====================

def phase1():
    base = get_base_path()
    main_report = find_main_report(base)
    template = os.path.join(base, "TCD Aging Accrual Report TEMPLATE.xlsx")

    print("=" * 55)
    print("Phase 1: Step 1 -> 2 -> 3 -> 3.5 -> 4")
    print(f"Working directory: {base}")
    print("=" * 55)

    if not main_report:
        print("[ERROR] Main report not found")
        return
    if not os.path.exists(template):
        print("[ERROR] Template not found")
        return

    print(f"Main report: {os.path.basename(main_report)}")

    # ====== Load SSR (once) ======
    print("[Load] Loading all SSR files into memory...")
    ssr_by_year, _ = load_all_ssr(base)
    print(f"  -> SSR covers years: {sorted(ssr_by_year.keys())}")

    # ====== Read main report ======
    print("[Read] Reading main report PH sheet...")
    df_ph = pd.read_excel(main_report, sheet_name="PH")
    print(f"  -> Total {len(df_ph)} rows")

    # ====== Step 1 ======
    filtered_mask = step1_filter(df_ph)
    df_ph_filtered = df_ph[filtered_mask].copy()

    # ====== Step 2 (in-memory) ======
    step2_fill_ir(df_ph_filtered, ssr_by_year)

    # ====== Step 3 (vectorized) ======
    bt_data = step3_build_bt(df_ph_filtered)

    # ====== Step 3.5 (in-memory) ======
    bucket_sheets = step35_buckets(bt_data)

    # ====== Step 4 (in-memory) ======
    sap_check_df = step4_sap_check(bt_data)

    # ====== Write all at once ======
    write_all_sheets(template, df_ph_filtered, bt_data, bucket_sheets, sap_check_df)

    print("\n[Phase 1 Completed]")
    print("Breakpoint: Please manually send the SAP CHECK email; run Phase 2 after receiving the reply")


if __name__ == "__main__":
    phase1()