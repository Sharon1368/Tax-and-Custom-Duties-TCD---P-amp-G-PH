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


DESKTOP = get_desktop_path()


def get_base_path():
    base = os.path.join(DESKTOP, "TCD AGING REPORT")
    if not os.path.exists(base):
        base = os.path.join(DESKTOP, "TCD AGING")
    return base


# ==================== Utility Functions ====================

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
        s = to_iso_date(v) if not isinstance(v, datetime) else v.strftime("%Y-%m-%d")
        m = re.search(r"(\d{4})", s)
        return int(m.group(1)) if m else None
    return None


def norm_col(c):
    # Replace all non-standard whitespace characters with normal space
    c = re.sub(r"[\t\n\r\xa0\u3000]", " ", str(c))
    return re.sub(r"\s+", " ", c).strip().lower()


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


def safe_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    s = str(v)
    if s in ("None", "nan", ""):
        return ""
    return s


BUCKET_SHEETS = ["A. 0-30", "B. 31-60", "D. 91-120", "G. OVER 180"]


# ==================== Step 5: EXPORT → SAP-RAW ====================
def step5_export_to_sapraw(template):
    print("[Step 5] EXPORT → SAP-RAW...")

    base = get_base_path()
    export_files = glob.glob(os.path.join(base, "EXPORT*.xlsx"))
    if not export_files:
        print("[ERROR] Cannot find EXPORT file")
        return False

    export_file = export_files[0]
    print(f"  → Found: {os.path.basename(export_file)}")

    df_export = pd.read_excel(export_file, header=4, dtype={"Material": str})
    df_export = df_export.dropna(how='all').reset_index(drop=True)

    if "Material" in df_export.columns:
        df_export["Material"] = df_export["Material"].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() != "" else x
        )

    print(f"  → EXPORT total {len(df_export)} rows, {len(df_export.columns)} columns")

    end_col = min(39, len(df_export.columns))
    df_sapraw = df_export.iloc[:, 2:end_col].copy()

    with pd.ExcelWriter(template, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_sapraw.to_excel(writer, sheet_name="SAP-RAW", index=False)

    return True


# ==================== Step 6: SAP-RAW → SAP-BT ====================
def step6_sap_raw_to_sap_bt(template):
    print("[Step 6] SAP-RAW → SAP-BT...")

    df_raw = pd.read_excel(template, sheet_name="SAP-RAW", header=0)
    print(f"  → SAP-RAW total {len(df_raw)} rows")

    sapbt_columns = [
        "Purchasing Document Number",
        "Company Code",
        "Material",
        "Import Declaration Number",
        "Duty Coding",
        "TMS LOAD ID",
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
    ]

    col_map = {
        "Purchasing Doc.": "Purchasing Document Number",
        "Company Code": "Company Code",
        "Material": "Material",
        "Import Declaration": "Import Declaration Number",
        "Duty Coding": "Duty Coding",
        "TMS LOAD ID": "TMS LOAD ID",
        "Date": "Import Date",
        "Quantity": "Qty",
        "Base Unit": "UOM",
        "Duty Rate": "Duty Rate",
        "Duty Currency": "Duty Currency",
        "Basic Duty Amount": "Basic Duty Amount",
        "Other Taxes Amount": "Other Taxes Amount",
        "VAT Amount": "VAT Amount",
        "Country of origin": "Country of Origin",
        "Dutiable Amount": "Dutiable Amount",
        "HTS Code": "HTS Code",
    }

    df_sapbt = pd.DataFrame(columns=sapbt_columns)

    for src_col, tgt_col in col_map.items():
        if src_col in df_raw.columns:
            df_sapbt[tgt_col] = df_raw[src_col]
        else:
            print(f"  [WARN] SAP-RAW missing column: {src_col}")
            df_sapbt[tgt_col] = None

    print(f"  → SAP-BT total {len(df_sapbt)} rows, {len(df_sapbt.columns)} columns")

    def excel_concat(*args, remove_commas=False):
        parts = []
        for v in args:
            if v is None or (isinstance(v, float) and pd.isna(v)):
                parts.append("")
            else:
                s = str(v)
                s = s.replace(",", "")
                if "." in s:
                    s = s.rstrip("0").rstrip(".")
                if remove_commas:
                    s = s.replace(",", "")
                parts.append(s)
        return "".join(parts)

    print("  → Computing R_ref and S_ref...")

    df_sapbt["R_ref"] = df_sapbt.apply(
        lambda r: excel_concat(
            r["Purchasing Document Number"],
            r["Material"],
            r["Import Declaration Number"],
            r["Duty Rate"],
            r["Basic Duty Amount"],
            r["Other Taxes Amount"],
            r["VAT Amount"],
            r["Dutiable Amount"],
            r["HTS Code"],
            remove_commas=True
        ), axis=1
    )

    df_sapbt["S_ref"] = df_sapbt.apply(
        lambda r: excel_concat(
            r["Purchasing Document Number"],
            r["Material"],
            r["Import Declaration Number"],
        ), axis=1
    )

    with pd.ExcelWriter(template, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_sapbt.to_excel(writer, sheet_name="SAP-BT", index=False)

    print(f"  → SAP-BT written ({len(df_sapbt)} rows, with static R_ref / S_ref values)")


# ==================== Step 7: Amount Mismatch ====================
def step7_amount_mismatch(template):
    print("[Step 7] Amount mismatch processing...")

    df_sapbt = pd.read_excel(template, sheet_name="SAP-BT", header=0)

    sap_r_set = set()
    sap_s_set = set()
    for _, row in df_sapbt.iterrows():
        r_val = str(row["R_ref"]).strip() if pd.notna(row["R_ref"]) else ""
        s_val = str(row["S_ref"]).strip() if pd.notna(row["S_ref"]) else ""
        if r_val:
            sap_r_set.add(r_val)
        if s_val:
            sap_s_set.add(s_val)

    print(f"  SAP-BT R set size: {len(sap_r_set)}, S set size: {len(sap_s_set)}")

    df_bt = pd.read_excel(template, sheet_name="PH-BT", header=0)
    bt_cols = list(df_bt.columns)

    bt_ir_col = find_col(bt_cols, "Import Declaration Number")
    bt_ref1_col = find_col(bt_cols, "References for Amount mismatch")
    bt_ref2_col = find_col(bt_cols, "References for NO BT")

    if bt_ir_col is None:
        print("[ERROR] PH-BT missing Import Declaration Number column")
        return

    from openpyxl.utils import get_column_letter

    wb = load_workbook(template)
    ws_bt = wb["PH-BT"]

    bt_header = [cell.value for cell in ws_bt[1]]

    v_col_idx = bt_header.index("References for Amount mismatch") + 1
    v_col_letter = get_column_letter(v_col_idx)

    x_col_idx = bt_header.index("References for NO BT") + 1
    x_col_letter = get_column_letter(x_col_idx)

    if "COMPARE TO SAP" in bt_header:
        cmp1_col_idx = bt_header.index("COMPARE TO SAP") + 1
    else:
        cmp1_col_idx = len(bt_header) + 1
        ws_bt.cell(row=1, column=cmp1_col_idx, value="COMPARE TO SAP")
    cmp1_letter = get_column_letter(cmp1_col_idx)

    if "COMPARE TO SAP2" in bt_header:
        cmp2_col_idx = bt_header.index("COMPARE TO SAP2") + 1
    else:
        cmp2_col_idx = len(bt_header) + 1
        ws_bt.cell(row=1, column=cmp2_col_idx, value="COMPARE TO SAP2")
    cmp2_letter = get_column_letter(cmp2_col_idx)

    for row_num in range(2, ws_bt.max_row + 1):
        ws_bt[f"{cmp1_letter}{row_num}"] = (
            f'=ISNUMBER(MATCH({v_col_letter}{row_num},\'SAP-BT\'!R:R,0))'
        )
        ws_bt[f"{cmp2_letter}{row_num}"] = (
            f'=ISNUMBER(MATCH({x_col_letter}{row_num},\'SAP-BT\'!S:S,0))'
        )

    wb.save(template)
    wb.close()

    print("  → Triggering COMPARE TO SAP formula calculation...")
    try:
        import win32com.client as win32
        import pythoncom

        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        abs_path = os.path.abspath(template)
        wb = excel.Workbooks.Open(abs_path)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close()
        # Release COM objects only, do not close other Excel instances
        excel.DisplayAlerts = True
        del wb
        del excel
        pythoncom.CoUninitialize()
        print("  → Formula calculation completed")
    except Exception as e:
        print(f"  [WARN] Formula calculation failed: {e}")
        print("  → Please manually open and save the file once")

    df_bt_result = pd.read_excel(template, sheet_name="PH-BT", header=0)

    result_cols = list(df_bt_result.columns)
    cmp1_result_col = find_col(result_cols, "COMPARE TO SAP")
    remarks_col = find_col(result_cols, "4PL REMARKS")

    if cmp1_result_col is None:
        print("[ERROR] Cannot find COMPARE TO SAP column")
        return

    bt_by_ir = {}
    for idx, row in df_bt_result.iterrows():
        ir_val = str(row[bt_ir_col]).strip() if pd.notna(row[bt_ir_col]) else ""
        if ir_val:
            bt_by_ir[ir_val] = idx

    for sheet_name in BUCKET_SHEETS:
        try:
            df_bucket = pd.read_excel(template, sheet_name=sheet_name, header=0)
        except Exception:
            print(f"  [WARN] Cannot read {sheet_name}")
            continue

        bk_cols = list(df_bucket.columns)
        bk_ir_col = find_col(bk_cols, "Import Declaration Number")
        bk_rm_col = find_col(bk_cols, "Remarks and Next Step")
        bk_status_col = find_col(bk_cols, "4PL STATUS")

        if bk_ir_col is None:
            print(f"  [WARN] {sheet_name}: Cannot find Import Declaration Number")
            continue

        if bk_status_col is None:
            df_bucket["4PL STATUS"] = ""
            bk_status_col = "4PL STATUS"

        done_count = 0
        mismatch_count = 0

        for idx, row in df_bucket.iterrows():
            ir_val = str(row[bk_ir_col]).strip() if pd.notna(row[bk_ir_col]) else ""
            rm_val = str(row[bk_rm_col]).strip() if bk_rm_col is not None and pd.notna(row[bk_rm_col]) else ""

            if "amount mismatch" not in rm_val.lower():
                continue

            if ir_val in bt_by_ir:
                bt_idx = bt_by_ir[ir_val]
                compare_val = df_bt_result.at[bt_idx, cmp1_result_col]

                if compare_val == True:
                    df_bucket.at[idx, bk_status_col] = "Done"
                    if remarks_col is not None:
                        df_bt_result[remarks_col] = df_bt_result[remarks_col].astype(object)
                        df_bt_result.at[bt_idx, remarks_col] = "Amount Match"
                    done_count += 1
                else:
                    if remarks_col is not None:
                        df_bt_result[remarks_col] = df_bt_result[remarks_col].astype(object)
                        df_bt_result.at[bt_idx, remarks_col] = "Amount mismatch"
                    mismatch_count += 1

        with pd.ExcelWriter(template, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_bucket.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"  → {sheet_name}: Done={done_count}, Amount mismatch={mismatch_count}")

    with pd.ExcelWriter(template, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_bt_result.to_excel(writer, sheet_name="PH-BT", index=False)

    print("  → Step 7 completed")


def step8_no_broker(template, bt_template_file):
    print("[Step 8] No Broker lookup via Customs Download...")

    base = get_base_path()
    customs_dir = os.path.join(base, "Customs Download Consolidated")

    if not os.path.exists(customs_dir):
        print(f"[ERROR] Cannot find Customs Download folder: {customs_dir}")
        return

    customs_files = sorted(glob.glob(os.path.join(customs_dir, "Customs Download*.xlsx")))
    if not customs_files:
        print("[ERROR] Cannot find Customs Download files")
        return

    print(f"  → Found {len(customs_files)} Customs Download files")

    # ---- Build index for each file ----
    file_indexes = []
    for cf in customs_files:
        df_cd = pd.read_excel(cf, header=0)
        entry_col = find_col(df_cd.columns, "Entry No", "Entry Number", "Entry #")
        pay_col = find_col(df_cd.columns, "Payment Date", "Payment date")
        if not entry_col or not pay_col:
            print(f"  [WARN] {os.path.basename(cf)}: Missing columns, skipping")
            file_indexes.append(None)
            continue

        index = {}
        for idx, row in df_cd.iterrows():
            e = safe_str(row[entry_col])
            y = extract_year(row[pay_col])
            if e and y:
                key = (e, y)
                if key not in index:
                    index[key] = []
                index[key].append(idx)
        file_indexes.append((df_cd, entry_col, pay_col, index))
        print(f"  → {os.path.basename(cf)}: {len(df_cd)} rows, {len(index)} unique Entry+Year")

    # ---- Collect No Broker rows from bucket sheets ----
    needed = []
    for sheet_name in BUCKET_SHEETS:
        try:
            df_b = pd.read_excel(template, sheet_name=sheet_name, header=0)
        except Exception:
            continue

        ir_col = find_col(df_b.columns, "Import Declaration Number")
        rm_col = find_col(df_b.columns, "Remarks and Next Step")
        dt_col = find_col(df_b.columns, "Import Date")
        if not ir_col:
            continue

        for _, row in df_b.iterrows():
            ir = safe_str(row[ir_col])
            rm = safe_str(row[rm_col]) if rm_col else ""
            if "no broker" not in rm.lower():
                continue
            yr = extract_year(row[dt_col]) if dt_col else None
            if ir and yr:
                needed.append((ir, yr))

    if not needed:
        print("  → No No Broker data to query")
        return

    print(f"  → Need to query {len(needed)} No Broker records")

    # ---- Query each file ----
    matched_rows = []
    remaining = list(needed)

    for fi_idx, result in enumerate(file_indexes):
        if not remaining:
            break
        if result is None:
            continue

        df_cd, entry_col, pay_col, index = result
        found_keys = []

        for key in remaining:
            if key in index:
                for ridx in index[key]:
                    matched_rows.append(df_cd.iloc[ridx])
                found_keys.append(key)

        for k in found_keys:
            remaining.remove(k)

        print(f"  → File {fi_idx + 1}: found {sum(len(index[k]) for k in found_keys)} rows, {len(remaining)} still unmatched")

    if not matched_rows:
        print("  → No matching Customs data found")
    else:
        print(f"  → Total {len(matched_rows)} rows found")

        # ---- Write to BT TEMPLATE ACCRUAL.xlsx ----
        if not os.path.exists(bt_template_file):
            print(f"[ERROR] Cannot find BT TEMPLATE ACCRUAL: {bt_template_file}")
            return

        wb = load_workbook(bt_template_file)

        target_sheet = "PROCTER & GAMBLE PHILIPPINE IN"
        if target_sheet in wb.sheetnames:
            del wb[target_sheet]
        ws = wb.create_sheet(target_sheet)

        df_first = file_indexes[0]
        if df_first is not None:
            df_cd = df_first[0]
            cd_col_count = df_cd.shape[1]
            write_cols = min(59, cd_col_count)
            for col in range(1, write_cols + 1):
                ws.cell(row=1, column=col, value=str(df_cd.columns[col - 1]))

        for offset, src_row in enumerate(matched_rows):
            r = offset + 2
            for col in range(1, write_cols + 1):
                val = src_row.iloc[col - 1]
                if pd.isna(val):
                    val = None
                ws.cell(row=r, column=col, value=val)

        wb.save(bt_template_file)
        wb.close()

        print(f"  → PROCTER & GAMBLE PHILIPPINE IN written with {len(matched_rows)} rows")

    # ---- Build found_set ----
    found_set = set()
    for fi_idx, result in enumerate(file_indexes):
        if result is None:
            continue
        df_cd, entry_col, pay_col, index = result
        for key in index:
            found_set.add(key)

    # ---- Mark unmatched No Broker rows ----
    for sheet_name in BUCKET_SHEETS:
        try:
            df_b = pd.read_excel(template, sheet_name=sheet_name, header=0)
        except Exception:
            continue

        ir_col = find_col(df_b.columns, "Import Declaration Number")
        rm_col = find_col(df_b.columns, "Remarks and Next Step")
        dt_col = find_col(df_b.columns, "Import Date")

        if not ir_col or not rm_col:
            continue

        has_no_broker = any(
            "no broker" in safe_str(row[rm_col]).lower()
            for _, row in df_b.iterrows()
        )
        if not has_no_broker:
            continue

        wb_local = load_workbook(template)
        ws = wb_local[sheet_name]
        ws_header = [cell.value for cell in ws[1]]

        def find_col_idx_safe(header, keyword):
            norm_kw = keyword.strip().lower().replace(" ", "")
            for i, h in enumerate(header, 1):
                if h is None:
                    continue
                hn = str(h).strip().lower().replace(" ", "")
                if hn == norm_kw or norm_kw in hn:
                    return i
            return None

        ir_idx = find_col_idx_safe(ws_header, "importdeclarationnumber")
        rm_idx = find_col_idx_safe(ws_header, "remarksandnextstep")
        dt_idx = find_col_idx_safe(ws_header, "importdate")
        st_idx = find_col_idx_safe(ws_header, "4plstatus")

        next_col = len(ws_header) + 1

        if ir_idx is None:
            ir_idx = next_col
            ws.cell(row=1, column=ir_idx, value="Import Declaration Number")
            print(f"  → {sheet_name}: Auto-added Import Declaration Number column")
            next_col += 1

        if rm_idx is None:
            rm_idx = next_col
            ws.cell(row=1, column=rm_idx, value="Remarks and Next Step")
            print(f"  → {sheet_name}: Auto-added Remarks and Next Step column")
            next_col += 1

        if dt_idx is None:
            dt_idx = next_col
            ws.cell(row=1, column=dt_idx, value="Import Date")
            print(f"  → {sheet_name}: Auto-added Import Date column")
            next_col += 1

        if st_idx is None:
            st_idx = next_col
            ws.cell(row=1, column=st_idx, value="4PL STATUS")
            print(f"  → {sheet_name}: Auto-added 4PL STATUS column")
            next_col += 1

        remarks_idx = next_col
        oss_remarks_idx = next_col + 1
        oss_status_idx = next_col + 2

        ws.cell(row=1, column=remarks_idx, value="4PL Remarks")
        ws.cell(row=1, column=oss_remarks_idx, value="OSS REMARKS")
        ws.cell(row=1, column=oss_status_idx, value="OSS STATUS")

        not_found_count = 0
        for r in range(2, ws.max_row + 1):
            rm_val = safe_str(ws.cell(row=r, column=rm_idx).value)
            if "no broker" not in rm_val.lower():
                continue

            ir_val = safe_str(ws.cell(row=r, column=ir_idx).value)
            dt_val = ws.cell(row=r, column=dt_idx).value
            yr = extract_year(dt_val)

            if (ir_val, yr) not in found_set:
                ws.cell(row=r, column=remarks_idx, value="Import Declaration Number not found in Customs Download")
                ws.cell(row=r, column=st_idx, value="Pending")
                not_found_count += 1

        wb_local.save(template)
        wb_local.close()
        print(f"  → {sheet_name}: Marked {not_found_count} No Broker records as not found")

    print("  → Step 8 completed")


def step9_bt_acc(template, bt_template_file):
    print("[Step 9] PROCTER & GAMBLE PHILIPPINE IN → BT ACC...")

    png_sheet = "PROCTER & GAMBLE PHILIPPINE IN"
    df_png = pd.read_excel(bt_template_file, sheet_name=png_sheet, header=0)
    print(f"  → {png_sheet} total {len(df_png)} rows")

    # Find columns
    entry_col = find_col(df_png.columns, "Entry No", "Entry Number", "Entry #")
    pay_col = find_col(df_png.columns, "Payment Date", "Payment date")
    dr_col = find_col(df_png.columns, "Duty Rate", "Duty rate")
    bda_col = find_col(df_png.columns, "Individual Customs Duty")
    vat_col = find_col(df_png.columns, "Individual VAT")
    co_col = find_col(df_png.columns, "Country of Origin")
    da_col = find_col(df_png.columns, "Individual Dutiable Value PHP")
    hts_col = find_col(df_png.columns, "Tariff Heading")

    # Build data DataFrame (only write value columns A-Q, indices 0-16)
    data = {}
    data["Purchasing Document Number"] = [None] * len(df_png)
    data["Company Code"] = ["403"] * len(df_png)
    data["Material"] = [None] * len(df_png)
    data["Import Declaration Number"] = df_png[entry_col].tolist() if entry_col else [None] * len(df_png)
    data["Duty Coding"] = ["B"] * len(df_png)
    data["TMSLoad#"] = [None] * len(df_png)
    data["Import Date"] = df_png[pay_col].tolist() if pay_col else [None] * len(df_png)
    data["Duty Rate"] = df_png[dr_col].tolist() if dr_col else [None] * len(df_png)
    data["Duty Currency"] = ["PHP"] * len(df_png)
    data["Basic Duty Amount"] = df_png[bda_col].tolist() if bda_col else [None] * len(df_png)
    data["Other Taxes Amount"] = [None] * len(df_png)
    data["VAT Amount"] = df_png[vat_col].tolist() if vat_col else [None] * len(df_png)
    data["Country of Origin"] = df_png[co_col].tolist() if co_col else [None] * len(df_png)
    data["Dutiable Amount"] = df_png[da_col].tolist() if da_col else [None] * len(df_png)
    data["HTS Code"] = df_png[hts_col].tolist() if hts_col else [None] * len(df_png)

    df_data = pd.DataFrame(data)

    wb = load_workbook(bt_template_file)
    ws = wb["BT ACC"]

    start_row = 3
    # Clear old data
    for r in range(start_row, ws.max_row + 1):
        for c in range(1, 31):
            ws.cell(row=r, column=c, value=None)

    # Write value columns A-Q
    col_names_aq = [
        "Purchasing Document Number",  # A: 0
        "Company Code",                 # B: 1
        "Material",                     # C: 2
        "Import Declaration Number",     # D: 3
        "Duty Coding",                  # E: 4
        "TMSLoad#",                     # F: 5
        "Import Date",                  # G: 6
        None,                           # H: 7 - formula
        None,                           # I: 8 - formula
        "Duty Rate",                    # J: 9
        "Duty Currency",                # K: 10
        "Basic Duty Amount",            # L: 11
        "Other Taxes Amount",           # M: 12
        "VAT Amount",                   # N: 13
        "Country of Origin",            # O: 14
        "Dutiable Amount",              # P: 15
        "HTS Code",                     # Q: 16
    ]

    for i in range(len(df_png)):
        r = start_row + i
        for col_idx, col_name in enumerate(col_names_aq, 1):
            if col_name is None:
                continue
            ws.cell(row=r, column=col_idx, value=df_data.iloc[i][col_name])

        # H column (8): Qty formula (INDEX+MATCH)
        ws.cell(row=r, column=8, value=f'=IFERROR(INDEX(\'EKPO\'!C:C,MATCH(A{r}&C{r},\'EKPO\'!E:E,0)),"")')
        # I column (9): UOM formula (INDEX+MATCH)
        ws.cell(row=r, column=9, value=f'=IFERROR(INDEX(\'EKPO\'!D:D,MATCH(A{r}&C{r},\'EKPO\'!E:E,0)),"")')

    # Formula columns R-AC (indices 17-28)
    formula_map = {
        17: '=SUM(L{row}:N{row})',
        18: '=SUMIFS(R:R,D:D,D{row})',
        19: "=INDEX('PROCTER & GAMBLE PHILIPPINE IN'!BF:BF,MATCH(D{row},'PROCTER & GAMBLE PHILIPPINE IN'!D:D,0))",
        20: '=T{row}-S{row}',
        21: '=COUNTIFS(D:D,D{row})',
        22: "=COUNTIF('PROCTER & GAMBLE PHILIPPINE IN'!D:D,D{row})",
        23: '=V{row}=W{row}',
        24: "=VLOOKUP(D{row},'Entry Num'!A:B,2,0)",
        25: f"=IFERROR(INDEX('EKPO'!C:C,MATCH(A{{row}}&C{{row}},'EKPO'!E:E,0)),\"\")",  # Z
        26: f"=IFERROR(INDEX('EKPO'!D:D,MATCH(A{{row}}&C{{row}},'EKPO'!E:E,0)),\"\")",  # AA
        27: '=Z{row}=H{row}',
        28: '=AA{row}=I{row}',
    }

    for i in range(len(df_png)):
        r = start_row + i
        for col_idx, formula_template in formula_map.items():
            ws.cell(row=r, column=col_idx + 1, value=formula_template.format(row=r))

    wb.save(bt_template_file)
    wb.close()

    print(f"  → BT ACC written ({len(df_png)} rows, starting from row 3, headers in rows 1-2 preserved)")
    print("  → Step 9 completed")


# ==================== Helper: build_3pl_map ====================
def norm_entry_key(val):
    """Normalize Entry No for matching: remove hyphens/spaces/underscores, uppercase, strip trailing .0"""
    if val is None:
        return ''
    s = str(val).strip()
    if s.lower() in ('nan', 'none', '', 'n/a'):
        return ''
    s = s.replace('-', '').replace('_', '').replace(' ', '').upper()
    s = re.sub(r'\.0$', '', s)
    return s


def clean_entry_no(val):
    """Unified cleaning of Entry No: remove hyphens/underscores/spaces, uppercase, strip leading/trailing non-alphanumeric, strip trailing .0"""
    if val is None:
        return ''
    s = str(val).strip()
    if s.lower() == 'nan' or s == 'N/A' or s == '':
        return ''
    s = s.replace('-', '').replace('_', '').replace(' ', '').upper()
    s = re.sub(r'^[^A-Za-z0-9]+', '', s)
    s = re.sub(r'[^A-Za-z0-9]+$', '', s)
    s = re.sub(r'\.0$', '', s)
    return s


def build_3pl_map(aging_root):
    """
    Build 3PL mapping from TCD files
    Returns: {normalized_entry_no: broker_name}
    """
    tcd_files = glob.glob(os.path.join(aging_root, "*TCD*.xlsx"))
    tcd_files = [f for f in tcd_files if not os.path.basename(f).startswith("~$")]

    # Priority order
    priority = ["CEVA", "DPWORLD", "MAERSK", "TW"]
    tcd_files.sort(key=lambda f: next(
        (i for i, k in enumerate(priority) if k in os.path.basename(f).upper()),
        len(priority)
    ))

    three_pl_map = {}

    for f in tcd_files:
        try:
            df = pd.read_excel(f, engine="openpyxl", dtype=str)
        except Exception as e:
            print(f"  [WARN] Failed to read TCD: {os.path.basename(f)} -> {e}")
            continue

        entry_col = None
        broker_col = None
        for c in df.columns:
            cn = str(c).strip().lower()
            if "import declaration" in cn or cn in ("entry no", "entry number", "entry #"):
                entry_col = c
            if "broker" in cn or "3pl" in cn:
                broker_col = c

        if not entry_col:
            continue

        for _, row in df.iterrows():
            k = norm_entry_key(row[entry_col])
            if not k:
                continue

            if broker_col:
                v = str(row[broker_col]).strip()
            else:
                v = ""

            if not v or v.lower() == "nan":
                # Infer from filename
                name = os.path.basename(f).upper()
                if "CEVA" in name:
                    v = "CEVA"
                elif "DPWORLD" in name:
                    v = "DPWORLD"
                elif "MAERSK" in name:
                    v = "MAERSK"
                elif "TW" in name:
                    v = "TRANSWORLD"
                else:
                    v = "UNKNOWN"

            # First hit kept (priority already sorted)
            if k not in three_pl_map:
                three_pl_map[k] = v

    print(f"  3PL map built: {len(three_pl_map)} entries")
    return three_pl_map


# ==================== Step 10: PO Lookup + 3PL + RMPM/FG ====================
def step10_po_lookup(template, bt_template_file):
    print("[Step 10] PO Lookup + 3PL + RMPM/FG ...")

    base = get_base_path()
    ssr_dir = os.path.join(base, "SSR")
    aging_root = base

    # ========== 1. Read BT ACC Entry No and Import Date ==========
    df_btacc = pd.read_excel(bt_template_file, sheet_name="BT ACC", header=1)
    ir_col = find_col(df_btacc.columns, "Import Declaration Number", "Entry No")
    date_col = find_col(df_btacc.columns, "Import Date")
    if ir_col is None:
        print("[ERROR] BT ACC cannot find Import Declaration Number")
        return

    entry_date_map = {}
    entry_bt_rows = {}
    for _, row in df_btacc.iterrows():
        e = clean_entry_no(row[ir_col])
        d = row[date_col] if date_col else None
        if e:
            if e not in entry_date_map:
                entry_date_map[e] = d
                entry_bt_rows[e] = 0
            else:
                if d and entry_date_map[e]:
                    try:
                        if pd.to_datetime(d, dayfirst=True) < pd.to_datetime(entry_date_map[e], dayfirst=True):
                            entry_date_map[e] = d
                    except Exception:
                        pass
            entry_bt_rows[e] += 1

    unique_entries = list(entry_date_map.keys())
    print(f"  BT ACC: {len(df_btacc)} rows, {len(unique_entries)} unique entries")

    # ========== 2. PROCTER → Entry No → Bill of Lading ==========
    png_sheet = "PROCTER & GAMBLE PHILIPPINE IN"
    df_png = pd.read_excel(bt_template_file, sheet_name=png_sheet, header=0)
    png_entry_col = find_col(df_png.columns, "Entry No", "Entry Number", "Entry #")
    bol_col = find_col(df_png.columns, "Bill of Lading")
    if png_entry_col is None or bol_col is None:
        print("[ERROR] PROCTER cannot find Entry No or Bill of Lading column")
        return

    png_map = {}
    for _, row in df_png.iterrows():
        e = clean_entry_no(row[png_entry_col])
        b = safe_str(row[bol_col])
        if e and b:
            png_map.setdefault(e, set()).add(b)
    print(f"  PROCTER: {len(png_map)} entries have BL")

    # ========== 3. RMPM Dashboard → Raw Data (fixed header row 2) ==========
    rmpm_map = {}
    rmpm_files = glob.glob(os.path.join(base, "*RMPM Dashboard PH Import*.xlsx"))
    if rmpm_files:
        df_rmpm = pd.read_excel(rmpm_files[0], sheet_name="Raw Data", header=1, dtype=str)

        bl_col = None
        po_col = None
        for c in df_rmpm.columns:
            cn = str(c).strip()
            if cn == "BL #" or cn == "BL":
                bl_col = c
            if cn == "PO":
                po_col = c

        print(f"  RMPM found columns: BL={bl_col}, PO={po_col}")

        if bl_col and po_col:
            count = 0
            for _, row in df_rmpm.iterrows():
                bl = safe_str(row[bl_col])
                po = safe_str(row[po_col])
                if bl and po and po.upper() != "N/A":
                    rmpm_map.setdefault(bl, []).append(po)
                    count += 1
            print(f"  RMPM: {len(rmpm_map)} unique BL, {count} total records")
        else:
            print(f"  [WARN] RMPM cannot find BL or PO column")
    else:
        print("  [WARN] Cannot find RMPM Dashboard file")

    # ========== 4. APAC → Tracker sheet ==========
    apac_map = {}
    apac_files = glob.glob(os.path.join(base, "*APAC standard PA tracker*.xlsx"))
    if apac_files:
        try:
            df_apac = pd.read_excel(apac_files[0], sheet_name="Tracker", header=0)
            bl_col = find_col(df_apac.columns, "Bill of Lading", "BL #", "BL")
            last_col = find_col(df_apac.columns, "PO# (LastLeg)", "PO#(LastLeg)")
            if bl_col and last_col:
                for _, row in df_apac.iterrows():
                    bl = safe_str(row[bl_col])
                    po = safe_str(row[last_col])
                    if bl and po and po.upper() != "N/A":
                        apac_map.setdefault(bl, []).append(po)
            print(f"  APAC: {len(apac_map)} unique BL")
        except Exception as e:
            print(f"  [WARN] Failed to read APAC: {e}")
    else:
        print("  [WARN] Cannot find APAC tracker file")

    # ========== 5. SSR: match by file name year ==========
    ssr_results = {}

    ssr_files = []
    if os.path.exists(ssr_dir):
        all_files = glob.glob(os.path.join(ssr_dir, "*.xlsx"))
        all_files = [f for f in all_files if not os.path.basename(f).startswith("~$")]
        all_files.sort()

        for f in all_files:
            fname = os.path.basename(f)
            m = re.search(r"DMS_PH_(\d{2})\d{4}", fname)
            if m:
                yy = int(m.group(1))
                full_year = 2000 + yy if yy < 50 else 1900 + yy
            else:
                full_year = None
            ssr_files.append((f, full_year))

# ==================== Step 10: PO Lookup + 3PL + RMPM/FG (continued) ====================

        print(f"  SSR files: {len(ssr_files)}")
        for f, yr in ssr_files:
            print(f"    {os.path.basename(f)} ({yr})")
    else:
        print(f"  [WARN] Cannot find SSR directory: {ssr_dir}")

    # Group files by year
    year_files = {}
    for fpath, yr in ssr_files:
        if yr not in year_files:
            year_files[yr] = []
        year_files[yr].append(fpath)

    # Determine year for each entry's Import Date
    entry_year_map = {}
    for entry in unique_entries:
        if entry in entry_date_map and entry_date_map[entry]:
            try:
                y = pd.to_datetime(entry_date_map[entry], dayfirst=True).year
                entry_year_map[entry] = y
            except Exception:
                entry_year_map[entry] = None
        else:
            entry_year_map[entry] = None

    # Pre-load all SSR file indexes
    print("  → Loading all SSR file indexes...")
    file_indexes = {}  # file_path -> (index_dict, year)

    for fpath, yr in ssr_files:
        try:
            df = pd.read_excel(fpath, sheet_name="SSR", header=0, dtype=str, engine="openpyxl")
        except Exception as e:
            print(f"    [WARN] Failed to read {os.path.basename(fpath)}: {e}")
            continue

        e_col = find_col(df.columns, "Import Customs Declaration #", "Entry No", "Entry Number", "Entry #")
        if e_col is None:
            print(f"    [SKIP] {os.path.basename(fpath)}: cannot find Entry column")
            continue
        last_col = find_col(df.columns, "PO# (Last Leg)", "PO# (LastLeg)")
        second_col = find_col(df.columns, "PO# (2nd Leg)", "PO# (2nd Leg)")
        po_col = last_col or second_col
        cargo_col = find_col(df.columns, "Cargo Type")
        if po_col is None:
            print(f"    [SKIP] {os.path.basename(fpath)}: cannot find PO column")
            continue

        file_idx = {}
        for _, row in df.iterrows():
            ev = clean_entry_no(row[e_col])
            if not ev:
                continue
            po_val = safe_str(row[po_col]).strip()
            cargo_val = safe_str(row[cargo_col]).strip() if cargo_col else ""
            if ev not in file_idx:
                file_idx[ev] = []
            if po_val and po_val.upper() != "N/A":
                if po_val not in [p for p, _ in file_idx[ev]]:
                    file_idx[ev].append((po_val, cargo_val))

        file_indexes[fpath] = (file_idx, yr)
        print(f"    → {os.path.basename(fpath)}: {len(file_idx)} entries")

    print(f"  Total loaded {len(file_indexes)} SSR file indexes")

    # Query: first match by file name year, then fallback to other years
    for entry in unique_entries:
        entry_yr = entry_year_map.get(entry)

        found = False
        if entry_yr and entry_yr in year_files:
            for fpath in year_files[entry_yr]:
                idx, yr = file_indexes.get(fpath, ({}, None))
                if entry in idx:
                    ssr_results[entry] = idx[entry]
                    found = True
                    print(f"    {entry}: found {len(idx[entry])} PO(s) in {os.path.basename(fpath)} ({yr})")
                    break

        if not found:
            for yr in sorted(year_files.keys()):
                if entry_yr and yr == entry_yr:
                    continue
                for fpath in year_files[yr]:
                    idx, _ = file_indexes.get(fpath, ({}, None))
                    if entry in idx:
                        ssr_results[entry] = idx[entry]
                        found = True
                        print(f"    {entry}: found {len(idx[entry])} PO(s) in {os.path.basename(fpath)} ({yr})")
                        break
                if found:
                    break

    stats_found = len(ssr_results)
    stats_not_found = len(unique_entries) - stats_found
    total_pos = sum(len(v) for v in ssr_results.values())
    print(f"  SSR query complete: found {stats_found} entries ({total_pos} POs), not found {stats_not_found} entries")

    # ========== 6. Three-level query: RMPM → APAC → SSR ==========
    result = {}
    result_cargo = {}  # entry -> cargo_type
    stats = {"rmpm": 0, "apac": 0, "ssr": 0, "none": 0}

    for entry in unique_entries:
        bl_set = png_map.get(entry, set())
        po_list = []

        # Level 1: RMPM
        for bl in bl_set:
            if bl in rmpm_map:
                po_list.extend(rmpm_map[bl])
                stats["rmpm"] += len(rmpm_map[bl])

        # Level 2: APAC
        if not po_list:
            for bl in bl_set:
                if bl in apac_map:
                    po_list.extend(apac_map[bl])
                    stats["apac"] += len(apac_map[bl])

        # Level 3: SSR (with Cargo Type)
        if not po_list and entry in ssr_results:
            for po_val, cargo_val in ssr_results[entry]:
                po_list.append(po_val)
                if entry not in result_cargo:
                    result_cargo[entry] = cargo_val
            stats["ssr"] += 1

        # Deduplicate preserving order
        po_list = list(dict.fromkeys(po_list))

        if not po_list:
            stats["none"] += 1

        result[entry] = po_list

    print(
        f"  Stats: RMPM={stats['rmpm']} POs, APAC={stats['apac']} POs, SSR={stats['ssr']} entries, Not found={stats['none']} entries")

    # ========== 7. Build 3PL mapping from DMS files ==========
    print("  → Building 3PL mapping from DMS files...")

    three_pl_map = {}

    # Collect all DMS files
    dms_files = []
    if os.path.exists(ssr_dir):
        all_files = glob.glob(os.path.join(ssr_dir, "*.xlsx"))
        all_files = [f for f in all_files if not os.path.basename(f).startswith("~$")]
        all_files.sort()
        dms_files = all_files

    if not dms_files:
        dms_files = glob.glob(os.path.join(aging_root, "*DMS_PH*.*"))

    print(f"  Found {len(dms_files)} DMS files for 3PL mapping")

    for fpath in dms_files:
        try:
            try:
                df = pd.read_excel(fpath, sheet_name="SSR", header=0, dtype=str, engine="openpyxl")
            except Exception:
                df = pd.read_excel(fpath, header=0, dtype=str, engine="openpyxl")

            entry_col = None
            broker_col = None

            for c in df.columns:
                cn = str(c).strip().lower()
                if "import customs declaration" in cn or cn in ("entry no", "entry number", "entry #"):
                    entry_col = c
                if "broker" in cn or "3pl" in cn or "agent" in cn:
                    broker_col = c

            if not entry_col:
                print(f"    [SKIP] {os.path.basename(fpath)}: cannot find Entry column")
                continue

            fname_upper = os.path.basename(fpath).upper()

            for _, row in df.iterrows():
                k = clean_entry_no(row[entry_col])
                if not k:
                    continue

                if broker_col:
                    v = safe_str(row[broker_col])
                else:
                    v = ""

                if not v or v.lower() == "nan":
                    if "CEVA" in fname_upper:
                        v = "CEVA"
                    elif "DPWORLD" in fname_upper:
                        v = "DPWORLD"
                    elif "MAERSK" in fname_upper:
                        v = "MAERSK"
                    elif "TW" in fname_upper or "TRANSWORLD" in fname_upper:
                        v = "TRANSWORLD"
                    elif "ATI" in fname_upper:
                        v = "ATI"
                    elif "YUSEN" in fname_upper:
                        v = "YUSEN"
                    else:
                        v = "UNKNOWN"

                if k not in three_pl_map:
                    three_pl_map[k] = v

        except Exception as e:
            print(f"    [WARN] Failed to read {os.path.basename(fpath)}: {e}")
            continue

    print(f"  3PL map built: {len(three_pl_map)} entries")

    # ========== 8. Write Entry Num sheet ==========
    wb = load_workbook(bt_template_file)
    ws_entry = wb["Entry Num"]
    entry_header = [cell.value for cell in ws_entry[1]]

    entry_idx = None
    po_idx = None
    rmfg_idx = None
    remarks_idx = None
    payment_idx = None
    threepl_idx = None
    sad_idx = None
    bl_idx = None

    for i, h in enumerate(entry_header, 1):
        if h is None:
            continue
        hn = str(h).strip().lower()
        if "entry" in hn and "import declaration" not in hn:
            entry_idx = i
        elif "po" == hn or "purchasing" in hn or "po#" in hn:
            po_idx = i
        elif "rmpm" in hn or "cargo" in hn or "fg" in hn:
            rmfg_idx = i
        elif "remark" in hn:
            remarks_idx = i
        elif "payment" in hn:
            payment_idx = i
        elif "3pl" in hn or "broker" in hn:
            threepl_idx = i
        elif "sad" in hn:
            sad_idx = i
        elif "bl" == hn or "bill of lading" in hn:
            bl_idx = i

    print(f"  [DEBUG] Column positions: ENTRY={entry_idx}, 3PL={threepl_idx}, RMPM/FG={rmfg_idx}, "
          f"REMARKS={remarks_idx}, Payment={payment_idx}, PO={po_idx}")

    if entry_idx is None:
        print("[ERROR] Entry Num sheet missing ENTRY column")
        wb.close()
        return

    def ensure_col(name, preferred_idx, keyword_check):
        col_idx = None
        for i, h in enumerate(entry_header, 1):
            if h and keyword_check(str(h).strip().lower()):
                col_idx = i
                break
        if col_idx is None:
            col_idx = preferred_idx
            ws_entry.cell(row=1, column=col_idx, value=name)
        return col_idx

    entry_idx = ensure_col("ENTRY", 1, lambda x: "entry" in x and "import declaration" not in x)
    threepl_idx = ensure_col("3PL", 2, lambda x: "3pl" in x or "broker" in x)
    rmfg_idx = ensure_col("RMPM/FG", 3, lambda x: "rmpm" in x or "cargo" in x or "fg" in x)
    remarks_idx = ensure_col("REMARKS", 4, lambda x: "remark" in x)
    sad_idx = ensure_col("SAD Updated", 5, lambda x: "sad" in x)
    payment_idx = ensure_col("Payment Date", 6, lambda x: "payment" in x)
    bl_idx = ensure_col("BL", 7, lambda x: x == "bl" or "bill of lading" in x)
    po_idx = ensure_col("PO", 8, lambda x: x == "po" or "purchasing" in x or "po#" in x)

    print(f"  [DEBUG] Final columns: ENTRY={entry_idx}, 3PL={threepl_idx}, RMPM/FG={rmfg_idx}, "
          f"REMARKS={remarks_idx}, Payment={payment_idx}, PO={po_idx}")

    current_row = 2

    for entry in unique_entries:
        po_list = result.get(entry, [])
        pay_date = entry_date_map.get(entry)
        cargo_val = result_cargo.get(entry, "")
        three_pl_val = three_pl_map.get(clean_entry_no(entry), "")

        if not po_list:
            ws_entry.cell(row=current_row, column=entry_idx, value=entry)
            ws_entry.cell(row=current_row, column=threepl_idx, value=three_pl_val)
            ws_entry.cell(row=current_row, column=rmfg_idx, value=cargo_val)
            ws_entry.cell(row=current_row, column=remarks_idx, value="No PO found")
            ws_entry.cell(row=current_row, column=payment_idx, value=pay_date)
            ws_entry.cell(row=current_row, column=bl_idx, value=None)
            ws_entry.cell(row=current_row, column=po_idx, value=None)
            current_row += 1
        else:
            for po_val in po_list:
                ws_entry.cell(row=current_row, column=entry_idx, value=entry)
                ws_entry.cell(row=current_row, column=threepl_idx, value=three_pl_val)
                ws_entry.cell(row=current_row, column=rmfg_idx, value=cargo_val)
                ws_entry.cell(row=current_row, column=remarks_idx, value=None)   # No longer writes "Lines mismatch"
                ws_entry.cell(row=current_row, column=payment_idx, value=pay_date)
                ws_entry.cell(row=current_row, column=bl_idx, value=None)
                ws_entry.cell(row=current_row, column=po_idx, value=po_val)
                current_row += 1

    print(f"  → Entry Num sheet written, total {current_row - 2} data rows")

    # ========== 9. Write Purchasing Document Number into BT ACC ==========
    print("  → Writing Purchasing Document Number into BT ACC...")

    ws_bt = wb["BT ACC"]
    bt_header = [cell.value for cell in ws_bt[2]]

    bt_entry_idx = None
    bt_pdoc_idx = None
    for i, h in enumerate(bt_header, 1):
        hn = norm_col(str(h)) if h else ""
        if "import declaration number" in hn or "entry no" in hn:
            bt_entry_idx = i
        if "purchasing document number" in hn:
            bt_pdoc_idx = i

    if bt_entry_idx and bt_pdoc_idx:
        entry_bt_indices = {}
        for r in range(3, ws_bt.max_row + 1):
            ev = clean_entry_no(ws_bt.cell(row=r, column=bt_entry_idx).value)
            if ev:
                entry_bt_indices.setdefault(ev, []).append(r)

        update_count = 0
        skip_count = 0

        for entry, bt_indices in entry_bt_indices.items():
            po_list = result.get(entry, [])
            po_count = len(po_list)
            bt_count = len(bt_indices)

            if bt_count == po_count and po_count > 0:
                for i, r in enumerate(bt_indices):
                    ws_bt.cell(row=r, column=bt_pdoc_idx, value=po_list[i])
                update_count += bt_count
            else:
                skip_count += 1

        print(f"  → BT ACC: updated {update_count} rows, skipped {skip_count} entries (row count mismatch)")
    else:
        print(f"  [WARN] BT ACC missing columns: entry={bt_entry_idx}, pdoc={bt_pdoc_idx}")

    # ★★★ Unified save ★★★
    wb.save(bt_template_file)
    wb.close()
    print("  → Step 10 completed")


# ==================== Main Pipeline ====================
import time


def phase2():
    base = get_base_path()
    template = os.path.join(base, "TCD Aging Accrual Report TEMPLATE.xlsx")
    bt_template_file = os.path.join(base, "BT TEMPLATE ACCRUAL.xlsx")

    print("=" * 55)
    print("Phase 2: Step 5 → 6 → 7 → 8 → 9 → 10")
    print(f"Working directory: {base}")
    print("=" * 55)

    if not os.path.exists(template):
        print("[ERROR] Cannot find template:", template)
        return

    t_start = time.time()

    t_step = time.time()
    ok = step5_export_to_sapraw(template)
    print(f"  ⏱ Step 5 elapsed: {time.time() - t_step:.2f}s")
    if not ok:
        return

    t_step = time.time()
    step6_sap_raw_to_sap_bt(template)
    print(f"  ⏱ Step 6 elapsed: {time.time() - t_step:.2f}s")

    t_step = time.time()
    step7_amount_mismatch(template)
    print(f"  ⏱ Step 7 elapsed: {time.time() - t_step:.2f}s")

    t_step = time.time()
    step8_no_broker(template, bt_template_file)
    print(f"  ⏱ Step 8 elapsed: {time.time() - t_step:.2f}s")

    t_step = time.time()
    step9_bt_acc(template, bt_template_file)
    print(f"  ⏱ Step 9 elapsed: {time.time() - t_step:.2f}s")

    t_step = time.time()
    step10_po_lookup(template, bt_template_file)
    print(f"  ⏱ Step 10 elapsed: {time.time() - t_step:.2f}s")

    print(f"\n  ⏱ Phase 2 total elapsed: {time.time() - t_start:.2f}s")
    print("\n[Phase 2 Step 5-10 Completed]")


if __name__ == "__main__":
    phase2()