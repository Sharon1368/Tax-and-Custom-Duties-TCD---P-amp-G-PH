"""
TCD-PH v2 Full Automation (Patched)
  Patched version with user-requested changes.
  PDF download (download_all) and PDF reading (read_total_global_tax) are UNTOUCHED.
  Only modified: Step 3, Step 6, Step 10(formulas), Step 11(QTY/UOM), Export, process_pdfs_and_bt
  Execution order changed: formulas -> QTY/UOM -> AMT DIFF/PDF -> UPLOADED -> Export
"""
import os, re, time, sys, zipfile, glob, subprocess
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import pdfplumber
import pytesseract
import pandas as pd
from datetime import datetime
import warnings
from playwright.sync_api import sync_playwright

warnings.filterwarnings('ignore')

# ============================================================
# Config
# ============================================================
TCD_DIR = Path.home() / "Desktop" / "TCD"
PDF_DIR = TCD_DIR / "PDF"
BT_EXCEL_PATH = TCD_DIR / "BT TEMPLATE.xlsx"
ENTRY_NUM_SHEET = "Entry Num"
BT_SHEET = "BT"
CDP_URL = "http://localhost:9222"

tess_path = TCD_DIR / "Tesseract-OCR" / "tesseract.exe"
if tess_path.exists():
    pytesseract.pytesseract.tesseract_cmd = str(tess_path)

TCD_COLS = [
    'Purchasing Document Number', 'Company Code', 'Material',
    'Import Declaration Number', 'Duty Coding', 'TMS Load #', 'Import Date',
    'Qty', 'UOM', 'Duty Rate', 'Duty Currency', 'Basic Duty Amount',
    'Other Taxes Amount', 'VAT Amount', 'Country of Origin',
    'Dutiable Amount', 'HTS Code'
]

CLIENTS = {
    "ATI": {
        "url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FATI%2FFY2526%20Raw%20Shipping%20Documents",
        "keyword": "C319574"
    },
    "CEVA": {
        "url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?csf=1&web=1&id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FCEVA",
        "keyword": "C18964"
    },
    "TRANSWORLD": {
        "url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?csf=1&web=1&id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FTW%20(NEW)",
        "keyword": "C18437"
    },
    "MAERSK": {
        "url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?e=bzBDoH&sharingv2=true&fromShare=true&at=9&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNjA5MTcxMjIyMiIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D&CID=161445a2%2Da04b%2Ff000%2Dcbbf%2D47299236efcc&cidOR=SPO&FolderCTID=0x012000CCD2F3217E34DD4781511DEE93F8125E&id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FMAERSK%2FDocument%20Retention",
        "keyword": "C320527"
    },
    "YUSEN": {
        "url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?csf=1&web=1&id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FYusen",
        "keyword": "C14915"
    }
}


# ============================================================
# Utility functions
# ============================================================
def clean_entry_no(val):
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

def get_val(v):
    if v is None:
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0

def find_files(pattern):
    results = [f for f in TCD_DIR.glob(pattern)]
    return results

def extract_date_range(csv_file):
    """Extract everything after 'PROCTER & GAMBLE PHILIPPINES INC.' in CSV filename"""
    stem = csv_file.stem
    prefix = "PROCTER & GAMBLE PHILIPPINES INC."
    if stem.startswith(prefix):
        date_range = stem[len(prefix):].strip()
        date_range = date_range.lstrip("._").strip()
    else:
        m = re.search(r'PROCTER\s*&\s*GAMBLE\s*PHILIPPINES\s*INC\.\s*(.+)', stem, re.IGNORECASE)
        if m:
            date_range = m.group(1).strip().lstrip("._").strip()
        else:
            date_range = "EXPORT"
    if not date_range:
        date_range = "EXPORT"
    return date_range


# ============================================================
# Step 0: Auto-start Edge debug mode
# ============================================================
def ensure_edge_debug():
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    os.chdir(str(TCD_DIR))

    try:
        import urllib.request
        urllib.request.urlopen("http://localhost:9222/json/version", timeout=2)
        print("[OK] Edge debug port already open, skipping launch")
        return
    except:
        pass

    print("Closing all Edge instances...")
    subprocess.run(["taskkill", "/f", "/im", "msedge.exe"],
                   capture_output=True, shell=True)
    time.sleep(2)

    print("Starting Edge in debug mode...")
    edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    profile_dir = r"C:\EdgeDebugProfile"
    cmd = [
        edge_path,
        "--remote-debugging-port=9222",
        f"--user-data-dir={profile_dir}",
        "--remote-allow-origins=*",
        "--no-first-run",
        "https://pgone.sharepoint.com"
    ]
    subprocess.Popen(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(5)

    print("Please login to SharePoint in the browser window")
    input("Press Enter after login...")


# ============================================================
# Step 1: Clean template
# ============================================================
def step1_clean_template(wb):
    print("\n[Step 1] Clean main template")

    ws_in = wb['PROCTER & GAMBLE PHILIPPINE IN']
    ws_in.auto_filter.ref = None
    max_row = ws_in.max_row
    if max_row > 1:
        ws_in.delete_rows(2, max_row - 1)
    print(f"  [OK] PROCTER & GAMBLE PHILIPPINE IN cleaned, {ws_in.max_row} rows remaining")

    ws_bt = wb[BT_SHEET]
    ws_bt.auto_filter.ref = None
    max_row = ws_bt.max_row
    if max_row > 2:
        ws_bt.delete_rows(3, max_row - 2)
    print(f"  [OK] BT cleaned, {ws_bt.max_row} rows remaining")

    return wb


# ============================================================
# Step 2: Process CSV file (now includes row height setting)
# ============================================================
def step2_process_csv(wb):
    print("\n[Step 2] Process CSV file")

    csv_files = [f for f in TCD_DIR.glob("*PROCTER*CSV")]
    if not csv_files:
        # 兜底：更宽泛匹配
        csv_files = [f for f in TCD_DIR.glob("*PROCTER*.*") if f.suffix.upper() == '.CSV']
    if not csv_files:
        print("  [ERROR] No CSV file found")
        return wb, None

    CSV_FILE = csv_files[0]
    print(f"  CSV file found: {CSV_FILE.name}")

    df_csv = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
    print(f"  CSV rows: {len(df_csv)}, columns: {len(df_csv.columns)}")

    df_csv['Entry No'] = df_csv['Entry No'].astype(str).str.replace('-', '', regex=False)

    valid_dates = pd.to_datetime(df_csv['Arrival Date'], errors='coerce').dropna()
    if len(valid_dates) > 0:
        recent_date = valid_dates.max().strftime('%m/%d/%Y')
        print(f"  Most recent date: {recent_date}")
        df_csv['Arrival Date'] = df_csv['Arrival Date'].fillna(recent_date)

    ws_in = wb['PROCTER & GAMBLE PHILIPPINE IN']

    for col_idx, col_name in enumerate(df_csv.columns, 1):
        ws_in.cell(row=1, column=col_idx, value=col_name)

    for row_idx, row_data in df_csv.iterrows():
        for col_idx, value in enumerate(row_data, 1):
            ws_in.cell(row=row_idx + 2, column=col_idx, value=value)

    print(f"  [OK] CSV data written, {len(df_csv)} rows total")
    return wb, df_csv


# ============================================================
# Step 2b: Uniform row height (moved from old Step 8b)
# ============================================================
def step2b_uniform_row_height(wb):
    print("\n[Step 2b] Uniform row height for PROCTER & GAMBLE PHILIPPINE IN")
    ws_in = wb['PROCTER & GAMBLE PHILIPPINE IN']
    for row in range(1, ws_in.max_row + 1):
        ws_in.row_dimensions[row].height = 15
    print(f"  [OK] Row height adjusted to 15, {ws_in.max_row} rows total")
    return wb


# ============================================================
# Step 3: Process EKPO file —— CHANGED: skip rows where Del. Indicator contains L
# ============================================================
def step3_process_ekpo(wb):
    print("\n[Step 3] Process EKPO file")

    ekpo_files = [f for f in TCD_DIR.glob("*EKPO*.*")]
    if not ekpo_files:
        print("  [ERROR] No EKPO file found")
        return wb

    EKPO_FILE = ekpo_files[0]
    print(f"  EKPO file found: {EKPO_FILE.name}")

    df_ekpo = pd.read_csv(
        EKPO_FILE,
        sep='\t',
        encoding='utf-16',
        skiprows=3,
        dtype=str
    )
    print(f"  EKPO total rows: {len(df_ekpo)}, total columns: {len(df_ekpo.columns)}")

    # 定位 Del. Indicator 列
    del_indicator_col = None
    for c in df_ekpo.columns:
        if str(c).strip() == 'Del. Indicator':
            del_indicator_col = c
            break
    if del_indicator_col:
        print(f"  [INFO] Del. Indicator column found: '{del_indicator_col}'")
    else:
        print("  [WARN] Del. Indicator column NOT found, will not filter")

    ws_ekpo = wb['EKPO']
    ws_ekpo.delete_rows(2, ws_ekpo.max_row - 1)

    headers_ekpo = ['Purchasing Doc.', 'Material', 'PO Quantity', 'Order Unit', 'Ref']
    for col_idx, header in enumerate(headers_ekpo, 1):
        ws_ekpo.cell(row=1, column=col_idx, value=header)

    data_count = 0
    skip_count = 0
    for row_idx, row_data in df_ekpo.iterrows():
        try:
            purchasing_doc = row_data['Purchasing Doc.']
            material = row_data['Material']
            po_qty = row_data['PO Quantity']
            order_unit = row_data['Order Unit']

            if pd.isna(purchasing_doc) or str(purchasing_doc).strip() == '':
                continue

            # ===== CHANGED: skip if Del. Indicator contains L =====
            if del_indicator_col is not None:
                del_val = str(row_data[del_indicator_col]).strip().upper()
                if 'L' in del_val:
                    skip_count += 1
                    continue

            if isinstance(po_qty, str):
                po_qty = po_qty.replace(',', '').strip()
                try:
                    po_qty = float(po_qty)
                except ValueError:
                    pass

            ref = f"{purchasing_doc}{material}"

            ws_ekpo.cell(row=data_count + 2, column=1, value=purchasing_doc)
            ws_ekpo.cell(row=data_count + 2, column=2, value=material)
            ws_ekpo.cell(row=data_count + 2, column=3, value=po_qty)
            ws_ekpo.cell(row=data_count + 2, column=4, value=order_unit)
            ws_ekpo.cell(row=data_count + 2, column=5, value=ref)

            data_count += 1
        except Exception as e:
            continue

    print(f"  [OK] EKPO data written, {data_count} rows total")
    if skip_count > 0:
        print(f"  [INFO] Skipped {skip_count} rows (Del. Indicator contains 'L')")
    return wb


# ============================================================
# Step 4: Process 3PL files (UNTOUCHED)
# ============================================================
def step4_process_3pl(wb):
    print("\n[Step 4] Process four 3PL files")

    tcd_files_config = {
        'CEVA_TCD': TCD_DIR / "CEVA_TCD.xlsx",
        'DPWORLD_TCD': TCD_DIR / "DPWORLD_TCD.xlsx",
        'MAERSK_TCD': TCD_DIR / "MAERSK_TCD.xlsx",
        'TW_TCD': TCD_DIR / "TW_TCD.xlsx"
    }

    for sheet_name, file_path in tcd_files_config.items():
        if not file_path.exists():
            print(f"  [WARN] {sheet_name} file not found, skipping")
            continue

        print(f"  --- Processing {sheet_name} ---")
        df_tcd = pd.read_excel(file_path, engine='openpyxl')
        df_tcd = df_tcd.iloc[:, :17]
        df_tcd.columns = TCD_COLS
        print(f"  Rows: {len(df_tcd)}, Columns: {len(df_tcd.columns)}")

        df_tcd['Import Declaration Number'] = df_tcd['Import Declaration Number'].astype(str).str.replace(r'[-\s]', '', regex=True)

        df_tcd['Import Date'] = pd.to_datetime(
            df_tcd['Import Date'].astype(str).str.strip(),
            errors='coerce'
        ).apply(lambda x: x.strftime('%Y%m%d') if pd.notna(x) else '')

        ws_tcd = wb[sheet_name]
        ws_tcd.delete_rows(2, ws_tcd.max_row - 1)

        for col_idx, col_name in enumerate(df_tcd.columns, 1):
            ws_tcd.cell(row=1, column=col_idx, value=col_name)

        for row_idx, row_data in df_tcd.iterrows():
            for col_idx, value in enumerate(row_data, 1):
                ws_tcd.cell(row=row_idx + 2, column=col_idx, value=value)

        print(f"  [OK] {sheet_name} data written, {len(df_tcd)} rows total")

    print("  [OK] All 3PL files processed")
    return wb


# ============================================================
# Step 5: Process Entry Num sheet (UNTOUCHED)
# ============================================================
def step5_process_entry_num(wb, df_csv):
    print("\n[Step 5] Process Entry Num sheet")

    ws_in = wb['PROCTER & GAMBLE PHILIPPINE IN']

    if df_csv is not None:
        entry_nos = df_csv['Entry No'].astype(str).str.strip().tolist()
    else:
        entry_nos = []
        for row in range(2, ws_in.max_row + 1):
            entry_no = ws_in.cell(row=row, column=4).value
            if entry_no and entry_no not in entry_nos:
                entry_nos.append(entry_no)

    csv_entry_nos = [clean_entry_no(e) for e in entry_nos]
    seen = set()
    unique_entry_nos = []
    for e in csv_entry_nos:
        if e and e not in seen:
            seen.add(e)
            unique_entry_nos.append(e)

    print(f"  CSV unique Entry No count: {len(unique_entry_nos)}")

    dms_files = [f for f in TCD_DIR.glob("*DMS_PH*.*")]
    df_ssr = pd.DataFrame()
    if dms_files:
        try:
            df_ssr = pd.read_excel(dms_files[0], sheet_name='SSR', engine='openpyxl')
            print(f"  SSR sheet rows: {len(df_ssr)}, columns: {len(df_ssr.columns)}")
        except Exception as e:
            print(f"  [WARN] Failed to read DMS file: {e}")
    else:
        print("  [WARN] No DMS_PH file found")

    ssr_dict = {}
    for _, row in df_ssr.iterrows():
        entry_no_raw = row.get('Import Customs Declaration #', '')
        entry_no_key = clean_entry_no(entry_no_raw)
        if entry_no_key and entry_no_key not in ssr_dict:
            broker = row.get('Broker (Import)', '')
            cargo_type = row.get('Cargo Type (FG/RMPM)', '')
            ssr_dict[entry_no_key] = {'broker': broker, 'cargo_type': cargo_type}

    print(f"  SSR data dictionary built, {len(ssr_dict)} unique Entry No entries")

    in_header_map = {}
    for cell in ws_in[1]:
        if cell.value is not None:
            in_header_map[str(cell.value).strip()] = cell.column

    paydate_col = in_header_map.get('Payment Date')
    bl_col = in_header_map.get('Bill of Lading')

    in_data = {}
    for row in range(2, ws_in.max_row + 1):
        entry_raw = ws_in.cell(row=row, column=4).value
        entry_key = clean_entry_no(entry_raw)
        if not entry_key:
            continue
        if entry_key not in in_data:
            payment_date = ws_in.cell(row=row, column=paydate_col).value if paydate_col else ''
            bl = ws_in.cell(row=row, column=bl_col).value if bl_col else ''
            in_data[entry_key] = {'payment_date': payment_date, 'bl': bl}

    print(f"  PROCTER data dictionary built, {len(in_data)} unique Entry No entries")

    ws_entry = wb[ENTRY_NUM_SHEET]
    ws_entry.delete_rows(2, ws_entry.max_row - 1)

    entry_headers = ['Entry', '3PL', 'RMPM/FG', 'REMARKS', 'SAD Updated', 'Payment Date', 'BL']
    for col_idx, header in enumerate(entry_headers, 1):
        ws_entry.cell(row=1, column=col_idx, value=header)

    row_num = 2
    match_count = 0
    no_match_list = []

    for entry_no in unique_entry_nos:
        ssr_match = ssr_dict.get(entry_no, {})
        broker = ssr_match.get('broker', '')
        cargo_type = ssr_match.get('cargo_type', '')

        if ssr_match:
            match_count += 1
        else:
            no_match_list.append(entry_no)

        in_match = in_data.get(entry_no, {})
        payment_date = in_match.get('payment_date', '')
        bl = in_match.get('bl', '')

        remarks = ''
        if cargo_type:
            cargo_lower = str(cargo_type).lower()
            if 'sample' in cargo_lower or 'spare part' in cargo_lower:
                remarks = 'NO ACTION'

        ws_entry.cell(row=row_num, column=1, value=entry_no)
        ws_entry.cell(row=row_num, column=2, value=broker)
        ws_entry.cell(row=row_num, column=3, value=cargo_type)
        ws_entry.cell(row=row_num, column=4, value=remarks)
        ws_entry.cell(row=row_num, column=5, value='')
        ws_entry.cell(row=row_num, column=6, value=payment_date)
        ws_entry.cell(row=row_num, column=7, value=bl)

        row_num += 1

    print(f"  [OK] Entry Num sheet filled, {len(unique_entry_nos)} records total")
    print(f"  Matched SSR data: {match_count} records")
    if no_match_list:
        print(f"  [WARN] Entry No count not matched to SSR: {len(no_match_list)}")
        print(f"  First 10: {no_match_list[:10]}")

    return wb


# ============================================================
# Step 6: Fill BT sheet
# ============================================================
def step6_fill_bt(wb, df_csv):
    print("\n[Step 6] Fill BT sheet")

    # Build TCD data dictionary in priority order
    tcd_order = ['CEVA_TCD', 'DPWORLD_TCD', 'MAERSK_TCD', 'TW_TCD']
    tcd_data = {}

    for sheet_name in tcd_order:
        if sheet_name not in wb.sheetnames:
            continue
        ws_tcd = wb[sheet_name]
        for row in range(2, ws_tcd.max_row + 1):
            imp_dec_no = ws_tcd.cell(row=row, column=4).value
            if imp_dec_no:
                entry_no_key = str(imp_dec_no).strip().replace('-', '').replace(' ', '')
                if entry_no_key not in tcd_data:
                    row_values = []
                    for col in range(1, 18):
                        row_values.append(ws_tcd.cell(row=row, column=col).value)
                    tcd_data[entry_no_key] = row_values

    print(f"  TCD data dictionary built, {len(tcd_data)} unique Entry No entries")

    # Get Entry No list from CSV
    if df_csv is not None:
        csv_entry_nos = df_csv['Entry No'].astype(str).str.strip().tolist()
    else:
        csv_entry_nos = []
    print(f"  CSV has {len(csv_entry_nos)} Entry No entries")

    # Clear existing data in BT sheet
    ws_bt = wb[BT_SHEET]
    if ws_bt.max_row > 2:
        ws_bt.delete_rows(3, ws_bt.max_row - 2)

    # Reference to Entry Num sheet
    ws_entry = wb[ENTRY_NUM_SHEET]

    # Write data
    write_count = 0
    not_found = []

    for entry_no in csv_entry_nos:
        if entry_no in tcd_data:
            row_values = tcd_data[entry_no]
            for col_idx, value in enumerate(row_values, 1):
                ws_bt.cell(row=write_count + 3, column=col_idx, value=value)
            write_count += 1
        else:
            not_found.append(entry_no)
            # Mark "3PL not found" in Entry Num REMARKS column (only if empty)
            for r in range(2, ws_entry.max_row + 1):
                entry_val = str(ws_entry.cell(row=r, column=1).value or '').strip()
                if entry_val == entry_no:
                    current_remark = str(ws_entry.cell(row=r, column=4).value or '').strip()
                    if current_remark == '' or current_remark == 'None':
                        ws_entry.cell(row=r, column=4, value='3PL not found')
                        print(f"  [3PL not found] {entry_no} -> REMARKS marked")
                    break

    print(f"  [OK] BT sheet filled with {write_count} records")
    if not_found:
        print(f"  [WARN] Unmatched Entry No count: {len(not_found)}")


    return wb


# ============================================================
# Step 7: Fill Company Code (UNTOUCHED)
# ============================================================
def step7_fill_company_code(wb):
    print("\n[Step 7] Fill BT sheet Company Code")
    ws_bt = wb[BT_SHEET]
    for row in range(3, ws_bt.max_row + 1):
        ws_bt.cell(row=row, column=2, value=403)
    print(f"  [OK] Company Code filled with 403, {ws_bt.max_row - 2} rows total")
    return wb


# ============================================================
# Step 8: Write formulas —— CHANGED: col 25/26 NOT overwritten
# ============================================================
def step8_write_formulas(wb):
    print("\n[Step 8] Write formulas to BT sheet columns R~AB")
    ws_bt = wb[BT_SHEET]
    last_row = ws_bt.max_row

    formula_map = {
        18: '=SUM(L{row}:N{row})',
        19: '=SUMIFS(R:R,D:D,D{row})',
        20: "=INDEX('PROCTER & GAMBLE PHILIPPINE IN'!BF:BF,MATCH(D{row},'PROCTER & GAMBLE PHILIPPINE IN'!D:D,0))",
        21: '=T{row}-S{row}',
        22: '=COUNTIFS(D:D,D{row})',
        23: "=COUNTIF('PROCTER & GAMBLE PHILIPPINE IN'!D:D,D{row})",
        24: "=VLOOKUP(D{row},'Entry Num'!A:B,2,0)",
        # 25: EKPO QTY —— 由 Step 9 (QTY/UOM) 填写，不覆盖
        # 26: EKPO UOM —— 由 Step 9 (QTY/UOM) 填写，不覆盖
        27: '=Y{row}=H{row}',
        28: '=Z{row}=I{row}',
    }

    formula_headers = {
        18: 'BT DUTY+TAX+VAT AMT',
        19: 'AMT PER ENTRY BT',
        20: 'Total D/T Paid CDEC',
        21: 'AMT DIFF',
        22: 'BT COUNT',
        23: 'CDEC COUNT',
        24: 'BROKER',
        25: 'EKPO QTY',
        26: 'EKPO UOM',
        27: 'QTY Check',
        28: 'UOM Check',
    }

    for col, header in formula_headers.items():
        ws_bt.cell(row=2, column=col, value=header)

    for row in range(3, last_row + 1):
        for col, formula in formula_map.items():
            ws_bt.cell(row=row, column=col, value=formula.format(row=row))

    print(f"  [OK] Formulas written to columns R~AB (cols 25/26 preserved), {last_row - 2} rows total")
    return wb


# ============================================================
# Step 9: QTY/UOM overwrite —— CHANGED: A+C -> EKPO Ref, compare & overwrite, red font
# ============================================================
def step9_qty_uom_overwrite(wb):
    print("\n[Step 9] QTY / UOM Check & Overwrite")

    ws_bt = wb[BT_SHEET]
    ws_ekpo = wb['EKPO']

    # Build EKPO mapping: Ref (E列, col 5) -> {qty: col3, uom: col4}
    ekpo_map = {}
    for row in range(2, ws_ekpo.max_row + 1):
        e_val = str(ws_ekpo.cell(row=row, column=5).value or '').strip()
        qty = ws_ekpo.cell(row=row, column=3).value
        uom = ws_ekpo.cell(row=row, column=4).value
        if e_val and e_val not in ekpo_map:
            ekpo_map[e_val] = {'qty': qty, 'uom': uom}

    print(f"  EKPO mapping built: {len(ekpo_map)} unique Ref entries")

    red_font = Font(color="9C0006", bold=True)
    qty_overwrite = 0
    uom_overwrite = 0

    for row in range(3, ws_bt.max_row + 1):
        a_val = str(ws_bt.cell(row=row, column=1).value or '').strip()
        c_val = str(ws_bt.cell(row=row, column=3).value or '').strip()
        key = a_val + c_val  # A + C 两列拼接（无加号）

        bt_qty = ws_bt.cell(row=row, column=8).value   # BT Qty
        bt_uom = ws_bt.cell(row=row, column=9).value   # BT UOM

        ekpo_data = ekpo_map.get(key, {})
        ekpo_qty = ekpo_data.get('qty')
        ekpo_uom = ekpo_data.get('uom')

        # 先把 EKPO 的值写回 BT 的 EKPO QTY / EKPO UOM（25/26列）
        ws_bt.cell(row=row, column=25, value=ekpo_qty)
        ws_bt.cell(row=row, column=26, value=ekpo_uom)

        # 对比 Qty vs EKPO QTY
        need_qty_overwrite = False
        if ekpo_qty is not None and bt_qty is not None:
            try:
                if float(bt_qty) != float(ekpo_qty):
                    need_qty_overwrite = True
            except (ValueError, TypeError):
                need_qty_overwrite = True
        elif ekpo_qty is not None or bt_qty is not None:
            need_qty_overwrite = True

        if need_qty_overwrite:
            ws_bt.cell(row=row, column=25, value=bt_qty)
            ws_bt.cell(row=row, column=25).font = red_font
            qty_overwrite += 1

        # 对比 UOM vs EKPO UOM
        need_uom_overwrite = False
        if ekpo_uom is not None and bt_uom is not None:
            if str(bt_uom).strip().upper() != str(ekpo_uom).strip().upper():
                need_uom_overwrite = True
        elif ekpo_uom is not None or bt_uom is not None:
            need_uom_overwrite = True

        if need_uom_overwrite:
            ws_bt.cell(row=row, column=26, value=bt_uom)
            ws_bt.cell(row=row, column=26).font = red_font
            uom_overwrite += 1

    print(f"  [OK] QTY overwritten {qty_overwrite} rows (red), UOM overwritten {uom_overwrite} rows (red)")
    return wb


# ============================================================
# Step 10: AMT DIFF processing —— CHANGED: REMARKS append instead of overwrite
# ============================================================
def step10_amt_diff_processing(wb):
    print("\n[Step 10] AMT DIFF processing & CDEC overwrite")

    ws_bt = wb[BT_SHEET]
    ws_entry = wb[ENTRY_NUM_SHEET]
    ws_in = wb["PROCTER & GAMBLE PHILIPPINE IN"]

    cdec_t_map = {}
    for row in range(2, ws_in.max_row + 1):
        d_val = str(ws_in.cell(row=row, column=4).value or '').strip()
        bf_val = get_val(ws_in.cell(row=row, column=58).value)
        d_clean = clean_entry_no(d_val)
        if d_clean and d_clean not in cdec_t_map:
            cdec_t_map[d_clean] = bf_val

    cdec_detail_map = {}
    for row in range(2, ws_in.max_row + 1):
        d_val = clean_entry_no(ws_in.cell(row=row, column=4).value)
        if d_val:
            dutiable = get_val(ws_in.cell(row=row, column=33).value)
            duty = get_val(ws_in.cell(row=row, column=35).value)
            vat = get_val(ws_in.cell(row=row, column=53).value)
            if d_val not in cdec_detail_map:
                cdec_detail_map[d_val] = []
            cdec_detail_map[d_val].append((dutiable, duty, vat))

    bt_data = {}
    for row in range(3, ws_bt.max_row + 1):
        l_val = get_val(ws_bt.cell(row=row, column=12).value)
        m_val = get_val(ws_bt.cell(row=row, column=13).value)
        n_val = get_val(ws_bt.cell(row=row, column=14).value)
        r_val = l_val + m_val + n_val
        d_val = clean_entry_no(ws_bt.cell(row=row, column=4).value)
        bt_data[row] = {'d': d_val, 'l': l_val, 'm': m_val, 'n': n_val, 'r': r_val}

    entry_s_sum = {}
    for row, data in bt_data.items():
        entry_s_sum[data['d']] = entry_s_sum.get(data['d'], 0) + data['r']

    diff_zero_entries = set()
    diff_nonzero_entries = set()

    for row, data in bt_data.items():
        d_val = data['d']
        t_val = cdec_t_map.get(d_val, 0)
        s_val = entry_s_sum.get(d_val, 0)
        u_val = t_val - s_val

        if abs(u_val) < 1e-9:
            diff_zero_entries.add(d_val)
        else:
            diff_nonzero_entries.add(d_val)

    print(f"  Entries with AMT DIFF = 0: {len(diff_zero_entries)}")
    print(f"  Entries with AMT DIFF != 0: {len(diff_nonzero_entries)}")

    # ===== CHANGED: append "Done" to REMARKS instead of overwriting =====
    for row in range(2, ws_entry.max_row + 1):
        entry_val = clean_entry_no(ws_entry.cell(row=row, column=1).value)
        if entry_val in diff_zero_entries:
            current_remark = str(ws_entry.cell(row, column=4).value or '').strip()
            if current_remark == '' or current_remark == 'None':
                ws_entry.cell(row=row, column=4, value='Done')
            else:
                ws_entry.cell(row=row, column=4, value=f"{current_remark}/Done")

    print(f"  [OK] Entry Num REMARKS column appended Done (AMT DIFF = 0)")

    for entry_no in diff_nonzero_entries:
        bt_rows = [r for r in range(3, ws_bt.max_row + 1)
                   if clean_entry_no(ws_bt.cell(r, 4).value) == entry_no]
        cdec_rows = cdec_detail_map.get(entry_no, [])

        if len(bt_rows) == len(cdec_rows) and len(bt_rows) > 0:
            for i, bt_row_idx in enumerate(bt_rows):
                if i < len(cdec_rows):
                    dutiable, duty, vat = cdec_rows[i]
                    ws_bt.cell(row=bt_row_idx, column=16, value=dutiable)
                    ws_bt.cell(row=bt_row_idx, column=12, value=duty)
                    ws_bt.cell(row=bt_row_idx, column=14, value=vat)

            new_s = 0
            for bt_row_idx in bt_rows:
                l = get_val(ws_bt.cell(bt_row_idx, 12).value)
                m = get_val(ws_bt.cell(bt_row_idx, 13).value)
                n = get_val(ws_bt.cell(bt_row_idx, 14).value)
                new_s += l + m + n

            t_val = cdec_t_map.get(entry_no, 0)
            new_diff = t_val - new_s

            if abs(new_diff) < 1e-9:
                for row in range(2, ws_entry.max_row + 1):
                    if clean_entry_no(ws_entry.cell(row, 1).value) == entry_no:
                        current_remark = str(ws_entry.cell(row, column=4).value or '').strip()
                        if current_remark == '' or current_remark == 'None':
                            ws_entry.cell(row=row, column=4, value='Done')
                        else:
                            ws_entry.cell(row=row, column=4, value=f"{current_remark}/Done")
                        break
                print(f"  [Done] {entry_no} AMT DIFF = 0 after CDEC overwrite")
            else:
                print(f"  [Pending] {entry_no} AMT DIFF still {new_diff}, requires PDF operation")
        else:
            print(f"  [Skip] {entry_no} BT {len(bt_rows)} rows vs CDEC {len(cdec_rows)} rows mismatch")

    return wb


# ============================================================
# Phase 1: Batch download PDFs (UNTOUCHED - DO NOT MODIFY)
# ============================================================
def download_all(browser, entry_list):
    context = browser.contexts[0]
    not_found = []

    for entry_no, broker in entry_list:
        pdf_target = PDF_DIR / f"{entry_no}.pdf"
        if pdf_target.exists():
            print(f"  [Skip] {entry_no}.pdf already exists")
            continue

        broker_u = broker.strip().upper() if broker else ""
        if broker_u not in CLIENTS:
            print(f"  [WARN] {entry_no} BROKER='{broker}' not in config")
            not_found.append(entry_no)
            continue

        url = CLIENTS[broker_u]["url"]
        page = context.new_page()
        try:
            downloaded = False
            page.goto(url, timeout=60000)
            page.wait_for_load_state("networkidle")
            time.sleep(3)

            if "login.microsoftonline.com" in page.url:
                print(f"  [ERROR] {entry_no} session expired")
                not_found.append(entry_no)
                continue

            sb = page.locator('input[placeholder*="Search" i], input[title*="Search" i]').first
            if sb.is_visible():
                sb.fill(entry_no)
                sb.press("Enter")
                page.wait_for_timeout(5000)

            exact = None
            folder = None
            for r in page.locator('div[role="row"], tr[role="row"]').all():
                try:
                    t = r.inner_text()
                except Exception:
                    continue
                if entry_no not in t:
                    continue
                if f"{entry_no}.pdf" in t:
                    exact = r
                    break
                if folder is None:
                    folder = r

            target_row = exact
            if not target_row and folder:
                folder.click()
                page.wait_for_timeout(3000)

                for r in page.locator('div[role="row"], tr[role="row"]').all():
                    try:
                        t = r.inner_text()
                    except Exception:
                        continue
                    if f"{entry_no}.pdf" in t:
                        target_row = r
                        break

                if not target_row:
                    sb2 = page.locator('input[placeholder*="Search" i]').first
                    if sb2.is_visible():
                        sb2.fill(entry_no)
                        sb2.press("Enter")
                        page.wait_for_timeout(5000)

                    for r in page.locator('div[role="row"], tr[role="row"]').all():
                        try:
                            t = r.inner_text()
                        except Exception:
                            continue
                        if f"{entry_no}.pdf" in t:
                            target_row = r
                            break

            if not target_row:
                print(f"  [PDF not found] {entry_no}")
                not_found.append(entry_no)
            else:
                target_row.click()
                page.wait_for_timeout(1500)
                for sel in ['button:has-text("Download")',
                            '[data-automationid="DownloadButton"]',
                            '[aria-label*="Download" i]']:
                    btn = page.locator(sel).first
                    if btn.is_visible():
                        try:
                            with page.expect_download(timeout=60000) as di:
                                btn.click()
                            dl = di.value
                            raw = dl.suggested_filename
                            tmp = PDF_DIR / f"_{entry_no}_tmp"
                            if raw.lower().endswith(".zip"):
                                zip_path = tmp.with_suffix(".zip")
                                dl.save_as(str(zip_path))
                                with zipfile.ZipFile(zip_path) as zf:
                                    for n in zf.namelist():
                                        if n.lower().endswith(".pdf"):
                                            with zf.open(n) as s, open(pdf_target, "wb") as t:
                                                t.write(s.read())
                            else:
                                dl.save_as(str(pdf_target))
                            for x in PDF_DIR.glob(f"_{entry_no}_tmp*"):
                                x.unlink()
                            print(f"  [OK] Downloaded {entry_no}.pdf")
                            downloaded = True
                            break
                        except Exception as e:
                            print(f"  [WARN] {entry_no} download error: {e}")
                if not downloaded:
                    not_found.append(entry_no)
        finally:
            page.close()

    return not_found


# ============================================================
# Phase 2: Read PDFs & write BT —— CHANGED: REMARKS append
# ============================================================
TOTAL_GLOBAL_TAX_PATTERNS = [
    re.compile(r"Total\s*Global\s*Taxes?[:\s]*([\d,]+\.?\d*)", re.IGNORECASE),
    re.compile(r"Tota[l1]\s*Glob[aA][lI1]\s*Taxe[s5][:\s]*([\d,]+\.?\d*)", re.IGNORECASE),
    re.compile(r"Total\s*Globa[lI1]\s*Taxe[s5][:\s]*([\d,]+\.?\d*)", re.IGNORECASE),
]

def read_total_global_tax(pdf_path):
    """UNTOUCHED - DO NOT MODIFY"""
    patterns = [
        re.compile(r"Total\s*Global\s*Taxes?[:\s]*([\d,]+\.?\d*)", re.IGNORECASE),
        re.compile(r"Tota[l1]\s*Glob[aA][lI1]\s*Taxe[s5][:\s]*([\d,]+\.?\d*)", re.IGNORECASE),
    ]

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                for pat in patterns:
                    m = pat.search(txt)
                    if m:
                        try:
                            val = float(m.group(1).replace(",", ""))
                            return val
                        except ValueError:
                            pass

            for page in pdf.pages:
                try:
                    height = page.height
                    crop_box = (0, height * 0.50, page.width, height)
                    cropped_page = page.within_bbox(crop_box)
                    img = cropped_page.to_image(resolution=200).original
                    ocr_txt = pytesseract.image_to_string(img, lang="eng")
                    for pat in patterns:
                        m = pat.search(ocr_txt)
                        if m:
                            try:
                                val = float(m.group(1).replace(",", ""))
                                return val
                            except ValueError:
                                pass
                except Exception as e:
                    print(f"  [OCR WARN] {pdf_path.name} page {page.page_number}: {e}")

    except Exception as e:
        print(f"  [WARN] {pdf_path.name} pdfplumber failed: {e}")
        try:
            import fitz
            doc = fitz.open(pdf_path)
            full_text = ""
            for page in doc:
                full_text += page.get_text()
            doc.close()

            for pat in patterns:
                m = pat.search(full_text)
                if m:
                    return float(m.group(1).replace(",", ""))
        except:
            pass

    return None

def process_pdfs_and_bt(wb, entry_list, not_found):
    """CHANGED: REMARKS append instead of overwrite; PDF prefix matching"""
    ws_bt = wb[BT_SHEET]
    ws_entry = wb[ENTRY_NUM_SHEET]
    ws_in = wb["PROCTER & GAMBLE PHILIPPINE IN"]

    cdec_t = {}
    for row in range(2, ws_in.max_row + 1):
        d = clean_entry_no(ws_in.cell(row=row, column=4).value)
        if d and d not in cdec_t:
            cdec_t[d] = get_val(ws_in.cell(row=row, column=58).value)

    for entry_no, broker in entry_list:
        if entry_no in not_found:
            continue

        # ========== 修改开始：前缀匹配 ==========
        matching_pdfs = list(PDF_DIR.glob(f"{entry_no}*.pdf"))
        if not matching_pdfs:
            # 标记 PDF not found
            for r in range(2, ws_entry.max_row + 1):
                if clean_entry_no(ws_entry.cell(r, 1).value) == entry_no:
                    current_remark = str(ws_entry.cell(r, 4).value or '').strip()
                    if current_remark == '' or current_remark == 'None':
                        ws_entry.cell(r, 4, "PDF not found")
                    else:
                        ws_entry.cell(r, 4, f"{current_remark}/PDF not found")
                    break
            continue

        if len(matching_pdfs) > 1:
            # 按修改时间降序，取最新的
            matching_pdfs.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            print(f"  [WARN] Multiple PDFs found for {entry_no}, using {matching_pdfs[0].name}")
        pdf_path = matching_pdfs[0]
        # ========== 修改结束 ==========

        tax_val = read_total_global_tax(pdf_path)
        if tax_val is None:
            for r in range(2, ws_entry.max_row + 1):
                if clean_entry_no(ws_entry.cell(r, 1).value) == entry_no:
                    current_remark = str(ws_entry.cell(r, 4).value or '').strip()
                    if current_remark == '' or current_remark == 'None':
                        ws_entry.cell(r, 4, "PDF value not found")
                    else:
                        ws_entry.cell(r, 4, f"{current_remark}/PDF value not found")
                    print(f"  [PDF value not found] {entry_no}")
                    break
            continue

        # 后续更新 BT 和 REMARKS 的逻辑保持不变
        bt_rows = [r for r in range(3, ws_bt.max_row + 1)
                   if clean_entry_no(ws_bt.cell(r, 4).value) == entry_no]
        if not bt_rows:
            continue

        for i, r in enumerate(bt_rows):
            ws_bt.cell(r, 13, tax_val if i == 0 else 0)

        s = sum(get_val(ws_bt.cell(r, 12).value) +
                get_val(ws_bt.cell(r, 13).value) +
                get_val(ws_bt.cell(r, 14).value) for r in bt_rows)
        t = cdec_t.get(entry_no, 0)
        diff = t - s

        remark = "Done" if abs(diff) < 1e-9 else "AMT DIFF not 0, pls manual check."
        for r in range(2, ws_entry.max_row + 1):
            if clean_entry_no(ws_entry.cell(r, 1).value) == entry_no:
                current_remark = str(ws_entry.cell(r, 4).value or '').strip()
                if current_remark == '' or current_remark == 'None':
                    ws_entry.cell(r, 4, remark)
                else:
                    ws_entry.cell(r, 4, f"{current_remark}/{remark}")
                break
        print(f"  [{remark}] {entry_no}  diff={diff}")

    return wb


# ============================================================
# Step 11: Mark UPLOADED = Done for AMT DIFF = 0
# ============================================================
def step11_mark_uploaded(wb):
    print("\n[Step 11] Mark UPLOADED = Done for AMT DIFF = 0")
    ws_bt = wb[BT_SHEET]

    # 找 UPLOADED 列
    uploaded_col = None
    for cell in ws_bt[2]:
        if cell.value and str(cell.value).strip().upper() == 'UPLOADED':
            uploaded_col = cell.column
            break

    if not uploaded_col:
        print("  [WARN] UPLOADED column not found, skipping")
        return wb

    # AMT DIFF 是第 21 列 (U列)
    marked = 0
    for row in range(3, ws_bt.max_row + 1):
        diff_val = ws_bt.cell(row=row, column=22).value
        try:
            if abs(float(diff_val)) < 1e-9:
                ws_bt.cell(row=row, column=uploaded_col, value='Done')
                marked += 1
        except (ValueError, TypeError):
            pass

    print(f"  [OK] UPLOADED marked Done for {marked} rows")
    return wb


# ============================================================
# Step 12: Export —— CHANGED: filename from CSV, UPLOADED=Done rows
# ============================================================
def step12_export(wb):
    print("\n[Step 12] Export")

    # 找 CSV 文件
    csv_file = None
    csv_files = [f for f in TCD_DIR.glob("*PROCTER*CSV")]
    if not csv_files:
        csv_files = [f for f in TCD_DIR.glob("*PROCTER*.*") if f.suffix.upper() == '.CSV']
    if csv_files:
        csv_file = csv_files[0]

    if not csv_file:
        print("  [ERROR] PROCTER*.csv not found, cannot determine export filename")
        return

    # 提取日期范围
    date_range = extract_date_range(csv_file)
    print(f"  Date range extracted: {date_range}")
    print(f"  Source CSV: {csv_file.name}")

    out_name = f"BT_{date_range}.xlsx"
    out_path = TCD_DIR / out_name

    ws_bt = wb[BT_SHEET]
    ws_entry = wb[ENTRY_NUM_SHEET]
    ws_in = wb["PROCTER & GAMBLE PHILIPPINE IN"]

    # 找 UPLOADED 列
    uploaded_col = None
    for cell in ws_bt[2]:
        if cell.value and str(cell.value).strip().upper() == 'UPLOADED':
            uploaded_col = cell.column
            break

    # CDEC T 值
    cdec_t = {}
    for row in range(2, ws_in.max_row + 1):
        d = clean_entry_no(ws_in.cell(row, 4).value)
        if d and d not in cdec_t:
            cdec_t[d] = get_val(ws_in.cell(row, 58).value)

    # 收集 Done 的 entry
    done_entries = set()
    for r in range(2, ws_entry.max_row + 1):
        remark = str(ws_entry.cell(r, 4).value or '').strip()
        # 检查是否包含 Done（可能在 / 分隔的任一段中）
        parts = [p.strip() for p in remark.split('/')]
        if 'Done' in parts:
            entry = clean_entry_no(ws_entry.cell(r, 1).value)
            if entry:
                done_entries.add(entry)

    print(f"  Entries marked as Done: {len(done_entries)}")

    # 导出 Done 的行
    done_rows = []
    for r in range(3, ws_bt.max_row + 1):
        entry = clean_entry_no(ws_bt.cell(r, 4).value)
        if not entry:
            continue

        # 条件1: 从 Entry Num REMARKS 判断 Done
        if entry in done_entries:
            row_vals = [ws_bt.cell(r, c).value for c in range(1, 18)]
            done_rows.append(row_vals)
            continue

        # 条件2: UPLOADED 列 = Done
        if uploaded_col:
            uploaded_val = str(ws_bt.cell(r, uploaded_col).value or '').strip()
            if uploaded_val == 'Done':
                row_vals = [ws_bt.cell(r, c).value for c in range(1, 18)]
                done_rows.append(row_vals)
                continue



    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = BT_SHEET
    headers = [ws_bt.cell(2, c).value for c in range(1, 18)]
    ws_out.append(headers)
    for row_vals in done_rows:
        ws_out.append(row_vals)

    wb_out.save(out_path)
    print(f"  [OK] Exported {len(done_rows)} rows -> {out_path}")


# ============================================================
# Main flow —— CHANGED execution order
# ============================================================
def main():
    print("=" * 55)
    print("TCD-PH v2 - Full Automation Pipeline (Patched)")
    print("=" * 55)

    PDF_DIR.mkdir(parents=True, exist_ok=True)
    os.chdir(str(TCD_DIR))

    # Step 0: Auto-start Edge
    # ensure_edge_debug()

    # ============================================================
    # Step 1-7: Template Preparation
    # ============================================================
    print("\n" + "=" * 55)
    print("PHASE: Template Preparation (Steps 1-7)")
    print("=" * 55)

    # Step 1: Clean template
    wb = load_workbook(BT_EXCEL_PATH)
    wb = step1_clean_template(wb)
    wb.save(BT_EXCEL_PATH)

    # Step 2: Process CSV
    wb = load_workbook(BT_EXCEL_PATH)
    wb, df_csv = step2_process_csv(wb)
    wb = step2b_uniform_row_height(wb)
    wb.save(BT_EXCEL_PATH)

    # Step 3: Process EKPO (filter Del. Indicator L)
    wb = load_workbook(BT_EXCEL_PATH)
    wb = step3_process_ekpo(wb)
    wb.save(BT_EXCEL_PATH)

    # Step 4: Process 3PL files
    wb = load_workbook(BT_EXCEL_PATH)
    wb = step4_process_3pl(wb)
    wb.save(BT_EXCEL_PATH)

    # Step 5: Process Entry Num sheet
    wb = load_workbook(BT_EXCEL_PATH)
    wb = step5_process_entry_num(wb, df_csv)
    wb.save(BT_EXCEL_PATH)

    # Step 6: Fill BT sheet
    wb = load_workbook(BT_EXCEL_PATH)
    wb = step6_fill_bt(wb, df_csv)
    wb.save(BT_EXCEL_PATH)

    # Step 7: Fill Company Code
    wb = load_workbook(BT_EXCEL_PATH)
    wb = step7_fill_company_code(wb)
    wb.save(BT_EXCEL_PATH)

    # ============================================================
    # Step 8: Write formulas (cols 25/26 preserved)
    # ============================================================
    print("\n" + "=" * 55)
    print("PHASE: Formulas & QTY/UOM (Steps 8-9)")
    print("=" * 55)

    wb = load_workbook(BT_EXCEL_PATH)
    wb = step8_write_formulas(wb)
    wb.save(BT_EXCEL_PATH)

    # ============================================================
    # Step 9: QTY/UOM overwrite
    # ============================================================
    wb = load_workbook(BT_EXCEL_PATH)
    wb = step9_qty_uom_overwrite(wb)
    wb.save(BT_EXCEL_PATH)

    # ============================================================
    # Step 10: AMT DIFF + PDF processing
    # ============================================================
    print("\n" + "=" * 55)
    print("PHASE: AMT DIFF Analysis & PDF (Step 10)")
    print("=" * 55)

    wb = load_workbook(BT_EXCEL_PATH)
    wb = step10_amt_diff_processing(wb)

    # Collect entries that still need PDF processing
    ws_entry = wb[ENTRY_NUM_SHEET]
    still_not_zero = []
    for r in range(2, ws_entry.max_row + 1):
        entry = clean_entry_no(ws_entry.cell(r, 1).value)
        remark = str(ws_entry.cell(r, 4).value or '').strip()
        parts = [p.strip() for p in remark.split('/')]
        broker = str(ws_entry.cell(r, 2).value or '').strip()
        if entry and 'Done' not in parts and 'NO ACTION' not in parts:
            still_not_zero.append((entry, broker))

    print(f"\nEntries requiring PDF processing: {len(still_not_zero)}")
    wb.save(BT_EXCEL_PATH)

    if still_not_zero:
        not_found = []
        # print("\n" + "=" * 55)
        # print("PHASE: Download PDFs")
        # print("=" * 55)
        #
        # with sync_playwright() as p:
        #     browser = p.chromium.connect_over_cdp(CDP_URL)
        #     not_found = download_all(browser, still_not_zero)
        #     browser.close()
        #
        # # Mark not_found (append REMARKS)
        # wb = load_workbook(BT_EXCEL_PATH)
        # ws_entry = wb[ENTRY_NUM_SHEET]
        # for entry_no in not_found:
        #     for r in range(2, ws_entry.max_row + 1):
        #         if clean_entry_no(ws_entry.cell(r, 1).value) == entry_no:
        #             current_remark = str(ws_entry.cell(r, 4).value or '').strip()
        #             if current_remark == '' or current_remark == 'None':
        #                 ws_entry.cell(r, 4, "PDF not found")
        #             else:
        #                 ws_entry.cell(r, 4, f"{current_remark}/PDF not found")
        #             break
        # wb.save(BT_EXCEL_PATH)

        print("\n" + "=" * 55)
        print("PHASE: Read PDFs & Update BT")
        print("=" * 55)

        wb = load_workbook(BT_EXCEL_PATH)
        wb = process_pdfs_and_bt(wb, still_not_zero, not_found)
        wb.save(BT_EXCEL_PATH)

    # ============================================================
    # Step 11: Mark UPLOADED = Done
    # ============================================================
    print("\n" + "=" * 55)
    print("PHASE: Mark UPLOADED (Step 11)")
    print("=" * 55)

    wb = load_workbook(BT_EXCEL_PATH)
    wb = step11_mark_uploaded(wb)
    wb.save(BT_EXCEL_PATH)

    # ============================================================
    # Step 12: Export
    # ============================================================
    print("\n" + "=" * 55)
    print("PHASE: Export (Step 12)")
    print("=" * 55)

    wb = load_workbook(BT_EXCEL_PATH)
    step12_export(wb)

    print("\n" + "=" * 55)
    print("All done!")
    print("=" * 55)
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()
