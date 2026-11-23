
import os, re, time, sys, zipfile, subprocess, shutil
from pathlib import Path
from openpyxl import load_workbook
import pdfplumber, pytesseract
from playwright.sync_api import sync_playwright
import pandas as pd
import urllib.request
from concurrent.futures import ThreadPoolExecutor, TimeoutError

def ensure_edge_debug():
    try:
        urllib.request.urlopen("http://localhost:9222/json/version", timeout=2)
        print("[OK] Edge debug port already open")
        return
    except:
        pass
    subprocess.run(["taskkill", "/f", "/im", "msedge.exe"], capture_output=True, shell=True)
    time.sleep(2)
    edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    profile_dir = r"C:\EdgeDebugProfile"
    cmd = [edge_path, "--remote-debugging-port=9222", f"--user-data-dir={profile_dir}",
           "--remote-allow-origins=*", "--no-first-run",
           "https://pgone.sharepoint.com"]
    subprocess.Popen(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(5)
    print("Please login to SharePoint in the browser window")
    input("Press Enter after login...")

TCD_DIR = Path.home() / "Desktop" / "TCD AGING"
BT_EXCEL_PATH = TCD_DIR / "BT TEMPLATE ACCRUAL.xlsx"
ACC_EXCEL_PATH = TCD_DIR / "BT TEMPLATE ACCRUAL.xlsx"
AGING_EXCEL_PATH = TCD_DIR / "TCD Aging Accrual Report TEMPLATE.xlsx"
DOCS_DIR = TCD_DIR / "Docs"
CDP_URL = "http://127.0.0.1:9222"
tess_path = TCD_DIR / "Tesseract-OCR" / "tesseract.exe"
if tess_path.exists():
    pytesseract.pytesseract.tesseract_cmd = str(tess_path)

CLIENTS = {
    "ATI": {"url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FATI%2FFY2526%20Raw%20Shipping%20Documents"},
    "CEVA": {"url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?csf=1&web=1&id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FCEVA"},
    "TRANSWORLD": {"url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?csf=1&web=1&id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FTW%20(NEW)"},
    "MAERSK": {"url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?e=bzBDoH&sharingv2=true&fromShare=true&at=9&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNjA5MTcxMjIyMiIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D&CID=161445a2%2Da04b%2Ff000%2Dcbbf%2D47299236efcc&cidOR=SPO&FolderCTID=0x012000CCD2F3217E34DD4781511DEE93F8125E&id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FMAERSK%2FDocument%20Retention"},
    "YUSEN": {"url": "https://pgone.sharepoint.com/sites/OneSupplyServicesManila/FAP%20AMA%20East%20Operations/Forms/AllItems.aspx?csf=1&web=1&id=%2Fsites%2FOneSupplyServicesManila%2FFAP%20AMA%20East%20Operations%2FRetention%2FVault%20and%20Retention%2FPhilippines%2FShipping%20Documents%2FBROKERS%2FYusen"}
}
AGING_SHEETS = ["A. 0-30", "B. 31-60", "D. 91-120", "G. OVER 180"]

def clean_entry_no(val):
    if val is None: return ''
    s = str(val).strip()
    if s.lower() == 'nan' or s == 'N/A' or s == '': return ''
    s = s.replace('-','').replace('_','').replace(' ','').upper()
    s = re.sub(r'^[^A-Za-z0-9]+','',s); s = re.sub(r'[^A-Za-z0-9]+$','',s)
    s = re.sub(r'\.0$','',s)
    return s

def get_val(v):
    if v is None: return 0.0
    try: return float(v)
    except: return 0.0

def find_col(ws, header_name, header_rows=(1,2)):
    target = re.sub(r'\s+', '', header_name.strip()).lower()
    for hr in header_rows:
        for c in range(1, ws.max_column + 1):
            hdr = ws.cell(hr, c).value
            if hdr:
                cell_str = re.sub(r'\s+', '', str(hdr).strip()).lower()
                if target in cell_str or cell_str in target:
                    return c

    for hr in header_rows:
        cols_preview = [repr(ws.cell(hr,c).value) for c in range(1, min(ws.max_column+1, 19))]
        print(f"  Row {hr}: {cols_preview}")

    return None

def download_entry_folder(browser, entry_no, broker):
    context = browser.contexts[0]
    broker_u = (broker or "").strip().upper()
    if broker_u not in CLIENTS:
        print(f"  [WARN] {entry_no} broker '{broker}' not in config"); return False
    cfg = CLIENTS[broker_u]; url = cfg["url"]; keyword = cfg.get("keyword","")
    entry_dir = DOCS_DIR / entry_no
    if entry_dir.exists():
        print(f"  [Skip] {entry_no} folder already exists"); return True
    page = context.new_page()
    try:
        page.goto(url, timeout=60000); page.wait_for_load_state("networkidle"); time.sleep(3)
        if "login.microsoftonline.com" in page.url:
            print(f"  [Session expired] {entry_no}"); return False
        sb = page.locator('input[placeholder*="Search" i], input[title*="Search" i]').first
        if sb.is_visible():
            search_term = f"{keyword} {entry_no}" if keyword else entry_no
            sb.fill(search_term); sb.press("Enter"); page.wait_for_timeout(5000)
        folder_row = None
        for r in page.locator('div[role="row"], tr[role="row"]').all():
            try: t = r.inner_text()
            except: continue
            if entry_no not in t: continue
            if re.search(r"\.(pdf|xlsx?|zip|docx?)$", t, re.I): continue
            folder_row = r; break
        if not folder_row:
            print(f"  [Folder not found] {entry_no}"); return False
        folder_row.click(); page.wait_for_timeout(3000)
        try:
            file_items = page.locator('div[role="row"], tr[role="row"]').all()
            has_content = False
            for item in file_items:
                try: text = item.inner_text()
                except: continue
                if text and text.strip() and "Name" not in text and "Modified" not in text:
                    has_content = True; break
            if not has_content:
                print(f"  [Folder empty] {entry_no}"); return False
        except: pass
        downloaded = False
        for sel in ['button:has-text("Download")','[data-automationid="DownloadButton"]','[aria-label*="Download" i]']:
            btn = page.locator(sel).first
            if not btn.is_visible(): continue
            try:
                zip_path = DOCS_DIR / f"{entry_no}.zip"
                if zip_path.exists(): zip_path.unlink()
                with page.expect_download(timeout=180000) as di: btn.click()
                dl = di.value
                err = dl.failure()
                if err: print(f"  [Download failed] {entry_no}: {err}"); break
                dl.save_as(str(zip_path))
                if not zip_path.exists() or zip_path.stat().st_size == 0:
                    print(f"  [Empty file] {entry_no}")
                    if zip_path.exists(): zip_path.unlink(); break
                if not zipfile.is_zipfile(zip_path): print(f"  [Not a zip] {entry_no}"); break
                entry_dir = DOCS_DIR / entry_no; entry_dir.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(zip_path,'r') as zf:
                    temp_dir = DOCS_DIR / f"_{entry_no}_temp"; temp_dir.mkdir(parents=True, exist_ok=True)
                    zf.extractall(str(temp_dir))
                    for item in temp_dir.rglob("*"):
                        if item.is_file():
                            dest = entry_dir / item.relative_to(temp_dir).name
                            if dest.exists(): dest.unlink()
                            item.rename(dest)
                    shutil.rmtree(temp_dir)
                print(f"  [OK] Downloaded & extracted {entry_no}"); downloaded = True; break
            except Exception as e:
                print(f"  [Download error] {entry_no}: {e}"); break
        return downloaded
    finally:
        page.close()

def find_pdf_in_folder(entry_dir, keyword):
    if not entry_dir.exists(): return None
    for f in entry_dir.rglob("*.pdf"):
        if keyword.lower() in f.stem.lower(): return f
    return None

def ocr_with_timeout(img, lang='eng', config='', timeout=30):
    with ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(pytesseract.image_to_string, img, lang=lang, config=config)
        try:
            return future.result(timeout=timeout)
        except TimeoutError:
            print(f"  [WARN] OCR timed out after {timeout}s")
            return ""

def read_total_global_tax(pdf_path):
    patterns = [
        re.compile(r"Total\s*Global\s*Taxes?[:\s]*([\d,]+\.?\d*)", re.I),
        re.compile(r"Tota[l1]\s*Glob[aA][lI1]\s*Taxe[s5][:\s]*([\d,]+\.?\d*)", re.I),
    ]
    header_pattern = re.compile(r"Total\s+Global\s+Taxes?", re.I)
    num_pattern = re.compile(r"([\d,]+\.?\d{0,2})")
    def extract_tax_from_next_line(text):
        lines = text.split('\n')
        for i,line in enumerate(lines):
            if header_pattern.search(line):
                nums = num_pattern.findall(line)
                if nums:
                    vals = [float(n.replace(',','')) for n in nums if float(n.replace(',',''))>0]
                    if vals: return vals[-1]
                for j in range(i+1, min(i+4,len(lines))):
                    nl = lines[j].strip()
                    if not nl: continue
                    nums = num_pattern.findall(nl)
                    if nums:
                        vals = [float(n.replace(',','')) for n in nums if float(n.replace(',',''))>0]
                        if vals: return vals[-1]
                return None
        return None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ''.join((page.extract_text() or '')+'\n' for page in pdf.pages)
            for pat in patterns:
                m = pat.search(full_text)
                if m:
                    val = float(m.group(1).replace(',',''))
                    if val>0: print(f"  [TAX] {pdf_path.name} => {val}"); return val
            val = extract_tax_from_next_line(full_text)
            if val is not None: print(f"  [TAX] {pdf_path.name} => {val}"); return val
    except Exception as e: print(f"  [WARN] pdfplumber failed: {e}")
    try:
        from PIL import ImageEnhance
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                img = page.to_image(resolution=300).original
                enhancer = ImageEnhance.Contrast(img); img_a = enhancer.enhance(2.0)
                ocr_a = ocr_with_timeout(img_a, lang='eng')
                ocr_b = ocr_with_timeout(img, lang='eng', config='--psm 6')
                img_gray = img.convert('L'); img_bin = img_gray.point(lambda x:0 if x<128 else 255,'1')
                ocr_c = ocr_with_timeout(img_bin, lang='eng', config='--psm 11')
                for ocr_text in [ocr_a, ocr_b, ocr_c]:
                    if not ocr_text.strip(): continue
                    for pat in patterns:
                        m = pat.search(ocr_text)
                        if m:
                            val = float(m.group(1).replace(',',''))
                            if val>0: print(f"  [TAX via OCR] {pdf_path.name} => {val}"); return val
                    val = extract_tax_from_next_line(ocr_text)
                    if val is not None: print(f"  [TAX via OCR] {pdf_path.name} => {val}"); return val
    except Exception as e: print(f"  [WARN] OCR failed: {e}")
    try:
        import fitz
        doc = fitz.open(pdf_path); full_text = ''.join(page.get_text() for page in doc); doc.close()
        for pat in patterns:
            m = pat.search(full_text)
            if m:
                val = float(m.group(1).replace(',',''))
                if val>0: print(f"  [TAX via PyMuPDF] {pdf_path.name} => {val}"); return val
        val = extract_tax_from_next_line(full_text)
        if val is not None: print(f"  [TAX via PyMuPDF] {pdf_path.name} => {val}"); return val
    except Exception as e: print(f"  [WARN] PyMuPDF failed: {e}")
    print(f"  [FAIL] No Total Global Taxes found in {pdf_path.name}"); return None

def read_total_global_tax_from_folder(entry_no):
    entry_dir = DOCS_DIR / entry_no
    if not entry_dir.exists(): print(f"  [No folder] {entry_no}"); return None
    pdf_path = find_pdf_in_folder(entry_dir, "Customs Declaration Sheet")
    if pdf_path: print(f"  Reading: {pdf_path.name}"); return read_total_global_tax(pdf_path)
    pdf_path = find_pdf_in_folder(entry_dir, "_SAD")
    if pdf_path: print(f"  Reading: {pdf_path.name}"); return read_total_global_tax(pdf_path)
    print(f"  [No SAD PDF] in {entry_dir}"); return None

def get_bt_acc_rows(entry_no):
    if not ACC_EXCEL_PATH.exists(): return []
    wb = load_workbook(ACC_EXCEL_PATH, read_only=True)
    if "BT ACC" not in wb.sheetnames: wb.close(); return []
    ws = wb["BT ACC"]
    id_col = find_col(ws, "Import Declaration Number")
    if not id_col: wb.close(); return []
    rows = [r for r in range(2, ws.max_row+1) if clean_entry_no(ws.cell(r,id_col).value)==entry_no]
    wb.close(); return rows

def get_ekpo_count(entry_no):
    wb = load_workbook(BT_EXCEL_PATH, read_only=True)
    if "EKPO" not in wb.sheetnames: wb.close(); return None
    ws = wb["EKPO"]; id_col = find_col(ws,"Import Declaration Number") or 1
    count = sum(1 for r in range(2, ws.max_row+1) if clean_entry_no(ws.cell(r,id_col).value)==entry_no)
    wb.close(); return count

def mark_lines_mismatch(entry_no):
    if not AGING_EXCEL_PATH.exists(): return
    wb = load_workbook(AGING_EXCEL_PATH); marked=0
    for sheet in AGING_SHEETS:
        if sheet not in wb.sheetnames: continue
        ws = wb[sheet]
        id_col = find_col(ws,"Import Declaration Number")
        if not id_col: continue
        status_col = find_col(ws,"4PL STATUS") or ws.max_column+1
        remark_col = find_col(ws,"4PL Remarks") or status_col+1
        ws.cell(1,status_col,"4PL STATUS"); ws.cell(1,remark_col,"4PL Remarks")
        for r in range(2, ws.max_row+1):
            if clean_entry_no(ws.cell(r,id_col).value)==entry_no:
                ws.cell(r,status_col,"pending"); ws.cell(r,remark_col,"Lines mismatch with ekpo"); marked+=1
    wb.save(AGING_EXCEL_PATH); print(f"  [Marked] {entry_no} -> {marked} rows pending in aging report")

def update_entry_remark(ws_entry, entry_no, new_remark_part):
    updated=False
    for r in range(2, ws_entry.max_row+1):
        if clean_entry_no(ws_entry.cell(r,1).value)==entry_no:
            cur = str(ws_entry.cell(r,4).value or '').strip()
            if cur=='' or cur=='None': ws_entry.cell(r,4,new_remark_part)
            else:
                parts = [p.strip() for p in cur.split('/') if p.strip()]
                if new_remark_part not in parts: parts.append(new_remark_part)
                ws_entry.cell(r,4,'/'.join(parts))
            updated=True
    if not updated: print(f"  [WARN] No row found for {entry_no} when updating Remark")

def update_sad_updated(ws_entry, entry_no, value):
    sad_col = find_col(ws_entry,"SAD Updated")
    if sad_col is None: sad_col = ws_entry.max_column+1; ws_entry.cell(1,sad_col,"SAD Updated")
    updated_any=False
    for r in range(2, ws_entry.max_row+1):
        if clean_entry_no(ws_entry.cell(r,1).value)==entry_no:
            cur = str(ws_entry.cell(r,sad_col).value or '').strip()
            if cur=='' or cur=='None': ws_entry.cell(r,sad_col,value)
            else:
                parts = [p.strip() for p in cur.split('/') if p.strip()]
                if value not in parts: parts.append(value)
                ws_entry.cell(r,sad_col,'/'.join(parts))
            updated_any=True
    if not updated_any: print(f"  [WARN] No row found for {entry_no} when updating SAD Updated")

def load_ekpo_data():
    ekpo_files = list(TCD_DIR.glob("*EKPO*.*"))
    if not ekpo_files: print("  [ERROR] No EKPO file found"); return {}
    ekpo_file = ekpo_files[0]; print(f"  Loading EKPO data from: {ekpo_file.name}")
    df=None
    for enc in ['utf-16','utf-8','cp1252']:
        try: df=pd.read_csv(ekpo_file,sep='\t',encoding=enc,skiprows=3,dtype=str); break
        except: continue
    if df is None: print("  [ERROR] Failed to read EKPO file"); return {}
    print(f"  EKPO raw rows: {len(df)}, columns: {len(df.columns)}")
    del_col=None
    for c in df.columns:
        if str(c).strip()=='Del. Indicator': del_col=c; break
    if del_col: print(f"  Del. Indicator column found: '{del_col}'")
    else: print("  [WARN] Del. Indicator column NOT found, no filtering applied")
    purch_col=None; mat_col=None
    for c in df.columns:
        cs=str(c).strip()
        if 'Purchasing Doc' in cs: purch_col=c
        elif cs=='Material': mat_col=c
    if not purch_col or not mat_col: print(f"  [ERROR] Required columns not found"); return {}
    ekpo_data={}; skipped=0
    for _,row in df.iterrows():
        purch=str(row[purch_col]).strip(); mat=str(row[mat_col]).strip()
        if not purch or purch=='nan': continue
        if del_col:
            del_val=str(row[del_col]).strip().upper()
            if 'L' in del_val: skipped+=1; continue
        ekpo_data.setdefault(purch,[]).append((purch,mat))
    total=sum(len(v) for v in ekpo_data.values())
    print(f"  EKPO data loaded: {total} rows after filtering")
    if skipped: print(f"  Skipped {skipped} rows due to Del. Indicator containing 'L'")
    return ekpo_data

def get_entry_po_list(ws_entry, entry_no):
    po_col = find_col(ws_entry,"PO")
    if po_col is None:
        for c in range(1, ws_entry.max_column+1):
            if 'po' in str(ws_entry.cell(1,c).value or '').strip().lower(): po_col=c; break
    if po_col is None: print(f"  [ERROR] Cannot find PO column"); return []
    po_set=set()
    for r in range(2, ws_entry.max_row+1):
        if clean_entry_no(ws_entry.cell(r,1).value)==entry_no:
            v=str(ws_entry.cell(r,po_col).value or '').strip()
            if v and v!='nan': po_set.add(v)
    return list(po_set)

def update_entry_remarks_all_rows(ws_entry, entry_no, remark):
    updated=False
    for r in range(2, ws_entry.max_row+1):
        if clean_entry_no(ws_entry.cell(r,1).value)==entry_no:
            cur=str(ws_entry.cell(r,4).value or '').strip()
            if cur=='' or cur=='None': ws_entry.cell(r,4,remark)
            else:
                parts=[p.strip() for p in cur.split('/') if p.strip()]
                if remark not in parts: parts.append(remark)
                ws_entry.cell(r,4,'/'.join(parts))
            updated=True
    if not updated: print(f"  [WARN] No row found for {entry_no} when updating Remarks (all rows)")


def check_lines_and_update(entry_no, ws_entry, ws_bt, ekpo_data_dict):
    """
    返回 True 表示 lines matched，False 表示 mismatched
    """
    po_list = get_entry_po_list(ws_entry, entry_no)
    if not po_list:
        print(f"  [WARN] No PO found for {entry_no}")
        return False

    ekpo_rows = []
    for po in po_list:
        if po in ekpo_data_dict:
            ekpo_rows.extend(ekpo_data_dict[po])

    ekpo_count = len(ekpo_rows)

    id_col = find_col(ws_bt, "Import Declaration Number")
    if id_col is None:
        print(f"  [ERROR] Cannot find ID col in BT ACC")
        return False

    bt_acc_rows = [r for r in range(3, ws_bt.max_row + 1)
                   if clean_entry_no(ws_bt.cell(r, id_col).value) == entry_no]
    bt_acc_count = len(bt_acc_rows)

    print(f"  Line check: BT ACC rows={bt_acc_count}, EKPO rows={ekpo_count}")

    purch_doc_col = find_col(ws_bt, "Purchasing Document Number") or ws_bt.max_column + 1
    material_col = find_col(ws_bt, "Material") or ws_bt.max_column + 1
    ws_bt.cell(1, purch_doc_col, "Purchasing Document Number")
    ws_bt.cell(1, material_col, "Material")

    if ekpo_count == bt_acc_count:
        for i, r in enumerate(bt_acc_rows):
            if i < len(ekpo_rows):
                ws_bt.cell(r, purch_doc_col, ekpo_rows[i][0])  # 采购凭证号保持原样（如需去零可加 int()）
                ws_bt.cell(r, material_col, int(ekpo_rows[i][1]))  # ★ 物料号去前导零
        update_entry_remarks_all_rows(ws_entry, entry_no, "Lines matched")
        print(f"  [Lines matched] {entry_no} -> data written")
        return True
    else:
        update_entry_remarks_all_rows(ws_entry, entry_no, "Lines mismatched")
        mark_lines_mismatch(entry_no)
        print(f"  [Lines mismatched] {entry_no}")
        return False

def write_done_to_aging(entry_no):
    """将 'done' 写入 Aging Report 对应 sheet 的 4PL STATUS 列"""
    if not AGING_EXCEL_PATH.exists():
        print(f"  [WARN] Aging file not found: {AGING_EXCEL_PATH}")
        return
    wb = load_workbook(AGING_EXCEL_PATH)
    written = 0
    for sheet_name in AGING_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        id_col = find_col(ws, "Import Declaration Number")
        if id_col is None:
            continue
        status_col = find_col(ws, "4PL STATUS")
        if status_col is None:
            status_col = ws.max_column + 1
            ws.cell(1, status_col, "4PL STATUS")
        for r in range(2, ws.max_row + 1):
            if clean_entry_no(ws.cell(r, id_col).value) == entry_no:
                ws.cell(r, status_col, "done")
                written += 1
    if written > 0:
        wb.save(AGING_EXCEL_PATH)
        print(f"  [Writeback] {entry_no} -> wrote 'done' to {written} row(s) in aging report")
    else:
        print(f"  [Writeback] {entry_no} -> no matching rows found in aging sheets")
    wb.close()


def append_ekpo_sheet(wb):
    print("\n--- Appending EKPO sheet ---")
    ekpo_files = list(TCD_DIR.glob("*EKPO*.*"))
    if not ekpo_files:
        print("  [ERROR] No EKPO file found for appending")
        return
    ekpo_file = ekpo_files[0]
    print(f"  Reading EKPO file: {ekpo_file.name}")

    df = None
    for enc in ['utf-16', 'utf-8', 'cp1252']:
        try:
            df = pd.read_csv(ekpo_file, sep='\t', encoding=enc, skiprows=3, dtype=str)
            break
        except:
            continue
    if df is None:
        print("  [ERROR] Cannot read EKPO file")
        return

    del_col = None
    for c in df.columns:
        if str(c).strip() == 'Del. Indicator':
            del_col = c
            break

    purch_col = None
    mat_col = None
    qty_col = None
    unit_col = None
    for c in df.columns:
        cs = str(c).strip()
        if 'Purchasing Doc' in cs:
            purch_col = c
        elif cs == 'Material':
            mat_col = c
        elif cs == 'PO Quantity':
            qty_col = c
        elif cs == 'Order Unit':
            unit_col = c

    if not all([purch_col, mat_col]):
        print(f"  [ERROR] Missing required columns")
        return

    records = []
    for _, row in df.iterrows():
        purch = str(row[purch_col]).strip()
        if not purch or purch == 'nan':
            continue
        if del_col and 'L' in str(row[del_col]).strip().upper():
            continue
        mat = str(row[mat_col]).strip() if mat_col else ''
        qty = str(row[qty_col]).strip() if qty_col else ''
        unit = str(row[unit_col]).strip() if unit_col else ''
        records.append((purch, mat, qty, unit, f"{purch}{mat}"))

    if not records:
        print("  [INFO] No new records to append")
        return

    if "EKPO" not in wb.sheetnames:
        ws = wb.create_sheet("EKPO")
        for ci, h in enumerate(['Purchasing Doc.', 'Material', 'PO Quantity', 'Order Unit', 'Ref'], 1):
            ws.cell(1, ci, h)
        next_row = 2
    else:
        ws = wb["EKPO"]
        next_row = ws.max_row + 1
        if ws.cell(1, 1).value is None:
            for ci, h in enumerate(['Purchasing Doc.', 'Material', 'PO Quantity', 'Order Unit', 'Ref'], 1):
                ws.cell(1, ci, h)
            next_row = 2

    for i, (p, m, q, u, r) in enumerate(records):
        row = next_row + i
        ws.cell(row, 1, p)  # 采购凭证号保持原样
        ws.cell(row, 2, int(m))  # ★ 物料号去前导零
        ws.cell(row, 3, q)
        ws.cell(row, 4, u)
        ws.cell(row, 5, r)

    print(f"  [OK] Appended {len(records)} rows to EKPO sheet (starting at row {next_row})")

def fmt_time(seconds):
    m, s = divmod(seconds, 60)
    return f"{int(m):02d}:{s:06.3f}"

def main():
    t_total_start = time.time()
    ensure_edge_debug()
    print("="*60); print("Phase 3: Two‑phase + writeback done to aging"); print("="*60)
    DOCS_DIR.mkdir(parents=True, exist_ok=True); os.chdir(str(TCD_DIR))

    wb = load_workbook(BT_EXCEL_PATH); ws_entry = wb['Entry Num']
    entry_list=[]; seen=set()
    for r in range(2, ws_entry.max_row+1):
        entry=clean_entry_no(ws_entry.cell(r,1).value); broker=str(ws_entry.cell(r,2).value or '').strip()
        remark=str(ws_entry.cell(r,4).value or '').strip()
        if entry and 'Done' not in remark and 'NO ACTION' not in remark and entry not in seen:
            seen.add(entry); entry_list.append((entry,broker))
    wb.close()
    print(f"\nEntries to process: {len(entry_list)}")
    if not entry_list: print("Nothing to do."); return

    # PHASE 1: Download
    print("\n========== PHASE 1: Download all folders ==========")
    t_phase1_start = time.time()
    playwright = sync_playwright().start()
    browser = playwright.chromium.connect_over_cdp(CDP_URL)
    print("Browser connected.")
    try:
        for entry_no, broker in entry_list:
            entry_dir = DOCS_DIR / entry_no
            if entry_dir.exists():
                print(f"  [Skip existing] {entry_no}")
                continue
            print(f"\n--- Downloading {entry_no} ---")
            ok = download_entry_folder(browser, entry_no, broker)
            if not ok:
                print(f"  [Download failed] {entry_no}")
    finally:
        browser.close()
        playwright.stop()
    t_phase1_end = time.time()
    print(f"\n[PHASE 1 DONE] Duration: {fmt_time(t_phase1_end-t_phase1_start)}")

    # PHASE 2: Process
    print("\n========== PHASE 2: Process all entries ==========")
    t_phase2_start = time.time()
    wb = load_workbook(BT_EXCEL_PATH)
    ws_bt = wb['BT ACC']
    ws_entry = wb['Entry Num']
    ws_in = wb["PROCTER & GAMBLE PHILIPPINE IN"]

    bt_index = {}
    for r in range(3, ws_bt.max_row+1):
        val = clean_entry_no(ws_bt.cell(r,4).value)
        if val: bt_index.setdefault(val,[]).append(r)
    cdec_t = {}
    for row in range(2, ws_in.max_row+1):
        d = clean_entry_no(ws_in.cell(row,4).value)
        if d and d not in cdec_t: cdec_t[d] = get_val(ws_in.cell(row,58).value)

    ekpo_data_dict = load_ekpo_data()

    for entry_no, broker in entry_list:
        print(f"\n--- Processing {entry_no} ---")
        entry_dir = DOCS_DIR / entry_no
        if not entry_dir.exists():
            print(f"  [No folder] {entry_no}, skip reading PDF")
            if ekpo_data_dict:
                lines_matched = check_lines_and_update(entry_no, ws_entry, ws_bt, ekpo_data_dict)
                # 如果 lines matched 但没有 SAD，不写 done
            continue

        tax_val = read_total_global_tax_from_folder(entry_no)
        bt_rows = bt_index.get(entry_no, [])
        if not bt_rows:
            print(f"  [SKIP] No BT ACC rows found for {entry_no}")
            if ekpo_data_dict: check_lines_and_update(entry_no, ws_entry, ws_bt, ekpo_data_dict)
            continue
        print(f"  BT ACC rows: {bt_rows}")
        s = sum(get_val(ws_bt.cell(r,12).value)+get_val(ws_bt.cell(r,13).value)+get_val(ws_bt.cell(r,14).value) for r in bt_rows)
        t = cdec_t.get(entry_no,0); diff = t-s
        if tax_val is not None:
            for i,r in enumerate(bt_rows): ws_bt.cell(r,13,tax_val if i==0 else 0)
            print(f"  [Wrote tax={tax_val}] {entry_no}")
            s_new = sum(get_val(ws_bt.cell(r,12).value)+get_val(ws_bt.cell(r,13).value)+get_val(ws_bt.cell(r,14).value) for r in bt_rows)
            diff_new = t-s_new
            sad_done = abs(diff_new) < 1e-9
            if sad_done:
                update_sad_updated(ws_entry,entry_no,"Done")
                print(f"  [Done] {entry_no}  new_diff={diff_new}")
            else:
                update_sad_updated(ws_entry,entry_no,"AMT DIFF not 0")
                print(f"  [AMT DIFF not 0] {entry_no}  new_diff={diff_new}")
        else:
            print(f"  [Customs Declaration Sheet not found] {entry_no}  diff={diff}")
            sad_done = False  # 没有 SAD 则不算 Done

        # 执行 lines 检查，并获取是否 matched
        lines_matched = False
        if ekpo_data_dict:
            lines_matched = check_lines_and_update(entry_no, ws_entry, ws_bt, ekpo_data_dict)

        # 如果同时满足 SAD Done 且 Lines matched，写回 aging report
        if sad_done and lines_matched:
            write_done_to_aging(entry_no)

    append_ekpo_sheet(wb)
    wb.save(BT_EXCEL_PATH)
    t_phase2_end = time.time()
    print(f"\n[PHASE 2 DONE] Duration: {fmt_time(t_phase2_end-t_phase2_start)}")
    print(f"\n[DONE] Phase 3 complete. Total duration: {fmt_time(t_phase2_end-t_total_start)}")

if __name__ == "__main__":
    main()