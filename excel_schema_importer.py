# -*- coding: utf-8 -*-
"""
excel_schema_importer.py
ម៉ូឌុលឆ្លាតវៃសម្រាប់សិក្សា និងស្រង់ទិន្នន័យដោយស្វ័យប្រវត្តិតាមរចនាសម្ព័ន្ធឯកសារ
«បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» និងឯកសារខាងក្រៅគ្រប់ប្រភេទ (Excel, Word, PDF, រូបភាព OCR, CSV)
"""

import io
import re
import sqlite3
from datetime import datetime, date
import openpyxl
import pandas as pd
import numpy as np

# Mappings ឈ្មោះសាលារៀនពីអក្សរឡាតាំងទៅជាភាសាខ្មែរ
SCHOOL_NAME_MAP = {
    'sleng': 'ស្លែងស្ពាន',
    'phomdey': 'ភ្នំដី',
    'chronieng': 'ច្រនៀង',
    'chomkar': 'ចំការចេក',
    'thlok': 'ថ្លុក',
    'phomsrie': 'ភ្នំស្រី',
    'dongkaor': 'ដង្កោ',
    'sala': 'សាលា',
    'dunsok': 'ដូនសោក',
    'rmeat': 'រមៀត',
    'char': 'ចារ'
}

def clean_school_name(sheet_name_or_val):
    """បម្លែងឈ្មោះសន្លឹក ឬឈ្មោះសាលាអោយត្រឹមត្រូវជាភាសាខ្មែរ"""
    if not sheet_name_or_val:
        return ""
    raw = str(sheet_name_or_val).strip()
    clean = raw.lower().replace('r-', '').replace('r_', '').strip()
    if clean in SCHOOL_NAME_MAP:
        return SCHOOL_NAME_MAP[clean]
    # Check if already Khmer
    return raw

def format_date_str(val):
    """បម្លែងកាលបរិច្ឆេទទៅជា YYYY-MM-DD"""
    if val is None or val == "":
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    m = re.search(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return s[:10]

def parse_excel_workbook(file_bytes_or_path, db_conn=None):
    """
    សិក្សា និងទាញយករចនាសម្ព័ន្ធសន្លឹកកិច្ចការទាំងអស់ក្នុងឯកសារ Excel (.xlsm, .xlsx, .xls)
    កំណត់សម្គាល់ប្រភេទសន្លឹកដោយស្វ័យប្រវត្តិ៖
    - daily_matrix: កំណត់ត្រាប្រចាំថ្ងៃតាមសាលា (Sleng, PhomDey...)
    - monthly_claim: សំណើទូទាត់ប្រចាំខែ (R-Sleng, R-PhomDey...)
    - farmer_purchases: បញ្ជីទិញទំនិញពីកសិករ/ជំពាក់ (Feed, Viget...)
    - price_catalog: តារាងតម្លៃទំនិញ (តម្លៃទំនិញ...)
    """
    if isinstance(file_bytes_or_path, (bytes, bytearray)):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes_or_path), data_only=True)
    else:
        wb = openpyxl.load_workbook(file_bytes_or_path, data_only=True)

    results = {
        "file_type": "excel_workbook",
        "sheet_names": wb.sheetnames,
        "sheets": {}
    }

    # ទាញយកតម្លៃទំនិញជាមុនសម្រាប់ប្រើប្រាស់បូកសរុប (ពី SQLite, សន្លឹកតម្លៃទំនិញ, និង R-* sheets)
    product_prices = {}

    # ក. ពី Database SQLite
    try:
        conn_local = db_conn or sqlite3.connect("school_pos.db")
        c = conn_local.cursor()
        rows = c.execute("SELECT name, phase1_price, phase2_price FROM products").fetchall()
        for r in rows:
            if r[0]:
                product_prices[r[0]] = (r[1] or 0.0, r[2] or 0.0)
        if not db_conn:
            conn_local.close()
    except Exception:
        pass

    # ខ. ពីសន្លឹក «តម្លៃទំនិញ» ក្នុង Workbook ផ្ទាល់
    for s in wb.sheetnames:
        if 'តម្លៃទំនិញ' in s or 'price' in s.lower():
            p_ws = wb[s]
            for r in range(4, p_ws.max_row + 1):
                p_name = p_ws.cell(r, 2).value
                p_p1 = p_ws.cell(r, 3).value
                p_p2 = p_ws.cell(r, 4).value
                if p_name and str(p_name).strip():
                    try:
                        p1_val = float(p_p1 or 0)
                        p2_val = float(p_p2 or 0)
                        product_prices[str(p_name).strip()] = (p1_val, p2_val)
                    except Exception:
                        pass

    # គ. ពីសន្លឹកសំណើទូទាត់ R-* (ដើម្បីបានតម្លៃជាក់លាក់តាមសាលា)
    for s in wb.sheetnames:
        if s.lower().startswith('r-'):
            r_ws = wb[s]
            for r in range(11, min(30, r_ws.max_row + 1)):
                r_name = r_ws.cell(r, 2).value
                r_up = r_ws.cell(r, 5).value
                if r_name and r_up:
                    try:
                        up_val = float(r_up)
                        if up_val > 0:
                            product_prices[str(r_name).strip()] = (up_val, up_val)
                    except Exception:
                        pass

    for s_name in wb.sheetnames:
        ws = wb[s_name]
        lower_s = s_name.lower().strip()

        # ១. សំណើទូទាត់ប្រចាំខែ (Monthly Claim Proposal - R-*)
        if lower_s.startswith('r-') or 'សំណើ' in s_name or 'ទូទាត់' in s_name:
            claim_info = parse_monthly_claim_sheet(ws, s_name)
            if claim_info:
                results["sheets"][s_name] = claim_info
            continue

        # ២. បញ្ជីទិញពីកសិករ/ជំពាក់ (Feed, Viget)
        if lower_s in ['feed', 'viget'] or 'កសិករ' in s_name:
            farmer_info = parse_farmer_sheet(ws, s_name)
            if farmer_info:
                results["sheets"][s_name] = farmer_info
            continue

        # ៣. តារាងតម្លៃទំនិញ (តម្លៃទំនិញ)
        if 'តម្លៃទំនិញ' in s_name or 'price' in lower_s:
            catalog_info = parse_price_catalog_sheet(ws, s_name)
            if catalog_info:
                results["sheets"][s_name] = catalog_info
            continue

        # ៤. កំណត់ត្រាប្រចាំថ្ងៃតាមសាលា (School Daily Matrix)
        # ពិនិត្យមើលវត្តមានក្បាលជួរដេក ថ្ងៃដែលបានទិញដាក់ ឬ លេខវិក័យប័ត្រ ឬឈ្មោះសាលា
        is_matrix = False
        r2_vals = [str(ws.cell(2, c).value or '') for c in range(1, 15)]
        if any('ថ្ងៃដែលបានទិញ' in v for v in r2_vals) or any('លេខវិក័យប័ត្រ' in v for v in r2_vals):
            is_matrix = True
        elif lower_s in SCHOOL_NAME_MAP:
            is_matrix = True

        if is_matrix:
            matrix_info = parse_school_daily_matrix_sheet(ws, s_name, product_prices)
            if matrix_info:
                results["sheets"][s_name] = matrix_info
            continue

    return results

def parse_school_daily_matrix_sheet(ws, sheet_name, product_prices=None):
    """
    ស្រង់ទិន្នន័យពីសន្លឹកតារាងម៉ាទ្រីសប្រចាំថ្ងៃ (Sleng, PhomDey, Chronieng...)
    """
    clean_sch = clean_school_name(sheet_name)

    # ស្វែងរកមុខទំនិញនៅជួរដេកទី ២៩ (Row 29)
    items = []
    for c in range(4, ws.max_column + 1):
        val = ws.cell(row=29, column=c).value
        if val is not None:
            s_val = str(val).strip()
            # មិនយកកូដ 0, None, ថ្ងៃខែ, សរុប
            if s_val and s_val not in ['0', 'None', 'ថ្ងៃខែ', 'សរុប'] and not s_val.startswith('202'):
                items.append((c, s_val))

    # ប្រសិនបើ Row 29 គ្មានទំនិញ ស្វែងរកនៅ Row 2 ឬ 3
    if not items:
        for c in range(4, ws.max_column + 1):
            val = ws.cell(row=2, column=c).value or ws.cell(row=3, column=c).value
            if val is not None:
                s_val = str(val).strip()
                if s_val and s_val not in ['0', 'None', '']:
                    items.append((c, s_val))

    if not items:
        return None

    # លេខយោងចុងខែមុន
    ref_last_month = ws.cell(row=30, column=3).value or ""

    flattened_rows = []
    matrix_rows = []

    # ចាប់ផ្ដើមពីជួរដេក ៣១ ដល់ជួរដេកចុងក្រោយ
    for r in range(31, ws.max_row + 1):
        d_buy = ws.cell(row=r, column=1).value
        d_eat = ws.cell(row=r, column=2).value
        v_no = ws.cell(row=r, column=3).value

        if d_buy is None and v_no is None:
            continue

        d_buy_str = format_date_str(d_buy)
        d_eat_str = format_date_str(d_eat) if d_eat else ""
        v_no_str = str(v_no).strip() if v_no is not None else ""

        # ទិន្នន័យសម្រាប់ Matrix Grid
        row_grid = {
            "កាលបរិច្ឆេទដាក់": d_buy_str,
            "កាលបរិច្ឆេទហូប": d_eat_str,
            "លេខវិក័យប័ត្រ": v_no_str
        }

        has_item_data = False
        row_daily_total = 0.0

        for c, item_name in items:
            q = ws.cell(row=r, column=c).value
            q_val = 0.0
            if q is not None and str(q).strip() not in ['', '0', '0.0']:
                try:
                    q_val = float(q)
                except Exception:
                    q_val = 0.0

            row_grid[item_name] = q_val if q_val > 0 else 0.0

            if q_val > 0:
                has_item_data = True
                # តម្លៃឯកតាពី catalog
                u_p = 0.0
                if product_prices and item_name in product_prices:
                    u_p = product_prices[item_name][0] or product_prices[item_name][1] or 0.0

                t_p = q_val * u_p
                row_daily_total += t_p

                flattened_rows.append({
                    "កាលបរិច្ឆេទ": d_buy_str,
                    "ថ្ងៃហូប": d_eat_str,
                    "សាលារៀន": clean_sch,
                    "លេខវិក័យប័ត្រ": v_no_str,
                    "មុខទំនិញ": item_name,
                    "បរិមាណ": q_val,
                    "តម្លៃរាយ (៛)": u_p,
                    "សរុប (៛)": t_p
                })

        if has_item_data or d_buy_str:
            row_grid["សរុបប្រចាំថ្ងៃ (៛)"] = row_daily_total
            matrix_rows.append(row_grid)

    df_flat = pd.DataFrame(flattened_rows) if flattened_rows else pd.DataFrame(
        columns=["កាលបរិច្ឆេទ", "ថ្ងៃហូប", "សាលារៀន", "លេខវិក័យប័ត្រ", "មុខទំនិញ", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)"]
    )
    df_matrix = pd.DataFrame(matrix_rows) if matrix_rows else pd.DataFrame()

    return {
        "sheet_type": "daily_matrix",
        "sheet_name": sheet_name,
        "school_name": clean_sch,
        "title_kh": f"📅 កំណត់ត្រាប្រចាំថ្ងៃសាលា {clean_sch}",
        "ref_last_month": ref_last_month,
        "items_count": len(items),
        "items_list": [it[1] for it in items],
        "df": df_flat,
        "df_matrix": df_matrix,
        "total_records": len(flattened_rows)
    }

def parse_monthly_claim_sheet(ws, sheet_name):
    """
    ស្រង់ទិន្នន័យពីសន្លឹកសំណើទូទាត់ប្រចាំខែ (R-Sleng, R-PhomDey...)
    """
    clean_sch = clean_school_name(sheet_name)

    dist = ws.cell(3, 3).value or "ស្រីស្នំ"
    comm = ws.cell(4, 3).value or "ស្លែងស្ពាន"
    sch = ws.cell(5, 3).value or clean_sch
    v_no = ws.cell(5, 7).value or ""

    d_start = format_date_str(ws.cell(7, 1).value)
    d_end = format_date_str(ws.cell(7, 5).value)

    sup = ws.cell(8, 3).value or "សាត ក្រូត"
    sup_addr = ws.cell(8, 5).value or "ភូមិខ្វែក ឃុំមោង"
    sup_phone = ws.cell(8, 7).value or "090 854 133"

    items = []
    tot_exact = 0.0
    tot_round = 0.0

    for r in range(11, ws.max_row + 1):
        c1 = ws.cell(r, 1).value
        c1_str = str(c1 or '').strip()

        # ប្រសិនបើជួបជួរដេកសរុបទឹកប្រាក់
        if 'សរុបទឹកប្រាក់' in c1_str:
            for col_idx in [6, 5, 7, 4]:
                val = ws.cell(r, col_idx).value
                if val is not None and isinstance(val, (int, float)):
                    tot_exact = float(val)
                    break
            continue

        # ប្រសិនបើជួបជួរដេកថវិកាសរុបស្នើសុំទូទាត់ (បង្គត់លេខ)
        if 'ថវិកាសរុប' in c1_str or 'បង្គត់លេខ' in c1_str:
            for col_idx in [5, 6, 4, 7]:
                val = ws.cell(r, col_idx).value
                if val is not None and isinstance(val, (int, float)):
                    tot_round = float(val)
                    break
            continue

        # បញ្ឈប់ប្រសិនបើដល់សម្គាល់ ឬ ហត្ថលេខា
        if 'សម្គាល់' in c1_str or 'បានឃើញ' in c1_str:
            break

        # ទាញយកទំនិញ
        name = ws.cell(r, 2).value
        v_ref = ws.cell(r, 3).value
        qty = ws.cell(r, 4).value
        u_price = ws.cell(r, 5).value
        total = ws.cell(r, 6).value

        if name and str(name).strip() and qty is not None:
            s_name = str(name).strip()
            if s_name not in ['0', 'None', '', 'បរិយាយមុខទំនិញ'] and not s_name.startswith('('):
                try:
                    q_val = float(qty)
                    u_val = float(u_price or 0)
                    t_val = float(total or (q_val * u_val))
                except Exception:
                    continue

                if q_val > 0 or u_val > 0:
                    items.append({
                        "លរ": c1 if c1 is not None and isinstance(c1, int) else len(items) + 1,
                        "មុខទំនិញ": s_name,
                        "លេខយោងបង្កាន់ដៃ": str(v_ref or "").strip(),
                        "បរិមាណ": q_val,
                        "តម្លៃរាយ (៛)": u_val,
                        "សរុប (៛)": t_val,
                        "សាលារៀន": clean_sch
                    })

    if tot_exact == 0.0 and items:
        tot_exact = sum(i["សរុប (៛)"] for i in items)
    if tot_round == 0.0 and tot_exact > 0:
        val = int(round(tot_exact))
        rem = val % 100
        base = (val // 100) * 100
        tot_round = base if rem < 50 else base + 100

    df_items = pd.DataFrame(items) if items else pd.DataFrame(
        columns=["លរ", "មុខទំនិញ", "លេខយោងបង្កាន់ដៃ", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)", "សាលារៀន"]
    )

    return {
        "sheet_type": "monthly_claim",
        "sheet_name": sheet_name,
        "school_name": clean_sch,
        "title_kh": f"📑 សំណើទូទាត់ប្រចាំខែសាលា {clean_sch}",
        "district": str(dist).strip(),
        "commune": str(comm).strip(),
        "voucher_no": str(v_no).strip(),
        "d_start": d_start,
        "d_end": d_end,
        "supplier_name": str(sup).strip(),
        "supplier_address": str(sup_addr).strip(),
        "supplier_phone": str(sup_phone).strip(),
        "total_exact": float(tot_exact or 0),
        "total_rounded": float(tot_round or 0),
        "df": df_items,
        "total_items": len(items)
    }



def parse_farmer_sheet(ws, sheet_name):
    """
    ស្រង់ទិន្នន័យពីសន្លឹកទិញពីកសិករ (Feed, Viget)
    """
    rows = []
    for r in range(4, ws.max_row + 1):
        d_val = ws.cell(r, 1).value
        farmer = ws.cell(r, 2).value
        item = ws.cell(r, 3).value
        qty = ws.cell(r, 4).value
        u_price = ws.cell(r, 5).value
        total = ws.cell(r, 6).value
        note = ws.cell(r, 7).value
        pay_date = ws.cell(r, 8).value
        paid = ws.cell(r, 9).value
        remain = ws.cell(r, 10).value

        if farmer and item and qty is not None:
            try:
                q_val = float(qty)
                u_val = float(u_price or 0)
                tot_val = float(total or (q_val * u_val))
            except Exception:
                continue

            rows.append({
                "កាលបរិច្ឆេទ": format_date_str(d_val),
                "ឈ្មោះកសិករ/អ្នកលក់": str(farmer).strip(),
                "មុខទំនិញ": str(item).strip(),
                "បរិមាណ": q_val,
                "តម្លៃរាយ (៛)": u_val,
                "សរុប (៛)": tot_val,
                "ចំណាំ": str(note or "ជំពាក់").strip(),
                "ថ្ងៃទូទាត់": format_date_str(pay_date),
                "ប្រាក់បានទូទាត់": float(paid or 0),
                "ប្រាក់នៅសល់": float(remain or tot_val)
            })

    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["កាលបរិច្ឆេទ", "ឈ្មោះកសិករ/អ្នកលក់", "មុខទំនិញ", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)", "ចំណាំ", "ថ្ងៃទូទាត់", "ប្រាក់បានទូទាត់", "ប្រាក់នៅសល់"]
    )

    return {
        "sheet_type": "farmer_purchases",
        "sheet_name": sheet_name,
        "title_kh": f"🛒 បញ្ជីទិញពីកសិករ ({sheet_name})",
        "df": df,
        "total_records": len(rows)
    }

def parse_price_catalog_sheet(ws, sheet_name):
    """
    ស្រង់ទិន្នន័យពីសន្លឹកតម្លៃទំនិញ (តម្លៃទំនិញ)
    """
    rows = []
    for r in range(4, ws.max_row + 1):
        num = ws.cell(r, 1).value
        name = ws.cell(r, 2).value
        p1 = ws.cell(r, 3).value
        p2 = ws.cell(r, 4).value

        if name and str(name).strip():
            try:
                p1_val = float(p1 or 0)
                p2_val = float(p2 or 0)
            except Exception:
                p1_val, p2_val = 0.0, 0.0

            rows.append({
                "លរ": num if num is not None else len(rows) + 1,
                "មុខទំនិញ": str(name).strip(),
                "តម្លៃវគ្គ១ (៛)": p1_val,
                "តម្លៃវគ្គ២ (៛)": p2_val
            })

    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["លរ", "មុខទំនិញ", "តម្លៃវគ្គ១ (៛)", "តម្លៃវគ្គ២ (៛)"]
    )

    return {
        "sheet_type": "price_catalog",
        "sheet_name": sheet_name,
        "title_kh": f"📦 តារាងតម្លៃទំនិញ ({sheet_name})",
        "df": df,
        "total_records": len(rows)
    }

def parse_external_file(uploaded_file, db_conn=None):
    """
    អនុគមន៍ស្វែងរក និងស្រង់ទិន្នន័យឆ្លាតវៃពីឯកសារគ្រប់ប្រភេទ (Excel, Word, PDF, រូបភាព OCR, CSV)
    """
    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    # ១. Excel Workbooks (.xlsx, .xlsm, .xls)
    if filename.endswith(('.xlsm', '.xlsx', '.xls')):
        return parse_excel_workbook(file_bytes, db_conn=db_conn)

    # ២. PDF Documents (.pdf)
    elif filename.endswith('.pdf'):
        return parse_pdf_document(file_bytes, uploaded_file.name)

    # ៣. Word Documents (.docx, .doc)
    elif filename.endswith(('.docx', '.doc')):
        return parse_word_document(file_bytes, uploaded_file.name)

    # ៤. Image OCR (.png, .jpg, .jpeg, .webp, .bmp)
    elif filename.endswith(('.png', '.jpg', '.jpeg', '.webp', '.bmp')):
        return parse_image_ocr(file_bytes, uploaded_file.name)

    # ៥. CSV (.csv)
    elif filename.endswith('.csv'):
        return parse_csv_file(file_bytes, uploaded_file.name)

    return {"file_type": "unknown", "sheets": {}}

def parse_pdf_document(file_bytes, filename):
    """ស្រង់ទិន្នន័យតារាង និងអត្ថបទពី PDF"""
    import pdfplumber
    all_rows = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                tbls = page.extract_tables()
                for tbl in tbls:
                    for r in tbl:
                        cleaned = [(c.strip() if c else "") for c in r]
                        if any(cleaned):
                            all_rows.append(cleaned)
                if not tbls:
                    txt = page.extract_text()
                    if txt:
                        for line in txt.split("\n"):
                            parts = [p.strip() for p in re.split(r'[,|\t]', line) if p.strip()]
                            if parts:
                                all_rows.append(parts)

        df = standardize_imported_table(all_rows)
        return {
            "file_type": "pdf",
            "sheet_names": ["ទិន្នន័យ PDF"],
            "sheets": {
                "ទិន្នន័យ PDF": {
                    "sheet_type": "generic_table",
                    "title_kh": f"📄 ឯកសារ PDF: {filename}",
                    "df": df,
                    "total_records": len(df)
                }
            }
        }
    except Exception as e:
        return {"file_type": "pdf", "error": str(e), "sheets": {}}

def parse_word_document(file_bytes, filename):
    """ស្រង់ទិន្នន័យតារាងពីឯកសារ Word (.docx, .doc)"""
    from docx import Document
    all_rows = []
    try:
        doc = Document(io.BytesIO(file_bytes))
        for tbl in doc.tables:
            for row in tbl.rows:
                r_data = [cell.text.strip() for cell in row.cells]
                if any(r_data):
                    all_rows.append(r_data)
        if not all_rows:
            lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for line in lines:
                parts = [p.strip() for p in re.split(r'[,|\t]', line) if p.strip()]
                if parts:
                    all_rows.append(parts)

        df = standardize_imported_table(all_rows)
        return {
            "file_type": "word",
            "sheet_names": ["ទិន្នន័យ Word"],
            "sheets": {
                "ទិន្នន័យ Word": {
                    "sheet_type": "generic_table",
                    "title_kh": f"📝 ឯកសារ Word: {filename}",
                    "df": df,
                    "total_records": len(df)
                }
            }
        }
    except Exception as e:
        return {"file_type": "word", "error": str(e), "sheets": {}}

def parse_image_ocr(file_bytes, filename):
    """ស្រង់ទិន្នន័យពីវិក្កយបត្រ ឬតារាងស្កេនរូបភាពតាម RapidOCR"""
    from PIL import Image
    from rapidocr_onnxruntime import RapidOCR
    try:
        img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        img_np = np.array(img)
        ocr = RapidOCR()
        res, _ = ocr(img_np)
        all_rows = []
        if res:
            items = []
            for box, text, score in res:
                y_center = (box[0][1] + box[2][1]) / 2.0
                x_left = box[0][0]
                items.append({"text": text.strip(), "x": x_left, "y": y_center})

            items.sort(key=lambda i: i["y"])
            line_groups = []
            curr_group = []
            curr_y = None

            for it in items:
                if curr_y is None or abs(it["y"] - curr_y) < 18:
                    curr_group.append(it)
                    curr_y = it["y"]
                else:
                    curr_group.sort(key=lambda i: i["x"])
                    line_groups.append([i["text"] for i in curr_group])
                    curr_group = [it]
                    curr_y = it["y"]

            if curr_group:
                curr_group.sort(key=lambda i: i["x"])
                line_groups.append([i["text"] for i in curr_group])

            all_rows = line_groups

        df = standardize_imported_table(all_rows)
        return {
            "file_type": "image_ocr",
            "sheet_names": ["ទិន្នន័យស្កេនរូបភាព"],
            "sheets": {
                "ទិន្នន័យស្កេនរូបភាព": {
                    "sheet_type": "generic_table",
                    "title_kh": f"📷 ស្កេនរូបភាព (OCR): {filename}",
                    "df": df,
                    "total_records": len(df)
                }
            }
        }
    except Exception as e:
        return {"file_type": "image_ocr", "error": str(e), "sheets": {}}

def parse_csv_file(file_bytes, filename):
    """ស្រង់ទិន្នន័យពី CSV"""
    for enc in ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']:
        try:
            df_raw = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
            df = standardize_imported_dataframe(df_raw)
            return {
                "file_type": "csv",
                "sheet_names": ["ទិន្នន័យ CSV"],
                "sheets": {
                    "ទិន្នន័យ CSV": {
                        "sheet_type": "generic_table",
                        "title_kh": f"📊 ឯកសារ CSV: {filename}",
                        "df": df,
                        "total_records": len(df)
                    }
                }
            }
        except Exception:
            continue
    return {"file_type": "csv", "sheets": {}}

def standardize_imported_table(raw_rows):
    """
    បម្លែងជួរដេកឆៅពី PDF/Word/OCR អោយក្លាយជា DataFrame ស្តង់ដារ
    """
    if not raw_rows:
        return pd.DataFrame(columns=["កាលបរិច្ឆេទ", "សាលារៀន", "មុខទំនិញ", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)", "លេខវិក័យប័ត្រ"])

    max_c = max(len(r) for r in raw_rows)
    padded = [r + [""] * (max_c - len(r)) for r in raw_rows]
    df_raw = pd.DataFrame(padded[1:], columns=padded[0]) if len(padded) > 1 else pd.DataFrame(padded)
    return standardize_imported_dataframe(df_raw)

def standardize_imported_dataframe(df_raw):
    """
    កំណត់សម្គាល់ជួរឈរសំខាន់ៗ និងគណនាតម្លៃដោយស្វ័យប្រវត្តិ
    """
    if df_raw.empty:
        return pd.DataFrame(columns=["កាលបរិច្ឆេទ", "សាលារៀន", "មុខទំនិញ", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)", "លេខវិក័យប័ត្រ"])

    col_map = {}
    for c in df_raw.columns:
        s = str(c).lower().strip()
        if any(k in s for k in ['date', 'ថ្ងៃ', 'កាលបរិច្ឆេទ']):
            col_map[c] = "កាលបរិច្ឆេទ"
        elif any(k in s for k in ['school', 'សាលា']):
            col_map[c] = "សាលារៀន"
        elif any(k in s for k in ['item', 'មុខទំនិញ', 'បរិយាយ', 'ទំនិញ']):
            col_map[c] = "មុខទំនិញ"
        elif any(k in s for k in ['qty', 'បរិមាណ', 'ចំនួន']):
            col_map[c] = "បរិមាណ"
        elif any(k in s for k in ['price', 'unit', 'តម្លៃរាយ', 'តម្លៃឯកតា']):
            col_map[c] = "តម្លៃរាយ (៛)"
        elif any(k in s for k in ['total', 'សរុប', 'ទឹកប្រាក់']):
            col_map[c] = "សរុប (៛)"
        elif any(k in s for k in ['voucher', 'លេខវិក័យ', 'សក្ខីប័ត្រ']):
            col_map[c] = "លេខវិក័យប័ត្រ"

    df_renamed = df_raw.rename(columns=col_map)
    std_cols = ["កាលបរិច្ឆេទ", "សាលារៀន", "មុខទំនិញ", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)", "លេខវិក័យប័ត្រ"]

    for col in std_cols:
        if col not in df_renamed.columns:
            df_renamed[col] = "" if col in ["កាលបរិច្ឆេទ", "សាលារៀន", "មុខទំនិញ", "លេខវិក័យប័ត្រ"] else 0.0

    # បម្លែងប្រភេទលេខ និងគណនាតម្លៃ
    for col in ["បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)"]:
        df_renamed[col] = pd.to_numeric(df_renamed[col].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0.0)

    # បើគ្មានសរុប គណនាពី បរិមាណ * តម្លៃរាយ
    mask_zero_total = (df_renamed["សរុប (៛)"] == 0) & (df_renamed["បរិមាណ"] > 0) & (df_renamed["តម្លៃរាយ (៛)"] > 0)
    df_renamed.loc[mask_zero_total, "សរុប (៛)"] = df_renamed.loc[mask_zero_total, "បរិមាណ"] * df_renamed.loc[mask_zero_total, "តម្លៃរាយ (៛)"]

    return df_renamed[std_cols]

def import_records_to_database(df_records, db_path="school_pos.db", target_table="daily_records"):
    """
    ចម្លងទិន្នន័យដែលបានផ្ទៀងផ្ទាត់ និងកែសម្រួលរួចចូលក្នុង Database ដោយផ្ទាល់
    """
    if df_records is None or df_records.empty:
        return 0, "គ្មានទិន្នន័យសម្រាប់ចម្លងចូលទេ"

    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    imported_count = 0

    try:
        if target_table == "daily_records":
            for _, r in df_records.iterrows():
                d_val = str(r.get("កាលបរិច្ឆេទ", "")).strip()
                s_name = str(r.get("សាលារៀន", "")).strip()
                i_name = str(r.get("មុខទំនិញ", "")).strip()
                q_val = float(r.get("បរិមាណ", 0) or 0)
                u_val = float(r.get("តម្លៃរាយ (៛)", 0) or 0)
                tot_val = float(r.get("សរុប (៛)", 0) or (q_val * u_val))
                v_no = str(r.get("លេខវិក័យប័ត្រ", "")).strip()
                eat_d = str(r.get("ថ្ងៃហូប", "")).strip()

                if not d_val or not s_name or not i_name or q_val <= 0:
                    continue

                # ពិនិត្យមើលវត្តមានត្រួតគ្នា (Duplicate check)
                existing = c.execute(
                    "SELECT id FROM daily_records WHERE date=? AND school_name=? AND item_name=?",
                    (d_val, s_name, i_name)
                ).fetchone()

                if existing:
                    c.execute(
                        """
                        UPDATE daily_records 
                        SET quantity=?, unit_price=?, total_price=?, voucher_no=?, consumption_date=?
                        WHERE id=?
                        """,
                        (q_val, u_val, tot_val, v_no, eat_d, existing[0])
                    )
                else:
                    c.execute(
                        """
                        INSERT INTO daily_records (date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no, consumption_date)
                        VALUES (?,?,?,?,?,?,?,?,?)
                        """,
                        (d_val, s_name, i_name, "វគ្គ១", q_val, u_val, tot_val, v_no, eat_d)
                    )

                # Don't auto-insert school into schools table
                imported_count += 1

        elif target_table == "purchases":
            for _, r in df_records.iterrows():
                d_val = str(r.get("កាលបរិច្ឆេទ", "")).strip()
                farmer = str(r.get("ឈ្មោះកសិករ/អ្នកលក់", r.get("អ្នកផ្គត់ផ្គង់", ""))).strip()
                i_name = str(r.get("មុខទំនិញ", "")).strip()
                q_val = float(r.get("បរិមាណ", 0) or 0)
                u_val = float(r.get("តម្លៃរាយ (៛)", 0) or 0)
                tot_val = float(r.get("សរុប (៛)", 0) or (q_val * u_val))
                note = str(r.get("ចំណាំ", "ជំពាក់")).strip()

                if not d_val or not i_name or q_val <= 0:
                    continue

                status = "បានទូទាត់" if "ទូទាត់" in note and "ជំពាក់" not in note else "ជំពាក់"

                c.execute(
                    """
                    INSERT INTO purchases (date, item_name, unit_price, quantity, total_price, supplier_name, status, note)
                    VALUES (?,?,?,?,?,?,?,?)
                    """,
                    (d_val, i_name, u_val, q_val, tot_val, farmer, status, note)
                )
                imported_count += 1

        conn.commit()
        return imported_count, f"បានចម្លងទិន្នន័យជោគជ័យចំនួន {imported_count} ជួរដេក!"
    except Exception as e:
        conn.rollback()
        return 0, f"កំហុសក្នុងការចម្លងទិន្នន័យ: {e}"
    finally:
        conn.close()


def render_school_daily_matrix_tab(conn, cursor, get_districts_fn, get_communes_fn, get_schools_by_commune_fn, get_all_schools_fn, get_products_map_fn, format_riel_fn, to_excel_fn, parse_date_safe_fn, user_scope=None):
    """
    បង្ហាញផ្ទាំង «តារាងបញ្ជាទិញប្រចាំថ្ងៃតាមសាលា» (School Daily Matrix)
    គំរូដូចក្នុងឯកសារ «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» (Sleng, PhomDey, Chronieng, Chomkar...)
    """
    import os
    import streamlit as st

    st.subheader("📅 តារាងបញ្ជាទិញប្រចាំថ្ងៃតាមសាលា (School Daily Matrix)")
    st.info(
        "💡 ទម្រង់ដូចក្នុងឯកសារ «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» (សន្លឹក Sleng, PhomDey, Chronieng, Chomkar, Thlok...) | "
        "ជួរដេក = ថ្ងៃដាក់/ថ្ងៃហូប/លេខវិក័យប័ត្រ | ជួរឈរ = មុខទំនិញស្បៀង | អាចកែសម្រួលផ្ទាល់លើតារាង គណនាតម្លៃសរុបប្រចាំថ្ងៃ និងសរុបប្រចាំខែស្វ័យប្រវត្តិ"
    )

    is_admin = True
    user_prov, user_dist, user_comm, user_sch, user_loc_code = None, None, None, None, None
    if user_scope:
        is_admin, user_prov, user_dist, user_comm, user_sch, user_loc_code = user_scope

    # ជួរជ្រើសរើសទីតាំង សាលា ខែ និងឆ្នាំ
    mc1, mc2, mc3, mc4 = st.columns([1.2, 1.2, 1.4, 1.2])
    with mc1:
        d_list = get_districts_fn()
        if not is_admin and user_dist:
            m_dist = st.selectbox("ក្រុង/ស្រុក", [user_dist], key="mat_d_sel")
        else:
            m_dist = st.selectbox("ក្រុង/ស្រុក", ["-- ទាំងអស់ --"] + d_list, key="mat_d_sel")
    with mc2:
        if not is_admin and user_comm:
            m_comm = st.selectbox("ឃុំ/សង្កាត់", [user_comm], key="mat_c_sel")
        else:
            c_list = get_communes_fn(district=m_dist if m_dist != "-- ទាំងអស់ --" else None)
            m_comm = st.selectbox("ឃុំ/សង្កាត់", ["-- ទាំងអស់ --"] + c_list, key="mat_c_sel")
    with mc3:
        if not is_admin:
            if user_sch:
                m_schools = [user_sch]
            elif user_comm:
                m_schools = get_schools_by_commune_fn(user_comm)
            else:
                m_schools = get_all_schools_fn()
        elif m_comm != "-- ទាំងអស់ --":
            m_schools = get_schools_by_commune_fn(m_comm)
        else:
            m_schools = get_all_schools_fn()

        m_school = st.selectbox("🏫 សាលាបឋមសិក្សា", m_schools if m_schools else ["គ្មានសាលា"], key="mat_s_sel")
    with mc4:
        m_c_yr, m_c_mo = st.columns(2)
        with m_c_yr:
            m_year = st.selectbox("ឆ្នាំ", [2026, 2025, 2024], index=0, key="mat_y_sel")
        with m_c_mo:
            m_month = st.selectbox("ខែ", [f"{m:02d}" for m in range(1, 13)], index=2, key="mat_m_sel")

    target_prefix = f"{m_year}-{m_month}"

    # ទាញយកទិន្នន័យពី Database
    cur_records = cursor.execute(
        """
        SELECT date, consumption_date, voucher_no, item_name, quantity, unit_price, total_price 
        FROM daily_records 
        WHERE school_name=? AND date LIKE ?
        ORDER BY date ASC, voucher_no ASC
        """,
        (m_school, f"{target_prefix}%")
    ).fetchall()

    temp_key = f"temp_matrix_{m_school}_{target_prefix}"

    # ប្រសិនបើមិនទាន់មានទិន្នន័យក្នុង Database ផ្តល់ជម្រើសទាញពីឯកសារ Excel_2026 ឬ បង្កើតគំរូទទេ
    if not cur_records and temp_key not in st.session_state:
        st.warning(f"⚠️ មិនទាន់មានទិន្នន័យប្រចាំថ្ងៃសម្រាប់សាលា **{m_school}** ក្នុងខែ **{target_prefix}** ក្នុង Database នៅឡើយទេ។")
        col_init1, col_init2 = st.columns(2)
        with col_init1:
            if os.path.exists("បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm"):
                if st.button(f"📥 ផ្ទុកទិន្នន័យពីឯកសារ Excel_2026 សម្រាប់សាលា {m_school}", key="btn_load_sch_xlsm"):
                    wb_data = parse_excel_workbook("បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm", db_conn=conn)
                    target_sheet = None
                    for s_k, s_v in wb_data["sheets"].items():
                        if s_v.get("sheet_type") == "daily_matrix" and s_v.get("school_name") == m_school:
                            target_sheet = s_v
                            break
                    if target_sheet and not target_sheet["df"].empty:
                        cnt, msg = import_records_to_database(target_sheet["df"], target_table="daily_records")
                        st.success(f"🎉 បានផ្ទុកទិន្នន័យពីឯកសារ Excel ចំនួន {cnt} កំណត់ត្រាសម្រាប់សាលា {m_school}!")
                        st.rerun()
                    else:
                        st.info(f"ពុំមានសន្លឹកកិច្ចការផ្គូផ្គងនឹងសាលា {m_school} ក្នុងឯកសារ Excel ទេ។")
        with col_init2:
            if st.button("📋 បង្កើតទម្រង់តារាងទទេ (31 ថ្ងៃ) សម្រាប់ខែនេះ", key="btn_gen_empty_mat"):
                prod_map = get_products_map_fn(school_name=m_school)
                sample_items = list(prod_map.keys())[:15] if prod_map else ["ប្រេងឆា", "អំបិលអ៊ីយូត", "សាច់ជ្រូក៣ជាន់", "ស៊ុតទា", "ត្រីរស់", "ត្រកួន", "ស្ពៃចង្កឹះ"]
                init_rows = []
                for d_idx in range(1, 32):
                    d_str = f"{m_year}-{m_month}-{d_idx:02d}"
                    try:
                        date(int(m_year), int(m_month), d_idx)
                        eat_str = f"{m_year}-{m_month}-{min(31, d_idx+1):02d}"
                        v_str = f"{d_idx:03d}"
                        row_dict = {"កាលបរិច្ឆេទដាក់": d_str, "ថ្ងៃហូប": eat_str, "លេខវិក័យប័ត្រ": v_str}
                        for it in sample_items:
                            row_dict[it] = 0.0
                        row_dict["សរុបប្រចាំថ្ងៃ (៛)"] = 0.0
                        init_rows.append(row_dict)
                    except ValueError:
                        pass
                st.session_state[temp_key] = pd.DataFrame(init_rows)
                st.rerun()

    # រៀបចំទម្រង់ DataFrame Matrix
    matrix_df = None
    if temp_key in st.session_state:
        matrix_df = st.session_state[temp_key]
    elif cur_records:
        df_raw = pd.DataFrame(cur_records, columns=["date", "consumption_date", "voucher_no", "item_name", "quantity", "unit_price", "total_price"])
        pivot_q = df_raw.pivot_table(index=["date", "consumption_date", "voucher_no"], columns="item_name", values="quantity", aggfunc="sum").fillna(0.0)
        pivot_tot = df_raw.groupby(["date", "consumption_date", "voucher_no"])["total_price"].sum()
        pivot_df = pivot_q.reset_index()
        pivot_df = pivot_df.rename(columns={"date": "កាលបរិច្ឆេទដាក់", "consumption_date": "ថ្ងៃហូប", "voucher_no": "លេខវិក័យប័ត្រ"})
        pivot_df["សរុបប្រចាំថ្ងៃ (៛)"] = pivot_tot.values
        matrix_df = pivot_df

    if matrix_df is not None and not matrix_df.empty:
        st.markdown(f"##### 📋 តារាងបញ្ជាទិញប្រចាំថ្ងៃ - សាលាបឋមសិក្សា **{m_school}** ({target_prefix})")
        st.caption("✏️ លោកអ្នកអាចកែសម្រួលចំនួនបរិមាណ (គ.ក/គ្រាប់) ដោយផ្ទាល់លើក្រឡាតារាងខាងក្រោម៖")

        edited_mat = st.data_editor(
            matrix_df,
            use_container_width=True,
            num_rows="dynamic",
            key=f"ed_mat_{m_school}_{target_prefix}"
        )

        prod_map = get_products_map_fn(school_name=m_school)
        item_cols = [c for c in edited_mat.columns if c not in ["កាលបរិច្ឆេទដាក់", "ថ្ងៃហូប", "លេខវិក័យប័ត្រ", "សរុបប្រចាំថ្ងៃ (៛)"]]

        calc_daily_totals = []
        for _, row in edited_mat.iterrows():
            row_tot = 0.0
            for it in item_cols:
                q_val = float(row.get(it, 0) or 0)
                u_val = float(prod_map.get(it, {}).get("p1", 0) or 0)
                row_tot += q_val * u_val
            calc_daily_totals.append(row_tot)
        edited_mat["សរុបប្រចាំថ្ងៃ (៛)"] = calc_daily_totals

        total_days = len(edited_mat[edited_mat["សរុបប្រចាំថ្ងៃ (៛)"] > 0])
        total_month_budget = sum(calc_daily_totals)
        m1, m2, m3 = st.columns(3)
        m1.metric("📅 ចំនួនថ្ងៃមានទិន្នន័យ", f"{total_days} ថ្ងៃ")
        m2.metric("📦 ចំនួនមុខទំនិញក្នុងតារាង", f"{len(item_cols)} មុខ")
        m3.metric("💰 ថវិកាសរុបប្រចាំខែ (ប៉ាន់ស្មាន)", format_riel_fn(total_month_budget))

        with st.expander("📊 មើលផលបូកបរិមាណមុខទំនិញសរុបប្រចាំខែ (Monthly Item Totals)", expanded=False):
            col_sums = {it: edited_mat[it].sum() for it in item_cols}
            df_col_sums = pd.DataFrame([{"មុខទំនិញ": k, "បរិមាណសរុប (គ.ក)": v, "តម្លៃរាយ": prod_map.get(k, {}).get("p1", 0), "សរុប (៛)": v * prod_map.get(k, {}).get("p1", 0)} for k, v in col_sums.items() if v > 0])
            if not df_col_sums.empty:
                st.dataframe(df_col_sums, use_container_width=True, hide_index=True)

        b_c1, b_c2 = st.columns([1.5, 1])
        with b_c1:
            if st.button("💾 រក្សាទុកតារាងប្រចាំថ្ងៃចូល Database", key=f"btn_save_mat_{m_school}", use_container_width=True):
                saved_cnt = 0
                for _, row in edited_mat.iterrows():
                    d_put = str(row.get("កាលបរិច្ឆេទដាក់", "")).strip()
                    d_eat = str(row.get("ថ្ងៃហូប", "")).strip()
                    v_no = str(row.get("លេខវិក័យប័ត្រ", "")).strip()
                    if not d_put:
                        continue

                    for it in item_cols:
                        q_val = float(row.get(it, 0) or 0)
                        if q_val > 0:
                            u_val = float(prod_map.get(it, {}).get("p1", 0) or 0)
                            tot_val = q_val * u_val
                            existing_rec = cursor.execute("SELECT id FROM daily_records WHERE date=? AND school_name=? AND item_name=?", (d_put, m_school, it)).fetchone()
                            if existing_rec:
                                cursor.execute(
                                    "UPDATE daily_records SET quantity=?, unit_price=?, total_price=?, voucher_no=?, consumption_date=? WHERE id=?",
                                    (q_val, u_val, tot_val, v_no, d_eat, existing_rec[0])
                                )
                            else:
                                cursor.execute(
                                    "INSERT INTO daily_records (date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no, consumption_date) VALUES (?,?,?,?,?,?,?,?,?)",
                                    (d_put, m_school, it, "វគ្គ១", q_val, u_val, tot_val, v_no, d_eat)
                                )
                            saved_cnt += 1
                conn.commit()
                st.success(f"🎉 បានរក្សាទុកជោគជ័យ! ចំនួន {saved_cnt} កំណត់ត្រាត្រូវបានបញ្ចូល/ធ្វើបច្ចុប្បន្នភាពក្នុង Database!")
                st.rerun()

        with b_c2:
            mat_excel_bytes = to_excel_fn(edited_mat)
            st.download_button(
                "📤 នាំចេញតារាងម៉ាទ្រីសជា Excel",
                data=mat_excel_bytes,
                file_name=f"តារាងបញ្ជាទិញ_{m_school}_{target_prefix}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


def render_universal_importer_tab(conn, cursor, format_riel_fn):
    """
    បង្ហាញផ្ទាំង «នាំចូលឯកសារពីខាងក្រៅ» (Universal Multi-format Importer)
    គាំទ្រ Excel (.xlsm, .xlsx, .xls), Word (.docx, .doc), PDF (.pdf), រូបភាព OCR (.png, .jpg, .jpeg) និង CSV
    """
    import os
    import streamlit as st

    st.subheader("📥 នាំចូលឯកសារពីខាងក្រៅ (Excel, XLSM, Word, PDF, រូបភាព OCR, CSV)")
    st.info(
        "💡 ប្រព័ន្ធសិក្សាពីរចនាសម្ព័ន្ធឯកសារស្វ័យប្រវត្តិ (គាំទ្រ .xlsm, .xlsx, .doc, .docx, .pdf, .png, .jpg, .csv) | "
        "អាចអានឯកសារ «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» ចាប់យករាល់តារាងសាលារៀន (Sleng, PhomDey...) សំណើទូទាត់ (R-*) និងបញ្ជីទិញកសិករ (Viget/Feed) | "
        "ទិន្នន័យទាំងអស់អាច **កែសម្រួលបានភ្លាមៗ** និង **គណនាតម្លៃសរុបស្វ័យប្រវត្តិ** មុននឹងចម្លងចូលក្នុង Database"
    )

    col_up_src1, col_up_src2 = st.columns([1.5, 1])
    with col_up_src1:
        up_univ_file = st.file_uploader(
            "📂 ជ្រើសរើសឯកសារសម្រាប់នាំចូល (xlsm, xlsx, docx, doc, pdf, png, jpg, csv)",
            type=["xlsm", "xlsx", "xls", "docx", "doc", "pdf", "png", "jpg", "jpeg", "csv"],
            key="up_univ_file"
        )
    with col_up_src2:
        st.markdown("###### ⚡ ឬប្រើប្រាស់ឯកសារប្រព័ន្ធដែលមានស្រាប់:")
        if os.path.exists("បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm"):
            if st.button("📖 សិក្សាឯកសារ «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» ភ្លាមៗ", key="btn_load_ws_file", use_container_width=True):
                st.session_state["use_workspace_file"] = "បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm"

    file_to_parse = None
    if up_univ_file:
        file_to_parse = up_univ_file
    elif st.session_state.get("use_workspace_file") and os.path.exists(st.session_state["use_workspace_file"]):
        file_to_parse = st.session_state["use_workspace_file"]

    if file_to_parse:
        with st.spinner("🔍 ប្រព័ន្ធកំពុងសិក្សា និងវិភាគរចនាសម្ព័ន្ធឯកសារ..."):
            if isinstance(file_to_parse, str):
                with open(file_to_parse, "rb") as f_in:
                    parsed_res = parse_excel_workbook(f_in.read(), db_conn=conn)
                file_display_name = file_to_parse
            else:
                parsed_res = parse_external_file(file_to_parse, db_conn=conn)
                file_display_name = file_to_parse.name

        if parsed_res and parsed_res.get("sheets"):
            total_s = len(parsed_res["sheets"])
            st.success(f"✅ បានសិក្សា និងស្រង់ចេញនូវសន្លឹកទិន្នន័យចំនួន **{total_s}** ពីឯកសារ `{file_display_name}`!")

            matrix_sheets = [k for k, v in parsed_res["sheets"].items() if v.get("sheet_type") == "daily_matrix"]
            if matrix_sheets:
                with st.container(border=True):
                    st.markdown(f"#### ⚡ មុខងារនាំចូលស្វ័យប្រវត្តិគ្រប់សាលាទាំងអស់ (Batch Import - {len(matrix_sheets)} សាលា)")
                    st.caption(f"រកឃើញតារាងសាលារៀនចំនួន {len(matrix_sheets)}: {', '.join(matrix_sheets)}")
                    if st.button(f"🚀 ចម្លងកំណត់ត្រាគ្រប់សាលាទាំង {len(matrix_sheets)} ចូល Database ក្នុងពេលតែមួយ", key="btn_batch_import_all", type="primary", use_container_width=True):
                        tot_imported = 0
                        import_details = []
                        for m_s_name in matrix_sheets:
                            s_info = parsed_res["sheets"][m_s_name]
                            if not s_info["df"].empty:
                                c_cnt, _ = import_records_to_database(s_info["df"], target_table="daily_records")
                                tot_imported += c_cnt
                                import_details.append(f"• សាលា **{s_info['school_name']}**: {c_cnt} កំណត់ត្រា")
                        st.success(f"🎉 បាននាំចូលកំណត់ត្រាសរុបចំនួន **{tot_imported}** ជួរដេក ដោយជោគជ័យ!")
                        st.markdown("\n".join(import_details))
                        st.rerun()

            st.divider()
            st.markdown("#### 🔍 ពិនិត្យ កែសម្រួល និងចម្លងចូលប្រព័ន្ធតាមសន្លឹកនីមួយៗ (Interactive Sheet Editor)")

            sheet_options = list(parsed_res["sheets"].keys())
            def format_sheet_label(k):
                info = parsed_res["sheets"][k]
                stype = info.get("sheet_type", "")
                recs = len(info.get("df", []))
                if stype == "daily_matrix":
                    return f"📅 {k} ({info.get('school_name', '')}) - កំណត់ត្រាប្រចាំថ្ងៃ [{recs} ជួរ]"
                elif stype == "monthly_claim":
                    return f"📑 {k} ({info.get('school_name', '')}) - សំណើទូទាត់ [{recs} មុខ]"
                elif stype == "farmer_purchases":
                    return f"🛒 {k} - បញ្ជីទិញកសិករ [{recs} ជួរ]"
                elif stype == "price_catalog":
                    return f"📦 {k} - តារាងតម្លៃទំនិញ [{recs} មុខ]"
                return f"📄 {k} [{recs} ជួរ]"

            sel_sheet = st.selectbox("ជ្រើសរើសសន្លឹកទិន្នន័យដែលត្រូវពិនិត្យ៖", sheet_options, format_func=format_sheet_label, key="sel_parsed_sheet")
            curr_sheet_data = parsed_res["sheets"][sel_sheet]

            st.markdown(f"##### {curr_sheet_data.get('title_kh', sel_sheet)}")
            df_to_edit = curr_sheet_data.get("df", pd.DataFrame())

            if not df_to_edit.empty:
                edited_sheet_df = st.data_editor(
                    df_to_edit,
                    use_container_width=True,
                    num_rows="dynamic",
                    key=f"editor_sheet_{sel_sheet}"
                )

                if "បរិមាណ" in edited_sheet_df.columns and "តម្លៃរាយ (៛)" in edited_sheet_df.columns and "សរុប (៛)" in edited_sheet_df.columns:
                    edited_sheet_df["សរុប (៛)"] = (
                        pd.to_numeric(edited_sheet_df["បរិមាណ"], errors="coerce").fillna(0) *
                        pd.to_numeric(edited_sheet_df["តម្លៃរាយ (៛)"], errors="coerce").fillna(0)
                    ).round(2)

                sum_qty = edited_sheet_df["បរិមាណ"].sum() if "បរិមាណ" in edited_sheet_df.columns else 0
                sum_tot = edited_sheet_df["សរុប (៛)"].sum() if "សរុប (៛)" in edited_sheet_df.columns else 0
                c_m1, c_m2, c_m3 = st.columns(3)
                c_m1.metric("ជួរដេកសរុប", f"{len(edited_sheet_df)} ជួរ")
                c_m2.metric("បរិមាណសរុប", f"{sum_qty:g}")
                c_m3.metric("ទឹកប្រាក់សរុប", format_riel_fn(sum_tot))

                s_type = curr_sheet_data.get("sheet_type")
                target_tbl = "purchases" if s_type == "farmer_purchases" else "daily_records"

                if st.button(f"💾 ចម្លងទិន្នន័យសន្លឹក [{sel_sheet}] នេះចូលក្នុង Database", key=f"btn_import_{sel_sheet}", use_container_width=True):
                    imp_cnt, imp_msg = import_records_to_database(edited_sheet_df, target_table=target_tbl)
                    st.success(f"🎉 {imp_msg}")
                    st.rerun()
            else:
                st.info(f"សន្លឹក [{sel_sheet}] នេះគ្មានជួរដេកទិន្នន័យទេ។")
        else:
            st.warning("មិនអាចទាញយកទិន្នន័យពីឯកសារនេះបានទេ។ សូមពិនិត្យមើលទ្រង់ទ្រាយឯកសារឡើងវិញ។")

