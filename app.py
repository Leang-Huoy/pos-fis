import hashlib
import io
import os
import re
import sqlite3
from datetime import date
from docx import Document
from fpdf import FPDF
import numpy as np
import pandas as pd
import pdfplumber
from PIL import Image
from rapidocr_onnxruntime import RapidOCR
import streamlit as st
import excel_schema_importer as esi
from catalog_data import SUPPLIER_PRODUCT_CATALOG

# កំណត់ទំព័រ Interface
st.set_page_config(
    page_title="ប្រព័ន្ធគ្រប់គ្រង POS និងសាលារៀន",
    page_icon="🛒",
    layout="wide",
)

# បង្កើត Database Connection
conn = sqlite3.connect("school_pos.db", check_same_thread=False)
cursor = conn.cursor()


# មុខងារជំនួយសម្រាប់ទ្រង់ទ្រាយរូបិយប័ណ្ណប្រាក់រៀល (៛)
def format_riel(amount):
  if amount is None:
    return "0 ៛"
  try:
    val = float(amount)
    if val % 1 == 0:
      return f"{int(val):,} ៛"
    else:
      return f"{val:,.2f} ៛"
  except (ValueError, TypeError):
    return f"{amount} ៛"


# មុខងារបន្ថែមលេខរៀង (ល.រ ពីលេខ ១ ដល់ N) ទៅក្នុង DataFrame
def add_row_numbers(df: pd.DataFrame, col_name="ល.រ") -> pd.DataFrame:
  if df is None or df.empty:
    return df
  df_copy = df.copy()
  if col_name in df_copy.columns:
    df_copy = df_copy.drop(columns=[col_name])
  df_copy.insert(0, col_name, range(1, len(df_copy) + 1))
  return df_copy


# មុខងារជំនួយសម្រាប់ស្វែងរក Font ខ្មែរ (Cross-platform: Windows / Linux / Cloud)
def find_khmer_font():
  import os
  candidates = [
      r"C:\Windows\Fonts\KhmerOS.ttf",
      r"C:\Windows\Fonts\khmeros.ttf",
      "/usr/share/fonts/truetype/khmeros/KhmerOS.ttf",
      "/usr/share/fonts/truetype/khmer/KhmerOS.ttf",
      "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
      "KhmerOS.ttf",
  ]
  for c in candidates:
    if os.path.exists(c):
      return c
  return None


# មុខងារជំនួយសម្រាប់ស្វែងរក Headless Browser (Windows Edge / Chrome / Linux Chromium)
def find_headless_browser():
  import shutil
  import os
  for name in ["chromium", "chromium-browser", "google-chrome", "google-chrome-stable", "msedge"]:
    found = shutil.which(name)
    if found:
      return found
  win_paths = [
      r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
      r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
      r"C:\Program Files\Google\Chrome\Application\chrome.exe",
      r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
  ]
  return next((p for p in win_paths if os.path.exists(p)), None)


# បញ្ជីប្រភេទមុខទំនិញស្ដង់ដារ (សម្រាប់វិក្កយបត្រ និងសំណើទូទាត់)
STANDARD_CATEGORIES = ["អង្ករ", "អំបិល", "ប្រេងឆា", "ត្រី សាច់ ស៊ុត", "បន្លែ"]

# បញ្ជីមុខទំនិញស្ដង់ដារទាំង ៦១ មុខ (ផ្អែកលើរូបភាព និងតារាងតម្លៃស្ដង់ដារ)
STANDARD_PRODUCT_CATALOG = [
    {"name": "អង្ករចម្រុះ", "category": "អង្ករ", "price_phase1": 2100.0, "price_phase2": 2100.0},
    {"name": "ប្រេងឆា", "category": "ប្រេងឆា", "price_phase1": 6499.0, "price_phase2": 6499.0},
    {"name": "អំបិលអ៊ីយូត", "category": "អំបិល", "price_phase1": 1000.0, "price_phase2": 1000.0},
    {"name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 15950.0, "price_phase2": 15950.0},
    {"name": "ស៊ុតទា", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 580.0, "price_phase2": 580.0},
    {"name": "ត្រីប្រា", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 7290.0, "price_phase2": 7290.0},
    {"name": "ត្រីរស់", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 10000.0, "price_phase2": 10000.0},
    {"name": "ត្រីអណ្តែង", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 7990.0, "price_phase2": 7990.0},
    {"name": "ត្រកួន", "category": "បន្លែ", "price_phase1": 2495.0, "price_phase2": 2495.0},
    {"name": "ស្លឹកបាស", "category": "បន្លែ", "price_phase1": 3999.0, "price_phase2": 3999.0},
    {"name": "ស្លឹកម្រុំ", "category": "បន្លែ", "price_phase1": 2440.0, "price_phase2": 2440.0},
    {"name": "ស្ពៃក្រញាញ់", "category": "បន្លែ", "price_phase1": 3999.0, "price_phase2": 3999.0},
    {"name": "ស្ពៃជើងទា", "category": "បន្លែ", "price_phase1": 3699.0, "price_phase2": 3699.0},
    {"name": "ស្ពៃចង្កឹះ", "category": "បន្លែ", "price_phase1": 3950.0, "price_phase2": 3950.0},
    {"name": "ស្ពៃខ្មៅ", "category": "បន្លែ", "price_phase1": 3490.0, "price_phase2": 3490.0},
    {"name": "ផ្ទីដូង", "category": "បន្លែ", "price_phase1": 2890.0, "price_phase2": 2890.0},
    {"name": "ត្រួយល្ពៅ", "category": "បន្លែ", "price_phase1": 2590.0, "price_phase2": 2590.0},
    {"name": "ស្លឹកងប់", "category": "បន្លែ", "price_phase1": 1999.0, "price_phase2": 1999.0},
    {"name": "ផ្លែល្ពៅ", "category": "បន្លែ", "price_phase1": 2790.0, "price_phase2": 2790.0},
    {"name": "ផ្លែឃ្លោក", "category": "បន្លែ", "price_phase1": 1550.0, "price_phase2": 1890.0},
    {"name": "ផ្លែត្រឡាច", "category": "បន្លែ", "price_phase1": 2450.0, "price_phase2": 2450.0},
    {"name": "ត្រប់វែង", "category": "បន្លែ", "price_phase1": 2490.0, "price_phase2": 2490.0},
    {"name": "ត្រប់ស្រួយ", "category": "បន្លែ", "price_phase1": 3250.0, "price_phase2": 3250.0},
    {"name": "ប៉េងប៉ោះ", "category": "បន្លែ", "price_phase1": 3495.0, "price_phase2": 3495.0},
    {"name": "ននោងមូល", "category": "បន្លែ", "price_phase1": 2499.0, "price_phase2": 2499.0},
    {"name": "ននោងជ្រុង", "category": "បន្លែ", "price_phase1": 2390.0, "price_phase2": 2390.0},
    {"name": "ល្ហុងខ្ចី", "category": "បន្លែ", "price_phase1": 1495.0, "price_phase2": 1495.0},
    {"name": "សណ្តែកគួរ", "category": "បន្លែ", "price_phase1": 3999.0, "price_phase2": 3999.0},
    {"name": "ត្រយូងចេក", "category": "បន្លែ", "price_phase1": 1500.0, "price_phase2": 1500.0},
    {"name": "សណ្តែកដីលីង", "category": "បន្លែ", "price_phase1": 7990.0, "price_phase2": 7990.0},
    {"name": "ការ៉ុត", "category": "បន្លែ", "price_phase1": 2999.0, "price_phase2": 2999.0},
    {"name": "ផ្កាខាត់ណា", "category": "បន្លែ", "price_phase1": 4890.0, "price_phase2": 4890.0},
    {"name": "ដើមខាត់ណា", "category": "បន្លែ", "price_phase1": 2500.0, "price_phase2": 2500.0},
    {"name": "ដំឡូងពណលឿង", "category": "បន្លែ", "price_phase1": 1500.0, "price_phase2": 1500.0},
    {"name": "សាច់គោ (សាច់យាវ)", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 25000.0, "price_phase2": 25000.0},
    {"name": "ត្រីពោធិ៍", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 10000.0, "price_phase2": 10000.0},
    {"name": "ត្រីខ្យា", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 12000.0, "price_phase2": 12000.0},
    {"name": "ត្រីសណ្ដាយ", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 18000.0, "price_phase2": 18000.0},
    {"name": "ត្រីផ្ទក់", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 10000.0, "price_phase2": 10000.0},
    {"name": "សាច់មាន់", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 9500.0, "price_phase2": 9500.0},
    {"name": "ឆ្អឹងជំនីជ្រូក", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 15980.0, "price_phase2": 15980.0},
    {"name": "ឆ្អឹងជ្រូកស៊ុប", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 14590.0, "price_phase2": 14590.0},
    {"name": "ជើងជ្រូក", "category": "ត្រី សាច់ ស៊ុត", "price_phase1": 13590.0, "price_phase2": 13590.0},
    {"name": "ព្រលឹត", "category": "បន្លែ", "price_phase1": 1500.0, "price_phase2": 1500.0},
    {"name": "ត្រួយស្អំ", "category": "បន្លែ", "price_phase1": 7000.0, "price_phase2": 7000.0},
    {"name": "ត្រួយននោង", "category": "បន្លែ", "price_phase1": 4000.0, "price_phase2": 4000.0},
    {"name": "ត្រួយ ឬផ្កាអង្គារដី", "category": "បន្លែ", "price_phase1": 3000.0, "price_phase2": 3000.0},
    {"name": "ផ្សិតអំបោះ", "category": "បន្លែ", "price_phase1": 7000.0, "price_phase2": 7000.0},
    {"name": "ស្ពៃក្ដោប", "category": "បន្លែ", "price_phase1": 2999.0, "price_phase2": 2999.0},
    {"name": "ស្ពៃបូកគោ", "category": "បន្លែ", "price_phase1": 3999.0, "price_phase2": 3999.0},
    {"name": "សណ្ដែកបណ្ដុះ", "category": "បន្លែ", "price_phase1": 3000.0, "price_phase2": 3000.0},
    {"name": "ក្ដិបឪឡឹក", "category": "បន្លែ", "price_phase1": 2500.0, "price_phase2": 2500.0},
    {"name": "ពោតបារាំង", "category": "បន្លែ", "price_phase1": 3000.0, "price_phase2": 3000.0},
    {"name": "ស្នៀតពោត", "category": "បន្លែ", "price_phase1": 3000.0, "price_phase2": 3000.0},
    {"name": "ផ្លែពោត", "category": "បន្លែ", "price_phase1": 3800.0, "price_phase2": 3800.0},
    {"name": "ផ្លែត្នោតខ្ចី", "category": "បន្លែ", "price_phase1": 5000.0, "price_phase2": 5000.0},
    {"name": "ផ្លែម្នាស់ទុំ", "category": "បន្លែ", "price_phase1": 2700.0, "price_phase2": 2700.0},
    {"name": "ផ្លែត្រសក់", "category": "បន្លែ", "price_phase1": 2800.0, "price_phase2": 2800.0},
    {"name": "មើមត្រាវ", "category": "បន្លែ", "price_phase1": 3000.0, "price_phase2": 3000.0},
    {"name": "មើមឆៃថាវ", "category": "បន្លែ", "price_phase1": 3000.0, "price_phase2": 3000.0},
    {"name": "ទំពាំង", "category": "បន្លែ", "price_phase1": 1800.0, "price_phase2": 1800.0},
]
STANDARD_PRODUCT_ITEMS = [item["name"] for item in STANDARD_PRODUCT_CATALOG]


def classify_item_category(item_name: str) -> str:
  """កំណត់ប្រភេទមុខទំនិញស្វ័យប្រវត្តិតាមឈ្មោះ៖ អង្ករ, អំបិល, ប្រេងឆា, ត្រី សាច់ ស៊ុត, បន្លែ"""
  if not item_name:
    return "បន្លែ"
  name = str(item_name).strip()
  # ពិនិត្យក្នុងកាតាឡុកស្ដង់ដារជាមុន
  for item in STANDARD_PRODUCT_CATALOG:
    if item["name"] == name:
      return item["category"]

  # 1. អង្ករ
  if any(k in name for k in ["អង្ករ", "បាយ"]):
    return "អង្ករ"
  # 2. អំបិល
  if any(k in name for k in ["អំបិល", "អ៊ីយ៉ូត", "អ៊ីយូត"]):
    return "អំបិល"
  # 3. ប្រេងឆា
  if any(k in name for k in ["ប្រេង", "ប្រេងឆា", "ខ្លាញ់"]):
    return "ប្រេងឆា"
  # បន្លែដែលឈ្មោះមានពាក្យសត្វ (ដូចជា ស្ពៃជើងទា, ស្ពៃបូកគោ, សណ្តែកគួរ)
  if any(k in name for k in [
      "ស្ពៃ", "ត្រកួន", "ស្លឹក", "ផ្ទី", "ត្រួយ", "ផ្លែ", "ត្រប់", "ប៉េងប៉ោះ",
      "ននោង", "ល្ហុង", "សណ្តែក", "សណ្ដែក", "ត្រយូង", "ការ៉ុត", "ខាត់ណា",
      "ដំឡូង", "ព្រលឹត", "ផ្សិត", "ក្ដិប", "ពោត", "មើម", "ទំពាំង"
  ]):
    return "បន្លែ"
  # 4. ត្រី សាច់ ស៊ុត
  if any(k in name for k in [
      "សាច់", "ជ្រូក", "គោ", "មាន់", "ទា", "ត្រី", "ស៊ុត", "ពង",
      "ប្រហិត", "ក្រៀម", "ងៀត", "ឆ្អើរ", "ក្តាម", "បង្គា", "មឹក", "ឆ្អឹង", "ជើង"
  ]):
    return "ត្រី សាច់ ស៊ុត"
  return "បន្លែ"


def format_category_badge(cat: str) -> str:
  """រៀបចំ Icon និងស្លាកសម្គាល់ប្រភេទមុខទំនិញឱ្យស្អាត"""
  c = str(cat or "").strip()
  if "អង្ករ" in c:
    return "🍚 អង្ករ"
  elif "អំបិល" in c:
    return "🧂 អំបិល"
  elif "ប្រេង" in c:
    return "🫗 ប្រេងឆា"
  elif any(k in c for k in ["សាច់", "ត្រី", "ស៊ុត"]):
    return "🥩 ត្រី សាច់ ស៊ុត"
  elif "បន្លែ" in c:
    return "🥬 បន្លែ"
  elif c:
    return f"🏷️ {c}"
  return "⚪ មិនទាន់កំណត់"


# មុខងារ Hash Password
def hash_password(password):
  return hashlib.sha256(password.encode()).hexdigest()


# បង្កើត និង Update តារាងទាំងអស់
def init_db():
  # តារាងអ្នកប្រើប្រាស់ (Users)
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        full_name TEXT,
        role TEXT
    )""")

  # បង្កើត Admin លំនាំដើមបើមិនទាន់មាន
  admin_pass = hash_password("admin123")
  cursor.execute(
      """
        INSERT OR IGNORE INTO users (username, password, full_name, role) 
        VALUES ('admin', ?, 'Administrator', 'Admin')
    """,
      (admin_pass,),
  )

  # តារាងទីតាំងរដ្ឋបាល
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS locations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        province TEXT,
        district TEXT,
        commune TEXT,
        village TEXT
    )""")

  # តារាងសាលារៀន
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS schools (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        commune TEXT,
        district TEXT,
        province TEXT,
        village TEXT
    )""")

  # Upgrade schema schools បើខ្វះ column
  cursor.execute("PRAGMA table_info(schools)")
  existing_cols = [c[1] for c in cursor.fetchall()]
  if "province" not in existing_cols:
    cursor.execute("ALTER TABLE schools ADD COLUMN province TEXT")
  if "district" not in existing_cols:
    cursor.execute("ALTER TABLE schools ADD COLUMN district TEXT")
  if "village" not in existing_cols:
    cursor.execute("ALTER TABLE schools ADD COLUMN village TEXT")

  # បំពេញ province/district សម្រាប់ទិន្នន័យចាស់ដែលមានតែ commune
  cursor.execute("""
    UPDATE schools 
    SET province = (SELECT province FROM locations WHERE locations.commune = schools.commune LIMIT 1),
        district = (SELECT district FROM locations WHERE locations.commune = schools.commune LIMIT 1),
        village = (SELECT village FROM locations WHERE locations.commune = schools.commune LIMIT 1)
    WHERE (province IS NULL OR province = '')
  """)

  # តារាងមុខទំនិញ (មានតម្លៃមធ្យម price_avg និងកាលបរិច្ឆេទវគ្គ១-២)
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_name TEXT,
        commune TEXT,
        price_phase1 REAL,
        price_phase2 REAL,
        price_avg REAL,
        phase1_start TEXT,
        phase1_end TEXT,
        phase2_start TEXT,
        phase2_end TEXT
    )""")

  cursor.execute("PRAGMA table_info(products)")
  prod_cols = [c[1] for c in cursor.fetchall()]
  if "price_avg" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN price_avg REAL")
  if "phase1_start" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN phase1_start TEXT")
  if "phase1_end" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN phase1_end TEXT")
  if "phase2_start" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN phase2_start TEXT")
  if "phase2_end" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN phase2_end TEXT")
  if "category" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN category TEXT DEFAULT ''")
  if "province" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN province TEXT DEFAULT ''")
  if "district" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN district TEXT DEFAULT ''")
  if "school_name" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN school_name TEXT DEFAULT ''")
  if "price_level" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN price_level TEXT DEFAULT 'commune'")
  if "supplier_name" not in prod_cols:
    cursor.execute("ALTER TABLE products ADD COLUMN supplier_name TEXT DEFAULT ''")

  # Update price_level and province/district for products
  cursor.execute("UPDATE products SET price_level='commune' WHERE price_level IS NULL OR price_level=''")
  cursor.execute("""
    UPDATE products 
    SET province = (SELECT province FROM locations WHERE locations.commune = products.commune LIMIT 1),
        district = (SELECT district FROM locations WHERE locations.commune = products.commune LIMIT 1)
    WHERE (province IS NULL OR province = '') AND commune IS NOT NULL AND commune != ''
  """)

  # គណនាតម្លៃមធ្យមសម្រាប់ទិន្នន័យចាស់ដែលមិនទាន់មាន price_avg
  cursor.execute("""
    UPDATE products 
    SET price_avg = ROUND((COALESCE(price_phase1, 0) + COALESCE(price_phase2, 0)) / 2.0, 2)
    WHERE price_avg IS NULL
  """)

  # កំណត់កាលបរិច្ឆេទលំនាំដើមសម្រាប់ទិន្នន័យមុខទំនិញដែលមិនទាន់មានកាលបរិច្ឆេទ
  cur_year = date.today().year
  cursor.execute(f"""
    UPDATE products 
    SET phase1_start = COALESCE(NULLIF(phase1_start, ''), '{cur_year}-01-01'),
        phase1_end   = COALESCE(NULLIF(phase1_end, ''), '{cur_year}-06-30'),
        phase2_start = COALESCE(NULLIF(phase2_start, ''), '{cur_year}-07-01'),
        phase2_end   = COALESCE(NULLIF(phase2_end, ''), '{cur_year}-12-31')
    WHERE phase1_start IS NULL OR phase1_start = ''
  """)

  # បំពេញ category ស្វ័យប្រវត្តសម្រាប់មុខទំនិញដែលមិនទាន់មានប្រភេទសម្គាល់
  cursor.execute("SELECT id, item_name FROM products WHERE category IS NULL OR category = ''")
  for pid, pname in cursor.fetchall():
    auto_cat = classify_item_category(pname)
    cursor.execute("UPDATE products SET category=? WHERE id=?", (auto_cat, pid))

  # តារាងកត់ត្រាប្រចាំថ្ងៃ (មានលេខសក្ខីប័ត្រ voucher_no)
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS daily_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        school_name TEXT,
        item_name TEXT,
        phase TEXT,
        quantity REAL,
        unit_price REAL,
        total_price REAL,
        voucher_no TEXT
    )""")

  cursor.execute("PRAGMA table_info(daily_records)")
  dr_cols = [c[1] for c in cursor.fetchall()]
  if "voucher_no" not in dr_cols:
    cursor.execute("ALTER TABLE daily_records ADD COLUMN voucher_no TEXT")

  # បំពេញលេខសក្ខីប័ត្រស្វ័យប្រវត្ត ចាប់ផ្ដើមពី 001 សម្រាប់ទិន្នន័យចាស់ដែលមិនទាន់មាន
  schools_list = [r[0] for r in cursor.execute("SELECT DISTINCT school_name FROM daily_records WHERE school_name IS NOT NULL").fetchall()]
  for sch in schools_list:
    dates = [r[0] for r in cursor.execute("SELECT DISTINCT date FROM daily_records WHERE school_name=? AND (voucher_no IS NULL OR voucher_no='') ORDER BY date ASC", (sch,)).fetchall()]
    if dates:
      existing_v = [r[0] for r in cursor.execute("SELECT DISTINCT voucher_no FROM daily_records WHERE school_name=? AND voucher_no IS NOT NULL AND voucher_no!=''", (sch,)).fetchall()]
      nums = []
      for v in existing_v:
        digits = re.findall(r'\d+', str(v))
        if digits:
          nums.append(int(digits[-1]))
      cur_idx = max(nums) if nums else 0
      for d_val in dates:
        cur_idx += 1
        v_str = f"{cur_idx:03d}"
        cursor.execute("UPDATE daily_records SET voucher_no=? WHERE school_name=? AND date=? AND (voucher_no IS NULL OR voucher_no='')", (v_str, sch, d_val))

  # តារាងទិញទំនិញចូល
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS purchases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        item_name TEXT,
        unit_price REAL,
        quantity REAL,
        total_price REAL,
        supplier_name TEXT,
        status TEXT
    )""")

  # តារាងចំណូលចំណាយ
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        type TEXT,
        category TEXT,
        amount REAL,
        description TEXT
    )""")
  # តារាងអ្នកផ្គត់ផ្គង់តាមសាលារៀន (School Suppliers)
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS suppliers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_name TEXT,
        supplier_name TEXT,
        village TEXT,
        commune TEXT,
        district TEXT,
        province TEXT,
        phone TEXT,
        signature_data TEXT,
        supplied_categories TEXT
    )""")

  cursor.execute("PRAGMA table_info(suppliers)")
  sup_cols = [c[1] for c in cursor.fetchall()]
  if "supplied_categories" not in sup_cols:
    cursor.execute("ALTER TABLE suppliers ADD COLUMN supplied_categories TEXT DEFAULT ''")
  if "supply_level" not in sup_cols:
    cursor.execute("ALTER TABLE suppliers ADD COLUMN supply_level TEXT DEFAULT 'school'")
  if "target_commune" not in sup_cols:
    cursor.execute("ALTER TABLE suppliers ADD COLUMN target_commune TEXT DEFAULT ''")
  if "target_district" not in sup_cols:
    cursor.execute("ALTER TABLE suppliers ADD COLUMN target_district TEXT DEFAULT ''")
  if "target_province" not in sup_cols:
    cursor.execute("ALTER TABLE suppliers ADD COLUMN target_province TEXT DEFAULT ''")

  cursor.execute("""
    UPDATE suppliers 
    SET supplied_categories = 'អង្ករ, អំបិល, ប្រេងឆា, ត្រី សាច់ ស៊ុត, បន្លែ'
    WHERE supplied_categories IS NULL OR supplied_categories = ''
  """)

  # បញ្ចូលអ្នកផ្គត់ផ្គង់គំរូប្រសិនបើទទេ
  cursor.execute("SELECT COUNT(*) FROM suppliers")
  if cursor.fetchone()[0] == 0:
    cursor.execute("""
      INSERT INTO suppliers (school_name, supplier_name, village, commune, district, province, phone, signature_data, supplied_categories)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, ("ច្រឡង", "សាត ក្រូត", "ភូមិខ្មែរ", "រោង", "ស្រីស្នំ", "សៀមរាប", "090 854 133", "", "អង្ករ, អំបិល, ប្រេងឆា, ត្រី សាច់ ស៊ុត, បន្លែ"))
    cursor.execute("""
      INSERT INTO suppliers (school_name, supplier_name, village, commune, district, province, phone, signature_data, supplied_categories)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, ("សាលាបឋមសិក្សា ស្លែងស្ពាន", "សាត ក្រូត", "ភូមិខ្មែរ", "រោង", "ស្រីស្នំ", "សៀមរាប", "090 854 133", "", "អង្ករ, អំបិល, ប្រេងឆា, ត្រី សាច់ ស៊ុត, បន្លែ"))

  conn.commit()


init_db()


# ================= Helper Functions សម្រាប់អ្នកផ្គត់ផ្គង់តាមសាលា =================
def get_supplier_for_school(school_name):
  """ទាញយកព័ត៌មានអ្នកផ្គត់ផ្គង់សម្រាប់សាលារៀនណាមួយ"""
  if not school_name:
    return None
  row = cursor.execute("""
    SELECT supplier_name, village, commune, district, province, phone, signature_data, id, supplied_categories
    FROM suppliers
    WHERE school_name = ?
    ORDER BY id DESC LIMIT 1
  """, (school_name,)).fetchone()
  if row:
    addr_parts = []
    if row[1]: addr_parts.append(row[1])
    if row[2]: addr_parts.append(f"ឃុំ{row[2]}" if not row[2].startswith("ឃុំ") else row[2])
    return {
        "id": row[7],
        "supplier_name": row[0],
        "village": row[1] or "",
        "commune": row[2] or "",
        "district": row[3] or "",
        "province": row[4] or "",
        "phone": row[5] or "",
        "signature_data": row[6] or "",
        "address": " ".join(addr_parts) if addr_parts else "ភូមិខ្មែរ ឃុំរោង",
        "supplied_categories": row[8] or "អង្ករ, អំបិល, ប្រេងឆា, ត្រី សាច់ ស៊ុត, បន្លែ"
    }
  # Fallback ទៅអ្នកផ្គត់ផ្គង់ចុងក្រោយក្នុង DB
  row_any = cursor.execute("""
    SELECT supplier_name, village, commune, district, province, phone, signature_data, id, supplied_categories
    FROM suppliers
    ORDER BY id DESC LIMIT 1
  """).fetchone()
  if row_any:
    addr_parts = []
    if row_any[1]: addr_parts.append(row_any[1])
    if row_any[2]: addr_parts.append(f"ឃុំ{row_any[2]}" if not row_any[2].startswith("ឃុំ") else row_any[2])
    return {
        "id": row_any[7],
        "supplier_name": row_any[0],
        "village": row_any[1] or "",
        "commune": row_any[2] or "",
        "district": row_any[3] or "",
        "province": row_any[4] or "",
        "phone": row_any[5] or "",
        "signature_data": row_any[6] or "",
        "address": " ".join(addr_parts) if addr_parts else "ភូមិខ្មែរ ឃុំរោង",
        "supplied_categories": row_any[8] or "អង្ករ, អំបិល, ប្រេងឆា, ត្រី សាច់ ស៊ុត, បន្លែ"
    }
  return {
      "id": 0,
      "supplier_name": "សាត ក្រូត",
      "village": "ភូមិខ្មែរ",
      "commune": "រោង",
      "district": "ស្រីស្នំ",
      "province": "សៀមរាប",
      "phone": "090 854 133",
      "signature_data": "",
      "address": "ភូមិខ្មែរ ឃុំរោង",
      "supplied_categories": "អង្ករ, អំបិល, ប្រេងឆា, ត្រី សាច់ ស៊ុត, បន្លែ"
  }


class SupplierRow(dict):
  """រចនាសម្ព័ន្ធទិន្នន័យអ្នកផ្គត់ផ្គង់ដែលគាំទ្រទាំង Tuple Index [0] និង Dictionary Key ['supplier_name']"""
  def __init__(self, t):
    keys = [
        "id", "school_name", "supplier_name", "village", "commune", "district", "province",
        "phone", "signature_data", "supplied_categories", "supply_level",
        "target_commune", "target_district", "target_province"
    ]
    d = {k: (t[i] if i < len(t) and t[i] is not None else "") for i, k in enumerate(keys)}
    super().__init__(d)
    self._tuple = list(t)

  def __getitem__(self, item):
    if isinstance(item, int):
      return self._tuple[item] if item < len(self._tuple) else None
    return super().__getitem__(item)


def get_all_suppliers():
  """ទាញយកបញ្ជីអ្នកផ្គត់ផ្គង់ទាំងអស់"""
  cursor.execute("""
    SELECT id, school_name, supplier_name, village, commune, district, province, phone, signature_data, supplied_categories,
           COALESCE(supply_level, 'school') as supply_level,
           COALESCE(target_commune, '') as target_commune,
           COALESCE(target_district, '') as target_district,
           COALESCE(target_province, '') as target_province
    FROM suppliers
    ORDER BY id DESC
  """)
  return [SupplierRow(r) for r in cursor.fetchall()]


def save_or_update_supplier(
    school_name,
    supplier_name,
    village,
    commune,
    district,
    province,
    phone,
    signature_data=None,
    supplied_categories=None,
    supply_level="school",
    target_commune="",
    target_district="",
    target_province="",
    supplier_id=None
):
  """រក្សាទុក ឬកែប្រែអ្នកផ្គត់ផ្គង់តាមសាលា ឬតាមឃុំ"""
  sc = supplied_categories or "អង្ករ, អំបិល, ប្រេងឆា, ត្រី សាច់ ស៊ុត, បន្លែ"
  if supplier_id:
    if signature_data is not None:
      cursor.execute("""
        UPDATE suppliers 
        SET school_name=?, supplier_name=?, village=?, commune=?, district=?, province=?, phone=?,
            signature_data=?, supplied_categories=?, supply_level=?, target_commune=?, target_district=?, target_province=?
        WHERE id=?
      """, (school_name, supplier_name, village, commune, district, province, phone, signature_data, sc, supply_level, target_commune, target_district, target_province, supplier_id))
    else:
      cursor.execute("""
        UPDATE suppliers 
        SET school_name=?, supplier_name=?, village=?, commune=?, district=?, province=?, phone=?,
            supplied_categories=?, supply_level=?, target_commune=?, target_district=?, target_province=?
        WHERE id=?
      """, (school_name, supplier_name, village, commune, district, province, phone, sc, supply_level, target_commune, target_district, target_province, supplier_id))
    conn.commit()
    return supplier_id
  else:
    # Check if exists by name & school/commune
    existing = cursor.execute("""
      SELECT id FROM suppliers 
      WHERE supplier_name=? AND (school_name=? OR (school_name='' AND target_commune=?))
    """, (supplier_name, school_name, target_commune)).fetchone()
    if existing:
      s_id = existing[0]
      cursor.execute("""
        UPDATE suppliers 
        SET school_name=?, village=?, commune=?, district=?, province=?, phone=?,
            signature_data=COALESCE(?, signature_data), supplied_categories=?, supply_level=?, target_commune=?, target_district=?, target_province=?
        WHERE id=?
      """, (school_name, village, commune, district, province, phone, signature_data, sc, supply_level, target_commune, target_district, target_province, s_id))
      conn.commit()
      return s_id
    else:
      cursor.execute("""
        INSERT INTO suppliers (
          school_name, supplier_name, village, commune, district, province, phone,
          signature_data, supplied_categories, supply_level, target_commune, target_district, target_province
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      """, (school_name, supplier_name, village, commune, district, province, phone, signature_data or "", sc, supply_level, target_commune, target_district, target_province))
      conn.commit()
      return cursor.lastrowid


def delete_supplier(supplier_id):
  """លុបអ្នកផ្គត់ផ្គង់តាម ID"""
  cursor.execute("DELETE FROM suppliers WHERE id=?", (supplier_id,))
  conn.commit()


def get_catalog_base_price(item_name, school_name=None, commune=None):
  """
  ស្វែងរកតម្លៃគោលនៃមុខទំនិញ៖
  ១. ពិនិត្យមើលតម្លៃគោលក្នុង products (price_level='school' ឬ 'commune')
  ២. ប្រសិនបើគ្មាន ប្រើប្រាស់តម្លៃ base_avg ពី SUPPLIER_PRODUCT_CATALOG (ឯកសារ Excel)
  """
  if school_name:
    try:
      row = cursor.execute(
          "SELECT price_avg, price_phase1 FROM products WHERE item_name=? AND school_name=? AND price_level='school'",
          (item_name, school_name)
      ).fetchone()
      if row and (row[0] or row[1]):
        return float(row[0] or row[1])
    except Exception:
      pass
  if commune:
    try:
      row = cursor.execute(
          "SELECT price_avg, price_phase1 FROM products WHERE item_name=? AND commune=? AND (price_level='commune' OR school_name IS NULL OR school_name='')",
          (item_name, commune)
      ).fetchone()
      if row and (row[0] or row[1]):
        return float(row[0] or row[1])
    except Exception:
      pass
  
  for it in SUPPLIER_PRODUCT_CATALOG:
    if it["name"].strip() == item_name.strip():
      return float(it.get("base_avg", 0) or it.get("default_p1", 0))
  return 0.0


def evaluate_supplier_price_status(avg_price, base_price):
  """
  គណនាស្ថានភាពប្រៀបធៀបតម្លៃមធ្យម ទៅនឹងតម្លៃគោល៖
  - ខ្ពស់ជាងតម្លៃគោល 10% (> +10%): ពណ៌ក្រហម (#fee2e2, border #ef4444, text #991b1b)
  - ទាបជាងតម្លៃគោល 10% (< -10%): ពណ៌លឿង (#fef9c3, border #eab308, text #854d0e)
  - ស្ថិតនៅចន្លោះតម្លៃគោល ±10%: ពណ៌ត្រួយចេក (#ecfccb, border #84cc16, text #365314)
  """
  if not base_price or base_price <= 0:
    return {
        "status": "normal",
        "diff_pct": 0.0,
        "bg_color": "#f8fafc",
        "border_color": "#cbd5e1",
        "text_color": "#334155",
        "badge_text": "⚪ គ្មានតម្លៃគោលកំណត់",
        "badge_class": "badge-neutral"
    }

  upper_limit = base_price * 1.10
  lower_limit = base_price * 0.90
  diff_pct = ((avg_price - base_price) / base_price) * 100.0

  if avg_price > upper_limit:
    return {
        "status": "high",
        "diff_pct": diff_pct,
        "bg_color": "#fee2e2",
        "border_color": "#ef4444",
        "text_color": "#991b1b",
        "badge_text": f"🔴 ថ្លៃជាងតម្លៃគោល ១០%+ (+{diff_pct:.1f}%)",
        "badge_class": "badge-red"
    }
  elif avg_price < lower_limit:
    return {
        "status": "low",
        "diff_pct": diff_pct,
        "bg_color": "#fef9c3",
        "border_color": "#eab308",
        "text_color": "#854d0e",
        "badge_text": f"🟡 ថោកជាងតម្លៃគោល ១០%+ ({diff_pct:.1f}%)",
        "badge_class": "badge-yellow"
    }
  else:
    sign = "+" if diff_pct > 0 else ""
    return {
        "status": "normal_lime",
        "diff_pct": diff_pct,
        "bg_color": "#ecfccb",
        "border_color": "#84cc16",
        "text_color": "#365314",
        "badge_text": f"🟢 ត្រួយចេក សមស្រប (±១០%) [{sign}{diff_pct:.1f}%]",
        "badge_class": "badge-lime"
    }


def get_supplier_prices_map(supplier_name):
  """ទាញយកតម្លៃទំនិញដែលបានកំណត់របស់អ្នកផ្គត់ផ្គង់ពីតារាង products"""
  res = {}
  if not supplier_name:
    return res
  try:
    rows = cursor.execute("""
      SELECT item_name, price_phase1, price_phase2, price_avg
      FROM products
      WHERE supplier_name=? AND price_level='supplier'
    """, (supplier_name,)).fetchall()
    for r in rows:
      res[r[0]] = {
          "p1": float(r[1] or 0),
          "p2": float(r[2] or 0),
          "avg": float(r[3] or 0)
      }
  except Exception:
    pass
  return res


def save_all_supplier_prices(supplier_name, price_dict, supply_level="school", target_school="", target_commune="", target_district="", target_province=""):
  """រក្សាទុកតម្លៃទំនិញទាំងអស់សម្រាប់អ្នកផ្គត់ផ្គង់នេះ"""
  saved_count = 0
  for it in SUPPLIER_PRODUCT_CATALOG:
    name = it["name"]
    category = it["category"]
    p_info = price_dict.get(name, {"p1": it["default_p1"], "p2": it["default_p2"]})
    p1 = float(p_info.get("p1", it["default_p1"]))
    p2 = float(p_info.get("p2", it["default_p2"]))
    avg_p = round((p1 + p2) / 2.0, 2)
    sch_col = target_school if supply_level == "school" else ""
    com_col = target_commune

    existing = cursor.execute("""
      SELECT id FROM products 
      WHERE supplier_name=? AND item_name=? AND price_level='supplier'
    """, (supplier_name, name)).fetchone()

    if existing:
      cursor.execute("""
        UPDATE products 
        SET price_phase1=?, price_phase2=?, price_avg=?, category=?, commune=?, district=?, province=?, school_name=?, price_level='supplier'
        WHERE id=?
      """, (p1, p2, avg_p, category, com_col, target_district, target_province, sch_col, existing[0]))
    else:
      cursor.execute("""
        INSERT INTO products (
          item_name, commune, district, province, school_name,
          price_phase1, price_phase2, price_avg, category,
          price_level, supplier_name
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'supplier', ?)
      """, (name, com_col, target_district, target_province, sch_col, p1, p2, avg_p, category, supplier_name))
    saved_count += 1

  conn.commit()
  return saved_count



def file_to_base64_img(uploaded_file):
  """បំប្លែង File Uploaded ទៅជា Data URL Base64 សម្រាប់ប្រើក្នុង HTML និង PDF"""
  if uploaded_file is None:
    return None
  try:
    import base64
    b64_str = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
    mime_type = getattr(uploaded_file, "type", "image/png") or "image/png"
    return f"data:{mime_type};base64,{b64_str}"
  except Exception:
    return None


def process_signature_image(
    image_input,
    remove_bg: bool = True,
    threshold: int = 215,
    softness: int = 40,
    recolor_choice: str = "blue",
    custom_hex: str = "#0B3C95",
    contrast_boost: float = 1.2
) -> str:
  """
  បម្លែងរូបភាពហត្ថលេខា/ត្រា៖
  1. ✨ Remove Background: លុបផ្ទៃក្រដាសស ឬស្រអាប់ចេញឱ្យថ្លា (Transparent PNG)
  2. 🎨 Recolor: កែប្រែពណ៌ទឹកប៊ិច (ខៀវផ្លូវការ, ខ្មៅដិត, ត្រាក្រហម, ឬពណ៌តាមចិត្ត)
  3. ✂️ Auto-Crop: កាត់គែមទទេជុំវិញចេញឱ្យល្មមស្អាត
  4. 📦 Return: Data URL Base64 ('data:image/png;base64,...')
  """
  if image_input is None:
    return None
  try:
    import io
    import base64
    from PIL import Image
    import numpy as np

    # ១. អានរូបភាពចូលជា PIL Image
    if isinstance(image_input, str):
      if image_input.startswith("data:image"):
        raw_b64 = image_input.split(",", 1)[1]
        img_bytes = base64.b64decode(raw_b64)
        img = Image.open(io.BytesIO(img_bytes))
      else:
        return image_input
    elif hasattr(image_input, "getvalue"):
      img = Image.open(io.BytesIO(image_input.getvalue()))
    elif isinstance(image_input, bytes):
      img = Image.open(io.BytesIO(image_input))
    elif isinstance(image_input, Image.Image):
      img = image_input
    else:
      return None

    # បម្លែងជា RGBA
    img = img.convert("RGBA")
    arr = np.array(img, dtype=np.float32)
    r, g, b, a = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2], arr[:, :, 3]

    # គណនាពន្លឺ (Luminance) របស់ pixel
    lum = 0.299 * r + 0.587 * g + 0.114 * b

    # ២. លុបផ្ទៃខាងក្រោយ (Remove Background)
    if remove_bg:
      t_high = float(threshold)
      t_low = max(0.0, t_high - float(softness))
      denom = max(1.0, t_high - t_low)
      # ផ្ទៃក្រដាស (lum >= t_high) => alpha 0; ស្នាមប៊ិច (lum <= t_low) => alpha 255
      alpha_bg = np.clip((t_high - lum) / denom * 255.0, 0.0, 255.0)
      final_alpha = np.minimum(a, alpha_bg)
    else:
      final_alpha = a

    # ៣. កែប្រែពណ៌ទឹកប៊ិច (Recolor)
    palette = {
        "blue": (11, 60, 149),     # #0B3C95 Official Royal Blue
        "black": (17, 24, 39),     # #111827 Deep Black
        "red": (220, 38, 38)       # #DC2626 Official Stamp Red
    }

    if recolor_choice in palette:
      tr, tg, tb = palette[recolor_choice]
      out_r = np.full_like(r, tr)
      out_g = np.full_like(g, tg)
      out_b = np.full_like(b, tb)
    elif recolor_choice == "custom":
      c = str(custom_hex).lstrip("#")
      if len(c) == 6:
        tr, tg, tb = tuple(int(c[i:i+2], 16) for i in (0, 2, 4))
      else:
        tr, tg, tb = (11, 60, 149)
      out_r = np.full_like(r, tr)
      out_g = np.full_like(g, tg)
      out_b = np.full_like(b, tb)
    else:  # "keep" ពណ៌ដើម
      out_r = np.clip(r * contrast_boost, 0.0, 255.0)
      out_g = np.clip(g * contrast_boost, 0.0, 255.0)
      out_b = np.clip(b * contrast_boost, 0.0, 255.0)

    # ផ្គុំរូបភាពឡើងវិញ
    out_arr = np.dstack([out_r, out_g, out_b, final_alpha]).astype(np.uint8)
    out_img = Image.fromarray(out_arr, mode="RGBA")

    # ៤. កាត់គែមទទេ (Auto-crop transparent bounding box)
    bbox = out_img.getbbox()
    if bbox:
      w, h = out_img.size
      box = (max(0, bbox[0] - 6), max(0, bbox[1] - 6), min(w, bbox[2] + 6), min(h, bbox[3] + 6))
      out_img = out_img.crop(box)

    # រក្សាទុកជា Base64 PNG
    buf = io.BytesIO()
    out_img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    return f"data:image/png;base64,{b64}"
  except Exception as e:
    return None


def render_signature_uploader_with_tools(
    label: str,
    key_prefix: str,
    default_sig_b64=None,
    allow_use_saved: bool = False,
    saved_sig_b64=None,
    allow_blank_choice: bool = True,
    default_blank: bool = False,
    default_recolor: str = "blue"
):
  """
  Component បង្ហាញ UI សម្រាប់ Upload ហត្ថលេខា ជាមួយមុខងារ៖
  - 📤 File uploader
  - ✨ Remove background (លុបផ្ទៃខាងក្រោយ ថ្លា)
  - 🎨 Recolor ទឹកប៊ិច (ខៀវផ្លូវការ, ខ្មៅ, ក្រហម, ពណ៌តាមចិត្ត, ពណ៌ដើម)
  - 🎚️ Slider កម្រិតសម្អាតក្រដាសស
  - 👁️ Live preview
  """
  st.markdown(f"**{label}**")

  use_saved = False
  if allow_use_saved and saved_sig_b64 and str(saved_sig_b64).startswith("data:image"):
    use_saved = st.checkbox("ប្រើហត្ថលេខាដែលបានរក្សាទុក", value=True, key=f"{key_prefix}_chk_saved")

  up_file = st.file_uploader(
      "📤 បញ្ចូលរូបភាព (PNG/JPG/WebP)",
      type=["png", "jpg", "jpeg", "webp"],
      key=f"{key_prefix}_up_file",
      help="អាចបញ្ចូលរូបថតហត្ថលេខាលើក្រដាស ប្រព័ន្ធនឹងលុបផ្ទៃស និងកែពណ៌ស្វ័យប្រវត្តិ"
  )

  is_blank = False
  if allow_blank_choice:
    is_blank = st.checkbox("ទុកចន្លោះចុចៗ (.........) ស៊ីញ៉េដៃ", value=default_blank, key=f"{key_prefix}_blank")

  if is_blank:
    st.caption("ℹ️ បង្ហាញបន្ទាត់ចុចៗ ................. សម្រាប់ចុះហត្ថលេខាផ្ទាល់ដៃ")
    return None

  # កំណត់ប្រភពរូបភាព (Source Image)
  source_img = None
  if up_file is not None:
    source_img = up_file
  elif use_saved and saved_sig_b64:
    source_img = saved_sig_b64
  elif default_sig_b64:
    source_img = default_sig_b64

  if source_img is None:
    st.caption("ℹ️ បង្ហាញបន្ទាត់ចុចៗ ................. សម្រាប់ចុះហត្ថលេខាផ្ទាល់ដៃ")
    return None

  # បង្ហាញឧបករណ៍កែសម្រួល (Remove Background & Recolor)
  with st.expander("🛠️ ឧបករណ៍លុបផ្ទៃក្រោយ & កែប្រែពណ៌ហត្ថលេខា", expanded=(up_file is not None)):
    c1, c2 = st.columns(2)
    with c1:
      rm_bg = st.checkbox(
          "✨ លុបផ្ទៃខាងក្រោយ (ថ្លា)",
          value=True,
          key=f"{key_prefix}_rm_bg",
          help="លុបផ្ទៃក្រដាសស ឬស្រអាប់ចេញឱ្យថ្លា (Transparent)"
      )
      thresh = 215
      if rm_bg:
        thresh = st.slider(
            "កម្រិតសម្អាតក្រដាសស",
            min_value=140,
            max_value=250,
            value=215,
            step=5,
            key=f"{key_prefix}_thresh",
            help="លេខកាន់តែធំ សម្អាតផ្ទៃសកាន់តែជ្រៅ"
        )
    with c2:
      color_mode = st.selectbox(
          "🎨 ពណ៌ទឹកប៊ិច / ត្រា",
          options=["blue", "black", "red", "keep", "custom"],
          index=0 if default_recolor == "blue" else (2 if default_recolor == "red" else 1),
          format_func=lambda x: {
              "blue": "🔵 ទឹកប៊ិចខៀវ (ផ្លូវការ)",
              "black": "⚫ ទឹកប៊ិចខ្មៅដិត",
              "red": "🔴 ត្រាក្រហម / ប៊ិចក្រហម",
              "keep": "🔄 រក្សាពណ៌ដើម",
              "custom": "🎨 ជ្រើសរើសពណ៌តាមចិត្ត..."
          }.get(x, x),
          key=f"{key_prefix}_col_mode"
      )
      custom_hex = "#0B3C95"
      if color_mode == "custom":
        custom_hex = st.color_picker("ជ្រើសរើសពណ៌", value="#0B3C95", key=f"{key_prefix}_custom_hex")

  processed_sig = process_signature_image(
      source_img,
      remove_bg=rm_bg,
      threshold=thresh,
      recolor_choice=color_mode,
      custom_hex=custom_hex
  )

  if processed_sig:
    st.image(processed_sig, width=135, caption="✅ ហត្ថលេខាសកម្ម (ផ្ទៃថ្លា & ពណ៌ស្អាត)")
    return processed_sig

  return None


def format_supplier_address(sup_dict):
  """រៀបចំអាសយដ្ឋានអ្នកផ្គត់ផ្គង់ជាទម្រង់អានស្រួល: ភូមិ... ឃុំ... ស្រុក... ខេត្ត..."""
  if not sup_dict:
    return ""
  parts = []
  if sup_dict.get("village"):
    v = str(sup_dict["village"]).strip()
    if v:
      parts.append(v if v.startswith("ភូមិ") else f"ភូមិ{v}")
  if sup_dict.get("commune"):
    c = str(sup_dict["commune"]).strip()
    if c:
      parts.append(c if c.startswith("ឃុំ") or c.startswith("សង្កាត់") else f"ឃុំ{c}")
  if sup_dict.get("district"):
    d = str(sup_dict["district"]).strip()
    if d:
      parts.append(d if d.startswith("ស្រុក") or d.startswith("ខណ្ឌ") or d.startswith("ក្រុង") else f"ស្រុក{d}")
  if sup_dict.get("province"):
    p = str(sup_dict["province"]).strip()
    if p:
      parts.append(p if p.startswith("ខេត្ត") or p.startswith("រាជធានី") else f"ខេត្ត{p}")
  return " ".join(parts) if parts else sup_dict.get("address", "")

def get_provinces():
  cursor.execute(
      "SELECT DISTINCT province FROM locations WHERE province IS NOT NULL AND"
      " TRIM(province) != '' ORDER BY province"
  )
  return [r[0] for r in cursor.fetchall()]


def get_districts(province=None):
  if province and province not in ["-- ជ្រើសរើស --", "➕ វាយបញ្ចូលខេត្តថ្មី..."]:
    cursor.execute(
        "SELECT DISTINCT district FROM locations WHERE province=? AND district"
        " IS NOT NULL AND TRIM(district) != '' ORDER BY district",
        (province,),
    )
  else:
    cursor.execute(
        "SELECT DISTINCT district FROM locations WHERE district IS NOT NULL AND"
        " TRIM(district) != '' ORDER BY district"
    )
  return [r[0] for r in cursor.fetchall()]


def get_communes(province=None, district=None):
  if (
      province
      and district
      and province not in ["-- ជ្រើសរើស --", "➕ វាយបញ្ចូលខេត្តថ្មី..."]
      and district not in ["-- ជ្រើសរើស --", "➕ វាយបញ្ចូលស្រុកថ្មី..."]
  ):
    cursor.execute(
        "SELECT DISTINCT commune FROM locations WHERE province=? AND"
        " district=? AND commune IS NOT NULL AND TRIM(commune) != '' ORDER BY"
        " commune",
        (province, district),
    )
  elif district and district not in [
      "-- ជ្រើសរើស --",
      "➕ វាយបញ្ចូលស្រុកថ្មី...",
  ]:
    cursor.execute(
        "SELECT DISTINCT commune FROM locations WHERE district=? AND commune IS"
        " NOT NULL AND TRIM(commune) != '' ORDER BY commune",
        (district,),
    )
  else:
    cursor.execute(
        "SELECT DISTINCT commune FROM locations WHERE commune IS NOT NULL AND"
        " TRIM(commune) != '' ORDER BY commune"
    )
  return [r[0] for r in cursor.fetchall()]


def get_villages(commune=None):
  if commune and commune not in ["-- ជ្រើសរើស --", "➕ វាយបញ្ចូលឃុំថ្មី..."]:
    cursor.execute(
        "SELECT DISTINCT village FROM locations WHERE commune=? AND village IS"
        " NOT NULL AND TRIM(village) != '' ORDER BY village",
        (commune,),
    )
  else:
    cursor.execute(
        "SELECT DISTINCT village FROM locations WHERE village IS NOT NULL AND"
        " TRIM(village) != '' ORDER BY village"
    )
  return [r[0] for r in cursor.fetchall()]


def get_filtered_schools(province=None, district=None, commune=None):
  """ទាញយកបញ្ជីសាលារៀនចម្រាញ់តាម ខេត្ត ស្រុក ឃុំ"""
  query = "SELECT DISTINCT name FROM schools WHERE TRIM(name) != ''"
  params = []
  if province and province not in ["-- ទាំងអស់ --", "-- ជ្រើសរើស --", "➕ វាយបញ្ចូលខេត្តថ្មី..."]:
    query += " AND province=?"
    params.append(province)
  if district and district not in ["-- ទាំងអស់ --", "-- ជ្រើសរើស --", "➕ វាយបញ្ចូលស្រុកថ្មី..."]:
    query += " AND district=?"
    params.append(district)
  if commune and commune not in ["-- ទាំងអស់ --", "-- ជ្រើសរើស --", "-- ជ្រើសរើសឃុំ --", "➕ វាយបញ្ចូលឃុំថ្មី..."]:
    query += " AND commune=?"
    params.append(commune)
  query += " ORDER BY name"
  cursor.execute(query, tuple(params))
  return [r[0] for r in cursor.fetchall()]


def get_schools_by_commune(commune=None):
  if commune and commune not in ["-- ជ្រើសរើស --", "➕ វាយបញ្ចូលឃុំថ្មី..."]:
    cursor.execute(
        "SELECT DISTINCT name FROM schools WHERE commune=? AND TRIM(name) !=''"
        " ORDER BY name",
        (commune,),
    )
  else:
    cursor.execute(
        "SELECT DISTINCT name FROM schools WHERE TRIM(name) != '' ORDER BY name"
    )
  return [r[0] for r in cursor.fetchall()]


def get_all_schools():
  cursor.execute(
      "SELECT DISTINCT name FROM schools WHERE TRIM(name) != '' ORDER BY name"
  )
  return [r[0] for r in cursor.fetchall()]


def save_location(province, district, commune, village=""):
  p = province.strip() if province else ""
  d = district.strip() if district else ""
  c = commune.strip() if commune else ""
  v = village.strip() if village else ""
  if not p or not d or not c:
    return False, "សូមបំពេញ ខេត្ត ស្រុក និងឃុំ ឱ្យបានគ្រប់គ្រាន់!"
  exists = cursor.execute(
      "SELECT id FROM locations WHERE province=? AND district=? AND commune=?"
      " AND (village=? OR (village IS NULL AND ?=''))",
      (p, d, c, v, v),
  ).fetchone()
  if not exists:
    cursor.execute(
        "INSERT INTO locations (province, district, commune, village) VALUES"
        " (?,?,?,?)",
        (p, d, c, v),
    )
    conn.commit()
  return True, "បានរក្សាទុកទីតាំងជោគជ័យ!"


def save_school(name, commune, district="", province="", village=""):
  n = name.strip() if name else ""
  c = commune.strip() if commune else ""
  d = district.strip() if district else ""
  p = province.strip() if province else ""
  v = village.strip() if village else ""
  if not n or not c:
    return False, "សូមបំពេញ ឈ្មោះសាលា និង ឃុំ/សង្កាត់!"
  exists = cursor.execute(
      "SELECT id FROM schools WHERE name=? AND commune=?", (n, c)
  ).fetchone()
  if not exists:
    cursor.execute(
        "INSERT INTO schools (name, commune, district, province, village)"
        " VALUES (?,?,?,?,?)",
        (n, c, d, p, v),
    )
    conn.commit()
  return True, "បានរក្សាទុកសាលារៀនជោគជ័យ!"


# ================= Helper Functions សម្រាប់កាលបរិច្ឆេទ និងវគ្គនីមួយៗ =================
def parse_date_safe(d_val, default_date=None):
  """បំប្លែងទិន្នន័យកាលបរិច្ឆេទទៅជា datetime.date ដោយសុវត្ថិភាព"""
  if isinstance(d_val, date):
    return d_val
  if not d_val:
    return default_date or date.today()
  try:
    parts = str(d_val).strip()[:10].split("-")
    if len(parts) == 3:
      return date(int(parts[0]), int(parts[1]), int(parts[2]))
  except Exception:
    pass
  return default_date or date.today()


def get_commune_phase_dates(commune=None, school_name=None):
  """ទាញយកកាលបរិច្ឆេទវគ្គ១ និងវគ្គ២ លំនាំដើមពីទំនិញក្នុងសាលា ឬឃុំ"""
  cur_year = date.today().year
  def_p1_s = f"{cur_year}-01-01"
  def_p1_e = f"{cur_year}-06-30"
  def_p2_s = f"{cur_year}-07-01"
  def_p2_e = f"{cur_year}-12-31"

  if school_name and school_name not in ["-- ជ្រើសរើសសាលា --", "➕ វាយបញ្ចូលសាលាថ្មី..."]:
    row = cursor.execute(
        """SELECT phase1_start, phase1_end, phase2_start, phase2_end 
           FROM products 
           WHERE school_name=? AND price_level='school' AND phase1_start IS NOT NULL AND TRIM(phase1_start) != ''
           ORDER BY id DESC LIMIT 1""",
        (school_name,),
    ).fetchone()
    if row and row[0] and row[1]:
      return row[0], row[1], (row[2] or def_p2_s), (row[3] or def_p2_e)

  if commune and commune not in ["-- ជ្រើសរើសឃុំ --", "-- ទាំងអស់ --", "➕ វាយបញ្ចូលឃុំថ្មី..."]:
    row = cursor.execute(
        """SELECT phase1_start, phase1_end, phase2_start, phase2_end 
           FROM products 
           WHERE commune=? AND phase1_start IS NOT NULL AND TRIM(phase1_start) != ''
           ORDER BY id DESC LIMIT 1""",
        (commune,),
    ).fetchone()
    if row and row[0] and row[1]:
      return row[0], row[1], (row[2] or def_p2_s), (row[3] or def_p2_e)

  # Check any product in the system
  row_any = cursor.execute(
      """SELECT phase1_start, phase1_end, phase2_start, phase2_end 
         FROM products 
         WHERE phase1_start IS NOT NULL AND TRIM(phase1_start) != ''
         ORDER BY id DESC LIMIT 1"""
  ).fetchone()
  if row_any and row_any[0] and row_any[1]:
    return row_any[0], row_any[1], (row_any[2] or def_p2_s), (row_any[3] or def_p2_e)

  return def_p1_s, def_p1_e, def_p2_s, def_p2_e


def get_products_map(school_name=None, commune=None, supplier_name=None):
  """
  ទាញយកបញ្ជីមុខទំនិញ និងតម្លៃសម្រាប់សាលា ឃុំ ឬអ្នកផ្គត់ផ្គង់៖
  ១. យកតម្លៃកំណត់តាមឃុំ (Commune Level) ជាគោល
  ២. ប្រសិនបើសាលានោះមានកំណត់តម្លៃដោយឡែក (School Level) យកតម្លៃសាលាមកជំនួស (Override)
  ៣. ប្រសិនបើមានកំណត់តាមអ្នកផ្គត់ផ្គង់ (Supplier Level) យកមកជំនួស/បន្ថែម
  ៤. បំពេញទំនិញទាំងអស់ពីកាតាឡុកស្ដង់ដារ (STANDARD_PRODUCT_CATALOG) ដើម្បីកុំឲ្យបាត់មុខទំនិញ
  """
  res_map = {}
  
  # ១. ទាញយកតាមឃុំ
  if commune and commune not in ["-- ជ្រើសរើសឃុំ --", "-- ទាំងអស់ --", "➕ វាយបញ្ចូលឃុំថ្មី..."]:
    comm_rows = cursor.execute(
        """SELECT item_name, price_phase1, price_phase2, price_avg, 
                  phase1_start, phase1_end, phase2_start, phase2_end, category 
           FROM products 
           WHERE commune=? AND (school_name IS NULL OR school_name = '' OR price_level='commune')""",
        (commune,)
    ).fetchall()
    for r in comm_rows:
      res_map[r[0]] = {
          "price_phase1": float(r[1] or 0),
          "price_phase2": float(r[2] or 0),
          "price_avg": float(r[3] or 0),
          "p1_start": r[4], "p1_end": r[5],
          "p2_start": r[6], "p2_end": r[7],
          "category": r[8] or classify_item_category(r[0]),
          "scope": "ឃុំ"
      }

  # ២. បើមានសាលា ហើយសាលានោះមានកំណត់តម្លៃដោយឡែក (School Level)
  if school_name and school_name not in ["-- ជ្រើសរើសសាលា --", "➕ វាយបញ្ចូលសាលាថ្មី..."]:
    sch_rows = cursor.execute(
        """SELECT item_name, price_phase1, price_phase2, price_avg, 
                  phase1_start, phase1_end, phase2_start, phase2_end, category 
           FROM products 
           WHERE school_name=? AND price_level='school'""",
        (school_name,)
    ).fetchall()
    for r in sch_rows:
      res_map[r[0]] = {
          "price_phase1": float(r[1] or 0),
          "price_phase2": float(r[2] or 0),
          "price_avg": float(r[3] or 0),
          "p1_start": r[4], "p1_end": r[5],
          "p2_start": r[6], "p2_end": r[7],
          "category": r[8] or classify_item_category(r[0]),
          "scope": "សាលារៀន"
      }

  # ៣. ប្រសិនបើមានកំណត់តាមអ្នកផ្គត់ផ្គង់ (Supplier Level)
  if supplier_name and supplier_name not in ["-- ជ្រើសរើសអ្នកផ្គត់ផ្គង់ --", ""]:
    try:
      sup_rows = cursor.execute(
          """SELECT item_name, price_phase1, price_phase2, price_avg, 
                    phase1_start, phase1_end, phase2_start, phase2_end, category 
             FROM products 
             WHERE supplier_name=? AND (
                 (school_name=? AND school_name IS NOT NULL AND school_name != '') OR 
                 (commune=? AND commune IS NOT NULL AND commune != '') OR
                 price_level='supplier' OR school_name IS NULL OR school_name=''
             )""",
          (supplier_name, school_name or "", commune or "")
      ).fetchall()
      for r in sup_rows:
        res_map[r[0]] = {
            "price_phase1": float(r[1] or 0),
            "price_phase2": float(r[2] or 0),
            "price_avg": float(r[3] or 0),
            "p1_start": r[4], "p1_end": r[5],
            "p2_start": r[6], "p2_end": r[7],
            "category": r[8] or classify_item_category(r[0]),
            "scope": "អ្នកផ្គត់ផ្គង់"
        }
    except Exception:
      pass

  # ៤. បំពេញទំនិញដែលមានស្រាប់ក្នុងប្រព័ន្ធបើទទេ
  if not res_map:
    all_rows = cursor.execute(
        """SELECT item_name, price_phase1, price_phase2, price_avg, 
                  phase1_start, phase1_end, phase2_start, phase2_end, category 
           FROM products GROUP BY item_name"""
    ).fetchall()
    for r in all_rows:
      res_map[r[0]] = {
          "price_phase1": float(r[1] or 0),
          "price_phase2": float(r[2] or 0),
          "price_avg": float(r[3] or 0),
          "p1_start": r[4], "p1_end": r[5],
          "p2_start": r[6], "p2_end": r[7],
          "category": r[8] or classify_item_category(r[0]),
          "scope": "ប្រព័ន្ធ"
      }

  # ៥. បំពេញគ្រប់មុខទំនិញដែលនៅសល់ពីកាតាឡុកស្ដង់ដារ (STANDARD_PRODUCT_CATALOG)
  cur_y = date.today().year
  for item in STANDARD_PRODUCT_CATALOG:
    if item["name"] not in res_map or res_map[item["name"]]["price_phase1"] <= 0:
      res_map[item["name"]] = {
          "price_phase1": item["price_phase1"],
          "price_phase2": item["price_phase2"],
          "price_avg": (item["price_phase1"] + item["price_phase2"]) / 2.0,
          "p1_start": f"{cur_y}-01-01", "p1_end": f"{cur_y}-06-30",
          "p2_start": f"{cur_y}-07-01", "p2_end": f"{cur_y}-12-31",
          "category": item["category"],
          "scope": "កាតាឡុកស្ដង់ដារ"
      }

  return res_map


def detect_phase_from_date(check_date, p1_start=None, p1_end=None, p2_start=None, p2_end=None):
  """
  ពិនិត្យស្វ័យប្រវត្តថាកាលបរិច្ឆេទត្រូវនឹង វគ្គ១ ឬ វគ្គ២
  ត្រឡប់: (phase_name, reason_description)
  """
  if isinstance(check_date, date):
    d_str = str(check_date)
  else:
    d_str = str(check_date).strip()[:10]

  if p1_start and p1_end and (p1_start <= d_str <= p1_end):
    return "វគ្គ១", f"កាលបរិច្ឆេទ {d_str} ស្ថិតក្នុងចន្លោះវគ្គ១ ({p1_start} ដល់ {p1_end})"
  if p2_start and p2_end and (p2_start <= d_str <= p2_end):
    return "វគ្គ២", f"កាលបរិច្ឆេទ {d_str} ស្ថិតក្នុងចន្លោះវគ្គ២ ({p2_start} ដល់ {p2_end})"

  # Partial comparisons
  if p1_start and d_str >= p1_start and (not p1_end or d_str <= p1_end):
    return "វគ្គ១", f"កាលបរិច្ឆេទ {d_str} ត្រូវនឹងវគ្គ១ (ចាប់ពី {p1_start})"
  if p2_start and d_str >= p2_start and (not p2_end or d_str <= p2_end):
    return "វគ្គ២", f"កាលបរិច្ឆេទ {d_str} ត្រូវនឹងវគ្គ២ (ចាប់ពី {p2_start})"

  # Default semester by month
  try:
    month = int(d_str.split("-")[1])
    if 1 <= month <= 6:
      return "វគ្គ១", f"កាលបរិច្ឆេទ {d_str} ស្ថិតក្នុងឆមាសទី១ (ស្វ័យប្រវត្ត: វគ្គ១)"
    else:
      return "វគ្គ២", f"កាលបរិច្ឆេទ {d_str} ស្ថិតក្នុងឆមាសទី២ (ស្វ័យប្រវត្ត: វគ្គ២)"
  except Exception:
    return "វគ្គ១", "លំនាំដើម: វគ្គ១"


# ================= មុខងារ Universal File Importer & OCR =================
def extract_table_from_file(uploaded_file):
  """ស្រង់ទិន្នន័យពីឯកសារជាច្រើនប្រភេទ: xlsx, docx, csv, pdf, png, jpg..."""
  filename = uploaded_file.name.lower()
  file_bytes = uploaded_file.getvalue()

  # ១. Excel (.xlsx, .xls)
  if filename.endswith((".xlsx", ".xls")):
    try:
      return pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
      st.error(f"កំហុសក្នុងការអានឯកសារ Excel: {e}")
      return pd.DataFrame()

  # ២. CSV (.csv)
  elif filename.endswith(".csv"):
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252"]:
      try:
        return pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
      except Exception:
        continue
    return pd.DataFrame()

  # ៣. Word (.docx, .doc)
  elif filename.endswith((".docx", ".doc")):
    try:
      doc = Document(io.BytesIO(file_bytes))
      tables_data = []
      for tbl in doc.tables:
        for row in tbl.rows:
          r_data = [cell.text.strip() for cell in row.cells]
          if any(r_data):
            tables_data.append(r_data)
      if tables_data:
        max_c = max(len(r) for r in tables_data)
        tables_data = [r + [""] * (max_c - len(r)) for r in tables_data]
        if len(tables_data) > 1:
          return pd.DataFrame(tables_data[1:], columns=tables_data[0])
        else:
          return pd.DataFrame(tables_data)
      else:
        # Paragraphs lines
        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        rows = [
            [part.strip() for part in re.split(r"[,|\t]", line) if part.strip()]
            for line in lines
        ]
        if rows:
          max_c = max(len(r) for r in rows)
          rows = [r + [""] * (max_c - len(r)) for r in rows]
          return (
              pd.DataFrame(rows[1:], columns=rows[0])
              if len(rows) > 1
              else pd.DataFrame(rows)
          )
    except Exception as e:
      st.error(f"កំហុសក្នុងការអានឯកសារ Word: {e}")
      return pd.DataFrame()

  # ៤. PDF (.pdf)
  elif filename.endswith(".pdf"):
    try:
      all_rows = []
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
                parts = [p.strip() for p in re.split(r"[,|\t]", line) if p.strip()]
                if not parts:
                  parts = [p.strip() for p in line.split() if p.strip()]
                if parts:
                  all_rows.append(parts)
      if all_rows:
        max_c = max(len(r) for r in all_rows)
        all_rows = [r + [""] * (max_c - len(r)) for r in all_rows]
        return (
            pd.DataFrame(all_rows[1:], columns=all_rows[0])
            if len(all_rows) > 1
            else pd.DataFrame(all_rows)
        )
    except Exception as e:
      st.error(f"កំហុសក្នុងការអានឯកសារ PDF: {e}")
      return pd.DataFrame()

  # ៥. រូបភាព (Image OCR: .png, .jpg, .jpeg, .webp, .bmp)
  elif filename.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp")):
    try:
      img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
      img_np = np.array(img)
      ocr = RapidOCR()
      res, _ = ocr(img_np)
      if res:
        items = []
        for box, text, score in res:
          y_center = (box[0][1] + box[2][1]) / 2.0
          x_left = box[0][0]
          items.append({"text": text.strip(), "x": x_left, "y": y_center})

        # តម្រៀបតាមបន្ទាត់ផ្ដេក Y
        items.sort(key=lambda it: it["y"])
        rows = []
        cur_row = []
        cur_y = None
        for it in items:
          if cur_y is None or abs(it["y"] - cur_y) < 22:
            cur_row.append(it)
            cur_y = it["y"]
          else:
            cur_row.sort(key=lambda x: x["x"])
            rows.append([x["text"] for x in cur_row])
            cur_row = [it]
            cur_y = it["y"]
        if cur_row:
          cur_row.sort(key=lambda x: x["x"])
          rows.append([x["text"] for x in cur_row])

        if rows:
          max_c = max(len(r) for r in rows)
          rows = [r + [""] * (max_c - len(r)) for r in rows]
          return (
              pd.DataFrame(rows[1:], columns=rows[0])
              if len(rows) > 1
              else pd.DataFrame(rows)
          )
    except Exception as e:
      st.error(f"កំហុសក្នុងការអានរូបភាព (OCR): {e}")
      return pd.DataFrame()

  return pd.DataFrame()


# មុខងារបង្កើតគំរូ Excel សម្រាប់ Download
def generate_sample_excel(columns, sample_row):
  df_sample = pd.DataFrame([sample_row], columns=columns)
  out = io.BytesIO()
  with pd.ExcelWriter(out, engine="openpyxl") as writer:
    df_sample.to_excel(writer, index=False, sheet_name="Template")
  return out.getvalue()


# មុខងារ Export ជា Excel (មានលេខរៀង ល.រ ពីលេខ ១ ដល់ N)
def to_excel(df: pd.DataFrame) -> bytes:
  output = io.BytesIO()
  df_export = add_row_numbers(df)
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_export.to_excel(writer, index=False, sheet_name="Report")
  return output.getvalue()


# មុខងារ Export ជា PDF សាមញ្ញ (មានលេខរៀង No. ពីលេខ ១ ដល់ N)
def generate_simple_pdf(title: str, subtitle: str, df: pd.DataFrame) -> bytes:
  import os
  df_pdf = add_row_numbers(df)
  pdf = FPDF()
  pdf.add_page()

  khmer_font_path = find_khmer_font()
  has_khmer_font = khmer_font_path is not None

  if has_khmer_font:
    pdf.add_font("KhmerOS", fname=khmer_font_path)
    font_main = "KhmerOS"
  else:
    font_main = "Helvetica"

  pdf.set_font(font_main, size=16 if not has_khmer_font else 14)
  safe_title = title if has_khmer_font else title.encode("latin-1", "replace").decode("latin-1")
  pdf.cell(0, 10, safe_title, ln=True, align="C")

  pdf.set_font(font_main, size=11 if not has_khmer_font else 10)
  safe_subtitle = subtitle if has_khmer_font else subtitle.encode("latin-1", "replace").decode("latin-1")
  pdf.cell(0, 8, safe_subtitle, ln=True, align="C")
  pdf.ln(6)

  # Header តារាង
  pdf.set_font(font_main, size=9 if not has_khmer_font else 9)
  col_widths = [180 / len(df_pdf.columns)] * len(df_pdf.columns)
  for i, col in enumerate(df_pdf.columns):
    col_str = str(col) if has_khmer_font else str(col).replace("៛", "Riel").replace("ល.រ", "No.")
    if not has_khmer_font:
      col_str = col_str.encode("latin-1", "replace").decode("latin-1")
    pdf.cell(
        col_widths[i],
        8,
        col_str,
        border=1,
        align="C",
    )
  pdf.ln()

  # Data rows
  pdf.set_font(font_main, size=8 if not has_khmer_font else 8)
  for _, row in df_pdf.iterrows():
    for i, (col, val) in enumerate(row.items()):
      if col == "ល.រ" or col == "No.":
        text_val = str(val)
      elif isinstance(val, (int, float)):
        text_val = f"{int(val):,} ៛" if val % 1 == 0 else f"{val:,.2f} ៛"
        if not has_khmer_font:
          text_val = text_val.replace("៛", "Riel")
      else:
        text_val = str(val)
        if not has_khmer_font:
          text_val = (
              text_val.replace("៛", "Riel")
              .encode("latin-1", "replace")
              .decode("latin-1")
          )
      pdf.cell(col_widths[i], 7, text_val, border=1, align="C")
    pdf.ln()

  return bytes(pdf.output())


# ================= មុខងារបង្កើតប័ណ្ណទទួលស្បៀង (វិក្កយបត្រ ឧបសម្ពន្ធ ៣) =================
def to_khmer_num(val):
  """បំប្លែងលេខអារ៉ាប់ទៅជាលេខខ្មែរ (០-៩)"""
  khmer_digits = {'0': '០', '1': '១', '2': '២', '3': '៣', '4': '៤', '5': '៥', '6': '៦', '7': '៧', '8': '៨', '9': '៩'}
  return "".join(khmer_digits.get(c, c) for c in str(val))


def format_khmer_date(d_val):
  """បំប្លែងកាលបរិច្ឆេទទៅជាទម្រង់ខ្មែរ: ថ្ងៃទី DD ខែ MM ឆ្នាំ YYYY"""
  if not d_val:
    return ""
  if isinstance(d_val, str):
    parts = d_val.strip()[:10].split("-")
    if len(parts) == 3:
      y, m, d = parts[0], parts[1], parts[2]
    else:
      return d_val
  else:
    y, m, d = f"{d_val.year}", f"{d_val.month:02d}", f"{d_val.day:02d}"
  return f"ថ្ងៃទី {to_khmer_num(d)} ខែ {to_khmer_num(m)} ឆ្នាំ {to_khmer_num(y)}"


def get_school_location_info(school_name):
  """ទាញយកព័ត៌មាន ខេត្ត ស្រុក ឃុំ ភូមិ របស់សាលារៀន"""
  if not school_name:
    return "", "", "", ""
  row = cursor.execute(
      "SELECT province, district, commune, village FROM schools WHERE name=? LIMIT 1",
      (school_name,)
  ).fetchone()
  if row:
    p, d, c, v = row[0] or "", row[1] or "", row[2] or "", row[3] or ""
    if (not d or not p) and c:
      loc_row = cursor.execute(
          "SELECT province, district FROM locations WHERE commune=? LIMIT 1",
          (c,)
      ).fetchone()
      if loc_row:
        p = p or loc_row[0] or ""
        d = d or loc_row[1] or ""
    return p, d, c, v
  return "", "", "", ""


def get_or_create_voucher_no(school_name, date_str):
  """
  ទាញយក ឬបង្កើតលេខសក្ខីប័ត្រស្វ័យប្រវត្ត ចាប់ផ្ដើមពី 001 អាស្រ័យតាមសាលានីមួយៗ
  """
  if not school_name:
    return "001"
  d_str = str(date_str).strip()[:10]
  # ១. ពិនិត្យមើលថាតើថ្ងៃនេះ និងសាលានេះ មានលេខសក្ខីប័ត្ររួចហើយឬនៅ
  row = cursor.execute(
      """SELECT voucher_no FROM daily_records 
         WHERE school_name=? AND date=? AND voucher_no IS NOT NULL AND TRIM(voucher_no) != '' 
         LIMIT 1""",
      (school_name, d_str)
  ).fetchone()
  if row and row[0]:
    return str(row[0]).strip()

  # ២. បើមិនទាន់មាន រកលេខសក្ខីប័ត្រធំបំផុតរបស់សាលានោះ រួចបូក ១ (ចាប់ពី 001)
  rows = cursor.execute(
      """SELECT DISTINCT voucher_no FROM daily_records 
         WHERE school_name=? AND voucher_no IS NOT NULL AND TRIM(voucher_no) != ''""",
      (school_name,)
  ).fetchall()
  nums = []
  for r in rows:
    val = str(r[0]).strip()
    matches = re.findall(r'\d+', val)
    if matches:
      nums.append(int(matches[-1]))

  next_val = max(nums) + 1 if nums else 1
  return f"{next_val:03d}"


def save_voucher_no(school_name, date_str, voucher_no):
  """រក្សាទុក ឬកែប្រែលេខសក្ខីប័ត្រសម្រាប់កំណត់ត្រានៅថ្ងៃនោះរបស់សាលា"""
  if not school_name or not date_str:
    return
  cursor.execute(
      "UPDATE daily_records SET voucher_no=? WHERE school_name=? AND date=?",
      (str(voucher_no).strip(), school_name, str(date_str).strip()[:10])
  )
  conn.commit()


def generate_annex3_html(district, commune, school_name, voucher_no, invoice_date, df_items,
                        supplier_name="សាត ក្រូត", comment="",
                        supplier_sig=None, director_sig=None, receiver_sig=None,
                        consumption_date=None):
  """បង្កើតកូដ HTML តាមគំរូផ្លូវការ ឧបសម្ពន្ធ ៣ (ប័ណ្ណទទួលស្បៀង) សម្រាប់ទំហំក្រដាស A5 ដោយគ្មានហត្ថលេខាក្លែងក្លាយ និងគាំទ្រការបញ្ចូលហត្ថលេខាពិត"""
  if isinstance(df_items, pd.DataFrame):
    items_list = df_items.to_dict('records') if not df_items.empty else []
  elif isinstance(df_items, list):
    items_list = df_items
  else:
    items_list = []

  tot_qty = 0.0
  tot_amount = 0.0
  kh_digits = ['១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩', '១០']

  rows_html = ""
  for r_idx in range(10):
    k_num = kh_digits[r_idx]
    if r_idx < len(items_list):
      item = items_list[r_idx]
      i_name = str(item.get("មុខទំនិញ", item.get("item_name", item.get("name", "")))).strip()
      try:
        i_qty = float(item.get("បរិមាណ", item.get("quantity", item.get("qty", 0))) or 0)
      except Exception:
        i_qty = 0.0
      try:
        i_unit = float(item.get("តម្លៃរាយ (៛)", item.get("unit_price", 0)) or 0)
      except Exception:
        i_unit = 0.0
      try:
        i_tot = float(item.get("សរុប (៛)", item.get("total_price", i_qty * i_unit)) or (i_qty * i_unit))
      except Exception:
        i_tot = 0.0

      tot_qty += i_qty
      tot_amount += i_tot

      qty_str = f"{i_qty:g}" if i_qty % 1 != 0 else f"{int(i_qty)}"
      unit_str = f"{int(i_unit):,} ៛" if i_unit % 1 == 0 else f"{i_unit:,.2f} ៛"
      tot_str = f"{int(i_tot):,} ៛" if i_tot % 1 == 0 else f"{i_tot:,.2f} ៛"

      rows_html += f"""
      <tr>
        <td class="col-num">{k_num}</td>
        <td class="col-desc">{i_name}</td>
        <td class="col-qty">{qty_str}</td>
        <td class="col-price">{unit_str}</td>
        <td class="col-total">{tot_str}</td>
      </tr>
      """
    else:
      rows_html += f"""
      <tr class="empty-row">
        <td class="col-num">{k_num}</td>
        <td class="col-desc"></td>
        <td class="col-qty"></td>
        <td class="col-price"></td>
        <td class="col-total"></td>
      </tr>
      """

  tot_qty_str = f"{tot_qty:g}" if tot_qty % 1 != 0 else f"{int(tot_qty)}"
  tot_amount_str = f"{int(tot_amount):,} ៛" if tot_amount % 1 == 0 else f"{tot_amount:,.2f} ៛"
  khmer_date_str = format_khmer_date(invoice_date)
  # បង្ហាញត្រឹមតែកាលបរិច្ឆេទខាងមុខ ដោយលុបពាក្យថ្ងៃដាក់ចេញ
  date_display_str = khmer_date_str
  v_display = f"{voucher_no}" if voucher_no else "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

  html_template = f"""<!DOCTYPE html>
<html lang="km">
<head>
<meta charset="utf-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Battambang:wght@400;700&family=Kantumruy+Pro:ital,wght@0,400;0,600;0,700;1,400&family=Moul&display=swap');
  
  @page {{
    size: A5 portrait;
    margin: 4mm 6mm 4mm 6mm;
  }}

  @media print {{
    .no-print {{
      display: none !important;
    }}
    @page {{
      size: A5 portrait;
      margin: 4mm 6mm 4mm 6mm;
    }}
    body {{
      margin: 0 !important;
      padding: 0 !important;
      width: 100% !important;
    }}
  }}
  
  * {{
    box-sizing: border-box;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }}
  
  body {{
    font-family: 'Khmer OS Battambang', 'Kantumruy Pro', 'Battambang', Arial, sans-serif;
    font-size: 10px;
    line-height: 1.2;
    color: #000;
    background: #fff;
    margin: 0;
    padding: 2px 4px;
    width: 100%;
  }}
  
  .muol {{
    font-family: 'Khmer OS Muol Light', 'Moul', 'Khmer OS Muol', serif;
    font-weight: normal;
  }}
  
  .top-bar {{
    position: relative;
    text-align: center;
    margin-bottom: 2px;
  }}
  
  .annex-tag {{
    position: absolute;
    right: 0;
    top: 0;
    font-size: 11px;
    font-weight: 700;
    color: #111;
  }}
  
  .country-title {{
    font-size: 13px;
    letter-spacing: 0.5px;
    line-height: 1.3;
  }}
  
  .motto-title {{
    font-size: 12px;
    letter-spacing: 0.5px;
    line-height: 1.3;
    margin-bottom: 3px;
  }}
  
  .meta-grid {{
    display: table;
    width: 100%;
    margin-top: 2px;
    margin-bottom: 2px;
  }}
  
  .meta-left {{
    display: table-cell;
    width: 60%;
    vertical-align: top;
  }}
  
  .meta-right {{
    display: table-cell;
    width: 40%;
    vertical-align: top;
    text-align: right;
  }}
  
  .info-table {{
    border-collapse: collapse;
  }}
  
  .info-table td {{
    padding: 1px 3px;
    font-size: 10.5px;
  }}
  
  .info-table td.lbl {{
    font-weight: bold;
    width: 90px;
    white-space: nowrap;
  }}
  
  .info-table td.val {{
    font-weight: 600;
    color: #0c4a8a;
    padding-left: 6px;
  }}
  
  .voucher-box {{
    display: inline-block;
    padding-right: 10px;
  }}
  
  .voucher-lbl {{
    font-size: 11px;
    font-weight: 700;
    margin-right: 8px;
  }}
  
  .voucher-num {{
    font-size: 13.5px;
    font-weight: 700;
    color: #0c4a8a;
    letter-spacing: 1.5px;
    min-width: 40px;
    display: inline-block;
    text-align: right;
  }}
  
  .doc-title-row {{
    text-align: center;
    margin-top: 3px;
    margin-bottom: 2px;
  }}
  
  .doc-title {{
    font-size: 15px;
    line-height: 1.3;
    display: inline-block;
  }}
  
  .date-row {{
    text-align: right;
    margin-bottom: 4px;
    padding-right: 4px;
  }}
  
  .date-text {{
    font-size: 10px;
    font-style: italic;
    color: #1a4d80;
    font-weight: 500;
  }}
  
  /* Main Table */
  .items-table {{
    width: 100%;
    border-collapse: collapse;
    border: 1.5px solid #000;
    margin-bottom: 3px;
  }}
  
  .items-table th, .items-table td {{
    border: 1px solid #000;
    padding: 2px 4px;
    font-size: 10px;
    height: 17px;
  }}
  
  .items-table th {{
    font-weight: 700;
    text-align: center;
    background: #fbfbfb;
  }}
  
  .col-num {{
    width: 7%;
    text-align: center;
    font-weight: 600;
  }}
  
  .col-desc {{
    width: 39%;
    text-align: left;
    padding-left: 6px;
    color: #0c4a8a;
    font-weight: 500;
  }}
  
  .col-qty {{
    width: 14%;
    text-align: center;
    color: #0c4a8a;
  }}
  
  .col-price {{
    width: 19%;
    text-align: right;
    padding-right: 8px;
    color: #0c4a8a;
  }}
  
  .col-total {{
    width: 21%;
    text-align: right;
    padding-right: 8px;
    color: #0c4a8a;
    font-weight: 600;
  }}
  
  .total-row td {{
    font-weight: 700;
    font-size: 10.5px;
    background-color: #fafafa;
  }}
  
  .note-text {{
    font-size: 9.5px;
    margin-top: 3px;
    margin-bottom: 4px;
    line-height: 1.3;
  }}
  
  .note-label {{
    font-weight: 700;
    margin-right: 8px;
  }}
  
  /* Signatures */
  .sig-table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 3px;
    margin-bottom: 4px;
    table-layout: fixed;
  }}
  
  .sig-table td {{
    vertical-align: top;
    text-align: center;
    padding: 0 2px;
  }}
  
  .sig-role {{
    font-weight: 700;
    font-size: 10.5px;
    line-height: 1.3;
  }}
  
  .sig-name {{
    font-weight: 700;
    font-size: 10px;
    margin-top: 1px;
  }}
  
  .sig-space {{
    height: 34px;
    display: flex;
    align-items: center;
    justify-content: center;
  }}
  
  .sig-dots {{
    letter-spacing: 1.5px;
    color: #444;
    font-size: 9px;
  }}
  
  /* Comment Box */
  .comment-box {{
    border: 1px dashed #666;
    padding: 3px 6px;
    margin-top: 3px;
    margin-bottom: 3px;
    border-radius: 2px;
  }}
  
  .comment-title {{
    font-weight: 700;
    font-size: 9.5px;
    margin-bottom: 2px;
  }}
  
  .comment-line {{
    border-bottom: 1px dotted #888;
    height: 12px;
  }}
  
  .doc-copy-mark {{
    text-align: center;
    color: #b91c1c;
    font-size: 9.5px;
    font-weight: 700;
    margin-top: 3px;
    letter-spacing: 0.5px;
  }}
</style>
</head>
<body>

<div class="no-print" style="display: flex; justify-content: flex-end; margin-bottom: 5px; padding: 2px 0;">
  <button onclick="window.print()" style="background: #0284c7; color: #fff; border: 1px solid #0284c7; padding: 5px 14px; border-radius: 4px; font-family: inherit; font-size: 11px; font-weight: bold; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.12);">
    🖨️ បោះពុម្ពប័ណ្ណទទួលស្បៀង A5 (Print)
  </button>
</div>

<div class="top-bar">
  <div class="annex-tag">ឧបសម្ពន្ធ ៣</div>
  <div class="muol country-title">ព្រះរាជាណាចក្រកម្ពុជា</div>
  <div class="muol motto-title">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
</div>

<div class="meta-grid">
  <div class="meta-left">
    <table class="info-table">
      <tr>
        <td class="lbl">ក្រុង/ស្រុក</td>
        <td class="val">{district if district else '&nbsp;'}</td>
      </tr>
      <tr>
        <td class="lbl">ឃុំ/សង្កាត់</td>
        <td class="val">{commune if commune else '&nbsp;'}</td>
      </tr>
      <tr>
        <td class="lbl">សាលាបឋមសិក្សា</td>
        <td class="val">{school_name if school_name else '&nbsp;'}</td>
      </tr>
    </table>
  </div>
  <div class="meta-right">
    <div class="voucher-box">
      <span class="voucher-lbl">លេខសក្ខីប័ត្រ ៖</span>
      <span class="voucher-num">{v_display}</span>
    </div>
  </div>
</div>

<div class="doc-title-row">
  <div class="muol doc-title">បង្កាន់ដៃទទួលស្បៀង</div>
</div>

<div class="date-row">
  <span class="date-text">{date_display_str}</span>
</div>

<table class="items-table">
  <thead>
    <tr>
      <th style="width: 6.5%;">លរ</th>
      <th style="width: 41%;">បរិយាយមុខទំនិញ</th>
      <th style="width: 15%;">បរិមាណ</th>
      <th style="width: 17.5%;">តម្លៃឯកតា</th>
      <th style="width: 20%;">តម្លៃសរុប</th>
    </tr>
  </thead>
  <tbody>
    {rows_html}
    <tr class="total-row">
      <td colspan="2" style="text-align: center; font-weight: 700;">តម្លៃសរុប </td>
      <td style="text-align: center; color: #0c4a8a;">{tot_qty_str}</td>
      <td></td>
      <td style="text-align: right; padding-right: 12px; color: #0c4a8a;">{tot_amount_str}</td>
    </tr>
  </tbody>
</table>

<div class="note-text">
  <span class="note-label">សម្គាល់៖</span>រាល់តម្លៃស្បៀងស្នើសុំទូទាត់ត្រូវស្របតាមកិច្ចសន្យាផ្គត់ផ្គង់ស្បៀង។
</div>

<table class="sig-table">
  <tr>
    <td style="width: 33.3%;">
      <div class="sig-role">បានឃើញ និងឯកភាព</div>
      <div class="sig-name">ប្រធាន គមស (នាយក/នាយិកាសាលា)</div>
      <div class="sig-space">
        {f'<img src="{director_sig}" style="max-height: 34px; max-width: 90px; object-fit: contain;" />' if director_sig else ''}
      </div>
      {f'<div class="sig-dots">...............................</div>' if not director_sig else ''}
    </td>
    <td style="width: 33.3%;">
      <div class="sig-role">អ្នកទទួល</div>
      <div class="sig-name">នាយឃ្លាំងឬបេឡាធិការ</div>
      <div class="sig-space">
        {f'<img src="{receiver_sig}" style="max-height: 34px; max-width: 90px; object-fit: contain;" />' if receiver_sig else ''}
      </div>
      {f'<div class="sig-dots">...............................</div>' if not receiver_sig else ''}
    </td>
    <td style="width: 33.3%;">
      <div class="sig-role">អ្នកប្រគល់</div>
      <div class="sig-name">អ្នកផ្គត់ផ្គង់ស្បៀង</div>
      <div class="sig-space">
        {f'<img src="{supplier_sig}" style="max-height: 34px; max-width: 90px; object-fit: contain;" />' if supplier_sig else ''}
      </div>
      {f'<div class="sig-dots">...............................</div>' if not supplier_sig else ''}
      <div class="sig-name" style="color: #0c4a8a; font-weight: bold; margin-top: 1px;">{supplier_name}</div>
    </td>
  </tr>
</table>

<div class="comment-box">
  <div class="comment-title">យោបល់ចំពោះទំនិញ៖ <span style="font-weight: normal;">{comment}</span></div>
  <div class="comment-line"></div>
  <div class="comment-line"></div>
</div>

<div class="doc-copy-mark">ច្បាប់ដើមសម្រាប់សាលា</div>

</body>
</html>
"""
  return html_template


def generate_annex3_pdf(district, commune, school_name, voucher_no, invoice_date, df_items,
                       supplier_name="សាត ក្រូត", comment="",
                       supplier_sig=None, director_sig=None, receiver_sig=None,
                       consumption_date=None):
  """បង្កើតឯកសារ PDF ផ្លូវការ (ប័ណ្ណទទួលស្បៀង - ឧបសម្ពន្ធ ៣) ទំហំក្រដាស A5 ដោយប្រើប្រាស់ Headless Edge ឬ Chromium លើ Windows"""
  import subprocess
  import tempfile
  import os

  html = generate_annex3_html(
      district, commune, school_name, voucher_no, invoice_date, df_items,
      supplier_name, comment, supplier_sig, director_sig, receiver_sig,
      consumption_date=consumption_date
  )

  browser_exe = find_headless_browser()

  if browser_exe:
    try:
      with tempfile.TemporaryDirectory() as tmp_dir:
        html_path = os.path.join(tmp_dir, "receipt.html")
        pdf_path = os.path.join(tmp_dir, "receipt.pdf")
        with open(html_path, "w", encoding="utf-8") as f:
          f.write(html)

        for h_flag in ["--headless=new", "--headless"]:
          cmd = [
              browser_exe,
              h_flag,
              "--no-sandbox",
              "--disable-dev-shm-usage",
              "--disable-gpu",
              "--run-all-compositor-stages-before-draw",
              "--no-pdf-header-footer",
              f"--print-to-pdf={pdf_path}",
              html_path,
          ]
          res = subprocess.run(cmd, capture_output=True, timeout=20)
          if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
            with open(pdf_path, "rb") as f:
              return f.read()
    except Exception:
      pass

  # Fallback to generate_simple_pdf
  return generate_simple_pdf(f"Food Receipt - {school_name} ({voucher_no})", f"Date: {invoice_date}", df_items if isinstance(df_items, pd.DataFrame) else pd.DataFrame(df_items))


def write_invoice_block_to_ws(
    ws,
    start_row: int,
    district: str,
    commune: str,
    school_name: str,
    voucher_no: str,
    invoice_date,
    items_list: list,
    supplier_name: str = "សាត ក្រូត",
    comment: str = "",
    supplier_sig=None,
    director_sig=None,
    receiver_sig=None,
    consumption_date=None,
    copy_label: str = "ច្បាប់ដើមសម្រាប់សាលា"
):
  """
  សរសេរប្លុកវិក្កយបត្រ (បង្កាន់ដៃទទួលស្បៀង - ឧបសម្ពន្ធ ៣) ចំនួន ២៧ ជួរដេក 
  តាមទម្រង់សន្លឹក «Invoice 1» នៃឯកសារ «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» បេះបិទ ១០០%
  """
  from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
  import openpyxl.drawing.image
  import base64
  import io

  # Fonts
  font_muol = Font(name="Khmer OS Muol Light", size=10, bold=True)
  font_muol_title = Font(name="Khmer OS Muol", size=10, bold=True)
  font_siemreap_bold = Font(name="Khmer OS Siemreap", size=9, bold=True)
  font_siemreap = Font(name="Khmer OS Siemreap", size=9)
  font_hand = Font(name="AKbalthom KhmerHand", size=9)
  font_battambang = Font(name="Khmer OS Battambang", size=9)
  font_note = Font(name="Khmer OS Siemreap", size=8)
  font_note_bold = Font(name="Khmer OS Siemreap", size=8, bold=True)
  font_copy = Font(name="Khmer OS Siemreap", size=8, italic=True)

  # Borders
  thin = Side(border_style="thin", color="000000")
  border_all = Border(top=thin, bottom=thin, left=thin, right=thin)

  # Alignments
  align_center = Alignment(horizontal="center", vertical="center")
  align_left = Alignment(horizontal="left", vertical="center")
  align_right = Alignment(horizontal="right", vertical="center")

  r = start_row

  # R1: C:G merged "ព្រះរាជាណាចក្រកម្ពុជា", H "ឧបសម្ពន្ធ ៣"
  ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
  ws.cell(row=r, column=3, value="ព្រះរាជាណាចក្រកម្ពុជា").font = font_muol
  ws.cell(row=r, column=3).alignment = align_center
  ws.cell(row=r, column=8, value="ឧបសម្ពន្ធ ៣").font = font_siemreap
  ws.cell(row=r, column=8).alignment = align_right
  ws.row_dimensions[r].height = 22.0

  # R2: C:G merged "ជាតិ សាសនា​ ព្រះមហាក្សត្រ"
  r += 1
  ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
  ws.cell(row=r, column=3, value="ជាតិ សាសនា ព្រះមហាក្សត្រ").font = font_muol
  ws.cell(row=r, column=3).alignment = align_center
  ws.row_dimensions[r].height = 17.5

  # R3: A "ក្រុង/ស្រុក", C:D merged district
  r += 1
  ws.cell(row=r, column=1, value="ក្រុង/ស្រុក").font = font_siemreap_bold
  ws.cell(row=r, column=1).alignment = align_left
  ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
  ws.cell(row=r, column=3, value=district or "").font = font_hand
  ws.cell(row=r, column=3).alignment = align_left
  ws.row_dimensions[r].height = 17.5

  # R4: A "ឃុំ/សង្កាត់", C:D merged commune, G "លេខសក្ខីប័ត្រ ៖", H voucher_no
  r += 1
  ws.cell(row=r, column=1, value="ឃុំ/សង្កាត់").font = font_siemreap_bold
  ws.cell(row=r, column=1).alignment = align_left
  ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
  ws.cell(row=r, column=3, value=commune or "").font = font_hand
  ws.cell(row=r, column=3).alignment = align_left

  ws.cell(row=r, column=7, value="លេខសក្ខីប័ត្រ ៖").font = font_siemreap
  ws.cell(row=r, column=7).alignment = align_right
  ws.cell(row=r, column=8, value=str(voucher_no or "")).font = font_siemreap_bold
  ws.cell(row=r, column=8).alignment = align_center
  ws.row_dimensions[r].height = 17.5

  # R5: A "សាលាបឋមសិក្សា", C:D merged school_name
  r += 1
  ws.cell(row=r, column=1, value="សាលាបឋមសិក្សា").font = font_siemreap_bold
  ws.cell(row=r, column=1).alignment = align_left
  ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
  ws.cell(row=r, column=3, value=school_name or "").font = font_hand
  ws.cell(row=r, column=3).alignment = align_left
  ws.row_dimensions[r].height = 17.5

  # R6: A:H merged "បង្កាន់ដៃទទួលស្បៀង"
  r += 1
  ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
  ws.cell(row=r, column=1, value="បង្កាន់ដៃទទួលស្បៀង").font = font_muol_title
  ws.cell(row=r, column=1).alignment = align_center
  ws.row_dimensions[r].height = 19.5

  # R7: A:H merged Date
  r += 1
  ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
  date_label = format_khmer_date(invoice_date)
  ws.cell(row=r, column=1, value=date_label).font = font_hand
  ws.cell(row=r, column=1).alignment = align_right
  ws.row_dimensions[r].height = 16.0

  # R8: Table Header (A=លរ, B:E merged=បរិយាយមុខទំនិញ, F=បរិមាណ, G=តម្លៃឯកតា, H=តម្លៃសរុប)
  r += 1
  ws.cell(row=r, column=1, value="លរ").font = font_siemreap_bold
  ws.cell(row=r, column=1).alignment = align_center
  ws.cell(row=r, column=1).border = border_all

  ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
  for c in range(2, 6):
    ws.cell(row=r, column=c).border = border_all
  ws.cell(row=r, column=2, value="បរិយាយមុខទំនិញ").font = font_siemreap_bold
  ws.cell(row=r, column=2).alignment = align_center

  ws.cell(row=r, column=6, value="បរិមាណ").font = font_siemreap_bold
  ws.cell(row=r, column=6).alignment = align_center
  ws.cell(row=r, column=6).border = border_all

  ws.cell(row=r, column=7, value="តម្លៃឯកតា").font = font_siemreap_bold
  ws.cell(row=r, column=7).alignment = align_center
  ws.cell(row=r, column=7).border = border_all

  ws.cell(row=r, column=8, value="តម្លៃសរុប").font = font_siemreap_bold
  ws.cell(row=r, column=8).alignment = align_center
  ws.cell(row=r, column=8).border = border_all
  ws.row_dimensions[r].height = 21.0

  # R9 to R18: 10 Item Rows
  tot_qty = 0.0
  tot_amt = 0.0
  for item_idx in range(10):
    r += 1
    row_num = item_idx + 1
    ws.row_dimensions[r].height = 20.5

    ws.cell(row=r, column=1, value=row_num).font = font_siemreap
    ws.cell(row=r, column=1).alignment = align_center
    ws.cell(row=r, column=1).border = border_all

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    for c in range(2, 6):
      ws.cell(row=r, column=c).border = border_all

    ws.cell(row=r, column=6).border = border_all
    ws.cell(row=r, column=6).alignment = align_center
    ws.cell(row=r, column=6).font = font_battambang

    ws.cell(row=r, column=7).border = border_all
    ws.cell(row=r, column=7).alignment = align_right
    ws.cell(row=r, column=7).font = font_siemreap
    ws.cell(row=r, column=7).number_format = '_(#,###_)\\ \\៛'

    ws.cell(row=r, column=8).border = border_all
    ws.cell(row=r, column=8).alignment = align_right
    ws.cell(row=r, column=8).font = font_siemreap
    ws.cell(row=r, column=8).number_format = '_(#,###_)\\ \\៛'

    if item_idx < len(items_list):
      item = items_list[item_idx]
      name = str(item.get("មុខទំនិញ", item.get("item_name", item.get("name", "")))).strip()
      try:
        qty = float(item.get("បរិមាណ", item.get("quantity", item.get("qty", 0))) or 0)
      except Exception:
        qty = 0.0
      try:
        unit_p = float(item.get("តម្លៃរាយ (៛)", item.get("unit_price", 0)) or 0)
      except Exception:
        unit_p = 0.0
      try:
        total_p = float(item.get("សរុប (៛)", item.get("total_price", qty * unit_p)) or (qty * unit_p))
      except Exception:
        total_p = 0.0

      ws.cell(row=r, column=2, value=name).font = font_hand
      ws.cell(row=r, column=2).alignment = align_left
      ws.cell(row=r, column=6, value=qty if qty % 1 != 0 else int(qty))
      ws.cell(row=r, column=7, value=unit_p)
      ws.cell(row=r, column=8, value=total_p)

      tot_qty += qty
      tot_amt += total_p

  # R19: Total Row
  r += 1
  ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
  for c in range(1, 6):
    ws.cell(row=r, column=c).border = border_all
  ws.cell(row=r, column=1, value="តម្លៃសរុប ").font = font_siemreap_bold
  ws.cell(row=r, column=1).alignment = align_center

  ws.cell(row=r, column=6, value=tot_qty if tot_qty % 1 != 0 else int(tot_qty)).font = font_siemreap_bold
  ws.cell(row=r, column=6).alignment = align_center
  ws.cell(row=r, column=6).border = border_all

  ws.cell(row=r, column=7).border = border_all

  ws.cell(row=r, column=8, value=tot_amt).font = font_siemreap_bold
  ws.cell(row=r, column=8).alignment = align_right
  ws.cell(row=r, column=8).border = border_all
  ws.cell(row=r, column=8).number_format = '_(#,###_)\\ \\៛'
  ws.row_dimensions[r].height = 21.0

  # R20: Note
  r += 1
  ws.cell(row=r, column=1, value="សម្គាល់៖").font = font_note_bold
  ws.cell(row=r, column=1).alignment = align_left
  ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
  ws.cell(row=r, column=2, value="រាល់តម្លៃស្បៀងស្នើសុំទូទាត់ត្រូវស្របតាមកិច្ចសន្យាផ្គត់ផ្គង់ស្បៀង។").font = font_note
  ws.cell(row=r, column=2).alignment = align_left
  ws.row_dimensions[r].height = 19.5

  # R21: Signatures Header
  r += 1
  ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
  ws.cell(row=r, column=1, value="បានឃើញ និងឯកភាព").font = font_siemreap_bold
  ws.cell(row=r, column=1).alignment = align_center

  ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
  ws.cell(row=r, column=5, value="អ្នកទទួល").font = font_siemreap_bold
  ws.cell(row=r, column=5).alignment = align_center

  ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
  ws.cell(row=r, column=7, value="អ្នកប្រគល់").font = font_siemreap_bold
  ws.cell(row=r, column=7).alignment = align_center
  ws.row_dimensions[r].height = 22.5

  # R22: Roles
  r += 1
  ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
  ws.cell(row=r, column=1, value="ប្រធាន គមស (នាយក/នាយិកាសាលា)").font = font_siemreap
  ws.cell(row=r, column=1).alignment = align_center

  ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
  ws.cell(row=r, column=5, value="នាយឃ្លាំងឬបេឡាធិការ").font = font_siemreap
  ws.cell(row=r, column=5).alignment = align_center

  ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
  ws.cell(row=r, column=7, value="អ្នកផ្គត់ផ្គង់ស្បៀង").font = font_siemreap
  ws.cell(row=r, column=7).alignment = align_center
  ws.row_dimensions[r].height = 20.0

  # R23 & R24: Signature space / Images / Dots
  sig_r = r + 1
  ws.row_dimensions[sig_r].height = 36.0
  ws.merge_cells(start_row=sig_r, start_column=1, end_row=sig_r, end_column=4)
  ws.cell(row=sig_r, column=1, value="...............................").alignment = align_center
  ws.cell(row=sig_r, column=1).font = font_note

  ws.merge_cells(start_row=sig_r, start_column=5, end_row=sig_r, end_column=6)
  ws.cell(row=sig_r, column=5, value="...............................").alignment = align_center
  ws.cell(row=sig_r, column=5).font = font_note

  ws.merge_cells(start_row=sig_r, start_column=7, end_row=sig_r, end_column=8)
  ws.cell(row=sig_r, column=7, value="...............................").alignment = align_center
  ws.cell(row=sig_r, column=7).font = font_note

  # Embed signatures if provided
  if director_sig and str(director_sig).startswith("data:image"):
    try:
      b64_d = str(director_sig).split(",", 1)[1]
      img_d = openpyxl.drawing.image.Image(io.BytesIO(base64.b64decode(b64_d)))
      img_d.width = 90
      img_d.height = 36
      ws.add_image(img_d, f"B{sig_r}")
    except Exception:
      pass

  if receiver_sig and str(receiver_sig).startswith("data:image"):
    try:
      b64_rc = str(receiver_sig).split(",", 1)[1]
      img_rc = openpyxl.drawing.image.Image(io.BytesIO(base64.b64decode(b64_rc)))
      img_rc.width = 90
      img_rc.height = 36
      ws.add_image(img_rc, f"E{sig_r}")
    except Exception:
      pass

  if supplier_sig and str(supplier_sig).startswith("data:image"):
    try:
      b64_s = str(supplier_sig).split(",", 1)[1]
      img_s = openpyxl.drawing.image.Image(io.BytesIO(base64.b64decode(b64_s)))
      img_s.width = 90
      img_s.height = 36
      ws.add_image(img_s, f"G{sig_r}")
    except Exception:
      pass

  r += 2

  # R25: Supplier Name under signature
  r += 1
  ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
  ws.cell(row=r, column=7, value=supplier_name or "").font = font_siemreap_bold
  ws.cell(row=r, column=7).alignment = align_center
  ws.row_dimensions[r].height = 20.0

  # R26: Spacer row
  r += 1
  ws.row_dimensions[r].height = 14.0

  # R27: Bottom copy label
  r += 1
  ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
  ws.cell(row=r, column=1, value=copy_label).font = font_copy
  ws.cell(row=r, column=1).alignment = align_center
  ws.row_dimensions[r].height = 17.5


def generate_annex3_excel(
    district, commune, school_name, voucher_no, invoice_date, df_items,
    supplier_name="សាត ក្រូត", comment="",
    supplier_sig=None, director_sig=None, receiver_sig=None,
    consumption_date=None, copy_label="ច្បាប់ដើមសម្រាប់សាលា"
):
  """បង្កើតឯកសារ Excel ផ្លូវការ ឧបសម្ពន្ធ ៣ (បង្កាន់ដៃទទួលស្បៀង) តាមគំរូសន្លឹក Invoice 1 បេះបិទ ១០០%"""
  import openpyxl

  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Invoice 1"

  ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
  ws.page_setup.paperSize = ws.PAPERSIZE_A5
  ws.sheet_properties.pageSetUpPr.fitToPage = True
  ws.page_setup.fitToWidth = 1
  ws.page_setup.fitToHeight = 1

  col_widths = {
      'A': 8.18,
      'B': 14.36,
      'C': 3.5,
      'D': 9.54,
      'E': 3.5,
      'F': 9.82,
      'G': 11.45,
      'H': 16.18
  }
  for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

  if isinstance(df_items, pd.DataFrame):
    items_list = df_items.to_dict('records') if not df_items.empty else []
  elif isinstance(df_items, list):
    items_list = df_items
  else:
    items_list = []

  write_invoice_block_to_ws(
      ws=ws,
      start_row=1,
      district=district,
      commune=commune,
      school_name=school_name,
      voucher_no=voucher_no,
      invoice_date=invoice_date,
      items_list=items_list,
      supplier_name=supplier_name,
      comment=comment,
      supplier_sig=supplier_sig,
      director_sig=director_sig,
      receiver_sig=receiver_sig,
      consumption_date=consumption_date,
      copy_label=copy_label
  )

  out = io.BytesIO()
  wb.save(out)
  return out.getvalue()


def generate_all_school_invoices_excel(
    district, commune, school_name, month_prefix,
    supplier_name="សាត ក្រូត", comment="",
    supplier_sig=None, director_sig=None, receiver_sig=None
):
  """
  បង្កើតឯកសារ Excel សៀវភៅបង្កាន់ដៃប្រចាំខែទាំងអស់ (Full Invoice 1 Workbook)
  ដោយរៀបចំប្លុកវិក្កយបត្រនីមួយៗ (២៧ ជួរដេកក្នុងមួយប្លុក) បន្តបន្ទាប់គ្នាបញ្ឈរ 
  ដូចសន្លឹកកិច្ចការ «Invoice 1» នៃឯកសារ «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» ទាំងស្រុង!
  """
  import openpyxl

  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Invoice 1"

  ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
  ws.sheet_properties.pageSetUpPr.fitToPage = True
  ws.page_setup.fitToWidth = 1

  col_widths = {
      'A': 8.18,
      'B': 14.36,
      'C': 3.5,
      'D': 9.54,
      'E': 3.5,
      'F': 9.82,
      'G': 11.45,
      'H': 16.18
  }
  for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

  # ទាញយកកាលបរិច្ឆេទទាំងអស់ក្នុងខែនោះសម្រាប់សាលានេះ
  dates_query = """
      SELECT DISTINCT date, voucher_no, consumption_date 
      FROM daily_records 
      WHERE school_name=? AND date LIKE ?
      ORDER BY date ASC
  """
  date_rows = cursor.execute(dates_query, (school_name, f"{month_prefix}%")).fetchall()

  if not date_rows:
    write_invoice_block_to_ws(
        ws=ws, start_row=1, district=district, commune=commune,
        school_name=school_name, voucher_no="001", invoice_date=f"{month_prefix}-01",
        items_list=[], supplier_name=supplier_name, comment=comment
    )
  else:
    for idx, (inv_d, v_no, c_d) in enumerate(date_rows):
      start_r = 1 + idx * 27
      items_q = """
          SELECT item_name as [មុខទំនិញ], quantity as [បរិមាណ], unit_price as [តម្លៃរាយ (៛)], total_price as [សរុប (៛)]
          FROM daily_records 
          WHERE school_name=? AND date=? 
          ORDER BY id ASC
      """
      df_day_items = pd.read_sql_query(items_q, conn, params=(school_name, inv_d))
      items_list = df_day_items.to_dict('records') if not df_day_items.empty else []

      write_invoice_block_to_ws(
          ws=ws,
          start_row=start_r,
          district=district,
          commune=commune,
          school_name=school_name,
          voucher_no=v_no if v_no else f"{idx+1:03d}",
          invoice_date=inv_d,
          items_list=items_list,
          supplier_name=supplier_name,
          comment=comment,
          supplier_sig=supplier_sig,
          director_sig=director_sig,
          receiver_sig=receiver_sig,
          consumption_date=c_d
      )

  out = io.BytesIO()
  wb.save(out)
  return out.getvalue()


# ================= Helper Functions សម្រាប់ សំណើទូទាត់ប្រចាំខែ =================
def format_khmer_date_range(d_start, d_end):
  """បំប្លែងចន្លោះកាលបរិច្ឆេទទៅជាអក្សរខ្មែរ: ចាប់ពីថ្ងៃទី ... ដល់ថ្ងៃទី ..."""
  if isinstance(d_start, str):
    try:
      p = d_start.strip()[:10].split("-")
      d_start = date(int(p[0]), int(p[1]), int(p[2]))
    except Exception:
      pass
  if isinstance(d_end, str):
    try:
      p = d_end.strip()[:10].split("-")
      d_end = date(int(p[0]), int(p[1]), int(p[2]))
    except Exception:
      pass
  ds_str = f"{d_start.day:02d}"
  ms_str = f"{d_start.month:02d}"
  ys_str = f"{d_start.year}"
  de_str = f"{d_end.day:02d}"
  me_str = f"{d_end.month:02d}"
  ye_str = f"{d_end.year}"
  return f"ចាប់ពីថ្ងៃទី {to_khmer_num(ds_str)} ខែ {to_khmer_num(ms_str)} ឆ្នាំ {to_khmer_num(ys_str)} ដល់ថ្ងៃទី {to_khmer_num(de_str)} ខែ {to_khmer_num(me_str)} ឆ្នាំ {to_khmer_num(ye_str)}"


def round_khmer_currency(amount):
  """
  ក្បួនបង្គត់លេខផ្លូវការ៖
  * បង្គត់ស្មើ ០ បើខ្ទង់ដប់តិចជាង ៥០ រៀល (ឧ. 2,375,702 -> 2,375,700)
  * បង្គត់ឡើងមួយរយ បើខ្ទង់ដប់ស្មើឬច្រើនជាង ៥០ រៀល (ឧ. 2,375,750 -> 2,375,800)
  """
  if amount is None:
    return 0
  val = int(round(float(amount)))
  remainder = val % 100
  base = (val // 100) * 100
  if remainder < 50:
    return base
  else:
    return base + 100


def get_monthly_claim_items(school_name, d_start, d_end):
  """
  ទាញយក និងបូកសរុបទំនិញសម្រាប់សំណើទូទាត់ប្រចាំខែ
  ព្រមទាំងចាប់យកលេខយោងបង្កាន់ដៃទទួលទំនិញ ចាប់ផ្ដើមក្នុងខែ និងចុងខែ (គំរូ 001 - 016)
  """
  if not school_name:
    return []
  s_str = str(d_start)[:10]
  e_str = str(d_end)[:10]

  rows = cursor.execute(
      """SELECT item_name, quantity, unit_price, total_price, voucher_no, date 
         FROM daily_records 
         WHERE school_name=? AND date >= ? AND date <= ? 
         ORDER BY date ASC, id ASC""",
      (school_name, s_str, e_str)
  ).fetchall()

  if not rows:
    return []

  def parse_v_num(v_str):
    digits = re.findall(r'\d+', str(v_str))
    return int(digits[-1]) if digits else 0

  all_vouchers = []
  for r in rows:
    v = str(r[4] or "").strip()
    if v:
      all_vouchers.append(v)
  sorted_all_v = sorted(list(set(all_vouchers)), key=parse_v_num)
  month_fallback_ref = f"{sorted_all_v[0]} - {sorted_all_v[-1]}" if len(sorted_all_v) > 1 else (sorted_all_v[0] if sorted_all_v else "001 - 016")

  from collections import OrderedDict
  items_dict = OrderedDict()

  for r in rows:
    i_name = str(r[0] or "").strip()
    if not i_name:
      continue
    qty = float(r[1] or 0)
    u_price = float(r[2] or 0)
    t_price = float(r[3] or (qty * u_price))
    v_no = str(r[4] or "").strip()

    if i_name not in items_dict:
      items_dict[i_name] = {
          "name": i_name,
          "qty": 0.0,
          "unit_prices": [],
          "total_price": 0.0,
          "vouchers": []
      }

    items_dict[i_name]["qty"] += qty
    items_dict[i_name]["total_price"] += t_price
    if u_price > 0:
      items_dict[i_name]["unit_prices"].append(u_price)
    if v_no:
      items_dict[i_name]["vouchers"].append(v_no)

  result = []
  for i_name, data in items_dict.items():
    # ចាប់យកលេខយោងបង្កាន់ដៃទទួលទំនិញ ចាប់ផ្ដើម និងចុងបញ្ចប់ ក្នុងខែ
    v_set = sorted(list(set(data["vouchers"])), key=parse_v_num)
    if len(v_set) > 1:
      v_ref = f"{v_set[0]} - {v_set[-1]}"
    elif len(v_set) == 1:
      v_ref = v_set[0]
    else:
      v_ref = month_fallback_ref

    u_list = data["unit_prices"]
    if u_list:
      final_u_price = u_list[-1]
    elif data["qty"] > 0:
      final_u_price = round(data["total_price"] / data["qty"], 2)
    else:
      final_u_price = 0.0

    row_c = cursor.execute("SELECT category FROM products WHERE item_name=? LIMIT 1", (i_name,)).fetchone()
    c_val = row_c[0] if (row_c and row_c[0]) else classify_item_category(i_name)
    result.append({
        "name": i_name,
        "category": c_val,
        "voucher_ref": v_ref,
        "qty": round(data["qty"], 2),
        "unit_price": final_u_price,
        "total_price": round(data["total_price"], 2)
    })

  return result


def generate_monthly_claim_html(district, commune, school_name, voucher_no, d_start, d_end, items,
                                supplier_name="សាត ក្រូត", supplier_address="ភូមិខ្មែរ ឃុំរោង",
                                supplier_phone="090 854 133", checked_cats=None,
                                supplier_sig=None, director_sig=None, preparer_sig=None):
  """បង្កើតកូដ HTML តាមគំរូផ្លូវការ សំណើសុំទូទាត់ប្រចាំខែ (ទំហំ A4) ដោយគ្មានហត្ថលេខាក្លែងក្លាយ និងគាំទ្រការបញ្ចូលហត្ថលេខាពិត"""
  if checked_cats is None:
    checked_cats = ["អង្ករ", "អំបិល", "ប្រេងឆា", "សាច់ ត្រី ស៊ុត", "បន្លែ"]

  date_range_kh = format_khmer_date_range(d_start, d_end)
  kh_digits = ['១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩', '១០',
               '១១', '១២', '១៣', '១៤', '១៥', '១៦', '១៧', '១៨', '១៩', '២០',
               '២១', '២២', '២៣', '២៤', '២៥', '២៦', '២៧', '២៨', '២៩', '៣០']

  if isinstance(items, pd.DataFrame):
    items_list = items.to_dict('records') if not items.empty else []
  elif isinstance(items, list):
    items_list = items
  else:
    items_list = []

  total_amount = 0.0
  rows_html = []

  for idx, it in enumerate(items_list):
    r_num_kh = kh_digits[idx] if idx < len(kh_digits) else to_khmer_num(idx + 1)
    name = str(it.get('name', it.get('item_name', ''))).strip()
    v_ref = str(it.get('voucher_ref', it.get('ref_no', ''))).strip()
    try:
      qty = float(it.get('qty', it.get('quantity', 0)) or 0)
    except Exception:
      qty = 0.0
    try:
      u_price = float(it.get('unit_price', 0) or 0)
    except Exception:
      u_price = 0.0
    try:
      t_price = float(it.get('total_price', qty * u_price) or (qty * u_price))
    except Exception:
      t_price = 0.0

    total_amount += t_price
    q_str = f"{qty:g}" if qty % 1 != 0 else f"{int(qty)}"
    u_str = f"{int(u_price):,} ៛" if u_price % 1 == 0 else f"{u_price:,.2f} ៛"
    t_str = f"{int(t_price):,} ៛" if t_price % 1 == 0 else f"{t_price:,.2f} ៛"

    rows_html.append(f"""
    <tr>
      <td class="col-num">{r_num_kh}</td>
      <td class="col-name">{name}</td>
      <td class="col-ref">{v_ref}</td>
      <td class="col-qty">{q_str}</td>
      <td class="col-price">{u_str}</td>
      <td class="col-total">{t_str}</td>
    </tr>
    """)

  rounded_amount = round_khmer_currency(total_amount)
  tot_amt_str = f"{int(total_amount):,} ៛" if total_amount % 1 == 0 else f"{total_amount:,.2f} ៛"
  rounded_amt_str = f"{int(rounded_amount):,} ៛"
  v_display = f"{voucher_no}" if voucher_no else "0008"

  def chk(cat_name):
    cat_words = set(cat_name.split())
    for c in (checked_cats or []):
      if c == cat_name or cat_name in c or c in cat_name:
        return "🗹"
      if len(cat_words & set(c.split())) >= 2:
        return "🗹"
    return "☐"

  html = f"""<!DOCTYPE html>
<html lang="km">
<head>
<meta charset="utf-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Battambang:wght@400;700&family=Kantumruy+Pro:ital,wght@0,400;0,600;0,700;1,400&family=Moul&display=swap');
  
  @page {{
    size: A4 portrait;
    margin: 10mm 14mm 8mm 14mm;
  }}
  
  * {{
    box-sizing: border-box;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }}
  
  body {{
    font-family: 'Khmer OS Battambang', 'Kantumruy Pro', 'Battambang', Arial, sans-serif;
    font-size: 12px;
    line-height: 1.3;
    color: #000;
    background: #fff;
    margin: 0;
    padding: 6px 14px;
    width: 100%;
  }}
  
  .muol {{
    font-family: 'Khmer OS Muol Light', 'Moul', 'Khmer OS Muol', serif;
    font-weight: normal;
  }}
  
  .top-center {{
    text-align: center;
    margin-bottom: 2px;
  }}
  
  .country-title {{
    font-size: 14.5px;
    line-height: 1.35;
  }}
  
  .motto-title {{
    font-size: 13.5px;
    line-height: 1.35;
    margin-bottom: 4px;
  }}
  
  .meta-grid {{
    display: table;
    width: 100%;
    margin-top: 2px;
    margin-bottom: 2px;
  }}
  
  .meta-left {{
    display: table-cell;
    width: 60%;
    vertical-align: top;
  }}
  
  .meta-right {{
    display: table-cell;
    width: 40%;
    vertical-align: middle;
    text-align: right;
  }}
  
  .info-table {{
    border-collapse: collapse;
  }}
  
  .info-table td {{
    padding: 1px 2px;
    font-size: 12px;
  }}
  
  .info-table td.lbl {{
    font-weight: bold;
    width: 110px;
    white-space: nowrap;
  }}
  
  .info-table td.val {{
    font-weight: 600;
    color: #0c4a8a;
    padding-left: 10px;
  }}
  
  .voucher-box {{
    display: inline-block;
    padding-right: 20px;
  }}
  
  .voucher-lbl {{
    font-size: 12.5px;
    font-weight: 700;
    margin-right: 14px;
  }}
  
  .voucher-num {{
    font-size: 14px;
    font-weight: 700;
    color: #0c4a8a;
    letter-spacing: 1.5px;
    display: inline-block;
  }}
  
  .doc-title-row {{
    text-align: center;
    margin-top: 4px;
    margin-bottom: 2px;
  }}
  
  .doc-title {{
    font-size: 16px;
    line-height: 1.3;
  }}
  
  .date-range-row {{
    text-align: center;
    margin-bottom: 6px;
    font-weight: 700;
    font-size: 12px;
  }}
  
  .supplier-bar {{
    display: table;
    width: 100%;
    font-size: 11.5px;
    font-weight: 600;
    margin-bottom: 5px;
  }}
  
  .sup-col {{
    display: table-cell;
    vertical-align: middle;
  }}
  
  /* Main Table */
  .claim-table {{
    width: 100%;
    border-collapse: collapse;
    border: 1.5px solid #000;
    margin-bottom: 4px;
  }}
  
  .claim-table th, .claim-table td {{
    border: 1px solid #000;
    padding: 2.5px 4px;
    font-size: 11.5px;
    height: 20px;
  }}
  
  .claim-table th {{
    font-weight: 700;
    text-align: center;
    background: #fbfbfb;
    vertical-align: middle;
  }}
  
  .col-num {{
    width: 5%;
    text-align: center;
    font-weight: 600;
  }}
  
  .col-name {{
    width: 23%;
    text-align: left;
    padding-left: 6px;
    color: #0c4a8a;
    font-weight: 500;
  }}
  
  .col-ref {{
    width: 17%;
    text-align: center;
    color: #0c4a8a;
    font-weight: 600;
    font-size: 11px;
  }}
  
  .col-qty {{
    width: 11%;
    text-align: center;
    color: #0c4a8a;
  }}
  
  .col-price {{
    width: 15%;
    text-align: right;
    padding-right: 8px;
    color: #0c4a8a;
  }}
  
  .col-total {{
    width: 29%;
    text-align: right;
    padding-right: 8px;
    color: #0c4a8a;
    font-weight: 600;
  }}
  
  .checkboxes-box {{
    font-size: 10.5px;
    line-height: 1.3;
    font-weight: normal;
    text-align: left;
    padding: 2px 4px;
  }}
  
  .chk-item {{
    display: inline-block;
    margin-right: 8px;
    white-space: nowrap;
  }}
  
  .total-row td {{
    font-weight: 700;
    font-size: 12px;
  }}
  
  .note-text {{
    font-size: 10.5px;
    margin-top: 4px;
    margin-bottom: 8px;
    line-height: 1.35;
  }}
  
  /* Signatures */
  .sig-table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 6px;
    table-layout: fixed;
  }}
  
  .sig-table td {{
    vertical-align: top;
    text-align: center;
    padding: 0 4px;
  }}
  
  .sig-role {{
    font-weight: 700;
    font-size: 11.5px;
    line-height: 1.3;
  }}
  
  .sig-name {{
    font-weight: 700;
    font-size: 11.5px;
    margin-top: 2px;
  }}
  
  .sig-space {{
    height: 44px;
    display: flex;
    align-items: center;
    justify-content: center;
  }}
  
  .sig-dots {{
    margin-top: 28px;
    letter-spacing: 2px;
    color: #444;
  }}
</style>
</head>
<body>

<div class="top-center">
  <div class="muol country-title">ព្រះរាជាណាចក្រកម្ពុជា</div>
  <div class="muol motto-title">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
</div>

<div class="meta-grid">
  <div class="meta-left">
    <table class="info-table">
      <tr>
        <td class="lbl">ក្រុង/ស្រុក/ខណ្ឌ</td>
        <td class="val">{district if district else '&nbsp;'}</td>
      </tr>
      <tr>
        <td class="lbl">ឃុំ/សង្កាត់</td>
        <td class="val">{commune if commune else '&nbsp;'}</td>
      </tr>
      <tr>
        <td class="lbl">សាលាបឋមសិក្សា</td>
        <td class="val">{school_name if school_name else '&nbsp;'}</td>
      </tr>
    </table>
  </div>
  <div class="meta-right">
    <div class="voucher-box">
      <span class="voucher-lbl">លេខសក្ខីប័ត្រ ៖</span>
      <span class="voucher-num">{v_display}</span>
    </div>
  </div>
</div>

<div class="doc-title-row">
  <div class="muol doc-title">សំណើសុំទូទាត់ប្រចាំខែ</div>
</div>

<div class="date-range-row">
  {date_range_kh}
</div>

<div class="supplier-bar">
  <div class="sup-col" style="width: 35%;">ឈ្មោះអ្នកផ្គត់ផ្គង់: <span style="color:#0c4a8a;">{supplier_name}</span></div>
  <div class="sup-col" style="width: 35%; text-align: center;">អាសយដ្ឋាន: <span style="color:#0c4a8a;">{supplier_address}</span></div>
  <div class="sup-col" style="width: 30%; text-align: right;">លេខទូរស័ព្ទ: <span style="color:#0c4a8a;">{supplier_phone}</span></div>
</div>

<table class="claim-table">
  <thead>
    <tr>
      <th rowspan="2" style="width: 5%;">ល.រ</th>
      <th rowspan="2" style="width: 23%;">បរិយាយមុខទំនិញ</th>
      <th rowspan="2" style="width: 17%;">លេខយោងក្នុងបង្កាន់ដៃ<br>ទទួលទំនិញ</th>
      <th colspan="2" style="width: 26%;">ទំនិញទទួលបាន</th>
      <th style="width: 29%; padding: 3px 4px;">
        <div style="font-weight: 700; margin-bottom: 2px;">សរុបចំណាយ ( គូស 🗹 ខាងក្រោមនេះ )</div>
        <div class="checkboxes-box">
          <span class="chk-item">{chk('អង្ករ')} អង្ករ</span>
          <span class="chk-item">{chk('អំបិល')} អំបិល</span>
          <span class="chk-item">{chk('ប្រេងឆា')} ប្រេងឆា</span><br>
          <span class="chk-item">{chk('សាច់ ត្រី ស៊ុត')} សាច់ ត្រី ស៊ុត</span>
          <span class="chk-item">{chk('បន្លែ')} បន្លែ</span>
        </div>
      </th>
    </tr>
    <tr>
      <th style="width: 11%;">បរិមាណ<br>(គ.ក)</th>
      <th style="width: 15%;">តម្លៃឯកតា</th>
      <th style="width: 29%; border-top: none;">សរុប (៛)</th>
    </tr>
  </thead>
  <tbody>
    {''.join(rows_html)}
    <tr class="total-row">
      <td colspan="5" style="text-align: right; font-weight: 700; padding-right: 15px;">សរុបទឹកប្រាក់</td>
      <td style="text-align: right; padding-right: 8px; font-weight: 700; color: #0c4a8a;">{tot_amt_str}</td>
    </tr>
    <tr class="total-row" style="background-color: #fcfcfc;">
      <td colspan="5" style="text-align: right; font-weight: 700; padding-right: 15px;">ថវិកាសរុបស្នើសុំទូទាត់ (បង្គត់លេខ) *</td>
      <td style="text-align: right; padding-right: 8px; font-weight: 700; color: #0c4a8a; font-size: 13px;">{rounded_amt_str}</td>
    </tr>
  </tbody>
</table>

<div class="note-text">
  <strong>សម្គាល់៖</strong> *ថវិកាសរុបត្រូវទូទាត់ស្នើសុំត្រូវបង្គត់៖ បង្គត់ស្មើ ០ បើខ្ទង់ដប់តិចជាង ៥០ រៀល បង្គត់ឡើងមួយ បើខ្ទង់ដប់ស្មើឬច្រើនជាង ៥០ រៀល ។
</div>

<table class="sig-table">
  <tr>
    <td style="width: 33.3%;">
      <div class="sig-role">បានឃើញ និងឯកភាពដោយ</div>
      <div class="sig-name">ប្រធាន គ.ម.ស (នាយក ឬនាយិកាសាលា)</div>
      <div class="sig-space">
        {f'<img src="{director_sig}" style="max-height: 44px; max-width: 120px; object-fit: contain;" />' if director_sig else ''}
      </div>
      <div class="sig-dots">...............................</div>
    </td>
    <td style="width: 33.3%;">
      <div class="sig-role">ផ្ទៀងផ្ទាត់ដោយ</div>
      <div class="sig-name">អ្នកផ្គត់ផ្គង់</div>
      <div class="sig-space">
        {f'<img src="{supplier_sig}" style="max-height: 44px; max-width: 120px; object-fit: contain;" />' if supplier_sig else ''}
      </div>
      {f'<div class="sig-dots">...............................</div>' if not supplier_sig else ''}
      <div class="sig-name" style="color: #0c4a8a; font-weight: bold; margin-top: 2px;">{supplier_name}</div>
    </td>
    <td style="width: 33.3%;">
      <div class="sig-role">រៀបចំដោយ</div>
      <div class="sig-name">នាយឃ្លាំង/បេឡាធិការ</div>
      <div class="sig-space">
        {f'<img src="{preparer_sig}" style="max-height: 44px; max-width: 120px; object-fit: contain;" />' if preparer_sig else ''}
      </div>
      <div class="sig-dots">...............................</div>
    </td>
  </tr>
</table>

</body>
</html>
"""
  return html


def generate_monthly_claim_pdf(district, commune, school_name, voucher_no, d_start, d_end, items,
                               supplier_name="សាត ក្រូត", supplier_address="ភូមិខ្មែរ ឃុំរោង",
                               supplier_phone="090 854 133", checked_cats=None,
                               supplier_sig=None, director_sig=None, preparer_sig=None):
  """បង្កើតឯកសារ PDF ផ្លូវការ (សំណើសុំទូទាត់ប្រចាំខែ) ដោយប្រើប្រាស់ Headless Edge ឬ Chromium"""
  import subprocess
  import tempfile
  import os

  html = generate_monthly_claim_html(
      district, commune, school_name, voucher_no, d_start, d_end, items,
      supplier_name, supplier_address, supplier_phone, checked_cats,
      supplier_sig, director_sig, preparer_sig
  )

  browser_exe = find_headless_browser()

  if browser_exe:
    try:
      with tempfile.TemporaryDirectory() as tmp_dir:
        html_path = os.path.join(tmp_dir, "claim.html")
        pdf_path = os.path.join(tmp_dir, "claim.pdf")
        with open(html_path, "w", encoding="utf-8") as f:
          f.write(html)

        cmd = [
            browser_exe,
            "--headless",
            "--disable-gpu",
            "--run-all-compositor-stages-before-draw",
            "--no-pdf-header-footer",
            f"--print-to-pdf={pdf_path}",
            html_path,
        ]
        res = subprocess.run(cmd, capture_output=True, timeout=20)
        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
          with open(pdf_path, "rb") as f:
            return f.read()
    except Exception:
      pass

  # Fallback to simple PDF
  df_items = pd.DataFrame(items) if not isinstance(items, pd.DataFrame) else items
  return generate_simple_pdf(f"Monthly Claim - {school_name}", f"Period: {d_start} to {d_end}", df_items)


def generate_monthly_claim_excel(district, commune, school_name, voucher_no, d_start, d_end, items,
                                 supplier_name="សាត ក្រូត", supplier_address="ភូមិខ្មែរ ឃុំរោង",
                                 supplier_phone="090 854 133", checked_cats=None):
  """បង្កើតឯកសារ Excel ផ្លូវការ សម្រាប់សំណើសុំទូទាត់ប្រចាំខែ"""
  import openpyxl
  from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

  if isinstance(items, pd.DataFrame):
    items_list = items.to_dict('records') if not items.empty else []
  elif isinstance(items, list):
    items_list = items
  else:
    items_list = []

  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "សំណើទូទាត់ប្រចាំខែ"
  ws.views.sheetView[0].showGridLines = True

  ws.page_setup.paperSize = ws.PAPERSIZE_A4
  ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
  ws.page_margins.left = 0.4
  ws.page_margins.right = 0.4
  ws.page_margins.top = 0.4
  ws.page_margins.bottom = 0.4

  ws.column_dimensions['A'].width = 6
  ws.column_dimensions['B'].width = 24
  ws.column_dimensions['C'].width = 18
  ws.column_dimensions['D'].width = 12
  ws.column_dimensions['E'].width = 15
  ws.column_dimensions['F'].width = 22

  font_family = "Khmer OS Battambang"
  font_muol = "Khmer OS Muol Light"

  thin_side = Side(border_style="thin", color="000000")
  double_side = Side(border_style="double", color="000000")
  cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
  tot_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)

  # 1. Header Center
  ws.merge_cells('A1:F1')
  ws['A1'] = "ព្រះរាជាណាចក្រកម្ពុជា"
  ws['A1'].font = Font(name=font_muol, size=13, bold=True)
  ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

  ws.merge_cells('A2:F2')
  ws['A2'] = "ជាតិ សាសនា ព្រះមហាក្សត្រ"
  ws['A2'].font = Font(name=font_muol, size=12, bold=True)
  ws['A2'].alignment = Alignment(horizontal="center", vertical="center")

  # 2. Metadata
  ws['A4'] = "ក្រុង/ស្រុក/ខណ្ឌ"
  ws['A4'].font = Font(name=font_family, size=10.5, bold=True)
  ws['B4'] = f":  {district}" if district else ":"
  ws['B4'].font = Font(name=font_family, size=10.5, color="0C4A8A", bold=True)

  ws['E4'] = "លេខសក្ខីប័ត្រ ៖"
  ws['E4'].font = Font(name=font_family, size=10.5, bold=True)
  ws['E4'].alignment = Alignment(horizontal="right", vertical="center")
  ws['F4'] = f"{voucher_no}" if voucher_no else "0008"
  ws['F4'].font = Font(name=font_family, size=12, bold=True, color="0C4A8A")
  ws['F4'].alignment = Alignment(horizontal="center", vertical="center")

  ws['A5'] = "ឃុំ/សង្កាត់"
  ws['A5'].font = Font(name=font_family, size=10.5, bold=True)
  ws['B5'] = f":  {commune}" if commune else ":"
  ws['B5'].font = Font(name=font_family, size=10.5, color="0C4A8A", bold=True)

  ws['A6'] = "សាលាបឋមសិក្សា"
  ws['A6'].font = Font(name=font_family, size=10.5, bold=True)
  ws['B6'] = f":  {school_name}" if school_name else ":"
  ws['B6'].font = Font(name=font_family, size=10.5, color="0C4A8A", bold=True)

  # 3. Title
  ws.merge_cells('A8:F8')
  ws['A8'] = "សំណើសុំទូទាត់ប្រចាំខែ"
  ws['A8'].font = Font(name=font_muol, size=14, bold=True)
  ws['A8'].alignment = Alignment(horizontal="center", vertical="center")
  ws.row_dimensions[8].height = 26

  # 4. Date Range
  ws.merge_cells('A9:F9')
  ws['A9'] = format_khmer_date_range(d_start, d_end)
  ws['A9'].font = Font(name=font_family, size=11, bold=True)
  ws['A9'].alignment = Alignment(horizontal="center", vertical="center")

  # 5. Supplier info bar
  ws.merge_cells('A10:B10')
  ws['A10'] = f"ឈ្មោះអ្នកផ្គត់ផ្គង់:  {supplier_name}"
  ws['A10'].font = Font(name=font_family, size=10, bold=True)

  ws.merge_cells('C10:D10')
  ws['C10'] = f"អាសយដ្ឋាន:  {supplier_address}"
  ws['C10'].font = Font(name=font_family, size=10, bold=True)
  ws['C10'].alignment = Alignment(horizontal="center", vertical="center")

  ws.merge_cells('E10:F10')
  ws['E10'] = f"លេខទូរស័ព្ទ:  {supplier_phone}"
  ws['E10'].font = Font(name=font_family, size=10, bold=True)
  ws['E10'].alignment = Alignment(horizontal="right", vertical="center")
  ws.row_dimensions[10].height = 20

  # 6. Table Headers
  ws.merge_cells('A11:A12')
  ws['A11'] = "ល.រ"
  ws.merge_cells('B11:B12')
  ws['B11'] = "បរិយាយមុខទំនិញ"
  ws.merge_cells('C11:C12')
  ws['C11'] = "លេខយោងក្នុងបង្កាន់ដៃទទួលទំនិញ"

  ws.merge_cells('D11:E11')
  ws['D11'] = "ទំនិញទទួលបាន"
  ws['D12'] = "បរិមាណ (គ.ក)"
  ws['E12'] = "តម្លៃឯកតា"

  def chk_e(c_name):
    c_words = set(c_name.split())
    for x in (checked_cats or []):
      if x == c_name or c_name in x or x in c_name:
        return "🗹"
      if len(c_words & set(x.split())) >= 2:
        return "🗹"
    return "☐"
  ws['F11'] = f"សរុបចំណាយ ({chk_e('អង្ករ')} អង្ករ  {chk_e('អំបិល')} អំបិល  {chk_e('ប្រេងឆា')} ប្រេងឆា  {chk_e('សាច់')} សាច់ ត្រី ស៊ុត  {chk_e('បន្លែ')} បន្លែ)"
  ws['F12'] = "សរុប (៛)"

  for r_i in [11, 12]:
    for c_i in range(1, 7):
      cell = ws.cell(row=r_i, column=c_i)
      cell.font = Font(name=font_family, size=10, bold=True)
      cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
      cell.border = cell_border
      cell.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

  ws.row_dimensions[11].height = 22
  ws.row_dimensions[12].height = 20

  # 7. Data rows
  start_r = 13
  kh_digits = ['១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩', '១០',
               '១១', '១២', '១៣', '១៤', '១៥', '១៦', '១៧', '១៨', '១៩', '២០',
               '២១', '២២', '២៣', '២៤', '២៥', '២៦', '២៧', '២៨', '២៩', '៣០']

  for idx, it in enumerate(items_list):
    r = start_r + idx
    ws.row_dimensions[r].height = 20
    r_num_kh = kh_digits[idx] if idx < len(kh_digits) else to_khmer_num(idx + 1)
    name = it.get('name', it.get('item_name', ''))
    v_ref = it.get('voucher_ref', it.get('ref_no', ''))
    qty = float(it.get('qty', it.get('quantity', 0)) or 0)
    u_price = float(it.get('unit_price', 0) or 0)

    c_a = ws.cell(row=r, column=1, value=r_num_kh)
    c_a.alignment = Alignment(horizontal="center", vertical="center")

    c_b = ws.cell(row=r, column=2, value=name)
    c_b.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c_b.font = Font(name=font_family, size=10, color="0C4A8A")

    c_c = ws.cell(row=r, column=3, value=v_ref)
    c_c.alignment = Alignment(horizontal="center", vertical="center")
    c_c.font = Font(name=font_family, size=10, color="0C4A8A", bold=True)

    c_d = ws.cell(row=r, column=4, value=qty)
    c_d.alignment = Alignment(horizontal="center", vertical="center")
    c_d.number_format = "#,##0.##"
    c_d.font = Font(name=font_family, size=10, color="0C4A8A")

    c_e = ws.cell(row=r, column=5, value=u_price)
    c_e.alignment = Alignment(horizontal="right", vertical="center")
    c_e.number_format = '#,##0" ៛"'
    c_e.font = Font(name=font_family, size=10, color="0C4A8A")

    c_f = ws.cell(row=r, column=6, value=f"=D{r}*E{r}")
    c_f.alignment = Alignment(horizontal="right", vertical="center")
    c_f.number_format = '#,##0" ៛"'
    c_f.font = Font(name=font_family, size=10, color="0C4A8A", bold=True)

    for col_idx in range(1, 7):
      ws.cell(row=r, column=col_idx).border = cell_border

  # 8. Totals
  end_data_r = start_r + len(items_list) - 1 if items_list else start_r
  tot_r1 = end_data_r + 1
  ws.merge_cells(f'A{tot_r1}:E{tot_r1}')
  c_tot1_lbl = ws.cell(row=tot_r1, column=1, value="សរុបទឹកប្រាក់")
  c_tot1_lbl.font = Font(name=font_family, size=11, bold=True)
  c_tot1_lbl.alignment = Alignment(horizontal="right", vertical="center")

  if items_list:
    c_tot1_val = ws.cell(row=tot_r1, column=6, value=f"=SUM(F{start_r}:F{end_data_r})")
  else:
    c_tot1_val = ws.cell(row=tot_r1, column=6, value=0)
  c_tot1_val.font = Font(name=font_family, size=11, bold=True, color="0C4A8A")
  c_tot1_val.alignment = Alignment(horizontal="right", vertical="center")
  c_tot1_val.number_format = '#,##0" ៛"'

  for c_i in range(1, 7):
    ws.cell(row=tot_r1, column=c_i).border = cell_border
  ws.row_dimensions[tot_r1].height = 22

  # Rounded total row
  tot_r2 = tot_r1 + 1
  ws.merge_cells(f'A{tot_r2}:E{tot_r2}')
  c_tot2_lbl = ws.cell(row=tot_r2, column=1, value="ថវិកាសរុបស្នើសុំទូទាត់ (បង្គត់លេខ) *")
  c_tot2_lbl.font = Font(name=font_family, size=11, bold=True)
  c_tot2_lbl.alignment = Alignment(horizontal="right", vertical="center")

  calc_tot = sum([float(it.get('total_price', float(it.get('qty', 0)) * float(it.get('unit_price', 0)))) for it in items_list])
  c_tot2_val = ws.cell(row=tot_r2, column=6, value=round_khmer_currency(calc_tot))
  c_tot2_val.font = Font(name=font_family, size=12, bold=True, color="0C4A8A")
  c_tot2_val.alignment = Alignment(horizontal="right", vertical="center")
  c_tot2_val.number_format = '#,##0" ៛"'

  for c_i in range(1, 7):
    ws.cell(row=tot_r2, column=c_i).border = tot_border
  ws.row_dimensions[tot_r2].height = 24

  # 9. Note
  note_r = tot_r2 + 1
  ws.merge_cells(f'A{note_r}:F{note_r}')
  ws.cell(row=note_r, column=1, value="សម្គាល់៖ *ថវិកាសរុបត្រូវទូទាត់ស្នើសុំត្រូវបង្គត់៖ បង្គត់ស្មើ ០ បើខ្ទង់ដប់តិចជាង ៥០ រៀល បង្គត់ឡើងមួយ បើខ្ទង់ដប់ស្មើឬច្រើនជាង ៥០ រៀល ។").font = Font(name=font_family, size=9.5, italic=True)
  ws.row_dimensions[note_r].height = 20

  # 10. Signatures
  sig_r1 = note_r + 2
  ws.merge_cells(f'A{sig_r1}:B{sig_r1}')
  ws.cell(row=sig_r1, column=1, value="បានឃើញ និងឯកភាពដោយ").font = Font(name=font_family, size=10.5, bold=True)
  ws.cell(row=sig_r1, column=1).alignment = Alignment(horizontal="center", vertical="center")

  ws.merge_cells(f'C{sig_r1}:D{sig_r1}')
  ws.cell(row=sig_r1, column=3, value="ផ្ទៀងផ្ទាត់ដោយ").font = Font(name=font_family, size=10.5, bold=True)
  ws.cell(row=sig_r1, column=3).alignment = Alignment(horizontal="center", vertical="center")

  ws.merge_cells(f'E{sig_r1}:F{sig_r1}')
  ws.cell(row=sig_r1, column=5, value="រៀបចំដោយ").font = Font(name=font_family, size=10.5, bold=True)
  ws.cell(row=sig_r1, column=5).alignment = Alignment(horizontal="center", vertical="center")

  sig_r2 = sig_r1 + 1
  ws.merge_cells(f'A{sig_r2}:B{sig_r2}')
  ws.cell(row=sig_r2, column=1, value="ប្រធាន គ.ម.ស ( នាយក ឬនាយិកាសាលា )").font = Font(name=font_family, size=10, bold=True)
  ws.cell(row=sig_r2, column=1).alignment = Alignment(horizontal="center", vertical="center")

  ws.merge_cells(f'C{sig_r2}:D{sig_r2}')
  ws.cell(row=sig_r2, column=3, value="អ្នកផ្គត់ផ្គង់").font = Font(name=font_family, size=10, bold=True)
  ws.cell(row=sig_r2, column=3).alignment = Alignment(horizontal="center", vertical="center")

  ws.merge_cells(f'E{sig_r2}:F{sig_r2}')
  ws.cell(row=sig_r2, column=5, value="នាយឃ្លាំង/បេឡាធិការ").font = Font(name=font_family, size=10, bold=True)
  ws.cell(row=sig_r2, column=5).alignment = Alignment(horizontal="center", vertical="center")

  sig_name_r = sig_r2 + 3
  ws.merge_cells(f'A{sig_name_r}:B{sig_name_r}')
  ws.cell(row=sig_name_r, column=1, value="......................................................").alignment = Alignment(horizontal="center", vertical="center")

  ws.merge_cells(f'C{sig_name_r}:D{sig_name_r}')
  c_sn = ws.cell(row=sig_name_r, column=3, value=supplier_name)
  c_sn.font = Font(name=font_family, size=11, bold=True, color="0C4A8A")
  c_sn.alignment = Alignment(horizontal="center", vertical="center")

  ws.merge_cells(f'E{sig_name_r}:F{sig_name_r}')
  ws.cell(row=sig_name_r, column=5, value="......................................................").alignment = Alignment(horizontal="center", vertical="center")

  out = io.BytesIO()
  wb.save(out)
  return out.getvalue()


# ================= គ្រប់គ្រង LOGIN SESSION =================
if "logged_in" not in st.session_state:
  st.session_state["logged_in"] = False
  st.session_state["user_info"] = None


def login(username, password):
  hp = hash_password(password)
  res = cursor.execute(
      "SELECT id, username, full_name, role FROM users WHERE username=? AND"
      " password=?",
      (username, hp),
  ).fetchone()
  if res:
    st.session_state["logged_in"] = True
    st.session_state["user_info"] = {
        "id": res[0],
        "username": res[1],
        "name": res[2],
        "role": res[3],
    }
    return True
  return False


def logout():
  st.session_state["logged_in"] = False
  st.session_state["user_info"] = None
  st.rerun()


# ================= ទំព័រ LOGIN =================
if not st.session_state.get("logged_in") or not st.session_state.get(
    "user_info"
):
  st.markdown(
      "<h2 style='text-align: center;'>🔐 ចូលប្រើប្រព័ន្ធ POS</h2>",
      unsafe_allow_html=True,
  )
  col1, col2, col3 = st.columns([1, 1.2, 1])

  with col2:
    with st.container(border=True):
      u_name = st.text_input("ឈ្មោះគណនី (Username)")
      u_pass = st.text_input("ពាក្យសម្ងាត់ (Password)", type="password")

      if st.button("ចូលប្រព័ន្ធ (Login)", use_container_width=True):
        if login(u_name, u_pass):
          st.success("ចូលប្រព័ន្ធជោគជ័យ!")
          st.rerun()
        else:
          st.error("ឈ្មោះគណនី ឬពាក្យសម្ងាត់មិនត្រឹមត្រូវ!")

      st.caption("គណនីលំនាំដើម: **admin** / ពាក្យសម្ងាត់: **admin123**")
  st.stop()

# ================= ម៉ឺនុយចំហៀង (SIDEBAR) ពេល LOGIN រួច =================
user_info = st.session_state.get("user_info") or {
    "name": "User",
    "role": "Guest",
    "username": "guest",
}
st.sidebar.markdown(f"### 👤 {user_info['name']}")
st.sidebar.caption(
    f"តួនាទី: **{user_info['role']}** | គណនី: `{user_info['username']}`"
)
if st.sidebar.button("🚪 ចាកចេញ (Logout)", use_container_width=True):
  logout()

st.sidebar.divider()
menu = st.sidebar.radio(
    "ជ្រើសរើសផ្នែក៖",
    [
        "📊 Dashboard",
        "📍 គ្រប់គ្រងទីតាំង និងសាលារៀន",
        "🚚 បញ្ជីគ្រប់គ្រងអ្នកផ្គត់ផ្គង់",
        "📦 បញ្ជីមុខទំនិញ និងតម្លៃ (តាមសាលា / តាមឃុំ)",
        "📝 កត់ត្រា និងចេញវិក្កយបត្រប្រចាំថ្ងៃ",
        "📑 សំណើទូទាត់ប្រចាំខែ",
        "🛒 បញ្ជីទិញទំនិញចូល & ជំពាក់អ្នកផ្គត់ផ្គង់",
        "💰 ចំណូល និងចំណាយ",
        "👥 គ្រប់គ្រងអ្នកប្រើប្រាស់",
    ],
)

# ================= ១. DASHBOARD =================
if menu == "📊 Dashboard":
  st.title("📊 ផ្ទាំងគ្រប់គ្រងទូទៅ (Dashboard)")

  total_sales = (
      cursor.execute("SELECT SUM(total_price) FROM daily_records").fetchone()[0]
      or 0.0
  )
  total_purchases = (
      cursor.execute("SELECT SUM(total_price) FROM purchases").fetchone()[0]
      or 0.0
  )
  total_debt = (
      cursor.execute(
          "SELECT SUM(total_price) FROM purchases WHERE status='ជំពាក់'"
      ).fetchone()[0]
      or 0.0
  )

  c1, c2, c3 = st.columns(3)
  c1.metric("ចំណូលសរុបពីសាលារៀន", format_riel(total_sales))
  c2.metric("ការទិញទំនិញសរុប", format_riel(total_purchases))
  c3.metric(
      "បំណុលជំពាក់អ្នកផ្គត់ផ្គង់",
      format_riel(total_debt),
      delta_color="inverse",
  )

  st.divider()
  st.subheader("សកម្មភាពលក់ចុងក្រោយ")
  df_recent = pd.read_sql_query(
      "SELECT date as [កាលបរិច្ឆេទ], school_name as [សាលារៀន], item_name as"
      " [មុខទំនិញ], phase as [វគ្គ], quantity as [បរិមាណ], unit_price as"
      " [តម្លៃរាយ (៛)], total_price as [សរុប (៛)] FROM daily_records ORDER BY"
      " id DESC LIMIT 8",
      conn,
  )
  st.dataframe(
      add_row_numbers(df_recent), use_container_width=True, hide_index=True
  )

# ================= ២. ទីតាំង និងសាលា (មានមុខងារ កែប្រែ, លុប, និង នាំចូលពីក្រៅ) =================
elif menu == "📍 គ្រប់គ្រងទីតាំង និងសាលារៀន":
  st.title("📍 គ្រប់គ្រង ភូមិ/ឃុំ/ស្រុក/ខេត្ត និងសាលារៀន")
  tab1, tab2 = st.tabs(["📍 គ្រប់គ្រងទីតាំងរដ្ឋបាល", "🏫 គ្រប់គ្រងសាលារៀន"])

  # ----------------- TAB 1: ទីតាំងរដ្ឋបាល -----------------
  with tab1:
    sub_t1, sub_t2, sub_t3 = st.tabs([
        "➕ បញ្ចូលទីតាំងថ្មី",
        "✏️ កែប្រែ / 🗑️ លុបទីតាំង",
        "📥 នាំចូលទីតាំងពីក្រៅ (Excel, Word, CSV, PDF, រូបភាព)",
    ])

    with sub_t1:
      st.subheader("➕ បញ្ចូលទីតាំងរដ្ឋបាលថ្មី (ខេត្ត ស្រុក ឃុំ ភូមិ)")
      c_p, c_d, c_c, c_v = st.columns(4)

      prov_list = get_provinces()
      p_options = (
          ["-- ជ្រើសរើសខេត្ត --"] + prov_list + ["➕ វាយបញ្ចូលខេត្តថ្មី..."]
      )
      with c_p:
        sel_p = st.selectbox("១. ខេត្ត/ក្រុង", p_options, key="t1_p")
        if sel_p == "➕ វាយបញ្ចូលខេត្តថ្មី..." or (
            not prov_list and sel_p == "-- ជ្រើសរើសខេត្ត --"
        ):
          act_p = st.text_input("ឈ្មោះខេត្តថ្មី", key="t1_p_in").strip()
        elif sel_p != "-- ជ្រើសរើសខេត្ត --":
          act_p = sel_p
        else:
          act_p = ""

      dist_list = get_districts(act_p) if act_p else []
      d_options = (
          ["-- ជ្រើសរើសស្រុក --"] + dist_list + ["➕ វាយបញ្ចូលស្រុកថ្មី..."]
      )
      with c_d:
        sel_d = st.selectbox("២. ស្រុក/ខណ្ឌ", d_options, key="t1_d")
        if sel_d == "➕ វាយបញ្ចូលស្រុកថ្មី..." or (
            not dist_list and sel_d == "-- ជ្រើសរើសស្រុក --"
        ):
          act_d = st.text_input("ឈ្មោះស្រុកថ្មី", key="t1_d_in").strip()
        elif sel_d != "-- ជ្រើសរើសស្រុក --":
          act_d = sel_d
        else:
          act_d = ""

      comm_list = get_communes(act_p, act_d) if (act_p and act_d) else []
      c_options = (
          ["-- ជ្រើសរើសឃុំ --"] + comm_list + ["➕ វាយបញ្ចូលឃុំថ្មី..."]
      )
      with c_c:
        sel_c = st.selectbox("៣. ឃុំ/សង្កាត់", c_options, key="t1_c")
        if sel_c == "➕ វាយបញ្ចូលឃុំថ្មី..." or (
            not comm_list and sel_c == "-- ជ្រើសរើសឃុំ --"
        ):
          act_c = st.text_input("ឈ្មោះឃុំថ្មី", key="t1_c_in").strip()
        elif sel_c != "-- ជ្រើសរើសឃុំ --":
          act_c = sel_c
        else:
          act_c = ""

      vill_list = get_villages(act_c) if act_c else []
      v_options = (
          ["-- ជ្រើសរើសភូមិ --"] + vill_list + ["➕ វាយបញ្ចូលភូមិថ្មី..."]
      )
      with c_v:
        sel_v = st.selectbox("៤. ភូមិ", v_options, key="t1_v")
        if sel_v == "➕ វាយបញ្ចូលភូមិថ្មី..." or (
            not vill_list and sel_v == "-- ជ្រើសរើសភូមិ --"
        ):
          act_v = st.text_input("ឈ្មោះភូមិថ្មី", key="t1_v_in").strip()
        elif sel_v != "-- ជ្រើសរើសភូមិ --":
          act_v = sel_v
        else:
          act_v = ""

      if st.button("💾 រក្សាទុកទីតាំងរដ្ឋបាល", use_container_width=True):
        ok, msg = save_location(act_p, act_d, act_c, act_v)
        if ok:
          st.success(f"បានរក្សាទុកទីតាំងជោគជ័យ!")
          st.rerun()
        else:
          st.error(msg)

    with sub_t2:
      st.subheader("✏️ កែប្រែ ឬ 🗑️ លុបទីតាំងរដ្ឋបាល")
      all_locs = cursor.execute(
          "SELECT id, province, district, commune, village FROM locations"
          " ORDER BY id DESC"
      ).fetchall()
      if all_locs:
        loc_dict = {
            f"ID {r[0]}: ខេត្ត {r[1]} > ស្រុក {r[2]} > ឃុំ {r[3]} >"
            f" ភូមិ {r[4] or 'គ្មាន'}": r
            for r in all_locs
        }
        loc_choice = st.selectbox(
            "ជ្រើសរើសទីតាំងដើម្បីកែប្រែ ឬលុប",
            list(loc_dict.keys()),
            key="sel_loc_edit",
        )
        cur_loc = loc_dict[loc_choice]

        c_ep, c_ed, c_ec, c_ev = st.columns(4)
        with c_ep:
          ed_p = st.text_input("ខេត្ត/ក្រុង", value=cur_loc[1], key="ed_p")
        with c_ed:
          ed_d = st.text_input("ស្រុក/ខណ្ឌ", value=cur_loc[2], key="ed_d")
        with c_ec:
          ed_c = st.text_input("ឃុំ/សង្កាត់", value=cur_loc[3], key="ed_c")
        with c_ev:
          ed_v = st.text_input(
              "ភូមិ", value=cur_loc[4] if cur_loc[4] else "", key="ed_v"
          )

        b_c1, b_c2 = st.columns(2)
        with b_c1:
          if st.button("💾 រក្សាទុកការកែប្រែទីតាំង", use_container_width=True):
            if not ed_p.strip() or not ed_d.strip() or not ed_c.strip():
              st.error("សូមកុំទុកឱ្យ ខេត្ត ស្រុក និងឃុំ ទទេ!")
            else:
              cursor.execute(
                  "UPDATE locations SET province=?, district=?, commune=?,"
                  " village=? WHERE id=?",
                  (
                      ed_p.strip(),
                      ed_d.strip(),
                      ed_c.strip(),
                      ed_v.strip(),
                      cur_loc[0],
                  ),
              )
              conn.commit()
              st.success("បានកែប្រែទីតាំងដោយជោគជ័យ!")
              st.rerun()

        with b_c2:
          if st.button(
              f"🗑️ លុបទីតាំង (ID: {cur_loc[0]})",
              use_container_width=True,
              type="primary",
          ):
            cursor.execute("DELETE FROM locations WHERE id=?", (cur_loc[0],))
            conn.commit()
            st.success(f"បានលុបទីតាំង ID {cur_loc[0]} ដោយជោគជ័យ!")
            st.rerun()
      else:
        st.info("មិនទាន់មានទិន្នន័យទីតាំងក្នុងប្រព័ន្ធនៅឡើយទេ។")

    with sub_t3:
      st.subheader("📥 នាំចូលទីតាំងរដ្ឋបាលពីក្រៅ (Excel, Word, CSV, PDF, Image)")
      st.info(
        "💡 គាំទ្រការនាំចូលពី: **Excel (.xlsx, .xls)**, **Word (.docx)**,"
        " **CSV**, **PDF**, និង **រូបភាពថតតារាង (PNG, JPG)**។"
        " លោកអ្នកអាចកែប្រែទិន្នន័យក្នុងតារាងមុនពេលចុចយល់ព្រមនាំចូល!"
      )

      # Sample template download
      loc_tpl = generate_sample_excel(
          ["ខេត្ត/ក្រុង", "ស្រុក/ខណ្ឌ", "ឃុំ/សង្កាត់", "ភូមិ"],
          ["សៀមរាប", "ស្រីស្នំ", "ស្លែងស្ពាន", "ស្លែងស្ពាន"],
      )
      st.download_button(
          "📥 ទាញយកគំរូឯកសារ Excel (Location Template)",
          data=loc_tpl,
          file_name="Template_Locations.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

      up_loc_file = st.file_uploader(
          "ជ្រើសរើសឯកសារទីតាំង (xlsx, docx, csv, pdf, png, jpg)",
          type=["xlsx", "xls", "docx", "doc", "csv", "pdf", "png", "jpg", "jpeg"],
          key="up_loc_file",
      )

      if up_loc_file:
        raw_loc_df = extract_table_from_file(up_loc_file)
        if not raw_loc_df.empty:
          st.success(f"✅ អានទិន្នន័យបានជោគជ័យ! ចំនួន {len(raw_loc_df)} ជួរ")

          # Normalize to standard columns
          cols = list(raw_loc_df.columns)
          map_p = next(
              (c for c in cols if any(k in str(c).lower() for k in ["ខេត្ត", "prov"])),
              cols[0] if len(cols) > 0 else "",
          )
          map_d = next(
              (c for c in cols if any(k in str(c).lower() for k in ["ស្រុក", "dist"])),
              cols[1] if len(cols) > 1 else "",
          )
          map_c = next(
              (c for c in cols if any(k in str(c).lower() for k in ["ឃុំ", "comm"])),
              cols[2] if len(cols) > 2 else "",
          )
          map_v = next(
              (c for c in cols if any(k in str(c).lower() for k in ["ភូមិ", "vill"])),
              cols[3] if len(cols) > 3 else "",
          )

          formatted_loc_df = pd.DataFrame({
              "ខេត្ត/ក្រុង": raw_loc_df[map_p] if map_p in raw_loc_df else "",
              "ស្រុក/ខណ្ឌ": raw_loc_df[map_d] if map_d in raw_loc_df else "",
              "ឃុំ/សង្កាត់": raw_loc_df[map_c] if map_c in raw_loc_df else "",
              "ភូមិ": raw_loc_df[map_v] if map_v in raw_loc_df else "",
          })

          st.markdown("##### ✏️ ផ្ទៀងផ្ទាត់ និងកែប្រែទិន្នន័យមុនពេលរក្សាទុក:")
          edited_loc_df = st.data_editor(
              formatted_loc_df, num_rows="dynamic", use_container_width=True, key="ed_loc_import"
          )

          if st.button("📥 យល់ព្រមនាំចូលទីតាំងទាំងអស់ចូលក្នុង Database", key="btn_save_loc_import", use_container_width=True):
            count = 0
            for _, r in edited_loc_df.iterrows():
              p = str(r.get("ខេត្ត/ក្រុង", "")).strip()
              d = str(r.get("ស្រុក/ខណ្ឌ", "")).strip()
              c = str(r.get("ឃុំ/សង្កាត់", "")).strip()
              v = str(r.get("ភូមិ", "")).strip()
              if p and d and c:
                save_location(p, d, c, v)
                count += 1
            st.success(f"🎉 បាននាំចូលទីតាំងដោយជោគជ័យចំនួន {count} ទីតាំង!")
            st.rerun()
        else:
          st.warning("មិនអាចទាញយកទិន្នន័យតារាងពីឯកសារនេះបានទេ សូមពិនិត្យមើលទ្រង់ទ្រាយឯកសារ!")

    st.divider()
    st.subheader("📋 បញ្ជីទីតាំងរដ្ឋបាលទាំងអស់")
    df_loc = pd.read_sql_query(
        "SELECT province as [ខេត្ត/ក្រុង], district as [ស្រុក/ខណ្ឌ], commune as"
        " [ឃុំ/សង្កាត់], village as [ភូមិ] FROM locations ORDER BY province,"
        " district, commune, village",
        conn,
    )
    st.dataframe(
        add_row_numbers(df_loc), use_container_width=True, hide_index=True
    )

  # ----------------- TAB 2: សាលារៀន -----------------
  with tab2:
    sch_sub1, sch_sub2, sch_sub3 = st.tabs([
        "➕ បញ្ចូលសាលាថ្មី",
        "✏️ កែប្រែ / 🗑️ លុបសាលារៀន",
        "📥 នាំចូលសាលាពីក្រៅ (Excel, Word, CSV, PDF, រូបភាព)",
    ])

    with sch_sub1:
      st.subheader("🏫 បញ្ចូលសាលារៀនថ្មី (ស្វ័យប្រវត្តិចាប់តាមឃុំ)")
      sc_p, sc_d, sc_c = st.columns(3)
      prov_list2 = get_provinces()
      with sc_p:
        s_p_sel = st.selectbox(
            "ខេត្ត/ក្រុង",
            ["-- ជ្រើសរើសខេត្ត --"]
            + prov_list2
            + ["➕ វាយបញ្ចូលខេត្តថ្មី..."],
            key="t2_p",
        )
        if s_p_sel == "➕ វាយបញ្ចូលខេត្តថ្មី...":
          s_act_p = st.text_input("វាយបញ្ចូលខេត្ត", key="t2_p_in").strip()
        elif s_p_sel != "-- ជ្រើសរើសខេត្ត --":
          s_act_p = s_p_sel
        else:
          s_act_p = ""

      dist_list2 = get_districts(s_act_p) if s_act_p else []
      with sc_d:
        s_d_sel = st.selectbox(
            "ស្រុក/ខណ្ឌ",
            ["-- ជ្រើសរើសស្រុក --"]
            + dist_list2
            + ["➕ វាយបញ្ចូលស្រុកថ្មី..."],
            key="t2_d",
        )
        if s_d_sel == "➕ វាយបញ្ចូលស្រុកថ្មី...":
          s_act_d = st.text_input("វាយបញ្ចូលស្រុក", key="t2_d_in").strip()
        elif s_d_sel != "-- ជ្រើសរើសស្រុក --":
          s_act_d = s_d_sel
        else:
          s_act_d = ""

      comm_list2 = (
          get_communes(s_act_p, s_act_d) if (s_act_p and s_act_d) else []
      )
      with sc_c:
        s_c_sel = st.selectbox(
            "ឃុំ/សង្កាត់",
            ["-- ជ្រើសរើសឃុំ --"]
            + comm_list2
            + ["➕ វាយបញ្ចូលឃុំថ្មី..."],
            key="t2_c",
        )
        if s_c_sel == "➕ វាយបញ្ចូលឃុំថ្មី...":
          s_act_c = st.text_input("វាយបញ្ចូលឃុំ", key="t2_c_in").strip()
        elif s_c_sel != "-- ជ្រើសរើសឃុំ --":
          s_act_c = s_c_sel
        else:
          s_act_c = ""

      if s_act_c:
        schools_in_c = get_schools_by_commune(s_act_c)
        if schools_in_c:
          st.success(
              f"🎯 ស្វ័យប្រវត្តិចាប់ឃើញសាលាក្នុងឃុំ **{s_act_c}** ចំនួន"
              f" **{len(schools_in_c)}** ៖ "
              + " | ".join([f"`{s}`" for s in schools_in_c])
          )
        else:
          st.info(
              f"ℹ️ មិនទាន់មានសាលាក្នុងឃុំ **{s_act_c}** នៅឡើយទេ។"
              " សូមវាយបញ្ចូលបន្ថែមខាងក្រោម៖"
          )

      col_sn, col_sv = st.columns([2, 1])
      with col_sn:
        new_sch_name = st.text_input(
            "✏️ ឈ្មោះសាលាបឋមសិក្សា (វាយបញ្ចូលបន្ថែមបើគ្មានក្នុងប្រព័ន្ធ)",
            key="t2_sch_name",
        )
      with col_sv:
        new_sch_vill = st.text_input("ភូមិ (ស្រេចចិត្ត)", key="t2_sch_vill")

      if st.button("💾 រក្សាទុកសាលារៀន", use_container_width=True):
        if not s_act_c:
          st.error("សូមជ្រើសរើស ឬវាយបញ្ចូលឃុំ/សង្កាត់សិន!")
        elif not new_sch_name.strip():
          st.error("សូមវាយបញ្ចូលឈ្មោះសាលារៀន!")
        else:
          if s_act_p and s_act_d and s_act_c:
            save_location(s_act_p, s_act_d, s_act_c, new_sch_vill)
          ok, msg = save_school(
              new_sch_name, s_act_c, s_act_d, s_act_p, new_sch_vill
          )
          if ok:
            st.success(
                f"បានរក្សាទុកសាលា '{new_sch_name}' ក្នុងឃុំ '{s_act_c}'"
                " ដោយជោគជ័យ!"
            )
            st.rerun()
          else:
            st.error(msg)

    with sch_sub2:
      st.subheader("✏️ កែប្រែ ឬ 🗑️ លុបសាលារៀន")
      all_schools_data = cursor.execute(
          "SELECT id, name, commune, district, province, village FROM schools"
          " ORDER BY name"
      ).fetchall()
      if all_schools_data:
        sch_dict = {
            f"ID {r[0]}: {r[1]} (ឃុំ {r[2]} | ស្រុក {r[3] or '-'} |"
            f" ខេត្ត {r[4] or '-'})": r
            for r in all_schools_data
        }
        sch_choice = st.selectbox(
            "ជ្រើសរើសសាលាដើម្បីកែប្រែ ឬលុប",
            list(sch_dict.keys()),
            key="sel_sch_edit",
        )
        cur_sch = sch_dict[sch_choice]

        c_en, c_ec = st.columns([2, 1])
        with c_en:
          ed_sname = st.text_input(
              "ឈ្មោះសាលារៀន", value=cur_sch[1], key="ed_sname"
          )
        with c_ec:
          ed_scomm = st.text_input(
              "ឃុំ/សង្កាត់", value=cur_sch[2], key="ed_scomm"
          )

        c_ep, c_ed, c_ev = st.columns(3)
        with c_ep:
          ed_sprov = st.text_input(
              "ខេត្ត/ក្រុង",
              value=cur_sch[4] if cur_sch[4] else "",
              key="ed_sprov",
          )
        with c_ed:
          ed_sdist = st.text_input(
              "ស្រុក/ខណ្ឌ",
              value=cur_sch[3] if cur_sch[3] else "",
              key="ed_sdist",
          )
        with c_ev:
          ed_svill = st.text_input(
              "ភូមិ", value=cur_sch[5] if cur_sch[5] else "", key="ed_svill"
          )

        sb_c1, sb_c2 = st.columns(2)
        with sb_c1:
          if st.button("💾 រក្សាទុកការកែប្រែសាលា", use_container_width=True):
            if not ed_sname.strip() or not ed_scomm.strip():
              st.error("ឈ្មោះសាលា និង ឃុំ មិនអាចទុកឱ្យទទេបានទេ!")
            else:
              cursor.execute(
                  "UPDATE schools SET name=?, commune=?, district=?,"
                  " province=?, village=? WHERE id=?",
                  (
                      ed_sname.strip(),
                      ed_scomm.strip(),
                      ed_sdist.strip(),
                      ed_sprov.strip(),
                      ed_svill.strip(),
                      cur_sch[0],
                  ),
              )
              conn.commit()
              st.success(f"បានកែប្រែព័ត៌មានសាលា ID {cur_sch[0]} ដោយជោគជ័យ!")
              st.rerun()

        with sb_c2:
          if st.button(
              f"🗑️ លុបសាលានេះ (ID: {cur_sch[0]})",
              use_container_width=True,
              type="primary",
          ):
            cursor.execute("DELETE FROM schools WHERE id=?", (cur_sch[0],))
            conn.commit()
            st.success(f"បានលុបសាលារៀន ID {cur_sch[0]} រួចរាល់!")
            st.rerun()
      else:
        st.info("មិនទាន់មានទិន្នន័យសាលារៀនក្នុងប្រព័ន្ធនៅឡើយទេ។")

    with sch_sub3:
      st.subheader("📥 នាំចូលសាលារៀនពីក្រៅ (Excel, Word, CSV, PDF, Image)")
      st.info(
        "💡 គាំទ្រការនាំចូលពី: **Excel**, **Word**, **CSV**, **PDF**, និង **រូបភាព (PNG, JPG)**។"
        " លោកអ្នកអាចផ្ទៀងផ្ទាត់ និងកែសម្រួលមុនពេលរក្សាទុក!"
      )

      # Sample template download
      sch_tpl = generate_sample_excel(
          ["ឈ្មោះសាលារៀន", "ឃុំ/សង្កាត់", "ស្រុក/ខណ្ឌ", "ខេត្ត/ក្រុង", "ភូមិ"],
          ["សាលាបឋមសិក្សា ស្លែងស្ពាន", "ស្លែងស្ពាន", "ស្រីស្នំ", "សៀមរាប", "ស្លែងស្ពាន"],
      )
      st.download_button(
          "📥 ទាញយកគំរូឯកសារ Excel (School Template)",
          data=sch_tpl,
          file_name="Template_Schools.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

      up_sch_file = st.file_uploader(
          "ជ្រើសរើសឯកសារសាលារៀន (xlsx, docx, csv, pdf, png, jpg)",
          type=["xlsx", "xls", "docx", "doc", "csv", "pdf", "png", "jpg", "jpeg"],
          key="up_sch_file",
      )

      if up_sch_file:
        raw_sch_df = extract_table_from_file(up_sch_file)
        if not raw_sch_df.empty:
          st.success(f"✅ អានទិន្នន័យបានជោគជ័យ! ចំនួន {len(raw_sch_df)} ជួរ")

          cols = list(raw_sch_df.columns)
          map_sn = next(
              (c for c in cols if any(k in str(c).lower() for k in ["សាលា", "school", "name"])),
              cols[0] if len(cols) > 0 else "",
          )
          map_sc = next(
              (c for c in cols if any(k in str(c).lower() for k in ["ឃុំ", "comm"])),
              cols[1] if len(cols) > 1 else "",
          )
          map_sd = next(
              (c for c in cols if any(k in str(c).lower() for k in ["ស្រុក", "dist"])),
              cols[2] if len(cols) > 2 else "",
          )
          map_sp = next(
              (c for c in cols if any(k in str(c).lower() for k in ["ខេត្ត", "prov"])),
              cols[3] if len(cols) > 3 else "",
          )
          map_sv = next(
              (c for c in cols if any(k in str(c).lower() for k in ["ភូមិ", "vill"])),
              cols[4] if len(cols) > 4 else "",
          )

          formatted_sch_df = pd.DataFrame({
              "ឈ្មោះសាលារៀន": raw_sch_df[map_sn] if map_sn in raw_sch_df else "",
              "ឃុំ/សង្កាត់": raw_sch_df[map_sc] if map_sc in raw_sch_df else "",
              "ស្រុក/ខណ្ឌ": raw_sch_df[map_sd] if map_sd in raw_sch_df else "",
              "ខេត្ត/ក្រុង": raw_sch_df[map_sp] if map_sp in raw_sch_df else "",
              "ភូមិ": raw_sch_df[map_sv] if map_sv in raw_sch_df else "",
          })

          st.markdown("##### ✏️ ផ្ទៀងផ្ទាត់ និងកែប្រែទិន្នន័យសាលារៀន:")
          edited_sch_df = st.data_editor(
              formatted_sch_df, num_rows="dynamic", use_container_width=True, key="ed_sch_import"
          )

          if st.button("📥 យល់ព្រមនាំចូលសាលារៀនទាំងអស់ចូលក្នុង Database", key="btn_save_sch_import", use_container_width=True):
            count = 0
            for _, r in edited_sch_df.iterrows():
              sn = str(r.get("ឈ្មោះសាលារៀន", "")).strip()
              sc = str(r.get("ឃុំ/សង្កាត់", "")).strip()
              sd = str(r.get("ស្រុក/ខណ្ឌ", "")).strip()
              sp = str(r.get("ខេត្ត/ក្រុង", "")).strip()
              sv = str(r.get("ភូមិ", "")).strip()
              if sn and sc:
                if sp and sd and sc:
                  save_location(sp, sd, sc, sv)
                save_school(sn, sc, sd, sp, sv)
                count += 1
            st.success(f"🎉 បាននាំចូលសាលារៀនចំនួន {count} សាលាដោយជោគជ័យ!")
            st.rerun()
        else:
          st.warning("មិនអាចទាញយកទិន្នន័យតារាងពីឯកសារនេះបានទេ!")

    st.divider()
    st.subheader("📋 បញ្ជីសាលារៀនទាំងអស់ក្នុងប្រព័ន្ធ")
    df_sch = pd.read_sql_query(
        "SELECT name as [ឈ្មោះសាលារៀន], commune as [ឃុំ/សង្កាត់], district as"
        " [ស្រុក/ខណ្ឌ], province as [ខេត្ត/ក្រុង], village as [ភូមិ] FROM"
        " schools ORDER BY id DESC",
        conn,
    )
    st.dataframe(
        add_row_numbers(df_sch), use_container_width=True, hide_index=True
    )

# ================= ៣. បញ្ជីគ្រប់គ្រងអ្នកផ្គត់ផ្គង់ =================
elif menu in ["🚚 បញ្ជីគ្រប់គ្រងអ្នកផ្គត់ផ្គង់", "🚚 គ្រប់គ្រងអ្នកផ្គត់ផ្គង់តាមសាលា", "🚚 គ្រប់គ្រងអ្នកផ្គត់ផ្គង់"]:
  st.title("🚚 បញ្ជីគ្រប់គ្រងអ្នកផ្គត់ផ្គង់ និងកំណត់តម្លៃទំនិញ")
  st.info(
      "💡 គ្រប់គ្រងព័ត៌មានផ្ទាល់ខ្លួនអ្នកផ្គត់ផ្គង់, អាស័យដ្ឋាន, ហត្ថលេខា (ជាមួយប៊ូតុងលុបផ្ទៃខាងក្រោយ), ពត៌មានផ្គត់ផ្គង់ (កម្រិតឃុំ ឬសាលា) ព្រមទាំងកំណត់តម្លៃវគ្គ១/វគ្គ២ និងផ្ទៀងផ្ទាត់ធៀបនឹងតម្លៃគោលដោយស្វ័យប្រវត្តិ។"
  )

  sup_tab1, sup_tab2, sup_tab3 = st.tabs([
      "📋 បញ្ជីអ្នកផ្គត់ផ្គង់ទាំងអស់",
      "➕ បន្ថែម / កែប្រែព័ត៌មាន & តម្លៃទំនិញ",
      "📥 នាំចូលពីក្រៅ (Excel, CSV, Word, PDF)"
  ])

  # ----------------- TAB 1: បញ្ជីអ្នកផ្គត់ផ្គង់ទាំងអស់ -----------------
  with sup_tab1:
    st.subheader("📋 បញ្ជីអ្នកផ្គត់ផ្គង់ស្បៀងទាំងអស់ក្នុងប្រព័ន្ធ")
    all_sups = get_all_suppliers()
    if all_sups:
      sup_display_list = []
      for s in all_sups:
        addr_parts = []
        if s["village"]: addr_parts.append(str(s["village"]))
        if s["commune"]: addr_parts.append(f"ឃុំ{s['commune']}" if not str(s['commune']).startswith("ឃុំ") else str(s['commune']))
        if s["district"]: addr_parts.append(f"ស្រុក{s['district']}" if not str(s['district']).startswith("ស្រុក") and not str(s['district']).startswith("ក្រុង") else str(s['district']))
        if s["province"]: addr_parts.append(f"ខេត្ត{s['province']}" if not str(s['province']).startswith("ខេត្ត") else str(s['province']))
        full_addr = " ".join(addr_parts) if addr_parts else "-"

        lvl_label = "🏛️ តាមឃុំ" if s["supply_level"] == "commune" else "🏫 តាមសាលា"
        target_dest = (s["target_commune"] or s["commune"] or "-") if s["supply_level"] == "commune" else (s["school_name"] or "-")
        sig_stat = "✅ មានហត្ថលេខា" if (s["signature_data"] and str(s["signature_data"]).startswith("data:image")) else "⚪ គ្មានហត្ថលេខា"

        p_cnt_row = cursor.execute("SELECT COUNT(*) FROM products WHERE supplier_name=? AND price_level='supplier'", (s["supplier_name"],)).fetchone()
        p_count = p_cnt_row[0] if p_cnt_row else 0
        price_stat = f"✅ កំណត់រួច ({p_count} មុខ)" if p_count > 0 else "⚪ ប្រើតម្លៃគោល"

        sup_display_list.append({
            "ID": s["id"],
            "ឈ្មោះអ្នកផ្គត់ផ្គង់": s["supplier_name"],
            "លេខទូរស័ព្ទ": s["phone"] or "-",
            "អាស័យដ្ឋានអ្នកផ្គត់ផ្គង់": full_addr,
            "កម្រិតផ្គត់ផ្គង់": lvl_label,
            "គោលដៅផ្គត់ផ្គង់": target_dest,
            "ហត្ថលេខា": sig_stat,
            "ស្ថានភាពតម្លៃ": price_stat
        })

      df_sups = pd.DataFrame(sup_display_list)
      st.dataframe(add_row_numbers(df_sups), use_container_width=True, hide_index=True)

      with st.expander("👁️ មើលព័ត៌មានលម្អិត & រូបភាពហត្ថលេខារបស់អ្នកផ្គត់ផ្គង់", expanded=False):
        c_pick = st.selectbox(
            "ជ្រើសរើសអ្នកផ្គត់ផ្គង់ដើម្បីពិនិត្យលម្អិត៖",
            [f"👤 {s['supplier_name']} ({'សាលា: ' + s['school_name'] if s['supply_level'] != 'commune' and s['school_name'] else 'ឃុំ: ' + s['target_commune']}) - ID: {s['id']}" for s in all_sups],
            key="view_sup_detail_pick"
        )
        if c_pick:
          picked_id = int(c_pick.split("ID: ")[-1])
          picked_s = next((s for s in all_sups if s["id"] == picked_id), None)
          if picked_s:
            cd_col1, cd_col2 = st.columns([2, 1])
            with cd_col1:
              st.markdown(f"#### 👤 {picked_s['supplier_name']}")
              st.markdown(f"📞 **លេខទូរស័ព្ទ៖** `{picked_s['phone'] or '-'}`")
              st.markdown(f"📍 **អាស័យដ្ឋានផ្ទាល់ខ្លួន៖** {picked_s['village']} {picked_s['commune']} {picked_s['district']} {picked_s['province']}")
              st.markdown(f"🎯 **កម្រិតផ្គត់ផ្គង់៖** `{'តាមឃុំ (Commune Level)' if picked_s['supply_level'] == 'commune' else 'តាមសាលា (School Level)'}`")
              st.markdown(f"🏫 **គោលដៅផ្គត់ផ្គង់៖** `{picked_s['school_name'] if picked_s['supply_level'] != 'commune' else picked_s['target_commune']}`")
            with cd_col2:
              st.markdown("**✍️ ហត្ថលេខាអ្នកផ្គត់ផ្គង់៖**")
              if picked_s["signature_data"] and str(picked_s["signature_data"]).startswith("data:image"):
                st.image(picked_s["signature_data"], width=170)
              else:
                st.caption("*(គ្មានហត្ថលេខាឌីជីថល - ទុកចន្លោះចុះហត្ថលេខាផ្ទាល់ដៃ)*")
    else:
      st.info("💡 មិនទាន់មានអ្នកផ្គត់ផ្គង់ក្នុងប្រព័ន្ធនៅឡើយទេ។ សូមចុចផ្ទាំង «➕ បន្ថែម / កែប្រែព័ត៌មាន & តម្លៃទំនិញ» ដើម្បីបង្កើតថ្មី។")

  # ----------------- TAB 2: បន្ថែម / កែប្រែព័ត៌មាន & តម្លៃទំនិញ -----------------
  with sup_tab2:
    all_sups_edit = get_all_suppliers()
    sup_select_options = ["➕ បង្កើតអ្នកផ្គត់ផ្គង់ថ្មី..."] + [
        f"👤 {s['supplier_name']} ({'សាលា: ' + s['school_name'] if s['supply_level'] != 'commune' and s['school_name'] else 'ឃុំ: ' + (s['target_commune'] or s['commune'])}) - ID: {s['id']}"
        for s in all_sups_edit
    ]
    
    st.markdown("#### 🎯 ជ្រើសរើសប្រតិបត្តិការ")
    chosen_op = st.selectbox(
        "ជ្រើសរើសអ្នកផ្គត់ផ្គង់ដើម្បីកែប្រែ ឬបង្កើតថ្មី៖",
        sup_select_options,
        key="sup_manage_op_select"
    )
    is_new = (chosen_op == "➕ បង្កើតអ្នកផ្គត់ផ្គង់ថ្មី...")
    cur_sup = None
    if not is_new:
      target_id = int(chosen_op.split("ID: ")[-1])
      cur_sup = next((s for s in all_sups_edit if s["id"] == target_id), None)

    sup_id_for_key = cur_sup["id"] if cur_sup else "new"

    st.divider()

    # 1. ព័ត៌មានផ្ទាល់ខ្លួន
    st.markdown("#### 👤 ១. ព័ត៌មានផ្ទាល់ខ្លួនអ្នកផ្គត់ផ្គង់")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
      val_name = st.text_input("ឈ្មោះអ្នកផ្គត់ផ្គង់ *", value=(cur_sup["supplier_name"] if cur_sup else ""), placeholder="ឧ. សាត ក្រូត", key=f"inp_sup_name_{sup_id_for_key}").strip()
    with col_p2:
      val_phone = st.text_input("លេខទូរស័ព្ទ", value=(cur_sup["phone"] if cur_sup else ""), placeholder="ឧ. 090 854 133", key=f"inp_sup_phone_{sup_id_for_key}").strip()

    # 2. អាស័យដ្ឋានអ្នកផ្គត់ផ្គង់
    st.markdown("#### 📍 ២. អាស័យដ្ឋានអ្នកផ្គត់ផ្គង់ (ផ្ទាល់ខ្លួន)")
    col_a1, col_a2, col_a3, col_a4 = st.columns(4)
    with col_a1:
      val_prov = st.text_input("ខេត្ត", value=(cur_sup["province"] if cur_sup and cur_sup["province"] else "សៀមរាប"), key=f"inp_sup_prov_{sup_id_for_key}").strip()
    with col_a2:
      val_dist = st.text_input("ក្រុង/ស្រុក/ខណ្ឌ", value=(cur_sup["district"] if cur_sup and cur_sup["district"] else "ស្រីស្នំ"), key=f"inp_sup_dist_{sup_id_for_key}").strip()
    with col_a3:
      val_comm = st.text_input("ឃុំ/សង្កាត់", value=(cur_sup["commune"] if cur_sup and cur_sup["commune"] else "រោង"), key=f"inp_sup_comm_{sup_id_for_key}").strip()
    with col_a4:
      val_vill = st.text_input("ភូមិ", value=(cur_sup["village"] if cur_sup and cur_sup["village"] else "ភូមិខ្មែរ"), key=f"inp_sup_vill_{sup_id_for_key}").strip()

    # 3. ហត្ថលេខាអ្នកផ្គត់ផ្គង់
    st.markdown("#### ✍️ ៣. ហត្ថលេខាអ្នកផ្គត់ផ្គង់ (មានប៊ូតុងលុបផ្ទៃខាងក្រោយ & កែប្រែពណ៌)")
    existing_sig_val = cur_sup["signature_data"] if cur_sup else None
    
    col_sig_tool, col_sig_prev = st.columns([2, 1])
    with col_sig_tool:
      uploaded_sig_data = render_signature_uploader_with_tools(
          label="បញ្ចូល ឬផ្លាស់ប្ដូររូបភាពហត្ថលេខា (PNG / JPG / WebP)",
          key_prefix=f"tool_sig_{sup_id_for_key}",
          default_sig_b64=existing_sig_val,
          allow_use_saved=True if existing_sig_val else False,
          saved_sig_b64=existing_sig_val,
          allow_blank_choice=True,
          default_recolor="blue"
      )
    with col_sig_prev:
      st.markdown("**👁️ ការបង្ហាញហត្ថលេខា៖**")
      if uploaded_sig_data and str(uploaded_sig_data).startswith("data:image"):
        st.image(uploaded_sig_data, width=150, caption="ហត្ថលេខាសកម្ម")
      elif existing_sig_val and str(existing_sig_val).startswith("data:image"):
        st.image(existing_sig_val, width=150, caption="ហត្ថលេខាដើម")
      else:
        st.caption("*(ទុកចន្លោះ ........... ចុះហត្ថលេខាផ្ទាល់ដៃ)*")

    # 4. ព័ត៌មានអំពីការផ្គត់ផ្គង់
    st.markdown("#### 🚚 ៤. ព័ត៌មានអំពីការផ្គត់ផ្គង់ (កម្រិតផ្គត់ផ្គង់ & អាស័យដ្ឋានគោលដៅ)")
    default_scope_idx = 1 if (cur_sup and cur_sup["supply_level"] == "school") else 0
    c_scope_rad, c_target_loc = st.columns([1.5, 2.5])
    with c_scope_rad:
      chosen_scope = st.radio(
          "🎯 ជម្រើសកម្រិតផ្គត់ផ្គង់៖",
          ["🏛️ តាមឃុំ (Commune Level)", "🏫 តាមសាលា (School Level)"],
          index=default_scope_idx,
          key=f"scope_radio_{sup_id_for_key}"
      )
      is_commune_level = "តាមឃុំ" in chosen_scope

    with c_target_loc:
      st.markdown("**📍 អាស័យដ្ឋានគោលដៅផ្គត់ផ្គង់៖**")
      col_loc1, col_loc2, col_loc3 = st.columns(3)
      with col_loc1:
        cur_t_prv = cur_sup["target_province"] if cur_sup and cur_sup["target_province"] else (cur_sup["province"] if cur_sup else "សៀមរាប")
        provinces_all = get_provinces()
        prv_idx = provinces_all.index(cur_t_prv) if cur_t_prv in provinces_all else 0
        tgt_province = st.selectbox("ខេត្តគោលដៅ", provinces_all, index=prv_idx, key=f"t_prv_{sup_id_for_key}")
      with col_loc2:
        dist_options = get_districts(province=tgt_province)
        cur_t_dst = cur_sup["target_district"] if cur_sup and cur_sup["target_district"] else (cur_sup["district"] if cur_sup else "")
        dst_idx = dist_options.index(cur_t_dst) if cur_t_dst in dist_options else 0
        tgt_district = st.selectbox("ក្រុង/ស្រុកគោលដៅ", dist_options if dist_options else ["ស្រីស្នំ"], index=dst_idx, key=f"t_dst_{sup_id_for_key}")
      with col_loc3:
        comm_options = get_communes(district=tgt_district)
        cur_t_com = cur_sup["target_commune"] if cur_sup and cur_sup["target_commune"] else (cur_sup["commune"] if cur_sup else "")
        com_idx = comm_options.index(cur_t_com) if cur_t_com in comm_options else 0
        tgt_commune = st.selectbox("ឃុំ/សង្កាត់គោលដៅ", comm_options if comm_options else ["រោង"], index=com_idx, key=f"t_com_{sup_id_for_key}")

    if not is_commune_level:
      school_candidates = get_schools_by_commune(tgt_commune) if tgt_commune else get_all_schools()
      cur_target_sch = cur_sup["school_name"] if cur_sup else ""
      sch_idx = school_candidates.index(cur_target_sch) if cur_target_sch in school_candidates else 0
      col_s1, col_s2 = st.columns([2, 1])
      with col_s1:
        sel_sch = st.selectbox("🏫 សាលារៀនដែលត្រូវផ្គត់ផ្គង់ *", school_candidates + ["➕ វាយបញ្ចូលសាលាថ្មី..."], index=sch_idx, key=f"t_sch_sel_{sup_id_for_key}")
      with col_s2:
        if sel_sch == "➕ វាយបញ្ចូលសាលាថ្មី...":
          tgt_school = st.text_input("វាយបញ្ចូលឈ្មោះសាលាថ្មី", key=f"t_sch_new_{sup_id_for_key}").strip()
        else:
          tgt_school = sel_sch
    else:
      tgt_school = ""
      st.info(f"🏛️ អ្នកផ្គត់ផ្គង់នេះ ផ្គត់ផ្គង់គ្រប់សាលារៀនទាំងអស់នៅក្នុង **ឃុំ{tgt_commune}**")

    # 5. កំណត់តម្លៃទំនិញផ្គត់ផ្គង់ & ប្រៀបធៀបតម្លៃគោល
    st.markdown("#### 💰 ៥. កំណត់តម្លៃទំនិញផ្គត់ផ្គង់ & ផ្ទៀងផ្ទាត់ធៀបនឹងតម្លៃគោល")
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); border: 1.5px solid #cbd5e1; border-radius: 10px; padding: 12px 16px; margin-bottom: 14px;">
      <div style="font-weight: 700; color: #1e293b; margin-bottom: 6px; font-size: 14px;">🎯 គោលការណ៍ហាយឡាយពណ៌ប្រៀបធៀបតម្លៃមធ្យម ធៀបនឹងតម្លៃគោល៖</div>
      <div style="display: flex; gap: 15px; flex-wrap: wrap; font-size: 13px;">
        <span style="background: #fee2e2; border: 1px solid #ef4444; color: #991b1b; padding: 4px 10px; border-radius: 6px; font-weight: 600;">🔴 ពណ៌ក្រហម៖ ខ្ពស់ជាងតម្លៃគោលលើសពី ១០% (&gt; +10%)</span>
        <span style="background: #fef9c3; border: 1px solid #eab308; color: #854d0e; padding: 4px 10px; border-radius: 6px; font-weight: 600;">🟡 ពណ៌លឿង៖ ទាបជាងតម្លៃគោលលើសពី ១០% (&lt; -10%)</span>
        <span style="background: #ecfccb; border: 1px solid #84cc16; color: #365314; padding: 4px 10px; border-radius: 6px; font-weight: 600;">🟢 ពណ៌ត្រួយចេក៖ ស្ថិតនៅចន្លោះតម្លៃគោល ±១០% (កម្រិតសមស្រប)</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Initialize price state for this supplier
    mat_key = f"sup_mat_{sup_id_for_key}"
    if mat_key not in st.session_state:
      st.session_state[mat_key] = {}
      saved_sup_prices = get_supplier_prices_map(cur_sup["supplier_name"]) if cur_sup else {}
      for it in SUPPLIER_PRODUCT_CATALOG:
        inm = it["name"]
        if inm in saved_sup_prices:
          st.session_state[mat_key][inm] = {
              "p1": float(saved_sup_prices[inm]["p1"]),
              "p2": float(saved_sup_prices[inm]["p2"])
          }
        else:
          st.session_state[mat_key][inm] = {
              "p1": float(it["default_p1"]),
              "p2": float(it["default_p2"])
          }

    # Group & Category Filter Buttons
    col_btn_grp, col_btn_cat = st.columns(2)
    with col_btn_grp:
      sel_grp = st.segmented_control(
          "🏷️ ប៊ូតុងជ្រើសរើសក្រុមទំនិញ៖",
          options=["🌟 ទាំងអស់", "🌾 ស្បៀងគោល", "🥦 បន្លែគោល", "🥩 ស្បៀងបន្ថែម", "🥕 បន្លែបន្ថែម"],
          default="🌟 ទាំងអស់",
          key=f"ctrl_grp_{sup_id_for_key}"
      ) or "🌟 ទាំងអស់"
    with col_btn_cat:
      sel_cat = st.segmented_control(
          "🛒 ប៊ូតុងជ្រើសរើសប្រភេទទំនិញ៖",
          options=["🌟 ទាំងអស់", "🍚 អង្ករ", "🫗 ប្រេងឆា", "🧂 អំបិល", "🥩 ត្រី សាច់ ស៊ុត", "🥬 បន្លែ"],
          default="🌟 ទាំងអស់",
          key=f"ctrl_cat_{sup_id_for_key}"
      ) or "🌟 ទាំងអស់"

    clean_grp = sel_grp.replace("🌾 ", "").replace("🥦 ", "").replace("🥩 ", "").replace("🥕 ", "").replace("🌟 ", "").strip()
    clean_cat = sel_cat.replace("🍚 ", "").replace("🫗 ", "").replace("🧂 ", "").replace("🥩 ", "").replace("🥬 ", "").replace("🌟 ", "").strip()

    filtered_prods = []
    for it in SUPPLIER_PRODUCT_CATALOG:
      if clean_grp != "ទាំងអស់" and it["group"] != clean_grp:
        continue
      if clean_cat != "ទាំងអស់" and it["category"] != clean_cat:
        continue
      filtered_prods.append(it)

    # Calculate real-time statistics across all 61 items
    cnt_red = 0
    cnt_yellow = 0
    cnt_lime = 0
    for it in SUPPLIER_PRODUCT_CATALOG:
      inm = it["name"]
      p_vals = st.session_state[mat_key].get(inm, {"p1": it["default_p1"], "p2": it["default_p2"]})
      p1_v = float(p_vals.get("p1", it["default_p1"]))
      p2_v = float(p_vals.get("p2", it["default_p2"]))
      avg_v = round((p1_v + p2_v) / 2.0, 2)
      base_v = get_catalog_base_price(inm, tgt_school if not is_commune_level else None, tgt_commune)
      ev_st = evaluate_supplier_price_status(avg_v, base_v)
      if ev_st["status"] == "high": cnt_red += 1
      elif ev_st["status"] == "low": cnt_yellow += 1
      elif ev_st["status"] == "normal_lime": cnt_lime += 1

    # KPI summary bar
    kp1, kp2, kp3, kp4 = st.columns(4)
    with kp1:
      st.metric("📦 កំពុងបង្ហាញ", f"{len(filtered_prods)} មុខ", delta=f"សរុប {len(SUPPLIER_PRODUCT_CATALOG)} មុខ")
    with kp2:
      st.metric("🔴 ថ្លៃជាងគោល >10%", f"{cnt_red} មុខ", delta="ក្រហម" if cnt_red > 0 else "គ្មាន", delta_color="inverse")
    with kp3:
      st.metric("🟡 ថោកជាងគោល >10%", f"{cnt_yellow} មុខ", delta="លឿង" if cnt_yellow > 0 else "គ្មាន", delta_color="off")
    with kp4:
      st.metric("🟢 ត្រួយចេក (±10%)", f"{cnt_lime} មុខ", delta="សមស្រប", delta_color="normal")

    st.markdown("---")

    if not filtered_prods:
      st.warning(f"⚠️ គ្មានមុខទំនិញដែលត្រូវគ្នានឹងក្រុម «{clean_grp}» និងប្រភេទ «{clean_cat}» ទេ។ សូមជ្រើសរើស «🌟 ទាំងអស់»។")
    else:
      # Display products vertically
      for it in filtered_prods:
        inm = it["name"]
        unit = it["unit"]
        grp = it["group"]
        cat = it["category"]
        base_p = get_catalog_base_price(inm, tgt_school if not is_commune_level else None, tgt_commune)

        # Retrieve current prices
        cur_p_entry = st.session_state[mat_key].get(inm, {"p1": it["default_p1"], "p2": it["default_p2"]})
        cur_p1 = float(cur_p_entry.get("p1", it["default_p1"]))
        cur_p2 = float(cur_p_entry.get("p2", it["default_p2"]))
        cur_avg = round((cur_p1 + cur_p2) / 2.0, 2)
        cur_eval = evaluate_supplier_price_status(cur_avg, base_p)

        # Render styled item container
        st.markdown(f"""
        <div style="background-color: {cur_eval['bg_color']}; border: 1.5px solid {cur_eval['border_color']}; border-radius: 8px; padding: 10px 14px; margin-top: 12px; margin-bottom: 6px;">
          <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
            <div>
              <span style="font-size: 16px; font-weight: 700; color: #0f172a;">{inm}</span>
              <span style="font-size: 13px; color: #475569; margin-left: 6px; font-weight: 500;">({unit})</span>
              <span style="background: rgba(0,0,0,0.06); padding: 2px 8px; border-radius: 12px; font-size: 11px; margin-left: 8px; font-weight: 600; color: #334155;">{grp}</span>
              <span style="background: rgba(0,0,0,0.04); padding: 2px 8px; border-radius: 12px; font-size: 11px; margin-left: 4px; color: #64748b;">{cat}</span>
              <div style="font-size: 12px; color: #475569; margin-top: 3px;">
                🏷️ តម្លៃគោល៖ <b>{format_riel(base_p)}</b> &nbsp;|&nbsp; ចន្លោះ ±10%៖ <b>{format_riel(base_p * 0.9)}</b> ~ <b>{format_riel(base_p * 1.1)}</b>
              </div>
            </div>
            <div style="text-align: right;">
              <div style="font-size: 15px; font-weight: 800; color: {cur_eval['text_color']};">
                តម្លៃមធ្យម៖ {format_riel(cur_avg)}
              </div>
              <div style="font-size: 12px; font-weight: 700; color: {cur_eval['text_color']}; margin-top: 2px;">
                {cur_eval['badge_text']}
              </div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        col_in1, col_in2, col_disp_avg = st.columns([2, 2, 2])
        with col_in1:
          val_p1 = st.number_input(
              f"តម្លៃវគ្គ១ (៛) - {inm}",
              min_value=0.0,
              max_value=500000.0,
              step=100.0,
              value=cur_p1,
              key=f"inp_p1_{mat_key}_{inm}",
              help="តម្លៃក្នុងវគ្គទី ១"
          )
          st.session_state[mat_key][inm]["p1"] = val_p1
        with col_in2:
          val_p2 = st.number_input(
              f"តម្លៃវគ្គ២ (៛) - {inm}",
              min_value=0.0,
              max_value=500000.0,
              step=100.0,
              value=cur_p2,
              key=f"inp_p2_{mat_key}_{inm}",
              help="តម្លៃក្នុងវគ្គទី ២"
          )
          st.session_state[mat_key][inm]["p2"] = val_p2
        with col_disp_avg:
          re_avg = round((val_p1 + val_p2) / 2.0, 2)
          st.text_input(
              f"តម្លៃមធ្យមគណនាស្វ័យប្រវត្តិ - {inm}",
              value=f"មធ្យម៖ {format_riel(re_avg)}",
              disabled=True,
              key=f"disp_avg_{mat_key}_{inm}"
          )

    # Action buttons
    st.divider()
    b_c1, b_c2, b_c3 = st.columns([2, 1.5, 1.5])
    with b_c1:
      btn_save_sup = st.button("💾 រក្សាទុកព័ត៌មាន & តម្លៃទំនិញអ្នកផ្គត់ផ្គង់", type="primary", use_container_width=True, key=f"btn_save_{sup_id_for_key}")
    with b_c2:
      btn_reset_defaults = st.button("🔄 យកតម្លៃគោលលំនាំដើម (Reset)", use_container_width=True, key=f"btn_reset_{sup_id_for_key}")
    with b_c3:
      btn_del_sup = st.button(f"🗑️ លុបអ្នកផ្គត់ផ្គង់នេះ", use_container_width=True, key=f"btn_del_{sup_id_for_key}") if not is_new and cur_sup else False

    if btn_reset_defaults:
      for it in SUPPLIER_PRODUCT_CATALOG:
        inm = it["name"]
        st.session_state[mat_key][inm] = {
            "p1": float(it["default_p1"]),
            "p2": float(it["default_p2"])
        }
      st.success("✅ បានកំណត់តម្លៃទំនិញទាំងអស់មកតាមតម្លៃគោលលំនាំដើមវិញ!")
      st.rerun()

    if btn_del_sup and cur_sup:
      delete_supplier(cur_sup["id"])
      cursor.execute("DELETE FROM products WHERE supplier_name=? AND price_level='supplier'", (cur_sup["supplier_name"],))
      conn.commit()
      st.success(f"🗑️ បានលុបអ្នកផ្គត់ផ្គង់ «{cur_sup['supplier_name']}» និងទិន្នន័យតម្លៃពាក់ព័ន្ធដោយជោគជ័យ!")
      st.rerun()

    if btn_save_sup:
      if not val_name:
        st.error("⚠️ សូមវាយបញ្ចូលឈ្មោះអ្នកផ្គត់ផ្គង់!")
      elif not is_commune_level and not tgt_school:
        st.error("⚠️ សូមជ្រើសរើស ឬវាយបញ្ចូលឈ្មោះសាលារៀនគោលដៅ!")
      else:
        final_sig_data = uploaded_sig_data if uploaded_sig_data is not None else (existing_sig_val or "")
        sup_lvl_str = "commune" if is_commune_level else "school"
        saved_id = save_or_update_supplier(
            school_name=tgt_school if not is_commune_level else "",
            supplier_name=val_name,
            village=val_vill,
            commune=val_comm,
            district=val_dist,
            province=val_prov,
            phone=val_phone,
            signature_data=final_sig_data,
            supply_level=sup_lvl_str,
            target_commune=tgt_commune,
            target_district=tgt_district,
            target_province=tgt_province,
            supplier_id=cur_sup["id"] if cur_sup else None
        )
        saved_prods_count = save_all_supplier_prices(
            supplier_name=val_name,
            price_dict=st.session_state[mat_key],
            supply_level=sup_lvl_str,
            target_school=tgt_school,
            target_commune=tgt_commune,
            target_district=tgt_district,
            target_province=tgt_province
        )
        st.success(f"🎉 បានរក្សាទុកអ្នកផ្គត់ផ្គង់ «{val_name}» និងតម្លៃទំនិញចំនួន {saved_prods_count} មុខដោយជោគជ័យ!")
        st.rerun()

  # ----------------- TAB 3: នាំចូលអ្នកផ្គត់ផ្គង់ពីក្រៅ -----------------
  with sup_tab3:
    st.subheader("📥 នាំចូលបញ្ជីអ្នកផ្គត់ផ្គង់ពីក្រៅ (Excel, CSV, Word, PDF, រូបភាព)")
    st.info("💡 អាចនាំចូលទិន្នន័យអ្នកផ្គត់ផ្គង់ និងអាសយដ្ឋានដោយស្វ័យប្រវត្តិពីឯកសារខាងក្រៅ")
    sup_tpl = generate_sample_excel(
        ["សាលារៀន", "ឈ្មោះអ្នកផ្គត់ផ្គង់", "ភូមិ", "ឃុំ/សង្កាត់", "ក្រុង/ស្រុក", "ខេត្ត", "លេខទូរស័ព្ទ"],
        ["សាលាបឋមសិក្សា ច្រឡង", "សាត ក្រូត", "ភូមិខ្មែរ", "រោង", "ស្រីស្នំ", "សៀមរាប", "090 854 133"]
    )
    st.download_button(
        "📥 ទាញយកគំរូឯកសារ Excel (Suppliers Template)",
        data=sup_tpl,
        file_name="Template_Suppliers.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    up_sup_file = st.file_uploader(
        "ជ្រើសរើសឯកសារអ្នកផ្គត់ផ្គង់ (xlsx, docx, csv, pdf, png, jpg)",
        type=["xlsx", "xls", "docx", "doc", "csv", "pdf", "png", "jpg", "jpeg"],
        key="up_sup_file_tab3"
    )
    if up_sup_file:
      raw_sup_df = extract_table_from_file(up_sup_file)
      if not raw_sup_df.empty:
        st.success(f"✅ អានទិន្នន័យបានជោគជ័យ! ចំនួន {len(raw_sup_df)} ជួរ")
        cols = list(raw_sup_df.columns)
        map_sch = next((c for c in cols if any(k in str(c).lower() for k in ["សាលា", "school"])), cols[0] if len(cols) > 0 else "")
        map_sup = next((c for c in cols if any(k in str(c).lower() for k in ["ផ្គត់ផ្គង់", "supplier", "ឈ្មោះ", "name"])), cols[1] if len(cols) > 1 else "")
        map_vil = next((c for c in cols if any(k in str(c).lower() for k in ["ភូមិ", "village"])), "")
        map_com = next((c for c in cols if any(k in str(c).lower() for k in ["ឃុំ", "commune"])), "")
        map_dis = next((c for c in cols if any(k in str(c).lower() for k in ["ស្រុក", "district"])), "")
        map_prv = next((c for c in cols if any(k in str(c).lower() for k in ["ខេត្ត", "province"])), "")
        map_phn = next((c for c in cols if any(k in str(c).lower() for k in ["ទូរស័ព្ទ", "phone", "tel"])), "")

        formatted_sup_df = pd.DataFrame({
            "សាលារៀន": raw_sup_df[map_sch] if map_sch in raw_sup_df else "",
            "ឈ្មោះអ្នកផ្គត់ផ្គង់": raw_sup_df[map_sup] if map_sup in raw_sup_df else "",
            "ភូមិ": raw_sup_df[map_vil] if map_vil in raw_sup_df else "ភូមិខ្មែរ",
            "ឃុំ/សង្កាត់": raw_sup_df[map_com] if map_com in raw_sup_df else "រោង",
            "ក្រុង/ស្រុក": raw_sup_df[map_dis] if map_dis in raw_sup_df else "ស្រីស្នំ",
            "ខេត្ត": raw_sup_df[map_prv] if map_prv in raw_sup_df else "សៀមរាប",
            "លេខទូរស័ព្ទ": raw_sup_df[map_phn] if map_phn in raw_sup_df else ""
        })

        st.markdown("##### ✏️ ផ្ទៀងផ្ទាត់ និងកែសម្រួលមុនពេលរក្សាទុក:")
        edited_sup_df = st.data_editor(formatted_sup_df, num_rows="dynamic", use_container_width=True, key="ed_sup_import_tab3")
        if st.button("📥 យល់ព្រមនាំចូលអ្នកផ្គត់ផ្គង់ទាំងអស់ចូល Database", key="btn_save_sup_import_tab3", use_container_width=True):
          count_i = 0
          for _, r in edited_sup_df.iterrows():
            i_sch = str(r.get("សាលារៀន", "")).strip()
            i_sup = str(r.get("ឈ្មោះអ្នកផ្គត់ផ្គង់", "")).strip()
            i_vil = str(r.get("ភូមិ", "")).strip()
            i_com = str(r.get("ឃុំ/សង្កាត់", "")).strip()
            i_dis = str(r.get("ក្រុង/ស្រុក", "")).strip()
            i_prv = str(r.get("ខេត្ត", "")).strip()
            i_phn = str(r.get("លេខទូរស័ព្ទ", "")).strip()
            if i_sch and i_sup:
              save_or_update_supplier(i_sch, i_sup, i_vil, i_com, i_dis, i_prv, i_phn)
              count_i += 1
          st.success(f"🎉 បាននាំចូលអ្នកផ្គត់ផ្គង់ចំនួន {count_i} ដោយជោគជ័យ!")
          st.rerun()


# ================= ៤. បញ្ជីមុខទំនិញ និងកំណត់តម្លៃ (តាមសាលា ឬ តាមឃុំ) =================
elif menu in ["📦 បញ្ជីមុខទំនិញ និងតម្លៃ (តាមសាលា / តាមឃុំ)", "📦 បញ្ជីមុខទំនិញតាមឃុំ (វគ្គ១/វគ្គ២/មធ្យម)"]:
  st.title("📦 បញ្ជីមុខទំនិញ និងកំណត់តម្លៃ (តាមសាលា ឬ តាមឃុំ)")
  st.info(
      "💡 កំណត់មុខទំនិញ និងតម្លៃទៅតាម **សាលារៀនរៀងៗខ្លួន** ឬ **តាមឃុំ/សង្កាត់រួម** ដោយមានការចាត់ថ្នាក់ប្រភេទស្វ័យប្រវត្ត (អង្ករ, ប្រេងឆា, អំបិល, ត្រី សាច់ ស៊ុត, បន្លែ) និងគណនាតម្លៃមធ្យមស្វ័យប្រវត្តិ។"
  )

  # ជម្រើសកម្រិតនៃការកំណត់តម្លៃ
  level_choice = st.radio(
      "🎯 ជម្រើសកម្រិតនៃការកំណត់តម្លៃទំនិញ៖",
      [
          "🏛️ កំណត់តាមឃុំ/សង្កាត់ (Commune Level - អនុវត្តចំពោះគ្រប់សាលាក្នុងឃុំ)", 
          "🏫 កំណត់តាមសាលារៀន (School Level - តម្លៃជាក់លាក់តាមសាលានីមួយៗ)",
          "🚚 កំណត់តាមអ្នកផ្គត់ផ្គង់ (Supplier Level - តម្លៃតាមកិច្ចសន្យាអ្នកផ្គត់ផ្គង់)",
      ],
      index=0,
      horizontal=True,
      key="prod_level_choice"
  )
  is_school_mode = "តាមសាលារៀន" in level_choice
  is_supplier_mode = "តាមអ្នកផ្គត់ផ្គង់" in level_choice

  all_provinces = get_provinces()

  if is_supplier_mode:
    all_sups = get_all_suppliers()
    sup_names = sorted(list(set([s["supplier_name"] for s in all_sups if s.get("supplier_name")])))
    col_sup1, col_sup2 = st.columns([1.5, 1.5])
    with col_sup1:
      sup_choice = st.selectbox(
          "🚚 ជ្រើសរើសអ្នកផ្គត់ផ្គង់",
          ["-- ជ្រើសរើសអ្នកផ្គត់ផ្គង់ --"] + sup_names + ["➕ វាយឈ្មោះអ្នកផ្គត់ផ្គង់ថ្មី..."],
          key="prod_sup_choice"
      )
      if sup_choice == "➕ វាយឈ្មោះអ្នកផ្គត់ផ្គង់ថ្មី...":
        sup_val = st.text_input("វាយឈ្មោះអ្នកផ្គត់ផ្គង់ថ្មី", key="prod_sup_in").strip()
      elif sup_choice != "-- ជ្រើសរើសអ្នកផ្គត់ផ្គង់ --":
        sup_val = sup_choice
      else:
        sup_val = ""
    with col_sup2:
      sup_sch_choice = st.selectbox(
          "🏫 អនុវត្តចំពោះសាលា (ជម្រើសបន្ថែម)",
          ["-- គ្រប់សាលារបស់អ្នកផ្គត់ផ្គង់នេះ --"] + get_all_schools(),
          key="prod_sup_sch"
      )
      s_val = sup_sch_choice if sup_sch_choice != "-- គ្រប់សាលារបស់អ្នកផ្គត់ផ្គង់នេះ --" else ""
    c_val = ""
    p_val = None
    d_val = None
    if sup_val:
      st.success(f"📍 **គោលដៅកំណត់តម្លៃ៖** អ្នកផ្គត់ផ្គង់ **«{sup_val}»** {f'(សាលា: {s_val})' if s_val else '(គ្រប់សាលាដែលផ្គត់ផ្គង់)'}")
    else:
      st.warning("⚠️ សូមជ្រើសរើស **អ្នកផ្គត់ផ្គង់** ខាងលើដើម្បីចាប់ផ្ដើមកំណត់តម្លៃ!")

  elif is_school_mode:
    col_mp, col_md, col_mc, col_ms = st.columns([1, 1, 1, 1.3])
    with col_mp:
      p_choice = st.selectbox("ខេត្ត/ក្រុង", ["-- ទាំងអស់ --"] + all_provinces, key="prod_p_sch")
      p_val = p_choice if p_choice != "-- ទាំងអស់ --" else None

    all_districts = get_districts(p_val) if p_val else get_districts()
    with col_md:
      d_choice = st.selectbox("ស្រុក/ខណ្ឌ", ["-- ទាំងអស់ --"] + all_districts, key="prod_d_sch")
      d_val = d_choice if d_choice != "-- ទាំងអស់ --" else None

    all_communes = get_communes(p_val, d_val) if (p_val and d_val) else get_communes()
    with col_mc:
      c_choice = st.selectbox("ឃុំ/សង្កាត់", ["-- ទាំងអស់ --"] + all_communes + ["➕ វាយបញ្ចូលឃុំថ្មី..."], key="prod_c_sch")
      if c_choice == "➕ វាយបញ្ចូលឃុំថ្មី...":
        c_val = st.text_input("វាយបញ្ចូលឃុំថ្មី", key="prod_c_in_sch").strip()
      elif c_choice != "-- ទាំងអស់ --":
        c_val = c_choice
      else:
        c_val = ""

    schools_in_scope = get_filtered_schools(p_val, d_val, c_val)
    with col_ms:
      s_choice = st.selectbox(
          "🏫 សាលារៀន",
          ["-- ជ្រើសរើសសាលា --"] + schools_in_scope + ["➕ វាយបញ្ចូលសាលាថ្មី..."],
          key="prod_s_sch"
      )
      if s_choice == "➕ វាយបញ្ចូលសាលាថ្មី...":
        s_val = st.text_input("វាយឈ្មោះសាលាថ្មី", key="prod_s_in_sch").strip()
      elif s_choice != "-- ជ្រើសរើសសាលា --":
        s_val = s_choice
      else:
        s_val = ""

    if s_val:
      st.success(f"📍 **គោលដៅកំណត់តម្លៃ៖** សាលារៀន **«{s_val}»** {f'(ឃុំ: {c_val})' if c_val else ''} {f'(ស្រុក: {d_val})' if d_val else ''} {f'(ខេត្ត: {p_val})' if p_val else ''}")
    else:
      st.warning("⚠️ សូមជ្រើសរើស **សាលារៀន** ខាងលើដើម្បីចាប់ផ្ដើមកំណត់តម្លៃ ឬពិនិត្យមុខទំនិញ!")

  else:
    col_mp, col_md, col_mc = st.columns(3)
    with col_mp:
      p_choice = st.selectbox("ខេត្ត/ក្រុង", ["-- ទាំងអស់ --"] + all_provinces, key="prod_p_comm")
      p_val = p_choice if p_choice != "-- ទាំងអស់ --" else None

    all_districts = get_districts(p_val) if p_val else get_districts()
    with col_md:
      d_choice = st.selectbox("ស្រុក/ខណ្ឌ", ["-- ទាំងអស់ --"] + all_districts, key="prod_d_comm")
      d_val = d_choice if d_choice != "-- ទាំងអស់ --" else None

    all_communes = get_communes(p_val, d_val) if (p_val and d_val) else get_communes()
    with col_mc:
      c_choice = st.selectbox(
          "ឃុំ/សង្កាត់",
          ["-- ជ្រើសរើសឃុំ --"] + all_communes + ["➕ វាយបញ្ចូលឃុំថ្មី..."],
          key="prod_c_comm",
      )
      if c_choice == "➕ វាយបញ្ចូលឃុំថ្មី...":
        c_val = st.text_input("វាយបញ្ចូលឃុំថ្មី", key="prod_c_in_comm").strip()
      elif c_choice != "-- ជ្រើសរើសឃុំ --":
        c_val = c_choice
      else:
        c_val = ""
    s_val = ""

    if c_val:
      st.success(f"📍 **គោលដៅកំណត់តម្លៃ៖** ឃុំ/សង្កាត់ **«{c_val}»** {f'(ស្រុក: {d_val})' if d_val else ''} {f'(ខេត្ត: {p_val})' if p_val else ''} — គ្រប់សាលាក្នុងឃុំនេះនឹងប្រើប្រាស់តម្លៃនេះជាគោល")
    else:
      st.warning("⚠️ សូមជ្រើសរើស **ឃុំ/សង្កាត់** ខាងលើដើម្បីចាប់ផ្ដើមកំណត់តម្លៃ ឬពិនិត្យមុខទំនិញ!")

  if is_supplier_mode:
    target_label = f"អ្នកផ្គត់ផ្គង់ «{sup_val}»" + (f" (សាលា «{s_val}»)" if s_val else "")
  elif is_school_mode:
    target_label = f"សាលារៀន «{s_val}»" if s_val else ""
  else:
    target_label = f"ឃុំ «{c_val}»" if c_val else ""


  prod_tab1, prod_tab2, prod_tab3, prod_tab4 = st.tabs([
      "➕ បន្ថែមមុខទំនិញថ្មី",
      "✏️ កែប្រែ / 🗑️ លុបមុខទំនិញ",
      "📅 កំណត់កាលបរិច្ឆេទវគ្គរួម",
      "📥 នាំចូលមុខទំនិញពីក្រៅ (Excel, Word, CSV, PDF, រូបភាព)",
  ])

  def_p1_s, def_p1_e, def_p2_s, def_p2_e = get_commune_phase_dates(c_val, s_val if is_school_mode else None)

  # ----------------- TAB 1: បន្ថែមមុខទំនិញថ្មី -----------------
  with prod_tab1:
    st.subheader(
        f"➕ បន្ថែមមុខទំនិញសម្រាប់: {target_label if target_label else '(សូមជ្រើសរើសគោលដៅខាងលើ)'}"
    )

    col_pname, col_pcat = st.columns([2, 1.2])
    with col_pname:
      std_options = ["-- ជ្រើសរើសមុខទំនិញស្ដង់ដារ (៦១ មុខដូចក្នុងរូប) --"] + STANDARD_PRODUCT_ITEMS + ["➕ វាយបញ្ចូលឈ្មោះទំនិញផ្សេងទៀត..."]
      sel_std_item = st.selectbox(
          "📦 ជ្រើសរើសមុខទំនិញ (គំរូស្ដង់ដារដូចក្នុងរូប)",
          std_options,
          key="sel_std_item"
      )
      if sel_std_item == "➕ វាយបញ្ចូលឈ្មោះទំនិញផ្សេងទៀត...":
        p_name = st.text_input("វាយឈ្មោះមុខទំនិញថ្មី", key="add_p_name_custom").strip()
      elif sel_std_item != "-- ជ្រើសរើសមុខទំនិញស្ដង់ដារ (៦១ មុខដូចក្នុងរូប) --":
        p_name = sel_std_item
      else:
        p_name = ""

    with col_pcat:
      auto_cat = classify_item_category(p_name) if p_name else STANDARD_CATEGORIES[0]
      cat_opts = STANDARD_CATEGORIES + ["ផ្សេងៗ"]
      cat_idx = cat_opts.index(auto_cat) if auto_cat in cat_opts else 0
      p_cat = st.selectbox(
          "🏷️ ប្រភេទសម្គាល់ (ចាត់ថ្នាក់ស្វ័យប្រវត្តិ)",
          cat_opts,
          index=cat_idx,
          key=f"add_pcat_{p_name[:8] if p_name else 'def'}"
      )
      if p_name:
        st.caption(f"💡 ប្រព័ន្ធចាត់ថ្នាក់ជា៖ **{format_category_badge(p_cat)}** ដោយស្វ័យប្រវត្តិ")

    # យកតម្លៃលំនាំដើមពីកាតាឡុកបើមាន
    cat_match = next((item for item in STANDARD_PRODUCT_CATALOG if item["name"] == p_name), None)
    def_price_p1 = float(cat_match["price_phase1"]) if cat_match else 0.0
    def_price_p2 = float(cat_match["price_phase2"]) if cat_match else def_price_p1

    col_ph1, col_ph2 = st.columns(2)
    with col_ph1:
      st.markdown("##### 🟢 កំណត់តម្លៃ និងកាលបរិច្ឆេទ វគ្គ១")
      p_price1 = st.number_input(
          "តម្លៃវគ្គ ១ (៛)",
          min_value=0.0,
          value=def_price_p1,
          step=100.0,
          format="%.0f",
          key=f"add_p1_{p_name[:8] if p_name else 'def'}",
      )
      c1_s, c1_e = st.columns(2)
      with c1_s:
        p1_start = st.date_input(
            "ថ្ងៃចាប់ផ្ដើម វគ្គ១",
            value=parse_date_safe(def_p1_s),
            key="add_p1_s",
        )
      with c1_e:
        p1_end = st.date_input(
            "ថ្ងៃបញ្ចប់ វគ្គ១",
            value=parse_date_safe(def_p1_e),
            key="add_p1_e",
        )

    with col_ph2:
      st.markdown("##### 🔵 កំណត់តម្លៃ និងកាលបរិច្ឆេទ វគ្គ២")
      p_price2 = st.number_input(
          "តម្លៃវគ្គ ២ (៛)",
          min_value=0.0,
          value=def_price_p2,
          step=100.0,
          format="%.0f",
          key=f"add_p2_{p_name[:8] if p_name else 'def'}",
      )
      c2_s, c2_e = st.columns(2)
      with c2_s:
        p2_start = st.date_input(
            "ថ្ងៃចាប់ផ្ដើម វគ្គ២",
            value=parse_date_safe(def_p2_s),
            key="add_p2_s",
        )
      with c2_e:
        p2_end = st.date_input(
            "ថ្ងៃបញ្ចប់ វគ្គ២",
            value=parse_date_safe(def_p2_e),
            key="add_p2_e",
        )

    auto_avg = (p_price1 + p_price2) / 2.0
    st.metric("📊 តម្លៃមធ្យម (គណនាស្វ័យប្រវត្តិ)", format_riel(auto_avg))

    if st.button("💾 រក្សាទុកមុខទំនិញ", use_container_width=True, type="primary"):
      if is_supplier_mode and not sup_val:
        st.error("សូមជ្រើសរើស ឬវាយបញ្ចូលឈ្មោះអ្នកផ្គត់ផ្គង់សិន!")
      elif is_school_mode and not s_val:
        st.error("សូមជ្រើសរើស ឬវាយបញ្ចូលឈ្មោះសាលារៀនសិន!")
      elif not is_school_mode and not is_supplier_mode and not c_val:
        st.error("សូមជ្រើសរើស ឬវាយបញ្ចូលឃុំសិន!")
      elif not p_name.strip():
        st.error("សូមជ្រើសរើស ឬវាយបញ្ចូលឈ្មោះមុខទំនិញ!")
      elif p1_start > p1_end:
        st.error("កាលបរិច្ឆេទចាប់ផ្ដើមវគ្គ១ មិនអាចធំជាងថ្ងៃបញ្ចប់បានទេ!")
      elif p2_start > p2_end:
        st.error("កាលបរិច្ឆេទចាប់ផ្ដើមវគ្គ២ មិនអាចធំជាងថ្ងៃបញ្ចប់បានទេ!")
      else:
        target_prov = p_val or ""
        target_dist = d_val or ""
        target_comm = c_val or ""
        target_sch = s_val or ""
        target_sup = sup_val if is_supplier_mode else ""
        target_lvl = "supplier" if is_supplier_mode else ("school" if is_school_mode else "commune")

        # ពិនិត្យថាមានទំនិញនេះស្រាប់ក្នុង Scope នេះឬនៅ
        chk_sql = "SELECT id FROM products WHERE item_name=? AND price_level=?"
        chk_params = [p_name.strip(), target_lvl]
        if is_supplier_mode:
          chk_sql += " AND supplier_name=? AND school_name=?"
          chk_params.extend([target_sup, target_sch])
        elif is_school_mode:
          chk_sql += " AND school_name=?"
          chk_params.append(target_sch)
        else:
          chk_sql += " AND commune=?"
          chk_params.append(target_comm)
        row_chk = cursor.execute(chk_sql, tuple(chk_params)).fetchone()

        if row_chk:
          cursor.execute(
              """
              UPDATE products 
              SET price_phase1=?, price_phase2=?, price_avg=?,
                  phase1_start=?, phase1_end=?, phase2_start=?, phase2_end=?,
                  category=?, province=?, district=?, commune=?, school_name=?, price_level=?, supplier_name=?
              WHERE id=?
              """,
              (p_price1, p_price2, auto_avg, str(p1_start), str(p1_end), str(p2_start), str(p2_end),
               p_cat, target_prov, target_dist, target_comm, target_sch, target_lvl, target_sup, row_chk[0])
          )
          st.success(f"🎉 បានធ្វើបច្ចុប្បន្នភាពមុខទំនិញ '{p_name}' [{format_category_badge(p_cat)}] សម្រាប់ {target_label} (តម្លៃមធ្យម: {format_riel(auto_avg)}) រួចរាល់!")
        else:
          cursor.execute(
              """
              INSERT INTO products (item_name, commune, price_phase1, price_phase2, price_avg,
                                    phase1_start, phase1_end, phase2_start, phase2_end, category,
                                    province, district, school_name, price_level, supplier_name)
              VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
              """,
              (p_name.strip(), target_comm, p_price1, p_price2, auto_avg,
               str(p1_start), str(p1_end), str(p2_start), str(p2_end), p_cat,
               target_prov, target_dist, target_sch, target_lvl, target_sup)
          )
          st.success(f"🎉 បានរក្សាទុកមុខទំនិញ '{p_name}' [{format_category_badge(p_cat)}] សម្រាប់ {target_label} (តម្លៃមធ្យម: {format_riel(auto_avg)}) ដោយជោគជ័យ!")
        conn.commit()
        st.rerun()

    # ប៊ូតុងមួយឃ្លីកបញ្ចូល ៦១ មុខ
    st.divider()
    with st.expander("⚡ បញ្ចូលបញ្ជីមុខទំនិញស្ដង់ដារទាំង ៦១ មុខ (ដូចក្នុងរូបភាព) ដោយចុចតែ ១ ឃ្លីក", expanded=False):
      active_dest = f"សាលារៀន «{s_val}»" if (is_school_mode and s_val) else (f"ឃុំ «{c_val}»" if c_val else "")
      st.markdown(f"""
      មុខងារនេះនឹងជួយបញ្ចូល **មុខទំនិញស្ដង់ដារទាំង ៦១ មុខ** ដូចក្នុងរូបភាពគំរូ ព្រមទាំងតម្លៃវគ្គ១-វគ្គ២ ស្ដង់ដារ និងចាត់ថ្នាក់ប្រភេទស្វ័យប្រវត្ត (**អង្ករ, ប្រេងឆា, អំបិល, ត្រី សាច់ ស៊ុត, បន្លែ**) ចូលទៅកាន់ **{active_dest if active_dest else '(សូមជ្រើសរើសគោលដៅ)'}** ដោយស្វ័យប្រវត្តិ មិនបាច់វាយម្តងមួយៗឡើយ!
      """)
      if st.button(f"🚀 បញ្ចូលមុខទំនិញទាំង ៦១ មុខភ្លាមៗ ទៅកាន់ {active_dest if active_dest else 'គោលដៅ'}", key="btn_batch_seed_61", type="secondary"):
        if is_school_mode and not s_val:
          st.error("សូមជ្រើសរើសសាលារៀនជាមុនសិន!")
        elif not is_school_mode and not c_val:
          st.error("សូមជ្រើសរើសឃុំជាមុនសិន!")
        else:
          target_prov = p_val or ""
          target_dist = d_val or ""
          target_comm = c_val or ""
          target_sch = s_val if is_school_mode else ""
          target_lvl = "school" if is_school_mode else "commune"

          add_c = 0
          upd_c = 0
          for item in STANDARD_PRODUCT_CATALOG:
            it_name = item["name"]
            it_cat = item["category"]
            it_p1 = item["price_phase1"]
            it_p2 = item["price_phase2"]
            it_avg = (it_p1 + it_p2) / 2.0

            chk_sql = "SELECT id FROM products WHERE item_name=? AND price_level=?"
            chk_p = [it_name, target_lvl]
            if is_school_mode:
              chk_sql += " AND school_name=?"
              chk_p.append(target_sch)
            else:
              chk_sql += " AND commune=?"
              chk_p.append(target_comm)
            r_ex = cursor.execute(chk_sql, tuple(chk_p)).fetchone()

            if r_ex:
              cursor.execute(
                  """
                  UPDATE products 
                  SET price_phase1=?, price_phase2=?, price_avg=?,
                      phase1_start=?, phase1_end=?, phase2_start=?, phase2_end=?,
                      category=?, province=?, district=?, commune=?, school_name=?, price_level=?
                  WHERE id=?
                  """,
                  (it_p1, it_p2, it_avg, def_p1_s, def_p1_e, def_p2_s, def_p2_e,
                   it_cat, target_prov, target_dist, target_comm, target_sch, target_lvl, r_ex[0])
              )
              upd_c += 1
            else:
              cursor.execute(
                  """
                  INSERT INTO products (item_name, commune, price_phase1, price_phase2, price_avg,
                                        phase1_start, phase1_end, phase2_start, phase2_end, category,
                                        province, district, school_name, price_level)
                  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                  """,
                  (it_name, target_comm, it_p1, it_p2, it_avg,
                   def_p1_s, def_p1_e, def_p2_s, def_p2_e, it_cat,
                   target_prov, target_dist, target_sch, target_lvl)
              )
              add_c += 1
          conn.commit()
          st.success(f"🎉 ជោគជ័យ! បានបញ្ចូលមុខទំនិញថ្មី {add_c} មុខ និងធ្វើបច្ចុប្បន្នភាព {upd_c} មុខ សម្រាប់ {active_dest}!")
          st.rerun()

  # ----------------- TAB 2: កែប្រែ / លុបមុខទំនិញ -----------------
  with prod_tab2:
    st.subheader("✏️ កែប្រែ ឬ 🗑️ លុបមុខទំនិញ")
    prod_query = """
        SELECT id, item_name, commune, price_phase1, price_phase2, price_avg, 
               phase1_start, phase1_end, phase2_start, phase2_end, category,
               province, district, school_name, price_level
        FROM products
    """
    params = []
    if is_school_mode and s_val:
      prod_query += " WHERE school_name=? AND price_level='school'"
      params.append(s_val)
    elif not is_school_mode and c_val:
      prod_query += " WHERE commune=? AND (school_name IS NULL OR school_name = '' OR price_level='commune')"
      params.append(c_val)
    prod_query += " ORDER BY id DESC"

    all_prods = cursor.execute(prod_query, tuple(params)).fetchall()
    if all_prods:
      prod_dict = {}
      for r in all_prods:
        lvl_str = f"🏫 {r[13]}" if (r[14] == "school" and r[13]) else f"🏛️ ឃុំ {r[2]}"
        lbl = f"ID {r[0]}: {r[1]} [{format_category_badge(r[10])}] ({lvl_str} | វគ្គ១: {format_riel(r[3])} | វគ្គ២: {format_riel(r[4])})"
        prod_dict[lbl] = r

      p_sel = st.selectbox(
          "ជ្រើសរើសមុខទំនិញដើម្បីកែប្រែ ឬលុប",
          list(prod_dict.keys()),
          key="sel_prod_edit",
      )
      cur_p = prod_dict[p_sel]

      col_en, col_ecat, col_elvl = st.columns([2, 1.2, 1])
      with col_en:
        ed_pname = st.text_input("ឈ្មោះមុខទំនិញ", value=cur_p[1], key="ed_pname")
      with col_ecat:
        cur_cat = cur_p[10] if len(cur_p) > 10 and cur_p[10] else classify_item_category(cur_p[1])
        cat_opts = STANDARD_CATEGORIES + ["ផ្សេងៗ"]
        cat_idx = cat_opts.index(cur_cat) if cur_cat in cat_opts else 0
        ed_pcat = st.selectbox("🏷️ ប្រភេទសម្គាល់", cat_opts, index=cat_idx, key=f"ed_pcat_{cur_p[0]}")
      with col_elvl:
        cur_lvl_is_sch = (cur_p[14] == "school")
        ed_lvl = st.selectbox("កម្រិតកំណត់តម្លៃ", ["🏛️ តាមឃុំ", "🏫 តាមសាលា"], index=1 if cur_lvl_is_sch else 0, key=f"ed_lvl_{cur_p[0]}")

      col_ecomm, col_esch = st.columns(2)
      with col_ecomm:
        ed_pcomm = st.text_input("ឃុំ/សង្កាត់", value=cur_p[2] or "", key=f"ed_pcomm_{cur_p[0]}")
      with col_esch:
        ed_psch = st.text_input("សាលារៀន (បើកំណត់តាមសាលា)", value=cur_p[13] or "", key=f"ed_psch_{cur_p[0]}")

      col_eph1, col_eph2 = st.columns(2)
      with col_eph1:
        st.markdown("##### 🟢 កែប្រែតម្លៃ និងកាលបរិច្ឆេទ វគ្គ១")
        ed_pr1 = st.number_input(
            "តម្លៃវគ្គ ១ (៛)",
            min_value=0.0,
            value=float(cur_p[3] or 0),
            step=100.0,
            format="%.0f",
            key=f"ed_pr1_{cur_p[0]}",
        )
        c_ed1_s, c_ed1_e = st.columns(2)
        with c_ed1_s:
          ed_p1_s = st.date_input(
              "ថ្ងៃចាប់ផ្ដើម វគ្គ១",
              value=parse_date_safe(cur_p[6], date(date.today().year, 1, 1)),
              key=f"ed_p1_s_{cur_p[0]}",
          )
        with c_ed1_e:
          ed_p1_e = st.date_input(
              "ថ្ងៃបញ្ចប់ វគ្គ១",
              value=parse_date_safe(cur_p[7], date(date.today().year, 6, 30)),
              key=f"ed_p1_e_{cur_p[0]}",
          )

      with col_eph2:
        st.markdown("##### 🔵 កែប្រែតម្លៃ និងកាលបរិច្ឆេទ វគ្គ២")
        ed_pr2 = st.number_input(
            "តម្លៃវគ្គ ២ (៛)",
            min_value=0.0,
            value=float(cur_p[4] or 0),
            step=100.0,
            format="%.0f",
            key=f"ed_pr2_{cur_p[0]}",
        )
        c_ed2_s, c_ed2_e = st.columns(2)
        with c_ed2_s:
          ed_p2_s = st.date_input(
              "ថ្ងៃចាប់ផ្ដើម វគ្គ២",
              value=parse_date_safe(cur_p[8], date(date.today().year, 7, 1)),
              key=f"ed_p2_s_{cur_p[0]}",
          )
        with c_ed2_e:
          ed_p2_e = st.date_input(
              "ថ្ងៃបញ្ចប់ វគ្គ២",
              value=parse_date_safe(cur_p[9], date(date.today().year, 12, 31)),
              key=f"ed_p2_e_{cur_p[0]}",
          )

      ed_avg = (ed_pr1 + ed_pr2) / 2.0
      st.metric("📊 តម្លៃមធ្យមថ្មី (គណនាស្វ័យប្រវត្តិ)", format_riel(ed_avg))

      col_pb1, col_pb2 = st.columns(2)
      with col_pb1:
        if st.button("💾 រក្សាទុកការកែប្រែទំនិញ", use_container_width=True, type="primary", key=f"btn_save_ed_{cur_p[0]}"):
          new_lvl = "school" if "សាលា" in ed_lvl else "commune"
          if not ed_pname.strip():
            st.error("ឈ្មោះទំនិញមិនអាចទុកឱ្យទទេបានទេ!")
          elif new_lvl == "commune" and not ed_pcomm.strip():
            st.error("សូមបំពេញឈ្មោះឃុំ!")
          elif new_lvl == "school" and not ed_psch.strip():
            st.error("សូមបំពេញឈ្មោះសាលារៀន!")
          elif ed_p1_s > ed_p1_e:
            st.error("កាលបរិច្ឆេទចាប់ផ្ដើមវគ្គ១ មិនអាចធំជាងថ្ងៃបញ្ចប់បានទេ!")
          elif ed_p2_s > ed_p2_e:
            st.error("កាលបរិច្ឆេទចាប់ផ្ដើមវគ្គ២ មិនអាចធំជាងថ្ងៃបញ្ចប់បានទេ!")
          else:
            cursor.execute(
                """
                UPDATE products 
                SET item_name=?, commune=?, school_name=?, price_level=?,
                    price_phase1=?, price_phase2=?, price_avg=?,
                    phase1_start=?, phase1_end=?, phase2_start=?, phase2_end=?, category=? 
                WHERE id=?
                """,
                (
                    ed_pname.strip(),
                    ed_pcomm.strip(),
                    ed_psch.strip() if new_lvl == "school" else "",
                    new_lvl,
                    ed_pr1,
                    ed_pr2,
                    ed_avg,
                    str(ed_p1_s),
                    str(ed_p1_e),
                    str(ed_p2_s),
                    str(ed_p2_e),
                    ed_pcat,
                    cur_p[0],
                ),
            )
            conn.commit()
            st.success(f"🎉 បានកែប្រែមុខទំនិញ ID {cur_p[0]} ({ed_pname} [{format_category_badge(ed_pcat)}]) ដោយជោគជ័យ!")
            st.rerun()

      with col_pb2:
        if st.button(f"🗑️ លុបមុខទំនិញនេះ (ID: {cur_p[0]})", use_container_width=True, key=f"btn_del_p_{cur_p[0]}"):
          cursor.execute("DELETE FROM products WHERE id=?", (cur_p[0],))
          conn.commit()
          st.success(f"បានលុបមុខទំនិញ ID {cur_p[0]} រួចរាល់!")
          st.rerun()
    else:
      st.info(f"មិនទាន់មានមុខទំនិញសម្រាប់ {target_label if target_label else 'គោលដៅនេះ'} ទេ។")

  # ----------------- TAB 3: កំណត់កាលបរិច្ឆេទវគ្គរួម -----------------
  with prod_tab3:
    st.subheader(f"📅 កំណត់កាលបរិច្ឆេទវគ្គរួម {f'សម្រាប់ {target_label}' if target_label else ''}")
    st.info(
        "💡 មុខងារនេះជួយឱ្យអ្នកអាចកំណត់កាលបរិច្ឆេទចាប់ផ្ដើម និងបញ្ចប់ នៃវគ្គ១ និងវគ្គ២ ដល់មុខទំនិញទាំងអស់ក្នុងគោលដៅតែម្ដង យ៉ាងងាយស្រួល និងរហ័ស!"
    )
    if is_school_mode and not s_val:
      st.warning("⚠️ សូមជ្រើសរើសសាលារៀនខាងលើជាមុនសិន ដើម្បីអនុវត្តការកំណត់កាលបរិច្ឆេទវគ្គរួម!")
    elif not is_school_mode and not c_val:
      st.warning("⚠️ សូមជ្រើសរើសឃុំ/សង្កាត់ខាងលើជាមុនសិន ដើម្បីអនុវត្តការកំណត់កាលបរិច្ឆេទវគ្គរួម!")
    else:
      def_b_p1_s, def_b_p1_e, def_b_p2_s, def_b_p2_e = get_commune_phase_dates(c_val, s_val if is_school_mode else None)
      col_b1, col_b2 = st.columns(2)
      with col_b1:
        st.markdown("##### 🟢 កាលបរិច្ឆេទ វគ្គ១ (ឆមាសទី១)")
        b1_s = st.date_input(
            "ថ្ងៃចាប់ផ្ដើម វគ្គ១",
            value=parse_date_safe(def_b_p1_s, date(date.today().year, 1, 1)),
            key="batch_p1_s",
        )
        b1_e = st.date_input(
            "ថ្ងៃបញ្ចប់ វគ្គ១",
            value=parse_date_safe(def_b_p1_e, date(date.today().year, 6, 30)),
            key="batch_p1_e",
        )
      with col_b2:
        st.markdown("##### 🔵 កាលបរិច្ឆេទ វគ្គ២ (ឆមាសទី២)")
        b2_s = st.date_input(
            "ថ្ងៃចាប់ផ្ដើម វគ្គ២",
            value=parse_date_safe(def_b_p2_s, date(date.today().year, 7, 1)),
            key="batch_p2_s",
        )
        b2_e = st.date_input(
            "ថ្ងៃបញ្ចប់ វគ្គ២",
            value=parse_date_safe(def_b_p2_e, date(date.today().year, 12, 31)),
            key="batch_p2_e",
        )

      if is_school_mode:
        num_prods = cursor.execute("SELECT COUNT(*) FROM products WHERE school_name=? AND price_level='school'", (s_val,)).fetchone()[0]
      else:
        num_prods = cursor.execute("SELECT COUNT(*) FROM products WHERE commune=? AND (school_name IS NULL OR school_name='' OR price_level='commune')", (c_val,)).fetchone()[0]

      st.write(f"📦 ចំនួនមុខទំនិញបច្ចុប្បន្នក្នុង {target_label}: **{num_prods}** មុខ")

      if st.button(
          f"💾 អនុវត្តកាលបរិច្ឆេទទាំងនេះចំពោះគ្រប់មុខទំនិញក្នុង {target_label}",
          type="primary",
          use_container_width=True,
          key="btn_apply_batch_dates",
      ):
        if num_prods == 0:
          st.warning("មិនទាន់មានមុខទំនិញក្នុងគោលដៅនេះនៅឡើយទេ!")
        elif b1_s > b1_e:
          st.error("កាលបរិច្ឆេទចាប់ផ្ដើមវគ្គ១ មិនអាចធំជាងថ្ងៃបញ្ចប់បានទេ!")
        elif b2_s > b2_e:
          st.error("កាលបរិច្ឆេទចាប់ផ្ដើមវគ្គ២ មិនអាចធំជាងថ្ងៃបញ្ចប់បានទេ!")
        else:
          if is_school_mode:
            cursor.execute(
                """
                UPDATE products 
                SET phase1_start=?, phase1_end=?, phase2_start=?, phase2_end=? 
                WHERE school_name=? AND price_level='school'
                """,
                (str(b1_s), str(b1_e), str(b2_s), str(b2_e), s_val),
            )
          else:
            cursor.execute(
                """
                UPDATE products 
                SET phase1_start=?, phase1_end=?, phase2_start=?, phase2_end=? 
                WHERE commune=? AND (school_name IS NULL OR school_name='' OR price_level='commune')
                """,
                (str(b1_s), str(b1_e), str(b2_s), str(b2_e), c_val),
            )
          conn.commit()
          st.success(
              f"🎉 បានធ្វើបច្ចុប្បន្នភាពកាលបរិច្ឆេទ វគ្គ១ ({b1_s} ដល់ {b1_e}) និង វគ្គ២ ({b2_s} ដល់ {b2_e}) ដល់មុខទំនិញទាំង {num_prods} នៃ {target_label} ដោយជោគជ័យ!"
          )
          st.rerun()

  # ----------------- TAB 4: នាំចូលមុខទំនិញពីក្រៅ -----------------
  with prod_tab4:
    st.subheader("📥 នាំចូលមុខទំនិញពីក្រៅ (Excel, Word, CSV, PDF, រូបភាព)")
    st.info(
        "💡 គាំទ្រការនាំចូលមុខទំនិញសម្រាប់ **សាលារៀន** ឬ **ឃុំ/សង្កាត់** ដោយចាត់ថ្នាក់ប្រភេទស្វ័យប្រវត្ត និងគណនាតម្លៃមធ្យមស្វ័យប្រវត្តិ។"
    )

    cur_y = date.today().year
    prod_tpl = generate_sample_excel(
        [
            "ឈ្មោះមុខទំនិញ",
            "ប្រភេទសម្គាល់",
            "កម្រិតកំណត់តម្លៃ (សាលា/ឃុំ)",
            "សាលារៀន",
            "ឃុំ/សង្កាត់",
            "តម្លៃវគ្គ១ (៛)",
            "ថ្ងៃចាប់ផ្ដើមវគ្គ១",
            "ថ្ងៃបញ្ចប់វគ្គ១",
            "តម្លៃវគ្គ២ (៛)",
            "ថ្ងៃចាប់ផ្ដើមវគ្គ២",
            "ថ្ងៃបញ្ចប់វគ្គ២",
        ],
        [
            "អង្ករចម្រុះ",
            "អង្ករ",
            "តាមឃុំ",
            "",
            c_val if c_val else "ស្លែងស្ពាន",
            2100,
            f"{cur_y}-01-01",
            f"{cur_y}-06-30",
            2100,
            f"{cur_y}-07-01",
            f"{cur_y}-12-31",
        ],
    )
    st.download_button(
        "📥 ទាញយកគំរូឯកសារ Excel (Product Template)",
        data=prod_tpl,
        file_name="Template_Products.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    up_prod_file = st.file_uploader(
        "ជ្រើសរើសឯកសារមុខទំនិញ (xlsx, docx, csv, pdf, png, jpg)",
        type=["xlsx", "xls", "docx", "doc", "csv", "pdf", "png", "jpg", "jpeg"],
        key="up_prod_file",
    )

    if up_prod_file:
      raw_prod_df = extract_table_from_file(up_prod_file)
      if not raw_prod_df.empty:
        st.success(f"✅ អានទិន្នន័យបានជោគជ័យ! ចំនួន {len(raw_prod_df)} ជួរ")

        cols = list(raw_prod_df.columns)
        map_pn = next(
            (c for c in cols if any(k in str(c).lower() for k in ["ទំនិញ", "item", "product", "name"])),
            cols[0] if len(cols) > 0 else "",
        )
        map_pcat = next(
            (c for c in cols if any(k in str(c).lower() for k in ["ប្រភេទ", "សម្គាល់", "category", "cat"])),
            "",
        )
        map_plvl = next(
            (c for c in cols if any(k in str(c).lower() for k in ["កម្រិត", "level"])),
            "",
        )
        map_psch = next(
            (c for c in cols if any(k in str(c).lower() for k in ["សាលា", "school"])),
            "",
        )
        map_pc = next(
            (c for c in cols if any(k in str(c).lower() for k in ["ឃុំ", "comm"])),
            cols[1] if len(cols) > 1 else "",
        )
        map_pp1 = next(
            (c for c in cols if any(k in str(c).lower() for k in ["វគ្គ១", "phase1", "price1", "តម្លៃ១"])),
            cols[2] if len(cols) > 2 else "",
        )
        map_pp2 = next(
            (c for c in cols if any(k in str(c).lower() for k in ["វគ្គ២", "phase2", "price2", "តម្លៃ២"])),
            cols[3] if len(cols) > 3 else "",
        )
        map_p1s = next(
            (c for c in cols if any(k in str(c).lower() for k in ["ចាប់ផ្ដើម១", "p1_start", "start1", "ផ្ដើម១"])),
            "",
        )
        map_p1e = next(
            (c for c in cols if any(k in str(c).lower() for k in ["បញ្ចប់១", "p1_end", "end1"])),
            "",
        )
        map_p2s = next(
            (c for c in cols if any(k in str(c).lower() for k in ["ចាប់ផ្ដើម២", "p2_start", "start2", "ផ្ដើម២"])),
            "",
        )
        map_p2e = next(
            (c for c in cols if any(k in str(c).lower() for k in ["បញ្ចប់២", "p2_end", "end2"])),
            "",
        )

        p1_series = pd.to_numeric(raw_prod_df[map_pp1] if map_pp1 in raw_prod_df else 0, errors="coerce").fillna(0)
        p2_series = pd.to_numeric(raw_prod_df[map_pp2] if map_pp2 in raw_prod_df else 0, errors="coerce").fillna(0)
        avg_series = ((p1_series + p2_series) / 2.0).round(2)

        comm_default = raw_prod_df[map_pc] if map_pc in raw_prod_df else (c_val if c_val else "")
        sch_default = raw_prod_df[map_psch] if map_psch in raw_prod_df else (s_val if is_school_mode and s_val else "")

        cat_series = []
        for idx_r, row_r in raw_prod_df.iterrows():
          c_val_found = str(row_r[map_pcat]).strip() if (map_pcat and map_pcat in row_r) else ""
          if not c_val_found or c_val_found.lower() == 'nan':
            c_val_found = classify_item_category(str(row_r[map_pn])) if map_pn in row_r else "បន្លែ"
          cat_series.append(c_val_found)

        formatted_prod_df = pd.DataFrame({
            "ឈ្មោះមុខទំនិញ": raw_prod_df[map_pn] if map_pn in raw_prod_df else "",
            "ប្រភេទសម្គាល់": cat_series,
            "សាលារៀន": sch_default,
            "ឃុំ/សង្កាត់": comm_default,
            "តម្លៃវគ្គ១ (៛)": p1_series,
            "ថ្ងៃចាប់ផ្ដើមវគ្គ១": raw_prod_df[map_p1s] if map_p1s and map_p1s in raw_prod_df else def_p1_s,
            "ថ្ងៃបញ្ចប់វគ្គ១": raw_prod_df[map_p1e] if map_p1e and map_p1e in raw_prod_df else def_p1_e,
            "តម្លៃវគ្គ២ (៛)": p2_series,
            "ថ្ងៃចាប់ផ្ដើមវគ្គ២": raw_prod_df[map_p2s] if map_p2s and map_p2s in raw_prod_df else def_p2_s,
            "ថ្ងៃបញ្ចប់វគ្គ២": raw_prod_df[map_p2e] if map_p2e and map_p2e in raw_prod_df else def_p2_e,
            "តម្លៃមធ្យម (៛)": avg_series,
        })

        st.markdown("##### ✏️ ផ្ទៀងផ្ទាត់ កែប្រែ និងគណនាតម្លៃមធ្យមស្វ័យប្រវត្ត:")
        edited_prod_df = st.data_editor(
            formatted_prod_df, num_rows="dynamic", use_container_width=True, key="ed_prod_import"
        )

        edited_prod_df["តម្លៃមធ្យម (៛)"] = (
            (pd.to_numeric(edited_prod_df["តម្លៃវគ្គ១ (៛)"], errors="coerce").fillna(0) +
             pd.to_numeric(edited_prod_df["តម្លៃវគ្គ២ (៛)"], errors="coerce").fillna(0)) / 2.0
        ).round(2)

        if st.button("📥 យល់ព្រមនាំចូលមុខទំនិញទាំងអស់ចូល Database", key="btn_save_prod_import", use_container_width=True, type="primary"):
          count = 0
          for _, r in edited_prod_df.iterrows():
            in_name = str(r.get("ឈ្មោះមុខទំនិញ", "")).strip()
            in_cat = str(r.get("ប្រភេទសម្គាល់", "")).strip() or classify_item_category(in_name)
            in_sch = str(r.get("សាលារៀន", "")).strip()
            in_comm = str(r.get("ឃុំ/សង្កាត់", "")).strip() or c_val
            in_p1 = float(r.get("តម្លៃវគ្គ១ (៛)", 0) or 0)
            in_p2 = float(r.get("តម្លៃវគ្គ២ (៛)", 0) or 0)
            in_avg = float(r.get("តម្លៃមធ្យម (៛)", 0) or ((in_p1 + in_p2) / 2.0))
            in_p1s = str(r.get("ថ្ងៃចាប់ផ្ដើមវគ្គ១", def_p1_s)).strip()[:10]
            in_p1e = str(r.get("ថ្ងៃបញ្ចប់វគ្គ១", def_p1_e)).strip()[:10]
            in_p2s = str(r.get("ថ្ងៃចាប់ផ្ដើមវគ្គ២", def_p2_s)).strip()[:10]
            in_p2e = str(r.get("ថ្ងៃបញ្ចប់វគ្គ២", def_p2_e)).strip()[:10]
            in_lvl = "school" if in_sch else "commune"

            if in_name and (in_comm or in_sch):
              # Check existing
              ch_sql = "SELECT id FROM products WHERE item_name=? AND price_level=?"
              ch_p = [in_name, in_lvl]
              if in_lvl == "school":
                ch_sql += " AND school_name=?"
                ch_p.append(in_sch)
              else:
                ch_sql += " AND commune=?"
                ch_p.append(in_comm)
              r_ex = cursor.execute(ch_sql, tuple(ch_p)).fetchone()
              if r_ex:
                cursor.execute(
                    """
                    UPDATE products 
                    SET price_phase1=?, price_phase2=?, price_avg=?,
                        phase1_start=?, phase1_end=?, phase2_start=?, phase2_end=?,
                        category=?, commune=?, school_name=?, price_level=?
                    WHERE id=?
                    """,
                    (in_p1, in_p2, in_avg, in_p1s, in_p1e, in_p2s, in_p2e, in_cat, in_comm, in_sch, in_lvl, r_ex[0])
                )
              else:
                cursor.execute(
                    """
                    INSERT INTO products (item_name, commune, school_name, price_level,
                                          price_phase1, price_phase2, price_avg,
                                          phase1_start, phase1_end, phase2_start, phase2_end, category)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (in_name, in_comm, in_sch, in_lvl, in_p1, in_p2, in_avg, in_p1s, in_p1e, in_p2s, in_p2e, in_cat),
                )
              count += 1
          conn.commit()
          st.success(f"🎉 បាននាំចូលមុខទំនិញចំនួន {count} ដោយជោគជ័យ!")
          st.rerun()
      else:
        st.warning("មិនអាចទាញយកទិន្នន័យតារាងពីឯកសារនេះបានទេ!")

  # ----------------- បញ្ជីមុខទំនិញទាំងអស់ក្នុងប្រព័ន្ធ -----------------
  st.divider()
  st.subheader(
      f"📋 បញ្ជីមុខទំនិញ និងកាលបរិច្ឆេទវគ្គ"
      + (f" សម្រាប់ {target_label}" if target_label else " ទាំងអស់ក្នុងប្រព័ន្ធ")
  )
  prod_sql = """
      SELECT item_name as [មុខទំនិញ], 
             category as [ប្រភេទសម្គាល់],
             CASE WHEN price_level='school' THEN '🏫 តាមសាលា' ELSE '🏛️ តាមឃុំ' END as [កម្រិត],
             COALESCE(school_name, '-') as [សាលារៀន],
             commune as [ឃុំ/សង្កាត់], 
             price_phase1 as [តម្លៃវគ្គ១ (៛)], 
             COALESCE(phase1_start || ' ដល់ ' || phase1_end, '-') as [កាលបរិច្ឆេទវគ្គ១],
             price_phase2 as [តម្លៃវគ្គ២ (៛)],
             COALESCE(phase2_start || ' ដល់ ' || phase2_end, '-') as [កាលបរិច្ឆេទវគ្គ២],
             price_avg as [តម្លៃមធ្យម (៛)],
             phase1_start, phase1_end, phase2_start, phase2_end
      FROM products 
  """
  if is_school_mode and s_val:
    prod_sql += " WHERE school_name=? AND price_level='school' ORDER BY id DESC"
    df_prod = pd.read_sql_query(prod_sql, conn, params=(s_val,))
  elif not is_school_mode and c_val:
    prod_sql += " WHERE commune=? AND (school_name IS NULL OR school_name='' OR price_level='commune') ORDER BY id DESC"
    df_prod = pd.read_sql_query(prod_sql, conn, params=(c_val,))
  else:
    prod_sql += " ORDER BY price_level, commune, school_name, id DESC"
    df_prod = pd.read_sql_query(prod_sql, conn)

  if not df_prod.empty:
    today_str = str(date.today())
    def calc_active_status(row):
      p1_s, p1_e = str(row.get("phase1_start") or ""), str(row.get("phase1_end") or "")
      p2_s, p2_e = str(row.get("phase2_start") or ""), str(row.get("phase2_end") or "")
      if p1_s and p1_e and (p1_s <= today_str <= p1_e):
        return "🟢 កំពុងអនុវត្ត វគ្គ១"
      elif p2_s and p2_e and (p2_s <= today_str <= p2_e):
        return "🔵 កំពុងអនុវត្ត វគ្គ២"
      elif p1_s and today_str >= p1_s and (not p1_e or today_str <= p1_e):
        return "🟢 វគ្គ១"
      elif p2_s and today_str >= p2_s and (not p2_e or today_str <= p2_e):
        return "🔵 វគ្គ២"
      return "⚪ ក្រៅកាលបរិច្ឆេទ"

    df_prod["ប្រភេទសម្គាល់"] = df_prod["ប្រភេទសម្គាល់"].apply(lambda c: format_category_badge(c) if c else "⚪ មិនទាន់កំណត់")
    df_prod["វគ្គបច្ចុប្បន្ន (ស្វ័យប្រវត្ត)"] = df_prod.apply(calc_active_status, axis=1)
    disp_cols = [
        "មុខទំនិញ", "ប្រភេទសម្គាល់", "កម្រិត", "សាលារៀន", "ឃុំ/សង្កាត់",
        "តម្លៃវគ្គ១ (៛)", "កាលបរិច្ឆេទវគ្គ១",
        "តម្លៃវគ្គ២ (៛)", "កាលបរិច្ឆេទវគ្គ២",
        "តម្លៃមធ្យម (៛)", "វគ្គបច្ចុប្បន្ន (ស្វ័យប្រវត្ត)"
    ]
    df_prod_display = df_prod[[c for c in disp_cols if c in df_prod.columns]]
    st.dataframe(
        add_row_numbers(df_prod_display), use_container_width=True, hide_index=True
    )
  else:
    st.info("មិនទាន់មានមុខទំនិញនៅឡើយទេ។")

# ================= ៤. កត់ត្រា និងចេញវិក្កយបត្រប្រចាំថ្ងៃ (ឧបសម្ពន្ធ ៣) =================
elif menu == "📝 កត់ត្រា និងចេញវិក្កយបត្រប្រចាំថ្ងៃ":
  st.title("📝 វិក្កយបត្រផ្លូវការ (ឧបសម្ពន្ធ ៣) និងកត់ត្រាប្រចាំថ្ងៃ")

  rec_main_tab1, rec_main_tab2, rec_main_tab3, rec_main_tab4 = st.tabs([
      "🧾 ប័ណ្ណទទួលស្បៀង (វិក្កយបត្រ ឧបសម្ពន្ធ ៣ / Invoice 1)",
      "📅 តារាងបញ្ជាទិញប្រចាំថ្ងៃតាមសាលា (School Daily Matrix)",
      "➕ កត់ត្រាការលក់ប្រចាំថ្ងៃ (ទម្រង់ទោល)",
      "📥 នាំចូលឯកសារពីខាងក្រៅ (Excel, XLSM, Word, PDF, រូបភាព OCR)",
  ])

  # ----------------- TAB 1: វិក្កយបត្រ ឧបសម្ពន្ធ ៣ ផ្លូវការ -----------------
  with rec_main_tab1:
    st.subheader("🧾 បង្កាន់ដៃទទួលស្បៀង (វិក្កយបត្រផ្លូវការ ឧបសម្ពន្ធ ៣ - តាមគំរូសន្លឹក Invoice 1)")
    st.info("💡 គំរូទម្រង់ផ្លូវការតាមសន្លឹកកិច្ចការ «Invoice 1» នៃឯកសារ Excel «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» បេះបិទ ១០០% | លេខសក្ខីប័ត្របង្កើតឡើងដោយស្វ័យប្រវត្ត ចាប់ផ្ដើមពី 001 តាមសាលានីមួយៗ")

    # ១. ជួរជ្រើសរើសទីតាំងតៗគ្នា (Cascading: District -> Commune -> School) និងកាលបរិច្ឆេទ
    col_d, col_c, col_s, col_dt = st.columns([1, 1, 1.3, 1])

    with col_d:
      dist_list = get_districts()
      inv_dist = st.selectbox("ក្រុង/ស្រុក", ["-- ទាំងអស់ --"] + dist_list, key="annex3_dist")

    with col_c:
      filtered_communes = get_communes(district=inv_dist if inv_dist != "-- ទាំងអស់ --" else None)
      inv_comm = st.selectbox("ឃុំ/សង្កាត់", ["-- ទាំងអស់ --"] + filtered_communes, key="annex3_comm")

    with col_s:
      if inv_comm != "-- ទាំងអស់ --":
        school_options = get_schools_by_commune(inv_comm)
      elif inv_dist != "-- ទាំងអស់ --":
        rows_s = cursor.execute(
            "SELECT DISTINCT name FROM schools WHERE district=? OR commune IN (SELECT commune FROM locations WHERE district=?)",
            (inv_dist, inv_dist)
        ).fetchall()
        school_options = [r[0] for r in rows_s]
      else:
        school_options = get_all_schools()

      if not school_options:
        school_options = get_all_schools()

      inv_school = st.selectbox("សាលាបឋមសិក្សា", school_options if school_options else ["គ្មានសាលា"], key="annex3_school")

    # ទាញយកទីតាំងពិតរបស់សាលាពី Database
    prov_db, dist_db, comm_db, vill_db = get_school_location_info(inv_school)
    act_district = dist_db if dist_db else (inv_dist if inv_dist != "-- ទាំងអស់ --" else "")
    act_commune = comm_db if comm_db else (inv_comm if inv_comm != "-- ទាំងអស់ --" else "")

    with col_dt:
      school_dates = [r[0] for r in cursor.execute("SELECT DISTINCT date FROM daily_records WHERE school_name=? ORDER BY date DESC", (inv_school,)).fetchall()]
      dt_c1, dt_c2 = st.columns(2)
      with dt_c1:
        inv_date = st.date_input("ថ្ងៃដាក់ (Purchase)", date.today(), key="annex3_date")
      with dt_c2:
        rec_eat_row = cursor.execute("SELECT consumption_date FROM daily_records WHERE school_name=? AND date=? AND consumption_date IS NOT NULL AND consumption_date != '' LIMIT 1", (inv_school, str(inv_date))).fetchone()
        default_eat_d = parse_date_safe(rec_eat_row[0]) if (rec_eat_row and rec_eat_row[0]) else date.fromordinal(inv_date.toordinal() + 1)
        inv_eat_date = st.date_input("ថ្ងៃស៊ី/ញ៉ាំ (Consumption)", default_eat_d, key="annex3_eat_date")
      if school_dates and str(inv_date) not in school_dates:
        st.caption(f"📅 ថ្ងៃមានទិន្នន័យចុងក្រោយ: `{school_dates[0]}`")

    # ២. ជួរលេខសក្ខីប័ត្រស្វ័យប្រវត្ត ចាប់ពី 001 តាមសាលា និងឈ្មោះអ្នកផ្គត់ផ្គង់
    col_v1, col_v2, col_v3 = st.columns([1, 1.2, 1.5])

    # គណនាលេខសក្ខីប័ត្រអូតូ ចាប់ផ្ដើមពី 001 តាមសាលានីមួយៗ
    auto_v_no = get_or_create_voucher_no(inv_school, str(inv_date))

    # ទាញយកព័ត៌មានអ្នកផ្គត់ផ្គង់ដែលផ្គត់ផ្គង់សាលានេះពី Database
    school_sup_info = get_supplier_for_school(inv_school)
    default_sup_name = school_sup_info.get("supplier_name", "សាត ក្រូត") if school_sup_info else "សាត ក្រូត"
    saved_sup_sig = school_sup_info.get("signature_data") if school_sup_info else None

    with col_v1:
      cur_voucher_no = st.text_input("លេខសក្ខីប័ត្រ (អូតូតាមសាលា)", value=auto_v_no, key=f"annex3_vno_{inv_school}_{inv_date}")
      st.caption("🔢 ចាប់ផ្ដើមពី `001` ដោយឡែកតាមសាលានីមួយៗ")

    with col_v2:
      supplier_name = st.text_input("ឈ្មោះអ្នកផ្គត់ផ្គង់ស្បៀង", value=default_sup_name, key=f"annex3_sup_{inv_school}_{inv_date}")
      if school_sup_info:
        sup_addr_badge = format_supplier_address(school_sup_info)
        st.caption(f"🚚 📞 `{school_sup_info.get('phone', 'N/A')}` | 🏠 {sup_addr_badge}")

    with col_v3:
      inv_comment = st.text_input("យោបល់ចំពោះទំនិញ (ប្រសិនបើមាន)", value="", key=f"annex3_cmt_{inv_school}_{inv_date}")

    # មុខងារបញ្ចូល និងគ្រប់គ្រងហត្ថលេខាលើវិក្កយបត្រ A5 (Signatures)
    with st.expander("✍️ មុខងារបញ្ចូល និងគ្រប់គ្រងហត្ថលេខាលើវិក្កយបត្រ (Signatures)", expanded=False):
      st.markdown("###### ជ្រើសរើស ឬបញ្ចូលហត្ថលេខាសម្រាប់ឯកសារវិក្កយបត្រ A5")
      st.caption("💡 អាចប្រើហត្ថលេខាដែលបានរក្សាទុករបស់អ្នកផ្គត់ផ្គង់, Upload ហត្ថលេខាថ្មី, ឬទុកចន្លោះចុចៗ (.........) សម្រាប់ចុះហត្ថលេខាផ្ទាល់ដៃលើក្រដាស។")
      s_col1, s_col2, s_col3 = st.columns(3)

      with s_col1:
        active_sup_sig = render_signature_uploader_with_tools(
            label="១. ហត្ថលេខាអ្នកផ្គត់ផ្គង់ (អ្នកប្រគល់)",
            key_prefix=f"inv_sup_{inv_school}_{inv_date}",
            allow_use_saved=True,
            saved_sig_b64=saved_sup_sig,
            allow_blank_choice=True,
            default_blank=False,
            default_recolor="blue"
        )

      with s_col2:
        active_dir_sig = render_signature_uploader_with_tools(
            label="២. ហត្ថលេខា/ត្រានាយកសាលា (បានឃើញ និងឯកភាព)",
            key_prefix=f"inv_dir_{inv_school}_{inv_date}",
            allow_use_saved=False,
            allow_blank_choice=True,
            default_blank=False,
            default_recolor="red"
        )

      with s_col3:
        active_rec_sig = render_signature_uploader_with_tools(
            label="៣. ហត្ថលេខាអ្នកទទួលស្បៀង (នាយឃ្លាំង/បេឡាធិការ)",
            key_prefix=f"inv_rec_{inv_school}_{inv_date}",
            allow_use_saved=False,
            allow_blank_choice=True,
            default_blank=False,
            default_recolor="blue"
        )

    # ទាញយកទំនិញសម្រាប់ថ្ងៃ និងសាលានោះ
    query = """SELECT item_name as [មុខទំនិញ], phase as [វគ្គ], quantity as [បរិមាណ], 
                      unit_price as [តម្លៃរាយ (៛)], total_price as [សរុប (៛)], id 
               FROM daily_records 
               WHERE school_name=? AND date=? 
               ORDER BY id ASC"""
    df_inv_full = pd.read_sql_query(query, conn, params=(inv_school, str(inv_date)))
    df_inv = df_inv_full[["មុខទំនិញ", "វគ្គ", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)"]] if not df_inv_full.empty else pd.DataFrame()

    # ៣. ប្រអប់បន្ថែម ឬកែប្រែទំនិញដោយផ្ទាល់លើវិក្កយបត្រនេះ
    with st.expander("✏️ បន្ថែម / កែសម្រួលទំនិញក្នុងវិក្កយបត្រនេះ", expanded=False):
      col_add1, col_add2, col_add3, col_add4, col_add5 = st.columns([2, 1, 1, 1.2, 1])
      with col_add1:
        prod_info_map = get_products_map(school_name=inv_school, commune=act_commune, supplier_name=supplier_name)
        prod_map = {p: (info["price_phase1"], info["price_phase2"]) for p, info in prod_info_map.items()}
        prod_choices = list(prod_map.keys()) + ["➕ វាយទំនិញថ្មី..."]
        quick_item_sel = st.selectbox("មុខទំនិញ", prod_choices, key="quick_add_item")
        if quick_item_sel == "➕ វាយទំនិញថ្មី...":
          quick_item_name = st.text_input("វាយឈ្មោះទំនិញថ្មី", key="quick_add_new_name").strip()
        else:
          quick_item_name = quick_item_sel

      with col_add2:
        quick_phase = st.selectbox("វគ្គ", ["វគ្គ១", "វគ្គ២"], key="quick_add_phase")

      with col_add3:
        quick_qty = st.number_input("បរិមាណ", min_value=0.1, value=1.0, step=0.1, key="quick_add_qty")

      with col_add4:
        def_price = 0.0
        if quick_item_name in prod_map:
          p1, p2 = prod_map[quick_item_name]
          def_price = p1 if quick_phase == "វគ្គ១" else p2
        quick_price = st.number_input("តម្លៃរាយ (៛)", min_value=0.0, value=def_price, step=100.0, format="%.0f", key="quick_add_price")

      with col_add5:
        st.write("")
        st.write("")
        if st.button("➕ បន្ថែម", key="btn_quick_add", use_container_width=True):
          if not quick_item_name:
            st.error("សូមជ្រើសរើស ឬវាយបញ្ចូលឈ្មោះទំនិញ!")
          else:
            q_tot = quick_qty * quick_price
            cursor.execute(
                """INSERT INTO daily_records (date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (str(inv_date), inv_school, quick_item_name, quick_phase, quick_qty, quick_price, q_tot, cur_voucher_no)
            )
            conn.commit()
            st.success(f"បានបន្ថែម {quick_item_name}!")
            st.rerun()

      if not df_inv_full.empty:
        col_price_info, col_price_btn = st.columns([2.2, 1.2])
        with col_price_info:
          st.caption(f"🏷️ **តម្លៃសម្រាប់គណនា៖** ផ្អែកលើអ្នកផ្គត់ផ្គង់ **«{supplier_name}»** | សាលា **«{inv_school}»** | ឃុំ **«{act_commune}»**")
        with col_price_btn:
          if st.button("🔄 គណនាតម្លៃឡើងវិញស្វ័យប្រវត្ត", key="btn_refresh_inv_prices", use_container_width=True):
            recalc_cnt = 0
            for _, r_it in df_inv_full.iterrows():
              it_nm = r_it["មុខទំនិញ"]
              it_ph = r_it["វគ្គ"] if "វគ្គ" in r_it and r_it["វគ្គ"] else "វគ្គ១"
              if it_nm in prod_info_map:
                u_p = prod_info_map[it_nm]["price_phase1"] if it_ph == "វគ្គ១" else prod_info_map[it_nm]["price_phase2"]
                if u_p > 0:
                  tot_p = float(r_it["បរិមាណ"]) * u_p
                  cursor.execute("UPDATE daily_records SET unit_price=?, total_price=? WHERE id=?", (u_p, tot_p, r_it["id"]))
                  recalc_cnt += 1
            conn.commit()
            st.success(f"✅ បានកែសម្រួលតម្លៃ {recalc_cnt} មុខទំនិញទៅតាមតម្លៃអ្នកផ្គត់ផ្គង់ សាលា និងឃុំ!")
            st.rerun()

        st.markdown("###### បញ្ជីទំនិញបច្ចុប្បន្ន (អាចលុបបាន):")
        for _, row_item in df_inv_full.iterrows():
          c_del1, c_del2, c_del3, c_del4 = st.columns([3, 1.5, 2, 1])
          with c_del1:
            st.write(f"• **{row_item['មុខទំនិញ']}** ({row_item['វគ្គ']})")
          with c_del2:
            st.write(f"បរិមាណ: **{row_item['បរិមាណ']}**")
          with c_del3:
            st.write(f"សរុប: **{format_riel(row_item['សរុប (៛)'])}**")
          with c_del4:
            if st.button("🗑️ លុប", key=f"del_item_{row_item['id']}"):
              cursor.execute("DELETE FROM daily_records WHERE id=?", (row_item['id'],))
              conn.commit()
              st.rerun()

    # បង្កើតកូដ HTML វិក្កយបត្រផ្លូវការ A5
    html_code = generate_annex3_html(
        district=act_district,
        commune=act_commune,
        school_name=inv_school,
        voucher_no=cur_voucher_no,
        invoice_date=inv_date,
        df_items=df_inv,
        supplier_name=supplier_name,
        comment=inv_comment,
        supplier_sig=active_sup_sig,
        director_sig=active_dir_sig,
        receiver_sig=active_rec_sig,
        consumption_date=inv_eat_date
    )

    # ៤. ប៊ូតុងបញ្ជា ទាញយកឯកសារ (Print, PDF, Excel, HTML) និងរក្សាទុក
    col_btn1, col_btn_print, col_btn2, col_btn3, col_btn4 = st.columns([1.1, 1.1, 1.1, 1.1, 0.9])

    with col_btn1:
      if st.button("💾 រក្សាទុកសក្ខីប័ត្រ", key="btn_save_vno", use_container_width=True):
        save_voucher_no(inv_school, str(inv_date), cur_voucher_no)
        st.success(f"បានរក្សាទុកលេខសក្ខីប័ត្រ {cur_voucher_no} សម្រាប់សាលា {inv_school} កាលបរិច្ឆេទ {inv_date}!")
        st.rerun()

    with col_btn_print:
      clean_html_str = html_code.replace("</script>", "<\\/script>")
      print_btn_code = f"""
      <!DOCTYPE html>
      <html>
      <head><meta charset="utf-8">
      <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        .p-btn {{
          width: 100%;
          height: 38px;
          background: #0284c7;
          color: #ffffff;
          border: none;
          border-radius: 6px;
          font-size: 13.5px;
          font-weight: 600;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 6px;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          transition: background-color 0.15s ease;
        }}
        .p-btn:hover {{
          background: #0369a1;
        }}
      </style>
      </head>
      <body>
        <button class="p-btn" onclick="openPrint()">🖨️ បោះពុម្ព (Print)</button>
        <script>
          function openPrint() {{
            var pw = window.open('', '_blank');
            if (pw) {{
              pw.document.open();
              pw.document.write({json.dumps(clean_html_str)});
              pw.document.close();
              pw.focus();
              setTimeout(function() {{
                pw.print();
              }}, 450);
            }} else {{
              alert('សូមអនុញ្ញាត Popups ក្នុងកម្មវិធីរុករក (Browser) ដើម្បីបោះពុម្ព!');
            }}
          }}
        </script>
      </body>
      </html>
      """
      st.components.v1.html(print_btn_code, height=42)

    with col_btn2:
      pdf_bytes = generate_annex3_pdf(
          district=act_district,
          commune=act_commune,
          school_name=inv_school,
          voucher_no=cur_voucher_no,
          invoice_date=inv_date,
          df_items=df_inv,
          supplier_name=supplier_name,
          comment=inv_comment,
          supplier_sig=active_sup_sig,
          director_sig=active_dir_sig,
          receiver_sig=active_rec_sig,
          consumption_date=inv_eat_date
      )
      st.download_button(
          "📄 ទាញយកជា PDF",
          data=pdf_bytes,
          file_name=f"ឧបសម្ពន្ធ៣_{inv_school}_{inv_date}_{cur_voucher_no}.pdf",
          mime="application/pdf",
          use_container_width=True
      )

    with col_btn3:
      excel_bytes = generate_annex3_excel(
          district=act_district,
          commune=act_commune,
          school_name=inv_school,
          voucher_no=cur_voucher_no,
          invoice_date=inv_date,
          df_items=df_inv,
          supplier_name=supplier_name,
          comment=inv_comment,
          supplier_sig=active_sup_sig,
          director_sig=active_dir_sig,
          receiver_sig=active_rec_sig,
          consumption_date=inv_eat_date
      )
      st.download_button(
          "📥 ទាញយកជា Excel",
          data=excel_bytes,
          file_name=f"Invoice1_{inv_school}_{inv_date}_{cur_voucher_no}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          use_container_width=True
      )

    with col_btn4:
      st.download_button(
          "🌐 ទាញយក HTML",
          data=html_code.encode("utf-8"),
          file_name=f"ឧបសម្ពន្ធ៣_{inv_school}_{inv_date}_{cur_voucher_no}.html",
          mime="text/html",
          use_container_width=True
      )

    # ៥. មុខងារទាញយកសៀវភៅបង្កាន់ដៃប្រចាំខែទាំងអស់ (Full Invoice 1 Workbook)
    cur_month_str = str(inv_date)[:7]
    with st.expander("📚 ទាញយកសៀវភៅបង្កាន់ដៃប្រចាំខែទាំងអស់ (Full Monthly Invoice 1 Workbook)", expanded=False):
      st.info(f"💡 ទាញយកបង្កាន់ដៃទាំងអស់ក្នុងខែ `{cur_month_str}` សម្រាប់សាលា «{inv_school}» ដោយរៀបចំជាប្លុកបញ្ឈរ (២៧ ជួរដេកក្នុងមួយវិក្កយបត្រ) ដូចសន្លឹកកិច្ចការ «Invoice 1» នៃឯកសារ Excel «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» ទាំងស្រុង!")
      all_invoices_excel = generate_all_school_invoices_excel(
          district=act_district,
          commune=act_commune,
          school_name=inv_school,
          month_prefix=cur_month_str,
          supplier_name=supplier_name,
          comment=inv_comment,
          supplier_sig=active_sup_sig,
          director_sig=active_dir_sig,
          receiver_sig=active_rec_sig
      )
      st.download_button(
          f"📥 ទាញយកសៀវភៅបង្កាន់ដៃខែ {cur_month_str} ទាំងអស់ (ដូច Invoice 1 ទាំងស្រុង)",
          data=all_invoices_excel,
          file_name=f"Invoice_1_{inv_school}_{cur_month_str}_All.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          use_container_width=True
      )

    # ៦. បង្ហាញទិដ្ឋភាពពិតនៃប័ណ្ណទទួលស្បៀង (Live Preview លើអេក្រង់)
    st.markdown("#### 👁️ ទិដ្ឋភាពពិតនៃប័ណ្ណទទួលស្បៀង (Live Preview - តាមគំរូ Invoice 1)")
    st.components.v1.html(html_code, height=940, scrolling=True)

  # ----------------- TAB 2: តារាងបញ្ជាទិញប្រចាំថ្ងៃតាមសាលា (School Daily Matrix) -----------------
  with rec_main_tab2:
    esi.render_school_daily_matrix_tab(
        conn, cursor,
        get_districts, get_communes, get_schools_by_commune, get_all_schools,
        get_products_map, format_riel, to_excel, parse_date_safe
    )

  # ----------------- TAB 3: កត់ត្រាការលក់ប្រចាំថ្ងៃ (ទម្រង់ទោល) -----------------
  with rec_main_tab3:
    st.subheader("➕ កត់ត្រាការលក់ប្រចាំថ្ងៃ (ទម្រង់ទោល)")
    c_rd1, c_rd2 = st.columns(2)
    with c_rd1:
      rec_date = st.date_input("កាលបរិច្ឆេទដាក់ (Purchase Date)", date.today(), key="rec_entry_date")
    with c_rd2:
      rec_eat_date = st.date_input("កាលបរិច្ឆេទហូប (Consumption Date)", date.fromordinal(rec_date.toordinal() + 1), key="rec_entry_eat_date")

    st.markdown("##### 📍 ជ្រើសរើសទីតាំង និងសាលារៀន")
    direct_search = st.toggle("🔍 ស្វែងរកតាមឈ្មោះសាលាផ្ទាល់", value=False, key="rec_entry_toggle")

    chosen_school = ""
    chosen_commune = ""

    if direct_search:
      all_s = get_all_schools()
      s_pick = st.selectbox(
          "ជ្រើសរើសសាលា",
          ["-- ជ្រើសរើស --"] + all_s + ["➕ វាយបញ្ចូលសាលាថ្មី..."],
          key="rec_direct_sch",
      )
      if s_pick == "➕ វាយបញ្ចូលសាលាថ្មី...":
        chosen_school = st.text_input(
            "✏️ វាយបញ្ចូលឈ្មោះសាលាថ្មី", key="rec_d_s_in"
        ).strip()
        chosen_commune = st.text_input(
            "ស្ថិតក្នុងឃុំ/សង្កាត់", key="rec_d_c_in"
        ).strip()
      elif s_pick != "-- ជ្រើសរើស --":
        chosen_school = s_pick
        r_c = cursor.execute(
            "SELECT commune, district, province FROM schools WHERE name=?",
            (chosen_school,),
        ).fetchone()
        if r_c:
          chosen_commune = r_c[0] or ""
          loc_desc = f"📍 ឃុំ: **{chosen_commune}**"
          if r_c[1]:
            loc_desc += f" | ស្រុក: **{r_c[1]}**"
          if r_c[2]:
            loc_desc += f" | ខេត្ត: **{r_c[2]}**"
          st.caption(loc_desc)
    else:
      # Cascading: Province -> District -> Commune -> Auto-detect School!
      c_p1, c_p2 = st.columns(2)
      with c_p1:
        p_list = get_provinces()
        rec_p_sel = st.selectbox(
            "ខេត្ត/ក្រុង",
            ["-- ទាំងអស់/ជ្រើសរើស --"] + p_list + ["➕ វាយបញ្ចូលថ្មី..."],
            key="rec_p_sel",
        )
        if rec_p_sel == "➕ វាយបញ្ចូលថ្មី...":
          rec_act_p = st.text_input("វាយខេត្តថ្មី", key="rec_p_in").strip()
        elif rec_p_sel != "-- ទាំងអស់/ជ្រើសរើស --":
          rec_act_p = rec_p_sel
        else:
          rec_act_p = ""

      with c_p2:
        d_list = (
            get_districts(rec_act_p) if rec_act_p else get_districts()
        )
        rec_d_sel = st.selectbox(
            "ស្រុក/ខណ្ឌ",
            ["-- ទាំងអស់/ជ្រើសរើស --"] + d_list + ["➕ វាយបញ្ចូលថ្មី..."],
            key="rec_d_sel",
        )
        if rec_d_sel == "➕ វាយបញ្ចូលថ្មី...":
          rec_act_d = st.text_input("វាយស្រុកថ្មី", key="rec_d_in").strip()
        elif rec_d_sel != "-- ទាំងអស់/ជ្រើសរើស --":
          rec_act_d = rec_d_sel
        else:
          rec_act_d = ""

      c_c1, c_s1 = st.columns(2)
      with c_c1:
        c_list = (
            get_communes(rec_act_p, rec_act_d)
            if (rec_act_p and rec_act_d)
            else (
                get_communes(district=rec_act_d)
                if rec_act_d
                else get_communes()
            )
        )
        rec_c_sel = st.selectbox(
            "ឃុំ/សង្កាត់",
            ["-- ជ្រើសរើសឃុំ --"] + c_list + ["➕ វាយបញ្ចូលថ្មី..."],
            key="rec_c_sel",
        )
        if rec_c_sel == "➕ វាយបញ្ចូលថ្មី...":
          chosen_commune = st.text_input(
              "វាយឃុំថ្មី", key="rec_c_in"
          ).strip()
        elif rec_c_sel != "-- ជ្រើសរើសឃុំ --":
          chosen_commune = rec_c_sel
        else:
          chosen_commune = ""

      with c_s1:
        schools_in_commune = (
            get_schools_by_commune(chosen_commune) if chosen_commune else []
        )
        s_options = (
            ["-- ជ្រើសរើសសាលា --"]
            + schools_in_commune
            + ["➕ វាយបញ្ចូលសាលាថ្មី..."]
        )
        rec_s_sel = st.selectbox(
            "🏫 ឈ្មោះសាលា (ស្វ័យប្រវត្ត)", s_options, key="rec_s_sel"
        )
        if rec_s_sel == "➕ វាយបញ្ចូលសាលាថ្មី..." or (
            not schools_in_commune and chosen_commune
        ):
          chosen_school = st.text_input(
              "✏️ វាយបញ្ចូលឈ្មោះសាលាថ្មី", key="rec_s_in"
          ).strip()
        elif rec_s_sel != "-- ជ្រើសរើសសាលា --":
          chosen_school = rec_s_sel
        else:
          chosen_school = ""

    st.divider()
    st.markdown("##### 📦 ជ្រើសរើសមុខទំនិញ និងគណនាតម្លៃស្វ័យប្រវត្ត")

    # ទាញយកទំនិញព្រមទាំងកាលបរិច្ឆេទវគ្គតាមសាលា ឃុំ ឬអ្នកផ្គត់ផ្គង់
    sch_sup_tab2 = get_supplier_for_school(chosen_school) if chosen_school else None
    sup_name_tab2 = sch_sup_tab2.get("supplier_name") if sch_sup_tab2 else None
    prod_info_map = get_products_map(school_name=chosen_school, commune=chosen_commune, supplier_name=sup_name_tab2)

    item_options = (
        list(prod_info_map.keys()) + ["➕ វាយបញ្ចូលមុខទំនិញថ្មី..."]
        if prod_info_map
        else (STANDARD_PRODUCT_ITEMS + ["➕ វាយបញ្ចូលមុខទំនិញថ្មី..."])
    )
    rec_item_sel = st.selectbox("📦 មុខទំនិញ", item_options, key="rec_item_sel")

    # គណនាវគ្គស្វ័យប្រវត្តផ្អែកលើកាលបរិច្ឆេទ (rec_date) និងទំនិញ
    auto_phase = "វគ្គ១"
    auto_reason = ""
    selected_p1_s, selected_p1_e, selected_p2_s, selected_p2_e = None, None, None, None

    if rec_item_sel in prod_info_map:
      item_data = prod_info_map[rec_item_sel]
      selected_p1_s = item_data["p1_start"]
      selected_p1_e = item_data["p1_end"]
      selected_p2_s = item_data["p2_start"]
      selected_p2_e = item_data["p2_end"]
      auto_phase, auto_reason = detect_phase_from_date(
          rec_date, selected_p1_s, selected_p1_e, selected_p2_s, selected_p2_e
      )
    else:
      # Fallback to commune dates or system dates
      comm_p1_s, comm_p1_e, comm_p2_s, comm_p2_e = get_commune_phase_dates(chosen_commune)
      selected_p1_s, selected_p1_e, selected_p2_s, selected_p2_e = comm_p1_s, comm_p1_e, comm_p2_s, comm_p2_e
      auto_phase, auto_reason = detect_phase_from_date(
          rec_date, comm_p1_s, comm_p1_e, comm_p2_s, comm_p2_e
      )

    col_ph_rad, col_ph_banner = st.columns([1, 1.8])
    with col_ph_rad:
      rec_phase = st.radio(
          "វគ្គអនុវត្ត (គណនាស្វ័យប្រវត្តិ)",
          ["វគ្គ១", "វគ្គ២"],
          index=1 if auto_phase == "វគ្គ២" else 0,
          horizontal=True,
          key=f"rec_phase_dyn_{rec_date}_{rec_item_sel}",
      )
    with col_ph_banner:
      st.info(f"⚡ **ស្វ័យប្រវត្តិ:** {auto_reason} ➔ **{auto_phase}**")

    if rec_item_sel == "➕ វាយបញ្ចូលមុខទំនិញថ្មី...":
      rec_item = st.text_input("វាយឈ្មោះទំនិញថ្មី", key="rec_new_item").strip()
      unit_p = st.number_input(
          f"តម្លៃឯកតា ({rec_phase}) (៛)",
          min_value=0.0,
          value=0.0,
          step=100.0,
          format="%.0f",
      )
    else:
      rec_item = rec_item_sel
      item_data = prod_info_map[rec_item]
      unit_p = item_data["price_phase1"] if rec_phase == "វគ្គ១" else item_data["price_phase2"]
      st.info(f"🏷️ តម្លៃឯកតាស្វ័យប្រវត្តិ ({rec_phase}): **{format_riel(unit_p)}**")

    qty = st.number_input("ចំនួនបរិមាណ", min_value=0.01, value=1.0)
    total_p = unit_p * qty
    st.markdown(f"#### 💰 តម្លៃសរុប (ស្វ័យប្រវត្តិ): **{format_riel(total_p)}**")

    if st.button("កត់ត្រាចូលប្រព័ន្ធ", use_container_width=True):
      if not chosen_school:
        st.error("សូមជ្រើសរើស ឬវាយបញ្ចូលឈ្មោះសាលារៀនសិន!")
      elif not rec_item:
        st.error("សូមជ្រើសរើស ឬវាយបញ្ចូលមុខទំនិញ!")
      else:
        save_school(chosen_school, chosen_commune)

        if (
            rec_item_sel == "➕ វាយបញ្ចូលមុខទំនិញថ្មី..."
            and chosen_commune
            and rec_item
        ):
          p1 = unit_p if rec_phase == "វគ្គ១" else 0.0
          p2 = unit_p if rec_phase == "វគ្គ២" else 0.0
          p_avg_val = (p1 + p2) / 2.0
          cursor.execute(
              """INSERT OR IGNORE INTO products (item_name, commune, price_phase1, price_phase2, price_avg, phase1_start, phase1_end, phase2_start, phase2_end) 
                 VALUES (?,?,?,?,?,?,?,?,?)""",
              (rec_item, chosen_commune, p1, p2, p_avg_val, selected_p1_s, selected_p1_e, selected_p2_s, selected_p2_e),
          )

        auto_v_no = get_or_create_voucher_no(chosen_school, rec_date)
        cursor.execute(
            """
                INSERT INTO daily_records (date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no, consumption_date)
                VALUES (?,?,?,?,?,?,?,?,?)
            """,
            (
                str(rec_date),
                chosen_school,
                rec_item,
                rec_phase,
                qty,
                unit_p,
                total_p,
                auto_v_no,
                str(rec_eat_date),
            ),
        )
        conn.commit()
        st.success(
            f"🎉 បានកត់ត្រាជោគជ័យ! សាលា: {chosen_school} | ទំនិញ: {rec_item} |"
            f" លេខសក្ខីប័ត្រ: `{auto_v_no}` | សរុប: {format_riel(total_p)}"
        )
        st.info("💡 អ្នកអាចមើល និងទាញយកប័ណ្ណទទួលស្បៀងផ្លូវការ (ឧបសម្ពន្ធ ៣) នៅក្នុងផ្ទាំង '🧾 ប័ណ្ណទទួលស្បៀង (វិក្កយបត្រ ឧបសម្ពន្ធ ៣)' ខាងលើ!")
        st.rerun()

  # ----------------- TAB 4: នាំចូលឯកសារពីខាងក្រៅ -----------------
  with rec_main_tab4:
    esi.render_universal_importer_tab(conn, cursor, format_riel)

# ================= ៥. សំណើទូទាត់ប្រចាំខែ =================
elif menu == "📑 សំណើទូទាត់ប្រចាំខែ":
  st.title("📑 សំណើសុំទូទាត់ប្រចាំខែ (Monthly Payment Proposal)")
  st.info(
      "💡 គំរូទម្រង់ផ្លូវការ សំណើសុំទូទាត់ប្រចាំខែ តាមឯកសារស្កេន ១០០% | "
      "ជួរឈរ **'លេខយោងក្នុងបង្កាន់ដៃទទួលទំនិញ'** ចាប់យកលេខសក្ខីប័ត្រដំបូង និងចុងក្រោយក្នុងខែស្វ័យប្រវត្តិ (គំរូ `001 - 016`) | "
      "បង្គត់ទឹកប្រាក់ស្វ័យប្រវត្តិតាមច្បាប់ហិរញ្ញវត្ថុ"
  )

  # មុខងារនាំចូលសំណើទូទាត់ពីឯកសារខាងក្រៅ
  with st.expander("📥 នាំចូលសំណើទូទាត់ពីឯកសារខាងក្រៅ (Excel, Word, PDF, រូបភាព)", expanded=False):
    c_cl_up1, c_cl_up2 = st.columns([1.5, 1])
    with c_cl_up1:
      up_claim_file = st.file_uploader(
          "📂 ជ្រើសរើសឯកសារសំណើទូទាត់ (xlsm, xlsx, docx, pdf, png, jpg)",
          type=["xlsm", "xlsx", "xls", "docx", "doc", "pdf", "png", "jpg", "jpeg"],
          key="up_claim_file"
      )
    with c_cl_up2:
      st.markdown("###### ⚡ ឬប្រើឯកសារដែលមានស្រាប់:")
      if os.path.exists("បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm"):
        if st.button("📖 ផ្ទុកពី «បញ្ជីមុខម្ហូប_2026.xlsm»", key="btn_quick_claim_load", use_container_width=True):
          st.session_state["claim_load_source"] = "បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm"

    claim_parse_src = up_claim_file or st.session_state.get("claim_load_source")
    if claim_parse_src:
      if isinstance(claim_parse_src, str):
        with open(claim_parse_src, "rb") as f_c:
          cl_parsed = esi.parse_excel_workbook(f_c.read(), db_conn=conn)
      else:
        cl_parsed = esi.parse_external_file(claim_parse_src, db_conn=conn)

      if cl_parsed and cl_parsed.get("sheets"):
        claim_sheet_opts = [k for k, v in cl_parsed["sheets"].items() if v.get("sheet_type") == "monthly_claim"]
        if not claim_sheet_opts:
          claim_sheet_opts = list(cl_parsed["sheets"].keys())

        pick_cl_sheet = st.selectbox("ជ្រើសរើសសន្លឹកសំណើទូទាត់៖", claim_sheet_opts, key="pick_cl_sheet_sel")
        chosen_sheet_data = cl_parsed["sheets"][pick_cl_sheet]

        if st.button(f"⚡ បញ្ចូលទិន្នន័យពីសន្លឹក [{pick_cl_sheet}] ទៅក្នុងទម្រង់សំណើទូទាត់", key="btn_apply_cl_sheet", use_container_width=True):
          imported_items_list = []
          for _, r_it in chosen_sheet_data["df"].iterrows():
            it_name_clean = str(r_it.get("មុខទំនិញ", "")).strip()
            if it_name_clean:
              imported_items_list.append({
                  "name": it_name_clean,
                  "category": classify_item_category(it_name_clean),
                  "voucher_ref": str(r_it.get("លេខយោងបង្កាន់ដៃ", "")).strip(),
                  "qty": float(r_it.get("បរិមាណ", 0) or 0),
                  "unit_price": float(r_it.get("តម្លៃរាយ (៛)", 0) or 0),
                  "total_price": float(r_it.get("សរុប (៛)", 0) or 0)
              })
          if chosen_sheet_data.get("school_name"):
            st.session_state["claim_school_preset"] = chosen_sheet_data["school_name"]
          if chosen_sheet_data.get("voucher_no"):
            st.session_state["claim_vno_preset"] = chosen_sheet_data["voucher_no"]
          st.session_state["claim_imported_items_preset"] = imported_items_list
          st.success(f"🎉 បានទាញយកទិន្នន័យមុខទំនិញចំនួន {len(imported_items_list)} មុខ ពីសន្លឹក {pick_cl_sheet} ដោយជោគជ័យ!")
          st.rerun()

  # ១. ជួរជ្រើសរើសទីតាំងតៗគ្នា (District -> Commune -> School)
  c_d, c_c, c_s = st.columns([1, 1, 1.3])
  with c_d:
    dist_list = get_districts()
    claim_dist_sel = st.selectbox("ក្រុង/ស្រុក/ខណ្ឌ", ["-- ទាំងអស់ --"] + dist_list, key="claim_dist_sel")
  with c_c:
    filtered_communes = get_communes(district=claim_dist_sel if claim_dist_sel != "-- ទាំងអស់ --" else None)
    claim_comm_sel = st.selectbox("ឃុំ/សង្កាត់", ["-- ទាំងអស់ --"] + filtered_communes, key="claim_comm_sel")
  with c_s:
    if claim_comm_sel != "-- ទាំងអស់ --":
      school_options = get_schools_by_commune(claim_comm_sel)
    elif claim_dist_sel != "-- ទាំងអស់ --":
      rows_s = cursor.execute(
          "SELECT DISTINCT name FROM schools WHERE district=? OR commune IN (SELECT commune FROM locations WHERE district=?)",
          (claim_dist_sel, claim_dist_sel),
      ).fetchall()
      school_options = [r[0] for r in rows_s]
    else:
      school_options = get_all_schools()

    if not school_options:
      school_options = get_all_schools()

    def_sch_idx = 0
    if st.session_state.get("claim_school_preset") in school_options:
      def_sch_idx = school_options.index(st.session_state["claim_school_preset"])
    claim_school = st.selectbox("សាលាបឋមសិក្សា", school_options if school_options else ["គ្មានសាលា"], index=def_sch_idx, key="claim_school_sel")

  prov_db, dist_db, comm_db, vill_db = get_school_location_info(claim_school)
  act_district = dist_db if dist_db else (claim_dist_sel if claim_dist_sel != "-- ទាំងអស់ --" else "ស្រីស្នំ")
  act_commune = comm_db if comm_db else (claim_comm_sel if claim_comm_sel != "-- ទាំងអស់ --" else "ស្លែងស្ពាន")

  # ២. ជួរកាលបរិច្ឆេទចាប់ផ្ដើម និងបញ្ចប់
  today = date.today()
  col_dt1, col_dt2, col_vno = st.columns([1, 1, 1])
  with col_dt1:
    def_s_date = date(2026, 7, 31) if today.year == 2026 else date(today.year, max(1, today.month - 1), 1)
    claim_start_date = st.date_input("ចាប់ពីថ្ងៃទី (Start Date)", value=def_s_date, key="claim_d_start")
  with col_dt2:
    def_e_date = date(2026, 8, 31) if today.year == 2026 else date(today.year, today.month, 1)
    claim_end_date = st.date_input("ដល់ថ្ងៃទី (End Date)", value=def_e_date, key="claim_d_end")
  with col_vno:
    vno_def = st.session_state.get("claim_vno_preset", "0008")
    claim_voucher_no = st.text_input("លេខសក្ខីប័ត្រសំណើ (Voucher No)", value=vno_def, key=f"claim_vno_{claim_school}")

  # ៣. ព័ត៌មានអ្នកផ្គត់ផ្គង់ (ឈ្មោះ, អាសយដ្ឋាន, លេខទូរស័ព្ទ)
  claim_sup_info = get_supplier_for_school(claim_school)
  if claim_sup_info:
    claim_def_name = claim_sup_info.get("supplier_name", "សាត ក្រូត")
    claim_def_addr = format_supplier_address(claim_sup_info) or "ភូមិខ្មែរ ឃុំរោង"
    claim_def_phone = claim_sup_info.get("phone", "090 854 133")
    saved_claim_sig = claim_sup_info.get("signature_data")
  else:
    claim_def_name = "សាត ក្រូត"
    claim_def_addr = "ភូមិខ្មែរ ឃុំរោង"
    claim_def_phone = "090 854 133"
    saved_claim_sig = None

  col_sup1, col_sup2, col_sup3 = st.columns([1, 1.2, 1])
  with col_sup1:
    supplier_name = st.text_input("ឈ្មោះអ្នកផ្គត់ផ្គង់", value=claim_def_name, key=f"claim_sup_name_{claim_school}")
  with col_sup2:
    supplier_address = st.text_input("អាសយដ្ឋាន", value=claim_def_addr, key=f"claim_sup_addr_{claim_school}")
  with col_sup3:
    supplier_phone = st.text_input("លេខទូរស័ព្ទ", value=claim_def_phone, key=f"claim_sup_phone_{claim_school}")

  # មុខងារគ្រប់គ្រង និងបញ្ចូលហត្ថលេខាលើសំណើសុំទូទាត់ប្រចាំខែ (A4 Monthly Claim Signatures)
  with st.expander("✍️ មុខងារបញ្ចូល និងគ្រប់គ្រងហត្ថលេខាលើសំណើទូទាត់ (Signatures)", expanded=False):
    st.markdown("###### ជ្រើសរើស ឬបញ្ចូលហត្ថលេខាសម្រាប់ឯកសារសំណើទូទាត់ប្រចាំខែ A4")
    st.caption("💡 អាចប្រើហត្ថលេខាដែលបានរក្សាទុករបស់អ្នកផ្គត់ផ្គង់, Upload ហត្ថលេខាថ្មី, ឬទុកចន្លោះចុចៗ (.........) សម្រាប់ចុះហត្ថលេខាផ្ទាល់ដៃលើក្រដាស A4។")
    cs_col1, cs_col2, cs_col3 = st.columns(3)

    with cs_col1:
      active_claim_sup_sig = render_signature_uploader_with_tools(
          label="១. ហត្ថលេខាអ្នកផ្គត់ផ្គង់ (អ្នកស្នើសុំ)",
          key_prefix=f"claim_sup_{claim_school}_{claim_start_date}",
          allow_use_saved=True,
          saved_sig_b64=saved_claim_sig,
          allow_blank_choice=True,
          default_blank=False,
          default_recolor="blue"
      )

    with cs_col2:
      active_claim_dir_sig = render_signature_uploader_with_tools(
          label="២. ហត្ថលេខា/ត្រានាយកសាលា (បានឃើញ និងឯកភាព)",
          key_prefix=f"claim_dir_{claim_school}_{claim_start_date}",
          allow_use_saved=False,
          allow_blank_choice=True,
          default_blank=False,
          default_recolor="red"
      )

    with cs_col3:
      active_claim_prep_sig = render_signature_uploader_with_tools(
          label="៣. ហត្ថលេខាអ្នកធ្វើតារាង / គណនេយ្យករ",
          key_prefix=f"claim_prep_{claim_school}_{claim_start_date}",
          allow_use_saved=False,
          allow_blank_choice=True,
          default_blank=False,
          default_recolor="blue"
      )

  # ៤. ទាញយក និងគ្រប់គ្រងទិន្នន័យមុខទំនិញជាមុន ដើម្បីចាប់យកប្រភេទមុខទំនិញស្វ័យប្រវត្តិ
  session_key = f"claim_items_data_{claim_school}_{claim_start_date}_{claim_end_date}"
  if "claim_imported_items_preset" in st.session_state and st.session_state["claim_imported_items_preset"]:
    st.session_state[session_key] = st.session_state.pop("claim_imported_items_preset")
  elif session_key not in st.session_state:
    db_items = get_monthly_claim_items(claim_school, claim_start_date, claim_end_date)
    st.session_state[session_key] = db_items

  current_items = list(st.session_state[session_key])

  # ស្វែងរកប្រភេទសម្គាល់ស្វ័យប្រវត្តិពីទិន្នន័យមុខទំនិញក្នុងសំណើ និងអ្នកផ្គត់ផ្គង់
  detected_cats = set()
  for it in current_items:
    it_n = it.get("name", "")
    row_c = cursor.execute("SELECT category FROM products WHERE item_name=? LIMIT 1", (it_n,)).fetchone()
    c_name = row_c[0] if (row_c and row_c[0]) else classify_item_category(it_n)
    if c_name:
      detected_cats.add(c_name)

  # បើទំនិញទទេ យកតាមអ្នកផ្គត់ផ្គង់ដែលបានកំណត់ក្នុង School Supplier
  sup_cats_raw = [c.strip() for c in (claim_sup_info.get("supplied_categories") or "").split(",") if c.strip()]
  if not detected_cats and sup_cats_raw:
    detected_cats = set(sup_cats_raw)

  # ៥. ប្រភេទចំណាយដែលត្រូវគូសធីក (🗹) លើក្បាលតារាង (ចាប់យកដោយស្វ័យប្រវត្ត)
  st.markdown("##### 🏷️ ប្រភេទមុខទំនិញដែលត្រូវគូសធីក (🗹 ក្នុងក្បាលតារាងសំណើទូទាត់):")
  if detected_cats:
    st.caption(f"🤖 **ប្រព័ន្ធចាប់យកស្វ័យប្រវត្តិតាមមុខទំនិញ:** {' • '.join([format_category_badge(c) for c in detected_cats])}")
  else:
    st.caption("ℹ️ ធីកជ្រើសរើសប្រភេទមុខទំនិញដែលបានផ្គត់ផ្គង់ក្នុងខែនេះ")

  auto_rice = ("អង្ករ" in detected_cats) if detected_cats else True
  auto_salt = ("អំបិល" in detected_cats) if detected_cats else True
  auto_oil = ("ប្រេងឆា" in detected_cats) if detected_cats else True
  auto_meat = any(k in c for c in detected_cats for k in ["សាច់", "ត្រី", "ស៊ុត"]) if detected_cats else True
  auto_veg = ("បន្លែ" in detected_cats) if detected_cats else True

  chk_c1, chk_c2, chk_c3, chk_c4, chk_c5 = st.columns(5)
  with chk_c1:
    chk_rice = st.checkbox("អង្ករ", value=auto_rice, key=f"chk_cat_rice_{claim_school}_{len(current_items)}")
  with chk_c2:
    chk_salt = st.checkbox("អំបិល", value=auto_salt, key=f"chk_cat_salt_{claim_school}_{len(current_items)}")
  with chk_c3:
    chk_oil = st.checkbox("ប្រេងឆា", value=auto_oil, key=f"chk_cat_oil_{claim_school}_{len(current_items)}")
  with chk_c4:
    chk_meat = st.checkbox("សាច់ ត្រី ស៊ុត", value=auto_meat, key=f"chk_cat_meat_{claim_school}_{len(current_items)}")
  with chk_c5:
    chk_veg = st.checkbox("បន្លែ", value=auto_veg, key=f"chk_cat_veg_{claim_school}_{len(current_items)}")

  checked_categories = []
  if chk_rice: checked_categories.append("អង្ករ")
  if chk_salt: checked_categories.append("អំបិល")
  if chk_oil: checked_categories.append("ប្រេងឆា")
  if chk_meat: checked_categories.append("សាច់ ត្រី ស៊ុត")
  if chk_veg: checked_categories.append("បន្លែ")

  # ប៊ូតុងជំនួយ: ផ្ទុកទិន្នន័យគំរូ ១៨ មុខ (ដូចឯកសារស្កេន) ឬ ទាញយកពី Database ឡើងវិញ
  col_act1, col_act2 = st.columns(2)
  with col_act1:
    if st.button("🌟 ផ្ទុកទិន្នន័យគំរូ (១៨ មុខទំនិញដូចឯកសារស្កេន 001 - 016)", key="btn_load_sample_claim", use_container_width=True):
      st.session_state[session_key] = [
          {"name": "ប្រេងឆា", "category": "ប្រេងឆា", "voucher_ref": "001 - 016", "qty": 23.7, "unit_price": 6499.0, "total_price": 154026.0},
          {"name": "អំបិលអ៊ីយ៉ូត", "category": "អំបិល", "voucher_ref": "001 - 016", "qty": 5.0, "unit_price": 1000.0, "total_price": 5000.0},
          {"name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី សាច់ ស៊ុត", "voucher_ref": "001 - 016", "qty": 19.0, "unit_price": 15950.0, "total_price": 303050.0},
          {"name": "ស៊ុតទា", "category": "ត្រី សាច់ ស៊ុត", "voucher_ref": "001 - 016", "qty": 305.0, "unit_price": 580.0, "total_price": 176900.0},
          {"name": "ត្រីរស់", "category": "ត្រី សាច់ ស៊ុត", "voucher_ref": "001 - 016", "qty": 57.2, "unit_price": 10000.0, "total_price": 572000.0},
          {"name": "ត្រីអណ្តែង", "category": "ត្រី សាច់ ស៊ុត", "voucher_ref": "001 - 016", "qty": 20.8, "unit_price": 7990.0, "total_price": 166192.0},
          {"name": "ត្រកួន", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 44.4, "unit_price": 2495.0, "total_price": 110778.0},
          {"name": "ស្លឹកបាស", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 8.8, "unit_price": 3999.0, "total_price": 35191.0},
          {"name": "ស្លឹកម្រុំ", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 8.8, "unit_price": 2440.0, "total_price": 21472.0},
          {"name": "ស្ពៃក្រញាញ់", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 44.4, "unit_price": 3999.0, "total_price": 177556.0},
          {"name": "ស្ពៃចង្កឹះ", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 22.0, "unit_price": 3950.0, "total_price": 86900.0},
          {"name": "ផ្លែល្ពៅ", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 8.8, "unit_price": 2790.0, "total_price": 24552.0},
          {"name": "ផ្លែត្រឡាច", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 16.8, "unit_price": 2450.0, "total_price": 41160.0},
          {"name": "ប៉េងប៉ោះ", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 16.8, "unit_price": 3495.0, "total_price": 58716.0},
          {"name": "ល្ហុងខ្ចី", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 8.8, "unit_price": 1495.0, "total_price": 13156.0},
          {"name": "សណ្តែកកួរ", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 30.8, "unit_price": 3999.0, "total_price": 123169.0},
          {"name": "ការ៉ុត", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 11.5, "unit_price": 2999.0, "total_price": 34489.0},
          {"name": "ផ្កាខាត់ណា", "category": "បន្លែ", "voucher_ref": "001 - 016", "qty": 55.5, "unit_price": 4890.0, "total_price": 271395.0},
      ]
      st.rerun()

  with col_act2:
    if st.button("🔄 ទាញយកទិន្នន័យពី Database ឡើងវិញ (Refresh from DB)", key="btn_refresh_claim", use_container_width=True):
      st.session_state[session_key] = get_monthly_claim_items(claim_school, claim_start_date, claim_end_date)
      st.rerun()

  # ៦. ប្រអប់បន្ថែម ឬកែប្រែទំនិញ និងលេខយោងក្នុងតារាង
  with st.expander("✏️ បន្ថែម / កែសម្រួលមុខទំនិញ និងលេខយោងក្នុងតារាង (Custom Editor)", expanded=False):
    st.info("💡 អ្នកអាចបន្ថែម កែឈ្មោះ កែបរិមាណ តម្លៃឯកតា ឬកែប្រែលេខយោង (ឧ. `001 - 016`) ដោយផ្ទាល់នៅក្នុងតារាងខាងក្រោម:")
    for it_obj in current_items:
      if "category" not in it_obj or not it_obj["category"]:
        row_c = cursor.execute("SELECT category FROM products WHERE item_name=? LIMIT 1", (it_obj.get("name", ""),)).fetchone()
        it_obj["category"] = row_c[0] if (row_c and row_c[0]) else classify_item_category(it_obj.get("name", ""))

    raw_edit_df = pd.DataFrame(current_items) if current_items else pd.DataFrame(columns=["name", "category", "voucher_ref", "qty", "unit_price", "total_price"])
    rename_cols = {
        "name": "បរិយាយមុខទំនិញ",
        "category": "ប្រភេទសម្គាល់",
        "voucher_ref": "លេខយោងក្នុងបង្កាន់ដៃទទួលទំនិញ",
        "qty": "បរិមាណ (គ.ក)",
        "unit_price": "តម្លៃឯកតា (៛)",
        "total_price": "សរុប (៛)"
    }
    df_for_editor = raw_edit_df.rename(columns=rename_cols)
    for col_k in rename_cols.values():
      if col_k not in df_for_editor.columns:
        df_for_editor[col_k] = ""
    df_for_editor = df_for_editor[list(rename_cols.values())]

    edited_claim_df = st.data_editor(
        df_for_editor,
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_{session_key}"
    )

    # ធ្វើបច្ចុប្បន្នភាពទិន្នន័យ
    updated_items = []
    for _, r in edited_claim_df.iterrows():
      item_n = str(r.get("បរិយាយមុខទំនិញ", "")).strip()
      if not item_n:
        continue
      v_ref_val = str(r.get("លេខយោងក្នុងបង្កាន់ដៃទទួលទំនិញ", "")).strip()
      try:
        q_val = float(r.get("បរិមាណ (គ.ក)", 0) or 0)
      except Exception:
        q_val = 0.0
      try:
        u_val = float(r.get("តម្លៃឯកតា (៛)", 0) or 0)
      except Exception:
        u_val = 0.0
      t_val = round(q_val * u_val, 2)
      cat_val = str(r.get("ប្រភេទសម្គាល់", "")).strip() or classify_item_category(item_n)
      updated_items.append({
          "name": item_n,
          "category": cat_val,
          "voucher_ref": v_ref_val,
          "qty": q_val,
          "unit_price": u_val,
          "total_price": t_val
      })

    if updated_items != current_items:
      st.session_state[session_key] = updated_items
      current_items = updated_items

    # មុខងាររក្សាទុកចូល Database ប្រសិនបើចង់កត់ត្រាចូលជា Daily Records
    if st.button("💾 រក្សាទុកមុខទំនិញទាំងអស់នេះចូលកំណត់ត្រា Database", key="btn_save_claim_to_db"):
      count_saved = 0
      for it_s in current_items:
        cursor.execute(
            """INSERT INTO daily_records (date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no)
               VALUES (?,?,?,?,?,?,?,?)""",
            (str(claim_end_date), claim_school, it_s["name"], "វគ្គ១", it_s["qty"], it_s["unit_price"], it_s["total_price"], it_s["voucher_ref"])
        )
        count_saved += 1
      conn.commit()
      st.success(f"🎉 បានកត់ត្រាមុខទំនិញទាំង {count_saved} ចូលប្រព័ន្ធដោយជោគជ័យ!")

  # ៧. បង្ហាញ Card សង្ខេបទឹកប្រាក់ និងការបង្គត់
  exact_sum = sum(float(it["total_price"]) for it in current_items) if current_items else 0.0
  rounded_sum = round_khmer_currency(exact_sum)

  col_m1, col_m2, col_m3 = st.columns(3)
  with col_m1:
    st.metric("📦 ចំនួនមុខទំនិញសរុប", f"{len(current_items)} មុខ")
  with col_m2:
    st.metric("💰 សរុបទឹកប្រាក់ (Exact Sum)", format_riel(exact_sum))
  with col_m3:
    st.metric("🎯 ថវិកាសរុបស្នើសុំទូទាត់ (បង្គត់លេខ)", format_riel(rounded_sum))
  st.caption("ℹ️ * សម្គាល់៖ បង្គត់ស្មើ ០ បើខ្ទង់ដប់តិចជាង ៥០ រៀល | បង្គត់ឡើងមួយ បើខ្ទង់ដប់ស្មើឬច្រើនជាង ៥០ រៀល ។")

  st.divider()

  # ៨. ប៊ូតុងទាញយកឯកសារផ្លូវការ (PDF, Excel, HTML)
  claim_html_code = generate_monthly_claim_html(
      district=act_district,
      commune=act_commune,
      school_name=claim_school,
      voucher_no=claim_voucher_no,
      d_start=claim_start_date,
      d_end=claim_end_date,
      items=current_items,
      supplier_name=supplier_name,
      supplier_address=supplier_address,
      supplier_phone=supplier_phone,
      checked_cats=checked_categories,
      supplier_sig=active_claim_sup_sig,
      director_sig=active_claim_dir_sig,
      preparer_sig=active_claim_prep_sig
  )

  col_btn1, col_btn2, col_btn3 = st.columns(3)
  with col_btn1:
    claim_pdf_bytes = generate_monthly_claim_pdf(
        district=act_district,
        commune=act_commune,
        school_name=claim_school,
        voucher_no=claim_voucher_no,
        d_start=claim_start_date,
        d_end=claim_end_date,
        items=current_items,
        supplier_name=supplier_name,
        supplier_address=supplier_address,
        supplier_phone=supplier_phone,
        checked_cats=checked_categories,
        supplier_sig=active_claim_sup_sig,
        director_sig=active_claim_dir_sig,
        preparer_sig=active_claim_prep_sig
    )
    st.download_button(
        "📄 ទាញយកជា PDF (សំណើទូទាត់ប្រចាំខែ)",
        data=claim_pdf_bytes,
        file_name=f"សំណើទូទាត់ប្រចាំខែ_{claim_school}_{claim_voucher_no}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

  with col_btn2:
    claim_excel_bytes = generate_monthly_claim_excel(
        district=act_district,
        commune=act_commune,
        school_name=claim_school,
        voucher_no=claim_voucher_no,
        d_start=claim_start_date,
        d_end=claim_end_date,
        items=current_items,
        supplier_name=supplier_name,
        supplier_address=supplier_address,
        supplier_phone=supplier_phone,
        checked_cats=checked_categories
    )
    st.download_button(
        "📥 ទាញយកជា Excel (សំណើទូទាត់ប្រចាំខែ)",
        data=claim_excel_bytes,
        file_name=f"សំណើទូទាត់ប្រចាំខែ_{claim_school}_{claim_voucher_no}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

  with col_btn3:
    st.download_button(
        "🌐 ទាញយកជា HTML",
        data=claim_html_code.encode("utf-8"),
        file_name=f"សំណើទូទាត់ប្រចាំខែ_{claim_school}_{claim_voucher_no}.html",
        mime="text/html",
        use_container_width=True
    )

  # ៩. Live Preview
  st.markdown("#### 👁️ ទិដ្ឋភាពពិតនៃសំណើសុំទូទាត់ប្រចាំខែ (Live Preview)")
  st.components.v1.html(claim_html_code, height=1080, scrolling=True)

# ================= ៦. បញ្ជីទិញទំនិញ & ជំពាក់ =================
elif menu == "🛒 បញ្ជីទិញទំនិញចូល & ជំពាក់អ្នកផ្គត់ផ្គង់":
  st.title("🛒 ការទិញទំនិញ និងគ្រប់គ្រងបំណុលអ្នកផ្គត់ផ្គង់")
  tab_buy, tab_debt, tab_import_buy = st.tabs(
      ["កត់ត្រាការទិញទំនិញចូល", "បញ្ជីជំពាក់អ្នកផ្គត់ផ្គង់", "📥 នាំចូលពីឯកសារ (Feed / Viget / Excel / OCR)"]
  )

  with tab_buy:
    with st.form("buy_form"):
      b_date = st.date_input("ថ្ងៃខែឆ្នាំទិញ", date.today())
      b_item = st.text_input("មុខទំនិញ")
      b_qty = st.number_input("ចំនួន", min_value=1.0, value=1.0)
      b_price = st.number_input(
          "តម្លៃរាយ (៛)", min_value=0.0, step=100.0, format="%.0f"
      )
      b_supplier = st.text_input("ឈ្មោះអ្នកផ្គត់ផ្គង់")
      b_status = st.selectbox("ស្ថានភាពទូទាត់", ["ទូទាត់រួច", "ជំពាក់"])

      if st.form_submit_button("រក្សាទុកការទិញ"):
        total_cost = b_qty * b_price
        cursor.execute(
            """
                INSERT INTO purchases (date, item_name, unit_price, quantity, total_price, supplier_name, status)
                VALUES (?,?,?,?,?,?,?)
            """,
            (
                str(b_date),
                b_item,
                b_price,
                b_qty,
                total_cost,
                b_supplier,
                b_status,
            ),
        )

        if b_status == "ទូទាត់រួច":
          cursor.execute(
              "INSERT INTO transactions (date, type, category, amount,"
              " description) VALUES (?,?,?,?,?)",
              (
                  str(b_date),
                  "ចំណាយ",
                  "ទិញទំនិញ",
                  total_cost,
                  f"ទិញពី {b_supplier}",
              ),
          )
        conn.commit()
        st.success(
            f"បានកត់ត្រាការទិញចូលជោគជ័យ! សរុប: {format_riel(total_cost)}"
        )

    df_p = pd.read_sql_query(
        "SELECT date as [កាលបរិច្ឆេទ], item_name as [មុខទំនិញ], unit_price as"
        " [តម្លៃរាយ (៛)], quantity as [ចំនួន], total_price as [សរុប (៛)],"
        " supplier_name as [អ្នកផ្គត់ផ្គង់], status as [ស្ថានភាព] FROM"
        " purchases ORDER BY id DESC",
        conn,
    )
    st.dataframe(add_row_numbers(df_p), use_container_width=True, hide_index=True)

  with tab_debt:
    st.subheader("បញ្ជីអ្នកផ្គត់ផ្គង់ដែលត្រូវទូទាត់ (ជំពាក់)")
    df_debts = pd.read_sql_query(
        "SELECT id as ID, date as [កាលបរិច្ឆេទ], supplier_name as"
        " [អ្នកផ្គត់ផ្គង់], item_name as [មុខទំនិញ], total_price as [បំណុលជំពាក់"
        " (៛)] FROM purchases WHERE status='ជំពាក់'",
        conn,
    )
    st.dataframe(
        add_row_numbers(df_debts), use_container_width=True, hide_index=True
    )

    if not df_debts.empty:
      st.divider()
      st.markdown("##### ធ្វើបច្ចុប្បន្នភាពការទូទាត់បំណុល")
      debt_id = st.selectbox(
          "ជ្រើសរើស ID ដើម្បីប្តូរជា 'ទូទាត់រួច'", df_debts["ID"]
      )
      if st.button("សម្គាល់ថាបានទូទាត់រួច"):
        cursor.execute(
            "UPDATE purchases SET status='ទូទាត់រួច' WHERE id=?", (debt_id,)
        )
        conn.commit()
        st.success("បានកែប្រែស្ថានភាពជោគជ័យ!")
        st.rerun()

  with tab_import_buy:
    st.subheader("📥 នាំចូលកំណត់ត្រាទិញទំនិញ/ជំពាក់ពីឯកសារខាងក្រៅ (Feed, Viget, Excel, Word, PDF, OCR)")
    st.info("💡 គាំទ្រការទាញយកទិន្នន័យពីសន្លឹក Feed និង Viget ក្នុង «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» ឬឯកសារខាងក្រៅ កែសម្រួលបាន និងគណនាតម្លៃសរុបស្វ័យប្រវត្ត។")

    c_b_u1, c_b_u2 = st.columns([1.5, 1])
    with c_b_u1:
      up_buy_file = st.file_uploader(
          "📂 ជ្រើសរើសឯកសារសម្រាប់នាំចូលការទិញ",
          type=["xlsm", "xlsx", "xls", "docx", "doc", "pdf", "png", "jpg", "jpeg", "csv"],
          key="up_buy_file"
      )
    with c_b_u2:
      st.markdown("###### ⚡ ឬផ្ទុកពីឯកសារដែលមានស្រាប់:")
      if os.path.exists("បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm"):
        if st.button("📖 ផ្ទុកពី «បញ្ជីមុខម្ហូប_2026.xlsm» (Feed/Viget)", key="btn_load_buy_ws", use_container_width=True):
          st.session_state["buy_load_source"] = "បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm"

    buy_file_src = up_buy_file or st.session_state.get("buy_load_source")
    if buy_file_src:
      with st.spinner("🔍 កំពុងវិភាគឯកសារ..."):
        if isinstance(buy_file_src, str):
          with open(buy_file_src, "rb") as f_b:
            b_parsed = esi.parse_excel_workbook(f_b.read(), db_conn=conn)
        else:
          b_parsed = esi.parse_external_file(buy_file_src, db_conn=conn)

      if b_parsed and b_parsed.get("sheets"):
        farmer_sheets = [k for k, v in b_parsed["sheets"].items() if v.get("sheet_type") == "farmer_purchases"]
        if not farmer_sheets:
          farmer_sheets = list(b_parsed["sheets"].keys())

        pick_b_sheet = st.selectbox("ជ្រើសរើសសន្លឹកទិន្នន័យទិញទំនិញ៖", farmer_sheets, key="pick_b_sheet")
        b_sheet_data = b_parsed["sheets"][pick_b_sheet]

        st.markdown(f"##### {b_sheet_data.get('title_kh', pick_b_sheet)}")
        df_buy_edit = b_sheet_data.get("df", pd.DataFrame())

        if not df_buy_edit.empty:
          ed_buy_df = st.data_editor(
              df_buy_edit,
              use_container_width=True,
              num_rows="dynamic",
              key=f"ed_buy_{pick_b_sheet}"
          )

          if "បរិមាណ" in ed_buy_df.columns and "តម្លៃរាយ (៛)" in ed_buy_df.columns and "សរុប (៛)" in ed_buy_df.columns:
            ed_buy_df["សរុប (៛)"] = (
                pd.to_numeric(ed_buy_df["បរិមាណ"], errors="coerce").fillna(0) *
                pd.to_numeric(ed_buy_df["តម្លៃរាយ (៛)"], errors="coerce").fillna(0)
            ).round(2)

          sum_b_tot = ed_buy_df["សរុប (៛)"].sum() if "សរុប (៛)" in ed_buy_df.columns else 0
          st.metric("💰 ទឹកប្រាក់សរុបនៃការទិញ", format_riel(sum_b_tot))

          if st.button(f"💾 ចម្លងទិន្នន័យទិញពី [{pick_b_sheet}] ចូលក្នុង Database", key=f"btn_save_buy_{pick_b_sheet}", use_container_width=True):
            cnt_b, msg_b = esi.import_records_to_database(ed_buy_df, target_table="purchases")
            st.success(f"🎉 {msg_b}")
            st.rerun()
        else:
          st.info(f"សន្លឹក [{pick_b_sheet}] គ្មានជួរដេកទិន្នន័យទិញទេ។")

# ================= ៧. ចំណូល និងចំណាយ =================
elif menu == "💰 ចំណូល និងចំណាយ":
  st.title("💰 បញ្ជីគ្រប់គ្រងចំណូល និងចំណាយ")
  col_t1, col_t2 = st.columns([1, 2])

  with col_t1:
    with st.form("trans_form"):
      t_date = st.date_input("កាលបរិច្ឆេទ", date.today())
      t_type = st.selectbox("ប្រភេទប្រតិបត្តិការ", ["ចំណូល", "ចំណាយ"])
      t_cat = st.text_input("ប្រភេទទូទៅ (ឧ. លក់, ថ្លៃដឹក, ទឹកភ្លើង...)")
      t_amt = st.number_input(
          "ចំនួនទឹកប្រាក់ (៛)", min_value=0.0, step=100.0, format="%.0f"
      )
      t_desc = st.text_area("ពិពណ៌នាបន្ថែម")

      if st.form_submit_button("កត់ត្រាប្រតិបត្តិការ"):
        cursor.execute(
            "INSERT INTO transactions (date, type, category, amount,"
            " description) VALUES (?,?,?,?,?)",
            (str(t_date), t_type, t_cat, t_amt, t_desc),
        )
        conn.commit()
        st.success("បានបញ្ចូលទិន្នន័យរួចរាល់!")

  with col_t2:
    df_trans = pd.read_sql_query(
        "SELECT date as [កាលបរិច្ឆេទ], type as [ប្រភេទ], category as [ប្រភេទទូទៅ],"
        " amount as [ចំនួនទឹកប្រាក់ (៛)], description as [ពិពណ៌នា] FROM"
        " transactions ORDER BY id DESC",
        conn,
    )
    st.dataframe(
        add_row_numbers(df_trans), use_container_width=True, hide_index=True
    )

    inc = (
        cursor.execute(
            "SELECT SUM(amount) FROM transactions WHERE type='ចំណូល'"
        ).fetchone()[0]
        or 0.0
    )
    exp = (
        cursor.execute(
            "SELECT SUM(amount) FROM transactions WHERE type='ចំណាយ'"
        ).fetchone()[0]
        or 0.0
    )
    st.info(
        f"ចំណូលសរុប: **{format_riel(inc)}** | ចំណាយសរុប: **{format_riel(exp)}**"
        f" | សល់សុទ្ធ: **{format_riel(inc - exp)}**"
    )

# ================= ៨. គ្រប់គ្រងអ្នកប្រើប្រាស់ =================
elif menu == "👥 គ្រប់គ្រងអ្នកប្រើប្រាស់":
  st.title("👥 គ្រប់គ្រងគណនីអ្នកប្រើប្រាស់")

  if user_info.get("role") != "Admin":
    st.error("មានតែ Admin ទេដែលអាចចូលផ្នែកនេះបាន!")
  else:
    with st.form("new_user_form"):
      st.subheader("បង្កើតគណនីថ្មី")
      new_u = st.text_input("ឈ្មោះគណនី (Username)")
      new_p = st.text_input("ពាក្យសម្ងាត់", type="password")
      new_fn = st.text_input("ឈ្មោះពេញ")
      new_r = st.selectbox("តួនាទី", ["Staff", "Admin"])

      if st.form_submit_button("បង្កើតគណនី"):
        if new_u and new_p:
          try:
            cursor.execute(
                "INSERT INTO users (username, password, full_name, role) VALUES"
                " (?,?,?,?)",
                (new_u, hash_password(new_p), new_fn, new_r),
            )
            conn.commit()
            st.success("បានបង្កើតគណនីថ្មីដោយជោគជ័យ!")
          except sqlite3.IntegrityError:
            st.error("ឈ្មោះគណនីនេះមានរួចហើយ!")
        else:
          st.warning("សូមបំពេញឈ្មោះគណនី និងពាក្យសម្ងាត់!")

    st.subheader("បញ្ជីគណនីទាំងអស់")
    df_users = pd.read_sql_query(
        "SELECT username as [គណនី], full_name as [ឈ្មោះពេញ], role as [តួនាទី] FROM"
        " users",
        conn,
    )
    st.dataframe(
        add_row_numbers(df_users), use_container_width=True, hide_index=True
    )
