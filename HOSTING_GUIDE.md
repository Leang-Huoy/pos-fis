# ការណែនាំអំពីការ Hosting ប្រព័ន្ធ POS FIS លើ Streamlit Community Cloud (ឥតគិតថ្លៃ ១០០%)

ប្រព័ន្ធនេះត្រូវបានរៀបចំរួចរាល់ទាំងស្រុងសម្រាប់ដំណើរការនៅលើ **Streamlit Community Cloud** ដែលជាសេវាកម្ម Cloud Hosting ឥតគិតថ្លៃ ១០០% ផ្លូវការរបស់ Streamlit (Snowflake)៖
- **ឥតគិតថ្លៃ ១០០%** មិនអស់ប្រាក់ប្រចាំខែ
- **ដំណើរការ ២៤/៧** មិនបាច់បើកកុំព្យូទ័រចោល
- **មាន SSL/HTTPS ស្រាប់** សុវត្ថិភាព និងអាចចូលប្រើបានពីទូរស័ព្ទដៃ ថេប្លេត និងកុំព្យូទ័រគ្រប់ទីកន្លែង
- **មានប្រព័ន្ធទិន្នន័យ (SQLite)** និងមុខងារស្រង់ឯកសារ Excel/PDF/OCR ស្រាប់

---

## ជំហានទី ១៖ បង្កើត Repository ថ្មីនៅលើ GitHub

1. ចូលទៅកាន់គេហទំព័រ [https://github.com/new](https://github.com/new)
2. ត្រង់ **Repository name** វាយបញ្ចូល៖ `pos-fis` (ឬឈ្មោះណាមួយដែលលោកអ្នកចង់បាន)
3. ជ្រើសរើស **Public** (ដើម្បីអាចប្រើ Free Cloud Hosting បាន)
4. ចុចប៊ូតុងពណ៌បៃតង **"Create repository"**

---

## ជំហានទី ២៖ Push កូដពីកុំព្យូទ័រនេះទៅកាន់ GitHub

លោកអ្នកគ្រាន់តែចុចពីរដង (Double Click) លើឯកសារ៖
👉 **`push_to_github.bat`** (នៅក្នុង Folder `POS FIS`)

- ប្រព័ន្ធនឹងសួររក Repo URL (ចុច **Enter** ដើម្បីប្រើ `https://github.com/Leang-Huoy/pos-fis.git` ជាស្រេច)
- ប្រសិនបើមានផ្ទាំង Browser លោតឡើង សូមចុច **"Sign in with your browser"** ដើម្បីបញ្ជាក់សិទ្ធិ
- កូដទាំងអស់រួមមាន `app.py`, `excel_schema_importer.py`, `requirements.txt`, `packages.txt`, និង `school_pos.db` នឹងត្រូវបញ្ចូលទៅកាន់ GitHub ដោយស្វ័យប្រវត្តិ។

---

## ជំហានទី ៣៖ Hosting លើ Streamlit Community Cloud

1. ចូលទៅកាន់គេហទំព័រ៖ **[https://share.streamlit.io/](https://share.streamlit.io/)**
2. ចុច **"Continue with GitHub"** (ដើម្បី Sign in ជាមួយគណនី GitHub `Leang-Huoy`)
3. ចុចប៊ូតុង **"Create app"** (ឬ **"New app"**) នៅផ្នែកខាងលើ
4. ជ្រើសរើសជម្រើស៖
   - **Repository:** `Leang-Huoy/pos-fis`
   - **Branch:** `main`
   - **Main file path:** `app.py`
   - **App URL (Custom domain):** លោកអ្នកអាចកែឈ្មោះតំណភ្ជាប់បាន ឧ. `pos-fis-khmer.streamlit.app`
5. ចុចប៊ូតុងពណ៌ខៀវ **"Deploy!"**

---

## ជំហានទី ៤៖ រីករាយជាមួយប្រព័ន្ធដែលបាន Hosting រួចរាល់!

- Streamlit Cloud នឹងដំឡើង Packages ពី `requirements.txt` និង `packages.txt` ដោយស្វ័យប្រវត្តិក្នុ​ងរយៈពេលប្រហែល ១-២ នាទី។
- បន្ទាប់មក លោកអ្នកនឹងទទួលបាន Link ផ្លូវការ (ឧ. `https://pos-fis-khmer.streamlit.app`) ដែលអាចចែករំលែកឱ្យក្រុមការងារ នាយកសាលា ឬអ្នកផ្គត់ផ្គង់ចូលប្រើប្រាស់បានភ្លាមៗ!
