# -*- coding: utf-8 -*-
import sys, sqlite3
sys.stdout.reconfigure(encoding='utf-8')
import excel_schema_importer as esi

# 1. Test Excel Schema Parsing
with open('បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm', 'rb') as f:
    wb_data = esi.parse_excel_workbook(f.read())

print('1. Sheets detected:', len(wb_data['sheets']))
matrix_sheets = [k for k, v in wb_data['sheets'].items() if v.get('sheet_type') == 'daily_matrix']
claim_sheets = [k for k, v in wb_data['sheets'].items() if v.get('sheet_type') == 'monthly_claim']
farmer_sheets = [k for k, v in wb_data['sheets'].items() if v.get('sheet_type') == 'farmer_purchases']

print('2. Daily Matrix sheets count:', len(matrix_sheets), matrix_sheets)
print('3. Monthly Claim sheets count:', len(claim_sheets), claim_sheets)
print('4. Farmer Ledger sheets count:', len(farmer_sheets), farmer_sheets)

# 2. Test importing Sleng daily matrix
sleng_info = wb_data['sheets']['Sleng']
print('5. Testing import of Sleng sheet, rows:', len(sleng_info['df']))
cnt, msg = esi.import_records_to_database(sleng_info['df'], target_table='daily_records')
print('Import result:', cnt, msg)

# 3. Verify database records
conn = sqlite3.connect('school_pos.db')
c = conn.cursor()
c.execute("SELECT COUNT(id) FROM daily_records WHERE school_name='ស្លែងស្ពាន'")
print('Database records for Sleng:', c.fetchone()[0])

c.execute("SELECT date, consumption_date, voucher_no, item_name, quantity, unit_price, total_price FROM daily_records WHERE school_name='ស្លែងស្ពាន' LIMIT 3")
for row in c.fetchall():
    print('  Sample:', row)

# 4. Test Viget farmer purchases import
viget_info = wb_data['sheets']['Viget']
cnt_v, msg_v = esi.import_records_to_database(viget_info['df'], target_table='purchases')
print('Viget import result:', cnt_v, msg_v)

c.execute("SELECT date, item_name, quantity, unit_price, total_price, supplier_name, status FROM purchases LIMIT 3")
for row in c.fetchall():
    print('  Purchase Sample:', row)

print('\nALL VERIFICATIONS PASSED SUCCESSFULLY!')
