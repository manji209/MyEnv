import pyodbc
import xlrd
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()

# read data
#data = pd.read_excel('Data/import_order_entery_template.xlsx')

# Open the workbook and define the worksheet
# book = xlrd.open_workbook("Data/import_order_entery_template.xlsx")
book = xlrd.open_workbook("Import/New_Item_Templates_Test.xlsx")
sheet = book.sheet_by_name("Sheet1")
#sheet.getCells().deleteRows(sheet.nrows+1, 1, True)
total_qty = 0
total_sales = 0
ord_no = 0

query = """INSERT INTO [dbo].[ITMFIL00] (
            [item_no]
           ,[item_desc_1]
           ,[bs_acct_pc1]
           ,[bs_acct_pc2]
           ,[bs_acct_main]
           ,[bs_acct_sub]
           ,[item_no_2]
           ,[item_prod_cat]
           ,[item_prod_sub_cat]
           ,[item_no_3]
           ,[item_prime_vend]
           ,[item_no_4]
           ,[item_status]
           ,[item_no_5]
           ,[item_desc_2]
           ,[item_typ]
           ,[item_vend_prod_no]
           ,[item_vend_lead_day]
           ,[item_vend_min_ord]
           ,[stock_unit_of_meas]
           ,[prc_unit_of_meas]
           ,[conv_fac]
           ,[item_prc_cod]
           ,[commis_cod]
           ,[item_prc_1]
           ,[item_prc_2]
           ,[item_prc_3]
           ,[item_prc_4]
           ,[item_prc_5]
           ,[item_avg_cost]
           ,[item_standard_cost]
           ,[replacement_cost]
           ,[txbl_cod_1]
           ,[txbl_cod_2]
           ,[txbl_cod_3]
           ,[txbl_cod_4]
           ,[txbl_cod_5]
           ,[back_ord_cod]
           ,[sls_acct_pc1]
           ,[sls_acct_pc2]
           ,[sls_acct_main]
           ,[sls_acct_sub]
           ,[expense_acct_pc1]
           ,[expense_acct_pc2]
           ,[expense_acct_main]
           ,[expense_acct_sub]
           ,[cr_memo_acct_pc1]
           ,[cr_memo_acct_pc2]
           ,[cr_memo_acct_main]
           ,[cr_memo_acct_sub]
           ,[qty_on_hand]
           ,[cur_prd_qty_on_hnd]
           ,[qty_commitd]
           ,[qty_on_ord]
           ,[qty_on_bk_ord]
           ,[qty_on_work_ord]
           ,[warty_grace_prd]
           ,[item_prefer_unit]
           ,[weight]
           ,[height]
           ,[width]
           ,[depth]
           ,[jc_cat_no]
           ,[item_dat_created]
           ,[item_track_flg]
           ,[usr_def_qty_1]
           ,[usr_def_qty_2]
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 
           ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
           ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""

# grab existing row count in the database for validation later
cur.execute("SELECT count(*) FROM dbo.ICPHXF00")
before_import = cur.fetchone()

for r in range(1, sheet.nrows):
    item_no = sheet.cell(r,0).value
    item_desc_1 = sheet.cell(r,1).value
    bs_acct_pc1 = sheet.cell(r,2).value
    bs_acct_pc2 = sheet.cell(r,3).value
    bs_acct_main = sheet.cell(r,4).value
    bs_acct_sub = sheet.cell(r,5).value
    item_no_2 = sheet.cell(r,6).value
    item_prod_cat = sheet.cell(r, 7).value
    item_prod_sub_cat = sheet.cell(r, 8).value
    item_no_3 = sheet.cell(r,9).value
    item_prime_vend = sheet.cell(r,10).value
    item_no_4 = sheet.cell(r,11).value
    item_status = sheet.cell(r,12).value
    item_no_5 = sheet.cell(r, 13).value
    item_desc_2 = sheet.cell(r, 14).value
    item_typ = sheet.cell(r, 15).value
    item_vend_prod_no = sheet.cell(r, 16).value
    item_vend_lead_day = sheet.cell(r, 17).value
    item_vend_min_ord = sheet.cell(r, 18).value
    stock_unit_of_meas = sheet.cell(r, 19).value
    prc_unit_of_meas = sheet.cell(r, 20).value
    conv_fac = sheet.cell(r, 21).value
    item_prc_cod = sheet.cell(r, 22).value
    commis_cod = sheet.cell(r, 23).value
    item_prc_1 = sheet.cell(r, 24).value
    item_prc_2 = sheet.cell(r, 25).value
    item_prc_3 = sheet.cell(r, 26).value
    item_prc_4 = sheet.cell(r, 27).value
    item_prc_5 = sheet.cell(r, 28).value
    item_avg_cost = sheet.cell(r, 29).value
    item_standard_cost = sheet.cell(r, 30).value
    replacement_cost = sheet.cell(r, 31).value
    txbl_cod_1 = sheet.cell(r, 32).value
    txbl_cod_2 = sheet.cell(r, 33).value
    txbl_cod_3 = sheet.cell(r, 34).value
    txbl_cod_4 = sheet.cell(r, 35).value
    txbl_cod_5 = sheet.cell(r, 36).value
    back_ord_cod = sheet.cell(r, 37).value
    sls_acct_pc1 = sheet.cell(r, 38).value
    sls_acct_pc2 = sheet.cell(r, 39).value
    sls_acct_main = sheet.cell(r, 40).value
    sls_acct_sub = sheet.cell(r, 41).value
    expense_acct_pc1 = sheet.cell(r, 42).value
    expense_acct_pc2 = sheet.cell(r, 43).value
    expense_acct_main = sheet.cell(r, 44).value
    expense_acct_sub = sheet.cell(r, 45).value
    cr_memo_acct_pc1 = sheet.cell(r, 46).value
    cr_memo_acct_pc2 = sheet.cell(r, 47).value
    cr_memo_acct_main = sheet.cell(r, 48).value
    cr_memo_acct_sub = sheet.cell(r, 49).value
    qty_on_hand = sheet.cell(r, 50).value
    cur_prd_qty_on_hnd = sheet.cell(r, 51).value
    qty_commitd = sheet.cell(r, 52).value
    qty_on_ord = sheet.cell(r, 53).value
    qty_on_bk_ord = sheet.cell(r, 54).value
    qty_on_work_ord = sheet.cell(r, 55).value
    warty_grace_prd = sheet.cell(r, 56).value
    item_prefer_unit = sheet.cell(r, 57).value
    weight = sheet.cell(r, 58).value
    height = sheet.cell(r, 59).value
    width = sheet.cell(r, 60).value
    depth = sheet.cell(r, 61).value
    jc_cat_no = sheet.cell(r, 62).value
    item_dat_created = sheet.cell(r, 63).value
    item_track_flg = sheet.cell(r, 64).value
    usr_def_qty_1 = sheet.cell(r, 65).value
    usr_def_qty_21 = sheet.cell(r, 66).value

    # Assign values from each row
    values = (item_no, item_desc_1, bs_acct_pc1, bs_acct_pc2, bs_acct_main, bs_acct_sub, item_no_2, item_prod_cat,
              item_prod_sub_cat, item_no_3, item_prime_vend, item_no_4,
              item_status, item_no_5, item_desc_2, item_typ, item_vend_prod_no, item_vend_lead_day, item_vend_min_ord,
              stock_unit_of_meas, prc_unit_of_meas, conv_fac, item_prc_cod, commis_cod, item_prc_1, item_prc_2, item_prc_3,
              item_prc_4, item_prc_5, item_avg_cost, item_standard_cost, replacement_cost, txbl_cod_1, txbl_cod_2, txbl_cod_3, txbl_cod_4, txbl_cod_5,
              back_ord_cod, sls_acct_pc1, sls_acct_pc2, sls_acct_main, sls_acct_sub, expense_acct_pc1, expense_acct_pc2, expense_acct_main, expense_acct_sub,
              cr_memo_acct_pc1, cr_memo_acct_pc2, cr_memo_acct_main, cr_memo_acct_sub, qty_on_hand, cur_prd_qty_on_hnd, qty_commitd, qty_on_ord, qty_on_bk_ord,
              qty_on_work_ord, warty_grace_prd, item_prefer_unit, weight, height, width, depth, jc_cat_no, item_dat_created, item_track_flg, usr_def_qty_1, usr_def_qty_21)

    # Execute sql Query
    cur.execute(query, values)



# Commit the transaction
conn.commit()

# If you want to check if all rows are imported
cur.execute("SELECT count(*) FROM dbo.ITMFIL00")
result = cur.fetchone()

print((result[0] - before_import[0]))  # should be True

# Close the database connection
conn.close()