import pyodbc
import xlrd
import pandas as pd
import numpy as np

book = xlrd.open_workbook("Import/test_sample_RECEIVE_B.xlsx")
sheet = book.sheet_by_name("Sheet1")

#REMEMBER TO CHANGE THE trx_dat and trx_dat_a to current date!!!!!!!!!!!!!!!

# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()


query = """INSERT INTO [dbo].[INVTRX00] (
            [item_no]
            ,[whs]
            ,[trx_typ]
            ,[trx_dat]
            ,[seq_no]
            ,[batch_no]
            ,[usr_id]
            ,[lev_no]
            ,[item_no_alt]
            ,[trx_dat_a]
            ,[trx_typ_a]
            ,[seq_no_alt]
            ,[doc_no]
            ,[corr_flg]
            ,[qty]
            ,[actual_unit_cost]
            ,[unit_prc]
            ,[new_prc_1]
            ,[new_prc_2]
            ,[new_prc_3]
            ,[new_prc_4]
            ,[new_prc_5]
            ,[commis_pct]
            ,[work_ord_no]
            ,[component_seq_no]
            ,[invc_hst_lin_no]
            ,[track_qty]
          ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""

# grab existing row count in the database for validation later
cur.execute("SELECT count(*) FROM dbo.INVTRX00")
before_import = cur.fetchone()


for r in range(1, sheet.nrows):
    item_no = sheet.cell(r,0).value
    whs = ''
    trx_typ = 80
    trx_dat = sheet.cell(r,5).value - 2
    seq_no = sheet.cell(r,6).value
    batch_no = 0.0
    usr_id = ''
    lev_no = 0.0
    item_no_alt = item_no
    trx_dat_a = sheet.cell(r,5).value - 2
    trx_typ_a = 80
    seq_no_alt = sheet.cell(r,6).value
    doc_no = sheet.cell(r,4).value
    corr_flg = 'N'
    qty = sheet.cell(r,1).value
    actual_unit_cost = sheet.cell(r,2).value
    unit_prc = sheet.cell(r,3).value
    new_prc_1 = 0.0
    new_prc_2 = 0.0
    new_prc_3 = 0.0
    new_prc_4 = 0.0
    new_prc_5 = 0.0
    commis_pct = 0.0
    work_ord_no = 0.0
    component_seq_no = 0.0
    invc_hst_lin_no = 0.0
    track_qty = 0.0

    values = (item_no,whs, trx_typ, trx_dat, seq_no, batch_no, usr_id, lev_no, item_no_alt, trx_dat_a, trx_typ_a, seq_no_alt,
              doc_no, corr_flg, qty, actual_unit_cost, unit_prc, new_prc_1, new_prc_2, new_prc_3, new_prc_4, new_prc_5,
              commis_pct, work_ord_no, component_seq_no, invc_hst_lin_no, track_qty)

    cur.execute(query, values)
    print(item_no)

conn.commit()

# If you want to check if all rows are imported
cur.execute("SELECT count(*) FROM dbo.INVTRX00")
result = cur.fetchone()



print((result[0] - before_import[0]))  # should be True

# Close the database connection
conn.close()