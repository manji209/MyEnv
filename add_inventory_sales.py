import pyodbc
import xlrd
import pandas as pd
import numpy as np

book = xlrd.open_workbook("Import/INVENTORY_SALES_103019.xlsx")
sheet = book.sheet_by_name("Sheet1")

#REMEMBER TO CHANGE THE trx_dat and trx_dat_a to current date!!!!!!!!!!!!!!!

# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

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
          ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""

# grab existing row count in the database for validation later
cur.execute("SELECT count(*) FROM dbo.INVTRX00")
before_import = cur.fetchone()


for r in range(1, sheet.nrows):
    item_no = sheet.cell(r,0).value
    whs = ''
    trx_typ = 80
    trx_dat = "10-31-19"
    seq_no = sheet.cell(r,5).value
    batch_no = 0.0
    usr_id = ''
    lev_no = 0.0
    item_no_alt = item_no
    trx_dat_a = "10-31-19"
    trx_typ_a = 80
    seq_no_alt = sheet.cell(r,5).value
    doc_no = 191031
    corr_flg = 'N'
    qty = sheet.cell(r,1).value
    actual_unit_cost = sheet.cell(r,4).value
    unit_prc = sheet.cell(r,2).value
    new_prc_1 = 0.0
    new_prc_2 = 0.0
    new_prc_3 = 0.0
    new_prc_4 = 0.0
    new_prc_5 = 0.0

    values = (item_no,whs, trx_typ, trx_dat, seq_no, batch_no, usr_id, lev_no, item_no_alt, trx_dat_a, trx_typ_a, seq_no_alt,
              doc_no, corr_flg, qty, actual_unit_cost, unit_prc, new_prc_1, new_prc_2, new_prc_3, new_prc_4, new_prc_5)

    cur.execute(query, values)

conn.commit()

# If you want to check if all rows are imported
cur.execute("SELECT count(*) FROM dbo.INVTRX00")
result = cur.fetchone()



print((result[0] - before_import[0]))  # should be True

# Close the database connection
conn.close()