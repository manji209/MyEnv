import pyodbc
import xlrd
import pandas as pd
import numpy as np

book = xlrd.open_workbook("Data/INV_EXTRA.xlsx")
sheet = book.sheet_by_name("Sheet3")

# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

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
            ,[corr_flg]
            ,[qty]
            ,[actual_unit_cost]
            ,[unit_prc]
            ,[new_prc_1]
            ,[new_prc_2]
            ,[new_prc_3]
            ,[new_prc_4]
            ,[new_prc_5]
          ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""

# grab existing row count in the database for validation later
cur.execute("SELECT count(*) FROM dbo.INVTRX00")
before_import = cur.fetchone()


for r in range(1, sheet.nrows):

    item_no = sheet.cell(r,0).value
    whs = ''
    trx_typ = 0.0
    trx_dat = '2019-05-17'
    seq_no = 0.0
    batch_no = 0.0
    usr_id = ''
    lev_no = 0.0
    item_no_alt = item_no
    trx_dat_a = '2019-05-17'
    trx_typ_a = 0.0
    seq_no_alt = 0.0
    corr_flg = 'N'
    qty = sheet.cell(r,1).value
    actual_unit_cost = 0.0
    unit_prc = 0.0
    new_prc_1 = sheet.cell(r,2).value
    new_prc_2 = 0.0
    new_prc_3 = 0.0
    new_prc_4 = 0.0
    new_prc_5 = 0.0

    print("Row: ", trx_dat_a)

    values = (item_no,whs, trx_typ, trx_dat, seq_no, batch_no, usr_id, lev_no, item_no_alt, trx_dat_a, trx_typ_a, seq_no_alt,
              corr_flg, qty, actual_unit_cost, unit_prc, new_prc_1, new_prc_2, new_prc_3, new_prc_4, new_prc_5)

    cur.execute(query, values)

conn.commit()

# If you want to check if all rows are imported
cur.execute("SELECT count(*) FROM dbo.INVTRX00")
result = cur.fetchone()



print((result[0] - before_import[0]))  # should be True

# Close the database connection
conn.close()