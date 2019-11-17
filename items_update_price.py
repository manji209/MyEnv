import pyodbc
import xlrd
import pandas as pd
import numpy as np

# Update inventory from Excel sheet
book = xlrd.open_workbook("Data/INV_EXTRA.xlsx")
sheet = book.sheet_by_name("Sheet3")


conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = "UPDATE dbo.ITMFIL00 SET replacement_cost=? WHERE item_no=?"

# grab existing row count in the database for validation later
cur.execute("SELECT count(*) FROM dbo.ITMFIL00")
before_import = cur.fetchone()

for r in range(1, sheet.nrows):

    item_no = sheet.cell(r,0).value
    replacement_cost = sheet.cell(r,3).value

    print("Row: ", item_no)

    values = (replacement_cost, item_no)

    cur.execute(sqlstring, values)


# If you want to check if all rows are imported
cur.execute("SELECT count(*) FROM dbo.INVTRX00")
result = cur.fetchone()

cur.close()
conn.commit()
