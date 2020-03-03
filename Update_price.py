import pyodbc
import xlrd
import pandas as pd
import numpy as np

book = xlrd.open_workbook("Import/Update_Price.xlsx")
sheet = book.sheet_by_name("Sheet1")



# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()

sqlupdate = "UPDATE dbo.ITMFIL00 SET item_prc_1=? WHERE item_no=?"


cur.execute("SELECT count(*) FROM dbo.ITMFIL00")
before_import = cur.fetchone()


for r in range(1, sheet.nrows):
    item_no = sheet.cell(r,0).value
    prc = sheet.cell(r,1).value

    values = [prc, item_no]

    try:
        cur.execute(sqlupdate, values)
    except Exception as e:
        print(e)

conn.commit()

# If you want to check if all rows are imported
cur.execute("SELECT count(*) FROM dbo.ITMFIL00")
result = cur.fetchone()



print((result[0] - before_import[0]))  # should be True

# Close the database connection
conn.close()