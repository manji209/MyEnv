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


query = """UPDATE dbo.ITMFIL00 SET item_standard_cost=? WHERE item_no=?"""

# grab existing row count in the database for validation later
cur.execute("SELECT count(*) FROM dbo.ITMFIL00")
before_import = cur.fetchone()


for r in range(1, sheet.nrows):
    item_no = sheet.cell(r,0).value
    item_standard_cost = sheet.cell(r,7).value

    values = (item_standard_cost, item_no)

    cur.execute(query, values)
    print(item_no)

conn.commit()

# If you want to check if all rows are imported
cur.execute("SELECT count(*) FROM dbo.INVTRX00")
result = cur.fetchone()



print((result[0] - before_import[0]))  # should be True

# Close the database connection
cur.close()
conn.commit()
conn.close()