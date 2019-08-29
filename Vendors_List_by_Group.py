import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


cur = conn.cursor()


sqlstring = """SELECT vend_no, vend_nam,  addr_1, addr_2, city, state, zip_cod, pmt_grp FROM dbo.VENFIL00 ORDER BY pmt_grp"""




df = pd.read_sql(sqlstring, conn)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/Vendors_list_by_group.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', index=False)

writer.save()





