import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
# conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

# Monthly Sales Figure
sqlstring = """SELECT item_no, desc_1, desc_2, qty_ord, unit_prc, ord_no, seq_no 
                FROM dbo.LINITM00 
                WHERE ord_no=?"""

sqlstring_items = """SELECT item_no, replacement_cost
                    FROM dbo.ITMFIL00"""

ord_no = 113819

df = pd.read_sql(sqlstring, conn, params={ord_no})
df['item_no'] = df['item_no'].str.strip()
df['cost'] = ""
#df.applymap(lambda x: x.strip() if type(x)==str else x)

df_items = pd.read_sql(sqlstring_items, conn)
df_items['item_no'] = df_items['item_no'].str.strip()

for row in df.itertuples(index=True):
    rep_cost = df_items.loc[df_items['item_no'] == row.item_no, 'replacement_cost']
    df.set_value(row.Index, 'cost', rep_cost.iloc[0])

df = df[['item_no', 'desc_1', 'desc_2', 'qty_ord', 'unit_prc', 'cost', 'ord_no', 'seq_no']]

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/Cost_eval_113819.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', index=False)


writer.save()
