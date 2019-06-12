import pyodbc
import pandas as pd


# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


cur = conn.cursor()

order_no = 107203

setprice_fname = "Data/107203_Price.xlsx"
setprice_df = pd.read_excel(setprice_fname, sheet_name='Sheet1')


sqlupdate = "UPDATE dbo.LINITM00 SET unit_prc=? WHERE item_no=? AND ord_no=?"


for idx, row in setprice_df.iterrows():
    values2 = [row.Price, row.Item, order_no]
    try:
        cur.execute(sqlupdate, values2)
    except Exception as e:
        print(e)

conn.commit()




