import pyodbc
import pandas as pd


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


cur = conn.cursor()


setprice_fname = "Import/LocationCodes.xlsx"
setprice_df = pd.read_excel(setprice_fname, sheet_name='Consolidated')


sqlstring = """SELECT item_no, item_location_cod
                FROM dbo.STAFIL00
                WHERE item_no=?"""


sqlupdate = "UPDATE dbo.STAFIL00 SET item_location_cod=? WHERE item_no=?"


for idx, row in setprice_df.iterrows():
    value = (row.item_no)
    cur.execute(sqlstring, value)
    result = cur.fetchone()
    if result[1] != '#NA#':
        print(result[0])
        continue
    else:
        print("Sanity Check")
        values2 = [row.location_code, row.item_no]
        try:
            cur.execute(sqlupdate, values2)
        except Exception as e:
            print(e)

conn.commit()




