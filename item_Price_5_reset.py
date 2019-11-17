import pyodbc
import pandas as pd


def roundup(num):
    num += .25
    #nearest = .5
    return round((num * 2), 0) / 2

def greater(standard, replacement):
    if standard > replacement:
        return standard
    else:
        return replacement


# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()


sqlstring = """SELECT item_no, item_prc_1, item_standard_cost, replacement_cost FROM dbo.ITMFIL00"""

df = pd.read_sql(sqlstring, conn)

df.columns = df.columns.str.strip()

setprice_fname = "Data/Set_Price_06082019.xlsx"
setprice_df = pd.read_excel(setprice_fname, sheet_name='Sheet1')

sqlupdate = "UPDATE dbo.ITMFIL00 SET item_prc_5=? WHERE item_no=?"

for idx, row in df.iterrows():
    rounded = 0.0
    values = [rounded, row.item_no.strip()]
    try:
        cur.execute(sqlupdate, values)

    except Exception as e:
        print(e)


conn.commit()


