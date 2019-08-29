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
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


#DEMO Database For Tony Pricing
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()


sqlstring = """SELECT item_no, item_prc_1, item_standard_cost, replacement_cost FROM dbo.ITMFIL00"""

df = pd.read_sql(sqlstring, conn)

df.columns = df.columns.str.strip()

setprice_fname = "Data/Set_Price_08152019_updated.xlsx"
setprice_df = pd.read_excel(setprice_fname, sheet_name='Sheet1')

sqlupdate = "UPDATE dbo.ITMFIL00 SET item_prc_1=? WHERE item_no=?"

for idx, row in df.iterrows():
    # 21 percent markups
    cost = greater(row.item_standard_cost, row.replacement_cost)
    markup = cost / .77
    if markup > row.item_prc_1:
        rounded = roundup(markup)
    elif cost <= 8.5:
        rounded = roundup(row.item_prc_1 + ((cost / .785) - cost))
    elif cost > 8.50 and cost <= 11:
        rounded = roundup(row.item_prc_1 + ((cost / .815) - cost))
    elif cost > 11 and cost <= 16:
        rounded = roundup(row.item_prc_1 + ((cost / .825) - cost))
    elif cost > 16 and cost <= 20.5:
        rounded = roundup(row.item_prc_1 + ((cost / .845) - cost))
    elif cost > 20.5 and cost <= 26.5:
        rounded = roundup(row.item_prc_1 + ((cost / .865) - cost))
    elif cost > 26.5 and cost <= 47:
        rounded = roundup(row.item_prc_1 + ((cost / .875) - cost))
    elif cost > 47 and cost <= 60:
        rounded = roundup(row.item_prc_1 + ((cost / .885) - cost))
    else:
        # Additional 6 percent markup if imported
        rounded = roundup(row.item_prc_1 + ((cost / .90) - cost))

    values = [rounded, row.item_no.strip()]
    try:
        cur.execute(sqlupdate, values)

    except Exception as e:
        print(e)

for idx2, row2 in setprice_df.iterrows():
    values2 = [row2.price, row2.item_no.strip()]
    try:
        cur.execute(sqlupdate, values2)

    except Exception as e:
        print(e)


conn.commit()


