import pyodbc
import pandas as pd


def roundup(num):
    num += .25
    #nearest = .5
    return round((num * 2), 0) / 2


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


cur = conn.cursor()

order_no = 107904
sqlstring = """SELECT ord_no, seq_no, item_no, desc_1, qty_ord, unit_prc, unit_cost FROM dbo.LINITM00 WHERE ord_no=?"""

df = pd.read_sql(sqlstring, conn, params={order_no})

df.columns = df.columns.str.strip()

sqlstring2 = "SELECT bill_to_nam FROM dbo.ORDHDR00 WHERE ord_no=?"
values = [order_no]
cur.execute(sqlstring2, values)
result = cur.fetchone()


fname = str(order_no) + '_' + result[0].strip() + '_Markup.xlsx'

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(('Out/'+fname), engine='xlsxwriter')

#df.to_excel(writer, sheet_name='Original', index=False)

df['MarkUp_21'] = ""
df['Total'] = ""


setprice_fname = "Data/Set_Price_06082019.xlsx"
setprice_df = pd.read_excel(setprice_fname, sheet_name='Sheet1')

found = False

for idx, row in df.iterrows():
    for idx2, row2 in setprice_df.iterrows():
        if row.item_no.strip() == row2.item_no:
            #Set sale price to fixxed price
            if row.unit_prc == 0:
                df.loc[idx, 'MarkUp_21'] = 0
                df.loc[idx, 'Total'] = row.unit_prc * row.qty_ord
                found = True
                break
            else:
                df.loc[idx, 'MarkUp_21'] = row2.price
                df.loc[idx, 'Total'] = row2.price * row.qty_ord
                found = True
                break

    if found:
        found = False
        continue
    else:
        print("Sanity Check 1")
        if row.unit_prc == 0:
            df.loc[idx, 'MarkUp_21'] = 0
            df.loc[idx, 'Total'] = row.unit_prc * row.qty_ord
            found = False
        else:
            # 21 percent markups
            markup = row.unit_cost / .77
            if markup > row.unit_prc:
                rounded = roundup(markup)
                df.loc[idx, 'MarkUp_21'] = rounded
                df.loc[idx, 'Total'] = rounded * row.qty_ord
                found = False
            else:
                # Additional 6 percent markup if imported
                rounded = roundup(row.unit_prc + ((row.unit_cost / .90) - row.unit_cost))
                df.loc[idx, 'MarkUp_21'] = rounded
                df.loc[idx, 'Total'] = rounded * row.qty_ord
                found = False



sqlupdate = "UPDATE dbo.LINITM00 SET unit_prc=? WHERE seq_no=? AND ord_no=?"

total_qty = 0
total_sales = 0

for idx, row in df.iterrows():
    values2 = [row.MarkUp_21, row.seq_no, order_no]
    try:
        cur.execute(sqlupdate, values2)
        total_qty += row.qty_ord
        total_sales += row.Total
    except Exception as e:
        print(e)

conn.commit()


query_order = """UPDATE [dbo].[ORDHDR00]
                SET [tot_qty] =?
                ,[tot_sls_amt] =?
                ,[tot_gross_amt] =?
                WHERE [ord_no] =?"""

values3 = [total_qty, total_sales, total_sales, order_no]
cur.execute(query_order, values3)
conn.commit()


#df['MarkUp_21'] = df['MarkUp_21'].map('${:,.2f}'.format)

#df.to_excel(writer, sheet_name='MarkedUp', index=False)

df.to_excel(writer, sheet_name='Original', index=False)
writer.save()



