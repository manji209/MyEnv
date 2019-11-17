#Get list of items a customer bought withing the last 5 invoices
import pyodbc
import pandas as pd


# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


cur = conn.cursor()

# Get customer list for Sales Rep
sales_rep = 14
sqlstring = """SELECT cust_no FROM dbo.CUSFIL00 WHERE sls_rep=?
            ORDER BY cust_no ASC"""

cust_list_df = pd.read_sql(sqlstring, conn, params={sales_rep})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/CustomerList_year_qr.xlsx', engine='xlsxwriter')


# Get list of past 5 invoices
# From list of invoices get all items purchased and pivot
def get_items(cust_no):
    inv_sqlstring = """SELECT DISTINCT item_no FROM dbo.IHSLIN00 WHERE cust_no=?
                        AND (invc_dat >= DATEADD(DAY, -365, GETDATE()))
                        ORDER BY item_no ASC"""

    items_df = pd.read_sql(inv_sqlstring, conn, params={cust_no})
    items_df['item_no'] = items_df['item_no'].str.strip()
    items_df['@qr_image'] = ''

    for idx, row in items_df.iterrows():
        items_df.loc[idx, '@qr_image'] = '/QRcodes/' + row.item_no + '.png'

    return items_df


for row in cust_list_df.itertuples(index=False):
    cust_df = get_items(row.cust_no)
    if cust_df.size > 30:
        cust_df.to_excel(writer, sheet_name=row.cust_no, index=False)


writer.save()