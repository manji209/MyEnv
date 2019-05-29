import pyodbc
from openpyxl import Workbook, load_workbook
import datetime
import xlrd
import pandas as pd
import numpy as np

#'DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#'DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#'DRIVER={SQL Server Native Client 11.0};SERVER=MANJI-RYZEN\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

# create Workbook object
wb=Workbook()

# Get current date for use in filename
d = datetime.datetime.today()
# Converting date into DD-MM-YYYY format
cur_date = d.strftime('%d-%m-%Y')

# set file path
filepath="Data/Invoice_Check_" + cur_date + ".xlsx"

sheet=wb.active

# Init variables for row values
inv_desc1 = ''
inv_desc2 = ''
inv_price = ''
desc1 = ''
desc2 = ''
price = ''
rep = ''
found = False

header = ('Invoice #', 'Order #', 'Customer Name', 'Item #', 'INV-Desc-1', 'INV-Desc-2', 'INV-Price', 'Desc-1', 'Desc-2', 'Price', 'Sales Rep')

# Add Header to Workbook
sheet.append(header)




# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

#sqlstring = 'SELECT invc_no, invc_dat, cust_no, item_no, desc_1, desc_2, ord_no, sls_rep FROM dbo.IHSLIN00 WHERE invc_dat = CONVERT(VARCHAR(10), getdate(), 111)'
sqlstring = """SELECT invc_no, invc_dat, cust_no, item_no, desc_1, desc_2, prc, ord_no, sls_rep FROM dbo.IHSLIN00 WHERE invc_dat = '2019-05-23'"""

df = pd.read_sql(sqlstring,conn)

# Select Item info based on Invoice Item list
sqlstring2 = """SELECT item_desc_1, item_desc_2, item_prc_1 FROM dbo.ITMFIL00 WHERE item_no=?"""

def init_var():
    inv_desc1 = ''
    inv_desc2 = ''
    invc_price = ''
    desc1 = ''
    desc2 = ''
    price = ''
    rep = ''
    found = False


#Returns the Name associated with the Sales Rep #
def get_name(num):
    switcher = {
            5: "Chi Nong",
            7: "Michelle Nong",
            11: "George Nguyen",
            12: "Terry Nguyen",
            14: "Larry Nguyen",
            15: "Linh Ung",
            16: "Pierre Bach",
            17: "Kenny Nguyen",
            18: "Phat Tran",
            19: "Minh Bui",
            20: "Sang Tran",
            21: "Tony Thai"
        }
    return switcher.get(num, "NA#")

def get_cust_name(num):
    sql = """SELECT cust_nam FROM dbo.CUSFIL00 WHERE cust_no=?"""
    value = [num]
    cur.execute(sql, value)
    results = cur.fetchone()
    return results[0]


for row in df.itertuples(index=True):
    # Init variables for row values
    inv_desc1 = ''
    inv_desc2 = ''
    inv_price = ''
    cust_no = ''
    desc1 = ''
    desc2 = ''
    price = ''
    rep = ''
    found = False

    value = [row.item_no]
    if not row.item_no.strip():
        continue
    try:
        cur.execute(sqlstring2, value)
        results = cur.fetchone()
        # Check for
        if (row.desc_1 != results[0]): #or (row.desc_2 != results[1]):
            inv_desc1 = row.desc_1
            desc1 = results[0]
            found = True

        if (row.desc_2 != results[1]):
            inv_desc2 = row.desc_2
            desc2 = results[1]
            found = True

        if float(row.prc) != float(results[2]):
            inv_price = float(row.prc)
            price = float(results[2])
            found = True

        if found:
            invoice = row.invc_no
            order = row.ord_no
            item = row.item_no
            cust_no = get_cust_name(row.cust_no)
            rep = get_name(int(row.sls_rep))
            row_add = (invoice, order, cust_no, item, inv_desc1, inv_desc2, inv_price, desc1, desc2, price, rep)
            sheet.append(row_add)

    except Exception as e:
        print(e)

# save workbook
wb.save(filepath)


#print(df)
'''
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/history_out.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Inv History', index=False)

writer.save()

'''