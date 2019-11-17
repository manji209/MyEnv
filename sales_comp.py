from difflib import SequenceMatcher
import pyodbc
from openpyxl import load_workbook
import datetime
import pandas as pd
import numpy as np


# Prompt for start Date and End Date.  Also prompt for sheet name to process

sheet_name = input("Enter Sheet Name:  ")
order_num = input("Please Enter Order Number: ")

# Load Excel File for Processing
wb = load_workbook('Data/Sales_Comp_2019.xlsx')
sheet = wb[sheet_name]


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


#Retrieve line items from order
sqlstring = """SELECT item_no, qty_ord, desc_1, desc_2, unit_prc
                FROM dbo.LINITM00
                WHERE ord_no =?"""


df_items = pd.read_sql(sqlstring, conn, params={int(order_num)})

#Retrieve cust no
sqlstring = """SELECT ord_no, ord_dat, cust_no
                FROM dbo.ORDHDR00
                WHERE ord_no =?"""

df_cust = pd.read_sql(sqlstring, conn, params={int(order_num)})

df_items['total'] = ""
df_items['order_num'] = ""
df_items['date'] = ""
df_items['cust_no'] = ""

print(df_cust)
print(df_items)

for idx, row in df_items.iterrows():
    df_items.loc[idx, 'total'] = row.qty_ord * row.unit_prc
    df_items.loc[idx, 'order_num'] = df_cust['ord_no'].iloc[0]
    df_items.loc[idx, 'date'] = df_cust['ord_dat'].iloc[0]
    df_items.loc[idx, 'cust_no'] = df_cust['cust_no'].iloc[0]


for row in df_items.itertuples(index=False):
    sheet.append(row)



wb.save(filename="Data/Sales_Comp_2019.xlsx")
