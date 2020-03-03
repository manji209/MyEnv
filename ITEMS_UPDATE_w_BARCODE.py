import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np

def connect_db():
    connect = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
        #'DRIVER={SQL Server Native Client 11.0};SERVER=MANJI-RYZEN\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
        #'DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


    return connect

# Connect to SQL Server and set cursor
conn = connect_db()

try:

    cur = conn.cursor()
    sqlstring = """SELECT keyword_token, keyword_cust_no, ic_filler_1, ic_filler_2  FROM dbo.CKEYWF00 WHERE keyword_typ = 'B'"""
    df_barcodes = pd.read_sql(sqlstring, conn)
    df_barcodes['item_no'] = df_barcodes.keyword_token
    df_barcodes['item_no'] = df_barcodes['item_no'].str.strip()

    sqlstring2 = """SELECT item_no, item_desc_1, item_desc_2 FROM dbo.ITMFIL00"""
    df_items = pd.read_sql(sqlstring2, conn)
    df_items['item_no'] = df_items['item_no'].str.strip()

    df_join = pd.merge(df_barcodes, df_items, on='item_no', how='right')
    # Remove all NaN and replace with blank space
    df_join = df_join.fillna('')

    # Initialize barcode column
    df_join['barcode'] = ''

    # Concatenate all barcode fields to one and then strip whitespace
    for index, row in df_join.iterrows():
        bcode = row.keyword_cust_no + row.ic_filler_1 + row.ic_filler_2
        df_join.loc[index, 'barcode'] = bcode.strip()

    df_join.drop(['keyword_token', 'keyword_cust_no', 'ic_filler_1', 'ic_filler_2'], axis=1, inplace=True)
    # Reorder dataframe so that barcode would be first column for VSLOOKUP purposes
    df_join = df_join[['barcode', 'item_no', 'item_desc_1', 'item_desc_2']]
    #df_join.drop(['keyword_cust_no'], axis=1)
    #df_join.drop(['ic_filler_1'], axis=1)
    #df_join.drop(['ic_filler_2'], axis=1)



finally:
    conn.close()

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Data/PO_Template/ITEMS_BARCODE.xlsx', engine='xlsxwriter')

df_join.to_excel(writer, sheet_name='Sheet1', index=False)


writer.save()
