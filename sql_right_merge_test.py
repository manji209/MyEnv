import pyodbc
import pandas as pd
import json

def connect_db():
    connect = pyodbc.connect(
        #'DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
        #'DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
        #'DRIVER={SQL Server Native Client 11.0};SERVER=MANJI-RYZEN\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

        'DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
        #'DRIVER={SQL Server Native Client 11.0};SERVER=MANJI-RYZEN\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
        #'DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')


    return connect


def BarcodeSearchView():

    conn = connect_db()

    try:
        cur = conn.cursor()
        sqlstring = """SELECT keyword_token, keyword_cust_no, ic_filler_1, ic_filler_2  FROM dbo.CKEYWF00 WHERE keyword_typ = 'B'"""
        df_barcodes = pd.read_sql(sqlstring, conn)
        df_barcodes['item_no'] = df_barcodes.keyword_token
        df_barcodes['item_no'] = df_barcodes['item_no'].str.strip()

        sqlstring2 = """SELECT item_no, item_desc_1, item_desc_2, qty_on_hand, item_prc_1 FROM dbo.ITMFIL00"""
        df_items = pd.read_sql(sqlstring2, conn)
        df_items['item_no'] = df_items['item_no'].str.strip()

        df_inner = pd.merge(df_barcodes, df_items, on='item_no', how='right')
        # Remove all NaN and replace with blank space
        df_inner = df_inner.fillna('')

        # Initialize barcode column
        df_inner['barcode'] = ''

        # Concatenate all barcode fields to one and then strip whitespace
        for index, row in df_inner.iterrows():
            bcode = row.keyword_cust_no + row.ic_filler_1 + row.ic_filler_2
            df_inner.loc[index, 'barcode'] = bcode.strip()

        print('Sanity', list(df_inner.columns))
        print(df_inner)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter('Out/Barcode_View_Test.xlsx', engine='xlsxwriter')

        # df.to_excel(writer, sheet_name='Sheet1', index=False)
        df_inner.to_excel(writer, sheet_name='Sheet1')

        writer.save()





    finally:
        conn.close()


BarcodeSearchView()