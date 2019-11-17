import pyodbc
import xlrd
import pandas as pd
import numpy as np

# Update inventory from Excel sheet

file_name = "Import/Items_past_year_test_backup08282019.xlsx"
df = pd.read_excel(file_name, sheet_name='w-barcode')
# Remove all NaN and replace with blank space
df = df.fillna('')
# df['barcode'] = df['barcode'].astype('str')

#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()

# Retrieve list of existing barcodes
sqlstring = "SELECT keyword_token FROM dbo.CKEYWF00 WHERE keyword_typ = 'B'"
df_barcodes = pd.read_sql(sqlstring,conn)
df_barcodes['keyword_token'] = df_barcodes['keyword_token'].str.strip()

# Update barcode if code exists
sqlupdate = "UPDATE dbo.CKEYWF00 SET keyword_cust_no=?, ic_filler_1=?, ic_filler_2=' ' WHERE keyword_typ='B' AND keyword_token=?"
#sqlstring = "UPDATE dbo.CKEYWF00 SET keyword_cust_no=? WHERE keyword_token=?"

# Insert new barcode if code does not exist
sqlinsert = """INSERT INTO [dbo].[CKEYWF00] (
            [keyword_typ]
           ,[keyword_token]
           ,[keyword_cust_no]
           ,[ic_filler_1]
           ,[ic_filler_2]
           ,[keyword_typ_alt]
           ) VALUES (?,?,?,?,?,?)"""

found = False

for index, row in df.iterrows():
    for index2 , row2 in df_barcodes.iterrows():
        if row.item_no == row2.keyword_token:
            #df_barcodes.drop(df_barcodes.index[index2])
            found = True
            break

    if found:
        found = False
        values = [row.barcode2, row.barcode3, row.item_no]
        try:
            cur.execute(sqlupdate, values)
        except Exception as e:
            print(e)
    else:
        values2 = ['B', row.item_no, row.barcode2, row.barcode3, ' ', 'B']
        try:
            cur.execute(sqlinsert, values2)
        except Exception as e:
            print(e)



cur.close()
conn.commit()
