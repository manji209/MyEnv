import pyodbc
import xlrd
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()


query = """INSERT INTO [dbo].[STAFIL00] (
            [item_no]
            ,[whs]
            ,[whs_alt]
          ) VALUES (?,?,?)"""

# grab existing row count in the database for validation later
cur.execute("SELECT count(*) FROM dbo.STAFIL00")
before_import = cur.fetchone()

item_no = 'DVH02'
whs = ''
whs_alt = ''
    # Assign values from each row
values = (item_no,whs,whs_alt)

    # Execute sql Query
cur.execute(query, values)



# Commit the transaction
conn.commit()

# If you want to check if all rows are imported
cur.execute("SELECT count(*) FROM dbo.STAFIL00")
result = cur.fetchone()

print((result[0] - before_import[0]))  # should be True

# Close the database connection
conn.close()