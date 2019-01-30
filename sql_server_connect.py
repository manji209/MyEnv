import pyodbc

conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sql = "SELECT top 1000 * FROM dbo.CUSFIL00"
cur.execute(sql)

for row in cur:
    print('row = %r' % (row,))

