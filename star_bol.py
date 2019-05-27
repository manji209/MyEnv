import pyodbc
import xlrd
import pandas as pd
import numpy as np

# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={Advantage StreamlineSQL ODBC};DataDirectory=D:\BOL\data;ServerTypes=1;')
#conn = pyodbc.connect('DRIVER={Advantage StreamlineSQL ODBC};DataDirectory=\\DINHPC\BOL\data\BOLH.adt;DefaultType=Advantage;ServerTypes=3;')
#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()
