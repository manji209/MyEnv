import xlrd
import pandas as pd

# Assign spreadsheet filename to `file`
file = 'inventory_list.xlsx'

# Load spreadsheet
df = pd.ExcelFile(file)
df1 = xl.parse(category)
#df = pd.DataFrame(columns=['ITEM ID', 'DESCRIPTION', 'IN STOCK'])

