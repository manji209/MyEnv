import pyodbc
import xlrd
import pandas as pd
import numpy as np

# Open the workbook and define the worksheet
# book = xlrd.open_workbook("Data/import_order_entery_template.xlsx")

template_file = 'Import/LineItem_Import_Template_109792.xlsx'


# Commit each item.  First pivot Items by Item # and add quantity.  Convert Sheet to Dataframe then Pivot
df_commit = pd.read_excel(template_file)
# Pivot df_commit
df_commit_piv = df_commit.pivot_table(values=['qty_ord'], index=['item_no'], aggfunc={'qty_ord':np.sum})
print(df_commit_piv)
print(df_commit_piv.columns.values)

for i, row in df_commit_piv.iterrows():
    print(row['qty_ord'], i)


