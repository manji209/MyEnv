import xlrd
import pandas as pd
import numpy as np
import datetime

file_name = "Data/Inv_History_2018.xlsx"
dfs = pd.read_excel(file_name, sheet_name="2018")


print(dfs.columns.values)
#pivot_dfs = pd.pivot_table(dfs, index=['SALES REP', 'INVOICE DATE'], values=['QTY', 'TOTAL PRICE'], aggfunc={'QTY' : np.sum, 'TOTAL PRICE' : np.sum})

#pivot_dfs = pd.pivot_table(dfs, index=['SALES REP', 'INVOICE DATE'], aggfunc='first')


'''
qty_df = (dfs.groupby(['SALES REP', pd.Grouper(freq='M', key='INVOICE DATE')])
             ['QTY']
             .sum()
             .unstack(fill_value=0))
             
'''

sales_df = (dfs.groupby(['SALES REP', pd.Grouper(freq='M', key='INVOICE DATE')])
             ['TOTAL PRICE']
             .sum()
             .unstack(fill_value=0))


invoice_df = (dfs.groupby(['SALES REP', pd.Grouper(freq='M', key='INVOICE DATE')])
             ['INVOICE #']
             .nunique()
             .unstack(fill_value=0))


account_df = (dfs.groupby(['SALES REP', pd.Grouper(freq='M', key='INVOICE DATE')])
             ['CUSTOMER #']
             .nunique()
             .unstack(fill_value=0))

sales_df.columns = sales_df.columns.date
invoice_df.columns = invoice_df.columns.date
account_df.columns = account_df.columns.date

num_invoices = dfs['INVOICE #'].nunique()
print('Total Inovices: ', num_invoices)
print(sales_df.columns.values)

'''
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Sales_Figures_2018.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
#pivot_dfs.to_excel(writer, sheet_name='Sales_Figures', index=True)
qty_df.to_excel(writer, sheet_name='Qty Sum', index=True)
sales_df.to_excel(writer, sheet_name='Sales Sum', index=True)



writer.save()
'''

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('2_Sales_Figures_2018.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
#pivot_dfs.to_excel(writer, sheet_name='Sales_Figures', index=True)
#qty_df.to_excel(writer, sheet_name='Qty Sum', index=True)
#sales_df.to_excel(writer, sheet_name='Sales Sum', index=True)

#invoice_df.to_excel(writer, sheet_name='Invoice Totals', startrow=3 , startcol=0, index=True)
#account_df.to_excel(writer, sheet_name='Account Totals', startrow=13 , startcol=0, index=True)
#sales_df.to_excel(writer, sheet_name='Sales Sum', startrow=26 , startcol=0, index=True)

invoice_df.to_excel(writer, sheet_name='2018', startrow=4 , startcol=0, index=True)
account_df.to_excel(writer, sheet_name='2018', startrow=18 , startcol=0, index=True)
sales_df.to_excel(writer, sheet_name='2018', startrow=32 , startcol=0, index=True)

writer.save()
#print(dfs.columns.values)