import xlrd
import pandas as pd
import numpy as np
import datetime

yr = 2017

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Sales_Figures_2017_2018.xlsx', engine='xlsxwriter')

def run_report(writer, yr):

    file_name = "Data/Inv_History_" + str(yr) + ".xlsx"
    dfs = pd.read_excel(file_name, sheet_name=str(yr))


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
   
    # Convert the dataframe to an XlsxWriter Excel object.
    #pivot_dfs.to_excel(writer, sheet_name='Sales_Figures', index=True)
    qty_df.to_excel(writer, sheet_name='Qty Sum', index=True)
    sales_df.to_excel(writer, sheet_name='Sales Sum', index=True)
    
 
    '''

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    #writer = pd.ExcelWriter('3_Sales_Figures_2017.xlsx', engine='xlsxwriter')


    # Write the dataframes into the same sheet
    invoice_df.to_excel(writer, sheet_name=str(yr), startrow=4 , startcol=0, index=True)
    account_df.to_excel(writer, sheet_name=str(yr), startrow=18 , startcol=0, index=True)
    sales_df.to_excel(writer, sheet_name=str(yr), startrow=32 , startcol=0, index=True)


    worksheet = writer.sheets[str(yr)]

    # ADD the Header for each Dataframe and include the Yearly Totals column
    worksheet.write('A4', 'Invoice Totals')
    worksheet.write('N5', 'Yearly Totals')

    # Fill in the Yearly Totals with the SUM formula
    for row_num in range(6, 14):
        worksheet.write_formula(row_num-1, 13, '=SUM($B%d:$M%d)' % (row_num, row_num))

    # ADD the Header for each Dataframe and include the Yearly Totals column
    worksheet.write('A18', 'Accounts Totals')
    worksheet.write('N19', 'Yearly Totals')

    # Fill in the Yearly Totals with the SUM formula
    for row_num in range(20, 28):
        worksheet.write_formula(row_num-1, 13, '=SUM($B%d:$M%d)' % (row_num, row_num))

    # ADD the Header for each Dataframe and include the Yearly Totals column
    worksheet.write('A32', 'Sales Totals')
    worksheet.write('N33', 'Yearly Totals')

    # Fill in the Yearly Totals with the SUM formula
    for row_num in range(34, 42):
        worksheet.write_formula(row_num-1, 13, '=SUM($B%d:$M%d)' % (row_num, row_num))




#print(dfs.columns.values)

for i in range(1, 3):
    run_report(writer, yr)
    yr += 1

writer.save()