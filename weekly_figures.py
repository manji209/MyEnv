import xlrd
import pandas as pd
import numpy as np
import datetime



# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Weekly_Figure.xlsx', engine='xlsxwriter')

# Get file name from user.
# fname = input("Please enter the File name to be processed: ")
fname = 'Week_1.xlsx'
s_name = "Sheet1"

def run_report(writer):

    dfs = pd.read_excel("Data/" + fname, sheet_name=s_name)

    invoice_df = (dfs.groupby(['SALES REP', pd.Grouper(freq='D', key='INVOICE DATE')])
                  ['INVOICE #']
                  .nunique()
                  .unstack(fill_value=0))

    item_df = (dfs.groupby(['SALES REP', pd.Grouper(freq='D', key='INVOICE DATE')])
                  ['ITEM #']
                  .nunique()
                  .unstack(fill_value=0))



    invoice_df.columns = invoice_df.columns.day_name()
    item_df.columns = item_df.columns.day_name()

    # Add Totals Column
    invoice_df['Weekly Total'] = ""
    item_df['Weekly Total'] = ""

    # Calculate the totals for Weekly Total
    for index, row in invoice_df.iterrows():
        invoice_df.loc[index, 'Weekly Total'] = row.Monday + row.Tuesday + row.Wednesday + row.Thursday + row.Friday

    # Calculate the totals for Weekly Total
    for index, row in item_df.iterrows():
        item_df.loc[index, 'Weekly Total'] = row.Monday + row.Tuesday + row.Wednesday + row.Thursday + row.Friday



    '''

    # Convert the dataframe to an XlsxWriter Excel object.
    #pivot_dfs.to_excel(writer, sheet_name='Sales_Figures', index=True)
    qty_df.to_excel(writer, sheet_name='Qty Sum', index=True)
    sales_df.to_excel(writer, sheet_name='Sales Sum', index=True)


    '''

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    # writer = pd.ExcelWriter('3_Sales_Figures_2017.xlsx', engine='xlsxwriter')

    # Write the dataframes into the same sheet
    invoice_df.to_excel(writer, sheet_name=s_name, startrow=4, startcol=0, index=True)
    item_df.to_excel(writer, sheet_name=s_name, startrow=18, startcol=0, index=True)

    '''
    worksheet = writer.sheets[s_name]

    # ADD the Header for each Dataframe and include the Yearly Totals column
    worksheet.write('A4', 'Daily Invoice Totals')
    worksheet.write('G5', 'Weekly Totals')

    # Fill in the Yearly Totals with the SUM formula
    for row_num in range(6, 12):
        worksheet.write_formula(row_num - 1, 11, '=SUM($B%d:$M%d)' % (row_num, row_num))

    # ADD the Header for each Dataframe and include the Yearly Totals column
    worksheet.write('A18', 'Accounts Totals')
    worksheet.write('N19', 'Yearly Totals')

    # Fill in the Yearly Totals with the SUM formula
    for row_num in range(20, 28):
        worksheet.write_formula(row_num - 1, 13, '=SUM($B%d:$M%d)' % (row_num, row_num))

    # ADD the Header for each Dataframe and include the Yearly Totals column
    worksheet.write('A32', 'Sales Totals')
    worksheet.write('N33', 'Yearly Totals')

    # Fill in the Yearly Totals with the SUM formula
    for row_num in range(34, 42):
        worksheet.write_formula(row_num - 1, 13, '=SUM($B%d:$M%d)' % (row_num, row_num))

    '''
# print(dfs.columns.values)

run_report(writer)

writer.save()