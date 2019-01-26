import xlrd
import pandas as pd
import numpy as np


file_name = "Data/CUSTOMER_address_correction.xlsx"
df = pd.read_excel(file_name, sheet_name='Sheet1')


class Customer:
    def __init__(self):
        self.city = ""
        self.state = ""
        self.zip = ""



# Calculate the totals for Weekly Total
for index, row in df.iterrows():
    #df[["city", "state", "zipcode"]] = df.addr_2.str.extract('^(.+)\,[ ]([a-zA-z]{2})[ ]([0-9]{5})',expand=True)
    df[["city", "state", "zipcode"]] = df.addr_2.str.extract('^(.+?)\,[ ]{0,1}([a-zA-z]{2})[ ]([0-9]{5})', expand=True)
    #addr = row.addr_2

    #invoice_df.loc[index, 'Weekly Total'] = row.Monday + row.Tuesday + row.Wednesday + row.Thursday + row.Friday


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/address_fixxed.xlsx', engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.save()

#print(df)