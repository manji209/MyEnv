import xlrd
import pandas as pd
import numpy as np


file_name = "Data/CUSTOMER_phone.xlsx"
df = pd.read_excel(file_name, sheet_name='Sheet1')
df['phone'] = df['phone'].astype(str)
df['fax'] = df['fax'].astype(str)

# Calculate the totals for Weekly Total
for index, row in df.iterrows():
    df.at[index, 'phone'] = str(row['phone']).replace('(', '').replace(')', '-')
    # Add dashes to fax number
    if len(row['fax']) == 10:
        d = row['fax']
        d = '-'.join([d[:3], d[3:6], d[6:]])
        df.at[index, 'fax'] = d

for index, row in df.iterrows():
    if len(row['fax']) == 8:
        area = row['phone'][:3]
        df.at[index, 'fax'] = area + '-' + row['fax']




# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/phone_fixxed.xlsx', engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.save()


phone = '(405)624-6668'
print(phone.replace('(', '').replace(')', '-'))

#print(df)