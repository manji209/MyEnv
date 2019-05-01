from difflib import SequenceMatcher
import pyodbc
from openpyxl import load_workbook
import datetime
import pandas as pd
import numpy as np


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def get_markup(p1, p2):
    if p1 == float(0):
        return 0
    else:
        return (float(p1) - float(p2))/ float(p2)

def get_name(num):
    switcher = {
            5: "Chi Nong",
            7: "Michelle Nong",
            11: "George Nguyen",
            12: "Terry Nguyen",
            14: "Larry Nguyen",
            15: "Linh Ung",
            16: "Pierre Bach",
            17: "Kenny Nguyen",
            18: "Phat Tran",
            19: "Minh Bui",
            20: "Sang Tran"
        }
    return switcher.get(num, "NA#")


# Load Excel File for Processing
wb = load_workbook('Data/Vendor_Map.xlsx')
sheet = wb['19-1-25']


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT item_no, item_desc_1, item_desc_2, replacement_cost
                FROM dbo.ITMFIL00"""


df = pd.read_sql(sqlstring, conn)
'''
# Prompt for start Date and End Date.  Also prompt for sheet name to process
start_date = input("Input Start Date yyyy-mm-dd:  ")
end_date = input("Input End Date yyyy-mm-dd:  ")
sheet_name = input("Enter Sheet Name:  ")

df = pd.read_sql(sqlstring, conn, params=(start_date, end_date))

# Load Excel Sheet
sheet = wb[sheet_name]
'''

df['item_desc_2'] = df['item_desc_2'].apply(str)

index = 0
none = 0
best = 0
ratio = 0


for row in sheet.iter_rows():
    index +=1
    if none >= 4:
        break

    if (row[5].value != None):
        if index != 1:
            none = 0
            for row_df in df.itertuples(index=True):

                vendor = str(row[4].value) + str(row[6].value)
                lucky = row_df.item_desc_1.rstrip() + str(row_df.replacement_cost)
                #lucky = row_df.item_desc_1.rstrip() + row_df.item_desc_2.rstrip()
                ratio = similar(vendor, lucky)
                #ratio = similar(row[4].value, row_df.item_desc_1.rstrip())
                if ratio > best:
                    best = ratio
                    best_row = row_df

            # Write to Excel Sheet
            row[9].value = best_row.item_no.rstrip()
            row[10].value = best_row.item_desc_1
            row[11].value = best_row.item_desc_2

            print(row, 'ratio: ', ratio, best_row)


            best = 0
    else:
        none += 1


wb.save(filename="Data/Vendor_Map.xlsx")
