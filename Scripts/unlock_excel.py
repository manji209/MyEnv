import openpyxl
import pandas as pd
import numpy as np
import datetime

wb_book = openpyxl.load_workbook("Out/LineItem_Import_3.xlsx")
#wb_book.protection.disable()
#wb_book.protection.sheet = False
sheet = wb_book.active
sheet.protection.sheet = False
sheet.protection.disable()

rowindex = 1
maxrow = 605
for row in sheet.iter_rows(min_row=1, max_row=605):
    if row[2].value is None:
        sheet.delete_rows(rowindex, (maxrow-rowindex))
        break
    else:
        rowindex += 1


wb_book.save("Out/LineItem_Import_3.xlsx")