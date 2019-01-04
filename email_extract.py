import csv
import xlrd
import pandas as pd

datafile = open('Customer_Email.csv', 'r')
line_reader = list(csv.reader(datafile))
email_list = []

for l in line_reader:
    for item in l:
        if "@" in item:
            email_list.append(item)

pd.DataFrame(email_list).to_excel('out_email.xlsx', header=False, index=False)

print(email_list)

