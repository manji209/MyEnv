import pyodbc
import xlrd
import socket
import time
import pandas as pd
import numpy as np

book = xlrd.open_workbook("Data/SAP_FRZ_ITEM_LABEL_FULL.xlsx")
sheet = book.sheet_by_name("Sheet1")

#Printer Info

host = "192.168.1.248"
port = 9100

for r in range(1, sheet.nrows):
    mysocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    item_no = sheet.cell(r,0).value
    desc1 = sheet.cell(r,1).value
    desc2 = sheet.cell(r, 2).value
    batch = sheet.cell(r,4).value
    qr = sheet.cell(r,5).value
    exp_date = sheet.cell(r, 6).value

    item_template = f"""^XA
    ~TA000
    ~JSN
    ^LT0
    ^MNW
    ^MTT
    ^PON
    ^PMN
    ^LH0,0
    ^JMA
    ^PR8,8
    ~SD15
    ^JUS
    ^LRN
    ^CI27
    ^PA0,1,1,0
    ^XZ
    ^XA
    ^MMT
    ^PW812
    ^LL1218
    ^LS0
    ^FT232,1216^A0B,186,185^FB1216,1,48,C^FH\^CI28^FD{item_no}^FS^CI27
    ^FT401,1175^A0B,68,68^FH\^CI28^FD{desc1}^FS^CI27
    ^FT496,1175^A0B,56,56^FH\^CI28^FD{desc2}^FS^CI27
    ^FO309,58^GB0,1145,2^FS
    ^FO539,372^GB0,827,2^FS
    ^FT631,1180^A0B,62,61^FH\^CI28^FDExp Date:^FS^CI27
    ^FT738,1180^A0B,62,61^FH\^CI28^FDBatch#^FS^CI27
    ^FT637,925^A0B,68,68^FH\^CI28^FD{exp_date}^FS^CI27
    ^FT738,987^A0B,62,61^FH\^CI28^FD{batch}^FS^CI27
    ^FT525,383^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^PQ1,0,1,Y
    ^XZ
    """

    template_bytes = bytes(item_template, 'utf8')

    try:
        result = mysocket.connect((host, port))
        print(result)
        mysocket.send(template_bytes)
        mysocket.close()
        time.sleep(1)
    except Exception as e:
        print("something's wrong with %s:%d. Exception is %s" % (host, port, e))


