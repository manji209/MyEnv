import pyodbc
import xlrd
import socket
import time
import pandas as pd
import numpy as np

book = xlrd.open_workbook("Data/SAP_FRZ_ITEM_LABEL_FULL.xlsx")
sheet = book.sheet_by_name("Sheet1")

#Printer Info

host = "192.168.1.249"
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
    ^LL406
    ^LS0
    ^FT0,133^A0N,118,119^FB812,1,30,C^FH\^CI28^FD{item_no}^FS^CI27
    ^FT582,403^BQN,2,7
    ^FH\^FDLA,{qr}^FS
    ^FO8,163^GB796,0,2^FS
    ^FT19,239^A0N,34,33^FH\^CI28^FD{desc1}^FS^CI27
    ^FT19,283^A0N,28,28^FH\^CI28^FD{desc2}^FS^CI27
    ^FO8,305^GB549,0,2^FS
    ^FT19,339^A0N,28,28^FH\^CI28^FDExp Date:^FS^CI27
    ^FT19,385^A0N,28,28^FH\^CI28^FDBatch#^FS^CI27
    ^FT135,339^A0N,28,28^FH\^CI28^FD{exp_date}^FS^CI27
    ^FT135,385^A0N,28,28^FH\^CI28^FD{batch}^FS^CI27
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


