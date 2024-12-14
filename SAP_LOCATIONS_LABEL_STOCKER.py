import pyodbc
import pandas as pd
import numpy as np
import xlrd
import socket
import time


book = xlrd.open_workbook("Data/MISC_LOCATIONS.xlsx")
sheet = book.sheet_by_name("Misc")

#Printer Info

host = "192.168.1.249"
port = 9100

for r in range(1, sheet.nrows):
    mysocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    #level = int(sheet.cell(r,4).value)
    level = sheet.cell(r, 3).value
    location = sheet.cell(r,4).value
    qr = sheet.cell(r,5).value

    location_template = f"""^XA
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
    ^FT171,399^A0B,141,142^FB399,1,36,C^FH\^CI28^FD{level}^FS^CI27
    ^FT213,339^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^FT491,339^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^FT742,403^A0B,39,38^FB403,1,10,C^FH\^CI28^FD{location}^FS^CI27
    ^PQ1,0,1,Y
    ^XZ
    """

    template_bytes = bytes(location_template, 'utf8')

    try:
        result = mysocket.connect((host, port))
        print(result)
        mysocket.send(template_bytes)
        mysocket.close()
        time.sleep(1)
    except Exception as e:
        print("something's wrong with %s:%d. Exception is %s" % (host, port, e))


