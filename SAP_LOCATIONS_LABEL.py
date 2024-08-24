import pyodbc
import xlrd
import socket
import time
import pandas as pd
import numpy as np

book = xlrd.open_workbook("Data/SAP_ROW_B_LOCATIONS.xlsx")
sheet = book.sheet_by_name("Sheet1")

#Printer Info
'''mysocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
host = "192.168.1.248"
port = 9100'''
host = "192.168.1.248"
port = 9100

for r in range(1, sheet.nrows):
    mysocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    wh_name = sheet.cell(r,1).value
    row = sheet.cell(r,2).value
    slot = int(sheet.cell(r,3).value)
    level = int(sheet.cell(r,4).value)
    location = sheet.cell(r,5).value
    qr = sheet.cell(r,7).value
    arrow = sheet.cell(r,8).value

    location_down_template = f"""^XA
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
    ^FO259,15^GB0,1179,2^FS
    ^FO168,913^GB612,0,2^FS
    ^FO168,305^GB612,0,2^FS
    ^FO293,15^GB0,1179,2^FS
    ^FT287,1218^A0B,28,28^FB327,1,7,C^FH\^CI28^FDROW^FS^CI27
    ^FT289,1218^A0B,28,28^FB1214,1,7,C^FH\^CI28^FDSLOT^FS^CI27
    ^FT289,328^A0B,28,28^FB328,1,7,C^FH\^CI28^FDLEVEL^FS^CI27
    ^FT681,1218^A0B,381,380^FB1216,1,97,C^FH\^CI28^FD{slot}^FS^CI27
    ^FT236,1218^A0B,62,61^FB1216,1,16,C^FH\^CI28^FD{location}^FS^CI27
    ^FO166,302^GB0,613,2^FS
    ^FO512,915^GB0,276,2^FS
    ^FO511,31^GB0,276,2^FS
    ^FO524,100^GFA,349,6112,32,:Z64:eJztmLENwjAQRUMoKBkhIzBCRotLJBZBNIyAR/EIKV1YNkVM4ZjjFS6ioPvtixX/f/adkq5b1L8W3TtBaVEAHiU+5QckPmZuBH7O3AJ3Aj9mPgP30gY1AA1g3wGcNAANYN8BDBrAtgEc/iyAx1rPMoAkKQCPwBNxA9wCd0XctWZY74EH4BF4Im6AW+AO+AzcAw/AI/BE3AC3wN3P+i+c1re8n/bfmB/lT/Wj+tP5ofNH57fx/tD984333zX2H+pfjf0z99/rWrfMpQFGA5Dmx1DYrzUW9mtNhf1apf1KNH97ta/21f53qX21r/YFrvab7F+2tf/5kF3/v38DUVVVMQ==:F822
    ^FT467,338^A0B,169,170^FB338,1,43,C^FH\^CI28^FD{level}^FS^CI27
    ^FT469,1218^A0B,169,170^FB330,1,43,C^FH\^CI28^FD{row}^FS^CI27
    ^FT548,1188^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^FT30,1188^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^FT30,299^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^FT148,1218^A0B,68,68^FB1217,1,17,C^FH\^CI28^FD{wh_name}^FS^CI27
    ^FT70,1218^A0B,23,23^FB1217,1,6,C^FH\^CI28^FDWAREHOUSE^FS^CI27
    ^FO43,528^GB0,164,2^FS
    ^FO76,528^GB0,164,2^FS
    ^FO44,530^GB124,0,2^FS
    ^FO44,687^GB124,0,2^FS
    ^PQ1,0,1,Y
    ^XZ
    """

    location_up_template = f"""^XA
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
    ^FO259,15^GB0,1179,2^FS
    ^FO168,913^GB612,0,2^FS
    ^FO168,305^GB612,0,2^FS
    ^FO293,15^GB0,1179,2^FS
    ^FT287,1218^A0B,28,28^FB327,1,7,C^FH\^CI28^FDROW^FS^CI27
    ^FT289,1218^A0B,28,28^FB1214,1,7,C^FH\^CI28^FDSLOT^FS^CI27
    ^FT289,328^A0B,28,28^FB328,1,7,C^FH\^CI28^FDLEVEL^FS^CI27
    ^FT681,1218^A0B,381,380^FB1216,1,97,C^FH\^CI28^FD{slot}^FS^CI27
    ^FT236,1218^A0B,62,61^FB1216,1,16,C^FH\^CI28^FD{location}^FS^CI27
    ^FO166,302^GB0,613,2^FS
    ^FO512,915^GB0,276,2^FS
    ^FO511,31^GB0,276,2^FS
    ^FT467,338^A0B,169,170^FB338,1,43,C^FH\^CI28^FD{level}^FS^CI27
    ^FT469,1218^A0B,169,170^FB330,1,43,C^FH\^CI28^FD{row}^FS^CI27
    ^FT548,1188^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^FT30,1188^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^FT548,305^BQN,2,10
    ^FH\^FDLA,{qr}^FS
    ^FT148,1218^A0B,68,68^FB1217,1,17,C^FH\^CI28^FD{wh_name}^FS^CI27
    ^FT70,1218^A0B,23,23^FB1217,1,6,C^FH\^CI28^FDWAREHOUSE^FS^CI27
    ^FO43,528^GB0,164,2^FS
    ^FO76,528^GB0,164,2^FS
    ^FO44,530^GB124,0,2^FS
    ^FO44,687^GB124,0,2^FS
    ^FO28,72^GFA,381,6112,32,:Z64:eJztmMENwjAMRYM49MgIGSWjpaNwRCwRRmABpI7QY4WgIMWpaCnfPxInkH3kybjvt42lOveqo9Tefa7NQ6oHfFt4B3hT+AnwXeEt4L5wgF0SPKL+KPyG+svfDwCbvumbPuCmb/qmD7jp/4L+FeBJf10SyBbyLvMG8n6hj7gnPJH5kfRDLBzrC8f6Mh/rSz/WF+4JT2R+JP0YZ67oZ67o5/mKfu5X9DP3hCcyP5L+8OV87fpr/LX8avLX7l/V/Q+Ee8KVAKqefyWAqvdPCaDu/Q+Ee8JxAHXnHw6gX5y/98t7nRfn99Cuys0DgOd/ED4iPgWAONtfbP+x/cn2L9vfFoAFYAFYABaABWAB/HEA5fv9Yf7bE7ZsHyg=:3F21
    ^PQ1,0,1,Y
    ^XZ
    """

    if arrow == "DOWN":
        final_template = location_down_template
    else:
        final_template = location_up_template

    template_bytes = bytes(final_template, 'utf8')

    try:
        result = mysocket.connect((host, port))
        print(result)
        mysocket.send(template_bytes)
        mysocket.close()
        time.sleep(1)
    except Exception as e:
        print("something's wrong with %s:%d. Exception is %s" % (host, port, e))


