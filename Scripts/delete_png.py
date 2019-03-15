import os
import csv
import pandas as pd

d = "D:/PNG/"

with open("Data/png.csv", 'r') as f:
    reader = csv.reader(f)
    your_list = list(reader)

flattened_list = []
#flatten the lis
for x in your_list:
    for y in x:
        flattened_list.append(y)


print(flattened_list)

for item in flattened_list:
    if os.path.exists("D:/PNG/" + item + ".PNG"):
        os.remove("D:/PNG/" + item + ".PNG")
    else:
        print("The file does not exist")


