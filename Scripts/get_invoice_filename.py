
import os
import pandas as pd

def list_files(dir):
    r = []
    for root, dirs, files in os.walk(dir):
        for name in files:
            # Strip the file extension
            r.append(os.path.splitext(name)[0])
    return r

# master_list = []
pics = []

d = ["D:/PycharmProjects/luckyPORTAL/luckyPORTAL/media/invoice"]

# Traverse each directory to get filenames
for i in range(0, len(d)):
    pics = pics + list_files(d[i])

# Remove duplicates
pics = list(set(pics))

# Sort
pics.sort()

df = pd.DataFrame(pics)
df.to_csv('../OUT/invoice.csv', index=False)

print(pics)