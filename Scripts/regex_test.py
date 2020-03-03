import re
import datefinder

result = []
temp_data = []


sample = "[BMAI020 - A/N] (ETD: 11/1 - ETA: 11/19) LA LUCKY"
sample2 = "[TP4952/19 - A/N] (ETD: 10/9 - ETA: 11/1) LA LUCKY"
sample3 = "[TNT002 - A/N] (ETD: 9/29 - ETA: 10/16) LA LUCKY"
sample4 = "[TP58/19 - DOCUMENTS] (ETD: 11/13 - 12/6) LA LUCKY"
sample5 =""

if "ETD" in sample5:
    print("ok")

first = re.search("[^\s]+", sample4)

temp_data.append(first.group()[1:])

print(first.group()[1:])

matches = datefinder.find_dates(sample)

for match in matches:
    print(match.date().strftime("%m/%d"))
    temp_data.append(match.date().strftime("%m/%d"))

sub_string = sample3.partition("]")[2]
print (temp_data)
print (sub_string)