import csv
import datefinder
import datetime

datafile = open('FullInventoryList.csv', 'r')
line_reader = list(csv.reader(datafile))
line_list = []
product_list = []
# Initialize list with headers
product_list.append(['SKU', 'DESC-1', 'DESC-2'])


class Item:
    def __init__(self):
        item_number = ''
        description1 = ''
        description2 = ''
        date = ''
        quantity = ''


line_counter = 0
list_products = []

for line in line_reader:
    line_counter += 1
    if len(line) == 0 or line[0] == '':
        continue
    elif line[0][:2].isupper():
        print(line[0])
        print(line)


def check_num(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


#row_count = sum(1 for row in line_reader)
for i in range(0, len(line_reader)):

    if len(line_reader[i]) == 0 or line_reader[i][0] == '':
        continue
    elif line_reader[i][0][:2].isupper():
        list_products.append(line_reader[i])
        list_products.append(line_reader[i + 1])



print('This is the clean list: ')
#print('Row count: ', row_count)
print('Line count: ', line_counter)
for item in list_products:
    print(item)



for x in range(0, len(list_products), 2):
    # Get Item number then pop it off the top of list
    sku = list_products[x][0]

    list_products[x].pop(0)
    desc_one = ''
    desc_two = ''


    # Get Description 1 by going through string until float is found
    for s in list_products[x]:
        if check_num(list_products[x][0]):
            break
        elif list_products[x][0] == '':
            list_products[x].pop(0)
        else:
            desc_one = desc_one + " " + list_products[x][0]
            list_products[x].pop(0)


    # Get Description 2 by going through string until float is found
    for s in list_products[x+1]:
        if check_num(list_products[x+1][0]):
            break
        elif list_products[x + 1][0] == '':
            list_products[x + 1].pop(0)
        else:
            desc_two = desc_two + " " + list_products[x + 1][0]
            list_products[x + 1].pop(0)


    # Add each item to a new organized list
    curr_list = [sku, desc_one, desc_two]
    product_list.append(curr_list)


item_num = 0
for item in product_list:
    item_num += 1
    print(item)


print('Total items: ', item_num)

# Save to CSV file
with open('InventoryListItems.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerows(product_list)

