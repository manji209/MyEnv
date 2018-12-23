import csv
import datefinder
import datetime

datafile = open('FullInventoryList.csv', 'r')
line_reader = list(csv.reader(datafile))
line_list = []
product_list = []
# Initialize list with headers
product_list.append(['SKU', 'DESC-1', 'DESC-2', 'IN STOCK'])


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


'''
def check_num(s):
    return isinstance(s, float)

def get_qty(string_list):
    for s in string_list:
        temp = s.replace("-", "")
        if isinstance(temp, int):
            return s
    return 'NA'
    
'''

def get_qty(string_list):
    for s in string_list:
        try:
            int(s.replace("-", ""))
            return s
        except ValueError:
            continue

    return '0'

# Go through line_reader and find SKU that starts with 2 capital letters.  Add the following 2 lines into a list
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


# Go through the list_products to extract all necessary info.  The list is paired so iterate by 2.
for x in range(0, len(list_products), 2):
    # Get Item number then pop it off the top of list
    sku = list_products[x][0]

    list_products[x].pop(0)
    desc_one = ''
    desc_two = ''

    double_found = False
    # Get Description 1 by going through string until float is found.  Next item is Quantity in Hand
    for s in list_products[x]:
        if check_num(s.replace("-", "").replace(",", "")) and s.find('.') >= 0:
            double_found = True
        elif check_num(s.replace("-", "").replace(",", "")) and double_found:
            qty = s
            break
        elif s == '':
            continue
        else:
            desc_one = desc_one + " " + s


    # Get Description 2 by going through string until float is found
    for s in list_products[x+1]:
        if check_num(s.replace("-", "").replace(",", "")) and s.find('.') >= 0:
            break
        elif s == '':
            continue
        else:
            desc_two = desc_two + " " + s

    '''
    
    # Get Description 2 by going through string until float is found
    for s in list_products[x+1]:
        if check_num(list_products[x+1][0]):
            break
        elif list_products[x + 1][0] == '':
            list_products[x + 1].pop(0)
        else:
            desc_two = desc_two + " " + list_products[x + 1][0]
            list_products[x + 1].pop(0)

    '''

    # Move the negative sign to the left side if negative value found
    if qty.find('-') >= 0:
        qty = '-' + qty.replace("-", "")


    # Add each item to a new organized list
    curr_list = [sku, desc_one, desc_two, qty]
    product_list.append(curr_list)


item_num = 0
for item in product_list:
    item_num += 1
    print(item)


print('Total items: ', item_num)

# Save to CSV file
with open('InventoryListItems2.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerows(product_list)


if isinstance(float('130'), float):
    print('Number float found: ')
else:
    print('Number not found: ')

