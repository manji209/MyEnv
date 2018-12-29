import csv
import xlrd
import pandas as pd
import numpy as np
import datefinder
import datetime
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from openpyxl import Workbook

# Based on history.py
#12/15/18
fname = 'Inv_Hist_by_Inv_full_dec.csv'
datafile = open(fname, 'r')
line_reader = list(csv.reader(datafile))

# Column names for dataframe
labels = ['ORDER #', 'INVOICE #', 'DATE', 'CUSTOMER ID', 'SALES REP', 'SKU #', 'DESC-1', 'DESC-2',
           'QUANTITY', 'UNIT PRICE', 'CREDIT MEMO']

pages = []

invoice_history = []

count = 0
sub_count = 0
line_item = 0


class Page:
    def __init__(self):
        self.list_items = []


class Product:
    def __init__(self):
        self.sku = ''
        self.description = ''
        self.description2 = ''
        self.quantity = ''
        self.unit_price = ''


class Order:
    def __init__(self):
        self.invoice_num = ''
        self.order_num = ''
        self.customer_num = ''
        self.date = ''
        self.sales_rep = ''
        self.credit_memo = ''


def check_num(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


def check_double(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

# Return 2 if double found, 1 if integer found, 0 if not a number
def find_num(s):
    if check_double(s.replace("-", "").replace(",", "")) and s.find('.') >= 0:
        return 2
    elif check_double(s.replace("-", "").replace(",", "")):
        return 1
    else:
        return 0

#Returns the Name associated with the Sales Rep #
def get_name(num):
    switcher = {
            5: "Chi Nong",
            7: "Michelle Nong",
            11: "George Nguyen",
            12: "Terry Nguyen",
            14: "Larry Nguyen",
            15: "Linh Ung",
            16: "Pierre Bach",
            17: "Kenny Nguyen",
            18: "Phat Tran",
            19: "Minh Bui",
            20: "Sang Tran"
        }
    return switcher.get(num, "NA#")



# Return the date string if found
def find_dates(date_string):
    temp_string = date_string.replace(":", "")
    matches = datefinder.find_dates(temp_string, strict=True)
    for match in matches:
        return match

    return " "


def find_quantity(string):
    hyphen_found = string.find('-', len(string)-1, len(string))
    test_string = string
    if hyphen_found >= 0:
        return check_num(test_string.replace("-", ""))
    elif check_num(string) or check_double(string):
        return True
    else:
        return False


def find_unit_price(string):
    hyphen_found = string.find('-', len(string)-1, len(string))
    test_string = string
    if hyphen_found >= 0:
        return check_double(test_string.replace("-", ""))
    else:
        return check_double(string)


# Go thru first column of dataframe to identify non-duplicates.  Return list of non-duplicates in reverse order
def del_non_dupe(pivot_df):
    print('Pivot index size: ', len(pivot_df.index.values))
    pop_list = []
    found = False
    for i in range(0, len(pivot_df.index.values) - 1):
        a = pivot_df.index[i][0]
        b = pivot_df.index[i + 1][0]
        if a == b:
            found = True
            continue
        elif found:
            found = False
            continue
        else:
            pop_list.insert(0, i)

    # Delete non dupe rows from list above. In reverse order
    pivot_df.drop(pivot_df.index[pop_list], inplace=True)


# Go through each page and add relevant lines into a temp list then empty p.list_items.  Put all items
# in temp list back into p.list_items.
def remove_lines(p):
    item_found = False
    temp_list = []
    for i in range(0, len(p.list_items)):
        if  p.list_items[i][0] == 'Invoice' or p.list_items[i][0] == 'Order':
            temp_list.append(p.list_items[i])
            #p.list_items.pop(0)
        elif check_num(p.list_items[i][0]) and len(p.list_items[i]) > 1:
            if p.list_items[i][1] != "":
                temp_list.append(p.list_items[i])
                item_found = True
                #p.list_items.pop(0)
            else:
                continue
        elif item_found:
            temp_list.append(p.list_items[i])
            item_found = False
            #p.list_items.pop(0)
        else:
            #p.list_items.pop(0)
            continue

    # Put back the saved line on top of page after the unnecessary lines have been removed
    del p.list_items[:]
    for i in temp_list:
        p.list_items.append(i)

# format the Unit Price Field
def format_currency(x):
    return "${:.2f}".format((x / 10))


# Convert Unit Price to Currency field in Excel
def set_currency(writer, sheet_name, column):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    money_fmt = workbook.add_format({'num_format': '$#,##0.00', 'bold': True})
    worksheet.set_column(column, 12, money_fmt)


def extract_info(p):
    # Credit Memo flag to False until credit memo page is found
    cmemo = False
    # Sometimes the Customer # is blank so test for it.  If it is then Customer number is next string
    customer_blank = False
    global line_item

    # If page has no info then continue
    if len(p.list_items) == 2:
        return False

    info = Order()



    # Go through Invoice line to extract the appropriate data
    for string in p.list_items[0]:

        if customer_blank:
            info.customer_num = string
            break

        if check_num(string):
            info.invoice_num = string
        elif find_dates(string) != " ":
            info.date = find_dates(string)
        elif string.find('Customer:') >= 0:
            info.customer_num = string.rsplit("Customer:")[-1].strip()
            if len(info.customer_num):
                break
            else:
                customer_blank = True
        else:
            continue

    if check_num(p.list_items[0][-1]):
        info.sales_rep = get_name(int(p.list_items[0][-1]))
    elif check_num(p.list_items[0][-2]):
        # Check if the page is a Credit Memo
        info.sales_rep = get_name(int(p.list_items[0][-2]))
        info.credit_memo = 'CREDIT MEMO'
        cmemo = True

    elif check_num(p.list_items[0][-3]):
    # Check if the page is a Credit Memo
        info.sales_rep = get_name(int(p.list_items[0][-3]))
        info.credit_memo = 'CREDIT MEMO'
        cmemo = True
    else:
        # Check if the page is a Credit Memo
        info.sales_rep = "NA"
        info.credit_memo = 'CREDIT MEMO'
        cmemo = True

    '''
    # Get Sales rep number then convert to name
    if check_num(p.list_items[0][-1]):
        info.sales_rep = get_name(int(p.list_items[0][-1]))
    else:
        info.sales_rep = get_name(int(p.list_items[0][-2]))
        info.credit_memo = 'CREDIT MEMO'
        cmemo = True
    
    elif check_num(p.list_items[0][-2]):
        # Check if the page is a Credit Memo
        info.sales_rep = get_name(int(p.list_items[0][-2]))
        info.credit_memo = 'CREDIT MEMO'
        cmemo = True
    elif check_num(p.list_items[0][-3]):
        # Check if the page is a Credit Memo
        info.sales_rep = get_name(int(p.list_items[0][-3]))
        info.credit_memo = 'CREDIT MEMO'
        cmemo = True
    else:
        # Check if the page is a Credit Memo
        info.sales_rep = "NA"
        info.credit_memo = 'CREDIT MEMO'
        cmemo = True
    '''
    # Remove line with Invoice number
    p.list_items.pop(0)

    # Process Order line to get Order number
    info.order_num = p.list_items[0][2]
    # Remove line with Order number
    p.list_items.pop(0)

    for x in range(0, len(p.list_items), 2):
        item = Product()
        # Check if line contains order info by seeing if first element is a sequence number
        if check_num(p.list_items[x][0]):
            if len(p.list_items[x]) == 1 or p.list_items[x][0] == '' or p.list_items[x][1] == '':
                continue
            else:
                item.sku = p.list_items[x][1]
                p.list_items[x].pop(0)
                p.list_items[x].pop(0)

                for i in range(0, len(p.list_items[x])):
                    if p.list_items[x][0] == "":
                        p.list_items[x].pop(0)
                        continue
                    elif find_quantity(p.list_items[x][0]):
                        # item.quantity = '-' + line[0].replace("-", "")
                        item.quantity = p.list_items[x][0]

                        # Move the negative sign to the left side if negative value found
                        if item.quantity.find('-') >= 0:
                            temp_line = '-' + item.quantity.replace("-", "")
                            item.quantity = temp_line

                        # Check if item is a credit memo.  If so, make the quantity a negative number
                        if cmemo and item.quantity.find('-') == -1:
                            item.quantity = '-' + item.quantity

                        p.list_items[x].pop(0)
                        break

                    else:
                        item.description = item.description + p.list_items[x][0]
                        p.list_items[x].pop(0)

                # Go through rest of list to find Unit Price.
                for i in range(0, len(p.list_items[x])):
                    if check_num(p.list_items[x][i]):
                        # if next number found is a integer than set it to the quantity
                        item.quantity = p.list_items[x][i]
                    elif find_unit_price(p.list_items[x][i]):

                        item.unit_price = p.list_items[x][i]
                        # Move the negative sign to the left side if negative value found
                        if item.unit_price.find('-') >= 0:
                            temp_line = '-' + item.unit_price.replace("-", "")
                            item.unit_price = temp_line
                        break

        else:
            continue

        next_seq = True

        try:
            for w in p.list_items[x+1]:
                # Check if next line is also a sequence number
                if next_seq and check_num(w):
                    item.description2 = 'NOT Available'
                    break
                else:
                    next_seq = False

                if find_num(w) == 2:
                    break
                elif w == "":
                    continue
                else:
                    item.description2 = item.description2 + " " + w
        except IndexError:
            item.description2 = 'NOT Available'

        if item.sku != '':
            line_item += 1
            invoice_history.append([info.order_num, info.invoice_num, info.date, info.customer_num, info.sales_rep,
                                    item.sku, item.description.strip(), item.description2.strip(), item.quantity, item.unit_price,
                                    info.credit_memo])

    return True


'''
    # Get account info first then go line by line to get product info
    for line in p.list_items:

        item = Product()
        # Check if line contains order info by seeing if first element is a sequence number
        if check_num(line[0]):
            if len(line) == 1 or line[0] == '' or line[1] == '':
                continue
            else:
                item.sku = line[1]
                line.pop(0)
                line.pop(0)

                for i in range(0, len(line)):
                    if line[0] == "":
                        line.pop(0)
                        continue
                    elif find_quantity(line[0]):
                        # item.quantity = '-' + line[0].replace("-", "")
                        item.quantity = line[0]

                        # Move the negative sign to the left side if negative value found
                        if item.quantity.find('-') >= 0:
                            temp_line = '-' + item.quantity.replace("-", "")
                            item.quantity = temp_line

                        # Check if item is a credit memo.  If so, make the quantity a negative number
                        if cmemo and item.quantity.find('-') == -1:
                            item.quantity = '-' + item.quantity

                        line.pop(0)
                        break

                    else:
                        item.description = item.description + line[0]
                        line.pop(0)

                # Go through rest of list to find Unit Price.
                for i in range(0, len(line)):
                    if check_num(line[i]):
                        # if next number found is a integer than set it to the quantity
                        item.quantity = line[i]
                    elif find_unit_price(line[i]):

                        item.unit_price = line[i]
                        # Move the negative sign to the left side if negative value found
                        if item.unit_price.find('-') >= 0:
                            temp_line = '-' + item.unit_price.replace("-", "")
                            item.unit_price = temp_line
                        break

        else:
            item.description2 = 'Test'


        if item.sku != '':
            line_item += 1
            invoice_history.append([info.order_num, info.invoice_num, info.date, info.customer_num, info.sales_rep,
                                    item.sku, item.description, item.description2, item.quantity, item.unit_price, info.credit_memo])


'''

# Break down the CSV file into list of pages
for item in line_reader:
    sub_count += 1
    if len(item) > 0 and item[0].find('Date') >= 0:
        #print('Found')
        p = Page()
        p.list_items.append(item)
    elif len(item) > 0:
        #print('line')
        p.list_items.append(item)
    else:
        #print('Not Found')
        remove_lines(p)
        if extract_info(p):
            pages.append(p)
        else:
            continue

    # Add last page
    if sub_count == len(line_reader):
        remove_lines(p)
        if extract_info(p):
            pages.append(p)
        else:
            continue


print('Pages Created: ')

'''
# Break down the CSV file into list of pages
for item in line_reader:
    sub_count += 1
    if len(item) > 0 and item[0].find('Date') >= 0:
        #print('Found')
        p = Page()
        p.list_items.append(item)
    elif len(item) > 0:
        #print('line')
        p.list_items.append(item)
    else:
        #print('Not Found')
        remove_lines(p)
        extract_info(p)
        pages.append(p)

    # Add last page
    if sub_count == len(line_reader):
        remove_lines(p)
        extract_info(p)
        pages.append(p)


print('Pages Created: ')

'''

'''
for p in pages:
    print(*p.list_items, sep="\n")

with open('out_file.txt', 'w') as f:
    for p in pages:
        for ln in p.list_items:
            f.write("%s\n" % ln)

'''



df = pd.DataFrame(invoice_history, columns=labels)
df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce').dt.date

larry_df = df.loc[df['SALES REP'] == 'Larry Nguyen']

larry_cust = larry_df['CUSTOMER ID']


print(df)
print(larry_df)
print(larry_cust)


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('out_'+fname+'.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Invoice Info', index=False)
#df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce').dt.date




test_pivot_df = pd.pivot_table(df, index=['CUSTOMER ID', 'SALES REP', 'DATE'], aggfunc='first')
test_pivot_df.drop(columns=['ORDER #', 'INVOICE #'], inplace=True)

test_larry_df = pd.pivot_table(larry_df, index=['CUSTOMER ID', 'SKU #', 'DATE'], aggfunc='first')
test_larry_df.drop(columns=['ORDER #', 'INVOICE #'], inplace=True)

test_pivot_df.to_excel(writer, sheet_name='Customer ID Audit')
larry_df.to_excel(writer, sheet_name='Larry Nguyen')

larry_cust.to_excel(writer, sheet_name='Customer')
test_larry_df.to_excel(writer, sheet_name='Larry Nguyen Pivot')





writer.save()

'''
df = pd.DataFrame(invoice_history, columns=labels)



df['ORDER #'] = pd.to_numeric(df['ORDER #'], errors='coerce')
df['INVOICE #'] = pd.to_numeric(df['INVOICE #'], errors='coerce')
df['QUANTITY'] = pd.to_numeric(df['QUANTITY'], errors='coerce')



#df['UNIT PRICE']= df['UNIT PRICE'].apply(format_currency)
df['UNIT PRICE'] = pd.to_numeric(df['UNIT PRICE'], errors='coerce').map(('${:,.2f}'.format))

df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce').dt.date

print('Dataframe to Numeric: ')

test_pivot_df = pd.pivot_table(df, index=['SKU #', 'DESCRIPTION'], aggfunc='first')
test_pivot_df.drop(columns=['ORDER #', 'QUANTITY'], inplace=True)

test_pivot_df2 = pd.pivot_table(df, index=['SKU #', 'UNIT PRICE', 'DESCRIPTION'], aggfunc='first')
test_pivot_df2.drop(columns=['ORDER #'], inplace=True)



print('Pivot Table Created: ')


# Delete non-repeating row items
del_non_dupe(test_pivot_df)
del_non_dupe(test_pivot_df2)


print('Delete non_dupe: ')



# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pivot_sample7.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Invoice Info', index=False)


test_pivot_df.to_excel(writer, sheet_name='Item Number Audit')
test_pivot_df2.to_excel(writer, sheet_name='Price Audit')
#test_pivot_df3.to_excel(writer, sheet_name='Pivot3')
#test_pivot_df4.to_excel(writer, sheet_name='Pivot4')
#test_pivot_df5.to_excel(writer, sheet_name='Pivot5')
#test_pivot_df6.to_excel(writer, sheet_name='Pivot6')

#set_currency(writer, 'Price Audit', 'B:B')
#set_currency(writer, 'Invoice Info', 'I:I')
#set_currency(writer, 'Item Number Audit', 'H:H')
#set_currency(writer, 'Price Audit', 'B:B')


print('Write to Excel: ')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

#print(invoice_history)
print('Pages Found: ', len(pages))
print('Line Items: ', line_item)
'''
