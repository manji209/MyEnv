# Import pandas
import xlrd
#import numpy as np
import pandas as pd
import datefinder
import datetime
import re
import calendar
from calendar import month_name
from time import strptime


print("My name is Dinh")
# Assign spreadsheet filename to `file`
file = 'master1.xlsm'

# Load spreadsheet
xl = pd.ExcelFile(file)

# Print the sheet names
print(xl.sheet_names)

sheet_list = 'FROZEN'
is_frozen = False

# Load a sheet into a DataFrame by name: df1
#df1 = xl.parse('Frozen-Master')
# print(df1)

# Get input from user.  Header of Month to be processed
month = input("Please enter the Month to be processed: ")

default_date = ''


def init_empty_df():
    df = pd.DataFrame(columns=[month, 'ITEM CODE', 'BRAND', 'PRODUCT DESCRIPTION', 'PACKAGING SPECS',
                               '@IMAGES','INCOMING', 'NOTE-1', 'NOTE-2'])
    return df


# Sort selected column, extract necessary column values and return new dataframe
def new_sorted_df(by_month):
    global df1
    # Sort selected column
    df1.sort_values(by_month, inplace=True)
    # Reset the Row index from 0
    df1 = df1.reset_index(drop=True)
    # Add additional columns
    df1['@IMAGES'] = ""
    df1['INCOMING'] = ""
    df1['NOTE-1'] = ""
    df1['NOTE-2'] = ""

    # start from the first row
    num_item = 0

    # Determine end of item in column
    for item in df1[month]:
        num_item += 1
        if pd.isnull(item):
            break
    print("Number of Items: ", num_item)



    # Retrieve the items of the column into a new dataframe
    df_new = df1.loc[0:num_item - 2,
              [month, 'ITEM CODE', 'BRAND', 'PRODUCT DESCRIPTION', 'PACKAGING SPECS', '@IMAGES',
               'INCOMING', 'NOTE-1', 'NOTE-2']]

    return df_new

'''
df_date = new_sorted_df(month)
print(df_date)
# Initialize empty dataframe for additional rows for extra dates
df_append = init_empty_df()
'''

# Go through dates and duplicate entry if multiple dates found
def find_append_dates(sorted_df):
    global df_append
    global df_date
    row_item = 0
    for item in sorted_df[month]:
        #total_item += 1
        date_string = item.replace(',', 'a')
        # Get Note1 info
        note_one = get_note_one(date_string)
        if is_frozen:
            note_two = 'FROZEN'
        else:
            note_two = ''
        #note_two = 'Noted'
        # Get list of dates
        incoming_dates = get_list_dates(date_string)
        num_dates = len(incoming_dates)
        if num_dates == 0:
            incoming_date = get_default_date()
            enter_data(row_item, incoming_date, note_one, note_two)
        elif num_dates == 1:
            incoming_date = incoming_dates[0]
            enter_data(row_item, incoming_date, note_one, note_two)
        else:
            df_copy = df_date.loc[row_item]
            enter_append_data(df_copy, incoming_dates, note_one, note_two)
            #enter_append_data(df_copy, incoming_dates, note_one, note_two)

        row_item += 1


# Enter additional row data for sorted dataframe if one or less dates found
def enter_data(row_num, date, note1, note2):
    global df_date
    df_date.loc[row_num, 'INCOMING'] = date
    df_date.loc[row_num, 'NOTE-1'] = note1
    df_date.loc[row_num, 'NOTE-2'] = note2

# Enter addition row data for appended dataframe if more thane one dates found
def enter_append_data(copy, date_list, note1, note2):
#def enter_append_data(copy, date_list, note1, note2):
    global df_append
    copy2 = copy
    copy2['NOTE-1'] = note1
    copy2['NOTE-2'] = note2
    #for item in date_list:
    for x in range(1, len(date_list)):
        copy2['INCOMING'] = date_list[x]
        df_append = df_append.append(copy, ignore_index=True)


# Return string item for NOTE-1.  Either *JA*, *NEW*, or  null
def get_note_one(note_string):
    result_new = note_string.find('NEW')
    result_ja = note_string.find('JA')
    note_one = ''
    if result_new == 0:
        note_one = '*NEW*'
        return note_one
    elif result_ja == 0:
        note_one = '*JA*'
        return note_one
    else:
        return note_one


'''
# Check if item is FROZEN
def is_frozen(note_string):
    result_frozen = note_string.find('FROZEN')
    if result_frozen >= 0:
        return True
'''

# Return number of dates found in a string
def find_dates(date_string):
    matches = datefinder.find_dates(date_string, strict=True)
    num_match = 0
    for match in matches:
        num_match += 1
    return num_match


# Return list of dates found
def get_list_dates(date_string):
    list_dates = []
    matches = datefinder.find_dates(date_string, strict=True)
    for match in matches:
        list_dates.append(datetime.date.strftime(match, "%m/%d/%y"))
    return list_dates


def get_default_date():
    m = get_month_name(month)
    d = datetime.date.today()
    month_number = strptime(m, '%B').tm_mon
    date_format = datetime.date(d.year, month_number, calendar.monthrange(d.year, d.month)[-1])
    return datetime.date.strftime(date_format, "%m/%d/%y")


def get_month_name(s):
    pattern = "|".join(month_name[1:])
    return re.search(pattern, s, re.IGNORECASE).group(0)


# Load a sheet into a DataFrame by name: df1
df1 = xl.parse('Grocery-Master')
df_date = new_sorted_df(month)
print(df_date)
# Initialize empty dataframe for additional rows for extra dates
df_append = init_empty_df()

# Initialize empty dataframe to store both Grocery and Frozen
df_combined = init_empty_df()


find_append_dates(df_date)
print(df_date)
print(df_append)
print(df_append.loc[:, 'INCOMING'])
print(df_date.loc[:, 'INCOMING'])
#print(df_date.columns.values)

df_combined = df_combined.append(df_date, ignore_index=True)
print('This is combined dataframe: ')
print(df_combined)

is_frozen = True
# Load a sheet into a DataFrame by name: df1
df1 = xl.parse('Frozen-Master')
df_date = new_sorted_df(month)
print(df_date)

find_append_dates(df_date)
print(df_date)
print(df_append)
print(df_append.loc[:, 'INCOMING'])
print(df_date.loc[:, 'INCOMING'])
#print(df_date.columns.values)

df_combined = df_combined.append(df_date,ignore_index=True)
print('This is combined dataframe 2: ')
print(df_combined)

df_combined = df_combined.append(df_append,ignore_index=True)
print('This is combined dataframe 3: ')
print(df_combined)
'''
# Load a sheet into a DataFrame by name: df1
df1 = xl.parse('Frozen-Master')
df_frozen = new_sorted_df(month)
print(df_frozen)
# Initialize empty dataframe for additional rows for extra dates
#df_append = init_empty_df()



find_append_dates(df_frozen)
print(df_frozen)
print(df_append)
print(df_append.loc[:, 'INCOMING'])
print(df_frozen.loc[:, 'INCOMING'])


#df2 = df1.iloc[0:5, [0,2,4]]
#print(df2)


#df2 = xl.parse('Frozen-Master')
'''
'''  
date_text = 'NEW7/01/18-MT073, 7/02/18-FRZTP006, 7/02/18-FRZTP006'
date_text = date_text.replace(',', 'a')
print(date_text)
#date_text = 'NEW Your bill is due 7/02/18-FRZTP006 7/01/18-MT073 Your appointment is on July 14th, 2016. Your bill is due 05/05/2016'
matches = datefinder.find_dates(date_text)

num_match = 0

for match in matches:
    print(match)
    num_match += 1
    date_format = datetime.date.strftime(match, "%m/%d/%y")
    print(date_format)
    #print(datetime.date.strftime(match, "%m/%d/%y"))

print('Number of dates: ', num_match)

result = date_text.find('NEW')
print('Word found :', result)


'''
'''
date_text = 'JA7/01/18-MT073, NEW7/02/18-FRZTP006, 7/02/18-FRZTP006'
date_text = date_text.replace(',', 'a')
#result = date_text.find('NEW')
result = get_note_one(date_text)
print('Word found :', result)
print(get_default_date())
month1 = 'July2nd'
print(get_month_name(month1))
'''