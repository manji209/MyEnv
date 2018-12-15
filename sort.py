
import xlrd
import pandas as pd
import datefinder
import datetime
import re
import calendar
from calendar import month_name
from time import strptime

# Assign spreadsheet filename to `file`
file = 'master1.xlsm'

# Load spreadsheet
xl = pd.ExcelFile(file)

# Get input from user.  Header of Month to be processed
month = input("Please enter the Month to be processed: ")


class Products:
    def __init__(self, category, df):
        self.category = category
        self.df = df
        self.is_frozen = self.set_frozen(category)

    def enter_data(self, row_num, date, note1, note2, image):
        self.df.loc[row_num, 'INCOMING'] = date
        self.df.loc[row_num, 'NOTE-1'] = note1
        self.df.loc[row_num, 'NOTE-2'] = note2
        self.df.loc[row_num, '@IMAGES'] = image

    def set_frozen(self, category):
        if self.category.lower().find('frozen') >= 0:
            return True
        else:
            return False


def init_empty_df():
    df = pd.DataFrame(columns=[month, 'ITEM CODE', 'BRAND', 'PRODUCT DESCRIPTION', 'PACKAGING SPECS', '@IMAGES',
                               'INCOMING', 'NOTE-1', 'NOTE-2', 'ORIGIN', 'CATEGORY', 'SUBCATEGORY'])
    return df

# Sort selected column, extract necessary column values and return new dataframe
def new_sorted_df(by_month, category):
    #global df1
    df1 = xl.parse(category)
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
    for item in df1[by_month]:
        num_item += 1
        if pd.isnull(item):
            break
    print("Number of Items: ", num_item)



    # Retrieve the items of the column into a new dataframe
    df_new = df1.loc[0:num_item - 2,
              [by_month, 'ITEM CODE', 'BRAND', 'PRODUCT DESCRIPTION', 'PACKAGING SPECS', '@IMAGES',
               'INCOMING', 'NOTE-1', 'NOTE-2', 'ORIGIN', 'CATEGORY', 'SUBCATEGORY']]


    return df_new


# Go through dates and duplicate entry if multiple dates found
def find_append_dates(product):
    row_item = 0
    for item in product.df[month]:

        date_string = item.replace(',', 'a')
        # Get Note1 info
        note_one = get_note_one(date_string)
        # Note_one determines @image
        if note_one != '':
            image = str(product.df.loc[row_item, 'ITEM CODE']) + '.JPG'
        else:
            image = ''

        if product.is_frozen:
            note_two = 'FROZEN'
        else:
            note_two = ''

        # Get list of dates
        incoming_dates = get_list_dates(date_string)
        num_dates = len(incoming_dates)
        if num_dates == 0:
            incoming_date = get_default_date()
            product.enter_data(row_item, incoming_date, note_one, note_two, image)
        elif num_dates == 1:
            incoming_date = incoming_dates[0]
            product.enter_data(row_item, incoming_date, note_one, note_two, image)
        else:
            df_copy = product.df.loc[row_item]
            enter_append_data(df_copy, incoming_dates, note_one, note_two, image)

        row_item += 1


# Enter additional row data for sorted dataframe if one or less dates found
def enter_data(row_num, date, note1, note2):
    global df_date
    df_date.loc[row_num, 'INCOMING'] = date
    df_date.loc[row_num, 'NOTE-1'] = note1
    df_date.loc[row_num, 'NOTE-2'] = note2


# Enter addition row data for appended dataframe if more than one dates found
def enter_append_data(copy, date_list, note1, note2, image):
    global df_append
    copy2 = copy
    copy2['NOTE-1'] = note1
    copy2['NOTE-2'] = note2
    copy2['@IMAGES'] = image

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


# Combine Grocery, Frozen and extra products into one dataframe
def append_df(df1, df2, df3):
    combined_df = init_empty_df()
    combined_df = combined_df.append(df1, ignore_index=True)
    combined_df = combined_df.append(df2, ignore_index=True)
    combined_df = combined_df.append(df3, ignore_index=True)
    # Rename Headers to match CSV output file
    combined_df = combined_df.rename(index=str, columns={'ITEM CODE':'SKU','PRODUCT DESCRIPTION':'DESC-1', 'PACKAGING SPECS':'PACK'})
    return combined_df


# Output to CSV files.  Three files outputted.  1-Combined list, 2-List of NEW and JA items, 3-All remaining items
def output_csv(df_combined):
    # Remove Original Incoming Month Column
    del df_combined[month]
    df_combined.sort_values('NOTE-1', ascending=False, inplace=True)
    df_combined = df_combined.reset_index(drop=True)
    # Output Combined CSV
    df_combined.to_csv('Combined' + month + '.csv', index=False)

    # Output CSV containing NOTE-1 with *JA* and *NEW*
    # Determine end of item in column
    num_item = 0
    for item in df_combined['NOTE-1']:
        num_item += 1
        if item == '':
            break
    print("Number of Note-1: ", num_item)

    # Retrieve the items of the column into a new dataframe
    df_new = df_combined.loc[0:num_item - 2,
             ['SKU', 'BRAND', 'DESC-1', 'PACK', '@IMAGES',
              'INCOMING', 'NOTE-1', 'NOTE-2', 'ORIGIN', 'CATEGORY', 'SUBCATEGORY']]

    df_new.to_csv('New' + month + '.csv', index=False)

    # Output CSV containing NOTE-1 that is empty
    df_empty = df_combined.loc[num_item - 1: df_combined.index[-1],
               ['SKU', 'BRAND', 'DESC-1', 'PACK', '@IMAGES',
                'INCOMING', 'NOTE-1', 'NOTE-2', 'ORIGIN', 'CATEGORY', 'SUBCATEGORY']]

    df_empty.to_csv('Empty' + month + '.csv', index=False)


df_grocery = new_sorted_df(month, 'Grocery-Master')

grocery = Products('Grocery-Master', df_grocery)
print(grocery.df)
# Initialize empty dataframe for additional rows for extra dates
df_append = init_empty_df()
find_append_dates(grocery)
print(grocery.df)
print(df_append)

df_frozen = new_sorted_df(month, 'Frozen-Master')
frozen = Products('Frozen-Master', df_frozen)
find_append_dates(frozen)
print(frozen.df)
print(df_append)
print(frozen.df['@IMAGES'])
print(frozen.df['CATEGORY'])
combined_set = append_df(grocery.df, frozen.df, df_append)
output_csv(combined_set)

