#!/usr/bin/python

import sys
import xlrd
import csv
import json

# list struct: user_param[]
# [ Save_which_Type, "path/to/excel", sheet_index, "path/to/save/file" ]
user_param = [];

# list struct: excel_basic_information[]
# [NumColumn, NumRow ]
excel_basic_information = [];

# ignore the asking string.
onelineparam = 0;

def add_user_param (user_param, asking_string, index, onelineparam):
    if onelineparam == 1:
        temp = sys.argv[index + 1];
    else:
        temp = input (asking_string);

    user_param.insert(index, temp);
    return

def print_list (user_param):
    print("User param Length is %d\n" % user_param.__len__());
    for val in user_param:
        print val
    return

def excel_file_init():
    # print("Init the excel data handle\n");
    # open the excel file to book object
    temphandle = xlrd.open_workbook(user_param[1]);
    return temphandle

def excel_sheet_option(user_param, excel_handle):
    if onelineparam != 1:
        print("Your input excel file has %d's sheet:\n" % excel_handle.nsheets);
        numberofsheet = excel_handle.nsheets - 1;
        print("Please choose the sheet index(from 0 ~ %d) you want to parse\n" % numberofsheet);
        print("Sheet list:\n");
        index = 0;
        for element in excel_handle.sheet_names():
            print("%d: %s" % (index, element));
            index = index + 1;

    add_user_param(user_param, "?", 2, onelineparam);

    if not excel_handle.sheet_loaded(int(user_param[2])):
        print("load the sheet again\n");
        excel_sheet_init(excel_handle);

    temphandle = excel_handle.sheet_by_index(int(user_param[2]));

    return temphandle

def excel_sheet_init(excel_handle):
    # print("Init the excel sheet handle\n");
    # load all sheet.
    excel_handle.sheets();
    return

def parse_excel_row_number(excel_basic_information, excel_sheet_handle):
    excel_basic_information.insert(1, excel_sheet_handle.nrows);
    return

def parse_excel_column_number(excel_basic_information, excel_sheet_handle):
    excel_basic_information.insert(0, excel_sheet_handle.ncols);
    return

def create_excel_content_data_list(excel_basic_information):
    temp = [];
    for index in range(excel_basic_information[1]):
        temp.append([]);
    return temp

def parse_content_data(excel_sheet_handle, excel_content_list, excel_basic_information):
    for row in range(excel_basic_information[1]):
        for col in range(excel_basic_information[0]):
            if excel_sheet_handle.cell_type(row, col) is 3:
                date_tuple = xlrd.xldate.xldate_as_tuple(excel_sheet_handle.cell_value(row, col), 0);
                date_format = `date_tuple[0]` + '-' + `date_tuple[1]` + '-' + `date_tuple[2]`;
                excel_content_list[row].append(date_format);
            elif excel_sheet_handle.cell_type(row, col) is 2:
                excel_content_list[row].append(int(excel_sheet_handle.cell_value(row, col)));
            else:
                excel_content_list[row].append(excel_sheet_handle.cell_value(row, col).encode("UTF-8"));
            #print("%d x %d: %r" % (row, col, excel_content_list[row][col]));
    return

def create_json():
    add_user_param(user_param, "Create Json Format\nSelect where do you want to save file\n?", 3, onelineparam);
    # consist the dictionary. 
    jsonlist = [];
    jsondict = {};
    for row in range(1, excel_basic_information[1], 1):
        for col in range(excel_basic_information[0]):
            jsondict[excel_content_list[0][col]] = excel_content_list[row][col];
        jsonlist.append(jsondict.copy());

    # due to the sheet name is not global variable. so we have to assign the name to be the default
    objname = "Sheet" + `int(user_param[2])`;

    with open(user_param[3], 'w') as jsonout:
        json.dump({objname:jsonlist}, jsonout, indent=4);
    return

def create_csv():
    add_user_param(user_param, "Create CSV Format\nSelect where do you want to save file\n?", 3, onelineparam);

    with open(user_param[3], 'wb') as csvout:
        csvwriter = csv.writer(csvout, delimiter=',', quotechar="'", quoting=csv.QUOTE_ALL);

        for row in range(excel_basic_information[1]):
                csvwriter.writerow(excel_content_list[row]);

    return

def create_xml():
    add_user_param(user_param, "Create XML Format\nSelect where do you want to save file\n?", 3, onelineparam);
    return

def create_buffering():
    #print("Create buffering format");
    for index in range(excel_basic_information[1]):
        print("%r" % str(excel_content_list[index]));
    #exit(excel_content_list);
    return

def create_translator():
    add_user_param(user_param, "Create Translator.xml Format\nSelect where do you want to save file\n?", 3, onelineparam);
    return

def error_saveformat():
    print("Gone wrong format");
    return

if len(sys.argv) > 2:
    onelineparam = 1;

add_user_param(user_param, "Welcome to python parse Excel file\n1. to Json Format file\n2. to CSV Format file\n3. to XML Format file\n4. to buffer array\n5. to translator.xml\n?", 0, onelineparam);
add_user_param(user_param, "Please enter the full path to the excel file location. (e.g. \"/Path/TO/FILE/WITH/QUOTE\")\n?", 1, onelineparam);
#print_list(user_param);

# initialise the excel book object
excel_handle = excel_file_init();

# initialise the excel sheet object
excel_sheet_init(excel_handle);
excel_sheet_handle = excel_sheet_option(user_param, excel_handle);

# parse the needed information
parse_excel_column_number(excel_basic_information, excel_sheet_handle);
parse_excel_row_number(excel_basic_information, excel_sheet_handle);

# create the 2D list according to row and column number
excel_content_list = create_excel_content_data_list(excel_basic_information);
parse_content_data(excel_sheet_handle, excel_content_list, excel_basic_information);

# Consist and Save to specific format
saveformat = int(user_param[0])

{
        # Just put the object address in the switch case
        1: create_json,
        2: create_csv,
        3: create_xml,
        4: create_buffering,
        5: create_translator
}.get(saveformat, error_saveformat)()
