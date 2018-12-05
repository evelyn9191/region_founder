"""Fill name of regions to chosen Excel sheet based on city location."""

import os
import pandas as pd
from shutil import copyfile
from openpyxl import load_workbook


def input_file_check():
    """Get input file name. Check if it exists.

    :return: Name of the input file.
    """
    input_file = input("Put the .xlsx file with shops to the same directory, where you put "
                       "the file you are currently running. What is the name "
                       "of the .xlsx file? (Write it as filename.xlsx) ")
    if os.path.exists(input_file) is False:
        print("Such file does not exist. Did you put it into a right directory? "
              "Is the name of the file right? Try again.")
        input_file_check()
    if input_file.lower().endswith('.xlsx') is False:
        print("I cannot process such file. Give me a file with .xlsx extension "
              "(Excel file of version 2007 and later.)")
        input_file_check()
    return input_file


def user_data():
    """Get input data.

    Get specific descriptions of relevant columns and country name.

    :return: Input data.
    """
    country = input("What is the country where the shops are located? (Type CZ or SR) ")
    cities = input("What is the column that contains names of the cities? (for example: D) ")
    regions = input("What is the column where the names of the regions should go? (for example: E) ")
    start = input("What is the number of the row with the first city? (for example: 5) ")
    finish = input("What is the number of the row with the last city? (for example: 25) ")
    return {"country_name": country,
            "cities_column": cities,
            "regions_column": regions,
            "first_row": start,
            "last_row": finish
            }

"""
def correct_data_check(input_file, user_data):
    # Check if input data are correct.
    # TODO: Check if file with regions is there
    # TODO: Check if input is right. Check if numbers of areas match.
    # TODO: Redo correct_data_check
    # TODO: If first time value in user_data is wrong, program keeps working with the wrong
    # value even if it was fixed in correct_data_check
    check_sheet_name = load_workbook(input_file, read_only=True)
    while user_data["sheet_name"] not in check_sheet_name.sheetnames:
        print("The name of the sheet is not correct.")
        new_sheet = input("Write it again: ")
        user_data["sheet_name"] = new_sheet
    df = pd.read_excel(input_file, user_data["sheet_name"])
    while user_data["street_column"] not in df:
        print("The name of the column with streets is not correct.")
        new_street = input("Write it again: ")
        user_data["street_column"] = new_street
    while user_data["city_column"] not in df:
        print("The name of the column with cities is not correct.")
        new_city = input("Write it again: ")
        user_data["city_column"] = new_city
    while user_data["postal_column"] not in df:
        print("The name of the column with postal code is not correct.")
        new_postal = input("Write it again: ")
        user_data["postal_column"] = new_postal
    while user_data["gps_column"] not in df:
        print("The name of the column with gps is not correct.")
        new_gps = input("Write it again: ")
        user_data["gps_column"] = new_gps
    return user_data
"""

def get_region(input_file, user_data):
    """Process :input_file according to :user_data.

    Open the input file, read data, match cities with regions for each line,
    and store regions.

    :param dict[str] user_data: Input data with info about processing the file.
    :param str input_file: Name of the file to process.
    """
    df = pd.read_excel(input_file, [0])
    df_regions = pd.read_excel("kraje.xlsx", [user_data["country_name"]])

    start = int(user_data["first_row"])
    finish = int(user_data["last_row"])
    all_rows = range(start, finish)

    file_name_split = os.path.splitext(input_file)
    root_name = file_name_split[1]
    copyfile(input_file, "{}_regions_added.xlsx".format(root_name))
    output_file = "{}_regions_added.xlsx".format(root_name)
    wb = load_workbook(output_file)
    ws = wb.active

    for row_number in all_rows:
        if df_regions["Okres"] in df[user_data["cities_column"]]: # TODO: KeyError: "Okres" - or anything inside the parentheses
            ws.cell(row=user_data["first_row"], column=user_data["regions_column"]).value == df_regions[row_number, 1]
    wb.save(output_file)
    print("Regions successfully matched to cities in", output_file)


if __name__ == "__main__":
    input_file = input_file_check()
    user_data = user_data()
    # clean_user_data = correct_data_check(input_file=input_file, user_data=user_data)
    get_region(input_file=input_file, user_data=user_data)