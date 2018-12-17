"""Fill name of regions to chosen Excel sheet based on city location."""

import os
from sys import exit
import pandas as pd
from shutil import copyfile
from openpyxl.utils import column_index_from_string


def regions_file_check():
    """Check if dataframe with regions exists."""
    if os.path.exists("kraje.xlsx") is False:
        print("The file kraje.xlsx that was originally distributed with this program is missing. "
              "Put the file to the folder where this script is and start the script again.")
        exit()


def input_file_check():
    """Get input file name. Check if it exists.
    :return Name of the input file.
    """
    input_file = input("Put the .xlsx file with shops to the same directory, where you put "
                       "the file you are currently running. What is the name "
                       "of the .xlsx file? (Write it as filename.xlsx) ")
    while True:
        if os.path.exists(input_file) is False:
            print("Such file does not exist. Did you put it into a right directory? "
                  "Is the name of the file right?")
            input_file = input("Try again: ")
        if input_file.lower().endswith('.xlsx') is False:
            print("I cannot process such file. Give me a file with .xlsx extension "
                  "(Excel file of version 2007 and later.)")
            input_file = input("Try again: ")
        break

    return input_file


def user_data():
    """Get input data.
    Get specific descriptions of relevant columns and country name.
    :return Input data.
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

def output_file_name(input_file):
    """Create output file.
    Create output file as a copy of the input file with the right name.
    :return Name of the output file.
    """
    file_name_split = os.path.splitext(input_file)
    root_name = file_name_split[0]
    copyfile(input_file, "{}_regions_added.xlsx".format(root_name))
    output_file = "{}_regions_added.xlsx".format(root_name)
    return output_file


def get_region(input_file, user_data, output_file):
    """Process :input_file according to :user_data and compare it with regions dataframe.
    Open the input file, read data, match cities with regions for each line by comparing
    them to the regions dataframe, and store region names to a new file.
    :param dict[str] user_data: Input data with info about processing the file.
    :param str input_file: Name of the file to process.
    :param str output_file: Name of the file where to save final data.
    """
    df = pd.read_excel(input_file, header=None)
    df_regions = pd.read_excel("kraje.xlsx", user_data["country_name"])
    cities_column_number = column_index_from_string(user_data["cities_column"]) - 1
    regions_column_number = column_index_from_string(user_data["regions_column"]) - 1

    start = int(user_data["first_row"]) - 1
    finish = int(user_data["last_row"]) - 1
    all_rows = range(start, finish)

    writer = pd.ExcelWriter(output_file)
    # wb = load_workbook(output_file) - to be deleted
    # ws = wb.active - to be deleted - to be deleted

    for row_number in all_rows:
        line = df.iloc[row_number, cities_column_number]
        match = df_regions.iloc[:,0].apply(lambda district: district in line)
        region_name = df_regions.loc[match, 'Kraj']
        if not region_name.empty:
            df.iloc[row_number, regions_column_number] = region_name.iloc[0]
    print(df) # TODO: the function is not generating regions
    df.to_excel(writer)
    print("Regions successfully matched to cities in", output_file)


if __name__ == "__main__":
    regions_file_check()
    input_file = input_file_check()
    user_data = user_data()
    output_file = output_file_name(input_file=input_file)
    get_region(input_file=input_file, user_data=user_data, output_file=output_file)
