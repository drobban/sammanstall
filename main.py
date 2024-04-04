from glob import glob
from openpyxl import load_workbook
from pprint import pprint
import json


def list_input():
    files = glob("./input/*Camilla*.xlsx")
    return files


def read_dog(sheet, min_row, max_row, min_col, max_col):
    cell_range = sheet.iter_rows(
        min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
    )

    return [x.value for [x] in cell_range]


def read_dogs(filename):
    # All our values of interest is found in the sheet
    # named 'Inmatning_av_värden'
    sheet_name = "Inmatning_av_värden"

    wb = load_workbook(filename=filename)

    sheet = wb[sheet_name]

    # Our columns of interest
    cols = range(4, 122, 9)
    # First dog located at
    # value_collection = read_dog(sheet, min_row=11, max_row=55, min_col=4, max_col=4)

    # To read all of our dogs, we iterate cols.
    value_collection = []
    for col in cols:
        value_range = read_dog(sheet, min_row=11, max_row=55, min_col=col, max_col=col)
        value_collection.append(value_range)

    return value_collection


files = list_input()
print(files[0])
dogs = read_dogs(files[0])

pprint(dogs)
