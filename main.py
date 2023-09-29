from tkinter import filedialog
import openpyxl
import numpy as np
import re
import datetime


class Assignment:

    def __init__(self, _subject, _day, _month, _type):
        self.subject = _subject
        self.type = _type
        self.date = self.create_date_of_month_day(_day, _month)

    def create_date_of_month_day(self, _day, _month):
        formatted_day = int(self.add_zero_to_single_digit_numbers(_day))
        current_year = int(datetime.datetime.now().year)

        return datetime.datetime(year=current_year, month=_month, day=formatted_day)

    def add_zero_to_single_digit_numbers(self, input_string):
        if len(input_string) == 1:
            return "0" + input_string
        else:
            return input_string

# dictionary
dictionary = {
    "EBA E1": "Englisch",
    "0WRRE": "Recht",
    "0WRWI": "Wirtschaft",
    "0MEDTSM": "Mobile Computing",
    "0D": "Deutsch",
    "1SEW": "Softwareentwicklung",
    "1INSY": "Datenbanken",
}

# get all dictionary keys
dictionary_key = list(dictionary.keys())

available_months = list()

# TODO: Change later on to -> ""
file_path = "test.xlsx"

while file_path == "":
    print("No file selected.\nRestarting...")
    # start up dialog
    file_path = filedialog.askopenfilename(
        title="Open Text File",
        filetypes=(("Excel Files", "*.xlsx"),))

wb = openpyxl.load_workbook(file_path)

print("You have successfully selected: " + file_path)

ws = wb['Sheet1']

# final values
starting_date_cell_column = 2
ending_date_cell_column = 11
date_row = 2

# all list
assignment_list = set()

# add data to 2d array
for i in range(2, ending_date_cell_column):
    for j in range(3, 33):
        if ws.cell(row=j, column=i).value is not None:
            _day = ""
            _month = ws.cell(row=j, column=i).value
            _subject = ""
            _type = ""
            assignment_list.add(Assignment(_subject, _day, _month, _type))


