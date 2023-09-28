from tkinter import filedialog
import openpyxl
import numpy as np

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

event_dictionary = dict()

shape = (ending_date_cell_column, 33)
dates_for_month = np.empty(shape, dtype='str')

count_nulls = 0




# add data to 2d array
for i in range(2, ending_date_cell_column):
    for j in range(3, 33):
        current_value = ws.cell(row=j, column=i).value
        count_nulls += 1
        if current_value is not None:
            dates_for_month[i - 2][j - 3 - count_nulls] = current_value
    count_nulls = 0


# get available months
for i in range(2, ending_date_cell_column):
    available_months.append(ws.cell(row=2, column=i).value)

for month, dates in zip(available_months, dates_for_month):
    event_dictionary[month] = dates

print(event_dictionary)