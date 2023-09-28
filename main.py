from tkinter import filedialog
import openpyxl

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

file_name = ""

while file_name == "":
    print("No file selected.\nRestarting...")
    # start up dialog
    file_name = filedialog.askopenfilename(
        title="Open Text File",
        filetypes=(("Excel Files", "*.xls"),))

print(file_name)


