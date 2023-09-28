from tkinter import filedialog

# start up dialog
file_name = filedialog.askopenfilename(
    title="Open Text File",
    filetypes=(("Excel Files", "*.xls"),))


print(file_name)
