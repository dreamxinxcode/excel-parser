import openpyxl
import os

directory = os.path.dirname(os.path.realpath(__file__))
target = input("Target string: ")
replace = input("Replacment string ")

for files in os.walk(directory):
    for file in files:
        if file.endswith(".xlsx"):
            print(file)
            wb = openpyxl.load_workbook(file)
            for sheet in wb.worksheets:
                for cell in sheet.iter_rows('C{}:C{}'.format(sheet.min_row,sheet.max_row)):
                    print(cell + "\n")
                    if cell.value == target:
                        cell.value = replace

    wb.save(wb)
