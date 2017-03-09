from openpyxl import load_workbook
import glob
import re
import os
# todo 
# add settings dialog to get location of excel file, also allow for selection of sheet to search within the excel file
xlFile = '' # location of excel file to search
xlSheet = '' # sheet name
wb = load_workbook(xlFile)
sh = wb.get_sheet_by_name(xlSheet)
for row_index in range(sh.get_highest_row()):
    if sh.cell(row=row_index, column=0).value == name:
        print(row_index)