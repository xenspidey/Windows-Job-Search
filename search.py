from openpyxl import load_workbook
import glob
import re
import os
from tkinter import *
from tkinter.tix import ComboBox
from tkinter.filedialog import askopenfilenames, askopenfilename, asksaveasfilename
import json
# todo 
# add settings dialog to get location of excel file, also allow for selection of sheet to search within the excel file
key_xlfile = 'xlfile'
value_xlFile = ''
value_xlSheet = ''
searchTerm = ''

def settings():
    '''function for dealing with persistant settings'''

    def read_settings():
        '''read settings from file'''
        with open('config.json', 'w+') as f:
            config = json.load(f)
            value_xlFile = config['xlfile']
            value_xlSheet = config['xlsheet']

    def write_settings():
        '''write config settings to file'''
        config = {'xlfile': '', 'xlsheet': ''}
        with open('config.json', 'w+') as f:
            json.dump(config, f)

    def showSettingsdlg():

        def browse():
            file = askopenfilename(parent=Settingsdlg, filetypes=[('Excel', '.xls, xlsx, xlsm')])
            text_xlFile.delete(0, END)
            text_xlFile.insert(0, file)
            wb = load_workbook(file)
            sheets = wb.get_sheet_names()
            for sheet in sheets:
                combo_sheets.insert(0, sheet)


        Settingsdlg = Tk()
        Settingsdlg.geometry('300x200')
        text_xlFile = Entry(parent=Settingsdlg)
        button_Browse = Button(Settingsdlg, text='Browse...', command=browse())
        combo_sheets = ComboBox(parent=Settingsdlg)
        text_xlFile.pack(pady=10, padx=10)
        button_Browse.pack(pady=10, padx=10)
        combo_sheets.pack(pady=10, padx=10)
        Settingsdlg.mainloop()


    try: 
        read_settings()
    except:
        write_settings()
        read_settings()

    if value_xlFile == '' or value_xlSheet == '':
        showSettingsdlg()


def search():
    '''function to search file'''
    def load_searchfile():
        '''load file'''
        wb = load_workbook(value_xlFile)
        sh = wb.get_sheet_by_name(value_xlSheet)
        for row_index in range(sh.get_highest_row()):
            if sh.cell(row=row_index, column=0).value == searchTerm:
                print(row_index)


def main():
    settings()
    root = Tk()
    root.geometry('300x300+1000+500')
    button = Button(root, text='settings', command=lambda: settings())
    button.pack(pady=10, padx=10)
    root.mainloop()

if __name__ == '__main__':
    main()
