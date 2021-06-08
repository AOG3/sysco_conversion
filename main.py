#!/usr/bin/env python
import csv
import datetime as dt
import time
import PySimpleGUI as sg
import tkinter as tk
import os

# Role: Open and read file
def open_file(csvfile):
    with open(csvfile) as file:
        file = list(csv.reader(file))
        return file

# Get customer name to get part of the new file name
# Role: Filename creation
def customer_name_filename(body):
    name = body[0][0]
    name = str(name).lower()
    new_name = None
    if 'bay shore' in name:
        new_name = 'BS'
    elif 'massapequa' in name:
        new_name = 'Tapped'
    else:
        new_name = 'DJB'
    return new_name

# Get datetime for filename
# Role: Filename creation
def get_dates(body):
    dates = str(body[0][4])
    try:
        dates = dt.datetime.strptime(dates, '%Y-%m-%d')
    except:
        dates = dt.datetime.strptime(dates, '%m/%d/%Y')
    dates = dates.strftime('%d.%m.%y')
    return dates

# This function takes in the csv file and strips the unnessary columns
# Role: Alter Columns
def strip_columns(file, newfilename):
    with open(file, "r") as source:
        reader = csv.reader(source)

        with open(newfilename, "w") as result:
            writer = csv.writer(result)
            for r in reader:
 
                writer.writerow((r[6], r[9], r[10], r[11], r[12], r[13], r[14], r[17]))

def main():
    
    # GUI
    #################################
    sg.ChangeLookAndFeel('BrownBlue')
    form = sg.FlexForm('Sysco csv to xlsx conversion tool')
    layout = [
          [sg.Text('Your Files', size=(15, 1), auto_size_text=False, justification='right'),
          sg.InputText('Click "Browse" to input file'), sg.FileBrowse()],
          [sg.Submit(), sg.Cancel()]
        ]
    button, values = form.Layout(layout).Read()
    file = values[0]
    ################################

    if ".csv" not in file:
        return  sg.Popup("Please input a .csv file")
    else:
        csvfile = open_file(file)
            
        header = csvfile[0]
        body = csvfile[1:]

        # creates the new file title
        new_filename = 'Sysco' + get_dates(body) + customer_name_filename(body) +'.xlsx'

        split_idx = header.index("Split")
        qty_idx = header.index("Qty")
        pack_idx = header.index("Pack")

        # This makes the Pack values, Qty values on Split
        for row in body:
            split = row[split_idx]
            qty = row[qty_idx]
            if split != '':
                row[pack_idx] = qty
        
        # this makes a new file from the results of the new_filename var
        strip_columns(file, new_filename) 

        sg.Popup("Conversion Successul!\n\nYour new file\n\n'{}'\n\nwill be in the\n\n'{}' folder.".format(new_filename, os.getcwd()))

if __name__ == "__main__":
    main()