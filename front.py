import PySimpleGUI as sg
import pandas as pd
import os
import numpy as np
r = 1

import ctypes
import platform
def make_dpi_aware():
    if int(platform.release()) >= 8:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)
make_dpi_aware()

def get_scaling():
    # called before window created
    root = sg.tk.Tk()
    scaling = root.winfo_fpixels('1i')/72
    root.destroy()
    return scaling

print(get_scaling())
original = 3.3366174055829227
my_width = 3840
my_height = 2160
scaling_old = get_scaling()
width, height = sg.Window.get_screen_size()

scaling = scaling_old * min(width / my_width, height / my_height)


def pivot(dataframe, values, index, columns):
    check = True
    if dataframe is not None:
        v = values.split(',')
        i = index.split(',')
        c = columns.split(',')
        bag = set()
        for elt in v:
            bag.add(elt)
        for elt in i:
            bag.add(elt)
        for elt in c:
            bag.add(elt)
        for elt in bag:
            if elt not in xl.keys():
                print("Warning : invalid column name :" + elt)
                check = False
        if check:
            table = pd.pivot(dataframe, values=v, index=i, columns=c)
            print("\n --- PIVOT --- \n")
            print(table)
            return table


def drop(dataframe, names):
    check = True
    if dataframe is not None:
        n = names.split(",")
        for elt in n:
            if elt not in dataframe.keys():
                check = False
                print("Error with columns name :" + elt)
        if check:
            cols = []
            for elt in n:
                cols.append(elt)
            dropped = dataframe.drop(columns = cols)
            return dropped




sg.theme("Black")
sg.set_options(font=("Consolas", int(13*r)), scaling=scaling)

function_keys = ['-NULL-', '-PIVOT-', '-SHOWCOLUMNS-', '-DROPCOLUMNS-']

frame1 = [[sg.Radio('Show columns name', 1, key=function_keys[2])],
          [sg.Radio('Drop colums', 1, key=function_keys[3]), sg.Frame('Names', [[sg.Input(key=function_keys[3]+'NAMES-', size=(int(10*r),int(10*r)))]], font=('Consolas', int(10*r)))],
          [sg.Radio('Pivot',1, key=function_keys[1], default=True), sg.Frame('Values', [[sg.Input(key=function_keys[0]+'VALUES-', size=(int(10*r),int(10*r)))]], font=('Consolas', int(10*r))), sg.Frame('Index', [[sg.Input(key=function_keys[0]+'INDEX-', size=(int(10*r),int(10*r)))]], font=('Consolas', int(10*r))), sg.Frame('Columns', [[sg.Input(key=function_keys[0]+'COLUMNS-', size=(int(10*r),int(10*r)))]], font=('Consolas', int(10*r)))],
          [sg.Radio('Coming Soon', 1, key=function_keys[0])]]

file_open = [[sg.Input(key='-INPUT-'),sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"), ("ALL Files", "*.*"))),sg.Button("Open"),]]

col1 = [[sg.Button('Execute')],
        [sg.Button('Save output')]]


layout = [  [sg.Frame('', file_open), sg.Column([[sg.Button('Clear output'), sg.Button('Help')]], element_justification="r", expand_x=True)],
            [sg.Output(size=(int(100*r),int(10*r)), key='-OUTPUT-')],
            [sg.Frame('Basic functions', frame1, size=(int(1500*r),int(450*r)), pad=(int(100*r),int(20*r))), sg.Column(col1, element_justification='l')]  ]

window = sg.Window('Small & Simple', layout)
filename = ''
res = None
xl = None

while True:
    event, values = window.read()

    if event is None:
        break
    if event == sg.WINDOW_CLOSED:
        break
    if event == 'Open':
        if type(values['-INPUT-']) == str and values['-INPUT-'] != filename:
            filename = values['-INPUT-']
        if os.path.isfile(filename):
            try:
                xl = pd.read_excel(filename)
                print('File preview :')
                print("")
                print(xl)
            except Exception as e:
                print("Error: ", e)
    if event == 'Execute':
        if values[function_keys[1]]:
            try:
                res = pivot(xl, values[function_keys[1]+'VALUES-'], values[function_keys[1]+'INDEX-'], values[function_keys[1]+'COLUMNS-'])
            except:
                print("Error")
        if values[function_keys[2]]:
            try:
                print(xl.columns)
            except:
                print("Error with file")
        if values[function_keys[3]]:
            try:
                xl = drop(xl, values[function_keys[3]+'NAMES-'])
                print(" --- Columns dropped ---")
                print(xl)
            except:
                print("Error")
    if event == 'Clear output':
            window['-OUTPUT-'].update('')
    if event == 'Save output':
        if res is not None:
            file_name = sg.PopupGetFile('Please enter filename to save', save_as=True)
            if file_name[-5:] != '.xlsx':
                file_name+='.xlsx'
            res.to_excel(file_name)
            print("")
            print(" --- File saved at :" + file_name)
        else:
            print("Warning : No valid output")
    if event == 'Help':
        sg.PopupQuickMessage("If you don't understand something, call me <3")
window.close()