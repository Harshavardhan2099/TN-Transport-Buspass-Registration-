import PySimpleGUI as sg
import pandas as pd
#Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'Data.xlsx'
dd = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text(' ', size = (10)), sg.Text('TamilNadu Transport Corporation - Kumbakonam')],
    [sg.Text('Ranees Govt. Hr. Sec. School', size = (40,1)), sg.Text('School Code: 199')],
    [sg.Text('Name', size = (15, 1)), sg.InputText(key = 'Name')],
    [sg.Text('Class', size = (15, 1)), 
            sg.Combo(['6','7','8','9','10','11','12'], key = 'Class')],
    [sg.Text('Section', size = (15,1)), sg.InputText(key = 'Section')],
    [sg.Text('Address 1', size = (15, 1)), sg.InputText(key = 'Address1')],
    [sg.Text('Address 2', size = (15, 1)), sg.InputText(key = 'Address2')],
    [sg.Text('Address 3', size = (15, 1)), sg.InputText(key = 'Address3')],
    [sg.Text('From', size = (15, 1)), sg.InputText(key = 'From')],
    [sg.Text('To', size = (15, 1)), sg.InputText(key = 'To')],

    [sg.Submit(), sg.Exit()]
]

window = sg.Window('Bus-Pass Registration', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        dd = dd._append(values, ignore_index = True)
        dd.to_excel(EXCEL_FILE, index = False)
        sg.popup('Data saved!')

window.close()