# updated: 4/24/23 (v2)

from pathlib import Path

import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('LightBlue2')

layout = [
    [sg.Text('Monthly Payment: $'), sg.Text(size=(10,1), key='output')]
    [sg.Text('Please fill out the following fields:')],
    [sg.Frame('Vehicle Price',[[
        sg.Slider(range=(0, 80000), orientation='h', size=(50, 10), default_value=16000),]])],
    [sg.Text('Down Payment', size=(15,1)), sg.InputText(key='Down Payment')],
    [sg.Text('Trade-In Value', size=(15,1)), sg.InputText(key='Trade-In Value')],
    [sg. Text('Credit Score', size=(15,1)), sg.Combo(['300-599', '600-659', '660-719', '720-850'], key='Credit Score')],
    [sg. Text('Loan Term', size=(15,1)), sg.Combo(['72 months', '60 months', '48 months', '36 months'], key='Loan Term')],
    #[sg.Text('I speak', size=(15,1)),
    #    sg.Checkbox('German', key='German'),
    #   sg.Checkbox('Spanish', key='Spanish'),
    #    sg.Checkbox('English', key='English')],
    #[sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)], initial_value=0, key='Children')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]    
]


window = sg.Window('Monthly Car Payment Form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel('Data saved!')
        sg.popup('Data saved!')
        clear_input()
window.close()        