# updated: 4/24/23 (v3)

from pathlib import Path
import os
import PySimpleGUI as sg
import pandas as pd
import openpyxl as op
import csv

# Add some color to the window
sg.theme('LightBlue2')

current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
#current_dir = os.getcwd()
EXCEL_FILE = current_dir / 'DataSheet.xlsx'
#EXCEL_FILE = os.path.join(current_dir, "DataSheet.xlsx")
df = pd.read_excel(EXCEL_FILE)

# Layout for first window
layout = [
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Please fill out the following fields:')],
    [sg.Frame('Vehicle Price',[[
        sg.Slider(range=(0, 80000), key='Price', orientation='h', size=(50, 10), default_value=16000),]])],
    [sg.Text('Down Payment', size=(15,1)), sg.InputText(key='Down Payment')],
    [sg.Text('Trade-In Value', size=(15,1)), sg.InputText(key='Trade-In Value')],
    [sg. Text('Credit Score', size=(15,1)), sg.Combo(['300-599', '600-659', '660-719', '720-850'], key='Credit Score')],
    [sg. Text('Loan Term', size=(15,1)), sg.Combo(['36 months', '48 months', '60 months', '72 months'], key='Loan Term')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()] 

]

# Layout for second window
layout1 = [
    [sg.Text('Monthly Payment: $'), sg.Text(size=(70,1), key='Monthly')]
]

# Display first Window
window = sg.Window('Monthly Car Payment Form', layout)

# Clear data fields
def clear_input(values):
    for key in values:
        window[key]('')
    return None

while True:
    # Read data fields
    event, values = window.read()
    # Turn user input into variables
    name = str(values['Name'])
    price = float(values['Price'])
    down = float(values['Down Payment'])
    trade = float(values['Trade-In Value'])
    score = str(values['Credit Score'])
    term = str(values['Loan Term'])
    pv = price - down - trade

    if event in (sg.WIN_CLOSED, 'Exit'):
        break

    if event == 'Clear':
        clear_input(values)

    if event == "Submit":
        window1 = sg.Window('Your Payment is:', layout1, finalize=True)
        if term == "36 months":
            term = 36
        elif term == "48 months":
            term = 48
        elif term == "60 months":
            term = 60
        elif term == "72 months":
            term = 72
        else:
            print("error")

        if score == "300-599":
            rate = 0.1
        elif score == "600-659":
            rate = 0.07
        elif score == "660-719":
            rate = 0.05
        elif score == "720-850":
            rate = 0.03
        # monthly rate 
        r = rate/12
        payment = (r * pv) / (1 - (1 + r) ** (-term))      
        window1['Monthly'].update(value=f'{payment:.2f}')

        data = [
 {'Name': name, 'Price': price, 'Down Payment': down, 'Trade-In Value': trade, 'Credit Score': score, 'Loan Term': term, 'Monthly': payment}
 ]
        new_record = pd.DataFrame(data, index=[0])
        df = pd.concat([df, new_record], ignore_index=False)
        df.to_excel(EXCEL_FILE, sheet_name='name', index=False)
    sg.popup('Data saved!')