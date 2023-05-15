# updated: 4/24/23 (v3)

from pathlib import Path
import PySimpleGUI as sg
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import mysql.connector
from AnalyzingPython import Ana, DataAnalysisfun, FinalFun

# Add some color to the window
sg.theme('LightBlue2')
 # connect to mySQL database
# connection = mysql.connector.connect(host="localhost", user="guest", passwd="iamaguest", database = "payments")
# cursor = connection.cursor()
# connect to Excel
current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Car.xlsx'
df = pd.read_excel(EXCEL_FILE)

def create_Layout():
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
    return sg.Window('Monthly Car Payment Form', layout)


def create_payment_window():
    layout1 = [
        [sg.Text('Monthly Payment: $'), sg.Text(size=(70,1), key='Monthly')],
        [sg.Text('Total Payment: $'), sg.Text(size=(70,1), key='Total')]
    ]
    return sg.Window('Your Payment is:', layout1, finalize=True)



while True:
    window = create_Layout()
    def clear_input(values):
        for key in values:
            window[key]('')
        return None
    # Read data fields
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == 'Clear':
        clear_input(values)
        continue

    # Turn user input into variables
    name = str(values['Name'])
    price = float(values['Price'])
    down = float(values['Down Payment'])
    trade = float(values['Trade-In Value'])
    score = str(values['Credit Score'])
    term = str(values['Loan Term'])
    pv = price - down - trade

    
    if event == "Submit":
        Payment_window = create_payment_window()
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
        tax = 0.08
        monthlyTax = (tax*price)/term
        temp1 = (r * pv) / (1 - (1 + r) ** (-term))
        payment = temp1 + monthlyTax
        total = (payment * term) + down
        if down == price:
            payment = 0     
        Payment_window['Monthly'].update(value=f'{payment:.2f}')
        Payment_window['Total'].update(value=f'{total:.2f}')
        data = [{'Name': name, 'Price': price, 'Down Payment': down, 'Trade-In Value': trade, 'Credit Score': score, 'Loan Term': term, 'Monthly': payment, 'Total': total}]
        new_record = pd.DataFrame(data, index=[0])
        df = pd.concat([df, new_record], ignore_index=False)
        df.to_excel(EXCEL_FILE, sheet_name='name', index=False)
        sg.popup('Data saved!')
        # Add values to mySQL
        # query = "INSERT INTO customers (Name, Price, Down_Payment, Trade_In_Value, Credit_Score, Loan_Term, Monthly_Payment, Total) values (%s, %s, %s, %s, %s, %s, %s, %s)" 
        # cursor.execute(query, (name, price, down, trade, score, term, payment, total))
        # print(cursor.rowcount, "Record added to mySQL.")
      
# Commit to mySQL
# connection.commit()
# connection.close()
             

FinalFun()