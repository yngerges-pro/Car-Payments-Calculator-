# updated: 5/14/23 (v3)

from pathlib import Path
import PySimpleGUI as sg
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import mysql.connector
from AnalyzingPython import Ana, DataAnalysisfun, function
import tkinter as tk

current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Car.xlsx'
Design = current_dir/'form2.png'
df = pd.read_excel(EXCEL_FILE)

def create_payment_window(pay, tot):
    win = tk.Tk()
    win.geometry("300x100")
    win.config(bg = "#FFEEF2")
    win.title("Monthly Payment")
    pay = tk.Label(win, text = 'Monthly Payment: $'+ str(f'{pay:.2f}'), bg= "#FFEEF2" ,fg = "#98847F", font = ('Lato', 12))
    pay.pack(anchor="w")
    tot = tk.Label(win, text = 'Total Payment: $'+ str(f'{tot:.2f}'), bg= "#FFEEF2" ,fg = "#98847F", font = ('Lato', 12))
    tot.pack(anchor="w")
    return win


def create():
    window = tk.Tk()
    window.geometry("600x700")
    window.title("Monthly Car Payment Form")
    window.config(bg="white")

    canvas = tk.Canvas(window, width = 550, bg="white", height = 820)

    img = tk.PhotoImage(file= Design)
    canvas.create_image(300, 300, image =img)
    canvas.pack()

    window.configure(bg = "white")

    nameTxtBox = tk.Text(window, height = 1, width = 20, fg = "#98847F", bg= "white", font = ('Lato', 15))
    nameTxtBox.place(x = 250, y =145)
    slider = tk.Scale(window, from_=16000, to=86000, fg = "#98847F", bg= "white", orient="horizontal", length = 180)
    slider.pack()
    slider.place(x= 290, y =240)
    downTxtBox = tk.Text(window, height = 1, width = 16, fg = "#98847F", bg= "white", font = ('Lato', 15))
    downTxtBox.place(x = 295, y =300)

    tradeTxtBox = tk.Text(window, height = 1, width = 16, fg = "#98847F", bg= "white", font = ('Lato', 15))
    tradeTxtBox.place(x = 295, y =345)

    op = ['300-599', '600-659', '660-719', '720-850']
    variable = tk.StringVar(window)
    variable.set(op[0])
    credit = tk.OptionMenu( window , variable , *op)
    credit.pack()
    credit.config(background="white", fg="#98847F")
    credit.place(x= 295, y =395)

    op2 = ['36 months', '48 months', '60 months', '72 months']
    variable2 = tk.StringVar(window)
    variable2.set(op2[0])
    term = tk.OptionMenu( window , variable2 , *op2 )
    term.pack()
    term.config(background="white", fg="#98847F")
    term.place(x= 295, y =445)

    def clear_input():
        nameTxtBox.delete("1.0","end")
        downTxtBox.delete("1.0","end")
        tradeTxtBox.delete("1.0","end")
        slider.set(16000)
        variable.set(op[0])
        variable2.set(op2[0])
  

    def submit_info():
        # connect to mySQL database
        connection = mysql.connector.connect(host="localhost", user="guest", passwd="iamaguest", database = "payments")
        cursor = connection.cursor()
        current_dir = Path(file).parent if 'file' in locals() else Path.cwd()
        name = str(nameTxtBox.get("1.0","end"))
        down = float(downTxtBox.get("1.0","end"))
        trade = float(tradeTxtBox.get("1.0","end"))
        price = float(slider.get())
        score = variable.get()
        term = variable2.get()
        clear_input()
        pv = price - down - trade
            
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
    

        Payment_window = create_payment_window(payment, total)
        
        data = [{'Name': name, 'Price': price, 'Down Payment': down, 'Trade-In Value': trade, 'Credit Score': score, 'Loan Term': term, 'Monthly': payment, 'Total': total}]
        new_record = pd.DataFrame(data, index=[0])
        global df
        df = pd.concat([df, new_record], ignore_index=False)
        df.to_excel(EXCEL_FILE, sheet_name='name', index=False)
        save = tk.Label(Payment_window, text = 'Data Saved!', bg= "#FFEEF2" ,fg = "#98847F", font = ('Lato', 12))
        save.pack(anchor="w")
        okay = tk.Button(Payment_window, text ="Ok", command = Payment_window.destroy, fg = "#98847F", background="#FFEEF2")
        okay.pack()
        # Add values to mySQL
        query = "INSERT INTO customers (Name, Price, Down_Payment, Trade_In_Value, Credit_Score, Loan_Term, Monthly_Payment, Total) values (%s, %s, %s, %s, %s, %s, %s, %s)" 
        cursor.execute(query, (name, price, down, trade, score, term, payment, total))
        print(cursor.rowcount, "Record added to mySQL.")
        # Commit to mySQL
        connection.commit()
        connection.close()
        window.destroy()
        create()
        

        
    def analyze_data():
        function()

    submit = tk.Button(window, text ="Submit", command = submit_info, fg = "#98847F", background="#FFEEF2")
    submit.place(x= 180, y =515)
    clear = tk.Button(window, text ="Clear", command = clear_input, fg = "#98847F", background="#FFEEF2")
    clear.place(x= 260, y =515)
    exit_ = tk.Button(window, text ="Exit", command = window.destroy, fg = "#98847F", background="#FFEEF2")
    exit_.place(x= 330, y =515)
    analyze = tk.Button(window, text ="Analyze", command = analyze_data, fg = "#98847F", background="#FFEEF2")
    analyze.place(x= 400, y =515)

    return window.mainloop()


def main():
    create()

main()

      

             

