import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from pathlib import Path
import PySimpleGUI as sg


current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Car.xlsx'
df = pd.read_excel(EXCEL_FILE)

def Ana():
    totalListed = []
    downListed = []
    for i in range(32):
        downListed.append(df['Down Payment'][i])
        totalListed.append(df['Total'][i])
    

    n = 31
    scoreListed = []
    totalforScores = []
    for _ in range(4):
        n += 1
        scoreListed.append(df['Credit Score'][n])
        totalforScores.append(df['Total'][n])

    a = 35
    totalforTerm = []
    TermListed = []
    for _ in range(4):
        a += 1
        TermListed.append(df['Loan Term'][a])
        totalforTerm.append(df['Total'][a])

    fig, axs = plt.subplots(2, 2, figsize=(8, 4))

    
    axs[0][0].scatter(downListed,totalListed, c='red')
    axs[0][0].set_title("Total Payment vs Down Payment")

    

    axs[0][1].scatter(scoreListed,totalforScores, c='blue')
    axs[0][1].set_title("Total Payment vs Credit Score")

    
    
    axs[1][0].scatter(TermListed,totalforTerm, c='green')
    axs[1][0].set_title("Total Payment vs Loan Term")

    axs[1][1].remove()

    fig.suptitle("Graph")

    fig.tight_layout()
    plt.show()

#----------------------------------slopes-------------------------------------------------
    slope, _ = np.polyfit(downListed, totalListed, 1)
    # # # Print the slope
    scoreListed = [300,600,660,720]
    slopeforScore, _ = np.polyfit(scoreListed, totalforScores, 1)
    slopeforterm, _ = np.polyfit(TermListed,totalforTerm,1)
    
#----------------------------------slopes-------------------------------------------------
    return totalListed,totalforScores,totalforTerm, slope, slopeforScore, slopeforterm


def DataAnalysisfun():
    totalListed,totalforScores,totalforTerm, slope, slopeforScore, slopeforterm = Ana()
    max_val_downPayment = np.max(totalListed)
    min_val_downPayment = np.min(totalListed)

    max_val_score = np.max(totalforScores)
    min_val_score = np.min(totalforScores)
    print(min_val_score,max_val_score)

    max_val_Term = np.max(totalforTerm)
    min_val_Term = np.min(totalforTerm)
    return min_val_downPayment, max_val_downPayment, min_val_score, max_val_score, min_val_Term,max_val_Term, slope, slopeforScore, slopeforterm 

def function():
    layout3 = [
            [sg.Text('The minimum total payment for down payment: $'), sg.Text(size=(15,1), key='minDown')],
            [sg.Text('The maximum total payment for down payment: $'), sg.Text(size=(15,1), key='maxDown')],
            [sg.Text('The minimum total payment for credit score: $'), sg.Text(size=(15,1), key='minScore')],
            [sg.Text('The maximum total payment for credit score: $'), sg.Text(size=(15,1), key='maxScore')],
            [sg.Text('The minimum total payment for loan term: $'), sg.Text(size=(15,1), key='minTerm')],
            [sg.Text('The maximum total payment for loan term: $'), sg.Text(size=(15,1), key='maxTerm')],
            [sg.Text('The slope of total payment for down payment: '), sg.Text(size=(15,1), key='slope')],
            [sg.Text('The slope of total payment for credit score: '), sg.Text(size=(15,1), key='slopeforScore')],
            [sg.Text('The slope of total payment for loan term: '), sg.Text(size=(15,1), key='slopeforterm')],
            [sg.Exit()]
        ]
    return sg.Window('Data Analysis', layout3, finalize=True)

def FinalFun():
    min_val_downPayment, max_val_downPayment, min_val_score, max_val_score, min_val_Term,max_val_Term, slope, slopeforScore, slopeforterm = DataAnalysisfun()    
    while True:
        DataAnalysis = function()
        DataAnalysis['minDown'].update(value=f'{min_val_downPayment:.2f}')
        DataAnalysis['maxDown'].update(value=f'{max_val_downPayment:.2f}')
        DataAnalysis['minScore'].update(value=f'{min_val_score:.2f}')
        DataAnalysis['maxScore'].update(value=f'{max_val_score:.2f}')
        DataAnalysis['minTerm'].update(value=f'{min_val_Term:.2f}')
        DataAnalysis['maxTerm'].update(value=f'{max_val_Term:.2f}')
        DataAnalysis['slope'].update(value=f'{slope:.2f}')
        DataAnalysis['slopeforScore'].update(value=f'{slopeforScore:.2f}')
        DataAnalysis['slopeforterm'].update(value=f'{slopeforterm:.2f}')
        event, values = DataAnalysis.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            DataAnalysis.close()
            break
    

