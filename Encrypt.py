import tkinter as tk
import pandas as pd 
import numpy as np 
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showinfo
from tkinter import filedialog as fd
import random
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell

def rephrase(a):
    b=[]
    a=list(str(a))
    for item in a:
        b.append(a[random.randint(0,len(a))-1])
    return "".join(b)



class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # configure the root window
        self.title('Encrypt Data')
        self.geometry('220x100')

        # label
        self.label = ttk.Label(self, text='Hello!')
        self.label.grid(row=0,column=0,columnspan=2)

        #Buttons
        self.button = ttk.Button(self, text='Encrypt',width=20)
        self.button['command'] = self.button_clicked
        self.button.grid(row=1,column=0,pady=5,padx=50)
        
        self.button2 = ttk.Button(self,text='Select Excel file',width=20)
        self.button2.grid(row=2,column=0,pady=5,padx=50)
        self.button2['command']=self.button1_clicked

    def button1_clicked(self):
        self.source = fd.askopenfilename()
        self.workbook= load_workbook(self.source)
        self.sheet=self.workbook.active

    def button_clicked(self):
        def rephrase(a):
            b=[]
            a=list(str(a))
            for item in a:
                b.append(a[random.randint(0,len(a))-1])
            return "".join(b)
        for sheet in self.workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    cell.value=rephrase(cell.value)
        self.workbook.save("Output.xlsx")
if __name__ == "__main__":
    app = App()
    app.mainloop()