from msilib import text
from msilib.schema import ComboBox
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.ttk import Combobox
import openpyxl

import stockDataApp

class gui():
    filepath = "" 
    sheetName = ""
    sheetList = []
    log = ""

    def __init__(self) -> None:

        def runFunc():
            self.sheetName = sheetSelection.get()
            if(self.sheetName=="Select" or ""): logInfo['text'] = logInfo['text'] + '\nMissing Sheet Selection'
            stockDataApp.fillRawData(self.filepath,self.sheetName)
            print("Completed")
            root.destroy()
        def excelCheck(filepath):
            try: 
                wb = openpyxl.load_workbook(filepath)
                sheetSelection['values'] = wb.sheetnames
               
            except:
                logInfo['text'] = "File selected is not an excel"
                
        def updateFilePath():
            self.filepath = askopenfilename()
            excelEntry['text'] = self.filepath
            excelCheck(self.filepath)

        root = Tk(className = " Stock Data Application")
        root.configure(background='white')
        #root.geometry('500x200')
        frame = Frame(root,bg='white')
        frame.grid()

        titleLabel = Label(frame, text='Stock Data Application',bg='white',anchor=CENTER)
        titleLabel.grid(column=0,row=0)

        fileButton = Button(frame, text='Select Excel File',command= updateFilePath,bg='white')
        fileButton.grid(column=0,row=1)

        excelEntry = Label(frame,text='Filepath appears here',bg='white')
        excelEntry.grid(column=1,row=1) 

        sheetLabel = Label(frame,text="Select sheet", bg='white')
        sheetLabel.grid(column=0,row=2)

        variable = StringVar(root)
        variable.set("Select") # default value
        sheetSelection = Combobox(frame,textvariable=variable, background='white')
        sheetSelection.grid(column=1,row=2)
        

        logBox = Label(frame, text="Feedback: ",bg='white')
        logBox.grid(column=0, row=3)

        logInfo = Label(frame, text=self.log, bg='white')
        logInfo.grid(column=1,row=3)

        runButton = Button(frame,text="Run", width=5, command=runFunc)
        runButton.grid(column=3, row=3)
        root.mainloop()



#############################################################################################################################
# Main method
#############################################################################################################################

if __name__ =='__main__':
    gui()

