#Pre-requisites
#install openpyxl
#pip install openpyxl
#or
#py -m pip install openpyxl
#py -m pip install xlrd
#virtual enviorments can mess up with python modules, 2nd one works with pyenv
import xlrd
from xlrd import *
from tkinter import *
from tkinter import filedialog

retN = 0

def AnalizeFile():
    book = xlrd.open_workbook("2019-2-listado_6201_C1.xls")
    print("The number of worksheets is {0}".format(book.nsheets))
    print("Worksheet name(s): {0}".format(book.sheet_names()))

    sh = book.sheet_by_index(0)

    #print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
    #print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3))
    i = 11
    for i in range(11, sh.nrows):
        buffer = sh.row(i)
        print(buffer[0])
        

def GenerateReport():
    reg = 0
    quit()

def openFile():
    filepath = filedialog.askopenfilename()
    print(filepath)
    AnalizeFile()



window = Tk()
buttonOpen = Button(text="Open", command = openFile)
buttonOpen.pack()
buttonReport = Button(text="Generate Report", command = GenerateReport)
buttonReport.pack()
window.mainloop()
