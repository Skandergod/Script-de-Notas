#Pre-requisites
#install openpyxl
#pip install openpyxl
#or
#py -m pip install openpyxl
#virtual enviorments can mess up with python modules, 2nd one works with pyenv
import openpyxl
from openpyxl import Workbook, load_workbook

wb = Workbook()

ws = wb.active

wb.save("sample.xlsx")

print("This line will be printed.") 