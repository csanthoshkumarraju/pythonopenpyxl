# pythonopenpyxl

import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
# creating workbook
wb = openpyxl.Workbook()
wb.save('firstpythonexcel.xlsx')
# opening a workbook
wb = openpyxl.load_workbook('Untitled spreadsheet.xlsx')
for sheet in wb:
    print(sheet.tittle)
# writing ,creating deleting saving
wb1 = Workbook()
ws1 = wb1.create_sheet('A sheet',0)
ws2 = wb1.create_sheet('B sheet',0)
for sheet in wb1:
    print(sheet.tittle)
wb1.remove((wb['A sheet']))
for sheet in wb1:
    print(sheet.tittle)
del wb1['B sheet']
for sheet in wb1:
    print(sheet.tittle)
wb1.save('create_sheets.xlsx')
# copying sheets
wb2 = load_workbook('create_sheets.xlsx')
for sheet in wb2:
    print(sheet.tittle)
source = wb2["sheet1"]
new_sheet = wb.copy_worksheet(source)
for sheet in wb2:
    print(sheet.tittle)
wb1.save('copied_sheet.xlsx')
#  getting sheet by index and name
wb3 = load_workbook('create_sheets.xlsx')
for sheet in wb2:
    print(sheet.tittle)
ws3 = wb3.worksheets[0]
ws4 = wb3.worksheets[1]
worksheets = wb.sheetnames
print(type(worksheets))
print(worksheets)
# READING CELLS
wb4 = openpyxl.Workbook()
ws5 = wb4.worksheets[0]
ws5["A1"].value = 56
ws5["B1"].value = 57
ws5["C1"].value = 58
print(ws5["A1"].value) #hit enter to print
print(ws5["B1"].value)
print(ws5["C1"].value)
ws5.cell(row = 2,column =1).value = 1234
ws5.cell(row = 2,column =2).value = 12345
ws5.cell(row = 2,column =3).value = 123456
print(ws5["A2"].value)
print(ws5["B2"].value)
print(ws5["C2"].value)
#  getting column name
from openpyxl.utils.cell import get_column_letter
get_column_letter(3)
get_column_letter(13)
# offset
wb5 = Workbook()
ws6 = wb5.worksheets[0]
ws5["A3"].value = "Train"
ws5.cell(1,1).offset(0,1).value = "train_Cart"
print(ws5["A3"].value, ws5["B1"].value)
# reading range of cells
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
def save_wb(wb6, filename):
    wb6.save(filename)
def create_sheets(wb6, sheet_name_list):
    for sheet1 in sheet_name_list:
        wb6.create_sheet(sheet1)
if __name__ == "__main__":
    filename = "readingrangeofcells.xlsx"
    wb6 = Workbook()
    create_sheets(wb6, ["sheet2","sheet3","sheet4"])
    wsl = wb6["sheet"]
    wsl["A1"] = "A1"
    wsl["A2"] = "A2"
    wsl["B1"] = "B1"
    wsl["B2"] = "B2"
    wsl["C1"] = "C1"
    wsl["C2"] = "C2"
    cell_range = wsl['A1':'C2']
    for c1,c2 in cell_range:
        print(c1.value,c2.value)
    print()
#  looping over cells
import openpyxl
from openpyxl import Workbook
if __name__ == "__main__":
    filename = "loopingcells.xlsx"
    wb7 = Workbook()
    wsl = wb7["sheet"]
    for row in range(1,3):
        for col in range(1,3):
            cell = wsl.cell(row = col , column = row)
            print(cell.coordinate,end='')
        print()
    wb7.save(filename)

#  iter rows and columns
import openpyxl
from openpyxl import Workbook
def create_sheets(wb8, sheet_name_list):
    for sheet_name in sheet_name_list:
        wb.create_chartsheet(sheet_name)

if __name__ == "__main__":
    filename = "iterrowsandcolumns.xlsx"
    wb8 = Workbook()
    create_sheets(wb8, ["sheet2", "sheet3", "sheet4"])
    wsl = wb8["sheet"]
    print("iterated rows")
    for row in wsl.iter_rows(min_row=1, min_col =1,max_col =4, max_row =3):
        for cell in row:
            print(cell.coordinate,end=" ")
        print()
    print("_"* 40)
    print("iterated cols:")
    for column in wsl.iter_cols(max_col =4, max_row =3):
        for cell in column:
            print(cell.coordinate,end=" ")
        print()
    wb8.save(filename)
# delete and insert rows
from openpyxl import workbook
def set_values(ws):
    ws.delete_cols(1,100)
    counter = 1
    for row in ws.iter_rows(min_row = 1, max_col=ws.max_column,max_row=ws.max-row):
        for cell in row:
            cell.value = counter
            counter += 1
def print_rows(ws):
    row_string = ""
    for row in ws.iter_rows(min_row=1,max_col=ws.max_column,max_row=ws.max_row):
        for cell in row:
            row_string += "{:<3}".format(str(cell.value) + ' ')
        row_string  += "\n"
    print(row_string)
if __name__ == "__main__":
    filename = "deleteinsertrowsandcolumns.xlsx"
    wb9 = Workbook()
    wsl = wb9["sheet"]
    set_values(wsl)
    print_rows(wsl)
wsl.insert_rows(0)
print_rows(wsl)
wsl.delete_rows(0)
print_rows(wsl)
wsl.insert_cols(4)
print_rows(wsl)
set_values(wsl)
print_rows(wsl)
wsl.delete_rows(0,4)
print_rows(wsl)
wb9.save(filename)

# append methods for rows
from openpyxl import Workbook

def print_rows (WS):
    row_string=""
    for row in ws.iter_rows (min_row=1, max_col=ws.max_column, max_row=ws.max_row) :
        for cell in row:
            row_string + "(:<8)".format (str (cell.value) + " ")
        row_string += "\n"
    print (row_string)
if __name__ == "__main__":
    filename = "Append.xlsx"
    wb = Workbook ()
    ws1 = wb ["Sheet"]
    print_rows (wsl)
sales_data = [["North", 670_000, 230_000],\["West", 111_000, 95_000],
["South", 340_000, 550_000],\
["East", 456_000, 123_000]]
wsl.append( ["Sales", 2018, 2019] )
print_rows (wsl)
for row in sales_data:
    wsl.append(row)
print_rows (wsl)
wb.save (filename)
# 




