# pythonopenpyxl
source document //
https://openpyxl.readthedocs.io
document //
 # ** I have learnt this codes from Conny Soderholm **
 # ** these codes for my refrence only These arec not my own codes took from Conny Soderholm Udemy course. **
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
# #!/usr/bin/env python3

"""
Moving and copying ranges of cells
"""

from openpyxl import Workbook

def set_values(ws):
    ws.delete_cols(1,100)
    counter = 1
    for row in ws.iter_rows(min_row=1, max_col=10, max_row=10):
        for cell in row:
            cell.value = counter
            counter += 1

def print_rows(ws):
    row_string = ""
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            row_string += "{:<6}".format(str(cell.value) + " ")
        row_string += "\n"
    print(row_string)

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Moving_copying_ranges.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]

    # Insert values from 1 to 100 into a grid of 10x10 cells
    set_values(ws1)
    print_rows(ws1)
    print("*"*30)

##    # move the whole range ten rows down
##    ws1.move_range("A1:J10", rows=10, cols=0)
##
##    print_rows(ws1)

##    #Move cell A1 30 rows down
##    #ws1._move_cell(row, column, row_offset, col_offset)
##    ws1._move_cell(1, 1, 30, 0)
##    print_rows(ws1)
##
##    # reset values
##    set_values(ws1)
##
##    print_rows(ws1)
##
##    # copy cell A1's value
##    old_cell = ws1.cell(row=1, column=1)
##    new_cell = ws1.cell(row=12, column=1, value= old_cell.value)
##
##    print_rows(ws1)
    
##     Copy the first row to row 15
    rows = ws1.iter_rows(min_row=0, max_row=1)

    for row in rows:
        for cell in row:
            new_cell = ws1.cell(row=15, column=cell.col_idx, value= cell.value)

    print_rows(ws1)

   #!/usr/bin/env python3

"""
Insert formulas to the spreadsheet
"""

from openpyxl import Workbook

def print_rows(ws):
    row_string = ""
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            row_string += "{:<16}".format(str(cell.value) + " ")
        row_string += "\n"
    print(row_string)

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Formulae.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]

    """
    NB you must use the English name for a function
    and function arguments must be separated by commas
    and not other punctuation such as semi-colons.
    """

    # Insert values
    for i in range(1,11):
        ws1.cell(row=i, column=1).value = i*i
        ws1.cell(row=i,column=2).value = i/2

    print_rows(ws1)
    #=SUM(A1:A10)

    # Define the first and last cell used in the formula
    first_cell = ws1.cell(row=1, column=1) #A1
    last_cell = ws1.cell(row=10,column=1) #A10

    # Create the formula =SUM(A1:A10)
    ws1.cell(row=11, column=1).value = "=SUM(" +str(first_cell.coordinate) + ":" + str(last_cell.coordinate) +")"
    print_rows(ws1)

    

    # Move the formula one step to the right, and transpose the formula
    ws1._move_cell(row=11,column=1,row_offset=0,col_offset=1,translate=True)
    print_rows(ws1)
    
    wb.save(filename)
  #!/usr/bin/env python3

"""
Creating tables
"""

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

def print_rows(ws):
    row_string = ""
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            row_string += "{:<6}".format(str(cell.value) + " ")
        row_string += "\n"
    print(row_string)

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Tables.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]
    
    # Create Sales data, a list of lists
    # Using underscore for easier separation of thousands
    sales_data = [["North",  670_000, 230_000],\
                  ["South", 340_000, 550_000],\
                  ["West", 111_000, 95_000],
                  ["East", 456_000, 123_000]]

    # add column headings. String type only
    ws1.append(["Sales", "2018", "2019"])
    for row in sales_data:
        ws1.append(row)

    print_rows(ws1)

    # Create a table. Remember that all headers need to be of string type, None type is not acceptable
    sales_table = Table(displayName="SalesTable", ref="A1:C5")

    # Add a default style
    style = TableStyleInfo(name="TableStyleMedium8", showRowStripes=True)
    sales_table.tableStyleInfo = style
    # Add the table to the sheet
    ws1.add_table(sales_table)

    print_rows(ws1)

    wb.save(filename)
    #!/usr/bin/env python3

"""
Cell formatting
"""

from openpyxl import load_workbook, Workbook
from copy import copy
from openpyxl.styles import colors, Font
from openpyxl.styles.fills import PatternFill
from openpyxl.styles.borders import Border, Side, BORDER_THIN, BORDER_THICK, BORDER_DASHDOT, BORDER_DOUBLE
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from openpyxl.styles.protection import Protection
from openpyxl.styles.alignment import Alignment

def print_rows(ws):
    row_string = ""
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            row_string += "{:<3}".format(str(cell.value) + " ")
        row_string += "\n"
    print(row_string)

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Cell formatting.xlsx"
    wb = load_workbook("Cell formatting_original.xlsx")
    ws1 = wb["Sheet"]

    # Insert values from 1 to 100 into a grid of 10x10 cells
    print_rows(ws1)    

    font_cell = ws1.cell(row=2, column=2)           #B2
    border_cell = ws1.cell(row=3, column=4)         #D3
    fill_cell = ws1.cell(row=4, column=6)           #F4
    number_format_cell = ws1.cell(row=5, column=8)  #H5
    alignment_cell = ws1.cell(row=8,column=9)       #I8
    # See openpyxl.styles.colors for a list of colors
    # See openpyxl.styles.fonts for all elements of Font
    font_cell.font = Font(name='Arial', size=18, b=True, i=True, color=colors.COLOR_INDEX[12])
    
    # openpyxl.styles.borders
    borders = Border(left=Side(border_style=BORDER_THIN, color='00000000'),\
                         right=Side(border_style=BORDER_THICK, color='00000000'),\
                         top=Side(border_style=BORDER_DASHDOT, color='00000000'),\
                         bottom=Side(border_style=BORDER_DOUBLE, color='00000000'))
    border_cell.border = borders
    

    # openpyxl.styles.fills
    red_color = colors.Color(rgb='00FF0000')
    solid_red_fill = PatternFill(patternType='solid', fgColor=red_color)
    fill_cell.fill = solid_red_fill
    

    # openpyxl.styles.numbers
    number_format_cell.number_format = FORMAT_PERCENTAGE_00
    
    # Create a Workbook password and lock the structure
    wb2 = Workbook()
    wb2.security.workbookPassword = '1234'
    wb2.security.lockStructure = True
    wb2.save("Cell formatting2.xlsx")

    # Protect the sheet
    ws1.protection.sheet = True
    ws1.protection.password = '1234'
    ws1.protection.enable()

    
    # Unlock cell A1
    ws1.cell(row=1,column=1).protection = Protection(locked=False, hidden=False)
    wb.save(filename)

    # openpyxl.styles.alignment
    
##    alignment_cell.alignment = Alignment(horizontal="center", vertical=None, textRotation=0,\
##                                         wrapText=None, shrinkToFit=None, indent=0,\
##                                         relativeIndent=0, justifyLastLine=None,\
##                                         readingOrder=0, text_rotation=None,\
##                                         wrap_text=None, shrink_to_fit=None, mergeCell=None)
##
##    wb.save(filename)
##            
##
##            

 #!/usr/bin/env python3

"""
Copying cell formatting
https://openpyxl.readthedocs.io/en/stable/styles.html?highlight=styles
"""

from openpyxl import Workbook
from copy import copy
from openpyxl.styles import Font


def set_values(ws):
    ws.delete_cols(1, 100)
    counter = 1
    for row in ws.iter_rows(min_row=1, max_col=10, max_row=10):
        for cell in row:
            cell.value = counter
            counter += 1


def print_rows(ws):
    row_string = ""
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            row_string += "{:<3}".format(str(cell.value) + " ")
        row_string += "\n"
    print(row_string)


if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Moving_copying_ranges.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]

    # Insert values from 1 to 100 into a grid of 10x10 cells
    set_values(ws1)
    print_rows(ws1)

    # move the whole range ten rows down
    ws1.move_range("A1:J10", rows=10, cols=0)

    print_rows(ws1)

    # reset values
    set_values(ws1)

    print_rows(ws1)

    # copy cell A1's value and formatting to cell A12
    old_cell = ws1.cell(row=1, column=1)
    old_cell.font = Font(name="Arial", size=18, color="FF0000")
    # old_cell.font = Font(name='Arial', size=18, color="FF0000")
    new_cell = ws1.cell(row=12, column=1, value=ws1.cell(row=1, column=1).value)
    new_cell.font = copy(old_cell.font)
    new_cell.border = copy(old_cell.border)
    new_cell.fill = copy(old_cell.fill)
    new_cell.number_format = copy(old_cell.number_format)
    new_cell.protection = copy(old_cell.protection)
    new_cell.alignment = copy(old_cell.alignment)

    wb.save(filename)          
    
#!/usr/bin/env python3

"""
Merging cells
"""

from openpyxl import Workbook

def set_values(ws):
    ws.delete_cols(1,100)
    counter = 1
    for row in ws.iter_rows(min_row=1, max_col=10, max_row=10):
        for cell in row:
            cell.value = counter
            counter += 1

def print_rows(ws):
    row_string = ""
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            row_string += "{:<3}".format(str(cell.value) + " ")
        row_string += "\n"
    print(row_string)

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Merge_cells.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]

    # Insert values from 1 to 100 into a grid of 10x10 cells
    set_values(ws1)
    # Merge cells A1:B1
    ws1.merge_cells('A1:B1')
    wb.save(filename)
    #ws1.unmerge_cells('A2:D2')
    # Merge cells A4:C4
    #ws1.merge_cells(start_row=4, start_column=1, end_row=4, end_column=3)
    
    print_rows(ws1)

    wb.save(filename)
     #!/usr/bin/env python3

"""
Filtering
"""

from openpyxl import Workbook

def set_values(ws):
    data = [
    ["Fruit", "Quantity"],
    ["Kiwi", 1],
    ["Grape", 15],
    ["Apple", 3],
    ["Peach", 6],
    ["Pomegranate", 3],
    ["Pear", 7],
    ["Tangerine", 4],
    ["Blueberry", 58],
    ["Mango", 3],
    ["Watermelon", 19],
    ["Blackberry", 3],
    ["Orange", 25],
    ["Raspberry", 9],
    ["Banana", 7]
    ]
    for r in data:
        ws.append(r)

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Filtering.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]
    # Insert values from 1 to 100 into a grid of 10x10 cells
    set_values(ws1)

    # Set autofilter
    ws1.auto_filter.ref = "A1:B15"
    #ws1.auto_filter.add_filter_column(0, ["Kiwi", "Apple", "Mango"])
    ws1.auto_filter.add_sort_condition("B2:B15", descending=False)
    # OpenPyXL does not apply the filter, you have to do that manually.
    # This will add the relevant instructions to the file but will neither actually filter nor sort.

    wb.save(filename)
     #pip install pypiwin32 if you haven't installed the module
import win32com.client
import os

xl = win32com.client.Dispatch("Excel.Application")

workbook_name = "sort_workbook.xlsx"
absolute_path = os.path.abspath(workbook_name)


wb = xl.Workbooks.open(absolute_path)


ws = wb.Worksheets('Sheet1')

ws.Range('A1:A10').Sort(Key1=ws.Range('A1'), Order1=1, Orientation=1)

wb.Save()
xl.Application.Quit()
#!/usr/bin/env python3

"""
Freeze panes
"""

from openpyxl import load_workbook

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Freeze_panes.xlsx"
    wb = load_workbook(filename)
    ws1 = wb["Sheet1"]

    # Freeze the top row
    ws1.freeze_panes = "A1"

    """

    freeze_panes settings             Rows and columns frozen
    
    sheet.freeze_panes = 'A2'         Row 1

    sheet.freeze_panes = 'B1'         Column A

    sheet.freeze_panes = 'C1'         Columns A and B

    sheet.freeze_panes = 'C2'         Row 1 and columns A and B

    sheet.freeze_panes = 'A1' or      No Frozen panes
    sheet.freeze_panes = None
    """

    wb.save(filename)
            
#!/usr/bin/env python3

"""
Freeze panes
"""

from openpyxl import load_workbook

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Freeze_panes.xlsx"
    wb = load_workbook(filename)
    ws1 = wb["Sheet1"]

    # Freeze the top row
    ws1.freeze_panes = "A1"

    """

    freeze_panes settings             Rows and columns frozen
    
    sheet.freeze_panes = 'A2'         Row 1

    sheet.freeze_panes = 'B1'         Column A

    sheet.freeze_panes = 'C1'         Columns A and B

    sheet.freeze_panes = 'C2'         Row 1 and columns A and B

    sheet.freeze_panes = 'A1' or      No Frozen panes
    sheet.freeze_panes = None
    """

    wb.save(filename)
            
#!/usr/bin/env python3

"""
Page setup
"""

from openpyxl import load_workbook
from openpyxl.worksheet.page import PrintPageSetup

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Page_setup.xlsx"
    wb = load_workbook(filename)
    ws1 = wb["Sheet"]

    # openpyxl.worksheet.page module
    # Parameters from "Source code for openpyxl.worksheet.worksheet"
    ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
    ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
    ws1.page_setup.fitToHeight = 0
    ws1.page_setup.fitToWidth = 1

    ws1.page_setup = PrintPageSetup(worksheet=None, orientation=ws1.ORIENTATION_PORTRAIT, paperSize=ws1.PAPERSIZE_LETTER,\
                                    scale=None, fitToHeight=None, fitToWidth=None, firstPageNumber=None,\
                                    useFirstPageNumber=None, paperHeight=None, paperWidth=None, pageOrder=None,\
                                    usePrinterDefaults=None, blackAndWhite=None, draft=None, cellComments=None,\
                                    errors=None, horizontalDpi=None, verticalDpi=None, copies=None, id=None)


    wb.save(filename)
      #!/usr/bin/env python3

"""
Fold (Outline)
"""

from openpyxl import Workbook

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Fold.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]

    ws1.column_dimensions.group('A','D', hidden=True)
    ws1.row_dimensions.group(1,10, hidden=True)

    wb.save(filename)
            
#!/usr/bin/env python3

"""
Charts
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, ScatterChart, PieChart, Reference, Series
from openpyxl.chart.series import DataPoint
from random import randint

def set_values(ws):
    counter = 2
    for column in ws1.iter_cols(min_row=2, min_col=1, max_col=1, max_row=11):
        for cell in column:
            cell.value = counter
            counter += 2
    for row in ws.iter_rows(min_row=2, min_col = 2, max_col=4, max_row=11):
        for cell in row:
            cell.value = randint(0,500)

def print_rows(ws):
    row_string = ""
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            row_string += "{:<10}".format(str(cell.value) + " ")
        row_string += "\n"
    print(row_string)

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Charts.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]

    # Insert values from 1 to 100 into a grid of 10x10 cells
    headers = ["Number", "Torque", "Power", "Consumption"]
    ws1.append(headers)
    set_values(ws1)
    print_rows(ws1)

    series1 = Reference(ws1, min_col=1, min_row=1, max_col=1, max_row=11)
    series2 = Reference(ws1, min_col=2, min_row=1, max_col=2, max_row=11)
    series3 = Reference(ws1, min_col=3, min_row=1, max_col=3, max_row=11)
    series4 = Reference(ws1, min_col=4, min_row=1, max_col=4, max_row=11)
    series5 = Reference(ws1, min_col=2, min_row=2, max_col=2, max_row=6)

    # Add a Bar chart
    bar_chart = BarChart()
    bar_chart.add_data(series1, titles_from_data=True)
    bar_chart.add_data(series2, titles_from_data=True)
    bar_chart.title = "Bar Chart"
    bar_chart.style = 11
    bar_chart.x_axis.title = 'Size'
    bar_chart.y_axis.title = 'Percentage'
    ws1.add_chart(bar_chart, "A16")
    
    # Add a Scatter chart
    scatter_chart = ScatterChart()
    scatter_chart.title = "Scatter Chart"
    scatter_chart.style = 14
    scatter_chart.x_axis.title = 'Size'
    scatter_chart.y_axis.title = 'Percentage'
    series = Series(series1, series2, title_from_data=True)
    scatter_chart.series.append(series)
    ws1.add_chart(scatter_chart, "G1")
    
    # Add a Pie chart
    pie_chart = PieChart()
    labels = Reference(ws1, min_col=1, min_row=1, max_col=4, max_row=1)
    pie_chart.add_data(series5, titles_from_data=True)
    pie_chart.set_categories(labels)
    pie_chart.title = "Pie Chart"

    # Cut the first slice out of the pie
    pie_slice = DataPoint(idx=0, explosion=40)
    pie_chart.series[0].data_points = [pie_slice]
    ws1.add_chart(pie_chart, "K16")

    wb.save(filename)
            

    #!/usr/bin/env python3

"""
Chartsheet
"""

from openpyxl import Workbook
from openpyxl.chart import AreaChart, Reference

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Chartsheets.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]
    # Create a chartsheet
    cs = wb.create_chartsheet()

    rows = [
        ["Bricks", 3],
        ["Tiles", 2],
        ["Blocks", 4],
        ["Grass", 8],
        ["Plates", 8],
        ["Soil", 1],
    ]

    for row in rows:
        ws1.append(row)

    # Titles
    chart = AreaChart()
    chart.title = "Area Chart"
    chart.style = 13
    chart.x_axis.title = 'Item'
    chart.y_axis.title = 'Share of area'

    # Add the data to the chart
    data = Reference(ws1, min_col=2, min_row=1, max_row=6)
    categories = Reference(ws1, min_col=1, min_row=1, max_row=6)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(categories)
    # Add the chart to the chartsheet
    #cs.add_chart(chart)

    wb.save(filename)
            
#!/usr/bin/env python3

"""
Inserting images
"""

from openpyxl import Workbook
from openpyxl.drawing.image import Image
# Remember to install Pillow; pip install Pillow

if __name__ == "__main__":
    # Create a workbook and sheets
    filename = "Images.xlsx"
    wb = Workbook()
    ws1 = wb["Sheet"]

    # create an image
    img = Image('logo.png')
    # add to worksheet and anchor next to cells
    ws1.add_image(img, 'A1')
    print("Image added to workbook")

    wb.save(filename)
            
#!/usr/bin/env python3

"""
Open file and save file dialogs with tkinter
Icon by Paulo Ruberto
http://www.iconarchive.com/show/custom-round-yosemite-icons-by-pauloruberto/Python-icon.html

If you are using VSCode:
1) add root.mainloop() as the last line in the script, and 
2) comment out root.withdraw() as you will need to 'X' out of this window otherwise VSCode will hang.
Thank you James Strayer for the comment about this!
"""
import tkinter as tk
from tkinter import filedialog
from os import getcwd
from openpyxl import load_workbook

if __name__ == "__main__": 
    root = tk.Tk()
    root.withdraw() # Hides the root window
    #root.wm_iconbitmap('py.ico')
    
    root.filename =  filedialog.asksaveasfilename(initialdir=getcwd(),\
                                                  title = "Select file",\
                                                  filetypes = (("png files","*.png"),\
                                                               ("all files","*.*")))

    print ("Save as file path:", root.filename)

##    root.filename =  filedialog.askopenfilename(initialdir=getcwd(),\
##                                                title = "Select file",\
##                                                filetypes = (("Excel workbooks","*.xlsx"),\
##                                                             ("all files","*.*")))
##    print ("Open file path:", root.filename)
##
##    
##    wb = load_workbook(root.filename)
##    print("Sheets in workbook", wb.worksheets)

    #!/usr/bin/env python3

"""
Creating files and folders
"""
import os
# Get the path
path = os.getcwd()  

if __name__ == "__main__":
    content = ["This", "is", "a", "test", "file"]
    with open("testfile.txt", "w") as f: 
        for word in content: 
            f.write(word+"\n")

    with open("testfile.txt", "r") as f:
        for line in f:
            line = line.strip()
            print(line)
            try:  
                os.mkdir(path+"\\"+line)
            except OSError:  
                continue
                #print("Cannot create folder", path+"\\"+line)
            finally:
                print("Finishing up...")
            #!/usr/bin/env python3

"""
Creating files and folders
"""
import os
# Get the path
path = os.getcwd()  

if __name__ == "__main__":
    content = ["This", "is", "a", "test", "file"]
    with open("testfile.txt", "w") as f: 
        for word in content: 
            f.write(word+"\n")

    with open("testfile.txt", "r") as f:
        for line in f:
            line = line.strip()
            print(line)
            try:  
                os.mkdir(path+"\\"+line)
            except OSError:  
                continue
                #print("Cannot create folder", path+"\\"+line)
            finally:
                print("Finishing up...")
            

    #!/usr/bin/env python3

"""
Getting paths and filenames.
"""
import pathlib

files= ["1.txt", "2.txt", "3.txt", "4.txt"]

if __name__ == "__main__":
    for i, file_path in enumerate(files):
        path = pathlib.Path.joinpath(pathlib.Path.cwd(), files[i])
        with open(path, "r") as f:
            print(f.read())
#!/usr/bin/env python3

"""
Timing your scripts
"""
import time

if __name__ == "__main__":
    start = time.time()
    a = range(100_000_00)
    b = []
    for i in a:
        b.append(i*2)
    end = time.time()
    print("The script took", f"{end - start:.3f}", "seconds to complete")      
    #!/usr/bin/env python3

"""
Start Excel App for viewing your sheets
"""
import os

if __name__ == "__main__":
    # Open workbook
    file_name = "8.6_Open_me.xlsx"
    os.system("start EXCEL.EXE " + file_name)
            
          

    

    
    

           

    
             

    




