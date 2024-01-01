# import openpyxl
# # creating a source file.
# from openpyxl import Workbook
# wb = Workbook()
# sheet = wb.active
# data = (
#     ('id','va1','va2','va3','va4','va5','va6','va7'),
#     (1,2,3,4,5,6,7,8),
# (42,30,64,92,93,85,0,93),
# (45,92,81,43,2,44,92,0),
# (32,21,0,0,0,40,84,7),
# (37,90,4,46,80,94,60,79),
# (5,56,18,98,53,0,0,0),
# (5,78,85,56,93,24,78,85),
# (27,31,26,77,55,28,89,73),
# (8,31,59,56,15,81,0,60),
# (42,0,94,20,0,30,53,0),
# (25,92,92,0,72,88,94,84),
# (10,43,53,66,49,0,46,80),
# (6,95,0,52,24,71,6,0),
# (9,33,53,0,97,68,0,19),
# (34,0,25,54,60,71,68,86),
# (19,34,74,43,0,15,40,83),
# (11,93,9,79,59,6,9,47),
# (41,97,3,10,81,93,89,0),
# (17,80,72,49,79,3,60,30),
# (44,38,0,45,0,90,0,17),
# (41,0,23,31,39,29,94,55),
# (25,21,20,87,38,7,89,87),
# (54,9,15,75,0,2,54,59),
# (45,28,40,0,91,99,32,74),
# (45,76,0,40,33,77,0,35),
# (4,45,80,9,71,18,26,0),
# (3,0,40,59,0,0,62,49),
# (3,70,74,69,92,95,56,95),
# (50,44,75,15,90,82,0,72),
# (4,57,8,26,34,68,30,38),
# (62,49,95,38,20,82,46,94),
# (38,18,67,24,97,17,24,0),
# (3,99,17,2,26,20,51,44),
# (21,72,9,0,62,6,1,34),
# (30,61,31,36,11,47,35,70),
# (1,69,67,39,98,12,44,91),
# (61,0,32,85,92,82,33,89),
# (39,80,91,11,15,25,83,0),
# (29,18,12,24,11,37,64,98),
# (33,36,62,63,70,27,61,43),
# (13,65,62,88,41,53,14,20),
# (41,6,13,53,26,67,86,75),
# (44,6,24,97,61,1,1,34),
# (30,7,68,47,72,67,90,0),
# (43,0,13,24,89,67,58,84),
# (51,0,10,89,95,87,14,12),
# (51,67,14,0,90,0,19,29),
# (6,0,62,15,13,18,0,87),
# (42,7,3,2,87,7,20,32),
# (37,37,80,81,18,39,30,83),
# (53,71,95,67,56,35,82,0),
# (33,34,84,88,93,25,86,53),
# (23,4,68,57,0,65,37,20),
# (28,100,0,85,97,6,39,90),
# (10,36,8,0,18,27,76,57),
# (52,8,60,74,4,0,53,54),
# (10,41,42,35,68,94,18,0),
# (62,0,43,73,76,75,43,55),
# (49,24,0,0,63,48,0,74),
# (35,46,57,62,57,94,16,44),
# (14,100,72,2,0,22,35,21)
# )
# for i in data:
#     sheet.append(i)
# wb.save('source.xlsx')

# creating a lines file.
# import openpyxl
# from openpyxl import load_workbook,Workbook
# wb = load_workbook('source.xlsx')
# sheet = wb.active
# wb1 = Workbook()
# sheet1 = wb1.active
# hl = ['price','servicetype','item_code']
# sheet1.append(hl)
# pl = []
# for i in sheet.iter_cols(min_row = 2,min_col = 2,max_row = sheet.max_row,max_col = sheet.max_column):
#     for j in i:
#         pl.append(j.value)
# column_index = 1
# # Start appending data from row 2
# start_row = 2
# for index, value in enumerate(pl):
#     sheet1.cell(row=start_row + index, column=column_index, value=value)
# zl = []
# for a in sheet.iter_cols(min_row = 1,min_col = 2,max_row = 1,max_col = sheet.max_column):
#     for b in a:
#         zl.append(b.value)
# row_count = sheet.max_row
# column_letter1 = 'B'
# start_row = 2
# for value in zl:
#     for _ in range(row_count -1):
#         sheet1[f'{column_letter1}{start_row}'] = value
#         start_row += 1
# il = []
# for d in sheet.iter_rows(min_row = 2,min_col = 1,max_row = sheet.max_row,max_col = 1):
#     for e in d:
#         il.append(e.value)
# column_letter4 = 'C'
# start_row = 2
# column_count = sheet.max_column
# for _ in range(column_count - 1):
#     for value4 in il:
#         sheet1[f'{column_letter4}{start_row}'] = value4
#         start_row += 1
# wb1.save('linesandattributesfile.xlsx')









