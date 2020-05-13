import pandas as pd
import xlrd
import datetime
import xlwt
import openpyxl
class STUDENT(object):
    def __init__(self, name = '', group='', points = [], midpoint = int , respoint = int):
         self.name = name
         self.group = group
         self.points = points
         self.midpoint = midpoint
         self.respoint = respoint
    def show_student(self):
        discription = (self.name + " " + self.group + " middle point is " + str(self.midpoint) + " result point is " + str(self.respoint))
        print(discription)
students = []
print('Enter filename(number of group): ')
file_name = str(input())
group = file_name
file_name = "D:" + file_name + ".xlsx"
rb = xlrd.open_workbook(file_name,formatting_info=False)
if rb == 0:
    print('Open error')
sheet = rb.sheet_by_index(0)            #выбираем активный лист
row = []
row_number = sheet.nrows
col_number = sheet.ncols
print("Col: " + str(col_number))
print("Row: " + str(row_number))
#for i in range(1, sheet.nrows):
#     if sheet.cell(rowx=i, colx=1).value == ''
i = 2
kolvo = 0
#print("U1 is " + str(sheet.cell(rowx=0, colx=20).value))
while sheet.cell(rowx=0, colx=i).value != 'Средний балл' and sheet.cell(rowx=1, colx=i).value != 'Middle point':
    kolvo += 1
    i += 1
#print('Kolvo is ' + str(kolvo))
if row_number > 0:
    for i in range(1, row_number - 1):
        Student = STUDENT(sheet.cell(rowx=i+1, colx=2).value, group, sheet.cell(rowx=i+1, colx=2+kolvo+1).value,
                          sheet.cell(rowx=i+1, colx=2+kolvo+2).value)
        students.append(Student)
        r = sheet.row_values(i)
        row.append(r)
else:
        print("File is empty or incorrect")
df = pd.DataFrame(row)
df.columns = df.iloc[0]
df = df.reindex(df.index.drop(0))
print(df)