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
        discription = (str(self.name) + " " + str(self.group) + " middle point is " + str(self.midpoint) + " result point is " + str(self.respoint))
        print(discription)
students = []
print('Enter the way to your file(without name of this file): ')
way = str(input())
print('Enter filename(number of group): ')
file_name = str(input())
group = file_name
file_name = way + file_name + ".xlsx"
rb = xlrd.open_workbook(file_name,formatting_info=False)
if rb == 0:
    print('Open error')
sheet = rb.sheet_by_index(0)            #выбираем активный лист
row = []
row_number = sheet.nrows
col_number = sheet.ncols
#print("Col: " + str(col_number))
#print("Row: " + str(row_number))
#for i in range(1, sheet.nrows):
#     if sheet.cell(rowx=i, colx=1).value == ''
i = 2
kolvo = 0
#print("U1 is " + str(sheet.cell(rowx=0, colx=20).value))
while sheet.cell(rowx=0, colx=i).value != 'Средний балл' and sheet.cell(rowx=1, colx=i).value != 'Middle point':
    kolvo += 1
    i += 1
#print('Kolvo is ' + str(kolvo))
k = 0
if row_number > 0:
    for i in range(1, row_number):
        array = [0]*kolvo
        j = 1
        k = 0
        while sheet.cell(rowx=0, colx=j+1).value != 'Средний балл' and sheet.cell(rowx=1, colx=j).value != 'Middle point':
            j += 1
            array[k] = str(sheet.cell(rowx=i, colx=j).value)
            k+=1
        Student = STUDENT(sheet.cell(rowx=i, colx=1).value, group, array, sheet.cell(rowx=i, colx=kolvo+2).value,
                          sheet.cell(rowx=i, colx=1+kolvo+2).value)
        students.append(Student)
        Student.show_student()
        #r = sheet.row_values(i)
        #row.append(r)
        array.clear()
else:
        print("File is empty or incorrect")
#df = pd.DataFrame(row)
#df.columns = df.iloc[0]
#df = df.reindex(df.index.drop(0))
#print(df)