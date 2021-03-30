#This is to extract data using openpyxl

import openpyxl
w=openpyxl.load_workbook("please paste the path of excel sheet here and add \ with \")
print(type(w))
print(w.active)
s1=w.active
print(type(s1))

row=s1.max_row
column=s1.max_column
print(row,"   ",column)
for i in range(1,row+1):
    for j in range(1,column+1):
        print(s1.cell(i,j).value,end=" ")
    print()
