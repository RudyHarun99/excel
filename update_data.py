import xlrd
file=xlrd.open_workbook('data.xlsx')
sheet=file.sheet_by_name('DataKaryawan')

cols=[]
for i in range(sheet.nrows):
    cols.append(sheet.row_values(i))
# print(cols)

x=int(input('Nomor : '))
y=input('Nama : ')
z=input('Kota : ')
a=[x,y,z]
# print(a)
cols.append(a)
# print(cols)

import xlsxwriter
file=xlsxwriter.Workbook('data.xlsx')
sheet=file.add_worksheet('DataKaryawan')

# write data
row=0
for x,y,z in cols:
    sheet.write(row,0,x)
    sheet.write(row,1,y)
    sheet.write(row,2,z)
    row+=1
file.close()