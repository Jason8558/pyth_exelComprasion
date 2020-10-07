import sys, os
import xlrd

def getSheetName(workbook, sheet):
    pointSheetObj = []
    pointSheets = workbook.sheet_names()
    sheet_name = ''
    for i in pointSheets: 
        pointSheetObj.append(str(i))
    for s in range(len(pointSheetObj)):
    	if (str(s) == str(sheet)):
    		sheet_name = str(pointSheetObj[s])    	
    return sheet_name

os.system('cls')

print ('---WELCOME TO PENTAGON---\n')
print ('-------------------------')

book1 = xlrd.open_workbook("book1.xlsx")
book2 = xlrd.open_workbook("book2.xlsx")

print('Выберите лист из книги 1')
sheet_1 = input()
sheet1 = book1.sheet_by_index(int(sheet_1))
print('Выбран лист ' + str(getSheetName(book1,sheet_1)) + '\n')

print('Выберите лист из книги 2')
sheet_2 = input()
sheet2 = book2.sheet_by_index(int(sheet_2))
print('Выбран лист ' + str(getSheetName(book2,sheet_2)) + '\n')

print('Выберите сравниваемый столбец из ' + str(getSheetName(book1,sheet_1)) + '\n')
f_col = input()
fcol = sheet1.col_values(int(f_col))
print('Выбран столбец: ' + '\n' + str(fcol[0]) + '\n')

print('Выберите сравниваемый столбец из' + str(getSheetName(book2,sheet_2)) + '\n')
s_col = input()
scol = sheet2.col_values(int(s_col))
print('Выбран столбец: ' + '\n' + str(scol[0]) + '\n')

print('Выберите сопоставляемый столбец из ' + str(getSheetName(book2,sheet_2)) + ' с данными из ' + str(getSheetName(book1,sheet_1)) )
t_col = input()
tcol = sheet2.col_values(int(t_col))
print('Выбран столбец: ' + '\n' + str(tcol[0]) + '\n')


match_list = []
for v1 in range(len(fcol)):
	for v2 in range(len(scol)):
		if (fcol[v1] == scol[v2]):
			match_list.append(fcol[v1] + '=' + tcol[v2])
match_file = open('match.txt', 'w')
for el in match_list:
	match_file.write(str(el + '\n'))
