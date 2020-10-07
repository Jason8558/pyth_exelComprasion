import sys, os
import xlrd
from PyQt5 import QtWidgets
from gui import Ui_MainWindow

class mywindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(mywindow, self).__init__()

        os.system('cls')
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.book1_open.clicked.connect(self.SelectBook)
        self.ui.book2_open.clicked.connect(self.SelectBook2)
        self.ui.sheet1_list.currentTextChanged.connect(self.SelectColumnFromSheet1)
        self.ui.sheet2_list.currentTextChanged.connect(self.SelectColumnFromSheet2)
        self.ui.accept_cols.clicked.connect(self.acceptColumnFromSheet)
        self.ui.compare.clicked.connect(self.Compare)

    def SelectBook(self):
        self.ui.sheet1_cols.clear()
        self.ui.sheet1_list.clear()

        b_name = QtWidgets.QFileDialog.getOpenFileName(self,'','',"exel (*.xlsx *.xls)")[0]

        if (b_name == ''):
            self.ui.book1_filename.setText('Файл не выбран!')
        else:
            self.ui.book1_filename.setText(b_name)
            o_book = xlrd.open_workbook(b_name)
            sheets = []
            sheets = o_book.sheet_names()
            self.ui.sheet1_list.addItems(sheets)
            o_book.release_resources()

    def SelectBook2(self):
        self.ui.sheet2_cols1.clear()
        self.ui.sheet2_cols2.clear()
        self.ui.sheet2_list.clear()

        b_name = QtWidgets.QFileDialog.getOpenFileName(self,'','',"exel (*.xlsx *.xls)")[0]

        if (b_name == ''):
            self.ui.book2_filename.setText('Файл не выбран!')
        else:
            self.ui.book2_filename.setText(b_name)
            o_book = xlrd.open_workbook(b_name)
            sheets = []
            sheets = o_book.sheet_names()
            self.ui.sheet2_list.addItems(sheets)
            o_book.release_resources()

    def SelectColumnFromSheet1(self):
        fname = self.ui.book1_filename.text()
        sname = self.ui.sheet1_list.currentText()
        if (fname != '') and (fname != 'Файл не выбран!') and (sname != ''):
            o_book = xlrd.open_workbook(fname)
            sheet1 = o_book.sheet_by_name(self.ui.sheet1_list.currentText())
            sheet1_cols_list = []
            for c in range(sheet1.ncols):
                cname = sheet1.col_values(c)
                sheet1_cols_list.append(str(cname[0]) + ' (' + str(c) + ')')
            self.ui.sheet1_cols.clear()
            self.ui.sheet1_cols.addItems(sheet1_cols_list)
            o_book.release_resources()
        else: 
            
            self.ui.sheet1_list.clear()

    def SelectColumnFromSheet2(self):
        fname = self.ui.book2_filename.text()
        sname = self.ui.sheet2_list.currentText()
        if (fname != '') and (fname != 'Файл не выбран!') and (sname != ''):
            o_book = xlrd.open_workbook(self.ui.book2_filename.text())
            sheet2 = o_book.sheet_by_name(self.ui.sheet2_list.currentText())
            sheet2_cols1_list = []
            for c in range(sheet2.ncols):
                cname = sheet2.col_values(c)
                sheet2_cols1_list.append(str(cname[0]) + ' (' + str(c) + ')')
            self.ui.sheet2_cols1.clear()
            self.ui.sheet2_cols1.addItems(sheet2_cols1_list)
            c = 0
            sheet2_cols2_list = []
            for c in range(sheet2.ncols):
                cname = sheet2.col_values(c)
                sheet2_cols2_list.append(str(cname[0]) + ' (' + str(c) + ')')
            self.ui.sheet2_cols2.clear()
            self.ui.sheet2_cols2.addItems(sheet2_cols2_list)
            o_book.release_resources()
        else: 

            self.ui.sheet2_list.clear()

    def acceptColumnFromSheet(self):
        self.ui.sel_col1.setText(str(self.ui.sheet1_cols.currentRow()))
        self.ui.sel_col2.setText(str(self.ui.sheet2_cols1.currentRow()))
        self.ui.sel_col3.setText(str(self.ui.sheet2_cols2.currentRow()))

    def Compare(self):
        book1 = self.ui.book1_filename.text()                
        book2 = self.ui.book2_filename.text()

        if (book1 == '') or (book1 == 'Файл не выбран!') or (book2 == '') or (book2 == 'Файл не выбран!'):
            self.ui.err_msg.setText('ФАЙЛЫ НЕ ВЫБРАНЫ')
        else:

        
            o_book1 = xlrd.open_workbook(self.ui.book1_filename.text())
            o_book2 = xlrd.open_workbook(self.ui.book2_filename.text())

            sheet_1 = o_book1.sheet_by_name(self.ui.sheet1_list.currentText())
            sheet_2 = o_book2.sheet_by_name(self.ui.sheet2_list.currentText())
            col1 = []
            col2 = []
            col3 = []
            col1 = sheet_1.col_values(int(self.ui.sel_col1.text()))
            col2 = sheet_2.col_values(int(self.ui.sel_col2.text()))
            col3 = sheet_2.col_values(int(self.ui.sel_col3.text()))

            match_list = []
            for v1 in range(len(col1)):
                for v2 in range(len(col2)):
                    if (col1[v1] == col2[v2]):
                        if (col1[v1] != '') and (col2[v2] != '') and (col3[v2] != ''):
                            match_list.append(str(col1[v1]) + '=' + str(col3[v2]))
            match_file = open('match.txt', 'w')
            for el in match_list:
                match_file.write(str(el + '\n'))










    # self.ui.book1_filename.set_text(b1_name)
    # b2_name = QtWidgets.QFileDialog.getOpenFileName()[0]
    # self.ui.book2_filename.set_text(b2_name)



# def getSheetName(workbook, sheet):
#     pointSheetObj = []
#     pointSheets = workbook.sheet_names()
#     sheet_name = ''
#     for i in pointSheets:
#         pointSheetObj.append(str(i))
#     for s in range(len(pointSheetObj)):
#     	if (str(s) == str(sheet)):
#     		sheet_name = str(pointSheetObj[s])
#     return sheet_name






# os.system('cls')

# print ('---WELCOME TO PENTAGON---\n')
# print ('-------------------------')

# book1 = xlrd.open_workbook("book1.xlsx")
# book2 = xlrd.open_workbook("book2.xlsx")

# print('Выберите лист из книги 1')
# sheet_1 = input()
# sheet1 = book1.sheet_by_index(int(sheet_1))
# print('Выбран лист ' + str(getSheetName(book1,sheet_1)) + '\n')

# print('Выберите лист из книги 2')
# sheet_2 = input()
# sheet2 = book2.sheet_by_index(int(sheet_2))
# print('Выбран лист ' + str(getSheetName(book2,sheet_2)) + '\n')

# print('Выберите сравниваемый столбец из ' + str(getSheetName(book1,sheet_1)) + '\n')
# f_col = input()
# fcol = sheet1.col_values(int(f_col))
# print('Выбран столбец: ' + '\n' + str(fcol[0]) + '\n')

# print('Выберите сравниваемый столбец из' + str(getSheetName(book2,sheet_2)) + '\n')
# s_col = input()
# scol = sheet2.col_values(int(s_col))
# print('Выбран столбец: ' + '\n' + str(scol[0]) + '\n')

# print('Выберите сопоставляемый столбец из ' + str(getSheetName(book2,sheet_2)) + ' с данными из ' + str(getSheetName(book1,sheet_1)) )
# t_col = input()
# tcol = sheet2.col_values(int(t_col))
# print('Выбран столбец: ' + '\n' + str(tcol[0]) + '\n')


# match_list = []
# for v1 in range(len(fcol)):
# 	for v2 in range(len(scol)):
# 		if (fcol[v1] == scol[v2]):
# 			match_list.append(fcol[v1] + '=' + tcol[v2])
# match_file = open('match.txt', 'w')
# for el in match_list:
# 	match_file.write(str(el + '\n'))

app = QtWidgets.QApplication([])
application = mywindow()
application.show()


sys.exit(app.exec())
