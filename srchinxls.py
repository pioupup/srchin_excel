#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys
import os
import xlrd
from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit, QTextEdit, QGridLayout, QPushButton, QApplication, QFileDialog)
from PyQt5.QtGui import QFont

class Example(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.path_Title = QLabel('Path:', self)
        self.path_Edit = QLineEdit(os.getcwd())
        self.srch_wrd_Title = QLabel('Keywords:', self)
        self.srch_wrd_Edit = QLineEdit("")

        self.path_Btn = QPushButton('Browse', self)
        self.path_Btn.clicked.connect(self.showDialog)
        self.path_Btn.resize(self.path_Btn.sizeHint())

        self.start_Btn = QPushButton('Start', self, default = True)
        self.start_Btn.clicked.connect(self.start_search)
        self.start_Btn.resize(self.start_Btn.sizeHint())
        self.start_Btn.setAutoDefault(True)
        self.srch_wrd_Edit.returnPressed.connect(self.start_Btn.click)

        self.srch_List = QTextEdit()

        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.path_Title, 1, 0)
        grid.addWidget(self.path_Edit, 1, 1)
        grid.addWidget(self.path_Btn, 1, 2)
        grid.addWidget(self.srch_wrd_Title, 2, 0)
        grid.addWidget(self.srch_wrd_Edit, 2, 1)
        grid.addWidget(self.start_Btn, 2, 2)
        grid.addWidget(self.srch_List, 3, 0, 1, 3)

        self.setLayout(grid)

        self.setGeometry(200, 200, 800, 700)
        self.setWindowTitle('Search in EXEL files')
        self.show()

    def showDialog(self):
        select_dir = QFileDialog.getExistingDirectory(self,
                        'Browse directory', './')
        self.path_Edit.clear()
        self.path_Edit.insert(select_dir)

    def search_in_path(self, s_path):
        list_files = [] # list for finding files
        for rootdir, dirs, files in os.walk(s_path):
            for file in files:
                if file.split('.')[-1] == 'xls' or file.split('.')[-1] == 'xlsx':
                    # append file path if it xls or xlsx
                    list_files.append(os.path.join(rootdir, file))
        return list_files

    def srch_txt_in_xlsfile(self, xlsfile, search_text):
        wrkbook = xlrd.open_workbook(xlsfile)
        # watching sheets in file
        self.srch_List.setFontPointSize(11)
        self.srch_List.append(xlsfile)
        self.srch_List.setFontPointSize(9)
        for sheet_nmb in range(wrkbook.nsheets):
            sheet = wrkbook.sheet_by_index(sheet_nmb)
            for rownum in range(sheet.nrows):   # going to rows
                for colnum in range(sheet.ncols):   # going to cols
                    # coincidence with searching string
                    if search_text in str(sheet.cell(rownum, colnum).value):
                        row_val = ""
                        for cols in range(sheet.ncols):
                            row_val += str(sheet.cell(rownum,cols).value) + " "
                        self.srch_List.append(row_val)
                        # ~ self.srch_List.append("[Column:" +
                            # ~ str(colnum+1) + " Row:" + str(rownum+1) +
                            # ~ " Sheet:" + str(sheet_nmb) + " Path:" +
                            # ~ str(xlsfile) + "]")
        self.srch_List.append("-" * 150)

    def start_search(self):
        self.start_Btn.setEnabled(0)
        s_path = self.path_Edit.text()
        if os.path.exists(s_path) == False:
            s_path = os.getcwd()
            self.path_Edit.clear()
            self.path_Edit.insert(s_path)

        search_text = self.srch_wrd_Edit.text()
        if search_text == "":
            search_text = "кардан"
            self.srch_wrd_Edit.insert(search_text)
        self.srch_List.clear()
        self.srch_List.append("Starting search...")
        for wrk_file in self.search_in_path(s_path):
            self.srch_txt_in_xlsfile(wrk_file, search_text)
        self.srch_List.append("Search finished...")
        self.start_Btn.setEnabled(1)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
