import sys
import openpyxl
import os
from PyQt5.QtWidgets import *
from PyQt5 import uic

import time

form_class = uic.loadUiType("Excel.ui")[0]

MAX_COL_SIZE = 10
MAX_ROW_SIZE = 10

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super(QMainWindow, self).__init__()
        self.setupUi(self)
        self.pushLoadButton.clicked.connect(self.Load_btn_clicked)
        self.pushSaveButton.clicked.connect(self.Save_btn_clicked)
        self.pushTestButton.clicked.connect(self.Test_btn_clicked)
        
        self.tableWidget.setColumnCount(MAX_COL_SIZE)
        self.tableWidget.setRowCount(MAX_ROW_SIZE)
        
        self.tableWidget.setHorizontalHeaderLabels(['A','B','C','D','E']) 
        
        self.file_name = [] #Excel file name
        self.wb = [] #work book
        self.ws = [] #work sheet
       
        
    def Load_btn_clicked(self):
        file_dlg = QFileDialog()
        self.file_name = file_dlg.getOpenFileName(self, 'Open file', os.getcwd(), "Excel files (*.xlsx)")
        print (self.file_name)
        self.wb = openpyxl.load_workbook(self.file_name[0]) # data_only=False is deafault value
                                                            # data_only=True read just result of function
        self.wb.guess_types = True #type guess enable
        self.ws = self.wb.active
        
        global MAX_ROW_SIZE
        global MAX_COL_SIZE

        MAX_ROW_SIZE = self.ws.max_row
        MAX_COL_SIZE = self.ws.max_column
        
        #change size from sheet size
        self.tableWidget.setColumnCount(MAX_COL_SIZE)
        self.tableWidget.setRowCount(MAX_ROW_SIZE)
        #######
        ## TODO : From MAX SIZE need to update tableWidget.setHorizontalHeaderLabels ( List of label )
        #######
        
        for row in range(1, MAX_ROW_SIZE+1):
            for col in range(1, MAX_COL_SIZE+1):
                cell_txt = self.ws.cell(column = col , row = row).value
                #print(cell_txt) #for debug
                if cell_txt is not None:
                    if(is_number(cell_txt)): #nubmer
                        self.tableWidget.setItem( row-1,col-1, QTableWidgetItem( str(cell_txt) ) )
                    else:
                        self.tableWidget.setItem( row-1,col-1, QTableWidgetItem( cell_txt ) )
                
        
    def Save_btn_clicked(self):
        #print(self.tableWidget.itemAt(1,1).text())
        cell_txt = self.tableWidget.currentItem().text()
        global MAX_ROW_SIZE
        global MAX_COL_SIZE
        
        for row in range(1, MAX_ROW_SIZE+1):
            for col in range(1, MAX_COL_SIZE+1):
                item = self.tableWidget.item(row-1,col-1)
                if item is not None: #QtableWidget item Exist
                    cell_txt = self.tableWidget.item(row-1,col-1).text()
                    if(is_number(cell_txt)): #nubmer
                        self.ws.cell(column = col , row = row, value = cell_txt).number_format
                    else: #text
                        self.ws.cell(column = col , row = row, value = cell_txt)
                else: #Need to make QtableWidget item 
                    self.ws.cell(column = col , row = row, value = "")
        
        self.wb.save(self.file_name[0])
        QMessageBox.about(self, "message", "Save Success!!!")
        
    def Test_btn_clicked(self):
        #cell_txt = self.tableWidget.currentItem().text()
        cur_col = self.tableWidget.currentColumn()
        cur_row = self.tableWidget.currentRow()
        print("%s %s" %(cur_row,cur_col))
        
        #QTableWidget Read Test OK
        # item = self.tableWidget.item(cur_row,cur_col) 
        # if item is not None: #QtableWidget item Exist
            # item.setText("TEST")
        # else: #Need to make QtableWidget item 
            # self.tableWidget.setItem(cur_row,cur_col, QTableWidgetItem("TEST"))
            #cell_txt = self.tableWidget.item(cur_row,cur_col).text()
            #QMessageBox.about(self, "message", cell_txt)
            
        #QTableWidget Write Test
        item = self.tableWidget.item(cur_row,cur_col) 
        if item is not None: #QtableWidget item Exist
            item.setText("TEST")
        else: #Need to make QtableWidget item 
            self.tableWidget.setItem(cur_row,cur_col, QTableWidgetItem("TEST"))
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()



