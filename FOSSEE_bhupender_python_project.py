import sys
import os
import csv
from PyQt5.QtWidgets import QTableWidget, QTabWidget, QVBoxLayout, QWidget, QApplication, QMainWindow, QTableWidgetItem, QFileDialog
from PyQt5.QtWidgets import qApp, QAction, QMessageBox, QAbstractScrollArea
import xlrd
from json import dumps
from os import remove
class Table(QTableWidget):
    ''' it initialize the class QTableWidgets'''
    def __init__(self, r, c):
        ''' instantiate  the table object with rows and columns

        :param r:  rows
        :param c:  columns

        :return: None
        '''
        super().__init__(r, c)


        self.ncols = 26
        self.nrows = 26
        self.interface()

    def interface(self):
        ''' initialize graphical Interface

        :return: None
        '''
        self.show()

    def warning_message(self,bad_values,flag):
        ''' initialize the QMessageBox object for warning message

        :param bad_values: list of (rows,cols) containing bad values/duplicate
        :param flag: Flags the condition for bad values

        :return: None
        '''
        msg = QMessageBox()
        msg.setWindowTitle("Warning!")
        temp_str = ''
        for row,col in bad_values:
            temp_str += '({0:d}, {1:d}), '.format(row+1,col+1)
        temp_str = 'cell {0:s} Contain bad values'.format(temp_str)
        if flag == 'DuplicateID':
            if len(bad_values) > 0:
                temp_str += '\nColumn(ID) contains Duplicate Values'
            else:
                temp_str = 'Column(ID) contains Duplicate Values'
        msg.setText(temp_str)
        x = msg.exec_()

class spread_work():
    '''class for creating new spreadsheet'''
    def __init__(self):
        ''' Constructor for instantiating the table object

        :return : None
        '''
        super().__init__()
        self.widget = Table(1, 1)
        self.headers = []
        self.widget.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
    def setColumnHeaders(self):
        ''' Sets the Headers for the table Object

        :return: None
        '''
        self.widget.setHorizontalHeaderLabels(self.headers)
        ''' Sets the Headers for the table Object

        :return : None
        '''
class Main(QMainWindow):
    ''' this class inheritate from the QMainWindow'''
    def __init__(self):
        ''' Constructor for instantiating the table object

        :return: None
        '''
        super().__init__()
        self.setWindowTitle('SpreadSheet')
        self.resize(800, 600)
        self.work_books = 0
        self.Tab_widget = QTabWidget()
        self.vBox = QVBoxLayout()
        widget = QWidget()
        widget.setLayout(self.vBox)
        self.listing_book = []
        self.new_sheet()
        self.vBox.addWidget(self.Tab_widget)
        self.setCentralWidget(widget)
        self.show()
        self.setMenu()
    def setTab(self,name,index=0):
        ''' In Qtabwidget it add Tabs

        :param index: current index of Tab in QTabWidget
        :param name: sets the title of tab according to csv file name

        :return: None
        '''
        self.Tab_widget.addTab(self.listing_book[index].widget,name)

    def setMenu(self):
        ''' Sets the interface, menu, toolbar

        :return: None
        '''
        
        stop = QAction('Quit', self)
        loading  = QAction('Load inputs', self)
        validation = QAction('Validate', self)
        submitt = QAction('Submit', self)
       
        loading.triggered.connect(self.new_sheet)
        validation.triggered.connect(self.validate_sheet)
        submitt.triggered.connect(self.final_save)
        stop.triggered.connect(self.quit_app)
       
        toolbar = QMainWindow.addToolBar(self,'Toolbar')
        toolbar.addAction(loading)
        toolbar.addAction(validation)
        toolbar.addAction(submitt)
        toolbar.addAction(stop)


    def converter(self,path):
        ''' it change file format from xlsx to csv

        :param path: Directory path for the file

        :return: None
        '''
        wb = xlrd.open_workbook(path,on_demand=True)
        self.work_books = wb.nsheets
        for i in range(self.work_books):
            sh = wb.sheet_by_index(i)
            your_csv_file = open('your_csv_file'+str(i)+'.csv', 'w')
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL,lineterminator='\n')

            for rownum in range(sh.nrows):
                wr.writerow(sh.row_values(rownum))                    
            your_csv_file.close()
        return [['your_csv_file'+str(i)+'.csv', wb.sheet_names()[i] ] for i in range(self.work_books)]

    def new_sheet(self):
        ''' this function open the new xlsx and csv file

        :return : workbooklist - contains class sheet object list
                  paths - list of path for each csv file
        '''
        paths = ['']
        path = QFileDialog.getOpenFileName(self, 'Open CSV/XLS', os.getenv('HOME'), 'CSV(*.csv *.xlsx)')
        if path[0] == '':
            return
        paths[0] = path[0]
        sheet_names = [path[0].split('/')[-1].split('.')[0]]
        if path[0].split('/')[-1].split('.')[-1] == 'xlsx':
            paths.clear()
            paths,sheet_names = list(zip(*self.converter(path[0])))
        else:
            self.work_books = 1
                #changes needed to open multiple WB
        self.listing_book.clear()
        for g in range(self.work_books):
            self.listing_book.append(spread_work())
            self.setTab(sheet_names[g],g)
        def Sheets(Workbooks,paths):
            ''' repopulates the each table with each csv file 

            :param Workbooks: list of class sheet objects
            :param paths: list of paths to the each csv files
                        
            :return: None
            '''
            for Workbook,path in zip(Workbooks,paths):
                with open(path, newline='',encoding='utf-8',errors='ignore') as csv_file:
                    Workbook.widget.setRowCount(0)
                    #Workbook.widget.setColumnCount(10)
                    my_file = csv.reader(csv_file, dialect='excel')
                    fields = next(my_file)
                    Workbook.widget.setColumnCount(len(list(filter(lambda x: x != "", fields))))
                    Workbook.headers = list(filter(lambda x: x != "", fields))
                    Workbook.widget.ncols = len(fields)
                    Workbook.setColumnHeaders()
                    for row_data in my_file:
                        row = Workbook.widget.rowCount()
                        Workbook.widget.insertRow(row)
                        for column, stuff in enumerate(row_data):
                            item = QTableWidgetItem(stuff)
                            Workbook.widget.setItem(row, column, item)
                    Workbook.widget.nrows = Workbook.widget.rowCount()
                Workbook.widget.resizeColumnsToContents()
            for w in range(self.work_books):
                try:
                   remove('your_csv_file{0:d}.csv'.format(w))
                except:
                      pass


        return Sheets(self.listing_book,paths)
    def validate_sheet(self):
        ''' validates the workbook table if contains bad values calls message box

        :return : None
        '''
        Bad_val = []
        ID_col = []
        flag='NoDuplicateID'
        cur_workbook = self.Tab_widget.currentIndex()
        ncols = self.listing_book[cur_workbook].widget.ncols 
        nrows = self.listing_book[cur_workbook].widget.nrows 
        for i in range(0,ncols):
            for j in range(0,nrows):
                if self.listing_book[cur_workbook].widget.item(j, i) is not None:
                    if (self.listing_book[cur_workbook].widget.item(j, i).text()).replace('.','').isnumeric() == False:
                        Bad_val.append([i,j])
                    elif i == 0:
                        ID_col.append(self.listing_book[cur_workbook].widget.item(j, i).text())
        #checking duplicate IDs
        if len(ID_col) > len(set(ID_col)):
            flag = 'DuplicateID'
        if len(Bad_val) > 0 or flag == 'DuplicateID':
            self.listing_book[cur_workbook].widget.warning_message(Bad_val,flag)
       
    def final_save(self):
        ''' creates the text file for workbook

        :return: None
        '''
        workbook_index = self.Tab_widget.currentIndex()
        workbookTitle = self.Tab_widget.tabText(workbook_index)
        ncols = self.listing_book[workbook_index].widget.ncols 
        nrows = self.listing_book[workbook_index].widget.nrows
        Dictonary = {}
        for i in range(0,nrows):
            for j in range(0,ncols):
                if self.listing_book[workbook_index].widget.item(i, j) is not None:
                   Dictonary[self.listing_book[workbook_index].widget.horizontalHeaderItem(j).text()] = self.listing_book[workbook_index].widget.item(i, j).text()
            if self.listing_book[workbook_index].widget.item(i, 0) is not None:
                with open('{0:s}_{1:s}.txt'.format(workbookTitle,self.listing_book[workbook_index].widget.item(i, 0).text()), 'w') as file:
                     file.write(dumps(Dictonary).replace('{','').replace('}',''))
                Dictonary.clear()

    def quit_app(self):
        ''' close the window and closes the app

        :return: None
        '''
        qApp.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = Main()
    sys.exit(app.exec_())
