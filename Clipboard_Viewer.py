# -*- coding: utf-8 -*-
"""
Created on Sun Jan  3 16:20:07 2021

@author: deepspkd
"""

# PBF - Payment Budget and File 

import sys
from PyQt5.QtWidgets import (QApplication, QMessageBox, QDialog, QMainWindow, QLabel, QMdiArea, QDockWidget,
QAction, QToolBar, QStatusBar, QDesktopWidget, QTabWidget, QListWidget, QWidget, QTableWidgetItem, QTableWidget,
QFormLayout, QTextEdit, QLineEdit, QMdiSubWindow, QTableView)
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon, QBrush, QColor
from PyQt5.Qt import QApplication, QClipboard
from PyQt5 import QtCore

import pandas as pd 
from bs4 import BeautifulSoup, NavigableString, Tag
from io import StringIO
from dateutil.parser import parse
import os

import pathlib


def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try: 
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False

def convertDate(string):
    '''
    This will convert Date format to dd,mm,yyy
    
    '''
    return parse(string).strftime('%d-%m-%Y')
    

class TableModel(QtCore.QAbstractTableModel):

    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def data(self, index, role):
        if role == Qt.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, index):
        return self._data.shape[1]

    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])

            if orientation == Qt.Vertical:
                return str(self._data.index[section])

class PBF (QMainWindow):
    
    SavedBankBookDF = pd.DataFrame()
    SavedCashBookDF = pd.DataFrame()
    SavedJournalDF = pd.DataFrame()
    BudgetFile = pd.DataFrame()
    
    currentCopiedDF = pd.DataFrame()
    currentCopiedILGMSItem = ''
    
    currentWidget = ""
    
    
    
    
    def __init__(self):
        super().__init__()
        self.initializeUI()

    def initializeUI(self):
        self.setWindowTitle("PBF 1.00")    # Set Title of Application

        #Adjust Screen size to available size

        desktop = QDesktopWidget().screenGeometry()
        self.height = desktop.height()
        self.width = desktop.width()
        self.setGeometry(0,0,self.width,self.height)
        self.setWindowState(Qt.WindowMaximized) # Set initial state as "Maximised"

        '''
        Various Window Elements are initialised by its own functions.

        '''

        #self.mdi_Back = QBrush(QColor("blue"))
        #self.mdiArea = QMdiArea()
        #self.mdiArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        #self.mdiArea.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        #self.mdiArea.setBackground(self.mdi_Back)
        
        
        
        self.setCurrentWidget(self.currentWidget)
        
        QApplication.clipboard().clear()
        QApplication.clipboard().dataChanged.connect(self.GetClipboardData)

        #self.createTextView()
        self.createToolBar()
        self.createMenu()
        self.readSavedExcels()
        
        #self.createDockViews() #create indicator, watchlist, studies Tabbed windows and dock it to mainwindow
        #self.show()
        
    def readSavedExcels(self):
        
        currentWorkingDir = os.getcwd()
        
        fileBankBookName = currentWorkingDir + "\\" + "BankBook.xlsx"
        fileCashBookName = currentWorkingDir + "\\" + "CashBook.xlsx"
        fileJournalName = currentWorkingDir + "\\" + "Journal.xlsx"
        
        print(fileBankBookName)
                
        if(os.path.isfile(fileBankBookName) == True):
            self.SavedBankBookDF = pd.read_excel(fileBankBookName)
            #self.setCurrentView("Table",self.SavedBankBookDF)
            
        if(os.path.isfile(fileCashBookName) == True):
            self.SavedCashBookDF = pd.read_excel(fileCashBookName)
            #self.setCurrentView("Table",self.SavedCashBookDF)
            
        if(os.path.isfile(fileJournalName) == True):
            self.SavedJournalDF = pd.read_excel(fileJournalName)
        
        self.BudgetFile = pd.read_excel("Budget.xls", sheet_name = 'BS-2')
        #self.setCurrentView("Table",self.BudgetFile)
        #print(self.BudgetFile)
        
    def setCurrentWidget(self,widgetType):
        if widgetType == "" or widgetType == "Table":
            self.MainTextView = QTextEdit()
            self.setCentralWidget(self.MainTextView)
            self.currentWidget = "Text"
        elif widgetType == "Text":
            self.MainTableView = QTableView()
            self.setCentralWidget(self.MainTableView)
            self.currentWidget = "Table"
            

            
        
    def setCurrentView(self,viewType,viewData):
        print("Called")
        if viewType == "Text":
            if viewType != self.currentWidget:
                self.setCurrentWidget(self.currentWidget)
                
            self.MainTextView.setPlainText(viewData.to_string())   
            
        elif viewType == "Table":
            if viewType != self.currentWidget:
                self.setCurrentWidget(self.currentWidget)
                
            if isinstance(viewData, str):
                data = StringIO(viewData)
                df = pd.read_csv(data,"\t",engine='python')
                #df['count'] = df.groupby(df.columns.tolist()).cumcount()
                
                model = TableModel(df)
                self.MainTableView.setModel(model)
            
                self.currentCopiedDF = df #Data Frame for Saving Copied information
            else:
                model = TableModel(viewData)
                self.MainTableView.setModel(model)
            
            
            
        
    '''    
    def createTextView(self):
            
        self.MainTextEdit = QTextEdit()
        self.mdiArea.addSubWindow(self.MainTextEdit)
    '''
    ''' 
    def CreateTable(self):
        showTable = QTableView()
        model = TableModel(self.clipboardTable)
        showTable.setModel(model)
        self.mdiArea.addSubWindow(showTable)
        showTable.show()
    '''
    def createMenu(self):

        # Menu Creation for Clipboard Viewer

        #Create actions for file menu

        self.Save_ILGMS_Data_act = QAction("Save ILGMS Data", self)
        self.Save_ILGMS_Data_act.setEnabled(False) 
        
        
        #Creating Menu Bar
        
        ilgmsDataViewer_Menu = self.menuBar()
        ilgmsDataViewer_Menu.setNativeMenuBar(False)

        # Create File menu and Add actions

        file_menu = ilgmsDataViewer_Menu.addMenu("File")
        file_menu.addAction(self.Save_ILGMS_Data_act)
        self.Save_ILGMS_Data_act.triggered.connect(self.SaveILGMSData)

    def SaveILGMSData(self):
        print(self.currentCopiedILGMSItem)
        
        if self.currentCopiedILGMSItem == "BankBook":
            if len(self.SavedBankBookDF) == 0:
                self.currentCopiedDF.to_excel("BankBook.xlsx", index=False)
            else:
                print(self.currentCopiedDF)
                tempDF = pd.concat([self.SavedBankBookDF,self.currentCopiedDF]).drop_duplicates()
                self.SavedBankBookDF = tempDF
                self.SavedBankBookDF.to_excel("BankBook.xlsx",index=False)
                
        elif self.currentCopiedILGMSItem == "CashBook":
            
            self.currentCopiedDF['Count'] = self.currentCopiedDF.groupby(self.currentCopiedDF.columns.tolist()).cumcount()          
            
            if len(self.SavedCashBookDF) == 0:
                self.currentCopiedDF.to_excel("CashBook.xlsx", index=False)
            else:
                tempDF = pd.concat([self.SavedCashBookDF,self.currentCopiedDF]).drop_duplicates()
                self.SavedCashBookDF = tempDF
                self.SavedCashBookDF.to_excel("CashBook.xlsx",index=False)
                
        elif self.currentCopiedILGMSItem == "Journal":
            if len(self.SavedJournalDF) == 0:
                self.currentCopiedDF.to_excel("Journal.xlsx", index=False)
            else:
                tempDF = pd.concat([self.SavedJournalDF,self.currentCopiedDF]).drop_duplicates()
                self.SavedJournalDF = tempDF
                self.SavedJournalDF.to_excel("CashBook.xlsx",index=False)
        
        self.currentCopiedDF = pd.DataFrame()
        self.currentCopiedILGMSItem = ""
        self.readSavedExcels()
        
        
    def createToolBar(self):
        
        self.Copy_Journal_Act = QAction("Journal", self)
        self.Copy_CashBook_Act = QAction("CashBook",self)
        self.Copy_BankBook_Act = QAction("BankBook", self)

        self.Copy_Journal_Act.setEnabled(False)
        self.Copy_CashBook_Act.setEnabled(False)
        self.Copy_BankBook_Act.setEnabled(False)

        
        ilgmsDataViewer_ToolBar = QToolBar("ILGMSDataViewer")
        ilgmsDataViewer_ToolBar.setIconSize(QSize(24,24))
        self.addToolBar(ilgmsDataViewer_ToolBar)
        
        ilgmsDataViewer_ToolBar.addAction(self.Copy_Journal_Act)
        ilgmsDataViewer_ToolBar.addAction(self.Copy_BankBook_Act)
        ilgmsDataViewer_ToolBar.addAction(self.Copy_CashBook_Act)
        
        self.Copy_Journal_Act.triggered.connect(self.copyJournalData)
        self.Copy_CashBook_Act.triggered.connect(self.copyCashBookData)
        self.Copy_BankBook_Act.triggered.connect(self.copyBankBookData)
        
    def copyCashBookData(self):
        print("CashBook")
        clipboardText = QApplication.clipboard().text()
        txt = self.TakeCashBookData(clipboardText)
        self.currentCopiedILGMSItem = "CashBook"
        self.setCurrentView("Table",txt)
        
    def copyJournalData(self):
        print("Journal")
        clipboardHtml = QApplication.clipboard().mimeData().html()
        txt = self.TakeJournalData(clipboardHtml)
        self.currentCopiedILGMSItem = "Journal"
        self.setCurrentView("Table",txt)
        
    def copyBankBookData(self):
        print("BankBook")
        clipboardText = QApplication.clipboard().text()
        txt = self.TakeBankBookData(clipboardText)
        self.currentCopiedILGMSItem = "BankBook"
        self.setCurrentView("Table",txt)
    
           
    def GetClipboardData(self):
        
        if QApplication.clipboard().text():
            self.Copy_Journal_Act.setEnabled(True)
            self.Copy_CashBook_Act.setEnabled(True)
            self.Copy_BankBook_Act.setEnabled(True)
            self.Save_ILGMS_Data_act.setEnabled(True)
        else:
            self.Copy_Journal_Act.setEnabled(False)
            self.Copy_CashBook_Act.setEnabled(False)
            self.Copy_BankBook_Act.setEnabled(False)
            self.Save_ILGMS_Data_act.setEnabled(False)

        '''    
        clipboardText = QApplication.clipboard().text()
        clipboardHtml = QApplication.clipboard().mimeData().html()
       
        
        csvStr = StringIO(clipboardText)
        htmlStr = StringIO(clipboardHtml)        
        self.clipboardTable = pd.read_table(csvStr,header=0,sep="\n",error_bad_lines = False)
        self.CreateTable()
        '''
        
    def TakeJournalData(self,clipboardhtmlText):

        soup = BeautifulSoup(clipboardhtmlText,'html.parser')        

        JournalText = ""  #This Captures Journal items from Copied Text
        
        # Code which Strips HTML and gets the data
        
        for tag in soup.descendants: 
            
            SpaceChecker = "" # Variable to verify Space and :
            
            # IDEA : Get to the last item in the htmltree and Take Text from it
            if(isinstance(tag, NavigableString)):
                
                SpaceChecker = str(tag.string)
                if SpaceChecker.strip() != "" and SpaceChecker[-1] != ":":
                    JournalText += SpaceChecker.strip() + "\t"
                else:
                    JournalText += SpaceChecker.strip()
                
            elif(tag.text == ""):
                JournalText += "0\t"
            
            # Has to write code to Output data in needed format
        crudeList = list(JournalText.split("\t"))
        selectList = []
        
        titleList = ['JournalNo','Date','Type','Code','Head','Dr','Cr','Narration']
        
        journalNumbers = 0        
        journalDataCounter = 0
        journalMaxColumnCounter = 0
        journalColumnCounter = 0
        individualJournalData = ""
        
        for index in range(len(crudeList)):
            if "JOURNAL NO:" in crudeList[index]:
                individualJournalData += crudeList[index].replace("JOURNAL NO:", "")
            elif is_date(crudeList[index]) == True:
                individualJournalData += "\t" + convertDate(str(crudeList[index]))
            elif "Type:" in crudeList[index]:
                individualJournalData += "\t" + crudeList[index].replace("Type:", "")
            elif crudeList[index - (journalDataCounter + 1)] == "Cr":
                if "@Total:" in crudeList[index+1]:
                    if journalDataCounter < 4:
                        individualJournalData += "\t - \t - \t - \t -"
                    journalDataCounter = 0
                                       
                else:                
                    journalDataCounter += 1
                    
                    if journalDataCounter > 1:                        
                        if journalDataCounter / 4 == (journalDataCounter // 4) + 0.25:
                            individualJournalData += "-" + "\t -" + "\t -"
                            
                    
                    if ".00" in crudeList[index]:
                        individualJournalData += "\t" + crudeList[index].replace(",","")
                        
                    else:
                        individualJournalData += "\t" + str(crudeList[index])
                        
                        
                    if journalDataCounter % 4 == 0:
                        if "@Total:" not in crudeList[index+2]:
                            individualJournalData += "\t -"
                            selectList.append(individualJournalData)
                            individualJournalData = ""
                            
                
            elif "Narration:" in crudeList[index]:
                individualJournalData += "\t" + crudeList[index].replace("Narration:", "")
                selectList.append(individualJournalData)
                individualJournalData = ""
                journalNumbers += 1

        selectList.insert(0,"\t".join(titleList))
        
        formattedText = "\n"
        formattedText = formattedText.join(selectList)
        
        return formattedText
        
        
    def TakeBankBookData(self,clipboardText):
        
        crudeList = list(clipboardText.split("\n"))
        
        removalList = []
        appendList = []
        dateList = []
        
        for index in range(len(crudeList)):
            if "Opening" in crudeList[index]:
                removalList.append(index)
            elif "Daily" in crudeList[index]:
                removalList.append(index)
                removalList.append(index+1)
                removalList.append(index+2)
        
        deleteadjustment = 0
        
        for index in range(len(removalList)):
            crudeList.pop(removalList[index] - deleteadjustment)
            deleteadjustment += 1
            
        for index in range(len(crudeList)):            
            if "Type :" in crudeList[index]:
                appendList.append(index)
                appendList.append(index+1)
                crudeList[index-1] = str(crudeList[index-1]) + "\t" + str(crudeList[index]) + str(crudeList[index+1]) 
    
        deleteadjustment = 0
        
        for index in range(len(appendList)):
            crudeList.pop(appendList[index] - deleteadjustment)
            deleteadjustment += 1
        
        
        for index in range(len(crudeList)):            
            if is_date(crudeList[index]) == True:
                dateList.append(index)
                if len(dateList) == 1:
                    crudeList[1] = "Date" + "\t" + str(crudeList[1]).replace("Description", "Description \t Cheque")
                    
                    
            elif len(dateList) > 0:
                crudeList[index] = convertDate(crudeList[dateList[-1]].replace("\t",""))+ "\t" + str(crudeList[index])

        deleteadjustment = 0
        
        for index in range(len(dateList)):
            crudeList.pop(dateList[index] - deleteadjustment)
            deleteadjustment += 1
            
        bankNameString  = crudeList[0].split("___")
        bankNameString1 = bankNameString[1].split("(")
        bankNameString.pop(1)
        bankNameString1[1] = bankNameString1[1].replace(")","")
        bankNameString.extend(bankNameString1)

        for index in range(len(bankNameString)):
            bankNameString[index].strip()
            
        crudeList.pop(0)
        
        for index in range(len(crudeList)):
            if index == 0:
                crudeList[index] = crudeList[index] + "\t" + "Bank Head" + "\t" + "Bank Name" + "\t" + "Account No"
            else:
                for nextIndex in range(len(bankNameString)):
                    crudeList[index] = crudeList[index] + "\t" + bankNameString[nextIndex]
            
        fromattedText = "\n"
        fromattedText = fromattedText.join(crudeList)
        return fromattedText
    
    def TakeCashBookData(self,clipboardText):
        
        crudeList = list(clipboardText.split("\n"))
        
        removalList = []
        dateList = []
        
        for index in range(len(crudeList)):
            if "Opening" in crudeList[index]:
                removalList.append(index)
            elif "Daily" in crudeList[index]:
                removalList.append(index)
                removalList.append(index+1)
                removalList.append(index+2)
        
        deleteadjustment = 0
        
        for index in range(len(removalList)):
            crudeList.pop(removalList[index] - deleteadjustment)
            deleteadjustment += 1
            
        for index in range(len(crudeList)):            
            if is_date(crudeList[index]) == True:
                dateList.append(index)
                if len(dateList) == 1:
                    crudeList[0] = "Date" + "\t" + str(crudeList[0])
                
            elif len(dateList) > 0:
                crudeList[index] = convertDate(crudeList[dateList[-1]].replace("\t",""))+ "\t" + str(crudeList[index])

        deleteadjustment = 0
        
        for index in range(len(dateList)):
            crudeList.pop(dateList[index] - deleteadjustment)
            deleteadjustment += 1
        
        fromattedText = "\n"
        fromattedText = fromattedText.join(crudeList)
        return fromattedText
        

if __name__ == "__main__":
    app = QApplication([])
    window = PBF()
    window.show()
    sys.exit(app.exec_())