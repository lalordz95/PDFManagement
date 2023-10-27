from fun_reporter import *

import PyQt5
import gui_report
import sys
import os
import time
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *

from PyQt5.QtWidgets import QFileDialog
from pathlib import Path
import glob
import traceback

global path, flag_path, validfolder

class hilo_ejecucion(QThread): #Thread that mantains responsivness of the interface

    valorbar = pyqtSignal(int) #Connection to the progress bar object
    valorpath = pyqtSignal(str) #Connection to the label object that shows selected path
    valorstring = pyqtSignal(str) #Connection to the label object that shows information to the user
    
    def run(self):
        
        global path
        
        create_report(self, path) #function defined in fun_reporter.py script

         
class reporter_gui(QtWidgets.QMainWindow, gui_report.Ui_MainWindow):
    
    def __init__(self, parent = None): #Initializing the gui/interface. code for interface in gui_report.py script
        super(reporter_gui, self).__init__(parent)
        self.setupUi(self)
        
        init_conect(self) #Connection with the objects of the interface in fun_reporter.py
        
    def bt_selectf(self): #method connected to the "Select Folder" button
    
        global path, flag_path, validfolder
        
        flag_path = 0 #Flags
        validfolder = 0
        
        path_str = str(QFileDialog.getExistingDirectory(self, "Select Directory")) #Opens the browser to select the folder
        self.l_pathfolder.setText(path_str) #sets the text of the label "l_pathfolder" with the path of the selected folder
        checkpath = self.l_pathfolder.text() #string variable of the path of the folder
        path = Path(self.l_pathfolder.text()) #conversion from text to Path variable
        
        if checkpath == "": #Validation related to closing the browser before selecting a folder
            self.l_out.setText("Error: The folder could not be loaded successfully. Please select it again.") #Message to the user in label "l_out" object 
            self.l_pathfolder.setText("Error: Incorrect folder path") #Message to the user in label "l_pathfolder" object
            self.pBar.setValue(int(25))
        
        else:
            
            if os.path.isdir(path) == True: #Validates the existence of the path in the system
                flag_path = 1
                
            else:
                self.l_pathfolder.setText("No folder path selected")
                self.pBar.setValue(int(25))
                
            pdfCounter = 0 #

            for folder, subfolders, files in os.walk(path): #loop through subfolders and files to check if the path contains PDF files
                for file in files:
                    if file.endswith(('.pdf', '.PDF')):
                        pdfCounter = pdfCounter + 1
                        
            if pdfCounter > 0: #If at least 1 PDF file was found, a report can be generated, otherwise the program asks for a different folder
                self.l_out.setText("Folder selected successfully! Please click on Run AutoPDFReport")
                self.pBar.setValue(int(0))
                validfolder = 1 #flag for validfolder
                
            else:
                self.l_out.setText("Error: The folder does not contain PDF files. Select a different folder.")
                self.pBar.setValue(int(25))
                
        return path, flag_path, validfolder
    
    def bt_runreporter(self): #method that connects to the thread in line 19
    
        global flag_path, validfolder
         
        if flag_path == 1: #If the user selected a valid folder flag_path = 1, line 59
            
            if validfolder == 1: #If the folder contains at least 1 PDF validfolder = 1, line 75
                flag_path = 2 #Changes the value of the flag, so if the user re-runs the process, the program will ask for a new folder, because report is already generated
                self.thread = hilo_ejecucion() #Commands and connections with the objects that starts the thread
                self.thread.valorbar.connect(self.valorbarra)
                self.thread.valorstring.connect(self.textout)
                self.thread.valorpath.connect(self.textpath)
                self.thread.start()
                
        elif flag_path == 2: #Validates if the report was already created, asks for a new folder
            self.l_out.setText("Error: Select a new folder before generating a new report.")
            self.pBar.setValue(int(25))
        else: #Validates if the user already selected a folder
            self.l_out.setText("Error: Please select a folder first.")
            self.pBar.setValue(int(25))
            
        return flag_path
        
    def valorbarra(self, val): #Connections to the objects of the interface
        self.pBar.setValue(val)
    def textout(self, val):
        self.l_out.setText(val)
    def textpath(self, val):
        self.l_pathfolder.setText(val)
            
                
        
if __name__ == '__main__': #Starts the execution of the program
    a = QtWidgets.QApplication(sys.argv)
    app = reporter_gui()
    app.show()
    a.exec_()