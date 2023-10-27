from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(788, 385)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.b_selectfolder = QtWidgets.QPushButton(self.centralwidget)
        self.b_selectfolder.setGeometry(QtCore.QRect(610, 20, 161, 51))
        self.b_selectfolder.setObjectName("b_selectfolder") #definition of the select folder button
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 30, 581, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.l_pathfolder = QtWidgets.QLabel(self.centralwidget) #definition of the path label
        self.l_pathfolder.setGeometry(QtCore.QRect(20, 90, 741, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.l_pathfolder.setFont(font) 
        self.l_pathfolder.setAlignment(QtCore.Qt.AlignCenter)
        self.l_pathfolder.setObjectName("l_pathfolder")
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(20, 90, 751, 41))
        self.groupBox.setTitle("")
        self.groupBox.setObjectName("groupBox")
        self.b_runreport = QtWidgets.QPushButton(self.centralwidget) #definition of the "Run process" button
        self.b_runreport.setGeometry(QtCore.QRect(230, 150, 321, 81))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.b_runreport.setFont(font)
        self.b_runreport.setObjectName("b_runreport")
        self.l_out = QtWidgets.QLabel(self.centralwidget) #definition of the information label
        self.l_out.setGeometry(QtCore.QRect(20, 310, 751, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.l_out.setFont(font)
        self.l_out.setText("")
        self.l_out.setAlignment(QtCore.Qt.AlignCenter)
        self.l_out.setObjectName("l_out")
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_2.setGeometry(QtCore.QRect(20, 250, 751, 111))
        self.groupBox_2.setTitle("")
        self.groupBox_2.setObjectName("groupBox_2")
        self.pBar = QtWidgets.QProgressBar(self.groupBox_2)
        self.pBar.setGeometry(QtCore.QRect(10, 20, 731, 31))
        self.pBar.setProperty("value", 0)
        self.pBar.setObjectName("pBar")
        self.groupBox.raise_()
        self.groupBox_2.raise_()
        self.b_selectfolder.raise_()
        self.label.raise_()
        self.l_pathfolder.raise_()
        self.b_runreport.raise_()
        self.l_out.raise_()
        MainWindow.setCentralWidget(self.centralwidget)
        self.pbar = QtWidgets.QStatusBar(MainWindow)
        self.pbar.setObjectName("pbar")
        MainWindow.setStatusBar(self.pbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "AutoPDFReport by Lalo Rodriguez")) #definition of the first messages of the objects
        self.b_selectfolder.setText(_translate("MainWindow", "Select Folder"))
        self.label.setText(_translate("MainWindow", "Please select the folder that contains the PDF assessments that you want check for missing answers."))
        self.l_pathfolder.setText(_translate("MainWindow", "Select a folder"))
        self.l_out.setText(_translate("MainWindow", "Waiting to start the process..."))
        self.b_runreport.setText(_translate("MainWindow", "Run AutoPDFReport"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
