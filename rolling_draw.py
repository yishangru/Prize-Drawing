import os
import sys
import math
import xlrd
import random
import tkinter
import win32api
import tkinter.messagebox as mb
from PyQt5 import QtCore, QtGui, QtWidgets

logo1 = "sustc.jpg"
logo2 = "logo.png"
studentList = "namelist.xlsx"

class Rolling_Dialog(object):
    def setupUi(self, Dialog):
        self.factor1 = (win32api.GetSystemMetrics(0)-80*win32api.GetSystemMetrics(0)/win32api.GetSystemMetrics(1))/2000
        self.factor2 = (win32api.GetSystemMetrics(1)-80)/1100
        self.Dialog = Dialog
        Dialog.setObjectName("Dialog")
        Dialog.resize(int(self.factor1*1910), int(self.factor2*1098))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Dialog.sizePolicy().hasHeightForWidth())
        Dialog.setSizePolicy(sizePolicy)
        Dialog.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.rolling = QtWidgets.QPushButton(Dialog)
        self.rolling.setGeometry(QtCore.QRect(int(self.factor1*20), int(self.factor2*1000), int(self.factor1*1871), int(self.factor2*81)))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.rolling.sizePolicy().hasHeightForWidth())
        self.rolling.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial Rounded MT Bold")
        font.setPointSize(int(self.factor1*20))
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(int(self.factor1*50))
        self.rolling.setFont(font)
        self.rolling.setStyleSheet("border-color: rgb(255, 255, 255);\n"
                                   "border:5px solid;border-radius:10px;\n"
                                   "border-color: rgb(0, 170, 255);\n"
                                   "font: 20pt \"Arial Rounded MT Bold\";\n"
                                   "background-color: rgb(240, 240, 240);")
        self.rolling.setObjectName("rolling")
        self.logo1 = QtWidgets.QLabel(Dialog)
        self.logo1.setGeometry(QtCore.QRect(int(self.factor1*20), int(self.factor2*30), int(self.factor1*950), int(self.factor2*270)))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.logo1.sizePolicy().hasHeightForWidth())
        self.logo1.setSizePolicy(sizePolicy)
        self.logo1.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.logo1.setText("")
        self.logo1.setObjectName("logo1")
        self.com = QtWidgets.QLabel(Dialog)
        self.com.setGeometry(QtCore.QRect(int(self.factor1*1480), int(self.factor2*160), int(self.factor1*411), int(self.factor2*141)))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.com.sizePolicy().hasHeightForWidth())
        self.com.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("华文彩云")
        font.setPointSize(int(self.factor1*25))
        font.setBold(True)
        font.setWeight(int(self.factor1*75))
        self.com.setFont(font)
        self.com.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.com.setAlignment(QtCore.Qt.AlignBottom | QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft)
        self.com.setObjectName("com")
        self.logo2 = QtWidgets.QLabel(Dialog)
        self.logo2.setGeometry(QtCore.QRect(int(self.factor1*980), int(self.factor2*30), int(self.factor1*270), int(self.factor2*270)))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.logo2.sizePolicy().hasHeightForWidth())
        self.logo2.setSizePolicy(sizePolicy)
        self.logo2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.logo2.setText("")
        self.logo2.setObjectName("logo2")
        self.stu_list = QtWidgets.QTableWidget(Dialog)
        self.stu_list.setGeometry(QtCore.QRect(int(self.factor1*20), int(self.factor2*310), int(self.factor1*1871), int(self.factor2*671)))
        self.stu_list.setStyleSheet("border-color: rgb(255, 255, 255);\n"
                                    "background-color: rgb(255, 255, 255);\n"
                                    "border:3px solid;border-radius:10px;\n"
                                    "border-color: rgb(0, 170, 127);\n"
                                    "font: 10pt \"Myriad Pro\";")
        self.stu_list.setObjectName("stu_list")
        self.stu_list.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        # disable showing sequence number
        self.stu_list.verticalHeader().setVisible(False)
        self.stu_list.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.stu_list.setSelectionMode(QtWidgets.QTableWidget.NoSelection)
        self.stu_list.setColumnCount(4)
        self.stu_list.setRowCount(0)
        self.stu_list.setColumnWidth(0, int(self.factor1*466))
        self.stu_list.setColumnWidth(1, int(self.factor1*466))
        self.stu_list.setColumnWidth(2, int(self.factor1*466))
        self.stu_list.setColumnWidth(3, int(self.factor1*466))
        self.stu_list.horizontalHeader().setFixedHeight(int(self.factor2*60))
        item = QtWidgets.QTableWidgetItem()
        self.stu_list.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.stu_list.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.stu_list.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.stu_list.setHorizontalHeaderItem(3, item)


        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

        self.on_select = -1
        self.rolling.clicked.connect(self.rolling_selection)
        self.setting_pic()
        self.loading_student_list()
        self.display_student_list()

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "比赛抽奖"))
        self.rolling.setText(_translate("Dialog", "Let\'s Rolling!"))
        self.com.setText(_translate("Dialog", "南方科技大学\n"
                                              "程序设计竞赛"))
        font =QtGui.QFont()
        font.setBold(True)
        font.setFamily("仿宋")
        font.setPointSize(int(self.factor1*20))

        #self.table.horizontalHeader().setStyleSheet('QHeaderView::section{background:gray}')
        item = self.stu_list.horizontalHeaderItem(0)
        item.setText(_translate("Dialog", "参赛者"))
        item.setFont(font)
        item = self.stu_list.horizontalHeaderItem(1)
        item.setFont(font)
        item.setText(_translate("Dialog", "参赛者"))
        item = self.stu_list.horizontalHeaderItem(2)
        item.setFont(font)
        item.setText(_translate("Dialog", "参赛者"))
        item = self.stu_list.horizontalHeaderItem(3)
        item.setFont(font)
        item.setText(_translate("Dialog", "参赛者"))

    def setting_pic(self):
        jpg_logo1 = QtGui.QPixmap(logo1).scaled(self.logo1.width(), self.logo1.height())
        self.logo1.setPixmap(jpg_logo1)
        jpg_logo2 = QtGui.QPixmap(logo2).scaled(self.logo2.width(), self.logo2.height())
        self.logo2.setPixmap(jpg_logo2)

    def loading_student_list(self):
        self.comp_student_list = list()
        if os.path.exists(os.path.abspath(studentList)):
            opendata = xlrd.open_workbook(filename=os.path.abspath(studentList))
            table_open = opendata.sheets()[0]
            for i in range(1, table_open.nrows):
                if not str(table_open.cell(i, 3).value) == "":
                    try:
                        stu_id_inter = int(table_open.cell(i, 3).value)
                        self.comp_student_list.append(str(stu_id_inter))
                    except:
                        self.comp_student_list.append(str(table_open.cell(i, 3).value))

    def display_student_list(self):
        font =QtGui.QFont()
        font.setFamily("Myriad Pro")
        font.setPointSize(int(self.factor1*26))
        rows_count = int(math.ceil(len(self.comp_student_list)/4))
        count = 0
        for i in range(rows_count):
            row_count = self.stu_list.rowCount()
            self.stu_list.insertRow(row_count)
            self.stu_list.setRowHeight(i, int(self.factor2*50))
            for j in range(4):
                if count < len(self.comp_student_list):
                    item = QtWidgets.QTableWidgetItem(str(self.comp_student_list[count]))  # doc name
                    item.setFont(font)
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    self.stu_list.setItem(row_count, j, item)
                    count += 1
                else:
                    break

    def rolling_selection(self):
        self.rolling.setEnabled(False)

        self.selected_item = 0
        self.Bigtimer = QtCore.QTimer(self.Dialog)
        self.Bigtimer.timeout.connect(self.big_time_timeout)
        self.counter = 0
        self.biger_timeout = 2000
        self.smaller_timeout = 10
        if not self.on_select == -1:
            self.stu_list.item(self.on_select // 4, self.on_select - (self.on_select // 4) * 4).setBackground(
                QtGui.QColor('White'))
        self.on_select = random.randint(0, len(self.comp_student_list)-1) #LawnGreen LimeGreen
        self.stu_list.item(self.on_select // 4, self.on_select - (self.on_select // 4) * 4).setBackground(QtGui.QColor('IndianRed'))
        scrollBar = self.stu_list.verticalScrollBar()
        scrollBar.setValue(self.on_select // 4)

        self.Bigtimer.start(self.biger_timeout)
        self.Smalltimer = QtCore.QTimer(self.Dialog)
        self.Smalltimer.timeout.connect(self.small_time_timeout)
        self.Smalltimer.start(self.smaller_timeout)

    def big_time_timeout(self):
        self.Bigtimer.stop()
        if self.counter < 345:
            self.Bigtimer.start(self.biger_timeout)
        else:
            root = tkinter.Tk()
            root.iconify()
            self.phone_get = self.stu_list.item(self.on_select // 4, self.on_select - 4 * (self.on_select // 4)).text()
            mb.showinfo("获奖者产生", "获奖者 ---- " + str(self.phone_get))
            root.destroy()
            self.rolling.setEnabled(True)

    def small_time_timeout(self):
        self.stu_list.item(self.on_select // 4, self.on_select - (self.on_select // 4) * 4).setBackground(
            QtGui.QColor('White'))
        if self.counter < 50:
            self.counter += 1
            self.on_select = random.randint(0, len(self.comp_student_list)-1)
            if self.counter == 49:
                self.Smalltimer.stop()
                self.smaller_timeout = 20
                self.Smalltimer.start(self.smaller_timeout)
        elif self.counter < 150:
            self.counter += 1
            self.on_select = random.randint(0, len(self.comp_student_list) - 1)
            if self.counter == 149:
                self.Smalltimer.stop()
                self.smaller_timeout = 50
                self.Smalltimer.start(self.smaller_timeout)
        elif self.counter < 250:
            self.counter += 1
            self.on_select = random.randint(0, len(self.comp_student_list) - 1)
            if self.counter == 249:
                self.Smalltimer.stop()
                self.smaller_timeout = 100
                self.Smalltimer.start(self.smaller_timeout)
        elif self.counter < 300:
            self.counter += 1
            self.on_select = random.randint(0, len(self.comp_student_list) - 1)
            if self.counter == 299:
                self.Smalltimer.stop()
                self.smaller_timeout = 200
                self.Smalltimer.start(self.smaller_timeout)
        elif self.counter < 330:
            self.counter += 1
            self.on_select = random.randint(0, len(self.comp_student_list) - 1)
            if self.counter == 329:
                self.Smalltimer.stop()
                self.smaller_timeout = 500
                self.Smalltimer.start(self.smaller_timeout)
        elif self.counter < 340:
            self.counter += 1
            self.on_select = random.randint(0, len(self.comp_student_list) - 1)
            if self.counter == 339:
                self.Smalltimer.stop()
                self.smaller_timeout = 1000
                self.Smalltimer.start(self.smaller_timeout)
        elif self.counter < 345:
            self.counter += 1
            self.on_select = random.randint(0, len(self.comp_student_list) - 1)
            if self.counter == 345:
                self.Smalltimer.stop()

        self.stu_list.item(self.on_select // 4, self.on_select - (self.on_select // 4) * 4).setBackground(
            QtGui.QColor('IndianRed'))
        if self.counter >= 150:
            scrollBar = self.stu_list.verticalScrollBar()
            scrollBar.setValue(self.on_select // 4)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    holderWindow = QtWidgets.QDialog()
    ui = Rolling_Dialog()
    ui.setupUi(holderWindow)
    holderWindow.show()
    sys.exit(app.exec_())