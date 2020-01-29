# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created by: PyQt5 UI code generator 5.13.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
import openpyxl
import os
import time
from datetime import datetime
import logging
from socket import gethostname
import qtawesome
from shutil import copyfile

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(781, 581)
        font = QtGui.QFont()
        font.setFamily("Russo One")
        MainWindow.setFont(font)
        MainWindow.setWindowOpacity(1.0)

        MainWindow.setStyleSheet("""
                            * {
                                font-family: "Russo one";
                            }
                            QMainWindow{background-color: #1e1e2f;}
                            QLineEdit {
                                color: #a7a7ba;
                                border: 1px solid #353a53;
                                background-color: #27293d;
                                padding: 3px;
                                margin-bottom: 3px;
                                }
                            QLabel {
                                color: #a7a7ba;
                            }
                            QTextEdit {
                                background-color: #27293d;
                                color: #a7a7ba;
                                border: 1px solid #353a53;
                                padding: 3px;
                            }
                            QTreeView {
                                background-color: #27293d;
                                color: #a7a7ba;
                                border: 1px solid #353a53;
                            }
                            QProgressBar {
                                background-color: #27293d;
                                color: #a7a7ba;
                            }
                            QPushButton {
                                background-color: #3bb001;
                                border: 1px solid #353a53;
                                padding: 5px;
                                font-size: 15px;
                                font-weight: 900;
                                color: #FFFFFF;
                            }
                            """)

        MainWindow.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        MainWindow.setDocumentMode(False)
        MainWindow.setTabShape(QtWidgets.QTabWidget.Triangular)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.part_number_label = QtWidgets.QLabel(self.centralwidget)
        self.part_number_label.setObjectName("part_number_label")
        self.gridLayout_2.addWidget(self.part_number_label, 0, 0, 1, 1)
        self.new_part_number_label = QtWidgets.QLabel(self.centralwidget)
        self.new_part_number_label.setObjectName("new_part_number_label")
        self.gridLayout_2.addWidget(self.new_part_number_label, 0, 1, 1, 1)
        self.treeWidget = QtWidgets.QTreeWidget(self.centralwidget)
        self.treeWidget.setObjectName("treeWidget")
        self.treeWidget.headerItem().setText(0, "1")
        self.gridLayout_2.addWidget(self.treeWidget, 0, 2, 11, 1)
        self.part_number_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.part_number_textbox.setObjectName("part_number_textbox")
        self.gridLayout_2.addWidget(self.part_number_textbox, 1, 0, 1, 1)
        self.new_part_number_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.new_part_number_textbox.setObjectName("new_part_number_textbox")
        self.gridLayout_2.addWidget(self.new_part_number_textbox, 1, 1, 1, 1)
        self.supplier_label = QtWidgets.QLabel(self.centralwidget)
        self.supplier_label.setObjectName("supplier_label")
        self.gridLayout_2.addWidget(self.supplier_label, 2, 0, 1, 1)
        self.new_supplier_label = QtWidgets.QLabel(self.centralwidget)
        self.new_supplier_label.setObjectName("new_supplier_label")
        self.gridLayout_2.addWidget(self.new_supplier_label, 2, 1, 1, 1)
        self.supplier_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.supplier_textbox.setObjectName("supplier_textbox")
        self.gridLayout_2.addWidget(self.supplier_textbox, 3, 0, 1, 1)
        self.new_supplier_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.new_supplier_textbox.setObjectName("new_supplier_textbox")
        self.gridLayout_2.addWidget(self.new_supplier_textbox, 3, 1, 1, 1)
        self.price_label = QtWidgets.QLabel(self.centralwidget)
        self.price_label.setObjectName("price_label")
        self.gridLayout_2.addWidget(self.price_label, 4, 0, 1, 1)
        self.new_price_label = QtWidgets.QLabel(self.centralwidget)
        self.new_price_label.setObjectName("new_price_label")
        self.gridLayout_2.addWidget(self.new_price_label, 4, 1, 1, 1)
        self.price_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.price_textbox.setObjectName("price_textbox")
        self.gridLayout_2.addWidget(self.price_textbox, 5, 0, 1, 1)
        self.new_price_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.new_price_textbox.setObjectName("new_price_textbox")
        self.gridLayout_2.addWidget(self.new_price_textbox, 5, 1, 1, 1)
        self.description_label = QtWidgets.QLabel(self.centralwidget)
        self.description_label.setObjectName("description_label")
        self.gridLayout_2.addWidget(self.description_label, 6, 0, 1, 1)
        self.new_description_label = QtWidgets.QLabel(self.centralwidget)
        self.new_description_label.setObjectName("new_description_label")
        self.gridLayout_2.addWidget(self.new_description_label, 6, 1, 1, 1)
        self.description_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.description_textbox.setObjectName("description_textbox")
        self.gridLayout_2.addWidget(self.description_textbox, 7, 0, 1, 1)
        self.new_description_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.new_description_textbox.setObjectName("new_description_textbox")
        self.gridLayout_2.addWidget(self.new_description_textbox, 7, 1, 1, 1)
        self.quantity_label = QtWidgets.QLabel(self.centralwidget)
        self.quantity_label.setObjectName("quantity_label")
        self.gridLayout_2.addWidget(self.quantity_label, 8, 0, 1, 1)
        self.new_quantity_label = QtWidgets.QLabel(self.centralwidget)
        self.new_quantity_label.setObjectName("new_quantity_label")
        self.gridLayout_2.addWidget(self.new_quantity_label, 8, 1, 1, 1)
        self.quantity_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.quantity_textbox.setObjectName("quantity_textbox")
        self.gridLayout_2.addWidget(self.quantity_textbox, 9, 0, 1, 1)
        self.new_quantity_textbox = QtWidgets.QLineEdit(self.centralwidget)
        self.new_quantity_textbox.setObjectName("new_quantity_textbox")
        self.gridLayout_2.addWidget(self.new_quantity_textbox, 9, 1, 1, 1)
        self.start_button = QtWidgets.QPushButton(self.centralwidget)
        self.start_button.setObjectName("start_button")
        self.gridLayout_2.addWidget(self.start_button, 10, 0, 1, 2)
        self.start_button.clicked.connect(self.updater)

        self.console_output = QtWidgets.QTextEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("xos4 Terminus")
        self.console_output.setFont(font)
        self.console_output.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.console_output.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.console_output.setObjectName("console_output")
        self.gridLayout_2.addWidget(self.console_output, 11, 0, 1, 3)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.gridLayout_2.addWidget(self.progressBar, 12, 0, 1, 3)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 781, 19))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        self.menuHelp_2 = QtWidgets.QMenu(self.menubar)
        self.menuHelp_2.setObjectName("menuHelp_2")
        MainWindow.setMenuBar(self.menubar)
        self.actionBackup = QtWidgets.QAction(MainWindow)
        self.actionBackup.setObjectName("actionBackup")
        self.actionQuit = QtWidgets.QAction(MainWindow)
        self.actionQuit.setObjectName("actionQuit")
        self.actionDark_Mode = QtWidgets.QAction(MainWindow)
        self.actionLight_Mode = QtWidgets.QAction(MainWindow)
        self.actionDark_Mode.setObjectName("actionDark_Mode")
        self.actionLight_Mode.setObjectName("actionLight_Mode")
        self.menuFile.addAction(self.actionQuit)
        self.actionQuit.triggered.connect(self.close_app)
        self.actionBackup.triggered.connect(self.create_backup)
        self.actionDark_Mode.triggered.connect(self.dark_mode)
        self.actionLight_Mode.triggered.connect(self.light_mode)
        self.menuFile.addAction(self.actionBackup)
        self.menuHelp.addAction(self.actionDark_Mode)
        self.menuHelp.addAction(self.actionLight_Mode)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menubar.addAction(self.menuHelp_2.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate(
            "MainWindow", "Silver Streak - Template Updater"))
        self.part_number_label.setText(_translate("MainWindow", "Part number:"))
        self.new_part_number_label.setText(
            _translate("MainWindow", "New Part Number:"))
        self.supplier_label.setText(_translate("MainWindow", "Supplier:"))
        self.new_supplier_label.setText(
            _translate("MainWindow", "New Supplier:"))
        self.price_label.setText(_translate("MainWindow", "Price:"))
        self.new_price_label.setText(_translate("MainWindow", "New Price:"))
        self.description_label.setText(_translate("MainWindow", "Description:"))
        self.new_description_label.setText(
            _translate("MainWindow", "New Description:"))
        self.quantity_label.setText(_translate("MainWindow", "Quantity:"))
        self.new_quantity_label.setText(
            _translate("MainWindow", "New Quantity:"))
        self.start_button.setText(_translate("MainWindow", "Start"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuHelp.setTitle(_translate("MainWindow", "Settings"))
        self.menuHelp_2.setTitle(_translate("MainWindow", "Help"))
        self.actionQuit.setText(_translate("MainWindow", "Quit"))
        self.actionBackup.setText(_translate("MainWindow", "Backup"))
        self.actionDark_Mode.setText(_translate("MainWindow", "Dark Mode"))
        self.actionLight_Mode.setText(_translate("MainWindow", "Light Mode"))

    def updater(self):
        logging.basicConfig(filename='{}-{}.log'.format(str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())), str(gethostname())), level=logging.DEBUG)
        logging.info("Starting new session from {}".format(gethostname()))
        self.completed = 0

        while self.completed < 100:
            self.completed += 1
            self.progressBar.setValue(self.completed)

        DIRECTORY = os.path.dirname(os.path.realpath(__file__))  
        EXTENTIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")

        TARGET = str(self.part_number_textbox.text())
        TARGET_REPLACEMENT = str(self.new_part_number_textbox.text())

        SUPPLIER = self.supplier_textbox.text()
        SUPPLIER_REPLACEMENT = self.new_supplier_textbox.text()

        DESCRIPTION = self.description_textbox.text()
        DESCRIPTION_REPLACEMENT = self.new_description_textbox.text()

        PRICE = float(self.price_textbox.text())
        PRICE_REPLACEMENT = float(self.new_price_textbox.text())

        QUANTITY = self.quantity_textbox.text()
        QUANTITY_REPLACEMENT = self.new_quantity_textbox.text()

        for (root, files) in os.walk(DIRECTORY):
            for file in files:
                if (file.endswith(EXTENTIONS)):
                    path = os.path.join(root, file)
                    self.console_output.append(str("Opening: {0}".format(file)))
                    logging.info("Opening: {0}".format(file))
                    wb = openpyxl.load_workbook(path, data_only=True)
                    ws = wb.active
                    target_in_wb = False
                    for ws in wb.worksheets:
                        for row in ws.iter_rows():
                            target_in_row = False
                            supplier_in_row = False
                            description_in_row = False
                            price_in_row = False
                            quantity_in_row = False
                            for cell in row:
                                if (cell.value == TARGET):
                                    self.console_output.append(
                                        str("PART {} STRING FOUND".format(TARGET)))
                                    logging.info("PART {} STRING FOUND".format(TARGET))
                                    self.console_output.append("Replacing {0} with {1} on row {2}".format(
                                        TARGET, TARGET_REPLACEMENT, ws._current_row))
                                    logging.info("Replacing {0} with {1} on row {2}".format(
                                        TARGET, TARGET_REPLACEMENT, ws._current_row))
                                    cell.value = TARGET_REPLACEMENT
                                    target_in_wb = True

                                    for cell in row:

                                        if (QUANTITY != ""):
                                            if (cell.value == QUANTITY):
                                                self.console_output.append(
                                                    "QUANTITY STRING FOUND")
                                                logging.info("QUANTITY STRING FOUND")
                                                self.console_output.append("Replacing {0} with {1} on row {2}".format(
                                                    QUANTITY, QUANTITY_REPLACEMENT, ws._current_row))
                                                logging.info("Replacing {0} with {1} on row {2}".format(
                                                    QUANTITY, QUANTITY_REPLACEMENT, ws._current_row))
                                                cell.value = QUANTITY_REPLACEMENT
                                                quantity_in_row = True

                                        if (DESCRIPTION != ""):
                                            if (cell.value == DESCRIPTION):
                                                self.console_output.append(
                                                    "DESCRIPTION STRING FOUND")
                                                logging.info("DESCRIPTION STRING FOUND")
                                                self.console_output.append("Replacing {0} with {1} on row {2}".format(
                                                    DESCRIPTION, DESCRIPTION_REPLACEMENT, ws._current_row))
                                                logging.info("Replacing {0} with {1} on row {2}".format(
                                                    DESCRIPTION, DESCRIPTION_REPLACEMENT, ws._current_row))
                                                cell.value = DESCRIPTION_REPLACEMENT
                                                description_in_row = True

                                        if (TARGET != ""):
                                            if (cell.value == TARGET):
                                                self.console_output.append("Replacing {0} with {1} on row {2}".format(
                                                    TARGET, TARGET_REPLACEMENT, ws._current_row))
                                                logging.info("Replacing {0} with {1} on row {2}".format(
                                                    TARGET, TARGET_REPLACEMENT, ws._current_row))
                                                cell.value = TARGET_REPLACEMENT
                                                supplier_in_row = True

                                        if (SUPPLIER != ""):
                                            if (cell.value == SUPPLIER):
                                                self.console_output.append(
                                                    "SUPPLIER STRING FOUND")
                                                logging.info("SUPPLIER STRING FOUND")
                                                self.console_output.append("Replacing {0} with {1} on row {2}".format(
                                                    SUPPLIER, SUPPLIER_REPLACEMENT, ws._current_row))
                                                logging.info("Replacing {0} with {1} on row {2}".format(
                                                    SUPPLIER, SUPPLIER_REPLACEMENT, ws._current_row))
                                                cell.value = SUPPLIER_REPLACEMENT
                                                supplier_in_row = True

                                        if (PRICE != ""):
                                            if (cell.value == PRICE):
                                                self.console_output.append(
                                                    "PRICE STRING FOUND")
                                                logging.info("PRICE STRING FOUND")
                                                self.console_output.append("Replacing {0} with {1} on row {2}".format(
                                                    PRICE, PRICE_REPLACEMENT, ws._current_row))
                                                logging.info("Replacing {0} with {1} on row {2}".format(
                                                    PRICE, PRICE_REPLACEMENT, ws._current_row))
                                                cell.value = PRICE_REPLACEMENT
                                                price_in_row = True

                                    if (target_in_row == False):
                                        self.console_output.append(
                                            "PART NOT FOUND")
                                        logging.info("PART NOT FOUND")
                                        pass
                                        if (supplier_in_row == False):
                                            self.console_output.append(
                                                "Supplier string not found")
                                            logging.info(
                                                "Supplier string not found")
                                            if (description_in_row == False):
                                                self.console_output.append(
                                                    "Description string not found")
                                                logging.info("Description string not found")
                                                if (price_in_row == False):
                                                    self.console_output.append(
                                                        "Price string not found")
                                                    logging.info("Price string not found")
                                                    if (quantity_in_row == False):
                                                        self.console_output.append(
                                                            "Quantity string not found")
                                                        logging.info("Quantity string not found")

                    if (target_in_wb == False):
                        self.console_output.append("PART NOT FOUND")
                        logging.info("PART NOT FOUND")

                    self.console_output.append(
                        "Saving: {} at {}\n".format(file, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                    logging.info(
                        "Saving: {} at {}\n".format(file, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                    wb.save(file)


    def create_backup(self):
        DIRECTORY = os.path.dirname(os.path.realpath(__file__))  
        EXTENTIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")
        self.console_output.append(str("Starting backup"))
        logging.info("Starting backup")

        if not os.path.exists('backup'):
            os.makedirs('backup')

        files = os.walk(DIRECTORY)
        for file in files:
            if (file.endswith(EXTENTIONS)):
                self.console_output.append(str("Copying: {0}".format(file)))
                logging.info("Copying: {0}".format(file))
                copyfile(file, "backup")

        

    def light_mode(self):
        app.setStyleSheet('QMainWindow{background-color: white;}')


    def dark_mode(self):
        app.setStyleSheet("""
                            QMainWindow{background-color: #1e1e2f;}
                            QLineEdit {
                                color: #a7a7ba;
                                border: 1px solid #353a53;
                                background-color: #27293d;
                                }
                            QLabel {
                                color: #a7a7ba;
                            }
                            QTextEdit {
                                background-color: #27293d;
                                color: #a7a7ba;
                                border: 1px solid #353a53;
                            }
                            QTreeView {
                                background-color: #27293d;
                                color: #a7a7ba;
                                border: 1px solid #353a53;
                            }
                            """)

    def close_app(self):
        sys.exit()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
