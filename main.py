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


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(781, 581)
        font = QtGui.QFont()
        font.setFamily("Russo One")
        MainWindow.setFont(font)
        MainWindow.setWindowOpacity(1.0)
        MainWindow.setStyleSheet("")
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
        self.console_output = QtWidgets.QListView(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("xos4 Terminus")
        self.console_output.setFont(font)
        self.console_output.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.console_output.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.console_output.setObjectName("console_output")
        self.gridLayout_2.addWidget(self.console_output, 11, 0, 1, 3)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setProperty("value", 24)
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
        self.actionQuit = QtWidgets.QAction(MainWindow)
        self.actionQuit.setObjectName("actionQuit")
        self.actionDark_Mode = QtWidgets.QAction(MainWindow)
        self.actionDark_Mode.setObjectName("actionDark_Mode")
        self.menuFile.addAction(self.actionQuit)
        self.actionQuit.triggered.connect(self.close_app)
        self.actionDark_Mode.triggered.connect(self.dark_mode)

        self.menuHelp.addAction(self.actionDark_Mode)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menubar.addAction(self.menuHelp_2.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate(
            "MainWindow", "Silver Streak - Template Updater"))
        self.part_number_label.setText(_translate("MainWindow", "Part number"))
        self.new_part_number_label.setText(
            _translate("MainWindow", "NEW PART NUMBER"))
        self.supplier_label.setText(_translate("MainWindow", "SUPPLIER"))
        self.new_supplier_label.setText(
            _translate("MainWindow", "NEW SUPPLIER"))
        self.price_label.setText(_translate("MainWindow", "PRICE"))
        self.new_price_label.setText(_translate("MainWindow", "NEW PRICE"))
        self.description_label.setText(_translate("MainWindow", "DESCRIPTION"))
        self.new_description_label.setText(
            _translate("MainWindow", "NEW DESCRIPTION"))
        self.quantity_label.setText(_translate("MainWindow", "QUANTITY"))
        self.new_quantity_label.setText(
            _translate("MainWindow", "NEW QUANTITY"))
        self.start_button.setText(_translate("MainWindow", "START"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuHelp.setTitle(_translate("MainWindow", "Settings"))
        self.menuHelp_2.setTitle(_translate("MainWindow", "Help"))
        self.actionQuit.setText(_translate("MainWindow", "Quit"))
        self.actionDark_Mode.setText(_translate("MainWindow", "Dark Mode"))

    def updater(self):
        DIRECTORY = os.path.dirname(os.path.realpath(__file__))
        EXTENTIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")

        TARGET = self.part_number_textbox.text()
        TARGET_REPLACEMENT = self.new_part_number_textbox.text()

        SUPPLIER = self.supplier_textbox.text()
        SUPPLIER_REPLACEMENT = self.new_supplier_textbox.text()

        DESCRIPTION = self.description_textbox.text()
        DESCRIPTION_REPLACEMENT = self.new_description_textbox.text()

        PRICE = float(self.price_textbox.text())
        PRICE_REPLACEMENT = float(self.new_price_textbox.text())

        QUANTITY = self.quantity_textbox.text()
        QUANTITY_REPLACEMENT = self.new_quantity_textbox.text()

        for (root, dirs, files) in os.walk(DIRECTORY):
            for file in files:
                if (file.endswith(EXTENTIONS)):
                    start_time = time.time()
                    path = os.path.join(root, file)
                    print(
                        "\033[1m\033[96mOpening:\033[0m \033[1m\033[93m{0}\033[0m".format(file))
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
                                    print(
                                        "\033[1m\033[92mPART STRING FOUND\033[0m")
                                    print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                        TARGET, TARGET_REPLACEMENT, ws._current_row))
                                    cell.value = TARGET_REPLACEMENT
                                    target_in_wb = True

                                    for cell in row:

                                        if (QUANTITY != ""):
                                            if (cell.value == QUANTITY):
                                                print(
                                                    "\033[1m\033[92mQUANTITY STRING FOUND\033[0m")
                                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                                    QUANTITY, QUANTITY_REPLACEMENT, ws._current_row))
                                                cell.value = QUANTITY_REPLACEMENT
                                                quantity_in_row = True

                                        if (DESCRIPTION != ""):
                                            if (cell.value == DESCRIPTION):
                                                print(
                                                    "\033[1m\033[92mDESCRIPTION STRING FOUND\033[0m")
                                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                                    DESCRIPTION, DESCRIPTION_REPLACEMENT, ws._current_row))
                                                cell.value = DESCRIPTION_REPLACEMENT
                                                description_in_row = True

                                        if (TARGET != ""):
                                            if (cell.value == TARGET):
                                                print(
                                                    "\033[1m\033[92mTARGET STRING FOUND\033[0m")
                                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                                    TARGET, TARGET_REPLACEMENT, ws._current_row))
                                                cell.value = TARGET_REPLACEMENT
                                                supplier_in_row = True

                                        if (SUPPLIER != ""):
                                            if (cell.value == SUPPLIER):
                                                print(
                                                    "\033[1m\033[92mSUPPLIER STRING FOUND\033[0m")
                                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                                    SUPPLIER, SUPPLIER_REPLACEMENT, ws._current_row))
                                                cell.value = SUPPLIER_REPLACEMENT
                                                supplier_in_row = True

                                        if (PRICE != ""):
                                            if (cell.value == PRICE):
                                                print(
                                                    "\033[1m\033[92mPRICE STRING FOUND\033[0m")
                                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                                    PRICE, PRICE_REPLACEMENT, ws._current_row))
                                                cell.value = PRICE_REPLACEMENT
                                                price_in_row = True

                                    if (target_in_row == False):
                                        print(
                                            "\033[1m\033[91mPART NOT FOUND\033[0m")
                                        pass
                                        if (supplier_in_row == False):
                                            print(
                                                "\033[1m\033[91mSupplier string not found\033[0m")
                                            if (description_in_row == False):
                                                print(
                                                    "\033[1m\033[91mDescription string not found\033[0m")
                                                if (price_in_row == False):
                                                    print(
                                                        "\033[1m\033[91mPrice string not found\033[0m")
                                                    if (quantity_in_row == False):
                                                        print(
                                                            "\033[1m\033[91mQuantity string not found\033[0m")

                    if (target_in_wb == False):
                        print("\033[1m\033[91mPART NOT FOUND\033[0m")

                    print(
                        "\033[1m\033[96mSaving:\033[0m \033[1m\033[93m{}\033[0m at \033[1m\033[96m{}\033[0m\n".format(file, datetime.now()))
                    wb.save(file)

        print(
            "\033[95m[\033[0m\033[96m*\033[0m\033[95m]\033[0m \033[1m\033[96mDone in %s\033[0m" % round((time.time() - start_time), 2))

    def light_mode(self):
        app.setStyleSheet('QMainWindow{background-color: white;}')

    def dark_mode(self):
        app.setStyleSheet('QMainWindow{background-color: #1E1E1E;}')

    def close_app(self):
        sys.exit()
        #self.part_number_label.setText(self.part_number_textbox.text())

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
