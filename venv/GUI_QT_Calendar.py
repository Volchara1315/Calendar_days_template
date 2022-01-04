from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox
import os
from ForDateTime import cr_l_cal
from ExcelDateCalendar import sav_l_cal_excel


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(840, 600)
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        MainWindow.setFont(font)
        MainWindow.setStyleSheet("background-color: rgb(208, 208, 208);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setStyleSheet("background-color: rgb(67, 67, 67);")
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(250, 480, 400, 60))
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("color: rgb(255, 30, 30);\n"
                                      "border-color: rgb(255, 0, 0);\n"
                                      "background-color: rgb(24, 24, 24);")
        self.pushButton.setObjectName("pushButton")
        self.title_label = QtWidgets.QLabel(self.centralwidget)
        self.title_label.setGeometry(QtCore.QRect(0, 0, 840, 80))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.title_label.sizePolicy().hasHeightForWidth())
        self.title_label.setSizePolicy(sizePolicy)
        self.title_label.setMaximumSize(QtCore.QSize(840, 16777215))
        font = QtGui.QFont()
        font.setPointSize(20)
        font.setBold(True)
        font.setItalic(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.title_label.setFont(font)
        self.title_label.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.title_label.setAutoFillBackground(False)
        self.title_label.setStyleSheet("color: rgb(255, 30, 30);\n"
                                       "border-color: rgb(255, 0, 0);\n"
                                       "background-color: rgb(24, 24, 24);")
        self.title_label.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.title_label.setLineWidth(0)
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.title_label.setObjectName("title_label")
        self.label_input_year = QtWidgets.QLabel(self.centralwidget)
        self.label_input_year.setGeometry(QtCore.QRect(40, 125, 280, 50))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.label_input_year.setFont(font)
        self.label_input_year.setStyleSheet("color: rgb(255, 30, 30);\n"
                                            "border-color: rgb(255, 0, 0);\n"
                                            "background-color: rgb(24, 24, 24);")
        self.label_input_year.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.label_input_year.setAlignment(QtCore.Qt.AlignCenter)
        self.label_input_year.setWordWrap(False)
        self.label_input_year.setObjectName("label_input_year")
        self.label_input_name = QtWidgets.QLabel(self.centralwidget)
        self.label_input_name.setGeometry(QtCore.QRect(40, 235, 280, 50))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.label_input_name.setFont(font)
        self.label_input_name.setStyleSheet("color: rgb(255, 30, 30);\n"
                                            "border-color: rgb(255, 0, 0);\n"
                                            "background-color: rgb(24, 24, 24);")
        self.label_input_name.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.label_input_name.setAlignment(QtCore.Qt.AlignCenter)
        self.label_input_name.setObjectName("label_input_name")
        self.label_input_path = QtWidgets.QLabel(self.centralwidget)
        self.label_input_path.setGeometry(QtCore.QRect(40, 335, 280, 60))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.label_input_path.setFont(font)
        self.label_input_path.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_input_path.setStyleSheet("color: rgb(255, 30, 30);\n"
                                            "border-color: rgb(255, 0, 0);\n"
                                            "background-color: rgb(24, 24, 24);")
        self.label_input_path.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.label_input_path.setAlignment(QtCore.Qt.AlignCenter)
        self.label_input_path.setObjectName("label_input_path")
        self.lineEdit_input_year = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_input_year.setGeometry(QtCore.QRect(340, 110, 450, 50))
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_input_year.setFont(font)
        self.lineEdit_input_year.setStyleSheet("background-color: rgb(208, 208, 208);")
        self.lineEdit_input_year.setMaxLength(4)
        self.lineEdit_input_year.setObjectName("lineEdit_input_year")
        self.lineEdit_input_name = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_input_name.setGeometry(QtCore.QRect(340, 220, 450, 50))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_input_name.setFont(font)
        self.lineEdit_input_name.setStyleSheet("background-color: rgb(208, 208, 208);")
        self.lineEdit_input_name.setObjectName("lineEdit_input_name")
        self.lineEdit_input_path = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_input_path.setGeometry(QtCore.QRect(340, 325, 450, 50))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_input_path.setFont(font)
        self.lineEdit_input_path.setStyleSheet("background-color: rgb(208, 208, 208);")
        self.lineEdit_input_path.setObjectName("lineEdit_input_path")
        self.label_detailed_year = QtWidgets.QLabel(self.centralwidget)
        self.label_detailed_year.setGeometry(QtCore.QRect(360, 165, 420, 30))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_detailed_year.setFont(font)
        self.label_detailed_year.setStyleSheet("color: rgb(255, 30, 30);\n"
                                               "border-color: rgb(255, 0, 0);\n"
                                               "background-color: rgb(24, 24, 24);")
        self.label_detailed_year.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.label_detailed_year.setAlignment(QtCore.Qt.AlignCenter)
        self.label_detailed_year.setObjectName("label_detailed_year")
        self.label_detailed_name = QtWidgets.QLabel(self.centralwidget)
        self.label_detailed_name.setGeometry(QtCore.QRect(360, 275, 420, 30))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_detailed_name.setFont(font)
        self.label_detailed_name.setStyleSheet("color: rgb(255, 30, 30);\n"
                                               "border-color: rgb(255, 0, 0);\n"
                                               "background-color: rgb(24, 24, 24);")
        self.label_detailed_name.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.label_detailed_name.setAlignment(QtCore.Qt.AlignCenter)
        self.label_detailed_name.setObjectName("label_detailed_name")
        self.label_detailed_path = QtWidgets.QLabel(self.centralwidget)
        self.label_detailed_path.setGeometry(QtCore.QRect(360, 380, 420, 51))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_detailed_path.setFont(font)
        self.label_detailed_path.setStyleSheet("color: rgb(255, 30, 30);\n"
                                               "border-color: rgb(255, 0, 0);\n"
                                               "background-color: rgb(24, 24, 24);")
        self.label_detailed_path.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.label_detailed_path.setAlignment(QtCore.Qt.AlignCenter)
        self.label_detailed_path.setObjectName("label_detailed_path")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusBar = QtWidgets.QStatusBar(MainWindow)
        self.statusBar.setObjectName("statusBar")
        MainWindow.setStatusBar(self.statusBar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.func_cl_btn()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "CREATING A TEMPLATE FOR ACTIVE INCOME STATEMENT"))
        self.pushButton.setText(_translate("MainWindow", "ЗБЕРЕГТИ В EXCEL ФАЙЛ"))
        self.title_label.setText(_translate("MainWindow", "СТВОРЕННЯ ШАБЛОНУ ДЛЯ ЗВІТУ АКТИВНОГО ДОХОДУ "))
        self.label_input_year.setText(_translate("MainWindow", "Введіть номер року:"))
        self.label_input_name.setText(_translate("MainWindow", "Введіть назву файлу:"))
        self.label_input_path.setText(_translate("MainWindow", "Введіть шлях для \n"
                                                               "збереженняв excel файлу:"))
        self.label_detailed_year.setText(_translate("MainWindow", "Введіть номер року, в діапазоні (1-9999)"))
        self.label_detailed_name.setText(
            _translate("MainWindow", "В назві файла не має бути наступні знаки: \\  / : * ? \" < > |"))
        self.label_detailed_path.setText(_translate("MainWindow", "Приклад шляху для Windows: D:\\Program\\Qt\n"
                                                                  " Приклад шляху для Linux: D:/Program/Qt"))

    def func_cl_btn(self):
        self.pushButton.clicked.connect(lambda: self.func_data_proc())

    def func_data_proc(self):
        num_er = 0
        n_year = 0
        str_name_file = ""
        str_path_file = ""

        # Перевірка поля вводу для номера року
        num_er = self.func_check_input_year(num_er)
        if num_er == 0:
            n_year = int(self.lineEdit_input_year.text())

        #Перевірка поля вводу для назви файлу
        num_er = self.func_check_input_name(num_er)
        if num_er == 0:
            str_name_file = self.lineEdit_input_name.text()

        # Перевірка поля вводу для шляху к каталогу
        num_er = self.func_check_input_path(num_er)
        if num_er == 0:
            str_path_file = self.lineEdit_input_path.text()

        if n_year != 0 and str_name_file != "" and str_path_file != "":
            str_path_file = str_path_file + '\\' + str_name_file + ".xlsx"
            self.func_save_excel(n_year, str_path_file, str_name_file)

    def func_save_excel(self, n_year, str_path_file, str_name_file):
        sav_l_cal_excel(cr_l_cal(n_year), n_year, str_path_file)
        self.func_successful_message(str_name_file)

    def func_error_message(self, str_title, str_text):
        error_message = QMessageBox()
        error_message.setWindowTitle(str_title)
        error_message.setText(str_text)
        error_message.setIcon(QMessageBox.Warning)
        error_message.exec_()

    def func_successful_message(self, str_name_file):
        successful_message = QMessageBox()
        successful_message.setWindowTitle('Успіх')
        successful_message.setText("Шаблон успішно збережений в файл " + str_name_file + ".xlsx")
        successful_message.exec_()

    def func_check_input_year(self, num_er):
        # Перевірка поля вводу для номера року
        str_in = self.lineEdit_input_year.text()
        if str_in == '':
            self.func_error_message("Помилка вводу номера року",
                                    'Введіть номера року')
            num_er = 1
        else:
            if str_in.isdigit():
                n_year = int(str_in)
                if n_year >= 1 and n_year <= 9999:
                    num_er = 0
                else:
                    self.func_error_message("Помилка вводу номера року",
                                            "Введіть номер року, в діапазоні (1-9999)")
                    num_er = 1
            else:
                self.func_error_message("Помилка вводу номера року",
                                        "Введіть номер року, у вигляді позитивного цілого числа")
                num_er = 1

        return num_er

    def func_check_input_name(self, num_er):
        # Перевірка поля вводу для назви файлу

        i = 0
        list_sym_ban = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        str_in = self.lineEdit_input_name.text()

        if str_in == '' and num_er == 0:
            self.func_error_message("Помилка вводу назви файла",
                                    'Введіть назву файла')
            num_er = 2
        if num_er == 0:
            for str_sym in str_in:
                for list_sym in list_sym_ban:
                    if str_sym == list_sym:
                        i = 1
            if i == 1:
                self.func_error_message("Помилка вводу назви файла",
                                        'У назві файла наступні символи заборонені: \\  / : * ? " < > |')
                num_er = 2
                print("str_in" + str_in)
            else:
                num_er = 0
                str_name_file = str_in

        return num_er

    def func_check_input_path(self, num_er):
        # Перевірка поля вводу для шляху к каталогу

        list_sym_dir_ban = ['\\', '/', '.']
        testpath = ""

        if self.lineEdit_input_path.text() != '' and num_er == 0:
            i = 0
            testpath = self.lineEdit_input_path.text()
            if os.path.exists(testpath):
                if os.path.isdir(testpath):
                    for str_sym in list_sym_dir_ban:
                        if str_sym == testpath[len(testpath) - 1]:
                            i = 1
                    if i == 0:
                        num_er = 0
                    else:
                        self.func_error_message("Помилка вводу шляху для збереженняв excel файлу:",
                                                'Заоронені символи в кінець шляху каталога . / \\')
                        num_er = 3
                if os.path.isfile(testpath):
                    self.func_error_message("Помилка вводу шляху для збереженняв excel файлу:",
                                            'Введіть шлях до каталога, а не до файлу')
                    num_er = 3
            else:
                self.func_error_message("Помилка вводу шляху для збереженняв excel файлу:",
                                        'Каталог не знайден')
                num_er = 3
        else:
            self.func_error_message("Помилка вводу шляху для збереженняв excel файлу:",
                                    'Введіть повну адресу каталога для збереженняв excel файлу')
            num_er = 3

        return num_er


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
