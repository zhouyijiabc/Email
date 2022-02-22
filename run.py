from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
import sys
import os
from qt_material import apply_stylesheet
from auto_send_ui import Ui_MainWindow
from mail_handle import EmailHandle


class Run:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.main = QMainWindow()
        self.window = Ui_MainWindow()
        self.window.setupUi(self.main)
        self.window.file_path_lineEdit_4.setText(os.getcwd())

        self.window.ok_pushButton.clicked.connect(self.get_data)

        self.apply_stylesheet = apply_stylesheet(self.app, theme='dark_teal.xml')
        self.main.show()
        sys.exit(self.app.exec_())

    def get_data(self):
        self.email_addr = self.window.email_addr_lineEdit.text()
        self.passwd = self.window.passwd_lineEdit_2.text()
        self.host = self.window.host_lineEdit_3.text()
        self.file_path = self.window.file_path_lineEdit_4.text()

        print(self.email_addr, self.passwd, self.host)
        print(self.file_path)

    def send_mail(self):
        send = EmailHandle(self.email_addr, self.passwd, self.host)
        send.batch_send('email.xlsx')



if __name__ == '__main__':
    ds = Run()
    ds()
