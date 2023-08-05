import sys
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
import datetime
from gui import Ui_MainWindow

from control import write_to_excel, write_to_excel_out_of_date

class ExcelReaderWriter:
    def __init__(self):
        self.main_win = QMainWindow()
        self.uic = Ui_MainWindow()
        self.uic.setupUi(self.main_win)
        self.uic.btnSave.clicked.connect(self.save_excel)
        self.uic.btnLoad.clicked.connect(self.show_excel)


    def show(self):
        self.main_win.show()

    def getFileDialog(self, title, file_filter):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        return QFileDialog.getOpenFileName(self, title, '', file_filter, options=options)
    def show_excel(self):
        df = pd.read_excel('data.xlsx')
        
        row = 0
        self.uic.tableWidget.setRowCount(len(df))
        while(row < len(df)):
            self.uic.tableWidget.setItem(row,0,QtWidgets.QTableWidgetItem(str(df['STT'][row])))
            self.uic.tableWidget.setItem(row,1,QtWidgets.QTableWidgetItem(df['Dang Vien'][row]))
            self.uic.tableWidget.setItem(row,2,QtWidgets.QTableWidgetItem(df['Ten Thuoc'][row]))
            self.uic.tableWidget.setItem(row,3,QtWidgets.QTableWidgetItem(df['Ngay Nhap'][row]))
            self.uic.tableWidget.setItem(row,4,QtWidgets.QTableWidgetItem(df['Hoat Chat'][row]))
            self.uic.tableWidget.setItem(row,5,QtWidgets.QTableWidgetItem(df['Don vi san xuat'][row]))
            self.uic.tableWidget.setItem(row,6,QtWidgets.QTableWidgetItem(df['Dia chi'][row]))
            row = row + 1
    def save_excel(self):
        STT = 1
        txt_dangVien = self.uic.txt_DangVien.toPlainText()
        txt_tenThuoc = self.uic.txt_tenThuoc.toPlainText()
        txt_dateTime = self.uic.txt_dateNhap.text()
        txt_HoatChat = self.uic.txt_HoatChat.currentText()
        txt_donViSanXuat = self.uic.txt_DonViSanXuat.toPlainText()
        txt_diaChi = self.uic.txt_DiaChi.toPlainText()

        date_Nhap = datetime.datetime.strptime(txt_dateTime, "%d/%m/%Y")
        

        data = [STT,txt_dangVien,txt_tenThuoc,txt_dateTime,txt_HoatChat,txt_donViSanXuat,txt_diaChi]
        
        
        if (datetime.datetime.today() - date_Nhap).days > 20 :
            write_to_excel_out_of_date(data)
            message_box = QMessageBox()
            message_box.setWindowTitle("Information")
            message_box.setText("Da them vao file qua han thanh cong")
            message_box.setIcon(QMessageBox.Information)
            message_box.setStandardButtons(QMessageBox.Ok)
            message_box.setDefaultButton(QMessageBox.Ok)
            message_box.exec_()
        else : 
            write_to_excel(data)
            message_box = QMessageBox()
            message_box.setWindowTitle("Information")
            message_box.setText("Da them vao file data thanh cong.")
            message_box.setIcon(QMessageBox.Information)
            message_box.setStandardButtons(QMessageBox.OK)
            message_box.setDefaultButton(QMessageBox.Ok)
            message_box.exec_()
if __name__ == '__main__':

    app = QApplication(sys.argv)
    main_win = ExcelReaderWriter()
    main_win.show()
    sys.exit(app.exec())
