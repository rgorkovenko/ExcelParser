# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QWidget, QErrorMessage, QPushButton, QFileDialog,
                             QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QSizePolicy)

from PyQt5.QtCore import *
from lib.ExcelController import ExcelController


class MainForm(QWidget):

    def __init__(self):
        super().__init__(flags=Qt.Window)

        self.excel_controller = None

        # init components
        self.load_excel_btn = QPushButton("OK")
        self.table = QTableWidget()
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # init signals and slots
        self.init_signals_slots()

        # loading components on form
        self.init_ui()

    def init_signals_slots(self):
        self.load_excel_btn.clicked.connect(self.open_excel_file)

    def init_ui(self):
        h_top = QHBoxLayout()
        h_top.addWidget(self.load_excel_btn, alignment=Qt.AlignLeft | Qt.AlignTop)

        v_box = QVBoxLayout()
        v_box.addLayout(h_top)
        v_box.addWidget(self.table)

        self.setLayout(v_box)

        self.setWindowTitle('ExcelParser')

        # todo set user settings
        self.setGeometry(300, 300, 500, 500)

        self.show()

    def open_excel_file(self):
        excel_path = QFileDialog.getOpenFileName(filter="Excel files (*.xls *.xlsx)")[0]

        self.excel_controller = ExcelController()

        err = self.excel_controller.create_excel()
        if err is not None:
            self.show_error(err)
            return

        # Загрузка данных в таблицу
        self.table_load_data(self.excel_controller.load_data(excel_path))

    def table_load_data(self, data):
        if len(data) == 0:
            return

        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(data[0]))

        for i in range(len(data)):
            for j in range(len(data[i])):
                self.table.setItem(i, j, QTableWidgetItem(data[i][j]))

    @staticmethod
    def show_error(message):
        error = QErrorMessage()
        error.showMessage(message)
        error.exec_()

    def closeEvent(self, event):
        # Закрываем соединение с екселем
        self.excel_controller.close_excel()
