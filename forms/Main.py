# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QWidget, QErrorMessage, QPushButton, QFileDialog,
                             QVBoxLayout, QHBoxLayout, QLabel, QTableWidget)

from PyQt5.QtCore import *
from lib.ExcelController import ExcelController


class MainForm(QWidget):

    def __init__(self):
        super().__init__(flags=Qt.Window)

        self.excel_controller = None

        # init components
        self.label = QLabel("text")
        self.load_excel_btn = QPushButton("OK")
        self.table = QTableWidget()

        # init signals and slots
        self.init_signals_slots()

        # loading components on form
        self.init_ui()

    def init_signals_slots(self):
        self.load_excel_btn.clicked.connect(self.open_excel_file)

    def init_ui(self):
        h_top = QHBoxLayout()
        h_top.addWidget(self.load_excel_btn, alignment=Qt.AlignLeft | Qt.AlignTop)
        h_top.addWidget(self.label, alignment=Qt.AlignRight | Qt.AlignTop)

        h_table = QHBoxLayout()
        h_table.addWidget(self.table, alignment=Qt.AlignCenter)

        v_box = QVBoxLayout()
        v_box.addLayout(h_top)
        v_box.addLayout(h_table)

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

        data = self.excel_controller.load_data(excel_path)
        print(data)

    @staticmethod
    def show_error(message):
        error = QErrorMessage()
        error.showMessage(message)
        error.exec_()

    def closeEvent(self, event):
        # Закрываем соединение с екселем
        self.excel_controller.close_excel()
