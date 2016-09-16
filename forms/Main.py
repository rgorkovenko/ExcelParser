# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QWidget, QErrorMessage, QPushButton, QFileDialog,
                             QVBoxLayout, QHBoxLayout, QLabel)

from PyQt5.QtCore import *
from lib.ExcelController import ExcelController


class MainForm(QWidget):

    def __init__(self):
        super().__init__(flags=Qt.Window)

        self.excel_file = ''

        # init components
        self.label = QLabel("text")
        self.load_excel_btn = QPushButton("OK")

        # init signals and slots
        self.init_signals_slots()

        # loading components on form
        self.init_ui()

    def init_signals_slots(self):
        self.load_excel_btn.clicked.connect(self.open_excel_file)

    def init_ui(self):
        h_box = QHBoxLayout()
        h_box.addWidget(self.load_excel_btn, alignment=Qt.AlignLeft | Qt.AlignTop)
        h_box.addWidget(self.label, alignment=Qt.AlignRight | Qt.AlignTop)

        v_box = QVBoxLayout()
        v_box.addLayout(h_box)

        self.setLayout(v_box)

        self.setGeometry(300, 300, 300, 150)
        self.setWindowTitle('ExcelParser')

        self.show()

    def open_excel_file(self):
        self.excel_file = QFileDialog.getOpenFileName(filter="Excel files (*.xls *.xlsx)")[0]

        excel_controller = ExcelController()
        err, excel = excel_controller.create_excel()

        if err is not None:
            self.show_error(err)
            return

        self.label.setText(self.excel_file)

    @staticmethod
    def show_error(message):
        error = QErrorMessage()
        error.showMessage(message)
        error.exec_()
