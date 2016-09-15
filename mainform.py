# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QWidget, QPushButton, QFileDialog,
                             QVBoxLayout, QHBoxLayout, QLabel)

from PyQt5.QtCore import *
from lib.ExcelController import ExcelController


class MainForm(QWidget):
    excel_file = ''

    def __init__(self):
        super().__init__(flags=Qt.Window)

        # init components
        self.label = QLabel("text")
        self.load_excel_btn = QPushButton("OK")
        self.load_excel_btn.clicked.connect(self.open_excel_file)

        # loading components on form
        self.init_ui()

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
        excel = excel_controller.load_file(self.excel_file)
        self.label.setText(self.excel_file)
