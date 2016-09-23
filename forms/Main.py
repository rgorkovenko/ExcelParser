# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QWidget, QErrorMessage, QPushButton, QFileDialog,
                             QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QSizePolicy,
                             QSplitter)

from PyQt5.QtCore import *
from lib.ExcelController import ExcelController


class MainForm(QWidget):

    def __init__(self):
        super().__init__(flags=Qt.Window)

        self.excel_controller = None

        # init components
        self.load_excel_btn = QPushButton("OK")
        self.work_table = QTableWidget()
        self.work_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.result_table = QTableWidget()
        self.result_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # init signals and slots
        self.init_signals_slots()

        # loading components on form
        self.init_ui()

    def init_signals_slots(self):
        self.load_excel_btn.clicked.connect(self.open_excel_file)

    def load_window_settings(self):
        settings = QSettings('ExcelParser')
        self.setGeometry(settings.value('geometry', defaultValue=QRect(200, 200, 200, 200), type=QRect))

        state = settings.value('window_state')
        if state == Qt.WindowNoState:
            self.setWindowState(Qt.WindowNoState)
        elif state == Qt.WindowMaximized:
            self.setWindowState(Qt.WindowMaximized)
        elif state == Qt.WindowMinimized:
            self.setWindowState(Qt.WindowNoState)

    def save_window_settings(self):
        settings = QSettings('ExcelParser')
        settings.clear()

        settings.setValue('geometry', self.geometry())
        settings.setValue('window_state', self.windowState())

    def init_ui(self):
        h_top = QHBoxLayout()
        h_top.addWidget(self.load_excel_btn, alignment=Qt.AlignLeft | Qt.AlignTop)

        v_box = QVBoxLayout()
        v_box.addLayout(h_top)

        splitter_tables = QSplitter(Qt.Vertical)

        splitter_tables.addWidget(self.work_table)
        splitter_tables.addWidget(self.result_table)

        v_box.addWidget(splitter_tables)

        self.setLayout(v_box)
        self.setWindowTitle('ExcelParser')

        self.load_window_settings()

        self.show()

    def open_excel_file(self):
        excel_path = QFileDialog.getOpenFileName(filter="Excel files (*.xls *.xlsx)")[0]
        if not excel_path:
            return

        self.excel_controller = ExcelController()

        err = self.excel_controller.create_excel()
        if err is not None:
            self.show_error(err)
            return

        # Loading data in table
        self.table_load_data(self.excel_controller.load_data(excel_path))

    def table_load_data(self, data):
        if len(data) == 0:
            return

        self.work_table.setRowCount(len(data))
        self.work_table.setColumnCount(len(data[0]))

        for i in range(len(data)):
            for j in range(len(data[i])):
                item = data[i][j]
                self.work_table.setItem(i, j, QTableWidgetItem(item['value']))
                self.work_table.setSpan(i, j, item['merged_y'], item['merged_x'])

    @staticmethod
    def show_error(message):
        error = QErrorMessage()
        error.showMessage(message)
        error.exec_()

    def closeEvent(self, event):
        # Close excel connection
        if self.excel_controller:
            self.excel_controller.close_excel()

        # Save window settings
        self.save_window_settings()
