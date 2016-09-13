# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import QWidget


class MainForm(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setGeometry(300, 300, 300, 220)
        self.setWindowTitle('Icon')

        self.show()
