# -*- coding: utf-8 -*-

import sys

from mainform import MainForm
from PyQt5.QtWidgets import QApplication


if __name__ == '__main__':
    app = QApplication(sys.argv)

    ex = MainForm()

    sys.exit(app.exec_())
