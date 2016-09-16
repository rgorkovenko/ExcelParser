# -*- coding: utf-8 -*-

import sys

from PyQt5.QtWidgets import QApplication

from forms.Main import MainForm


if __name__ == '__main__':
    app = QApplication(sys.argv)

    ex = MainForm()

    sys.exit(app.exec_())
