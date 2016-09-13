# -*- coding: utf-8 -*-

import sys
import comtypes.client as cc

from mainform import MainForm
from PyQt5.QtWidgets import QApplication


if __name__ == '__main__':
    Excel = cc.CreateObject("Excel.Application")
    Excel.Workbooks.Open('D:test.xlsx')
    print(Excel.Version)
    print(Excel.Cells.CurrentRegion.Rows.Count)
    print(Excel.Cells(1, 1).Value())
    app = QApplication(sys.argv)

    ex = MainForm()

    sys.exit(app.exec_())
