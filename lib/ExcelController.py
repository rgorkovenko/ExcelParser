import comtypes.client as cc


class ExcelController:
    def __init__(self):
        self.excel = None

    def create_excel(self):
        error = None
        try:
            self.excel = cc.CreateObject("Excel.Application")
        except OSError:
            error = "Excel not found in your system"

        return error, self.excel

    def test_excel(self, file_path):
        self.excel.Workbooks.Open(file_path)
        print(self.excel.Version)
        print(self.excel.Cells.CurrentRegion.Rows.Count)
        print(self.excel.Cells(1, 1).Value())
        for wb in self.excel.Workbooks:
            wb.Close(0)
            self.excel.Quit()

    def load_file(self, file_path):
        self.test_excel(file_path)
        return self.excel
