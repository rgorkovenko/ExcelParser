import comtypes.client as cc


class ExcelController:
    def __init__(self):
        self.excel = None

    def create_excel(self):
        error = None

        # если объект екселя уже создан
        if not self.excel:
            try:
                self.excel = cc.CreateObject("Excel.Application")
            except OSError:
                error = "Excel not found in your system"

        return error

    def load_data(self, file_path):
        self.excel.Workbooks.Open(file_path)

        # Пока обрабатываем только первую страницу
        data = self.excel.Worksheets[1].UsedRange.Formula
        self.close_excel()
        return data

    def close_excel(self):
        for wb in self.excel.Workbooks:
            wb.Close(0)
            self.excel.Quit()

        del self.excel
