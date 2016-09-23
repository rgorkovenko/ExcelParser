import comtypes.client as cc


class ExcelController:
    def __init__(self):
        self.excel = None

    def create_excel(self):
        error = None

        # If excel object was created
        if not self.excel:
            try:
                self.excel = cc.CreateObject("Excel.Application")
            except OSError:
                error = "Excel not found in your system"

        return error

    def load_data(self, file_path):
        self.excel.Workbooks.Open(file_path)

        # We process only the first page
        excel_range = self.excel.Worksheets[1].UsedRange

        # parsing
        result = []
        for row in range(excel_range.Rows.Count):
            result_row = []
            for col in range(excel_range.Columns.Count):
                cell = excel_range.Cells(row + 1, col + 1)
                item = {'value': cell.Value(),
                        'is_merged': cell.MergeCells,
                        'merged_x': 1,
                        'merged_y': 1}

                area = cell.MergeArea()
                if type(area) is tuple and cell.Value():
                    item['merged_x'] = len(area[0])
                    item['merged_y'] = len(area)

                result_row.append(item)
            result.append(result_row)

        # return result
        self.close_excel()
        return result

    def close_excel(self):
        if self.excel is None:
            return

        for wb in self.excel.Workbooks:
            wb.Close(0)
            self.excel.Quit()

        del self.excel
