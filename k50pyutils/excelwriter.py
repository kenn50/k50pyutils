import pandas as pd
import xlwings as xw

class ExcelWriter:
    def __init__(self, path: str):
        self.wb = xw.Book(path)

    def __call__(self, df: pd.DataFrame, sheet_name: str = None):
        """Write DataFrame to A1 on a given sheet.
        
        - If sheet_name is None: use first sheet.
        - If sheet_name exists: use it.
        - If sheet_name doesn't exist: create at end.
        """
        if sheet_name is None:
            sheet = self.wb.sheets[0]
        else:
            if sheet_name in [s.name for s in self.wb.sheets]:
                sheet = self.wb.sheets[sheet_name]
            else:
                sheet = self.wb.sheets.add(name=sheet_name, after=self.wb.sheets[-1])

        sheet["A1"].expand().clear()
        sheet["A1"].options(index=False).value = df
        sheet.tables.add(sheet["A1"].expand())
