import pandas as pd
import xlwings as xw

class ExcelWriter:
    def __init__(self, path: str):
        self.wb = xw.Book(path)

    def __call__(self, df: pd.DataFrame, sheet_name: str = None):
        if sheet_name is None:
            sheet = self.wb.sheets[0]
        else:
            if sheet_name in [s.name for s in self.wb.sheets]:
                sheet = self.wb.sheets[sheet_name]
            else:
                sheet = self.wb.sheets.add(name=sheet_name, after=self.wb.sheets[-1])

        # ✅ Convert un-writable objects to strings
        for col in df.columns:
            if df[col].apply(lambda x: isinstance(x, (list, tuple, np.ndarray))).any():
                df[col] = df[col].astype(str)

        sheet["A1"].expand().clear()
        sheet["A1"].options(index=False).value = df
        sheet.tables.add(sheet["A1"].expand())

