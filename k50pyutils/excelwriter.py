import pandas as pd
import xlwings as xw

class ExcelWriter:
    def __init__(self, path: str):
        """Initialize with a path to an Excel workbook."""
        self.wb = xw.Book(path)

    def __call__(self, df: pd.DataFrame, sheet_name: str = None):
        """Write DataFrame to A1 on a given sheet, including index.
        
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

        # Clear previous content
        sheet["A1"].expand().clear()

        # ✅ Write DataFrame with index
        sheet["A1"].options(index=False).value = df

        # Add Excel Table object
        sheet.tables.add(sheet["A1"].expand())

        # Auto-fit columns for readability
        sheet.autofit("columns")

    def get(self, sheet_name: str = None) -> pd.DataFrame:
        """Return the first Excel Table in the given sheet as a DataFrame.
        
        - If sheet_name is None: use first sheet.
        - Raises ValueError if no tables are found.
        """
        sheet = self.wb.sheets[0] if sheet_name is None else self.wb.sheets[sheet_name]

        if not sheet.tables:
            raise ValueError(f"No tables found in sheet '{sheet.name}'")

        table = sheet.tables[0]

        # ✅ Read as DataFrame (keep index if present in Excel table)
        df = table.range.options(pd.DataFrame, index=False, expand="table").value
        return df
