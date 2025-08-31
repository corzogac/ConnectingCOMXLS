from pathlib import Path
import pythoncom
import win32com.client as win32

class ExcelTankSession:
    def __init__(self, xlsm_path: Path):
        self.xlsm_path = Path(xlsm_path).resolve()
        self.excel = None
        self.wb = None

    def __enter__(self):
        pythoncom.CoInitialize()
        self.excel = win32.gencache.EnsureDispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.wb = self.excel.Workbooks.Open(str(self.xlsm_path), ReadOnly=False)
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.wb is not None:
                self.wb.Close(SaveChanges=False)
        finally:
            if self.excel is not None:
                self.excel.Quit()
            pythoncom.CoUninitialize()

    # One-time input write
    def write_precip(self, series):
        ws = self.wb.Worksheets("input")
        ws.Cells.ClearContents()
        ws.Range("A1").Value = "Precip_mm"
        n = len(series)
        col = [[float(x)] for x in series]
        ws.Range("A2").Resize(n, 1).Value = col

    # Run model with parameters; returns list of Q (and optionally S if you need)
    def run_once(self, f=1.0, k1=0.05, h1=20.0, k2=0.3, S0=0.0):
        # call the VBA: "'WorkbookName'!Module1.RunTankModel"
        qualified_macro = f"'{self.wb.Name}'!Module1.RunTankModel"
        _n = self.excel.Application.Run(qualified_macro, f, k1, h1, k2, S0)

        ws = self.wb.Worksheets("discharge")
        # xlUp = -4162 (avoid depending on constants module)
        last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        if last < 2:
            return []

        arr = ws.Range(f"A2:B{last}").Value  # (Q,S)
        q_series = [row[0] for row in arr]
        return q_series

if __name__ == "__main__":
    # ---- example usage ----
    xlsm = Path("TankModel.xlsm")  # adjust if your file is elsewhere
    # Example synthetic precipitation (mm)
    precip = [0, 0, 5, 12, 7, 0, 0, 0, 18, 25, 2, 0, 0, 9, 0, 0, 0, 4, 0, 0]

    with ExcelTankSession(xlsm) as xls:
        # optional: ensure workbook is set up (only needed once)
        xls.excel.Application.Run(f"'{xls.wb.Name}'!Module1.SetupTankModel")
        # write input once
        xls.write_precip(precip)

        # sweep one parameter (k1) and collect results
        k1_values = [0.02, 0.04, 0.06, 0.08]
        results = {}
        for k1 in k1_values:
            q = xls.run_once(f=1.0, k1=k1, h1=15.0, k2=0.4, S0=0.0)
            results[k1] = q

        # minimal output preview
        for k1, q in results.items():
            print(f"k1={k1:.2f} → first 5 Q: {q[:5]}")
