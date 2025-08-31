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

    def write_precip(self, series):
        ws = self.wb.Worksheets("input")
        ws.Cells.ClearContents()
        ws.Range("A1").Value = "Precip_mm"
        n = len(series)
        ws.Range("A2").Resize(n, 1).Value = [[float(x)] for x in series]

    def run_once(self, f=1.0, k1=0.05, h1=20.0, k2=0.3, S0=0.0):
        macro = f"'{self.wb.Name}'!Module1.RunTankModel"
        _n = self.excel.Application.Run(macro, f, k1, h1, k2, S0)
        ws = self.wb.Worksheets("discharge")
        last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp
        if last < 2:
            return []
        arr = ws.Range(f"A2:B{last}").Value
        return [row[0] for row in arr]  # Q only

if __name__ == "__main__":
    xlsm = Path("excel/TankModel.xlsm")
    precip = [0,0,5,12,7,0,0,0,18,25,2,0,0,9,0,0,0,4,0,0]

    with ExcelTankSession(xlsm) as xls:
        xls.excel.Application.Run(f"'{xls.wb.Name}'!Module1.SetupTankModel")
        xls.write_precip(precip)

        k1_values = [0.02, 0.04, 0.06, 0.08]
        for k1 in k1_values:
            q = xls.run_once(f=1.0, k1=k1, h1=15.0, k2=0.4, S0=0.0)
            print(f"k1={k1:.2f}  first 5 Q: {q[:5]}")
