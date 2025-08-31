# scripts/run_macro_xlwings.py
from pathlib import Path
import xlwings as xw

def run_macro(xlsm: str, macro_path: str, save=False):
    xlsm = str(Path(xlsm).resolve())
    with xw.App(visible=False, add_book=False) as app:
        wb = app.books.open(xlsm)
        # macro_path like "Module1.RunSimulation"
        macro = wb.macro(macro_path)
        result = macro()
        if save:
            wb.save()
        wb.close()
    return result
