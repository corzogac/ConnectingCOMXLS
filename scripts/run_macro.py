from pathlib import Path
import pythoncom
import win32com.client as win32

def run_excel_macro(xlsm_path: Path, macro: str, save: bool=False, *macro_args):
    pythoncom.CoInitialize()  # required for notebooks/threads
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        xlsm_path = Path(xlsm_path).resolve()
        wb = excel.Workbooks.Open(str(xlsm_path), ReadOnly=False)
        qualified = f"'{wb.Name}'!{macro}"  # e.g. Module1.RunTankModel
        result = excel.Application.Run(qualified, *macro_args)
        if save:
            wb.Save()
        return result
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("xlsm", help="Path to .xlsm")
    p.add_argument("macro", help="ModuleName.MacroName")
    p.add_argument("--save", action="store_true")
    args, rest = p.parse_known_args()
    run_excel_macro(Path(args.xlsm), args.macro, args.save, *rest)
