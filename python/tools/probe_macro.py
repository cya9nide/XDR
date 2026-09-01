"""Minimal macro-execution probe — captures the real COM error and tests a fresh open."""
import sys
import time

import win32com.client

OUT = r"..\excel\XDR.xlsm"


def main() -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1
    wb = None
    try:
        wb = excel.Workbooks.Open(OUT)
        print("opened:", wb.Name)
        # macro execution test with full error capture
        try:
            excel.Run("XDRMain.TestPaint")
            print("TestPaint: OK")
        except Exception as e:
            print("TestPaint FAILED:", repr(e))
            # try to get the inner com_error detail
            if hasattr(e, "excepinfo"):
                print("excepinfo:", e.excepinfo)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()


if __name__ == "__main__":
    main()