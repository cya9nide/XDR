"""Probe: does the VBA module compile + run via COM? Reports real state."""
import sys
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
        try:
            excel.Run("XDRMain.TestPaint")
            print("MACRO RUN OK — module compiles and TestPaint executed")
        except Exception as e:
            print("MACRO RUN FAILED:", repr(e))
            # the com_error carries the VBE message if it's a compile error
            if hasattr(e, "excepinfo") and e.excepinfo:
                print("detail:", e.excepinfo[2] if len(e.excepinfo) > 2 else e.excepinfo)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()


if __name__ == "__main__":
    main()