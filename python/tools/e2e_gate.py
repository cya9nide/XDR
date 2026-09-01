"""E2E gate via default-instance route — Dispatch (not DispatchEx) honors HKCU trusted locations."""
import time
from pathlib import Path

import win32com.client

OUT = r"..\excel\XDR.xlsm"


def main() -> None:
    excel = win32com.client.Dispatch("Excel.Application")  # default instance — honors HKCU TLS
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1
    wb = None
    try:
        wb = excel.Workbooks.Open(OUT)
        print("opened:", wb.Name)

        # 1. macro execution probe
        try:
            excel.Run("XDRMain.TestPaint")
            print("TestPaint: OK")
        except Exception as e:
            print("TestPaint FAILED:", repr(e))
            return

        # 2. set refresh + start loop
        wb.Names("refresh_ms").RefersToRange.Value = 250
        excel.Run("XDRMain.StartLoop")
        print("loop started; waiting 3.5s...")
        time.sleep(3.5)

        wf = wb.Worksheets("Waterfall")
        sdr = wb.Worksheets("SDR")
        print("status:", sdr.Range("B6").Value)
        print("fps_cell:", sdr.Range("B7").Value)
        print("seq_cell:", sdr.Range("B8").Value)
        print("render_ms_cell:", sdr.Range("B9").Value)
        print("paints_cell:", sdr.Range("B10").Value)

        # sample colors
        sample = ["B2", "B32", "B65", "H20", "N40", "Z50", "DU2", "DU32", "DU65"]
        colored = 0
        for addr in sample:
            c = wf.Range(addr).Interior.Color
            if c not in (0xFFFFFF, 0):
                colored += 1
            print(f"cell {addr}: color {c}")
        print(f"colored samples: {colored}/{len(sample)}")

        excel.Run("XDRMain.StopLoop")
        print("--- E2E done ---")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()


if __name__ == "__main__":
    main()