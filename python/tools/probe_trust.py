"""Inspect Trust Center macro policy from the live COM instance."""
import sys

import win32com.client


def main() -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1
    try:
        # Try common property names for the Trust Center macro state
        for attr in ("TrustCenter", "TrustCenterAccess", "MacroSettings"):
            try:
                v = getattr(excel.Application, attr)
                print(f"{attr}:", v)
            except Exception as e:
                print(f"{attr}: unavailable ({type(e).__name__})")
        # The Workbook-level AutomationSecurity reflects the app policy
        print("workbook-level AutomationSecurity:", excel.AutomationSecurity)
        # Try to read the VBAWarnings-equivalent from the object model
        try:
            print("app macro setting:", excel.Application.MacroSecurity)
        except Exception as e:
            print("app MacroSecurity unavailable:", type(e).__name__)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()