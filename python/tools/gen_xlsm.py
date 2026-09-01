"""Build the XDR.xlsm workbook skeleton via Excel COM."""
import os
import sys
from pathlib import Path

import win32com.client  # noqa: F401 — imported by gen_xlsm via pywin32

OUT = Path(r"..\excel\XDR.xlsm")
# Excel COM colors are BGR ints (0x00BBGGRR)
DARK = 0x202020        # dark gray
RED = 0x4455DD         # red-ish (BGR)
GREEN = 0x55CC55       # green-ish (BGR)

SHEETS = {
    "SDR": ["Frequency (MHz)", "100.0", "freq_mhz"],
    "Waterfall": ["Top Level", "0.5", "waterfall_top"],
    "Settings": ["Refresh (ms)", "250", "refresh_ms"],
    "Diagnostics": ["FPS", "--", "fps_cell"],
}

VBA = r"""
Attribute VB_Name = "XDRMain"
Option Explicit

' Excel Defined Radio - main module (skeleton)
' Rendering logic lands in Phase 1.

Public Const S_DATA_DIR As String = "..\xdr_data\"
Public Const S_FRAME_FILE As String = "frames.bin"

'Named-range helpers
Public Function GetRange(name As String) As Range
    On Error Resume Next
    Set GetRange = ThisWorkbook.Names(name).RefersToRange
End Function

Public Function GetVal(name As String) As Variant
    On Error Resume Next
    Dim r As Range
    Set r = ThisWorkbook.Names(name).RefersToRange
    If Not r Is Nothing Then
        GetVal = r.Value
    Else
        GetVal = Empty
    End If
End Function

Public Sub SetVal(name As String, v As Variant)
    On Error Resume Next
    Dim r As Range
    Set r = ThisWorkbook.Names(name).RefersToRange
    If Not r Is Nothing Then r.Value = v
End Sub

'Stub renderer; real implementation in later phases
Public Sub RenderFrame(frameBytes() As Byte, ByVal frameSeq As Long)
    ' TODO Phase 1: parse frame, push into ring, paint cells
End Sub

'Placeholder: read frames.bin and dispatch (Phase 1 fills the reader)
Public Sub ReadFrameFile()
    Dim f As Integer
    Dim data() As Byte
    Dim path As String
    path = S_DATA_DIR & S_FRAME_FILE
    If Dir(path) = "" Then Exit Sub
    f = FreeFile
    Open path For Binary Access Read As #f
    ReDim data(0 To 551)
    Get #f, , data
    Close #f
    ' TODO: parse struct; call RenderFrame
End Sub

Public Sub StartLoop()
    'Schedule the render loop (Phase 1 wires Application.OnTime)
    SetVal "fps_cell", "LOOP READY"
    SetVal "status_cell", "Idle"
End Sub

Public Sub StopLoop()
    SetVal "fps_cell", "--"
    SetVal "status_cell", "Stopped"
End Sub
"""


def main() -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Add()
        ws_main = wb.Worksheets(1)
        ws_main.Name = "SDR"
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)).Name = "Waterfall"
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)).Name = "Settings"
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)).Name = "Diagnostics"

        # Header across sheets
        for name in SHEETS:
            ws = wb.Worksheets(name)
            ws.Range("A1").Value = "EXCEL DEFINED RADIO - " + name.upper()
            ws.Range("A1").Font.Bold = True
            ws.Range("A1").Font.Color = RED
            ws.Range("A1").Interior.Color = DARK
            ws.Range("A1:A1").Font.Size = 14

        # Named controls on SDR sheet
        ws = wb.Worksheets("SDR")
        ws.Range("B3").Value = 100.0
        ws.Range("B3").NumberFormat = "0.0"
        wb.Names.Add("freq_mhz", ws.Range("B3"))
        ws.Range("C3").Value = "MHz"
        ws.Range("B4").Value = 250
        wb.Names.Add("refresh_ms", ws.Range("B4"))
        ws.Range("C4").Value = "ms"
        # status/fps named cells (used by VBA)
        ws.Range("B6").Value = "Idle"
        wb.Names.Add("status_cell", ws.Range("B6"))
        ws.Range("B7").Value = "--"
        wb.Names.Add("fps_cell", ws.Range("B7"))

        # Diagnostics sheet values
        wsD = wb.Worksheets("Diagnostics")
        wsD.Range("A1").Value = "FPS"
        wsD.Range("A2").Value = "Frame"
        wsD.Range("B1").Value = "--"
        wsD.Range("B2").Value = 0

        # Embed VBA
        vbproj = wb.VBProject
        mod = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
        mod.Name = "XDRMain"
        mod.CodeModule.AddFromString(VBA)

        OUT.parent.mkdir(parents=True, exist_ok=True)
        wb.SaveAs(str(OUT), FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
        print(f"WROTE {OUT} ({OUT.stat().st_size} bytes)")

        # Export the module for readable git diffs
        exp = OUT.with_suffix(".bas")
        mod.Export(str(exp))
        print(f"EXPORTED {exp}")

        # Reopen to verify integrity
        wb.Close(SaveChanges=False)
        wb = None
        wb2 = excel.Workbooks.Open(str(OUT))
        names = sorted(n.Name for n in wb2.Names)
        print("REOPEN OK; named ranges:", ", ".join(names))
        wb2.Close(SaveChanges=False)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()


if __name__ == "__main__":
    main()