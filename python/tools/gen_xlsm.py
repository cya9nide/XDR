"""Build the XDR.xlsm workbook via Excel COM — Phase 1 layout + full VBA."""
import sys
from pathlib import Path

OUT = Path(r"..\excel\XDR.xlsm")
BAS = OUT.with_suffix(".bas")

# Excel COM colors are BGR ints (0x00BBGGRR)
DARK = 0x202020        # dark gray
RED = 0x4455DD         # red-ish
GREEN = 0x55CC55       # green-ish
AMBER = 0x00CCFF       # amber (BGR = 00 CC FF)
GRAY = 0x808080        # mid gray

# (label, value, target named range) → control cells on the SDR sheet
CONTROLS = [
    ("Frequency (MHz)", 100.0, "freq_mhz"),
    ("Refresh (ms)", 250, "refresh_ms"),
]

# Waterfall color ramp: low → high (BGR). 8 stops.
RAMP = [
    0x000000,
    0x000060,
    0x0000C8,
    0x0060C8,
    0x00C8C8,
    0x00C860,
    0xC8C800,
    0xC8C8C8,
]

VBA = r"""
Option Explicit
Const MODULE_TEST As String = "imported-ok"
' Phase 1: fake SDR waterfall in cells.
' Scarecrow: reads frames.csv, paints a 128x64 waterfall via cell colors.

Public Const S_DATA_DIR As String = "..\xdr_data\"
Public Const S_FRAME_FILE As String = "frames.csv"
Public Const WS_WATERFALL As String = "Waterfall"
Public Const BINS As Long = 128
Public Const ROWS As Long = 64
Public Const FIRST_COL As Long = 2
Public Const FIRST_ROW As Long = 2
Public Const FPS_AVG_N As Long = 30

Private m_frameSeq As Long
Private m_ring() As Double        ' ring buffer: ROWS x BINS
Private m_lastPaintMs As Double
Private m_paintCount As Long
Private m_diagWindow() As Double  ' rolling render-time window
Private m_running As Boolean
Private m_loopScheduled As Boolean

' ── named-range helpers ────────────────────────────────────────────────
Public Function GetRange(name As String) As Range
    On Error Resume Next
    Set GetRange = ThisWorkbook.Names(name).RefersToRange
End Function

Public Function GetVal(name As String) As Variant
    On Error Resume Next
    Dim r As Range
    Set r = ThisWorkbook.Names(name).RefersToRange
    If Not r Is Nothing Then GetVal = r.Value Else GetVal = Empty
End Function

Public Sub SetVal(name As String, v As Variant)
    On Error Resume Next
    Dim r As Range
    Set r = ThisWorkbook.Names(name).RefersToRange
    If Not r Is Nothing Then r.Value = v
End Sub

' ── waterfall core ─────────────────────────────────────────────────────
' Paints (rows x bins) cells at (firstRow, firstCol).
' Takes a 2D array (bins, rows) ascol; a later row index = older data.
Public Sub PaintWaterfall(frameData() As Double, ByVal rowCount As Long)
    Dim i As Long, j As Long, c As Long
    Dim r As Range, cell As Range
    On Error Resume Next
    Set r = ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + rowCount - 1, FIRST_COL + BINS - 1))
    r.Clear
    For j = 0 To rowCount - 1
        For i = 0 To BINS - 1
            c = ColorFor(frameData(i, j))
            Set cell = r.Cells(j + 1, i + 1)
            cell.Interior.Color = c
        Next i
    Next j
    On Error GoTo 0
End Sub

' Maps a 0..1 magnitude to a BGR color via the fixed ramp.
Public Function ColorFor(m As Double) As Long
    If m <= 0# Then ColorFor = GetRampColor(0): Exit Function
    If m >= 1# Then ColorFor = GetRampColor(7): Exit Function
    Dim pos As Double
    Dim lo As Long, hi As Long, t As Double
    pos = m * 6#
    lo = Int(pos): hi = lo + 1: t = pos - lo
    ColorFor = BlendColor(GetRampColor(lo), GetRampColor(hi), t)
End Function

Private Function GetRampColor(idx As Long) As Long
    Select Case idx
        Case 0: GetRampColor = &H0&
        Case 1: GetRampColor = &H6000&
        Case 2: GetRampColor = &HC80000&
        Case 3: GetRampColor = &HC86000&
        Case 4: GetRampColor = &H00C8C8&
        Case 5: GetRampColor = &H00C860&
        Case 6: GetRampColor = &HC8C800&
        Case 7: GetRampColor = &HC8C8C8&
        Case Else: GetRampColor = &H0&
    End Select
End Function

Public Function BlendColor(a As Long, b As Long, t As Double) As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = a And &HFF: g1 = (a \ &H100) And &HFF: b1 = (a \ &H10000) And &HFF
    r2 = b And &HFF: g2 = (b \ &H100) And &HFF: b2 = (b \ &H10000) And &HFF
    BlendColor = RGB( _
        r1 + (r2 - r1) * t, _
        g1 + (g2 - g1) * t, _
        b1 + (b2 - b1) * t)
End Function

' ── CSV reader (Phase 1 scarecrow) ─────────────────────────────────────
' Overwrite-in-place: read once per frame, parse, push latest into ring.
Public Sub ReadCSV()
    Dim p As String
    Dim fso As Object, ts As Object
    Dim line As String
    Dim parts() As String
    Dim i As Long, seq As Long
    Dim tStart As Double
    p = S_DATA_DIR & S_FRAME_FILE
    If Dir(p) = "" Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(p, 1)  ' ForReading
    line = ts.ReadAll
    ts.Close
    parts = Split(line, ",")
    If UBound(parts) <> BINS - 1 Then Exit Sub
    tStart = Timer
    seq = m_frameSeq + 1
    m_frameSeq = seq
    m_paintCount = m_paintCount + 1
    ' Shift ring down one row; new frame becomes the top row.
    Dim j As Long
    For j = ROWS - 1 To 1 Step -1
        For i = 0 To BINS - 1
            m_ring(i, j) = m_ring(i, j - 1)
        Next i
    Next j
    For i = 0 To BINS - 1
        m_ring(i, 0) = CDbl(parts(i))
    Next i
    PaintWaterfall m_ring, ROWS
    Dim dt As Double
    dt = (Timer - tStart) * 1000#
    m_lastPaintMs = dt
    If m_paintCount <= FPS_AVG_N Then
        If m_paintCount = 1 Then
            ReDim m_diagWindow(0 To FPS_AVG_N - 1)
        End If
        m_diagWindow(m_paintCount - 1) = dt
    Else
        ' keep a small rolling average via shift
        Dim k As Long
        For k = 0 To FPS_AVG_N - 2
            m_diagWindow(k) = m_diagWindow(k + 1)
        Next k
        m_diagWindow(FPS_AVG_N - 1) = dt
    End If
    UpdateDiag
End Sub

Public Function AvgRenderMs() As Double
    Dim s As Double, c As Long, k As Long
    s = 0: c = 0
    For k = 0 To FPS_AVG_N - 1
        If m_diagWindow(k) > 0 Then s = s + m_diagWindow(k): c = c + 1
    Next k
    If c = 0 Then AvgRenderMs = 0 Else AvgRenderMs = s / c
End Function

Public Sub UpdateDiag()
    Dim fps As Double
    Dim seq As Long
    On Error Resume Next
    seq = m_frameSeq
    If m_lastPaintMs > 0 Then
        fps = 1000# / m_lastPaintMs
    Else
        fps = 0
    End If
    SetVal "fps_cell", Format(fps, "0.0")
    SetVal "seq_cell", seq
    SetVal "render_ms_cell", Format(AvgRenderMs(), "0.00")
    SetVal "paints_cell", m_paintCount
    If m_running Then SetVal "status_cell", "RUNNING" Else SetVal "status_cell", "IDLE"
    On Error GoTo 0
End Sub

' ── loop control ───────────────────────────────────────────────────────
Public Sub StartLoop()
    If m_running Then Exit Sub
    m_running = True
    m_loopScheduled = False
    m_frameSeq = 0: m_paintCount = 0
    ReDim m_ring(0 To BINS - 1, 0 To ROWS - 1)
    ReDim m_diagWindow(0 To FPS_AVG_N - 1)
    SetVal "status_cell", "STARTING"
    Call ScheduleTick(0.05)
End Sub

Public Sub StopLoop()
    m_running = False
    SetVal "status_cell", "STOPPED"
End Sub

Private Sub ScheduleTick(delaySec As Double)
    If Not m_running Then Exit Sub
    m_loopScheduled = True
    Application.OnTime Now + delaySec, "XDRMain.Tick", Schedule:=True
End Sub

Public Sub Tick()
    m_loopScheduled = False
    If Not m_running Then Exit Sub
    ReadCSV
    Dim refreshMs As Double
    refreshMs = Val(GetVal("refresh_ms") & "")
    If refreshMs < 50 Then refreshMs = 250
    Dim delaySec As Double
    delaySec = refreshMs / 1000#
    Call ScheduleTick(delaySec)
End Sub

Public Sub TestPaint()
    ' Paint a gradient test frame so you can see the ramp without Python.
    Dim arr() As Double
    Dim i As Long, j As Long
    ReDim arr(0 To BINS - 1, 0 To ROWS - 1)
    For j = 0 To ROWS - 1
        For i = 0 To BINS - 1
            arr(i, j) = (i + j * 0.5) / (BINS + ROWS * 0.5)
        Next i
    Next j
    PaintWaterfall arr, ROWS
End Sub

Public Sub HardReset()
    m_running = False
    m_frameSeq = 0: m_paintCount = 0
    Erase m_ring
    Erase m_diagWindow
    SetVal "fps_cell", "--"
    SetVal "seq_cell", 0
    SetVal "render_ms_cell", "--"
    SetVal "paints_cell", 0
    SetVal "status_cell", "IDLE"
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + ROWS - 1, FIRST_COL + BINS - 1))
    r.Clear
    On Error GoTo 0
End Sub

Public Sub InitDisplay()
    SetVal "status_cell", "READY (Phase 1)"
    SetVal "fps_cell", "--"
    SetVal "paints_cell", 0
    SetVal "seq_cell", 0
    SetVal "render_ms_cell", "--"
End Sub
"""


def main() -> None:
    import win32com.client

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

        # header across sheets
        for name in ("SDR", "Waterfall", "Settings", "Diagnostics"):
            ws = wb.Worksheets(name)
            ws.Range("A1").Value = "EXCEL DEFINED RADIO - " + name.upper()
            ws.Range("A1").Font.Bold = True
            ws.Range("A1").Font.Color = RED
            ws.Range("A1").Interior.Color = DARK
            ws.Range("A1:A1").Font.Size = 14

        # ── SDR sheet: controls + status ──
        ws = wb.Worksheets("SDR")
        r = 3
        for label, value, nm in CONTROLS:
            ws.Cells(r, 1).Value = label
            ws.Cells(r, 2).Value = value
            ws.Cells(r, 2).NumberFormat = "0.0" if isinstance(value, float) else "0"
            wb.Names.Add(nm, ws.Cells(r, 2))
            ws.Cells(r, 3).Value = ""  # units column blank for now
            r += 1
        # status + diagnostics on SDR too
        ws.Cells(6, 1).Value = "Status"
        ws.Cells(7, 1).Value = "FPS (paint)"
        ws.Cells(8, 1).Value = "Frame"
        ws.Cells(9, 1).Value = "Render (ms)"
        ws.Cells(10, 1).Value = "Paints"
        ws.Cells(6, 2).Value = "IDLE"
        ws.Cells(7, 2).Value = "--"
        ws.Cells(8, 2).Value = 0
        ws.Cells(9, 2).Value = "--"
        ws.Cells(10, 2).Value = 0
        wb.Names.Add("status_cell", ws.Cells(6, 2))
        wb.Names.Add("fps_cell", ws.Cells(7, 2))
        wb.Names.Add("seq_cell", ws.Cells(8, 2))
        wb.Names.Add("render_ms_cell", ws.Cells(9, 2))
        wb.Names.Add("paints_cell", ws.Cells(10, 2))

        # ── Waterfall sheet: render region (128x64 at B2) ──
        wf = wb.Worksheets("Waterfall")
        wf.Range("B2:DU65").ColumnWidth = 2.1
        wf.Range("B2:DU65").RowHeight = 12
        wf.Range("A1").Value = "WATERFALL (128 bins x 64 rows)"  # keep header simple
        # label row
        wf.Cells(1, 1).Value = "Freq Label"
        wf.Cells(1, 1).Font.Bold = True

        # ── Settings sheet ──
        st = wb.Worksheets("Settings")
        st.Range("A1").Value = "Refresh (ms)"
        st.Range("B1").Value = 250
        st.Range("C1").Value = "100-1000"
        st.Range("A2").Value = "Theme"
        st.Range("B2").Value = "Classic"
        st.Range("C2").Value = "Classic | Vapor | Mono"

        # ── Diagnostics sheet ──
        dg = wb.Worksheets("Diagnostics")
        dg.Range("A1").Value = "Metric"
        dg.Range("B1").Value = "Value"
        dg.Range("A2").Value = "Frames read"
        dg.Range("B2").Value = 0
        dg.Range("A3").Value = "Last render (ms)"
        dg.Range("B3").Value = "--"
        dg.Range("A4").Value = "Avg render (ms)"
        dg.Range("B4").Value = "--"
        dg.Range("A5").Value = "Status"
        dg.Range("B5").Value = "IDLE"
        dg.Range("A6").Value = "Mode"
        dg.Range("B6").Value = "Phase 1 (CSV)"

        # ── VBA ──
        vbproj = wb.VBProject
        mod = vbproj.VBComponents.Add(1)
        mod.Name = "XDRMain"
        mod.CodeModule.AddFromString(VBA)

        OUT.parent.mkdir(parents=True, exist_ok=True)
        wb.SaveAs(str(OUT), FileFormat=52)
        print(f"WROTE {OUT} ({OUT.stat().st_size} bytes)")

        exp = OUT.with_suffix(".bas")
        mod.Export(str(exp))
        print(f"EXPORTED {exp}")

        wb.Close(SaveChanges=False)
        wb = None
        wb2 = excel.Workbooks.Open(str(OUT))
        names = sorted(n.Name for n in wb2.Names)
        print("REOPEN OK; named ranges:", ", ".join(names))
        print("sheets:", ", ".join(s.Name for s in wb2.Worksheets))
        # verify VBA module survived
        try:
            mod2 = wb2.VBProject.VBComponents("XDRMain")
            print("VBA module OK, lines:", mod2.CodeModule.CountOfLines)
        except Exception as e:
            print("VBA check failed:", e)
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