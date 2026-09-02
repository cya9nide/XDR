Attribute VB_Name = "XDRMain"

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
Private m_lastLine As String
Private m_palette(0 To 255) As Long   ' precomputed color ramp lookup
Private m_paletteInit As Boolean
Private m_cfReady As Boolean          ' region has a 3-color-scale CF rule
Private m_rateCount As Long           ' rolling fps meter
Private m_rateStart As Double
Private m_fpsRate As Double

' kernel32 sleep for render-loop pacing (32/64-bit safe)
#If VBA7 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' -- ring lifecycle -----------------------------------------------------
' Lazy-init so ReadCSV / StepOnce / StartLoop all work without ordering tricks.
Private Sub EnsureRing()
    On Error Resume Next
    Dim x As Long
    x = UBound(m_ring, 1)
    If Err.Number <> 0 Then
        Err.Clear
        ReDim m_ring(0 To BINS - 1, 0 To ROWS - 1)
        ReDim m_diagWindow(0 To FPS_AVG_N - 1)
    End If
    On Error GoTo 0
    If Not m_paletteInit Then
        InitPalette
        m_paletteInit = True
    End If
    If Not m_cfReady Then
        EnsureCF
    End If
End Sub

' -- named-range helpers ------------------------------------------------
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

' -- waterfall core -----------------------------------------------------
Public Sub PaintWaterfall(frameData() As Double, ByVal rowCount As Long)
    On Error GoTo EH
    Call PaintWaterfallCF(frameData, rowCount)
    Exit Sub
EH:
    ' CF path failed — fall back to the proven per-cell loop
    On Error Resume Next
    SetVal "err_cell", "CF paint failed (" & Err.Number & "); using per-cell"
    On Error GoTo 0
    PaintWaterfallCells frameData, rowCount
End Sub

' CF render: write raw values (one COM array write, universally supported) and
' let a 3-color-scale conditional format color the cells natively. Fastest
' reliable path; Excel's own render engine does the coloring.
Public Sub EnsureCF()
    If m_cfReady Then Exit Sub
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + ROWS - 1, FIRST_COL + BINS - 1))
    r.FormatConditions.Delete
    With r.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = &H0&            ' black (BGR)
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 50
        .ColorScaleCriteria(2).FormatColor.Color = &HC80000        ' red-ish
        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .ColorScaleCriteria(3).FormatColor.Color = &HC8C8C8        ' white-ish
    End With
    m_cfReady = True
    On Error GoTo 0
End Sub

Public Sub PaintWaterfallCF(frameData() As Double, ByVal rowCount As Long)
    Dim vals() As Double
    Dim i As Long, j As Long
    Dim flush As Boolean
    ReDim vals(1 To rowCount, 1 To BINS)
    For j = 0 To rowCount - 1
        For i = 0 To BINS - 1
            vals(j + 1, i + 1) = frameData(i, j)
        Next i
    Next j
    On Error Resume Next
    ' write raw values in one array call
    ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + rowCount - 1, FIRST_COL + BINS - 1)).Value = vals
    ' flip calc state to force CF re-eval visually
    If Application.Calculation = xlCalculationManual Then
        Application.Calculate
        flush = True
    End If
    On Error GoTo 0
    If flush Then Application.Calculation = xlCalculationManual
    SetVal "err_cell", ""
End Sub

' Per-cell fallback (proven in Phase 1) — slow, always works.
Public Sub PaintWaterfallCells(frameData() As Double, ByVal rowCount As Long)
    Dim i As Long, j As Long
    Dim r As Range
    Set r = ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + rowCount - 1, FIRST_COL + BINS - 1))
    For j = 0 To rowCount - 1
        For i = 0 To BINS - 1
            r.Cells(j + 1, i + 1).Interior.Color = ColorFor(frameData(i, j))
        Next i
    Next j
End Sub

' Maps a 0..1 magnitude to a BGR color via a 256-entry precomputed ramp.
Public Sub InitPalette()
    Dim i As Long
    For i = 0 To 255
        m_palette(i) = ColorForRamp(i / 255#)
    Next i
End Sub

Public Function ColorFor(m As Double) As Long
    If m <= 0# Then ColorFor = m_palette(0): Exit Function
    If m >= 1# Then ColorFor = m_palette(255): Exit Function
    ColorFor = m_palette(Int(m * 255#))
End Function

' Direct ramp (no palette), used by palette init itself.
Public Function ColorForRamp(m As Double) As Long
    If m <= 0# Then ColorForRamp = GetRampColor(0): Exit Function
    If m >= 1# Then ColorForRamp = GetRampColor(7): Exit Function
    Dim pos As Double
    Dim lo As Long, hi As Long, t As Double
    pos = m * 6#
    lo = Int(pos): hi = lo + 1: t = pos - lo
    ColorForRamp = BlendColor(GetRampColor(lo), GetRampColor(hi), t)
End Function

Private Function GetRampColor(idx As Long) As Long
    Select Case idx
        Case 0: GetRampColor = &H0&
        Case 1: GetRampColor = &H6000&
        Case 2: GetRampColor = &HC80000
        Case 3: GetRampColor = &HC86000
        Case 4: GetRampColor = &HC8C8&
        Case 5: GetRampColor = &HC860&
        Case 6: GetRampColor = &HC8C800
        Case 7: GetRampColor = &HC8C8C8
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

' -- CSV reader (Phase 1 scarecrow) -------------------------------------
' Overwrite-in-place: read once per frame, parse, push latest into ring.
Public Sub ReadCSV()
    On Error GoTo EH
    EnsureRing
    Dim p As String
    Dim fso As Object, ts As Object
    Dim line As String
    Dim parts() As String
    Dim i As Long, seq As Long
    Dim tStart As Double
    p = S_DATA_DIR & S_FRAME_FILE
    If Dir(p) = "" Then
        SetVal "err_cell", "NO FILE: " & p
        Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(p, 1)  ' ForReading
    line = ts.ReadAll
    ts.Close
    ' skip if nothing changed since last read (writer hasn't pushed a new frame)
    If line = m_lastLine Then Exit Sub
    m_lastLine = line
    parts = Split(line, ",")
    If UBound(parts) <> BINS - 1 Then
        SetVal "err_cell", "COUNT: " & UBound(parts) + 1 & " parts, want " & BINS
        Exit Sub
    End If
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
    Exit Sub
EH:
    SetVal "err_cell", Err.Number & ": " & Err.Description
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
    Dim seq As Long
    On Error Resume Next
    seq = m_frameSeq
    ' true throughput: count paints over a rolling ~1s window
    If m_rateStart = 0 Then m_rateStart = Timer
    m_rateCount = m_rateCount + 1
    Dim dtW As Double
    dtW = Timer - m_rateStart
    If dtW >= 1 Then
        m_fpsRate = m_rateCount / dtW
        m_rateStart = Timer
        m_rateCount = 0
    End If
    SetVal "fps_cell", Format(m_fpsRate, "0.0")
    SetVal "seq_cell", seq
    SetVal "render_ms_cell", Format(AvgRenderMs(), "0.00")
    SetVal "paints_cell", m_paintCount
    If m_running Then SetVal "status_cell", "RUNNING" Else SetVal "status_cell", "IDLE"
    On Error GoTo 0
End Sub

' -- loop control -------------------------------------------------------
' -- loop control (DoEvents render loop — the cursed while-loop) -------
' Runs until the user types STOP in the stop_flag cell (SDR B12) or calls StopLoop.
Public Sub StartLoop()
    If m_running Then Exit Sub
    m_running = True
    m_lastLine = ""
    m_rateCount = 0: m_rateStart = 0: m_fpsRate = 0
    Call EnsureRing
    SetVal "err_cell", ""
    SetVal "stop_flag", ""
    SetVal "status_cell", "RUNNING"
    Dim refreshMs As Double
    Dim tNext As Double
    tNext = 0
    Do While m_running
        If UCase(Trim(CStr(GetVal("stop_flag") & ""))) = "STOP" Then Exit Do
        If Timer >= tNext Then
            Call ReadCSV
            refreshMs = RefreshInterval()
            tNext = Timer + refreshMs / 1000#
        Else
            Sleep 10
        End If
        DoEvents
    Loop
    m_running = False
    SetVal "status_cell", "STOPPED"
    SetVal "stop_flag", ""
End Sub

Private Function RefreshInterval() As Double
    Dim v As Variant
    v = GetVal("refresh_ms")
    If IsEmpty(v) Or Val(v) < 50 Then
        RefreshInterval = 250
    Else
        RefreshInterval = Val(v)
    End If
End Function

Public Sub StopLoop()
    m_running = False
    SetVal "stop_flag", "STOP"
    SetVal "status_cell", "STOPPING"
End Sub

Public Sub StepOnce()
    ' One synchronous read+paint for debugging: reports any error to err_cell.
    On Error GoTo EH
    SetVal "err_cell", ""
    Call EnsureRing
    Call ReadCSV
    SetVal "status_cell", "STEP OK"
    Exit Sub
EH:
    SetVal "err_cell", Err.Number & ": " & Err.Description
    SetVal "status_cell", "ERROR"
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
    Call EnsureRing
End Sub


