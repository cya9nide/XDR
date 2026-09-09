Attribute VB_Name = "XDRMain"

Option Explicit
Const MODULE_TEST As String = "imported-ok"
' Fake SDR waterfall in cells.
' Scarecrow: reads frames.csv, paints a 128x64 waterfall via cell colors.

Public Const S_DATA_DIR As String = "..\xdr_data\"
Public Const S_FRAME_FILE As String = "frames.bin"
Public Const WS_WATERFALL As String = "Waterfall"
' frame flags (protocol.py)
Public Const FLAG_SR_VALID As Long = 1
Public Const FLAG_LIVE As Long = 2
Public Const bins As Long = 128
Public Const ROWS As Long = 64
Public Const FIRST_COL As Long = 2
Public Const FIRST_ROW As Long = 2
Public Const FPS_AVG_N As Long = 30

' frames.bin record � layout aligned so VBA reads it natively with Get #f,, udt
Public Type XDRFrame
    magic As Long          ' 0..3    ("XDR1")
    sequence As Long       ' 4..7
    timestamp As Double    ' 8..15
    bins(0 To bins - 1) As Single   ' 16..527 (512 bytes)
    sample_rate As Double  ' 528..535
    peak_idx As Long       ' 536..539
    peak_val As Single     ' 540..543
    flags As Long          ' 544..547
    reserved As Long       ' 548..551
End Type

Private m_frameSeq As Long
Private m_ring() As Double        ' ring buffer: ROWS x BINS
Private m_lastPaintMs As Double
Private m_paintCount As Long
Private m_diagWindow() As Double  ' rolling render-time window
Private m_running As Boolean
Private m_loopScheduled As Boolean
Private m_lastSeq As Long
Private m_gapCount As Long                ' skipped-sequence counter (diagnostics)
Private m_peakIdx As Long
Private m_peakVal As Single
Private m_palette(0 To 255) As Long   ' precomputed color ramp lookup
Private m_paletteInit As Boolean
Private m_cfReady As Boolean          ' region has a 3-color-scale CF rule
Private m_rateCount As Long           ' rolling fps meter
Private m_rateStart As Double
Private m_fpsRate As Double
Private m_startPump As Boolean        ' heartbeat says "keep the pump going"

' kernel32 sleep for render-loop pacing (32/64-bit safe)
#If VBA7 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Shell for launching the engine headless (Form-button start; no Alt-F11)
#If VBA7 Then
Public Declare PtrSafe Function ShellExecuteW Lib "shell32" ( _
    ByVal hwnd As LongPtr, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
Public Declare Function ShellExecuteW Lib "shell32" ( _
    ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

' -- ring lifecycle -----------------------------------------------------
' Lazy-init so ReadCSV / StepOnce / StartLoop all work without ordering tricks.
Private Sub EnsureRing()
    On Error Resume Next
    Dim x As Long
    x = UBound(m_ring, 1)
    If Err.Number <> 0 Then
        Err.Clear
        ReDim m_ring(0 To bins - 1, 0 To ROWS - 1)
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
    ' CF path failed � fall back to the proven per-cell loop
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
    Call ApplyTheme
End Sub

Public Sub PaintWaterfallCF(frameData() As Double, ByVal rowCount As Long)
    Dim vals() As Double
    Dim i As Long, j As Long
    Dim flush As Boolean
    ReDim vals(1 To rowCount, 1 To bins)
    For j = 0 To rowCount - 1
        For i = 0 To bins - 1
            vals(j + 1, i + 1) = frameData(i, j)
        Next i
    Next j
    On Error Resume Next
    ' write raw values in one array call
    ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + rowCount - 1, FIRST_COL + bins - 1)).Value = vals
    ' flip calc state to force CF re-eval visually
    If Application.Calculation = xlCalculationManual Then
        Application.Calculate
        flush = True
    End If
    On Error GoTo 0
    If flush Then Application.Calculation = xlCalculationManual
    SetVal "err_cell", ""
End Sub

' Per-cell fallback — slow, always works.
Public Sub PaintWaterfallCells(frameData() As Double, ByVal rowCount As Long)
    Dim i As Long, j As Long
    Dim r As Range
    Set r = ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + rowCount - 1, FIRST_COL + bins - 1))
    For j = 0 To rowCount - 1
        For i = 0 To bins - 1
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

' -- cmd.bin writer (Excel ? Python control) ---------------------------
' Writes the 16-byte XDRCMD record: freq_hz u32, sample_rate u32, gain_x10 u32, refresh_ms u32.
' Called from SDR sheet Worksheet_Change; also callable as XDRMain.WriteCmd 0,0,0,0 to clear.
Public Sub WriteCmd(Optional freqHz As Long = 0, Optional srHz As Long = 0, _
                    Optional gainX10 As Long = 0, Optional refreshMs As Long = 0)
    On Error GoTo EH
    Dim f As Integer
    f = FreeFile
    Open S_DATA_DIR & "cmd.bin" For Binary Access Write As #f
    Dim b(1 To 16) As Byte
    Dim v As Long
    v = freqHz:   b(1) = v And &HFF: b(2) = (v \ &H100) And &HFF: b(3) = (v \ &H10000) And &HFF: b(4) = (v \ &H1000000) And &HFF
    v = srHz:     b(5) = v And &HFF: b(6) = (v \ &H100) And &HFF: b(7) = (v \ &H10000) And &HFF: b(8) = (v \ &H1000000) And &HFF
    v = gainX10:  b(9) = v And &HFF: b(10) = (v \ &H100) And &HFF: b(11) = (v \ &H10000) And &HFF: b(12) = (v \ &H1000000) And &HFF
    v = refreshMs: b(13) = v And &HFF: b(14) = (v \ &H100) And &HFF: b(15) = (v \ &H10000) And &HFF: b(16) = (v \ &H1000000) And &HFF
    Put #f, , b
    Close #f
    Exit Sub
EH:
    On Error Resume Next
    Close #f
    SetVal "err_cell", "WriteCmd: " & Err.Number & " " & Err.Description
End Sub

' -- SDR sheet Worksheet_Change: any control edit ? cmd.bin -------------
Public Sub SDRSheet_Change(ByVal Target As Range)
    On Error GoTo EH
    If Target Is Nothing Then Exit Sub
    If Target.Column <> 2 Then Exit Sub
        Dim r As Long
        r = Target.Row
        If r < 3 Or r > 6 Then Exit Sub   ' ONLY the control cells (rows 3-6); the
                                          ' status block lives in col B rows 8+ and
                                          ' must NOT re-trigger this handler
        Dim v As Variant
        v = Target.Value
        If Not IsNumeric(v) Then Exit Sub
        Select Case r
        Case 3   ' Frequency (MHz) ? freq_hz (�1e6)
            Call WriteCmd(CLng(CDbl(v) * 1000000#))
        Case 4   ' Sample Rate (MHz) ? sample_rate
            Call WriteCmd(0, CLng(CDbl(v) * 1000000#))
        Case 5   ' Gain (dB) ? gain_x10
            Call WriteCmd(0, 0, CLng(CDbl(v) * 10))
        Case 6   ' Refresh (ms) ? refresh_ms
            Call WriteCmd(0, 0, 0, CLng(v))
        Case Else
            Exit Sub
    End Select
    Exit Sub
EH:
    SetVal "err_cell", "Change: " & Err.Number & " " & Err.Description
End Sub

' -- preset auto-load -----------------------------------------
' Auto-loads the preset when the preset_idx cell (F23) is edited.
Public Sub PresetIdx_Change(ByVal Target As Range)
    On Error GoTo EH
    If Target Is Nothing Then Exit Sub
    If Target.Column <> 6 Then Exit Sub        ' col F
    If Target.Row <> 23 Then Exit Sub          ' F23 = preset_idx
    If Not IsNumeric(Target.Value) Then Exit Sub
    Call LoadPreset
    Exit Sub
EH:
    SetVal "err_cell", "PresetIdx: " & Err.Number & " " & Err.Description
End Sub
' -- binary reader --------------------------------------------
' Reads the latest frames.bin record via native Get (aligned UDT), dedupes by
' sequence, pushes into the ring, and paints via CF.
Public Sub ReadFrame()
    On Error GoTo EH
    EnsureRing
    Dim p As String
    Dim f As Integer
    Dim fr As XDRFrame
    p = S_DATA_DIR & S_FRAME_FILE
    If Dir(p) = "" Then
        SetVal "err_cell", "NO FILE: " & p
        SetVal "connected_cell", "NO"
        SetVal "srate_cell", "--"
        SetVal "signal_cell", "--"
        Exit Sub
    End If
    f = FreeFile
    Open p For Binary Access Read As #f
    Get #f, , fr
    Close #f
    If fr.magic <> &H58445231 Then
        SetVal "err_cell", "BAD MAGIC: " & Hex(fr.magic)
        SetVal "connected_cell", "NO"
        Exit Sub
    End If
    If fr.sequence = m_lastSeq Then Exit Sub   ' no new frame
    If fr.sequence > m_lastSeq + 1 Then m_gapCount = m_gapCount + (fr.sequence - m_lastSeq - 1)
    m_lastSeq = fr.sequence
    m_frameSeq = fr.sequence
    m_peakIdx = fr.peak_idx
    m_peakVal = fr.peak_val
    ' status panel — pull what the frame already carries
    SetVal "connected_cell", "YES"
    If (fr.flags And FLAG_SR_VALID) <> 0 Then
        SetVal "srate_cell", Format(fr.sample_rate / 1000000#, "0.0") & " MSPS"
    Else
        SetVal "srate_cell", "--"
    End If
    SetVal "signal_cell", Format(m_peakVal, "0.00")
    Call UpdateSMeter
    m_paintCount = m_paintCount + 1
    ' shift ring down one row; new frame becomes the top row
    Dim i As Long, j As Long
    For j = ROWS - 1 To 1 Step -1
        For i = 0 To bins - 1
            m_ring(i, j) = m_ring(i, j - 1)
        Next i
    Next j
    Dim tStart As Double
    tStart = Timer
    For i = 0 To bins - 1
        m_ring(i, 0) = fr.bins(i)
    Next i
    PaintWaterfall m_ring, ROWS
    m_lastPaintMs = (Timer - tStart) * 1000#
    ' rolling render-time window
    If m_paintCount <= FPS_AVG_N Then
        If m_paintCount = 1 Then ReDim m_diagWindow(0 To FPS_AVG_N - 1)
        m_diagWindow(m_paintCount - 1) = m_lastPaintMs
    Else
        Dim k As Long
        For k = 0 To FPS_AVG_N - 2
            m_diagWindow(k) = m_diagWindow(k + 1)
        Next k
        m_diagWindow(FPS_AVG_N - 1) = m_lastPaintMs
    End If
    UpdateDiag
    Exit Sub
EH:
    SetVal "err_cell", Err.Number & ": " & Err.Description
End Sub

' -- S-meter -------------------------------------------------
' Maps the live peak magnitude (m_peakVal, roughly 0..1.2) onto a 0-10
' bar of 10 cells (named range smeter_bar) + a numeric cell (smeter_cell).
' Called from ReadFrame every new frame.
Public Sub UpdateSMeter()
    On Error Resume Next
    Dim lvl As Double
    lvl = m_peakVal
    Dim n As Long
    n = Int((lvl / 1#) * 10#)
    If n < 0 Then n = 0
    If n > 10 Then n = 10
    SetVal "smeter_cell", n
    Dim r As Range, i As Long
    Set r = ThisWorkbook.Names("smeter_bar").RefersToRange
    For i = 1 To r.Cells.Count
        If i <= n Then
            r.Cells(i).Interior.Color = RGB(120, 200, 80)   ' lit (green)
        Else
            r.Cells(i).Interior.Color = RGB(60, 60, 60)     ' unlit (gray)
        End If
    Next i
    On Error GoTo 0
End Sub

' -- presets -------------------------------------------------
' Loads the preset selected by preset_idx from the named preset_freqs
' column, pushes it into the Frequency control + cmd.bin, and re-tunes.
Public Sub LoadPreset()
    On Error GoTo EH
    Dim i As Long, f As Variant
    i = CLng(GetVal("preset_idx"))
    Dim rng As Range
    Set rng = ThisWorkbook.Names("preset_freqs").RefersToRange
    If rng Is Nothing Then Exit Sub
    If i >= 1 And i <= rng.ROWS.Count Then
        f = rng.Cells(i, 1).Value
        If IsNumeric(f) And f > 0 Then
            SetVal "freq_mhz", CDbl(f)
            Call WriteCmd(CLng(CDbl(f) * 1000000#))
            SetVal "status_cell", "TUNED " & Format(f, "0.0") & " MHz"
            SetVal "err_cell", ""
        End If
    End If
    Exit Sub
EH:
    SetVal "err_cell", "LoadPreset: " & Err.Number & " " & Err.Description
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
    SetVal "frames_read_cell", m_paintCount
    SetVal "gaps_cell", m_gapCount
    SetVal "peak_cell", m_peakIdx & " (" & Format(m_peakVal, "0.00") & ")"
    If m_running Then SetVal "status_cell", "RUNNING" Else SetVal "status_cell", "IDLE"
    On Error GoTo 0
End Sub

' -- themes ---------------------------------------------------
' Reads the Theme cell (Settings B2, named theme_cell), clears the region,
' writes raw values, and re-paints the color scale for the chosen theme.
Public Sub ApplyTheme()
    On Error GoTo EH
    Dim theme As String
    theme = UCase(Trim(CStr(GetVal("theme_cell"))))
    If theme = "CLASSIC" Then Call ApplyThemeClassic
    If theme = "VAPOR" Then Call ApplyThemeVapor
    If theme = "MONO" Then Call ApplyThemeMono
    m_cfReady = True
    Exit Sub
EH:
    On Error Resume Next
    m_cfReady = True
    SetVal "err_cell", "ApplyTheme: " & Err.Number & " " & Err.Description
End Sub

Private Sub ApplyThemeClassic()
    On Error Resume Next
    Dim r As Range
    Set r = ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + ROWS - 1, FIRST_COL + bins - 1))
    r.FormatConditions.Delete
    With r.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = &H0&
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 40
        .ColorScaleCriteria(2).FormatColor.Color = &HC80000
        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .ColorScaleCriteria(3).FormatColor.Color = &HC8C8C8
    End With
    On Error GoTo 0
End Sub

Private Sub ApplyThemeVapor()
    On Error Resume Next
    Dim r As Range
    Set r = ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + ROWS - 1, FIRST_COL + bins - 1))
    r.FormatConditions.Delete
    With r.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = &H101020       ' dark desat purple
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 45
        .ColorScaleCriteria(2).FormatColor.Color = &HC03050        ' magenta/pink
        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .ColorScaleCriteria(3).FormatColor.Color = &HFF80FF        ' near-white pink
    End With
    On Error GoTo 0
End Sub

Private Sub ApplyThemeMono()
    On Error Resume Next
    Dim r As Range
    Set r = ThisWorkbook.Worksheets(WS_WATERFALL).Range( _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW, FIRST_COL), _
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + ROWS - 1, FIRST_COL + bins - 1))
    r.FormatConditions.Delete
    With r.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = &H101010
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 45
        .ColorScaleCriteria(2).FormatColor.Color = &H808080
        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .ColorScaleCriteria(3).FormatColor.Color = &HFFFFFF
    End With
    On Error GoTo 0
End Sub

' -- loop control -------------------------------------------------------
' -- loop control (DoEvents render loop � the cursed while-loop) -------
' Runs until the user types STOP in the stop_flag cell (SDR B12) or calls StopLoop.
Public Sub StartLoop()
    If m_running Then Exit Sub
    m_startPump = True
    m_running = True
    m_lastSeq = 0
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
            Call ReadFrame
            refreshMs = RefreshInterval()
            tNext = Timer + refreshMs / 1000#
        Else
            Sleep 10
        End If
        DoEvents
    Loop
    m_running = False
    m_startPump = False
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
    m_startPump = False
    SetVal "stop_flag", "STOP"
    SetVal "status_cell", "STOPPING"
End Sub

' Callback used by the heartbeat scheduler. Runs when Excel is idle so a
' button macro (which called StartLoop) never blocks the UI.
Public Sub Heartbeat()
    On Error Resume Next
    If m_startPump Then Call StartLoop
End Sub

Public Sub StartEngine()
    ' Launch the Python SDR engine headless (no console, one click), then
    ' arm an Idle-time start so the button macro RETURNS and never blocks COM.
    ' If the engine is already running, skip the launch (dongle is busy).
    On Error Resume Next
    Dim pid As Double
    If EngineRunning() Then
        SetVal "err_cell", ""
        SetVal "status_cell", "RUNNING"
        Call ArmStart
        Exit Sub
    End If
    Dim engDir As String, projRoot As String, q As String
    q = Chr(34)
    engDir = ThisWorkbook.Path & "\..\python"
    projRoot = ThisWorkbook.Path & "\.."
    pid = Shell("cmd /c cd /d " & q & engDir & q & " && uv run python main.py --mode sdr --fps 15 > " & q & projRoot & "\xdr_data\engine.log" & q & " 2>&1", vbHide)
    Err.Clear
    If pid = 0 Then
        SetVal "err_cell", "engine launch failed"
    Else
        SetVal "err_cell", ""
        SetVal "status_cell", "LAUNCHED"
        Call ArmStart
    End If
End Sub

' Arm the heartbeat: schedule a one-shot callback at idle (~1s out) so the
' button/COM call returns promptly. No named arg � OnTime is positional-only.
Public Sub ArmStart()
    m_startPump = True
    Application.OnTime Now + TimeValue("00:00:01"), "XDRMain.Heartbeat"
End Sub

Public Sub StopEngine()
    Call StopLoop
    Call KillEngine
    SetVal "status_cell", "STOPPED"
End Sub

' Terminate the XDR engine python process. WMI + command-line filter so we
' only kill OUR engine (main.py --mode sdr), never unrelated python.
Private Sub KillEngine()
    Dim objWMI As Object, colItems As Object, objItem As Object
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery( _
        "SELECT ProcessID FROM Win32_Process WHERE (Name LIKE 'python%') " & _
        "AND CommandLine LIKE '%main.py --mode sdr%'")
    If Not colItems Is Nothing Then
        For Each objItem In colItems
            objItem.Terminate
        Next
    End If
    Set objWMI = Nothing
End Sub

' True if the XDR engine is already running (launch guard).
Public Function EngineRunning() As Boolean
    Dim objWMI As Object, colItems As Object, objItem As Object
    EngineRunning = False
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery( _
        "SELECT ProcessID FROM Win32_Process WHERE (Name LIKE 'python%') " & _
        "AND CommandLine LIKE '%main.py --mode sdr%'")
    If Not colItems Is Nothing Then
        For Each objItem In colItems
            EngineRunning = True
            Exit For
        Next
    End If
    Set objWMI = Nothing
End Function

Public Sub StepAll()
    ' One full pump: read one frame + paint, without the loop.
    Call EnsureRing
    Call ReadFrame
    SetVal "status_cell", "STEPPED"
End Sub

Public Sub StepOnce()
    ' One synchronous read+paint for debugging: reports any error to err_cell.
    On Error GoTo EH
    SetVal "err_cell", ""
    Call EnsureRing
    Call ReadFrame
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
    ReDim arr(0 To bins - 1, 0 To ROWS - 1)
    For j = 0 To ROWS - 1
        For i = 0 To bins - 1
            arr(i, j) = (i + j * 0.5) / (bins + ROWS * 0.5)
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
        ThisWorkbook.Worksheets(WS_WATERFALL).Cells(FIRST_ROW + ROWS - 1, FIRST_COL + bins - 1))
    r.Clear
    On Error GoTo 0
End Sub

Public Sub InitDisplay()
    SetVal "status_cell", "READY"
    SetVal "fps_cell", "--"
    SetVal "paints_cell", 0
    SetVal "seq_cell", 0
    SetVal "render_ms_cell", "--"
    Call EnsureRing
End Sub


