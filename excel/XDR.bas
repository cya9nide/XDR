Attribute VB_Name = "XDRMain"

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

