Attribute VB_Name = "MdlFile"
Option Explicit

'File
Public FileName As String
Public Saved As Boolean

Sub SaveFile()
    On Error GoTo Err
    If FileName = "" Then
        FrmMain.CommonDialog1.ShowSave
        FileName = FrmMain.CommonDialog1.FileName
        SaveData
    Else
        SaveData
    End If
    Redraw
Err:
    Exit Sub
End Sub

Sub SaveData()
    Dim i As Integer
    If Dir(FileName) <> "" Then Kill FileName
    Open FileName For Binary As #1
        Put #1, , Alpha
        Put #1, , Beta
        Put #1, , K
        Put #1, , CenterX
        Put #1, , CenterY
        Put #1, , ConstG
        Put #1, , ConstK
        Put #1, , ConstGravity
        Put #1, , Consider1
        Put #1, , Consider2
        Put #1, , Consider3
        Put #1, , ShowAxis
        Put #1, , ShowGrid
        Put #1, , BgColor
        Put #1, , LineWidth
        Put #1, , RenderInterval
        Put #1, , RenderCount
        Put #1, , TimeRatio
        Put #1, , BallCount
        Put #1, , LineCount
        Put #1, , EFieldCount
        Put #1, , MFieldCount
        If BallCount > 0 Then
            For i = 1 To BallCount
                Put #1, , Balls(i)
            Next
        End If
        If LineCount > 0 Then
            For i = 1 To LineCount
                Put #1, , Lines(i)
            Next
        End If
        If EFieldCount > 0 Then
            For i = 1 To EFieldCount
                Put #1, , EFields(i)
            Next
        End If
        If MFieldCount > 0 Then
            For i = 1 To MFieldCount
                Put #1, , MFields(i)
            Next
        End If
    Close
    Saved = True
End Sub

Sub OpenFile()
    On Error GoTo Err
    FrmMain.CommonDialog1.ShowOpen
    FileName = FrmMain.CommonDialog1.FileName
    LoadFile
Err:
    Exit Sub
End Sub

Sub LoadFile()
    Dim i As Integer
    Open FileName For Binary As #1
    Get #1, , Alpha
    Get #1, , Beta
    Get #1, , K
    Get #1, , CenterX
    Get #1, , CenterY
    Get #1, , ConstG
    Get #1, , ConstK
    Get #1, , ConstGravity
    Get #1, , Consider1
    Get #1, , Consider2
    Get #1, , Consider3
    Get #1, , ShowAxis
    Get #1, , ShowGrid
    Get #1, , BgColor
    Get #1, , LineWidth
    Get #1, , RenderInterval
    Get #1, , RenderCount
    Get #1, , TimeRatio
    Get #1, , BallCount
    Get #1, , LineCount
    Get #1, , EFieldCount
    Get #1, , MFieldCount
    If BallCount > 0 Then
        For i = 1 To BallCount
            Get #1, , Balls(i)
        Next
    End If
    If LineCount > 0 Then
        For i = 1 To LineCount
            Get #1, , Lines(i)
        Next
    End If
    If EFieldCount > 0 Then
        For i = 1 To EFieldCount
            Get #1, , EFields(i)
        Next
    End If
    If MFieldCount > 0 Then
        For i = 1 To MFieldCount
            Get #1, , MFields(i)
        Next
    End If
    Close
    Saved = True
    Redraw
End Sub
