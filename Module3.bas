Attribute VB_Name = "MdlObj"
Option Explicit

'Obj
Public Balls(100) As ObjBall, BallCount As Integer
Public Lines(100) As ObjLine, LineCount As Integer
Public EFields(100) As ObjEField, EFieldCount As Integer
Public MFields(100) As ObjMField, MFieldCount As Integer

'Default
Public Ball0 As ObjBall, Line0 As ObjLine, EField0 As ObjEField, MField0 As ObjMField

'Dialog
Public DialogOK As Boolean
Public BallReturn As ObjBall, LineReturn As ObjLine, EFieldReturn As ObjEField, MFieldReturn As ObjMField

'ObjTypes
Public Type ObjBall
    M As Double
    Q As Double
    R As Double
    P As Vector
    V As Vector
    A As Vector
    F As Vector
    Color As Long
    PreP As Vector
    PreV As Vector
    Fixed As Boolean
End Type

Public Type ObjLine
    Ball1 As Integer
    Ball2 As Integer
    Length As Double
    K As Double
    Color As Long
    Style As Integer
End Type

Public Type ObjEField
    RangeStart As Vector
    RangeEnd As Vector
    E As Vector
    Color As Long
End Type

Public Type ObjMField
    RangeStart As Vector
    RangeEnd As Vector
    B As Vector
    Color As Long
End Type

Sub AddBall()
    If BallCount = 100 Then
        MsgBox "小球已达最大数量！", vbCritical
        Exit Sub
    End If
    BallReturn = Ball0
    SetBall
    If DialogOK Then
        BallCount = BallCount + 1
        Balls(BallCount) = BallReturn
        Redraw
    End If
End Sub

Sub AddLine()
    If LineCount = 100 Then
        MsgBox "弹性绳已达最大数量！", vbCritical
        Exit Sub
    End If
    If BallCount < 2 Then
        MsgBox "请先添加至少2个小球！", vbCritical
        Exit Sub
    End If
    LineReturn = Line0
    SetLine
    If DialogOK Then
        LineCount = LineCount + 1
        Lines(LineCount) = LineReturn
        Redraw
    End If
End Sub

Sub AddEField()
    If EFieldCount = 100 Then
        MsgBox "匀强电场已达最大数量！", vbCritical
        Exit Sub
    End If
    EFieldReturn = EField0
    SetEField
    If DialogOK Then
        EFieldCount = EFieldCount + 1
        EFields(EFieldCount) = EFieldReturn
        Redraw
    End If
End Sub

Sub AddMField()
    If MFieldCount = 100 Then
        MsgBox "匀强磁场已达最大数量！", vbCritical
        Exit Sub
    End If
    MFieldReturn = MField0
    SetMField
    If DialogOK Then
        MFieldCount = MFieldCount + 1
        MFields(MFieldCount) = MFieldReturn
        Redraw
    End If
End Sub

Sub DefineDefault()
    With Ball0
        .M = 1
        .Q = 0
        .R = 1
        .P = Vec0
        .V = Vec0
        .F = Vec0
        .Color = &HFF00&
        .Fixed = False
        .A = Vec0
        .PreP = Vec0
        .PreV = Vec0
    End With
    With Line0
        .Ball1 = 1
        .Ball2 = 2
        .Length = 10
        .K = 10
        .Style = 1
        .Color = &HFFFF&
    End With
    With EField0
        .E = Vec0
        .RangeStart = Vec(-30, -30, -30)
        .RangeEnd = Vec(30, 30, 30)
        .Color = &HFFFF00
    End With
    With MField0
        .B = Vec0
        .RangeStart = Vec(-30, -30, -30)
        .RangeEnd = Vec(30, 30, 30)
        .Color = &HFF00FF
    End With
End Sub

Sub SetBall()
    With FrmBallSetting
        .Text1.Text = BallReturn.M
        .Text2.Text = BallReturn.Q
        .Text3.Text = BallReturn.R
        .Text4.Text = BallReturn.P.X
        .Text5.Text = BallReturn.P.Y
        .Text6.Text = BallReturn.P.Z
        .Text7.Text = BallReturn.V.X
        .Text8.Text = BallReturn.V.Y
        .Text9.Text = BallReturn.V.Z
        .Text10.Text = BallReturn.F.X
        .Text11.Text = BallReturn.F.Y
        .Text12.Text = BallReturn.F.Z
        .Picture1.BackColor = BallReturn.Color
        .Check1.Value = IIf(BallReturn.Fixed, 1, 0)
        .Show 1
    End With
End Sub

Sub SetLine()
    With FrmRopeSetting
        .Combo2.ListIndex = LineReturn.Ball1 - 1
        .Combo3.ListIndex = LineReturn.Ball2 - 1
        .Text3.Text = LineReturn.Length
        .Text4.Text = LineReturn.K
        .Picture1.BackColor = LineReturn.Color
        .Combo1.ListIndex = LineReturn.Style
        .Show 1
    End With
End Sub

Sub SetEField()
    With FrmEFieldSetting
        .Text1.Text = EFieldReturn.E.X
        .Text2.Text = EFieldReturn.E.Y
        .Text3.Text = EFieldReturn.E.Z
        .Text4.Text = EFieldReturn.RangeStart.X
        .Text5.Text = EFieldReturn.RangeStart.Y
        .Text6.Text = EFieldReturn.RangeStart.Z
        .Text7.Text = EFieldReturn.RangeEnd.X
        .Text8.Text = EFieldReturn.RangeEnd.Y
        .Text9.Text = EFieldReturn.RangeEnd.Z
        .Picture1.BackColor = EFieldReturn.Color
        .Show 1
    End With
End Sub

Sub SetMField()
    With FrmMFieldSetting
        .Text1.Text = MFieldReturn.B.X
        .Text2.Text = MFieldReturn.B.Y
        .Text3.Text = MFieldReturn.B.Z
        .Text4.Text = MFieldReturn.RangeStart.X
        .Text5.Text = MFieldReturn.RangeStart.Y
        .Text6.Text = MFieldReturn.RangeStart.Z
        .Text7.Text = MFieldReturn.RangeEnd.X
        .Text8.Text = MFieldReturn.RangeEnd.Y
        .Text9.Text = MFieldReturn.RangeEnd.Z
        .Picture1.BackColor = MFieldReturn.Color
        .Show 1
    End With
End Sub
