Attribute VB_Name = "MdlSub"
Option Explicit

'API
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Const
Public Const Pi As Double = 3.14159265358979

'View
Public Alpha As Double, Beta As Double, K As Double, CenterX As Long, CenterY As Long

'Env
Public ConstG As Double, ConstK As Double, ConstGravity
Public Consider1 As Boolean, Consider2 As Boolean, Consider3 As Boolean
Public ShowAxis As Boolean, ShowGrid As Boolean, BgColor As Long, LineWidth As Integer
Public RenderInterval As Double, RenderCount As Long, TimeRatio As Double

'Tracer
Public Locus(65535) As Vector, LocusColor As Long, LocusCount As Integer
Public TracerWorking As Boolean

'HelpText
Public HelpText As String

'Sub & Function

Sub InitEnv()
    ConstG = 0.0000000000667
    ConstK = 9000000000#
    ConstGravity = 9.8
    Consider1 = True
    Consider2 = False
    Consider3 = False
    ShowAxis = True
    ShowGrid = False
    BgColor = 0
    LineWidth = 2
    RenderInterval = 0.04
    RenderCount = 100
    TimeRatio = 1
    FrmMain.Pic1.BackColor = BgColor
    Alpha = Pi / 180 * 75
    Beta = Pi / 180 * 30
    K = 0.1
    CenterX = FrmMain.Pic1.Width / 2
    CenterY = FrmMain.Pic1.Height / 2
    BallCount = 0
    LineCount = 0
    EFieldCount = 0
    MFieldCount = 0
    FileName = ""
    Saved = True
    FrmMain.Toolbar1.Buttons(5).Value = tbrUnpressed
    FrmMain.Timer1.Enabled = False
    Redraw
End Sub

Sub Kernel3D(Pos As Vector, PlotX As Double, PlotY As Double)
    With Pos
        PlotY = -(-Sin(Beta) * Cos(Alpha) * .Y - Sin(Beta) * Sin(Alpha) * .X + Cos(Beta) * .Z)
        PlotX = Sin(Alpha) * .Y - Cos(Alpha) * .X
    End With
End Sub

Sub Line3D(Vec1 As Vector, Vec2 As Vector, Color As Long)
    Dim ShowX As Double, ShowY As Double
    FrmMain.Pic1.ForeColor = Color
    Kernel3D Vec1, ShowX, ShowY
    FrmMain.Pic1.CurrentX = ShowX / K + CenterX
    FrmMain.Pic1.CurrentY = ShowY / K + CenterY
    Kernel3D Vec2, ShowX, ShowY
    FrmMain.Pic1.Line -(ShowX / K + CenterX, ShowY / K + CenterY)
End Sub

Sub Update()
    Dim i As Integer, T As Double, Count As Long
    T = RenderInterval * TimeRatio / RenderCount
    For Count = 1 To RenderCount
        If BallCount > 0 Then
            '计算匀变预测位移
            For i = 1 To BallCount
                With Balls(i)
                    .PreP = .P
                    .PreV = .V
                End With
            Next
            CalcF
            For i = 1 To BallCount
                With Balls(i)
                    If Not .Fixed Then
                        .PreP = VecAdd(.P, VecNum(VecAdd(VecNum(.V, T), VecNum(.A, T * T / 2)), 0.5))
                        .PreV = VecAdd(.V, VecNum(.A, T))
                    End If
                End With
            Next
            CalcF
            '运动
            For i = 1 To BallCount
                With Balls(i)
                    If Not .Fixed Then
                        .P = VecAdd(VecAdd(.P, VecNum(.V, T)), VecNum(.A, T * T / 2))
                        .V = VecAdd(.V, VecNum(.A, T))
                    End If
                End With
            Next
        End If
    Next
End Sub

Sub CalcF()
    Dim i As Integer, j As Integer
    Dim Dis As Double
    Dim F As Vector, Dif As Vector, AF As Double
    For i = 1 To BallCount
        Balls(i).A = Vec0
    Next
    '计算受力
    For i = 1 To BallCount
        With Balls(i)
            .A = VecAdd(.A, VecNum(.F, 1 / .M)) '外加恒力
            If Consider1 Then '重力
                .A.Z = .A.Z - ConstGravity
            End If
            If Consider2 And .Q <> 0 Then '电荷吸引
                For j = 1 To BallCount
                    If Balls(j).Q <> 0 And i <> j Then
                        Dis = VecDis(.PreP, Balls(j).PreP)
                        If Dis <> 0 Then
                            AF = ConstK * .Q * Balls(j).Q / Dis / Dis
                            .A = VecAdd(.A, VecNum(VecDir(VecSub(.PreP, Balls(j).PreP), AF), .M))
                        End If
                    End If
                Next
            End If
            If Consider3 Then '万有引力
                For j = 1 To BallCount
                    If i <> j Then
                        Dis = VecDis(.PreP, Balls(j).PreP)
                        If Dis <> 0 Then
                            AF = ConstG * Balls(j).M / Dis / Dis
                            .A = VecAdd(.A, VecDir(VecSub(Balls(j).PreP, .PreP), AF))
                        End If
                    End If
                Next
            End If
            If EFieldCount > 0 And .Q <> 0 Then '匀强电场
                For j = 1 To EFieldCount
                    If VecInRange(.PreP, EFields(i).RangeStart, EFields(i).RangeEnd) Then
                        .A = VecAdd(.A, VecNum(EFields(i).E, .Q / .M))
                    End If
                Next
            End If
            If MFieldCount > 0 And .Q <> 0 Then '匀强磁场
                For j = 1 To MFieldCount
                    If VecInRange(.PreP, MFields(i).RangeStart, MFields(i).RangeEnd) Then
                        .A = VecAdd(.A, VecNum(VecMul(.PreV, MFields(i).B), .Q / .M))
                    End If
                Next
            End If
        End With
    Next
    If LineCount > 0 Then '弹性绳
        For i = 1 To LineCount
            With Lines(i)
                Dis = VecDis(Balls(.Ball1).PreP, Balls(.Ball2).PreP)
                If .Style = 0 Or (.Style = 1 And Dis > .Length) Or (.Style = 2 And Dis > .Length) Then
                    Dif = VecSub(Balls(.Ball1).PreP, Balls(.Ball2).PreP)
                    F = VecNum(VecSub(VecDir(Dif, .Length), Dif), .K) 'F=-kx
                    Balls(.Ball1).A = VecAdd(Balls(.Ball1).A, VecNum(F, 1 / Balls(.Ball1).M))
                    Balls(.Ball2).A = VecSub(Balls(.Ball2).A, VecNum(F, 1 / Balls(.Ball2).M))
                End If
            End With
        Next
    End If
End Sub

Sub Redraw()
    Dim i As Integer
    Dim ShowX As Double, ShowY As Double
    Dim AxisNum As String
    FrmMain.Pic1.Cls
    FrmMain.Pic1.BackColor = BgColor
    FrmMain.Pic1.DrawWidth = LineWidth
    '绘制坐标系
    If ShowGrid Then
        FrmMain.Pic1.ForeColor = &H777777
        For i = -10 To 10
                Kernel3D Vec(50 * i, -500, 0), ShowX, ShowY
                FrmMain.Pic1.CurrentX = ShowX + CenterX
                FrmMain.Pic1.CurrentY = ShowY + CenterY
                Kernel3D Vec(50 * i, 500, 0), ShowX, ShowY
                FrmMain.Pic1.Line -(ShowX + CenterX, ShowY + CenterY)
                Kernel3D Vec(-500, 50 * i, 0), ShowX, ShowY
                FrmMain.Pic1.CurrentX = ShowX + CenterX
                FrmMain.Pic1.CurrentY = ShowY + CenterY
                Kernel3D Vec(500, 50 * i, 0), ShowX, ShowY
                FrmMain.Pic1.Line -(ShowX + CenterX, ShowY + CenterY)
        Next
    End If
    If ShowAxis Then
        AxisNum = Format(K * 500, "Scientific")
        Kernel3D Vec(500, 0, 0), ShowX, ShowY
        FrmMain.Pic1.ForeColor = &HFF&
        FrmMain.Pic1.Line (CenterX, CenterY)-(ShowX + CenterX, ShowY + CenterY)
        FrmMain.Pic1.Print " X " & AxisNum
        Kernel3D Vec(0, 500, 0), ShowX, ShowY
        FrmMain.Pic1.ForeColor = &HFF00&
        FrmMain.Pic1.Line (CenterX, CenterY)-(ShowX + CenterX, ShowY + CenterY)
        FrmMain.Pic1.Print " Y " & AxisNum
        Kernel3D Vec(0, 0, 500), ShowX, ShowY
        FrmMain.Pic1.ForeColor = &HFF0000
        FrmMain.Pic1.Line (CenterX, CenterY)-(ShowX + CenterX, ShowY + CenterY)
        FrmMain.Pic1.Print " Z " & AxisNum
    End If
    '绘制小球
    If BallCount > 0 Then
        For i = 1 To BallCount
            With Balls(i)
                Kernel3D .P, ShowX, ShowY
                ShowX = ShowX / K + CenterX
                ShowY = ShowY / K + CenterY
                FrmMain.Pic1.Circle (ShowX, ShowY), .R / K, .Color
            End With
        Next
    End If
    '绘制弹性绳
    If LineCount > 0 Then
        For i = 1 To LineCount
            With Lines(i)
                Line3D Balls(.Ball1).P, Balls(.Ball2).P, .Color
            End With
        Next
    End If
    '绘制电场
    If EFieldCount > 0 Then
        For i = 1 To EFieldCount
            With EFields(i)
                DrawBox .RangeStart, .RangeEnd, .Color
            End With
        Next
    End If
    '绘制磁场
    If MFieldCount > 0 Then
        For i = 1 To MFieldCount
            With MFields(i)
                DrawBox .RangeStart, .RangeEnd, .Color
            End With
        Next
    End If
    '轨迹
    If TracerWorking = True Then
        If LocusCount > 0 Then
            For i = 1 To LocusCount
                Line3D Locus(i - 1), Locus(i), LocusColor
            Next
        End If
    End If
    '角度滑杆
    FrmMain.Slider1.Value = Alpha / Pi * 180
    FrmMain.Slider2.Value = Beta / Pi * 180
    '状态栏
    FrmMain.StatusBar1.Panels(1).Text = IIf(Consider1, "重力 ", "") & IIf(Consider2, "电荷吸引 ", "") & IIf(Consider3, "万有引力", "")
    FrmMain.StatusBar1.Panels(2).Text = "Alpha=" & Int(Alpha / Pi * 180) & " Beta=" & Int(Beta / Pi * 180) & " K=" & K & " 平移=(" & CenterX & "," & CenterY & ")"
    FrmMain.StatusBar1.Panels(3).Text = IIf(Saved, "未修改,", "已修改,") & IIf(FileName = "", "未标题文件", FileName)
    '显示
    With FrmMain
        BitBlt .Pic2.hDC, 0, 0, .Pic1.ScaleWidth, .Pic1.ScaleHeight, .Pic1.hDC, 0, 0, &HCC0020
        .Pic2.Refresh
    End With
End Sub

Sub Swap(Str1 As String, Str2 As String)
    Dim Temp As String
    Temp = Str2
    Str2 = Str1
    Str1 = Temp
End Sub

Sub DrawBox(RangeStart As Vector, RangeEnd As Vector, Color As Long)
    Line3D Vec(RangeStart.X, RangeStart.Y, RangeStart.Z), Vec(RangeEnd.X, RangeStart.Y, RangeStart.Z), Color
    Line3D Vec(RangeStart.X, RangeStart.Y, RangeEnd.Z), Vec(RangeEnd.X, RangeStart.Y, RangeEnd.Z), Color
    Line3D Vec(RangeStart.X, RangeEnd.Y, RangeStart.Z), Vec(RangeEnd.X, RangeEnd.Y, RangeStart.Z), Color
    Line3D Vec(RangeStart.X, RangeEnd.Y, RangeEnd.Z), Vec(RangeEnd.X, RangeEnd.Y, RangeEnd.Z), Color
    Line3D Vec(RangeStart.X, RangeStart.Y, RangeStart.Z), Vec(RangeStart.X, RangeEnd.Y, RangeStart.Z), Color
    Line3D Vec(RangeStart.X, RangeStart.Y, RangeEnd.Z), Vec(RangeStart.X, RangeEnd.Y, RangeEnd.Z), Color
    Line3D Vec(RangeEnd.X, RangeStart.Y, RangeStart.Z), Vec(RangeEnd.X, RangeEnd.Y, RangeStart.Z), Color
    Line3D Vec(RangeEnd.X, RangeStart.Y, RangeEnd.Z), Vec(RangeEnd.X, RangeEnd.Y, RangeEnd.Z), Color
    Line3D Vec(RangeStart.X, RangeStart.Y, RangeStart.Z), Vec(RangeStart.X, RangeStart.Y, RangeEnd.Z), Color
    Line3D Vec(RangeStart.X, RangeEnd.Y, RangeStart.Z), Vec(RangeStart.X, RangeEnd.Y, RangeEnd.Z), Color
    Line3D Vec(RangeEnd.X, RangeStart.Y, RangeStart.Z), Vec(RangeEnd.X, RangeStart.Y, RangeEnd.Z), Color
    Line3D Vec(RangeEnd.X, RangeEnd.Y, RangeStart.Z), Vec(RangeEnd.X, RangeEnd.Y, RangeEnd.Z), Color
End Sub

Sub ViewChange(DegAlpha As Double, DegBeta As Double)
    Alpha = DegAlpha * Pi / 180
    Beta = DegBeta * Pi / 180
    CenterX = FrmMain.Pic1.Width / 2
    CenterY = FrmMain.Pic1.Height / 2
    Saved = False
    Redraw
End Sub
