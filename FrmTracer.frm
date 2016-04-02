VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTracer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "轨迹追踪器"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3900
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "轨迹"
      Height          =   1455
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      Width           =   1815
      Begin VB.CommandButton Command3 
         Caption         =   "清除轨迹"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "已采集："
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "颜色"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1320
      Top             =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开始追踪"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "追踪对象"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "刷新列表"
         Height          =   420
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Height          =   2040
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Selected As Integer

Private Sub Command1_Click()
    RefreshList
End Sub

Private Sub Command2_Click()
    If List1.ListIndex >= 0 Then
        If Timer1.Enabled = False Then
            If LocusCount > 0 Then
                If MsgBox("之前的轨迹将会丢失，继续？", vbYesNo) = vbNo Then Exit Sub
            End If
            LocusCount = 0
            Selected = List1.ListIndex + 1
            Locus(0) = Balls(Selected).P
            Timer1.Interval = RenderInterval * 1000
            Timer1.Enabled = True
            Frame1.Enabled = False
            Command2.Caption = "停止追踪"
        Else
            Timer1.Enabled = False
            Frame1.Enabled = True
            Command2.Caption = "开始追踪"
        End If
        Saved = False
    Else
        MsgBox "请先选择一个追踪对象！", vbExclamation
    End If
End Sub

Private Sub Command3_Click()
    If LocusCount > 0 Then
        If MsgBox("之前的轨迹将会丢失，继续？", vbYesNo) = vbNo Then Exit Sub
    End If
    LocusCount = 0
    Redraw
End Sub

Private Sub Form_Load()
    RefreshList
    LocusCount = 0
    TracerWorking = True
    With FrmMain
        .MenuFile.Enabled = False
        .MenuEdit.Enabled = False
        .MenuAdd.Enabled = False
        .MenuTool.Enabled = False
        .MenuRun.Enabled = False
        .Toolbar1.Buttons(1).Enabled = False
        .Toolbar1.Buttons(2).Enabled = False
        .Toolbar1.Buttons(3).Enabled = False
        .Toolbar1.Buttons(5).Enabled = False
        .Toolbar1.Buttons(6).Enabled = False
    End With
    LocusColor = Picture1.BackColor
    ShowGot
End Sub

Sub RefreshList()
    Dim i As Integer
    List1.Clear
    If BallCount > 0 Then
        For i = 1 To BallCount
            With Balls(i)
                List1.AddItem "[" & i & "] M=" & .M & " (" & .P.X & "," & .P.Y & "," & .P.Z & ")"
            End With
        Next
        List1.ToolTipText = ""
    Else
        MsgBox "请先添加小球！", vbExclamation
        Frame1.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Timer1.Enabled = True Then
        MsgBox "请先停止追踪！", vbExclamation
        Cancel = True
        Exit Sub
    ElseIf LocusCount > 0 Then
        If MsgBox("之前的轨迹将会丢失，继续？", vbYesNo) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    With FrmMain
        .MenuFile.Enabled = True
        .MenuEdit.Enabled = True
        .MenuAdd.Enabled = True
        .MenuTool.Enabled = True
        .MenuRun.Enabled = True
        .Toolbar1.Buttons(1).Enabled = True
        .Toolbar1.Buttons(2).Enabled = True
        .Toolbar1.Buttons(3).Enabled = True
        .Toolbar1.Buttons(5).Enabled = True
        .Toolbar1.Buttons(6).Enabled = True
    End With
    TracerWorking = False
    Redraw
End Sub

Private Sub List1_Click()
    List1.ToolTipText = List1.List(List1.ListIndex)
End Sub

Private Sub Picture1_Click()
    CommonDialog1.Color = Picture1.BackColor
    CommonDialog1.ShowColor
    Picture1.BackColor = CommonDialog1.Color
    LocusColor = Picture1.BackColor
    Redraw
End Sub

Private Sub Timer1_Timer()
    Update
    LocusCount = LocusCount + 1
    Locus(LocusCount) = Balls(Selected).P
    Redraw
    ShowGot
    If LocusCount = 65535 Then
        MsgBox "轨迹存储已满！", vbExclamation
        Command2_Click
    End If
End Sub

Sub ShowGot()
    Label1.Caption = "已采集=" & LocusCount
End Sub
