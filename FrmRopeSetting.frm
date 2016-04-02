VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRopeSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "弹性绳属性"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command4 
      Caption         =   "预览"
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "其它属性"
      Height          =   1095
      Left            =   3000
      TabIndex        =   11
      Top             =   120
      Width           =   2055
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "FrmRopeSetting.frx":0000
         Left            =   240
         List            =   "FrmRopeSetting.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "作用条件"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本属性"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "FrmRopeSetting.frx":0035
         Left            =   840
         List            =   "FrmRopeSetting.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "FrmRopeSetting.frx":006A
         Left            =   840
         List            =   "FrmRopeSetting.frx":0077
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "获取"
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "劲度系数"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "小球1"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "小球2"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "原长"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "颜色"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   360
      End
   End
End
Attribute VB_Name = "FrmRopeSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error GoTo Errh
    If Val(Text3.Text) < 0 Then
        MsgBox "原长必须大于等于0！", vbExclamation
        Exit Sub
    ElseIf Val(Text4.Text) <= 0 Then
        MsgBox "劲度系数必须为正数！", vbExclamation
        Exit Sub
    End If
    'check
    If Int(Val(Combo2.Text)) > BallCount Or Int(Val(Combo3.Text)) < 1 Then
        MsgBox "小球编号应介于1至" & BallCount, vbExclamation
        Exit Sub
    End If
    If Int(Val(Combo2.Text)) = Int(Val(Combo3.Text)) Then
        MsgBox "小球1和小球2不能相同！" & BallCount, vbExclamation
        Exit Sub
    End If
    'legal
    With LineReturn
        .Ball1 = Combo2.Text
        .Ball2 = Combo3.Text
        .Length = Text3.Text
        .K = Text4.Text
        .Color = Picture1.BackColor
        .Style = Combo1.ListIndex
    End With
    DialogOK = True
    Saved = False
    Unload Me
    Exit Sub
Errh:
    MsgBox "输入错误！", vbExclamation
End Sub

Private Sub Command2_Click()
    DialogOK = False
    Unload Me
End Sub

Private Sub Command3_Click()
    If Int(Val(Combo2.Text)) > BallCount Or Int(Val(Combo3.Text)) < 1 Then
        MsgBox "小球编号应介于1至" & BallCount, vbExclamation
        Exit Sub
    End If
    If Int(Val(Combo2.Text)) = Int(Val(Combo3.Text)) Then
        MsgBox "小球1和小球2不能相同！" & BallCount, vbExclamation
        Exit Sub
    End If
    Text3.Text = VecDis(Balls(Int(Val(Combo2.Text))).P, Balls(Int(Val(Combo3.Text))).P)
End Sub

Private Sub Command4_Click()
    If Int(Val(Combo2.Text)) = Int(Val(Combo3.Text)) Then
        MsgBox "小球1和小球2不能相同！" & BallCount, vbExclamation
        Exit Sub
    End If
    Redraw
    Line3D Balls(Int(Val(Combo2.Text))).P, Balls(Int(Val(Combo3.Text))).P, Picture1.BackColor
    With FrmMain
        BitBlt .Pic2.hDC, 0, 0, .Pic1.ScaleWidth, .Pic1.ScaleHeight, .Pic1.hDC, 0, 0, &HCC0020
        .Pic2.Refresh
    End With
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Combo2.Clear
    Combo3.Clear
    For i = 1 To BallCount
        Combo2.AddItem i
        Combo3.AddItem i
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then DialogOK = False
End Sub

Private Sub Picture1_Click()
    CommonDialog1.Color = Picture1.BackColor
    CommonDialog1.ShowColor
    Picture1.BackColor = CommonDialog1.Color
End Sub


