VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmBallSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "小球属性"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   8040
      TabIndex        =   32
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   7080
      TabIndex        =   31
      Top             =   2040
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "固定"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   "外加恒力"
      Height          =   1815
      Left            =   6720
      TabIndex        =   23
      Top             =   120
      Width           =   2175
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   720
         TabIndex        =   26
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   720
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   720
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "X坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Y坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Z坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "速度"
      Height          =   1815
      Left            =   4440
      TabIndex        =   16
      Top             =   120
      Width           =   2175
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   720
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   720
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   720
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "X坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Y坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Z坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "位移"
      Height          =   1815
      Left            =   2160
      TabIndex        =   9
      Top             =   120
      Width           =   2175
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   720
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Z坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Y坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本属性"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "半径"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "颜色"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "电荷量"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "质量"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   360
      End
   End
End
Attribute VB_Name = "FrmBallSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error GoTo Errh
    If Val(Text1.Text) <= 0 Then
        MsgBox "质量必须为正数！", vbExclamation
        Exit Sub
    ElseIf Val(Text3.Text) <= 0 Then
        MsgBox "半径必须为正数！", vbExclamation
        Exit Sub
    End If
    With BallReturn
        .M = Text1.Text
        .Q = Text2.Text
        .R = Text3.Text
        .P.X = Text4.Text
        .P.Y = Text5.Text
        .P.Z = Text6.Text
        .V.X = Text7.Text
        .V.Y = Text8.Text
        .V.Z = Text9.Text
        .F.X = Text10.Text
        .F.Y = Text11.Text
        .F.Z = Text12.Text
        .Color = Picture1.BackColor
        .Fixed = Check1.Value
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then DialogOK = False
End Sub

Private Sub Picture1_Click()
    CommonDialog1.Color = Picture1.BackColor
    CommonDialog1.ShowColor
    Picture1.BackColor = CommonDialog1.Color
End Sub
