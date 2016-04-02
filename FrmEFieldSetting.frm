VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmEFieldSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "匀强电场属性"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      ScaleHeight     =   225
      ScaleWidth      =   1065
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "范围"
      Height          =   1335
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   2280
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   2280
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   720
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   720
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   1920
         TabIndex        =   23
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   1920
         TabIndex        =   22
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Z坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Y坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "电场强度"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Z坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Y坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X坐标"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "颜色"
      Height          =   180
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   360
   End
End
Attribute VB_Name = "FrmEFieldSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error GoTo Errh
    If Val(Text4.Text) >= Val(Text7.Text) Or Val(Text5.Text) >= Val(Text8.Text) Or Val(Text6.Text) >= Val(Text9.Text) Then
        MsgBox "范围起始值必须小于终止值！", vbExclamation
        Exit Sub
    End If
    With EFieldReturn
        .E.X = Text1.Text
        .E.Y = Text2.Text
        .E.Z = Text3.Text
        .RangeStart.X = Text4.Text
        .RangeStart.Y = Text5.Text
        .RangeStart.Z = Text6.Text
        .RangeEnd.X = Text7.Text
        .RangeEnd.Y = Text8.Text
        .RangeEnd.Z = Text9.Text
        .Color = Picture1.BackColor
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
