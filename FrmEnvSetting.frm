VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmEnvSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境设置"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame5 
      Caption         =   "绘图参数"
      Height          =   615
      Left            =   1560
      TabIndex        =   25
      Top             =   2280
      Width           =   2055
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "线宽"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "物理常数"
      Height          =   2055
      Left            =   3720
      TabIndex        =   18
      Top             =   120
      Width           =   2055
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "重力加速度"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "静电力常数"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "万有引力常数"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1080
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "扫描"
      Height          =   2055
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   2055
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "真实一秒对应模拟秒数"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "每次渲染扫描次数"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "渲染间隔"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "显示"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
      Begin VB.CheckBox Check5 
         Caption         =   "坐标网格"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "坐标系"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "背景颜色"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "考虑"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox Check3 
         Caption         =   "万有引力"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "电荷吸引"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "重力"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmEnvSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Consider1 = Check1.Value
    Consider2 = Check2.Value
    Consider3 = Check3.Value
    ShowAxis = Check4.Value
    ShowGrid = Check5.Value
    RenderInterval = Val(Text1.Text)
    RenderCount = Val(Text2.Text)
    TimeRatio = Val(Text3.Text)
    BgColor = Picture1.BackColor
    ConstG = Text4.Text
    ConstK = Text5.Text
    ConstGravity = Text6.Text
    LineWidth = Text7.Text
    Saved = False
    Redraw
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Check1.Value = IIf(Consider1, 1, 0)
    Check2.Value = IIf(Consider2, 1, 0)
    Check3.Value = IIf(Consider3, 1, 0)
    Check4.Value = IIf(ShowAxis, 1, 0)
    Check5.Value = IIf(ShowGrid, 1, 0)
    Text1.Text = RenderInterval
    Text2.Text = RenderCount
    Text3.Text = TimeRatio
    Picture1.BackColor = BgColor
    Text4.Text = ConstG
    Text5.Text = ConstK
    Text6.Text = ConstGravity
    Text7.Text = LineWidth
End Sub

Private Sub Picture1_Click()
    CommonDialog1.Color = Picture1.BackColor
    CommonDialog1.ShowColor
    Picture1.BackColor = CommonDialog1.Color
End Sub
