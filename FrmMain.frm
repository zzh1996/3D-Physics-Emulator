VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "三维物理"
   ClientHeight    =   6750
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9540
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   840
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "三维物理文件(*.3dp)|*.3dp|所有文件(*.*)|*.*"
      FilterIndex     =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4200
      Top             =   3120
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   360
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   581
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   8775
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "新建"
            Object.ToolTipText     =   "新建"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "打开"
            Object.ToolTipText     =   "打开"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "保存"
            Object.ToolTipText     =   "保存"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "运行"
            Object.ToolTipText     =   "运行"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "属性"
            Object.ToolTipText     =   "属性"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "帮助"
            Object.ToolTipText     =   "帮助"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Beta:"
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Alpha:"
         Top             =   0
         Width           =   615
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   1
         Max             =   360
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   1
         Max             =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6375
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "已考虑作用力"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "视角信息"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "文件名"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3016
            Text            =   "负一的平方根 2013"
            TextSave        =   "负一的平方根 2013"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   8520
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":08CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":09DC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0AEE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0C00
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0D12
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0E24
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0F36
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu MenuFileNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu MenuFileOpen 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuFileSaveAs 
         Caption         =   "另存为(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu MenuFileExit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu MenuEditEnvSetting 
         Caption         =   "环境设置(&S)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuEditProperty 
         Caption         =   "属性编辑器(&P)"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu MenuView 
      Caption         =   "视图(&V)"
      Begin VB.Menu MenuViewReset 
         Caption         =   "视图复位(&R)"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu MenuView1 
         Caption         =   "正视图(&1)"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu MenuView2 
         Caption         =   "俯视图(&2)"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu MenuView3 
         Caption         =   "侧视图(&3)"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu MenuAdd 
      Caption         =   "添加(&A)"
      Begin VB.Menu MenuAddBall 
         Caption         =   "小球(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu MenuAddLine 
         Caption         =   "弹性绳(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu MenuAddE 
         Caption         =   "匀强电场(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu MenuAddB 
         Caption         =   "匀强磁场(&M)"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu MenuRun 
      Caption         =   "运行(&R)"
      Begin VB.Menu MenuRunStart 
         Caption         =   "开始/停止(&S)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MenuTool 
      Caption         =   "工具(&T)"
      Begin VB.Menu MenuToolTracer 
         Caption         =   "轨迹追踪器(&T)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuToolCalc 
         Caption         =   "计算器(&C)"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MenuHelpHelp 
         Caption         =   "帮助(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "关于(&A)"
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DragStartX As Long, DragStartY As Long, ZoomStartY As Long, ZoomStartK As Double
Dim Pressed As Boolean

Private Sub Form_Load()
    'Init
    Vec0.X = 0
    Vec0.Y = 0
    Vec0.Z = 0
    CommonDialog1.InitDir = App.Path
    Pressed = False
    TracerWorking = False
    DefineDefault
    HelpText = Mid(LoadResData(101, "CUSTOM"), 2)
    'Icons
'    FrmBallSetting.Icon = Me.Icon
'    FrmEFieldSetting.Icon = Me.Icon
'    FrmEnvSetting.Icon = Me.Icon
'    FrmHelp.Icon = Me.Icon
'    FrmMFieldSetting.Icon = Me.Icon
'    FrmPropertyEditor.Icon = Me.Icon
'    FrmRopeSetting.Icon = Me.Icon
'    FrmTracer.Icon = Me.Icon
    'New file
    InitEnv
    If Command <> "" Then
        If Left(Command, 1) = """" Then
            FileName = Mid(Command, 2, Len(Command) - 2)
        Else
            FileName = Command
        End If
        LoadFile
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Pic1.Left = 0
        Pic1.Top = Toolbar1.Height
        Pic1.Width = Me.ScaleWidth
        Pic1.Height = Me.ScaleHeight - Toolbar1.Height - StatusBar1.Height
        CenterX = FrmMain.Pic1.Width / 2
        CenterY = FrmMain.Pic1.Height / 2
        Redraw
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Timer1.Enabled = True Then
        MsgBox "请先停止运行！", vbExclamation
        Cancel = True
    ElseIf TracerWorking Then
        MsgBox "请先关闭轨迹追踪器！", vbExclamation
        Cancel = True
    ElseIf Not Saved Then
        Select Case MsgBox("文件未保存，是否保存？", vbYesNoCancel)
        Case vbYes
            SaveFile
        Case vbCancel
            Cancel = True
        End Select
    End If
End Sub

Private Sub MenuAddB_Click()
    AddMField
End Sub

Private Sub MenuAddBall_Click()
    AddBall
End Sub

Private Sub MenuAddE_Click()
    AddEField
End Sub

Private Sub MenuAddLine_Click()
    AddLine
End Sub

Private Sub MenuEditEnvSetting_Click()
    FrmEnvSetting.Show 1
End Sub

Private Sub MenuEditProperty_Click()
    FrmPropertyEditor.Show 1
End Sub

Private Sub MenuFileExit_Click()
    Unload Me
End Sub

Private Sub MenuFileNew_Click()
    If Not Saved Then
        Select Case MsgBox("文件未保存，是否保存？", vbYesNoCancel)
        Case vbYes
            SaveFile
            InitEnv
        Case vbNo
            InitEnv
        Case vbCancel
            Exit Sub
        End Select
    Else
        InitEnv
    End If
End Sub

Private Sub MenuFileOpen_Click()
    If Not Saved Then
        Select Case MsgBox("文件未保存，是否保存？", vbYesNoCancel)
        Case vbYes
            SaveFile
            OpenFile
            Toolbar1.Buttons(5).Value = tbrUnpressed
            Timer1.Enabled = False
        Case vbNo
            OpenFile
            Toolbar1.Buttons(5).Value = tbrUnpressed
            Timer1.Enabled = False
        Case vbCancel
            Exit Sub
        End Select
    Else
        OpenFile
        Toolbar1.Buttons(5).Value = tbrUnpressed
        Timer1.Enabled = False
    End If
End Sub

Private Sub MenuFileSave_Click()
    SaveFile
End Sub

Private Sub MenuFileSaveAs_Click()
    On Error GoTo Err
    CommonDialog1.ShowSave
    FileName = CommonDialog1.FileName
    SaveData
    Redraw
Err:
    Exit Sub
End Sub

Private Sub MenuHelpAbout_Click()
    MsgBox "三维物理V1.0 负一的平方根制作 2013年4月" & vbCrLf & "E-mail:fydpfg1996@163.com QQ:903806024"
End Sub

Private Sub MenuHelpHelp_Click()
    FrmHelp.Show 1
End Sub

Private Sub MenuRunStart_Click()
    If Timer1.Enabled = False Then
        MenuFile.Enabled = False
        MenuEdit.Enabled = False
        MenuAdd.Enabled = False
        MenuTool.Enabled = False
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(5).Value = tbrPressed
        Timer1.Interval = RenderInterval * 1000
        Timer1.Enabled = True
    Else
        MenuFile.Enabled = True
        MenuEdit.Enabled = True
        MenuAdd.Enabled = True
        MenuTool.Enabled = True
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(5).Value = tbrUnpressed
        Timer1.Enabled = False
    End If
End Sub

Private Sub MenuToolCalc_Click()
    Shell "calc"
End Sub

Private Sub MenuToolTracer_Click()
    FrmTracer.Show
End Sub

Private Sub MenuView1_Click()
    ViewChange 90, 0
End Sub

Private Sub MenuView2_Click()
    ViewChange 180, 90
End Sub

Private Sub MenuView3_Click()
    ViewChange 180, 0
End Sub

Private Sub MenuViewReset_Click()
    ViewChange 75, 30
End Sub

Private Sub Pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragStartX = X
    DragStartY = Y
    ZoomStartY = Y
    ZoomStartK = K
    Pressed = True
    Saved = False
End Sub

Private Sub Pic2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Pressed Then
        If Button = 2 Then
            CenterX = CenterX + X - DragStartX
            CenterY = CenterY + Y - DragStartY
            DragStartX = X
            DragStartY = Y
            Redraw
        ElseIf Button = 3 Then
            K = ZoomStartK * 1.01 ^ (Y - ZoomStartY)
            Redraw
        ElseIf Button = 1 Then
            Alpha = Alpha + (X - DragStartX) / 500
            If Alpha > Pi * 2 Then Alpha = Alpha - Pi * 2
            If Alpha < 0 Then Alpha = Alpha + Pi * 2
            Beta = Beta + (Y - DragStartY) / 500
            If Beta > Pi * 2 Then Beta = Beta - Pi * 2
            If Beta < 0 Then Beta = Beta + Pi * 2
            DragStartX = X
            DragStartY = Y
            Redraw
        End If
    End If
End Sub

Private Sub Pic2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pressed = False
End Sub

Private Sub Pic1_Resize()
    Pic2.Top = Pic1.Top
    Pic2.Left = Pic1.Left
    Pic2.Height = Pic1.Height
    Pic2.Width = Pic1.Width
End Sub

Private Sub Slider1_Change()
    Alpha = Slider1.Value / 180 * Pi
    Redraw
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Alpha = Slider1.Value / 180 * Pi
        Saved = False
        Redraw
    End If
End Sub

Private Sub Slider2_Change()
    Beta = Slider2.Value / 180 * Pi
    Redraw
End Sub

Private Sub Slider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Beta = Slider2.Value / 180 * Pi
        Saved = False
        Redraw
    End If
End Sub

Private Sub Timer1_Timer()
    Update
    Redraw
    Saved = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "新建"
            MenuFileNew_Click
        Case "打开"
            MenuFileOpen_Click
        Case "保存"
            MenuFileSave_Click
        Case "运行"
            MenuRunStart_Click
        Case "属性"
            MenuEditProperty_Click
        Case "帮助"
            MenuHelpHelp_Click
    End Select
End Sub

