VERSION 5.00
Begin VB.Form FrmPropertyEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "属性编辑器"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   10755
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame4 
      Caption         =   "匀强磁场"
      Height          =   3615
      Left            =   8040
      TabIndex        =   15
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List4 
         Height          =   2760
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command10 
         Caption         =   "属性"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         Caption         =   "复制"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command12 
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   3120
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "匀强电场"
      Height          =   3615
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List3 
         Height          =   2760
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command7 
         Caption         =   "属性"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "复制"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   3120
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "弹性绳"
      Height          =   3615
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List2 
         Height          =   2760
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "属性"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "复制"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   3120
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "小球"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command3 
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "复制"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "属性"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   2760
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmPropertyEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Selected As Integer

Private Sub Command1_Click()
    Selected = List1.ListIndex + 1
    BallReturn = Balls(Selected)
    SetBall
    If DialogOK Then
        Balls(Selected) = BallReturn
        Redraw
    End If
    RefreshList
End Sub

Private Sub Command2_Click()
    If BallCount = 100 Then
        MsgBox "小球已达最大数量！", vbCritical
        Exit Sub
    End If
    Selected = List1.ListIndex + 1
    BallReturn = Balls(Selected)
    SetBall
    If DialogOK Then
        BallCount = BallCount + 1
        Balls(BallCount) = BallReturn
        Redraw
    End If
    RefreshList
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    Selected = List1.ListIndex + 1
    If LineCount > 0 Then
        For i = 1 To LineCount
            If Lines(i).Ball1 = Selected Or Lines(i).Ball2 = Selected Then
                MsgBox "请先删除连接在小球上的弹性绳" & i & "！", vbExclamation
                Exit Sub
            End If
        Next
    End If
    If MsgBox("确定删除小球" & Selected & "？", vbYesNo) = vbNo Then Exit Sub
    BallCount = BallCount - 1
    If BallCount > 0 Then
        For i = Selected To BallCount
            Balls(i) = Balls(i + 1)
        Next
    End If
    If LineCount > 0 Then
        For i = 1 To LineCount
            With Lines(i)
                If .Ball1 > Selected Then .Ball1 = .Ball1 - 1
                If .Ball2 > Selected Then .Ball2 = .Ball2 - 1
            End With
        Next
    End If
    Saved = False
    Redraw
    RefreshList
End Sub

Private Sub Command4_Click()
    Selected = List2.ListIndex + 1
    LineReturn = Lines(Selected)
    SetLine
    If DialogOK Then
        Lines(Selected) = LineReturn
        Redraw
    End If
    RefreshList
End Sub

Private Sub Command5_Click()
    If LineCount = 100 Then
        MsgBox "弹性绳已达最大数量！", vbCritical
        Exit Sub
    End If
    Selected = List2.ListIndex + 1
    LineReturn = Lines(Selected)
    SetLine
    If DialogOK Then
        LineCount = LineCount + 1
        Lines(LineCount) = LineReturn
        Redraw
    End If
    RefreshList
End Sub

Private Sub Command6_Click()
    Dim i As Integer
    Selected = List2.ListIndex + 1
    If MsgBox("确定删除弹性绳" & Selected & "？", vbYesNo) = vbNo Then Exit Sub
    LineCount = LineCount - 1
    If LineCount > 0 Then
        For i = Selected To LineCount
            Lines(i) = Lines(i + 1)
        Next
    End If
    Saved = False
    Redraw
    RefreshList
End Sub

Private Sub Command7_Click()
    Selected = List3.ListIndex + 1
    EFieldReturn = EFields(Selected)
    SetEField
    If DialogOK Then
        EFields(Selected) = EFieldReturn
        Redraw
    End If
    RefreshList
End Sub

Private Sub Command8_Click()
    If EFieldCount = 100 Then
        MsgBox "匀强电场已达最大数量！", vbCritical
        Exit Sub
    End If
    Selected = List3.ListIndex + 1
    EFieldReturn = EFields(Selected)
    SetEField
    If DialogOK Then
        EFieldCount = EFieldCount + 1
        EFields(EFieldCount) = EFieldReturn
        Redraw
    End If
    RefreshList
End Sub

Private Sub Command9_Click()
    Dim i As Integer
    Selected = List3.ListIndex + 1
    If MsgBox("确定删除匀强电场" & Selected & "？", vbYesNo) = vbNo Then Exit Sub
    EFieldCount = EFieldCount - 1
    If EFieldCount > 0 Then
        For i = Selected To EFieldCount
            EFields(i) = EFields(i + 1)
        Next
    End If
    Saved = False
    Redraw
    RefreshList
End Sub

Private Sub Command10_Click()
    Selected = List4.ListIndex + 1
    MFieldReturn = MFields(Selected)
    SetMField
    If DialogOK Then
        MFields(Selected) = MFieldReturn
        Redraw
    End If
    RefreshList
End Sub

Private Sub Command11_Click()
    If MFieldCount = 100 Then
        MsgBox "匀强磁场已达最大数量！", vbCritical
        Exit Sub
    End If
    Selected = List4.ListIndex + 1
    MFieldReturn = MFields(Selected)
    SetMField
    If DialogOK Then
        MFieldCount = MFieldCount + 1
        MFields(MFieldCount) = MFieldReturn
        Redraw
    End If
    RefreshList
End Sub

Private Sub Command12_Click()
    Dim i As Integer
    Selected = List4.ListIndex + 1
    If MsgBox("确定删除匀强磁场" & Selected & "？", vbYesNo) = vbNo Then Exit Sub
    MFieldCount = MFieldCount - 1
    If MFieldCount > 0 Then
        For i = Selected To MFieldCount
            MFields(i) = MFields(i + 1)
        Next
    End If
    Saved = False
    Redraw
    RefreshList
End Sub

Private Sub Form_Load()
    RefreshList
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
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        List1.ToolTipText = ""
    Else
        Frame1.Enabled = False
    End If
    List2.Clear
    If LineCount > 0 Then
        For i = 1 To LineCount
            With Lines(i)
                List2.AddItem "[" & i & "] " & .Ball1 & "-" & .Ball2 & " L=" & .Length & " K=" & .K
            End With
        Next
        Command4.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        List2.ToolTipText = ""
    Else
        Frame2.Enabled = False
    End If
    List3.Clear
    If EFieldCount > 0 Then
        For i = 1 To EFieldCount
            With EFields(i)
                List3.AddItem "[" & i & "] E=(" & .E.X & "," & .E.Y & "," & .E.Z & ")"
            End With
        Next
        Command7.Enabled = False
        Command8.Enabled = False
        Command9.Enabled = False
        List3.ToolTipText = ""
    Else
        Frame3.Enabled = False
    End If
    List4.Clear
    If MFieldCount > 0 Then
        For i = 1 To MFieldCount
            With MFields(i)
                List4.AddItem "[" & i & "] B=(" & .B.X & "," & .B.Y & "," & .B.Z & ")"
            End With
        Next
        Command10.Enabled = False
        Command11.Enabled = False
        Command12.Enabled = False
        List4.ToolTipText = ""
    Else
        Frame4.Enabled = False
    End If
End Sub

Private Sub List1_Click()
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    List1.ToolTipText = List1.List(List1.ListIndex)
End Sub

Private Sub List1_DblClick()
    Command1_Click
End Sub

Private Sub List2_Click()
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    List2.ToolTipText = List2.List(List2.ListIndex)
End Sub

Private Sub List2_DblClick()
    Command4_Click
End Sub

Private Sub List3_Click()
    Command7.Enabled = True
    Command8.Enabled = True
    Command9.Enabled = True
    List3.ToolTipText = List3.List(List3.ListIndex)
End Sub

Private Sub List3_DblClick()
    Command7_Click
End Sub

Private Sub List4_Click()
    Command10.Enabled = True
    Command11.Enabled = True
    Command12.Enabled = True
    List4.ToolTipText = List4.List(List4.ListIndex)
End Sub

Private Sub List4_DblClick()
    Command10_Click
End Sub
