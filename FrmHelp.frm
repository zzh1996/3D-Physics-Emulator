VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "帮助 - 三维物理"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7110
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Text1.Text = HelpText
End Sub
