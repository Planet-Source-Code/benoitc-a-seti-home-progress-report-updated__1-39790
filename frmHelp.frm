VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   Caption         =   "How to use setiMstats"
   ClientHeight    =   5775
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9360
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4275
      Left            =   45
      TabIndex        =   1
      Top             =   1185
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   7541
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmHelp.frx":0000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "setiMstats"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   720
      Width           =   9255
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   15
      Picture         =   "frmHelp.frx":00D7
      Top             =   15
      Width           =   7500
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   Label1 = "setiMstats V" & App.Major & "." & App.Minor & "." & _
             App.Revision & " - 10/2002"
   Text1.LoadFile App.Path & "\help.rtf"

End Sub
Private Sub Form_Resize()

   If WindowState = vbMinimized Then Exit Sub

   On Error Resume Next
   Label1.Move 0, 720, ScaleWidth
   Text1.Move 30, 1115, ScaleWidth - 60, ScaleHeight - 1160
   On Error GoTo 0

End Sub
