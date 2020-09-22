VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SETI@Home information"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image5 
      Height          =   495
      Left            =   5760
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warning! The source code of this program is ONLY for private use..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   5595
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make sure to place 'Information.exe' in same direction with 'SETI@home.exe'"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   6525
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.octeam.cjb.net"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2520
      MouseIcon       =   "about.frx":1272
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3240
      Width           =   2625
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "overlord@sunpoint.net"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2520
      MouseIcon       =   "about.frx":1F3C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3000
      Width           =   2115
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visit our homepages:"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mail me for comments:"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   2115
      Left            =   120
      Picture         =   "about.frx":2C06
      Top             =   120
      Width           =   7350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SETI@Home information version:"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3240
      TabIndex        =   1
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows version programmed by OverLord, O.C. TeaM"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   5010
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   495
      Left            =   5760
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   4440
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Caption = App.Major & "." & App.Minor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Form2.Hide
End Sub


Private Sub Image5_Click()
Form2.Hide
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label25.Left = 5780
Label25.Top = 4460
Shape4.BorderColor = 16711680
End Sub


Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label25.Left = 5760
Label25.Top = 4440
Shape4.BorderColor = 16744576
End Sub


Private Sub Label6_Click()
Call ShellExecute(0&, vbNullString, "mailto:overlord@sunpoint.net", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = 16761024
End Sub


Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = 16777215
End Sub

Private Sub Label7_Click()
Call ShellExecute(0&, vbNullString, "http://www.octeam.cjb.net", vbNullString, vbNullString, vbNormalFocus)
End Sub


Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = 16761024
End Sub


Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = 16777215
End Sub


