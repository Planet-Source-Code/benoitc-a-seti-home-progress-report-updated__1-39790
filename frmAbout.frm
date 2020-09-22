VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   "About"
   ClientHeight    =   5295
   ClientLeft      =   4665
   ClientTop       =   2175
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5745
      TabIndex        =   2
      Top             =   4605
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "setiMstats V1.0 - 10/2002"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   2130
      MousePointer    =   10  'Up Arrow
      TabIndex        =   5
      Top             =   1965
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLI: Command Line CLient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   2985
      MousePointer    =   10  'Up Arrow
      TabIndex        =   4
      Top             =   4950
      Width           =   2445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1080
      X2              =   6285
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "All images in this application are copyrighted by the SETI Team: http://setiathome.berkeley.edu"
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
      Height          =   555
      Index           =   2
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   3
      Top             =   3825
      Width           =   6870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1080
      X2              =   6285
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© 2002 - BC Consulting www.bc-consult.com"
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
      Height          =   570
      Index           =   1
      Left            =   135
      MousePointer    =   10  'Up Arrow
      TabIndex        =   1
      Top             =   4560
      Width           =   2445
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
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
      Height          =   945
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   2535
      Width           =   6495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1080
      X2              =   6285
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Image Image1 
      Height          =   2115
      Left            =   30
      Picture         =   "frmAbout.frx":00D5
      Top             =   75
      Width           =   7350
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' * * * * * * * * * * * * * * * * * * * * * '
'                                           '
'       SETI@Home progress reports          '
'        for multiple CLI clients           '
'                                           '
'         © 2002 - BC Consulting            '
'           www.bc-consult.com              '
'                                           '
'------------------------------------------ '
'                                           '
'  Feel free to modify, but the original    '
'  copyright must stay                      '
'  All images copyrighted by the SETI team: '
'  http://setiathome.berkeley.edu           '
' * * * * * * * * * * * * * * * * * * * * * '

'frmAbout ---
'to satisfy my ego!!!!
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Form_Load()

   Label1(4) = "setiMstats V" & App.Major & "." & App.Minor & "." & _
             App.Revision & " - 10/2002"

End Sub

Private Sub Label1_Click(Index As Integer)

Dim i As Long
Dim L As String

   Select Case Index
   Case 1
      'Launch browser to www.BC-Consult.com
      L = "http://www.bc-consult.com/"
      i = ShellExecute(hWnd, "open", L, 0&, 0&, vbNormalFocus)
   Case 2
      'Launch browser to www.BC-Consult.com
      L = "http://setiathome.berkeley.edu/"
      i = ShellExecute(hWnd, "open", L, 0&, 0&, vbNormalFocus)
   End Select

End Sub
