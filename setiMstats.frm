VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStats 
   BackColor       =   &H00000000&
   Caption         =   "Stats for multi CLI clients"
   ClientHeight    =   4860
   ClientLeft      =   3720
   ClientTop       =   1800
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "setiMstats.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   9600
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4635
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   8100
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   4200
         Width           =   1290
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   6585
         TabIndex        =   2
         Top             =   4185
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   2880
         Left            =   45
         TabIndex        =   1
         Top             =   1245
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   5080
         _Version        =   393216
         BackColor       =   12632256
         ForeColor       =   12583104
         Rows            =   9
         BackColorFixed  =   16512
         ForeColorFixed  =   12648447
         BackColorSel    =   12632256
         BackColorBkg    =   0
         BackColorUnpopulated=   12632256
         GridColorFixed  =   0
         GridColorUnpopulated=   0
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         SelectionMode   =   2
         AllowUserResizing=   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   195
         TabIndex        =   6
         Top             =   795
         Width           =   1995
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   30
         Picture         =   "setiMstats.frx":1272
         Top             =   195
         Width           =   7500
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   7440
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Progress reports for multi CLI clients"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1050
         Left            =   1875
         TabIndex        =   4
         Top             =   1800
         Width           =   4965
      End
      Begin VB.Image Image1 
         Height          =   2115
         Left            =   0
         Picture         =   "setiMstats.frx":66C8
         Top             =   0
         Width           =   7350
      End
   End
   Begin VB.Menu mnu0 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuF 
         Caption         =   "&Setup..."
         Index           =   0
      End
      Begin VB.Menu mnuF 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuF 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin VB.Menu mnu0 
      Caption         =   "&Help"
      Index           =   1
      Begin VB.Menu mnuH 
         Caption         =   "&About..."
         Index           =   0
      End
      Begin VB.Menu mnuH 
         Caption         =   "&Help!..."
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmStats"
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
'------------------------------------------ '
'                                           '
'Version history                            '
'  1.0.1:  Added error handling on files    '
'          operations and possibility to    '
'          add item from other .sah files   '
'  1.0.0:  Original                         '
' * * * * * * * * * * * * * * * * * * * * * '

'frmStats ---
'This is the entry point to the program
'Display the progress reports or nothing
'If nothing is displayed, that's mean there
'is no CLI clients to watch...
'Use File->Setup to add CLI clients to the
'watch list...
'Click the 'Refresh' button to get the last
'updated progress report
'
'CLI=Command Line client
'===========================================
Option Explicit

Const mnuSETUP As Long = 0
Const mnuCLOSE As Long = 2
Const mnuABOUT As Long = 0
Const mnuHELP As Long = 1

Function GetHeaders() As Boolean
'Get all monitored headers

Dim n As Long, i As Long
Dim w As Long, w1 As Long

   n = Val(GetKey("Stats", "n"))
   If n = 0 Then
      MsgBox "Nothing to report on...", vbExclamation, "SETI stats"
      Exit Function
   End If

   ReDim ar(1 To n) As String
   w = 0
   With Grid1
   .Rows = n + 1
   .Col = 0
   For i = 1 To n
      ar(i) = GetKey("Stats", "st" & CStr(i))
      .Row = i
      .Text = GetKey("Stats", "d" & CStr(i))
      w1 = Me.TextWidth(.Text & "___")
      If w1 > w Then w = w1
   Next
   .ColAlignmentFixed = 7
   .ColAlignment = 1
   .ColWidth(0) = w
   End With

   GetHeaders = True

End Function
Private Sub Command1_Click(Index As Integer)

   Select Case Index
   Case 0
      Unload Me
   Case 1
      UpdateStats
   End Select

End Sub
Private Sub Form_Load()

Dim n As Long, i As Long
Dim CLIname As String

   INIF = App.Path & "\setiMstats.ini"
   If GetHeaders() Then
      'If exist .ini file read it
      n = Val(GetKey("Setup", "n"))
      If n > 0 Then
         'if exist CLI clients' profiles read them & show stats
         ReDim CLIs(n) As String
         With Grid1
         For i = 1 To n
            CLIs(i) = GetKey("Setup", "CLI" & CStr(i))
            If CLIs(i) <> "" Then
               CLIname = GetKey("Setup", "Name" & CStr(i))
               If CLIname <> "" Then
                  If i = 1 Then
                     .Col = 1
                  Else
                     .Cols = .Cols + 1
                     .Col = .Cols - 1
                     .ColAlignmentFixed = 7
                     .ColAlignment = 1
                  End If
                  .ColWidth(i) = 3270
                  .Row = 0: .Text = CLIname & "     "
                  FillStats CLIs(i), i
               End If
            End If
         Next
         End With
         Frame2.Visible = False
         Frame1.Visible = True

      'Otherwise do nothing!
      End If
   End If

End Sub
Private Sub Form_Resize()

   If WindowState = vbMinimized Then Exit Sub

   On Error Resume Next
   With Frame1
   .Move 30, 30, ScaleWidth - 60, ScaleHeight - 60
   Grid1.Move 45, 1245, .Width - 90, .Height - 435 - 1245
   Command1(0).Move .Width - Command1(0).Width - 45, .Height - 390
   Command1(1).Move 45, .Height - 390
   End With
   With Frame2
   .Left = (ScaleWidth - .Width) / 2
   .Top = (ScaleHeight - .Height) / 2
   End With
   On Error GoTo 0

End Sub
Private Sub Form_Unload(Cancel As Integer)

   'We do not have to kill COM objects or such rubish
   'so use brute force to close...
   End

End Sub

Private Sub mnuF_Click(Index As Integer)
'File menu

   Select Case Index
   Case mnuSETUP
   'Show setup form
      frmSetup.Show 1

   Case mnuCLOSE
   'Close app
      Unload Me
   End Select

End Sub
Private Sub mnuH_Click(Index As Integer)
'Help menu

   Select Case Index
   Case mnuABOUT
   'Show About form
      frmAbout.Show 1

   Case mnuHELP
   'Show help window
      frmHelp.Show
   End Select

End Sub


