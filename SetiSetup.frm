VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   4665
   ClientLeft      =   3885
   ClientTop       =   1515
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SetiSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Height          =   390
      Index           =   1
      Left            =   3660
      TabIndex        =   9
      Top             =   4185
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Index           =   1
      Left            =   1455
      TabIndex        =   7
      Top             =   4170
      Width           =   2145
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Done"
      Height          =   390
      Index           =   0
      Left            =   7440
      TabIndex        =   6
      Top             =   4215
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&V"
      Height          =   390
      Left            =   8055
      TabIndex        =   3
      Top             =   3735
      Width           =   435
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Index           =   0
      Left            =   1470
      TabIndex        =   2
      Top             =   3720
      Width           =   6585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2805
      Left            =   75
      TabIndex        =   0
      Top             =   435
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   4948
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   0
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8085
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "sah"
      DialogTitle     =   "Locate the CLI client folder"
      FileName        =   "*.sah"
      Filter          =   "SETI files|*.sah|All files|*.*"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CLI Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   150
      TabIndex        =   8
      Top             =   4245
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Monitored CLI clients :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   75
      TabIndex        =   5
      Top             =   60
      Width           =   2385
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CLI Folder :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   3795
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add a new CLI client to the list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   75
      TabIndex        =   1
      Top             =   3390
      Width           =   3210
   End
End
Attribute VB_Name = "frmSetup"
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

'frmSetup ---
'Allow to add new SETI folders to monitor
'To add a folder to the watch list:
'1) Enter the location of the seti folder
'   or click the 'V' button to navigate to
'   that folder, there select any file and
'   click 'Ok'
'2) Enter a name for that folder.
'   if it is a network a good name is the
'   network machine name, otherwise enter
'   a name that has significance to you!
'3) Click the 'Add' button
'   the new folder is added to the list
'   as well as to the main form list.
Option Explicit
Sub GetCLIclients()
'Display all registered clients

Dim n As Long, i As Long

   n = Val(GetKey("Setup", "n"))
   If n > 0 Then
      With Grid1
      For i = 1 To n
         If i > 1 Then .Rows = .Rows + 1
         .Row = i
         .Col = 0: .Text = GetKey("Setup", "CLI" & CStr(i))
         .Col = 1: .Text = GetKey("Setup", "Name" & CStr(i))
      Next
      End With
   End If

End Sub
Private Sub Command1_Click()
'Get folder

Dim i As Long
Dim L As String

   With CD1
   On Error Resume Next
   'show file dialog box
   .Action = 1
   If .FileName <> "" And InStr(.FileName, "*") = 0 And Err = 0 Then
      'fill in text box
      L = .FileName
      i = InStrRev(L, "\")
      L = Left$(L, i - 1)
      Text1(0).Text = L
   End If
   On Error GoTo 0
   End With

End Sub

Private Sub Command2_Click(Index As Integer)

Dim n As Long
Dim L As String

   Select Case Index
   Case 0
   'Done
      Unload Me

   Case 1
   'Add
      'Add the client to list on frmStats
      With frmStats!Grid1
      .Col = 1: .Row = 0
      If Trim$(.Text) <> "" Then
         .Cols = .Cols + 1
         .Col = .Cols - 1
      End If
      .Row = 0
      .Text = Text1(1).Text
      .ColAlignmentFixed = 7
      .ColAlignment = 1
      .ColWidth(.Col) = 3270
      End With

      'show in grid
      With Grid1
      .Row = 1: .Col = 0
      If .Text <> "" Then
         .Rows = .Rows + 1
         .Row = .Rows - 1
      End If
      .Col = 0: .Text = Text1(0).Text
      .Col = 1: .Text = Text1(1).Text
      '-- number of CLI clients
      n = .Rows - 1
      'save into .ini file
      SetKey "Setup", "n", CStr(n)
      L = "CLI" & CStr(n)
      SetKey "Setup", L, Text1(0).Text
      L = "Name" & CStr(n)
      SetKey "Setup", L, Text1(1).Text
      ReDim Preserve CLIs(n) As String
      CLIs(n) = Text1(0).Text
      End With
      'update stats
      If UpdateStats() And Not frmStats!Frame1.Visible Then frmStats!Frame1.Visible = True
   End Select

End Sub
Private Sub Form_Load()

   With Grid1
   .Row = 0
   .Col = 0:   .Text = " Folder name "
   .Col = 1:   .Text = " Given name "
   .ColWidth(0) = Me.TextWidth("_Folder name_") * 3
   .ColWidth(1) = Me.TextWidth("_Given name_")
   End With

   GetCLIclients

End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

   If KeyAscii = 13 Then KeyAscii = 0

End Sub
