Attribute VB_Name = "modSetiStats"
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

'modSetStats ---
'Support functions
'GetKeys:     read a value from the .ini file
'FillStats:   Update the stats of  CLI
'SetKey:      write a value to the .ini file
'ToNiceTime:  format a long to a time
'UpdateStats: Update the stats grid on the main form (frmStats)
'==============================================================
Option Explicit

Public INIF As String      '.ini file location
Public ar() As String      'names of the items to watch
Public CLIs() As String    'locations of the SETI folders to monitor

'.ini manipulation
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Function GetKey(S As String, K As String) As String
'Get a key from the .ini file

Dim i As Long
Dim L As String

   L = Space$(255)
   i = GetPrivateProfileString(S, K, "", L, 255, INIF)
   If i > 0 Then L = Left$(L, i) Else L = ""

   GetKey = L

End Function
Sub FillStats(Direc As String, Posi As Long)
'Update the stats of  CLI

Dim j As Long, j1 As Long, j0 As Long, j2 As Long
Dim ifi As Long, i As Long, n As Long
Dim perc As Double
Dim Fil As String, Ite As String
Dim Buff1 As String, Buff2 As String

   On Error Resume Next
   Fil = Direc & "\work_unit.sah"
   ifi = FreeFile
   Buff1 = Space$(FileLen(Fil))
   If Err = 0 Then
      Open Fil For Binary Access Read As ifi Len = 10240
      If Err Then
         'we cannot access the file, and we do not care about what error it is
         MsgBox "The SETI client on " & Direc & " cannot be accessed! (work_unit.sah)", vbExclamation, "SETI stats"
         Buff1 = ""
      Else
         Get #ifi, , Buff1
         Close ifi
         If Err Then Buff1 = ""
         'Only keep the header info from that file
         'we do not care about the data content!
         i = InStr(Buff1, "end_seti_header")
         If i = 0 Then
            MsgBox "Wrong WU file! " & Fil, vbCritical, "SETI Stats"
         Else
            Buff1 = "work_unit.sah" & vbCrLf & Left$(Buff1, i - 1)
         End If
      End If
   Else
      Buff1 = ""
   End If
   Err = 0

   Fil = Direc & "\state.sah"
   ifi = FreeFile
   Buff2 = Space$(FileLen(Fil))
   If Err = 0 Then
      Open Fil For Binary Access Read As ifi Len = 1024
      If Err Then
         'we cannot access the file, and we do not care about what error it is
         MsgBox "The SETI client on " & Direc & " cannot be accessed! (state.sah)", vbExclamation, "SETI stats"
      Else
         Get #ifi, , Buff2
         Close ifi
         If Err = 0 Then Buff1 = Buff1 & "state.sah" & vbCrLf & Buff2
      End If
   End If
   On Error GoTo 0

   'Add other sah files as defined in .ini
   n = Val(GetKey("SAH", "n"))
   If n > 0 Then
      On Error Resume Next
      For i = 1 To n
         Fil = Direc & "\" & GetKey("SAH", "file" & CStr(i))
         ifi = FreeFile
         Buff2 = Space$(FileLen(Fil))
         If Err = 0 Then
            Open Fil For Binary Access Read As ifi Len = 1024
            If Err = 0 Then
               Get #ifi, , Buff2
               Close ifi            '                filename       its content
               If Err = 0 Then Buff1 = Buff1 & vbCrLf & Fil & vbCrLf & Buff2
            End If
         End If
         Err = 0
      Next
      On Error GoTo 0
   End If

   If Buff1 <> "" Then
   'Loop through the buffer for the relevant
   'items to display
      With frmStats!Grid1
      .Col = Posi
      For i = 1 To UBound(ar)
         .Row = i

         j = InStr(ar(i), "|")
         If j > 0 Then
            'entry in the form: filename|item
            Ite = Mid$(ar(i), j + 1)               'item
            j = InStr(Buff1, Left$(ar(i), j - 1))  'from file
         Else
            Ite = ar(i)
            j = 1
         End If

         If j > 0 Then
            'extract line
            j = InStr(j, Buff1, Ite)
            If j > 0 Then
               j0 = InStr(j, Buff1, vbCr)
               j1 = InStr(j, Buff1, vbLf)
               If j0 > 0 And j1 > 0 Then
                  If j0 > j1 Then j0 = j1
               ElseIf j0 = 0 And j1 > 0 Then
                  j0 = j1
               End If
               If j0 > 0 Then
                  'extract value
                  Buff2 = Mid$(Buff1, j, j1 - j)
                  j0 = InStr(Buff2, "=")
                  j1 = InStr(Buff2, " ")
                  If j0 > 0 And j1 > 0 Then
                     If j0 > j1 Then j0 = j1
                  ElseIf j0 = 0 And j1 > 0 Then
                     j0 = j1
                  End If
                  If j0 > 0 Then
                     Buff2 = Mid$(Buff2, j0 + 1)
                     j = InStr(Buff2, "(")
                     If j > 0 Then
                        'extract date/time
                        Buff2 = Mid$(Buff2, j + 1)
                        j = InStr(Buff2, ")")
                        Buff2 = Left$(Buff2, j - 1)
                     End If
                     .Col = 0
                     If InStr(.Text, "%") > 0 Then
                     'represent value in %
                        If Int(Val(Buff2)) = 0 Then
                           perc = Val(Buff2) * 100
                           Buff2 = Format$(perc, "00.00")
                        End If
                        Buff2 = Buff2 & " %"
                     End If
                     If InStr(LCase$(Ite), "cpu") > 0 Then
                     'this is the time elapsed since start of
                     'processing this WU
                        Buff2 = ToNiceTime(Buff2)
                     End If
                     'Display value
                     .Col = Posi
                     .Text = Buff2
                  End If
               End If
            End If
         End If
      Next
      End With
   End If

End Sub
Function SetKey(S As String, K As String, V As String) As Long
'Set a key (K) to its value (V) in the .ini file

   SetKey = WritePrivateProfileString(S, K, V, INIF)

End Function
Function ToNiceTime(L As String) As String
'Format a long to a time

Dim j As Long, j0 As Long, j1 As Long, j2 As Long
Dim Buff2 As String
   j2 = 0
   j = Val(L) \ 3600                   'hours
   j0 = (Val(L) - j * 3600) \ 60       'minutes
   j1 = (Val(L) - j * 3600 - j0 * 60)  'seconds
   If j > 24 Then
   'Days
      j2 = j \ 24
      j = (j - j2 * 24)
   End If
   Buff2 = Format$(j, "##00") & ":" & Format$(j0, "00") & ":" & Format$(j1, "00")
   If j2 = 1 Then Buff2 = "1 day " & Buff2
   If j2 > 1 Then Buff2 = CStr(j2) & " days " & Buff2

   ToNiceTime = Buff2

End Function

Function UpdateStats() As Boolean
'Update the stats grid on the main form

Dim i As Long, n As Long

   n = UBound(CLIs)
   If n = 0 Then
      UpdateStats = False
      Exit Function
   End If

   For i = 1 To n
      FillStats CLIs(i), i
   Next
   UpdateStats = True

End Function
