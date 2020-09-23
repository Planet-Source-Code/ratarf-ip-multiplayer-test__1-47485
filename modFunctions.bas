Attribute VB_Name = "modFunctions"
Public Function Parse(sIn As String, sDel As String) As Variant
'parse function, copied from other submission to spend less time. Very great.
Dim i As Integer, X As Integer, s As Integer, t As Integer
i = 1: s = 1: t = 1: X = 1
ReDim tArr(1 To X) As Variant
If InStr(1, sIn, sDel) <> 0 Then
Do
ReDim Preserve tArr(1 To X) As Variant
tArr(i) = Mid(sIn, t, InStr(s, sIn, sDel) - t)
t = InStr(s, sIn, sDel) + Len(sDel)
s = t
If tArr(i) <> "" Then i = i + 1
X = X + 1
Loop Until InStr(s, sIn, sDel) = 0
ReDim Preserve tArr(1 To X) As Variant
tArr(i) = Mid(sIn, t, Len(sIn) - t + 1)
Else
tArr(1) = sIn
End If
Parse = tArr
End Function

Public Function receive(text As String)
'parse input in different parts and check them (check function)
a = Parse(text, "/")
a1 = a(1)
a2 = a(2)
a3 = a(3)
check a1, a2, a3
End Function

Public Function check(info As Variant, info2 As Variant, back As Variant)
'Input was split up in 3 parts by receive function, and now it's validated

Dim answer As String

If info = "1" And info2 = "1" And back = "startgame" Then
frmMain.Show
frmStartup.Hide
End If

If back = "xy" Then
Select Case frmMain.Tag
Case "host":
frmMain.player2.Left = info
frmMain.player2.Top = info2
Case "client":
frmMain.player1.Left = info
frmMain.player1.Top = info2
End Select
End If

If info = "1" And info2 = "1" And back = "exitmenu" Then
frmStartup.lblStatus.Caption = "player disconnected"
frmMain.Winsock1.Close
cmdStartHost.Enabled = True
cmdStopHost.Enabled = False
cmdJoin.Enabled = True
cmdJoinDisconnect.Enabled = False
cmdStartGame.Enabled = False
txtChat.Enabled = False
End If

If info2 = "1" And back = "chat" Then
frmStartup.lstChat.AddItem (info)
End If

End Function


Function endgame()
'prompt yes/no, if yes ends game
Dim answer As String
answer = MsgBox("Are you sure you want to exit?", vbOKCancel, "Exit")
If answer = vbOK Then
End
End If
End Function






