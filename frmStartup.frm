VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP multiplayer test"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Chat"
      Height          =   2535
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtChat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   3135
      End
      Begin VB.ListBox lstChat 
         Height          =   1425
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start game"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Join game"
      Height          =   1695
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdJoinDisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtIp 
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Text            =   "000.000.000.000"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdJoin 
         Caption         =   "Join"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Host game"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton cmdStopHost 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdStartHost 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "No connection"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   3735
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hostConnection As Boolean

Private Sub cmdExit_Click()
Dim answer As String
answer = MsgBox("Are you sure you want to exit?", vbOKCancel, "Exit")
If answer = vbOK Then
End
End If
End Sub

Private Sub cmdJoin_Click()
'try to connect to server
frmMain.Tag = "client"
frmMain.Winsock1.Close
connecttoserver frmMain.Winsock1, txtIp.text, lblStatus
End Sub

Private Sub cmdJoinDisconnect_Click()
'disconnect from server
frmMain.Winsock1.Close
cmdStartHost.Enabled = True
cmdStopHost.Enabled = False
cmdJoin.Enabled = True
cmdJoinDisconnect.Enabled = False
cmdStartGame.Enabled = False
txtChat.Enabled = False
lblStatus.Caption = "Disconnected"
End Sub

Private Sub cmdStartGame_Click()
'start game (only enabled if you are host, and if someone is connected)
frmMain.Winsock1.SendData ("1/1/startgame")
frmMain.Show
frmStartup.Hide
End Sub

Private Sub cmdStartHost_Click()
'start hosting
frmMain.Tag = "host"
cmdStartHost.Enabled = False
cmdStopHost.Enabled = True
cmdJoin.Enabled = False
cmdJoinDisconnect.Enabled = False
cmdStartGame.Enabled = False
txtChat.Enabled = False
frmMain.Winsock1.Close
frmMain.Winsock1.LocalPort = "666"
frmMain.Winsock1.Listen
lblStatus.Caption = "Host started"
End Sub

Private Sub cmdStopHost_Click()
'stop hosting
frmMain.Winsock1.Close
cmdStartHost.Enabled = True
cmdStopHost.Enabled = False
cmdJoin.Enabled = True
cmdJoinDisconnect.Enabled = False
cmdStartGame.Enabled = False
txtChat.Enabled = False
lblStatus.Caption = "Host stopped"
End Sub

Public Function connecttoserver(socket As Winsock, ip As String, Optional outputlabel As Label) As Boolean
'connect to server
outputlabel.Caption = "Connecting..."
socket.Close
socket.Connect ip, 666
Do Until socket.State <> 6
cmdStartHost.Enabled = False
cmdStopHost.Enabled = False
cmdJoin.Enabled = False
cmdJoinDisconnect.Enabled = True
cmdStartGame.Enabled = False
txtChat.Enabled = False
DoEvents
Loop
If socket.State = 7 Then
cmdStartHost.Enabled = False
cmdStopHost.Enabled = False
cmdJoin.Enabled = False
cmdJoinDisconnect.Enabled = True
cmdStartGame.Enabled = False
txtChat.Enabled = True
outputlabel.Caption = "Connected"
connecttoserver = True
Else
cmdStartHost.Enabled = True
cmdStopHost.Enabled = False
cmdJoin.Enabled = True
cmdJoinDisconnect.Enabled = False
cmdStartGame.Enabled = False
txtChat.Enabled = False
outputlabel.Caption = "Failed to connect"
socket.Close
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
'exit app is form is unloaded
End
End Sub

Private Sub txtChat_KeyUp(KeyCode As Integer, Shift As Integer)
'send chat string
If KeyCode = 13 Then
lstChat.AddItem (txtChat.text)
frmMain.Winsock1.SendData (txtChat.text & "/1/chat")
txtChat.text = ""
End If
End Sub
