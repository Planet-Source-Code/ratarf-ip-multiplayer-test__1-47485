VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image player2 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   4440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image player1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   3480
      Top             =   4560
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'validate user input
Select Case frmMain.Tag
Case "host": 'if the player is the host
frmMain.Winsock1.SendData (player1.Left & "/" & player1.Top & "/xy") 'send coordinates
Select Case KeyCode 'move the player process
Case 37:
player1.Left = player1.Left - 250
Case 38:
player1.Top = player1.Top - 250
Case 39:
player1.Left = player1.Left + 250
Case 40:
player1.Top = player1.Top + 250
Case 27:
endgame 'escape ends the game
End Select
Case "client": 'if the player is the client
frmMain.Winsock1.SendData (player2.Left & "/" & player2.Top & "/xy") 'send coordinates
Select Case KeyCode 'move the player process
Case 37:
player2.Left = player2.Left - 250
Case 38:
player2.Top = player2.Top - 250
Case 39:
player2.Left = player2.Left + 250
Case 40:
player2.Top = player2.Top + 250
Case 27:
endgame 'escape ends the game
End Select
End Select
End Sub

Private Sub Form_Load()
'load player pictures
player1.Picture = LoadPicture(App.Path & "\pics\player1.bmp")
player2.Picture = LoadPicture(App.Path & "\pics\player2.bmp")
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'close winsock, accept the incomming request and enable/disable some buttons
Winsock1.Close
Winsock1.Accept (requestID)
frmStartup.lblStatus.Caption = "Someone connected"
frmStartup.cmdStartHost.Enabled = False
frmStartup.cmdStopHost.Enabled = True
frmStartup.cmdJoin.Enabled = False
frmStartup.cmdJoinDisconnect.Enabled = False
frmStartup.cmdStartGame.Enabled = True
frmStartup.txtChat.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'when data arrives, receive it and check it (take a look at the module)
Dim data As String
Winsock1.GetData data
modFunctions.receive (data)
End Sub

