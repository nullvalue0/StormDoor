VERSION 5.00
Object = "{84AF4DF3-4B59-4D87-85BA-FA878460F831}#6.4#0"; "StormDoor.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTelDoor 
   Caption         =   "TelDoor"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "frmTelDoor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin StormDoorX.StormDoor StormDoor1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9128
   End
End
Attribute VB_Name = "frmTelDoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TelDoor is a simple door used to telnet to another server.
'You may alter this code and use it however you see fit.

Private Sub Form_Load()
Dim sTelnetServer As String
Me.Show
DoEvents
StormDoor1.OpenDropFile Command
StormDoor1.ChangeColor fBlue, bCyan
StormDoor1.Display " TelDoor v1.0                                                         nullvalue \n\n"
StormDoor1.ChangeColor fLightGreen, bBlack

sTelnetServer = "bbsmates.com"

StormDoor1.Display "Opening telnet session to " & sTelnetServer & "...\n\n"
StormDoor1.ChangeColor fGray, bBlack
Winsock1.Connect sTelnetServer, 23  'open a winsock connection to telnet port
End Sub

Private Sub StormDoor1_UserInput(data As String)
    'fired when the client types something at their terminal
    If Winsock1.State = 7 Then  '(if we've established a connection)
        Winsock1.SendData data  'send the typed data over the telneted connection
    End If
End Sub

Private Sub Winsock1_Close()
    'if we lose a connection to the telnet server, disconnect and end TelDoor.
    StormDoor1.Display "Connection to host lost.\n"
    StormDoor1.Quit
    End
End Sub

Private Sub Winsock1_Connect()
    'fired when a connection is fully established
    StormDoor1.ChangeColor fLightGreen, bBlack
    StormDoor1.Display "Connected...\n"
    StormDoor1.ChangeColor fGray, bBlack
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    'occurs when the telnet server sends data
    Dim sInBuffer As String, i
    Winsock1.GetData sInBuffer      'get the input from the server
    StormDoor1.Display sInBuffer    'now relay it back to the client
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'occurs when a winsock error occurs, like if we can't resolve a server hostname, or if a connection attempt times out
    StormDoor1.Display "Error: " & Description
End Sub
