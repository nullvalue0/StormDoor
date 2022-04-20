VERSION 5.00
Object = "{84AF4DF3-4B59-4D87-85BA-FA878460F831}#6.4#0"; "StormDoor.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin StormDoorX.StormDoor StormDoor1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8916
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Menu As String   'this holds which menu we're sitting at
Public InputBuffer As String    'this holds the input buffer

Private Sub Form_Load()
    StormDoor1.OpenDropFile Command
    StormDoor1.ClearDisplay
    ShowMainMenu
End Sub

Private Sub StormDoor1_UserInput(data As String)
    'this occurs automatically whenever a key is pressed by the client
    Select Case Menu
        Case "Main"
            Select Case UCase(data)
                Case "E"
                    StormDoor1.Display "\n\nEnter a string up to 10 characters long: "
                    Menu = "WaitForString"
                    InputBuffer = ""
                Case "Q"
                    StormDoor1.Quit
                    End
                Case Else
                    StormDoor1.Display "Please select a valid option.\n"
                    ShowMainMenu
            End Select
        Case "WaitForString"
            If InStr(1, data, vbCr) > 0 Then
                StormDoor1.Display "\n\nYou entered: " & InputBuffer & "\n\n------------------------------------------\n\n"
                ShowMainMenu
            ElseIf data = Chr(8) Then '(backspace)
                If InputBuffer <> "" Then
                    StormDoor1.Display Chr(27) & "[D " & Chr(27) & "[D"
                    InputBuffer = Left(InputBuffer, Len(InputBuffer) - 1)
                End If
            Else
                If Len(InputBuffer) < 10 Then
                    InputBuffer = InputBuffer & data
                    StormDoor1.Display data
                End If
            End If
    End Select
End Sub

Sub ShowMainMenu()
    StormDoor1.Display "Welcome to the main menu.\n\n"
    StormDoor1.Display "(E) Enter a string value\n"
    StormDoor1.Display "(Q) Quit\n\n"
    StormDoor1.Display "This main menu shows how you can just wait for 1 input character\n\n"
    StormDoor1.Display "Your Choice: (E,Q) "
    Menu = "Main"
End Sub
