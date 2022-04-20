VERSION 5.00
Object = "{84AF4DF3-4B59-4D87-85BA-FA878460F831}#6.3#0"; "StormDoor.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin StormDoorX.StormDoor StormDoor1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8916
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This short program shows how to place color on string you want at run time
'For Example if you wanted to set a color in a config file for users to be able to change
'for a certin string that would be displayed. Like a color change on a press any key prompt
'Color System Written by Nullvalue & example packaged up for you by Black Phantom

Dim intFgC As Integer 'Holds the ForeGround color
Dim intBgC As Integer 'Holds the BackGround color
Public ConfigColor As String 'Holds the text that tells us what color is being displayed

Private Sub Form_Load()
StormDoor1.OpenDropFile Command 'Open Drop File

intFgC = 0 'set starting forground color to black
intBgC = 0 'set starting background color to black
    
    For b = 1 To 7
        intBgC = intBgC + 1
            StormDoor1.ChangeColor GetFG(intFgC), GetBG(intBgC) 'Run thru FOR till all background colors
            StormDoor1.Display ConfigColor & "\n"               'have been displayed
    Next b

intBgC = 0 'reset that background color to black
    
    For f = 1 To 15
        intFgC = intFgC + 1
            StormDoor1.ChangeColor GetFG(intFgC), GetBG(intBgC) 'Run thru FOR till forground colors
            StormDoor1.Display ConfigColor & "\n"               'have been displayed
        Next f
End Sub

Function GetFG(colornum As Integer)
    Select Case colornum
        Case 0
            GetFG = 30
            ConfigColor = "[BLACK]"
        Case 1
            GetFG = 31
            ConfigColor = "[RED]"
        Case 2
            GetFG = 32
            ConfigColor = "[GREEN]"
        Case 3
            GetFG = 33
            ConfigColor = "[BROWN]"
        Case 4
            GetFG = 34
            ConfigColor = "[BLUE]"
        Case 5
            GetFG = 35
            ConfigColor = "[PURPLE]"
        Case 6
            GetFG = 36
            ConfigColor = "[CYAN]"
        Case 7
            GetFG = 37
            ConfigColor = "[GRAY]"
        Case 8
            GetFG = 130
            ConfigColor = "[DARK GRAY]"
        Case 9
            GetFG = 131
            ConfigColor = "[BRIGHT RED]"
        Case 10
            GetFG = 132
            ConfigColor = "[LIGHT GREEN]"
        Case 11
            GetFG = 133
            ConfigColor = "[YELLOW]"
        Case 12
            GetFG = 134
            ConfigColor = "[LIGHT BLUE]"
        Case 13
            GetFG = 135
            ConfigColor = "[LIGHT PURPLE]"
        Case 14
            GetFG = 136
            ConfigColor = "[LIGHT CYAN]"
        Case 15
            GetFG = 137
            ConfigColor = "[WHITE]"
        Case Else
            GetFG = 137
            ConfigColor = "[WHITE]"
    End Select
End Function
 
'this one is for background:
Function GetBG(colornum As Integer)
    Select Case colornum
        Case 0
            GetBG = 40
            ConfigColor = ConfigColor & " on [BLACK]"
        Case 1
            GetBG = 41
            ConfigColor = ConfigColor & " on [RED]"
        Case 2
            GetBG = 42
            ConfigColor = ConfigColor & " on [GREEN]"
        Case 3
            GetBG = 43
            ConfigColor = ConfigColor & " on [BROWN]"
        Case 4
            GetBG = 44
            ConfigColor = ConfigColor & " on [BLUE]"
        Case 5
            GetBG = 45
            ConfigColor = ConfigColor & " on [PURPLE]"
        Case 6
            GetBG = 46
            ConfigColor = ConfigColor & " on [CYAN]"
        Case 7
            GetBG = 47
            ConfigColor = ConfigColor & " on [GRAY]"
        Case Else
            GetBG = 40
            ConfigColor = ConfigColor & " on [RED]"
    End Select
End Function

