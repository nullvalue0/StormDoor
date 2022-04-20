VERSION 5.00
Object = "{7256D622-8402-414E-945A-4A0AA7300B90}#2.0#0"; "StormDoor.ocx"
Begin VB.Form frmMain 
   Caption         =   "Paper-Rock-Scissors"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim Nodes(255) As String, TempArr(255) As Integer, ChallengeNode As Integer, sResponse As String, k, iUserSel As Integer, iChalSel As Integer, WaitTime As Long, score As String

Private Sub Form_Load()
    Me.Show
    
    Call Game
End Sub

Private Sub Game(Optional Go As String)
    Dim i, iCompSel As Integer, a As Integer, sSel As String
    
    If Go = "MainMenu" Then 'The "Go" parameter is used when returning from the NodeRecv
        GoTo MainMenu       'event to go to a particular label
    End If
    
    StormDoor1.Echo = True  'Turn on client input echo, this means the client should have
                            'local echo turned off
    
    StormDoor1.OpenDropFile Command 'This command should be ran before anything else.
                                    'It opens the door file, reads its contects into the
                                    'ActiveX properties, and establishes all neccessary
                                    'connections to the client (player).
                                
    'Initiate the score file - if the file doesn't exist, then create it
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    If fso.FileExists(App.Path & "\players.cfg") = False Then fso.CreateTextFile App.Path & "\players.cfg"
    Set fso = Nothing
    
    'Initiate this user's score - if this is your first time playing the game, then create a
    'slot for this new player in the score file
    score = GetScore(StormDoor1.alias)
    If score = "" Then
        SaveNew StormDoor1.alias
        score = 0
    End If
    
    Me.Caption = "Paper-Rock-Scissors - Node: " & StormDoor1.ThisNode
    Nodes(StormDoor1.ThisNode) = StormDoor1.alias 'Nodes() is an array, which holds the Alias
                                                  'of who's on [index] node
MainMenu:   'This is where the game begins, displaying the main menu
    'Sends the ANSI command to clear the screen
    StormDoor1.ClearDisplay
    
    'Change the foreground color to Light Green
    StormDoor1.ChangeColor fLightGreen
    
    'Display a welcome message, including the player's alias
    StormDoor1.Display "Welcome to Paper-Rock-Scissors, " & StormDoor1.alias & "!\n\n"
    
    'Change foreground color to Light Purple
    StormDoor1.ChangeColor fLightPurple
    StormDoor1.Display "Please choose from the following menu:\n\n"
    StormDoor1.ChangeColor fLightBlue
    
    'Display the actual menu. "\n" may be used in the display string, to represent a carriage return & line feed
    StormDoor1.Display "    W) Who's in the game?\n    C) Play against the computer\n    P) Play against another node\n    S) Scores\n    A) About the game\n    Q) Quit\n\n"
    StormDoor1.ChangeColor fYellow
    StormDoor1.Display "Your choice? "
    
    'Wait for the player's keyboard input. Focus will only return back if one of the keys listed
    'are pressed. Seperate valid keys for this prompt with commas inside a string.
    k = StormDoor1.WaitForKey("W,C,P,A,S,Q")
    
    'Now k holds which key was actually pressed, so run the following code based on what was selected.
    Select Case k
    Case "W"
        'Player selected "W" for Who's in the game?
        StormDoor1.ChangeColor fWhite
        StormDoor1.Display "\n\n"
        
        'Loop through our Nodes array, displaying each user who is online.
        For i = 0 To 255
            'If Nodes([index]) is not empty, then there is a user on node (i).
            If Nodes(i) <> "" Then StormDoor1.Display Nodes(i) & vbCrLf
        Next i

        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "\n PRESS ANY KEY"
        
        'Wait for player's keyboard input. Since there are no WaitKeys assigned, focus will
        'return when any key is pressed.
        StormDoor1.WaitForKey
    Case "C"
        'Player selected "C" for Play Against the Computer.
PlayAgain:
        StormDoor1.ClearDisplay
        StormDoor1.ChangeColor fLightGreen
        StormDoor1.Display "Playing against the computer\n\n"
        StormDoor1.ChangeColor fLightPurple
GoAgain:
        StormDoor1.Display "Make your choice:\n\n"
        StormDoor1.ChangeColor fLightBlue
        
        'Display menu to for player to select Paper, Rock, Scissors, or Quit.
        StormDoor1.Display "    P) Paper\n    R) Rock\n    S) Scissors\n\n    Q) Quit to Main\n\n"
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "Your choice? "
        
        'Wait for the player's to select Paper, Rock, Scissors, or Quit.
        k = StormDoor1.WaitForKey("R,P,S,Q")
        Select Case k
        Case "R"
            'User selected Rock, just store this selection for now.
            iUserSel = 1
        Case "P"
            'User selected Paper, just store this selection for now.
            iUserSel = 2
        Case "S"
            'User selected Scissors, just store this selection for now.
            iUserSel = 3
        Case "Q"
            'User selected Quit, go back to the main menu. When GoTo MainMenu is called, the
            'Main Menu gets displayed again, and we start back at the beginning.
            GoTo MainMenu
        End Select
        
        'Generate a random number, from 1 to 3. This will be the computer's selection for
        'Paper, Rock, or Scissors
        Randomize
        iCompSel = CInt((Rnd * 2) + 1)
        
        'Display what was selected by both the player and the computer.
        StormDoor1.ChangeColor fWhite
        StormDoor1.Display "\n\n     You Selected: " & GetSelName(iUserSel) & vbCrLf
        StormDoor1.Display "Computer Selected: " & GetSelName(iCompSel) & "\n\n"
        
        'FYI:
        '   1=Rock
        '   2=Paper
        '   3=Scissors
        If iUserSel = iCompSel Then
            'If both the player's and the computer's selections are the same, then there was
            'a tie. Go to the GoAgain label, a try again.
            StormDoor1.ChangeColor fYellow
            StormDoor1.Display "Tie. Go again.\n\n"
            GoTo GoAgain
        ElseIf (iUserSel = 1 And iCompSel = 3) Or (iUserSel = 2 And iCompSel = 1) Or (iUserSel = 3 And iCompSel = 2) Then
            '   Rock beats Scissors               Paper beats Rock                   Scissors beats Paper
            StormDoor1.ChangeColor fLightGreen
            StormDoor1.Display "You Win!  :)\n\n"
            'Add 1 to the player's score
            score = score + 1
            'Save the player's new score into the score file
            SaveScore StormDoor1.alias, score
        ElseIf (iUserSel = 1 And iCompSel = 2) Or (iUserSel = 2 And iCompSel = 3) Or (iUserSel = 3 And iCompSel = 1) Then
            '   Rock loses to Paper                Paper loses to Scissors            Scissors loses to Rock
            StormDoor1.ChangeColor fLightRed
            StormDoor1.Display "You Lose.  :(\n\n"
            'Subtract from the player's score
            score = score - 1
            'Save the player's new score into the score file
            SaveScore StormDoor1.alias, score
        End If
            
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "\n Play Again? (Y/N) [Y] "
        
        'Ask the player if he wants to play again. The second parameter, [Default] is set
        'to "Y", meaning if the enter key is pressed, "Y" is automatically selected.
        k = StormDoor1.WaitForKey("Y,N", "Y")
        If k = "Y" Then
            'If Yes is selected, go to the PlayAgain label
            GoTo PlayAgain
        Else
            'If No is selected, go back to the Main Menu
            GoTo MainMenu
        End If
    Case "P"
        'Player selected "P" for Play against another node.
        'Ok, here's where it gets a little tricky...
        a = 0
        sSel = ""
        StormDoor1.ClearDisplay
        StormDoor1.ChangeColor fLightGreen
        StormDoor1.Display "Playing against another node.\n\n"
        StormDoor1.ChangeColor fLightPurple
        StormDoor1.Display "Select user: \n\n"
        StormDoor1.ChangeColor fLightBlue
        
        'Display the list of other online players to chose from.
        For i = 1 To 255
            'Loop through our array of online players.
            If Nodes(i) <> "" Then
                'i <> StormDoor1.ThisNode makes sure we don't display our own node.
                If i <> StormDoor1.ThisNode Then
                    'Count up the number of players actually online, store that number in 'a'
                    a = a + 1
                    'Print out the player's Alias
                    StormDoor1.Display "    " & a & ") " & Nodes(i) & vbCrLf
                    'Store that player's node in an array, so we know later which node is
                    'selected when someone chooses by Alias
                    TempArr(a) = i
                    'Build the WaitKeys string for later...
                    sSel = sSel & CStr(a) & ","
                End If
            End If
        Next i
        
        'If a=0, that means there were no other players online (beside yourself)
        If a = 0 Then
            'So tell the player that, and kick him back to the main menu
            StormDoor1.ClearDisplay
            StormDoor1.ChangeColor fLightGreen
            StormDoor1.Display "Playing against another node.\n\n"
            StormDoor1.ChangeColor fLightRed
            StormDoor1.Display "Sorry, there are no other players online right now.\n"
            StormDoor1.ChangeColor fYellow
            StormDoor1.Display "\n PRESS ANY KEY"
            StormDoor1.WaitForKey
            GoTo MainMenu
        End If
            
        'Add the Quit option to the WaitKeys string
        sSel = sSel & "Q"
        StormDoor1.Display "\n    Q) Quit to Main\n\n"
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "Your choice? "
        
        'Wait for the player to select the other node to play against.
        k = StormDoor1.WaitForKey(sSel)
        If k = "Q" Then
            'Selected Quit, go back to main menu
            GoTo MainMenu
        Else
            'Player must have selected a number for the player they wanted to challenge.
        
            'Get the node of the player we're challenging
            ChallengeNode = TempArr(k)
            
            sResponse = ""
            iChalSel = "0"
            'Send a message to the other node that we're requesting a challege.
            If StormDoor1.SendNode(ChallengeNode, "CHALLENGE") = False Then
                StormDoor1.ChangeColor fLightRed
                StormDoor1.Display "Sorry, cross-node communcation failed.\n"
                StormDoor1.ChangeColor fYellow
                StormDoor1.Display "\n PRESS ANY KEY"
                StormDoor1.WaitForKey
                GoTo MainMenu
            Else
                StormDoor1.ClearDisplay
                StormDoor1.ChangeColor fLightGreen
                
                'Wait for the other node to reply back whether he accepted the challenge,
                'or denied it.
                StormDoor1.Display "Waiting for other node to answer.\n\n"
                
                'This allows us to make sure that if the other node doesn't answer within
                '60 seconds, to exit out of this wait loop.
                'The Do While loop is looking to make sure that 60 seconds hasn't expired yet
                'and is also looking for there to be a value in sResponse. The value of
                'sResponse gets set in the NodeRecv when the other node selects Yes or No.
                WaitTime = Timer + 60
                Do While WaitTime > Timer And sResponse = ""
                    DoEvents
                Loop
                
                'If the value of sReponse is empty, that means we exited from the above loop
                'because the 60 seconds elapsed.
                If sResponse = "" Then
                    'So display a message stating that, and go back to the main menu.
                    StormDoor1.ChangeColor fLightRed
                    StormDoor1.Display "Sorry, the other node failed to acknowledge your request in a timely manner.\n"
                    StormDoor1.ChangeColor fYellow
                    StormDoor1.Display "\n PRESS ANY KEY"
                    StormDoor1.WaitForKey
                    GoTo MainMenu
                ElseIf sResponse = "NOT ACCEPTED" Then
                    'If sResponse is "NOT ACCEPTED", that means the other node pressed
                    '(N)o and declined the challenge.
                    StormDoor1.ChangeColor fLightRed
                    StormDoor1.Display "Sorry, the other node denied your request.\n"
                    StormDoor1.ChangeColor fYellow
                    StormDoor1.Display "\n PRESS ANY KEY"
                    StormDoor1.WaitForKey
                    GoTo MainMenu
                ElseIf sResponse = "ACCEPTED" Then
                    'If sResponse is "ACCEPTED", that means the other node pressed (Y)es
                    'and accepted the challenge.
                    StormDoor1.ChangeColor fLightGreen
                    
                    'Display who we're challenging.
                    StormDoor1.Display "Playing Against " & Nodes(ChallengeNode) & "\n\n"
                    StormDoor1.ChangeColor fLightPurple
GoAgainChal:
                    StormDoor1.Display "Make your choice:\n\n"
                    StormDoor1.ChangeColor fLightBlue
                    'Display the menu to chose from Paper, Rock, Scissors, or Quit.
                    StormDoor1.Display "    P) Paper\n    R) Rock\n    S) Scissors\n\n    Q) Quit to Main\n\n"
                    StormDoor1.ChangeColor fYellow
                    StormDoor1.Display "Your choice? "
                    'Wait for the player to select on of these options.
                    k = StormDoor1.WaitForKey("R,P,S,Q")
                    Select Case k
                        Case "R"
                            'User selected Rock, just store this selection for now.
                            iUserSel = 1
                        Case "P"
                            'User selected Paper, just store this selection for now.
                            iUserSel = 2
                        Case "S"
                            'User selected Scissors, just store this selection for now.
                            iUserSel = 3
                        Case "Q"
                            'User selected Quit. So first, we must tell the other node that
                            'we selected Quit.
                            StormDoor1.SendNode ChallengeNode, "QUIT"
                            
                            'Clear this variable since we're not playing him anymore.
                            ChallengeNode = 0
                            
                            'Go back to the main menu.
                            GoTo MainMenu
                    End Select
                    'Now that we've made our selection, wait until the other node makes a selection.
                    StormDoor1.Display "\n\nPlease wait for other node..."
                    'Same thing as before, wait for the other node to make a selection within
                    '60 seconds. iChalSel is assigned a value in the NodeRecv event when the other
                    'node makes a selection from his menu.
                    WaitTime = Timer + 60
                    Do While WaitTime > Timer And iChalSel = "0"
                        DoEvents
                    Loop
                    
                    'If iChalSel=0, that means the 60 seconds elapsed without the other node
                    'making a selection.
                    If iChalSel = 0 Then
                        'Tell the player that the other node didn't make a selection.
                        StormDoor1.ChangeColor fLightRed
                        StormDoor1.Display "Sorry, the other node failed to acknowledge your request in a timely manner.\n"
                        StormDoor1.ChangeColor fYellow
                        StormDoor1.Display "\n PRESS ANY KEY"
                        StormDoor1.WaitForKey
                        ChallengeNode = 0
                        GoTo MainMenu
                    End If
                    
                    'Now that we got here, that means both of us have made selections, so let's
                    'see who won.
                    
                    'First display our selections.
                    StormDoor1.ChangeColor fWhite
                    StormDoor1.Display "\n\n     You Selected: " & GetSelName(iUserSel) & vbCrLf
                    StormDoor1.Display Nodes(ChallengeNode) & " Selected: " & GetSelName(iChalSel) & "\n\n"
                    If iUserSel = iChalSel Then
                        'If the selections are equal, we tied.
                        StormDoor1.ChangeColor fYellow
                        StormDoor1.Display "Tie. Go again.\n\n"
                        'It is the challenging node's responsibility to tell the other node
                        'that we tied. The last character of this message to the other node
                        'will contain the player's selection, so the other node knows what
                        'the player selected.
                        StormDoor1.SendNode ChallengeNode, "TIE" & iUserSel
                        'Then go back and we both have to make new selections.
                        iChalSel = 0
                        GoTo GoAgainChal
                    ElseIf (iUserSel = 1 And iChalSel = 3) Or (iUserSel = 2 And iChalSel = 1) Or (iUserSel = 3 And iChalSel = 2) Then
                        '   Rock beats Scissors               Paper beats Rock                   Scissors beats Paper
                        StormDoor1.ChangeColor fLightGreen
                        'Player won, so display that.
                        StormDoor1.Display "You Win!  :)\n\n"
                        'And add 5 to his score.
                        score = score + 5
                        'Then save his new score to the score file.
                        SaveScore StormDoor1.alias, score
                        'It is the challenging node's responsibility to tell the other node
                        'that he lost.
                        'The last character of this message to the other node will contain the
                        'player's selection, so the other node knows what the player selected.
                        StormDoor1.SendNode ChallengeNode, "LOS" & iUserSel
                    ElseIf (iUserSel = 1 And iChalSel = 2) Or (iUserSel = 2 And iChalSel = 3) Or (iUserSel = 3 And iChalSel = 1) Then
                        '   Rock loses to Paper                Paper loses to Scissors            Scissors loses to Rock
                        StormDoor1.ChangeColor fLightRed
                        'Player lost, so display that.
                        StormDoor1.Display "You Lose.  :(\n\n"
                        'And subtract 5 from his score.
                        score = score - 5
                        'Then save his new score to the score file.
                        SaveScore StormDoor1.alias, score
                        'It is the challenging node's responsibility to tell the other node
                        'that he lost. The last character of this message to the other node
                        'will contain the player's selection, so the other node knows what
                        'the player selected.
                        StormDoor1.SendNode ChallengeNode, "WIN" & iUserSel
                    End If
                    'If we get here, that means there was a deciding end to the challenge
                    '(either this node won or lost), so wait for a key then go back to the
                    'main menu.
                    StormDoor1.ChangeColor fYellow
                    StormDoor1.Display "\n PRESS ANY KEY"
                    StormDoor1.WaitForKey
                    GoTo MainMenu
                End If
            End If
        End If
            
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "\n PRESS ANY KEY"
        StormDoor1.WaitForKey
    Case "S"
        'Player selected "S" for Scores.
        StormDoor1.ClearDisplay
        StormDoor1.ChangeColor fWhite
        StormDoor1.Display "Current Paper-Rock-Scissors Scores\n\n"
        Dim s As String, t
        s = ""
        'Open the score file.
        Open App.Path & "\players.cfg" For Input As #1
        Do While Not EOF(1)
            Line Input #1, t
            'store the file in s
            s = s & t & vbCrLf
        Loop
        Close #1
        s = Replace(s, "=", ": ")
        'Then basically, just display the file.
        StormDoor1.ChangeColor fLighCyan
        StormDoor1.Display s
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "\n PRESS ANY KEY"
        StormDoor1.WaitForKey
    Case "A"
        'Player selected "A" for About the game
        StormDoor1.ClearDisplay
        StormDoor1.ChangeColor fWhite
        StormDoor1.Display "About The Game\n\n\n"
        StormDoor1.ChangeColor fBrown
        StormDoor1.Display "        Paper-Rock-Scissors is a stupid game I have made to demonstrate\n"
        StormDoor1.Display "               the abilities of StormDoor, my ActiveX DoorKit\n\n"
        StormDoor1.Display "                                 -nullvalue\n\n"
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "\n PRESS ANY KEY"
        StormDoor1.WaitForKey
    Case "Q"
        'Player selected "Q" for Quit game
        StormDoor1.ClearDisplay
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "Quitting back to BBS..."
        DoEvents
        'When gracefully quitting the game, run this command to disconnect and clean everything up.
        StormDoor1.Quit
        'Shut down the door game.
        End
    End Select
    
    GoTo MainMenu

    Exit Sub
End Sub

Private Sub StormDoor1_NodeRecv(FromNode As Integer, data As String)
    'This event is fired when another node sends a message to this node.
    
    'Another node has challenged this node to play online.
    If data = "CHALLENGE" Then
        StormDoor1.ChangeColor fMagenta
        'Display who has challenged this node.
        StormDoor1.Display "\n\n " & Nodes(FromNode) & " has challenged you to an online contest.\n\n"
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "Do you accept? (Y/N) [Y] "
        'And ask the player if he want to accept or deny the challenge.
        'The third parameter for WaitForKey is the TimeLimit for this prompt. If the player
        'does not response with a valid selection within 55 seconds, k will contain and empty string.
        k = StormDoor1.WaitForKey("Y,N", "Y", 55)
        If k = "Y" Then
            'Player selected "Y" and accepted the challenge.
            If StormDoor1.SendNode(FromNode, "ACCEPTED") = False Then
                StormDoor1.ChangeColor fLightRed
                StormDoor1.Display "Sorry, cross-node communcation failed.\n"
                StormDoor1.ChangeColor fYellow
                StormDoor1.Display "\n PRESS ANY KEY"
                StormDoor1.WaitForKey
                Game "MainMenu"
            Else
                StormDoor1.ClearDisplay
                StormDoor1.ChangeColor fLightGreen
                
                'Display who we're playing against.
                StormDoor1.Display "Playing Against " & Nodes(FromNode) & "\n\n"
                StormDoor1.ChangeColor fLightPurple
GoAgain:
                'Prompt for this node to make it's selection of Paper, Rock, Scissors, or Quit.
                StormDoor1.Display "Make your choice:\n\n"
                StormDoor1.ChangeColor fLightBlue
                StormDoor1.Display "    P) Paper\n    R) Rock\n    S) Scissors\n\n    Q) Quit to Main\n\n"
                StormDoor1.ChangeColor fYellow
                StormDoor1.Display "Your choice? "
                'Wait for player's input.
                k = StormDoor1.WaitForKey("R,P,S,Q")
                Select Case k
                    Case "R"
                        'User selected Rock, so tell the challenging node that.
                        iUserSel = 1
                        StormDoor1.SendNode FromNode, "SEL1"
                    Case "P"
                        'User selected Paper, so tell the challenging node that.
                        iUserSel = 2
                        StormDoor1.SendNode FromNode, "SEL2"
                    Case "S"
                        'User selected Scissors, so tell the challenging node that.
                        iUserSel = 3
                        StormDoor1.SendNode FromNode, "SEL3"
                    Case "Q"
                        'User selected Quit, so tell the challenging node that.
                        StormDoor1.SendNode FromNode, "QUIT"
                        'Then go back to the main menu.
                        Game "MainMenu"
                End Select
                'Now wait for the challenging node to make his selection
                StormDoor1.Display "\n\nPlease wait for other node..."
                sResponse = ""
                'Wait up to 60 seconds to hear his response back. The challenging node's
                'response will tell this node whether we won, lost, or tied.
                'sResponse is assigned a value down below in this NodeRecv event.
                WaitTime = Timer + 60
                Do While WaitTime > Timer And sResponse = ""
                    DoEvents
                Loop
                'So if sResponse is empty, that means the 60 seconds elapsed.
                If sResponse = "" Then
                    StormDoor1.ChangeColor fLightRed
                    StormDoor1.Display "Sorry, the other node failed to acknowledge your request in a timely manner.\n"
                    StormDoor1.ChangeColor fYellow
                    StormDoor1.Display "\n PRESS ANY KEY"
                    StormDoor1.WaitForKey
                    ChallengeNode = 0
                    Game "MainMenu"
                Else
                    'We got a response back from the challenging node, so display our selections.
                    StormDoor1.ChangeColor fWhite
                    StormDoor1.Display "\n\n     You Selected: " & GetSelName(iUserSel) & vbCrLf
                    StormDoor1.Display Nodes(FromNode) & " Selected: " & GetSelName(Right(sResponse, 1)) & "\n\n"
                    
                    If Left(sResponse, 3) = "TIE" Then
                        'This means we tied, display that message and go back to make another selection.
                        StormDoor1.ChangeColor fYellow
                        StormDoor1.Display "Tie. Go again.\n\n"
                        sResponse = ""
                        GoTo GoAgain
                    ElseIf Left(sResponse, 3) = "WIN" Then
                        'This means the player won.
                        StormDoor1.ChangeColor fLightGreen
                        StormDoor1.Display "You Win!  :)\n\n"
                    ElseIf Left(sResponse, 3) = "LOS" Then
                        'This means the player lost.
                        StormDoor1.ChangeColor fLightRed
                        StormDoor1.Display "You Lose.  :(\n\n"
                    End If
                    StormDoor1.ChangeColor fYellow
                    StormDoor1.Display "\n PRESS ANY KEY"
                    StormDoor1.WaitForKey
                    Game "MainMenu"
                End If
            End If
        ElseIf k = "N" Then
            'Player selected "N" to deny the challenge, so tell the challenging node
            'that it was denied.
            If StormDoor1.SendNode(FromNode, "NOT ACCEPTED") = False Then
                StormDoor1.ChangeColor fLightRed
                StormDoor1.Display "Sorry, cross-node communcation failed.\n"
                StormDoor1.ChangeColor fYellow
                StormDoor1.Display "\n PRESS ANY KEY"
                StormDoor1.WaitForKey
                Game "MainMenu"
            Else
                'Then go back to the main menu.
                Game "MainMenu"
            End If
        Else
            'The player did not make a selection within 55 seconds, so kick them back to the main menu.
            StormDoor1.ChangeColor fRed
            StormDoor1.Display "\n\nSorry, you didn't response in time.\n"
            StormDoor1.ChangeColor fYellow
            StormDoor1.Display "\n PRESS ANY KEY"
            StormDoor1.WaitForKey
            Game "MainMenu"
        End If
    ElseIf FromNode = ChallengeNode And data = "ACCEPTED" Then
        'This sets the response variable when this node is waiting to hear back from challenged
        'node, whether or not they accept or deny the challenge.
        'This means the other node pressed "Y" and accepted the challenge.
        sResponse = "ACCEPTED"
    ElseIf FromNode = ChallengeNode And data = "NOT ACCEPTED" Then
        'This sets the response variable when this node is waiting to hear back from challenged
        'node, whether or not they accept or deny the challenge.
        'This means the other node pressed "N" and declined the challenge.
        sResponse = "NOT ACCEPTED"
    ElseIf (FromNode = ChallengeNode) And (data = "SEL1" Or data = "SEL2" Or data = "SEL3") Then
        'This sets the "challenger selected" variable when this node is waiting to hear back
        'from challenged node. This means the other node pressed "Y" and accepted the challenge.
        iChalSel = Right(data, 1)
    ElseIf Left(data, 3) = "TIE" Or Left(data, 3) = "LOS" Or Left(data, 3) = "WIN" Then
        sResponse = data
    ElseIf data = "QUIT" Then
        ChallengeNode = 0
        StormDoor1.ChangeColor fRed
        StormDoor1.Display "\n\nThe other player canceled this game.\n"
        StormDoor1.ChangeColor fYellow
        StormDoor1.Display "\n PRESS ANY KEY"
        StormDoor1.WaitForKey
        Game "MainMenu"
    Else
        StormDoor1.Display "Unknown Data: " & data & "\n"
    End If
End Sub

Private Sub StormDoor1_NodeSignon(Node As Integer, alias As String)
    'This event is fired when another node signs onto the game.
    'So we add his alias to our array of who's online.
    Nodes(Node) = alias
End Sub

Private Sub StormDoor1_NodeSignoff(Node As Integer, alias As String)
    'This event is fired when another node signs off from the game.
    'So just remove his alias from our array.
    Nodes(Node) = ""
End Sub

Private Sub StormDoor1_RanOutofTime()
    'This event is fired when the time limit is up. StormDoor keeps track of how much time the
    'player is allowed, based on what the doorfile said.
    StormDoor1.Display "Ran out of time... bye!"
    StormDoor1.Quit
    End
End Sub

Private Sub StormDoor1_ConnectionClosed()
    'This event is fired when the player unexpectedly drops his connection. In this case, just
    'shut down the door game.
    End
End Sub

Private Function GetSelName(val As Integer)
    'This is just used to display a player's selection
    If val = 1 Then
        GetSelName = "Rock"
    ElseIf val = 2 Then
        GetSelName = "Paper"
    ElseIf val = 3 Then
        GetSelName = "Scissors"
    End If
End Function

Private Function GetScore(alias As String)
    'Retrieves a player's score, if the player isn't found, it returns an empty string
    Dim s, i
    Open App.Path & "\players.cfg" For Input As #1
    Do While Not EOF(1)
        Line Input #1, s
        i = InStr(1, s, "=")
        If i > 0 Then
            If Left(s, i - 1) = alias Then
                GetScore = Mid(s, i + 1)
                Close #1
                Exit Function
            End If
        End If
    Loop
    Close #1
    GetScore = ""
End Function

Private Sub SaveScore(alias As String, newscore As String)
    'Save the player's score to the text file
    Dim s, i, t, sc
    Open App.Path & "\players.cfg" For Input As #1
    Do While Not EOF(1)
        Line Input #1, t
        s = s & t & vbCrLf
    Loop
    Close #1
    
    sc = GetScore(alias)
    
    s = Replace(s, alias & "=" & sc, alias & "=" & newscore)
    
    Open App.Path & "\players.cfg" For Output As #1
    Print #1, Left(s, Len(s) - 2)
    Close #1
End Sub

Private Sub SaveNew(alias As String)
    'Save a new record in the score file for a first-time player
    Open App.Path & "\players.cfg" For Append As #1
    Print #1, alias & "=0"
    Close #1
End Sub

