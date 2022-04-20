VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl StormDoor 
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   HitBehavior     =   0  'None
   ScaleHeight     =   5145
   ScaleWidth      =   9600
   ToolboxBitmap   =   "StormDoorX.ctx":0000
   Begin VB.Timer CheckNodeStatus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   0
   End
   Begin VB.Timer TimeLeftTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin MSWinsockLib.Winsock clientsock 
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock nodesock 
      Left            =   480
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer tmrBlink 
      Interval        =   200
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox picNorm 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9600
   End
   Begin VB.Timer NodeInput 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   240
      Top             =   120
   End
   Begin VB.PictureBox picBlnk 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   0
      Width           =   9600
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "StormDoor v0.3.3"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   8160
      TabIndex        =   14
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblTimeVal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblSecurityVal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblRealNameVal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblAliasVal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblBaudVal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblNodeVal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Baud Rate:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblNode 
      Alignment       =   1  'Right Justify
      Caption         =   "Node:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Remaining:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Security Level:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Real Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblAlias 
      Alignment       =   1  'Right Justify
      Caption         =   "Alias:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "StormDoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'StormDoor v0.3.3 beta

Option Explicit

Dim X As Integer, Y As Integer, iLastColor As Long
Dim SaveX As Integer, SaveY As Integer, f As String
Dim xOffSet As Integer, yOffSet As Integer
Public xScreenSize As Integer, yScreenSize As Integer
Dim bFontBold As Boolean, bFontItalic As Boolean
Dim bBlinkUsed As Boolean
Dim MyTime As Double, TempTime As Double
Public DisplayRate As Long
Dim ScrollTop As Integer, ScrollBottom As Integer, iLines As Integer
Public Echo As Boolean
Public Event NodeRecv(FromNode As Integer, data As String)
Public Event NodeSignon(node As Integer, Alias As String)
Public Event NodeSignoff(node As Integer, Alias As String)
Public Event UserInput(data As String)
Public Event RanOutofTime()
Public Event ConnectionClosed()

Public Event DebugDataChanges(debugdata As String)

Public debugdata As String
Public CommType As Integer
Public SocketHandle As Integer
Public BaudRate As Double
Public BBSID As String
Public UserRecord As Integer
Public RealName As String
Public Alias As String
Public SecurityLevel As Integer
Public TimeLeft As Integer
Public Emulation As Integer
Public ThisNode As Integer

Public NodeCount As Integer
Public NodeNumber As Integer
Public NodeAlias As String

Public expire As Date

Public Enum sdAnsiFG
    fBlack = 30
    fRed = 31
    fGreen = 32
    fBrown = 33
    fBlue = 34
    fMagenta = 35
    fCyan = 36
    fGray = 37
    fDarkGray = 130
    fLightRed = 131
    fLightGreen = 132
    fYellow = 133
    fLightBlue = 134
    fLightPurple = 135
    fLighCyan = 136
    fLightCyan = 136
    fWhite = 137
End Enum

Public Enum sdAnsiBG
    bBlack = 40
    bRed = 41
    bGreen = 42
    bYellow = 43
    bBlue = 44
    bMagenta = 45
    bCyan = 46
    bWhite = 47
End Enum

Private EndTime As Date
Private FirstNodeStatusCheck As Boolean
Private InputBuffer As String
Private lp As Integer
Private bLocal As Boolean


Function GetBG(colornum As Integer)
    Select Case colornum
        Case 0
            GetBG = 40
        Case 1
            GetBG = 41
        Case 2
            GetBG = 42
        Case 3
            GetBG = 43
        Case 4
            GetBG = 44
        Case 5
            GetBG = 45
        Case 6
            GetBG = 46
        Case Else
            GetBG = 40
    End Select
End Function

Public Function OpenDropFile(Optional path As String) As Boolean

    Dim sTemp As String
    
    If path = "" Then
        bLocal = True
        'Local mode, prompt for username
        Alias = InputBox("Starting in local mode, please enter your alias:", "StormDoor", "Sysop")
        ThisNode = 0
        TimeLeft = 1440
    ElseIf InStr(1, path, "-H") > 0 Then
        Alias = "unknown"
        ThisNode = 0
        TimeLeft = 1440
        SocketHandle = Trim(Replace(path, "-H", ""))
        MsgBox "'" & SocketHandle & "'"
        clientsock.Accept CLng(SocketHandle)
        If clientsock.State <> sckConnected Then
            Err.Raise 3331, "StormDoorX", "Could not connect to Telnet Socket"
            OpenDropFile = False
            Exit Function
        End If
    Else
        bLocal = False
        Open path For Input As #1
        Line Input #1, sTemp
        CommType = CInt(sTemp)
            
        Line Input #1, sTemp
        SocketHandle = CInt(sTemp)
            
        Line Input #1, sTemp
        BaudRate = CDbl(sTemp)
        
        Line Input #1, BBSID
            
        Line Input #1, sTemp
        UserRecord = CInt(sTemp)
            
        Line Input #1, RealName
                
        Line Input #1, Alias
            
        Line Input #1, sTemp
        SecurityLevel = CInt(sTemp)
            
        Line Input #1, sTemp
        TimeLeft = CInt(sTemp)
            
        Line Input #1, sTemp
        Emulation = CInt(sTemp)
            
        Line Input #1, sTemp
        ThisNode = CInt(sTemp)
        Close #1
        
        If CommType = 2 Then  'telnet
            clientsock.Accept SocketHandle
            If clientsock.State <> sckConnected Then
                Err.Raise 3331, "StormDoorX", "Could not connect to Telnet Socket"
                OpenDropFile = False
                Exit Function
            End If
        Else
        
        End If
        
        Display "\nPlease Wait, Loading...\n"
    End If
    
    lblBaudVal.Caption = BaudRate
    lblNodeVal.Caption = ThisNode
    lblRealNameVal.Caption = RealName
    lblAliasVal.Caption = Alias
    lblSecurityVal.Caption = SecurityLevel
    expire = DateAdd("n", TimeLeft, Now)
    Dim a, h, m, s
    a = DateDiff("s", Now, expire)
    If a > 1 Then
        h = Int(a / 3600)
        m = Int(a / 60) Mod 60
        s = a Mod 60
        lblTimeVal.Caption = Format(h & ":" & m & ":" & s, "h:mm:ss")
    End If
    
    DoEvents
    
    If Len(ThisNode) = 1 Then
        lp = "200" & ThisNode
    ElseIf Len(ThisNode) = 2 Then
        lp = "20" & ThisNode
    ElseIf Len(ThisNode) = 3 Then
        lp = "2" & ThisNode
    End If
    nodesock.LocalPort = lp
    nodesock.Bind
    
    EndTime = DateAdd("n", TimeLeft, Now)
    TimeLeftTimer.Enabled = True
    CheckNodeStatus.Enabled = True
    'NodeInput.Enabled = True
    
    FindDeadNodes
    SetNodeStatus ThisNode, Alias
    
    OpenDropFile = True
    
End Function

Private Sub CheckNodeStatus_Timer()
On Error Resume Next
    Dim stat As String, i As Integer
    If FirstNodeStatusCheck = True Then
        For i = 1 To 128
            NodeStatus(i) = GetSetting("StormDoorX", "NodeStatus", "Node" & Pad_String(CStr(i), 3, "0", 0), "")
            If NodeStatus(i) <> "" Then RaiseEvent NodeSignon(i, NodeStatus(i))
        Next i
        FirstNodeStatusCheck = False
    Else
        For i = 1 To 128
            stat = GetSetting("StormDoorX", "NodeStatus", "Node" & Pad_String(CStr(i), 3, "0", 0), "")
            If i <> ThisNode And stat <> NodeStatus(i) Then
                If stat = "" Then
                    RaiseEvent NodeSignoff(i, NodeStatus(i))
                Else
                    RaiseEvent NodeSignon(i, stat)
                End If
                NodeStatus(i) = stat
            End If
        Next i
    End If
End Sub

Private Sub clientsock_Close()
    RaiseEvent ConnectionClosed
End Sub

Private Sub clientsock_DataArrival(ByVal bytesTotal As Long)
    Dim sTemp As String
    clientsock.GetData sTemp
    RaiseEvent UserInput(sTemp)
    InputBuffer = InputBuffer & sTemp
    If Echo = True Then
        clientsock.SendData Replace(sTemp, Chr(8), Chr(8) & " " & Chr(8))
        DisplayLocal sTemp
    End If
End Sub

Private Sub nodesock_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Err_Handler
    Dim dat As String, frmnd As Integer, i As Integer

    nodesock.GetData dat
    
    i = InStr(1, dat, ",")
    'Find out which node this data came from
    frmnd = CInt(Left(dat, i - 1))
    'Now parse out the data
    dat = Mid(dat, i + 1)
    If dat = "testing node " & ThisNode Then
        'if the packet sent is a node test, just let it pass through,
        'which will return PASS to the test
    Else
        'Call the "data recv'd from another node" event
        RaiseEvent NodeRecv(frmnd, dat)
    End If

Exit Sub
Err_Handler:
    If Err.Number = 10054 Then
        'SendNode = False
        'Exit Function
    Else
        Err.Clear
        Resume Next
    End If
End Sub

Private Sub picBlnk_KeyPress(KeyAscii As Integer)
    RaiseEvent UserInput(Chr(KeyAscii))
    InputBuffer = InputBuffer & Chr(KeyAscii)
    If Echo = True Then DisplayLocal Chr(KeyAscii)
End Sub

Private Sub picNorm_KeyPress(KeyAscii As Integer)
    RaiseEvent UserInput(Chr(KeyAscii))
    InputBuffer = InputBuffer & Chr(KeyAscii)
    If Echo = True Then DisplayLocal Chr(KeyAscii)
End Sub

Private Sub TimeLeftTimer_Timer()
    Dim a, h, m, s
    a = DateDiff("s", Now, expire)
    If a > 1 Then
        h = Int(a / 3600)
        m = Int(a / 60) Mod 60
        s = a Mod 60
        lblTimeVal.Caption = Format(h & ":" & m & ":" & s, "h:mm:ss")
    End If
    If Now >= EndTime Then
        RaiseEvent RanOutofTime
        TimeLeftTimer.Enabled = False
    End If
End Sub

Public Function SendNode(ToNode As Integer, data As String) As Boolean
On Error GoTo Err_Handler

    If Len(ToNode) = 1 Then
        lp = "200" & ToNode
    ElseIf Len(ToNode) = 2 Then
        lp = "20" & ToNode
    ElseIf Len(ToNode) = 3 Then
        lp = "2" & ToNode
    End If
    
    nodesock.RemoteHost = "localhost"
    nodesock.RemotePort = lp
    nodesock.SendData ThisNode & "," & data
    SendNode = True

Exit Function

Err_Handler:
    If Err.Number = 10054 Then
        SendNode = False
        Exit Function
    Else
        Err.Clear
        Resume Next
    End If
End Function

Private Sub SetNodeStatus(node As Integer, name As String)
    SaveSetting "StormDoorX", "NodeStatus", "Node" & Pad_String(CStr(node), 3, "0", 0), name
End Sub

Private Sub txtScreen_KeyPress(KeyAscii As Integer)
    RaiseEvent UserInput(Chr(KeyAscii))
    InputBuffer = InputBuffer & Chr(KeyAscii)
    'txtScreen.Text = txtScreen.Text & Chr(KeyAscii)
End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer
    FirstNodeStatusCheck = True
    If GetSetting("StormDoorX", "NodeStatus", "Node001", "-never set-") = "-never set-" Then
        For i = 1 To 128
            SaveSetting "StormDoorX", "NodeStatus", "Node" & Pad_String(CStr(i), 3, "0", 0), ""
        Next i
    End If
    
    picNorm.ForeColor = GetColor(7)
    picBlnk.ForeColor = GetColor(7)
    picNorm.FontSize = 9
    picBlnk.FontSize = 9
    xOffSet = 8
    yOffSet = 12
    xScreenSize = 80
    yScreenSize = 24
    ScrollTop = 1
    ScrollBottom = yScreenSize
    
    bFontBold = False
    bFontItalic = False
    
    DisplayRate = 0
    
End Sub

Private Sub FindDeadNodes()
    Dim i As Integer
    
    CheckNodeStatus_Timer
    
    For i = 1 To 128
        If NodeStatus(i) <> "" Then
            If SendNode(i, "testing node " & i) = False Then SetNodeStatus i, ""
        End If
    Next i
End Sub

Private Sub UserControl_Terminate()
    If ThisNode <> 0 Then SetNodeStatus ThisNode, ""
End Sub

Public Sub Display(data As String)
    Dim disdata As String
    disdata = Replace(data, "\n", vbCrLf)
    If bLocal = False Then clientsock.SendData disdata
    disdata = Replace(data, "\n", vbCr)
    DisplayLocal disdata
End Sub

Public Sub ClearDisplay()
    If bLocal = False Then clientsock.SendData Chr(27) & "[2J"
    DisplayLocal Chr(27) & "[2J"
    picNorm.Cls
    picBlnk.Cls
End Sub

Public Sub ChangeColor(foreground As sdAnsiFG, Optional background As sdAnsiBG)
    If background = vbEmpty Then background = 40
    If foreground < 100 Then
        If bLocal = False Then clientsock.SendData Chr(27) & "[0;" & foreground & ";" & background & "m"
        DisplayLocal Chr(27) & "[0;" & foreground & ";" & background & "m"
    Else
        If bLocal = False Then clientsock.SendData Chr(27) & "[1;" & (foreground - 100) & ";" & background & "m"
        DisplayLocal Chr(27) & "[1;" & (foreground - 100) & ";" & background & "m"
    End If
End Sub

Public Sub DisplayANSI(path As String)
    On Error GoTo Err_Handler
    Dim sTemp As String, sANSI As String
    
    Open path For Input As #1
    Do While Not EOF(1)
        Line Input #1, sTemp
        sANSI = sANSI & sTemp & vbCrLf
    Loop
    Close #1
    
    Display sANSI

    Exit Sub
Err_Handler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub Quit()
    SetNodeStatus ThisNode, ""
    clientsock.Close
End Sub

Public Function GetInput() As String
    GetInput = InputBuffer
    InputBuffer = ""
End Function

Public Function WaitForKey(Optional WaitKeys As String, Optional Default As String, Optional TimeLimit As Integer)
On Error GoTo Err_Handler

    Dim WaitTime As Long
    If TimeLimit > 0 Then
        WaitTime = Timer + TimeLimit
    Else
        WaitTime = 0
    End If
    On Error GoTo Err_Handler
    InputBuffer = ""
    If WaitKeys = "" Then
        Do While InputBuffer = ""
            If WaitTime > 0 And Timer > WaitTime Then
                WaitForKey = ""
                InputBuffer = ""
                Exit Function
            End If
            DoEvents
            Sleep 1
        Loop
        WaitForKey = InputBuffer
        InputBuffer = ""
        Exit Function
    Else
        Dim keys, key
        keys = Split(WaitKeys, ",")
CheckAgain:
        For Each key In keys
            If InStr(1, UCase(InputBuffer), UCase(CStr(key))) > 0 Then
                WaitForKey = key
                InputBuffer = 0
                Exit Function
            ElseIf InStr(1, InputBuffer, vbCr) > 0 And Default <> "" Then
                WaitForKey = Default
                InputBuffer = 0
                Exit Function
            End If
        Next
        If WaitTime > 0 And Timer > WaitTime Then
            WaitForKey = ""
            InputBuffer = ""
            Exit Function
        End If
        DoEvents
        Sleep 1
        GoTo CheckAgain
    End If

    Exit Function
Err_Handler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub MoveToPos(line_num As Integer, col_num As Integer)
    If bLocal = False Then clientsock.SendData Chr(27) & "[" & line_num & ";" & col_num & "H"
    DisplayLocal Chr(27) & "[" & line_num & ";" & col_num & "H"
End Sub

Public Sub ClearToEndOfLine()
    If bLocal = False Then clientsock.SendData Chr(27) & "[K"
    DisplayLocal Chr(27) & "[K"
End Sub

Public Sub DisplayLocal(data As String)
Dim c, i
Dim tx As Integer, ty As Integer
Dim bEscMode As Boolean
Dim bBold As Boolean
Dim bBlink As Boolean
Dim sCmd As String
Dim Commands() As String
MyTime = Timer
iLines = 0
ScrollTop = 1
ScrollBottom = yScreenSize
bEscMode = False
bBold = False
bBlink = False
bBlinkUsed = False
picBlnk.Visible = False
picNorm.Visible = True
sCmd = ""
    For i = 1 To Len(data)
        If DisplayRate > 0 Then
            TempTime = Timer
            If TempTime < MyTime Then
              DoEvents
              TempTime = Timer
              If (MyTime > TempTime + 0.005) Then
                Sleep 1000 * (MyTime - TempTime)
                DoEvents
              End If
            End If
            If DisplayRate > 0 Then MyTime = MyTime + (10# / DisplayRate)
        End If
        
        c = Mid(data, i, 1)
        If c = Chr(27) Then
            bEscMode = True
        ElseIf bEscMode = True Then
            If c <> "[" Then
                Select Case c
                Case "m"
                    Dim cmd
                    Commands = Split(sCmd, ";")
                    sCmd = ""
                    For Each cmd In Commands
                        If cmd <> "" Then
                            Select Case cmd
                                Case 0
                                    picNorm.ForeColor = QBColor(7)
                                    picNorm.FillColor = QBColor(0)
                                    picBlnk.ForeColor = QBColor(7)
                                    picBlnk.FillColor = QBColor(0)
                                    iLastColor = 7
                                    bBold = False
                                    bBlink = False
                                Case 1
                                    picNorm.ForeColor = GetColor(iLastColor + 8)
                                    picBlnk.ForeColor = GetColor(iLastColor + 8)
                                    bBold = True
                                Case 5
                                    bBlink = True
                                    bBlinkUsed = True
                                Case 30 To 37
                                    If bBold = False Then
                                        picNorm.ForeColor = GetColor(cmd - 30)
                                        picBlnk.ForeColor = GetColor(cmd - 30)
                                    Else
                                        picNorm.ForeColor = GetColor((cmd - 30) + 8)
                                        picBlnk.ForeColor = GetColor(cmd - 30)
                                    End If
                                    iLastColor = cmd - 30
                                Case 40 To 47
                                    picNorm.FillColor = GetColor(cmd - 40)
                                    picBlnk.FillColor = GetColor(cmd - 40)
                            End Select
                        End If
                    Next
                    bEscMode = False
                Case "A"
                    If sCmd = "" Then
                        Y = Y - yOffSet
                    ElseIf IsNumeric(sCmd) = True Then
                        Y = Y - (sCmd * yOffSet)
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "B"
                    If sCmd = "" Then
                        Y = Y + yOffSet
                        iLines = iLines + 1
                    ElseIf IsNumeric(sCmd) = True Then
                        Y = Y + (sCmd * yOffSet)
                        iLines = iLines + sCmd
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "C"
                    If sCmd = "" Then
                        X = X + xOffSet
                    ElseIf IsNumeric(sCmd) = True Then
                        X = X + (sCmd * xOffSet)
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "D"
                    If sCmd = "" Then
                        X = X - xOffSet
                    ElseIf IsNumeric(sCmd) = True Then
                        X = X - (sCmd * xOffSet)
                    End If
                    If X < 0 Then X = 0
                    bEscMode = False
                    sCmd = ""
                Case "L"
                    If sCmd = "" Then sCmd = 1
                    BitBlt picNorm.hDC, 0, Y + (yOffSet * sCmd), xOffSet * xScreenSize, (yOffSet * yScreenSize) * sCmd, picNorm.hDC, 0, Y, vbSrcCopy
                    BitBlt picBlnk.hDC, 0, Y + (yOffSet * sCmd), xOffSet * xScreenSize, (yOffSet * yScreenSize) * sCmd, picBlnk.hDC, 0, Y, vbSrcCopy
                    picNorm.Line (0, Y)-(xOffSet * xScreenSize, (Y + (yOffSet * sCmd)) - 1), picNorm.FillColor, BF
                    picBlnk.Line (0, Y)-(xOffSet * xScreenSize, (Y + (yOffSet * sCmd)) - 1), picBlnk.FillColor, BF
                    bEscMode = False
                    sCmd = ""
                Case "M", "Y"
                Case "H"
                    'sCmd = Replace(sCmd, "(", "")
                    If sCmd <> "" Then
                        Dim l
                        l = InStr(1, sCmd, ";")
                        If l = 0 Then
                            Y = (sCmd - 1) * yOffSet
                            X = 0
                        ElseIf l = 1 Then
                            If sCmd = ";" Then
                                X = 0
                                Y = 0
                            Else
                                X = (Mid(sCmd, 2) - 1) * xOffSet
                                Y = 0
                            End If
                        Else
                            Y = CInt((Mid(sCmd, 1, l - 1)) - 1) * yOffSet
                            X = CInt((Mid(sCmd, l + 1)) - 1) * xOffSet
                        End If
                    Else
                        X = 0
                        Y = 0
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "K"
                    If sCmd = "" Or sCmd = "0" Then
                        picNorm.Line (X, Y)-(xOffSet * xScreenSize, Y + yOffSet - 1), picNorm.FillColor, BF
                        picBlnk.Line (X, Y)-(xOffSet * xScreenSize, Y + yOffSet - 1), picNorm.FillColor, BF
                    ElseIf sCmd = 1 Then
                        picNorm.Line (0, Y)-(X, Y + yOffSet - 1), picNorm.FillColor, BF
                        picBlnk.Line (0, Y)-(X, Y + yOffSet - 1), picNorm.FillColor, BF
                    ElseIf sCmd = 2 Then
                        picNorm.Line (0, Y)-(xOffSet * xScreenSize, Y + yOffSet - 1), picNorm.FillColor, BF
                        picBlnk.Line (0, Y)-(xOffSet * xScreenSize, Y + yOffSet - 1), picNorm.FillColor, BF
                    End If
                Case "u"
                    X = SaveX
                    Y = SaveY
                    bEscMode = False
                    sCmd = ""
                Case "s"
                    SaveX = X
                    SaveY = Y
                    bEscMode = False
                    sCmd = ""
                Case "L"
                    scrollup 1, ScrollTop, ScrollBottom
                    bEscMode = False
                    sCmd = ""
                Case "M"
                    scrolldown 1, ScrollTop, ScrollBottom
                    bEscMode = False
                    sCmd = ""
                Case "J"
                    picNorm.Cls
                    picBlnk.Cls
                    X = 0
                    Y = 0
                    picNorm.CurrentX = 0
                    picNorm.CurrentY = 0
                    picBlnk.CurrentX = 0
                    picBlnk.CurrentY = 0
                    bEscMode = False
                    sCmd = ""
                Case "r"
                    If sCmd = "" Then
                        ScrollTop = 1
                        ScrollBottom = yScreenSize
                    Else
                        l = InStr(1, sCmd, ";")
                        If l = 0 Then
                            ScrollTop = sCmd
                            ScrollBottom = yScreenSize
                        Else
                            ScrollTop = Mid(sCmd, 1, l - 1)
                            ScrollBottom = Mid(sCmd, l + 1)
                        End If
                    End If
                    bEscMode = False
                    sCmd = ""
                Case "l", "h"
                    bEscMode = False
                    sCmd = ""
                Case Is > "?"
                    'MsgBox "unrecognized command: " & sCmd & c
                    bEscMode = False
                    sCmd = ""
                Case Else
                    sCmd = sCmd & c
                End Select
            End If
        ElseIf c = Chr(13) Or c = Chr(10) Then
            If c = Chr(13) Then
                picNorm.CurrentX = X
                picNorm.CurrentY = Y
                picBlnk.CurrentX = X
                picBlnk.CurrentY = Y
                X = 0
                Y = Y + yOffSet
                iLines = iLines + 1
                If Y > ((yOffSet * ScrollBottom) - yOffSet) Then
                    scrollup 1, ScrollTop, ScrollBottom
                    Y = yOffSet * (ScrollBottom - 1)
                End If
            End If
        ElseIf c = Chr(8) Then
            If X > 0 Then
                X = X - xOffSet
                picNorm.Line (X, Y)-(X + (xOffSet - 1), Y + (yOffSet - 1)), picNorm.FillColor, BF
                picNorm.CurrentX = X
                picNorm.CurrentY = Y
                picNorm.Print " "
                
                picBlnk.Line (X, Y)-(X + (xOffSet - 1), Y + (yOffSet - 1)), picBlnk.FillColor, BF
                picBlnk.CurrentX = X
                picBlnk.CurrentY = Y
                picBlnk.ForeColor = picNorm.ForeColor
                If bBlink = False Then picBlnk.Print " "
            End If
        Else
        
            If X >= (xOffSet * xScreenSize) Then
                X = 0
                Y = Y + yOffSet
                iLines = iLines + 1
            End If
            
            If Y > ((yOffSet * ScrollBottom) - yOffSet) Then
                scrollup 1, ScrollTop, ScrollBottom
                Y = yOffSet * (ScrollBottom - 1)
            End If
            
            picNorm.Line (X, Y)-(X + (xOffSet - 1), Y + (yOffSet - 1)), picNorm.FillColor, BF
            picNorm.CurrentX = X
            picNorm.CurrentY = Y
            picNorm.Print c
            
            picBlnk.Line (X, Y)-(X + (xOffSet - 1), Y + (yOffSet - 1)), picBlnk.FillColor, BF
            picBlnk.CurrentX = X
            picBlnk.CurrentY = Y
            picBlnk.ForeColor = picNorm.ForeColor
            If bBlink = False Then picBlnk.Print c
            
            X = X + xOffSet
            
            tx = X / xOffSet
            ty = Y / yOffSet
        End If
    Next i
    If bBlinkUsed = True Then tmrBlink.Enabled = True
End Sub

Private Function GetColor(clr As Integer) As Long
    Select Case clr
        Case 0
            GetColor = QBColor(0)
        Case 1
            GetColor = QBColor(4)
        Case 2
            GetColor = QBColor(2)
        Case 3
            GetColor = QBColor(6)
        Case 4
            GetColor = QBColor(1)
        Case 5
            GetColor = QBColor(5)
        Case 6
            GetColor = QBColor(3)
        Case 7
            GetColor = QBColor(7)
        Case 8
            GetColor = QBColor(8)
        Case 9
            GetColor = QBColor(12)
        Case 10
            GetColor = QBColor(10)
        Case 11
            GetColor = QBColor(14)
        Case 12
            GetColor = QBColor(9)
        Case 13
            GetColor = QBColor(13)
        Case 14
            GetColor = QBColor(11)
        Case 15
            GetColor = QBColor(15)
    End Select
End Function

Public Sub scrollup(numlines As Integer, Top As Integer, bot As Integer)
    If numlines >= bot - Top + 1 Then
        ' Just clear from top to bottom
        picNorm.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (bot)), picNorm.FillColor, BF
        picBlnk.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (bot)), picBlnk.FillColor, BF
    Else
        ' Copy bot-top+1 - numlines lines from (top+numlines) to (top)
        BitBlt picNorm.hDC, 0, (Top - 1) * yOffSet, xOffSet * xScreenSize, _
          yOffSet * (bot - Top + 1 - numlines), picNorm.hDC, 0, yOffSet * (Top + numlines - 1), vbSrcCopy
        BitBlt picBlnk.hDC, 0, (Top - 1) * yOffSet, xOffSet * xScreenSize, _
          yOffSet * (bot - Top + 1 - numlines), picBlnk.hDC, 0, yOffSet * (Top + numlines - 1), vbSrcCopy
        ' Erase lines bot-numlines to bot
        picNorm.Line (0, yOffSet * (bot - numlines))-(xOffSet * xScreenSize, yOffSet * (bot)), picNorm.FillColor, BF
        picBlnk.Line (0, yOffSet * (bot - numlines))-(xOffSet * xScreenSize, yOffSet * (bot)), picBlnk.FillColor, BF
    End If
End Sub

Public Sub scrolldown(numlines As Integer, Top As Integer, bot As Integer)
    If numlines >= bot - Top + 1 Then
        ' Just clear from top to bottom
        picNorm.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (bot)), picNorm.FillColor, BF
        picBlnk.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (bot)), picBlnk.FillColor, BF
    Else
        ' Copy bot-top+1 - numlines lines from (top+numlines) to (top)
        BitBlt picNorm.hDC, 0, yOffSet * (Top + numlines - 1), xOffSet * xScreenSize, _
          yOffSet * (bot - Top + 1 - numlines), picNorm.hDC, 0, (Top - 1) * yOffSet, vbSrcCopy
        BitBlt picBlnk.hDC, 0, yOffSet * (Top + numlines - 1), xOffSet * xScreenSize, _
          yOffSet * (bot - Top + 1 - numlines), picBlnk.hDC, 0, (Top - 1) * yOffSet, vbSrcCopy
        ' Erase lines bot-numlines to bot
        picNorm.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (Top + numlines - 1)), picNorm.FillColor, BF
        picBlnk.Line (0, yOffSet * (Top - 1))-(xOffSet * xScreenSize, yOffSet * (Top + numlines - 1)), picBlnk.FillColor, BF
    End If
End Sub

Public Function GetUserInput()
    InputBuffer = ""
    Do
        If Right(InputBuffer, 1) = Chr(8) Then
            If Len(InputBuffer) = 1 Then
                InputBuffer = ""
            Else
                InputBuffer = Left(InputBuffer, Len(InputBuffer) - 2)
            End If
        End If
        If InStr(1, InputBuffer, vbCr) > 0 Then
            InputBuffer = Replace(InputBuffer, vbCr, "")
            InputBuffer = Replace(InputBuffer, vbLf, "")
            GetUserInput = InputBuffer
            Exit Function
        End If
        DoEvents
    Loop
End Function
