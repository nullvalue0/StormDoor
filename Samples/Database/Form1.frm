VERSION 5.00
Object = "{84AF4DF3-4B59-4D87-85BA-FA878460F831}#6.4#0"; "StormDoor.ocx"
Begin VB.Form Form1 
   Caption         =   "Database Example"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin StormDoorX.StormDoor StormDoor1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9128
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'these 2 lines need to be in the declarations area, so the
'database can be available to all sections of the code
Dim cn As ADODB.Connection  'create the database connection object
Dim rs As ADODB.Recordset   'create the database recordset object, this holds the actual data

Dim Menu As String          'holds which menu we're sitting at

Dim sInputBuffer As String  'holds the input buffer

Dim FirstName As String     'holds the first name temporarily while adding a new record
Dim LastName As String      'holds the last name temporarily while adding a new record
Dim PhoneNumber As String   'holds the phone temporarily while adding a new record


Private Sub Form_Load()
    'create a new instance of the database objects
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    'open the connection to the database file
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database.mdb"
    
    'initialize stormdoor
    StormDoor1.OpenDropFile Command
    StormDoor1.ClearDisplay
    
    DisplayMain
End Sub


Sub DisplayMain()
    Dim i As Integer
    StormDoor1.ClearDisplay
    
    'make sure we close the recordset is closed before we try to open it
    'or else we will get an error
    If rs.State = adStateOpen Then cn.Close
    'returns the rows matching this SQL statement
    rs.Open "SELECT * FROM TestTable ORDER BY RecordID", cn, adOpenStatic
    
    'show how many records we will display
    StormDoor1.MoveToPos 1, 2
    StormDoor1.ChangeColor fYellow, bBlack
    StormDoor1.Display "Displaying " & rs.RecordCount & " records:"
    
    'loop through all the records, displaying the 4 data fields horizontally
    For i = 1 To rs.RecordCount
        StormDoor1.ChangeColor fGreen, bBlack
        StormDoor1.MoveToPos i + 2, 5
        StormDoor1.Display rs.Fields("RecordID")
        
        'use IgnoreNulls() function around each textual field name
        'when displaying to ensure we don't get any errors
        
        StormDoor1.ChangeColor fLightBlue, bBlack
        StormDoor1.MoveToPos i + 2, 20
        StormDoor1.Display IgnoreNulls(rs.Fields("FirstName"))
        
        StormDoor1.ChangeColor fLighCyan, bBlack
        StormDoor1.MoveToPos i + 2, 40
        StormDoor1.Display IgnoreNulls(rs.Fields("LastName"))
        
        StormDoor1.ChangeColor fLightRed, bBlack
        StormDoor1.MoveToPos i + 2, 60
        StormDoor1.Display IgnoreNulls(rs.Fields("PhoneNumber"))
        
        'after displaying this record, move onto the next one
        rs.MoveNext
    Next i
    
    StormDoor1.Display "\n\n\n"
    StormDoor1.ChangeColor fWhite, bBlack
    StormDoor1.Display "(A) Add a record\n(D) Delete a record\n(Q) Quit\n\nCommand? "

    Menu = "Main"
    sInputBuffer = ""
End Sub

Private Sub StormDoor1_UserInput(data As String)
    Select Case Menu
        Case "Main"
            Select Case UCase(data)
                Case "A"
                    StormDoor1.Display "\n\nAdding a Record\nEnter the First Name: "
                    Menu = "AddNewFirstName"
                    sInputBuffer = ""
                Case "D"
                    StormDoor1.Display "\n\nEnter the RecordID for the record to delete: "
                    Menu = "Delete"
                Case "Q"
                    StormDoor1.Quit
                    End
            End Select
        Case "Delete"
            If InStr(1, data, vbCr) > 0 Then
                If rs.State = adStateOpen Then rs.Close
                If IsNumeric(sInputBuffer) Then
                    'Delete the desired record from the database
                    rs.Open "DELETE FROM TestTable WHERE RecordID = " & sInputBuffer, cn
                End If
                DisplayMain
            ElseIf data = Chr(8) Then
                If sInputBuffer <> "" Then
                    StormDoor1.Display Chr(27) & "[D " & Chr(27) & "[D"
                    sInputBuffer = Left(sInputBuffer, Len(sInputBuffer) - 1)
                End If
            Else
                If Len(sInputBuffer) < 3 Then
                    sInputBuffer = sInputBuffer & data
                    StormDoor1.Display data
                End If
            End If
        Case "AddNewFirstName"
            If InStr(1, data, vbCr) > 0 Then
                FirstName = sInputBuffer
                sInputBuffer = ""
                StormDoor1.Display "\nEnter the Last Name: "
                Menu = "AddNewLastName"
            ElseIf data = Chr(8) Then
                If sInputBuffer <> "" Then
                    StormDoor1.Display Chr(27) & "[D " & Chr(27) & "[D"
                    sInputBuffer = Left(sInputBuffer, Len(sInputBuffer) - 1)
                End If
            Else
                If Len(sInputBuffer) < 50 Then
                    sInputBuffer = sInputBuffer & data
                    StormDoor1.Display data
                End If
            End If
        Case "AddNewLastName"
            If InStr(1, data, vbCr) > 0 Then
                LastName = sInputBuffer
                sInputBuffer = ""
                StormDoor1.Display "\nEnter the Phone Number: "
                Menu = "AddNewPhone"
            ElseIf data = Chr(8) Then
                If sInputBuffer <> "" Then
                    StormDoor1.Display Chr(27) & "[D " & Chr(27) & "[D"
                    sInputBuffer = Left(sInputBuffer, Len(sInputBuffer) - 1)
                End If
            Else
                If Len(sInputBuffer) < 50 Then
                    sInputBuffer = sInputBuffer & data
                    StormDoor1.Display data
                End If
            End If
        Case "AddNewPhone"
            If InStr(1, data, vbCr) > 0 Then
                PhoneNumber = sInputBuffer
                sInputBuffer = ""
                
                If rs.State = adStateOpen Then rs.Close
                'Add this new record to the database
                '(Use the FixQuotes function around each inserted field value to ensure proper SQL conventions)
                rs.Open "INSERT INTO TestTable (FirstName,LastName,PhoneNumber) VALUES ('" & FixQuotes(FirstName) & "','" & FixQuotes(LastName) & "','" & FixQuotes(PhoneNumber) & "')", cn
                DisplayMain
            ElseIf data = Chr(8) Then
                If sInputBuffer <> "" Then
                    StormDoor1.Display Chr(27) & "[D " & Chr(27) & "[D"
                    sInputBuffer = Left(sInputBuffer, Len(sInputBuffer) - 1)
                End If
            Else
                If Len(sInputBuffer) < 50 Then
                    sInputBuffer = sInputBuffer & data
                    StormDoor1.Display data
                End If
            End If
    End Select
    
End Sub

Function IgnoreNulls(s)
    If IsNull(s) Then
        IgnoreNulls = ""
    Else
        IgnoreNulls = s
    End If
End Function

Function FixQuotes(s)
    FixQuotes = Replace(s, "'", "''")
End Function
