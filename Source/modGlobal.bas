Attribute VB_Name = "modGlobal"
Option Base 0
Public NodeStatus(128) As String

Public Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)

Public Declare Function BitBlt Lib "GDI32" ( _
   ByVal hDCDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
   As Long

Public Function Pad_String(sPadString As String, iNumChar As Integer, sPadChar As String, sAlignment0L1R As Integer)
    'PAD A STRING WITH A SPECIFIED CHARACTER, SO THAT IT RETURNS AS A SPECIFIED LENGTH
    Dim iLen As Integer
    iLen = Len(sPadString)
    If iLen >= iNumChar Then
        Pad_String = Left(sPadString, iNumChar)
    Else
        Do Until Len(sPadString) >= iNumChar
            If sAlignment0L1R = 0 Then
                sPadString = sPadChar & sPadString
            Else
                sPadString = sPadString & sPadChar
            End If
        Loop
        Pad_String = sPadString
    End If
End Function

