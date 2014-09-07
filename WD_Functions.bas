Attribute VB_Name = "Functions"
Option Explicit

Public Function File_Exists(ByVal sPathName As String, Optional Directory As Boolean) As Boolean
 
 'Read more at http://vbadud.blogspot.com/2007/04/vba-function-to-check-file-existence.html#y3dfrq8lAIHeTSGA.99
 
 'Returns True if the passed sPathName exist
 'Otherwise returns False
 
 On Error Resume Next
 If sPathName <> "" Then

    If IsMissing(Directory) Or Directory = False Then

        File_Exists = (Dir$(sPathName) <> "")
    Else

        File_Exists = (Dir$(sPathName, vbDirectory) <> "")
    End If

 End If
End Function

Public Function IsAppRunning(ByVal sAppName) As Boolean
    
'checks to see if the application passed is running.  the application is in the format
' AppName.Application, as in, "Excel.Application"

    Dim oApp As Object
    On Error Resume Next
    Set oApp = GetObject(, sAppName)
    If Not oApp Is Nothing Then
        Set oApp = Nothing
        IsAppRunning = True
    End If

End Function

Public Function xlIsDirty(ByVal ws As Excel.Worksheet) As Boolean
    
    'checks to see if the worksheet is dirty by looking at the range "A2:B3000"
    If ws.Application.WorksheetFunction.CountA(Range("A2:B3000")) = 0 Then
        xlIsDirty = False
    Else
        xlIsDirty = True
    End If
End Function
