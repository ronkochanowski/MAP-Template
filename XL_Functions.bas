Attribute VB_Name = "Functions"
'-----------------------------------------------------------------
Public Function ColumnLetter(Col As Long)
'-----------------------------------------------------------------
    ' returns the column letter for the passed column number
    Dim sColumn As String
    On Error Resume Next
    sColumn = Split(Columns(Col).Address(, False), ":")(1)
    On Error GoTo 0
    ColumnLetter = sColumn
End Function

'-----------------------------------------------------------------
Function LastInColumn(rng As Range)
'-----------------------------------------------------------------
    ' returns the contents of the last non-empty cell in a column
    Dim LastCell As Range
    Application.Volatile
    With rng.Parent
        With .Cells(.Rows.count, rng.Column)
            If Not IsEmpty(.Value) Then
                LastInColumn = .Value
            ElseIf IsEmpty(.End(xlUp)) Then
                LastInColumn = ""
            Else
                LastInColumn = .End(xlUp).Value
            End If
        End With
    End With
End Function

'-----------------------------------------------------------------
Function LastInRow(rng As Range)
'-----------------------------------------------------------------
    ' returns the contents of the last non-empty cell in a row
    Dim LastCell As Range
    Application.Volatile
    With rng.Parent
        With .Cells(rng.Row, .Columns.count)
            If Not IsEmpty(.Value) Then
                LastInRow = .Value
            ElseIf IsEmpty(.End(xlToLeft)) Then
                LastInRow = ""
            Else
                LastInRow = .End(xlToLeft).Value
            End If
        End With
    End With
End Function

'-----------------------------------------------------------------
Function NameExists(FindName As String) As Boolean
'-----------------------------------------------------------------
    Dim rng As Range
    Dim myName As String
    
    On Error Resume Next
    
    myName = ActiveWorkbook.Names(FindName).Name
    If Err.Number = 0 Then
        NameExists = True
    Else
        NameExists = False
    End If
End Function

'-----------------------------------------------------------------
Public Function IsAppRunning(ByVal sAppName) As Boolean
'-----------------------------------------------------------------
' checks to see if the application passed is running.  the application is in the format
' AppName.Application, as in, "Word.Application"

    Dim oApp As Object
    On Error Resume Next
    Set oApp = GetObject(, sAppName)
    If Not oApp Is Nothing Then
        Set oApp = Nothing
        IsAppRunning = True
    End If

End Function

'-----------------------------------------------------------------
Public Function GetRows() As Long
'-----------------------------------------------------------------
    GetRows = Worksheets("ProjectTimeline").Cells(Rows.count, 1).End(xlUp).Row

End Function

'-----------------------------------------------------------------
Public Function GetCols() As Long
'-----------------------------------------------------------------
    GetCols = Worksheets("ProjectTimeline").Cells(1, Columns.count).End(xlToLeft).Column

End Function

'-----------------------------------------------------------------
Public Function SheetExists(sname) As Boolean
'-----------------------------------------------------------------
    ' returns TRUE if sheet exists in the active workbook
    Dim x As Object
    On Error Resume Next
    Set x = ActiveWorkbook.Sheets(sname)
    
    If Err.Number = 0 Then
        SheetExists = True
    Else
        SheetExists = False
    End If
    
End Function

