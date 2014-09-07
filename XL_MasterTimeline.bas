Attribute VB_Name = "MasterTimeline"
Option Explicit

Sub MTLfromPTL()
    ' copy project timelines to master timeline and format
    MTL_InsertSheet
    MTL_CopyPTL
    MTL_DeleteHeaders
    MTL_InsertSummary
End Sub

Sub MTL_InsertSheet()
    ' subroutine that creates the Project Tracking worksheet to be inserted into the MAP
    With ThisWorkbook
        .Sheets.Add After:=Worksheets(Worksheets.count)
        .Sheets(.Sheets.count).Name = "MasterTimeline"
    End With
End Sub

Sub MTL_CopyPTL()
    ' copy the project timelines worksheet to master timeline worksheet
    Sheets("ProjectTimeline").Cells.Copy Destination:=Sheets("MasterTimeline").Cells
    Sheets("MasterTimeline").Select
End Sub

Sub MTL_DeleteHeaders()
    Dim cell As Range ' iteration range
    Dim myRng As Range ' working range
    
    ' myRng set to column A size of table plus the number of header 2 values. This
    ' to correct an issue with deleting rows and shrinking the size of myRng
    Set myRng = Sheets("MasterTimeline").Range("A1:A" & [PTL_Rows] + [PTL_Hd2])
        
    ' remove empty rows
    For Each cell In myRng
        With cell
            If .Value = "" Then
                .EntireRow.Delete
            End If
        End With
    Next

End Sub

Sub MTL_InsertSummary()
    Dim ProjHeads() As Variant ' variant to hold the individual project information
    Dim cell As Range   ' interation range
    Dim i As Integer    ' iteration counter
    Dim selRng As Range ' range of the selected area
    Dim sumRng As Range ' range of the summary area of the timeline
    Dim x As Integer    ' value of the color index in the current iteration
    Dim y As Double     ' value of the TintandShade of the ColorIndex
    Dim sCell As String ' starting cell of the range
    Dim eCell As String ' ending cell of the range
    
    ' set variables
    i = 0
    y = 0.75
    ActiveCell.CurrentRegion.Select
    Set selRng = Range(Selection.Address(rowabsolute:=False, columnabsolute:=False))
    
    ' ProjHeads elements:
    '   1=Header
    '   2=ColorIndex
    '   3=start range
    '   4=end range
    '   5=least column number of "X"
    '   6=greatest column number of "X"
    '   7=row placeholder
    ReDim ProjHeads([PTL_Hd2], 7)
    
    ' iterate thru selection to find headers, add them to the array
    For Each cell In Selection
        With cell
            If .Value = 2 And .Column = 1 Then
                i = i + 1
                x = .Interior.ColorIndex
                ProjHeads(i, 1) = .Offset(0, 1).Value
                ProjHeads(i, 2) = .Offset(0, 1).Interior.ColorIndex
                ProjHeads(i, 3) = "B" & .Row
                ProjHeads(i, 4) = ColumnLetter(selRng.Columns.count) & .Row
            ElseIf .Column > 2 And i > 0 Then
                ProjHeads(i, 7) = .Row
            End If
        End With
    Next
    
    ' using ProjHeads array, determine elements 5 and 6 from sort columns
    ' sort columns are set with the offset(0,2) which sets the first timeframe column number
    ' offset(0,3) sets the second sort key which represents the end timeframe column number
    ' min and max values are stored in the array to build the summary timeline
    For i = 1 To [PTL_Hd2]
        sCell = ColumnLetter(Range(ProjHeads(i, 4)).Offset(0, 2).Column) & Right(ProjHeads(i, 3), Len(ProjHeads(i, 3)) - 1)
        eCell = ColumnLetter(Range(ProjHeads(i, 4)).Offset(0, 3).Column) & ProjHeads(i, 7)
        ProjHeads(i, 5) = Application.WorksheetFunction.Min(Range(sCell, eCell))
        ProjHeads(i, 6) = Application.WorksheetFunction.Max(Range(sCell, eCell))
    Next i
    
    ' insert header2 names with a blank row at top of timeline to identify summary area
    Range("A2:B" & [PTL_Hd2] + 2).EntireRow.Insert (xlShiftDown)
    
    ' clean up the summary area
    Range("A1:B1").Clear
    Range("1:1").Interior.ColorIndex = 0
    Range("C2:" & ColumnLetter([PTL_Cols]) & [PTL_Hd2] + 1).Interior.ColorIndex = 0
    Range([PTL_Hd2] + 2 & ":" & [PTL_Hd2] + 2).Interior.ColorIndex = 0
    
    i = 0
    
    ' format each row to match it's original color from the project timeline
    ' PTL_Hd2 + 1 because we're starting at row 2
    For Each cell In Range("B2:B" & [PTL_Hd2] + 1)
        i = i + 1
        With cell
            .Value = ProjHeads(i, 1)
            .HorizontalAlignment = xlLeft
            .WrapText = True
            With .Font
                .Size = 12
                .Bold = True
            End With
        End With
    Next
    
    ' color the summary backgrounds
    i = 1
    For Each cell In Range("B2:B" & UBound(ProjHeads) + 1)
        With cell.Interior
            .ColorIndex = ProjHeads(i, 2)
            .TintAndShade = y
        End With
        i = i + 1
    Next
    
    ' build out summary timelines
    i = 0
    For Each cell In Range("B2:" & ColumnLetter([PTL_Cols]) & [PTL_Hd2] + 1)
        With cell
            If .Column = 2 Then
                i = i + 1
            ElseIf .Column >= ProjHeads(i, 5) And .Column <= ProjHeads(i, 6) Then
                .Interior.ColorIndex = ProjHeads(i, 2)
                .Interior.TintAndShade = y + 0.2
                .Value = "X"
                .Orientation = xlHorizontal
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End If
        End With
    Next
    
    ' add borders to summary area
    Set sumRng = Range("B1:" & Left(ProjHeads(1, 4), 1) & Range("B1").End(xlDown).Row)
    sumRng.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlColorIndexAutomatic
    With sumRng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With sumRng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    
    ' remove heading 2 rows from detail table
    For Each cell In Range("A:A")
        If cell.Value = 2 Then cell.EntireRow.Delete
    Next

    ' sort detail table
    Set selRng = selRng.Resize(selRng.Rows.count - ([PTL_Hd2] + 2), selRng.Columns.count + 3).Offset([PTL_Hd2] + 2, 0)
    With selRng
        .Sort Key1:=.Cells([PTL_Hd2] + 3, .Columns.count - 1), key2:=.Cells([PTL_Hd2] + 3, .Columns.count), Header:=xlNo
    End With
    
    'set the current cell to A1 when conversion complete
    Range("A1").Select
    
End Sub

Sub MTL_Export()
    Dim wdApp As Word.Application ' early binding variable to ms word application
    Dim wdDoc As Word.Document ' early binding variable to ms word document
    Dim wdRng As Word.Range ' early binding variable to ms word range
    Dim wdPar As Word.Paragraph ' early binding variable to ms word paragraph
    Dim myRange As Excel.Range ' working range for this procedure
    Dim myFile As String ' holds the path and file string of the document
    Dim eCol As Long ' ending column of the range to be copied
    Dim eRow As Long ' ending row of the range to be copied
    Dim sCell As String ' starting cell of the range to be copied
    Dim eCell As String ' ending cell of the range to be copied
    Dim f As Boolean
    
    ' set the excel range to copy
    eCol = Worksheets("MasterTimeline").Cells(1, Columns.count).End(xlToLeft).Column
    eRow = Worksheets("MasterTimeline").Cells(Rows.count, 1).End(xlUp).Row
    sCell = "B1"
    eCell = ColumnLetter(eCol) & eRow
    Set myRange = Range(sCell & ":" & eCell)
    myRange.Copy
    
    ' pointer to word docm file, word file has same name - different extension
    myFile = Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - 5) & ".docm"
    
    ' try to open the map file
    Set wdDoc = GetObject(myFile)
    If wdDoc Is Nothing Then
        Set wdApp = GetObject(, "Word.Application")
        If wdApp Is Nothing Then
            Set wdApp = CreateObject("Word.Application")
            If wdApp Is Nothing Then
                MsgBox "Failed to start Word!", vbCritical
                Exit Sub
            End If
            f = True
        End If
        Set wdDoc = wdApp.Documents.Open(myFile)
        If wdDoc Is Nothing Then
            MsgBox "Failed to open help document!", vbCritical
            If f Then
                wdApp.Quit
            End If
            Exit Sub
        End If
        wdApp.Visible = True
    Else
        With wdDoc.Parent
            .Visible = True
            .Activate
        End With
    End If
    
    Set wdRng = wdDoc.Bookmarks("Master_Timeline").Range
    With wdRng
        .Select
        .Collapse Direction:=wdCollapseEnd
        .PasteExcelTable LinkedToExcel:=True, WordFormatting:=False, RTF:=False
    End With
    
    'Clear out the clipboard, and turn screen updating back on.
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
    End With
    
    MsgBox "The MasterTimeline has been successfully " & vbNewLine & _
           "transferred to " & myFile, vbInformation
    
    'Save and show the Word doc.
    wdDoc.Save
    wdDoc.Activate
    
    'Null out your variables.
    Set wdRng = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
End Sub
