Attribute VB_Name = "ProjectTracker"
Option Explicit

Sub PT_Build()
    ' build the ProjectTracker worksheet
    PTL_StoreHeaders
    PTL_SetTargetDate
    PT_InsertSheet
    PT_CopyData
    PT_FormatHeadings
    PT_FormatColumns
    PT_DrawBorders
    PT_InsertTargetInfo
End Sub

Sub PT_InsertSheet()
    ' subroutine that creates the Project Tracking worksheet to be inserted into the MAP
    ' if the sheet name already exists, that current sheet gets deleted first
    With ThisWorkbook
        If SheetExists("ProjectTracker") Then
            Application.DisplayAlerts = False
            Worksheets("ProjectTracker").Delete
            Application.DisplayAlerts = True
            .Sheets.Add After:=Worksheets(Worksheets.count)
            .Sheets(.Sheets.count).Name = "ProjectTracker"
        Else
            .Sheets.Add After:=Worksheets(Worksheets.count)
            .Sheets(.Sheets.count).Name = "ProjectTracker"
        End If
    End With
End Sub

Sub PT_CopyData()
    Dim myRange As Range ' working range
    
    ' copy the range that references all of the project timelines
    ' column A holds the heading value from Word
    Set myRange = Worksheets("ProjectTimeline").Range("A1:B" & [PTL_Rows])
    myRange.Copy Destination:=Sheets("ProjectTracker").Range("A1")
    Sheets("ProjectTracker").Select
End Sub

Sub PT_FormatHeadings()
    ' Copy color and formats from main column to each projects header row
    ' Copy header labels into each projects header row
    
    ' Iterate through column A to find all instances of '2' (header rows)
    ' For each match, copy color and labels
    Dim Hdr As Variant
    Dim i As Integer
    Dim myRange As Range
    Dim destRange As Range
    
    Hdr = Array("Person Responsible", "Target Date", "Complete", "Notes")
    
    For i = 1 To [PTL_Rows]
        Set myRange = Range("A" & i)
        myRange.Select
        If myRange.Value = 2 Then
            With myRange
                .Offset(0, 1).Copy
                .Offset(0, 2).PasteSpecial (xlPasteFormats)
                .Offset(0, 3).PasteSpecial (xlPasteFormats)
                .Offset(0, 4).PasteSpecial (xlPasteFormats)
                .Offset(0, 5).PasteSpecial (xlPasteFormats)
                .Offset(0, 2).Resize(, 4).Value = Hdr
            End With
        End If
    Next i

End Sub

Sub PT_FormatColumns()
    ' format headings
    With Range("C:F")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 10
    End With
    
    ' set column widths
    Range("A1").ColumnWidth = 6
    Range("B1").ColumnWidth = 25
    Range("C1").ColumnWidth = 10
    Range("D1").ColumnWidth = 10
    Range("E1").ColumnWidth = 7.33
    Range("F1").ColumnWidth = 18.33
End Sub

Sub PT_DrawBorders()
    ' draw borders around the table area
    Dim myRange As Range
    
    Set myRange = Range("B1:F" & [PTL_Rows])
    myRange.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlColorIndexAutomatic
    With myRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With myRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    
End Sub

Sub PT_Export()
    ' export table to Word
    Dim wdApp As Word.Application ' early binding variable to ms word application
    Dim wdDoc As Word.Document ' early binding variable to ms word document
    Dim wdRng As Word.Range ' early binding variable to ms word range
    Dim wdPar As Word.Paragraph ' early binding variable to ms word paragraph
    Dim myRange As Excel.Range ' working range for this procedure
    Dim myFile As String
    Dim f As Boolean
    
    myFile = Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - 5) & ".docm"
    
    ' check globals
    ' set the row to -1 to match the spacing from Project Timeline
    ' otherwise a row would be missing when pasting
    If [PTL_Rows] = 0 Then
        [PTL_Rows] = Sheets("ProjectTracker").Cells(Rows.count, 1).End(xlUp).Row - 1
    End If
    
    ' set the excel range to copy
    Set myRange = Sheets("ProjectTracker").Range("B1:F" & [PTL_Rows] + 1)
    myRange.Copy
    
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
    
    Set wdRng = wdDoc.Bookmarks("Project_Tracking_Chart").Range
    
    ' check if there is text in the selection
    If wdRng.Text <> "" Then
        wdRng.Text = ""
    End If
    
    ' paste the link into word
    With wdRng
        .Collapse Direction:=wdCollapseStart
        .PasteExcelTable LinkedToExcel:=True, WordFormatting:=False, RTF:=False
    End With
    
    'Clear out the clipboard, and turn screen updating back on.
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
    End With
    
    MsgBox "The Project Tracking Chart has been successfully " & vbNewLine & _
           "transferred to " & myFile, vbInformation
    
    'Save the Word doc.
    wdDoc.Save
    wdDoc.Activate
    
    'Null out your variables.
    Set wdRng = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
End Sub

Sub PT_InsertTargetInfo()
    ' pull info from PTL_TargetDate into Project Tracker
    Dim i As Integer ' count through the rows on the sheet
    Dim n As Integer ' count through the columns in a row
    Dim aTarget() As Variant  ' array to hold the target date for each goal
    
    On Error Resume Next
    
    aTarget = [PTL_TargetDate]
    
    ' iterate through the timeline to find the targets
    For i = 1 To [PTL_Rows]
        ' only interested in Goal rows
        If Range("A" & i).Value = 3 Then
            ' pull array data to fill into table
            If aTarget(i) = "Complete" Then
                Range("E" & i).Value = "X"
            Else
                Range("D" & i).Value = aTarget(i)
            End If
        End If
    Next i

End Sub


