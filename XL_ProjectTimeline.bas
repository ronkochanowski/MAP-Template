Attribute VB_Name = "ProjectTimeline"
Option Explicit
Option Base 1

Sub PTLFormat()
    ' steps to format the timeline
    PTLInitializeGlobals
    PTLCleanUp
    PTLFormatRowHeaders
    PTLOrientColumnHeader
    PTLFormatColumnHeader
    PTLBorders
    PTL_StoreHeaders
    PTLInsertHeaders
    PTLCreateRefArray
    PTLInitializeGlobals
    PTLColorize
    PTLRemoveHeader
    PTLSetChange
End Sub

Sub PTLCleanUp()
    ' delete last row brought in from Word if it is "#N/A"
    With Cells(Rows.count, 1).End(xlUp)
        If WorksheetFunction.IsNA(.Value) Then
            .EntireRow.Delete
        End If
    End With
End Sub

Sub PTLFormatRowHeaders()
    Dim i As Integer ' iteration counter
    
    ' move through each row and format objectives and tasks
    With Sheets("ProjectTimeline")
        Cells(1, 1).Select
        For i = 1 To [PTL_Rows]
            With ActiveCell
                Select Case .Value
                    Case 1
                        With .Offset(0, 1)
                            .HorizontalAlignment = xlCenter
                            With .Font
                                .Size = 12
                                .Bold = True
                            End With
                        End With
                    Case 2
                        With .Offset(0, 1)
                            .HorizontalAlignment = xlLeft
                            .WrapText = True
                            With .Font
                                .Size = 12
                                .Bold = True
                            End With
                        End With
                    Case 3
                        With .Offset(0, 1)
                            .HorizontalAlignment = xlLeft
                            .WrapText = True
                            With .Font
                                .Size = 10
                                .Bold = False
                            End With
                        End With
                    Case Else
                End Select
            .Offset(1, 0).Select
            End With
        Next i
    End With
End Sub

Sub PTLOrientColumnHeader()
    Dim i As Integer ' iteration counter
    
    ' format top row headings inserting timeframes
    Sheets("ProjectTimeline").Range("A1").Select
    For i = 2 To [PTL_Cols]
        With ActiveCell.Offset(0, i)
            .Orientation = 90
            .ColumnWidth = 4
        End With
    Next i
End Sub

Sub PTLFormatColumnHeader()
    ' format header
    With Sheets("ProjectTimeline").Range("A1").EntireRow
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .WrapText = True
    End With
End Sub
    
Sub PTLBorders()
    Dim myRange As Range ' working range for this procedure
    
    Sheets("ProjectTimeline").Select
    
    ' set range of table and place borders
    Set myRange = Range("B1:" & ColumnLetter([PTL_Cols]) & [PTL_Rows])
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

Sub PTL_StoreHeaders()
    ' store the array to PTL_Hd1 for future use
    ' this combined with storing the timeframe indiacted
    ' will be used on the ProjectTracker sheet to pre-fill
    ' timings on that sheet
    Dim i As Integer
    Dim aHeaders() As Variant
    
    ReDim aHeaders([PTL_Cols])
    Sheets("ProjectTimeline").Activate
    
    ' cycle through the headers adding them to the array
    For i = 3 To [PTL_Cols]
        aHeaders(i) = ActiveSheet.Cells(1, i)
    Next i
    
    ' save the working array into the save array PTL_Hd1
    Names.Add Name:="PTL_Hd1", RefersTo:=aHeaders

End Sub

Sub PTLInsertHeaders()
    Dim myRange As Range ' working range for this procedure
    Dim i As Integer ' iteration counter
    
    ' set and copy the header row
    Set myRange = Range("C1:" & ColumnLetter([PTL_Cols]) & "1")
        
    ' move down thru each row and insert header row at heading 2
    With Sheets("ProjectTimeline")
        .Range("A1").Select
        For i = 1 To [PTL_Rows]
            Select Case ActiveCell.Value
                Case 2
                    With ActiveCell
                        .EntireRow.Insert
                        myRange.Copy
                        ActiveSheet.Paste Destination:=Range("C" & ActiveCell.Row).Offset(1, 0)
                        .Offset(1, 0).Select
                    End With
                Case Else
            End Select
            ' move to next row
            ActiveCell.Offset(1, 0).Select
        Next i
    End With
End Sub

Sub PTLColorize()
    Dim i As Integer    ' iteration counter
    Dim x As Integer    ' colorindex value
    Dim y As Double     ' TintandShade value
    Dim sRow As Long    ' holds the value of the starting row number of the range
    Dim eRow As Long    ' holds the value of the ending row number of the range
    Dim sLtr As String  ' holds the starting column letter of the range
    Dim eLtr As String  ' holds the ending column letter of the range
    Dim cRng As String  ' holds a string that represents the column range
    Dim rRng As String  ' holds a string that represents the row range
    Dim TL() As Variant ' recieves PTL_RefData from named store
    
    ' place the data from PTL_RefData into local array
    TL = [PTL_RefData]
    
    x = 35      ' set base colorindex
    y = 0.75    ' set TintandShade variable to a lighter shade -1 to 1 is darker to lighter
    
    ' work thru ranges to colorize
    For i = 1 To [PTL_Hd2]
        ' read row and column variables from global array
        sRow = Right(TL(i, 2), Len(TL(i, 2)) - 1)
        sLtr = Left(TL(i, 2), 1)
        eRow = Right(TL(i, 3), Len(TL(i, 3)) - 1)
        eLtr = Left(TL(i, 3), 1)
        
        ' set working area for this iteration
        cRng = sLtr & sRow & ":" & sLtr & eRow
        rRng = sLtr & sRow & ":" & eLtr & eRow
        
        ' only applying color indexes from 35 to 56
        If x > 56 Then
            x = 35
        Else
            With Range(cRng).Interior
                .ColorIndex = x
                .TintAndShade = y
            End With
            With Range(rRng).Interior
                .ColorIndex = x
                .TintAndShade = y
            End With
            
'            ' adjust font color to white if background color is dark
'            Select Case x
'                Case 35 To 40, 42 To 46, 48, 50
'                    Range(cRng).Font.ColorIndex = 1
'                    Range(rRng).Font.ColorIndex = 1
'                Case Else
'                    Range(cRng).Font.ColorIndex = 2
'                    Range(rRng).Font.ColorIndex = 2
'            End Select
        End If
        
        ' set next color index to use
        x = x + 1
    Next i
End Sub

Sub PTLRemoveHeader()
    ' clean up original header row
    Range("A1:A2").EntireRow.Delete
End Sub

Sub PTLSetChange()
    ' sets global boolean that is checked by the change event of the sheet
    ' set at the end of the formatting routines so that the Worksheet_Change event
    ' begins to function
    Names.Add Name:="bChange", RefersTo:="=True"
    Names.Add Name:="bFormatted", RefersTo:="=True"
End Sub

Sub PTLCreateRefArray()
    Dim myRange As Range ' working range for this procedure
    Dim TL() As Variant ' array holding various sheet elements that make the base data, passed to PTL_RefData
    Dim i As Integer ' iteration counter
    Dim n As Integer ' iteration counter (2 iterations occuring in this procedure)
    Dim ref1 As String ' heading information
    Dim ref2 As String ' column "B" row reference signifying the start cell of the range with headings
    Dim ref3 As String ' last column, last row of area signifying the end cell of the range
    Dim ref4 As String ' column "C" of area, signifying the start cell of the working area of the range
    Dim ref5 As Variant ' will contain the target date for the item, or 'Completed'
    Dim sRow As Long ' a starting point for when the loop finds a number 2 as a value, start row of range
    Dim eRow As Long ' the corresponding ending row of the range
    Dim eCol As String ' the column letter associated with the greatest column in the range
    
    ' set routine variables
    i = 0
    eCol = ColumnLetter([PTL_Cols])
    ReDim Preserve TL([PTL_Hd2], 4)
    
    ' begin loop through entire table to reference each project timeline
    ' [PTL_Hd2] is the number of seperate projects (Heading 2's found)
    For n = 1 To [PTL_Hd2]
    
        ' find header2 row
        Do Until Range("A1").Offset(i, 0).Value = 2
           i = i + 1
        Loop
        
        ' set variable to the found row number
        sRow = Range("A1").Offset(i, 0).Row
        
        ' start next search at next row
        i = i + 1
        
        ' find end row in block
        ' if the loop goes beyond the number of headers then we've gone too far
        Do Until Range("A1").Offset(i, 0).Value = 2
            If n < [PTL_Hd2] Then
                i = i + 1
            Else
                Exit Do
            End If
        Loop
        
        'handle the last row in the table
        If n = [PTL_Hd2] Then
            eRow = Cells(Rows.count, 1).End(xlUp).Row
        Else
            eRow = Range("A1").Offset(i, 0).Row - 2
        End If
        
        ' set range reference
        ref2 = "B" & sRow
        ref3 = eCol & eRow
        ref4 = Range(ref2).Offset(1, 0).Address(rowabsolute:=False, columnabsolute:=False)
        
        Set myRange = Range(ref2 & ":" & ref3)
        
        ' save heading to a string to search with
        ref1 = myRange(1, 1)
        
        ' populate array
        TL(n, 1) = ref1 ' heading 2 text
        TL(n, 2) = ref2 ' starting cell of individual project range
        TL(n, 3) = ref3 ' ending cell of individual project range
        TL(n, 4) = ref4 ' starting cell of working area within the project range
        
    Next n
    
    ' store the array to PTL_RefData for future use
    Names.Add Name:="PTL_RefData", RefersTo:=TL
End Sub

Sub PTLExport()
    Dim wdApp As Word.Application ' early binding variable to ms word application
    Dim wdDoc As Word.Document ' early binding variable to ms word document
    Dim wdRng As Word.Range ' early binding variable to ms word range
    Dim wdPar As Word.Paragraph ' early binding variable to ms word paragraph
    Dim myRange As Excel.Range ' working range for this procedure
    Dim myFile As String ' sting holding the full path of the file with the Word suffix attached
    Dim bMatch As Boolean ' used in a loop to determine when to break out of loop
    Dim i As Integer ' iteration counter
    Dim bRunning As Boolean ' identifies whether Word is running or not
    Dim TL() As Variant ' recieves PTL_RefData from named store, which hold project timeline header and range info
    
    ' place the data from PTL_RefData into local array
    PTLCreateRefArray
    TL = [PTL_RefData]
    
    On Error Resume Next
    
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
            bRunning = True
        End If
        Set wdDoc = wdApp.Documents.Open(myFile)
        If wdDoc Is Nothing Then
            MsgBox "Failed to open help document!", vbCritical
            If bRunning Then
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

    Set wdRng = wdDoc.Bookmarks("Action_Areas").Range
    With wdRng
        .Select
        .Collapse Direction:=wdCollapseStart
    End With
    
    ' Turn off screen updating.
    Application.ScreenUpdating = False
    
    ' loop through to the upper bounds of the array
    i = 1
    Do Until i > UBound(TL)
        bMatch = False
        ' copy the project range to be pasted into Word
        Set myRange = Range(TL(i, 2) & ":" & TL(i, 3))
        myRange.Copy
        
        ' reset the Find dialog info
        With wdRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute
        End With
        
        ' find the area matching hd2
        With wdRng.Find
            .ClearFormatting
            .Text = TL(i, 1)
            .Execute
        End With
        wdRng.Collapse Direction:=wdCollapseStart
        
        ' reset the Find dialog info
        With wdRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute
        End With
        
        ' find the next instance of Timeline Heading 3
        With wdRng.Find
            .ClearFormatting
            .Text = "Timeline"
            Do
                .Execute
                    ' check to make sure that this is the correct "Timeline"
                    Set wdPar = wdRng.Paragraphs(1)
                    If Left(wdPar.Range.Text, 8) = "Timeline" And wdPar.Style = "Heading 3" Then
                        bMatch = True
                        With wdRng
                            .Collapse Direction:=wdCollapseEnd
                            .InsertParagraphAfter
                            .PasteExcelTable LinkedToExcel:=True, WordFormatting:=False, RTF:=False
                        End With
                    wdRng.Collapse Direction:=wdCollapseEnd
                    End If
                    Set wdPar = Nothing
            Loop Until bMatch = True
        End With
        Set myRange = Nothing
        i = i + 1
    Loop
       
    'Save and close the Word doc.
    wdDoc.Save
    
    'Quit Word.
    wdApp.Quit
    
    'Null out your variables.
    Set wdRng = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    'Clear out the clipboard, and turn screen updating back on.
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
    End With
    
    MsgBox "The Timeline has been successfully " & vbNewLine & _
           "transferred to " & [wdFullName], vbInformation

End Sub

Sub ShtPTLWorkAreasRange()
    ' used to set mutliple ranges as the area to check when a cell is changed
    Dim i As Integer ' iteration counter
    Dim rng1 As Range ' range used to build a union range
    Dim rng2 As Range ' range used to build a union range
    Dim TL() As Variant ' recieves RefData from named store
    
    ' place the data from RefData into local array
    TL = [PTL_RefData]
    
    Set gWorkRange = Nothing
    
    For i = 1 To [PTL_Hd2] - 1
        If gWorkRange Is Nothing Then
            Set rng1 = Range(TL(i, 4) & ":" & TL(i, 3))
            Set rng2 = Range(TL(i + 1, 4) & ":" & TL(i + 1, 3))
            Set gWorkRange = Union(rng1, rng2)
        Else
            Set rng1 = gWorkRange
            Set rng2 = Range(TL(i + 1, 4) & ":" & TL(i + 1, 3))
            Set gWorkRange = Union(rng1, rng2)
        End If
    Next i
    
End Sub

Sub ShtPTLSetRange(ByVal Target As Range)
    ' used to set mutliple ranges as the area to check when a cell is changed
    Dim i As Integer ' iteration counter
    Dim xCol As Long
    Dim xRow As Long
    Dim sCol As String
    Dim myRange As Range ' working range within the procedure
    Dim sortRange As Range ' range to sort
    Dim sRng As Range ' start of range to perform work on
    Dim eRng As Range ' end off range to perform work on
    Dim cell As Range ' iteration range
    Dim bgColor As Variant ' background color index
    Dim fontColor As Variant ' font color index
        
    PTLCreateRefArray
    ShtPTLWorkAreasRange
    
    ' is the target in the working ranges
    If Intersect(Target, gWorkRange) Is Nothing Then Exit Sub
    
    ' set up the variables to work through and format the target cells
    xRow = Target.Row
    xCol = gWorkRange.Columns.count + 2
    sCol = ColumnLetter(xCol)
    Set myRange = Range("C" & xRow & ":" & sCol & xRow)
    Set sortRange = myRange.CurrentRegion.Resize(, myRange.Columns.count + 5)
    bgColor = Range("B" & xRow).Interior.ColorIndex
    fontColor = Range("B" & xRow).Font.ColorIndex
    Set sRng = Range(sCol & xRow).Offset(0, 2)
    Set eRng = Range(sCol & xRow).Offset(0, 3)
    i = 0
    
    Application.EnableEvents = False
    
    ' set the task timeframe start and task timeframe end cells
    For i = 1 To Target.Columns.count
        If i = 1 And Target.Columns.count = 1 Then
            With sRng
                .Value = Target.Column
                .Font.ColorIndex = 2
            End With
            With eRng
                .Value = Target.Column
                .Font.ColorIndex = 2
            End With
        ElseIf i = 1 Then
            With sRng
                .Value = Target.Column
                .Font.ColorIndex = 2
            End With
        ElseIf i = Target.Columns.count Then
            With eRng
                .Value = Target.Cells(Target.count).Column
                .Font.ColorIndex = 2
            End With
        Else
        End If
    Next i
    
    For i = 3 To xCol
        ' clear the contents and set the background color of myRange
        For Each cell In myRange
            With cell
                ' clear the cells
                .ClearContents
                .Interior.ColorIndex = xlColorIndexNone
            End With
        Next cell
                
        ' work with the selection that represents the working timeline
        For Each cell In Target
            With cell
                .Value = "X"
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = bgColor
                .Interior.TintAndShade = 0.95
                .Font.ColorIndex = fontColor
            End With
        Next cell
    Next i
    Application.EnableEvents = True

    ' sort the current objective timeline
    sortRange.Sort Key1:=sRng, key2:=eRng, order1:=xlAscending, order2:=xlAscending, Header:=xlYes
    
End Sub

Sub PTL_SetTargetDate()
    ' Find the set target date info from the goals
    Dim i As Integer ' count through the rows on the sheet
    Dim n As Integer ' count through the columns in a row
    Dim aTarget() As Variant  ' array to hold the target date for each goal
    Dim aHd() As Variant
    
    ReDim aTarget([PTL_Rows])
    aHd = [PTL_Hd1]
    
    ' iterate through the timeline to find the targets
    For i = 1 To [PTL_Rows]
        ' only interested in Goal rows
        If Range("A" & i).Value = 3 Then
            ' iterate through each column to find the target
            For n = 3 To [PTL_Cols]
                If Not IsEmpty(Worksheets("ProjectTimeline").Cells(i, n).Value) Then
                    ' make column adjustment when reading PTL_Hd1
                    ' this will overwrite the array ref until the last item in the row is found
                    aTarget(i) = aHd(n)
                End If
            Next n
        End If
    Next i
        
    ' store the array to PTL_TargetDate for future use
    Names.Add Name:="PTL_TargetDate", RefersTo:=aTarget
End Sub
