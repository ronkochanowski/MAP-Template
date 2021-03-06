Attribute VB_Name = "MAP_Excel"
Option Explicit
Option Base 1

' controller for creating array and moving to excel
Sub GetInfoForExcel()
    Call SaveWordDocument
    Call CheckForExcelFile
End Sub

' Save Word file based on MAP template to new file (forced save)
Sub SaveWordDocument()
    Dim dlgSaveAs As FileDialog
    
    ' has ActiveDocument been saved?
    If ActiveDocument.Saved Then
        ActiveDocument.Save
    Else
        ' save this document before moving forward
        Set dlgSaveAs = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
        
        With dlgSaveAs
            .Title = "Save MAP As..."
            .FilterIndex = 2
            .InitialFileName = Options.DefaultFilePath(wdDocumentsPath)
            .Show
            .Execute
        End With
        
    End If
End Sub

' check if Excel file exists with same name
Sub CheckForExcelFile()
    Dim sPath As String
    
    sPath = Left(ActiveDocument.FullName, Len(ActiveDocument.FullName) - 5) & ".xlsx"
    
    If File_Exists(sPath) Then
        ' call procedure to get the existing file and open it
        Call OpenExistingTimeline
    
    Else
        ' create a new Excel file based on the MAP Template template
        Call CreateNewTimeline
    
    End If
    
End Sub

' create a new workbook based on the Timline template
Sub CreateNewTimeline()
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlWS As Excel.Worksheet
    Dim xlLO As Excel.ListObject
    Dim sTemplate As String
    Dim sFullPath As String
    Dim xlRng As Excel.Range
    
    ' if develop mode then point to the develop mode excel template
    If gDevMode Then
        sTemplate = Options.DefaultFilePath(wdUserTemplatesPath) & "\Dev\MAP Template.xltx"
    Else
        sTemplate = Options.DefaultFilePath(wdUserTemplatesPath) & "\MAP Template.xltx"
    End If
    
    ' prepare Excel objects to write to
    Set xlApp = New Excel.Application
    Set xlWB = xlApp.Workbooks.Open(sTemplate)
    Set xlWS = xlWB.Sheets("Project Timeline")
    Set xlLO = xlWS.ListObjects("Table2")
    
    ' set path to save Excel file same as for Word file
    sFullPath = Left(ActiveDocument.FullName, Len(ActiveDocument.FullName) - 5)
    
    ' save Excel file
    xlWB.SaveAs FileName:=sFullPath, FileFormat:=xlOpenXMLWorkbook
    
    ' after saving the Project Timeline, create the array in Word
    Call CopyHeadingsToArray
    
    ' call funcetion to paste site number and abbreviation in excel header
    xlWS.Range("B1").Value = SiteAbbr

    ' paste headings from array into excel
    xlWS.Range("B5:C" & UBound(gHeadings, 2)) = xlApp.Transpose(gHeadings)
    
    MsgBox "The Timeline header information has been successfully " & vbNewLine & _
      "transferred to the Excel file" & vbNewLine & sFullPath & ".xlsx", vbInformation
    
    ' show Excel
    xlApp.Visible = True
    
    'Tidy up
    Set xlLO = Nothing
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    
End Sub

' find headings 2, 3 inside of Action Areas and build array
Sub CopyHeadingsToArray()
    Dim i As Long, a As Long, b As Long
    Dim myRange As Word.Range
    Dim myPar As Word.Paragraph
    
    ' turn off screen updating
    With Application
        .ScreenUpdating = False
        .StatusBar = "Grabbing information to pass to Excel..."
    End With
        
    ' move thru range and place headings into array
    i = 1
    Set myRange = ActiveDocument.Range.Bookmarks("Action_Areas").Range
    For Each myPar In myRange.Paragraphs
        If Left(myPar.Range.Text, 8) <> "Timeline" Then
            Select Case myPar.Style
                Case "Heading 2"
                    ReDim Preserve gHeadings(2, i)
                    gHeadings(1, i) = Trim(myPar.Range.Text)
                    i = i + 1
                Case "Heading 3"
                    ReDim Preserve gHeadings(2, i)
                    gHeadings(2, i) = Trim(myPar.Range.Text)
                    i = i + 1
                Case Else
            End Select
        End If
    Next
    
    ' turn on screen updating
    With Application
        .ScreenUpdating = True
        .StatusBar = ""
    End With
End Sub


Sub OpenExistingTimeline()
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlWS As Excel.Worksheet
    
    Dim bRunning As Boolean
    Dim sFullPath As String
    Dim sApp As String
    
    ' set the variable pointing to the file being looked for...
    sFullPath = Left(ActiveDocument.FullName, Len(ActiveDocument.FullName) - 5) & ".xlsx"
    sApp = "Excel.Application"
    
    ' check to see if Excel is running, if not, start it
    bRunning = IsAppRunning(sApp)
    If bRunning = False Then
        Set xlApp = New Excel.Application
    Else
        Set xlApp = GetObject(, sApp)
    End If
    
    'open the existing workbook
    Set xlWB = xlApp.Workbooks.Open(FileName:=sFullPath)
        
    'check if data exists on worksheet
    Set xlWS = xlWB.Sheets("Project Timeline")
    Set xlLO = xlWS.ListObjects("Table2")
    If xlIsDirty(xlWS) Then
    
        ' confirm that the data is to be cleared, if the answer is yes then clear and replace
        If MsgBox("Data has been previously placed on this worksheet." & vbCrLf _
         & "Do you want to overwrite what was previously entered?", vbYesNo, "Confirm Overwriting Data") _
         = vbYes Then
         
            ' after clearing the data, this would be the next statement. Recreate the array in Word
            Call CopyHeadingsToArray
            
            ' call funcetion to paste site number and abbreviation in excel header
            xlWS.Range("B1").Value = SiteAbbr
        
        
            ' paste headings from array into excel
            xlWS.Range("B5:C" & UBound(gHeadings, 2)) = xlApp.Transpose(gHeadings)
            
            ' show Excel
            xlApp.Visible = True
            
            ' tidy up
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing

        Else
            ' if the response to overwriting is No, then get out of here!
            Exit Sub
        End If
    Else
        ' after clearing the data, this would be the next statement. Recreate the array in Word
        Call CopyHeadingsToArray
    
        ' call funcetion to paste site number and abbreviation in excel header
        xlWS.Range("B1").Value = SiteAbbr
        
        ' paste headings from array into excel
        xlWS.Range("B5:C" & UBound(gHeadings, 2)) = xlApp.Transpose(gHeadings)
        
        ' show Excel
        xlApp.Visible = True
        
        ' tidy up
        Set xlWS = Nothing
        Set xlWB = Nothing
        Set xlApp = Nothing
    End If
    
End Sub
