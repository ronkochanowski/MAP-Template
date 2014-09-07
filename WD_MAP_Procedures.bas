Attribute VB_Name = "MAP_Procedures"
Option Explicit
Option Base 1

' action routine that builds an Project structure.
Sub CreateNewProject()
    ' run each subroutine in order to create a new section in the document
    Call InsertNewProject
    Call InsertProjectHeading
    Call BuildGoals
    Call InsertTimelineHeading
End Sub

' base routine to build from...called by CreateNewProject
Sub InsertNewProject()
    Dim myRange As Range
    Dim Sect As Long
    
    ' Determine the current section
    Sect = Selection.Information(wdActiveEndSectionNumber)
    
    ' add a new section after the current section
    Set myRange = ActiveDocument.Sections(Sect).Range
    myRange.Collapse direction:=wdCollapseEnd
    ActiveDocument.Sections.Add Range:=myRange, Start:=wdSectionNewPage
     
    ' move cursor to beginning of new section
    Selection.GoToNext what:=wdGoToSection
End Sub

Sub InsertNewGoal()

    'find next heading 3 occurance
    With Selection.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles("Heading 3")
        .Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
    End With
    
    'turn off Overtype if it is turned on
    If Application.Options.Overtype Then
        Application.Options.Overtype = False
    End If
        
    With Selection
    
        'check if the current selection is an insertion point
        If .Type = wdSelectionIP Then
            .TypeParagraph
            .MoveUp unit:=wdLine, Count:=1
            
        'otherwise, move to the beginning of the selection first
        ElseIf .Type = wdSelectionNormal Then
            .Collapse direction:=wdCollapseStart
            .TypeParagraph
            .MoveUp unit:=wdLine, Count:=1
        End If
        
    End With
    
End Sub
    
' base routine to build from...called by CreateNewProject
Sub InsertProjectHeading()
    Dim myRange As Range
    Dim Sect As Long
    
    ' Get the number of the current section and add one for the section just created
    Sect = Selection.Information(wdActiveEndSectionNumber)
     
    ' set Range
    Set myRange = ActiveDocument.Sections(Sect).Range
    myRange.Collapse direction:=wdCollapseStart
         
    ' add text
    With Selection
        ' Project Heading
        .Style = ActiveDocument.Styles("Heading 2")
        .TypeText Text:="Project Title"
        
        ' add a summary area as a content control
        .TypeParagraph
        .TypeParagraph
        .MoveLeft unit:=wdCharacter, Count:=1
        .Collapse direction:=wdCollapseStart
        
        ' insert content control
        Call InsertSummaryCC
        .MoveRight unit:=wdCharacter, Count:=1
        .TypeParagraph
        .TypeParagraph
        .MoveUp unit:=wdLine, Count:=1
    End With
End Sub

' base routine to build from...called by InsertProjectHeading
Sub InsertSummaryCC()
    Dim objCC As ContentControl

    ' insert content control to enter summary information
    Set objCC = Selection.ContentControls.Add(wdContentControlRichText)
    With objCC
        .Title = "Project Summary"
        .SetPlaceholderText Text:="Please enter the Project's Summary "
        .Tag = "summary"
    End With
    Set objCC = Nothing
    
End Sub

' base routine to build from...called by BuildGoals thru CreateNewProject
Sub InsertGoalHeading()
    ' Goal Heading
    With Selection
        ' insert heading
        .Style = ActiveDocument.Styles("Heading 3")
        .TypeText Text:="Goal Title"
        .TypeParagraph
    End With
    
End Sub

' base routine to build from...called by BuildGoals thru CreateNewProject
Sub InsertGoalPriority()
    
    ' Goal Element - Priority
    With Selection
        ' insert the header text
        .Style = ActiveDocument.Styles("Heading 4")
        .TypeText Text:="Priority:  "
        .TypeParagraph
        
        ' insert content control
        Call InsertPriorityCC
        .MoveLeft unit:=wdCharacter, Count:=2
        .InsertStyleSeparator
        .MoveRight unit:=wdCharacter, Count:=2
        .TypeParagraph

    End With
     
End Sub

' base routine to build from...called by BuildGoals thru CreateNewProject
Sub InsertPriorityCC()
    Dim objCC As ContentControl
    
    ' insert content control to pick priority
    Set objCC = Selection.ContentControls.Add(wdContentControlComboBox)
    With objCC
        .Title = "Goal Priority"
        .SetPlaceholderText Text:="Please select the Goal Priority "

         'List entries
        .DropdownListEntries.Add "Critical"
        .DropdownListEntries.Add "High"
        .DropdownListEntries.Add "Normal"
        .DropdownListEntries.Add "Low"
    End With
    Set objCC = Nothing
End Sub

' base routine to build from...called by BuildGoals thru CreateNewProject
Sub InsertGoalDependency()

    With Selection
        ' Goal Element - Dependency
        .Style = ActiveDocument.Styles("Heading 4")
        .TypeText Text:="Dependency:  "
        .TypeParagraph
        .MoveLeft unit:=wdCharacter, Count:=1
        .InsertStyleSeparator
        .Style = ActiveDocument.Styles("Normal")
        .TypeText Text:="Insert dependencies and restrictions to implementing this goal."
        .TypeParagraph
    End With

End Sub
            
' base routine to build from...called by BuildGoals thru CreateNewProject
Sub InsertGoalSupInfo()

    With Selection
        ' Goal Element - Supporting Information
        .Style = ActiveDocument.Styles("Heading 4")
        .TypeText Text:="Supporting Information:  "
        .TypeParagraph
        .MoveLeft unit:=wdCharacter, Count:=1
        .InsertStyleSeparator
        .Style = ActiveDocument.Styles("Normal")
        .TypeText Text:="Describe any information needed to help support this goal."
        .TypeParagraph
    End With
    
End Sub

' base routine to build from...called by BuildGoals thru CreateNewProject
Sub InsertGoalTask()

    With Selection
        ' Goal Element - Tasks
        .Style = ActiveDocument.Styles("Heading 4")
        .TypeText Text:="Tasks:"
        .TypeParagraph
        .MoveLeft unit:=wdCharacter, Count:=1
        .InsertStyleSeparator
        .TypeParagraph
        .Style = ActiveDocument.Styles("List Paragraph")
        With .Range.ListFormat
            .ListOutdent
            .ApplyBulletDefault DefaultListBehavior:=wdWord9ListBehavior
        End With
        .TypeText Text:="Bullet point any specific steps needed to accomplish this goal."
        .TypeParagraph
        .TypeText Text:="Use 'Action Styles' to bring greater attention to a particular point."
        .TypeParagraph
        .Style = ActiveDocument.Styles("Normal")
    End With
    
End Sub

Sub InsertTimelineHeading()

    With Selection
        ' Timeline Heading
        .Style = ActiveDocument.Styles("Heading 3")
        .TypeText Text:="Timeline"
        .TypeParagraph
    End With
    
End Sub

' action routine that deletes the current Project section
Sub DeleteProject()
    Dim myRange As Range
    Dim Sect As Long
    
     ' Determine the current section
    Sect = Selection.Information(wdActiveEndSectionNumber)
    
     ' delete the current section
    Set myRange = ActiveDocument.Sections(Sect).Range
    myRange.Delete
End Sub

' action routine that builds one or more Goal structures
Sub BuildGoals()
    Dim intGoals As Integer
    Dim i As Integer
    Dim msb As Integer
    
    ' Run the Error handler "ErrHandler" when an error occurs.
    On Error GoTo Errhandler

myInput:
    intGoals = InputBox("How many Goals need to be created for this Project? (Cannot be blank or zero)", "Number of Goals")
    
    ' create each Goal element in blocks based on the number entered
    If intGoals = 0 Then
        msb = MsgBox("Entry MUST be a greater than zero!", vbCritical, "Error")
        GoTo myInput
    ElseIf intGoals = 1 Then
        Call InsertGoalHeading
        Call InsertGoalPriority
        Call InsertGoalDependency
        Call InsertGoalSupInfo
        Call InsertGoalTask
    Else
        For i = 1 To intGoals
            Call InsertGoalHeading
            Call InsertGoalPriority
            Call InsertGoalDependency
            Call InsertGoalSupInfo
            Call InsertGoalTask
        Next i
    End If
    
Errhandler:
    ' if not an integer give the user the opportunity to correct
    If Err = 13 Then
        msb = MsgBox("Entry MUST be a number!", vbCritical, "Error")

        ' Resume at the line where the error occurred if the user
        ' clicks OK; otherwise delete the started Project
        ' and end the macro.
        If msb = vbOK Then
            Resume
        End If
        Call DeleteProject
    End If
End Sub
