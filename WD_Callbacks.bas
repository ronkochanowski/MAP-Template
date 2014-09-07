Attribute VB_Name = "Callbacks"
Option Explicit

'Callback for rxbutInsertProject onAction
Sub rxbutCreateProject_click(control As IRibbonControl)
    Call CreateNewProject
End Sub

'Callback for rxbutInsertGoal onAction
Sub rxbutCreateGoal_click(control As IRibbonControl)
    Call InsertNewGoal
    Call BuildGoals
End Sub

'Callback for rxbutDeleteProject onAction
Sub rxbutDeleteProject_click(control As IRibbonControl)
    Call DeleteProject
End Sub

'Callback for rxbutExcelEpxort onAction
Sub rxbutExcelEpxort_click(control As IRibbonControl)
    Call GetInfoForExcel
End Sub
