Attribute VB_Name = "Globals"
Option Explicit
Option Base 1

Public gWorkRange As Range ' range that refers to the various work areas within the timeline

Sub PTLInitializeGlobals()
    ' set workbook named ranges. using these because they hold their values when workbook is closed.
    With ActiveWorkbook.Names
        .Add Name:="PTL_Rows", RefersTo:="=GetRows()"
        .Add Name:="PTL_Cols", RefersTo:="=GetCols()"
        .Add Name:="PTL_Hd2", RefersTo:="=CountIf(ProjectTimeline!$A:$A, 2)"
        If Not [bFormatted] Then
            .Add Name:="bChange", RefersTo:="=False"
        End If
    End With
End Sub
