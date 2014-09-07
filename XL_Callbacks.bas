Attribute VB_Name = "Callbacks"
Option Explicit

'Callback for rxbutInsertObjective onAction
Sub rxbutPTLFormat_click(control As IRibbonControl)
    Call PTLFormat
End Sub

'Callback for rxbutExportTimeline onAction
Sub rxbutPTLExport_click(control As IRibbonControl)
    Call PTLExport
    Call MTL_Export
    Call PT_Export
End Sub

'Callback for rxbutPTLSetTimeframe onAction
Sub rxbutPTLSetTimeframe_click(control As IRibbonControl)
    Call ShtPTLSetRange(Selection)
End Sub

'Callback for rxbutBuild onAction
Sub rxbutBuild_click(control As IRibbonControl)
    Call MTLfromPTL
    Call PT_Build
End Sub
