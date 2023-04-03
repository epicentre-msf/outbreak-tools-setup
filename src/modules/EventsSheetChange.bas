Attribute VB_Name = "EventsSheetChange"
Option Explicit

'@ModuleDescription("Events for changes in a worksheet")
'@IgnoreModule ProcedureNotUsed

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Application.Cursor = xlNorthwestArrow
    checkUpdateStatus Me, Target
    'Only for analysis
    If Me.Name = "Analysis" Then
        CalculateAnalysis
        AddChoicesDropdown Target
    End If
    Application.Cursor = xlDefault
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_Activate()
    If Me.Name = "Analysis" Then EnterAnalysis
End Sub
