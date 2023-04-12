Attribute VB_Name = "EventsSheetChange"
Attribute VB_Description = "Events for changes in a worksheet"
Option Explicit

'@ModuleDescription("Events for changes in a worksheet")
'@IgnoreModule ProcedureNotUsed
'@Folder("Events")

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Application.Cursor = xlNorthwestArrow
    If Me.Name <> "__checkRep" Then
        EventsGlobal.checkUpdateStatus Me, Target
    Else
       EventsGlobal.FilterCheckingsSheet Target
    End If
    'Only for analysis
    If Me.Name = "Analysis" Then
       EventsAnalysis.CalculateAnalysis
       EventsAnalysis.AddChoicesDropdown Target
    End If
    Application.Cursor = xlDefault
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_Activate()
    If Me.Name = "Analysis" Then EventsAnalysis.EnterAnalysis
End Sub
