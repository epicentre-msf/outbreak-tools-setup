Attribute VB_Name = "EventsSheetChange"
Attribute VB_Description = "Events for changes in a worksheet"
Option Explicit

'@ModuleDescription("Events for changes in a worksheet")
'@IgnoreModule ProcedureNotUsed
'@Folder("Events")

'speed app
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Application.Cursor = xlNorthwestArrow

    BusyApp

    If Me.Name <> "__checkRep" Then
        EventsGlobal.checkUpdateStatus Me, Target
    Else
       EventsGlobal.FilterCheckingsSheet Target
    End If

    

    'Only for analysis
    If Me.Name = "Analysis" Then
       
       BusyApp
       EventsAnalysis.CalculateAnalysis
       
       BusyApp
       EventsAnalysis.AddChoicesDropdown Target

    End If

    Application.Cursor = xlDefault
    Application.EnableEvents = True
    NotBusyApp
End Sub


Private Sub Worksheet_Activate()
    BusyApp

    If Me.Name = "Analysis" Then EventsAnalysis.EnterAnalysis

    NotBusyApp
End Sub
