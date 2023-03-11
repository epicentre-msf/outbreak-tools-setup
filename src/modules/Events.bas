Attribute VB_Name = "Events"

Option Explicit

'Module for all the events - related actions in the setup file

'Import from another setup
Public Sub clickImport(ByRef Control As Office.IRibbonControl)
    [Imports].Show
End Sub

'add rows to listObject
Public Sub clickAddRows(ByRef Control As Office.IRibbonControl)
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    ManageRows sheetName:=sheetName, del:=False
End Sub

'resize the current listObject
Public Sub clickResize(ByRef Control As Office.IRibbonControl)
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    ManageRows sheetName:=sheetName, del:=True
End Sub

'clear data in the current setup
Public Sub clickClearSetup(ByRef Control As Office.IRibbonControl)
End Sub

'check the current setup for incoherences
Public Sub clickCheck(ByRef Control As Office.IRibbonControl)
End Sub

'speed app
Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'Add or Remove Rows to a table
Public Sub ManageRows(ByVal sheetName As String, Optional ByVal del As Boolean = False)
    Dim part As Object
    Dim sh As Worksheet
    Dim shpass As Worksheet
    Dim pass As IPasswords

    BusyApp
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(sheetName)
    Set shpass = ThisWorkbook.Worksheets("__pass")
    On Error GoTo 0

    If (sh Is Nothing) Or (shpass Is Nothing) Then Exit Sub

    '5 is the start line of the dictionary
    '4 is the start column of the dictionary
    Select Case sheetName
    Case "Dictionary"
        Set part = LLdictionary.Create(sh, 5, 1)
    Case "Choices"
        Set part = LLchoice.Create(sh, 4, 1)
    Case "Analysis"
        Set part = Analysis.Create(sh)
    End Select

    'Exit if unable to find the corresponding object
    If part Is Nothing Then Exit Sub
    Set pass = Passwords.Create(shpass)
    pass.UnProtect sh.Name

    If del Then
        part.RemoveRows
    Else
        part.AddRows
    End If

    pass.Protect sh.Name
    NotBusyApp
End Sub
