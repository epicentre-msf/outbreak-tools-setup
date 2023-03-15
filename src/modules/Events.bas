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

Public Sub clickUpdateTranslate()
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

'Fire this event when leaving the dictionary
Public Sub EnterAnalysis()

    Dim dict As ILLdictionary
    Dim drop As IDropdownLists
    Dim lst As BetterArray
    Dim vars As ILLVariables
    Dim upObj As IUpdatedValues

    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("Dictionary"), 5, 1)
    Set vars = LLVariables.Create(dict)
    Set drop = DropdownLists.Create(ThisWorkbook.Worksheets("__variables"))
    Set upObj = UpdatedValues.Create(ThisWorkbook.Worksheets("__updated"), "dict")

    BusyApp

    If upObj.IsUpdated("control_details") Then
        'Update geo vars
        Set lst = dict.GeoVars()
        drop.Update lst, "__geo_vars"
        'Update choices vars
        Set lst = dict.ChoicesVars()
        drop.Update lst, "__choice_vars"
    End If

    If upObj.IsUpdated("variable_type") Then
        'Update time vars
        Set lst = dict.TimeVars()
        drop.Update lst, "__time_vars"
    End If

    NotBusyApp
End Sub

Private Sub FormatLockCell(ByVal cellRng As Range, Optional ByVal Locked = True)
    cellRng.Font.Color = IIf(Locked, RGB(51, 142, 202), vbBlack)
    cellRng.Font.Italic = Locked
    cellRng.Locked = Locked
End Sub


'Add Dropdown on choices
Public Sub AddChoicesDropdown(ByVal Target As Range)

    Const LOBJNAME As String = "Tab_Graph_TimeSeries"
    Const LOBJTSNAME As String = "Tab_TimeSeries_Analysis"

    Dim sh As Worksheets
    Dim csTab As ICustomTable
    Dim tsTab As ICustomTable
    Dim colValue As String
    Dim drop As IDropdownLists
    Dim dropArray As BetterArray
    Dim choi As ILLchoice
    Dim seriestitleRng As Range
    Dim colValue As String
    Dim choiceName As String
    Dim cellRng As Range
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim sumLab As String

    Set sh = ThisWorkbook.Worksheets("Analysis")
    Set csTab = CustomTable.Create(sh.ListObjects(LOBJNAME), "series title")
    Set seriestitleRng = csTab.DataRange("series title")

    If InterSect(Target, seriestitleRng) Is Nothing Then Exit Sub

    'Create the choices object
    Set choi = LLchoice.Create(ThisWorkbook.Worksheets("Choices"), 4, 1)
    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("Choices"), 5, 1)
    Set vars = LLVariables.Create(dict)

    'Now get the value of column on the custom table and test it
    colValue = csTab.Value(colName:="column", keyName:=Target.Value)

    If colValue <> vbNullString Then
        choiceName = vars.Value(colName:="Control Details", varName:=colValue)
        Set dropArray = choi.Categories(choiceName)
        drop.Add dropArray, choiceName & "__"
        drop.Update dropArray, choiceName & "__"

        'get the cell Range for choices
        Set cellRng = csTab.CellRange("choice", Target.Row)
        drop.SetValidation cellRng, choiceName & "__"
        FormatLockCell cellRng, False

        'get the cell Range for plot values or percentage
        Set cellRng = csTab.CellRange("values or percentages")
        drop.SetValidation cellRng, "__perc_val"
        FormatLockCell cellRng, False
    Else
        'Get the cellRang for choice
        Set cellRng = csTab.CellRange("choice", Target.Row)
        Set tsTab = CustomTable.Create(sh.ListObjects(TSNAME), "title")
        sumLab = tsTab.Value(colName:="summary label", keyName:=Target.Value)
        FormatLockCell cellRng, True

        Set cellRng = csTab.CellRange("values or percentages")
        cellRng.Value = "values"
        FormatLockCell cellRng, True
    End If
End Sub
