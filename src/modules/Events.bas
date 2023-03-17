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

Public Sub clickUpdateTranslate(ByRef Control As Office.IRibbonControl)
    'remove update columns and add new columns to watch
    BusyApp
    CleanUpdateColumns
    UpdatedWatchedValues
    NotBusyApp
    MsgBox "Done!"
End Sub

'Clear the names of the columns to update
Private Sub CleanUpdateColumns()
    'Clear the update sheet
    Dim upSh As Worksheet
    Dim Lo As ListObject
    Dim wb As Workbook
    Dim namesRng As Range
    Dim counter As Long
    Set wb = ThisWorkbook
    Set upSh = wb.Worksheets("__updated")

    'Unlist all listObjects in the worksheet and delete all names
    For Each Lo In upSh.ListObjects
        Set namesRng = Lo.ListColumns("rngname").Range
        For counter = 1 To namesRng.Rows.Count
            On Error Resume Next
            wb.Names(namesRng.Cells(counter, 1).Value).Delete
            On Error GoTo 0
        Next
        Lo.Unlist
    Next
    upSh.Cells.Clear
End Sub

'Update the translation values
Private Sub UpdatedWatchedValues()
    Dim sh As Worksheet
    Dim sheetsList As BetterArray
    Dim counter As Long
    Dim sheetName As String

    Set sheetsList = New BetterArray
    sheetsList.Push "Dictionary", "Choices", "Exports", "Analysis"
    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        Set sh = ThisWorkbook.Worksheets(sheetName)
        'Write update status on each sheet
        writeUpdateStatus sh
    Next
End Sub

'Update status of columns to watch
Private Sub writeUpdateStatus(sh As Worksheet)
    Dim upSh As Worksheet
    Dim upId As String
    Dim upObj As IUpdatedValues
    Dim Lo As ListObject

    Set upSh = ThisWorkbook.Worksheets("__updated")
    upId = LCase(Left(sh.Name, 4))
    For Each Lo In sh.ListObjects
        If sh.Name = "Analysis" Then _
        upId = LCase(Replace(Lo.Name, "Tab_", ""))
        Set upObj = UpdatedValues.Create(upSh, upId)
        upObj.AddColumns Lo
    Next
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
    sh.EnableCalculation = False
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
    sh.EnableCalculation = True
    NotBusyApp
End Sub

'Fire this event when leaving the dictionary
Public Sub EnterAnalysis()

    Dim dict As ILLdictionary
    Dim drop As IDropdownLists
    Dim lst As BetterArray
    Dim upObj As IUpdatedValues

    BusyApp

    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("Dictionary"), 5, 1)
    Set drop = DropdownLists.Create(ThisWorkbook.Worksheets("__variables"))
    Set upObj = UpdatedValues.Create(ThisWorkbook.Worksheets("__updated"), "dict")


    If upObj.IsUpdated("control_details") Or upObj.IsUpdated("variable_name") Then
        'Update geo vars
        Set lst = dict.GeoVars()
        On Error Resume Next
        drop.Update lst, "__geo_vars"
        On Error GoTo 0
        'Update choices vars
        Set lst = dict.ChoicesVars()
        On Error Resume Next
        drop.Update lst, "__choice_vars"
        On Error GoTo 0
    End If

    If upObj.IsUpdated("variable_type") Or upObj.IsUpdated("variable_name") Then
        'Update time vars
        Set lst = dict.TimeVars()
        On Error Resume Next
        drop.Update lst, "__time_vars"
        On Error GoTo 0
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

    Dim sh As Worksheet
    Dim csTab As ICustomTable
    Dim tsTab As ICustomTable
    Dim drop As IDropdownLists
    Dim dropArray As BetterArray
    Dim choi As Object
    Dim seriestitleRng As Range
    Dim colValue As String
    Dim choiceName As String
    Dim cellRng As Range
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim sumLab As String
    Dim pass As IPasswords
    Dim wb As Workbook

    Set wb = ThisWorkbook

    Set sh = wb.Worksheets("Analysis")
    Set csTab = CustomTable.Create(sh.ListObjects(LOBJNAME), "series title")
    Set seriestitleRng = csTab.DataRange("series title")
    Set pass = Passwords.Create(wb.Worksheets("__pass"))
    Set drop = DropdownLists.Create(wb.Worksheets("__variables"))

    If InterSect(Target, seriestitleRng) Is Nothing Then Exit Sub
    BusyApp
    pass.UnProtect "Analysis"

    'Create the choices object
    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("Dictionary"), 5, 1)
    Set vars = LLVariables.Create(dict)


    'Now get the value of column on the custom table and test it
    colValue = csTab.Value(colName:="column", keyName:=Target.Value)

    If colValue <> vbNullString Then

        choiceName = Application.WorksheetFunction.Trim(vars.Value(colName:="Control Details", varName:=colValue))

        'Test if it is a choice formula, if it is the case you get the categories by another way
        If (InStr(1, choiceName, "CHOICE_FORMULA") = 1) Then
            Set choi = ChoiceFormula.Create(choiceName)
            choiceName = choi.choiceName()
            Set dropArray = choi.Categories()
        Else
            Set choi = LLchoice.Create(ThisWorkbook.Worksheets("Choices"), 4, 1)
            Set dropArray = choi.Categories(choiceName)
        End If

        If dropArray.Length = 0 Then
            NotBusyApp
            Exit Sub
        End If

        drop.Add dropArray, choiceName & "__"
        drop.Update dropArray, choiceName & "__"

        'get the cell Range for choices
        Set cellRng = csTab.CellRange("choice", Target.Row)
        cellRng.Value = ""
        drop.SetValidation cellRng, choiceName & "__", ignoreBlank:=False
        FormatLockCell cellRng, False

        'get the cell Range for plot values or percentage
        Set cellRng = csTab.CellRange("values or percentages", Target.Row)
        drop.SetValidation cellRng, "__perc_val"
        FormatLockCell cellRng, False
    Else
        'Get the cellRang for choice
        Set cellRng = csTab.CellRange("choice", Target.Row)
        cellRng.Validation.Delete
        Set tsTab = CustomTable.Create(sh.ListObjects(LOBJTSNAME), "title")
        sumLab = tsTab.Value(colName:="summary label", keyName:=Target.Value)
        cellRng.Value = sumLab
        FormatLockCell cellRng, True

        Set cellRng = csTab.CellRange("values or percentages", Target.Row)
        cellRng.Validation.Delete
        cellRng.Value = "values"
        FormatLockCell cellRng, True
    End If

    pass.Protect "Analysis"
    NotBusyApp
End Sub

'Check update status when something changes in a range
Public Sub checkUpdateStatus(ByVal sh As Worksheet, ByVal Target As Range)
    Dim upSh As Worksheet
    Dim upObj As IUpdatedValues
    Dim upId As String
    Dim Lo As ListObject

    BusyApp

    Set upSh = ThisWorkbook.Worksheets("__updated")
    upId = LCase(Left(sh.Name, 4))
    If sh.Name = "Analysis" Then
        For Each Lo In sh.ListObjects
            upId = upId & "_" & LCase(Replace(Lo.Name, "Tab_", ""))
            Set upObj = UpdatedValues.Create(upSh, upId)
            upObj.CheckUpdate sh, Target
        Next
    Else
        Set upObj = UpdatedValues.Create(upSh, upId)
        upObj.CheckUpdate sh, Target
    End If

    NotBusyApp
End Sub

'Calculate columns of analysis worksheet
Public Sub CalculateAnalysis(ByVal sh As Worksheet)
    Dim rng As Range
    Dim csTab As ICustomTable

    'I prefer not declaring a table for
    BusyApp
    sh.Range("__ana_series_title_").Calculate
    Set csTab = CustomTable.Create(sh.ListObjects("Tab_Graph_TimeSeries"))
    Set rng = csTab.DataRange("graph id")
    rng.Calculate
    Set rng = csTab.DataRange("series id")
    rng.Calculate
    Set rng = csTab.DataRange("graph order")
    rng.Calculate
    Set rng = csTab.DataRange("row")
    rng.Calculate
    Set rng = csTab.DataRange("column")
    rng.Calculate

    NotBusyApp
End Sub
