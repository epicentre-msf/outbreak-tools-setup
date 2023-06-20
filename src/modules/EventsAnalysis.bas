Attribute VB_Name = "EventsAnalysis"

Option Explicit
'@Folder("Events")

'All events on the analysis Worksheet
Private Const ANALYSISSHEET As String = "Analysis"
Private Const LOBJNAME As String = "Tab_Graph_TimeSeries"
Private Const LOBJTSNAME As String = "Tab_TimeSeries_Analysis"
Private Const UPDATEDSHEETNAME As String = "__updated"
Private Const PASSSHEETNAME As String = "__pass"
Private Const DICTIONARYSHEET As String = "Dictionary"
Private Const DROPDOWNSHEET As String = "__variables"
Private Const CHOICESHEET As String = "Choices"

'speed up the application
Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
End Sub

'Calculate columns of analysis worksheet
Public Sub CalculateAnalysis()
    Dim csTab As ICustomTable
    Dim wb As Workbook
    Dim sh As Worksheet

    Set wb = ThisWorkbook

    On Error Resume Next 'If the datarange is nothing, proceed to next line
    Set sh = wb.Worksheets(ANALYSISSHEET)
    BusyApp
    sh.Range("__ana_series_title_").Calculate
    Set csTab = CustomTable.Create(sh.ListObjects("Tab_Graph_TimeSeries"))
    csTab.DataRange("graph id").Calculate
    csTab.DataRange("series id").Calculate
    csTab.DataRange("graph order").Calculate
    csTab.DataRange("row").Calculate
    csTab.DataRange("column").Calculate
    On Error GoTo 0
    
    NotBusyApp
End Sub

'When you enter the analysis sheet, update dropdown for time variables, etc.
'Fire this event when leaving the dictionary
Public Sub EnterAnalysis()

    Dim dict As ILLdictionary
    Dim drop As IDropdownLists
    Dim lst As BetterArray
    Dim upObj As IUpdatedValues
    Dim wb As Workbook

    BusyApp
    Set wb = ThisWorkbook

    Set dict = LLdictionary.Create(wb.Worksheets(DICTIONARYSHEET), 5, 1)
    Set drop = DropdownLists.Create(wb.Worksheets(DROPDOWNSHEET))
    Set upObj = UpdatedValues.Create(wb.Worksheets(UPDATEDSHEETNAME), "dict")

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

'Add choices dropdowns on the analysis worksheet
Public Sub AddChoicesDropdown(ByVal Target As Range)

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
    Set sh = wb.Worksheets(ANALYSISSHEET)
    Set csTab = CustomTable.Create(sh.ListObjects(LOBJNAME), "series title")
    Set seriestitleRng = csTab.DataRange("series title")
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set drop = DropdownLists.Create(wb.Worksheets(DROPDOWNSHEET))

    If Intersect(Target, seriestitleRng) Is Nothing Then Exit Sub
    If Target.Rows.Count > 1 Then Exit Sub

    BusyApp
    pass.UnProtect ANALYSISSHEET

    'Create the choices object
    Set dict = LLdictionary.Create(wb.Worksheets(DICTIONARYSHEET), 5, 1)
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
            Set choi = LLchoice.Create(wb.Worksheets(CHOICESHEET), 4, 1)
            Set dropArray = choi.Categories(choiceName)
        End If

        'If there are no categories, just exit, something went wrong somewhere
        If dropArray.Length = 0 Then
            NotBusyApp
            Exit Sub
        End If

        drop.Add dropArray, choiceName & "__"
        drop.Update dropArray, choiceName & "__"

        'get the cell Range for choices
        Set cellRng = csTab.CellRange("choice", Target.Row)
        cellRng.Value = vbNullString
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

    pass.Protect ANALYSISSHEET, True
    NotBusyApp
End Sub

Private Sub FormatLockCell(ByVal cellRng As Range, Optional ByVal Locked As Boolean = True)
    cellRng.Font.color = IIf(Locked, RGB(51, 142, 202), vbBlack)
    cellRng.Font.Italic = Locked
    cellRng.Locked = Locked
End Sub
