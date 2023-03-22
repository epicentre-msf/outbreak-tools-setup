Attribute VB_Name = "PrepareSetup"

Option Explicit


'This module prepares the setup for usage and creates required elements for
'a fresh new setup without the codes for data management.

Private dropArray As BetterArray
Private drop As IDropdownLists
Private wb As Workbook
Private currsh As Worksheet
Private currTab As ICustomTable
Private currLo As ListObject
Private currUp As IUpdatedValues
Private pass As IPasswords

Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.CalculateBeforeSave = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
  End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub Initialize()
    Dim dropsh As Worksheet
    Set wb = ThisWorkbook
    Set dropsh = wb.Worksheets("__variables")
    'Initilialize the dropdown array and list
    Set dropArray = New BetterArray
    Set drop = DropdownLists.Create(dropsh)
    Set pass = Passwords.Create(wb.Worksheets("__pass"))
End Sub

Private Sub MoveToSheet(ByVal sheetName As String)
    Set currsh = wb.Worksheets(sheetName)
End Sub

Private Sub MoveToUp(ByVal upName As String)
    Set currUp = UpdatedValues.Create(wb.Worksheets("__updated"), upName)
End Sub

Private Sub MoveToTable(ByVal tabName As String)
    Set currLo = currsh.ListObjects(tabName)
    Set currTab = CustomTable.Create(currLo)
End Sub


'Function to add Elements to the dropdown list
Private Sub AddElements(ByVal dropdownName As String, ParamArray els() As Variant)
    Dim nbEls As Integer
    For nbEls = 0 To UBound(els())
        dropArray.Push els(nbEls)
    Next
    drop.Add dropArray, dropdownName
    dropArray.Clear
End Sub

Private Sub CreateDropdowns()

    'The first parameter or AddElements is the dropdown name, the others are
    'values to put in the dropdown

    'GLOBAL SETUP LEVEL --------------------------------------------------------
    '- yes_no dropdown
    AddElements "__yesno", "yes", "no"
    '- formats
    AddElements "__formats", "integer", "round0", "round1", _
                "round2", "round3", "percentage0", "percentage1", _
                "percentage2", "percentage3", "text", "euros", "dollars", _
                "dd/mm/yyyy", "d-mmm-yyyy", ""

    'DICTIONARY ----------------------------------------------------------------
    ' - variable status
    AddElements "__var_status", "mandatory", "optional", "hidden"
    '- variable_type
    AddElements "__var_type", "date", "integer", "text", "decimal"
    '- sheet_type
    AddElements "__sheet_type", "vlist1D", "hlist2D"
    '- control
    AddElements "__var_control", "choice_manual", _
                 "choice_formula", "formula", "geo", "hf", "custom", _
                 "list_auto", "case_when"
    '- alert
    AddElements "__var_alert", "error", "warning", "info"
    '- geo_variables
    AddElements "__geo_vars", "", ""
    '- choices_variables
    AddElements "__choice_vars", "", ""
    '- time_variables
    AddElements "__time_vars", "", ""

    'EXPORTS ------------------------------------------ -------------------------
    '- export_status
    AddElements "__export_status", "active", "inactive"
    '- export_format
    AddElements "__export_format", "xlsx", "xlsb"
    '- export_headers
    AddElements "__export_header", "variable names", "variable labels"

    'ANALYSIS ------------------------------------------------------------------
    '- percentage_ba
    AddElements "__percentage_ba", "no", "row", "column", "total"
    '- missing_ba
    AddElements "__missing_ba", "no", "row", "column", "all"
    '- percentage_ta
    AddElements "__percentage_ta", "no", "row", "column"
    '- percentage_vs_values
    AddElements "__perc_val", "percentages", "values"
    '- chart_type
    AddElements "__chart_type", "bar", "line", "point"
    '- axis_position
    AddElements "__axis_pos", "left", "right"
    '- swich between analysis tables
    AddElements "__swicth_tables", _
                "Add or remove rows of global summary", _
                "Add or remove rows of univariate analysis", _
                "Add or remove rows of bivariate analysis", _
                "Add or remove rows of time series analysis", _
                "Add or remove rows of labels for time series graphs", _
                "Add or remove rows of graph on time series", _
                "Add or remove rows of spatial analysis", _
                "Add or remove rows of spatio-temporal analysis", _
                "Add or remove rows of all tables"


    'Series and graphs titles
    AddElements "__graphs_titles", "", ""
    AddElements "__series_titles", "", ""

End Sub

Private Sub AddValidationsAndUpdates()

    'Dictionary dropdowns -----------------------------------------------------
    MoveToSheet "Dictionary"
    BusyApp
    pass.UnProtect "Dictionary"
    BusyApp
    MoveToTable "Tab_Dictionary"

    'Set validation on dictionary colnames elements
    'sheet type

    currTab.SetValidation colName:="sheet type", dropName:="__sheet_type", _
                        drop:=drop, alertType:="error"
    'variable status
    currTab.SetValidation colName:="status", dropName:="__var_status", _
                        drop:=drop, alertType:="error"
    'personal identifier
    currTab.SetValidation colName:="personal identifier", dropName:="__yesno", _
                         drop:=drop, alertType:="error"
    'variable type
    currTab.SetValidation colName:="variable type", dropName:="__var_type", drop:=drop, _
                        alertType:="error"
    'variable format
    currTab.SetValidation colName:="variable format", dropName:="__formats", _
                        drop:=drop, alertType:="info"
    'variable control
    currTab.SetValidation colName:="control", dropName:="__var_control", _
                        drop:=drop, alertType:="info"

    'print variable (add the variable to a print sheet)
    currTab.SetValidation colName:="print variable", dropName:="__yesno", _
                         drop:=drop, alertType:="info"
    'variable should be unique
    currTab.SetValidation colName:="unique", dropName:="__yesno", _
                        drop:=drop, alertType:="error"
    'Alert
    currTab.SetValidation colName:="alert", dropName:="__var_alert", _
                        drop:=drop, alertType:="error"
    'Lock cells on conditional formatting
    currTab.SetValidation colName:="lock cells", dropName:="__yesno", _
                        drop:=drop, alertType:="error"


    'Add watchers on columns
    MoveToUp "dict"
    currUp.AddColumns currLo
    BusyApp
    pass.Protect "Dictionary"
    BusyApp

    'Choices worksheet -----------------------------------------------------------------------------------------

    MoveToSheet "Choices"
    MoveToTable "Tab_Choices"
    MoveToUp "choi"
    currUp.AddColumns currLo

    'Exports dropdowns -----------------------------------------------------------------------------------------
    MoveToSheet "Exports"
    BusyApp
    pass.UnProtect "Exports"
    BusyApp
    MoveToTable "Tab_Export"

    currTab.SetValidation colName:="password", dropName:="__yesno", _
                        drop:=drop, alertType:="error"
    currTab.SetValidation colName:="status", dropName:="__export_status", _
                        drop:=drop, alertType:="error"
    currTab.SetValidation colName:="export metadata", dropName:="__yesno", _
                        drop:=drop, alertType:="error"
    currTab.SetValidation colName:="export translation", dropName:="__yesno", _
                        drop:=drop, alertType:="error"
    currTab.SetValidation colName:="file format", dropName:="__export_format", _
                        drop:=drop, alertType:="error"
    currTab.SetValidation colName:="export header", dropName:="__export_header", _
                        drop:=drop, alertType:="error"

    'Add Watchers on columns
    MoveToUp "expo"
    currUp.AddColumns currLo
    BusyApp
    pass.Protect "Exports"
    BusyApp

    'Analysis dropdowns ------------------------------------------------------------------------------------
    MoveToSheet "Analysis"
    BusyApp
    pass.UnProtect "Analysis"
    BusyApp

    'add validation on select table
    drop.SetValidation cellRng:=currsh.Range("RNG_SelectTable"), _
                       listName:="__swicth_tables", alertType:="error"

    'Global summary table
    MoveToTable "Tab_Global_Summary"
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                          alertType:="info"

    'columns to watch on global summary
    MoveToUp "global_summary"
    currUp.AddColumns currLo

    'Univariate analysis table
    MoveToTable "Tab_Univariate_Analysis"

    currTab.SetValidation colName:="add missing data", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                          alertType:="info"
    currTab.SetValidation colName:="add percentage", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="add graph", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="flip coordinates", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    'Group_by variable
    currTab.SetValidation colName:="row", dropName:="__choice_vars", drop:=drop, _
                          alertType:="error"

    'add columns
    MoveToUp "univariate_analysis"
    currUp.AddColumns currLo

    'Bivariate analysis table
    MoveToTable "Tab_Bivariate_Analysis"
    currTab.SetValidation colName:="add missing data", dropName:="__missing_ba", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                          alertType:="info"
    currTab.SetValidation colName:="add percentage", dropName:="__percentage_ba", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="add Graph", dropName:="__perc_val", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="flip coordinates", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="flip coordinates", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    'Row and columns groupby
    currTab.SetValidation colName:="row", dropName:="__choice_vars", drop:=drop, _
                          alertType:="error"
    currTab.SetValidation colName:="column", dropName:="__choice_vars", drop:=drop, _
                          alertType:="error"

    'add columns on bivariate analysis to watcher
    MoveToUp "bivariate_analysis"
    currUp.AddColumns currLo

    'Time Series analysis table
    MoveToTable "Tab_TimeSeries_Analysis"
    currTab.SetValidation colName:="add missing data", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                          alertType:="info"
    currTab.SetValidation colName:="add percentage", dropName:="__percentage_ta", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="add total", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    'Row and columns groupby
    currTab.SetValidation colName:="row", dropName:="__time_vars", drop:=drop, _
                          alertType:="error"
    'column group by is not mandatory for time series tables
    currTab.SetValidation colName:="column", dropName:="__choice_vars", drop:=drop, _
                          alertType:="info"

    'add column on time series analysis to watcher
    MoveToUp "timeseries_analysis"
    currUp.AddColumns currLo

    'add columns on graph on time series labels
    MoveToTable "Tab_Label_TSGraph"
    MoveToUp "label_tsgraph"
    currUp.AddColumns currLo

    'Graph on time series
    MoveToTable "Tab_Graph_TimeSeries"
    currTab.SetValidation colName:="plot values or percentages", _
                          dropName:="__perc_val", drop:=drop, _
                          alertType:="error"
    currTab.SetValidation colName:="chart type", dropName:="__chart_type", _
                          drop:=drop, alertType:="info"
    currTab.SetValidation colName:="y-axis", dropName:="__axis_pos", _
                          drop:=drop, alertType:="error"
    MoveToUp "graph_timeseries"
    currUp.AddColumns currLo

    'graph title and series title
    'Spatial Analysis
    MoveToTable "Tab_Spatial_Analysis"

    currTab.SetValidation colName:="row", dropName:="__geo_vars", _
                          drop:=drop, alertType:="error"
    'On spatial analysis column variables are not mandatory
    currTab.SetValidation colName:="column", dropName:="__choice_vars", _
                          drop:=drop, alertType:="info"
    currTab.SetValidation colName:="add missing data", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="add percentage", dropName:="__yesno", _
                          drop:=drop, alertType:="error"
    currTab.SetValidation colName:="add graph", dropName:="__perc_val", _
                        drop:=drop, alertType:="error"
    currTab.SetValidation colName:="flip coordinates", dropName:="__yesno", _
                        drop:=drop, alertType:="error"
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                        alertType:="info"

    MoveToUp "spatial_analysis"
    currUp.AddColumns currLo

    'Spatio-Temporal Analysis
    BusyApp
    pass.Protect "Analysis"
    BusyApp
End Sub

Public Sub ConfigureSetup()
    'Initialize elements
    BusyApp
    Initialize
    CreateDropdowns 'Create dropdowns for the setup
    AddValidationsAndUpdates  'Add the validations to each parts of the setup
    MsgBox "Done!"
End Sub

'Prepare the setup for production
Public Sub PrepareForProd()
    Dim wb As Workbook
    Dim pass As IPasswords
    Dim pwd As String
    Dim sh As Worksheet

    Set wb = ThisWorkbook
    'First write the password to the password sheet
    pwd = wb.Worksheets("Dev").Range("RNG_DevPasswd").Value
    wb.Worksheets("__pass").Range("RNG_DebuggingPassword").Value = pwd

    'Protect the worksheets
    Set sh = wb.Worksheets("__pass")
    Set pass = Passwords.Create(sh)
    'As Dictionary
    pass.Protect "Dictionary"
    'Choices
    pass.Protect "Choices"
    'Translations
    pass.Protect "Translations", True, True
    'Analysis
    pass.Protect "Analysis", True
    'Exports
    pass.Protect "Exports"
    'Hide some worksheeets
    pass.UnProtectWkb wb
    wb.Worksheets("__updated").Visible = xlSheetHidden
    wb.Worksheets("__pass").Visible = xlSheetHidden
    wb.Worksheets("__variables").Visible = xlSheetHidden
    wb.Worksheets("Dev").Visible = xlSheetHidden

    'Protect the workbook
    pass.ProtectWkb wb

    'Protect the project

End Sub
