Attribute VB_Name = "PrepareSetup"

Option Explicit


'This module prepares the setup for usage and creates required elements for
'a fresh new setup without the codes for data management.

Private dropArray As BetterArray
Private drop As IDropdownLists
Private wb As Workbook
Private currSh As Worksheet
Private currTab As ICustomTable
Private pass As IPasswords

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
    Set currSh = wb.Worksheet(sheetName)
End Sub

Private Sub MoveToTable(ByVal tabName As String)
    Set currTab = CustomTable.Create(currSh.ListObjects(tabName))
End Sub

'Function to add Elements to the dropdown list
Private Sub AddElements(ByVal dropdownName As String, ParamArray Els() As Variant)
    Dim nbEls As Integer
    For nbEls = 0 To UBound(Els())
        dropArray.Push Els(nbEls)
    Next
    drop.Add dropArray, dropdownName
    dropArray.Clear
End Sub

Private Sub CreateDropdowns()

    'The first parameter or AddElements is the dropdown name, the others are
    'values to put in the dropdown

    'Stop events and calculations
    BusyApp

    'GLOBAL SETUP LEVEL --------------------------------------------------------
    '- yes_no dropdown
    AddElements "__yesno", "yes", "no"
    '- formats
    AddElements "__formats", "round0", "round1", "round2", "round3", _
                "percentage0", "percentage1", "percentage2", _
                "percentage3", "text", "euros", "dollars", _
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
                 "choice_formula", "formula", "geo", "hf", "custom", "list_auto", _
                 "case_when"
    '- alert
    AddElements "__var_alert", "error", "warning", "info"
    '- geo_variables
    AddElements "__geo_vars", "", ""
    '- choices_variables
    AddElements "__choice_vars", "", ""
    '- time_variables
    AddElements "__time_vars", "", ""

    'EXPORTS -------------------------------------------------------------------
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
    AddElements "__percentage_ta", "no", "row", "variable labels"
    '- percentage_vs_values
    AddElements "__perc_val", "percentages", "values"
    '- chart_type
    AddElements "__chart_type", "bar", "line", "point"
    '- axis_position
    AddElements "__axis_pos", "left", "right"
    '- swich between analysis tables
    AddElements "__swicth_tables", _
                "Add or remove rows of Global Summary", _
                "Add or remove rows of Univariate Analysis", _
                "Add or remove rows of Bivariate Analysis", _
                "Add or remove rows of Time Series Analysis", _
                "Add or remove rows of Graph on Time Series Labels", _
                "Add or remove rows of Graph on Time Series", _
                "Add or remove rows of Spatial Analysis", _
                "Add or remove rows of Spatio-Temporal Analysis", _
                "Add or remove rows of all tables"


    'Series and graphs titles
    AddElements "__graphs_titles", "", ""
    AddElements "__series_titles", "", ""

    'Return the state after completion
    NotBusyApp
End Sub

Private Sub AddValidations()

    'Dictionary dropdowns -----------------------------------------------------
    MoveToSheet "Dictionary"
    MoveToTable "Tab_Dictionary"

    'Set validation on dictionary colnames elements
    'sheet type
    currTab.SetValidation colName:="sheet type", dropName:="__sheet_type", _
                        drop:=drop, alertType:="error", pass:=pass
    'variable status
    currTab.SetValidation colName:="status", dropName:="__var_status", _
                        drop:=drop, alertType:="error", pass:=pass
    'personal identifier
    currTab.SetValidation colName:="personal identifier", dropName:="__yesno", _
                         drop:=drop, alertType:="error", pass:=pass
    'variable type
    currTab.SetValidation colName:="type", dropName:="__var_type", drop:=drop, _
                        alertType:="error", pass:=pass
    'variable format
    currTab.SetValidation colName:="format", dropName:="__formats", _
                        drop:=drop, alertType:="info", pass:=pass
    'variable control
    currTab.SetValidation colName:="control", dropName:="__var_control", _
                        drop:=drop, alertType:="info", pass:=pass
    'variable should be unique
    currTab.SetValidation colName:="unique", dropName:="__yesno", _
                        drop:=drop, alertType:="error", pass:=pass
    'Alert
    currTab.SetValidation colName:="alert", dropName:="__var_alert", _
                        drop:=drop, alertType:="error", pass:=pass
    'Lock cells on conditional formatting
    currTab.SetValidation colName:="lock cells", dropName:="__var_status", _
                        drop:=drop, alertType:="error", pass:=pass

    'Exports dropdowns -----------------------------------------------------------------------------------------
    MoveToSheet "Exports"
    MoveToTable "Tab_Exports"

    currTab.SetValidation colName:="password", dropName:="__yesno", _
                        drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="status", dropName:="__export_status", _
                        drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="export metadata", dropName:="__yesno", _
                        drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="export translation", dropName:="__yesno", _
                        drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="file format", dropName:="__export_format", _
                        drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="export header", dropName:="__export_header", _
                        drop:=drop, alertType:="error", pass:=pass

    'Analysis dropdowns ------------------------------------------------------------------------------------
    MoveToSheet "Analysis"

    'Global summary table
    MoveToTable "Tab_Global_Summary"
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                          alertType:="info", pass:=pass
    'Univariate analysis table
    MoveToTable "Tab_Univariate_Analysis"

    currTab.SetValidation colName:="add missing data", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                          alertType:="info", pass:=pass
    currTab.SetValidation colName:="add percentage", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="add graph", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="flip coordinates", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    'Group_by variable
    currTab.SetValidation colName:="row", dropName:="__choices_vars", drop:=drop, _
                          alertType:="error", pass:=pass

    'Bivariate analysis table
    MoveToTable "Tab_Bivariate_Analysis"
    currTab.SetValidation colName:="add missing data", dropName:="__missing_ba", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                          alertType:="info", pass:=pass
    currTab.SetValidation colName:="add percentage", dropName:="__percentage_ba", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="add Graph", dropName:="__perc_val", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="flip coordinates", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    'Row and columns groupby
    currTab.SetValidation colName:="row", dropName:="__choices_vars", drop:=drop, _
                          alertType:="error", pass:=pass
    currTab.SetValidation colName:="column", dropName:="__choices_vars", drop:=drop, _
                          alertType:="error", pass:=pass

    'Time Series analysis table
    MoveToTable "Tab_TimeSeries_Analysis"
    currTab.SetValidation colName:="add missing data", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="format", dropName:="__formats", drop:=drop, _
                          alertType:="info", pass:=pass
    currTab.SetValidation colName:="add percentage", dropName:="__percentage_ta", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="add total", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    'Row and columns groupby
    currTab.SetValidation colName:="row", dropName:="__time_vars", drop:=drop, _
                          alertType:="error", pass:=pass
    currTab.SetValidation colName:="column", dropName:="__choice_vars", drop:=drop, _
                          alertType:="error", pass:=pass

    'Graph on time series
    MoveToTable "Tab_Graph_TimeSeries"
    currTab.SetValidation colName:="plot values or percentages", _
                          dropName:="__perc_val", drop:=drop, _
                          alertType:="error", pass:=pass
    currTab.SetValidation colName:="chart type", dropName:="__chart_type", _
                          drop:=drop, alertType:="info", pass:=pass
    currTab.SetValidation colName:="y-axis", dropName:="__axis_pos", _
                          drop:=drop, alertType:="error", pass:=pass
    'graph title and series title
    currTab.SetValidation colName:="graph title", dropName:="__graphs_titles", _
                          alertType:="error", pass:=pass
    currTab.SetValidation colName:="series title", dropName:="__series_titles", _
                          drop:=drop, alertType:="error", pass:=pass
    'Spatial Analysis
    MoveToTable "Tab_Spatial_Analysis"

    currTab.SetValidation colName:="row", dropName:="__geo_vars", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="column", dropName:="__choice_vars", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="add missing data", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="add percentage", dropName:="__yesno", _
                          drop:=drop, alertType:="error", pass:=pass
    currTab.SetValidation colName:="add graph", dropName:="__perc_val", _
                        drop:=drop, alertType:="error", pass:=pass

    'Spatio-Temporal Analysis

End Sub

Public Sub ConfigureSetup()
    'Initialize elements
    Initialize
    CreateDropdowns 'Create dropdowns for the setup
    AddValidations  'Add the validations to each parts of the setup
End Sub
