Attribute VB_Name = "PrepareSetup"

Option Explicit


'This module prepares the setup for usage and creates required elements for
'a fresh new setup without the codes for data management.

Private dropArray As BetterArray
Private drop As IDropdownLists

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
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set dropsh = wb.Worksheets("__variables")
    'Initilialize the dropdown array and list
    Set dropArray = New BetterArray
    Set drop = DropdownLists.Create(dropsh)
End Sub

'Function to add Elements to the dropdown list
Private Sub AddElements(ByVal dropdownName As String, ParamArray Els() As Variant)
    Initialize
    
    Dim nbEls As Integer
    For nbEls = 0 To UBound(Els())
        dropArray.Push Els(nbEls)
    Next
    drop.Add dropArray, dropdownName
    dropArray.Clear
End Sub

Public Sub ConfigureSetup()

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
    AddElements "var_type", "date", "integer", "text", "decimal"
    '- sheet_type
    AddElements "__sheet_type", "vlist1D", "hlist2D"
    '- control
    AddElements "__var_control", "choice_manual", _
                 "choice_formula", "formula", "geo", "hf", "custom", "list_auto", _
                 "case_when"
    '- alert
    AddElements "__alert", "error", "warning", "info"
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
    AddElements "__export_headers", "variable names", "variable labels"

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
                "Add or remove rows to Graph on Time Series Labels", _
                "Add or remove rows to Graph on Time Series", _
                "Add or remove rows of Spatial Analysis", _
                "Add or remove rows of Spatio-Temporal Analysis", _
                "Add or remove rows of all tables"

    'Return the state after completion
    NotBusyApp
End Sub
