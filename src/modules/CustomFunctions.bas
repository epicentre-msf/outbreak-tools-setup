Attribute VB_Name = "CustomFunctions"
'@IgnoreModule SheetAccessedUsingString

Option Explicit
'Custom functions for the setup

'@Description("Get the headers for the time series")
'@EntryPoint : Time series headers are used in the analysis worksheet
Public Function TimeSeriesHeader(ByVal timeVar As String, ByVal grpVar As String, _
                                 ByVal sumLab As String) As String
Attribute TimeSeriesHeader.VB_Description = "Get the headers for the time series"
    Application.Volatile

    Dim sh As Worksheet
    Dim csTab As ICustomTable
    Dim timeVarLab As String
    Dim colVarLab As String
    Dim header As String

    Set sh = ThisWorkbook.Worksheets("Dictionary")
    Set csTab = CustomTable.Create(sh.ListObjects(1), "variable name")

    timeVarLab = csTab.Value(colName:="Main Label", keyName:=timeVar)
    colVarLab = csTab.Value(colName:="Main Label", keyName:=grpVar)

    If (grpVar = vbNullString) Then
        header = sumLab & " " & ChrW(9472) & " " & timeVarLab
    Else
        header = sumLab & " " & ChrW(9472) & " " & timeVarLab & " " & ChrW(9472) & " " & colVarLab
    End If

    TimeSeriesHeader = header
End Function

'@Description("Get a graph value from the label on graph table")
'@EntryPoint : The function GraphValue is used only on analysis sheet/graph table
Public Function GraphValue(ByVal graphTitle As String, Optional ByVal graphCol As String = "Graph ID") As String
Attribute GraphValue.VB_Description = "Get a graph value from the label on graph table"
    Application.Volatile

    Const LOBJNAME As String = "Tab_Label_TSGraph"
    Dim csTab As ICustomTable
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("Analysis")
    Set csTab = CustomTable.Create(sh.ListObjects(LOBJNAME), "Graph title")

    GraphValue = csTab.Value(colName:=graphCol, keyName:=graphTitle)
End Function


'@Description("Get a time series value from the time series table")
'@EntryPoint : TSValue is used only on analysis sheet/ graph table
Public Function TSValue(ByVal tsTitle As String, Optional ByVal tsCol As String = "Series ID") As String
Attribute TSValue.VB_Description = "Get a time series value from the time series table"
    Application.Volatile

    Const LOBJNAME As String = "Tab_TimeSeries_Analysis"
    Dim csTab As ICustomTable
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("Analysis")
    Set csTab = CustomTable.Create(sh.ListObjects(LOBJNAME), "Title")

    TSValue = csTab.Value(colName:=tsCol, keyName:=tsTitle)
End Function
