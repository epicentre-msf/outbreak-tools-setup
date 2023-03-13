Attribute VB_Name = "CustomFunctions"

Option Explicit
'Custom functions for the setup

'Get the headers for the time series
Public Function TimeSeriesHeader(ByVal timeVar As String, ByVal grpVar As String, _
                                 ByVal sumLab As String) As String
    Application.Volatile


    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim sh As Worksheet
    Dim timeVarLab As String
    Dim colVarLab As String
    Dim header As String

    Set sh = ThisWorkbook.Worksheets("Dictionary")
    Set dict = LLdictionary.Create(sh, 5, 1)
    Set vars = LLVariables.Create(dict)

    timeVarLab = vars.Value(colName:="Main Label", varName:=timeVar)
    colVarLab = vars.Value(colName:="Main Label", varName:=grpVar)

    If (grpVar = vbNullString) Then
        header = sumLab & " " & ChrW(9472) & " " & timeVarLab
    Else
        header = sumLab & " " & ChrW(9472) & " " & timeVarLab & " " & ChrW(9472) & " " & colVarLab
    End If

    TimeSeriesHeader = header
End Function

'Graph Id, series Id and Graph order, Time variable, group by variable

Public Function GraphValue(ByVal graphTitle As String, Optional ByVal graphCol As String = "Graph ID") As String
    Application.Volatile

    Const LOBJNAME As String = "Tab_Label_TSGraph"
    Dim csTab As ICustomTable
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("Analysis")
    Set csTab = CustomTable.Create(sh.ListObjects(LOBJNAME), "Graph title")

    GraphValue = csTab.Value(colName:=graphCol, keyName:=graphTitle)
End Function


Public Function TSValue(ByVal tsTitle As String, Optional ByVal tsCol As String = "Series ID") As String
    Application.Volatile

    Const LOBJNAME As String = "Tab_TimeSeries_Analysis"
    Dim csTab As ICustomTable
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("Analysis")
    Set csTab = CustomTable.Create(sh.ListObjects(LOBJNAME), "Title")

    TSValue = csTab.Value(colName:=tsCol, keyName:=tsTitle)
End Function
