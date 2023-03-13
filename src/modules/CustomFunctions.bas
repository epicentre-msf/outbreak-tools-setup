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

