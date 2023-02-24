Attribute VB_Name = "Helpers"
Option Explicit

'This module contains functions used to control validations on graphs
'for time series.

'Test if a listobject exists
Public Function ListObjectExists(Wksh As Worksheet, sListObjectName As String) As Boolean
    ListObjectExists = False
    Dim Lo As ListObject
    On Error Resume Next
    Set Lo = Wksh.ListObjects(sListObjectName)
    ListObjectExists = (Not Lo Is Nothing)
    On Error GoTo 0
End Function

'Set a validation on a range
Sub SetValidation(oRange As Range, sValidList As String, sAlertType As Byte, Optional sMessage As String = vbNullString)

    With oRange.validation
        .Delete
        Select Case sAlertType
        Case 1                                   '"error"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sValidList
        Case 2                                   '"warning"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=sValidList
        Case Else                                'for all the others, add an information alert
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=sValidList
        End Select

        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .errorTitle = ""
        .InputMessage = ""
        .errorMessage = sMessage
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'Update the list of choices with choices values

Sub UpdateChoicesList(sChoice As String, choiceStart As Long, iChoiceCol As Integer, Optional addTotal As String = C_sNo)

    Dim listCounter As Long
    Dim choiceCounter As Long

    Const LabelColumn As Byte = 3

    listCounter = 2
    choiceCounter = choiceStart


    Do While SheetChoice.Cells(choiceCounter, 1) = sChoice
        sheetChoicesLists.Cells(listCounter, iChoiceCol).Value = SheetChoice.Cells(choiceCounter, LabelColumn).Value

        listCounter = listCounter + 1
        choiceCounter = choiceCounter + 1
    Loop

    If addTotal <> C_sNo Then
        sheetChoicesLists.Cells(listCounter, iChoiceCol).Value = C_sTotal
    End If
End Sub

'AddValidation'Add the choices
Sub AddChoices(sVarName As String, choicesAnalysisRow As Long, Optional addTotal As String = C_sNo)

    Dim iChoiceCol As Integer
    Dim iChoiceRow As Long
    Dim choiceStart As Long
    Dim sListObjectName As String
    Dim sChoice As String
    Dim LoRng As Range
    Dim choiRng As Range
    Dim varRng As Range
    Dim varRow As Long
    Dim listRng As Range
    Dim namesCol As Long
    Dim namesRow As Long

    Const choicesColumn As Byte = 14 'choice column in the dictionary (control details)
    Const choicesAnalysisCol As Byte = 8 'choice column in the analysis table

    'Range of the variable name column
    Set varRng = sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(1).DataBodyRange

    'Range of the choice column
    Set choiRng = SheetChoice.ListObjects(C_sTabChoices).ListColumns(1).DataBodyRange

    'If you can't find the variable name, just exist
    If varRng.Find(What:=sVarName, LookAt:=xlWhole, MatchCase:=True) Is Nothing Then Exit Sub
    varRow = varRng.Find(What:=sVarName, LookAt:=xlWhole, MatchCase:=True).Row

    'Get the choice corresponding to the variable
    sChoice = sheetDictionary.Cells(varRow, choicesColumn)

    'If the choice is a choice_formula, I need to extract the choice name from the formula
    If sheetDictionary.Cells(varRow, choicesColumn - 1).Value = "choice_formula" Then
        sChoice = Replace(Split(sChoice, ",")(0), "CHOICE_FORMULA(", "")
    End If

    'If you can't find the corresponding choice in the choice sheet, do nothing and just exit
    If choiRng.Find(What:=sChoice, LookAt:=xlWhole, MatchCase:=True) Is Nothing Then Exit Sub

    choiceStart = choiRng.Find(What:=sChoice, LookAt:=xlWhole, MatchCase:=True).Row
    sListObjectName = "lo" & "_" & sChoice

    'Add the list object if it does not exists
    With sheetChoicesLists

        If ListObjectExists(sheetChoicesLists, sListObjectName) Then
            'If the list object exists, convert to ranges and restart the process (maybe some updates have been made)
            iChoiceCol = .ListObjects(sListObjectName).Range.Column
            .ListObjects(sListObjectName).DataBodyRange.Clear
            .ListObjects(sListObjectName).Unlist
        Else
            iChoiceCol = .Cells(1, .Columns.Count).End(xlToLeft).Column + 2
        End If

        'Write the list of the choices
        .Cells(1, iChoiceCol).Value = sChoice
        UpdateChoicesList sChoice:=sChoice, choiceStart:=choiceStart, iChoiceCol:=iChoiceCol, addTotal:=addTotal

        'Add the list object to the list worksheet
        iChoiceRow = .Cells(.Rows.Count, iChoiceCol).End(xlUp).Row

        Set LoRng = .Range(.Cells(1, iChoiceCol), .Cells(iChoiceRow, iChoiceCol))

        'Add the list object here
        .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = sListObjectName

        'Add dynamic name for the choice
        'First delete the name if it exists
        On Error Resume Next
            ThisWorkbook.NAMES(sChoice).Delete
        On Error GoTo 0

        ThisWorkbook.NAMES.Add Name:=sChoice, RefersToR1C1:="=" & sListObjectName & "[" & sChoice & "]"
        'The listobject already exists, we will only focus on updating the choice
        With sheetLists
            Set listRng = .ListObjects(C_sTabListOfChoicesNames).Range
            namesCol = listRng.Column

             If listRng.Find(What:=sChoice, LookAt:=xlWhole, MatchCase:=True) Is Nothing Then
                namesRow = IIf(IsEmpty(.Cells(listRng.Rows.Count, namesCol)), listRng.Rows.Count, listRng.Rows.Count + 1)
                .Cells(namesRow, namesCol).Value = sChoice
            End If

        End With
    End With

    'Add the validation to the choice
    SetValidation oRange:=sheetAnalysis.Cells(choicesAnalysisRow, choicesAnalysisCol), sValidList:="=" & sChoice, sAlertType:=1

End Sub


Public Function Updated(ByVal rngName As String) As Boolean
    Updated = (sheetLists.Range(rngName).Value = "Updated")
End Function

Public Sub UpdateValue(ByVal Condition As Boolean, ByVal rngName As String)
    If Condition Then
        sheetLists.Range(rngName).Value = "Updated"
    Else
        sheetLists.Range(rngName).Value = "Not Updated"
    End If
End Sub


'Time Series Header
Public Function TimeSeriesHeader(ByVal timeVar As String, ByVal colVar As String, ByVal labelValue As String) As String

    Application.Volatile

    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim timeVarLabel As String
    Dim colVarLabel As String
    Dim headerLabel As String

    On Error GoTo Err

    'Get the label of the time variable in the dictionary
    Set dict = LLdictionary.Create(sheetDictionary, 4, 1)
    Set vars = LLVariables.Create(dict)
    timeVarLabel = vars.Value(varName:=timeVar, colName:="Main Label")
    colVarLabel = vars.Value(varName:=colVar, colName:="Main Label")

    If timeVarLabel <> vbNullString Then
        If colVar = vbNullString Then
            headerLabel = labelValue & " v.s " & timeVarLabel
        Else
            headerLabel = colVarLabel & " v.s " & timeVarLabel & " (" & labelValue & ")"
        End If
    End If

Err:
    TimeSeriesHeader = headerLabel
End Function


