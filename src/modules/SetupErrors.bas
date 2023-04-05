Attribute VB_Name = "SetupErrors"
Option Explicit

'Module for checkings in the Setup file
'This is a long long long module.

Private Const DICTSHEETNAME As String = "Dictionary"
Private Const CHOICESHEETNAME As String = "Choices"
Private Const EXPORTSHEETNAME As String = "Exports"

Private CheckTables As BetterArray
Private wb As Workbook
Private errTab As ICustomTable 'Custom table for Error Messages
Private pass As IPasswords
Private dict As ILLdictionary
Private formData As IFormulaData
Private choi As ILLchoice

Private Sub Initialize()
    Dim shform As Worksheet

    'Initialize formula
    Set wb = ThisWorkbook
    Set shform = wb.Worksheets("__formula")
    Set formData = FormulaData.Create(shform)

    'Initialize the checking
    Set CheckTables = New BetterArray
    Set errTab = CustomTable.Create(shform.ListObjects("Tab_Error_Messages"), idCol:="Key")
    Set pass = Passwords.Create(wb.Worksheets("__pass"))
    Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEETNAME), 5, 1)
    Set choi = LLchoice.Create(wb.Worksheets(CHOICESHEETNAME), 4, 1)
End Sub

Private Function FormulaMessage(ByVal formValue As String, _
                                ByVal keyName As String, _
                                Optional value_one As String = vbNullString, _
                                Optional value_two As String = vbNullString, _
                                Optional ByVal formulaType As String = "linelist") As String
    Dim setupForm As IFormulas
    Set setupForm = Formulas.Create(dict, formData, formValue)
    If Not setupForm.Valid(formulaType:=formulaType) Then _
    FormulaMessage = ConvertedMessage(keyName, value_one, value_two, _
                                     setupForm.Reason())

End Function

Private Function ConvertedMessage(ByVal keyName As String, _
                                  Optional value_one As String = vbNullString, _
                                  Optional value_two As String = vbNullString, _
                                  Optional value_three As String = vbNullString, _
                                  Optional value_four As String = vbNullString) As String
    Dim infoMessage As String

    infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
    infoMessage = Replace(infoMessage, "{$}", value_one)
    infoMessage = Replace(infoMessage, "{$$}", value_two)
    infoMessage = Replace(infoMessage, "{$$$}", value_three)
    infoMessage = Replace(infoMessage, "{$$$$}", value_four)

    ConvertedMessage = infoMessage
End Function

Private Sub CheckDictionary()

    Dim Check As IChecking
    Dim csTab As ICustomTable
    Dim expTab As ICustomTable
    Dim varRng As Range
    Dim sheetRng As Range
    Dim FUN As WorksheetFunction
    Dim varValue As String
    Dim sheetValue As String
    Dim shdict As Worksheet
    Dim shexp As Worksheet
    Dim infoMessage As String
    Dim keyName As String
    Dim cellRng As Range
    Dim controlDetailsValue As String
    Dim controlValue As String
    Dim setupForm As Object
    Dim checkingCounter As Long 'Counter As 0 for each variable
    Dim choiCategories As BetterArray
    Dim formCategories As BetterArray
    Dim controlsList As BetterArray
    Dim choiName As String
    Dim tabCounter As Long
    Dim catValue As String 'category value when dealing with choice formulas
    Dim expCounter As Long 'Counter for each of the exports
    Dim expRng As Range    'Range for exports
    Dim expStatusRng As Range
    Dim minValue As String
    Dim maxValue As String

    Set shdict = wb.Worksheets(DICTSHEETNAME)
    Set shexp = wb.Worksheets(EXPORTSHEETNAME)
    Set Check = Checking.Create(titleName:="Dictionary incoherences Type--Where?--Details")
    Set csTab = CustomTable.Create(shdict.ListObjects(1), idCol:="Variable Name")
    Set expTab = CustomTable.Create(shexp.ListObjects(1), idCol:="Export Number")
    Set FUN = Application.WorksheetFunction

    Set choiCategories = New BetterArray
    Set formCategories = New BetterArray
    Set controlsList = New BetterArray

    ' Some preparation steps: Resize the dictionary table, sort on sheetNames
    pass.UnProtect DICTSHEETNAME
    pass.UnProtect EXPORTSHEETNAME

    csTab.RemoveRows
    csTab.Sort "Sheet Name"

    Set varRng = csTab.DataRange("Variable Name")
    Set sheetRng = csTab.DataRange("Sheet Name")
    Set cellRng = varRng.Cells(varRng.Rows.Count, 1)
    controlsList.Push "choice_manual", "choice_formula", "formula", _
                      "geo", "hf", "custom", "list_auto", "case_when"

    'Errors on columns
    Do While cellRng.Row >= varRng.Row
        checkingCounter = 0 'checkingCounter is just an id for errors and checkings
        varValue = FUN.Trim(cellRng.Value)

        'Duplicates variable names
        If FUN.COUNTIF(varRng, varValue) > 1 Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-var-unique"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Variabel lenths < 4
        If Len(varValue) < 4 Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-var-length"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        sheetValue = shdict.Cells(cellRng.Row, sheetRng.Column)

        'Empty sheet names
        If sheetValue = vbNullString Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-empty-sheet"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, varValue)
            Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        controlValue = csTab.Value("Control", varValue)
        controlDetailsValue = csTab.Value("Control Details", varValue)

        'Unkown control
        If (Not controlsList.Includes(controlValue)) And (controlValue <> vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-unknown-control"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, controlValue, varValue)
            Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Incoherences in choice formula
        If (controlValue = "choice_formula") And (InStr(1, controlDetailsValue, "CHOICE_FORMULA") = 1) Then
            'choice formula
            Set setupForm = ChoiceFormula.Create(controlDetailsValue)

            'Test if the choice_name exists
            choiName = setupForm.choiceName()
            If Not choi.ChoiceExists(choiName) Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-choiform-empty"
                infoMessage = ConvertedMessage(keyName, cellRng.Row, choiName)
                Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            Else
                keyName = "dict-catnotfound"
                Set choiCategories = choi.Categories(choiName)
                Set formCategories = setupForm.Categories()

                For tabCounter = formCategories.LowerBound To formCategories.UpperBound
                    catValue = formCategories.Item(tabCounter)
                    If Not choiCategories.Includes(catValue) Then
                        checkingCounter = checkingCounter + 1
                        infoMessage = ConvertedMessage(keyName, cellRng.Row, catValue, choiName)
                        Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingNote
                    End If
                Next
            End If
        End If

        'Incorrect formulas
        If (controlValue = "formula") Then
            keyName = "dict-incor-form"
            infoMessage = FormulaMessage(controlDetailsValue, keyName, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        'Choices not present in choice sheet
        If (controlValue = "choice_manual") Then
            If Not choi.ChoiceExists(controlDetailsValue) Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-choi-empty"
                infoMessage = ConvertedMessage(keyName, cellRng.Row, controlDetailsValue)
                Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        'Incorrect Min/Max formulas
        minValue = csTab.Value("Min", varValue)
        maxValue = csTab.Value("Max", varValue)

        If (minValue <> vbNullString) Then
            keyName = "dict-incor-min"
            infoMessage = FormulaMessage(minValue, keyName, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (maxValue <> vbNullString) Then
            keyName = "dict-incor-max"
            infoMessage = FormulaMessage(maxValue, keyName, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        Set cellRng = cellRng.Offset(-1)
    Loop

    'Exports Range
    expTab.Sort "Export Number"

    For expCounter = 1 To 5
        Set expRng = csTab.DataRange("Export " & expCounter)
        Set expStatusRng = expTab.DataRange("Status")
        If (Not IsEmpty(expRng)) And (expStatusRng.Cells(expCounter, 1).Value <> "active") Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-export-na"
            infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
            infoMessage = Replace(infoMessage, "{$}", expCounter)
            Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingNote
        End If
    Next

    CheckTables.Push Check
    pass.Protect DICTSHEETNAME
    pass.Protect EXPORTSHEETNAME
End Sub


Private Sub CheckChoice()
    ' Dim Check As IChecking
    ' Dim csTabdict As ICustomTable
    ' Dim csTab
    ' Dim varRng As Range
    ' Dim FUN As WorksheetFunction
    ' Dim varValue As String
    ' Dim sheetValue As String
    ' Dim shdict As Worksheet
    ' Dim shchoi As Worksheet
    ' Dim shform As Worksheet
    ' Dim infoMessage As String
    ' Dim keyName As String
    ' Dim cellRng As Range
    ' Dim sortCols As BetterArray
    ' Dim dict As ILLdictionary
    ' Dim formData As IFormulaData
    ' Dim controlDetailsValue As String
    ' Dim controlValue As String
    ' Dim setupForm As IFormulas
    ' Dim counter As Long 'Counter As 0 for each variable

    ' Set shdict = wb.Worksheets(DICTSHEETNAME)
    ' Set shform = wb.Worksheets(FORMULASHEETNAME)
    ' Set shchoi = wb.Worksheets(CHOICESHEETNAME)
    ' Set Check = Checking.Create(titleName:="Choices incoherences Type--Concerned Sheet--Incoherences")
    ' Set csTabdict = CustomTable.Create(Lo, idCol:="Variable Name")
    ' Set dict = LLdictionary.Create(sh, 5, 1)
    ' Set FUN = Application.WorksheetFunction
    ' Set sortCols = New BetterArray
    ' Set formData = FormulaData.Create(shform)


End Sub

Private Sub CheckAnalysis()

End Sub

Private Sub PrintReport()
    Const CHECKSHEETNAME As String = "__checkRep"

    Dim checKout As ICheckingOutput
    Dim sh As Worksheet

    Set sh = wb.Worksheets(CHECKSHEETNAME)
    Set checKout = CheckingOutput.Create(sh)

    checKout.PrintOutput CheckTables
End Sub


Public Sub CheckTheSetup()
    Initialize
    CheckDictionary
    CheckChoice
    CheckAnalysis
    PrintReport
End Sub
