Attribute VB_Name = "SetupErrors"
Option Explicit

'Module for checkings in the Setup file
'This is a long long long module.

Private Const DICTSHEETNAME As String = "Dictionary"
Private Const CHOICESHEETNAME As String = "Choices"
Private Const EXPORTSHEETNAME As String = "Exports"
Private Const ANALYSISSHEETNAME As String = "Analysis"
Private Const TRANSLATIONSHEETNAME As String = "Translations"

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
                                Optional ByVal value_one As String = vbNullString, _
                                Optional ByVal value_two As String = vbNullString, _
                                Optional ByVal formulaType As String = "linelist") As String
    Dim setupForm As IFormulas
    Set setupForm = Formulas.Create(dict, formData, formValue)
    If Not setupForm.Valid(formulaType:=formulaType) Then _
    FormulaMessage = ConvertedMessage(keyName, value_one, value_two, _
                                     setupForm.Reason())

End Function

Private Function ConvertedMessage(ByVal keyName As String, _
                                  Optional ByVal value_one As String = vbNullString, _
                                  Optional ByVal value_two As String = vbNullString, _
                                  Optional ByVal value_three As String = vbNullString, _
                                  Optional ByVal value_four As String = vbNullString) As String
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
    Dim mainVarRng As Range
    Dim mainLabValue As String

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
    Set mainVarRng = csTab.DataRange("Main Label")
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

        'Empty variable name
        mainLabValue = shdict.Cells(cellRng.Row, mainVarRng.Column)

        If (mainLabValue = vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-main-lab"
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
                keyName = "dict-cat-notfound"
                Set choiCategories = choi.Categories(choiName)
                Set formCategories = setupForm.Categories()

                For tabCounter = formCategories.LowerBound To formCategories.UpperBound
                    catValue = formCategories.Item(tabCounter)
                    If Not choiCategories.Includes(catValue) Then
                        checkingCounter = checkingCounter + 1
                        infoMessage = ConvertedMessage(keyName, cellRng.Row, catValue, choiName)
                        Check.Add keyName & cellRng.Row & "-" & checkingCounter, infoMessage, checkingInfo
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
        If (FUN.CountBlank(expRng) <> expRng.Rows.Count) And (expStatusRng.Cells(expCounter, 1).Value <> "active") Then
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
    Dim Check As IChecking
    Dim shchoi As Worksheet
    Dim shdict As Worksheet
    Dim choiTab As ICustomTable
    Dim dictTab As ICustomTable
    Dim cntrlDetLst As BetterArray
    Dim choiLst As BetterArray
    Dim counter As Long
    Dim checkingCounter As Long
    Dim choiName As String
    Dim infoMessage As String
    Dim cellRng As Range
    Dim choiNameRng As Range
    Dim sortValue As String
    Dim choiLabValue As String
    Dim keyName As String

    Set shchoi = wb.Worksheets(CHOICESHEETNAME)
    Set shdict = wb.Worksheets(DICTSHEETNAME)
    Set choiTab = CustomTable.Create(shchoi.ListObjects(1))

    pass.UnProtect CHOICESHEETNAME
    'Sort the choices in choice sheet
    choi.Sort
    choiTab.RemoveRows

    Set Check = Checking.Create(titleName:="Choices incoherences Type--Where?--Details")
    
    Set dictTab = CustomTable.Create(shdict.ListObjects(1))
    Set choiLst = choi.AllChoices()
    Set cntrlDetLst = New BetterArray
    Set choiNameRng = choiTab.DataRange("List Name")
    Set cellRng = choiNameRng.Cells(choiNameRng.Rows.Count, 1)
    checkingCounter = 0

    cntrlDetLst.FromExcelRange dictTab.DataRange("Control Details")
    'choices not used
    For counter = choiLst.LowerBound To choiLst.UpperBound
        choiName = choiLst.Item(counter)
        If Not cntrlDetLst.Includes(choiName) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-unfound-choi"
            infoMessage = ConvertedMessage(keyName, choiName)

            Check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
        End If
    Next

    Do While cellRng.Row >= choiNameRng.Row
        choiName = cellRng.Value
        sortValue = cellRng.Offset(, 1).Value
        choiLabValue = cellRng.Offset(, 2).Value

        'Labels without choice name
        If (choiLabValue <> vbNullString) And (choiName = vbNullString) Then
            checkingCounter = checkingCounter + 1

            keyName = "choi-emptychoi-lab"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, choiLabValue)

            Check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Sort without choice name
        If (sortValue <> vbNullString) And (choiName = vbNullString) Then
            checkingCounter = checkingCounter + 1

            keyName = "choi-emptychoi-order"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, sortValue)

            Check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
        End If

        'Sort not filled
        If (sortValue = vbNullString) And (choiName <> vbNullString) Then
            checkingCounter = checkingCounter + 1

            keyName = "choi-empty-order"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, choiName)

            Check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
        End If

        'missing label for choice name (info)
        If (choiLabValue = vbNullString) And (choiName <> vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-mis-lab"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, choiName)

            Check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
        End If
        Set cellRng = cellRng.Offset(-1)
    Loop


    CheckTables.Push Check
    pass.Protect CHOICESHEETNAME
End Sub

'Checking on exports
Private Sub CheckExports()
    Dim expTab As ICustomTable
    Dim counter As Long
    Dim shexp As Worksheet
    Dim Check As IChecking
    Dim keyName As String
    Dim checkingCounter As Long
    Dim expStatus As String
    Dim cellRng As Range
    Dim infoMessage As String
    Dim keysLst As BetterArray
    Dim headersLst As BetterArray
    Dim headerCounter As Long
    Dim exportRng As Range
    Dim statusRng As Range
    Dim FUN As WorksheetFunction

    Set shexp = wb.Worksheets(EXPORTSHEETNAME)
    Set expTab = CustomTable.Create(shexp.ListObjects(1))
    Set keysLst = New BetterArray
    Set headersLst = New BetterArray
    Set statusRng = expTab.DataRange("Status")
    Set Check = Checking.Create(titleName:="Export incoherences type--Where?--Details")
    Set FUN = Application.WorksheetFunction

    headersLst.Push "Label Button", "Password", "Export Metadata", "Export Translation", _
                    "File Format", "File Name", "Export Header"
    keysLst.Push "exp-mis-lab", "exp-mis-pass", "exp-mis-meta", "exp-mis-trad", _
                 "exp-mis-form", "exp-mis-name", "exp-mis-head"

    checkingCounter = 0

    For counter = 1 To 5
        Set cellRng = expTab.CellRange("Status", counter + statusRng.Row - 1)
        expStatus = cellRng.Value
        For headerCounter = keysLst.LowerBound To keysLst.UpperBound
            'Empty label, password, metadata, translation file format or file name, file header
            If IsEmpty(expTab.CellRange(headersLst.Item(headerCounter), counter)) And (expStatus = "active") Then
                checkingCounter = checkingCounter + 1
                keyName = keysLst.Item(headerCounter)
                infoMessage = ConvertedMessage(keyName, cellRng.Row, counter)

                Check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
            End If
        Next

        'Active export not filled in the dictionary
        Set exportRng = dict.DataRange("Export " & counter, includeHeaders:=False)

        If (expStatus = "active") And (FUN.CountBlank(exportRng) = exportRng.Rows.Count) Then
            checkingCounter = checkingCounter + 1
            keyName = "exp-act-empty"
            infoMessage = ConvertedMessage(keyName, cellRng.Row, counter)

            Check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
        End If
    Next

    'Active export not filled in the dictionary
    CheckTables.Push Check
End Sub


'Checking on Translations
Private Sub CheckTranslations()
    Dim Lo As ListObject
    Dim shTrans As Worksheet
    Dim hRng As Range
    Dim messageMissing As String
    Dim nbMissing As Long
    Dim langName As String
    Dim Check As IChecking
    Dim counter As Long
    Dim colRng As Range


    Set shTrans = wb.Worksheets(TRANSLATIONSHEETNAME)
    Set Lo = shTrans.ListObjects(1)
    Set hRng = Lo.HeaderRowRange
    Set Check = Checking.Create(titleName:="Translation incoherences--Where?--Details")
    If (Not Lo.DataBodyRange Is Nothing) Then
        For counter = 1 To hRng.Columns.Count
            langName = hRng.Cells(1, counter).Value
            Set colRng = Lo.ListColumns(langName).DataBodyRange
            nbMissing = Application.WorksheetFunction.CountBlank(colRng)
            If nbMissing > 0 Then
                messageMissing = "Translations Sheet--" & nbMissing & _
                            " labels are missing for column " & _
                            langName & "."
                'Add the message to checkings
                Check.Add "trads-mis-labs-" & counter, messageMissing, checkingInfo
            End If
        Next
    End If

    CheckTables.Push Check
End Sub

'adding checks for analysis
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
    CheckExports
    CheckAnalysis
    CheckTranslations
    PrintReport
End Sub
