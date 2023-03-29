Attribute VB_Name = "SetupErrors"
Option Explicit

'Module for checkings in the Setup file
'This is a long long long module.

Private CheckTables As BetterArray
Private wb As Workbook
Private errTab As ICustomTable 'Custom table for Error Messages
Private formSh As Worksheet 'Formula worksheet
Private pass As IPasswords

Private Sub Initialize()
    Set wb = ThisWorkbook
    Set formSh = wb.Worksheets("__formula")

    'Initialize the checking
    Set CheckTables = New BetterArray
    Set errTab = CustomTable.Create(formSh.ListObjects("Tab_Error_Messages"), idCol:="Key")
    Set pass = Passwords.Create(wb.Worksheets("__pass"))
End Sub

Private Sub CheckDictionary()
    Const DICTSHEETNAME As String = "Dictionary"
    Const FORMULASHEETNAME As String = "__formula"

    Dim Check As IChecking
    Dim Lo As ListObject
    Dim csTab As ICustomTable
    Dim varRng As Range
    Dim sheetRng As Range
    Dim FUN As WorksheetFunction
    Dim varValue As String
    Dim sheetValue As String
    Dim sh As Worksheet
    Dim shform As Worksheet
    Dim infoMessage As String
    Dim keyName As String
    Dim cellRng As Range
    Dim dict As ILLdictionary
    Dim formData As IFormulaData
    Dim controlDetailsValue As String
    Dim controlValue As String
    Dim setupForm As IFormulas
    Dim counter As Long 'Counter As 0 for each variable

    Set sh = wb.Worksheets(DICTSHEETNAME)
    Set shform = wb.Worksheets(FORMULASHEETNAME)
    Set Lo = sh.ListObjects(1)
    Set Check = Checking.Create(titleName:="Dictionary incoherences Type--Concerned Sheet--Incoherences")
    Set csTab = CustomTable.Create(Lo, idCol:="Variable Name")
    Set dict = LLdictionary.Create(sh, 5, 1)
    Set FUN = Application.WorksheetFunction
    Set formData = FormulaData.Create(shform)

    'Resize the dictionary table
    pass.UnProtect DICTSHEETNAME
    csTab.RemoveRows
    csTab.Sort "Sheet Name"

    Set varRng = csTab.DataRange("Variable Name")
    Set sheetRng = csTab.DataRange("Sheet Name")
    Set cellRng = varRng.Cells(varRng.Rows.Count, 1)

    'Errors on columns
    Do While cellRng.Row >= varRng.Row
        counter = 0
        varValue = FUN.Trim(cellRng.Value)

        'Duplicates variable names
        If FUN.COUNTIF(varRng, varValue) > 1 Then
            counter = counter + 1
            keyName = "dict-var-unique"
            infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
            infoMessage = Replace(infoMessage, "{$$}", varValue)
            infoMessage = Replace(infoMessage, "{$}", cellRng.Row)
            Check.Add keyName & cellRng.Row & "-" & counter, infoMessage, checkingError
        End If

        'Variabel lenths < 4
        If Len(varValue) < 4 Then
            counter = counter + 1
            keyName = "dict-var-length"
            infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
            infoMessage = Replace(infoMessage, "{$$}", varValue)
            infoMessage = Replace(infoMessage, "{$}", cellRng.Row)
            Check.Add keyName & cellRng.Row & "-" & counter, infoMessage, checkingError
        End If

        sheetValue = sh.Cells(cellRng.Row, sheetRng.Column)

        'Empty sheet names
        If sheetValue = vbNullString Then
            counter = counter + 1
            keyName = "dict-empty-sheet"
            infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
            infoMessage = Replace(infoMessage, "{$$$}", cellRng.Row)
            infoMessage = Replace(infoMessage, "{$}", cellRng.Row)
            'Var value
            infoMessage = Replace(infoMessage, "{$$}", varValue)
            Check.Add keyName & cellRng.Row & "-" & counter, infoMessage, checkingError
        End If

        'Incorrect formulas
        controlValue = csTab.Value("Control", varValue)
        controlDetailsValue = csTab.Value("Control Details", varValue)

        If (controlValue = "formula") Then
            counter = counter + 1
            Set setupForm = Formulas.Create(dict, formData, controlDetailsValue)
            If Not setupForm.Valid(formulaType:="linelist") Then
                keyName = "dict-incor-form"
                infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
                infoMessage = Replace(infoMessage, "{$}", cellRng.Row)
                infoMessage = Replace(infoMessage, "{$$}", controlDetailsValue)
                infoMessage = Replace(infoMessage, "{$$$}", varValue)
                infoMessage = Replace(infoMessage, "{$$$$}", setupForm.Reason())

                Check.Add keyName & cellRng.Row & "-" & counter, infoMessage, checkingWarning
            End If
        End If

        Set cellRng = cellRng.Offset(-1)
    Loop

    CheckTables.Push Check
    pass.Protect DICTSHEETNAME
End Sub


Private Sub CheckChoice()
    ' Const DICTSHEETNAME As String = "Dictionary"
    ' Const FORMULASHEETNAME As String = "__formula"
    ' Const CHOICESHEETNAME As String = "Choices"

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
