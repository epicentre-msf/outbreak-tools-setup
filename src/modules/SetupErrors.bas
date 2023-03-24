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

    Dim Check As IChecking
    Dim Lo As ListObject
    Dim csTab As ICustomTable
    Dim varRng As Range
    Dim sheetRng As Range
    Dim FUN As WorksheetFunction
    Dim varValue As String
    Dim sheetValue As String
    Dim sh As Worksheet
    Dim infoMessage As String
    Dim keyName As String
    Dim cellRng As Range
    Dim sortCols As BetterArray

    Set sh = wb.Worksheets(DICTSHEETNAME)
    Set Lo = sh.ListObjects(1)
    Set Check = Checking.Create(titleName:="Dictionary incoherences Type--Concerned Sheet--Incoherences")
    Set csTab = CustomTable.Create(Lo, idCol:="Variable Name")
    Set FUN = Application.WorksheetFunction
    Set sortCols = New BetterArray

    'Resize the dictionary table
    pass.UnProtect DICTSHEETNAME
    csTab.RemoveRows
    csTab.Sort "Sheet Name"

    Set varRng = csTab.DataRange("Variable Name")
    Set sheetRng = csTab.DataRange("Sheet Name")
    Set cellRng = varRng.Cells(varRng.Rows.Count, 1)

    'Errors on columns
    Do While cellRng.Row >= varRng.Row

        varValue = FUN.Trim(cellRng.Value)

        'Duplicates variable names
        If FUN.COUNTIF(varRng, varValue) > 1 Then
            keyName = "dict-var-unique"
            infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
            infoMessage = Replace(infoMessage, "{$$}", varValue)
            infoMessage = Replace(infoMessage, "{$}", cellRng.Row)
            Check.Add keyName & cellRng.Row, infoMessage, checkingError
        End If

        'Variabel lenths < 4
        If Len(varValue) < 4 Then
            keyName = "dict-var-length"
            infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
            infoMessage = Replace(infoMessage, "{$$}", varValue)
            infoMessage = Replace(infoMessage, "{$}", cellRng.Row)
            Check.Add keyName & cellRng.Row, infoMessage, checkingError
        End If

        sheetValue = sh.Cells(cellRng.Row, sheetRng.Column)
        
        'Empty sheet names
        If sheetValue = vbNullString Then
            keyName = "dict-empty-sheet"
            infoMessage = errTab.Value(colName:="Message", keyName:=keyName)
            infoMessage = Replace(infoMessage, "{$$$}", cellRng.Row)
            infoMessage = Replace(infoMessage, "{$}", cellRng.Row)
            'Var value
            infoMessage = Replace(infoMessage, "{$$}", varValue)
            Check.Add keyName & cellRng.Row, infoMessage, checkingError
        End If

        Set cellRng = cellRng.Offset(-1)
    Loop

    CheckTables.Push Check
    pass.Protect DICTSHEETNAME
End Sub


Private Sub CheckChoice()

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
