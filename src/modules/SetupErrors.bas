Attribute VB_Name = "SetupErrors"
Option Explicit

'Module for checkings in the Setup file

Private CheckTables As BetterArray
Private wb As Workbook

Private Sub Initialize()
    Set wb = ThisWorkbook
    'Initialize the checking
    Set CheckTables = New BetterArray
End Sub

Private Sub CheckDictionary()
    Const DICTSHEETNAME As String = "Dictionary"

    Dim Check As IChecking
    Dim Lo As ListObject
    Dim csTab As ICustomTable
    Dim counter As Long
    Dim rng As Range
    Dim FUN As WorksheetFunction
    Dim rngValue As String
    Dim sh As Worksheet

    Set sh = wb.Worksheets("Dictionary")
    Set Lo = sh.ListObjects(1)
    Set Check = Checking.Create("Dictionary Checking ---")
    Set csTab = CustomTable.Create(Lo, idCol:="Variable Name")
    Set FUN = Application.WorksheetFunction

    'Empty sheetNames
    Set rng = csTab.DataRange("Variable Name")

    'Unique variable names
    For counter = 1 To rng.Rows.Count
        rngValue = rng.Cells(counter, 1).Value
        If FUN.COUNTIF(rng, rngValue) > 1 Then
            Check.Add rngValue & "-" & counter, _
                     "Variable " & rngValue & " is duplicate:" & _
                     "variable names should be unique", _
                    checkingError
        End If
    Next

    CheckTables.Push Check
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
