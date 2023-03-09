Attribute VB_Name = "Events"

Option Explicit

'Module for all the events - related actions in the setup file

'Import from another setup
Public Sub clickImport()
End Sub

'add rows to listObject
Public Sub clickAddRows()
End Sub

'resize the current listObject
Public Sub clickResize()
End Sub

'clear data in the current setup
Public Sub clickClearSetup()
End Sub

'check the current setup for incoherences
Public Sub clickCheck()
End Sub

'Add or Remove Rows to a table
Public Sub ManageRows(ByVal sheetName As String, Optional ByVal del As Boolean = False)
    Dim part As Object
    Dim sh As Worksheet
    Dim shpass As Worksheet
    Dim pass As IPasswords

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(sheetName)
    Set shpass = ThisWorkbook.Worksheets("__pass")
    On Error GoTo 0

    If (sh Is Nothing) Or (shpass Is Nothing) Then Exit Sub

    '4 is the start line of the dictionary
    '1 is the start column of the dictionary

    '2 is the start line of the choices
    '1 is the start column of the choices

    Select Case sheetName
    Case "Dictionary"
        Set part = LLdictionary.Create(sh, 4, 1)
    Case "Choices"
        Set part = LLchoice.Create(sh, 2, 1)
    Case "Analysis"
        'Set part = Analysis.Create(sh)
    End Select

    Set pass = Passwords.Create(shpass)

    If del Then
        part.RemoveRows pass:=pass
    Else
        part.AddRows pass:=pass
    End If
End Sub
