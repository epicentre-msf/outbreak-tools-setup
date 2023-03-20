Attribute VB_Name = "EventsRibbon"

Option Explicit

'Private constants for Ribbon Events
Private Const TRADSHEETNAME As String = "Translations"
Private Const TABTRANSLATION As String = "Tab_Translations"
Private Const PASSSHEETNAME As String = "__pass"
Private Const UPDATEDSHEETNAME As String = "__updated"

'Private Subs to speed up process
Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
End Sub

'Resize the listObjects in the current sheet
Public Sub clickResize(control As IRibbonControl)
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    ManageRows sheetName:=sheetName, del:=True
End Sub

'add rows to listObject
Public Sub clickAddRows(ByRef control As Office.IRibbonControl)
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    ManageRows sheetName:=sheetName, del:=False
End Sub

'Callback for editLang onChange
Public Sub clickAddLang(control As IRibbonControl, text As String)

    Dim pass As IPasswords
    Dim trads As ITranslations
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim askFirst As Long

    If text = vbNullString Then Exit Sub
    BusyApp

    'Ask before proceeding
    askFirst = MsgBox("Do you really want to add language(s) " & _
                      text & " to translations?", _
                      vbYesNo, "Confirm")

    If (askFirst = vbNo) Then Exit Sub

    Set wb = ThisWorkbook
    Set sh = wb.Worksheets(TRADSHEETNAME)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set trads = Translations.Create(sh, TABTRANSLATION)

    pass.UnProtect TRADSHEETNAME
    trads.AddTransLang text
    pass.Protect TRADSHEETNAME, True, True

    MsgBox "Done!"
    NotBusyApp
End Sub

'Callback for btnTransAdd onAction
Public Sub clickAddTrans(control As IRibbonControl)
    Dim pass As IPasswords
    Dim trads As ITranslations
    Dim wb As Workbook
    Dim tradsh As Worksheet
    Dim upsh As Worksheet
    Dim askFirst As Long

    BusyApp
    'Ask before proceeding
    askFirst = MsgBox("Do you want to update the translation sheet?", vbYesNo, "Confirm")

    If (askFirst = vbNo) Then Exit Sub

    Application.Cursor = xlWait
    'On Error GoTo errHand

    Set wb = ThisWorkbook
    Set tradsh = wb.Worksheets(TRADSHEETNAME)
    Set upsh = wb.Worksheets(UPDATEDSHEETNAME)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set trads = Translations.Create(tradsh, TABTRANSLATION)

    pass.UnProtect TRADSHEETNAME
    trads.UpdateTrans upsh
    pass.Protect TRADSHEETNAME, True, True

    'Set all updates to no (this sub is in the EventsGlobal module)
    SetAllUpdatedTo "no"
    NotBusyApp
    Application.Cursor = xlDefault
    Exit Sub

errHand:
    Application.Cursor = xlDefault
    MsgBox "An internal error occured", vbCritical + vbOkOnly
    NotBusyApp
End Sub

'Callback for btnTransUp onAction
Public Sub clickUpdateTranslate(control As IRibbonControl)
    'remove update columns and add new columns to watch
    BusyApp
    CleanUpdateColumns
    UpdatedWatchedValues
    NotBusyApp
    MsgBox "Done!"
End Sub

'Callback for btnChk onAction
Public Sub clickCheck(control As IRibbonControl)
End Sub

'Callback for btnImp onAction
Public Sub clickImport(control As IRibbonControl)
    PrepareForm cleanSetup:=False
    [Imports].Show
End Sub

'Callback for btnClear onAction
Public Sub clickClearSetup(control As IRibbonControl)
    PrepareForm cleanSetup:=True
    [Imports].Show
End Sub

'===== Auxilliary subs used in the process

'Add or Remove Rows to a table
Private Sub ManageRows(ByVal sheetName As String, Optional ByVal del As Boolean = False)
    Dim part As Object
    Dim sh As Worksheet
    Dim shpass As Worksheet
    Dim pass As IPasswords

    BusyApp

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(sheetName)
    sh.EnableCalculation = False
    Set shpass = ThisWorkbook.Worksheets(PASSSHEETNAME)
    On Error GoTo 0

    If (sh Is Nothing) Or (shpass Is Nothing) Then Exit Sub

    '5 is the start line of the dictionary
    '4 is the start column of the dictionary
    Select Case sheetName
    Case "Dictionary"
        Set part = LLdictionary.Create(sh, 5, 1)
    Case "Choices"
        Set part = LLchoice.Create(sh, 4, 1)
    Case "Analysis"
        Set part = Analysis.Create(sh)
    End Select

    'Exit if unable to find the corresponding object
    If part Is Nothing Then Exit Sub
    Set pass = Passwords.Create(shpass)
    pass.UnProtect sh.Name

    If del Then
        part.RemoveRows
    Else
        part.AddRows
    End If

    pass.Protect sh.Name, (sh.Name = "Analysis")
    sh.EnableCalculation = True
    NotBusyApp
End Sub


Private Sub CleanUpdateColumns()
    'Clear the update sheet
    Dim upsh As Worksheet
    Dim Lo As ListObject
    Dim wb As Workbook
    Dim namesRng As Range
    Dim counter As Long
    Set wb = ThisWorkbook
    Set upsh = wb.Worksheets(UPDATEDSHEETNAME)

    'Unlist all listObjects in the worksheet and delete all names
    For Each Lo In upsh.ListObjects
        Set namesRng = Lo.ListColumns("rngname").Range
        For counter = 1 To namesRng.Rows.Count
            On Error Resume Next
            wb.Names(namesRng.Cells(counter, 1).Value).Delete
            On Error GoTo 0
        Next
        Lo.Unlist
    Next
    upsh.Cells.Clear
End Sub

'Update the translation values
Private Sub UpdatedWatchedValues()
    Dim sh As Worksheet
    Dim sheetsList As BetterArray
    Dim counter As Long
    Dim sheetName As String

    Set sheetsList = New BetterArray
    sheetsList.Push "Dictionary", "Choices", "Exports", "Analysis"
    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        Set sh = ThisWorkbook.Worksheets(sheetName)
        'Write update status on each sheet
        writeUpdateStatus sh
    Next
End Sub

'Update status of columns to watch
Private Sub writeUpdateStatus(sh As Worksheet)
    Dim upsh As Worksheet
    Dim upId As String
    Dim upObj As IUpdatedValues
    Dim Lo As ListObject

    Set upsh = ThisWorkbook.Worksheets(UPDATEDSHEETNAME)
    upId = LCase(Left(sh.Name, 4))
    For Each Lo In sh.ListObjects
        If sh.Name = "Analysis" Then _
        upId = LCase(Replace(Lo.Name, "Tab_", ""))
        Set upObj = UpdatedValues.Create(upsh, upId)
        upObj.AddColumns Lo
    Next
End Sub

'Prepare the Import Form
Private Sub PrepareForm(Optional ByVal cleanSetup As Boolean = False)
    If cleanSetup Then
        [Imports].LoadButton.Visible = False
        [Imports].LabPath.Visible = False
        [Imports].InfoChoice.Caption = "Select what to Clear"
        [Imports].DictionaryCheck.Caption = "Clear Dictionary"
        [Imports].ChoiceCheck.Caption = "Clear Choices"
        [Imports].ExportsCheck.Caption = "Clear Exports"
        [Imports].AnalysisCheck.Caption = "Clear Analysis"
        [Imports].TranslationsCheck.Caption = "Clear Translation"
        [Imports].ConformityCheck.Visible = False
        [Imports].DoButton.Caption = "Clear"

        'Resize and change position of elements
        [Imports].Height = 400
        [Imports].InfoChoice.Top = 20
        [Imports].DictionaryCheck.Top = 50
        [Imports].ChoiceCheck.Top = 80
        [Imports].ExportsCheck.Top = 110
        [Imports].AnalysisCheck.Top = 140
        [Imports].TranslationsCheck.Top = 170
        [Imports].LabProgress.Top = 200
        [Imports].DoButton.Top = 270
        [Imports].Quit.Top = 310
    Else
        [Imports].InfoChoice.Caption = "Select what to Import"
        [Imports].DictionaryCheck.Caption = "Import Dictionary"
        [Imports].ChoiceCheck.Caption = "Import Choices"
        [Imports].ExportsCheck.Caption = "Import Exports"
        [Imports].AnalysisCheck.Caption = "Import Analysis"
        [Imports].TranslationsCheck.Caption = "Import Translation"
        [Imports].ConformityCheck.Visible = True
        [Imports].LoadButton.Visible = True
        [Imports].LabPath.Visible = True
        [Imports].DoButton.Caption = "Import"

        'resize the worksheet and position of elements
        [Imports].Height = 500
        [Imports].LoadButton.Top = 10
        [Imports].LabPath.Top = 55
        [Imports].InfoChoice.Top = 135
        [Imports].DictionaryCheck.Top = 170
        [Imports].ChoiceCheck.Top = 200
        [Imports].ExportsCheck.Top = 230
        [Imports].AnalysisCheck.Top = 260
        [Imports].TranslationsCheck.Top = 290
        [Imports].DoButton.Top = 350
        [Imports].LabProgress.Top = 390
        [Imports].Quit.Top = 440
    End If
End Sub
