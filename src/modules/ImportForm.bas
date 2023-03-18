Attribute VB_Name = "ImportForm"
Option Explicit

'Sub for functions on the import form

'Write the path to the new setup file to be imported
Public Sub NewSetupPath()
    Dim io As IOSFiles
    Set io = OSFiles.Create()
    'Load a setup file
    io.LoadFile "*.xlsb"
    If io.HasValidFile() Then [Imports].LabPath.Caption = "Path: " & io.File()
End Sub

'Import anaother setup to the new one

'speed app
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


Public Sub ImportOrCleanSetup()
    Dim importDict As Boolean
    Dim importChoi As Boolean
    Dim importExp As Boolean
    Dim importAna As Boolean
    Dim importTrans As Boolean
    Dim importPath As String
    Dim progObj As Object 'Progress label
    Dim importObj As ISetupImport 'Import Object
    Dim sheetsList As BetterArray
    Dim actsh As Worksheet
    Dim infoText As String
    Dim doLabel As String
    Dim pass As IPasswords

    BusyApp
    On Error GoTo ErrHand
    Set actsh = ActiveSheet
    importDict = [Imports].DictionaryCheck.Value
    importChoi = [Imports].ChoiceCheck.Value
    importExp = [Imports].ExportsCheck.Value
    importAna = [Imports].AnalysisCheck.Value
    importTrans = [Imports].TranslationsCheck.Value
    importPath = Application.WorksheetFunction.Trim( _
                Replace([Imports].LabPath, "Path: ", ""))

    Set progObj = [Imports].LabProgress
    Set pass = Passwords.Create(ThisWorkbook.Worksheets("__pass"))
    doLabel = [Imports].DoButton.Caption
    'freeze the pane for modifications
    progObj.Caption = ""
    Set importObj = SetupImport.Create(importPath, progObj)

    'Check import to be sure everything is fine (At least one import has to be made
    'and the file is correct (without missing parts)
    importObj.Check importDict, importChoi, importExp, importAna, _
                    importTrans, cleanSetup:=(doLabel = "Clear")
    Set sheetsList = New BetterArray

    'Add the sheets to import if required
    If importDict Then sheetsList.Push "Dictionary"
    If importChoi Then sheetsList.Push "Choices"
    If importExp Then sheetsList.Push "Exports"
    If importAna Then sheetsList.Push "Analysis"
    If importTrans Then sheetsList.Push "Translations"

    Select Case doLabel
    Case "Import"
        importObj.Import pass, sheetsList
         'Check the conformity of current setup file for errors
        If [Imports].ConformityCheck.Value Then
            'Check conformity of the current setup
        End If
        infoText = "Import Done!"
    Case "Clear"
        If MsgBox("Do you really want to clean the setup?", _
                 vbYesNo, "Confirmation") = vbYes Then
            importObj.Clean pass, sheetsList
            infoText = "Setup cleared!"
        End if
    End Select

    MsgBox infoText
    progObj.Caption = infoText
    actsh.Activate
ErrHand:
    NotBusyApp
End Sub

Public Sub PrepareForm(Optional ByVal cleanSetup As Boolean = False)
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

