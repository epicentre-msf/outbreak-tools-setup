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
    Dim wb As Workbook

    BusyApp
    On Error GoTo errHand
    Set actsh = ActiveSheet
    Set wb = ThisWorkbook
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
        If MsgBox("Do you really want to clear the setup?", _
                 vbYesNo, "Confirmation") = vbYes Then
            importObj.Clean pass, sheetsList
            infoText = "Setup cleared!"
        End If
    End Select

    MsgBox infoText
    progObj.Caption = infoText
    actsh.Activate
    SetAllUpdatedTo "yes"
    wb.Worksheets("Analysis").Calculate
errHand:
    NotBusyApp
End Sub

