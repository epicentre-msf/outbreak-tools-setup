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

Public Sub ImportSetup()
    Dim importDict As Boolean
    Dim importAna As Boolean
    Dim importTrans As Boolean
    Dim importPath As String
    Dim progObj As Object 'Progress label
    Dim importObj As ISetupImport 'Import Object
    Dim pass As IPasswords


    importDict = [Imports].DictionaryCheck.Value
    importAna = [Imports].AnalysisCheck.Value
    importTrans = [Imports].TranslationsCheck.Value
    importPath = Application.WorksheetFunction.Trim( _
                Replace([Imports].LabPath, "Path: ", ""))

    Set progObj = [Imports].LabProgress

    Set pass = Passwords.Create(ThisWorkbook.Worksheets("__pass"))

    'freeze the pane for modifications
    [Imports].Enabled = False
    [Imports].LabProgress.Caption = ""

    Set importObj = SetupImport.Create(importPath, progObj)

    'Check import to be sure everything is fine (At least one import has to be made
    'and the file is correct (without missing parts)
    importObj.Check importDict, importAna, importTrans

    importObj.Import pass

    'Check the conformity of current setup file for errors
    If [Imports].ConformityCheck.Value Then

    End If

    [Imports].Enabled = True
End Sub

