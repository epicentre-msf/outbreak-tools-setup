VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetupImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Imports")
'@ModuleDescription("Import another setup to this new one")
'@IgnoreModule

Option Explicit

'Exposed methods
Implements ISetupImport

Private Const className As String = "SetupImport"

Private Type TSetupImport
  inPath As String
  progress As Object
  inWb As Workbook
End Type

Private This As TSetupImport

Public Function Create(ByVal importPath As String, ByVal progrObj As Object)

  'Error Management when creating the imports.
  If (progrObj Is Nothing) Then ThrowError "Unable to initialize the class, information Object not found"

  With New SetupImport
    .Path = importPath
    Set .ProgressObject = progrObj
    Set Create = .Self()
  End With

End Function

'Get the path of the import of the setup
Public Property Get Path() As String
  Path = This.inPath
End Property

Public Property Let Path(ByVal importPath As String)
  This.inPath = importPath
End Property

Public Property Get ProgressObject() As Object
  Set ProgressObject = This.progress
End Property

Public Property Set ProgressObject(ByVal progrObj As Object)
  Set This.progress = progrObj
End Property

Public Property Get Self() As ISetupImport
  Set Self = Me
End Property

Private Sub WriteInfo(ByVal charValue As String)
  Dim progrObj As Object
  Set progrObj = ProgressObject()
  progrObj.Caption = charValue
End Sub

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
  Application.Calculation = xlCalculationAutomatic

End Sub

Private Property Get ImportWkb() As Workbook
  If This.inWb Is Nothing Then Set This.inWb = Workbooks.Open(Path())
  Set ImportWkb = This.inWb
End Property

'Write a progress when importing the data
Private Sub WriteProgress(ByVal numValue As Integer)

  Const NUMBEROFBARS As Integer = 100
  Dim CurrentStatus As Integer
  Dim progrObj As Object
  Set progrObj = ProgressObject()
  CurrentStatus = (NUMBEROFBARS) * Round(numValue / 100, 1)
  progrObj.Caption = "[" & String(CurrentStatus, "|") & SPACE(NUMBEROFBARS - CurrentStatus) & "]" & _
                        " " & CInt(numValue) & "%"
End Sub


'Check if the setup file to import is a valid setup file
Private Sub Check(ByVal importDict As Boolean, _
                  ByVal importAna As Boolean, _
                  ByVal importTrans As Boolean)

  Dim wb As Workbook
  Dim filePath As String

  Const NO_FILE_WARNING As String = "Error : no setup file was founded at the provided path"
  Const NO_FILE_WB As String = "Error : the specified file is not an excel workbook, or is broken, unable to open it."
  Const IMPORT_OPTIONS_WARNING As String = "Error: Please select at least one option for imports (Dictionary, Analysis or Translations)"

  If (Not importDict) And (Not importAna) And (Not importTrans) Then
    MsgBox IMPORT_OPTIONS_WARNING, vbExclamation + vbOKOnly
    WriteInfo IMPORT_OPTIONS_WARNING
    Exit Sub
  End If

  filePath = Path()

  'Test if the file exists
  If Dir(filePath) = vbNullString Then
    MsgBox NO_FILE_WARNING, vbExclamation + vbOKOnly
    WriteInfo NO_FILE_WARNING
    Exit Sub
  End If

  'Test if the workbook could be loaded without problems
  BusyApp
  On Error Resume Next
  Set wb = Workbooks.Open(filePath)
  On Error GoTo 0

  If wb Is Nothing Then
    MsgBox NO_FILE_WB, vbExclamation + vbOKOnly
    WriteInfo NO_FILE_WB
    Exit Sub
  End If

  CheckSheet wb, "Dictionary"
  CheckSheet wb, "Choices"
  CheckSheet wb, "Analysis"
  CheckSheet wb, "Exports"
  CheckSheet wb, "Translations"

  wb.Close saveChanges:=False
  NotBusyApp
End Sub


Private Sub ImportTheDictionary(ByVal pass As ILLPasswords)
  Dim part As Object
  Dim wb As Workbook
  Dim sheetName As String
  Dim sh As Worksheet
  Dim shImp As Worksheet

  'Import the dictionary
  Set wb = ImportWkb()
  sheetName = "Dictionary"
  Set sh = ThisWorkbook.Worksheets(sheetName)
  Set part = LLdictionary.Create(sh, 4, 1)
  Set shImp = wb.Worksheets(sheetName)

  part.Import shImp, pass
  'Import the choices

  'Import the exports

End Sub

'Import button
Private Sub Import(ByVal pass As IPasswords, ByVal dict As Boolean, ByVal ana As Boolean, ByVal trans As Boolean)

  If dict Then ImportTheDictionary pass

End Sub

'Errors
'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, className, "Error : " & errorMessage
End Sub

'Check if a sheet with a given name is present in the provided workbook
Private Sub CheckSheet(ByVal wb As Workbook, sheetName As String)
  Dim sh As Worksheet

  On Error Resume Next
  Set sh = wb.Worksheets(sheetName)
  On Error GoTo 0

  'Inform the user the sheet is not present
  If (sh Is Nothing) Then
    MsgBox "Error: the sheet " & sheetName & "is not present in the setup to import", vbOKOnly + vbExclamation
    WriteInfo "Error: the sheet " & sheetName & " is not present in the setup to import"
    Exit Sub
  End If
End Sub
'-------- INTERFACE
Public Sub ISetupImport_Check(ByVal dict As Boolean, ByVal ana As Boolean, ByVal trans As Boolean)
  CheckType dict, ana, trans
End Sub

Public Sub ISetupImport_Import(ByVal pass As IPasswords, ByVal dict As Boolean, ByVal ana As Boolean, ByVal trans As Boolean)
  Import pass, dict, ana, trans
End Sub
