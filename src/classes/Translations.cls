VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Translations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("")
'@IgnoreModule

Option Explicit
Implements ITranslation

Private Type TTrans
  lo As ListObject
End Type

Private this As TTrans

Public Function Create(ByVal Lo As ListObject)
  If Lo Is Nothing Then ThrowError "Please provide a listObject for translation"
  With New Translation
    Set .Table = Lo
    Set Create = .Self()
  End With
End Function

Public Property Get Table() As ListObject
  Set Table = this.lo
End Property

Public Property Set Table(ByVal lo As ListObject)
  Set this.lo = lo
End Property

Private Property Get TradSh() As Worksheet
  Dim lo As ListObject
  Set lo = Table()
  Set TradSh = lo.Parent
End Property

'where to start writing elements in the listObject
Public Property Get StartRange() As Range
  Dim cellRng As Range
  Dim lo As ListObject
  Set lo = Table()
  Set cellRng = Lo.Range().Cells(1, 1)
  Do while cellRng.Value <> vbNullString
    Set cellRng = cellRng.Offset(1)
  Loop
  Set StartRange = cellRng
End Property

'anchor range which is the column range of the table
Public Property KeyLangRange() As Range
  Dim rng As Range
  Dim Lo As ListObject

  Set Lo = Table()
  If (Lo.DataBodyRange Is Nothing) Then
    Set rng = Lo.Range.Cells(2, 1)
  Else
    Set rng = Lo.ListColumns(1).DataBodyRange
  End If
End Property

Public Property Get NumberTimes() As Long
  Dim nbTimes As Long
  Dim sh As Worksheet

  Set sh = Wksh()
  nbTimes = 0
  On Error Resume Next
  nbTimes = sh.Cells(1, 1).Value
  On Error GoTo 0
  NumberTimes = nbTimes
End Property

'Add a chunk (a word or a phrase in one cell to the translation table)
Private Sub AddChunk(ByVal lab As String, ByVal rngTag As String)
  Dim insertRng As Range 'where to insert the label
  'Test if the element to include Exists
  If chunkExists(lab) Then
    Set insertRng = chunkRange(lab)
  Else
    Set insertRng = StartRange()
    insertRng.Cells(1, 1).Value = lab
  End if
  insertRng.Offset(, -1).Value = rngTag
End Sub

Private Function chunkExists(ByVal lab As String)
  Dim langRng As Range
  Set langRng = KeyLangRange()
  chunkExists = Not (langRng.Find(What:=lab, LookAt:=xlWhole, MatchCase:=True) _ 
                     Is Nothing)
End Sub

Private Property Get chunkRange(ByVal lab As String) As Range
  Dim langRng As Range
  Set langRng = KeyLangRange()
  Set chunkRange = langRng.Find(What:=lab, LookAt:=xlWhole, MatchCase:=True)
End Sub

'Add a new language to the worksheet
Private Sub AddTransLang(ByVal langVal As String, ByVal pass As IPasswords)
  Dim sh As Worksheet
  Dim Lo As ListObject
  Dim hRng As Range
  Dim prevLoRng As Range
  Dim newLoRng As Range
  Dim askFirst As Long

  Set Lo = Table()
  Set sh = TradSh()

  asKFirst = MsgBox("Do you really want to add language " & langVal & " to translations?", _
                     vbYesNo, "Confirm the language")

  If (askFirst = vbNo) Then Exit Sub
  pass.UnProtect sh.Name
  Set hRng = Lo.HeaderRowRange

  'Previous range of the listObject
  Set prevLoRng = Lo.Range
  hRng.Cells(1, hRng.Columns.Count + 1).Value = langVal

  'Resize with new range of the listObject
  Set newLoRng = sh.Range(hRng.Cells(1, 1), prevLoRng.Cells(prevLoRng.Rows.Count, hRng.Columns.Count + 1))
  Lo.Resize newLoRng

  'Protect the worksheet
  pass.Protect sheetName:=sh.Name, strictProtect:=True
End Sub









'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Test if a range exists
Private Function rngExists(ByVal rngName As String) As Boolean
  Dim rng As Range
  On Error Resume Next
  Set rng = Range(rngName)
  On Error GoTo 0
  rngExists = (Not rng Is Nothing)
End Function


'Exposed methods
Public Sub ITranslations_UpdateValues(ByVal rngSh As Worksheet)
End Sub

Public Sub ITranslations_AddTransLang(ByVal langVal As String, _
                                      ByVal pass As IPasswords)
End Sub