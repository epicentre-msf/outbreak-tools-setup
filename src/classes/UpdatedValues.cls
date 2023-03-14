VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "UpdatedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Setup")
'@ModuleDescription("Watch updated values in the setup")
'@IgnoreModule
Option Explicit
Private Const CLASSNAME As String = "IUpdatedValues"
Private Type TUpValues
  upLo As ListObject
  sh As Worksheet
  upName As String
End Type

Private this As TUpValues

Public Function Create(ByVal sh As Worksheet, ByVal upName As String) As IUpdatedValues
  If sh Is Nothing Then ThrowError "Unable to create an instance of the class UpdatedValues, the worksheet does not exists"
  If upName = vbNullString Then ThrowError "Please provide a name for your object for UpdatedValues"

  With New UpdatedValues
    Set .Wksh = sh
    .ID = upName
    Set Create = .Self()
  End With
End Function

Public Property Get Self() As IDropdownLists
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get ID() As String
  ID = this.upName
End Property

Public Property Let ID(ByVal upName As String)
  this.upName = upName
End Property

Private Property Get LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()
    With sh
        colIndex = .Cells(1, .Columns.Count).End(xlToLeft).Column
        colIndex = IIf(colIndex = 1 And .Cells(1, 1).Value <> "", colIndex, colIndex + 2)
    End With
    LastColumn = colIndex
End Property

Private Property Get Table() As ListObject
  Dim sh As Worksheet
  Dim colIndex As Long
  Dim rng As Range
  Set sh = Wksh()

  If this.upLo Is Nothing Then
    If Not Exists(ID) Then
      colIndex = LastColumn()
      With sh
        .Cells(1, colIndex).Value = "colname"
        .Cells(1, colIndex + 1).Value = "rngname"
        .Cells(1, colIndex + 2).Value = "updated"
        Set rng = .Range(.Cells(1, colIndex), .Cells(2, colIndex + 2))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = CurrentLoName(ID)
      End With
    End If
    Set this.upLo = .ListObjects(CurrentLoName(ID))
  End If
  Set Table = this.upLo
End Property

Private Property Get IsUpdated(ByVal checkColName As String) As Boolean
  Dim upLo As ListObject
  Dim rng As Range
  Dim cellRng As Range
  Set upLo = Table()
  Set rng = upLo.ListColumns("rngname").Range
  Set cellRng = rng.Find(What:=checkColName, LookAt:=xlWhole, MatchCase:=True)
  If cellRng Is Nothing Then Exit Property
  IsUpdated = (cellRng.Cells(1, 2).Value = "yes")
End Property

Private Sub AddColumn(ByVal Lo As ListObject, colName As String, idName As String)

  Dim hRng As Range 'header row range of the listObject
  Dim idRng As Range
  Dim upLo As ListObject 'Current Updating listObject
  Dim wb As Workbook
  Dim sh As Worksheet

  If Lo Is Nothing Then ThrowError "ListObject is not initialize, please provide valid listObject"
  If Not Exists(ID) Then ThrowError "Unable to add elements to a non existant update table, check the id name"

  'Check if the column Exists in the header range
  Set hRng = Lo.HeaderRowRange
  If hRng.Find(What:=colName, lookAt:=xlPart, MatchCase:=False) Is Nothing _
   Then ThrowError colName & " not found in listobject " & Lo.Name

  'Check if the id already Exists, in that case exit
  Set idRng = upLo.ListColums("rngname").Range()
  If (Not idRng.Find(What:=idName, looAt:=xlWhole, MatchCase:=True) _
     Is Nothing) Then Exit Sub

  Set sh = Wksh()
  Set wb = sh.Parent

  'Now add values and names to the workbook
  wb.Names.Add Name:="RNG_" & idName & "_" & ID(), _
              RefersToR1C1:="=" & Lo.Name & "[" & colName & "]"
  'Write values to the listObject of the update
  WriteValues colName, "RNG_" & idName & "_" & ID()
End Sub


Private Sub WriteValues(ByVal colName As String, rngName As String, Optional ByVal statusType As String = "no")
  Dim upLo As ListObject
  Dim rowRng As Range
  Dim loRng As Range
  Dim sh As Worksheet

  Set upLo = Table()
  Set rowRng = upLo.HeaderRowRange

  Do While (Not IsEmpty(rowRng))
    Set rowRng = rowRng.Offset(1)
  Loop

  rowRng.Cells(1, 1).Value = colName
  rowRng.Cells(1, 2).Value = rngName
  rowRng.Cells(1, 3).Value = statusType

  'Resize the listObject
  Set sh = Wksh()
  Set loRng = sh.Range(upLo.HeaderRowRange.Cells(1, 1), rowRng.cells(1, 3))
  upLo.Resize loRng
End Sub

Private Sub WriteUpdate(ByVal idName As String, Byval pos As Long, Optional Byval statusType As String = "yes")
  Dim upLo As ListObject
  Dim idRng As Range
  'Find the colName in the listObject
  Set upLo = Table()
  Set idRng = upLo.ListColumns("rngname").Range
  idRng.Cells(pos, 2).Value = statusType
End Sub

Private Sub Clear()
  Dim upLo As ListObject
  Set upLo = Table()
  If Not upLo.DataBodyRange Is Nothing Then upLo.DataBodyRange.Delete
End Sub

'Check the update status
Private Sub CheckStatus(ByVal sh As Worksheet, ByVal Target As Range)
  Dim rng As Range
  Dim counter As Long
  Dim idRng As Range
  Dim upLo As ListObject
  Dim rngName As String
  Set upLo = Table()
  Set idRng = upLo.ListColumns("rngname").Range
  For counter = 1 To idRng.Rows.Count
    rngName = idRng.Cells(counter, 1).Value
    On Error Resume Next
      Set rng = sh.Range(rngName)
    On Error GoTo 0
    If Not (rng Is Nothing) Then
      If Not IsUpdated(rngName) Then
        If Not (InterSect(Target, rng) Is Nothing) Then
          WriteUpdate rngName, counter
          Exit Sub
        End If
      End If
      Set rng = Nothing
    End If
  Next
End Sub

Private Sub ResetUpdateStatus()
  Dim upLo As ListObject
  Dim idRng As Range
  Dim counter As Long
  Set upLo = Table()
  If upLo.DataBodyRange Is Nothing Then Exit Sub
  Set idRng = upLo.ListColumns("rngname").DataBodyRange
  For counter = 1 To idRng.Rows.Count
    idRng.Cells(counter, 2).Value = "no"
  Next
End Sub

Private Property Get CurrentLoName(ByVal listName As String) As String
    CurrentLoName = "UpLo_" & listName
End Property

'Test if a listobject exists in a sheet
Public Function Exists(ByVal listName As String) As Boolean
    Dim Lo As ListObject
    Dim loName As String
    Dim sh As Worksheet
    Set sh = Wksh()
    loName = CurrentLoName(listName)
    On Error Resume Next
    Set Lo = sh.ListObjects(loName)
    On Error GoTo 0
    Exists = (Not Lo Is Nothing)
End Function


'Errors
Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
