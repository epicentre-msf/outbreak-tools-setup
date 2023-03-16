VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("")
'@IgnoreModule

Option Explicit
Implements ICustomTable
Private Type TCustomTable
  Lo As ListObject
  idCol As String
  idPrefix As String
End Type

Private this As TCustomTable
Private Const CLASSNAME As String = "ICustomTable"
Private Const NBROWSADD As Long = 5 'Number of Rows to add when resizing

Public Function Create(ByVal Lo As ListObject, _
                      Optional ByVal idCol As String = vbNullString, _
                      Optional ByVal idPrefix As String = vbNullString) As ICustomTable
  If Lo Is Nothing Then ThrowError "Unable to initialize the class, no listobject provided"

  With New CustomTable
    Set .Table = Lo
    .IdValue = idCol
    .PrefixValue = idPrefix
    Set Create = .Self()
  End With
End Function


Public Property Get Self() As ICustomTable
  Set Self = Me
End Property

'Worksheet of the listObject
Private Function Wksh() As Worksheet
  Dim Lo As ListObject
  Set Lo = Table()
  Set Wksh = Lo.Parent
End Function


'Table listObject
Public Property Get Table() As ListObject
  Set Table = this.Lo
End Property

Public Property Set Table(ByVal Lo As ListObject)
  Set this.Lo = Lo
End Property

Public Property Get IdValue() As String
  IdValue = this.idCol
End Property

Public Property Get PrefixValue() As String
  PrefixValue = this.idPrefix
End Property

Public Property Let IdValue(ByVal idCol As String)
  this.idCol = idCol
End Property

Public Property Let PrefixValue(ByVal idPrefix As String)
  this.idPrefix = idPrefix
End Property

'HeaderRowRange
Private Property Get HeaderRowRange() As Range
  Dim Lo As ListObject
  Set Lo = Table()
  Set HeaderRowRange = Lo.HeaderRowRange()
End Property

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

'Resize a table
Private Sub Resize(Optional ByVal AddRows As Boolean = False, Optional ByVal totalRowCount As Long = 0)
  Dim Lo As ListObject
  'Begining of the tables
  Dim loRowHeader As Long
  Dim loColHeader  As Long
  Dim rowCounter As Long
  'End of the listobject table
  Dim loRowsEnd As Long
  Dim loColsEnd As Long
  Dim sh As Worksheet
  Set sh = Wksh()
  Set Lo = Table()
  'Remove unwanted rows from each listObject
  With sh
    'Rows and columns at the begining of the table to resize
    loRowHeader = Lo.Range.Row
    loColHeader = Lo.Range.Column
   'Rows and Columns at the end of the Table to resize
    loRowsEnd = loRowHeader + Lo.Range.Rows.Count - 1
    loColsEnd = loColHeader + Lo.Range.Columns.Count - 1
    rowCounter = loRowsEnd
    If AddRows Then
      loRowsEnd = loRowsEnd + 1 'Start at the bottom of the table
      For rowCounter = 1 To NBROWSADD + 1
            .Rows(loRowsEnd).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      Next
      loRowsEnd = loRowsEnd + NBROWSADD
      Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
      AddIds
    Else
      Do While (rowCounter > loRowHeader + 1)
        If (Application.WorksheetFunction.CountA(.Rows(rowCounter)) <= totalRowCount) Then
          .Rows(rowCounter).EntireRow.Delete
          'update the end rows
          loRowsEnd = loRowsEnd - 1
        End If
        rowCounter = rowCounter - 1
      Loop
      Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
    End If
  End With
End Sub

Private Sub AddIds()
  Dim rng As Range
  Dim counter As Long
  Dim idPref As String
  Dim ID As String

  ID = IdValue()
  If (ID = vbNullString) Then Exit Sub
  If Not (ColumnExists(ID)) Then Exit Sub
  idPref = PrefixValue()
  Set rng = DataRange(ID)
  For counter = 1 To rng.Rows.Count
    rng.Cells(counter, 1).Value = Application.WorksheetFunction.Trim(idPref & " " & counter)
  Next
End Sub


Private Sub SetValidation(ByVal colName As String, _
                          ByVal drop As IDropdownLists, _
                          ByVal dropName As String, _
                          Optional ByVal alertType As String = "info")

    Dim rng As Range
    Dim sh As Worksheet

    Set sh = Wksh()
    'Test if the colname exists before seting the validation
    If Not ColumnExists(colName) Then Exit Sub
    Set rng = DataRange(colName)
    'set validation for each cell in the datarange
    drop.SetValidation cellRng:=rng, listName:=dropName, _
                       alertType:=alertType
End Sub


'Column Range and column Index of a custom table
Private Function ColumnExists(ByVal colName As String) As Boolean
    Dim exist As Boolean
    Dim hRng As Range
    Set hRng = HeaderRowRange
    exist = False
    On Error Resume Next
    exist = (Not hRng.Find(What:=colName, lookAt:=xlPart, MatchCase:=False) Is Nothing)
    On Error GoTo 0
    ColumnExists = exist
End Function

Private Function ColumnIndex(ByVal colName) As Long
    Dim Index As Long
    Dim hRng As Range

    Set hRng = HeaderRowRange()
    If ColumnExists(colName) Then
        Index = hRng.Find(What:=colName, lookAt:=xlPart, MatchCase:=False).Column - hRng.Column + 1
    Else
        Index = -1
    End If
    ColumnIndex = Index
End Function

Private Property Get DataRange(ByVal colName As String, Optional ByVal includeHeaders As Boolean = False) As Range
  Dim Lo As ListObject
  Set Lo = Table()
  Dim rng As Range

  If Not ColumnExists(colName) Then Exit Property
  If includeHeaders Then
    Set rng = Lo.ListColumns(ColumnIndex(colName)).Range
  Else
    'DataBody range could be Nothing
    On Error Resume Next
    Set rng = Lo.ListColumns(ColumnIndex(colName)).DataBodyRange
    On Error GoTo 0
    'You can test here if it is Nothing, in that case the table is Empty, just return the second line
    If rng Is Nothing Then Set rng = Lo.ListColumns(ColumnIndex(colName)).Range.Cells(2, 1)
    Set DataRange = rng
  End If
End Property

Private Property Get CellRange(ByVal colName As String, ByVal lineNum As Long) As Range
  Dim cellPos As Long
  Dim rng As Range

  If Not ColumnExists(colName) Then Exit Property
  Set rng = DataRange(colName)
  cellPos = lineNum - rng.Row
  If cellPos < 0 Then Exit Property
  Set CellRange = rng.Cells(cellPos, 1)
End Property

Private Property Get Value(ByVal colName As String, ByVal keyName As String) As String

  Dim idRng As Range
  Dim colRng As Range
  Dim sh As Worksheet
  Dim FUN As WorksheetFunction
  Dim formValue As Variant

  If IdValue = vbNullString Then Exit Property

  Set FUN = Application.WorksheetFunction
  Set sh = Wksh()
  If (Not ColumnExists(colName)) Or (Not ColumnExists(IdValue)) Then Exit Property
  Set idRng = DataRange(IdValue)
  Set colRng = DataRange(colName)

  On Error Resume Next
  formValue = CStr(FUN.Index(colRng, FUN.Match(keyName, idRng, 0), 1))
  On Error GoTo 0

  If TypeName(formValue) = "String" Then Value = formValue
End Property

Public Sub ICustomTable_AddRows()
  Resize AddRows:=True
End Sub

Public Sub ICustomTable_RemoveRows(Optional ByVal totalCount As Long = 0)
  Resize AddRows:=False, totalRowCount:=totalCount
End Sub

Public Sub ICustomTable_SetValidation(ByVal colName As String, ByVal drop As IDropdownLists, _
                                      ByVal dropName As String, Optional ByVal alertType As String = "info")
    SetValidation colName, drop, dropName, alertType
End Sub

Public Property Get ICustomTable_DataRange(ByVal colName As String, Optional ByVal includeHeaders As Boolean = False) As Range
  Set ICustomTable_DataRange = DataRange(colName, includeHeaders)
End Property

Public Property Get ICustomTable_Value(ByVal colName As String, ByVal keyName As String) As String
  ICustomTable_Value = Value(colName, keyName)
End Property

Public Property Get ICustomTable_CellRange(ByVal colName As String, ByVal lineNum As Long) As Range
  Set ICustomTable_CellRange = CellRange(colName, lineNum)
End Property


'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub
