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
End Type

Private this As TCustomTable
Private Const CLASSNAME As String = "ICustomTable"
Private Const NBROWSADD As Long = 5 'Number of Rows to add when resizing

Public Function Create(ByVal Lo As ListObject) As ICustomTable
  If Lo Is Nothing Then ThrowError "Unable to initialize the class, no listobject provided"

  With New CustomTable
    Set .Table = Lo
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
Private Sub Resize(ByVal pass As IPasswords, Optional ByVal AddRows As Boolean = False, Optional ByVal totalRowCount As Long = 0)

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

    BusyApp

    pass.UnProtect sh.Name

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
        
        Else

            Do While (rowCounter > loRowHeader + 1)
            If (Application.WorksheetFunction.CountA(.Rows(rowCounter)) <= totalRowCount) Then
              .Rows(rowCounter).EntireRow.Delete
              'update the end rows
              loRowsEnd = loRowsEnd - 1
            End If
            rowCounter = rowCounter - 1
            Loop
        End If
    
        Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
    End With
    NotBusyApp
    pass.Protect sh.Name
End Sub


'Column Range and column Index of a custom table
Private Function ColumnExists(ByVal colName As String) As Boolean
    Dim hRng As Range
    Set hRng = HeaderRowRange
    ColumnExists = (Not hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=False) Is Nothing)
End Function

Private Function ColumnIndex(ByVal colName) As Long
    Dim Index As Long
    Dim hRng As Range

    Set hRng = HeaderRowRange()
    If ColumnExists(colName) Then
        Index = hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=False).Column - hRng.Column + 1
    Else
        Index = -1
    End If
    ColumnIndex = Index
End Function

Private Function DataRange(ByVal colName As String) As Range
  Dim Lo As ListObject
  Set Lo = Table()
  If ColumnExists(colName) Then
    Set DataRange = Lo.ListColumns(ColumnIndex(colName)).Range
  End If
End Function

Public Sub ICustomTable_AddRows(ByVal pass As IPasswords)
  Resize pass:=pass, AddRows:=True
End Sub

Public Sub ICustomTable_RemoveRows(ByVal pass As IPasswords, Optional ByVal totalCount As Long = 0)
  Resize AddRows:=False, totalRowCount:=totalCount, pass:=pass
End Sub


'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub
