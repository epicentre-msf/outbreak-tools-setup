VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manage Exports"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Manage Exports")

Option Explicit

Implements ILLExport

Private Type TLLExport
    expData As IDataSheet
End Type

Private this As TLLExport

Public Function Create(ByVal sh As Worksheet, ByVal expStartRow As Long, ByVal expStartCol As Long) As ILLExport
  
    'Set creation of datasheet
    Dim customData As IDataSheet
    Set customData = DataSheet.Create(sh, expStartRow, expStartCol)

    With New LLExport
        Set .Data = customData
        Set Create = .Self()
    End With

End Function

'Data for the exports
Public Property Get Data() As IDataSheet
    Set Data = this.expData
End Property

Public Property Set Data(ByVal customData As IDataSheet)
    Set this.expData = customData
End Property

Public Property Get Self() As ILLExport
    Set Self = Me
End Property

'Export specifications to another workbook
Private Sub ExportSpecs(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Dim expData As IDataSheet
    Set expData = Data()
    expData.Export toWkb, Hide:=Hide
End Sub

Private Sub Translate(ByVal trans As ITranslation)
    Dim expData As IDataSheet
    Dim rng As Range
    Set expData = Data()
    Set rng = expData.DataRange("label button")
    trans.TranslateRange rng
End Sub

'Add / Remove Rows
Private Sub ManageRows(Optional ByVal del As Boolean = False)
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim tabl As ICustomTable

    Set sh = Wksh()

    On Error Resume Next
    Set Lo = sh.ListObjects(1)
    On Error GoTo 0
    If Lo Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(Lo, "export number")
    If del Then
        tabl.RemoveRows totalCount:=1
    Else
        tabl.AddRows nbRows:=2
    End If
End Sub

Private Property Get NumberOfExports() As Long
    On Error Resume Next
    NumberOfExports = Data.DataRange.Rows.Count
    On Error GoTo 0
End Property

'Import from another worksheet
Private Sub ImportSpecs(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Dim expData As IDataSheet
    Set expData = Data()
    expData.Import fromWksh, fromStartRow, fromStartcol
End Sub

Private Sub ILLExport_Translate(ByVal trans As ITranslation)
    Translate trans
End Sub

Private Sub ILLExport_ExportSpecs(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    ExportSpecs toWkb, Hide:=Hide
End Sub

Private Sub ILLExport_ImportSpecs(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    ImportSpecs fromWksh, fromStartRow, fromStartcol
End Sub

'Remove Rows
Private Sub ILLExport_RemoveRows()
    ManageRows del:=True
End Sub

'Add Rows
Private Sub ILLExport_AddRows()
    ManageRows
End Sub

Private Property Get ILLExport_NumberOfExports() As Long
    ILLExport_NumberOfExports = NumberOfExports()
End Property

