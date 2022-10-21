VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Cancel Then Exit Sub
    
    Dim counter As Long
    Dim listChoices As String
    Dim listChoiceLo As ListObject
    
    Set listChoiceLo = sheetLists.ListObjects(C_sTabListOfChoicesNames)
    If listChoiceLo.DataBodyRange Is Nothing Then Exit Sub
    
    For counter = 1 To listChoiceLo.DataBodyRange.Rows.Count
        
        listChoices = sheetLists.Cells(counter + 1, listChoiceLo.DataBodyRange.Column).Value
        
        On Error Resume Next
            ThisWorkbook.Names(listChoices).Delete
        On Error GoTo 0
        
    Next
      sheetChoicesLists.Cells.Clear
End Sub

Private Sub Workbook_Open()
  Call SetAllUpdates(toValue:=True)
  sheetLists.Range("nbtimes").Value = 0
  
  sheetChoicesLists.Cells.Clear
  On Error Resume Next
  sheetLists.ListObjects(C_sTabListOfChoicesNames).DataBodyRange.Delete
  On Error GoTo 0
End Sub


