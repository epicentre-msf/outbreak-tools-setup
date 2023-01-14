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

Private Sub Workbook_Open()

  'Set all updates to true (no updates is done)
  SetAllUpdates toValue:=True
  sheetLists.Range("nbtimes").Value = 0
  
  On Error Resume Next
  sheetLists.ListObjects(C_sTabListOfChoicesNames).DataBodyRange.Delete
  On Error GoTo 0
End Sub


