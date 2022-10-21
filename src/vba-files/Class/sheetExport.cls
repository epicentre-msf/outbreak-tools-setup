VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheetExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Updated("Exp") Then UpdateValue Not (Intersect(Target, sheetExport.ListObjects(C_sTabExports).ListColumns(C_sExportHeaderLabelButton).Range) Is Nothing), "Exp"
End Sub



Private Sub Worksheet_Deactivate()
    sheetLists.Range("LastWorkSheet").Value = Me.Name
End Sub
