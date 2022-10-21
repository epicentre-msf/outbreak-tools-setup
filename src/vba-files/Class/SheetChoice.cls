VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
 
    If Not Updated("ChoiLabel") Then UpdateValue Not (Intersect(Target, SheetChoice.ListObjects(C_sTabChoices).ListColumns(C_sChoHeaderLabel).Range) Is Nothing), "ChoiLabel"

End Sub



Private Sub Worksheet_Deactivate()
    sheetLists.Range("LastWorkSheet").Value = Me.Name
End Sub
