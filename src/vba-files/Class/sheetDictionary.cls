VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheetDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

  If Not Updated("DictMainLabel") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderMainLabel).Range) Is Nothing), "DictMainLabel"
  If Not Updated("DictSubLabel") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderSubLabel).Range) Is Nothing), "DictSubLabel"
  If Not Updated("DictNote") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderNote).Range) Is Nothing), "DictNote"
  If Not Updated("DictSheetName") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderSheetName).Range) Is Nothing), "DictSheetName"
  If Not Updated("DictMainSection") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderMainSection).Range) Is Nothing), "DictMainSection"
  If Not Updated("DictSubSection") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderSubSection).Range) Is Nothing), "DictSubSection"
  If Not Updated("DictFormula") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderFormula).Range) Is Nothing), "DictFormula"
  If Not Updated("DictMessage") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderMessage).Range) Is Nothing), "DictMessage"
  If Not Updated("DictType") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderType).Range) Is Nothing), "DictType"
  If Not Updated("DictVarName") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderVarName).Range) Is Nothing), "DictVarName"
  If Not Updated("DictControl") Then UpdateValue Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderControl).Range) Is Nothing), "DictControl"
End Sub



Private Sub Worksheet_Deactivate()
    sheetLists.Range("LastWorkSheet").Value = Me.Name
End Sub
