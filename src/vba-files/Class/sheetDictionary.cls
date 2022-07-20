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

    If bUpdate Then Exit Sub

    bUpdate = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderVarName).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderControl).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderMainLabel).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderSubLabel).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderNote).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderSheetName).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderMainSection).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderSubSection).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderFormula).Range) Is Nothing And _
      Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary). _
        ListColumns(C_sDictHeaderMessage).Range) Is Nothing)
End Sub



