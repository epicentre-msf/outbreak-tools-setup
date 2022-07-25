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

      If Not bUpdateDictVarName Then bUpdateDictVarName = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderVarName).Range) Is Nothing)
      If Not bUpdateDictMainLabel Then bUpdateDictMainLabel = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderMainLabel).Range) Is Nothing)
      If Not bUpdateDictSubLabel Then bUpdateDictSubLabel = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderSubLabel).Range) Is Nothing)
      If Not bUpdateDictNote Then bUpdateDictNote = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderNote).Range) Is Nothing)
      If Not bUpdateDictSheetName Then bUpdateDictSheetName = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderSheetName).Range) Is Nothing)
      If Not bUpdateDictMainSection Then bUpdateDictMainSection = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderMainSection).Range) Is Nothing)
      If Not bUpdateDictSubSection Then bUpdateDictSubSection = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderSubSection).Range) Is Nothing)
      If Not bUpdateDictFormula Then bUpdateDictFormula = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderFormula).Range) Is Nothing)
      If Not bUpdateDictMessage Then bUpdateDictMessage = Not (Intersect(Target, sheetDictionary.ListObjects(C_sTabDictionary).ListColumns(C_sDictHeaderMessage).Range) Is Nothing)
End Sub



