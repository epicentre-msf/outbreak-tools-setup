VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheetAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Call UpdateVariablesList
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)

  If Not bUpdateAnaGS_SL Then bUpdateAnaGS_SL = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabGS).ListColumns(C_sAnaHeaderSL).Range) Is Nothing)
  If Not bUpdateAnaGS_SF Then bUpdateAnaGS_SF = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabGS).ListColumns(C_sAnaHeaderSF).Range) Is Nothing)

  If Not bUpdateAnaUA_SC Then bUpdateAnaUA_SC = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabUA).ListColumns(C_sAnaHeaderSC).Range) Is Nothing)
  If Not bUpdateAnaUA_SL Then bUpdateAnaUA_SL = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabUA).ListColumns(C_sAnaHeaderSL).Range) Is Nothing)
  If Not bUpdateAnaUA_SF Then bUpdateAnaUA_SF = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabUA).ListColumns(C_sAnaHeaderSF).Range) Is Nothing)


  If Not bUpdateAnaBA_SC Then bUpdateAnaBA_SC = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabBA).ListColumns(C_sAnaHeaderSC).Range) Is Nothing)
  If Not bUpdateAnaBA_SL Then bUpdateAnaBA_SL = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabBA).ListColumns(C_sAnaHeaderSL).Range) Is Nothing)
  If Not bUpdateAnaBA_SF Then bUpdateAnaBA_SF = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabBA).ListColumns(C_sAnaHeaderSF).Range) Is Nothing)

  If Not bUpdateAnaTA_SF Then bUpdateAnaTA_SF = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabTA).ListColumns(C_sAnaHeaderSF).Range) Is Nothing)
  If Not bUpdateAnaTA_SL Then bUpdateAnaTA_SL = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabTA).ListColumns(C_sAnaHeaderSL).Range) Is Nothing)
  If Not bUpdateAnaTA_SC Then bUpdateAnaTA_SC = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabTA).ListColumns(C_sAnaHeaderSC).Range) Is Nothing)

End Sub




