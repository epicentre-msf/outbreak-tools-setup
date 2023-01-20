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
     If Updated("DictType") Or Updated("DictVarName") Or Updated("DictControl") Then
        UpdateVariablesList
        UpdateValue False, "DictType"
        UpdateValue False, "DictVarName"
        UpdateValue False, "DictControl"
     End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)

  If Not Updated("AnaGS_SL") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabGS).ListColumns(C_sAnaHeaderSL).Range) Is Nothing), "AnaGS_SL"
  If Not Updated("AnaGS_SF") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabGS).ListColumns(C_sAnaHeaderSF).Range) Is Nothing), "AnaGS_SF"

  If Not Updated("AnaUA_SC") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabUA).ListColumns(C_sAnaHeaderSC).Range) Is Nothing), "AnaUA_SC"
  If Not Updated("AnaUA_SL") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabUA).ListColumns(C_sAnaHeaderSL).Range) Is Nothing), "AnaUA_SL"
  If Not Updated("AnaUA_SF") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabUA).ListColumns(C_sAnaHeaderSF).Range) Is Nothing), "AnaUA_SF"


  If Not Updated("AnaBA_SC") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabBA).ListColumns(C_sAnaHeaderSC).Range) Is Nothing), "AnaBA_SC"
  If Not Updated("AnaBA_SL") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabBA).ListColumns(C_sAnaHeaderSL).Range) Is Nothing), "AnaBA_SL"
  If Not Updated("AnaBA_SF") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabBA).ListColumns(C_sAnaHeaderSF).Range) Is Nothing), "AnaBA_SF"

  If Not Updated("AnaTA_SF") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabTA).ListColumns(C_sAnaHeaderSF).Range) Is Nothing), "AnaTA_SF"
  If Not Updated("AnaTA_SL") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabTA).ListColumns(C_sAnaHeaderSL).Range) Is Nothing), "AnaTA_SL"
  If Not Updated("AnaTA_SC") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabTA).ListColumns(C_sAnaHeaderSC).Range) Is Nothing), "AnaTA_SC"


  'Graph on time series labels and title
  If Not Updated("AnaGTS_GL") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabGTS).ListColumns(C_sAnaHeaderGL).Range) Is Nothing), "AnaGTS_GL"
  If Not Updated("AnaGTS_GT") Then UpdateValue Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabGTSLab).ListColumns(C_sAnaHeaderGT).Range) Is Nothing), "AnaGTS_GT"


  'Update the choices and percentage on graph table
  If Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabGTS).ListColumns(4).DataBodyRange) Is Nothing) Then
      AddGraphOptions Target
  End If
End Sub

Private Sub Worksheet_Deactivate()
    sheetLists.Range("LastWorkSheet").Value = Me.Name
End Sub




