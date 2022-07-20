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

    If bUpdate Then Exit Sub

    bUpdate = Not (Intersect(Target, sheetAnalysis.ListObjects(C_sTabGS). _
                ListColumns(C_sAnaHeaderSL).Range) Is Nothing And _
                Intersect(Target, sheetAnalysis.ListObjects(C_sTabUA). _
                ListColumns(C_sAnaHeaderSL).Range) Is Nothing And _
                Intersect(Target, sheetAnalysis.ListObjects(C_sTabUA). _
                ListColumns(C_sAnaHeaderSC).Range) Is Nothing And _
                Intersect(Target, sheetAnalysis.ListObjects(C_sTabBA). _
                ListColumns(C_sAnaHeaderSL).Range) Is Nothing And _
                Intersect(Target, sheetAnalysis.ListObjects(C_sTabBA). _
                ListColumns(C_sAnaHeaderSC).Range) Is Nothing)
End Sub




