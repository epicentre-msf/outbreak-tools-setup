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

    If bUpdate Then Exit Sub

    bUpdate = Not (Intersect(Target, SheetChoice.ListObjects(C_sTabChoices). _
                ListColumns(C_sChoHeaderLabelShort).Range) Is Nothing And _
                Intersect(Target, SheetChoice.ListObjects(C_sTabChoices). _
                ListColumns(C_sChoHeaderLabel).Range) Is Nothing)

End Sub



