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


Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Responses = 0

    On Error GoTo ManageError

    If bUpdate Then Call AddLabelsToTranslationTable("Close")

    If Responses = vbNo Then
        Cancel = True 'cancel closing the workbook
        Exit Sub
    End If

    ActiveWorkbook.Save

ManageError:
    Resume Next

End Sub


