Attribute VB_Name = "Tools"
Option Explicit

Sub BeginWork()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual

End Sub

Sub EndWork()

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub ResizeLo(Lo As ListObject, Optional AddRows As Boolean = True)

    'Begining of the tables
    Dim iRowHeader As Long
    Dim iColHeader  As Long
    Dim i As Long
    Dim iFirstColumnRow As Long

    'End of the listobject table
    Dim iRowsEnd As Long
    Dim iColsEnd As Long



    ActiveSheet.Unprotect C_sPassword

    'Rows and columns at the begining of the table to resize
    iRowHeader = Lo.Range.Row
    iColHeader = Lo.Range.Column

    'Rows and Columns at the end of the Table to resize
    iRowsEnd = iRowHeader + Lo.Range.Rows.Count
    iColsEnd = Lo.Range.Columns.Count

    If Not AddRows Then
        ActiveSheet.Rows(iRowsEnd + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        iFirstColumnRow = ActiveSheet.Cells(iRowHeader, iColHeader).End(xlDown).Row

        If iFirstColumnRow < iRowsEnd Then
            For i = iRowsEnd To iFirstColumnRow + 1 Step -1
                ActiveSheet.Rows(i).EntireRow.Delete
            Next

            iRowsEnd = iFirstColumnRow
        End If
    Else
        For i = 1 To C_iNbLinesLLData + 1
            ActiveSheet.Rows(iRowsEnd).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next
        iRowsEnd = iRowsEnd + C_iNbLinesLLData
    End If

    Lo.Resize Range(Cells(iRowHeader, iColHeader), Cells(iRowsEnd, iColHeader + iColsEnd - 1))


    Call ProtectSheet

End Sub

Sub ProtectSheet()
    ActiveSheet.Protect Password:=C_sPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                                 , AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                        AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
                        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                        AllowUsingPivotTables:=True

End Sub

'Resize the dictionary table object
Public Sub AddRowsDict()
    Call ResizeLo(sheetDictionary.ListObjects(C_sTabDictionary))
End Sub

'Resize the choices table object
Public Sub AddRowsChoices()
    Call ResizeLo(SheetChoice.ListObjects(C_sTabChoices))
End Sub

Public Sub AddRowsGS()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabGS))
End Sub

Public Sub AddRowsUA()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabUA))
End Sub

Public Sub AddRowsBA()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabBA))
End Sub

'resize the tables (delete empty rows at the bottom)----------------------------

'Resize the dictionary table object
Public Sub RemoveRowsDict()
    BeginWork

    Call ResizeLo(sheetDictionary.ListObjects(C_sTabDictionary), AddRows:=False)

    EndWork
End Sub

'Resize the choices table object
Public Sub RemoveRowsChoices()
    BeginWork

    Call ResizeLo(SheetChoice.ListObjects(C_sTabChoices), AddRows:=False)

    EndWork
End Sub

Public Sub RemoveRowsGS()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabGS), AddRows:=False)
End Sub

Public Sub RemoveRowsUA()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabUA), AddRows:=False)
End Sub

Public Sub RemoveRowsBA()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabBA), AddRows:=False)
End Sub

Public Sub AddRowsAna()

    BeginWork

    Select Case Application.WorksheetFunction.Trim(sheetAnalysis.Range("RNG_table_modify").Value)

    Case C_sModifyGS

        Call AddRowsGS

    Case C_sModifyUA

        Call AddRowsUA

    Case C_sModifyBA

        Call AddRowsBA

    Case Else

        AddRowsGS
        AddRowsUA
        AddRowsBA

    End Select

    EndWork

End Sub

Public Sub RemoveRowsAna()

    BeginWork

    Select Case Application.WorksheetFunction.Trim(sheetAnalysis.Range("RNG_table_modify").Value)

    Case C_sModifyGS

        Call RemoveRowsGS

    Case C_sModifyUA

        Call RemoveRowsUA

    Case C_sModifyBA

        Call RemoveRowsBA

    Case Else

        RemoveRowsGS
        RemoveRowsUA
        RemoveRowsBA

    End Select

    EndWork

End Sub

'Update the list of variables when moving to analysis

Sub UpdateVariablesList()

    Dim iControlColumn As Integer
    Dim iDictLength As Long
    Dim i As Long 'counters for the variables and for the list
    Dim j As Long
    Dim iVarColumn As Integer
    Dim iListVarColumn As Integer
    Dim Rng As Range


    BeginWork

    Set Rng = sheetDictionary.ListObjects(C_sTabDictionary).HeaderRowRange

    If Rng.Find(C_sDictHeaderControl, lookAt:=xlWhole, MatchCase:=True) Is Nothing Or _
         Rng.Find(C_sDictHeaderVarName, lookAt:=xlWhole, MatchCase:=True) Is Nothing Then
        Exit Sub
    Else
        iVarColumn = Rng.Find(C_sDictHeaderVarName, lookAt:=xlWhole, MatchCase:=True).Column
        iControlColumn = Rng.Find(C_sDictHeaderControl, lookAt:=xlWhole, MatchCase:=True).Column

    End If

    iDictLength = sheetDictionary.Cells(Rng.Row, iVarColumn).End(xlDown).Row
    iListVarColumn = sheetLists.ListObjects(C_sTabVarList).Range.Column

    j = 1
    With sheetDictionary
        For i = Rng.Row + 1 To iDictLength

            If .Cells(i, iControlColumn).Value = C_sDictControlChoice Or _
                .Cells(i, iControlColumn).Value = C_sDictControlFormulaChoice Then
                j = j + 1
                sheetLists.Cells(j, iListVarColumn).Value = .Cells(i, iVarColumn).Value
            End If
        Next
    End With

    EndWork

End Sub
