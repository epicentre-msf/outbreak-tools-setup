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

    If Not AddRows Then 'Remove rows
        ActiveSheet.Rows(iRowsEnd + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        'First Column Row is the last row of the first column of the listobject
        iFirstColumnRow = ActiveSheet.Cells(iRowHeader, iColHeader).End(xlDown).Row

        'If the listobject is empty, change the row end and start to resize to
        'only the first row

        If Not Lo.DataBodyRange Is Nothing Then
            If Application.WorksheetFunction.CountA(Lo.DataBodyRange) = 0 Then
                iFirstColumnRow = iRowHeader + 1
            End If
        End If


        If iFirstColumnRow < iRowsEnd Then
            For i = iRowsEnd To iFirstColumnRow + 1 Step -1
                ActiveSheet.Rows(i).EntireRow.Delete
            Next

            iRowsEnd = iFirstColumnRow
        End If
    Else 'Add rows
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

Public Sub AddRowsTA()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabTA))
End Sub

Public Sub AddRowsSA()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabSA))
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

Public Sub RemoveRowsSA()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabSA), AddRows:=False)
End Sub

Public Sub RemoveRowsTA()
    Call ResizeLo(sheetAnalysis.ListObjects(C_sTabTA), AddRows:=False)
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

    Case C_sModifySA

        Call AddRowsSA

    Case C_sModifyTA

        Call AddRowsTA

    Case Else

        AddRowsGS
        AddRowsUA
        AddRowsBA
        AddRowsTA
        AddRowsSA

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

     Case C_sModifySA

        Call RemoveRowsSA

    Case C_sModifyTA

        Call RemoveRowsTA

    Case Else

        RemoveRowsGS
        RemoveRowsUA
        RemoveRowsBA
        RemoveRowsTA
        RemoveRowsSA

    End Select

    EndWork

End Sub

'Update the list of variables when moving to analysis

Sub UpdateVariablesList()

    On Error GoTo errHand

    Dim iControlColumn As Integer
    Dim iTypeColumn As Integer
    Dim iDictLength As Long
    Dim i As Long 'counters for the variables and for the list
    Dim j As Long
    Dim k As Long
    Dim iVarColumn As Integer
    Dim iListVarColumn As Integer
    Dim iTimeVarColumn As Integer
    Dim Rng As Range


    BeginWork

    Set Rng = sheetDictionary.ListObjects(C_sTabDictionary).HeaderRowRange

    If Rng.Find(C_sDictHeaderControl, lookAt:=xlWhole, MatchCase:=True) Is Nothing Or _
         Rng.Find(C_sDictHeaderVarName, lookAt:=xlWhole, MatchCase:=True) Is Nothing Or _
         Rng.Find(C_sDictHeaderType, lookAt:=xlWhole, MatchCase:=True) Is Nothing Then
        Exit Sub
    Else
        iVarColumn = Rng.Find(C_sDictHeaderVarName, lookAt:=xlWhole, MatchCase:=True).Column
        iControlColumn = Rng.Find(C_sDictHeaderControl, lookAt:=xlWhole, MatchCase:=True).Column
        iTypeColumn = Rng.Find(C_sDictHeaderType, lookAt:=xlWhole, MatchCase:=True).Column
    End If

    With sheetDictionary
        iDictLength = .Cells(.Rows.Count, iVarColumn).End(xlUp).Row
    End With

    iListVarColumn = sheetLists.ListObjects(C_sTabVarList).Range.Column
    iTimeVarColumn = sheetLists.ListObjects(C_sTabTimeVar).Range.Column

    j = 1
    k = 1
    With sheetDictionary
        For i = Rng.Row + 1 To iDictLength

            If .Cells(i, iControlColumn).Value = C_sDictControlChoice Or _
                .Cells(i, iControlColumn).Value = C_sDictControlCaseWhen Then
                j = j + 1
                sheetLists.Cells(j, iListVarColumn).Value = .Cells(i, iVarColumn).Value
            End If

            'Add Dates vars
             If .Cells(i, iTypeColumn).Value = C_sDictTypeDate Then
                k = k + 1
                sheetLists.Cells(k, iTimeVarColumn).Value = .Cells(i, iVarColumn).Value
            End If
        Next
    End With

    EndWork

    Exit Sub

errHand:
    EndWork
    MsgBox Err.Description
    Exit Sub

End Sub

'Set All Updates to True / False

Sub SetAllUpdates(Optional toValue As Boolean = True)

    'Record updates for dictionary
     bUpdateDictVarName = toValue
     bUpdateDictMainLabel = toValue
     bUpdateDictSubLabel = toValue
     bUpdateDictNote = toValue
     bUpdateDictSheetName = toValue
     bUpdateDictMainSection = toValue
     bUpdateDictSubSection = toValue
     bUpdateDictFormula = toValue
     bUpdateDictMessage = toValue


    'Record updates for choices
     bUpdateChoiLabelShort = toValue
     bUpdateChoiLabel = toValue

    'Record updates for Exports
     bUpdateExp = toValue

    'Record updates for Analysis
     bUpdateAnaGS_SL = toValue
     bUpdateAnaGS_SF = toValue

     bUpdateAnaUA_SC = toValue
     bUpdateAnaUA_SL = toValue
     bUpdateAnaUA_SF = toValue

     bUpdateAnaBA_SC = toValue
     bUpdateAnaBA_SL = toValue
     bUpdateAnaBA_SF = toValue

     bUpdateAnaTA_SC = toValue
     bUpdateAnaTA_SL = toValue
     bUpdateAnaTA_SF = toValue


End Sub

