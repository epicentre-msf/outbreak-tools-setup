Attribute VB_Name = "Translation"
Option Explicit

Public Responses As Byte

'Write one value to the  translation table
Sub WriteTranslate(sLabel As String, Optional iColStart As Integer = 2)

    Dim Rng As Range
    Dim iRow As Long
    Dim sLab As String

    Dim iLineWrite As Long

    Set Rng = sheetTranslation.ListObjects(C_sTabTranslations).ListColumns(1).Range
    iLineWrite = Rng.Rows.Count + Rng.Row
    sLab = Application.WorksheetFunction.Trim(sLabel)
    If Not Rng.Find(What:=sLab, lookAt:=xlWhole, MatchCase:=True) Is Nothing Then
        iRow = Rng.Find(What:=sLab, lookAt:=xlWhole, MatchCase:=False).Row
        sheetTranslation.Cells(iRow, iColStart - 1).Value = 1
    Else
        sheetTranslation.Cells(iLineWrite, iColStart).Value = sLab
        sheetTranslation.Cells(iLineWrite, iColStart - 1).Value = 1
    End If

    Set Rng = Nothing
End Sub

'Split a formula to extract values inside the "", and add them to the translation table
Public Sub SplitAndWriteFormula(sFormula As String)

    Dim sText As String
    Dim iStart As Long
    Dim i As Long

    sText = Replace(sFormula, Chr(34) & Chr(34), "")

    If InStr(1, sText, Chr(34), 1) > 0 Then
        For i = 1 To Len(sText)
            If Mid(sText, i, 1) = Chr(34) Then
                If iStart = 0 Then
                    iStart = i + 1
                Else
                    Call WriteTranslate(Mid(sText, iStart, i - iStart))
                    iStart = 0
                End If
            End If
        Next
    End If
End Sub

'Extract values for one column with characters or formulas in it
Sub WriteColumn(Rng As Range, Optional ContainsFormula As Boolean = False)
    Dim c As Range 'cell value

    If Not Rng Is Nothing Then

        If ContainsFormula Then
            For Each c In Rng
                Call SplitAndWriteFormula(c.Value)
            Next
        Else
            For Each c In Rng
                If Not c Is Nothing Then Call WriteTranslate(c.Value)
            Next
        End If

    End If

    Set c = Nothing
End Sub


'Write values for the dictionary sheet

Sub WriteSheetColumn(Lo As ListObject, sColName As String, Optional ContainsFormula As Boolean = False)

    Dim iCol As Integer 'column to add to translation table
    Dim HeaderRng As Range 'Range of Headers to translate
    Dim ColumnRng As Range 'Range of column to Translate

    Set HeaderRng = Lo.HeaderRowRange

    'Make a test and exit sub
    If HeaderRng.Find(What:=sColName, lookAt:=xlWhole) Is Nothing Then
        MsgBox "Column " & sColName & " not found", vbOkOnly
        Exit Sub
    End If

    iCol = HeaderRng.Find(What:=sColName, lookAt:=xlWhole).Column - HeaderRng.Column + 1
    Set ColumnRng = Lo.ListColumns(iCol).DataBodyRange
    Call WriteColumn(ColumnRng, ContainsFormula)
End Sub

Sub WriteDictionary()

    Dim DictLo As ListObject
    Set DictLo = sheetDictionary.ListObjects(C_sTabDictionary)

    'Main Label
    WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderMainLabel

    'Sub Label
    WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderSubLabel

    'Note
    WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderNote

    'Sheet Name
    WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderSheetName

    'Main section
    WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderMainSection

    'Sub-section
    WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderSubSection

    'Formula

    WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderFormula, ContainsFormula:=True

    'Message
    WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderMessage

    Set DictLo = Nothing

End Sub


'Write values for the choices sheet

Sub WriteChoice()

    Dim ChoLo As ListObject

    Set ChoLo = SheetChoice.ListObjects(C_sTabChoices)

    'Label short
    WriteSheetColumn Lo:=ChoLo, sColName:=C_sChoHeaderLabelShort

    'Label
    WriteSheetColumn Lo:=ChoLo, sColName:=C_sChoHeaderLabel

    Set ChoLo = Nothing
End Sub


'Write values for the export sheet

Sub WriteExport()

    Dim ExpLo As ListObject

    Set ExpLo = sheetExport.ListObjects(C_sTabExports)
    'Label short
    WriteSheetColumn Lo:=ExpLo, sColName:=C_sExportHeaderLabelButton

    Set ExpLo = Nothing
End Sub


'Write values for the Analysis

Sub WriteAnalysis()

    Dim AnaLo As ListObject

    'Global summary, First column
    Set AnaLo = sheetAnalysis.ListObjects(C_sTabGS)

    WriteSheetColumn Lo:=AnaLo, sColName:=C_sAnaHeaderSL
    WriteSheetColumn Lo:=AnaLo, sColName:=C_sAnaHeaderSF, ContainsFormula:=True

    'Univariate analysis column
    Set AnaLo = sheetAnalysis.ListObjects(C_sTabUA)
    WriteSheetColumn Lo:=AnaLo, sColName:=C_sAnaHeaderSC
    WriteSheetColumn Lo:=AnaLo, sColName:=C_sAnaHeaderSF, ContainsFormula:=True
    WriteSheetColumn Lo:=AnaLo, sColName:=C_sAnaHeaderSL

    'Bivariate analysis column
    Set AnaLo = sheetAnalysis.ListObjects(C_sTabBA)
     WriteSheetColumn Lo:=AnaLo, sColName:=C_sAnaHeaderSC
    WriteSheetColumn Lo:=AnaLo, sColName:=C_sAnaHeaderSF, ContainsFormula:=True
    WriteSheetColumn Lo:=AnaLo, sColName:=C_sAnaHeaderSL

    Set AnaLo = Nothing

End Sub


'Write values for the Analysis sheet


Sub AddLabelsToTranslationTable(Optional sType As String)

    Dim iRow As Long 'where the translation table starts
    Dim iLastRow As Long 'where the translationt table ends
    Dim iNbColumns As Long 'Number of columns of the header range
    Dim iLastColumn As Long 'Last column of the Translation list object
    Dim iColStart As Integer
    Dim TransLo As ListObject
    Dim i As Long
    Dim sMessage As String
    Dim iNbBLanks As Long
    Dim idelRow As Long


    BeginWork
    sheetTranslation.Unprotect C_sPassword
    Application.Cursor = xlWait


    Set TransLo = sheetTranslation.ListObjects(C_sTabTranslations)



    iRow = TransLo.DataBodyRange.Row
    iLastRow = TransLo.DataBodyRange.Rows.Count + iRow

    iColStart = TransLo.Range.Column
    iNbColumns = TransLo.HeaderRowRange.Columns.Count
    iLastColumn = iColStart + iNbColumns - 1

    'Resize the listobject if needed to include new column if there is one
    With sheetTranslation
        i = .Cells(iRow - 1, .Columns.Count).End(xlToLeft).Column

        If i > iLastColumn Then
            TransLo.Resize Range(.Cells(iRow - 1, iColStart), .Cells(iLastRow - 1, i))

            'update some values previously set
            bUpdate = True
            iNbColumns = TransLo.HeaderRowRange.Columns.Count
            iLastColumn = iColStart + iNbColumns - 1
        End If

    End With


    sheetTranslation.Columns(iColStart - 1).ClearContents

    If bUpdate Then

        'Write label for each of the sheets
        Call WriteDictionary
        Call WriteChoice
        Call WriteExport
        Call WriteAnalysis

        'Delete rows not found
        For i = iRow To iLastRow
            If sheetTranslation.Cells(i, iColStart - 1).Value <> 1 Or sheetTranslation.Cells(i, iColStart).Value = vbNullString Then
                idelRow = sheetTranslation.Cells(i, iColStart - 1).Row
                sheetTranslation.Rows(idelRow).EntireRow.Delete
            End If
        Next

        sheetTranslation.Range(Cells(iRow, iColStart - 1), Cells(iLastRow, iColStart - 1)).ClearContents

        'sort the first column
        sheetTranslation.Sort.SortFields.Clear
        TransLo.DataBodyRange.Sort key1:=Cells(iRow, iColStart), Header:=xlYes, Orientation:=xlTopToBottom

    End If


    sMessage = vbNullString

    'Count blank labels
    For i = 1 To iNbColumns - 1

        iNbBLanks = Application.WorksheetFunction.CountBlank(TransLo.ListColumns(i + 1).Range)

        If iNbBLanks > 0 Then
            sMessage = sMessage & iNbBLanks & " labels are missing for column " & _
                sheetTranslation.Cells(iRow - 1, iColStart + i).Value & "." & Chr(10)
        End If
    Next


    Application.Cursor = xlDefault


    If sMessage <> vbNullString Then
        If sType = "Close" Then
            Responses = MsgBox(sMessage & Chr(10) & "Do you really want to close the workbook ?", vbYesNo, "verification of translations")
        Else
            MsgBox sMessage, vbCritical, "verification of translations"
        End If
    End If

    bUpdate = False

    sheetTranslation.Columns(iColStart - 1).ClearContents

    ActiveWorkbook.Save

    'Lock the first Column and protect the sheet
    Call LockFirstColumn
    Call ProtectTranslationSheet

    EndWork
End Sub


Sub LockFirstColumn()
    Dim iLastRow As Integer
    Dim rngFirstColumn As Range

    iLastRow = sheetTranslation.Cells(Rows.Count, 2).End(xlUp).Row

    With sheetTranslation
        Set rngFirstColumn = .Range(.Cells(5, 2), .Cells(iLastRow, 2))
    End With

    rngFirstColumn.Locked = True
End Sub


Sub ProtectTranslationSheet()
    sheetTranslation.Protect Password:=C_sPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=False, AllowInsertingHyperlinks:=True, _
        AllowDeletingRows:=False, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True, AllowDeletingColumns:=True
End Sub

Sub UpdateTranslation()
    Call AddLabelsToTranslationTable
End Sub
