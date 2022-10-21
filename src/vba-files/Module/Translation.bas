Attribute VB_Name = "Translation"
Option Explicit

Public Responses As Byte

'Write one value to the  translation table
Sub WriteTranslate(sLabel As String, sIndicator As String, Optional iColStart As Integer = 2)

    Dim Rng As Range
    Dim iRow As Long
    Dim sLab As String

    Dim iLineWrite As Long

    Set Rng = sheetTranslation.ListObjects(C_sTabTranslations).ListColumns(1).Range

    If Application.WorksheetFunction.CountBlank(Rng) = Rng.Rows.Count - 1 Then
        iLineWrite = Rng.Row + 1
    Else
        iLineWrite = Rng.Rows.Count + Rng.Row
    End If

    sLab = Application.WorksheetFunction.Trim(sLabel)
    If Not Rng.Find(What:=sLab, lookAt:=xlWhole, MatchCase:=True) Is Nothing Then
        iRow = Rng.Find(What:=sLab, lookAt:=xlWhole, MatchCase:=True).Row
        sheetTranslation.Cells(iRow, iColStart - 1).Value = sIndicator & "_" & sheetLists.Range("nbtimes").Value
    Else
        sheetTranslation.Cells(iLineWrite, iColStart).Value = sLab
        sheetTranslation.Cells(iLineWrite, iColStart - 1).Value = sIndicator & "_" & sheetLists.Range("nbtimes").Value
    End If

    Set Rng = Nothing
End Sub

'Split a formula to extract values inside the "", and add them to the translation table
Public Sub SplitAndWriteFormula(sFormula As String, sIndicator As String)

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
                    Call WriteTranslate(Mid(sText, iStart, i - iStart), sIndicator:=sIndicator)
                    iStart = 0
                End If
            End If
        Next
    End If
End Sub

'Extract values for one column with characters or formulas in it
Sub WriteColumn(Rng As Range, sIndicator As String, Optional ContainsFormula As Boolean = False)
    Dim c As Range 'cell value

    If Not Rng Is Nothing Then

        If ContainsFormula Then
            For Each c In Rng
                Call SplitAndWriteFormula(c.Value, sIndicator:=sIndicator)
            Next
        Else
            For Each c In Rng
                If Not c Is Nothing Then Call WriteTranslate(c.Value, sIndicator:=sIndicator)
            Next
        End If

    End If

    Set c = Nothing
End Sub


'Write values for the dictionary sheet

Sub WriteSheetColumn(Lo As ListObject, sColName As String, sIndicator As String, Optional ContainsFormula As Boolean = False)

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
    Call WriteColumn(Rng:=ColumnRng, sIndicator:=sIndicator, ContainsFormula:=ContainsFormula)
End Sub

Sub WriteDictionary()

    Dim DictLo As ListObject
    Dim sIndicator As String
    Set DictLo = sheetDictionary.ListObjects(C_sTabDictionary)
    sIndicator = "Dict"
    
    'Main Label
    If Updated("DictMainLabel") Then WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderMainLabel, sIndicator:=sIndicator & C_sDictHeaderMainLabel

    'Sub Label
    If Updated("DictSubLabel") Then WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderSubLabel, sIndicator:=sIndicator & C_sDictHeaderSubLabel

    'Note
    If Updated("DictNote") Then WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderNote, sIndicator:=sIndicator & C_sDictHeaderNote

    'Sheet Name
    If Updated("DictSheetName") Then WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderSheetName, sIndicator:=sIndicator & C_sDictHeaderSheetName

    'Main section
    If Updated("DictMainSection") Then WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderMainSection, sIndicator:=sIndicator & C_sDictHeaderMainSection

    'Sub-section
    If Updated("DictSubSection") Then WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderSubSection, sIndicator:=sIndicator & C_sDictHeaderSubSection

    'Formula

    If Updated("DictFormula") Then WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderFormula, sIndicator:=sIndicator & C_sDictHeaderFormula, ContainsFormula:=True

    'Message
    If Updated("DictMessage") Then WriteSheetColumn Lo:=DictLo, sColName:=C_sDictHeaderMessage, sIndicator:=sIndicator & C_sDictHeaderMessage

    Set DictLo = Nothing

End Sub


'Write values for the choices sheet

Sub WriteChoice()

    Dim ChoLo As ListObject
    Dim sIndicator As String
    Set ChoLo = SheetChoice.ListObjects(C_sTabChoices)
    sIndicator = "Choi"

    'Label
    If Updated("ChoiLabel") Then WriteSheetColumn Lo:=ChoLo, sIndicator:=sIndicator & C_sChoHeaderLabel, sColName:=C_sChoHeaderLabel

    Set ChoLo = Nothing
End Sub


'Write values for the export sheet

Sub WriteExport()

    Dim ExpLo As ListObject
    Dim sIndicator As String

    sIndicator = "Exp"

    Set ExpLo = sheetExport.ListObjects(C_sTabExports)
    'Label short
    If Updated("Exp") Then WriteSheetColumn Lo:=ExpLo, sIndicator:=sIndicator & C_sExportHeaderLabelButton, sColName:=C_sExportHeaderLabelButton

    Set ExpLo = Nothing
End Sub


'Write values for the Analysis

Sub WriteAnalysis()

    Dim AnaLo As ListObject
    Dim sIndicator As String


    'Global summary, First column
    Set AnaLo = sheetAnalysis.ListObjects(C_sTabGS)
    sIndicator = "AnaGS"
    If Updated("AnaGS_SL") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSL, sColName:=C_sAnaHeaderSL
    If Updated("AnaGS_SF") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSF, sColName:=C_sAnaHeaderSF, ContainsFormula:=True

    'Univariate analysis column
    sIndicator = "AnaUA"
    Set AnaLo = sheetAnalysis.ListObjects(C_sTabUA)
    If Updated("AnaUA_SC") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSC, sColName:=C_sAnaHeaderSC
    If Updated("AnaUA_SL") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSL, sColName:=C_sAnaHeaderSL
    If Updated("AnaUA_SF") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSF, sColName:=C_sAnaHeaderSF, ContainsFormula:=True

    'Bivariate analysis column
    sIndicator = "AnaBA"
    Set AnaLo = sheetAnalysis.ListObjects(C_sTabBA)
    If Updated("AnaBA_SC") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSC, sColName:=C_sAnaHeaderSC
    If Updated("AnaBA_SF") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSF, sColName:=C_sAnaHeaderSF, ContainsFormula:=True
    If Updated("AnaBA_SL") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSL, sColName:=C_sAnaHeaderSL

     'Time Series Analysis column
    Set AnaLo = sheetAnalysis.ListObjects(C_sTabTA)
    sIndicator = "AnaTA"
    If Updated("AnaTA_SC") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSC, sColName:=C_sAnaHeaderSC
    If Updated("AnaTA_SL") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSL, sColName:=C_sAnaHeaderSL
    If Updated("AnaTA_SF") Then WriteSheetColumn Lo:=AnaLo, sIndicator:=sIndicator & C_sAnaHeaderSF, sColName:=C_sAnaHeaderSF, ContainsFormula:=True

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
    Dim Rng As Range
    Dim RngSort As Range

    On Error GoTo errHand

    BeginWork
    Application.Cursor = xlWait
    sheetTranslation.Unprotect C_sPassword

    Set TransLo = sheetTranslation.ListObjects(C_sTabTranslations)


    iRow = TransLo.Range.Row + 1
    iLastRow = TransLo.Range.Rows.Count + iRow

    iColStart = TransLo.Range.Column
    iNbColumns = TransLo.HeaderRowRange.Columns.Count
    iLastColumn = iColStart + iNbColumns - 1

    'Resize the listobject if needed to include new column if there is one
    With sheetTranslation
        i = .Cells(iRow - 1, .Columns.Count).End(xlToLeft).Column

        If i > iLastColumn Then
            TransLo.Resize Range(.Cells(iRow - 1, iColStart), .Cells(iLastRow - 1, i))
            iNbColumns = TransLo.HeaderRowRange.Columns.Count
            iLastColumn = iColStart + iNbColumns - 1
        End If

    End With


    If sheetLists.Range("nbtimes").Value = 0 Then sheetTranslation.Columns(iColStart - 1).ClearContents
    sheetLists.Range("nbtimes").Value = sheetLists.Range("nbtimes").Value + 1

    'Write label for each of the sheets
    Call WriteDictionary
    Call WriteChoice
    Call WriteExport
    Call WriteAnalysis

    'Re-initialize the lastrow
    iLastRow = TransLo.Range.Rows.Count - 1 + iRow


     With sheetTranslation.Columns(iColStart - 1)
        .Interior.Color = vbWhite
         .Font.Color = vbWhite
         .FormulaHidden = True
     End With

     'Delete rows not found
    Call DeleteUnfoundLabels(iColStart, iRow - 1, iLastRow)

    'sort the first column
    sheetTranslation.Sort.SortFields.Clear

    'Unlist to sort using update values
    TransLo.Unlist

    With sheetTranslation
        .Cells(iRow - 1, iColStart - 1).Value = "TestValues"
        Set Rng = Range(.Cells(iRow - 1, iColStart - 1), .Cells(iLastRow, iLastColumn))
        Set RngSort = Range(.Cells(iRow - 1, iColStart), .Cells(iLastRow, iColStart))
    End With

    Rng.Sort key1:=RngSort, Header:=xlYes, Orientation:=xlTopToBottom

    With sheetTranslation
        iLastRow = .Cells(.Rows.Count, iColStart).End(xlUp).Row
        Set Rng = Range(.Cells(iRow - 1, iColStart), .Cells(iLastRow, iLastColumn))
        .ListObjects.Add(xlSrcRange, Rng, , xlYes, , "TableStyleLight8").Name = C_sTabTranslations
        Set TransLo = .ListObjects(C_sTabTranslations)

        Range(.Cells(iRow - 1, iColStart), .Cells(iLastRow, iLastColumn)).Locked = False
    End With

    sMessage = vbNullString

    'Count blank labels
    For i = 1 To iNbColumns - 1

        iNbBLanks = Application.WorksheetFunction.CountBlank(TransLo.ListColumns(i + 1).Range)

        If iNbBLanks > 0 Then
            sMessage = sMessage & iNbBLanks & " labels are missing for column " & _
                sheetTranslation.Cells(iRow - 1, iColStart + i).Value & "." & Chr(10)
        End If
    Next

    If sMessage <> vbNullString Then
        If sType = "Close" Then
            Responses = MsgBox(sMessage & Chr(10) & "Do you really want to close the workbook ?", vbYesNo, "verification of translations")
        Else
            MsgBox sMessage, vbCritical, "verification of translations"
        End If
    Else
        MsgBox "All labels are translated, there is no missing label left", vbInformation + vbOkOnly, "DONE!"
    End If

    Call SetAllUpdates(toValue:=False)

    ActiveWorkbook.Save
    'Lock the first Column and protect the sheet
    Call LockFirstColumn
    Call ProtectTranslationSheet

    Set Rng = Nothing
    Set RngSort = Nothing
    Application.Cursor = xlDefault
    EndWork

    Exit Sub

errHand:
    MsgBox "Unexpected error" & vbNewLine & Err.Description, vbCritical
    Call LockFirstColumn
    Call ProtectTranslationSheet
    Call SetAllUpdates(True)
    Application.Cursort = xlDefault
    EndWork
    Exit Sub
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


Sub DeleteUnfoundLabels(iColStart As Integer, iStartRow, iLastRow As Long)
    Dim i As Long
    Dim sValue As String
    Dim sValueFirstcol As String
    Dim DeleteRow As Boolean
    Dim iLast As Long

    i = iStartRow
    iLast = iLastRow

    Do While (i < iLast And i > iStartRow - 1)
        DeleteRow = False
        sValue = sheetTranslation.Cells(i + 1, iColStart - 1).Value

        sValueFirstcol = sheetTranslation.Cells(i + 1, iColStart).Value

        If sValueFirstcol = vbNullString Then DeleteRow = True

        If Not DeleteRow And Not (InStr(1, sValue, CStr(sheetLists.Range("nbtimes").Value)) > 0) Then
           'The update have been done previously
            DeleteRow = _
                (Updated("DictMainLabel") And InStr(1, sValue, "Dict" & C_sDictHeaderMainLabel) > 0) Or _
                (Updated("DictSubLabel") And InStr(1, sValue, "Dict" & C_sDictHeaderSubLabel) > 0) Or _
                (Updated("DictNote") And InStr(1, sValue, "Dict" & C_sDictHeaderNote) > 0) Or _
                (Updated("DictSheetName") And InStr(1, sValue, "Dict" & C_sDictHeaderSheetName) > 0) Or _
                (Updated("DictMainSection") And InStr(1, sValue, "Dict" & C_sDictHeaderMainSection) > 0) Or _
                (Updated("DictSubSection") And InStr(1, sValue, "Dict" & C_sDictHeaderSubSection) > 0) Or _
                (Updated("DictFormula") And InStr(1, sValue, "Dict" & C_sDictHeaderFormula) > 0) Or _
                (Updated("DictMessage") And InStr(1, sValue, "Dict" & C_sDictHeaderMessage) > 0) Or _
                (Updated("ChoiLabel") And InStr(1, sValue, "Choi" & C_sChoHeaderLabel) > 0) Or _
                (Updated("Exp") And InStr(1, sValue, "Exp" & C_sExportHeaderLabelButton) > 0) Or _
                (Updated("AnaGS_SL") And InStr(1, sValue, "AnaGS" & C_sAnaHeaderSL) > 0) Or _
                (Updated("AnaGS_SF") And InStr(1, sValue, "AnaGS" & C_sAnaHeaderSF) > 0) Or _
                (Updated("AnaUA_SC") And InStr(1, sValue, "AnaUA" & C_sAnaHeaderSC) > 0) Or _
                (Updated("AnaUA_SL") And InStr(1, sValue, "AnaUA" & C_sAnaHeaderSL) > 0) Or _
                (Updated("AnaUA_SF") And InStr(1, sValue, "AnaUA" & C_sAnaHeaderSF) > 0) Or _
                (Updated("AnaBA_SC") And InStr(1, sValue, "AnaBA" & C_sAnaHeaderSC) > 0) Or _
                (Updated("AnaBA_SL") And InStr(1, sValue, "AnaBA" & C_sAnaHeaderSL) > 0) Or _
                (Updated("AnaBA_SF") And InStr(1, sValue, "AnaBA" & C_sAnaHeaderSF) > 0) Or _
                (Updated("AnaTA_SC") And InStr(1, sValue, "AnaTA" & C_sAnaHeaderSC) > 0) Or _
                (Updated("AnaTA_SL") And InStr(1, sValue, "AnaTA" & C_sAnaHeaderSL) > 0) Or _
                (Updated("AnaTA_SF") And InStr(1, sValue, "AnaTA" & C_sAnaHeaderSF) > 0) Or _
                sValue = vbNullString
        End If


        If DeleteRow Then
            On Error Resume Next
                sheetTranslation.Rows(i + 1).EntireRow.Delete
            On Error GoTo 0
            iLast = iLast - 1
        End If

       If Not DeleteRow Then i = i + 1
    Loop
End Sub



Sub UpdateTranslation()
    Dim test As Byte

    ' If NoUpdate() Then

    '     Test = MsgBox("You haven't done any update, but you want to re-import translation labels. Do you really want to re-import all values ?", vbYesNo)

    '     If Test = vbYes Then
    '         SetAllUpdates (True)
    '     End If
    ' End If

    Call AddLabelsToTranslationTable
End Sub

