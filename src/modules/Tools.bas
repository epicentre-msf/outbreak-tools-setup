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

Sub ResizeLo(Lo As ListObject, Optional AddRows As Boolean = True, Optional totalRowCount As Long = 0)

    'Begining of the tables
    Dim loRowHeader As Long
    Dim loColHeader  As Long
    Dim rowCounter As Long

    'End of the listobject table
    Dim loRowsEnd As Long
    Dim loColsEnd As Long
    Dim Wksh As Worksheet

    Set Wksh = ActiveSheet

    With Wksh
        .Unprotect C_sPassword

        'Rows and columns at the begining of the table to resize
        loRowHeader = Lo.Range.Row
        loColHeader = Lo.Range.Column

        'Rows and Columns at the end of the Table to resize
        loRowsEnd = loRowHeader + Lo.Range.Rows.Count - 1
        loColsEnd = loColHeader + Lo.Range.Columns.Count - 1

        If Not AddRows Then 'Remove rows
            rowCounter = loRowsEnd
            Do While (rowCounter > loRowHeader + 1)
                If (Application.WorksheetFunction.CountA(.Rows(rowCounter)) <= totalRowCount) Then

                    .Rows(rowCounter).EntireRow.Delete

                    'update the end rows
                    loRowsEnd = loRowsEnd - 1
                End If

                rowCounter = rowCounter - 1
            Loop
        Else 'Add rows
            loRowsEnd = loRowsEnd + 1 'Start at the bottom of the table

            For rowCounter = 1 To C_iNbLinesLLData + 1
                .Rows(loRowsEnd).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Next
            loRowsEnd = loRowsEnd + C_iNbLinesLLData
        End If

            Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
    End With

    Call ProtectSheet
End Sub

Sub ProtectSheet()
    ActiveSheet.Protect Password:=C_sPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                        , AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                        AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
                        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                        AllowUsingPivotTables:=True

End Sub

'Add id to a range

Public Sub AddID(rng As Range, Optional sChar As String = "ID")

    'Increment a counter and write the values in each cells (ID_1, ID_2, etc.)
    Dim counter As Long
    Dim c As Range
    counter = 1

    ActiveSheet.Unprotect C_sPassword

    For Each c In rng
        c.Value = sChar & " " & counter
        counter = counter + 1
    Next

    ProtectSheet

End Sub


'Resize the dictionary table object
Public Sub AddRowsDict()
     ResizeLo Lo:=sheetDictionary.ListObjects(C_sTabDictionary)
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

    Dim IdRange As Range

    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabTA)
    Set IdRange = sheetAnalysis.ListObjects(C_sTabTA).ListColumns(1).DataBodyRange

    'Add the IDs using the Series
    AddID IdRange, sChar:=C_sSeries
End Sub

Public Sub AddRowsSA()
    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabSA)
End Sub

'Add row to graphs on time series
Public Sub AddRowsGTS()

    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabGTS)

End Sub

Public Sub AddRowsGTSLab()
    Dim IdRange As Range

    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabGTSLab)
    Set IdRange = sheetAnalysis.ListObjects(C_sTabGTSLab).ListColumns(1).DataBodyRange
    AddID IdRange, sChar:="Graph"

End Sub

'resize the tables (delete empty rows at the bottom)----------------------------

'Resize the dictionary table object
Public Sub RemoveRowsDict()
    BeginWork

        ResizeLo Lo:=sheetDictionary.ListObjects(C_sTabDictionary), AddRows:=False

    EndWork
End Sub

'Resize the choices table object
Public Sub RemoveRowsChoices()
    BeginWork

        ResizeLo Lo:=SheetChoice.ListObjects(C_sTabChoices), AddRows:=False

    EndWork
End Sub

Public Sub RemoveRowsGS()
    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabGS), AddRows:=False
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
    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabTA), AddRows:=False, totalRowCount:=2

End Sub

Public Sub RemoveRowsGTS()
    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabGTS), AddRows:=False, totalRowCount:=4
End Sub

Public Sub RemoveRowsGTSLab()
    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabGTSLab), AddRows:=False, totalRowCount:=1
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

    Case C_sModifyGTS

        Call AddRowsGTS
    
    Case C_sModifyGTSLab
    
        AddRowsGTSLab

    Case Else

        AddRowsGS
        AddRowsUA
        AddRowsBA
        AddRowsTA
        AddRowsGTS
        AddRowsGTSLab
        AddRowsSA

    End Select

    EndWork

End Sub

Public Sub RemoveRowsAna()

    BeginWork

    Select Case Application.WorksheetFunction.Trim(sheetAnalysis.Range("RNG_table_modify").Value)

    'Global Summary
    Case C_sModifyGS

        Call RemoveRowsGS

    'Univariate Analysis
    Case C_sModifyUA

        Call RemoveRowsUA

    'Bivariate Analysis
    Case C_sModifyBA

        Call RemoveRowsBA

    'Time Series analysis
    Case C_sModifyTA

        Call RemoveRowsTA

    'Graph on time series
    Case C_sModifyGTS

        Call RemoveRowsGTS
        
    Case C_sModifyGTSLab
        
        Call RemoveRowsGTSLab

    'Spatial Analysis
    Case C_sModifySA

        Call RemoveRowsSA
    Case Else

        RemoveRowsGS
        RemoveRowsUA
        RemoveRowsBA
        RemoveRowsTA
        RemoveRowsGTS
        RemoveRowsGTSLab
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
    Dim l As Long
    Dim iVarColumn As Integer
    Dim iListVarColumn As Integer
    Dim iTimeVarColumn As Integer
    Dim geoColumn As Long
    Dim rng As Range


    BeginWork

    Set rng = sheetDictionary.ListObjects(C_sTabDictionary).HeaderRowRange

    If rng.Find(C_sDictHeaderControl, lookAt:=xlWhole, MatchCase:=True) Is Nothing Or _
         rng.Find(C_sDictHeaderVarName, lookAt:=xlWhole, MatchCase:=True) Is Nothing Or _
         rng.Find(C_sDictHeaderType, lookAt:=xlWhole, MatchCase:=True) Is Nothing Then
        Exit Sub
    Else
        iVarColumn = rng.Find(C_sDictHeaderVarName, lookAt:=xlWhole, MatchCase:=True).Column
        iControlColumn = rng.Find(C_sDictHeaderControl, lookAt:=xlWhole, MatchCase:=True).Column
        iTypeColumn = rng.Find(C_sDictHeaderType, lookAt:=xlWhole, MatchCase:=True).Column
    End If

    With sheetDictionary
        iDictLength = .Cells(.Rows.Count, iVarColumn).End(xlUp).Row
    End With

    'Delte the ranges of variables and time series
    On Error Resume Next

        sheetLists.ListObjects(C_sTabVarList).DataBodyRange.Delete
        sheetLists.ListObjects(C_sTabTimeVar).DataBodyRange.Delete
         sheetLists.ListObjects(C_sTabGeoVar).DataBodyRange.Delete
         
    On Error GoTo errHand

    iListVarColumn = sheetLists.ListObjects(C_sTabVarList).Range.Column
    iTimeVarColumn = sheetLists.ListObjects(C_sTabTimeVar).Range.Column
    geoColumn = sheetLists.ListObjects(C_sTabGeoVar).Range.Column

    j = 1
    k = 1
    l = 1
    With sheetDictionary
        For i = rng.Row + 1 To iDictLength

            If .Cells(i, iControlColumn).Value = C_sDictControlChoice Or _
                .Cells(i, iControlColumn).Value = C_sDictControlChoiceForm Then
                j = j + 1
                sheetLists.Cells(j, iListVarColumn).Value = .Cells(i, iVarColumn).Value
            End If

            'Add Dates vars
             If .Cells(i, iTypeColumn).Value = C_sDictTypeDate Then
                k = k + 1
                sheetLists.Cells(k, iTimeVarColumn).Value = .Cells(i, iVarColumn).Value
            End If

            'Add Geo vars
             If .Cells(i, iControlColumn).Value = C_sDictControlGeo Or .Cells(i, iControlColumn).Value = C_sDictControlHf Then
                l = l + 1
                sheetLists.Cells(l, geoColumn).Value = .Cells(i, iVarColumn).Value
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
     UpdateValue toValue, "DictMainLabel"
     UpdateValue toValue, "DictSubLabel"
     UpdateValue toValue, "DictNote"
     UpdateValue toValue, "DictSheetName"
     UpdateValue toValue, "DictMainSection"
     UpdateValue toValue, "DictSubSection"
     UpdateValue toValue, "DictFormula"
     UpdateValue toValue, "DictMessage"


    'Record updates for choices
     UpdateValue toValue, "ChoiLabel"

    'Record updateValues for Exports
     UpdateValue toValue, "Exp"

    'Record updateValues for Analysis
     UpdateValue toValue, "AnaGS_SL"
     UpdateValue toValue, "AnaGS_SF"

     UpdateValue toValue, "AnaUA_SC"
     UpdateValue toValue, "AnaUA_SL"
     UpdateValue toValue, "AnaUA_SF"

     UpdateValue toValue, "AnaBA_SC"
     UpdateValue toValue, "AnaBA_SL"
     UpdateValue toValue, "AnaBA_SF"

     UpdateValue toValue, "AnaTA_SC"
     UpdateValue toValue, "AnaTA_SL"
     UpdateValue toValue, "AnaTA_SF"


     UpdateValue toValue, "AnaGTS_GL"
     UpdateValue toValue, "AnaGTS_GT"


End Sub


'Test if there is no update

Function NoUpdate() As Boolean
    NoUpdate = Not ( _
     Updated("DictMainLabel") Or _
     Updated("DictSubLabel") Or _
     Updated("DictNote") Or _
     Updated("DictSheetName") Or _
     Updated("DictMainSection") Or _
     Updated("DictSubSection") Or _
     Updated("DictFormula") Or _
     Updated("DictMessage") Or _
     Updated("ChoiLabel") Or _
     Updated("Exp") Or _
     Updated("AnaGS_SL") Or _
     Updated("AnaGS_SF") Or _
     Updated("AnaUA_SC") Or _
     Updated("AnaUA_SL") Or _
     Updated("AnaUA_SF") Or _
     Updated("AnaBA_SC") Or _
     Updated("AnaBA_SL") Or _
     Updated("AnaBA_SF") Or _
     Updated("AnaTA_SC") Or _
     Updated("AnaTA_SL") Or _
     Updated("AnaTA_SF"))
End Function


'Add options for the graphs (Choices, Percentages, etc.)
'depending on choices on series
Public Sub AddGraphOptions(rng As Range)

    'Values of row, column and serie for the graph Table
    Dim graphRow As Long
    Dim loRow As Long
    Dim graphSerie As String

    'Values of row, column and serie for the Time series table
    Dim tsRow As Long
    Dim tsGroupBy As String
    Dim tsAddPerc As String
    Dim tsAddTotal As String
    Dim tsTimeVar As String

    'Constants for columns on time series table
    Const tsGroupByColumn As Byte = 5
    Const tsAddPercColumn As Byte = 10
    Const tsAddTotalColumn As Byte = 11
    Const tsTimeVarColumn As Byte = 4

    'Contants for columns on graph table
    Const graphPercColumn As Byte = 10
    Const graphChoicesColumn As Byte = 8
    Const graphCol As Byte = 3

    graphRow = rng.Row
    graphSerie = sheetAnalysis.Cells(graphRow, graphCol).Value


    On Error GoTo errHand
    ActiveSheet.Unprotect C_sPassword
    
    'Events are stoped in the BeginWork process
    
    BeginWork
    Application.Cursor = xlIBeam

    With sheetAnalysis
        'remove previous data validation
        .Cells(graphRow, graphPercColumn).validation.Delete
        .Cells(graphRow, graphChoicesColumn).validation.Delete
        .Cells(graphRow, graphPercColumn).Value = ""
        .Cells(graphRow, graphChoicesColumn).Value = ""

        'Corresponding row in the time series table
        Debug.Print graphSerie
        tsRow = CInt(Application.WorksheetFunction.Trim(Replace(graphSerie, C_sSeries, ""))) + _
                .ListObjects(C_sTabTA).Range.Row

        tsGroupBy = .Cells(tsRow, tsGroupByColumn).Value
        tsAddPerc = .Cells(tsRow, tsAddPercColumn).Value
        tsAddTotal = .Cells(tsRow, tsAddTotalColumn).Value
        tsTimeVar = .Cells(tsRow, tsTimeVarColumn).Value
        
    End With

    'Set validation on percentage
    With sheetAnalysis.Cells(graphRow, graphPercColumn)
        If tsAddPerc <> C_sNo Then
            .Locked = False
            .Font.Color = RGB(34, 96, 162)
            .Font.Italic = False
            SetValidation sheetAnalysis.Cells(graphRow, graphPercColumn), sValidList:="=" & "perc_values", sAlertType:=1
        Else
            .Value = "values"
            .Font.Color = RGB(127, 127, 127)
            .Font.Italic = True
            .Locked = True
        End If
    End With


    'Add the choices
    AddChoices tsGroupBy, graphRow, tsAddTotal
    ProtectSheet

    Application.Cursor = xlDefault
    EndWork

    Exit Sub

errHand:
    MsgBox Err.Description
    ProtectSheet

    Application.Cursor = xlDefault
    EndWork
    Exit Sub

End Sub

'Clear names on graph on time series related to dropdowns

Public Sub ClearDropdowns()
    
    Dim counter As Long
    Dim listChoices As String
    Dim listChoiceLo As ListObject
    
    
    'remove the names related to all the choices in the workbook, for future uses.
    
    Set listChoiceLo = sheetLists.ListObjects(C_sTabListOfChoicesNames)
    If listChoiceLo.DataBodyRange Is Nothing Then Exit Sub
    
    For counter = 1 To listChoiceLo.DataBodyRange.Rows.Count
        
        listChoices = sheetLists.Cells(counter + 1, listChoiceLo.DataBodyRange.Column).Value
        
        On Error Resume Next
            ThisWorkbook.NAMES(listChoices).Delete
        On Error GoTo 0
        
    Next
    
    sheetChoicesLists.Cells.Clear
    
    MsgBox "Done!"
End Sub


Public Sub CheckSetup()
End Sub

Public Sub ImportSetup()
    Imports.Show
End Sub
