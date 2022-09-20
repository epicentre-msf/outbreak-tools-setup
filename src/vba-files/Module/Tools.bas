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

Sub ResizeLo(Lo As ListObject, Optional AddRows As Boolean = True, Optional iColStart As Integer = -1)

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
        'Or the column given by the user.
        If iColStart > 0 Then
            iFirstColumnRow = ActiveSheet.Cells(iRowHeader, iColStart).End(xlDown).Row
        Else
            iFirstColumnRow = ActiveSheet.Cells(iRowHeader, iColHeader).End(xlDown).Row
        End If

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

'Add id to a range

Public Sub AddID(Rng As Range, Optional sChar As String = "ID")

    'Increment a counter and write the values in each cells (ID_1, ID_2, etc.)
    Dim Counter As Long
    Dim c As Range
    Counter = 1
    
    ActiveSheet.Unprotect C_sPassword
    
    For Each c In Rng
        c.Value = sChar & " " & Counter
        Counter = Counter + 1
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

    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabTA), iColStart:=3
    
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
    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabTA), AddRows:=False, iColStart:=3
End Sub

Public Sub RemoveRowsGTS()
    ResizeLo Lo:=sheetAnalysis.ListObjects(C_sTabGTS), AddRows:=False
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

    Case Else

        AddRowsGS
        AddRowsUA
        AddRowsBA
        AddRowsTA
        AddRowsGTS
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

    'Spatial Analysis
    Case C_sModifySA

        Call RemoveRowsSA
    Case Else

        RemoveRowsGS
        RemoveRowsUA
        RemoveRowsBA
        RemoveRowsTA
        RemoveRowsGTS
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


'Test if there is no update

Function NoUpdate() As Boolean
    NoUpdate = Not ( _
     bUpdateDictVarName Or _
     bUpdateDictMainLabel Or _
     bUpdateDictSubLabel Or _
     bUpdateDictNote Or _
     bUpdateDictSheetName Or _
     bUpdateDictMainSection Or _
     bUpdateDictSubSection Or _
     bUpdateDictFormula Or _
     bUpdateDictMessage Or _
     bUpdateChoiLabel Or _
     bUpdateExp Or _
     bUpdateAnaGS_SL Or _
     bUpdateAnaGS_SF Or _
     bUpdateAnaUA_SC Or _
     bUpdateAnaUA_SL Or _
     bUpdateAnaUA_SF Or _
     bUpdateAnaBA_SC Or _
     bUpdateAnaBA_SL Or _
     bUpdateAnaBA_SF Or _
     bUpdateAnaTA_SC Or _
     bUpdateAnaTA_SL Or _
     bUpdateAnaTA_SF)
End Function


'Add options for the graphs (Choices, Percentages, etc.)
'depending on choices on series
Public Sub AddGraphOptions(Rng As Range)

    'Values of row, column and serie for the graph Table
    Dim graphRow As Long
    Dim graphCol As Integer
    Dim graphSerie As String

    'Values of row, column and serie for the Time series table
    Dim tsRow As Long
    Dim tsGroupBy As String
    Dim tsAddPerc As String
    Dim tsAddTotal As String
    
    'Constants for columns on time series table
    Const tsGroupByColumn As Byte = 5
    Const tsAddPercColumn As Byte = 9
    Const tsAddTotalColumn As Byte = 10
    
    'Contants for columns on graph table
    Const graphPercColumn As Byte = 5
    Const graphChoicesColumn As Byte = 4

    graphRow = Rng.Row
    graphCol = sheetAnalysis.ListObjects(C_sTabGTS).Range.Column
    graphSerie = sheetAnalysis.Cells(graphRow, graphCol).Value
    
    
    If graphSerie = vbNullString Then Exit Sub
    
    
    On Error GoTo errHand
    ActiveSheet.Unprotect C_sPassword

    BeginWork
    Application.Cursor = xlIBeam
    
    With sheetAnalysis
        'remove previous data validation
        .Cells(graphRow, graphPercColumn).Validation.Delete
        .Cells(graphRow, graphChoicesColumn).Validation.Delete
        .Cells(graphRow, graphPercColumn).Value = ""
        .Cells(graphRow, graphChoicesColumn).Value = ""
        
        'Corresponding row in the time series table
        
        tsRow = CInt(Application.WorksheetFunction.Trim(Replace(graphSerie, C_sSeries, ""))) + _
                .ListObjects(C_sTabTA).Range.Row
                
        tsGroupBy = .Cells(tsRow, tsGroupByColumn).Value
        tsAddPerc = .Cells(tsRow, tsAddPercColumn).Value
        tsAddTotal = .Cells(tsRow, tsAddTotalColumn).Value
    End With

    'Set validation on percentage
    With sheetAnalysis.Cells(graphRow, graphPercColumn)
        If tsAddPerc <> C_sNo Then
            .Locked = False
            .Font.Color = vbBlack
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
