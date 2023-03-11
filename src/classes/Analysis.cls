VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Analysis")
'@ModuleDescription("Analysis Class")
'@IgnoreModule

Implements IAnalysis

Private Type TAnalysis
    anash As Worksheet
End Type

Private this As TAnalysis

Private Const CLASSNAME As String = "Analysis"
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"

Option Explicit

Public Function Create(ByVal anash As Worksheet) As IAnalysis
  If anash Is Nothing Then ThrowError "Unable to create instance of class, Worksheet object not defined"
  CheckTablesRequirements
  With New Analysis
      Set .Wksh = anash
      Set Create = .Self
  End With
End Function


Public Property Get Wksh() As Worksheet
  Set Wksh = this.anash
End Property

Public Property Set Wksh(ByVal anash As Worksheet)
  Set this.anash = anash
End Property

Public Property Get Self() As IAnalysis
    Set Self = Me
End Property



'Add or remove Rows on worksheet analysis
Private Sub ManageRows(Optional ByVal del As Boolean = False)
  Dim sh As Worksheet
  Dim Lo As ListObject
  Dim tabl As ICustomTable
  Dim listLoNames As BetterArray
  Dim loName As String
  Dim counter As Long
  Dim selectedTable As String

  Set sh = Wksh()
  Set listLoNames = New BetterArray
  selectedTable = Replace(LCase(sh.Cells(1, 1).Value),  _
                          "add or remove rows of", "")
  selectedTable = Application.WorksheetFunction.Trim(selectedTable)

  Select Case selectedTable
  Case "global summary"
    listLoNames.Push GLOBALSUMMARY
  Case "univariate analysis"
    listLoNames.Push UNIVARIATE
  Case "bivariate analysis"
    listLoNames.Push BIVARIATE
  Case "time series analysis"
    listLoNames.Push TIMESERIES
  Case "graph on time series labels"
    listLoNames.Push TITLEGRAPHTIMESERIES
  Case "graph on time series"
    listLoNames.Push GRAPHTIMESERIES
  Case "spatial analysis"
    listLoNames.Push SPATIAL
  Case "spatio-temporal analysis"
    listLoNames.Push SPATIOTEMPORAL
  Case Else
    listLoNames.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, _
                     TITLEGRAPHTIMESERIES, GRAPHTIMESERIES, SPATIAL, _
                     SPATIOTEMPORAL
  End Select

  For counter = listLoNames.LowerBound To listLoNames.UpperBound
    loName = listLoNames.Item(counter)
    On Error Resume Next
      Set Lo = sh.ListObjects(loName)
      If  Lo Is Nothing Then Exit Sub
    On Error GoTo 0
    Set tabl = CustomTable.Create(Lo)
    If del Then
        tabl.RemoveRows
    Else
        tabl.AddRows
    End If
  Next
End Sub

'Errors
Private Sub CheckTablesRequirements()
  Dim loTable As BetterArray
  Dim counter As Long
  Dim Lo As ListObject
  Dim LoName As String

  Set loTable = New BetterArray

  loTable.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, _
               GRAPHTIMESERIES, TITLEGRAPHTIMESERIES, SPATIAL, SPATIOTEMPORAL

  For counter = loTable.LowerBound To loTable.UpperBound
      LoName = loTable.Items(counter)
      On Error Resume Next
      Set Lo = Wksh.ListObjects(LoName)
      If Lo Is Nothing Then _
       ThrowError " Listobject " & LoName _
        & " does not exist in analysis worksheet "
      On Error GoTo 0
  Next
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 515, CLASSNAME, errorMessage
End Sub

'Exposed methods
Public Sub IAnalysis_AddRows(ByVal pass As IPasswords)
  ManageRows pass:=pass
End Sub

Public Sub IAnalysis_RemoveRows(ByVal pass As IPasswords)
  ManageRows pass:=pass, del:=True
End Sub
