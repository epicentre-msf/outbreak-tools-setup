VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Dictionary Class"
'@Folder("Dictionary")
'@ModuleDescription("Dictionary class for the analysis")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Property Get Data() As IDataSheet: End Property
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False) As Range: End Property

Public Property Get Prepared() As Boolean: End Property


'Functions
Public Function ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean) As Boolean: End Function

Public Function UniqueValues(ByVal colName As String) As BetterArray: End Function
Public Function VariableExists(ByVal varName As String) As Boolean: End Function

Public Sub Clean(): End Sub

Public Sub Import(ByVal fromWksh As Worksheet, ByVal pass As IPasswords)
End Sub

Public Sub Export(ByVal toWkb As Workbook, Optional ByVal exportType As String = "__all__")
End Sub

'Add a column
Public Sub AddColumn(ByVal colName As String)
End Sub

'Add and remove rows
Public Sub AddRows(ByVal pass As IPasswords)
End Sub

Public Sub RemoveRows(ByVal pass As IPasswords)
End Sub