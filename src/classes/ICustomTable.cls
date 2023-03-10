VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Sub AddRows(ByVal pass As IPasswords): End Sub
Public Sub RemoveRows(ByVal pass As IPasswords, _
                     Optional ByVal totalCount As Long = 0)
End Sub

Public Sub SetValidation(ByVal colName As String, ByVal drop As IDropdownLists, _
                         ByVal dropName As String, ByVal pass As IPasswords, _
                         Optional ByVal alertType As String = "info")
End Sub
