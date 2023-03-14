VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IUpdatedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Setup")
'@ModuleDescription("Interface of Updating Class")
'@IgnoreModule

Option Explicit

'Exposed methods

Public Sub AddColumn(ByVal lo As ListObject, colName As String, idName As String): End Sub
Public Property Get UpdateStatus(ByVal colName As String) As Boolean : End Property
Public Sub CheckUpate(ByVal Target As Range):End Sub