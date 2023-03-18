VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISetupImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Exposed methods

Public Sub Check(ByVal impDict As Boolean, _
  ByVal impChoi As Boolean, _
  ByVal impExp As Boolean, _
  ByVal impAna As Boolean, _
  ByVal impTrans As Boolean, _
  Optional ByVal cleanSetup As Boolean = False): End Sub

Public Sub Import(ByVal pass As IPasswords, ByVal sheetsList As BetterArray): End Sub
Public Sub Clean(ByVal pass As IPasswords, ByVal sheetsList As BetterArray): End Sub
