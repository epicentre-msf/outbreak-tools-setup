VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Checking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Checkings")
'@ModuleDescription("Checking of the setup")
'@IgnoreModule

Option Explicit
Implements IChecking

Private Const className As String = "Checking"

Private Type TChecking
  title As String
  subTitle As String
  keysTable As BetterArray
  labelsTable As BetterArray
  scopesTable As BetterArray
End Type

Private This As TChecking

'Exposed methods
Public Function Create(ByVal titleName As String, Optional ByVal subtitleName As String = vbNullString) As IChecking
  If titleName = vbNullString Then ThrowError "Checkings should have a title"
  With New Checking
    .CheckingTitle = titleName
    .CheckingSubTitle = subtitleName
    Set Create = .Self()
  End With
End Function


Public Property Get Self() As IChecking
  Set Self = Me
End Property

'title and subTitle
Private Property Get CheckingTitle() As String
  CheckingTitle = This.title
End Property

Private Property Get CheckingSubTitle() As String
  CheckingSubTitle = This.subTitle
End Property

Private Property Let CheckingTitle(ByVal titleName As String)
  This.title = titleName
End Property

Private Property Let CheckingSubTitle(ByVal subtitleName As String)
  This.subTitle = subtitleName
End Property

'Heading Property
Private Property Get Heading(Optional ByVal subTitle As Boolean) As String
  If subTitle Then
    Heading = CheckingSubTitle()
    Exit Property
  End If
  Heading = CheckingTitle()
End Property


'Get the list of keys
Private Property Get ListOfKeys() As BetterArray

  'If the keysTable is unitialised, initialize it
  If This.keysTable Is Nothing Then
    Set This.keysTable = New BetterArray
    This.keysTable.LowerBound = 1
  End If
  Set ListOfKeys = This.keysTable.Clone()
End Property

'Test if a key already exists
Private Function keyExists(ByVal keyName As String) As Boolean
  Dim keyTabl As BetterArray

  If (keyName = vbNullString) Then
    keyExists = False
    Exit Function
  End If

  Set keyTabl = ListOfKeys()
  keyExists = keyTabl.Includes(keyName)
End Function

'Add new values to the checking

Private Sub Add(ByVal keyName As String, ByVal LABEL As String, _
                Optional ByVal scope As CheckingScope = checkingNote)

  If keyExists(keyName) Then ThrowError keyName & " already exists, keys should be unique."

  'Initialize the tables
  If (This.labelsTable Is Nothing) And (This.scopesTable Is Nothing) Then
    Set This.labelsTable = New BetterArray
    Set This.scopesTable = New BetterArray
    Set This.keysTable = New BetterArray

    This.labelsTable.LowerBound = 1
    This.scopesTable.LowerBound = 1
    This.keysTable.LowerBound = 1
  End If

  'Push the three values
  This.keysTable.Push keyName
  This.scopesTable.Push scope
  This.labelsTable.Push LABEL
End Sub


Private Property Get ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As CheckingValue = 1) As String
  Dim indexValue As Long
  Dim returnedValue As String
  Dim actualScope As CheckingScope

  If (Not keyExists(keyName)) Then ThrowError keyName & " does not exists in this checking"

  indexValue = This.keysTable.IndexOf(keyName)
  actualScope = This.scopesTable.IndexOf(keyName)

  Select Case retrievedValue

    Case checkingLabel

      returnedValue = This.labelsTable.Item(indexValue)

    Case checkingType

      Select Case actualScope

        Case checkingNote
          returnedValue = "Note"
        Case checkingError
          returnedValue = "Error"
        Case checkingWarning
          returnedValue = "Warning"
        Case checkingInfo
          returnedValue = "Info"
      End Select

    Case checkingColor

      Select Case actualScope

        Case checkingNote
          returnedValue = "blue"
        Case checkingError
          returnedValue = "red"
        Case checkingWarning
          returnedValue = "orange"
        Case checkingInfo
          returnedValue = "gray"
      End Select
  End Select

End Property


'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, className, "Error : " & errorMessage
End Sub


'Exposed methods
Public Sub IChecking_Add(ByVal keyName As String, ByVal LABEL As String, _
                Optional ByVal scope As CheckingScope = checkingNote)
End Sub
Public Property Get IChecking_ListOfKeys() As BetterArray
  Set IChecking_ListOfKeys = ListOfKeys()
End Property

'Heading could be subtitle or title.
Public Property Get IChecking_Heading(Optional ByVal subTitle As Boolean = False) As String
  IChecking_Heading = Heading(subTitle)
End Property

'You can retrieve the label, the scope or the color
Public Property Get IChecking_ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As CheckingValue = 1) As String
  IChecking_ValueOf = ValueOf(keyName, retriedValue)
End Property
