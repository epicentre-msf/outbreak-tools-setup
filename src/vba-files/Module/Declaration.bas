Attribute VB_Name = "Declaration"
Option Explicit

'constant linked to the different columns to be translated in the workbook sheets
Public Const sCstColDictionary As String = "Main label|Sub-label|Note|Sheet Name|Main section|Sub-section|Formula|Message"
Public Const sCstColChoices As String = "label_short|label"
Public Const sCstColGlobalSummary As String = "Summary label|Summary function"
Public Const sCstColExport As String = "Label button"
Public Const C_iNbLinesLLData As Integer = 5
Public Const C_iStartLinesTrans As Integer = 5

'determines whether to update the Translation sheet
Public bUpdate As Boolean

Public Reponse As Byte


'Different string constants
Public Const C_sPassword        As String = "1234"   'Password
Public Const C_sTabDictionary   As String = "Tab_Dictionary"
Public Const C_sTabChoices      As String = "Tab_Choices"
Public Const C_sTabExports      As String = "Tab_Export"
Public Const C_sTabGS           As String = "Tab_Global_Summary"
Public Const C_sTabUA           As String = "Tab_Univariate_Analysis"
Public Const C_sTabBA           As String = "Tab_Bivariate_Analysis"

Public Const C_sModifyGS        As String = "Add or remove rows of Global Summary"
Public Const C_sModifyUA        As String = "Add or remove rows of Univariate Analysis"
Public Const C_sModifyBA        As String = "Add or remove rows of Bivariate Analysis"

'Public Const C_sTab   As String = "Tab_"




