Attribute VB_Name = "Declaration"
Option Explicit

'constant linked to the different columns to be translated in the workbook sheets
Public Const C_iNbLinesLLData As Integer = 5
Public Const C_iStartLinesTrans As Integer = 5

'determines whether to update the Translation sheet or not. Checkings are done
' at every sheet level, on a worksheet_event call

'Record updates for dictionary

'Number of times we have run translations


'Different string constants
Public Const C_sPassword                    As String = "1234"   'Password
Public Const C_sTabDictionary               As String = "Tab_Dictionary"
Public Const C_sTabChoices                  As String = "Tab_Choices"
Public Const C_sTabExports                  As String = "Tab_Export"
Public Const C_sTabGS                       As String = "Tab_Global_Summary"
Public Const C_sTabUA                       As String = "Tab_Univariate_Analysis"
Public Const C_sTabBA                       As String = "Tab_Bivariate_Analysis"
Public Const C_sTabGTSLab                   As String = "Tab_Label_TSGraph"
Public Const C_sTabVarList                  As String = "var_list_table"
Public Const C_sTabTimeVar                  As String = "dates_var_table"
Public Const C_sTabTranslations             As String = "Tab_Translations"
Public Const C_sTabTA                       As String = "Tab_TimeSeries_Analysis"
Public Const C_sTabSA                       As String = "Tab_Spatial_Analysis"
Public Const C_sTabGTS                      As String = "Tab_Graph_TimeSeries"
Public Const C_sSeries                      As String = "Series"
Public Const C_sNo                          As String = "no"
Public Const C_sTotal                       As String = "Total"
Public Const C_sTabListOfChoicesNames       As String = "list_of_choices_names"
Public Const C_sTabGeoVar                   As String = "geo_var_table"

'Analysis dropdowns
Public Const C_sModifyGS                    As String = "Add or remove rows of Global Summary"
Public Const C_sModifyUA                    As String = "Add or remove rows of Univariate Analysis"
Public Const C_sModifyBA                    As String = "Add or remove rows of Bivariate Analysis"
Public Const C_sModifySA                    As String = "Add or remove rows of Spatial Analysis"
Public Const C_sModifyTA                    As String = "Add or remove rows of Time Series Analysis"
Public Const C_sModifyGTS                   As String = "Add or remove rows to Graph on Time Series"
Public Const C_sModifyGTSLab                As String = "Add or remove rows to Graph on Time Series Labels"


'Different types of controls
Public Const C_sDictControlChoice           As String = "choice_manual"
Public Const C_sDictControlChoiceForm       As String = "choice_formula"
Public Const C_sDictControlCaseWhen         As String = "case_when"
Public Const C_sDictControlGeo              As String = "geo"
Public Const C_sDictControlHf               As String = "hf"

'some types
Public Const C_sDictTypeDate                As String = "date"

'Some headers of the dictionary
Public Const C_sDictHeaderVarName           As String = "Variable Name"
Public Const C_sDictHeaderControl           As String = "Control"
Public Const C_sDictHeaderMainLabel         As String = "Main Label"
Public Const C_sDictHeaderSubLabel          As String = "Sub Label"
Public Const C_sDictHeaderNote              As String = "Note"
Public Const C_sDictHeaderSheetName         As String = "Sheet Name"
Public Const C_sDictHeaderMainSection       As String = "Main Section"
Public Const C_sDictHeaderSubSection        As String = "Sub Section"
Public Const C_sDictHeaderFormula           As String = "Control Details"
Public Const C_sDictHeaderMessage           As String = "Message"
Public Const C_sDictHeaderType              As String = "Type"


'Some headers of choice
Public Const C_sChoHeaderLabelShort         As String = "Label Short"
Public Const C_sChoHeaderLabel              As String = "Label"


'Some headers of export
Public Const C_sExportHeaderLabelButton     As String = "Label Button"

'Some headers of analysis
Public Const C_sAnaHeaderSF                 As String = "Summary Function"
Public Const C_sAnaHeaderSL                 As String = "Summary Label"
Public Const C_sAnaHeaderSC                 As String = "Section"





