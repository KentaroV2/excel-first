Attribute VB_Name = "CommonModule"
Option Explicit
'! This provides common definitions like public constants, public variables, and declaration of windows libraries.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
Public Const Font_Name As String = "Meiryo UI" '* Font name
Public Const Font_Size As Double = 10 '* Font size
Public Const Font_Color As Long = 0 '* Font color. RGB(0, 0, 0)
Public Const Row_Height As Double = 20 '* Row height
Public Const Column_Width As Double = 1.8 '* Column width
Public Const Horizontal_Alignment As Long = xlGeneral '* Horizontal alignment
Public Const Vertical_Alignment As Long = xlCenter '* Vertical alignment
'
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------
' Declare Sleep function to release CPU resources periodically.
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' Define public constants.
Public Const Excel_First_Name As String = "ExcelFirst" '* "Excel First" label
Public Const Logger_Name As String = "Logger" '* "Logger" label
Public Const Database_Connector_Name As String = "DatabaseConnector" '* "DatabaseConnector" label
Public Const Undefined As String = "Undefined" '* "undefined" label. This label is filled in empty variants or record values to ensure completeness.
Public Const True__ As String = "True" '* "True" label. This label as string type prevents Excel cells from interpreting boolean value as boolean type unintentoinally.
Public Const False__ As String = "False" '* "False" label. This label as string type prevents Excel cells from interpreting boolean value as boolean type unintentoinally.
Public Const Null__ As String = "Null" '* "Null" label. This label is filled in empty variants or record values to ensure completeness.
Public Const This As String = "This" '* "This" label. This label is used for representing "this" Workbook mainly.
Public Const Dot As String = "." '* "Dot" label. This label is used for defining delimiter between file name and file attribute mainly.
Public Const First_Level_Delimiter As String = vbTab '* First-level delimiter.
Public Const Second_Level_Delimiter As String = vbTab & vbTab '* Second-level delimiter.
Public Const Field_Date_Delimiter As String = "__" '* Field-Date delimiter. (ie. "ChargeAmount__Y2019M04")
Public Const And_Operator As String = "&&&&&" '* And operator (five ampersand marks in a row).
Public Const Or_Operator As String = "|||||" '* Or operator (five pipeline marks in a row).
Public Const Path_Separator As String = "\" '* Path separator.
Public Const Excel_2010_2007_File_Attribute As String = "xlsx" '* "Excel 2010 and 2007" file attribute.
Public Const Excel_97_2003_File_Attribute As String = "xls" '* "Excel 97 - 2003" file attribute.
Public Enum Logger_Level '* Logger level.
  Off '* Log nothing.
  Fatal
  Error__
  Warn
  Info '* Log start and end of functions.
  Debug__
  Trace
  All '* Log everything.
End Enum
Public Enum Size_Dimentions '* Size dimentions
  Row = 1
  Column
End Enum
Public Enum Exit_Status '* Exit status
  Success = (vbObjectError + 512) + 1
  ExcelWorkbook_Is_Not_Found
  ExcelWorkbook_Cannot_Be_Unbinded_Due_To_This_Workbook
  ExcelWorksheet_Is_Not_Found
  Workbook_Is_Not_Defined
  Workbook_Is_Not_Found
  Workbook_Is_Already_Existed
  Worksheet_Is_Not_Found
  Worksheet_Is_Not_Correct
  Worksheet_Is_Already_Existed
  SetOfFields_Is_Not_Defined
  DegreeDetails_Is_Not_Valid
  Database_Connection_Type_Is_Not_Valid
  Parameters_For_Database_Connection_Is_Not_Defined
  Miscellaneous
End Enum
Public Enum Worksheet_Row '* Worksheet row
  Top = 1
End Enum
Public Enum Worksheet_Column '* Worksheet column
  Left__ = 1
End Enum
Public Enum Table_Row '* Table row
  Field = 1
  TopOfRecord
End Enum
Public Enum Table_Column '* Table column
  Left__ = 1
End Enum
Public Enum Field_And_Value '* Field and value
  Field
  Value
End Enum
Public Enum Database_Connection_Type '* Database connection type.
  None
  MicrosoftExcelWorksheet
  MicrosoftAccess
  Oracle
  Miscellaneous
End Enum
Public Enum Cursor_Type '* Cursor type for Recordset object.
  adOpenForwardOnly = 0
  adOpenKeyset
  adOpenDynamic
  adOpenStatic
End Enum
Public Enum Lock_Type '* Lock type for Recordset object.
  adLockReadOnly = 1
  adLockPessimistic
  adLockOptimistic
  adLockBatchOptimistic
End Enum
Public Enum Command_Type '* Command type for Recordset object.
  adCmdText = 1
  adCmdTable = 2
  adCmdStoredProc = 4
  adCmdUnknown = 8
End Enum
Public Const Left_Parentheses As String = "(" '* Left parentheses.
Public Const Right_Parentheses As String = ")" '* Right parentheses.
