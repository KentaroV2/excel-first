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
Public Const Undefined As String = "Undefined" '* "undefined" label
Public Const This As String = "This" '* "this Excel Workbook" label
Public Const First_Level_Delimiter As String = vbTab '* First-level delimiter
Public Const Second_Level_Delimiter As String = vbTab & vbTab '* Second-level delimiter
Public Const Field_Date_Delimiter As String = "__" '* Field-Date delimiter (ie. "ChargeAmount__Y2019M04")
Public Const And_Operator As String = "&&&&&" '* And operator
Public Const Or_Operator As String = "|||||" '* Or operator
Public Enum Exit_Status '* Exit status
  Success = (vbObjectError + 512) + 1
  ExcelWorkbook_Is_Not_Found
  Workbook_Is_Not_Found
  Workbook_Is_Already_Existed
  Worksheet_Is_Not_Found
  Worksheet_Is_Already_Existed
  SetOfFields_Is_Not_Defined
End Enum
Public Enum WorksheetRow '* Worksheet row
  Top = 1
End Enum
Public Enum WorksheetColumn '* Worksheet column
  Left__ = 1
End Enum
Public Enum TableRow '* Table row
  Field = 1
  TopOfRecord
End Enum
Public Enum TableColumn '* Table column
  Left__ = 1
End Enum
Public Enum FieldAndValue '* Field and value
  Field
  Value
End Enum
Public Const Left_Parentheses As String = "(" '* Left parentheses
Public Const Right_Parentheses As String = ")" '* Right parentheses
