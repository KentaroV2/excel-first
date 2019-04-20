Attribute VB_Name = "CommonModule"
Option Explicit
'! This provides common definitions like public constants, public variables, and declaration of windows libraries.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
Public Const Font_Name As String = "Meiryo UI" '* Font name
Public Const Font_Size As Double = 10 '* Font name
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
Public Enum Exit_Status '* Exit status
  Success = (vbObjectError + 512) + 1
  ExcelWorkbook_Is_Not_Found
  Workbook_Is_Not_Found
  Workbook_Is_Already_Existed
  Worksheet_Is_Not_Found
  Worksheet_Is_Already_Existed
End Enum
Public Const Left_Parentheses As String = "(" '* Left parentheses
Public Const Right_Parentheses As String = ")" '* Right parentheses
