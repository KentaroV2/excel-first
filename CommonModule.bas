Attribute VB_Name = "CommonModule"
Option Explicit
'! This provides common definitions like public constants, public variables, and declaration of windows libraries.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------
' Declare Sleep function to release CPU resources periodically.
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' Define public constants.
Public Const First_Level_Delimiter As String = vbTab '* First-level delimiter
Public Const Exit_Status__Success As Long = 0 '* Exit status of success
Public Const Left_Parentheses As String = "(" '* Left parentheses
Public Const Right_Parentheses As String = ")" '* Right parentheses
