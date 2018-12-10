VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'! @file LoggerClass.cls
'! @brief Provides logging functions in reference to log4j.
'! @copyright MIT
Option Explicit
' --------------------------------------------------------------------------------------------------------------
' Edit the followings as needed:

' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
Private Const DELIMITER As String = vbTab '* Delimiter
Private MyFunctionNames As String '* A set of function names displayed with log
Private MyLevel As Long '* Log level designated by LoggerLevel enumeration

Private Sub Class_Initialize()
  MyFunctionNames = ""
  MyLevel = loggerLevel.OFF_LEVEL
End Sub

'* @fn Function SetLevel(level As Long) As Long
'* @brief Set logger level. Logger levels are defined by "loggerLevel" enumeration.
'* @param[in] level logger level defined by "loggerLevel" enumeration
'* @return exit status. Zero means successful exit.
Public Function SetLevel(level As Long) As Long
  MyLevel = level
  SetLevel = 0
End Function

Public Property Get GetLevel()
  GetLevel = MyLevel
End Property

Public Function StackFunctionName(functionName As String) As Long
  If (MyFunctionNames <> "") Then
    MyFunctionNames = MyFunctionNames & DELIMITER
  End If
  MyFunctionNames = MyFunctionNames + functionName
  StackFunctionName = 0
End Function

Public Function UnstackFunctionName() As Long
  If (MyFunctionNames <> "") Then
    Dim functionNames As Variant
    functionNames = Split(MyFunctionNames, DELIMITER)
    If (UBound(functionNames) > 0) Then
      ReDim Preserve functionNames(UBound(functionNames) - 1)
      MyFunctionNames = Join(functionNames, DELIMITER)
    Else
      MyFunctionNames = Empty
    End If
  End If
  UnstackFunctionName = 0
End Function

Private Function GetDateAndTime() As Long
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss")
  GetDateAndTime = 0
End Function

Private Function Log(level As String, message As String) As Long
  If (MyFunctionNames <> "") Then
    Dim functionNames As Variant
    functionNames = Split(MyFunctionNames, DELIMITER)
    Dim category As String
    category = functionNames(UBound(functionNames))
    Dim displayedIndent As String
    displayedIndent = String(UBound(functionNames) - LBound(functionNames) + 1, ">")
  Else
    category = "(undefined category)"
    displayedIndent = ""
  End If
  Debug.Print _
    Format(Now(), "yyyy-mm-dd hh:nn:ss ") & _
    "[" & level & "] " & _
    displayedIndent & " " & _
    category & " - " & message
    Log = 0
End Function

Public Function Fatal(message As String) As Long
  If (MyLevel >= loggerLevel.FATAL_LEVEL) Then
    Dim exitStatus As Long
    exitStatus = Me.Log("FATAL", message)
  End If
  Fatal = 0
End Function

Public Function Error(message As String) As Long
  If (MyLevel >= loggerLevel.ERROR_LEVEL) Then
    Dim exitStatus As Long
    exitStatus = Me.Log("ERROR", message)
  End If
  Error = 0
End Function

Public Function Warn(message As String) As Long
  If (MyLevel >= loggerLevel.WARN_LEVEL) Then
    Dim exitStatus As Long
    exitStatus = Me.Log("WARN", message)
  End If
  Warn = 0
End Function

Public Function Info(message As String) As Long
  If (MyLevel >= loggerLevel.INFO_LEVEL) Then
    Dim exitStatus As Long
    exitStatus = Me.Log("INFO", message)
  End If
  Info = 0
End Function

Public Function Debug_(message As String) As Long
  If (MyLevel >= loggerLevel.DEBUG_LEVEL) Then
    Dim exitStatus As Long
    exitStatus = Me.Log("DEBUG", message)
  End If
  Debug_ = 0
End Function

Public Function Trace(message As String) As Long
  If (MyLevel >= loggerLevel.TRACE_LEVEL) Then
    Dim exitStatus As Long
    exitStatus = Me.Log("TRACE", message)
  End If
  Trace = 0
End Function
