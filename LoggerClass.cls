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
'! This class provides logging functions in reference to log4j.
'! @copyright MIT
Option Explicit
' --------------------------------------------------------------------------------------------------------------
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
Private Const first_Level_Delimiter As String = vbTab '* First-level delimiter
Private Const exit_Status__Success As Long = 0 '* Exit status of success
' Define members.
Private myNames As String '* A set of names displayed with log
Private myLevel As Long '* Log level designated by LoggerLevel enumeration

'* Initialize class.
Private Sub Class_Initialize()
  myNames = ""
  myLevel = Logger_Level.Off
End Sub

'* Set logger level. Logger levels are defined by Logger_Level enumeration.
'* @param[in] Level_ logger Logger level defined by Logger_Level enumeration
'* @return exit status. Zero means success.
Public Function SetLevel(Level As Long) As Long
  myLevel = Level
  SetLevel = exit_Status__Success
End Function

'* Get logger level.
Public Property Get GetLevel()
  GetLevel = myLevel
End Property

Public Function StackName(Name As String) As Long
  StackName = exit_Status__Success
  If (myNames <> "") Then
    myNames = myNames & first_Level_Delimiter
  End If
  myNames = myNames + Name
  Dim exitStatus_ As Long
  exitStatus_ = Me.Info("Start.")
End Function

Public Function UnstackName() As Long
  If (myNames <> "") Then
    Dim exitStatus_ As Long
    exitStatus_ = Me.Info("End.")
    Dim names_ As Variant
    names_ = Split(myNames, first_Level_Delimiter)
    If (UBound(names_) > 0) Then
      ReDim Preserve names_(UBound(names_) - 1)
      myNames = Join(names_, first_Level_Delimiter)
    Else
      myNames = Empty
    End If
  End If
  UnstackName = exit_Status__Success
End Function

Private Function GetDateAndTime() As Long
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss")
  GetDateAndTime = exit_Status__Success
End Function

Public Function Log(Level As String, Message As String) As Long
  If (myNames <> "") Then
    Dim names_ As Variant
    names_ = Split(myNames, first_Level_Delimiter)
    Dim category_ As String
    category_ = names_(UBound(names_))
    Dim displayedIndent_ As String
    displayedIndent_ = String(UBound(names_) - LBound(names_) + 1, ">")
  Else
    category_ = "(undefined category)"
    displayedIndent_ = ""
  End If
  Debug.Print _
    Format(Now(), "yyyy-mm-dd hh:nn:ss ") & _
    "[" & Level & "] " & _
    displayedIndent_ & " " & _
    category_ & " - " & Message
  Log = exit_Status__Success
End Function

Public Function Fatal(Message As String) As Long
  If (myLevel >= Logger_Level.Fatal) Then
    Dim exitStatus_ As Long
    exitStatus_ = Me.Log("FATAL", Message)
  End If
  Me.ReleaseCpuResource
  Fatal = exit_Status__Success
End Function

Public Function Error(Message As String) As Long
  If (myLevel >= Logger_Level.Error__) Then
    Dim exitStatus_ As Long
    exitStatus_ = Me.Log("ERROR", Message)
  End If
  Me.ReleaseCpuResource
  Error = exit_Status__Success
End Function

Public Function Warn(Message As String) As Long
  If (myLevel >= Logger_Level.Warn) Then
    Dim exitStatus_ As Long
    exitStatus_ = Me.Log("WARN", Message)
  End If
  Me.ReleaseCpuResource
  Warn = exit_Status__Success
End Function

Public Function Info(Message As String) As Long
  If (myLevel >= Logger_Level.Info) Then
    Dim exitStatus_ As Long
    exitStatus_ = Me.Log("INFO", Message)
  End If
  Me.ReleaseCpuResource
  Info = exit_Status__Success
End Function

Public Function Debug__(Message As String) As Long
  If (myLevel >= Logger_Level.Debug__) Then
    Dim exitStatus_ As Long
    exitStatus_ = Me.Log("DEBUG", Message)
  End If
  Me.ReleaseCpuResource
  Debug__ = exit_Status__Success
End Function

Public Function Trace(Message As String) As Long
  If (myLevel >= Logger_Level.Trace) Then
    Dim exitStatus_ As Long
    exitStatus_ = Me.Log("TRACE", Message)
  End If
  Me.ReleaseCpuResource
  Trace = exit_Status__Success
End Function

Public Function ReleaseCpuResource() As Long
  Sleep 0
  DoEvents
  ReleaseCpuResource = exit_Status__Success
End Function
