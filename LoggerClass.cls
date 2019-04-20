VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'! This provides (a) original logging functions in reference to log4j and (b) object names, method names, and property names logging functions with log messages.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------
' Define members.
Private myName As String '* Name. This name is also used for logging purpose.
Private myParent As Object '* Parent object.
Private myStuckNames As String '* A set of names displayed with log message.
Private myLevel As Long '* Log level designated by Logger_Level enumeration.

'* Initialize the class.
Private Sub Class_Initialize()
  myName = Logger_Name
  myLevel = Logger_Level.Off
End Sub

'* Terminate this class.
Private Sub Class_Terminate()
  Set myParent = Nothing
End Sub

'* Set Name property.
Public Property Let Name(Name As String)
  myName = Name
End Property

'* Get Name property.
Public Property Get Name() As String
  Name = myName
End Property

'* Set Parent property.
Public Property Set Parent(Parent As Object)
  Set myParent = Parent
End Property

'* Get Parent property.
Public Property Get Parent() As Object
  Set Parent = myParent
End Property

'* Set logger level. Logger levels are defined by Logger_Level enumeration.
'* @param[in] Level Logger level defined by "Logger_Level" enumeration.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub SetLevel(Level As Long, Optional ExitStatus As Long)
  myLevel = Level
  ExitStatus = Exit_Status.Success
End Sub

'* Get logger level.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return Logger level defined by "Logger_Level" enumeration.
Public Property Get GetLevel(Optional ExitStatus As Long) As Long
  GetLevel = myLevel
  ExitStatus = Exit_Status.Success
End Property

'* Stack a name which logged messages indicate. The last stuck name (a top of pile) is always logged.
'* @param[in] StuckName Stuck name. The name consists of (a) object name and (b) method name or property name. When no name is stuck, "(undefined category)" is logged instead.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub StackName(StuckName As String, Optional ExitStatus As Long)
  If (myStuckNames <> "") Then
    myStuckNames = myStuckNames & First_Level_Delimiter
  End If
  myStuckNames = myStuckNames + StuckName
  Me.Info ("Start.")
  ExitStatus = Exit_Status.Success
End Sub

'* Unstack a name which logged messages indicate. A top of stuck names (a top of pile) is removed.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub UnstackName(Optional ExitStatus As Long)
  If (myStuckNames <> "") Then
    Me.Info ("End.")
    Dim names_ As Variant
    names_ = Split(myStuckNames, First_Level_Delimiter)
    If (UBound(names_) > 0) Then
      ReDim Preserve names_(UBound(names_) - 1)
      myStuckNames = Join(names_, First_Level_Delimiter)
    Else
      myStuckNames = Empty
    End If
  End If
  ExitStatus = Exit_Status.Success
End Sub

'* Log message with name including object name, method name, and property name.
'* @param[in] Level log level used to decide whether name and message should be logged or not with referring myLevel member.
'* @param[in] Message message logged with stuck name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Log(Level As String, Message As String, Optional ExitStatus As Long)
  If (myStuckNames <> "") Then
    Dim names_ As Variant
    names_ = Split(myStuckNames, First_Level_Delimiter)
    Dim category_ As String
    category_ = names_(UBound(names_))
    Dim displayedIndent_ As String
    displayedIndent_ = String(UBound(names_) - LBound(names_) + 1, ">")
  Else
    category_ = Undefined
    displayedIndent_ = ""
  End If
  Debug.Print _
    Format(Now(), "yyyy-mm-dd hh:nn:ss ") & _
    "[" & Level & "] " & _
    displayedIndent_ & " " & _
    category_ & " - " & Message
  ExitStatus = Exit_Status.Success
End Sub

'* Log fatal message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Fatal(Message As String, Optional ExitStatus As Long)
  If (myLevel >= Logger_Level.Fatal) Then
    Call Me.Log("FATAL", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
End Sub

'* Log error message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Error(Message As String, Optional ExitStatus As Long)
  If (myLevel >= Logger_Level.Error__) Then
    Call Me.Log("ERROR", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
End Sub

'* Log warning message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Warn(Message As String, Optional ExitStatus As Long)
  If (myLevel >= Logger_Level.Warn) Then
    Call Me.Log("WARN", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
End Sub

'* Log information message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Info(Message As String, Optional ExitStatus As Long)
  If (myLevel >= Logger_Level.Info) Then
    Call Me.Log("INFO", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
End Sub

'* Log debug message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Debug__(Message As String, Optional ExitStatus As Long)
  If (myLevel >= Logger_Level.Debug__) Then
    Call Me.Log("DEBUG", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
End Sub

'* Log trace message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Trace(Message As String, Optional ExitStatus As Long)
  If (myLevel >= Logger_Level.Trace) Then
    Call Me.Log("TRACE", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
End Sub

'* Release CPU resource periodically to prevent applications from freezing user operations.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub ReleaseCpuResource(Optional ExitStatus As Long)
  Sleep 0
  DoEvents
  ExitStatus = Exit_Status.Success
End Sub
