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
  Me.Name = Logger_Name
  myLevel = Logger_Level.Off
End Sub

'* Terminate this class.
Private Sub Class_Terminate()
  Set Me.Parent = Nothing
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
'* @param[in] Level. Logger level defined by "Logger_Level" enumeration.
'* @param[out] ExitStatus. (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
'*
Public Function SetLevel(Level As Long, Optional ExitStatus As Long) As LoggerClass
  myLevel = Level
  ExitStatus = Exit_Status.Success
  Set SetLevel = Me
End Function

'* Get logger level.
Public Property Get Level(Optional ExitStatus As Long) As Long
  Level = myLevel
  ExitStatus = Exit_Status.Success
End Property

'* Stack a name which logged messages indicate. The last stuck name (a top of pile) is always logged.
'* @param[in] StuckName Stuck name. The name consists of (a) object name and (b) method name or property name. When no name is stuck, "(undefined category)" is logged instead.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function StackName(StuckName As String, Optional ExitStatus As Long) As LoggerClass
  If (myStuckNames <> "") Then
    myStuckNames = myStuckNames & First_Level_Delimiter
  End If
  myStuckNames = myStuckNames + StuckName
  Me.Info ("Start.")
  ExitStatus = Exit_Status.Success
  Set StackName = Me
End Function

'* Unstack a name which logged messages indicate. A top of stuck names (a top of pile) is removed.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function UnstackName(Optional ExitStatus As Long) As LoggerClass
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
  Set UnstackName = Me
End Function

'* Log message with name including object name, method name, and property name.
'* @param[in] Level log level used to decide whether name and message should be logged or not with referring myLevel member.
'* @param[in] Message message logged with stuck name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function Log(Level As String, Message As String, Optional ExitStatus As Long) As LoggerClass
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
  Set Log = Me
End Function

'* Log fatal message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function Fatal(Message As String, Optional ExitStatus As Long) As LoggerClass
  If (myLevel >= Logger_Level.Fatal) Then
    Call Me.Log("FATAL", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
  Set Fatal = Me
End Function

'* Log error message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function Error(Message As String, Optional ExitStatus As Long) As LoggerClass
  If (myLevel >= Logger_Level.Error__) Then
    Call Me.Log("ERROR", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
  Set Error = Me
End Function

'* Log warning message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function Warn(Message As String, Optional ExitStatus As Long) As LoggerClass
  If (myLevel >= Logger_Level.Warn) Then
    Call Me.Log("WARN", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
  Set Warn = Me
End Function

'* Log information message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function Info(Message As String, Optional ExitStatus As Long) As LoggerClass
  If (myLevel >= Logger_Level.Info) Then
    Call Me.Log("INFO", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
  Set Info = Me
End Function

'* Log debug message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function Debug__(Message As String, Optional ExitStatus As Long) As LoggerClass
  If (myLevel >= Logger_Level.Debug__) Then
    Call Me.Log("DEBUG", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
  Set Debug__ = Me
End Function

'* Log trace message.
'* @param[in] Message message logged with name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function Trace(Message As String, Optional ExitStatus As Long) As LoggerClass
  If (myLevel >= Logger_Level.Trace) Then
    Call Me.Log("TRACE", Message)
  End If
  Me.ReleaseCpuResource
  ExitStatus = Exit_Status.Success
  Set Trace = Me
End Function

'* Release CPU resource periodically to prevent applications from freezing user operations.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return LoggerClass. Self.
Public Function ReleaseCpuResource(Optional ExitStatus As Long) As LoggerClass
  Sleep 0
  DoEvents
  ExitStatus = Exit_Status.Success
  Set ReleaseCpuResource = Me
End Function
