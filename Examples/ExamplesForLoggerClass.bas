Attribute VB_Name = "ExamplesForLoggerClass"
Option Explicit
'! This module provides some examples that help you understanding how to use the LoggerClass.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------

'* This subroutine teaches how to use the LoggerClass.
Sub ExampleForLoggerClass_UseLogger()
  
  ' Instantiate logger class.
  Dim logger_ As LoggerClass
  Set logger_ = New LoggerClass
  
  ' Set ERROR as logger level.
  Call logger_.SetLevel(Logger_Level.Error__)
  
  ' Write ERROR log.
  Call logger_.Error("This is an ERROR message.")  ' "yyyy-mm-dd hh:mm:ss [ERROR]  Undefined - This is an ERROR message."
  
End Sub

'* This subrouutine teaches how to write log under defined logger level.
Sub ExampleForLoggerClass_WriteLogs()
  
  ' Instantiate logger class.
  Dim logger_ As LoggerClass
  Set logger_ = New LoggerClass
  
  ' Use "With" statement for efficient programming.
  With logger_
    ' Set INFO as logger level.
    .SetLevel (Logger_Level.Info)
    ' Write logs under INFO level.
    .Fatal ("This is a FATAL message.") ' "yyyy-mm-dd hh:mm:ss [FATAL]  Undefined - This is a FATAL message."
    .Error ("This is an ERROR message.") ' "yyyy-mm-dd hh:mm:ss [ERROR]  Undefined - This is an ERROR message."
    .Warn ("This is a WARN message.") ' "yyyy-mm-dd hh:mm:ss [WARN]  Undefined - This is a WARN message."
    .Info ("This is an INFO message.") ' "yyyy-mm-dd hh:mm:ss [INFO]  Undefined - This is an INFO message."
    .Debug__ ("This is an DEBUG message.") ' Not logged.
    .Trace ("This is an TRACE message.") ' Not logged.
  End With
  
End Sub

'* This subrouutine teaches how to write logs with names.
Sub ExampleForLoggerClass_WriteLogsWithNames()
  
  ' Instantiate logger class.
  Dim logger_ As LoggerClass
  Set logger_ = New LoggerClass
  
  ' Use method chain for efficient programming.
  Call logger_ _
    .SetLevel(Logger_Level.Info) _
    .Info("This is an INFO message.") _
    .StackName("foo") _
    .Info("This is an INFO message.") _
    .StackName("bar") _
    .Info("This is an INFO message.") _
    .UnstackName _
    .Info("This is an INFO message.") _
    .UnstackName _
    .Info("This is an INFO message.")
    
End Sub
