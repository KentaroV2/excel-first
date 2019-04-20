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
  Dim Logger As LoggerClass
  Set Logger = New LoggerClass
  
  ' Set ERROR as logger level.
  Dim exitStatus_ As Long
  ' Call Logger.SetLevel(Logger_Level.Error__)
  Call Logger.SetLevel(Logger_Level.Error__, exitStatus_)
  Call Logger.Error(CStr(exitStatus_))
  
  ' Write ERROR log.
  Logger.Error ("This is an ERROR message.") ' "yyyy-mm-dd hh:mm:ss [ERROR]  (undefined) - This is an ERROR message."
  
End Sub

'* This subrouutine teaches how to write log under defined logger level.
Sub ExampleForLoggerClass_WriteLogs()
  
  ' Instantiate logger class.
  Dim Logger As LoggerClass
  Set Logger = New LoggerClass
  
  ' Use "With" statement to program in efficient way.
  With Logger
    ' Set INFO as logger level.
    .SetLevel (Logger_Level.Info)
    ' Write logs under INFO level.
    .Fatal ("This is a FATAL message.") ' "yyyy-mm-dd hh:mm:ss [FATAL]  (undefined) - This is a FATAL message."
    .Error ("This is an ERROR message.") ' "yyyy-mm-dd hh:mm:ss [ERROR]  (undefined) - This is an ERROR message."
     .Warn ("This is a WARN message.") ' "yyyy-mm-dd hh:mm:ss [WARN]  (undefined) - This is a WARN message."
    .Info ("This is an INFO message.") ' "yyyy-mm-dd hh:mm:ss [INFO]  (undefined) - This is an INFO message."
    .Debug__ ("This is an DEBUG message.") ' Not logged.
    .Trace ("This is an TRACE message.") ' Not logged.
  End With
  
End Sub

'* This subrouutine teaches how to write logs with names.
Sub ExampleForLoggerClass_WriteLogsWithNames()
  
  ' Instantiate logger class.
  Dim Logger As LoggerClass
  Set Logger = New LoggerClass
  
  ' Use "With" statement to program in efficient way.
  With Logger
    ' Set INFO as logger level.
    .SetLevel (Logger_Level.Info)
    ' Write logs under INFO level.
    .Info ("This is an INFO message.") ' "yyyy-mm-dd hh:mm:ss [INFO]  (undefined) - This is an INFO message."
    .StackName ("foo") ' "yyyy-mm-dd hh:mm:ss [INFO] > foo - Start."
    .Info ("This is an INFO message.") ' "yyyy-mm-dd hh:mm:ss [INFO] > foo - This is an INFO message."
    .StackName ("bar") ' "yyyy-mm-dd hh:mm:ss [INFO] >> bar - Start."
    .Info ("This is an INFO message.") ' "yyyy-mm-dd hh:mm:ss [INFO] >> bar - This is an INFO message."
    .UnstackName   ' "yyyy-mm-dd hh:mm:ss [INFO] >> bar - End."
    .Info ("This is an INFO message.") ' "yyyy-mm-dd hh:mm:ss [INFO] > foo - This is an INFO message."
    .UnstackName ' "yyyy-mm-dd hh:mm:ss [INFO] > foo - End."
    .Info ("This is an INFO message.") ' "yyyy-mm-dd hh:mm:ss [INFO]  (undefined) - This is an INFO message."
  End With

End Sub
