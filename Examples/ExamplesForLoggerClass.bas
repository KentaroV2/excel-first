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
  Logger.SetLevel (Logger_Level.Error__)
  ' Write ERROR log.
  Logger.Error ("This is an ERROR message.") ' "yyyy-mm-dd hh:mm:ss [ERROR]  (undefined category) - This is an ERROR message."
  
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
    .Fatal ("This is a FATAL message.") ' Logged.
    .Error ("This is an ERROR message.") ' Logged.
     .Warn ("This is a WARN message.") ' Logged.
    .Info ("This is an INFO message.") ' Logged.
    .Debug__ ("This is an DEBUG message.") ' (Not logged.)
    .Trace ("This is an TRACE message.") ' (Not logged.)
  End With
  
End Sub

'* This subrouutine teaches how to write logs with functions names.
Sub ExampleForLoggerClass_WriteLogsWithFunctionNames()
  
  ' Instantiate logger class.
  Dim Logger As LoggerClass
  Set Logger = New LoggerClass
  
  ' Use "With" statement to program in efficient way.
  With Logger
    ' Set INFO as logger level.
    .SetLevel (Logger_Level.Info)
    ' Write logs under INFO level.
    .Info ("This is an INFO message.") ' Write INFO log.
    .StackName ("foo") ' Stack function name. This line is expected to be written at the beginning of "foo" function.
    .Info ("This is an INFO message.") ' Write INFO log.
    .StackName ("bar") ' Stack function name. This line is expected to be written at the beginning of "bar" function.
    .Info ("This is an INFO message.") ' Write INFO log.
    .UnstackName   'Unstack function name. This line is expected to be written at the end of "bar" function.
    .Info ("This is an INFO message.") ' Write INFO log.
    .UnstackName 'Unstack function name. This line is expected to be written at the end of "foo" function.
    .Info ("This is an INFO message.") ' Write INFO log.
  End With

End Sub
