Attribute VB_Name = "ExamplesForLoggerClass"
'! @file ExamplesForLoggerClass.bas
'! @brief Provides some examples that help you understanding how to use the "LoggerClass" class.
'! @copyright MIT
Option Explicit

'* @fn Sub ExampleForLoggerClass_UseLogger()
'* @brief Teaches how to use logger.
Sub ExampleForLoggerClass_UseLogger()
  
  ' Instantiate logger class.
  Dim logger As LoggerClass
  Set logger = New LoggerClass
  
  ' Set ERROR as logger level.
  Dim exitStatus As Long
  exitStatus = logger.SetLevel(loggerLevel.ERROR_LEVEL)
  ' Write ERROR log.
  exitStatus = logger.Error("This is an ERROR message.") ' "yyyy-mm-dd hh:mm:ss [ERROR]  (undefined category) - This is an ERROR message."
  
End Sub

'* @fn Sub ExampleForLoggerClass_WriteLogs()
'* @brief Teaches how to write log under defined logger level.
Sub ExampleForLoggerClass_WriteLogs()
  
  ' Instantiate logger class.
  Dim logger As LoggerClass
  Set logger = New LoggerClass
  
  ' Use "With" statement to program in efficient way.
  With logger
    ' Set INFO as logger level.
    Dim exitStatus As Long
    exitStatus = .SetLevel(loggerLevel.INFO_LEVEL)
    ' Write logs under INFO level.
    exitStatus = .Fatal("This is a FATAL message.") ' Logged.
    exitStatus = .Error("This is an ERROR message.") ' Logged.
    exitStatus = .Warn("This is a WARN message.") ' Logged.
    exitStatus = .Info("This is an INFO message.") ' Logged.
    exitStatus = .Debug_("This is an DEBUG message.") ' (Not logged.)
    exitStatus = .Trace("This is an TRACE message.") ' (Not logged.)
  End With
  
End Sub

'* @fn Sub ExampleForLoggerClass_WriteLogsWithFunctionNames()
'* @brief Teaches how to write logs with functions names.
Sub ExampleForLoggerClass_WriteLogsWithFunctionNames()
  
  ' Instantiate logger class.
  Dim logger As LoggerClass
  Set logger = New LoggerClass
  
  ' Use "With" statement to program in efficient way.
  With logger
    ' Set INFO as logger level.
    Dim exitStatus As Long
    exitStatus = .SetLevel(loggerLevel.INFO_LEVEL)
    ' Write logs under INFO level.
    exitStatus = .Info("This is an INFO message.") ' Write INFO log.
    exitStatus = .StackFunctionName("foo") ' Stack function name. This line is expected to be written at the beginning of "foo" function.
    exitStatus = .Info("This is an INFO message.") ' Write INFO log.
    exitStatus = .StackFunctionName("bar") ' Stack function name. This line is expected to be written at the beginning of "bar" function.
    exitStatus = .Info("This is an INFO message.") ' Write INFO log.
    exitStatus = .UnstackFunctionName() 'Unstack function name. This line is expected to be written at the end of "bar" function.
    exitStatus = .Info("This is an INFO message.") ' Write INFO log.
    exitStatus = .UnstackFunctionName() 'Unstack function name. This line is expected to be written at the end of "foo" function.
    exitStatus = .Info("This is an INFO message.") ' Write INFO log.
  End With

End Sub