Attribute VB_Name = "LoggerModule"
'! @file LoggerModule.bas
'! This module provides logging functions in reference to log4j.
'! @copyright MIT
Option Explicit
' Declare Sleep function to release CPU resources periodically.
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' Define enumerations.
Public Enum Logger_Level '* Logger levels.
  Off
  Fatal
  Error__
  Warn
  Info
  Debug__
  Trace
  All
End Enum
