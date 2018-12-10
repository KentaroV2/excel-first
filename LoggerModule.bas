Attribute VB_Name = "LoggerModule"
'! @file LoggerModule.bas
'! @brief Provides logging functions in reference to log4j.
'! @copyright MIT
Option Explicit
' Define enumerations.
Public Enum loggerLevel '* Represents logger levels.
  OFF_LEVEL
  FATAL_LEVEL
  ERROR_LEVEL
  WARN_LEVEL
  INFO_LEVEL
  DEBUG_LEVEL
  TRACE_LEVEL
  ALL_LEVEL
End Enum
