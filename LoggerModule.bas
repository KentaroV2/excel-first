Attribute VB_Name = "LoggerModule"
Option Explicit
'! This provides (a) original logging functions in reference to log4j and (b) object names, method names, and property names logging functions with log messages.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------
' Define enumerations.
Public Enum Logger_Level '* Logger levels.
  Off '* Log nothing.
  Fatal
  Error__
  Warn
  Info '* Log start and end of functions.
  Debug__
  Trace
  All '* Log everything.
End Enum
