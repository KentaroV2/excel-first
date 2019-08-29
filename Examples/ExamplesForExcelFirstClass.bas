Attribute VB_Name = "ExamplesForExcelFirstClass"
 Option Explicit
'! This module provides some examples that help you understanding how to use the ExcelFirstClass.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
'
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------


'* This example teaches how to disable and enable screen updating.
Sub ExamplesForExcelFirstClass_DisableAndEnableScreenUpdating()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_DisableAndEnableScreenUpdating")
  
  ' Disable screen updating
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Enable screen updating
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to change date to string.
Sub ExamplesForExcelFirstClass_ChangeDateToString()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_ChangeDateToString")
  
  ' Change date to string.
  Dim now__ As Date
  now__ = Now
  Dim string__ As String
  Dim degreeDetails_ As String
  degreeDetails_ = "Y"
  Call excelFirst_.ChangeDateToString(now__, string__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & string__)
  degreeDetails_ = "M"
  Call excelFirst_.ChangeDateToString(now__, string__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & string__)
  degreeDetails_ = "D"
  Call excelFirst_.ChangeDateToString(now__, string__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & string__)
  degreeDetails_ = "h"
  Call excelFirst_.ChangeDateToString(now__, string__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & string__)
  degreeDetails_ = "m"
  Call excelFirst_.ChangeDateToString(now__, string__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & string__)
  degreeDetails_ = "s"
  Call excelFirst_.ChangeDateToString(now__, string__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & string__)
  degreeDetails_ = "(undefined)"
  Call excelFirst_.ChangeDateToString(now__, string__)
  Call logger_.Info(degreeDetails_ & ":" & string__)
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to change string to date.
Sub ExamplesForExcelFirstClass_ChangeStringToDate()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_ChangeStringToDate")
  
  ' Change date to string.
  Dim now__ As Date
  now__ = Now
  Dim date__ As Date
  Dim degreeDetails_ As String
  degreeDetails_ = "Y"
  Call excelFirst_.ChangeStringToDate( _
    "Y" & Format(now__, "yyyy"), _
    date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(date__, "yyyy/mm/dd hh:nn:ss"))
  degreeDetails_ = "M"
  Call excelFirst_.ChangeStringToDate( _
    "Y" & Format(now__, "yyyy") & _
    "M" & Format(now__, "mm"), _
    date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(date__, "yyyy/mm/dd hh:nn:ss"))
  degreeDetails_ = "D"
  Call excelFirst_.ChangeStringToDate( _
    "Y" & Format(now__, "yyyy") & _
    "M" & Format(now__, "mm") & _
    "D" & Format(now__, "dd"), _
    date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(date__, "yyyy/mm/dd hh:nn:ss"))
  degreeDetails_ = "h"
  Call excelFirst_.ChangeStringToDate( _
    "Y" & Format(now__, "yyyy") & _
    "M" & Format(now__, "mm") & _
    "D" & Format(now__, "dd") & _
    "h" & Format(now__, "hh"), _
    date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(date__, "yyyy/mm/dd hh:nn:ss"))
  degreeDetails_ = "m"
  Call excelFirst_.ChangeStringToDate( _
    "Y" & Format(now__, "yyyy") & _
    "M" & Format(now__, "mm") & _
    "D" & Format(now__, "dd") & _
    "h" & Format(now__, "hh") & _
    "m" & Format(now__, "nn"), _
    date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(date__, "yyyy/mm/dd hh:nn:ss"))
  degreeDetails_ = "s"
  Call excelFirst_.ChangeStringToDate( _
    "Y" & Format(now__, "yyyy") & _
    "M" & Format(now__, "mm") & _
    "D" & Format(now__, "dd") & _
    "h" & Format(now__, "hh") & _
    "m" & Format(now__, "nn") & _
    "s" & Format(now__, "ss"), _
    date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(date__, "yyyy/mm/dd hh:nn:ss"))
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

