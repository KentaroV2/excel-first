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

'* This example teaches how to bind ExcelWorkbook.
Sub ExamplesForExcelFirstClass_BindExcelWorkbook()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_BindExcelWorkbook")
  
  ' Bind ExcelWorkbook.
  Dim excelWorkbook_ As ExcelWorkbookClass
  Set excelWorkbook_ = excelFirst_.BindExcelWorkbook("").ExcelWorkbook("")
  Call logger_.Info("excelWorkbook_.Name = " & excelWorkbook_.Name)  ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorkbook - excelWorkbook_.Name = This"
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

'* This example teaches how to bind ExcelWorkbook and ExcelWorksheet.
'* @attention This example requires "Screen" Worksheet.
Sub ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = excelFirst_.ExcelWorkbook("").BindExcelWorksheet("Screen").ExcelWorksheet("Screen")
  logger_.Info ("excelWorksheet_.Name = " & excelWorksheet_.Name) ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet - excelWorksheet_.Name = Screen"
  
  ' Unstack name.
  logger_.UnstackName
  
End Sub

'* This example teaches how to clear Worksheet.
'* @warning This example requires a Worksheet called "Sheet1".
Sub ExamplesForExcelFirstClass_ClearWorksheet()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_ClearWorksheet")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = excelFirst_.ExcelWorkbook("").BindExcelWorksheet("Screen").ExcelWorksheet("Screen")
  
  ' Clear Worksheet.
  excelFirst_.ScreenUpdatingFlag = False
  Call excelWorksheet_.Clear
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

