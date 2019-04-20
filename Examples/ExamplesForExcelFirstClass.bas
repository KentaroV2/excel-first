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
  logger_.SetLevel (Logger_Level.Info)
  
  ' Stack name.
  logger_.StackName ("ExamplesForExcelFirstClass_DisableAndEnableScreenUpdating")
  
  ' Disable screen updating
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Enable screen updating
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  logger_.UnstackName
  
End Sub

'* This example teaches how to bind ExcelWorkbook.
Sub ExamplesForExcelFirstClass_BindExcelWorkbook()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  logger_.SetLevel (Logger_Level.Info)
  
  ' Stack name.
  logger_.StackName ("ExamplesForExcelFirstClass_BindExcelWorkbook")
  
  ' Bind ExcelWorkbook.
  Dim excelWorkbook_ As ExcelWorkbookClass
  Set excelWorkbook_ = excelFirst_.BindExcelWorkbook("")
  logger_.Info ("excelWorkbook_.Name = " & excelWorkbook_.Name) ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorkbook - excelWorkbook_.Name = This"
  
  ' Unstack name.
  logger_.UnstackName
  
End Sub

'* This example teaches how to bind ExcelWorkbook and ExcelWorksheet.
Sub ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  logger_.SetLevel (Logger_Level.Info)
  
  ' Stack name.
  logger_.StackName ("ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet")
  
  ' Bind ExcelWorkbook.
  Dim excelWorkbook_ As ExcelWorkbookClass
  Set excelWorkbook_ = excelFirst_.BindExcelWorkbook("")
  
  ' Bind ExcelWorksheet.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = excelWorkbook_.BindExcelWorksheet("Screen")
  logger_.Info ("excelWorksheet_.Name = " & excelWorksheet_.Worksheet.Name) ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet - excelWorksheet_.Name = Screen"
  
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
  logger_.SetLevel (Logger_Level.Info)
  
  ' Stack name.
  logger_.StackName ("ExamplesForExcelFirstClass_ClearWorksheet")
  
  ' Bind ExcelWorkbook.
  Dim excelWorkbook_ As ExcelWorkbookClass
  Set excelWorkbook_ = excelFirst_.BindExcelWorkbook("")
  
  ' Bind ExcelWorksheet.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = excelWorkbook_.BindExcelWorksheet("Screen")
  
  ' Clear Worksheet.
  excelWorksheet_.Clear
  
  ' Unstack name.
  logger_.UnstackName
  
End Sub

