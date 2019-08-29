Attribute VB_Name = "ExamplesForExcelWorkbookClass"
 Option Explicit
'! This module provides some examples that help you understanding how to use the ExcelWorkbookClass.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
'
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------


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
  Set excelWorkbook_ = _
    excelFirst_.BindExcelWorkbook(This) _
    .ExcelWorkbook(This)
  Call logger_.Info("excelWorkbook_.Name = " & excelWorkbook_.Name)  ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorkbook - excelWorkbook_.Name = This"
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to bind external ExcelWorkbook.
'* @attention This example requires a "test.xls" file on a directory where a file running this example locates.
Sub ExamplesForExcelFirstClass_BindExternalExcelWorkbook()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_BindExternalExcelWorkbook")
  
  ' Bind ExcelWorkbook.
  Dim excelWorkbook_ As ExcelWorkbookClass
  Set excelWorkbook_ = _
    excelFirst_.BindExcelWorkbook("test.xlsx") _
    .ExcelWorkbook("test.xlsx")
  Call logger_.Info("excelWorkbook_.Name = " & excelWorkbook_.Name)  ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorkbook - excelWorkbook_.Name = This"
  
  ' Unbind ExcelWorkbook.
  Call excelFirst_.ExcelWorkbook("test.xlsx").Unbind
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

