Attribute VB_Name = "ExamplesForExcelWorksheetClass"
 Option Explicit
'! This module provides some examples that help you understanding how to use the ExcelWorksheetClass.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
'
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------

'* This example teaches how to use CreateTable method.
'* @attention This example requires two Worksheets called "(SampleTable)" and "((SampleTable))".
Sub ExamplesForExcelWorksheetClass_CreateTable()
  
  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_CreateTable")
  
  ' Bind ExcelWorksheet.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_.ExcelWorkbook("").BindExcelWorksheetForTable("SampleTable"). _
    ExcelWorksheet("SampleTable")
  
  ' Disable screen updating.
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Create table.
  Call excelWorksheet_.CreateTable("item" & First_Level_Delimiter & "price")
  
  ' Enable screen updating.
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

'* This example teaches how to create new records with UpdateRecords method.
'* @attention This example requires two Worksheets called "(SampleTable)" and "((SampleTable))".
Sub ExamplesForExcelWorksheetClass_CreateNewRecordsWithUpdateRecords()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_CreateNewRecordsWithUpdateRecords")
  
  ' Bind ExcelWorksheet.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_.ExcelWorkbook("").BindExcelWorksheetForTable("SampleTable"). _
    ExcelWorksheet("SampleTable")
  
  ' Disable screen updating.
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Create table.
  With excelWorksheet_
    Call .CreateTable("item" & First_Level_Delimiter & "price")
  
  ' Create new records.
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(100))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "orange" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "cherry" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "plum" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(300))
    Dim updatedRecords_ As Range
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "grape" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(400), _
      , _
      updatedRecords_ _
      )
    Call logger_.Info("updatedRecords_.Rows.Count = " & CStr(updatedRecords_.Rows.Count)) ' "updatedRecords_.Rows.Count = 2"
    Call logger_.Info("updatedRecords_.Columns.Count = " & CStr(updatedRecords_.Columns.Count)) ' "updatedRecords_.Columns.Count = 2"
  End With
  
  ' Enable screen updating.
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to update records with UpdateRecords method.
'* @attention This example requires two Worksheets called "(SampleTable)" and "((SampleTable))".
Sub ExamplesForExcelWorksheetClass_UpdateRecordsWithUpdateRecords()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_UpdateRecordsWithUpdateRecords")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_.ExcelWorkbook("").BindExcelWorksheetForTable("SampleTable"). _
    ExcelWorksheet("SampleTable")
  
  ' Disable screen updating.
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Create table.
  With excelWorksheet_
    Call .CreateTable("item" & First_Level_Delimiter & "price")
  
  ' Create new records.
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(100))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "orange" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "cherry" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "plum" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(300))
    Dim updatedRecords_ As Range
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "grape" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(400), _
      , _
      updatedRecords_ _
      )
      
    ' Update records.
    Call .UpdateRecords( _
      "price" & First_Level_Delimiter & CStr(500), _
      "item" & First_Level_Delimiter & "*a*", _
      updatedRecords_ _
    )
    Call logger_.Info("updatedRecords_.Rows.Count = " & CStr(updatedRecords_.Rows.Count)) ' "updatedRecords_.Rows.Count = 4"
    Call logger_.Info("updatedRecords_.Columns.Count = " & CStr(updatedRecords_.Columns.Count)) ' "updatedRecords_.Columns.Count = 2"
  End With
  
  ' Enable screen updating.
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to insert new fields with UpdateRecords method.
'* @attention This example requires two Worksheets called "(SampleTable)" and "((SampleTable))".
Sub ExamplesForExcelWorksheetClass_InsertNewFieldsWithUpdateRecords()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_InsertNewFieldsWithUpdateRecords")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_.ExcelWorkbook("").BindExcelWorksheetForTable("SampleTable"). _
    ExcelWorksheet("SampleTable")
  
  ' Disable screen updating.
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Create table.
  With excelWorksheet_
    Call .CreateTable("item" & First_Level_Delimiter & "price__")
  
  ' Create new records.
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price__" & First_Level_Delimiter & CStr(100))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "orange" & _
        Second_Level_Delimiter & _
      "price__" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "cherry" & _
        Second_Level_Delimiter & _
      "price__" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "plum" & _
        Second_Level_Delimiter & _
      "price__" & First_Level_Delimiter & CStr(300))
    Dim updatedRecords_ As Range
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "grape" & _
        Second_Level_Delimiter & _
      "price__" & First_Level_Delimiter & CStr(400) _
      )
      
    ' Update records.
    Call .UpdateRecords( _
      "item__" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price__Y2019M04" & First_Level_Delimiter & CStr(120), _
      "item__" & First_Level_Delimiter & "apple" _
    )
    Call .UpdateRecords( _
      "item__" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price__Y2019M05" & First_Level_Delimiter & CStr(130), _
      "item__" & First_Level_Delimiter & "apple" _
    )
    Call .UpdateRecords( _
      "item__" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price__Y2019M03" & First_Level_Delimiter & CStr(110), _
      "item__" & First_Level_Delimiter & "apple", _
      updatedRecords_ _
    )
    Call logger_.Info("updatedRecords_.Rows.Count = " & CStr(updatedRecords_.Rows.Count)) ' "updatedRecords_.Rows.Count = 6"
    Call logger_.Info("updatedRecords_.Columns.Count = " & CStr(updatedRecords_.Columns.Count)) ' "updatedRecords_.Columns.Count = 5"
  End With
  
  ' Enable screen updating.
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

'* This example teaches how to use FilterRecords method.
'* @attention This example requires two Worksheets called "(SampleTable)" and "((SampleTable))".
Sub ExamplesForExcelWorksheetClass_FilterRecords()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_FilterRecords")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_.ExcelWorkbook("").BindExcelWorksheetForTable("SampleTable"). _
    ExcelWorksheet("SampleTable")
  
  ' Disable screen updating.
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Create table.
  With excelWorksheet_
    Call .CreateTable("item" & First_Level_Delimiter & "price")
  
  ' Create new records.
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(100))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "orange" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "cherry" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "plum" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(300))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "grape" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(400) _
      )
  
    ' Filter records.
    Dim filteredRecords_ As Range
    Call .FilterRecords( _
      "item" & First_Level_Delimiter & "*c*" & Or_Operator & "*g*", _
      "item" & First_Level_Delimiter & CStr(xlDescending), _
      "price", _
      filteredRecords_ _
    )
    Call logger_.Info("filteredRecords_.Rows.Count = " & CStr(filteredRecords_.Rows.Count)) ' "filteredRecords_.Rows.Count = 3"
    Call logger_.Info("filteredRecords_.Columns.Count = " & CStr(filteredRecords_.Columns.Count)) ' "filteredRecords_.Columns.Count = 2"
  End With
  
  ' Enable screen updating.
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

'* This example teaches how to use ReadRecords method.
'* @attention This example requires two Worksheets called "(SampleTable)" and "((SampleTable))".
Sub ExamplesForExcelWorksheetClass_ReadTable()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Off)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_ReadTable")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_.ExcelWorkbook("").BindExcelWorksheetForTable("SampleTable"). _
    ExcelWorksheet("SampleTable")
    
  ' Disable screen updating.
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Create table.
  With excelWorksheet_
    Call .CreateTable("item" & First_Level_Delimiter & "price")
  
  ' Create new records.
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(100))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "orange" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "cherry" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "plum" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(300))
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "grape" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(400) _
      )
  
    ' Filter records.
    Call .FilterRecords( _
      "item" & First_Level_Delimiter & "*c*" & Or_Operator & "*g*", _
      "item" & First_Level_Delimiter & CStr(xlDescending), _
      "price" _
    )
    
    ' Read records.
    Dim records_ As Range
    Call .ReadRecords(records_)
    Call logger_.Info("records_.Rows.Count = " & CStr(records_.Rows.Count)) ' "records_.Rows.Count = 3"
    Call logger_.Info("records_.Columns.Count = " & CStr(records_.Columns.Count)) ' "records_.Columns.Count = 2"
  End With
  
  ' Enable screen updating.
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


