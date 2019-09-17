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


'* This example teaches how to bind ExcelWorksheet.
'* @attention This example requires "SampleSheet" Worksheet.
Sub ExamplesForExcelFirstClass_BindExcelWorksheet()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet")
  
  ' Bind ExcelWorksheet.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet("SampleSheet") _
    .ExcelWorksheet("SampleSheet")
    ' The following line also work.
'  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleSheet", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.None) _
      ) _
    .ExcelWorksheet("SampleSheet")
  logger_.Info ("excelWorksheet_.Name = " & excelWorksheet_.Name) ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet - excelWorksheet_.Name = SampleSheet"
  logger_.Info ("excelWorksheet_.Worksheet.Name = " & excelWorksheet_.Worksheet.Name) ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorkbookAndExcelWorksheet - excelWorksheet_.Worksheet.Name = SampleSheet"
  
  ' Unbind ExcelWorksheet.
  excelFirst_.ExcelWorkbook(This) _
    .UnbindExcelWorksheet ("SampleSheet")
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to bind ExcelWorksheet to access the sheet as database.
'* @attention This example requires two Worksheets called "(SampleTable)" and "((SampleTable))".
Sub ExamplesForExcelFirstClass_BindExcelWorksheetToAccessTheSheetAsDatabase()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_BindExcelWorksheetToAccessTheSheetAsDatabase")
  
  ' Bind ExcelWorksheet.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This).BindExcelWorksheet( _
      "SampleTable", _
      "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftExcelWorksheet) _
    ) _
    .ExcelWorksheet("SampleTable")
  logger_.Info ("excelWorksheet_.Name = " & excelWorksheet_.Name) ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorksheetToAccessTheSheetAsDatabase - excelWorksheet_.Name = SampleTable"
  
  ' Unbind ExcelWorksheet.
  excelFirst_.ExcelWorkbook(This) _
    .UnbindExcelWorksheet ("SampleTable")
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to bind ExcelWorksheet to access Oracle database.
'* @attention This example requires "SampleTable" Worksheet.
'* @attention This example requires proper (a) data source, (b) user id and (c) password by replacing <data source>, <user id> and <password>, respectively.
Sub ExamplesForExcelFirstClass_BindExcelWorksheetToAccessOracleDatabase()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_BindExcelWorksheetToAccessOracleDatabase")
  
  ' Bind ExcelWorksheet.
  ' (Note) Replace the following <data source>, <user id>, and <password> with proper values.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.Oracle) & _
        Second_Level_Delimiter & _
        "DataSource" & First_Level_Delimiter & "<data source>" & _
        Second_Level_Delimiter & _
        "User" & First_Level_Delimiter & "<user id>" & _
        Second_Level_Delimiter & _
        "Password" & First_Level_Delimiter & "<password>" _
    ) _
    .ExcelWorksheet("SampleTable")
  logger_.Info ("excelWorksheet_.Name = " & excelWorksheet_.Name) ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorksheetToAccessOracleDatabase - excelWorksheet_.Name = SampleTable"
  
  ' Execute sample SQL.
  Call excelWorksheet_.ExecuteSQL("SELECT TABLE_NAME FROM ALL_TABLES ORDER BY OWNER,TABLE_NAME")
  
  ' Unbind ExcelWorksheet.
  excelFirst_.ExcelWorkbook(This) _
    .UnbindExcelWorksheet ("SampleTable")
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to bind ExcelWorksheet to access database on Microsoft Access.
'* @attention This example requires "SampleTable" Worksheet.
'* @attention This example requires proper (a) data source, (b) user id and (c) password by replacing <data source>, <user id> and <password>, respectively.
Sub ExamplesForExcelFirstClass_BindExcelWorksheetToAccessDatabaseOnMicrosoftAccess()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_BindExcelWorksheetToAccessDatabaseOnMicrosoftAccess")
  
  ' Bind ExcelWorksheet.
  ' (Note) Replace the following <data source>, <user id>, and <password> with proper values.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftAccess) & _
        Second_Level_Delimiter & _
        "DataSource" & First_Level_Delimiter & "<data source>" & _
        Second_Level_Delimiter & _
        "User" & First_Level_Delimiter & "<user id>" & _
        Second_Level_Delimiter & _
        "Password" & First_Level_Delimiter & "<password>" _
    ) _
    .ExcelWorksheet("SampleTable")
  logger_.Info ("excelWorksheet_.Name = " & excelWorksheet_.Name) ' "yyyy-mm-dd hh:mm:ss [INFO] > ExamplesForExcelFirstClass_BindExcelWorksheetToAccessOracleDatabase - excelWorksheet_.Name = SampleTable"
  
  ' Execute sample SQL.
  Call excelWorksheet_.ExecuteSQL("SELECT TABLE_NAME FROM ALL_TABLES ORDER BY OWNER,TABLE_NAME")
  
  ' Unbind ExcelWorksheet.
  excelFirst_.ExcelWorkbook(This) _
    .UnbindExcelWorksheet ("SampleTable")
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


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
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftExcelWorksheet) _
      ) _
    .ExcelWorksheet("SampleTable")

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
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftExcelWorksheet) _
      ) _
    .ExcelWorksheet("SampleTable")
  
  ' Disable screen updating.
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Create table.
  With excelWorksheet_
    Call .CreateTable("item" & First_Level_Delimiter & "price")
  
  ' Create new records.
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "apple" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(100) _
    )
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "orange" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200) _
    )
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "cherry" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(200) _
    )
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "plum" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(300) _
    )
    Dim updatedRecords_ As Range
    Call .UpdateRecords( _
      "item" & First_Level_Delimiter & "grape" & _
        Second_Level_Delimiter & _
      "price" & First_Level_Delimiter & CStr(400), _
      , _
      True, _
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
Sub ExamplesForExcelWorksheetClass_UpdateRecordsWithUpdateRecordsMethod()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_UpdateRecordsWithUpdateRecordsMethod")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftExcelWorksheet) _
      ) _
    .ExcelWorksheet("SampleTable")
  
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
      "price" & First_Level_Delimiter & CStr(400) _
    )
      
    ' Update records.
    Call .UpdateRecords( _
      "price" & First_Level_Delimiter & CStr(500), _
      "item" & First_Level_Delimiter & "*a*", _
      True, _
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
Sub ExamplesForExcelWorksheetClass_InsertNewFieldsWithUpdateRecordsMethod()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_InsertNewFieldsWithUpdateRecordsMethod")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftExcelWorksheet) _
      ) _
    .ExcelWorksheet("SampleTable")
  
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
      True, _
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
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftExcelWorksheet) _
      ) _
    .ExcelWorksheet("SampleTable")
  
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
      True, _
      filteredRecords_ _
    )
    If filteredRecords_ Is Nothing Then
      Stop
    End If
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
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_ReadTable")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftExcelWorksheet) _
      ) _
    .ExcelWorksheet("SampleTable")
    
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

'* This example teaches how to use DeleteRecords method.
'* @attention This example requires two Worksheets called "(SampleTable)" and "((SampleTable))".
Sub ExamplesForExcelWorksheetClass_DeleteRecords()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelWorksheetClass_DeleteRecords")
  
  ' Bind ExcelWorksheet. (Note) ExcelWorkbook that refers Application.ThisWorkbook is already created by ExcelFirst object.
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet( _
        "SampleTable", _
        "ConnectionType" & First_Level_Delimiter & CStr(Database_Connection_Type.MicrosoftExcelWorksheet) _
      ) _
    .ExcelWorksheet("SampleTable")
    
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
  
    ' Delete records.
    Dim deletedRecords_ As Range
    Call .DeleteRecords( _
      "item" & First_Level_Delimiter & "*c*" & Or_Operator & "*g*", _
      True, _
      deletedRecords_ _
    )
    
    ' Read records.
    Call logger_.Info("deletedRecords_.Rows.Count = " & CStr(deletedRecords_.Rows.Count)) ' "deletedRecords_.Rows.Count = 3"
    Call logger_.Info("deletedRecords_.Columns.Count = " & CStr(deletedRecords_.Columns.Count)) ' "deletedRecords_.Columns.Count = 2"
  End With
  
  ' Enable screen updating.
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to clear Worksheet.
'* @attention This example requires "Screen" Worksheet.
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
  Set excelWorksheet_ = _
    excelFirst_ _
    .ExcelWorkbook(This) _
      .BindExcelWorksheet("Screen") _
    .ExcelWorksheet("Screen")
  
  ' Clear Worksheet.
  excelFirst_.ScreenUpdatingFlag = False
  Call excelWorksheet_.Clear
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub
