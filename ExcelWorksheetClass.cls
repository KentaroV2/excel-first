VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelWorksheetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'! This class provides attributes and methods for Microsoft Excel Worksheet.
'! This class contains (a) Worksheet object, and (c) data manipulating functions.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------
' Define members.
Private myName As String '* Name (also needed for logging purpose). The name also represent Worksheet name.
Private myParent As Object '* Parent object. This object should be ExcelWorkbook.
Private myLogger As LoggerClass '* Logger object
Private myWorksheet As Worksheet '* Worksheet object.
Private myDatabaseConnectionType As Long '* Database connection type.
Private myParametersForDatabaseConnection As String '* Parameters for database connection.
Private myWorksheetForEditingTable As Worksheet '* Worksheet object for editing table.
Private myWorksheetForReadingTable As Worksheet '* Worksheet object for reading table.
Private myDatabaseConnection As Object '* Database connection object.
Private myRecordset As Object '* Recordset object.

'* Initialize this class.
Private Sub Class_Initialize()
  myName = Undefined
End Sub

'* Terminate this class.
Private Sub Class_Terminate()
  Set myWorksheet = Nothing
End Sub

'* Set Name property.
Public Property Let Name(Name As String)
  myName = Name
End Property

'* Get Name property.
Property Get Name() As String
  Name = myName
End Property

'* Set Parent property.
Public Property Set Parent(Parent As Object)
  Set myParent = Parent
End Property

'* Get Parent property.
Public Property Get Parent() As Object
  Set Parent = myParent
End Property

'* Set Logger property.
Public Property Set Logger(Logger As Object)
  Set myLogger = Logger
End Property

'* Get Logger property.
Public Property Get Logger() As LoggerClass
  Set Logger = myLogger
End Property

'* Set Worksheet property.
Property Set Worksheet(Worksheet__ As Worksheet)
  Set myWorksheet = Worksheet__
End Property

'* Get Worksheet property.
Property Get Worksheet() As Worksheet
  Set Worksheet = myWorksheet
End Property

'* Set DatabaseConnectionType property.
Public Property Let DatabaseConnectionType(DatabaseConnectionType As Long)
  myDatabaseConnectionType = DatabaseConnectionType
End Property

'* Get DatabaseConnectionType property.
Property Get DatabaseConnectionType() As Long
  DatabaseConnectionType = myDatabaseConnectionType
End Property

'* Set WorksheetForEditingTable property.
Property Set WorksheetForEditingTable(WorksheetForEditingTable As Worksheet)
  Set myWorksheetForEditingTable = WorksheetForEditingTable
End Property

'* Get WorksheetForEditingTable property.
Property Get WorksheetForEditingTable() As Worksheet
  Set WorksheetForEditingTable = myWorksheetForEditingTable
End Property

'* Set WorksheetForReadingTable property.
Property Set WorksheetForReadingTable(WorksheetForReadingTable As Worksheet)
  Set myWorksheetForReadingTable = WorksheetForReadingTable
End Property

'* Get WorksheetForReadingTable property.
Property Get WorksheetForReadingTable() As Worksheet
  Set WorksheetForReadingTable = myWorksheetForReadingTable
End Property

'* Bind Worksheet.
'* @param[in] Name. This name should be Worksheet name.
'* @param[in] ParametersForDatabaseConnection. (Optional) Parameters for database connection.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function Bind( _
  Name As String, _
  Optional ParametersForDatabaseConnection As String, _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "Bind")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  If (ParametersForDatabaseConnection <> "") Then
    Dim parameters_ As Variant
    parameters_ = Split(ParametersForDatabaseConnection, Second_Level_Delimiter)
    Dim databaseConnectionType_ As Long
    databaseConnectionType_ = Database_Connection_Type.Miscellaneous
    Dim indexForParameters_ As Long
    For indexForParameters_ = LBound(parameters_) To UBound(parameters_)
      Dim keyAndValue_ As Variant
      keyAndValue_ = Split(parameters_(indexForParameters_), First_Level_Delimiter)
      Dim key_ As String
      key_ = keyAndValue_(LBound(keyAndValue_))
      Dim value_ As String
      value_ = keyAndValue_(UBound(keyAndValue_))
      Select Case key_
        Case "ConnectionType"
          databaseConnectionType_ = CLng(value_)
        Case "DataSource"
          Dim dataSource_ As String
          dataSource_ = CStr(value_)
        Case "User"
          Dim user_ As String
          user_ = CStr(value_)
        Case "Password"
          Dim password_ As String
          password_ = CStr(value_)
      End Select
    Next
  Else
    databaseConnectionType_ = Database_Connection_Type.None
  End If
  Me.DatabaseConnectionType = databaseConnectionType_
  Dim parent_ As Object
  Set parent_ = Me.Parent
  Dim workbook_ As Workbook
  Set workbook_ = parent_.Workbook
  Dim worksheets_ As Sheets
  Set worksheets_ = workbook_.Worksheets
  Select Case databaseConnectionType_
    Case Database_Connection_Type.None
      Dim worksheet_ As Worksheet
      For Each worksheet_ In worksheets_
          If (worksheet_.Name = Name) Then
            Set Me.Worksheet = worksheet_
          End If
      Next
      If (Me.Worksheet Is Nothing) Then
        Call Err.Raise(Exit_Status.Worksheet_Is_Not_Found, Me, "Worksheet_Is_Not_Found")
      Else
        Set Me.WorksheetForEditingTable = Nothing
        Set Me.WorksheetForReadingTable = Nothing
      End If
    Case _
      Database_Connection_Type.MicrosoftExcelWorksheet, _
      Database_Connection_Type.Oracle, _
      Database_Connection_Type.MicrosoftAccess
      For Each worksheet_ In worksheets_
        If ( _
          worksheet_.Name = _
            Left_Parentheses & _
            Name & _
            Right_Parentheses _
        ) Then
          Set Me.WorksheetForEditingTable = worksheet_
        ElseIf ( _
          worksheet_.Name = _
            Left_Parentheses & _
            Left_Parentheses & _
            Name & _
            Right_Parentheses & _
            Right_Parentheses _
        ) Then
          Set Me.WorksheetForReadingTable = worksheet_
        End If
      Next
      Select Case databaseConnectionType_
        Case Database_Connection_Type.MicrosoftExcelWorksheet
          If ( _
            Me.WorksheetForEditingTable Is Nothing _
            Or _
            Me.WorksheetForReadingTable Is Nothing _
          ) Then
            Call Err.Raise(Exit_Status.Worksheet_Is_Not_Found, Me, "Worksheet_Is_Not_Found")
          Else
            Set Me.Worksheet = Nothing
          End If
        Case _
          Database_Connection_Type.Oracle, _
          Database_Connection_Type.MicrosoftAccess
          If (Me.WorksheetForReadingTable Is Nothing) Then
            Call Err.Raise(Exit_Status.Worksheet_Is_Not_Found, Me, "Worksheet_Is_Not_Found")
          Else
            Set Me.Worksheet = Nothing
          End If
          Select Case databaseConnectionType_
            Case Database_Connection_Type.Oracle
              Dim connection_ As String
              connection_ = "Provider=OraOLEDB.Oracle"
              If (dataSource_ <> "") Then
                connection_ = connection_ & ";" & "Data Source=" & dataSource_
                If (user_ <> "") Then
                  connection_ = connection_ & ";" & "User ID=" & user_
                  If (password_ <> "") Then
                    connection_ = connection_ & ";" & "Password=" & password_
                  End If
                End If
              End If
            Case Database_Connection_Type.MicrosoftAccess
              connection_ = "Provider=Microsoft.ACE.OLEDB.12.0"
              If (dataSource_ <> "") Then
                connection_ = connection_ & ";" & "Data Source=" & dataSource_
                If (user_ <> "") Then
                  connection_ = connection_ & ";" & "User Id=" & user_
                  If (password_ <> "") Then
                    connection_ = connection_ & ";" & "Password=" & password_
                  End If
                End If
              End If
            End Select
            If (dataSource_ = "") Then
              Call Err.Raise(Exit_Status.Parameters_For_Database_Connection_Is_Not_Defined, Me, "Parameters_For_Database_Connection_Is_Not_Defined")
            End If
          Set myDatabaseConnection = CreateObject("ADODB.Connection")
          Set myRecordset = CreateObject("ADODB.Recordset")
          myDatabaseConnection.Open connection_
        End Select
    Case Else
      Call Err.Raise(Exit_Status.Database_Connection_Type_Is_Not_Valid, Me, "Database_Connection_Type_Is_Not_Valid")
  End Select
  Me.Name = Name
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set Bind = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Unbind Worksheet.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function Unbind( _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "Unbind")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Select Case Me.DatabaseConnectionType
    Case _
      Database_Connection_Type.None, _
      Database_Connection_Type.MicrosoftExcelWorksheet
      ' Do nothing.
    Case _
      Database_Connection_Type.MicrosoftAccess, _
      Database_Connection_Type.Oracle
      myDatabaseConnection.Close
  End Select
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set Unbind = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Find keyword.
'* @param[in] What. Keyword.
'* @param[in] LookIn. (Optional) Look in. Default value is xlValues.
'* @param[in] FoundRow. (Optional) Look in. Default value is xlValues.
'* @param[in] FoundColumn. (Optional) Look in. Default value is xlValues.
'* @param[in] LookAt. (Optional) Look at. Default value is xlWhole.
'* @param[in] SearchOrder. (Optional) Search order. Default value is xlByRows.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function Find( _
  What As String, _
  Optional FoundRow As Long, _
  Optional FoundColumn As Long, _
  Optional LookIn As Long = xlValues, _
  Optional LookAt As Long = xlWhole, _
  Optional SearchOrder As Long = xlByRows, _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "Find")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Select Case Me.DatabaseConnectionType
    Case _
      Database_Connection_Type.None
      ' Do something.
      With Me.Worksheet
        Dim range_ As Range
        Set range_ = .Cells.Find( _
          What:=What, _
          After:=.Cells(.Cells.Rows.Count, .Cells.Columns.Count), _
          LookIn:=LookIn, _
          LookAt:=LookAt, _
          SearchOrder:=SearchOrder _
        )
      End With
      If (range_ Is Nothing) Then
        FoundRow = 0
        FoundColumn = 0
      Else
        FoundRow = range_.Row
        FoundColumn = range_.Column
      End If
    Case Else
      Call Err.Raise(Exit_Status.Database_Connection_Type_Is_Not_Valid, Me, "Database_Connection_Type_Is_Not_Valid")
  End Select
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set Find = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Clear Worksheet.
'* @param[in] FontName. (Optional) Font name.
'* @param[in] FontSize. (Optional) Font size.
'* @param[in] FontColor. (Optional) Font color.
'* @param[in] RowHeight. (Optional) Row height.
'* @param[in] ColumnWidth. (Optional) Column width.
'* @param[in] HorizontalAlignment. (Optional) Horizontal alignment
'* @param[in] VerticalAlignment. (Optional) Vertical alignment
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Clear( _
  Optional FontName As String = Font_Name, _
  Optional FontSize As Double = Font_Size, _
  Optional FontColor As Long = Font_Color, _
  Optional RowHeight As Double = Row_Height, _
  Optional ColumnWidth As Long = Column_Width, _
  Optional HorizontalAlignment As Long = Horizontal_Alignment, _
  Optional VerticalAlignment As Long = Vertical_Alignment, _
  Optional ExitStatus As Long _
)
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "Clear")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Dim activeWorksheet_ As Worksheet
  Set activeWorksheet_ = Me.Parent.Workbook.ActiveSheet
  With Me.Worksheet
    .Activate
    If (.AutoFilterMode) Then
      .Cells.AutoFilter
    End If
    .Cells.Delete
    With .Cells
      .Font.Name = FontName
      .Font.Size = FontSize
      .Font.Color = FontColor
      .RowHeight = RowHeight
      .ColumnWidth = ColumnWidth
      .HorizontalAlignment = HorizontalAlignment
      .VerticalAlignment = VerticalAlignment
    End With
  End With
  activeWorksheet_.Activate
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Exit Sub
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Sub

'* Bind Worksheet for table.
'* @param[in] Name. This name should be Worksheet name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
'Public Function BindWorksheetForTable(Name As String, Optional ExitStatus As Long) As ExcelWorksheetClass
'Initialize:
'  On Error GoTo Error
'  Dim logger_ As LoggerClass
'  Set logger_ = Me.Logger
'  Call logger_.StackName(Me.Name & "." & "BindWorksheetForTable")
'  Dim exitStatus_ As Long
'  exitStatus_ = Exit_Status.Success
'Execute:
'  Dim parent_ As Object
'  Set parent_ = Me.Parent
'  Dim workbook_ As Workbook
'  Set workbook_ = parent_.Workbook
'  Dim worksheets_ As Sheets
'  Set worksheets_ = workbook_.Worksheets
'  Dim worksheet_ As Worksheet
'  For Each worksheet_ In worksheets_
'      If ( _
'        worksheet_.Name = _
'          Left_Parentheses & _
'          Name & _
'          Right_Parentheses _
'      ) Then
'        Set Me.WorksheetForEditingTable = worksheet_
'      ElseIf ( _
'        worksheet_.Name = _
'          Left_Parentheses & _
'          Left_Parentheses & _
'          Name & _
'          Right_Parentheses & _
'          Right_Parentheses _
'      ) Then
'        Set Me.WorksheetForReadingTable = worksheet_
'      End If
'  Next
'  If ( _
'    Me.WorksheetForEditingTable Is Nothing _
'    Or _
'    Me.WorksheetForReadingTable Is Nothing _
'  ) Then
'    exitStatus_ = Exit_Status.Worksheet_Is_Not_Found
'  Else
'    Set Me.Worksheet = Nothing
'    Me.Name = Name
'  End If
'Exit__:
'  Call logger_.UnstackName
'  ExitStatus = exitStatus_
'  Set BindWorksheetForTable = Me
'  Exit Function
'Error:
'  exitStatus_ = Err.Number
'  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
'  GoTo Exit__
'End Function

'* Create table.
'* @param[in] SetOfFields. Set of fields.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function CreateTable( _
  SetOfFields As String, _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "CreateTable")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  With Me
    Set .Worksheet = .WorksheetForEditingTable
    .Clear
    Set .Worksheet = .WorksheetForReadingTable
    .Clear
    Set .Worksheet = Nothing
  End With
  If (SetOfFields = "") Then
    Call Err.Raise(Exit_Status.SetOfFields_Is_Not_Defined, , "SetOfFields_Is_Blank")
  End If
  Dim Fields As Variant
  Fields = Split(SetOfFields, First_Level_Delimiter)
  Dim worksheet_ As Worksheet
  '* @todo Should we clear not only text but also other elements like drawings?
  With Me.WorksheetForEditingTable
    .Cells.Clear
    .Range( _
        .Cells(Worksheet_Row.Top, Worksheet_Column.Left__), _
        .Cells(Worksheet_Row.Top, Worksheet_Column.Left__ + UBound(Fields) - LBound(Fields)) _
      ) = Fields
  End With
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set CreateTable = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Filter records.
'* @param[in] SetOfFiltedFields. (Optional) Set of filtered fields.
'* @param[in] SetOfSortedFields. (Optional) Set of sorted fields.
'* @param[in] DuplicatesRemovingField. (Optional) Duplicates removing field.
'* @param[in] IsResultRead. (Optional) Flag whether filtered result is read later on. Default value is False.
'* @param[out] FilteredRecords. (Optional) Set of filtered records.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function FilterRecords( _
  Optional SetOfFiltedFields As String, _
  Optional SetOfSortedFields As String, _
  Optional DuplicatesRemovingField As String, _
  Optional IsResultRead As Boolean = False, _
  Optional FilteredRecords As Range, _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "FilterRecords")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Dim worksheet_ As Worksheet
  Set worksheet_ = Me.WorksheetForEditingTable
  With worksheet_
    Dim fields_ As Variant
    fields_ = .Range( _
      .Cells(Table_Row.Field, Table_Column.Left__), _
      .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft) _
    )
    .Cells.AutoFilter
    If Not (.AutoFilterMode) Then
      .Cells.AutoFilter
    End If
    If (SetOfFiltedFields <> "") Then
      Dim filteredFields_ As Variant
      filteredFields_ = Split(SetOfFiltedFields, Second_Level_Delimiter)
      Dim indexOfFilteredFields_ As Long
      For indexOfFilteredFields_ = LBound(filteredFields_) To UBound(filteredFields_)
        Dim filteredField_ As Variant
        filteredField_ = Split(filteredFields_(indexOfFilteredFields_), First_Level_Delimiter)
        Dim indexOfFields_ As Long
        For indexOfFields_ = LBound(fields_, 2) To UBound(fields_, 2)
          If (filteredField_(Field_And_Value.Field) = fields_(Table_Row.Field, indexOfFields_)) Then
            Exit For
          End If
        Next
        If ( _
          indexOfFields_ >= LBound(fields_, 2) _
          And _
          indexOfFields_ <= UBound(fields_, 2) _
        ) Then
          If (InStr(filteredField_(Field_And_Value.Value), Or_Operator) = 0) Then
            Dim filters_ As Variant
            filters_ = Array(filteredField_(Field_And_Value.Value))
          Else
            filters_ = Split(filteredField_(Field_And_Value.Value), Or_Operator)
          End If
          Call .Cells.AutoFilter( _
            indexOfFields_, _
            filters_, _
            xlFilterValues _
          )
        End If
      Next
    End If
    With .AutoFilter.Sort
      With .SortFields
        .Clear
        If (SetOfSortedFields <> "") Then
          Dim sortedFields_ As Variant
          sortedFields_ = Split(SetOfSortedFields, Second_Level_Delimiter)
          Dim indexOfSortedFields_ As Long
          For indexOfSortedFields_ = LBound(sortedFields_) To UBound(sortedFields_)
            Dim sortedField_ As Variant
            sortedField_ = Split(sortedFields_(indexOfSortedFields_), First_Level_Delimiter)
            For indexOfFields_ = LBound(fields_, 2) To UBound(fields_, 2)
              If (sortedField_(Field_And_Value.Field) = fields_(Table_Row.Field, indexOfFields_)) Then
                Exit For
              End If
            Next
            If ( _
              indexOfFields_ >= LBound(fields_, 2) _
              And _
              indexOfFields_ <= UBound(fields_, 2) _
            ) Then
            .Add _
              Key:=worksheet_.Range( _
                worksheet_.Columns(indexOfFields_), _
                worksheet_.Columns(indexOfFields_) _
              ), _
              SortOn:=xlSortOnValues, _
              Order:=CLng(sortedField_(Field_And_Value.Value)), _
              DataOption:=xlSortNormal
            End If
          Next
          With worksheet_.AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
          End With
        End If
      End With
    End With
    Dim filteredAndSortedRecords As Range
    Set filteredAndSortedRecords = .Range( _
      .Cells(Table_Row.Field, Table_Column.Left__), _
      .Cells( _
        .Cells(.Rows.Count, Table_Column.Left__).End(xlUp).Row, _
        .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft).Column _
      ) _
    ).SpecialCells(xlCellTypeVisible)
  End With
  If ( _
    DuplicatesRemovingField <> "" _
    Or _
    IsResultRead _
  ) Then
    With Me.WorksheetForReadingTable
      .Cells.Clear
      Dim indexOfArea_ As Long
      Dim copiedRow_ As Long
      copiedRow_ = Table_Row.Field
      For indexOfArea_ = 1 To filteredAndSortedRecords.Areas.Count
        Dim row_ As Long
        row_ = filteredAndSortedRecords.Areas(indexOfArea_).Row
        Dim column_ As Long
        column_ = filteredAndSortedRecords.Areas(indexOfArea_).Column
        Dim rowsCount_ As Long
        rowsCount_ = filteredAndSortedRecords.Areas(indexOfArea_).Rows.Count
        Dim columnsCount_ As Long
        columnsCount_ = filteredAndSortedRecords.Areas(indexOfArea_).Columns.Count
        .Range( _
          .Cells(copiedRow_, column_), _
          .Cells(copiedRow_ + rowsCount_ - 1, columnsCount_) _
        ).Value = filteredAndSortedRecords.Areas(indexOfArea_).Value
        copiedRow_ = copiedRow_ + rowsCount_
      Next
      If (DuplicatesRemovingField <> "") Then
        For indexOfFields_ = LBound(fields_, 2) To UBound(fields_, 2)
          If (DuplicatesRemovingField = fields_(Table_Row.Field, indexOfFields_)) Then
            Exit For
          End If
        Next
        .Range( _
          .Cells(Table_Row.Field, Table_Column.Left__), _
          .Cells( _
            .Cells(.Rows.Count, Table_Column.Left__).End(xlUp).Row, _
            .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft).Column _
          ) _
        ).RemoveDuplicates _
          Columns:=Array(indexOfFields_), _
          Header:=xlYes
      End If
    End With
    If (IsResultRead) Then
      Call Me.ReadRecords(FilteredRecords)
    End If
  End If
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set FilterRecords = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Read records.
'* @param[out] Records. Set of records.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function ReadRecords( _
  Records As Range, _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "ReadRecords")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  With Me.WorksheetForReadingTable
    Set Records = _
      .Range( _
        .Cells(Table_Row.Field, Table_Column.Left__), _
        .Cells( _
          .Cells(.Rows.Count, Table_Column.Left__).End(xlUp).Row, _
          .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft).Column _
        ) _
      )
  End With
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set ReadRecords = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Update records.
'* @param[in] SetOfFieldsAndValues. Set of fields and values.
'* @param[in] SetOfFiltedFields. (Optional) Set of filtered fields. If this argument is not defined or blank, create a new record based on SetOfFieldsAndValues argument.
'* @param[in] IsResultRead. (Optional) Flag whether updated result is read later on. Default value is False.
'* @param[out] UpdatedRecords. (Optional) Set of updated records.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function UpdateRecords( _
  SetOfFieldsAndValues As String, _
  Optional SetOfFiltedFields As String, _
  Optional IsResultRead As Boolean = False, _
  Optional UpdatedRecords As Range, _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "UpdateRecords")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Dim worksheet_ As Worksheet
  Set worksheet_ = Me.WorksheetForEditingTable
  With worksheet_
    If (SetOfFieldsAndValues <> "") Then
      Dim fieldsAndValuesMap_ As Object
      Set fieldsAndValuesMap_ = CreateObject("Scripting.Dictionary")
      Dim fieldsAndValues_ As Variant
      fieldsAndValues_ = Split(SetOfFieldsAndValues, Second_Level_Delimiter)
      Dim indexOffieldsAndValues_ As Long
      For indexOffieldsAndValues_ = LBound(fieldsAndValues_) To UBound(fieldsAndValues_)
        Dim fieldAndValue_ As Variant
        fieldAndValue_ = Split(fieldsAndValues_(indexOffieldsAndValues_), First_Level_Delimiter)
        Dim updatedField_ As String
        updatedField_ = CStr(fieldAndValue_(Field_And_Value.Field))
        fieldsAndValuesMap_(updatedField_) = _
          CStr(fieldAndValue_(Field_And_Value.Value))
        If (InStr(updatedField_, Field_Date_Delimiter) <> 0) Then
          Dim fields_ As Variant
          fields_ = .Range( _
            .Cells(Table_Row.Field, Table_Column.Left__), _
            .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft) _
          )
          Dim indexOfFields_ As Long
          For indexOfFields_ = LBound(fields_, 2) To UBound(fields_, 2)
            Dim field_ As String
            field_ = fields_(Table_Row.Field, indexOfFields_)
            Dim ToBeInsertedFlag As Boolean
            ToBeInsertedFlag = False
            If (InStr(field_, Field_Date_Delimiter) <> 0) Then
              If (updatedField_ = field_) Then
                ToBeInsertedFlag = False
                Exit For
              ElseIf ( _
                Left( _
                  updatedField_, _
                  InStr(updatedField_, Field_Date_Delimiter) - 1 _
                ) = Left( _
                  field_, _
                  InStr(field_, Field_Date_Delimiter) - 1 _
                ) _
              ) Then
                ToBeInsertedFlag = True
                ' If updated field is lower than current field,
                If (updatedField_ < field_) Then
                  Exit For
                End If
              ElseIf (ToBeInsertedFlag) Then
                Exit For
              End If
            ElseIf (ToBeInsertedFlag) Then
              Exit For
            End If
          Next
          If (ToBeInsertedFlag) Then
            .Columns(indexOfFields_).Insert
            .Cells(Table_Row.Field, indexOfFields_) = CStr(updatedField_)
          End If
        End If
      Next
      If (SetOfFiltedFields <> "") Then
        Call Me.FilterRecords(SetOfFiltedFields)
        Dim rightColumnOfFields_ As Long
        rightColumnOfFields_ = .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft).Column
        Dim lastRowOfRecord_ As Long
        lastRowOfRecord_ = Table_Row.Field
        For indexOfFields_ = Table_Column.Left__ To rightColumnOfFields_
          lastRowOfRecord_ = CLng( _
            Application.WorksheetFunction.Max( _
              .Cells(.Rows.Count, indexOfFields_).End(xlUp).Row, _
              lastRowOfRecord_ _
            ) _
          )
        Next
        If (lastRowOfRecord_ > Table_Row.Field) Then
          fields_ = .Range( _
            .Cells(Table_Row.Field, Table_Column.Left__), _
            .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft) _
          )
          For indexOfFields_ = LBound(fields_, 2) To UBound(fields_, 2)
            field_ = CStr(fields_(Table_Row.Field, indexOfFields_))
            If (fieldsAndValuesMap_.Exists(field_)) Then
              .Range( _
                  .Cells(Table_Row.TopOfRecord, indexOfFields_), _
                  .Cells(lastRowOfRecord_, indexOfFields_) _
              ).SpecialCells(xlCellTypeVisible) = _
                fieldsAndValuesMap_(field_)
            End If
          Next
        End If
        If (IsResultRead) Then
          Dim updatedRecords_ As Range
          Set updatedRecords_ = _
            .Range( _
              .Cells(Table_Row.Field, Table_Column.Left__), _
              .Cells(Table_Row.Field, Table_Column.Left__) _
            ).CurrentRegion.SpecialCells(xlCellTypeVisible)
          With Me.WorksheetForReadingTable
            .Cells.Clear
            Dim indexOfArea_ As Long
            Dim copiedRow_ As Long
            copiedRow_ = Table_Row.Field
            For indexOfArea_ = 1 To updatedRecords_.Areas.Count
              Dim row_ As Long
              row_ = updatedRecords_.Areas(indexOfArea_).Row
              Dim column_ As Long
              column_ = updatedRecords_.Areas(indexOfArea_).Column
              Dim rowsCount_ As Long
              rowsCount_ = updatedRecords_.Areas(indexOfArea_).Rows.Count
              Dim columnsCount_ As Long
              columnsCount_ = updatedRecords_.Areas(indexOfArea_).Columns.Count
              .Range( _
                .Cells(copiedRow_, column_), _
                .Cells(copiedRow_ + rowsCount_ - 1, columnsCount_) _
              ).Value = updatedRecords_.Areas(indexOfArea_).Value
              copiedRow_ = copiedRow_ + rowsCount_
            Next
          End With
          Call Me.ReadRecords(UpdatedRecords)
        End If
      Else
        fields_ = .Range( _
          .Cells(Table_Row.Field, Table_Column.Left__), _
          .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft) _
        )
        Dim insertedRecord_ As Variant
        ReDim insertedRecord_(LBound(fields_, Size_Dimentions.Column) To UBound(fields_, Size_Dimentions.Column)) As Variant
        Dim indexOfFields As Long
        For indexOfFields = LBound(fields_, Size_Dimentions.Column) To UBound(fields_, Size_Dimentions.Column)
          If (fieldsAndValuesMap_.Exists(CStr(fields_(Table_Row.Field, indexOfFields)))) Then
            insertedRecord_(indexOfFields) = fieldsAndValuesMap_(CStr(fields_(Table_Row.Field, indexOfFields)))
          End If
        Next
        Dim rowOfInsertedRecord As Long
        rowOfInsertedRecord = .Cells(.Rows.Count, Table_Column.Left__).End(xlUp).Row + 1
        .Range( _
          .Cells(rowOfInsertedRecord, LBound(fields_, Size_Dimentions.Column)), _
          .Cells(rowOfInsertedRecord, UBound(fields_, Size_Dimentions.Column)) _
        ).Value = insertedRecord_
        If (IsResultRead) Then
          With Me.WorksheetForReadingTable
            .Cells.Clear
            .Range( _
              .Cells(Table_Row.Field, LBound(fields_, Size_Dimentions.Column)), _
              .Cells(Table_Row.Field, UBound(fields_, Size_Dimentions.Column)) _
            ).Value = fields_
            .Range( _
              .Cells(Table_Row.TopOfRecord, LBound(fields_, Size_Dimentions.Column)), _
              .Cells(Table_Row.TopOfRecord, UBound(fields_, Size_Dimentions.Column)) _
            ).Value = insertedRecord_
          End With
          Call Me.ReadRecords(UpdatedRecords)
        End If
      End If
    End If
  End With
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set UpdateRecords = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Delete records.
'* @param[in] SetOfFiltedFields (Optional). Set of filtered fields. If this argument is not defined or blank, create a new record based on SetOfFieldsAndValues argument.
'* @param[in] IsResultRead (Optional). Flag whether deleted result is read later on. Default value is False.
'* @param[out] DeletedRecords. (Optional) Set of deleted records
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function DeleteRecords( _
  SetOfFiltedFields As String, _
  Optional IsResultRead As Boolean = False, _
  Optional DeletedRecords As Range, _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "DeleteRecords")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Call Me.FilterRecords(SetOfFiltedFields)
  With Me.WorksheetForEditingTable
    Dim deletedRecords_ As Range
    Set deletedRecords_ = .Range( _
      .Cells(Table_Row.Field, Table_Column.Left__), _
      .Cells(Table_Row.Field, Table_Column.Left__) _
    ).CurrentRegion.SpecialCells(xlCellTypeVisible)
    If (IsResultRead) Then
      With Me.WorksheetForReadingTable
        .Cells.Clear
        Dim indexOfArea_ As Long
        Dim copiedRow_ As Long
        copiedRow_ = Table_Row.Field
        For indexOfArea_ = 1 To deletedRecords_.Areas.Count
          Dim row_ As Long
          row_ = deletedRecords_.Areas(indexOfArea_).Row
          Dim column_ As Long
          column_ = deletedRecords_.Areas(indexOfArea_).Column
          Dim rowsCount_ As Long
          rowsCount_ = deletedRecords_.Areas(indexOfArea_).Rows.Count
          Dim columnsCount_ As Long
          columnsCount_ = deletedRecords_.Areas(indexOfArea_).Columns.Count
          .Range( _
            .Cells(copiedRow_, column_), _
            .Cells(copiedRow_ + rowsCount_ - 1, columnsCount_) _
          ).Value = deletedRecords_.Areas(indexOfArea_).Value
          copiedRow_ = copiedRow_ + rowsCount_
        Next
        Call Me.ReadRecords(DeletedRecords)
      End With
    End If
    .Rows(CStr(Table_Row.TopOfRecord) & ":" & CStr(.AutoFilter.Range.Rows.Count)).SpecialCells(xlCellTypeVisible).Delete
  End With
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set DeleteRecords = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Execute SQL.
'* @param[in] Sql. SQL statement.
'* @param[out] ExecutedRecords. (Optional) Executed records.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function ExecuteSQL( _
  Sql As String, _
  Optional ExecutedRecords As Range, _
  Optional ExitStatus As Long _
) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "ExecuteSQL")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Select Case Me.DatabaseConnectionType
    Case _
      Database_Connection_Type.Oracle, _
      Database_Connection_Type.MicrosoftAccess
      With Me.WorksheetForEditingTable
        .Cells.Clear
        ' @todo Add argument that identifies read-only-mode and edit-mode.
'        Call myRecordset.Open(Sql, myDatabaseConnection, Cursor_Type.adOpenDynamic, Lock_Type.adLockOptimistic, Command_Type.adCmdText)
        Call myRecordset.Open(Sql, myDatabaseConnection, Cursor_Type.adOpenStatic, Lock_Type.adLockReadOnly, Command_Type.adCmdText)
        Dim indexOfField As Long
        For indexOfField = 1 To myRecordset.Fields.Count
          .Cells(Table_Row.Field, indexOfField).Value = myRecordset.Fields(indexOfField - 1).Name
        Next
        .Cells(Table_Row.TopOfRecord, Table_Column.Left__).CopyFromRecordset myRecordset
        Set ExecutedRecords = .Range( _
          .Cells(Table_Row.Field, Table_Column.Left__), _
          .Cells( _
            .Cells(.Rows.Count, Table_Column.Left__).End(xlUp).Row, _
            .Cells(Table_Row.Field, .Columns.Count).End(xlToLeft).Column _
          ) _
        )
      End With
      myRecordset.Close
    Case Else
      Call Err.Raise(Exit_Status.Database_Connection_Type_Is_Not_Valid, Me, "Database_Connection_Type_Is_Not_Valid")
  End Select
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set ExecuteSQL = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function
