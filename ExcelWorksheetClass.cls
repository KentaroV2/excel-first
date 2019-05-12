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
Private myLogger As Object '* Logger object
Private myWorksheet As Worksheet '* Worksheet object.
Private myWorksheetForEditingTable As Worksheet '* Worksheet object for editing table.
Private myWorksheetForReadingTable As Worksheet '* Worksheet object for reading table.

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
Public Property Get Logger() As Object
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
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function BindWorksheet(Name As String, Optional ExitStatus As Long) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Call Me.Logger.StackName(Me.Name & "." & "BindWorksheet")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Dim parent_ As Object
  Set parent_ = Me.Parent
  Dim workbook_ As Workbook
  Set workbook_ = parent_.Workbook
  Dim worksheets_ As Sheets
  Set worksheets_ = workbook_.Worksheets
  Dim worksheet_ As Worksheet
  For Each worksheet_ In worksheets_
      If (worksheet_.Name = Name) Then
        Set Me.Worksheet = worksheet_
      End If
  Next
  If (Me.Worksheet Is Nothing) Then
    exitStatus_ = Exit_Status.Worksheet_Is_Not_Found
  Else
    Set Me.WorksheetForEditingTable = Nothing
    Set Me.WorksheetForReadingTable = Nothing
    Me.Name = Name
  End If
Exit__:
  Call Me.Logger.UnstackName
  ExitStatus = exitStatus_
  Set BindWorksheet = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call Me.Logger.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
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
  myLogger.StackName (myName & "." & "Clear")
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
    With .Cells
      .Clear
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
  Call Me.Logger.UnstackName
  ExitStatus = exitStatus_
  Exit Sub
Error:
  exitStatus_ = Err.Number
  Call Me.Logger.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Sub

'* Bind Worksheet for table.
'* @param[in] Name. This name should be Worksheet name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function BindWorksheetForTable(Name As String, Optional ExitStatus As Long) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Call Me.Logger.StackName(Me.Name & "." & "BindWorksheetForTable")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Dim parent_ As Object
  Set parent_ = Me.Parent
  Dim workbook_ As Workbook
  Set workbook_ = parent_.Workbook
  Dim worksheets_ As Sheets
  Set worksheets_ = workbook_.Worksheets
  Dim worksheet_ As Worksheet
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
  If ( _
    Me.WorksheetForEditingTable Is Nothing _
    Or _
    Me.WorksheetForReadingTable Is Nothing _
  ) Then
    exitStatus_ = Exit_Status.Worksheet_Is_Not_Found
  Else
    Set Me.Worksheet = Nothing
    Me.Name = Name
  End If
Exit__:
  Call Me.Logger.UnstackName
  ExitStatus = exitStatus_
  Set BindWorksheetForTable = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call Me.Logger.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Create table.
'* @param[in] SetOfFields. Set of fields.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function CreateTable(SetOfFields As String, Optional ExitStatus As Long) As ExcelWorksheetClass
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
        .Cells(WorksheetRow.Top, WorksheetColumn.Left__), _
        .Cells(WorksheetRow.Top, WorksheetColumn.Left__ + UBound(Fields) - LBound(Fields)) _
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
'* @param[out] FilteredRecords. (Optional) Set of filtered records.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function FilterRecords(Optional SetOfFiltedFields As String, Optional SetOfSortedFields As String, Optional DuplicatesRemovingField As String, Optional FilteredRecords As Range, Optional ExitStatus As Long) As ExcelWorksheetClass
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
      .Cells(TableRow.Field, TableColumn.Left__), _
      .Cells(TableRow.Field, .Columns.Count).End(xlToLeft) _
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
          If (filteredField_(FieldAndValue.Field) = fields_(TableRow.Field, indexOfFields_)) Then
            Exit For
          End If
        Next
        If ( _
          indexOfFields_ >= LBound(fields_) _
          And _
          indexOfFields_ <= UBound(fields_) _
        ) Then
          If (InStr(filteredField_(FieldAndValue.Value), Or_Operator) = 0) Then
            Dim filters_ As Variant
            filters_ = Array(filteredField_(FieldAndValue.Value))
          Else
            filters_ = Split(filteredField_(FieldAndValue.Value), Or_Operator)
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
              If (sortedField_(FieldAndValue.Field) = fields_(TableRow.Field, indexOfFields_)) Then
                Exit For
              End If
            Next
            If ( _
              indexOfFields_ >= LBound(fields_) _
              And _
              indexOfFields_ <= UBound(fields_) _
            ) Then
            .Add _
              Key:=worksheet_.Range( _
                worksheet_.Columns(indexOfFields_), _
                worksheet_.Columns(indexOfFields_) _
              ), _
              SortOn:=xlSortOnValues, _
              Order:=CLng(sortedField_(FieldAndValue.Value)), _
              DataOption:=xlSortNormal
            End If
          Next
        End If
      End With
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
    End With
    Dim filteredAndSortedRecords As Range
    Set filteredAndSortedRecords = .Range( _
      .Cells(TableRow.Field, TableColumn.Left__), _
      .Cells( _
        .Cells(.Rows.Count, TableColumn.Left__).End(xlUp).Row, _
        .Cells(TableRow.Field, .Columns.Count).End(xlToLeft).Column _
      ) _
    ).SpecialCells(xlCellTypeVisible)
  End With
  With Me.WorksheetForReadingTable
    .Cells.Clear
    filteredAndSortedRecords.Copy .Range( _
      .Cells(TableRow.Field, TableColumn.Left__), _
      .Cells(TableRow.Field, TableColumn.Left__) _
    )
    If (DuplicatesRemovingField <> "") Then
      For indexOfFields_ = LBound(fields_, 2) To UBound(fields_, 2)
        If (DuplicatesRemovingField = fields_(TableRow.Field, indexOfFields_)) Then
          Exit For
        End If
      Next
      .Range( _
        .Cells(TableRow.Field, TableColumn.Left__), _
        .Cells( _
          .Cells(.Rows.Count, TableColumn.Left__).End(xlUp).Row, _
          .Cells(TableRow.Field, .Columns.Count).End(xlToLeft).Column _
        ) _
      ).RemoveDuplicates _
        Columns:=Array(indexOfFields_), _
        Header:=xlYes
'* @todo Fix a bug that defining two or more fields doesn't work properly (by removing duplicates field by field.).
'    If (SetOfRemovingDuplicates <> "") Then
'      Dim duplicatesRemovingFields_ As Variant
'      duplicatesRemovingFields_ = Split(SetOfRemovingDuplicates, First_Level_Delimiter)
'      Dim duplicatesRemovingColumns_() As Long
'      ReDim duplicatesRemovingColumns_(0)
'      Dim indexOfDuplicatesRemovingFields_ As Long
'      For indexOfDuplicatesRemovingFields_ = LBound(duplicatesRemovingFields_) To UBound(duplicatesRemovingFields_)
'        For indexOfFields_ = LBound(fields_, 2) To UBound(fields_, 2)
'          If (duplicatesRemovingFields_(indexOfDuplicatesRemovingFields_) = fields_(TableRow.Field, indexOfFields_)) Then
'            duplicatesRemovingColumns_(UBound(duplicatesRemovingColumns_)) = indexOfFields_
'            ReDim Preserve duplicatesRemovingColumns_(UBound(duplicatesRemovingColumns_) + 1)
'            Exit For
'          End If
'        Next
'      Next
'      ReDim Preserve duplicatesRemovingColumns_(UBound(duplicatesRemovingColumns_) - 1)
'      .Range( _
'        .Cells(TableRow.Field, TableColumn.Left), _
'        .Cells( _
'          .Cells(.Rows.Count, TableColumn.Left).End(xlUp).Row, _
'          .Cells(TableRow.Field, .Columns.Count).End(xlToLeft).Column _
'        ) _
'      ).RemoveDuplicates _
'        Columns:=.Evaluate(duplicatesRemovingColumns_), _
'        Header:=xlYes
    End If
  End With
  Call Me.ReadRecords(FilteredRecords)
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
Public Function ReadRecords(Records As Range, Optional ExitStatus As Long) As ExcelWorksheetClass
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
        .Cells(TableRow.Field, TableColumn.Left__), _
        .Cells( _
          .Cells(.Rows.Count, TableColumn.Left__).End(xlUp).Row, _
          .Cells(TableRow.Field, .Columns.Count).End(xlToLeft).Column _
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
'* @param[out] UpdatedRecords. (Optional) Set of updated records.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function UpdateRecords(SetOfFieldsAndValues As String, Optional SetOfFiltedFields As String, Optional UpdatedRecords As Range, Optional ExitStatus As Long) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "UpdateRecords")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  ' Filter records.
  Call Me.FilterRecords(SetOfFiltedFields)
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
        updatedField_ = CStr(fieldAndValue_(FieldAndValue.Field))
        fieldsAndValuesMap_(updatedField_) = _
          CStr(fieldAndValue_(FieldAndValue.Value))
        If (InStr(updatedField_, Field_Date_Delimiter) <> 0) Then
          Dim fields_ As Variant
          fields_ = .Range( _
            .Cells(TableRow.Field, TableColumn.Left__), _
            .Cells(TableRow.Field, .Columns.Count).End(xlToLeft) _
          )
          Dim indexOfFields_ As Long
          For indexOfFields_ = LBound(fields_, 2) To UBound(fields_, 2)
            Dim field_ As String
            field_ = fields_(TableRow.Field, indexOfFields_)
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
            .Cells(TableRow.Field, indexOfFields_) = CStr(updatedField_)
          End If
        End If
      Next
      If (SetOfFiltedFields <> "") Then
        Dim updatedRecords_ As Range
        Set updatedRecords_ = _
          .Range( _
            .Cells(TableRow.Field, TableColumn.Left__), _
            .Cells(TableRow.Field, TableColumn.Left__) _
          ).CurrentRegion.SpecialCells(xlCellTypeVisible)
        Dim updatedCell_ As Range
        For Each updatedCell_ In updatedRecords_
          If ( _
            (updatedCell_.Row >= TableRow.TopOfRecord) _
            And _
            (fieldsAndValuesMap_.Exists(CStr(updatedRecords_(TableRow.Field, updatedCell_.Column)))) _
          ) Then
            updatedCell_ = CStr(fieldsAndValuesMap_(CStr(updatedRecords_(TableRow.Field, updatedCell_.Column))))
          End If
        Next
      Else
        .Rows(TableRow.TopOfRecord).Insert
        fields_ = .Range( _
          .Cells(TableRow.Field, TableColumn.Left__), _
          .Cells(TableRow.Field, .Columns.Count).End(xlToLeft) _
        )
        Dim insertedRecord_ As Variant
        ReDim insertedRecord_(LBound(fields_, 2) To UBound(fields_, 2)) As String
        Dim indexOfFields As Long
        For indexOfFields = LBound(fields_, 2) To UBound(fields_, 2)
          If (fieldsAndValuesMap_.Exists(CStr(fields_(TableRow.Field, indexOfFields)))) Then
            insertedRecord_(indexOfFields) = CStr(fieldsAndValuesMap_(CStr(fields_(TableRow.Field, indexOfFields))))
          End If
        Next
        .Range( _
          .Cells(TableRow.TopOfRecord, LBound(fields_, 2)), _
          .Cells(TableRow.TopOfRecord, UBound(fields_, 2)) _
        ) = insertedRecord_
        Set updatedRecords_ = _
          .Range( _
            .Cells(TableRow.Field, LBound(fields_, 2)), _
            .Cells(TableRow.TopOfRecord, UBound(fields_, 2)) _
          )
      End If
    End If
  End With
  With Me.WorksheetForReadingTable
    .Cells.Clear
    updatedRecords_.Copy .Range( _
      .Cells(TableRow.Field, TableColumn.Left__), _
      .Cells(TableRow.Field, TableColumn.Left__) _
    )
  End With
  Call Me.ReadRecords(UpdatedRecords)
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
'* @param[in] SetOfFiltedFields. (Optional) Set of filtered fields. If this argument is not defined or blank, create a new record based on SetOfFieldsAndValues argument.
'* @param[out] DeletedRecords. Set of deleted records
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function DeleteRecords(SetOfFiltedFields As String, Optional DeletedRecords As Range, Optional ExitStatus As Long) As ExcelWorksheetClass
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
      .Cells(TableRow.Field, TableColumn.Left__), _
      .Cells(TableRow.Field, TableColumn.Left__) _
    ).CurrentRegion.SpecialCells(xlCellTypeVisible)
    With Me.WorksheetForReadingTable
      .Cells.Clear
      deletedRecords_.Copy .Range( _
        .Cells(TableRow.Field, TableColumn.Left__), _
        .Cells(TableRow.Field, TableColumn.Left__) _
      )
    End With
    deletedRecords_.Delete
  End With
  Call Me.ReadRecords(DeletedRecords)
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
