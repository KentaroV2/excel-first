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
'! This provides attributes and methods for Microsoft Excel Worksheet.
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
  Me.Worksheet.Activate
  ' Decorate Worksheet.
  With Me.Worksheet.Cells
    .Clear
    .Font.Name = FontName
    .Font.Size = FontSize
    .Font.Color = FontColor
    .RowHeight = RowHeight
    .ColumnWidth = ColumnWidth
    .HorizontalAlignment = HorizontalAlignment
    .VerticalAlignment = VerticalAlignment
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
  '* @todo do something.
End Function

'* Filter table.
'* @param[in] SetOfFilters. Set of filtered fields.
'* @param[in] SetOfSorts. Set of sorted fields.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function FilterTable(SetOfFiltedFields As String, SetOfSortedFields As String, Optional ExitStatus As Long) As ExcelWorksheetClass
  '* @todo do something.
End Function

'* Read table.
'* @param[out] SetOfRecords. Set of records
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function ReadTable(SetOfRecords As Range, Optional ExitStatus As Long) As ExcelWorksheetClass
  '* @todo do something.
End Function

'* Update table.
'* @param[in] SetOfFilters. Set of filtered fields.
'* @param[in] SetOfItems. Set of items.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function UpdateTable(SetOfFiltedFields As String, SetOfItems As String, Optional ExitStatus As Long) As ExcelWorksheetClass
  '* @todo do something.
End Function

'* Delete table.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorksheetClass. Self.
Public Function DeleteTable(Optional ExitStatus As Long) As ExcelWorksheetClass
  '* @todo do something.
End Function
