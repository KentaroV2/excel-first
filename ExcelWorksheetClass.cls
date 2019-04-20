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
Private myWorksheet As Worksheet '* Worksheet object.
Private myLogger As LoggerClass '* Logger object

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

'* Get Parent property.
Public Property Get Logger() As Object
  Set Logger = myLogger
End Property

'* Set Workbook property.
Property Set Worksheet(Worksheet__ As Worksheet)
  Set myWorksheet = Worksheet__
End Property

'* Get Workbook property.
Property Get Worksheet() As Worksheet
  Set Worksheet = myWorksheet
End Property

'* Bind Worksheet.
'* @param[in] Name. This name should be Worksheet name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Bind(Name As String, Optional ExitStatus As Long)
Initialize:
  On Error GoTo Error
  myLogger.StackName (myName & "." & "Bind")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Dim parent_ As Object
  Set parent_ = Me.Parent
  Dim workbook_ As Workbook
  Set workbook_ = parent_.Workbook
  Dim worksheets_ As Sheets
  Set worksheets_ = workbook_.Worksheets
  Dim worksheet_ As Worksheet, foundFlag_ As Boolean
  For Each worksheet_ In worksheets_
      If (worksheet_.Name = Name) Then
        foundFlag_ = True
        Dim foundWorksheet_ As Worksheet
        Set foundWorksheet_ = worksheet_
      End If
  Next
  If foundFlag_ = True Then
    Me.Logger.Info (foundWorksheet_.Name)
    Set Me.Worksheet = workbook_.Worksheets(foundWorksheet_.Name)
  Else
    exitStatus_ = Exit_Status.Worksheet_Is_Not_Found
  End If
Exit__:
  myLogger.UnstackName
  ExitStatus = exitStatus_
  Exit Sub
Error:
  ' Set return value as error number.
  exitStatus_ = Err.Number
  ' Log error.
  myLogger.Error ("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  ' Go to exit function.
  GoTo Exit__
End Sub

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
  Optional ExitStatus As Long)
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
  myLogger.UnstackName
  ExitStatus = exitStatus_
  Exit Sub
Error:
  ' Set return value as error number.
  exitStatus_ = Err.Number
  ' Log error.
  myLogger.Error ("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  ' Go to exit function.
  GoTo Exit__
End Sub
