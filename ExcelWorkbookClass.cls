VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelWorkbookClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'! This provides attributes and methods for Microsoft Excel Workbook.
'! This class contains (a) encryption attributes (b) Workbook object, and (c) data manipulating functions.
'! This class is created by the ExcelFirst class in order to bind a Workbook where the ExcelFirst class runs when the ExcelFirst class runs is initialized.
'! @copyright MIT

' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------
' Define members.
Private myName As String '* Name. This name represents file name except directory path and file extention; (ie. xls, xlsx). This name is also used for logging purpose.
Private myParent As Object '* Parent object.
Private myLogger As LoggerClass '* Logger object
Private myWorkbook As Workbook '* Workbook object
Private myPasswordFlag As Boolean '* Password flag
Private myExcelWorksheets As Object '* Set of ExcelWorksheet objects

'* Initialize this class.
Private Sub Class_Initialize()
  myName = Undefined
  myPasswordFlag = False
  Set myExcelWorksheets = CreateObject("Scripting.Dictionary")
End Sub

'* Terminate this class.
Private Sub Class_Terminate()
  ' Release all keys in myWorksheets object.
  Dim key_ As Variant
  For Each key_ In myExcelWorksheets
    Set myExcelWorksheets(key_) = Nothing
  Next
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
Property Set Workbook(Workbook__ As Workbook)
  Set myWorkbook = Workbook__
End Property

'* Get Workbook property.
Property Get Workbook() As Workbook
  Set Workbook = myWorkbook
End Property

'* Bind Workbook.
'* @param[in] Name. When this Name is blank, return a ThisWorkbook object (running this program). When this Name doesn't have any file extention like "xls" or "xlsx", find Excel file with these file extentions.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub Bind(Name As String, Optional ExitStatus As Long)
Initialize:
  On Error GoTo Error
  Me.Logger.StackName (myName & "." & "Bind")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  ' If the Name is blank,
  If (Name = "") Then
    ' Return a ThisWorkbook object.
    Set Me.Workbook = Application.ThisWorkbook
    myName = This
  Else
    '* @todo Do something.
    exitStatus_ = Exit_Status.Workbook_Is_Not_Found
  End If
Exit__:
  Me.Logger.UnstackName
  ExitStatus = exitStatus_
  Exit Sub
Error:
  ' Set return value as error number.
  exitStatus_ = Err.Number
  ' Log error.
  Me.Logger.Error ("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  ' Go to exit function.
  GoTo Exit__
End Sub

'* Bind ExcelWorksheet.
'* @param[in] Name. This name should be a Worksheet name.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorkbookClass
Public Function BindExcelWorksheet(Name As String, Optional ExitStatus As Long) As ExcelWorksheetClass
Initialize:
  On Error GoTo Error
  myLogger.StackName (myName & "." & "BindExcelWorksheet")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
  Set BindExcelWorksheet = Nothing
Execute:
  Dim excelWorksheet_ As ExcelWorksheetClass
  Set excelWorksheet_ = New ExcelWorksheetClass
  Set excelWorksheet_.Parent = Me
  Set excelWorksheet_.Logger = Me.Logger '* @todo Eliminate relations between Logger object and any other object.
  Call excelWorksheet_.Bind(Name, exitStatus_)
  If (exitStatus_ = Exit_Status.Worksheet_Is_Not_Found) Then
    Set excelWorksheet_ = Nothing
  End If
  Set BindExcelWorksheet = excelWorksheet_
Exit__:
  myLogger.UnstackName
  ExitStatus = exitStatus_
  Exit Function
Error:
  exitStatus_ = Err.Number
  myLogger.Error ("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

