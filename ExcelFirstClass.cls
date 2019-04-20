VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelFirstClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'! @file ExcelFirstClass.cls
'! This class provides a set of VBA (Visual Basic for Applications) libraries running on Microsoft Excel for
'! (a) taking advantage of original Excel functionalities and for
'! (b) smoothly deploying your applications to your business since most of companies use Microsoft Excel compliant with companies IT policies.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------
' Define members.
Private myName As String '* Name. This name is also used for logging purpose.
Private myParent As Object '* Parent object.
Private myExcelWorkbooks As Object ' Set of ExcelWorkbook objects.
Private myLogger As LoggerClass '* Logger object.
Private myScreenUpdatingFlag As Boolean '* Screen-updating flag
Private myEventsFlag As Boolean '* Events flag
Private myAutomaticCalculationFlag As Boolean '* Automatic-calculation flag

'* Initialize this class.
Private Sub Class_Initialize()
  myName = Excel_First_Name
  Set myLogger = New LoggerClass
  Set myExcelWorkbooks = CreateObject("Scripting.Dictionary")
  '* @todo Set thisWorkbook object here.
End Sub

'* Terminate this class.
Private Sub Class_Terminate()
  Set myParent = Nothing
  Dim key_ As Variant
  For Each key_ In myExcelWorkbooks
    Set myExcelWorkbooks(key_) = Nothing
  Next
  Set myLogger = Nothing
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

'* Get logger property.
Property Get Logger() As LoggerClass
    Set Logger = myLogger
End Property

'* Get screen updating flag.
Property Get ScreenUpdatingFlag() As Boolean
    ScreenUpdatingFlag = myScreenUpdatingFlag
End Property

'* Set screen updating flag.
Property Let ScreenUpdatingFlag(ScreenUpdatingFlag As Boolean)
  myScreenUpdatingFlag = ScreenUpdatingFlag
  Application.ScreenUpdating = ScreenUpdatingFlag
End Property

'* Get event flag.
Property Get EventsFlag() As Boolean
    EventsFlag = myEventsFlag
End Property

'* Set event flag.
Property Let EventsFlag(EventsFlag As Boolean)
  myEventsFlag = EventsFlag
  Application.EnableEvents = EventsFlag
End Property

'* Get automatic calculation flag.
Property Get AutomaticCalculationFlag() As Boolean
  AutomaticCalculationFlag = myAutomaticCalculationFlag
End Property

'* Set automatic calculation flag.
Property Let AutomaticCalculationFlag(AutomaticCalculationFlag As Boolean)
  myAutomaticCalculationFlag = AutomaticCalculationFlag
  If (AutomaticCalculationFlag = True) Then
    Dim calculationFlag_ As Long
    calculationFlag_ = xlCalculationManual
  Else
    calculationFlag_ = xlCalculationAutomatic
  End If
  Application.Calculation = calculationFlag_
End Property

'* Bind ExcelWorkbook.
'* @param[in] Name. When this Name is blank, return a ThisWorkbook object (running this program). When this Name doesn't have any file extention like "xls" or "xlsx", find Excel file with available file extention.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelWorkbookClass
Public Function BindExcelWorkbook(Name As String, Optional ExitStatus As Long) As ExcelWorkbookClass
Initialize:
  On Error GoTo Error
  myLogger.StackName (myName & "." & "BindExcelWorkbook")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
  Set BindExcelWorkbook = Nothing
Execute:
  Dim excelWorkbook_ As ExcelWorkbookClass
  Set excelWorkbook_ = New ExcelWorkbookClass
  Set excelWorkbook_.Parent = Me
  Set excelWorkbook_.Logger = Me.Logger '* @todo Eliminate relations between Logger object and any other object.
  Call excelWorkbook_.Bind(Name, exitStatus_)
  If (exitStatus_ = Exit_Status.Workbook_Is_Not_Found) Then
    Set excelWorkbook_ = Nothing
  End If
  Set BindExcelWorkbook = excelWorkbook_
Exit__:
  myLogger.UnstackName
  ExitStatus = exitStatus_
  Exit Function
Error:
  exitStatus_ = Err.Number
  myLogger.Error ("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Release CPU resource periodically to prevent applications from freezing user operations.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
Public Sub ReleaseCpuResource(Optional ExitStatus As Long)
  Sleep 0
  DoEvents
  ExitStatus = Exit_Status.Success
End Sub
