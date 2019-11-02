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
Private myLogger As LoggerClass '* Logger object.
Private myExcelWorkbooks As Object '* Set of ExcelWorkbook objects.
Private myScreenUpdatingFlag As Boolean '* Screen-updating flag
Private myEventsFlag As Boolean '* Events flag
Private myAutomaticCalculationFlag As Boolean '* Automatic-calculation flag

'* Initialize this class.
Private Sub Class_Initialize()
  Me.Name = Excel_First_Name
  Set Me.Parent = Nothing
  Set Me.Logger = New LoggerClass
  Set myExcelWorkbooks = CreateObject("Scripting.Dictionary")
  Call Me.BindExcelWorkbook("")
End Sub

'* Terminate this class.
Private Sub Class_Terminate()
  Set Me.Parent = Nothing
  Set Me.Logger = Nothing
  Dim key_ As Variant
  For Each key_ In myExcelWorkbooks
    Set myExcelWorkbooks(key_) = Nothing
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

'* Get Logger property.
Public Property Get Logger() As LoggerClass
  Set Logger = myLogger
End Property

'* Get ExcelWorkbook property.
Public Property Get ExcelWorkbook(ByVal Name As String) As ExcelWorkbookClass
  If (Name = "") Then
    Name = This
  End If
  If (myExcelWorkbooks.Exists(Name) = False) Then
    Set ExcelWorkbook = Nothing
  Else
    Set ExcelWorkbook = myExcelWorkbooks(Name)
  End If
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

'* Disable screen update, auto calculation, and event handling.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function DisableScreenUpdateAutoCalculationAndEventHandling( _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "DisableScreenUpdateAutoCalculationAndEventHandling")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Me.AutomaticCalculationFlag = False
  Me.EventsFlag = False
  Me.ScreenUpdatingFlag = False
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set DisableScreenUpdateAutoCalculationAndEventHandling = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Enable screen update, auto calculation, and event handling.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function EnableScreenUpdateAutoCalculationAndEventHandling( _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "EnableScreenUpdateAutoCalculationAndEventHandling")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Me.AutomaticCalculationFlag = True
  Me.EventsFlag = True
  Me.ScreenUpdatingFlag = True
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set EnableScreenUpdateAutoCalculationAndEventHandling = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Bind ExcelWorkbook.
'* @param[in] Name. When this Name is blank, set Application.ThisWorkbook (running this program) to ExcelWorkbook. When this Name doesn't have any file extention like "xls" or "xlsx", find Excel file with available file extention.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function BindExcelWorkbook( _
  ByVal Name As String, _
  Optional ReadOnly As Boolean = True, _
  Optional Password As String, _
  Optional WriteResPassword As String, _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "BindExcelWorkbook")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Dim excelWorkbook_ As ExcelWorkbookClass
  Set excelWorkbook_ = New ExcelWorkbookClass
  Set excelWorkbook_.Parent = Me
  Set excelWorkbook_.Logger = Me.Logger
  Call excelWorkbook_.Bind(Name, ReadOnly, Password, WriteResPassword, exitStatus_)
  If (exitStatus_ = Exit_Status.Workbook_Is_Not_Found) Then
    Set excelWorkbook_ = Nothing
  Else
    If (Name = "") Then
      Name = This
    End If
    Set myExcelWorkbooks(Name) = excelWorkbook_
  End If
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set BindExcelWorkbook = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Release CPU resource periodically to prevent applications from freezing user operations.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function ReleaseCpuResource( _
  Optional ExitStatus As Long _
) As ExcelFirstClass
  Sleep 0
  DoEvents
  ExitStatus = Exit_Status.Success
  Set ReleaseCpuResource = Me
End Function

'* Chage date to string.
'* @param[in] Date__. date.
'* @param[in] String__. string.
'* @param[in] DegreeDetails. Degree of details.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function ChangeDateToString( _
  date__ As Date, _
  String__ As String, _
  Optional DegreeDetails As String = "D", _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "ChangeDateToString")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  If ( _
    (Len(DegreeDetails) <> 1) _
    Or _
    (InStr("YMDhms", DegreeDetails) = 0) _
  ) Then
    Call Err.Raise(Exit_Status.DegreeDetails_Is_Not_Valid, , "Definition of DegreeDetails is not valid.")
  End If
  String__ = "Y" & Format(date__, "yyyy")
  If (DegreeDetails <> "Y") Then
    String__ = _
      String__ & _
      "M" & Format(date__, "mm")
    If (DegreeDetails <> "M") Then
      String__ = _
        String__ & _
        "D" & Format(date__, "dd")
      If (DegreeDetails <> "D") Then
        String__ = _
          String__ & _
          "h" & Format(date__, "hh")
        If (DegreeDetails <> "h") Then
          String__ = _
            String__ & _
            "m" & Format(date__, "nn")
          If (DegreeDetails <> "m") Then
            String__ = _
              String__ & _
              "s" & Format(date__, "ss")
          End If
        End If
      End If
    End If
  End If
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set ChangeDateToString = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Chage string to date.
'* @param[in] String__. string.
'* @param[in] Date__. date.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function ChangeStringToDate(String__ As String, date__ As Date, Optional ExitStatus As Long) As ExcelFirstClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "ChangeStringToDate")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  Dim now_ As Date
  now_ = Now
  If (InStr(String__, "Y") = 0) Then
    Dim year_ As Long
    year_ = Year(now_)
  Else
    year_ = CLng( _
      Mid( _
        String__, _
        InStr(String__, "Y") + 1, _
        4 _
      ) _
    )
  End If
  If (InStr(String__, "M") = 0) Then
    Dim month_ As Long
    month_ = 1
  Else
    month_ = CLng( _
      Mid( _
        String__, _
        InStr(String__, "M") + 1, _
        2 _
      ) _
    )
  End If
  If (InStr(String__, "D") = 0) Then
    Dim day_ As Long
    day_ = 1
  Else
    day_ = CLng( _
      Mid( _
        String__, _
        InStr(String__, "D") + 1, _
        2 _
      ) _
    )
  End If
  If (InStr(String__, "h") = 0) Then
    Dim hour_ As Long
    hour_ = 0
  Else
    hour_ = CLng( _
      Mid( _
        String__, _
        InStr(String__, "h") + 1, _
        2 _
      ) _
    )
  End If
  If (InStr(String__, "m") = 0) Then
    Dim minute_ As Long
    minute_ = 0
  Else
    minute_ = CLng( _
      Mid( _
        String__, _
        InStr(String__, "m") + 1, _
        2 _
      ) _
    )
  End If
  If (InStr(String__, "s") = 0) Then
    Dim second_ As Long
    second_ = 0
  Else
    second_ = CLng( _
      Mid( _
        String__, _
        InStr(String__, "s") + 1, _
        2 _
      ) _
    )
  End If
  date__ = CDate(year_ & "/" & month_ & "/" & day_ & " " & hour_ & ":" & minute_ & ":" & second_)
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set ChangeStringToDate = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Replace words with new words.
'* @param[in] ToBeReplacedString. To be replaced string.
'* @param[in] WordsReplacingMap. Words replacing map.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function ReplaceWordsWithNewWords( _
  ToBeReplacedString As String, _
  WordsReplacingMap As String, _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "ReplaceWordsWithNewWords")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  If (WordsReplacingMap = "") Then
    Call Err.Raise(Exit_Status.WordsReplacingMap_Is_Not_Defined, Me, "WordsReplacingMap_Is_Not_Defined")
  End If
  Dim wordsReplacingMap_ As Variant
  wordsReplacingMap_ = Split(WordsReplacingMap, Second_Level_Delimiter)
  Dim indexForWordsReplacingMap_ As Long
  For indexForWordsReplacingMap_ = LBound(wordsReplacingMap_) To UBound(wordsReplacingMap_)
    Dim keyAndValue_ As Variant
    keyAndValue_ = Split(wordsReplacingMap_(indexForWordsReplacingMap_), First_Level_Delimiter)
    ToBeReplacedString = Replace( _
      ToBeReplacedString, _
      keyAndValue_(LBound(keyAndValue_)), _
      keyAndValue_(UBound(keyAndValue_)) _
    )
  Next
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set ReplaceWordsWithNewWords = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Append key and value.
'* @param[in,out] KeysAndValues. Keys and values.
'* @param[in] Key. Key
'* @param[in] Value. Value.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function AppendKeyAndValue( _
  KeysAndValues As String, _
  Key As String, _
  Value As String, _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "AppendKeyAndValue")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  If (KeysAndValues <> "") Then
    KeysAndValues = KeysAndValues & Second_Level_Delimiter
  End If
  KeysAndValues = KeysAndValues & Key & First_Level_Delimiter & Value
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set AppendKeyAndValue = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

