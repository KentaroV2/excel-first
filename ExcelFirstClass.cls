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
'! This class provides a set of VBA (Visual Basic for Applications) libraries running on Microsoft Excel for:
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
Private myScreenUpdatingFlag As Boolean '* Screen-updating flag.
Private myEventsFlag As Boolean '* Events flag.
Private myAutomaticCalculationFlag As Boolean '* Automatic-calculation flag.
Private myFileSystem As Object '* File system object.

'* Initialize object.
Private Sub Class_Initialize()
  ' Set Name.
  Me.Name = Excel_First_Name
  ' Initialize Parent object.
  Set Me.Parent = Nothing
  ' Instantiate Logger object.
  Set Me.Logger = New LoggerClass
  ' Instantiate myExcelWorkbooks object.
  Set myExcelWorkbooks = CreateObject("Scripting.Dictionary")
  ' Bind ExcelWorkbook object where ExcelFirst object runs to this object.
  Call Me.BindExcelWorkbook(This)
  ' Instantiate file system object.
  Set myFileSystem = CreateObject("Scripting.FileSystemObject")
End Sub

'* Terminate object.
Private Sub Class_Terminate()
  ' Release ExcelWorkbooks objects.
  Dim key_ As Variant
  For Each key_ In myExcelWorkbooks
    Set myExcelWorkbooks(key_) = Nothing
  Next
  Set myExcelWorkbooks = Nothing
  ' Release Logger objects.
  Set Me.Logger = Nothing
  ' Release Parent objects.
  Set Me.Parent = Nothing
  ' Release file system objects.
  Set myFileSystem = Nothing
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
  ' If Name is blank,
  If (Name = "") Then
    ' Set Name to "This".
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

'* Get FileSystem property.
Public Property Get fileSystem() As Object
  Set fileSystem = myFileSystem
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
'* @param[in] DegreeDetails. Degree of details. Value should one of "Y" as year, "M" as month, "D" as date, "h" as hour, "m" as minute, and "s" as second.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function ChangeDateToString( _
  Date__ As Date, _
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
    Call Err.Raise(Exit_Status.DegreeDetails_Is_Not_Valid, Me, "Definition of DegreeDetails is not valid.")
  End If
  String__ = Format(Date__, "yyyy")
  If (DegreeDetails <> "Y") Then
    String__ = _
      String__ & _
      "-" & Format(Date__, "mm")
    If (DegreeDetails <> "M") Then
      String__ = _
        String__ & _
        "-" & Format(Date__, "dd")
      If (DegreeDetails <> "D") Then
        String__ = _
          String__ & _
          " " & Format(Date__, "hh")
        If (DegreeDetails <> "h") Then
          String__ = _
            String__ & _
            ":" & Format(Date__, "nn")
          If (DegreeDetails <> "m") Then
            String__ = _
              String__ & _
              ":" & Format(Date__, "ss")
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
Public Function ChangeStringToDate(String__ As String, Date__ As Date, Optional ExitStatus As Long) As ExcelFirstClass
Initialize:
  On Error GoTo Error
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  Call logger_.StackName(Me.Name & "." & "ChangeStringToDate")
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  If (InStr("String__", "'")) Then
    String__ = Replace(String__, "'", "")
  End If
  If ( _
    IsDate(String__) _
    And _
    String__ <> "" _
  ) Then
    Date__ = CDate(String__)
  Else
    Call Err.Raise(Exit_Status.String_For_Date_Is_Not_Defined, Me, "String_For_Date_Is_Not_Defined.")
'    Call Err.Raise(Exit_Status.String_For_Date_Is_Not_Defined, Me, "String_For_Date_Is_Not_Defined. String__ : '" & String__ & "'")
  End If
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

'* Export modules under specified directory path.
'* Exported modules are Standard Module(.bas), Class Module(.cls), and Form Module(.frm).
'* These modules don't include ActiveX Designer Module(.cls) nor Document Module(.cls).
'*
'* @param[in] DirectoryPath. Directory path where modules are exporeted.
'* @param[in] Filters (Optional). ie. "Examples*<first delimiter>/Examples/", "Ignore*<first delimiter>", "Special*<first delimiter>/Special/"
'* @param[out] ExitStatus (Optional). Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function ExportModules( _
  DirectoryPath As String, _
  Optional Filters As String, _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  ' Define error handling.
  On Error GoTo Error
  ' Set logger_.
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  ' Stack name for logging purpose.
  Call logger_.StackName(Me.Name & "." & "ExportModules")
  ' Set exitStatus_.
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  ' Set workbook_ to Workbook object where ExcelFirst is running on.
  Dim workbook_ As Workbook
  Set workbook_ = Me.ExcelWorkbook(This).Workbook
  ' Set filters_ splitted by given Filters.
  Dim filters_ As Variant
  filters_ = Split(Filters, Second_Level_Delimiter)
  ' Set moduleList_.
  Dim moduleList_ As Object
  Set moduleList_ = workbook_.VBProject.vbComponents
  ' Loop exporting modules.
  Dim module_ As Object
  For Each module_ In moduleList_
    ' Initialize extention_.
    Dim extention_  As String
    extention_ = ""
    ' Process depending on module type.
    Select Case module_.Type
      ' If vbext_ct_StdModule,
      Case vbext_ComponentType.vbext_ct_StdModule
        extention_ = "bas"
      ' If vbext_ct_ClassModule,
      Case vbext_ComponentType.vbext_ct_ClassModule
        extention_ = "cls"
      ' If vbext_ct_MSForm,
      Case vbext_ComponentType.vbext_ct_MSForm
        extention_ = "frm"
      ' If vbext_ct_ActiveXDesigner,
'      Case vbext_ComponentType.vbext_ct_ActiveXDesigner
'        extention_ = "cls"
'      ' If vbext_ct_Document,
'      Case vbext_ComponentType.vbext_ct_Document
'        extention_ = "cls"
    End Select
    ' If extention is defined,
    If (extention_ <> "") Then
      ' Set file name.
      Dim fileName_ As String
      fileName_ = module_.Name & Dot & extention_
      ' Initialize ignored_ flag.
      Dim ignored_ As Boolean
      ignored_ = False
      ' Initialize subDirectory_.
      Dim subDirectory_ As String
      subDirectory_ = ""
      ' Filter file.
      Dim filter_ As Variant
      For Each filter_ In filters_
        Dim keyValue_ As Variant
        keyValue_ = Split(filter_, First_Level_Delimiter)
        ' Process depending on filter's key.
        ' Set key.
        Dim key_ As String
        key_ = keyValue_(Key_Value.Key)
        ' Set value.
        Dim value_ As String
        value_ = keyValue_(Key_Value.Value)
        keyValue_ = Empty ' Released because or no more use.
        ' If file name matches filter,
        If (fileName_ Like key_) Then
          ' If value_ is NOT defined,
          If (value_ = "") Then
            ' Set ignored_ to TRUE.
            ignored_ = True
            ' Exit for-next loop.
            Exit For
          ' If value_ is defined,
          Else
            ' If value_ match directory,
            If (value_ Like "\*\") Then
              ' Set ignored_ to FALSE.
              ignored_ = False
              ' Remove first directory path.
              subDirectory_ = Right(value_, Len(value_) - 1)
            End If
          End If
        End If
      Next
      filter_ = Empty ' Released because or no more use.
      If Not (ignored_) Then
        ' Export module.
        Dim filePath_ As String
        filePath_ = DirectoryPath & Path_Separator & subDirectory_ & fileName_
        Call module_.Export(filePath_)
        logger_.Info ("Exported file is '" & filePath_ & "'.") ' Log information.
      End If
    End If
  Next
  filters_ = Empty ' Released because or no more use.
  Set module_ = Nothing ' Released because or no more use.
  Set moduleList_ = Nothing ' Released because or no more use.
Exit__:
  ' Release variants.
  filters_ = Empty
  filter_ = Empty
  keyValue_ = Empty
  ' Release objects.
  Set workbook_ = Nothing
  Set moduleList_ = Nothing
  Set module_ = Nothing
  ' Unstak name for logging purpose and release logger_.
  Call logger_.UnstackName
  Set logger_ = Nothing
  ' Set ExitStatus.
  ExitStatus = exitStatus_
  ' Return self.
  Set ExportModules = Me
  ' Exit function
  Exit Function
Error:
  ' Set exitStatus_ to error number.
  exitStatus_ = Err.Number
  ' Log error.
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  ' Go to exit.
  GoTo Exit__
End Function

'* Import modules under specified directory path.
'*
'* @param[in] DirectoryPath. Directory path where modules are imported.
'* @param[in] Filters (Optional). ie. "Examples*<first delimiter>/Examples/", "Ignore*<first delimiter>", "Special*<first delimiter>/Special/"
'* @param[out] ExitStatus (Optional). Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function ImportModules( _
  DirectoryPath As String, _
  Optional Filters As String, _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  ' Define error handling.
  On Error GoTo Error
  ' Set logger_.
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  ' Stack name for logging purpose.
  Call logger_.StackName(Me.Name & "." & "ImportModules")
  ' Set exitStatus_.
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  
  '* todo do something.
  
'  ' Set workbook_ to Workbook object where ExcelFirst is running on.
'  Dim workbook_ As Workbook
'  Set workbook_ = Me.ExcelWorkbook(This).Workbook
'  ' Set filters_ splitted by given Filters.
'  Dim filters_ As Variant
'  filters_ = Split(Filters, Second_Level_Delimiter)
'  ' Set vbComponents_.
'  Dim vbComponents_ As Object
'  Set vbComponents_ = workbook_.VBProject.vbComponents
'  ' Loop importing modules.
'  Dim filter_ As Variant
'  For Each filter_ In filters_
'    Dim keyValue_ As Variant
'    keyValue_ = Split(filter_, First_Level_Delimiter)
'    ' Process depending on filter's key.
'    ' Set key.
'    Dim key_ As String
'    key_ = keyValue_(Key_Value.Key)
'    ' Set value.
'    Dim value_ As String
'    value_ = keyValue_(Key_Value.Value)
'    ' If value_ specifiy sub directory,
'    If (value_ Like "\*\") Then
'
'    End If
'
'  Next
'    ' Initialize extention_.
'    Dim extention_  As String
'    extention_ = ""
'    ' Process depending on module type.
'    Select Case module_.Type
'      ' If vbext_ct_StdModule,
'      Case vbext_ComponentType.vbext_ct_StdModule
'        extention_ = "bas"
'      ' If vbext_ct_ClassModule,
'      Case vbext_ComponentType.vbext_ct_ClassModule
'        extention_ = "cls"
'      ' If vbext_ct_MSForm,
'      Case vbext_ComponentType.vbext_ct_MSForm
'        extention_ = "frm"
'      ' If vbext_ct_ActiveXDesigner,
'      Case vbext_ComponentType.vbext_ct_ActiveXDesigner
'        extention_ = "cls"
'      ' If vbext_ct_Document,
'      Case vbext_ComponentType.vbext_ct_Document
'        extention_ = "cls"
'    End Select
'    ' If extention is defined,
'    If (extention_ <> "") Then
'      ' Set file name.
'      Dim fileName_ As String
'      fileName_ = module_.Name & Dot & extention_
'      ' Initialize ignored_ flag.
'      Dim ignored_ As Boolean
'      ignored_ = False
'      ' Initialize subDirectory_.
'      Dim subDirectory_ As String
'      subDirectory_ = ""
'      ' Filter file.
'      Dim filter_ As Variant
'      For Each filter_ In filters_
'        Dim keyValue_ As Variant
'        keyValue_ = Split(filter_, First_Level_Delimiter)
'        ' Process depending on filter's key.
'        ' Set key.
'        Dim key_ As String
'        key_ = keyValue_(Key_Value.Key)
'        ' Set value.
'        Dim value_ As String
'        value_ = keyValue_(Key_Value.Value)
'        keyValue_ = Empty ' Release
'        ' If file name matches filter,
'        If (fileName_ Like key_) Then
'          ' If value_ is NOT defined,
'          If (value_ = "") Then
'            ' Set ignored_ to TRUE.
'            ignored_ = True
'            ' Exit for-next loop.
'            Exit For
'          ' If value_ is defined,
'          Else
'            ' If value_ match directory,
'            If (value_ Like "\*\") Then
'              ' Set ignored_ to FALSE.
'              ignored_ = False
'              ' Remove first directory path.
'              subDirectory_ = Right(value_, Len(value_) - 1)
'            End If
'          End If
'        End If
'      Next
'      filter_ = Empty ' Release
'      If Not (ignored_) Then
'        ' Export module.
'        Dim filePath_ As String
'        filePath_ = DirectoryPath & Path_Separator & subDirectory_ & fileName_
'        Call module_.Export(filePath_)
'        logger_.Info ("Exported file is '" & filePath_ & "'.") ' Log information.
'      End If
'    End If
'  Next
'  filters_ = Empty ' Release
'  Set module_ = Nothing ' Release
'  Set moduleList_ = Nothing ' Release
Exit__:
'  ' Release variants.
'  filters_ = Empty
'  filter_ = Empty
'  keyValue_ = Empty
'  ' Release objects.
'  Set workbook_ = Nothing
'  Set moduleList_ = Nothing
'  Set module_ = Nothing
'  ' Unstak name for logging purpose and release logger_.
'  Call logger_.UnstackName
'  Set logger_ = Nothing
'  ' Set ExitStatus.
'  ExitStatus = exitStatus_
'  ' Return self.
  Set ImportModules = Me
  ' Exit function
  Exit Function
Error:
  ' Set exitStatus_ to error number.
  exitStatus_ = Err.Number
  ' Log error.
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  ' Go to exit.
  GoTo Exit__
End Function

'* Get folders and files.
'* @param[in] Folder. Folder.
'* @param[out] FoundFolders. Folders.
'* @param[out] FoundFiles. Files.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function GetFoldersAndFiles( _
  Folder As String, _
  FoundFolders As String, _
  FoundFiles As String, _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  ' Define error handling.
  On Error GoTo Error
  ' Set logger_.
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  ' Stack name for logging purpose.
  Call logger_.StackName(Me.Name & "." & "GetFoldersAndFiles")
  ' Set exitStatus_.
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  ' Find folders.
  Dim foundFolders_ As Object
  FoundFolders = ""
  Dim fileSystem_ As Object
  Set fileSystem_ = Me.fileSystem
  For Each foundFolders_ In fileSystem_.GetFolder(Folder).SubFolders
    If (FoundFolders <> "") Then
      FoundFolders = FoundFolders + First_Level_Delimiter
    End If
    FoundFolders = FoundFolders + foundFolders_.Name
    logger_.Debug__ ("Folder name : " & foundFolders_.Name)
  Next
  
  ' Find files.
  Dim foundFiles_ As Object
  FoundFiles = ""
  For Each foundFiles_ In fileSystem_.GetFolder(Folder).Files
    Dim name_ As String
    name_ = foundFiles_.Name
    If ( _
      name_ <> "" _
      And _
      Left(name_, 1) <> "~" _
    ) Then
      If (FoundFiles <> "") Then
        FoundFiles = FoundFiles + First_Level_Delimiter
      End If
      FoundFiles = FoundFiles + name_
      logger_.Debug__ ("File name : " & name_)
    End If
  Next
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set GetFoldersAndFiles = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Examine folder or file.
'* @param[in,out] FolderOrFile. Folder or file. if target is link folder or file, return original folder or file.
'* @param[out] IsFolder. If target is folder or file, return TRUE or FALSE respectively.
'* @param[out] Exists. If folder or file exists, return TRUE.
'* @param[out] IsLinkBroken. If link is broken, return TRUE.
'* @param[out] IsPasswordProtected. If folder or file is password protected, return TRUE.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function ExamineFolderOrFile( _
  FolderOrFile As String, _
  Optional Exists As Boolean, _
  Optional IsFolder As Boolean, _
  Optional dateCreated As Date, _
  Optional dateLastModified As Date, _
  Optional dateLastAccessed As Date, _
  Optional IsLinkBroken As Boolean, _
  Optional IsPasswordProtected As Boolean, _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  ' Define error handling.
  On Error GoTo Error
  ' Set logger_.
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  ' Stack name for logging purpose.
  Call logger_.StackName(Me.Name & "." & "ExamineFolderOrFile")
  ' Set exitStatus_.
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  ' Set Exists and IsFolder flags.
  Dim fileSystem_ As Object
  Set fileSystem_ = Me.fileSystem
  If (fileSystem_.FileExists(FolderOrFile)) Then ' If file exists,
    Exists = True
    IsFolder = False
  ElseIf (fileSystem_.FolderExists(FolderOrFile)) Then ' If folder exists,
    Exists = True
    IsFolder = True
  Else ' If folder or file DOESN'T exist,
    Exists = False
  End If
  ' If folder or file exists,
  If Exists Then
    ' Set dates.
    If IsFolder Then ' If folder
      Dim folderOrFile_ As Object
      Set folderOrFile_ = fileSystem_.GetFolder(FolderOrFile)
    Else ' If file,
      Set folderOrFile_ = fileSystem_.GetFile(FolderOrFile)
    End If
    dateCreated = folderOrFile_.dateCreated
    dateLastModified = folderOrFile_.dateLastModified
    dateLastAccessed = folderOrFile_.dateLastAccessed
    ' Define IsLinkBroken flag.
    ' Initialize flag.
    IsLinkBroken = False
    ' If folder or file has link extention; "lnk",
    If Right(FolderOrFile, Len(Link_Extension)) = Link_Extension Then
      ' Search original folder or file.
      Dim windowsScriptingHost_ As Object
      Set windowsScriptingHost_ = CreateObject("WScript.Shell")
      Dim shortcutFolderOrFile_ As Object
      ' Get shortcut object.
      Set shortcutFolderOrFile_ = windowsScriptingHost_.CreateShortcut(FolderOrFile)
      ' Set original folder or file name.
      FolderOrFile = shortcutFolderOrFile_.TargetPath
      ' If folder or file (a) is nether current directory nor parent directoy and (b) doesn't exist,
      If ( _
        Right(FolderOrFile, Len(Current_Directory)) <> Current_Directory _
        And _
        Right(FolderOrFile, Len(Parent_Directory)) <> Parent_Directory _
        And _
        fileSystem_.FileExists(FolderOrFile) = False _
        And _
        fileSystem_.FolderExists(FolderOrFile) = False _
      ) Then
        ' Set flag to TRUE.
        IsLinkBroken = True
      End If
    End If
    ' Define IsPasswordProtected flag.
    IsPasswordProtected = False
  End If
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set ExamineFolderOrFile = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function

'* Get file hash.
'* @param[in] File. File name.
'* @param[out] Hash. Calculated hash.
'* @param[out] ExitStatus (Optional) Exit status. Zero means success.
'* @return ExcelFirstClass. Self.
Public Function GetFileHash( _
  File As String, _
  Hash As String, _
  Optional ExitStatus As Long _
) As ExcelFirstClass
Initialize:
  ' Define error handling.
  On Error GoTo Error
  ' Set logger_.
  Dim logger_ As LoggerClass
  Set logger_ = Me.Logger
  ' Stack name for logging purpose.
  Call logger_.StackName(Me.Name & "." & "GetFileHash")
  ' Set exitStatus_.
  Dim exitStatus_ As Long
  exitStatus_ = Exit_Status.Success
Execute:
  ' Examine existence of file.
  Dim exist_ As Boolean
  Call Me.ExamineFolderOrFile(File, exist_)
  ' If the file eixsts,
  If exist_ Then
    ' Get file size.
    Dim fileSystem_ As Object
    Set fileSystem_ = Me.fileSystem
    Dim file_ As Object
    Set file_ = fileSystem_.GetFile(File)
    Dim fileSize_ As Long
    fileSize_ = file_.Size
    ' Initialize file-bytes array.
    Dim fileNumber_ As Long
    ' Read file bytes.
    fileNumber_ = FreeFile
    Open File For Binary Access Read As fileNumber_
    Dim fileBytes_() As Byte
    ReDim fileBytes_(0 To LOF(fileNumber_) - 1&) As Byte
    Get fileNumber_, , fileBytes_
    Close fileNumber_
    Dim cryptography_ As Object
    Set cryptography_ = CreateObject("System.Security.Cryptography.SHA512Managed")
    'Convert the byte array to hash.
    Dim hashBytes_() As Byte
    hashBytes_ = cryptography_.ComputeHash_2((fileBytes_))
    Erase fileBytes_
    Set cryptography_ = Nothing ' Release
    Dim domDocument_ As Object
    Set domDocument_ = CreateObject("MSXML2.DOMDocument")
    With domDocument_
      .LoadXML "<root />"
      .DocumentElement.DataType = "bin.base64"
      .DocumentElement.nodeTypedValue = hashBytes_
    End With
    Hash = Replace(domDocument_.DocumentElement.Text, vbLf, "")
    Erase hashBytes_ ' Release
    Set domDocument_ = Nothing ' Release
  ' If the file DOESN'T exist,
  Else
    ' Set Hash to blank.
    Hash = ""
  End If
Exit__:
  Call logger_.UnstackName
  ExitStatus = exitStatus_
  Set GetFileHash = Me
  Exit Function
Error:
  exitStatus_ = Err.Number
  Call logger_.Error("Error #" & Err.Number & " ( " & Err.Source & " : " & Err.Description & " )")
  GoTo Exit__
End Function
