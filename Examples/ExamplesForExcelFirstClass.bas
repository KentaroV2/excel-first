Attribute VB_Name = "ExamplesForExcelFirstClass"
 Option Explicit
'! This module provides some examples that help you understanding how to use the ExcelFirstClass.
'! @copyright MIT
'
' Edit the followings as needed:
' --------------------------------------------------------------------------------------------------------------
'
' Don't edit the followings:
' --------------------------------------------------------------------------------------------------------------


'* This example teaches how to disable and enable screen updating.
Sub ExamplesForExcelFirstClass_DisableAndEnableScreenUpdating()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_DisableAndEnableScreenUpdating")
  
  ' Disable screen updating
  excelFirst_.ScreenUpdatingFlag = False
  
  ' Enable screen updating
  excelFirst_.ScreenUpdatingFlag = True
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to change date to string.
Sub ExamplesForExcelFirstClass_ChangeDateToString()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_ChangeDateToString")
  
  ' Change date to string.
  Dim now__ As Date
  now__ = Now
  Dim String__ As String
  Dim degreeDetails_ As String
  degreeDetails_ = "Y"
  Call excelFirst_.ChangeDateToString(now__, String__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & String__)
  degreeDetails_ = "M"
  Call excelFirst_.ChangeDateToString(now__, String__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & String__)
  degreeDetails_ = "D"
  Call excelFirst_.ChangeDateToString(now__, String__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & String__)
  degreeDetails_ = "h"
  Call excelFirst_.ChangeDateToString(now__, String__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & String__)
  degreeDetails_ = "m"
  Call excelFirst_.ChangeDateToString(now__, String__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & String__)
  degreeDetails_ = "s"
  Call excelFirst_.ChangeDateToString(now__, String__, degreeDetails_)
  Call logger_.Info(degreeDetails_ & ":" & String__)
  degreeDetails_ = "(undefined)"
  Call excelFirst_.ChangeDateToString(now__, String__)
  Call logger_.Info(degreeDetails_ & ":" & String__)
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to change string to date.
Sub ExamplesForExcelFirstClass_ChangeStringToDate()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_ChangeStringToDate")
  
  ' Change date to string.
  Dim now__ As Date
  now__ = Now
  Dim Date__ As Date
  Dim degreeDetails_ As String
  degreeDetails_ = "M"
  Call excelFirst_.ChangeStringToDate( _
    Format(now__, "yyyy") & _
    "-" & Format(now__, "mm"), _
    Date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(Date__, "yyyy/mm/dd hh:nn:ss"))
  degreeDetails_ = "D"
  Call excelFirst_.ChangeStringToDate( _
    Format(now__, "yyyy") & _
    "-" & Format(now__, "mm") & _
    "-" & Format(now__, "dd"), _
    Date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(Date__, "yyyy/mm/dd hh:nn:ss"))
  degreeDetails_ = "m"
  Call excelFirst_.ChangeStringToDate( _
    Format(now__, "yyyy") & _
    "-" & Format(now__, "mm") & _
    "-" & Format(now__, "dd") & _
    " " & Format(now__, "hh") & _
    ":" & Format(now__, "nn"), _
    Date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(Date__, "yyyy/mm/dd hh:nn:ss"))
  degreeDetails_ = "s"
  Call excelFirst_.ChangeStringToDate( _
    Format(now__, "yyyy") & _
    "-" & Format(now__, "mm") & _
    "-" & Format(now__, "dd") & _
    " " & Format(now__, "hh") & _
    ":" & Format(now__, "nn") & _
    ":" & Format(now__, "ss"), _
    Date__ _
  )
  Call logger_.Info(degreeDetails_ & ":" & Format(Date__, "yyyy/mm/dd hh:nn:ss"))
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to export modules.
Sub ExamplesForExcelFirstClass_ExportModules()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_ExportModules")
  ' Set directory path to desktop.
  Dim windowsScriptingHost_ As Object
  Set windowsScriptingHost_ = CreateObject("WScript.Shell")
  Dim directoryPath_ As String
  directoryPath_ = windowsScriptingHost_.SpecialFolders("Desktop")
  ' Export modules.
  Call excelFirst_.ExportModules( _
    directoryPath_ & Path_Separator & "ExcelFirst", _
    "ExamplesFor*" & First_Level_Delimiter & "\Examples\" _
  )
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub


'* This example teaches how to get folders and files on my document directory.
Sub ExamplesForExcelFirstClass_GetFoldersAndFiles()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_GetFoldersAndFiles")
  
  ' Get folders and files on my document directory.
  Dim windowsScriptingHost_ As Object
  Set windowsScriptingHost_ = CreateObject("WScript.Shell")
  Dim foundFolders_ As String
  Dim foundFiles_ As String
  Call excelFirst_.GetFoldersAndFiles( _
    windowsScriptingHost_.SpecialFolders("MyDocuments"), _
    foundFolders_, _
    foundFiles_ _
  )
  
  ' Display foler names.
  If (foundFolders_ <> "") Then
    Dim folders_ As Variant
    folders_ = Split(foundFolders_, First_Level_Delimiter)
    Dim folder_ As Variant
    For Each folder_ In folders_
      logger_.Info ("Folder name = " & folder_)
    Next
  End If
  
  ' Display file names.
  If (foundFiles_ <> "") Then
    Dim files_ As Variant
    files_ = Split(foundFiles_, First_Level_Delimiter)
    Dim file_ As Variant
    For Each file_ In files_
      logger_.Info ("File name = " & file_)
    Next
  End If
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

'* This example teaches how to examine folders and files on my document directory.
Sub ExamplesForExcelFirstClass_ExamineFolderOrFile()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_ExamineFolderOrFile")
  
  ' Get folders and files on my document directory.
  Dim windowsScriptingHost_ As Object
  Set windowsScriptingHost_ = CreateObject("WScript.Shell")
  Dim foundFolders_ As String
  Dim foundFiles_ As String
  Call excelFirst_.GetFoldersAndFiles( _
    windowsScriptingHost_.SpecialFolders("MyDocuments"), _
    foundFolders_, _
    foundFiles_ _
  )
  
  ' Display foler names.
  If (foundFolders_ <> "") Then
    Dim folders_ As Variant
    folders_ = Split(foundFolders_, First_Level_Delimiter)
    Dim folder_ As Variant
    For Each folder_ In folders_
      logger_.Info ("Folder name = " & folder_)
      Dim exist_ As Boolean
      Dim isFolder_ As Boolean
      Dim dateCreated_ As Date
      Dim dateLastModified_ As Date
      Dim dateLastAccessed_ As Date
      Dim isLinkBroken_ As Boolean
      Dim isPasswordProtected_ As Boolean
      Call excelFirst_.ExamineFolderOrFile(CStr(folder_), exist_, isFolder_, dateCreated_, dateLastModified_, dateLastAccessed_, isLinkBroken_, isPasswordProtected_)
      logger_.Info ("exist_ = " & exist_)
      logger_.Info ("isFolder_ = " & isFolder_)
      logger_.Info ("dateCreated_ = " & dateCreated_)
      logger_.Info ("dateLastModified_ = " & dateLastModified_)
      logger_.Info ("dateLastAccessed_ = " & dateLastAccessed_)
      logger_.Info ("IsLinkBroken = " & isLinkBroken_)
      logger_.Info ("IsPasswordProtected = " & isPasswordProtected_)
    Next
  End If
  
  ' Display file names.
  If (foundFiles_ <> "") Then
    Dim files_ As Variant
    files_ = Split(foundFiles_, First_Level_Delimiter)
    Dim file_ As Variant
    For Each file_ In files_
      logger_.Info ("File name = " & file_)
      Call excelFirst_.ExamineFolderOrFile(CStr(file_), exist_, isFolder_, dateCreated_, dateLastModified_, dateLastAccessed_, isLinkBroken_, isPasswordProtected_)
      logger_.Info ("exist_ = " & exist_)
      logger_.Info ("isFolder_ = " & isFolder_)
      logger_.Info ("dateCreated_ = " & dateCreated_)
      logger_.Info ("dateLastModified_ = " & dateLastModified_)
      logger_.Info ("dateLastAccessed_ = " & dateLastAccessed_)
      logger_.Info ("IsLinkBroken = " & isLinkBroken_)
      logger_.Info ("IsPasswordProtected = " & isPasswordProtected_)
    Next
  End If
  
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub

'* This example teaches how to user GetFileHas function.
'* @attention This example requires a "test.xls" file on a directory where a file running this example locates.
Sub ExamplesForExcelFirstClass_GetFileHash()

  ' Instantiate First class.
  Dim excelFirst_ As ExcelFirstClass
  Set excelFirst_ = New ExcelFirstClass
  
  ' Set INFO as logger level.
  Dim logger_ As LoggerClass
  Set logger_ = excelFirst_.Logger
  Call logger_.SetLevel(Logger_Level.Info)
  
  ' Stack name.
  Call logger_.StackName("ExamplesForExcelFirstClass_GetFileHash")
  
  ' Define target file.
  Dim targetFile_ As String
  targetFile_ = "test.xlsx"
  
  ' Get file hash.
  Dim hash_ As String
  Call excelFirst_.GetFileHash( _
    targetFile_, _
    hash_ _
  )
  logger_.Info ("targetFile_ = " & targetFile_)
  logger_.Info ("hash_ = " & hash_)
 
  ' Unstack name.
  Call logger_.UnstackName
  
End Sub
