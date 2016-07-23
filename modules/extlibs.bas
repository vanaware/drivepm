Attribute VB_Name = "extlibs"
Option Explicit

'Code bellow from http://peltiertech.com/save-retrieve-information-text-files/
Function SaveSetting(sFileName As String, sname As String, _
      Optional sValue As String) As Boolean
  
  Dim iFileNumA As Long
  Dim iFileNumB As Long
  Dim sFile As String
  Dim sXFile As String
  Dim sVarName As String
  Dim sVarValue As String
  Dim lErrLast As Long
  
  ' assume false unless variable is successfully saved
  SaveSetting = False
  
  ' add this workbook's path if not specified
  If Not IsFullName(sFileName) Then
    sFile = ThisWorkbook.path & "\" & sFileName
    sXFile = ThisWorkbook.path & "\X" & sFileName
  Else
    sFile = sFileName
    sXFile = FullNameToPath(sFileName) & "\X" & FullNameToFileName(sFileName)
  End If
  
  ' open text file to read settings
  If fileExists(sFile) Then
    'replace existing settings file
    iFileNumA = FreeFile
    Open sFile For Input As iFileNumA
    iFileNumB = FreeFile
    Open sXFile For Output As iFileNumB
      Do While Not EOF(iFileNumA)
        Input #iFileNumA, sVarName, sVarValue
        If sVarName <> sname Then
          Write #iFileNumB, sVarName, sVarValue
        End If
      Loop
      Write #iFileNumB, sname, sValue
      SaveSetting = True
    Close #iFileNumA
    Close #iFileNumB
    FileCopy sXFile, sFile
    kill sXFile
  Else
    ' make new file
    iFileNumB = FreeFile
    Open sFile For Output As iFileNumB
      Write #iFileNumB, sname, sValue
      SaveSetting = True
    Close #iFileNumB
  End If
  
End Function
Function GetSetting(sFile As String, sname As String, _
      Optional sValue As String) As Boolean
  
  Dim iFileNum As Long
  Dim sVarName As String
  Dim sVarValue As String
  Dim lErrLast As Long
  
  ' assume false unless variable is found
  GetSetting = False
  
  ' add this workbook's path if not specified
  If Not IsFullName(sFile) Then
    sFile = ThisWorkbook.path & "\" & sFile
  End If
  
  ' open text file to read settings
  If fileExists(sFile) Then
    iFileNum = FreeFile
    Open sFile For Input As iFileNum
      Do While Not EOF(iFileNum)
        Input #iFileNum, sVarName, sVarValue
        If sVarName = sname Then
          sValue = sVarValue
          GetSetting = True
          Exit Do
        End If
      Loop
    Close #iFileNum
  End If
  
End Function

Public Sub DebugLog(sLogEntry As String)
  ' write debug information to a log file

  Dim iFile As Integer
  Dim sDirectory As String
  
  sDirectory = ThisWorkbook.path & "\debuglog" & Format$(Now, "YYMMDD") & ".txt"

  iFile = FreeFile

  Open sFileName For Append As iFile
  Print #iFile, Now; " "; sLogEntry
  Close iFile

End Sub

Function IsFullName(sFile As String) As Boolean
  ' if sFile includes path, it contains path separator "\"
  IsFullName = InStr(sFile, "\") > 0
End Function

Function FullNameToPath(sFullName As String) As String
  ''' does not include trailing backslash
  Dim k As Integer
  For k = Len(sFullName) To 1 Step -1
    If Mid(sFullName, k, 1) = "\" Then Exit For
  Next k
  If k < 1 Then
    FullNameToPath = ""
  Else
    FullNameToPath = Mid(sFullName, 1, k - 1)
  End If
End Function

Function FullNameToFileName(sFullName As String) As String
  Dim k As Integer
  Dim sTest As String
  If InStr(1, sFullName, "[") > 0 Then
    k = InStr(1, sFullName, "[")
    sTest = Mid(sFullName, k + 1, InStr(1, sFullName, "]") - k - 1)
  Else
    For k = Len(sFullName) To 1 Step -1
      If Mid(sFullName, k, 1) = "\" Then Exit For
    Next k
    sTest = Mid(sFullName, k + 1, Len(sFullName) - k)
  End If
  FullNameToFileName = sTest
End Function

Function fileExists(ByVal FileSpec As String) As Boolean
   ' by Karl Peterson MS MVP VB
   Dim Attr As Long
   ' Guard against bad FileSpec by ignoring errors
   ' retrieving its attributes.
   On Error Resume Next
   Attr = GetAttr(FileSpec)
   If err.Number = 0 Then
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      fileExists = Not ((Attr And vbDirectory) = vbDirectory)
   End If
End Function

