Attribute VB_Name = "functions"
Option Explicit

'Code bellow re-used from Bruce McPherson,  desktop liberation, http://ramblings.mcpher.com

#If VBA7 And Win64 Then
Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal Operation As String, _
  ByVal fileName As String, _
  Optional ByVal parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMaximizedFocus _
  ) As LongLong
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As LongLong, ByVal dwflags As LongLong, _
    ByVal lpWideCharStr As LongLong, ByVal cchWideChar As LongLong, _
    ByVal lpMultiByteStr As LongLong, ByVal cchMultiByte As LongLong, _
    ByVal lpDefaultChar As LongLong, ByVal lpUsedDefaultChar As LongLong) As LongLong
#Else
Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal Operation As String, _
  ByVal fileName As String, _
  Optional ByVal parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMaximizedFocus _
  ) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwflags As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const CP_UTF8 = 65001
Public Const cFailedtoGetHandle = -1
Public Const config_txt = "DrivePM.config"


Public Function getGoogled(scope As String, _
                                Optional replacementpackage As cJobject = Nothing, _
                                Optional ClientID As String = vbNullString, _
                                Optional ClientSecret As String = vbNullString, _
                                Optional complain As Boolean = True, _
                                Optional cloneFromeScope As String = vbNullString, _
                                Optional apikey As String = vbNullString) As cOauth2
    Dim o2 As cOauth2
    Set o2 = New cOauth2
    With o2.googleAuth(scope, replacementpackage, ClientID, ClientSecret, complain, cloneFromeScope, apikey)
        If Not .hasToken And complain Then
            MsgBox ("Failed to authorize to google for scope " & scope & ":denied code " & o2.denied)
        End If
    End With
    
    Set getGoogled = o2
End Function

Public Function JSONParse(s As String, Optional jtype As eDeserializeType, Optional complain As Boolean = True) As cJobject
    Dim j As New cJobject
    Set JSONParse = j.Init(Nothing).parse(s, jtype, complain)
    j.tearDown
End Function

Public Function JSONStringify(j As cJobject, Optional blf As Boolean) As String
    JSONStringify = j.stringify(blf)
End Function

Public Function quote(s As String) As String
    quote = q & s & q
End Function
Public Function q() As String
    q = Chr(34)
End Function
Public Function qs() As String
    qs = Chr(39)
End Function
Public Function bracket(s As String) As String
    bracket = "(" & s & ")"
End Function

Public Function escapeify(s As String) As String
    escapeify = Replace( _
                    Replace( _
                        Replace( _
                            Replace( _
                                Replace(s, "\", "\\"), _
                                    q, "\" & q), _
                                "%", "\" & "%"), _
                            ">", "\>"), _
                        "<", "\<")
    'If s <> escapeify Then Debug.Print escapeify
End Function

Public Function unEscapify(s As String) As String
    unEscapify = Replace( _
                    Replace( _
                        Replace( _
                            Replace( _
                                Replace(s, "\" & q, q), _
                                 "\" & "%", "%"), _
                             "\>", ">"), _
                         "\<", "<"), _
                    "\\", "\")
End Function

Public Function UTF16To8(ByVal UTF16 As String) As String
Dim sBuffer As String
#If VBA7 And Win64 Then
    Dim lLength As LongLong
#Else
    Dim lLength As Long
#End If
If UTF16 <> "" Then
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, 0, 0, 0, 0)
    sBuffer = Space$(CLng(lLength))
    lLength = WideCharToMultiByte( _
        CP_UTF8, 0, StrPtr(UTF16), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF16To8 = Left$(sBuffer, CLng(lLength - 1))
Else
    UTF16To8 = ""
End If
End Function
'end of utf16to8


Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False, _
   Optional UTF8Encode As Boolean = True _
) As String

Dim StringValCopy As String: StringValCopy = _
    IIf(UTF8Encode, UTF16To8(StringVal), StringVal)
Dim StringLen As Long: StringLen = Len(StringValCopy)

If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

  If SpaceAsPlus Then Space = "+" Else Space = "%20"

  For i = 1 To StringLen
    Char = Mid$(StringValCopy, i, 1)
    CharCode = Asc(Char)
    Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        result(i) = Char
      Case 32
        result(i) = Space
      Case 0 To 15
        result(i) = "%0" & Hex(CharCode)
      Case Else
        result(i) = "%" & Hex(CharCode)
    End Select
  Next i
  URLEncode = Join(result, "")

End If
End Function

Public Function isSomething(o As Object) As Boolean
    isSomething = Not o Is Nothing
End Function

Public Function makeKey(v As Variant) As String
    makeKey = LCase(Trim(CStr(v)))
End Function

Public Function compareAsKey(a As Variant, b As Variant, Optional asKey As Boolean = True) As Boolean
    If (asKey And TypeName(a) = "String" And TypeName(b) = "String") Then
        compareAsKey = (makeKey(a) = makeKey(b))
    Else
        compareAsKey = (a = b)
    
    End If
End Function

' The below is taken from http://stackoverflow.com/questions/496751/base64-encode-string-in-vbscript
Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    ' function inserts line feeds so we need to get rid of them
    Base64Encode = Replace(oNode.text, vbLf, "")
    Set oNode = Nothing
    Set oXML = Nothing
End Function
'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function
' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  base64String = Replace(base64String, vbLf, "")
   
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

Function getMySheetId() As String
    Dim SheetID As String
    If GetConfig(config_txt, "SheetID", SheetID) Then
        getMySheetId = SheetID
    Else
        MsgBox ("failed on getting config for sheetid")
        Exit Function
    End If
End Function


'// below here are the useful functions you'll need
Function sheetExistsAtGoogle(sheetAccess As cSheetsV4, sheetName As String, Optional complain As Boolean = False) As cJobject
    Dim theSheet As String, s As String, job As cJobject, results As cJobject, result As cJobject
    theSheet = LCase(CStr(sheetName))
    '// get all the sheets at  the google end
    Set results = sheetAccess.getSheets()
    If (Not results.child("success").value) Then
        MsgBox "failed getting sheets meta data " & results.toString("response")
        Set sheetExistsAtGoogle = Nothing
        Exit Function
    End If
    For Each job In results.child("data").children(1).child("sheets").children
        s = job.toString("properties.title")
        If (LCase(s) = theSheet) Then
            Set sheetExistsAtGoogle = job
            Exit Function
        End If
    Next job
    ' we need to create it
      Set result = sheetAccess.insertSheet(sheetName)
      If (Not result.child("success").value) Then
          MsgBox "failed inserting sheet " & result.toString("response")
          Set sheetExistsAtGoogle = Nothing
          Exit Function
      Else
        'check again for sure
        Set results = sheetAccess.getSheets()
          For Each job In results.child("data").children(1).child("sheets").children
              s = job.toString("properties.title")
              If (LCase(s) = theSheet) Then
                  Set sheetExistsAtGoogle = job
                  Exit Function
              End If
          Next job
      End If
    If (complain) Then
        MsgBox ("sheet does not exist on Google " & theSheet)
    End If
    Set sheetExistsAtGoogle = Nothing
End Function


'Code bellow from http://peltiertech.com/save-retrieve-information-text-files/
Function SaveConfig(sFileName As String, sname As String, _
      Optional sValue As String) As Boolean
  
  Dim iFileNumA As Long
  Dim iFileNumB As Long
  Dim sFile As String
  Dim sXFile As String
  Dim sVarName As String
  Dim sVarValue As String
  Dim lErrLast As Long
  
  ' assume false unless variable is successfully saved
  SaveConfig = False
  
  ' add this workbook's path if not specified
  If Not IsFullName(sFileName) Then
    sFile = ActiveProject.Path & PathSeparator & sFileName
    sXFile = ActiveProject.Path & PathSeparator & "X" & sFileName
  Else
    sFile = sFileName
    sXFile = FullNameToPath(sFileName) & PathSeparator & "X" & FullNameToFileName(sFileName)
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
      SaveConfig = True
    Close #iFileNumA
    Close #iFileNumB
    FileCopy sXFile, sFile
    kill sXFile
  Else
    ' make new file
    iFileNumB = FreeFile
    Open sFile For Output As iFileNumB
      Write #iFileNumB, sname, sValue
      SaveConfig = True
    Close #iFileNumB
  End If
  
End Function

Function GetConfig(sFile As String, sname As String, _
      Optional sValue As String) As Boolean
  
  Dim iFileNum As Long
  Dim sVarName As String
  Dim sVarValue As String
  Dim lErrLast As Long
  
  ' assume false unless variable is found
  GetConfig = False
  
  ' add this workbook's path if not specified
  If Not IsFullName(sFile) Then
    sFile = ActiveProject.Path & PathSeparator & sFile
  End If
  
  ' open text file to read settings
  If fileExists(sFile) Then
    iFileNum = FreeFile
    Open sFile For Input As iFileNum
      Do While Not EOF(iFileNum)
        Input #iFileNum, sVarName, sVarValue
        If sVarName = sname Then
          sValue = sVarValue
          GetConfig = True
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
  
  sDirectory = ActiveProject.Path & "\debuglog" & Format$(Now, "YYMMDD") & ".txt"

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

Function A1Notation(values As Variant) As String
    A1Notation = Get_Alphabet(LBound(values, 2)) & LBound(values, 1) & ":" & Get_Alphabet(UBound(values, 2)) & UBound(values, 1)
End Function

'Returns the alphabet associated with the column
'intNumber: The column number
'Return Value: Alphabet associated with the column number
'LINK = http://software-solutions-online.com/retrieving-a-range-of-cells/
Private Function Get_Alphabet(ByVal intNumber As Integer) As String
    Dim result As String
    If intNumber \ 26 > 0 Then
        result = Get_Alphabet(intNumber \ 26)
    End If
    result = result & Strings.Trim(Chr((intNumber Mod 26) + 64))
    Get_Alphabet = result
End Function


'Returns Randon number between two integers
Public Function Randbetween(a As Double, b As Double) As Double
    Randomize
    Randbetween = Int(Rnd() * (b - a + 1) + a)
End Function

'Returns Hex Number with specified lenght
Public Function DEC2HEX(a As Double, b As Integer) As String
    DEC2HEX = Replace(Space(b - Len(Hex(a))), " ", "0") & Hex(a)
End Function

' Return an pseud guid number
' link = http://stackoverflow.com/questions/7031347/how-can-i-generate-guids-in-excel
Public Function genGuid() As String
'=CONCATENATE(DEC2HEX(RANDBETWEEN(0;4294967295);8);"-";DEC2HEX(RANDBETWEEN(0;65535);4);"-";DEC2HEX(RANDBETWEEN(16384;20479);4);"-";DEC2HEX(RANDBETWEEN(32768;49151);4);"-";DEC2HEX(RANDBETWEEN(0;65535);4);DEC2HEX(RANDBETWEEN(0;4294967295);8))
    Dim Guid As String
    Guid = DEC2HEX(Randbetween(0, 4294967295#), 8)
    Guid = Guid & "-" & DEC2HEX(Randbetween(0, 65535), 4)
    Guid = Guid & "-" & DEC2HEX(Randbetween(16384, 20479), 4)
    Guid = Guid & "-" & DEC2HEX(Randbetween(32768, 49151), 4)
    Guid = Guid & "-" & DEC2HEX(Randbetween(0, 65535), 4)
    Guid = Guid & DEC2HEX(Randbetween(0, 4294967295#), 8)
    genGuid = Guid
End Function

Sub forceUUID()
    Dim MyTask As Task
    Dim MyResource As Resource
    
    CustomFieldRename pjCustomTaskText30, "UUID"
    CustomFieldRename pjCustomResourceText30, "UUID"
    
    Set MyTask = ActiveProject.ProjectSummaryTask
    If getTaskID(MyTask) = "" Then setTaskID MyTask
    For Each MyTask In ActiveProject.Tasks
        'Debug.Print MyTask.GetField(pjTaskText30)
        If isSomething(MyTask) Then
            If getTaskID(MyTask) = "" Then
                setTaskID MyTask
            End If
        End If
    Next
 
    For Each MyResource In ActiveProject.Resources
        'Debug.Print MyTask.GetField(pjTaskText30)
        If isSomething(MyResource) Then
            If getResourceID(MyResource) = "" Then
                setResourceID MyResource
            End If
        End If
    Next
    
End Sub

Public Function getTaskbyID(MyProject As Project, TaskID As String) As Task
    Dim MyTask As Task
    For Each MyTask In MyProject.Tasks
        If isSomething(MyTask) Then
            If getTaskID(MyTask) = TaskID Then
                Set getTaskbyID = MyTask
                Exit Function
            End If
        End If
    Next
    Set MyTask = MyProject.ProjectSummaryTask
    If getTaskID(MyTask) = TaskID Then
        Set getTaskbyID = MyTask
        Exit Function
    End If
    Set getTaskbyID = Nothing
End Function

Public Function getResourcebyID(MyProject As Project, ResourceID As String) As Resource
    Dim MyResource As Resource
    For Each MyResource In MyProject.Resources
        If isSomething(MyResource) Then
            If getResourceID(MyResource) = ResourceID Then
                Set getResourcebyID = MyResource
                Exit Function
            End If
        End If
    Next
    Set getResourcebyID = Nothing
End Function

Public Function getTaskID(MyTask As Task) As String
    Dim TaskID As String
    TaskID = MyTask.GetField(pjTaskText30)
    If TaskID = "" Then
        setTaskID MyTask, genGuid
        TaskID = MyTask.GetField(pjTaskText30)
    End If
    getTaskID = TaskID
End Function

Public Function getResourceID(MyResource As Resource) As String
    Dim ResourceID As String
    ResourceID = MyResource.GetField(pjResourceText30)
    If ResourceID = "" Then
        setResourceID MyResource, genGuid
        ResourceID = MyResource.GetField(pjResourceText30)
    End If
    getResourceID = ResourceID
End Function

Private Sub setTaskID(MyTask As Task, Optional Guid As String = vbNullString)
    If Guid = vbNullString Then
        MyTask.SetField pjTaskText30, genGuid
    Else
        MyTask.SetField pjTaskText30, Guid
    End If
End Sub

Private Sub setResourceID(MyResource As Resource, Optional Guid As String = vbNullString)
    If Guid = vbNullString Then
        MyResource.SetField pjResourceText30, genGuid
    Else
        MyResource.SetField pjResourceText30, Guid
    End If
End Sub

Public Function getTaskCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("UUID (pjTaskText30)", pjTaskText30).add "read-only", True
        .add("pjTaskGuid", pjTaskGuid).add "read-only", True
        .add("pjTaskUniqueID", pjTaskUniqueID).add "read-only", True
        .add("pjTaskID", pjTaskID).add "read-only", True
        .add("pjTaskSummary", pjTaskSummary).add "read-only", True
        .add("pjTaskOutlineNumber", pjTaskOutlineNumber).add "read-only", True
        .add("pjTaskOutlineLevel", pjTaskOutlineLevel).add "read-only", True
        .add("pjTaskParentTask", pjTaskParentTask).add "read-only", True
        .add("pjTaskWBS", pjTaskWBS).add "read-only", True
        .add("pjTaskName", pjTaskName).add "read-only", True
        .add("pjTaskDuration", pjTaskDuration).add "read-only", False           'Read-only for summary tasks
        .add("pjTaskEstimated", pjTaskEstimated).add "read-only", True
        .add("pjTaskUniquePredecessors", pjTaskUniquePredecessors).add "read-only", True
        .add("pjTaskUniqueSuccessors", pjTaskUniqueSuccessors).add "read-only", True
        .add("pjTaskStart", pjTaskStart).add "read-only", True
        .add("pjTaskFinish", pjTaskFinish).add "read-only", True
        .add("pjTaskCreated", pjTaskCreated).add "read-only", True
        .add("pjTaskDeadline", pjTaskDeadline).add "read-only", True
        .add("pjTaskConstraintDate", pjTaskConstraintDate).add "read-only", True
        .add("pjTaskConstraintType", pjTaskConstraintType).add "read-only", True
        .add("pjTaskCalendar", pjTaskCalendar).add "read-only", True
        .add("pjTaskPercentComplete", pjTaskPercentComplete).add "read-only", True
        .add("pjTaskPercentWorkComplete", pjTaskPercentWorkComplete).add "read-only", True
        .add("pjTaskPhysicalPercentComplete", pjTaskPhysicalPercentComplete).add "read-only", True
        .add("pjTaskUpdateNeeded", pjTaskUpdateNeeded).add "read-only", True
        .add("pjTaskFixedCost", pjTaskFixedCost).add "read-only", True
        .add("pjTaskEffortDriven", pjTaskEffortDriven).add "read-only", True
        .add("pjTaskType", pjTaskType).add "read-only", True
        .add("pjTaskFixedDuration", pjTaskFixedDuration).add "read-only", True
        .add("pjTaskMilestone", pjTaskMilestone).add "read-only", True
    End With
    Set getTaskCols = Cols
End Function

Public Function getTaskField(MyTask As Task, FieldID As Long) As Variant
    Select Case FieldID
        Case pjTaskSummary
            getTaskField = MyTask.Summary
        Case pjTaskDuration
            getTaskField = MyTask.Duration 'in minutes
        Case pjTaskStart
            getTaskField = Date2Serial(MyTask.start)
        Case pjTaskFinish
            getTaskField = Date2Serial(MyTask.Finish)
        Case pjTaskCreated
            getTaskField = Date2Serial(MyTask.Created)
        Case pjTaskDeadline
            getTaskField = Date2Serial(MyTask.Deadline)
        Case pjTaskConstraintDate
            getTaskField = Date2Serial(MyTask.ConstraintDate)
        Case pjTaskConstraintDate
            getTaskField = Date2Serial(MyTask.ConstraintDate)
        Case pjTaskFixedCost
            getTaskField = MyTask.FixedCost
        Case pjTaskEstimated
            getTaskField = MyTask.Estimated
        Case pjTaskEffortDriven
            getTaskField = MyTask.EffortDriven
        Case pjTaskFixedDuration
            getTaskField = (MyTask.Type = pjFixedDuration)
        Case pjTaskType
            getTaskField = MyTask.Type
        Case pjTaskParentTask
            getTaskField = getTaskID(MyTask.OutlineParent)
        Case pjTaskUpdateNeeded
            getTaskField = MyTask.UpdateNeeded
        Case pjTaskMilestone
            getTaskField = MyTask.Milestone
        Case pjTaskPercentComplete
            getTaskField = MyTask.PercentComplete
        Case pjTaskPercentWorkComplete
            getTaskField = MyTask.PercentWorkComplete
        Case pjTaskPhysicalPercentComplete
            getTaskField = MyTask.PhysicalPercentComplete
        Case Else
            getTaskField = MyTask.GetField(FieldID)
    End Select
End Function

Public Function getResourceCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("UUID (pjResourceText30)", pjResourceText30).add "read-only", True
        .add("pjResourceGuid", pjResourceGuid).add "read-only", True
        .add("pjResourceUniqueID", pjResourceUniqueID).add "read-only", True
        .add("pjResourceID", pjResourceID).add "read-only", True
        .add("pjResourceName", pjResourceName).add "read-only", True
        .add("pjResourceBaseCalendar", pjResourceBaseCalendar).add "read-only", True
        .add("pjResourceCreated", pjResourceCreated).add "read-only", True
        .add("pjResourceStandardRate", pjResourceStandardRate).add "read-only", True
        
    End With
    Set getResourceCols = Cols
End Function

Public Function getResourceField(MyResource As Resource, FieldID As Long) As Variant
    Select Case FieldID
        Case pjResourceCreated
            getResourceField = Date2Serial(MyResource.Created)
        Case pjResourceStandardRate
            getResourceField = MyResource.StandardRate
        Case Else
            getResourceField = MyResource.GetField(FieldID)
    End Select
End Function

Public Function Date2Serial(d As Variant) As Variant
    Dim serial As Double
    Dim text As String
    
    On Error GoTo errorHandler
    serial = CDbl(d)
    Date2Serial = serial
    Exit Function
    
errorHandler:
    text = CStr(d)
    Date2Serial = text
End Function

