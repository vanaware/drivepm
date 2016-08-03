VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 3:54:18 PM : from manifest:7471153 gist https://gist.github.com/brucemcpherson/3423912/raw/cBrowser.cls
Option Explicit
' acknowledgement
' http://pastie.org/1192157 for basic authentication 'how to'
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
' v2.14
Private pLockActive As Boolean
Private pLockdown As Boolean
Private pSuccessCode As String
Private pDeniedCode As String
Private pResponseHeaders As String
Private pOptionUrl As String
Private pHtml As String
Private pText As String
Private WithEvents pIeOB As InternetExplorer
Attribute pIeOB.VB_VarHelpID = -1
Private pStatus As Long
Const SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID = 8192


' everything to do with accessing web pages from Excel
Public Property Get browser() As InternetExplorer
        Set browser = pIeOB.Application
        
End Property
Public Property Get isOk() As Boolean
    isOk = (pStatus = 200 Or pStatus = 201)
End Property
Public Property Get status() As Long
        status = pStatus
End Property
Public Property Get responseHeaders() As String
        responseHeaders = pResponseHeaders
End Property
Public Property Get optionURL() As String
        optionURL = pOptionUrl
End Property
Public Property Get successCode() As String
        successCode = pSuccessCode
End Property
Public Property Get deniedCode() As String
        deniedCode = pDeniedCode
End Property
Public Property Get text() As String
        text = pText
End Property
Public Property Get url() As String
    url = pHtml
End Property
Public Function Init() As cBrowser
    Set pIeOB = New InternetExplorer
    Set Init = Me
End Function

Public Function Navigate(fn As String, Optional lockDown As Boolean = False, _
                    Optional visible As Boolean = True) As cBrowser
    ' bring up the web page requested
    pHtml = fn
    pLockdown = lockDown
    pSuccessCode = vbNullString
    pDeniedCode = vbNullString
    With browser
        If lockDown Then
            .AddressBar = False
            .MenuBar = False
            .Resizable = False
        End If
        .visible = visible
        .Navigate2 pHtml
        pLockActive = True
        ' will fire document complete, then we can set this off
        Do
            DoEvents
            If Not pLockdown Then
                pLockActive = Not (.readyState = READYSTATE_COMPLETE And Not .Busy)
            End If
        Loop Until Not pLockActive

    End With
    Set Navigate = Me
    
End Function
Private Function getHttpObject(Optional timeout As Long = 0) As Object
    '// some installation dont have server object installed, so fall back
    Dim ob As Object
    On Error GoTo missing
    Set ob = CreateObject("Msxml2.ServerXMLHTTP.6.0")
    ob.setOption 2, ob.getOption(2) - SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID
    ' can have change of timeout for complex/long queries
    If timeout <> 0 Then
        ob.setTimeouts 0, 30 * 1000, 30 * 1000, timeout * 1000
    End If
    Set getHttpObject = ob
    Exit Function
    
missing:
    On Error GoTo screwed
    Set ob = CreateObject("Msxml2.XMLHTTP.6.0")
    Debug.Print "falling back to client xmlhttp - server is not installed on this machine"
    Set getHttpObject = ob
    Exit Function

screwed:
    MsgBox ("cant find either server or client xmlhttp - there missing files from your windows installtion")
    Exit Function
    
End Function
Public Function httpPost(fn As String, _
        Optional data As String = vbNullString, Optional isjson As Boolean = False, _
        Optional authHeader As String = vbNullString, _
        Optional additionalHeaders As cJobject = Nothing, _
        Optional method As String = "POST") As String
    pHtml = fn
    Dim v As Variant, ohttp As Object
    
    Set ohttp = getHttpObject
    If (ohttp Is Nothing) Then
        Exit Function
    End If
    
    With ohttp

        .Open method, pHtml, False
        If isjson Then
            .setRequestHeader "Content-Type", "application/json"
        Else
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        End If
        If (authHeader <> vbNullString) Then
            ohttp.setRequestHeader "Authorization", authHeader
        End If
        extraHeaders ohttp, additionalHeaders

        ohttp.send data
        storeStuff ohttp
    End With
    
    httpPost = pText
    Set ohttp = Nothing
End Function
Private Function extraHeaders(ohttp As Object, additionalHeaders As cJobject) As Object
    Dim job As cJobject
    ' any extra headers?
    If Not additionalHeaders Is Nothing Then
        For Each job In additionalHeaders.children
            ohttp.setRequestHeader job.Key, job.value
        Next job
    End If
    Set extraHeaders = ohttp
End Function

Public Function httpGET(fn As String, _
        Optional authUser As String = vbNullString, _
        Optional authPass As String = vbNullString, _
        Optional accept As String = vbNullString, _
        Optional timeout As Long = 0, _
        Optional authHeader As String = vbNullString, _
        Optional additionalHeaders As Object = Nothing, _
        Optional method As String = "GET") As String
    
    pHtml = fn

    Dim ohttp As Object, job As Object
    Set ohttp = getHttpObject(timeout)
    
    Call ohttp.Open(method, pHtml, False)
    If (authUser <> vbNullString) Then
    ' need to do basic authentication
    ' acknowledgement to http://pastie.org/1192157
        ohttp.setRequestHeader "Content-Type", "application/json"
        ohttp.setRequestHeader "Accept", "application/json"
        ohttp.setRequestHeader "Authorization", "Basic " + _
            Base64Encode(authUser + ":" + authPass)
    End If
    
    ' some times we need to set the accept header
    If accept <> vbNullString Then
        ohttp.setRequestHeader "Accept", accept
    End If


    ' this would be if we were doing an oauth2
    If (authHeader <> vbNullString) Then
        ohttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        ohttp.setRequestHeader "Authorization", authHeader
    End If

    
    extraHeaders ohttp, additionalHeaders
    'Debug.Print method, pHtml
    ohttp.send ""
    storeStuff ohttp
    httpGET = pText
    Set ohttp = Nothing
End Function
Private Sub storeStuff(o As Object)
    With o
        pStatus = .status
        pText = .responseText
        pResponseHeaders = .getAllResponseHeaders()
    End With
End Sub
Public Function element(eID As String) As Object
    On Error GoTo crappedout
    Set element = browser().Document.getElementById(eID)
    Exit Function
crappedout:
    Set element = Nothing
   
End Function
Public Function elementTags(tag As String) As Object
    On Error GoTo crappedout
    Set elementTags = browser().Document.getElementsByTagName(tag)
    Exit Function
crappedout:
    Set elementTags = Nothing
   
End Function
Public Property Get ElementText(eID As String) As String
    Dim e As Object
    Set e = element(eID)
    If (e Is Nothing) Then
        ElementText = ""
    Else
        ElementText = e.value
    End If
End Property

Public Sub kill()
    browser.Quit

End Sub

Private Sub Class_Initialize()
    pLockActive = False
End Sub

Private Sub Class_Terminate()
    Set pIeOB = Nothing
End Sub
Public Sub Teardown()
    If Not pIeOB Is Nothing Then
        kill
        Set pIeOB = Nothing
    End If
    
End Sub



Private Sub pIeOB_OnQuit()
    pLockActive = False
End Sub


Private Sub pIeOB_TitleChange(ByVal text As String)
    Dim s As String, f As String
    f = "Denied error="
    s = "Success code="
    If (pLockdown) Then
        If (Left(text, Len(s)) = s) Then
            pSuccessCode = Mid(text, Len(s) + 1)
            pLockActive = False
        ElseIf (Left(text, Len(f)) = f) Then
            pDeniedCode = Mid(text, Len(f) + 1)
            pLockActive = False
        End If
        If (Not pLockActive) Then
            With browser
                .visible = False
            End With
        End If
    End If

End Sub







