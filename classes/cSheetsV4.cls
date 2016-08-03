VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSheetsV4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pSheetId As String
Private pAccessToken As String
Private pBrowser As cBrowser
Private pRoot As String
Private pAuthName As String

Public Function setSheetId(id) As cSheetsV4
    pSheetId = id
    Set setSheetId = Me
End Function

Public Function getSheetId() As String
    getSheetId = pSheetId
End Function
Public Function setAuthName(nm) As cSheetsV4
    pAuthName = nm
    Set setAuthName = Me
End Function
Public Function getAuthName() As String
    getAuthName = pAuthName
End Function
Public Function getSheets() As cJobject
    Set getSheets = urlExecute(pSheetId, encodeFields(getSheetDefs()))
    'Debug.Print (getSheets.stringify)
End Function
Public Function sheetExists(sheetName As String) As Boolean
    Dim theSheet As String, s As String, job As cJobject, results As cJobject
    theSheet = LCase(CStr(sheetName))
    '// get all the sheets at  the google end
    Set results = getSheets()
    If (Not results.child("success").value) Then
        sheetExists = False
        Exit Function
    End If
    For Each job In results.child("data").children(1).child("sheets").children
        s = job.toString("properties.title")
        If (LCase(s) = theSheet) Then
            sheetExists = True
            Exit Function
        End If
    Next job
    sheetExists = False
End Function
Public Sub insertSheets(Tables As cJobject)
    Dim job As cJobject, tableName As String, results As cJobject
    'Add Sheet at Google if not exists
    For Each job In Tables.children
        tableName = job.child("range").value
        'Delete everything after ! sign
        If InStr(tableName, "!") > 0 Then tableName = Left(tableName, InStr(tableName, "!") - 1)
        If Not sheetExists(tableName) Then
            Set results = insertSheet(tableName)
        End If
    Next job
End Sub
Public Function insertSheet(sheetName As Variant) As cJobject
    Dim s As String, payload As New cJobject
    s = CStr(sheetName)
    
    With payload.Init(Nothing).add("requests").addArray
        With .add
            With .add("addSheet")
                With .add("properties")
                    .add "title", s
                End With
            End With
        End With
    End With
    Set insertSheet = urlExecute(pSheetId & ":batchUpdate", encodeFields(getInsertSheetDefs()), "POST", payload)
    'Debug.Print (payload.stringify)
    payload.Teardown
End Function





Public Sub pushDatatoSheets(Tables As cJobject)
    Dim payload As New cJobject
    Dim result As New cJobject
    ' add missing sheets
    insertSheets Tables
    '// append tables to sheets
    With payload.Init(Nothing)
        .add "valueInputOption", "RAW"
        '.add "valueInputOption", "USER_ENTERED"
        With .add("data").addArray
            .arrayAppend Tables
        End With
    End With
    ''// do the api call
    'Debug.Print (payload.stringify)
    Set result = urlExecute(pSheetId & "/values:batchUpdate", _
        encodeFields("responses"), "POST", payload)
    'Debug.Print (result.stringify)
    
    '// clean
    result.Teardown
    payload.Teardown
End Sub

Public Function pullDatafromSheets() As cJobject
    Dim ranges As cStringChunker, job As cJobject
    Dim result As cJobject
    Dim Tables As cJobject
    Dim tableName As String
    
    '// get all the sheets names at the google end
    Set result = getSheets()
    If (Not result.child("success").value) Then
        Exit Function
    End If

    Set ranges = New cStringChunker
    For Each job In result.child("data").children(1).child("sheets").children
        ranges.add(job.toString("properties.title")).add (",")
    Next job
    
    '// get data from all sheets at google
    Set result = getValues(ranges.chopWhile(",").toString)
    'Debug.Print result.stringify
    
    '// Populate Tables to return back
    Set Tables = New cJobject
    With Tables.Init(Nothing).addArray
        For Each job In result.child("data").children(1).child("valueRanges").children
            'Debug.Print joc.stringify
            .append job
        Next job
    End With
    'fixTablesName
    'Delete everything after ! sign
    For Each job In Tables.children
        tableName = job.child("range").value
        If InStr(tableName, "!") > 0 Then
            tableName = Left(tableName, InStr(tableName, "!") - 1)
            job.child("range").value = tableName
        End If
    Next job
    Set pullDatafromSheets = Tables
End Function



Public Function getValues(ranges As String) As cJobject
    Set getValues = urlExecute(pSheetId & "/values:batchGet", _
      getValuesParams & "&" & getRangesParams(ranges) & "&" & encodeFields(getValuesDefs()))
    'Debug.Print (getValues.stringify)
End Function

Public Function setValues(values As Variant, sheetName As String, rangeA1 As String) As cJobject
    Dim payload As cJobject
    '// convert array to object
    Set payload = getValuesPayload(values, sheetName, rangeA1)
    ''// do the api call
    Set setValues = urlExecute(pSheetId & "/values:batchUpdate", encodeFields("responses"), "POST", payload)
    'Debug.Print (payload.stringify)
    '// clean
    payload.Teardown
End Function

Public Function getValuesPayload(values As Variant, sheetName As String, rangeA1 As String) As cJobject
    Dim i As Long, j As Long, payload As New cJobject
    ' only using a:z for now TODO.. make it more specific
    With payload.Init(Nothing)
        '.add "valueInputOption", "USER_ENTERED"
        .add "valueInputOption", "RAW"
        With .add("data").addArray.add
            .add "range", sheetName & "!" & Replace(rangeA1, "$", "")
            With .add("values").addArray
                For i = LBound(values, 1) To UBound(values, 1)
                    With .add.addArray
                        For j = LBound(values, 2) To UBound(values, 2)
                            .add , values(i, j)
                        Next j
                    End With
                Next i
            End With
        End With
    End With
    'Debug.Print (payload.stringify)
    Set getValuesPayload = payload
End Function

Private Function getSheetDefs() As String
    getSheetDefs = "sheets(properties(sheetId,title))"
End Function

Private Function getInsertSheetDefs() As String
    getInsertSheetDefs = "replies(addSheet)"
End Function

Private Function getValuesDefs() As String
    getValuesDefs = "valueRanges(range,values)"
End Function

Private Function getValuesParams() As String
    'getValuesParams = "dateTimeRenderOptions=SERIAL_NUMBER&valueRenderOption=FORMATTED_VALUE"
    getValuesParams = "dateTimeRenderOptions=SERIAL_NUMBER&valueRenderOption=UNFORMATTED_VALUE"
End Function

Private Function getRangesParams(ranges As String) As String
    Dim a As Variant, i As Long, c As New cStringChunker
    a = Split(ranges, ",")
    For i = LBound(a) To UBound(a)
        c.add("ranges=").uri(CStr(a(i))).add "&"
    Next i
    getRangesParams = c.chopWhile("&").toString
End Function

'/**
'* encode partial field definitions
'* @param {[string]} varArray actually this is any number of args
'* @return {string} the encoded fields
'*/
Private Function encodeFields(ParamArray arguments() As Variant) As String
    Dim c As New cStringChunker, i
    encodeFields = ""
    For i = LBound(arguments) To UBound(arguments)
        c.uri(CStr(arguments(i))).add (",")
    Next i
    If (c.size > 0) Then
        encodeFields = "fields=" & c.chopWhile(",").toString
    End If
End Function

'/* execute a API request
'  * @param {string} urlTail the url appendage
'  * @param {string} [params] the params
'  * @param {string} [method=GET] what to do
'  * @param {cJobject} [payload] the payload
'  * @return {cjobject} a standard response object
'  */
Private Function urlExecute(urlTail As String, _
    Optional params As String = vbNullString, _
    Optional method As String = "GET", _
    Optional payload As cJobject = Nothing) As cJobject

    Dim url As String, body As String, ob As cJobject, paramstring As New cStringChunker
    Dim result As New cJobject, googled As cOauth2
    
    'Debug.Print "URL_EXECUTE VARS"
    'Debug.Print "URLTAIL= " & urlTail
    'Debug.Print "PARAMS= " & params
    'Debug.Print "METHOD= " & method
    
    body = ""
    If (isSomething(payload)) Then
        body = payload.stringify
    End If
    
    'Debug.Print "PAYLOAD(body)= " & (body)
            
    Set googled = getGoogled(pAuthName)
    
    paramstring.add("?key=").add googled.apikey
    If (params <> vbNullString) Then
        paramstring.add("&").add params
    End If
    
    url = pRoot & urlTail & paramstring.toString
    'Debug.Print "FINALURL= " & url
    'Debug.Print "AUTHHEADER= " & googled.authHeader
    
    If (method <> "GET") Then
        pBrowser.httpPost url, body, True, googled.authHeader, , method
    Else
        pBrowser.httpGET url, , , , , googled.authHeader
    End If
    
    result.Init(Nothing).add "response", pBrowser.text
    
    If (pBrowser.isOk()) Then
        Set ob = JSONParse(pBrowser.text)
        result.add("data").attach ob
        result.add "success", ob.childExists("error") Is Nothing
        
    Else
        result.add "success", False
    End If
    
    Set urlExecute = result
    'Debug.Print "URL EXECUTE RESULT"
    'Debug.Print (result.stringify)
End Function

Private Sub Class_Initialize()
    ' default url for sheets api
    pRoot = "https://sheets.googleapis.com/v4/spreadsheets/"
    Set pBrowser = New cBrowser

    
End Sub




