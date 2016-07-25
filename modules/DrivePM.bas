Attribute VB_Name = "DrivePM"


Option Explicit


'// this is about setting up for OAUTH2
'// all you have to do is to create a google console developer project
'// create some credentials of app type other
'// activeate the sheets api
'// copy the credentials to config file
'// and run this function once.
'// you can delete it once you;ve run it - its not needed any more if you successfully go through the google auth process

Const config_txt = "DrivePM.config"

Public Sub RunOnce()
    'your config here
    
    'SaveConfig config_txt, "SheetID", "xxx"
    'SaveConfig config_txt, "ClientID", "xxxx.apps.googleusercontent.com"
    'SaveConfig config_txt, "ClientSecret", "xxx"
    
    Dim ClientID As String
    Dim ClientSecret As String
    
    If GetConfig(config_txt, "ClientID", ClientID) And GetConfig(config_txt, "ClientSecret", ClientSecret) Then
        getGoogled "sheets", , ClientID, ClientSecret
    Else
        MsgBox ("failed on getting config for oauth clientid or/and clientsecret")
        Exit Sub
    End If
    
End Sub

'/**
' * put array contents to google sheets
' * @param {string} [sheetname]
' * @param {array} the valeus
' * @return boolean if succes
'*/
Private Function putTableToSheets(table As String, values As Variant) As Boolean
    Dim result As cJobject, sheetAccess As cSheetsV4
    '// set up an access object
    Set sheetAccess = New cSheetsV4
    sheetAccess.setAuthName("sheets").setSheetId (getMySheetId())
    '// put the sheet data .. will bacth this up later
    ' see if it exists at google
    If Not (sheetExistsAtGoogle(sheetAccess, table) Is Nothing) Then
        '// and write it
        Set result = sheetAccess.setValues(values, table, A1Notation(values))
        If (Not result.child("success").value) Then
            MsgBox ("failed on sheets API " + result.child("response").stringify)
            putTableToSheets = False
        Else
            putTableToSheets = True
        End If
    Else
        putTableToSheets = False
    End If
End Function


'/**
' * get the sheet contents from google
' * @param {string} [sheetname]
' * @return {array} the result
'*/
Private Function getTableFromSheets(table As String) As Variant
    Dim result As cJobject, sheetAccess As cSheetsV4
    Dim d() As Variant
    Dim values As cJobject, jor As cJobject, joc As cJobject
    '// set up an access object
    Set sheetAccess = New cSheetsV4
    sheetAccess.setAuthName("sheets").setSheetId (getMySheetId())
    '// get the data for the sheets that exist or were asked for
    '// see if they exist at google
    If Not (sheetExistsAtGoogle(sheetAccess, table) Is Nothing) Then
        '// now lets get the data in the selected sheets
        Set result = sheetAccess.getValues(table)
        If (Not result.child("success").value) Then
            MsgBox ("failed on sheets API " + result.child("response").stringify)
        Else
            '// now write the data to array
            Set values = result.child("data").children(1).child("valueRanges").children(1).child("values")
            ReDim d(0 To values.children.Count - 1, 0 To values.children(1).children.Count - 1)
            For Each jor In values.children
                If jor.children.Count - 1 > UBound(d, 2) Then
                    ReDim Preserve d(0 To values.children.Count - 1, 0 To jor.children.Count - 1)
                End If
                For Each joc In jor.children
                    d(jor.childIndex - 1, joc.childIndex - 1) = joc.value
                Next joc
            Next jor
        End If
    End If
    getTableFromSheets = d
End Function



Public Sub putTable()
    Dim result As Boolean
    result = putTableToSheets(ActiveSheet.NAME, ActiveSheet.UsedRange.value)

End Sub

Public Sub getTable()
    Dim result As Variant
    Dim sheet As Worksheet
    
    result = getTableFromSheets(ActiveSheet.NAME)
    If UBound(result) > 0 Then
        '// now we just need to write it out
        Set sheet = ActiveSheet
        ActiveSheet.Cells(1, 1).Resize(UBound(result, 1) + 1, UBound(result, 2) + 1).value = result
    End If

End Sub
