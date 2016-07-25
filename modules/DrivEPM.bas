Attribute VB_Name = "DrivePM"


Option Explicit


'// this is about setting up for OAUTH2
'// all you have to do is to create a google console developer project
'// create some credentials of app type other
'// activeate the sheets api
'// copy the credentials to config file
'// and run this function once.
'// you can delete it once you;ve run it - its not needed any more if you successfully go through the google auth process


Public Sub RunOnce()
    'your config here
    
    'SaveConfig "config.txt", "SheetID", "xxx"
    'SaveConfig "config.txt", "ClientID", "xxxx.apps.googleusercontent.com"
    'SaveConfig "config.txt", "ClientSecret", "xxx"
    
    Dim ClientID As String
    Dim ClientSecret As String
    
    If GetConfig("config.txt", "ClientID", ClientID) And GetConfig("config.txt", "ClientSecret", ClientSecret) Then
        getGoogled "sheets", , ClientID, ClientSecret
    Else
        MsgBox ("failed on getting config for oauth clientid or/and clientsecret")
        Exit Sub
    End If
    
End Sub


Function getMySheetId() As String
    Dim SheetID As String
    If GetConfig("config.txt", "SheetID", SheetID) Then
        getMySheetId = SheetID
    Else
        MsgBox ("failed on getting config for sheetid")
        Exit Function
    End If
End Function



'// below here are the useful functions you'll need
Private Function sheetExistsAtGoogle(sheetAccess As cSheetsV4, sheetName As String, Optional complain As Boolean = False) As cJobject
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


Private Function putTableToSheets(id As String, table As String, values As Variant) As cJobject
    Dim result As cJobject, sheetAccess As cSheetsV4
    '// set up an access object
    Set sheetAccess = New cSheetsV4
    sheetAccess.setAuthName("sheets").setSheetId (id)
    '// put the sheet data .. will bacth this up later
    ' see if it exists at google
    If Not (sheetExistsAtGoogle(sheetAccess, table) Is Nothing) Then

        '// and write it
        Set result = sheetAccess.setValues(values, table, A1Notation(values))
        If (Not result.child("success").value) Then
            MsgBox ("failed on sheets API " + result.child("response").stringify)
        End If
    End If
    Set putTableToSheets = result
End Function


'/**
' * get the sheet contents from google
' * @param {string} id the spreadsheet id
' * @param {string} [sheet]
' * @return {cjobject} the result
'*/
Private Function getTableFromSheets(id As String, table As String) As cJobject
    Dim result As cJobject, sheetAccess As cSheetsV4
    '// set up an access object
    Set sheetAccess = New cSheetsV4
    sheetAccess.setAuthName("sheets").setSheetId (id)
    '// get the data for the sheets that exist or were asked for
    '// see if they exist at google
    If Not (sheetExistsAtGoogle(sheetAccess, table) Is Nothing) Then
        '// now lets get the data in the selected sheets
        Set result = sheetAccess.getValues(table)
        If (Not result.child("success").value) Then
            MsgBox ("failed on sheets API " + result.child("response").stringify)
        End If
    End If
    Set getTableFromSheets = result
End Function

Private Function writeToSheets(data As cJobject)
    Dim d As Variant, sheet As Worksheet, values As cJobject, jor As cJobject, joc As cJobject
        '// now write the data
            Set values = data.children(1).child("values")
            ReDim d(0 To values.children.Count - 1, 0 To values.children(1).children.Count - 1)
            For Each jor In values.children
                For Each joc In jor.children
                    d(jor.childIndex - 1, joc.childIndex - 1) = joc.value
                Next joc
            Next jor
            
            '// now we just need to write it out
            Set sheet = ActiveSheet
            sheet.Cells(1, 1) _
            .Resize(values.children.Count, values.children(1).children.Count) _
            .value = d
End Function ' check if this childExists in current children

Public Sub putTable()
    Dim result As cJobject
    Set result = putTableToSheets(getMySheetId(), ActiveSheet.NAME, ActiveSheet.UsedRange.value)
    If (Not result Is Nothing) And (Not result.child("success").value) Then
        MsgBox ("failed on sheets API " + result.child("response"))
        Exit Sub
    End If
End Sub

Public Sub getTable()
    Dim result As cJobject
    
    Set result = getTableFromSheets(getMySheetId(), ActiveSheet.NAME)
    If (Not result.child("success").value) Then
        MsgBox ("failed on sheets API " + result.child("response"))
        Exit Sub
    End If
    writeToSheets result.child("data").children(1).child("valueRanges")
End Sub
