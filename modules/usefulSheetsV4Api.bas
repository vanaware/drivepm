Attribute VB_Name = "usefulSheetsV4Api"
Option Explicit
'// useful functions to go along with the cSheetsV4 class for accessing the Sheets API

'//DEMOS
Private Function getMySheetId() As String
'// make this into your own sheet id
'// simplest way to implement in a form would be to make a picker for google drive
    getMySheetId = "1V54F5b1e1bOcJXJ_McRkpyJ5Dx_ndGnjdiZpBeyA4L0"
End Function
'// this is about setting up for OAUTH2
'// all you have to do is to create a google console developer project
'// create some credentials of app type other
'// activeate the sheets api
'// copy the credentials to the below function (the ones here wont work- they are deactivated)
'// and run this function once.
'// you can delete it once you;ve run it - its not needed any more if you successfully go through the google auth process
Private Function sheetsOnceOff()
    
    getGoogled "sheets", , _
    "109xxfbhq.apps.googleusercontent.com", _
    "UxxCl"
    


End Function
'// this wil get the active sheet from google
'// go can replace the active sheet name with some other sheets, or a list of sheets separated by commas
Public Sub getThisSheet()
    Dim result As cJobject
    
    Set result = getStuffFromSheets(getMySheetId(), ActiveSheet.NAME)
    If (Not result.child("success").value) Then
        MsgBox ("failed on sheets API " + result.child("response"))
        Exit Sub
    End If
    writeToSheets result.child("data").children(1).child("valueRanges"), True
End Sub
'// this puts the active sheet to google
'// go can replace the active sheet name with some other sheets, or a list of sheets separated by commas
Public Sub putThisSheet()
    Dim result As cJobject
    Set result = putStuffToSheets(getMySheetId(), ActiveSheet.NAME)
    If (Not result.child("success").value) Then
        MsgBox ("failed on sheets API " + result.child("response"))
        Exit Sub
    End If
End Sub
'// below here are the useful functions you'll need
Private Function sheetExistsAtGoogle(results As cJobject, sheetName As Variant, Optional complain As Boolean = False) As cJobject

    Dim theSheet As String, s As String, job As cJobject
    theSheet = LCase(CStr(sheetName))
    
    For Each job In results.child("data").children(1).child("sheets").children
        s = job.toString("properties.title")
        If (LCase(s) = theSheet) Then
            Set sheetExistsAtGoogle = job
            Exit Function
        End If
    Next job
    If (complain) Then
        MsgBox ("sheet does not exist on Google " & theSheet)
    End If
    Set sheetExistsAtGoogle = Nothing
End Function
Public Function putStuffToSheets(id As String, sheetNames As String) As cJobject

    Dim names As Variant, i As Long, _
        results As cJobject, sheet As Worksheet, _
        r As Range, values As Variant, result As cJobject, sheetAccess As cSheetsV4
    
    names = Split(sheetNames, ",")
    
    '// set up an access object
    Set sheetAccess = New cSheetsV4
    sheetAccess.setAuthName("sheets").setSheetId (id)
    
    '// get all the sheets at  the google end
    Set results = sheetAccess.getSheets()
    If (Not results.child("success").value) Then
        MsgBox "failed getting sheets meta data " & results.toString("response")
        Set putStuffToSheets = results
        Exit Function
    End If
    
    '// put the sheet data .. will bacth this up later
    For i = LBound(names) To UBound(names)
        With sheetExists(names(i), True)
            ' see if it exists at google
            If (sheetExistsAtGoogle(results, names(i)) Is Nothing) Then
              ' we need to create it
                Set result = sheetAccess.insertSheet(names(i))
                If (Not result.child("success").value) Then
                    MsgBox "failed inserting sheet " & result.toString("response")
                    Set putStuffToSheets = result
                    Exit Function
                End If
            End If

            
            '// and write it
            Set result = sheetAccess.setValues(.UsedRange.value, .NAME, .UsedRange.Address)
            If (Not result.child("success").value) Then
                MsgBox ("failed on sheets API " + result.child("response").stringify)
            End If
        End With

    Next i
    
    Set putStuffToSheets = result

End Function
Private Function writeToSheets(data As cJobject, Optional overwrite As Boolean = False)
    Dim job As cJobject, a As Variant, d As Variant, _
      sheetName As String, sheet As Worksheet, ov As Boolean, _
      values As cJobject, jor As cJobject, joc As cJobject
      
    For Each job In data.children
        a = Split(LCase(job.toString("range")), "!")
        sheetName = CStr(a(0))
        '// see if the sheet exists
        Set sheet = sheetExists(sheetName)
        If (isSomething(sheet)) Then
            If (Not overwrite) Then
                ov = MsgBox("sheet already exists " & sheetName & " overwrite?", vbYesNo)
            End If
            If (ov Or overwrite) Then
                sheet.Cells.ClearContents
                
            Else
                Set sheet = Nothing
            End If
        Else
            Set sheet = sheets.add
            sheet.NAME = sheetName
        End If

        '// now write the data
        If (isSomething(sheet)) Then
            Set values = job.child("values")
            ReDim d(0 To values.children.Count - 1, 0 To values.children(1).children.Count - 1)
            For Each jor In values.children
                For Each joc In jor.children
                    d(jor.childIndex - 1, joc.childIndex - 1) = joc.value
                Next joc
            Next jor
            
            '// now we just need to write it out
            sheet.Cells(1, 1) _
            .Resize(values.children.Count, values.children(1).children.Count) _
            .value = d
            
        End If
        
    Next job
    
End Function ' check if this childExists in current children

Private Function sheetExists(sname As Variant, Optional complain As Boolean = False) As Worksheet
    Dim sheetName As String

    On Error GoTo handle
    sheetName = CStr(sname)
    Set sheetExists = sheets(sheetName)
    Exit Function
handle:
    Set sheetExists = Nothing
    If (complain) Then
        MsgBox ("sheet " & sheetName & " doesnt exist")
    End If
End Function
'/**
' * get the sheet contents from google
' * @param {string} id the spreadsheet id
' * @param {string} [sheets] separated buy commas - blank gets them all
' * @return {cjobject} the result
'*/
Private Function getStuffFromSheets(id As String, _
            Optional sheets As String = vbNullString) As cJobject
    Dim oauth As cOauth2
    Dim sheetAccess As cSheetsV4
    Dim results As cJobject, a As Variant, theSheet As String, _
      i As Long, job As cJobject, ranges As cStringChunker, c As cStringChunker


    '// set up an access object
    Set sheetAccess = New cSheetsV4
    sheetAccess.setAuthName("sheets").setSheetId (id)
    
    '// get what sheets exist
    Set results = sheetAccess.getSheets()
    If (Not results.child("success").value) Then
        MsgBox "failed getting sheets meta data " & results.toString("response")
        Set getStuffFromSheets = results
        Exit Function
    End If
    
    
    '// here's what they are
    '//Debug.Print JSONStringify(results.child("data"))
    
    '// get the data for the sheets that exist or were asked for
    '// see if they exist at google
    Set c = New cStringChunker
    If (sheets <> vbNullString) Then
        a = Split(sheets)
        For i = LBound(a) To UBound(a)
          If (isSomething(sheetExistsAtGoogle(results, a(i), True))) Then
            c.add(CStr(a(i))).add (",")
          End If
        Next i
        sheets = c.chopWhile(",").toString
    End If
    
    sheets = LCase(sheets) + ","
    
    Set ranges = New cStringChunker
    For Each job In results.child("data").children(1).child("sheets").children
        theSheet = job.toString("properties.title")
        If (sheets = "," Or InStr(1, sheets, LCase(theSheet) & ",") > 0) Then
          ranges.add(theSheet).add("!a:z").add (",")
        End If
    Next job

    '// now lets get the data in the selected sheets
    Set getStuffFromSheets = sheetAccess.getValues(ranges.chopWhile(",").toString)
     
End Function


