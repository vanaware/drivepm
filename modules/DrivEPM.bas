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



Public Sub teste()
    Dim sheetAccess As cSheetsV4
    Set sheetAccess = New cSheetsV4
    Dim result As New cJobject
    
    sheetAccess.setAuthName("sheets").setSheetId (getMySheetId())
    
    'Debug.Print sheetAccess.getTables.stringify
    With sheetAccess.addTable("Teste3!a1:z1000")
        With .add.addArray
            .add , "teste 1 - linha 1"
            .add , "teste 2 - linha 1"
        End With
        With .add.addArray
            .add , "teste 1 - linha 2"
            .add , "teste 2 - linha 2"
        End With
    End With
    'Debug.Print sheetAccess.getTables.stringify

    sheetAccess.fixTablesName
    'Debug.Print sheetAccess.getTables.stringify
    
    With sheetAccess.getTable("Teste3")
        With .add.addArray
            .add , "teste 1 - linha 3"
            .add , "teste 2 - linha 3"
        End With
        With .add.addArray
            .add , "teste 1 - linha 4"
            .add , "teste 2 - linha 4"
        End With
    End With

    'Debug.Print sheetAccess.getTables.stringify
    
    With sheetAccess.addTable("Teste2")
        With .add.addArray
            .add , "teste2 1 - linha 1"
            .add , "teste2 2 - linha 1"
        End With
        With .add.addArray
            .add , "teste2 1 - linha 2"
            .add , "teste2 2 - linha 2"
        End With
    End With
    Debug.Print sheetAccess.getTables.stringify
    

    Set result = sheetAccess.backupTables()
    
    Set result = sheetAccess.restoreTables()
    
    'sheetAccess.removeTable ("teste")
    Debug.Print sheetAccess.getTables.stringify

End Sub
