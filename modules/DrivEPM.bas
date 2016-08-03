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


Sub TesteProject()

    Dim MyProject As New cProject

    MyProject.ActivePrj = ActiveProject
    MyProject.pushtoTables
    MyProject.backupTables
    
    'Free Memory
    MyProject.Teardown
    Set MyProject = Nothing
End Sub


