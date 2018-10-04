VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mserv_Journal_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mySpace As Outlook.NameSpace
'Wer setzt diese Variable?
Rem diese Variable sollte im initialize gesetzt werden
Private myRootfolder As MAPIFolder
Private myWorkfolder As MAPIFolder

Private journal_OBJ As JournalItem
Private journal_MOD As JournalModule



Sub debug_me()

 Call writeLogging

End Sub

Public Function writeLogging() As String

Dim logItem_OBJ As JournalItem

    Set logItem_OBJ = JournalFolder.Items.Add(olJournalItem)
    logItem_OBJ.Type = "Note"
    
    Debug.Print "writeLogging.logItem_OBJ.Type=" & logItem_OBJ.Type

End Function


Private Function JournalFolder() As Outlook.Folder

    Dim folder_OBJ As Outlook.Folder
    Dim returnvalue_OBJ As Outlook.Folder
    
    Set mySpace = Application.GetNamespace("MAPI")
    
    Set myRootfolder = mySpace.Folders.Item("norman.gruener@telefonica.com").Folders.Item("Posteingang").Folders.Item("06_Datenbank").Folders.Item("12_Eventsource")
    Set myWorkfolder = myRootfolder.Folders.Item("Eventsource")
    
    Set JournalFolder = myWorkfolder
    

End Function

