VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_OutlookFolder_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem Klassenmodul aus Visual Basic for Applications / Outlook
Rem _______________________________________________________________________
Rem http://msdn.microsoft.com/en-us/library/office/ff865587(v=office.14).aspx

Option Explicit
Rem wird aktuell noch verwendet
Private user_NAME As String
Private folderhierarchy_MTX(3) As String

Rem ______________________________________________________________________
Private workSlot_OBJ As New edv_Workslot_handler
Private staging_OBJ As New edv_FilesystemFolder_wrapper
Rem wird verwendet
Private mySpace As Outlook.NameSpace
'Wer setzt diese Variable?
Rem diese Variable sollte im initialize gesetzt werden
Private myRootfolder As MAPIFolder
Private myWorkfolder As MAPIFolder
Private objMail As MailItem
Rem das konzept ist allerdings schon etwas älter:
Private stagingFileName_TXT(28) As String
Rem ______________________________________________________________________
Rem Konstanten
Private Const keepMailAsUnread As Boolean = True
Private Const debug_FLAG As Boolean = True
Private Const version_DBL = 3
Rem öffentliche Variablen
Public moduleversion_INT As Integer
Public newMailInFolder_INT As Integer
'verwendung?
Public newattachments_INT As Integer


Rem _________________________________________________
Rem Public Functions

Public Sub debug_me()

    Dim result_OBJ As Object
    Dim debugWS_OBJ As New edv_Workslot_handler
    
    Call debugWS_OBJ.Workslotindex(vbWednesday, 4)
    Call TextfileImportFromWorkslot(debugWS_OBJ)
    
    If result_OBJ Is Nothing Then Debug.Print "result_OBJ=Nothing"
        
    Debug.Print "debug_me"

End Sub

Public Function TextfileImportFromWorkslot(workSlot_ARG As edv_Workslot_handler, Optional execution_ARG As Integer = 9) As edv_Textfileimport_wrapper
Rem Diese Funktion ist sinnvoll, wenn Textdateien weiterverabeitet werden sollen
Rem daher sollte sie ausführlich verwendet werden
Rem zB .txt und .csv
Rem verwendet die Funktionen:
Rem
Dim resultvalue_OBJ As New edv_Textfileimport_wrapper
Dim source_OBJ As MailItem
Rem für das finden des richtigen Attachments sind nützlich:
Rem das Datum an dem die Mail ankam
Dim cr_DAT As Date
Dim date_PFX As String
Rem der Betreff der Mail
Dim mySubject_TXT As String
Rem der Name des Attachments:
Dim myfilename_TXT As String
Rem
Dim mailBodyText_PATH As String
Dim fileSequence_IDX As Integer
Dim IsNoAttachment_FLAG As Boolean
Dim SaveMailBody_FLAG As Boolean
Dim UseTime_FLAG As Boolean

On Error GoTo Errorhandler

    
    Rem initialize variables:
    fileSequence_IDX = 0
    IsNoAttachment_FLAG = True
    SaveMailBody_FLAG = False
    
    Rem ---------------------------------------------------------------------------------------------------------
    Rem die steuerung dessen, was in dieser Funktion passiert, wird vom execution_ARG übernommen
    Rem execution_ARG = 0: nur auf neue Mails prüfen
    Rem execution_ARG = 1: auf neue Mails mit attachment prüfen
    Debug.Print "edv.OutlookFolder_wrapper.TextfileImportFromWorkslot.execution_ARG=" & execution_ARG

    Set resultvalue_OBJ = Nothing

    If (myWorkfolder Is Nothing) Then
        Rem workslotfolder not yet defined, must be done now:
        Set myWorkfolder = WorkslotFolder_OBJ(workSlot_ARG)
    End If
    
    Rem ---------------------------------------------------------------------------------------------------------
    Rem Zwei neue Flags für die Nutzung mit Mails, die a) häufiger am Tag kommen und b) wo der Inhalt wichtiger ist
    Rem als das Attachment (weil z.B. keines vorhanden)
    Rem wird nur benötigt, wenn eine speicherung erwartet wird (execution ARG größer 1)
    If execution_ARG > 1 Then
        SaveMailBody_FLAG = True
        UseTime_FLAG = True
    End If
    
    Rem Hinweis: Wenn keine neuen Mails vorhanden sind, läuft die Schleife nicht durch, sondern wirft
    Rem sofort einen Fehler.
    Rem dieser wird dann im Errorhandler bewertet und darauf reagiert
    Rem wenn bei den Discriminatoren kein Preset aufaddiert wurde,
    Rem läuft die Schleife genau über die Anzahl der gefundenen Mails
    For fileSequence_IDX = 1 To newMailInFolder_INT + workSlot_ARG.PresetFile_CNT

        Set source_OBJ = NewMailFromFolder(myWorkfolder, Not keepMailAsUnread)

        cr_DAT = source_OBJ.CreationTime
    Rem Datumsreihenfolge im Prefix: MMDD
        date_PFX = Mid(cr_DAT, 4, 2) & Left(cr_DAT, 2)
        date_PFX = date_PFX & Mid(cr_DAT, 12, 2) & Mid(cr_DAT, 15, 2)
        mySubject_TXT = source_OBJ.Subject
        IsNoAttachment_FLAG = (source_OBJ.Attachments.Count = 0)
        'das hier näher analysieren
        'Call resultvalue_OBJ.textFileGenerated_Path(workSlot_ARG)
        mailBodyText_PATH = resultvalue_OBJ.textFileGenerated_Path(workSlot_ARG)
        
        Rem : Funktion zum abspeichern des Mailinhalts als Text
        If SaveMailBody_FLAG Then
            'Das hier speichert als reinen Text. Problematisch mit html mails
            'Call source_OBJ.SaveAs("C:\Users\Public\Documents\507\" & Left(cr_DAT, 5) & ".txt", olTXT)
            CreateObject("scripting.filesystemobject").CreateTextFile("C:\Users\Public\Documents\304_CT2020_FixedVoice\" & date_PFX & ".txt").Write source_OBJ.Body
  
        End If
        
        
        
    Rem bis hier sind die funktionen sehr ähnlich.
    Rem ___________________________________________________________________________________________________________
    Rem Führe dies nur durch, wenn attachment vorhanden ist und es abgespeichert werden soll
        
        If Not IsNoAttachment_FLAG And execution_ARG > 1 Then
            myfilename_TXT = source_OBJ.Attachments.Item(1).FileName
        
    Rem der textFileFullPath liefert nur ohne den staging pfad
        'myfilename_TXT = resultvalue_OBJ.textFileFullPath(workSlot_ARG) & myfilename_TXT
    
    Rem __________________________________________________________________________________________________________
    Rem nutzt textFileGenerated_Path
    Rem diese Funktion generiert Dateinamen nach einem intelligenten Scheam
    Rem generell ist das eine gute Idee, aber für den Fall, daß das Attachment noch gezippt ist, wäre das problematisch.
    Rem der Textfileimport_wrapper betreibt eine eigene Stagingmatrix.
            If Right(myfilename_TXT, 3) = "txt" Then myfilename_TXT = resultvalue_OBJ.textFileGenerated_Path(workSlot_ARG, fileSequence_IDX, cr_DAT)
            Call source_OBJ.Attachments.Item(1).SaveAsFile(myfilename_TXT)
        End If

            
    Next fileSequence_IDX
    
Errorhandler:

        If newMailInFolder_INT > 0 Then Debug.Print "edv_OutlookFilder_wrapper.source_OBJ.SenderName=" & source_OBJ.SenderName
        If newMailInFolder_INT = 0 Then Call resultvalue_OBJ.newMail_FLAG(False)
  
        Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.IsNoAttachment_FLAG=" & IsNoAttachment_FLAG

    If Err.Number <> 0 Then
        Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.Err.Number=" & Err.Number
        Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.Err.Description=" & Err.Description
        Debug.Print "edv.OutlookFolder_wrapper.TextfileImportFromWorkslot.SaveMailBody_FLAG=" & SaveMailBody_FLAG
        Debug.Print "edv_OutlookFolder_wrapper.newMailInFolder_INT=" & newMailInFolder_INT
        Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.cr_DAT=" & cr_DAT
        Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.newMail_IDX=" & fileSequence_IDX & " out of " & newMailInFolder_INT
        Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.source_OBJ.Attachments.Count=" & source_OBJ.Attachments.Count
        Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.source_OBJ.Attachments.myfilename_TXT=" & myfilename_TXT
        Debug.Print "edv.OutlookFolder_wrapper.TextfileImportFromWorkslot.UseTime_FLAG=" & UseTime_FLAG
    End If

    Set TextfileImportFromWorkslot = resultvalue_OBJ

End Function

'Public Function worksheetImportFromWorkslot(workSlot_ARG As edv_Workslot_handler) As edv_Worksheet_wrapper
'
'Const keepMailAsUnread As Boolean = True
'Rem zum Einschalten der Debug.Prints hier auf "True" setzen!
'Const debug_FLAG As Boolean = False
'Dim resultvalue_OBJ As New edv_Worksheet_wrapper
'Dim source_OBJ As MailItem
'Rem für das finden des richtigen Attachments sind nützlich:
'Rem das Datum an dem die Mail ankam
'Dim cr_DAT As Date
'Rem der Betreff der Mail
'Dim mySubject_TXT As String
'Rem der Name des Attachments:
'Dim myfilename_TXT As String
'Rem anzahl gefundener neuer Mails
'Dim newMail_CNT As Integer
'Dim fileSequence_IDX As Integer
'Dim IsNoAttachment_FLAG As Boolean
'
'On Error GoTo Errorhandler
'    Rem initialize variables:
'    newMail_CNT = -1
'    fileSequence_IDX = 0
'    IsNoAttachment_FLAG = True
'    Set resultvalue_OBJ = Nothing
'
'    If (myWorkfolder Is Nothing) Then
'        Rem workslotfolder not yet defined, must be done now:
'        Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.myWorkfolder Is Nothing"
'        Set myWorkfolder = WorkslotFolder_OBJ(workSlot_ARG)
'    End If
'
'    Rem wenn bei den Discriminatoren kein Preset aufaddiert wurde,
'    Rem läuft die Schleife genau über die Anzahl der gefundenen Mails
'    newMail_CNT = newMail_CNT + myWorkfolder.UnReadItemCount + workSlot_ARG.PresetFile_CNT
'    If newMail_CNT <> newMailsInFolder_INT Then
'        Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.newMail_CNT <> newMailsInFolder_INT"
'        Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.newMail_CNT=" & newMail_CNT
'        Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.newMailsInFolder_INT=" & newMailsInFolder_INT
'    End If
'
'
'    For fileSequence_IDX = 0 To newMail_CNT
'
'        Set source_OBJ = NewMailFromFolder(myWorkfolder, Not keepMailAsUnread)
'
'        cr_DAT = source_OBJ.CreationTime
'        mySubject_TXT = source_OBJ.Subject
'
'        IsNoAttachment_FLAG = (source_OBJ.Attachments.Count = 0)
'
'        myfilename_TXT = resultvalue_OBJ.workslotPathByWorkslot(workSlot_ARG)
'        myfilename_TXT = resultvalue_OBJ.worksheet_PATH(source_OBJ.Attachments.Item(1).FileName)
'
'        Call source_OBJ.Attachments.Item(1).SaveAsFile(myfilename_TXT)
'
'    Next fileSequence_IDX
'
'Errorhandler:
'
'If Err.Number <> 0 Or debug_FLAG Then
'    Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.IsNoAttachment_FLAG=" & IsNoAttachment_FLAG
'    Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.myfilename_TXT=" & myfilename_TXT
'    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.Err.Number=" & Err.Number
'    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.debug_FLAG=" & debug_FLAG
'    Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.newMail_CNT=" & newMail_CNT
'    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.source_OBJ.SenderName=" & source_OBJ.SenderName
'    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.cr_DAT=" & cr_DAT
'    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.newMail_IDX=" & fileSequence_IDX & " out of " & newMail_CNT
'    Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.source_OBJ.Attachments.Count=" & source_OBJ.Attachments.Count
'End If
'
'    Set worksheetImportFromWorkslot = resultvalue_OBJ
'
'
'End Function
'


Public Function pdfImportFromWorkslot(workSlot_ARG As edv_Workslot_handler) As String
Rem so einfach kann das sein....
    Dim source_OBJ As MailItem
    
    Set source_OBJ = NewMailFromFolder(WorkslotFolder_OBJ(workSlot_ARG))
    
    source_OBJ.Attachments.Item(1).SaveAsFile "C:\Users\Public\Documents\202_Benchmark\" & source_OBJ.Attachments.Item(1).FileName
    
    pdfImportFromWorkslot = "C:\Users\Public\Documents\202_Benchmark\" & source_OBJ.Attachments.Item(1).FileName

End Function


Public Function zipMicroServiceFromWorkslot(workSlot_ARG As edv_Workslot_handler) As mserv_zipFile_wrapper

On Error GoTo Errorhandler
    
    Dim source_OBJ As MailItem
    Dim result_OBJ As mserv_zipFile_wrapper
    Dim zipfile_NAME As String
    Dim isSourceOBJ As Boolean
    
    isSourceOBJ = False
        
    Set result_OBJ = New mserv_zipFile_wrapper
    Set source_OBJ = NewMailFromFolder(WorkslotFolder_OBJ(workSlot_ARG))
    
    isSourceOBJ = Not (source_OBJ Is Nothing)
    
    If isSourceOBJ Then
        zipfile_NAME = source_OBJ.Attachments.Item(1).FileName
        Debug.Print "edv_OutlookFolder_wrapper.result_OBJ.IsZipFileName(zipfile_NAME)=" & result_OBJ.IsZipFileName(zipfile_NAME)
        'source_OBJ.Attachments.Item(1).SaveAsFile
        Debug.Print "edv_OutlookFolder_wrapper.result_OBJ.IsZipFileName=" & result_OBJ.IsZipFileName
    End If
    
Errorhandler:

If debug_FLAG Then
    
    Debug.Print "edv_OutlookFolder_wrapper.zipMicroServiceFromWorkslot.Err.Number=" & Err.Number
    Debug.Print "edv_OutlookFolder_wrapper.zipMicroServiceFromWorkslot.Err.Description=" & Err.Description
    Debug.Print "edv_OutlookFolder_wrapper.zipMicroServiceFromWorkslot.workslot_ARG=" & workSlot_ARG.Workslotindex
    Debug.Print "edv_OutlookFolder_wrapper.zipMicroServiceFromWorkslot.isSourceOBJ=" & isSourceOBJ
    Debug.Print "edv_OutlookFolder_wrapper.zipfile_NAME=" & zipfile_NAME
    Debug.Print "edv_OutlookFolder_wrapper.zipMicroServiceFromWorkslot: TODO: source_OBJ.Attachments.Item(1).SaveAsFile workSlot_ARG"
    Debug.Print "edv_OutlookFolder_wrapper.zipMicroServiceFromWorkslot"
    
End If

Set zipMicroServiceFromWorkslot = result_OBJ
    
End Function

Rem ______________________________________________________________________________________________________


'Public Function StoreAttachment(Optional workSlot_ARG As edv_Workslot_handler) As edv_PowerPoint_wrapper
'Rem Wird von import_CCHWKL_mediator verwendet
'Rem
''da das hier nur PowerPoint als Attachment zurückliefert, sollte die Funktion auch so benannt werden, oder?
'On Error GoTo Errorhandler
'
''Const debug_FLAG As Boolean = True
'Const debug_FLAG As Boolean = False
'Rem die Zurückgelieferte Klasse:
'Dim resultobject_OBJ As New edv_PowerPoint_wrapper
'Rem das Objekt vom Typ MailItem, das im Grunde genommen ausgewertet wird
'Dim source_OBJ As MailItem
''was genau ist das?
'Dim new_FLAG As Boolean
'Rem der Name des Attachments:
'Dim myfilename_TXT As String
'
'Dim Anlagen_INT As Integer
'Dim indx As Integer
'Dim cr_DAT As Date
'
'    If debug_FLAG Then
'        Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.debug_FLAG=" & debug_FLAG
'        Set source_OBJ = NewMailFromFolder(WorkslotFolder_OBJ(workSlot_ARG), keepMailAsUnread)
'        Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.resultobject_OBJ.PresentationFull_PATH(" & "workslot_ARG" & ")=" & resultobject_OBJ.PresentationFull_PATH(workSlot_ARG)
'        Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.newMailsInFolder_INT+1=" & newMailsInFolder_INT + 1
'        source_OBJ.Attachments.Item(1).SaveAsFile resultobject_OBJ.addPresentationPath("", source_OBJ.Attachments.Item(1).FileName)
'        Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.Not debug_FLAG=" & Not debug_FLAG
'    End If
'
'If (myWorkfolder Is Nothing) Then
'    Set myWorkfolder = WorkslotFolder_OBJ(workSlot_ARG)
'End If
'
'myfilename_TXT = ""

'For Each objMail In myWorkfolder.Items
'
'    With objMail
'        new_FLAG = .UnRead
'        Anlagen_INT = .Attachments.Count
'        cr_DAT = .CreationTime
'
'        If new_FLAG Then
'            For indx = 1 To Anlagen_INT
'
'            Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.Alternative=" & True
'            'Alternative:
'            stagingFileName_TXT(newattachments_INT + 1) = .Attachments.Item(indx).FileName
'            myfilename_TXT = resultobject_OBJ.PresentationFull_PATH(workSlot_ARG)
'            myfilename_TXT = resultobject_OBJ.addPresentationPath("", stagingFileName_TXT(newattachments_INT + 1))
'            .Attachments.Item(indx).SaveAsFile myfilename_TXT
'             newattachments_INT = newattachments_INT + 1
'            Next indx
'        End If
'
'        .UnRead = False
'    End With
'Next objMail
'
'If newattachments_INT = 0 Then
'    stagingFileName_TXT(1) = "<no new attachments>"
'    Set resultobject_OBJ = Nothing
'End If
'
'Errorhandler:
'
'If (myWorkfolder Is Nothing) Or Err.Number <> 0 Or debug_FLAG Then
'    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.Now()=" & Now()
'    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.debug_FLAG=" & debug_FLAG
'    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.Err.Number=" & Err.Number
'    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.Err.Number=" & Err.Description
'    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.stagingfilename_TXT(1)=" & stagingFileName_TXT(1)
'    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.myfilename_TXT=" & myfilename_TXT
'    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.resultobject_OBJ="
'
'    'Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.myWorkfolder=Nothing"
'End If
'
'Set StoreAttachment = resultobject_OBJ
'
'End Function






Rem __________________________________________________________________________________


Private Function FolderFromEventsource(domainIndex_ARG As Integer) As Outlook.Folder

On Error GoTo Errorhandler

Dim folder_OBJ As Outlook.Folder
Dim returnvalue_OBJ As Outlook.Folder

Dim folder_cnt As Integer
Dim folderindex As Integer
Dim index_LEN As Integer

Set FolderFromEventsource = Nothing
Set returnvalue_OBJ = Nothing

    
Errorhandler:

    If Err.Number <> 0 Then
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.Err.Number=" & Err.Number
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.domainindex_ARG=" & domainIndex_ARG
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.myRootfolder.Name=" & myRootfolder.Name
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.returnvalue_OBJ.Name=" & returnvalue_OBJ.Name
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.returnvalue_OBJ.folder_cnt=" & folder_cnt
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.folderindex=" & folderindex
    End If
    
If Not (returnvalue_OBJ Is Nothing) Then

    Set FolderFromEventsource = returnvalue_OBJ
    
End If
    
End Function



Private Function WorkslotFolder_OBJ(workSlot_ARG As edv_Workslot_handler) As Outlook.Folder
Rem _________________________________________________________________________________________

Rem Diese Funktion liefert den Outlook Folder der zu einem bestimmetn Workslot gehört, zurück
Rem dazu werden die Ordner durchlaufen und bei Übereinstimmung mit dem Workslotprefix
Rem der Ordner zurückgeliefert

Rem wer ruft diese Funktion auf?
Rem StoreAttachment
Rem TextfileImportFromWorkslot
Rem worksheetImportFromWorkslot
Rem etc.etc.

On Error GoTo Errorhandler

Dim folder_OBJ As Outlook.Folder
Dim returnvalue_OBJ As Outlook.Folder

Dim folder_cnt As Integer
Dim folderindex As Integer

Set WorkslotFolder_OBJ = Nothing
Set returnvalue_OBJ = Nothing
'http://msdn.microsoft.com/en-us/library/office/ff865587(v=office.14).aspx


    folder_cnt = myRootfolder.Folders.Count
    For folderindex = 1 To folder_cnt

        Set myWorkfolder = myRootfolder.Folders.Item(folderindex)
        
        If workSlot_ARG.IsWorkslotprefix(myWorkfolder.Name) Then
            Set returnvalue_OBJ = myWorkfolder
            If (newMailInFolder_INT < 0) Then newMailInFolder_INT = returnvalue_OBJ.UnReadItemCount
            'newMailsInFolder_INT = returnvalue_OBJ.UnReadItemCount
            Exit For
            
        End If
        
    Next folderindex
    
Errorhandler:

If Err.Number <> 0 Or debug_FLAG Then
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.Err.Number=" & Err.Number
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.Err.Description=" & Err.Description
    Debug.Print "edv_OutlookFolder_wrapper.debug_FLAG=" & debug_FLAG
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.myRootfolder.Name=" & myRootfolder.Name
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.folder_cnt=" & folder_cnt
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.folderindex=" & folderindex
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.newMailInFolder_INT=" & newMailInFolder_INT
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.IsWorkslotprefix(" & myWorkfolder.Name & ")=" & workSlot_ARG.IsWorkslotprefix(myWorkfolder.Name)
End If
    
If Not (returnvalue_OBJ Is Nothing) Then

    'Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.returnvalue_OBJ.Name=" & returnvalue_OBJ.Name
    Set WorkslotFolder_OBJ = returnvalue_OBJ
    
End If
    
End Function

Private Function NewMailFromFolder(workfolder_ARG As Folder, Optional keepUnread_ARG As Boolean = False) As MailItem

On Error GoTo Errorhandler
Rem gibt ungelesene Mail aus Ordner zurück
Rem wichtige und relevante funktion
Rem wird von Textfileimport und von worksheetImport verwendet.


    For Each objMail In workfolder_ARG.Items

    With objMail
     
        If objMail.UnRead Then
            Set NewMailFromFolder = objMail
            objMail.UnRead = keepUnread_ARG
            Exit For
        End If
        
    End With
    Next objMail
    
Errorhandler:

Debug.Print "Err.Number=" & Err.Number
Debug.Print "Err.Description=" & Err.Description


End Function



Sub launch_Outlook()
'zu dokumentationszwecken, ist eigentlich unsinn, wenn outlook schon läuft

Dim var_DBL As Double
Dim outlook_PFAD As String
outlook_PFAD = "C:\Program Files\Microsoft Office\Office14\"

var_DBL = Shell(outlook_PFAD & "Outlook.exe", vbMaximizedFocus)

End Sub

Sub sendMessage(successmessage_ARG As String)

    Rem Send a mail using the DoCmd.SendObject method
    
    Rem i need a Object for handling the new mail:
    Dim newmail_OBJ As Object
    
    Rem i declare the Object to be a Mail Item:
    Set newmail_OBJ = Application.CreateItem(olMailItem)
    

    Rem i set the important properties:
    newmail_OBJ.To = "norman.gruener@telefonica.com"
    newmail_OBJ.Subject = "twitter"
    newmail_OBJ.Body = successmessage_ARG = successmessage_ARG & Chr(10)
    
    Rem i display it (optional, not necessary):
    Call newmail_OBJ.Display
    
    Rem i send it:
    Call newmail_OBJ.Send
End Sub


Private Sub Class_Initialize()
On Error GoTo Errorhandler

    Rem setzen der public variablen
    moduleversion_INT = 4
    newMailInFolder_INT = -1
    
    Rem setzen der Steuerungsvariablen
    
    Rem setzen objektspezifischer variablen
    user_NAME = "norman.gruener@telefonica.com"
    Set mySpace = Application.GetNamespace("MAPI")

    folderhierarchy_MTX(0) = "06_Datenbank"
    folderhierarchy_MTX(1) = "12_Eventsource"
    
    Set myRootfolder = mySpace.Folders.Item(user_NAME).Folders.Item("Posteingang").Folders.Item(folderhierarchy_MTX(0)).Folders.Item(folderhierarchy_MTX(1))
    Set myWorkfolder = Nothing
    
    Rem bei der Intialisierung weiss das Objekt noch nicht, welchen Workslot er erhalten wird.
    Rem daher kann leider hier der Workfolder noch nicht bestimmt werden
    
Errorhandler:
    If Err.Number <> 0 Then
        Debug.Print Err.Number
        Debug.Print "edv_OutlookFolder_wrapper.folderhierarchy_MTX(1)=" & folderhierarchy_MTX(1)
        Debug.Print "edv_OutlookFolder_wrapper.myWorkfolder=" & "Nothing"
    End If
    
    
End Sub

Private Sub Class_Terminate()

End Sub
