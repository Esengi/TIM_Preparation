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
'mal pr�fen, ob diese variablen noch alle ben�tigt werden!
Rem wird aktuell noch verwendet
Private user_NAME As String
Private folderhierarchy_MTX(3) As String
Rem diese globale variable ist dahingegen nicht gut programmiert:
Private dateprefix_TXT As String
'wer verwendet das hier?
'wird nur im initialize verwendet..
Private filename_TXT As String
Rem ______________________________________________________________________
Private workslot_OBJ As New edv_Workslot_handler
Rem wird verwendet
Private mySpace As Outlook.NameSpace
'Wer setzt diese Variable?
Rem diese Variable sollte im initialize gesetzt werden
Private myRootfolder As MAPIFolder
Private myWorkfolder As MAPIFolder
Private objMail As MailItem
Private newMailsInFolder_INT As Integer
Rem das konzept ist allerdings schon etwas �lter:
Private stagingFileName_TXT(28) As String
'wer verwendet das?
'keine weitere Verwendung
'au�er Initilize 11.09.17
Private folder_PATH(4) As String
Rem ______________________________________________________________________
Rem Konstanten
Private Const keepMailAsUnread As Boolean = True

'verwendung?
Public newattachments_INT As Integer

Rem _________________________________________________
Rem Public Functions



Public Function TextfileImportFromWorkslot(workslot_ARG As edv_Workslot_handler) As edv_Textfileimport_wrapper
Rem Diese Funktion ist sinnvoll, wenn Textdateien weiterverabeitet werden sollen
Rem daher sollte sie ausf�hrlich verwendet werden
Rem zB .txt und .csv
Rem verwendet die Funktionen:
Rem
Dim resultvalue_OBJ As New edv_Textfileimport_wrapper
Dim source_OBJ As MailItem
Rem f�r das finden des richtigen Attachments sind n�tzlich:
Rem das Datum an dem die Mail ankam
Dim cr_DAT As Date
Rem der Betreff der Mail
Dim mySubject_TXT As String
Rem der Name des Attachments:
Dim myfilename_TXT As String
Rem anzahl gefundener neuer Mails
Dim newMail_CNT As Integer
Dim fileSequence_IDX As Integer
Dim IsNoAttachment_FLAG As Boolean

Const debug_FLAG As Boolean = False

On Error GoTo Errorhandler
    Rem initialize variables:
    newMail_CNT = -1
    fileSequence_IDX = 0
    IsNoAttachment_FLAG = True
    Set resultvalue_OBJ = Nothing

    If (myWorkfolder Is Nothing) Then
        Rem workslotfolder not yet defined, must be done now:
        Set myWorkfolder = WorkslotFolder_OBJ(workslot_ARG)
    End If
    
    Rem wenn bei den Discriminatoren kein Preset aufaddiert wurde,
    Rem l�uft die Schleife genau �ber die Anzahl der gefundenen Mails
    newMail_CNT = newMail_CNT + myWorkfolder.UnReadItemCount + workslot_ARG.PresetFile_CNT
    Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.newMail_CNT=" & newMail_CNT
    Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.newMailsInFolder_INT+1=" & newMailsInFolder_INT + 1
    
    
    For fileSequence_IDX = 0 To newMail_CNT

        Set source_OBJ = NewMailFromFolder(myWorkfolder, Not keepMailAsUnread)

        cr_DAT = source_OBJ.CreationTime
        mySubject_TXT = source_OBJ.Subject

        IsNoAttachment_FLAG = (source_OBJ.Attachments.Count = 0)
        
    Rem bis hier sind die funktionen sehr �hnlich.
    Rem ___________________________________________________________________________________________________________
        
        myfilename_TXT = source_OBJ.Attachments.Item(1).FileName
    Rem der textFileFullPath liefert nur ohne den staging pfad
        myfilename_TXT = resultvalue_OBJ.textFileFullPath(workslot_ARG) & myfilename_TXT
    
    Rem __________________________________________________________________________________________________________
    Rem nutzt textFileGenerated_Path
    Rem diese Funktion generiert Dateinamen nach einem intelligenten Scheam
    Rem generell ist das eine gute Idee, aber f�r den Fall, da� das Attachment noch gezippt ist, w�re das problematisch.
    Rem der Textfileimport_wrapper betreibt eine eigene Stagingmatrix.
        If Right(myfilename_TXT, 3) = "txt" Then myfilename_TXT = resultvalue_OBJ.textFileGenerated_Path(workslot_ARG, fileSequence_IDX, cr_DAT)
        
        Call source_OBJ.Attachments.Item(1).SaveAsFile(myfilename_TXT)
    
    Next fileSequence_IDX
    
Errorhandler:

        Debug.Print "edv_OutlookFilder_wrapper.source_OBJ.SenderName=" & source_OBJ.SenderName
        Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.IsNoAttachment_FLAG=" & IsNoAttachment_FLAG

If Err.Number <> 0 Then
    Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.Err.Number=" & Err.Number
    Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.Err.Description=" & Err.Description
    Debug.Print "edv_OutlookFolder_wrapper.TextfileImportFromWorkslot.newMail_CNT=" & newMail_CNT
    Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.cr_DAT=" & cr_DAT
    Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.newMail_IDX=" & fileSequence_IDX & " out of " & newMail_CNT
    Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.source_OBJ.Attachments.Count=" & source_OBJ.Attachments.Count
    Debug.Print "edv_OutlookFilder_wrapper.TextfileImportFromWorkslot.source_OBJ.Attachments.myfilename_TXT=" & myfilename_TXT
End If

    Set TextfileImportFromWorkslot = resultvalue_OBJ

End Function

Public Function worksheetImportFromWorkslot(workslot_ARG As edv_Workslot_handler) As edv_Worksheet_wrapper

Const keepMailAsUnread As Boolean = True
Rem zum Einschalten der Debug.Prints hier auf "True" setzen!
Const debug_FLAG As Boolean = False
Dim resultvalue_OBJ As New edv_Worksheet_wrapper
Dim source_OBJ As MailItem
Rem f�r das finden des richtigen Attachments sind n�tzlich:
Rem das Datum an dem die Mail ankam
Dim cr_DAT As Date
Rem der Betreff der Mail
Dim mySubject_TXT As String
Rem der Name des Attachments:
Dim myfilename_TXT As String
Rem anzahl gefundener neuer Mails
Dim newMail_CNT As Integer
Dim fileSequence_IDX As Integer
Dim IsNoAttachment_FLAG As Boolean

On Error GoTo Errorhandler
    Rem initialize variables:
    newMail_CNT = -1
    fileSequence_IDX = 0
    IsNoAttachment_FLAG = True
    Set resultvalue_OBJ = Nothing

    If (myWorkfolder Is Nothing) Then
        Rem workslotfolder not yet defined, must be done now:
        Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.myWorkfolder Is Nothing"
        Set myWorkfolder = WorkslotFolder_OBJ(workslot_ARG)
    End If
    
    Rem wenn bei den Discriminatoren kein Preset aufaddiert wurde,
    Rem l�uft die Schleife genau �ber die Anzahl der gefundenen Mails
    newMail_CNT = newMail_CNT + myWorkfolder.UnReadItemCount + workslot_ARG.PresetFile_CNT
    If newMail_CNT <> newMailsInFolder_INT Then
        Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.newMail_CNT <> newMailsInFolder_INT"
        Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.newMail_CNT=" & newMail_CNT
        Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.newMailsInFolder_INT=" & newMailsInFolder_INT
    End If
    
    
    For fileSequence_IDX = 0 To newMail_CNT

        Set source_OBJ = NewMailFromFolder(myWorkfolder, Not keepMailAsUnread)

        cr_DAT = source_OBJ.CreationTime
        mySubject_TXT = source_OBJ.Subject
            
        IsNoAttachment_FLAG = (source_OBJ.Attachments.Count = 0)

        myfilename_TXT = resultvalue_OBJ.workslotPathByWorkslot(workslot_ARG)
        myfilename_TXT = resultvalue_OBJ.worksheet_PATH(source_OBJ.Attachments.Item(1).FileName)
          
        Call source_OBJ.Attachments.Item(1).SaveAsFile(myfilename_TXT)
    
    Next fileSequence_IDX

Errorhandler:

If Err.Number <> 0 Or debug_FLAG Then
    Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.IsNoAttachment_FLAG=" & IsNoAttachment_FLAG
    Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.myfilename_TXT=" & myfilename_TXT
    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.Err.Number=" & Err.Number
    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.debug_FLAG=" & debug_FLAG
    Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.newMail_CNT=" & newMail_CNT
    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.source_OBJ.SenderName=" & source_OBJ.SenderName
    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.cr_DAT=" & cr_DAT
    Debug.Print "edv_OutlookFilder_wrapper.worksheetImportFromWorkslot.newMail_IDX=" & fileSequence_IDX & " out of " & newMail_CNT
    Debug.Print "edv_OutlookFolder_wrapper.worksheetImportFromWorkslot.source_OBJ.Attachments.Count=" & source_OBJ.Attachments.Count
End If

    Set worksheetImportFromWorkslot = resultvalue_OBJ


End Function

Public Function ControlFromWorkslot(workslot_ARG As edv_Workslot_handler) As edv_Control_wrapper

Dim resultvalue_OBJ As New edv_Control_wrapper

On Error GoTo Errorhandler
    
    If (myWorkfolder Is Nothing) Then
        Rem workslotfolder not yet defined, must be done now:
        Set myWorkfolder = WorkslotFolder_OBJ(workslot_ARG)
    End If
    
    Call resultvalue_OBJ.addControlItem(1, "newMail_CNT", 1)
    
Errorhandler:

If Err.Number <> 0 Then

    Debug.Print "MailSubjectFromDomain.Err.Description=" & Err.Description

End If


End Function

Public Function pdfImportFromWorkslot(workslot_ARG As edv_Workslot_handler) As String
Rem so einfach kann das sein....
    Dim source_OBJ As MailItem
    
    Set source_OBJ = NewMailFromFolder(WorkslotFolder_OBJ(workslot_ARG))
    
    source_OBJ.Attachments.Item(1).SaveAsFile "C:\Users\Public\Documents\202_Benchmark\" & source_OBJ.Attachments.Item(1).FileName
    
    pdfImportFromWorkslot = "C:\Users\Public\Documents\202_Benchmark\" & source_OBJ.Attachments.Item(1).FileName

End Function


Public Function zipMicroServiceFromWorkslot(workslot_ARG As edv_Workslot_handler) As mserv_zipFile_wrapper

    Dim source_OBJ As MailItem
    Dim result_OBJ As mserv_zipFile_wrapper
    
    Set source_OBJ = NewMailFromFolder(WorkslotFolder_OBJ(workslot_ARG))
    
    If Not (source_OBJ Is Nothing) Then
        Debug.Print "edv.OutlookFolder_wrapper.zipMicroServiceFromWorkslot.source_OBJ<>Nothing"
        source_OBJ.Attachments.Item(1).SaveAsFile result_OBJ.zipFileGenerated_Path(workslot_ARG)
    
    End If
    
    Debug.Print "edv.OutlookFolder_wrapper.zipMicroServiceFromWorkslot"

End Function

Function FilenameFromStaging(outlookdomain_ARG As String, stagingpfad_ARG As String, Optional index_ARG As Integer = 1, Optional filename_ARG As String = "") As String
On Error GoTo Errorhandler
'sollte nicht mehr neu verwendet werden
'wird in CCH Weekly noch verwendet, dort allerdings schon z.T gekapselt
'wurde bereits auskommentiert 11.09.17

    Dim workCCHWKL As New edv_Workslot_handler
    Dim domainindex_INT As String
    Dim domainname_TXT As String
    Dim domain_SFX As String
    Dim Staging_PATH As String
    
    Dim filesPerWeek As Integer
    
    Call workCCHWKL.Workslotindex(vbMonday, 2)
    
    domainindex_INT = Left(outlookdomain_ARG, 3)

    If newattachments_INT = 0 Then
        Call StoreAttachment(workCCHWKL)
    End If
    
Errorhandler:

If Err.Number <> 0 Then

    Debug.Print "edv.OutlookFolder_wrapper.FilenameFromStaging.Now()=" & Now()
    Debug.Print "edv.OutlookFolder_wrapper.FilenameFromStaging.Err.Number=" & Err.Number
    Debug.Print "edv.OutlookFolder_wrapper.FilenameFromStaging.newattachments_INT=" & newattachments_INT
    Debug.Print "edv.OutlookFolder_wrapper.FilenameFromStaging.outlookdomain_ARG=" & outlookdomain_ARG
    Debug.Print "edv.OutlookFolder_wrapper.FilenameFromStaging.domainindex_INT=" & domainindex_INT

End If

    FilenameFromStaging = stagingFileName_TXT(index_ARG)

End Function

Rem ______________________________________________________________________________________________________


Public Function StoreAttachment(Optional workslot_ARG As edv_Workslot_handler) As edv_PowerPoint_wrapper
Rem Wird von import_CCHWKL_mediator verwendet
Rem
'da das hier nur PowerPoint als Attachment zur�ckliefert, sollte die Funktion auch so benannt werden, oder?
On Error GoTo Errorhandler

'Const debug_FLAG As Boolean = True
Const debug_FLAG As Boolean = False
Rem die Zur�ckgelieferte Klasse:
Dim resultobject_OBJ As New edv_PowerPoint_wrapper
Rem das Objekt vom Typ MailItem, das im Grunde genommen ausgewertet wird
Dim source_OBJ As MailItem
'was genau ist das?
Dim new_FLAG As Boolean
Rem der Name des Attachments:
Dim myfilename_TXT As String

Dim Anlagen_INT As Integer
Dim indx As Integer
Dim cr_DAT As Date

    If debug_FLAG Then
        Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.debug_FLAG=" & debug_FLAG
        Set source_OBJ = NewMailFromFolder(WorkslotFolder_OBJ(workslot_ARG), keepMailAsUnread)
        Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.resultobject_OBJ.PresentationFull_PATH(" & "workslot_ARG" & ")=" & resultobject_OBJ.PresentationFull_PATH(workslot_ARG)
        Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.newMailsInFolder_INT+1=" & newMailsInFolder_INT + 1
        source_OBJ.Attachments.Item(1).SaveAsFile resultobject_OBJ.addPresentationPath("", source_OBJ.Attachments.Item(1).FileName)
        Debug.Print "edv.OutlookFolder_wrapper.StoreAttachment.Not debug_FLAG=" & Not debug_FLAG
    End If
    
If (myWorkfolder Is Nothing) Then
    Set myWorkfolder = WorkslotFolder_OBJ(workslot_ARG)
End If

myfilename_TXT = ""
newattachments_INT = reset_stagingfilematrix()
For Each objMail In myWorkfolder.Items

    With objMail
        new_FLAG = .UnRead
        Anlagen_INT = .Attachments.Count
        cr_DAT = .CreationTime
       
        If new_FLAG Then
            For indx = 1 To Anlagen_INT
            
            Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.Alternative=" & True
            'Alternative:
            stagingFileName_TXT(newattachments_INT + 1) = .Attachments.Item(indx).FileName
            myfilename_TXT = resultobject_OBJ.PresentationFull_PATH(workslot_ARG)
            myfilename_TXT = resultobject_OBJ.addPresentationPath("", stagingFileName_TXT(newattachments_INT + 1))
            .Attachments.Item(indx).SaveAsFile myfilename_TXT
             newattachments_INT = newattachments_INT + 1
            Next indx
        End If
        
        .UnRead = False
    End With
Next objMail

If newattachments_INT = 0 Then
    stagingFileName_TXT(1) = "<no new attachments>"
    Set resultobject_OBJ = Nothing
End If

Errorhandler:

If (myWorkfolder Is Nothing) Or Err.Number <> 0 Or debug_FLAG Then
    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.Now()=" & Now()
    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.debug_FLAG=" & debug_FLAG
    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.Err.Number=" & Err.Number
    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.Err.Number=" & Err.Description
    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.stagingfilename_TXT(1)=" & stagingFileName_TXT(1)
    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.myfilename_TXT=" & myfilename_TXT
    Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.resultobject_OBJ="
    
    'Debug.Print "edv_OutlookFolder_wrapper.StoreAttachment.myWorkfolder=Nothing"
End If

Set StoreAttachment = resultobject_OBJ

End Function






Rem __________________________________________________________________________________


Private Function FolderFromEventsource(domainindex_ARG As Integer) As Outlook.Folder
'wer ruft diese Funktion eigentlich auf?
'nur noch Intern:
'MailSubjectFromDomain
'kann also auch Private sein?
On Error GoTo Errorhandler

Dim folder_OBJ As Outlook.Folder
Dim returnvalue_OBJ As Outlook.Folder

Dim folder_cnt As Integer
Dim folderindex As Integer
Dim index_LEN As Integer

Set FolderFromEventsource = Nothing
Set returnvalue_OBJ = Nothing

'auskommentiert 11.09.2017
'    index_LEN = Len(domainindex_ARG)
'
'    Set myRootfolder = mySpace.Folders.Item(user_NAME).Folders.Item("Posteingang").Folders.Item(folderhierarchy_MTX(0)).Folders.Item(folderhierarchy_MTX(1))
'    folder_cnt = myRootfolder.Folders.Count
'
'    For folderindex = 1 To folder_cnt
'
'        Set myWorkfolder = myRootfolder.Folders.Item(folderindex)
'
'        If Left(myWorkfolder.Name, index_LEN) = domainindex_ARG Then
'            Set returnvalue_OBJ = myWorkfolder
'            Exit For
'        End If
'
'    Next folderindex
    
Errorhandler:

    If Err.Number <> 0 Then
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.Err.Number=" & Err.Number
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.domainindex_ARG=" & domainindex_ARG
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.myRootfolder.Name=" & myRootfolder.Name
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.returnvalue_OBJ.Name=" & returnvalue_OBJ.Name
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.returnvalue_OBJ.folder_cnt=" & folder_cnt
        Debug.Print "edv_OutlookFolder_wrapper.FolderFromEventsource.folderindex=" & folderindex
    End If
    
If Not (returnvalue_OBJ Is Nothing) Then

    Set FolderFromEventsource = returnvalue_OBJ
    
End If
    
End Function



Private Function WorkslotFolder_OBJ(workslot_ARG As edv_Workslot_handler) As Outlook.Folder
Rem _________________________________________________________________________________________

Rem Diese Funktion liefert den Outlook Folder der zu einem bestimmetn Workslot geh�rt, zur�ck
Rem dazu werden die Ordner durchlaufen und bei �bereinstimmung mit dem Workslotprefix
Rem der Ordner zur�ckgeliefert

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

'Const debug_FLAG As Boolean = True
Const debug_FLAG As Boolean = False

Set WorkslotFolder_OBJ = Nothing
Set returnvalue_OBJ = Nothing
'http://msdn.microsoft.com/en-us/library/office/ff865587(v=office.14).aspx


    folder_cnt = myRootfolder.Folders.Count
    For folderindex = 1 To folder_cnt

        Set myWorkfolder = myRootfolder.Folders.Item(folderindex)
        
        If workslot_ARG.IsWorkslotprefix(myWorkfolder.Name) Then
            Set returnvalue_OBJ = myWorkfolder
            If (newMailsInFolder_INT < 0) Then newMailsInFolder_INT = newMailsInFolder_INT + returnvalue_OBJ.UnReadItemCount + workslot_ARG.PresetFile_CNT
            'newMailsInFolder_INT = returnvalue_OBJ.UnReadItemCount
            Exit For
            
        End If
        
    Next folderindex
    
Errorhandler:

If Err.Number <> 0 Or debug_FLAG Then
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.Err.Number=" & Err.Number
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.Err.Description=" & Err.Description
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.myRootfolder.Name=" & myRootfolder.Name
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.folder_cnt=" & folder_cnt
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.folderindex=" & folderindex
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.newMailsInFolder_INT=" & newMailsInFolder_INT
    Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.IsWorkslotprefix(" & myWorkfolder.Name & ")=" & workslot_ARG.IsWorkslotprefix(myWorkfolder.Name)
End If
    
If Not (returnvalue_OBJ Is Nothing) Then

    'Debug.Print "edv_OutlookFolder_wrapper.WorkslotFolder_OBJ.returnvalue_OBJ.Name=" & returnvalue_OBJ.Name
    Set WorkslotFolder_OBJ = returnvalue_OBJ
    
End If
    
End Function

Private Function NewMailFromFolder(workfolder_ARG As Folder, Optional keepUnread_ARG As Boolean = False) As MailItem
Rem gibt ungelesene Mail aus Ordner zur�ck
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

End Function


Private Function reset_stagingfilematrix()
Rem wird verwendet
'sollte aber nicht mehr neu in Funktionen integriert werden
'wo wird das noch verwendet?
'es wird Zeit f�r den R�ckbau
'verwendung: Store_Attachment

Dim indx_INT As Integer

For indx_INT = 0 To UBound(stagingFileName_TXT)
    stagingFileName_TXT(indx_INT) = ""
Next indx_INT

reset_stagingfilematrix = 0


End Function

Sub launch_Outlook()
'zu dokumentationszwecken, ist eigentlich unsinn, wenn outlook schon l�uft

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
    user_NAME = "norman.gruener@telefonica.com"
    dateprefix_TXT = ""
    
    Set mySpace = Application.GetNamespace("MAPI")
    'keine weitere Verwendung?
    '06.09.2017
    'folder_PATH(0) = mySpace.GetDefaultFolder(olFolderInbox).FolderPath
    folderhierarchy_MTX(0) = "06_Datenbank"
    folderhierarchy_MTX(1) = "12_Eventsource"
    'keine weitere Verwendung 11.09.17
    'filename_TXT = ""
    newMailsInFolder_INT = -1
    
    Set myRootfolder = mySpace.Folders.Item(user_NAME).Folders.Item("Posteingang").Folders.Item(folderhierarchy_MTX(0)).Folders.Item(folderhierarchy_MTX(1))
    Set myWorkfolder = Nothing
    
Errorhandler:
    If Err.Number <> 0 Then
        Debug.Print "edv_OutlookFolder_wrapper.folderhierarchy_MTX(1)=" & folderhierarchy_MTX(1)
        Debug.Print "edv_OutlookFolder_wrapper.myWorkfolder=" & "Nothing"
    End If
    
    
End Sub

Private Sub Class_Terminate()

    Call reset_stagingfilematrix

End Sub