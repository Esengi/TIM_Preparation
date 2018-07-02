VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Rem _________________________________________________
Rem Verweise
Rem Visual Basic for Applications
Rem Microsoft Outlook 15.0 Object Library
Rem OLE Automation
Rem Microsoft Office 15.0 Object Library
Rem Microsoft DAO 3.6 Object Library
Rem Microsoft Data Source Interfaces for ActiveX Data Binding
Rem Microsoft Jet and Replication Objects 2.6 Library
Rem Outlook SharePoint Social Provider
Rem Microsoft Access 15.0 Object Library
Rem Microsoft Forms 2.0 Object Library
Rem Microsoft PowerPoint 15.0 Object Library
Rem Microsoft ActiveX Data Objects 6.1 Library
Rem Microsoft XML, v6.0
Rem Microsoft HTML Object Library
Rem __________________________________________________

Option Explicit
Rem Klassenobjekte
Rem __________________________________________________
Dim tim_OBJ As New report_TIM_FV_mediator
Dim nprm_OBJ As New report_NPRMFV_mediator
Dim benchmarkPerfMon_OBJ As New report_TITANS_mediator
Dim controller_OBJ As New edv_Control_wrapper
Dim outlook_OBJ As New edv_OutlookFolder_wrapper
Dim folder_OBJ As New edv_FilesystemFolder_wrapper
Rem Statusmessagevariablen
Rem __________________________________________________
Dim statusmessage_TXT As String
Dim statusheader_TXT As String


Rem für Outlook Desktop geschrieben
Rem designed for Outlook Desktop 2013. Makro security set to "ask"
Public Sub rule_101_VerfügbarkeitenFixnet()
'Public Sub rule_101_VerfügbarkeitenFixnet(oMail As Outlook.MailItem)
Rem
    Call AddStatus("rule_101_VerfügbarkeitenFixnet.tim_OBJ.storeAvailabilityStatus_TXT=" & tim_OBJ.storeAvailabilityStatus_TXT)
    
Errorhandler:

    If Err.Number <> 0 Then
       Call AddStatus("rule_101_VerfügbarkeitenFixnet.Err.Number=" & Err.Number)
       Call AddStatus("rule_101_VerfügbarkeitenFixnet.Err.Number=" & Err.Description)
    End If

    Call DisplayStatusMessage
    
    Set tim_OBJ = Nothing

End Sub

Public Sub store_102_CCHWKL_Report()
'Public Sub store_102_CCHWKL_Report(oMail As Outlook.MailItem)
'Todo: das hier wird wohl nicht mehr benötigt?
On Error GoTo Errorhandler
Dim mediator_OBJ As New import_CCHWKL_mediator
Const debug_FLAG As Boolean = False
'Const debug_FLAG As Boolean = True


    If debug_FLAG Then
    Rem grundsätzlich funktioniert das!
        Debug.Print "ThisOutlookSession.store_102_CCHWKL_Report.debug_FLAG=" & debug_FLAG
        Call AddStatus("store_103_NOINFO_attachment.tim_OBJ.storeCCHWKLStatus_TXT=" & tim_OBJ.storeCCHWKLStatus_TXT)

    End If
    
    If Not debug_FLAG Then
    
        Debug.Print "ThisOutlookSession.store_102_CCHWKL_Report.debug_FLAG=" & debug_FLAG
        statusmessage_TXT = ""
        Call AddStatus("ThisOutlookSession.store_102_CCHWKL_Report.mediator_OBJ.IsHibernated=" & _
                        mediator_OBJ.IsHibernated & _
                        Chr(10))
    
    End If
    

     
Errorhandler:

If Err.Number <> 0 Then
    Debug.Print "ThisOutlookSession.store_102_CCHWKL_Report.Err.Description=" & Err.Description
    statusmessage_TXT = "ThisOutlookSession.store_102_CCHWKL_Report.Err.Number=" & Err.Number & Chr(10)
    statusmessage_TXT = statusmessage_TXT & "ThisOutlookSession.store_102_CCHWKL_Report.Err.Description=" & Err.Description & Chr(10)
End If

Call DisplayStatusMessage

'Set outlook_OBJ = Nothing
'Set folder_OBJ = Nothing
Set mediator_OBJ = Nothing

End Sub


Public Sub store_103_NOINFO_attachment()
'Public Sub store_103_NOINFO_attachment(oMail As Outlook.MailItem)
On Error GoTo Errorhandler

    Call AddStatus("store_103_NOINFO_attachment.tim_OBJ.storeNOInfomailStatus=" & tim_OBJ.storeNOInfomailStatus_TXT)
   
Errorhandler:

    If Err.Number <> 0 Then
       Call AddStatus("store_103_NOINFO_attachment.Err.Number=" & Err.Number)
       Call AddStatus("store_103_NOINFO_attachment.Err.Number=" & Err.Description)
    End If

    Call DisplayStatusMessage

    Set tim_OBJ = Nothing

End Sub

Public Sub store_104_Problemreport_attachment()
'Public Sub store_103_NOINFO_attachment(oMail As Outlook.MailItem)
On Error GoTo Errorhandler

    Dim tim_OBJ As New report_TIM_FV_mediator
    Call AddStatus("store_104_Problemreport.tim_OBJ.storeProblemreport_TXT=" & tim_OBJ.storeProblemreportStatus_TXT)
   
Errorhandler:

    If Err.Number <> 0 Then
       Call AddStatus("store_104_Problemreport_attachment.Err.Number=" & Err.Number)
       Call AddStatus("store_104_Problemreport.Err.Number=" & Err.Description)
    End If

    Call DisplayStatusMessage

    Set tim_OBJ = Nothing

End Sub



Public Sub store_202_IMSpdf()

    Call AddStatus("store_504_PerformanceMonitoring.benchmarkPerfMon_OBJ.storeIMS=" & benchmarkPerfMon_OBJ.storeIMS)

End Sub

Public Sub store_202_FBENCH_csv()
Rem wird wöchentlich geliefert
Rem Aktualisierung Dienstag Morgens

Dim importer_OBJ As New report_FBENCH_mediator
Dim statusmessage_TXT As String

    statusmessage_TXT = ""
    statusmessage_TXT = "store_202_FBENCH_csv.importer_OBJ.extractionStatus_TXT(1)=" & _
    importer_OBJ.extractionStatus_TXT(1)

    Call MsgBox(statusmessage_TXT, vbOKOnly, "store_202_FBENCH_csv")
    
    statusmessage_TXT = "store_202_FBENCH_csv.importer_OBJ.extractionStatus_TXT(2)=" & _
    importer_OBJ.extractionStatus_TXT(2)
    
    Call MsgBox(statusmessage_TXT, vbOKOnly, "store_202_FBENCH_csv")
    
    statusmessage_TXT = "store_202_FBENCH_csv.importer_OBJ.extractionStatus_TXT(3)=" & _
    importer_OBJ.extractionStatus_TXT(3)
    
    Call MsgBox(statusmessage_TXT, vbOKOnly, "store_202_FBENCH_csv")
    
End Sub




Public Sub store_205_DIAMTR_xls()

On Error GoTo Errorhandler

    Dim weeklyDiameterRep_OBJ As New edv_Workslot_handler
    Dim resultWorksheet_OBJ As New edv_Worksheet_wrapper
    
    Call weeklyDiameterRep_OBJ.Workslotindex(vbTuesday, 5)
    Call weeklyDiameterRep_OBJ.AddDomain("SAR", 0)

    Call AddStatus("Diameter Message Report")
    Call AddStatus("wird wöchentlich geliefert")
    Call AddStatus("Script kann nicht per Regel gestartet werden")
    
    Rem soll Attachment in Folder <tbd> kopieren
        
    Set resultWorksheet_OBJ = outlook_OBJ.worksheetImportFromWorkslot(weeklyDiameterRep_OBJ)
    Call AddStatus("resultWorksheet_OBJ.worksheet_PATH=" & resultWorksheet_OBJ.worksheet_PATH)
            
Errorhandler:

    If Err.Number <> 0 Then
       Call AddStatus("store_205_DIAMTR_xls.Err.Number=" & Err.Number)
       Call AddStatus("store_205_DIAMTR_xls.Err.Number=" & Err.Description)
    End If
    
    Call DisplayStatusMessage
 
End Sub


Public Sub store_504_PerformanceMonitoring()
Rem Dim benchmarkPerfMon_OBJ As New report_TITANS_mediator
On Error GoTo Errorhandler
    
    Call AddStatus("store_504_PerformanceMonitoring.benchmarkPerfMon_OBJ.storePerfMonitStatus=" & benchmarkPerfMon_OBJ.storePerfMonitStatus_TXT)
   
Errorhandler:

    If Err.Number <> 0 Then
       Call AddStatus("store_504_PerformanceMonitoring.Err.Number=" & Err.Number)
       Call AddStatus("store_504_PerformanceMonitoring.Err.Number=" & Err.Description)
    End If

    Call DisplayStatusMessage

    Set benchmarkPerfMon_OBJ = Nothing

End Sub





Public Sub DisplayForm_TIM()

    form_FN_TIM_Preparation.Show

End Sub


Public Sub DisplayForm_NPRM()

    form_NPRM_Preparation.Show

End Sub

Rem Inherited Class Subs
Rem __________________________________________________________--


Private Sub AddStatus(statusmessage_ARG As String, Optional statusheader_ARG As String = "Statusmessage")

    statusheader_TXT = statusheader_ARG
    statusmessage_TXT = statusmessage_TXT & statusmessage_ARG & Chr(10)

End Sub



Private Sub DisplayStatusMessage()

On Error GoTo Errorhandler

Call MsgBox(statusmessage_TXT, vbOKOnly, statusheader_TXT)

Errorhandler:
statusmessage_TXT = ""

End Sub


Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

End Sub

Private Sub Application_NewMail()

End Sub

Private Sub Application_Startup()

    statusmessage_TXT = ""

End Sub


