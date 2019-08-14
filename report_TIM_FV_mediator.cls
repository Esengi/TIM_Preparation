VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "report_TIM_FV_mediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem diese Klasse hat vor allem die Aufgabe, die zusammengehörigen Workslots zu bündeln
Private workSIN_Einzelmails_OBJ As New edv_Workslot_handler
' workAvailability As New edv_Workslot_handler

'Private workCCHWeekly As New edv_Workslot_handler
'Private workNOInfomail As New edv_Workslot_handler
'Private workProblemreport As New edv_Workslot_handler
Rem
Rem Funktionell gibt sie an eine aufrufende Prozedur
Rem oder Formular
Rem Stautusmessages zu dem ab, was sie so gerade gemacht hat
Rem
Private outlookFolder_OBJ As New edv_OutlookFolder_wrapper
'Private logger_OBJ As New edv_Control_wrapper
Private Enum execution_IDX
 justexecute = 0
 returnlabel = 1
 returnactive = 2
End Enum
Private statusmessage_TXT As String

'Private Const debug_FLAG As Boolean = False
Private Const debug_FLAG As Boolean = True

Rem _________

Public Function SIN_Status_TXT(Optional execution_ARG As Integer = 0) As String

On Error GoTo Errorhandler

Dim resultvalue_TXT As String
Dim result_OBJ As New edv_Textfileimport_wrapper
Dim unlinkedText_OBJ As New edv_UnlinkedTextfiles_wrapper

   
    Set result_OBJ = outlookFolder_OBJ.TextfileImportFromWorkslot(workSIN_Einzelmails_OBJ)
    
    Debug.Print "report_TIM_FV_mediator.SIN_Status_TXT.result_OBJ.newMail_FLAG=" & result_OBJ.newMail_FLAG
    Debug.Print "report_TIM_FV_mediator.SIN_Status_TXT.result_OBJ.textFileStaging_Path(1)=" & result_OBJ.textFileStaging_Path(1).Staging_PATH
    
    
Errorhandler:
    
    If Err.Number <> 0 Or debug_FLAG Then
    
        Debug.Print "report_TIM_FV_mediator.SIN_Status_TXT.Now()=" & Now()
        Debug.Print "report_TIM_FV_mediator.SIN_Status_TXT.Err.Number=" & Err.Number
        Debug.Print "report_TIM_FV_mediator.SIN_Status_TXT.Err.Description=" & Err.Description
        Debug.Print "report_TIM_FV_mediator.SIN_Status_TXT.Err.resultvalue_TXT=" & resultvalue_TXT
        
    End If

    SIN_Status_TXT = resultvalue_TXT

End Function

Public Function storeAvailabilityStatus_TXT(Optional execution_ARG As Integer = 0) As String

On Error GoTo Errorhandler
    storeAvailabilityStatus_TXT = "0"
    
    If outlookFolder_OBJ.HasNewItems(workAvailability) Then storeAvailabilityStatus_TXT = 1

    If execution_ARG < 2 Then storeAvailabilityStatus_TXT = "101_VerfügbarkeitenFixnet"
    
    If execution_ARG < 1 Then
        storeAvailabilityStatus_TXT = chartInformationStatus_TXT(workAvailability)
        storeAvailabilityStatus_TXT = storeAvailabilityStatus_TXT & Chr(10) & "attachmentStoreStatus_TXT:Done!"
    End If
    
Errorhandler:
    
    If debug_FLAG Then
    
        Debug.Print Err.Number
        Debug.Print Err.Description
        Debug.Print "report_TIM_FV_mediator.storeAvailabilityStatus_TXT().Now()=" & Now()
        Debug.Print "report_TIM_FV_mediator.storeAvailabilityStatus_TXT.execution_ARG=" & execution_ARG
        Debug.Print "report_TIM_FV_mediator.storeAvailabilityStatus_TXT.workAvailability.Workslotindex=" & workAvailability.Workslotindex
        Debug.Print "report_TIM_FV_mediator.storeAvailabilityStatus_TXT.execution_ARG=" & execution_ARG
        Debug.Print "report_TIM_FV_mediator.storeAvailabilityStatus_TXT()=1"
    
    End If

    
End Function

Public Function storeCCHWKLStatus_TXT() As String
Rem
'wer verwendet das:?
Dim pptCCHWeekly As New edv_PowerPoint_wrapper
    Debug.Print "storeCCHWKLStatus_TXT().Now()=" & Now()
    storeCCHWKLStatus_TXT = ""
    storeCCHWKLStatus_TXT = storeCCHWKLStatus_TXT & chartInformationStatus_TXT(workCCHWeekly)
    storeCCHWKLStatus_TXT = storeCCHWKLStatus_TXT & Chr(10) & ""
End Function

Public Function storeNOInfomailStatus_TXT(Optional execution_ARG As Integer = 0) As String

On Error GoTo Errorhandler

    storeNOInfomailStatus_TXT = "0"
    If outlookFolder_OBJ.HasNewItems(workNOInfomail) Then storeNOInfomailStatus_TXT = 1
        Debug.Print "report_TIM_FV_mediator.storeNOInfomailStatus_TXT.execution_ARG=" & execution_ARG
        Debug.Print "report_TIM_FV_mediator.storeNOInfomailStatus_TXT()=1"
    
    If execution_ARG < 2 Then
    
        storeNOInfomailStatus_TXT = "103_NO Infomail"
    
    End If
    
    If execution_ARG < 1 Then

        storeNOInfomailStatus_TXT = attachmentStoreStatus_TXT(workNOInfomail)
        storeNOInfomailStatus_TXT = storeNOInfomailStatus_TXT & Chr(10) & "storeNOInfomail:Done!"
    
    End If
    
Errorhandler:

    If debug_FLAG Then
    
        Debug.Print "report_TIM_FV_mediator.storeNOInfomailStatus_TXT()"
    
    End If

    
End Function


Public Function storeProblemreportStatus_TXT(Optional execution_ARG As Integer = 0) As String

    Dim storeProblemreport_TXT As String
  
    storeProblemreport_TXT = "0"
    If outlookFolder_OBJ.HasNewItems(workProblemreport) Then storeProblemreport_TXT = 1
        
    If execution_ARG < 2 Then
        storeProblemreport_TXT = "104_FN_Problemreport"
    End If
    
    
    If execution_ARG < 1 Then
        storeProblemreport_TXT = ""
        storeProblemreport_TXT = attachmentStoreStatus_TXT(workProblemreport)
        storeProblemreport_TXT = storeProblemreport_TXT & Chr(10) & "storeProblemreport_TXT:Done!"
    End If
    
    storeProblemreportStatus_TXT = storeProblemreport_TXT
    
    If debug_FLAG Then
        Debug.Print "report_TIM_FV_mediator.storeProblemreportStatus_TXT.Now=" & Now()
        Debug.Print "report_TIM_FV_mediator.storeProblemreportStatus_TXT.execution_ARG=" & execution_ARG
    End If

End Function

Public Sub launchStaging()

    Dim cchWeekly_OBJ As New edv_PowerPoint_wrapper
    Dim NOInfomail_OBJ As New edv_Worksheet_wrapper
    
    'todo: den slash am Pfadende in der Funktion anhängen lassen:
    'Call cchWeekly_OBJ.LaunchPowerpoint("C:\Users\Public\Documents\102_CCHWKL" & "\", "weekly_report_CW28_2017.pptx")
    
    'Call NOInfomail_OBJ.LaunchExcel("C:\Users\Public\Documents\103_NOINFO-mail\SIN-Überblick-2017-cw28.xlsx")
    'Call Availability.LaunchExcel("C:\Users\Public\Documents\103_NOINFO-mail\SIN-Überblick-2017-cw28.xlsx")
      
End Sub


Private Function attachmentStoreStatus_TXT(workSlot_ARG As edv_Workslot_handler) As String

Rem Prozedurvariablen ______________________________________________________________________________________

Debug.Print "report_TIM_FV_mediator.debug_FLAG=" & debug_FLAG

Rem wenn ein Worksheet (a/k/a "Excelsheet" o.ä.) erwartet wird
Dim worksheet_OBJ As New edv_Worksheet_wrapper

Dim workSlotControlItem3 As String
Dim IsWorksheet_FLAG As Boolean
Rem Initialisierungen_________________________________________________________________________________________

IsWorksheet_FLAG = False
workSlotControlItem3 = "<tbd>"

Rem Prozedur__________________________________________________________________________________________________

On Error GoTo Errorhandler
    workSlotControlItem3 = workSlot_ARG.ControlItem_TXT(3)
    IsWorksheet_FLAG = (workSlot_ARG.ControlItem_TXT(3) = "xls")

    'ist nicht so geschickt, daß hier die Eigenheiten eines bestimmten Workslots bekannt sein müssen
    'und dieser noch als Ausnahme gehandlet werden muss
    If workSlot_ARG.Workslotindex = 102 Then
        Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.workslot_ARG.Workslotindex=" & workSlot_ARG.Workslotindex
        'Set pptCCHWeekly = outlookFolder_OBJ.StoreAttachment(workslot_ARG)
    Else
        'hier wird davon ausgegangen, daß das Worksheet in einem Outlook Folder liegt
        'Das ist - angesichts der Tatsache, daß wir hier in Outlook programmieren - okay
        Set worksheet_OBJ = outlookFolder_OBJ.worksheetImportFromWorkslot(workSlot_ARG)
        attachmentStoreStatus_TXT = worksheet_OBJ.worksheet_PATH()
     End If
    
Errorhandler:

    If Err.Number <> 0 Or debug_FLAG Then

        Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.workSlotControlItem3=" & workSlotControlItem3
        Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.IsWorksheet_FLAG=" & IsWorksheet_FLAG
        Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.worksheet_OBj.workslotPathByWorkslot(workslot_ARG)=" & worksheet_OBJ.workslotPathByWorkslot(workSlot_ARG)
        Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.worksheet_OBj.worksheet_PATH()=" & worksheet_OBJ.worksheet_PATH()
        'Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.worksheet_OBj.worksheet_PATH()=" & pptCCHWeekly.PresentationFull_PATH
        Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.worksheet_OBj.IsLinked_FLAG=" & worksheet_OBJ.IsLinked_FLAG
    
    End If
    'wurde 26.06.17 auskommentiert da keine vernünftigen Ergebnisse rauskommen
    'Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.logger_OBJ.IsLogged=" & logger_OBJ.IsLogged(workSlot_ARG, attachmentStoreStatus_TXT)
    Debug.Print "report_TIM_FV_mediator.attachmentStoreStatus_TXT.workSlot_ARG.IsLogged(attachmentStoreStatus_TXT)=" & workSlot_ARG.IsLogged(attachmentStoreStatus_TXT)
         
End Function

Private Function attachmentObjectHandler(workSlot_ARG As edv_Workslot_handler) As Object
Rem Variablen
Dim resultWorksheet_OBJ As New edv_Worksheet_wrapper
Dim resultPowerpoint_OBJ As New edv_PowerPoint_wrapper
Dim statusmessage_TXT As String
Dim IsLogged_FLAG As Boolean
Rem ___________________________________________________________________________________________
On Error GoTo Errorhandler

    statusmessage_TXT = "report_TIM_FV_mediator.attachmentObjectHandler.IsCalledFrom=chartInformationStatus_TXT"
    
    If workSlot_ARG.ControlItem_TXT(3) = "xls" Then
        Set resultWorksheet_OBJ = outlookFolder_OBJ.worksheetImportFromWorkslot(workSlot_ARG)
        statusmessage_TXT = resultWorksheet_OBJ.worksheet_PATH
        IsLogged_FLAG = workSlot_ARG.IsLogged(statusmessage_TXT)
    End If
Rem
    statusmessage_TXT = "report_TIM_FV_mediator.attachmentObjectHandler.IsLogged_FLAG=" & IsLogged_FLAG
Rem _____________________________________________________________________________________________
Errorhandler:

    If Err.Number <> 0 Or debug_FLAG Then
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.Now()=" & Now()
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.debug_FLAG=" & debug_FLAG
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.Err.Number=" & Err.Number
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.Err.Description=" & Err.Description
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.workslot_ARG.Workslotindex=" & workSlot_ARG.Workslotindex
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.workslot_ARG.ControlItem_TXT(3)=" & workSlot_ARG.ControlItem_TXT(3)
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.resultWorksheet_OBJ.worksheet_PATH=" & resultWorksheet_OBJ.worksheet_PATH
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.resultWorksheet_OBJ.worksheet_PATH=Nothing?->Problem!"
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.resultWorksheet_OBJ.IsLogged_FLAG=" & IsLogged_FLAG
        Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.statusmessage_TXT=" & statusmessage_TXT
    End If
Rem _____________________________________________________________________________________________
    Set attachmentObjectHandler = resultWorksheet_OBJ

    Set resultWorksheet_OBJ = Nothing
    Debug.Print "report_TIM_FV_mediator.attachmentObjectHandler.resultWorksheet_OBJ=Nothing"
End Function

Private Function chartInformationStatus_TXT(workSlot_ARG As edv_Workslot_handler) As String

On Error GoTo Errorhandler

Dim myChart_OBJ As New edv_Chartinformation_handler
Dim resultvalue_TXT As String


    resultvalue_TXT = "<no data>"
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.resultvalue_TXT=" & resultvalue_TXT
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.workSlot_ARG.Workslotindex()=" & workSlot_ARG.Workslotindex()
    resultvalue_TXT = myChart_OBJ.DataSeriesName(1, "", "", attachmentObjectHandler(workSlot_ARG))
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.Left(resultvalue_TXT, 1)=" & Left(resultvalue_TXT, 1)
 
Errorhandler:

If Err.Number <> 0 Or debug_FLAG Then
        
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT Is called from " & ""
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.Now()=" & Now()
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.debug_FLAG=" & debug_FLAG
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.Err.Description=" & Err.Description
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.Err.Number=" & Err.Number
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.resultvalue_TXT=" & resultvalue_TXT
        
    Debug.Print "report_TIM_FV_mediator.chartInformationStatus_TXT.workSlot_ARG.IsLogged(attachmentStoreStatus_TXT)=" & workSlot_ARG.IsLogged(chartInformationStatus_TXT)

End If
     
    chartInformationStatus_TXT = resultvalue_TXT
    Set myChart_OBJ = Nothing

End Function


Private Sub Class_Initialize()

    statusmessage_TXT = ""
    
    Call workSIN_Einzelmails_OBJ.Workslotindex(vbMonday, 6)
    'Call workAvailability.Workslotindex(vbMonday, 1)
    'Call workCCHWeekly.Workslotindex(vbMonday, 2)
    'Call workNOInfomail.Workslotindex(vbMonday, 3)
    'Call workProblemreport.Workslotindex(vbMonday, 4)
    
    If debug_FLAG Then
        Debug.Print "report_TIM_FV_mediator.Initialize.Now()=" & Now()
        Debug.Print "report_TIM_FV_mediator.Initialize.debug_FLAG=" & debug_FLAG
    End If
End Sub

Private Sub Class_Terminate()

    On Error GoTo Errorhandler
    Const debug_FLAG As Boolean = False
Errorhandler:
    If debug_FLAG Then
        Debug.Print "report_TIM_FV_mediator.Class_Terminate.Err.Number=" & Err.Number
        Debug.Print "report_TIM_FV_mediator.Class_Terminate.Err.Description=" & Err.Description
        Debug.Print "report_TIM_FV_mediator.Class_Terminate=Done!"
    End If
End Sub

