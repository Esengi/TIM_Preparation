VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "report_TITANS_mediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private debug_FLAG As Boolean
Private workSlot_OBJ As New edv_Workslot_handler
Rem Ablageverzeichnisse
Rem -------------------
Rem SBGBGF
Private SBGBGFofflineZipFile_PFAD As New edv_FilesystemFolder_wrapper
Private SBGBGFonlineZipFile_PFAD As New edv_FilesystemFolder_wrapper
Private SBGBGFrawFile_PFAD As New edv_FilesystemFolder_wrapper
Rem -----------------------------------------------------------------
Rem Rohdateinamen
Rem -------------
Rem SBGBGF
'wo wird das verwendet?
Private SBGBGFofflineZipFile_NAME As String

Private unzipService_OBJ As New mserv_zipFile_wrapper
Private textFile_OBJ As New edv_Textfileimport_wrapper
Private mailfolder_OBJ As New edv_OutlookFolder_wrapper
'Private titanImport_OBJ As New edv_LinkTextfiles_wrapper
'Private hibernator_OBJ As New edv_SQL_wrapper
'Private groupfunction_OBJ As New edv_SQL_wrapper


Sub debug_me()

    Debug.Print extractionStatus_TXT()
   
End Sub

Function extractionStatus_TXT() As String
On Error GoTo Errorhandler

Dim offlineZipFile_PFAD As New edv_FilesystemFolder_wrapper
Dim rawfile_PFAD As New edv_FilesystemFolder_wrapper
Dim offlineZipFile_NAME As String
Dim resultvalue_TXT As String
    
    
    Rem zu extrahierendes
    Rem -----------------
    Set offlineZipFile_PFAD = SBGBGFofflineZipFile_PFAD
    Set rawfile_PFAD = SBGBGFrawFile_PFAD
    
    Rem welche Files sind aus dem Onlineverzeichnis zu holen?
    'das sollten uns die ungelesenen Mails aus dem Outlook Folder sagen
    
    
    Rem holen der gezippten Files aus dem Onlineverzeichnis
    'macht derzeit ein batchfile
    
    
    offlineZipFile_NAME = SBGBGFofflineZipFile_NAME
    
    Rem -----------------
    Rem extrahieren
    
    Call unzipService_OBJ.GetUnzippedTextfile(offlineZipFile_PFAD, offlineZipFile_NAME, rawfile_PFAD)
    
    Rem kopieren/umbenennen
    Rem nicht mehr notwendig, da wir nicht mehr in Datenbanken importieren wollen
    'Call textFile_OBJ.GetCopiedTextfile(workslot_OBJ, 1, "sbg")
    
    Rem zusammenkopieren
    Call SBGBGFrawFile_PFAD.IsBatchRun("5041_SBGBGF.bat")
    
Errorhandler:
    
    Debug.Print "report_TITANS_mediator.extractionStatus_TXT.Err.Number=" & Err.Number
    Debug.Print "report_TITANS_mediator.extractionStatus_TXT.Err.Description=" & Err.Description
    Debug.Print "report_TITANS_mediator.extractionStatus_TXT.offlineZipFile_PFAD.staging_PATH=" & offlineZipFile_PFAD.Staging_PATH
    Debug.Print "report_TITANS_mediator.extractionStatus_TXT.offlineZipFile_NAME=" & offlineZipFile_NAME
    Debug.Print "report_TITANS_mediator.extractionStatus_TXT.rawFile_PFAD.staging_PATH=" & rawfile_PFAD.Staging_PATH
    Debug.Print "report_TITANS_mediator.extractionStatus_TXT.resultvalue_TXT=" & resultvalue_TXT
    extractionStatus_TXT = resultvalue_TXT

End Function



Private Function AddStatusmessage(status_ARG As String) As Integer

    statusmessage_TXT = statusmessage_TXT & status_ARG & Chr(10)
    AddStatusmessage = 1

End Function

Private Function SBGBGFofflineZipFileName_TXT() As String

    Dim outlookFolder_OBJ As New edv_OutlookFolder_wrapper

    Dim resultvalue_TXT As String
    Dim firstDayOfCurrentWeek As Date
    Dim dayInFileName As Integer
    Dim monthInFileName As Integer
    Dim yearInFileName As Integer
    
    On Error GoTo Errorhandler
    
    firstDayOfCurrentWeek = Now() - Weekday(Now(), vbMonday) + 0
    
    Debug.Print "report_TITANS_mediator.SBGBGFofflineZipFileName_TXT.firstDayOfCurrentWeek=" & firstDayOfCurrentWeek
    
    yearInFileName = Year(firstDayOfCurrentWeek)
    monthInFileName = Month(firstDayOfCurrentWeek)
    dayInFileName = Day(firstDayOfCurrentWeek)
    
    SBGBGFofflineZipFileName_TXT = "sbg_dsltest."
    SBGBGFofflineZipFileName_TXT = SBGBGFofflineZipFileName_TXT & yearInFileName
    If monthInFileName < 10 Then SBGBGFofflineZipFileName_TXT = SBGBGFofflineZipFileName_TXT & "0"
    SBGBGFofflineZipFileName_TXT = SBGBGFofflineZipFileName_TXT & Month(firstDayOfCurrentWeek)
    If dayInFileName < 10 Then SBGBGFofflineZipFileName_TXT = SBGBGFofflineZipFileName_TXT & "0"
    SBGBGFofflineZipFileName_TXT = SBGBGFofflineZipFileName_TXT & dayInFileName
       
    Set unzipService_OBJ = outlookFolder_OBJ.zipMicroServiceFromWorkslot(workSlot_OBJ)
    SBGBGFofflineZipFileName_TXT = unzipService_OBJ.zippedTextFile

Errorhandler:
        
    Debug.Print "report_TITANS_mediator.SBGBGFofflineZipFileName_TXT.Err.Number=" & Err.Number
    Debug.Print "report_TITANS_mediator.SBGBGFofflineZipFileName_TXT.Err.Description=" & Err.Description
    Debug.Print "report_TITANS_mediator.SBGBGFofflineZipFileName_TXT.debug_FLAG=" & debug_FLAG

End Function


Private Sub Class_Initialize()

    debug_FLAG = True
    
    Call workSlot_OBJ.Workslotindex(vbFriday, 4)
    Call workSlot_OBJ.AddDomain("SBGBGF", 0)
    Call workSlot_OBJ.AddDiscriminator("CD", 0)
     
    Debug.Print "report_TITANS_mediator.Class_Initialize"
    Debug.Print "report_TITANS_mediator.debug_FLAG=" & debug_FLAG
    Debug.Print "report_TITANS_mediator.SBGBGFofflineZipFile_PFAD.IsExisting=" & SBGBGFofflineZipFile_PFAD.IsExisting("C:\Users\gruenern\OneDrive for Business\5041_SBGBGF\zipfiles\")
    Debug.Print "report_TITANS_mediator.SBGBGFonlineZipFile_PFAD.IsExisting=" & SBGBGFonlineZipFile_PFAD.IsExisting("S:\Email attachments from Flow\504_FBENCH_perfMonitoring\")

    Debug.Print "report_TITANS_mediator.SBGBGFrawFile_PFAD.IsExisting=" & SBGBGFrawFile_PFAD.IsExisting("C:\Users\gruenern\OneDrive for Business\5041_SBGBGF\textfiles\")
    SBGBGFofflineZipFile_NAME = SBGBGFofflineZipFileName_TXT
    Debug.Print "report_TITANS_mediator.extractionStatus_TXT.SBGBBGFofflineZipFile_NAME=" & SBGBGFofflineZipFile_NAME



End Sub

Private Sub Class_Terminate()

    Set SBGBGFofflineZipFile_PFAD = Nothing
    Set SBGBGFrawFile_PFAD = Nothing
    Set unzipService_OBJ = Nothing

End Sub
