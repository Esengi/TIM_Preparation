VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "report_FBENCH_mediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private workSlot304_OBJ As New edv_Workslot_handler
Private workSlot505_OBJ As New edv_Workslot_handler
Private outlook_OBJ As New edv_OutlookFolder_wrapper
Private rawFileFolder_OBJ As New edv_FilesystemFolder_wrapper
Private zipFileFolder_OBJ As New edv_FilesystemFolder_wrapper
Private unzipFile_OBJ As New mserv_zipFile_wrapper
Private ftpFile_OBJ As New mserv_ftpFile_wrapper

Rem XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Rem XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private unreadImportWeeks_INT As Integer

Rem XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Private Const stagingFolder_PATH As String = "C:\Users\gruenern\OneDrive for Business"
Private Const stagingFolder_PATH = "C:\Users\Public\Documents"
Private Const zipFileFolder_TXT As String = "\304_CT2020_FixedVoice\zipfiles"
Private Const rawFileFolder_TXT As String = "\304_CT2020_FixedVoice\textfiles"
Private Const moduleName_TXT As String = "report_FBENCH_mediator"
Private Const moduleVersion_DBL As Double = 3

'Private Const debug_FLAG As Boolean = True
Private Const debug_FLAG As Boolean = False

Sub debug_me()

    'Debug.Print extractionStatus_TXT(False)
    'Debug.Print extractionStatus_TXT
    Call RunShellCommands

End Sub

Public Function extractionStatus_TXT(Optional checkOnly_ARG As Boolean = False) As String
On Error GoTo Errorhandler

Dim resultFile_OBJ As New edv_UnlinkedTextfiles_wrapper

Dim resultvalue_TXT As String
Dim weekBack_IDX As Integer
Dim isOnlineFTP_FLAG As Integer

    weekBack_IDX = 1
    Rem zu debugging-zwecken ohne performante internetverbindung
    Rem in diesem Fall die zweite Zeile auskommentieren
    isOnlineFTP_FLAG = False
    isOnlineFTP_FLAG = True
    
    Rem wir möchten vielleicht nur geprüft haben, ob neue Mails vorliegen. Oder etwas anderes, das noch nicht defniert ist.
    If checkOnly_ARG Then
        isOnlineFTP_FLAG = False
        resultvalue_TXT = "<check_successful>"
        If unreadImportWeeks_INT = 0 Then resultvalue_TXT = "<check_unsuccessful>"
    Else
        If isOnlineFTP_FLAG Then Call ftpFile_OBJ.IsGet(datePrefix_TXT(weekBack_IDX), "")
        resultvalue_TXT = "<ftp_done>"
        Call unzipFile_OBJ.GetUnzippedTextfile(zipFileFolder_OBJ, zipFileName_TXT(weekBack_IDX), rawFileFolder_OBJ)
        resultvalue_TXT = "<zip_done>"
        Call RunShellCommands
        resultvalue_TXT = "<run_done>"
    End If
  
Errorhandler:

    If debug_FLAG Or Err.Number <> 0 Then
        Debug.Print moduleName_TXT & ".debug_FLAG=" & debug_FLAG
        Debug.Print moduleName_TXT & ".Err.Numer=" & Err.Number
        Debug.Print moduleName_TXT & ".Err.Numer=" & Err.Description
        Debug.Print moduleName_TXT & ".isOnlineFTP_FLAG=" & isOnlineFTP_FLAG
        Debug.Print moduleName_TXT & ".extractionStatus_TXT.=" & zipFileFolder_OBJ.staging_PATH
        Debug.Print moduleName_TXT & ".extractionStatus_TXT.resultvalue_TXT=" & resultvalue_TXT
    End If
    
    extractionStatus_TXT = resultvalue_TXT

End Function

Private Function missingWeeks_INT() As Integer
    
   
    'todo: herausfinden der Anzahl noch nicht per ftp abgeholter zipfiles

    'finde heraus, ob in dem Verzeichnis schon ein file liegt
    'welchen Datumsprefix sollte die Datei 1 (2, 3) haben?
    'Dim datumsPrefixLetzteDatei_TXT As String
    Dim datumsPrefixLetzteDatei_TXT As String
    
    datumsPrefixLetzteDatei_TXT = "19"
    
    'Call zipFileFolder_OBJ.IsExisting(
    'das sind dinge, die der Mediator erst mal wissen sollte
    'die er allerdings wenn sinnvoll
    'an andere Klassen auslagern kann
    'Mal einen WeekControlWrapper andenken?
    
    
    'Debug.Print "missingWeeks_INT" & zipFileFolder_OBJ.FullStaging_PATH

    missingWeeks_INT = 1
 
End Function

Private Function zipFileName_TXT(index_ARG As Integer) As String

On Error GoTo Errorhandler

    Dim resultvalue_TXT As String

    Const zipFileNameTrunk_TXT As String = "_RawData_CB_TE_WEEK.zip"
      
    resultvalue_TXT = datePrefix_TXT(index_ARG) & "0000" & zipFileNameTrunk_TXT
    zipFileName_TXT = resultvalue_TXT

Errorhandler:

    If debug_FLAG Or Err.Number <> 0 Then
        Debug.Print moduleName_TXT & ".zipFileName_TXT.resultvalue_TXT=" & resultvalue_TXT
    End If

End Function


Private Function datePrefix_TXT(weeksBack_ARG) As String

Rem diese Funktion wird im zipFileName_TXT und im extractionStatus vewendet
Rem sie liefert das für den Mediator spezifische Datumsformat

    Dim firstDayOfPreviousWeek As String
    Dim monthInFileName As String
    Dim dayInFileName As String
    Dim resultvalue_TXT As String
    
    firstDayOfPreviousWeek = Now() - Weekday(Now(), vbMonday) - (weeksBack_ARG * 7) + 1
    
    monthInFileName = Month(firstDayOfPreviousWeek)
    dayInFileName = Day(firstDayOfPreviousWeek)

    datePrefix_TXT = Right(Year(firstDayOfPreviousWeek), 2)
    resultvalue_TXT = Right(Year(firstDayOfPreviousWeek), 2)
    If monthInFileName < 10 Then datePrefix_TXT = datePrefix_TXT & "0"
    datePrefix_TXT = datePrefix_TXT & monthInFileName
    If dayInFileName < 10 Then datePrefix_TXT = datePrefix_TXT & "0"
    datePrefix_TXT = datePrefix_TXT & dayInFileName
    
    If debug_FLAG Then
    
        Debug.Print moduleName_TXT & ".debug_FLAG=" & debug_FLAG
        Debug.Print moduleName_TXT & ".datePrefix_TXT.firstDayOfPreviousWeek=" & firstDayOfPreviousWeek
        Debug.Print moduleName_TXT & ".datePrefix_TXT.firstDayOfPreviousWeek=" & resultvalue_TXT
    
    End If

End Function


Private Sub RunShellCommands()

On Error GoTo Errorhandler

Dim shellcommand_TXT As String
Dim batchFileName_TXT As String
Dim resultvalue_DBL As Double
    'Methode per DOS Commando. Crude, aber funktioniert.
    'wird erst mal aus dieser Klasse heraus ausgeführt, da die Klasse auch über Strukturen etc.
    'bescheid wissen sollte.
    shellcommand_TXT = "Copy "
    shellcommand_TXT = shellcommand_TXT & "C:\Users\Public\Documents\304_CT2020_FixedVoice\textfiles\*.CALL.E_VB_0002 "
    shellcommand_TXT = shellcommand_TXT & "C:\Users\Public\Documents\304_CT2020_FixedVoice\textdatenbank\5051_CALL_TXTs.txt"
    
    batchFileName_TXT = "C:\Users\Public\Documents\304_CT2020_FixedVoice\304_textdatenbank.bat"
    
    Debug.Print moduleName_TXT & ".RunShellCommands.shellcommand_TXT=" & shellcommand_TXT
    Debug.Print moduleName_TXT & ".RunShellCommands.batchFileName_TXT=" & batchFileName_TXT
    
    resultvalue_DBL = Shell(batchFileName_TXT, vbMaximizedFocus)
    
Errorhandler:
    
    If Err.Number <> 0 Then

        Debug.Print "RunShellCommands.Err.Number=" & Err.Number
        Debug.Print "RunShellCommands.Err.Description=" & Err.Description
    
    End If

End Sub


Private Sub SetCodeWarningMessage(operationname_ARG As Integer)

    Call MsgBox("Programming error", vbOKOnly, operationname_ARG & " is deprecated since, please debug")
    
End Sub



Private Sub SetStatusmessage(statusmessage_ARG As String)

    
Errorhandler:

End Sub


Private Sub Class_Initialize()
    Debug.Print "Class_Initialize.moduleName_TXT=" & moduleName_TXT
    Debug.Print "debug_FLAG=" & debug_FLAG
    
    Call workSlot304_OBJ.Workslotindex(vbWednesday, 4)
    'Call workSlot304_OBJ.AddCategory(0, 1)
    Call workSlot304_OBJ.AddDomain("zipfiles", 1)
    Call workSlot304_OBJ.AddDomain("textfiles", 2)
    Call zipFileFolder_OBJ.FullStaging_PATH(workSlot304_OBJ)

Debug.Print moduleName_TXT & "Err.Description " & missingWeeks_INT
    
    Call rawFileFolder_OBJ.FullStaging_PATH(workSlot304_OBJ, False, 2)
    
    Rem liefert Informationen aus dem Outlook Folder, in den die zafaco Reportnotifications verschoben werden
    Call outlook_OBJ.TextfileImportFromWorkslot(workSlot304_OBJ, 0)
    unreadImportWeeks_INT = outlook_OBJ.newMailInFolder_INT
    

    If debug_FLAG Or Err.Number <> 0 Then
    
        Debug.Print moduleName_TXT & ".Class_Initialize"
        Debug.Print moduleName_TXT & "Err.Number" & Err.Number
        Debug.Print moduleName_TXT & "Err.Description " & Err.Description
        Debug.Print moduleName_TXT & ".outlook_OBJ.newMailInFolder_INT=" & unreadImportWeeks_INT
        Debug.Print moduleName_TXT & ".stagingFolder_PATH=" & stagingFolder_PATH
        Debug.Print moduleName_TXT & ".rawFileFolder_OBJ=" & rawFileFolder_OBJ.staging_PATH
        Debug.Print moduleName_TXT & ".rawFileFolder_OBJ=" & rawFileFolder_OBJ.staging_PATH
         
    End If

End Sub

Private Sub Class_Terminate()

Rem hier wäre der Destruktor

End Sub
