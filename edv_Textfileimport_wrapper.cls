VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_Textfileimport_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private workslot_OBJ As New edv_Workslot_handler

Private Const staging_PFAD As String = "C:\Users\Public\Documents\"
Private Const discriminatorPosition_INT As Integer = 1
Rem ----------------
Rem verwendet 7zip als Software
Rem Warum hier der Pfad zu "Program Files" auf Deutsch genannt werden MUSS
Rem (sonst funktioniert das nicht)
Rem ist mir noch immer nicht völlig verständlich
Private Const programm_PFAD As String = "C:\Programme\7-zip\"
Rem ---------------------
Private Const programm_NAME As String = "7z.exe "
' was muss man eigentlich über einen Textfileimport genau wissen?
' Dateipfad und -Name, sowie der Typ.
Rem in dieser Variablen soll der Stagingpfad für den
Rem Workslot abgespeichert werden
Rem dieser ist eine "nicht-bäh" globale Variable, da sie einmal gesetzt und dann
Rem zur lebenszeit des Objekts nicht mehr verändert wird.
Private workslot_PATH As String
Private domain_PFAD As String
Rem -----------------------------
Private Discriminator_TXT As String
Private timeInventory_TXT As String
Private appendix_TXT As String
Private formatstring_TXT As String
Private stagingfilename_TXT(28) As String
Private stagingFiles_CNT As Integer
Private stagingInventory_FLAG As Boolean


Sub debug_me()
    '_____________________________
    Dim var_TXT As String
    Dim debug_DBL As Double
    Dim package_TXT As String
    Dim month_TXT As String
        
    Call workslot_OBJ.Workslotindex(vbTuesday, 2)
    Call workslot_OBJ.Workslotindex(vbTuesday, 1)
    'var_TXT = "1702060000_RawData_LA_TE_WEEK"
    var_TXT = "170227"
      
    workslot_PATH = "02_FN_Benchmark\"
    domain_PFAD = "CALL"
    Discriminator_TXT = "CB"
    timeInventory_TXT = "170227"
    '______________________________
    
    
    Debug.Print "IsNotExisting(""inventur.txt"")=" & IsNotExisting("inventur.txt")
    
    
End Sub

Public Function textFileGenerated_Path(workslot_ARG As edv_Workslot_handler, Optional sequence_IDX As Integer = 0, Optional date_ARG As Date) As String
Rem public
Rem wird verwendet von:
On Error GoTo Errorhandler

Dim statusmessage_TXT As String

    If workslot_OBJ Is Nothing Then Set workslot_OBJ = workslot_ARG
    
    textFileGenerated_Path = workslotPathByWorkslot(workslot_ARG)
    Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path=" & textFileGenerated_Path
    'todo: nach dem debuggen der Funktion wieder entfernen!
    'domain_PFAD = ""
    statusmessage_TXT = "Set initial Discriminator to " & workslot_ARG.Discriminator_TXT(0)
    Discriminator_TXT = workslot_ARG.Discriminator_TXT(0)
    If workslot_ARG.PresetFile_CNT > 0 Then Discriminator_TXT = workslot_ARG.Discriminator_TXT(sequence_IDX)
    Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.workslot_ARG.PresetFile_CNT=" & workslot_ARG.PresetFile_CNT
    
    appendix_TXT = ".csv"
    '___________________________________________________________
    'todo: mal bei gelegenheit in eine eigene Funktion auslagern
    timeInventory_TXT = Right(Year(date_ARG), 2)
    If Month(date_ARG) < 10 Then timeInventory_TXT = timeInventory_TXT & "0"
    timeInventory_TXT = timeInventory_TXT & Month(date_ARG)
    If Day(date_ARG) < 10 Then timeInventory_TXT = timeInventory_TXT & "0"
    timeInventory_TXT = timeInventory_TXT & Day(date_ARG)
    'der liefert noch gelegentlich ein "\" zurück
    'warum?
    textFileGenerated_Path = textFile_Path
    
    stagingfilename_TXT(sequence_IDX) = staging_PFAD & textFile_Path & textFile_Name(sequence_IDX)
    Debug.Print "edv_Textfileimport_wrapper." & stagingfilename_TXT(sequence_IDX)
    stagingFiles_CNT = stagingFiles_CNT + 1
   
Errorhandler:

    If Err.Number <> 0 Then
    
        Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.date_ARG=" & date_ARG
        Debug.Print "edv_Textfileimport_wrapper.timeInventory_TXT=" & timeInventory_TXT
        Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path=" & textFileGenerated_Path
        Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.Discriminator_TXT=" & Discriminator_TXT
            
    End If
    
    textFileGenerated_Path = stagingfilename_TXT(sequence_IDX)


End Function

Function textFileStaging_Path(sequence_IDX As Integer) As String
Rem gibt einen Pfad aus der Stagingmatrix per index zurück
Rem und prüft dabei auch das vorhandensein der Datei
Rem im Filesystem
    Dim resultvalue_TXT As String
    
    resultvalue_TXT = stagingfilename_TXT(sequence_IDX)
    VBA.FileSystem.ChDir (staging_PFAD & workslot_PATH)
    resultvalue_TXT = Dir(resultvalue_TXT)
    
    Debug.Print "edv_Textfileimport_wrapper.textFileStaging_Path.Dir(" & stagingfilename_TXT(sequence_IDX) & ")=" & resultvalue_TXT
    
    If Right(stagingfilename_TXT(sequence_IDX), Len(resultvalue_TXT)) = resultvalue_TXT Then resultvalue_TXT = stagingfilename_TXT(sequence_IDX)
    
    textFileStaging_Path = resultvalue_TXT
End Function

Function textFileFullPath(workSlotPath_ARG As String, Optional domainPath_ARG As String, Optional discriminator_ARG As String, Optional workslot_ARG As edv_Workslot_handler, Optional checkExistence_ARG As Boolean = True) As String
    On Error GoTo Errorhandler
    Rem zentrale öffentliche Funktion!
    'wer verwendet das heute (08.03.2017) schon?
    '1. outlookFolder_wrapper
    '2. import_FBENCH_mediator
    'weitere unterscheidung sollte sein: nur Pfad zurückgeben oder
    'kontrollieren, ob datei auch vorhanden ist!
    Dim resultvalue_TXT As String
    Dim search_PATH As String
    Rem Initialisieren
    resultvalue_TXT = "<file not found>"
    Rem
    workslot_PATH = workSlotPath_ARG
    If workslot_PATH = "" Then
        workslot_PATH = workslotPathByWorkslot(workslot_ARG)
    End If
    domain_PFAD = domainPath_ARG
    Discriminator_TXT = discriminator_ARG
    Rem Aufbau Suchpfad
    resultvalue_TXT = textFile_Path & textFile_Name
    'durch "IsNotExisting prüfen lassen, wenn schon...
    'wird wohl nicht mehr nötig sein, da funktionalität schon von woanders übernommen...
    If checkExistence_ARG Then resultvalue_TXT = FilenameBySearchstring(textFile_Path, discriminatorPosition_INT, discriminator_ARG)
    'Call IsNotExisting("", "")
    textFileFullPath = resultvalue_TXT
    
Errorhandler:
    
    Debug.Print "textFileFullPath." & Err.Number
    Debug.Print "textFileFullPath." & Err.Description
    Debug.Print "textFileFullPath.workslot_PATH=" & workslot_PATH
    Debug.Print "textFileFullPath.resultvalue_TXT=" & resultvalue_TXT
        
End Function

Rem ________________________________
Rem zipFunktionen sollten im mserv_zipFile_wrapper abgehandelt werden            





Function GetCopiedTextfile(rawfilepfad_ARG As String, taskpartposition_ARG As Integer, Optional sourcefileprefix_ARG As String = "") As String
On Error GoTo Errorhandler
Rem sucht in
Rem einem Verzeichnis rawfilepfad_ARG
Dim resultvalue_TXT As String
Dim rawFileFull_PATH As String
Dim textFileFull_PATH As String

    timeInventory_TXT = sourcefileprefix_ARG

    rawFileFull_PATH = workslot_PATH & rawfilepfad_ARG
    rawFileFull_PATH = workslot_PATH & rawfilepfad_ARG & FilenameBySearchstring(rawFileFull_PATH, taskpartposition_ARG, domain_PFAD)
    'die Funktion CopiedSourcefile sollte den textFileFull_PATH kennen?
    textFileFull_PATH = textFile_Path & textFile_Name
    textFileFull_PATH = workslot_PATH & domain_PFAD & "\" & Discriminator_TXT & sourcefileprefix_ARG & domain_PFAD
    'die Funktion CopiedSourcefile sollte den textFileFull_PATH kennen?
    resultvalue_TXT = CopiedSourcefile(rawFileFull_PATH, textFileFull_PATH)
            
Errorhandler:
        
    If (Err.Number <> 0) Then
        Debug.Print "edv_Textfileimport_wrapper.Err.Number=" & Err.Number
        Debug.Print "edv_Textfileimport_wrapper.Err.Description=" & Err.Description
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.staging_PFAD=" & staging_PFAD
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.rawfilepfad_ARG=" & rawfilepfad_ARG
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.resultvalue_TXT=" & resultvalue_TXT
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.rawFileFull_PATH=" & rawFileFull_PATH
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.textFileFull_PATH=" & textFileFull_PATH
    End If
        
        
    GetCopiedTextfile = resultvalue_TXT
    

End Function

Private Function CopiedSourcefile(rawfilepfad_ARG As String, sourcefilepfad_ARG As String) As String
Rem hat die Aufgabe, Filefunktionen vor der restlichen Klasse abzuschirmen
On Error GoTo Errorhandler
Dim resultvalue_TXT As String
Dim appendix_TXT As String
    resultvalue_TXT = "<unknown>"
    appendix_TXT = ".txt"
    'appendix_TXT = ""
    'VBA.FileSystem.ChDir (rawfilepfad_ARG)
    Rem C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7.1\VBE7.DLL
    Rem FileCopy kann nur einzelne Dateien kopieren, keine Verzeichnisse!
    Rem für die Verwendung mit Verzeichnissen wäre MkDir das korrekte Property
    Call VBA.FileSystem.FileCopy(staging_PFAD & rawfilepfad_ARG, staging_PFAD & sourcefilepfad_ARG & appendix_TXT)
    resultvalue_TXT = sourcefilepfad_ARG & appendix_TXT

Errorhandler:

If Err.Number <> 0 Then

    Debug.Print Err.Number
    Debug.Print Err.Description
    Debug.Print staging_PFAD & rawfilepfad_ARG
    Debug.Print staging_PFAD & sourcefilepfad_ARG & appendix_TXT
    
End If

CopiedSourcefile = resultvalue_TXT

End Function

Private Function workslotPathByWorkslot(workslot_ARG As edv_Workslot_handler, Optional index_ARG As Integer = 0) As String
On Error GoTo Errorhandler
    Rem hat die Fähigkeit einen Pfad anhand des Workslots zu suchen
    Rem wichtige Funktion
    Rem allerdings: sind die Argumente in der Funktion eigentlich noch notwendig?
    Rem denn: ein Textfileimport muß nicht mehr als einen Workslot kennen...
    
    Const file_ATT As Integer = vbDirectory
    Dim search_TXT As String

    search_TXT = workslot_ARG.Workslotindex & "*"
    
    If index_ARG <> 0 Then
        search_TXT = "0" & index_ARG & "*"
    End If
    
    VBA.FileSystem.ChDir (staging_PFAD)
    workslot_PATH = Dir(staging_PFAD & search_TXT, file_ATT)

    
Errorhandler:
'Fehlerbehandlung einfügen!
If Err.Number <> 0 Then
    Debug.Print "edv_Textfileimport_wrapper.workslotPathByWorkslot.index_ARG=" & index_ARG
    Debug.Print "edv_Textfileimport_wrapper.workslotPathByWorkslot.staging_PFAD=" & staging_PFAD & search_TXT
End If
If workslot_PATH <> "" Then workslot_PATH = workslot_PATH & "\"
If workslot_PATH = "" Then workslot_PATH = "<workslot index not found>"

    Rem ein workslot_PATH sollte immer mit einem "\" enden!
    workslotPathByWorkslot = workslot_PATH

End Function


Private Function textFile_Path() As String
Rem hat die Aufgabe, den dateipfad nach workslot und domain-verzeichnis herzustellen
'woher kommt das domain-verzeichnis?
On Error GoTo Errorhandler

    Dim resultvalue_TXT As String
    Dim domainPathFound_FLAG As Boolean
    
    resultvalue_TXT = "<no workslot path>"
    Debug.Print "edv_Textfileimport_wrapper.textFile_Path.workslot_PATH=" & workslot_PATH
    If workslot_PATH <> "" Then resultvalue_TXT = workslot_PATH
    Debug.Print "edv_Textfileimport_wrapper.textFile_Path.domain_PFAD=" & domain_PFAD
    If domain_PFAD <> "<no domain path>" Then resultvalue_TXT = resultvalue_TXT & domain_PFAD & "\"
    If IsNotExisting(domain_PFAD) Then domainPathFound_FLAG = False

Errorhandler:

    If Err.Number <> 0 Then
    
        Debug.Print "edv_Textfileimport_wrapper.Err.Number=" & Err.Number
        Debug.Print "edv_Textfileimport_wrapper.Err.Description=" & Err.Description
        Debug.Print "edv_Textfileimport_wrapper.domainPathFound_FLAG=" & domainPathFound_FLAG
        Debug.Print "edv_Textfileimport_wrapper.resultvalue_TXT=" & resultvalue_TXT
        resultvalue_TXT = "<error in textFile_Path>"
        
    End If
    
    textFile_Path = resultvalue_TXT

End Function

Private Function textFile_Name(Optional stagingIndex_ARG As Integer = 0) As String
Rem hat die Aufgabe, das gewünschte Format des Dateinamens festzulegen
'domain_PFAD richtig handlen!
On Error GoTo Errorhandler
    Const separator_CHR As String = "-"
    Const debugSubject_TXT As String = "VIPsessionData"
    
    Dim domainSubject_TXT As String
    Dim domainInfix_PFAD As String
    
    Dim resultvalue_TXT As String
    
    resultvalue_TXT = ""
    Rem unterscheidung staging oder time inventory
    If stagingInventory_FLAG Then timeInventory_TXT = "00000" & stagingIndex_ARG
    resultvalue_TXT = resultvalue_TXT & timeInventory_TXT & separator_CHR
    
    Rem discriminator

    resultvalue_TXT = resultvalue_TXT & Discriminator_TXT & separator_CHR
    
    Rem domain Subject
    
    domainInfix_PFAD = domain_PFAD
    If domain_PFAD = "<no domain path>" Then domainInfix_PFAD = "VIPsessionData"
    resultvalue_TXT = resultvalue_TXT & domainInfix_PFAD
    
    Rem appendix

    resultvalue_TXT = resultvalue_TXT & appendix_TXT
    
    
Errorhandler:

    textFile_Name = resultvalue_TXT
   
End Function

Private Function IsNotExisting(filename_ARG As String, Optional workSlotPath_ARG As String = "") As Boolean
Rem Parameter filename_ARG: ein Dateiname, z.B. "inventur.txt", aber kein ganzer Pfad
Rem Parameter workSlotPath_ARG:
Rem die Funktion gibt ein False zurück, wenn die Datei filename_ARG gefunden wird.
Rem
'wer ruft sie auf:
'a) unzippedPathName
'b) textFilePath
'todo: Fehlerbehandlung und Debug glattziehen
Dim result_FLAG As Boolean
Dim rawfile_NAME As String
Dim workSlotFull_PATH As String
Dim statusmessage_TXT As String
Dim filenNameLength_INT As String
Dim file_ATT As Integer
Dim periodPositon_INT As Integer
        
On Error GoTo Errorhandler
    Rem initialisierungen
    statusmessage_TXT = "<initializing variables>"
    result_FLAG = True
    rawfile_NAME = ""
    file_ATT = vbDirectory
    workSlotFull_PATH = staging_PFAD & workslot_PATH
    'sollte eigentlich nicht nötig sein. Die Klasse sollte zu diesem Zeitpunkt
    'den Workslot Pfad schon kennen und gesetzt haben:
    statusmessage_TXT = "<checking for workSlotPath_ARG>"
    If workSlotPath_ARG <> "" Then workSlotFull_PATH = staging_PFAD & workSlotPath_ARG
    Rem
    statusmessage_TXT = "<detecting period in filename>"
    If filename_ARG = "" Then filename_ARG = "<missing argument>"
    statusmessage_TXT = "<calculating period postition in filename>"
    periodPositon_INT = Len(filename_ARG) - 3
    If periodPositon_INT < 1 Then periodPositon_INT = 1
    statusmessage_TXT = "<finding character at period postition>"
    If Mid(filename_ARG, periodPositon_INT, 1) = "." Then file_ATT = vbNormal
    statusmessage_TXT = "<changing into workfolder>"
    VBA.FileSystem.ChDir (workSlotFull_PATH)
    statusmessage_TXT = "<List Folder Content>"
    rawfile_NAME = Dir(workSlotFull_PATH & "*.*", file_ATT)
    While rawfile_NAME <> ""
        If (rawfile_NAME = filename_ARG) Then result_FLAG = False
        rawfile_NAME = Dir()
    Wend
        
Errorhandler:

If Err.Number <> 0 Then
    Debug.Print "edv_Textfileimport_wrapper.IsNotExisting.Err.Number=" & Err.Number
    Debug.Print "edv_Textfileimport_wrapper.IsNotExisting.Err.Description=" & Err.Description
    Debug.Print "edv_Textfileimport_wrapper.IsNotExisting.statusmessage_TXT=" & statusmessage_TXT
    Debug.Print "edv_Textfileimport_wrapper.IsNotExisting.filename_ARG=" & filename_ARG
    Debug.Print "edv_Textfileimport_wrapper.IsNotExisting.workSlotFull_PATH=" & workSlotFull_PATH
    Debug.Print "edv_Textfileimport_wrapper.IsNotExisting.workSlotFull_PATH & filename_ARG & ""*.*""=" & workSlotFull_PATH & filename_ARG & "*.*"
    result_FLAG = True
    Debug.Print "edv_Textfileimport_wrapper.IsNotExisting.result_FLAG="
End If

    IsNotExisting = result_FLAG

End Function

Private Function FilenameBySearchstring(rawfilepfad_ARG As String, taskpartposition_ARG As Integer, taskdesc_ARG As String) As String
'wer verwendet das hier?
'GetCopiedTextfile

Dim fileNotFound_FLAG As Boolean
Dim rawfile_NAME As String
Dim rawfile_CNT As Integer
Dim rawfile_LIST() As String
Dim file_ATT As Integer
On Error GoTo Errorhandler

    fileNotFound_FLAG = True
    file_ATT = vbDirectory
    VBA.FileSystem.ChDir (staging_PFAD & rawfilepfad_ARG)
    rawfile_NAME = Dir(staging_PFAD & rawfilepfad_ARG & "*.*")

    While rawfile_NAME <> ""

        rawfile_CNT = rawfile_CNT + 1
        ReDim Preserve rawfile_LIST(1 To rawfile_CNT)
        rawfile_LIST(rawfile_CNT) = rawfile_NAME
                       
            If (taskdesc_ARG = Mid(rawfile_LIST(rawfile_CNT), taskpartposition_ARG, Len(taskdesc_ARG))) Then
                FilenameBySearchstring = rawfile_LIST(rawfile_CNT)
                fileNotFound_FLAG = False
            End If
            
        rawfile_NAME = Dir()
            
    Wend
Errorhandler:

If Err.Number <> 0 Or fileNotFound_FLAG Then

    FilenameBySearchstring = "<file not found>"
    Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.staging_PFAD & rawfilepfad_ARG=" & staging_PFAD & rawfilepfad_ARG
    Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.staging_PFAD & fileNotFound_FLAG=" & fileNotFound_FLAG
End If

End Function

Private Function reset_stagingfilematrix()
Rem wird verwendet

Dim indx_INT As Integer

For indx_INT = 0 To UBound(stagingfilename_TXT)
    stagingfilename_TXT(indx_INT) = ""
Next indx_INT

reset_stagingfilematrix = 0


End Function

Private Function runShell(textline_ARG As String) As Boolean
Rem soll das eigentliche Betriebssystem vom Script abschirmen
Dim resultvalue_DBL As Double

    resultvalue_DBL = Shell(textline_ARG, vbMaximizedFocus)

End Function

Private Sub Class_Initialize()
    Set workslot_OBJ = Nothing
    workslot_PATH = ""
    domain_PFAD = "<no domain path>"
    Discriminator_TXT = "<no discriminator>"
    appendix_TXT = ""
    formatstring_TXT = "Delimited"
    stagingInventory_FLAG = True
    stagingFiles_CNT = 0
    'formatstring_TXT = "Fixed"
    Call reset_stagingfilematrix
    Rem welche mediatoren verwenden überhaupt Textfiles, also z.B. .txt und .csv?
    Rem kommen diese Textfiles aus
    Rem     a) e-mails?
    Rem     b) ftp-downloads
    Rem     c) HTML-downloads
    Rem     d) anderen Quellen?
    
    Rem zip/unzipstatus?

End Sub

Private Sub Class_Terminate()
    workslot_PATH = ""
    domain_PFAD = "<no domain path>"
    Discriminator_TXT = "<no discriminator>"
    appendix_TXT = ""
    

'Call runShell("del /S " & cleanuptarget_ARG & "*.*")
'copy %cleanupweek%0000_RawData_CB_TE_WEEK.zip N:\12_Eventsource\02_FN_Benchmark\2016\%cleanupweek%0000_RawData_CB_TE_WEEK.zip

'del /S %cleanupweek%*.*

'rmdir /S /Q %cleanupweek%0000_RawData_CB_TE_WEEK
'rmdir /S /Q %cleanupweek%0000_RawData_WB_TE_WEEK
'rmdir /S /Q %cleanupweek%0000_RawData_LA_TE_WEEK

End Sub
