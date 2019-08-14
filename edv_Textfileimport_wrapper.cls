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
Private workSlot_OBJ As New edv_Workslot_handler


Private Const staging_PFAD As String = "C:\Users\gruenern\OneDrive for Business\"
Private Const discriminatorPosition_INT As Integer = 1
Rem ----------------
' was muss man eigentlich über einen Textfileimport genau wissen?
' welchen Dateityp hat sie?
Private appendix_TXT As String

' wo ist die Rohdatei und existiert sie?
Rem das klärt die filenameBySearchstring Funktion für uns
Private rawFileFolder_OBJ As New edv_FilesystemFolder_wrapper

Rem in dieser Variablen soll der Stagingpfad für den
Rem Workslot abgespeichert werden
Rem dieser ist eine "nicht-bäh" globale Variable, da sie einmal gesetzt und dann
Rem zur lebenszeit des Objekts nicht mehr verändert wird.
Private stagingFolder_OBJ As New edv_FilesystemFolder_wrapper

' Dateipfad und -Name, sowie der Typ.
Private domain_PFAD As String
Rem ---------------------------------------------------------------------------------
Private Discriminator_TXTString
Private formatstring_TXT As String
Private timeInventory_TXT As String

Private stagingFileName_TXT(28) As String
Private stagingFiles_CNT As Integer
'wozu war das noch mal gut?
Private stagingInventory_FLAG As Boolean
Private noNewMail_FLAG As Boolean
Private Const debug_FLAG As Boolean = True
'Private Const debug_FLAG As Boolean = False

Sub debug_me()

    '_____________________________
    Dim var_TXT As String
    Dim debug_DBL As Double
    Dim package_TXT As String
    Dim month_TXT As String
    
    'Debug.Print dayInventory_TXT(Now())
    
    Call runShell("copy C:\Users\gruenern\OneDrive for Business\5041_SBGBGF\textfiles\*.txt 5041_SBGBGF.txt")
         
    Debug.Print "edv_Textfileimport_wrapper.workslot_OBJ.Workslotindex(vbFriday, 4)=" & workSlot_OBJ.Workslotindex(vbFriday, 4)
    Debug.Print "edv_Textfileimport_wrapper.workslot_OBJ.Workslotindex(vbFriday, 4)=" & workSlot_OBJ.AddDomain("SBGBGF", 0)
    Debug.Print "edv_Textfileimport_wrapper.workslot_OBJ.Workslotindex(vbFriday, 4)=" & workSlot_OBJ.AddDiscriminator("CD", 0)
    Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile=" & GetCopiedTextfile(workSlot_OBJ, 1)
    'Debug.Print "edv_Textfileimport_wrapper.debug_me=" & filenameBySearchstring(workslot_OBJ, 1, "sbg")
    
   
    
    Debug.Print "IsNotExisting(""inventur.txt"")=" & IsNotExisting("inventur.txt")
    
    
End Sub

Public Function newMail_FLAG(Optional newMail_ARG As Boolean = False) As Boolean

    noNewMail_FLAG = Not newMail_ARG

    newMail_FLAG = Not noNewMail_FLAG

End Function


Rem --------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetCopiedTextfile(workSlot_ARG As edv_Workslot_handler, taskpartposition_ARG As Integer, Optional sourcefileprefix_ARG As String = "") As String
On Error GoTo Errorhandler

Dim resultvalue_TXT As String
Dim rawFileFull_PATH As String
Dim textFileFull_PATH As String

    timeInventory_TXT = sourcefileprefix_ARG
    
    Rem Prüfen, ob die Klasse bereits einen gültigen Workslot besitzt
    If workSlot_OBJ Is Nothing Then Debug.Print "edv_Textfileimport_wrapper.workslot_OBJ=Nothing"
    Debug.Print "edv_Textfileimport_wrapper.workslot_OBJ.Workslotindex=" & workSlot_OBJ.Workslotindex
    Debug.Print "edv_Textfileimport_wrapper.workslot_ARG.Workslotindex=" & workSlot_ARG.Workslotindex
    If workSlot_OBJ.Workslotindex = 0 Then Debug.Print "edv_Textfileimport_wrapper.workslot_OBJ.Workslotindex=0"
    
    Rem Workslot der Klasse überschreiben
    If workSlot_OBJ.Workslotindex = 0 Then Set workSlot_OBJ = workSlot_ARG
    
    Rem rawFileFull_PATH
    rawFileFull_PATH = ""
    rawFileFull_PATH = rawFileFull_PATH & rawFileFolder_OBJ.FullStaging_PATH(workSlot_OBJ)
    rawFileFull_PATH = rawFileFull_PATH & "\"
    'ist das erste Argument hier notwendig?
    rawFileFull_PATH = rawFileFull_PATH & filenameBySearchstring(1, "sbg")
    resultvalue_TXT = "<rawFileFull_PATH>"
    
    
    textFileFull_PATH = textFileGenerated_Path(workSlot_OBJ, 0)
    
    resultvalue_TXT = CopiedSourcefile(rawFileFull_PATH, textFileFull_PATH)
            
Errorhandler:
        
    If (Err.Number <> 0) Or debug_FLAG Then
        Debug.Print "edv_Textfileimport_wrapper.Err.Number=" & Err.Number
        Debug.Print "edv_Textfileimport_wrapper.Err.Description=" & Err.Description
        Debug.Print "edv_Textfileimport_wrapper.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.staging_PFAD=" & staging_PFAD
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.resultvalue_TXT=" & resultvalue_TXT
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.rawFileFull_PATH=" & rawFileFull_PATH
        Debug.Print "edv_Textfileimport_wrapper.GetCopiedTextfile.textFileFull_PATH=" & textFileFull_PATH

    End If
         
    GetCopiedTextfile = resultvalue_TXT

End Function

Public Function textFileStaging_Path(sequence_IDX As Integer) As edv_FilesystemFolder_wrapper
Rem gibt einen Pfad aus der Stagingmatrix per index zurück
Rem und prüft dabei auch das vorhandensein der Datei
Rem im Filesystem

'wer ruft das auf?
On Error GoTo Errorhandler

    Dim resultvalue_OBJ As New edv_FilesystemFolder_wrapper
    Dim resultvalue_TXT As String
    
    resultvalue_TXT = stagingFileName_TXT(sequence_IDX)
    Debug.Print "edv_Textfileimport_wrapper.textFileStaging_Path.textFile_Path=" & textFile_Path
    
    'VBA.FileSystem.ChDir (staging_PFAD & workslot_PATH)
    resultvalue_TXT = Dir(resultvalue_TXT)
    
    Debug.Print "edv_Textfileimport_wrapper.textFileStaging_Path.Dir(" & stagingFileName_TXT(sequence_IDX) & ")=" & resultvalue_TXT
    
    If Right(stagingFileName_TXT(sequence_IDX), Len(resultvalue_TXT)) = resultvalue_TXT Then resultvalue_TXT = stagingFileName_TXT(sequence_IDX)
    
Errorhandler:
    
    textFileStaging_Path = resultvalue_OBJ
    
End Function



Rem ---------------------------------------------------------------------------------------------------------------------------------------------------
Public Function textFileGenerated_Path(workSlot_ARG As edv_Workslot_handler, Optional sequence_IDX As Integer = 0, Optional date_ARG As Date) As String
Rem public
Rem wird verwendet von:
On Error GoTo Errorhandler

    Dim statusmessage_TXT As String
    Dim Discriminator_TXT As String

    'wo wird das workslot_OBJ eigentlich gesetzt?
    If workSlot_OBJ Is Nothing Then Debug.Print "edv_Textfileimport_wrapper.workslot_OBJ=Nothing"
    If workSlot_OBJ Is Nothing Then Set workSlot_OBJ = workSlot_ARG
    
    textFileGenerated_Path = workslotPathByWorkslot(workSlot_ARG)
    
    Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path=" & textFileGenerated_Path
    
    statusmessage_TXT = workSlot_ARG.domain_TXT(1)
        
    statusmessage_TXT = "Set initial Discriminator to " & workSlot_ARG.Discriminator_TXT(1)
    Discriminator_TXT = workSlot_ARG.Discriminator_TXT(0)
    Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.workslot_ARG.Discriminator_TXT(0)=" & workSlot_ARG.Discriminator_TXT(0)
    Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.workslot_OBJ.Discriminator_TXT(0)=" & workSlot_OBJ.Discriminator_TXT(0)
    Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.workslot_ARG.Discriminator_TXT(1)=" & workSlot_ARG.Discriminator_TXT(1)
    If workSlot_ARG.PresetFile_CNT > 0 Then Discriminator_TXT = workSlot_ARG.Discriminator_TXT(sequence_IDX)
    Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.workslot_ARG.PresetFile_CNT=" & workSlot_ARG.PresetFile_CNT
    
    Rem appendix_TXT ist eine globale Variable.
    appendix_TXT = ".txt"
    '___________________________________________________________

    
    stagingFileName_TXT(sequence_IDX) = textFile_Path & textfile_NAME(sequence_IDX)
    
    stagingFiles_CNT = stagingFiles_CNT + 1
   
Errorhandler:

    If Err.Number <> 0 Or debug_FLAG Then
    
        Debug.Print "edv_Textfileimport_wrapper.stagingfilename_TXT(sequence_IDX)=" & stagingFileName_TXT(sequence_IDX)
        Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.date_ARG=" & date_ARG
        Debug.Print "edv_Textfileimport_wrapper.timeInventory_TXT=" & timeInventory_TXT
        Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path=" & textFileGenerated_Path
        Debug.Print "edv_Textfileimport_wrapper.textFileGenerated_Path.Discriminator_TXT=" & Discriminator_TXT
            
    End If
    
    textFileGenerated_Path = stagingFileName_TXT(sequence_IDX)


End Function



Function textFileFullPath(workSlotPath_ARG As String, Optional domainPath_ARG As String, Optional discriminator_ARG As String, Optional workSlot_ARG As edv_Workslot_handler, Optional checkExistence_ARG As Boolean = True) As String
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
        workslot_PATH = workslotPathByWorkslot(workSlot_ARG)
    End If
    domain_PFAD = domainPath_ARG
    Discriminator_TXT = discriminator_ARG
    Rem Aufbau Suchpfad
    resultvalue_TXT = textFile_Path & textfile_NAME
    'durch "IsNotExisting prüfen lassen, wenn schon...
    'wird wohl nicht mehr nötig sein, da funktionalität schon von woanders übernommen...
    If checkExistence_ARG Then resultvalue_TXT = filenameBySearchstring(textFile_Path, discriminatorPosition_INT, discriminator_ARG)
    'Call IsNotExisting("", "")
    textFileFullPath = resultvalue_TXT
    
Errorhandler:
    
    Debug.Print "textFileFullPath." & Err.Number
    Debug.Print "textFileFullPath." & Err.Description
    Debug.Print "textFileFullPath.workslot_PATH=" & workslot_PATH
    Debug.Print "textFileFullPath.resultvalue_TXT=" & resultvalue_TXT
        
End Function


Rem ______________________________________________________________________________________________________

Private Function CopiedSourcefile(rawfilepfad_ARG As String, sourcefilepfad_ARG As String) As String
Rem hat die Aufgabe, Filefunktionen vor der restlichen Klasse abzuschirmen
On Error GoTo Errorhandler
Dim resultvalue_TXT As String
'diese erneute verwendung ist unnötig.
'Dim appendix_TXT As String
    resultvalue_TXT = "<unknown>"
    'appendix_TXT = ""
    'VBA.FileSystem.ChDir (rawfilepfad_ARG)
    Rem C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7.1\VBE7.DLL
    Rem FileCopy kann nur einzelne Dateien kopieren, keine Verzeichnisse!
    Rem für die Verwendung mit Verzeichnissen wäre MkDir das korrekte Property
    Call VBA.FileSystem.FileCopy(rawfilepfad_ARG, sourcefilepfad_ARG)
    resultvalue_TXT = sourcefilepfad_ARG

Errorhandler:

If Err.Number <> 0 Then

    Debug.Print "edv_Textfileimport_wrapper.CopiedSourcefile.Err.Number=" & Err.Number
    Debug.Print "edv_Textfileimport_wrapper.CopiedSourcefile.Err.Description=" & Err.Description
    Debug.Print "edv_Textfileimport_wrapper.CopiedSourcefile.staging_PFAD & rawfilepfad_ARG=" & staging_PFAD & rawfilepfad_ARG
    Debug.Print "edv_Textfileimport_wrapper.CopiedSourcefile.=" & staging_PFAD & sourcefilepfad_ARG
    
End If

CopiedSourcefile = resultvalue_TXT

End Function

Private Function workslotPathByWorkslot(workSlot_ARG As edv_Workslot_handler, Optional index_ARG As Integer = 0) As String
On Error GoTo Errorhandler
    Rem hat die Fähigkeit einen Pfad anhand des Workslots zu suchen
    Rem wichtige Funktion
    Rem allerdings: sind die Argumente in der Funktion eigentlich noch notwendig?
    Rem denn: ein Textfileimport muß nicht mehr als einen Workslot kennen...
    
    Const file_ATT As Integer = vbDirectory
    Dim workslot_PATH As String
    Dim search_TXT As String

    search_TXT = workSlot_ARG.Workslotindex & "*"
    
    If index_ARG <> 0 Then
        search_TXT = "0" & index_ARG & "*"
    End If
    'todo: staging_PFAD ist nicht richtig!
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
Rem dazu verwendet sie den Filesystemfolder_wrapper

On Error GoTo Errorhandler

    Dim resultvalue_TXT As String
    Dim workslot_IDX As Integer
    
    Dim domainPathFound_FLAG As Boolean
    
    Dim workslot_PATH As String
    
    Dim resultvalue_OBJ As New edv_FilesystemFolder_wrapper
    
    resultvalue_TXT = "<no workslot path>"
    
    workslot_IDX = workSlot_OBJ.Workslotindex
      
    Debug.Print "edv_Textfileimport_wrapper.textFile_Path.workslot_PATH=" & workslot_PATH

    'todo: hier sollte man steuern können, ob nur der workslot Pfad oder
    'auch die unterpfade angesteuert werden
    resultvalue_TXT = resultvalue_OBJ.FullStaging_PATH(workSlot_OBJ, True)
    
    If workslot_PATH <> "" Then resultvalue_TXT = workslot_PATH
    Debug.Print "edv_Textfileimport_wrapper.textFile_Path.domain_PFAD=" & domain_PFAD
    If domain_PFAD <> "<no domain path>" Then resultvalue_TXT = resultvalue_TXT & domain_PFAD & "\"
    If IsNotExisting(domain_PFAD) Then domainPathFound_FLAG = False
    
    

Errorhandler:

    If Err.Number <> 0 Or debug_FLAG Then
    
        Debug.Print "edv_Textfileimport_wrapper.textFile_Path.Err.Number=" & Err.Number
        Debug.Print "edv_Textfileimport_wrapper.textFile_Path.Err.Description=" & Err.Description
        Debug.Print "edv_Textfileimport_wrapper.textFile_Path.domainPathFound_FLAG=" & domainPathFound_FLAG
        Debug.Print "edv_Textfileimport_wrapper.textFile_Path.resultvalue_TXT=" & resultvalue_TXT
        Debug.Print "edv_Textfileimport_wrapper.textFile_Path.workslot_IDX=" & workslot_IDX
        'resultvalue_TXT = "<error in textFile_Path>"
        
    End If
    
    textFile_Path = resultvalue_TXT

    'Set staging_OBJ = resultvalue_OBJ

End Function

Private Function textfile_NAME(Optional stagingIndex_ARG As Integer = 0) As String
Rem hat die Aufgabe, das gewünschte Format des Dateinamens festzulegen
'domain_PFAD richtig handlen!
'appendix richtig handlen!
On Error GoTo Errorhandler
    Const separator_CHR As String = "-"
    Const debugSubject_TXT As String = "SBGBGF"
    
    Dim domainSubject_TXT As String
    Dim domainInfix_PFAD As String
    Dim Discriminator_TXT As String
    Dim resultvalue_TXT As String
    Dim discriminator_IDX As Integer
    
    resultvalue_TXT = ""
    Rem unterscheidung staging oder time inventory
    If stagingInventory_FLAG Then timeInventory_TXT = "00000" & stagingIndex_ARG
    resultvalue_TXT = resultvalue_TXT & timeInventory_TXT & separator_CHR
    
    Rem discriminator
    
    discriminator_IDX = 0
    Discriminator_TXT = workSlot_OBJ.Discriminator_TXT(discriminator_IDX)

    resultvalue_TXT = resultvalue_TXT & Discriminator_TXT & separator_CHR
    
    Rem domain Subject
    
    domainInfix_PFAD = domain_PFAD
    If domain_PFAD = "<no domain path>" Then domainInfix_PFAD = debugSubject_TXT
    resultvalue_TXT = resultvalue_TXT & domainInfix_PFAD
    
    Rem appendix

    resultvalue_TXT = resultvalue_TXT & appendix_TXT
    
     
Errorhandler:

    Debug.Print "edv_Textfileimport_wrapper.discriminator_IDX=" & discriminator_IDX
    Debug.Print "edv_Textfileimport_wrapper.Discriminator_TXT=" & Discriminator_TXT
    Debug.Print "edv_Textfileimport_wrapper.textFile_Name.resultvalue_TXT=" & resultvalue_TXT
    Debug.Print "edv_Textfileimport_wrapper.textFile_Name.appendix_TXT=" & appendix_TXT
    
    textfile_NAME = resultvalue_TXT
   
End Function

Private Function dayInventory_TXT(date_ARG As Date) As String

Dim resultvalue_TXT As String
    'todo: mal bei gelegenheit in eine eigene Funktion auslagern
    
    Debug.Print FileDateTime("C:\Users\gruenern\OneDrive for Business\5041_SBGBGF\000000-CD-SBGBGF.txt")
    
    resultvalue_TXT = Right(Year(date_ARG), 2)
    If Month(date_ARG) < 10 Then resultvalue_TXT = resultvalue_TXT & "0"
    resultvalue_TXT = resultvalue_TXT & Month(date_ARG)
    If Day(date_ARG) < 10 Then resultvalue_TXT = resultvalue_TXT & "0"
    resultvalue_TXT = resultvalue_TXT & Day(date_ARG)
    
    timeInventory_TXT = resultvalue_TXT

    dayInventory_TXT = resultvalue_TXT

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
Dim workslot_PATH As String
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

Private Function filenameBySearchstring(taskpartposition_ARG As Integer, taskdesc_ARG As String, Optional workSlot_ARG As edv_Workslot_handler) As String
'wer verwendet das hier?
'GetCopiedTextfile

Dim fileNotFound_FLAG As Boolean
Dim rawfile_PFAD As String
Dim rawfile_NAME As String
Dim rawfile_CNT As Integer
Dim rawfile_LIST() As String
Dim resultvalue_TXT As String
Dim file_ATT As Integer

On Error GoTo Errorhandler

    fileNotFound_FLAG = True
    file_ATT = vbDirectory
    resultvalue_TXT = "<file not found>"
    
    rawfile_PFAD = rawFileFolder_OBJ.FullStaging_PATH(workSlot_OBJ)
    VBA.FileSystem.ChDir (rawfile_PFAD & "\")
    'VBA.FileSystem.ChDir (rawfile_PFAD)
    rawfile_NAME = Dir(rawfile_PFAD & "\" & "*.*")

    While rawfile_NAME <> ""

        rawfile_CNT = rawfile_CNT + 1
        ReDim Preserve rawfile_LIST(1 To rawfile_CNT)
        rawfile_LIST(rawfile_CNT) = rawfile_NAME
                       
            If (taskdesc_ARG = Mid(rawfile_LIST(rawfile_CNT), taskpartposition_ARG, Len(taskdesc_ARG))) Then
                resultvalue_TXT = rawfile_LIST(rawfile_CNT)
                fileNotFound_FLAG = False
            End If
            
        rawfile_NAME = Dir()
            
    Wend
Errorhandler:

filenameBySearchstring = resultvalue_TXT

If Err.Number <> 0 Or fileNotFound_FLAG Or debug_FLAG Then

    Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.Err.Number=" & Err.Number
    Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.Err.Description=" & Err.Description
    Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.resultvalue_TXT=" & resultvalue_TXT
    
    If Not workSlot_ARG Is Nothing Then
        Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.workslot_ARG.Workslotindex=" & workSlot_ARG.Workslotindex
    Else
        Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.workslot_ARG=Nothing"
    End If
    If Not workSlot_OBJ Is Nothing Then
        Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.workslot_OBJ.Workslotindex=" & workSlot_OBJ.Workslotindex
    Else
        Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.workslot_OBJ=Nothing"
    End If

    Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.staging_PFAD & rawfile_PFAD=" & rawfile_PFAD
    Debug.Print "edv_Textfileimport_wrapper.FilenameBySearchstring.staging_PFAD & fileNotFound_FLAG=" & fileNotFound_FLAG
    
End If


End Function

Private Function reset_stagingfilematrix()
Rem wird verwendet

Dim indx_INT As Integer

For indx_INT = 0 To UBound(stagingFileName_TXT)
    stagingFileName_TXT(indx_INT) = ""
Next indx_INT

reset_stagingfilematrix = 0


End Function

Private Function runShell(textline_ARG As String) As Boolean
Rem soll das eigentliche Betriebssystem vom Script abschirmen
Dim resultvalue_DBL As Double

Debug.Print "runShell.textline_ARG=" & textline_ARG
    
    resultvalue_DBL = Shell(textline_ARG, vbMaximizedFocus)

End Function

Private Sub Class_Initialize()
    Set workSlot_OBJ = Nothing
    'workslot_PATH = ""
    domain_PFAD = "<no domain path>"
    'Discriminator_TXT = "<no discriminator>"
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
    Debug.Print "edv_Textfileimport_wrapper.Class_Initialize"
    Debug.Print "edv_Textfileimport_wrapper.staging_PFAD=" & staging_PFAD
    Debug.Print "edv_Textfileimport_wrapper.debug_FLAG=" & debug_FLAG
        'wo wird das workslot_OBJ eigentlich gesetzt?
        'genau hier. Okay, dann hab ich den Mechanismus nicht so ganz verstanden.
    If workSlot_OBJ Is Nothing Then Debug.Print "edv_Textfileimport_wrapper.workslot_OBJ=Nothing"

End Sub

Private Sub Class_Terminate()
    'workslot_PATH = ""
    domain_PFAD = "<no domain path>"
    'Discriminator_TXT = "<no discriminator>"
    appendix_TXT = ""
    

'Call runShell("del /S " & cleanuptarget_ARG & "*.*")
'copy %cleanupweek%0000_RawData_CB_TE_WEEK.zip N:\12_Eventsource\02_FN_Benchmark\2016\%cleanupweek%0000_RawData_CB_TE_WEEK.zip

'del /S %cleanupweek%*.*

'rmdir /S /Q %cleanupweek%0000_RawData_CB_TE_WEEK
'rmdir /S /Q %cleanupweek%0000_RawData_WB_TE_WEEK
'rmdir /S /Q %cleanupweek%0000_RawData_LA_TE_WEEK

End Sub
