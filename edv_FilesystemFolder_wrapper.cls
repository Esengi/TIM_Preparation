VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_FilesystemFolder_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem verwendet VBA.FileSystem

Private thisFileSystem_OBJ As FileSystemObject
Private workslot_OBJ As New edv_Workslot_handler

Private fullStaging_PFAD As String
Private stagingFile_NAME(28) As String
Private stagingFilesCount_INT As Integer


Private Const explorer_PFAD As String = "C:\Windows\explorer.exe"
Private Const stagingRoot_PFAD As String = "C:\Users\Public\Documents\"
Private Const oneDriveRoot_PFAD As String = "C:\Users\gruenern\OneDrive for Business\"
Private Const moduleVersion_DBL As Double = 0.3
Private Const debug_FLAG As Boolean = False

'Private Const debug_FLAG As Boolean = True

Rem -------------------------------------------------

Public Sub debug_me()

    Dim filepfad As String
    Dim file_NAME As String
    
    filepfad = stagingRoot_PFAD & "304_CT2020_FixedVoice\zipfiles"
    file_NAME = stagingRoot_PFAD & "304_CT2020_FixedVoice\zipfiles\1909160000_RawData_CB_TE_WEEK.zip"

    Debug.Print "thisFileSystem_OBJ.FileExists(filepfad)=" & thisFileSystem_OBJ.FileExists(file_NAME)
    Debug.Print "thisFileSystem_OBJ.FileExists(filepfad)=" & thisFileSystem_OBJ.FileExists(filepfad)
    Debug.Print "thisFileSystem_OBJ.FileExists(filepfad)=" & thisFileSystem_OBJ.GetFileName(filepfad & "\*")

End Sub

Rem -----------------------------------------------------

Public Function IsExisting(path_ARG As String) As Boolean
Rem Funktion, die das Vorhandensein des Pfads prüft
Rem oder eine Datei suchen kann?

On Error GoTo Errorhandler
    Dim testFileName_TXT As String
    Dim resultvalue_FLAG As Boolean
Rem initialisierungen
    resultvalue_FLAG = False
    testFileName_TXT = ""
    
    Rem Annahme, daß ein ChDir auf einen nicht vorhandenen
Rem Pfad zu einem Fehler führt
    VBA.FileSystem.ChDir (path_ARG)
Rem Bei Fehler im ChDir werden die beiden folgenden Kommandos übersprungen
    fullStaging_PFAD = path_ARG
    resultvalue_FLAG = True
Errorhandler:
        
    IsExisting = resultvalue_FLAG
    
    If debug_FLAG Or Err.Number <> 0 Then
        Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.Err.Number=" & Err.Number
        Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.Err.Description=" & Err.Description
        Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.path_ARG=" & path_ARG
        Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.resultvalue_FLAG=" & resultvalue_FLAG
        Debug.Print "edv_FilesystemFolder_wrapper.fullStaging_PFAD=" & fullStaging_PFAD
    End If

End Function

Public Function staging_PATH() As String

On Error GoTo Errorhandler

Dim resultvalue_TXT As String
Dim directoryname_TXT As String

    resultvalue_TXT = fullStaging_PFAD
    
    If fullStaging_PFAD = "" Then
        fullStaging_PFAD = "<emptyFilePath>"
        Err.Raise (9997)
    End If
  
Errorhandler:

    If debug_FLAG Or Err.Number <> 0 Then

    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH()"
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.Err.Number=" & Err.Number
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.Err.Description=" & Err.Description
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.debug_FLAG=" & debug_FLAG
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.resultvalue_TXT=" & resultvalue_TXT
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.fullStaging_PFAD=" & fullStaging_PFAD
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.directoryname_TXT=" & directoryname_TXT
    
    End If

    staging_PATH = resultvalue_TXT

End Function

Function IsActivated(path_ARG As String) As Boolean
Dim resultvalue_DBL As Double

    resultvalue_DBL = Shell("C:\Windows\explorer.exe " & path_ARG, vbMaximizedFocus)
    IsActivated = True
    
End Function


Rem -------------------------------------------------------------------------------------------

Public Function FullStaging_PATH(workSlot_ARG As edv_Workslot_handler, Optional workSlotOnly_ARG As Boolean = False, Optional domainIndex_ARG As Integer = 0) As String
Rem per Argument zu steuern, ob der workslot Pfad mit oder ohne unterpfad
Rem zurückgegeben werden soll.
On Error GoTo Errorhandler

    Dim resultvalue_TXT As String

    If fullStaging_PFAD = "" Or fullStaging_PFAD = "<emptyFilePath>" Then
        fullStaging_PFAD = stagingRoot_PFAD & domainStaging_PFAD(workSlot_ARG, workSlotOnly_ARG, domainIndex_ARG)
    End If

    FullStaging_PATH = fullStaging_PFAD

Errorhandler:
   
    If debug_FLAG Or Err.Number <> 0 Then
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH=" & Err.Number
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.Err.Description=" & Err.Description
        Debug.Print "edv_FilesystemFolder_wrapper.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.workSlotOnly_ARG=" & workSlotOnly_ARG
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.workSlot_ARG.Discriminator_TXT(1)=" & workSlot_ARG.Discriminator_TXT(1)
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.workSlot_ARG.Domain_TXT(1)=" & workSlot_ARG.domain_TXT(1)
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.fullStaging_PFAD=" & fullStaging_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.oneDriveRoot_PFAD=" & oneDriveRoot_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.domainStaging_PFAD()=" & domainStaging_PFAD(workSlot_ARG)
    End If

End Function

Rem -----------------------------------------------------------------------------------------------------------------------------------------

Public Function IsBatchRun(scriptname_ARG As String) As Boolean
Rem nicht besonders logisch, diese Funktion hier zu halten
Rem aber wir lassen sie mal b.a.w. hier dring
On Error GoTo Errorhandler

Dim path_var As String
Dim resultvalue_DBL As Double

Debug.Print "IsBatchRun.fullStaging_PFAD=" & fullStaging_PFAD

'Call Shell("powershell -noexit -file ""C:\Users\gruenern\OneDrive for Business\5041_SBGBGF\textfiles\5041_SBGBGF.ps1", 1)

path_var = ""
path_var = path_var & fullStaging_PFAD
path_var = path_var & "\"
path_var = path_var & scriptname_ARG & " "
'path_var = path_var & "160912 "
'path_var = path_var & "160815"

    'Debug.Print "requires sft.exe"
    resultvalue_DBL = Shell(path_var, vbMaximizedFocus)
    IsBatchRun = True
    
Errorhandler:

If Err.Number <> 0 Or debug_FLAG Then

    Debug.Print "edv_FilesystemFolder_wrapper.IsBatchRun.Err.Description=" & Err.Description
    Debug.Print "edv_FilesystemFolder_wrapper.path_var=" & path_var
End If

End Function

Function IsCleanup(cleanupCut_DATE As Date) As Boolean

On Error Resume Next

    Dim wildcards_TXT As String
    Dim cleanuptarget_ARG As String
    Dim isWildcarded_FLAG As Boolean
    
    wildcards_TXT = ""
    cleanuptarget_ARG = ""

    If cleanupCut_DATE < "01.01.2019" Then isWildcarded_FLAG = True
    
    Set folder_OBJ = filesystem_OBJ.GetFolder(test_PATH)
    
    Debug.Print "edv_FilesystemFolder_wrapper.debug_me.folder_OBJ.Files.Count=" & folder_OBJ.Files.Count

    For Each File In folder_OBJ.Files
    
        Debug.Print File.Name
        Debug.Print File.DateCreated
    
    Next
        
    If Right(cleanuptarget_ARG, 1) <> "*" Then
        wildcards_TXT = "\*.*"
    End If

    'VBA.FileSystem.Kill (staging_PFAD & cleanuptarget_ARG & wildcards_TXT)
    'VBA.FileSystem.RmDir (staging_PFAD & cleanuptarget_ARG)
    
    
    Debug.Print "edv_FilesystemFolder_wrapper.On Error Resume Next"
    Debug.Print "edv_FilesystemFolder_wrapper.IsCleanup.cleanuptarget_ARG=" & cleanuptarget_ARG
    Debug.Print "edv_FilesystemFolder_wrapper.IsCleanup.wildcards_TXT=" & wildcards_TXT
    Debug.Print "edv_FilesystemFolder_wrapper.IsCleanup.cleanuptarget_ARG=" & cleanuptarget_ARG
    
    IsCleanup = True

End Function

Private Function domainStaging_PFAD(workSlot_ARG As edv_Workslot_handler, Optional workSlotOnly_ARG As Boolean = False, Optional domainIndex_ARG As Integer = 0) As String
Rem Domain entspricht dem "Diensteanbieter"
Rem und schließt implizit nach der Geschäftslogik
Rem den Workslot bzw. die Kategorien mit ein.
On Error GoTo Errorhandler

Rem die Kategorie besteht aus Workslot und Diensteanbieternummer
Dim kategorie_INT As Integer
'das könnte auch eine globale Variable sein
Dim domain_IDX As Integer
'
Dim resultvalue_TXT As String
Dim root_PFAD As String
Dim isDefault_FLAG As Boolean

Const defaultResult_TXT As String = "5041_SBGBGF\"

    resultvalue_TXT = defaultResult_TXT
    isDefault_FLAG = True
    Rem setzen der Kategorie als Suchstring
    kategorie_INT = workSlot_ARG.Workslotindex
    domain_IDX = 1
    If domainIndex_ARG > 1 Then domain_IDX = domainIndex_ARG
    'hier eine Funktion, die aus kategorie und domain eine vierstellige Zahl macht?

    root_PFAD = stagingRoot_PFAD & "\"

    VBA.FileSystem.ChDir (root_PFAD)
    
    resultvalue_TXT = Dir(root_PFAD & "\" & kategorie_INT & "*", vbDirectory)
    isDefault_FLAG = False
    
    'If (fullStaging_PFAD = "") And (workSlot_ARG.Domain_TXT(1) <> "") Then
    If (workSlot_ARG.domain_TXT(domain_IDX) <> "") And Not workSlotOnly_ARG Then
        resultvalue_TXT = resultvalue_TXT & "\" & workSlot_ARG.domain_TXT(domain_IDX)
    End If

Errorhandler:

    domainStaging_PFAD = resultvalue_TXT
    
    If debug_FLAG Or Err.Number <> 0 Or isDefault_FLAG Then

        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.Err.Number=" & Err.Number
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.Err.Description=" & Err.Description
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.isDefault_FLAG=" & isDefault_FLAG
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.oneDriveRoot_PFAD=" & oneDriveRoot_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.stagingRoot_PFAD=" & stagingRoot_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.root_PFAD=" & root_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.resultvalue_TXT=" & resultvalue_TXT
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.workSlotOnly_ARG=" & workSlotOnly_ARG
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.kategorie_INT=" & kategorie_INT
    
    End If

End Function


Private Function testFileSystem()

'Dim logInfo As New System.IO.DirectoryInfo

'If my.Computer.FileSystem.DirectoryExists("C:\backup\logs") Then
'    logInfo = my.Computer.FileSystem.GetDirectoryInfo _
'    ("C:\backup\logs")
'End If

End Function


Private Sub Class_Initialize()

    fullStaging_PFAD = ""

    Set thisFileSystem_OBJ = CreateObject("scripting.filesystemobject")

    Debug.Print "edv_FilesystemFolder_wrapper.Class_Initialize"
    Debug.Print "edv_FilesystemFolder_wrapper.debug_FLAG=" & debug_FLAG
    If debug_FLAG Then
        Debug.Print "edv_FilesystemFolder_wrapper.stagingRoot_PFAD=" & stagingRoot_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.oneDriveRoot_PFAD=" & oneDriveRoot_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.stagingFilesCount_INT=" & stagingFilesCount_INT
        Debug.Print "edv_FilesystemFolder_wrapper.fullStaging_PFAD=" & fullStaging_PFAD
    End If

End Sub

Private Sub Class_Terminate()

    Debug.Print "edv_FilesystemFolder_wrapper.Class_Terminate"
    Debug.Print "edv_FilesystemFolder_wrapper.stagingFilesCount_INT=" & stagingFilesCount_INT
End Sub

