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

Private Const explorer_PFAD As String = "C:\Windows\explorer.exe"
Private Const stagingRoot_PFAD As String = "C:\Users\Public\Documents\"
Private Const oneDriveRoot_PFAD As String = "C:\Users\gruenern\OneDrive for Business\"
Private fullStaging_PFAD As String
Private stagingFile_NAME(28) As String
Private stagingFilesCount_INT As Integer

Private Const debug_FLAG As Boolean = False
'Private Const debug_FLAG As Boolean = True

Rem -------------------------------------------------

Public Sub debug_me()

    Rem Testverzeichnisse

    Dim fileSystem_OBJ As New Scripting.FileSystemObject
    Dim folder_OBJ As Scripting.Folder
    Dim test_PATH As String
    
    Rem zuweisungen
    'test_PATH = oneDriveRoot_PFAD & "5041_SBGBGF\zipfiles"
    test_PATH = oneDriveRoot_PFAD & "5041_SBGBGF\textfiles"
    
    fullStaging_PFAD = test_PATH
        
Debug.Print "edv_FilesystemFolder_wrapper.debug_me"
Debug.Print "edv_FilesystemFolder_wrapper.debug_FLAG=" & debug_FLAG
Debug.Print "edv_FilesystemFolder_wrapper.debug_me.test_PATH=" & test_PATH
    
    Call IsBatchRun("5041_SBGBGF.bat")
    
    Set folder_OBJ = fileSystem_OBJ.GetFolder(test_PATH)
    
    Debug.Print "edv_FilesystemFolder_wrapper.debug_me.folder_OBJ.Files.Count=" & folder_OBJ.Files.Count
    'die items im folder lassen sich nicht über den Zahlindex aufrufen. da muss man wohl selber was basteln...
    Debug.Print "edv_FilesystemFolder_wrapper.debug_me.folder_OBJ.Files.Item(1).DateCreated" & folder_OBJ.Files.Item("sbg_dsltest.20190310.txt.gz").DateCreated


    'Test mit existierendem Verzeichnis
    Debug.Print "edv_FilesystemFolder_wrapper.IsExisting(\SBGBGF=" & IsExisting("C:\Users\Public\Documents\504_FBENCH_perfMonitoring\SBGBGF")
    'Test mit NICHT existierendem Verzeichnis
    Debug.Print "edv_FilesystemFolder_wrapper.IsExisting(\SBCBGF=" & IsExisting("C:\Users\Public\Documents\504_FBENCH_perfMonitoring\SBCBGF")
    'Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH=" & staging_PATH

End Sub

Rem -----------------------------------------------------

Function IsExisting(path_ARG As String) As Boolean
Rem Funktion, die das Vorhandensein des Pfads prüft


On Error GoTo Errorhandler

    
    Dim resultvalue_FLAG As Boolean
Rem initialisierungen
    resultvalue_FLAG = False
        
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

Public Function Staging_PATH() As String

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
    Debug.Print "edv_FilesystemFolder_wrapper.Err.Number=" & Err.Number
    Debug.Print "edv_FilesystemFolder_wrapper.Err.Description=" & Err.Description
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.resultvalue_TXT=" & resultvalue_TXT
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.fullStaging_PFAD=" & fullStaging_PFAD
    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH.directoryname_TXT=" & directoryname_TXT
    
End If

Staging_PATH = resultvalue_TXT


End Function

Function IsActivated(path_ARG As String) As Boolean
Dim resultvalue_DBL As Double

    resultvalue_DBL = Shell("C:\Windows\explorer.exe " & path_ARG, vbMaximizedFocus)
    IsActivated = True
    
End Function


Rem -------------------------------------------------------------------------------------------

Public Function FullStaging_PATH(workSlot_ARG As edv_Workslot_handler, Optional workSlotOnly_ARG As Boolean = False) As String
Rem per Argument zu steuern, ob der workslot Pfad mit oder ohne unterpfad
Rem zurückgegeben werden soll.
On Error GoTo Errorhandler

    Dim resultvalue_TXT As String

    If fullStaging_PFAD = "" Or fullStaging_PFAD = "<emptyFilePath>" Then
        fullStaging_PFAD = oneDriveRoot_PFAD & domainStaging_PFAD(workSlot_ARG, workSlotOnly_ARG)
    End If

    FullStaging_PATH = fullStaging_PFAD

Errorhandler:

    Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.debug_FLAG=" & debug_FLAG
        If debug_FLAG Then
    
    
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH=" & Err.Number
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH" & Err.Description
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.workSlotOnly_ARG=" & workSlotOnly_ARG
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.workSlot_ARG.Discriminator_TXT(1)=" & workSlot_ARG.Discriminator_TXT(1)
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.workSlot_ARG.Domain_TXT(1)=" & workSlot_ARG.domain_TXT(1)
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.fullStaging_PFAD=" & fullStaging_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.oneDriveRoot_PFAD=" & oneDriveRoot_PFAD
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.domainStaging_PFAD()=" & domainStaging_PFAD(workSlot_ARG)

    End If

End Function

Rem -----------------------------------------------------------------------------------------------------------------------------------------

Function IsBatchRun(scriptname_ARG As String) As Boolean
On Error GoTo Errorhandler

Dim path_var As String
Dim resultvalue_DBL As Double

Debug.Print "IsBatchRun.fullStaging_PFAD=" & fullStaging_PFAD

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

If Err.Number <> 0 Then
    Debug.Print "IsBatchRun.Err.Description=" & Err.Description
    Debug.Print "path_var=" & path_var
End If

End Function

Function IsCleanup(cleanuptarget_ARG As String) As Boolean

On Error Resume Next

    Dim wildcards_TXT As String
    
    wildcards_TXT = ""
    
    If Right(cleanuptarget_ARG, 1) <> "*" Then
        wildcards_TXT = "\*.*"
    End If

    VBA.FileSystem.Kill (staging_PFAD & cleanuptarget_ARG & wildcards_TXT)
    VBA.FileSystem.RmDir (staging_PFAD & cleanuptarget_ARG)

End Function

Private Function domainStaging_PFAD(workSlot_ARG As edv_Workslot_handler, Optional workSlotOnly_ARG As Boolean = False) As String
Rem Domain entspricht dem "Diensteanbieter"
Rem und schließt implizit nach der Geschäftslogik
Rem den Workslot bzw. die Kategorien mit ein.
On Error GoTo Errorhandler

Rem die Kategorie besteht aus Workslot und Diensteanbieternummer
Dim kategorie_INT As Integer
Dim resultvalue_TXT As String
Dim isDefault_FLAG As Boolean

Const defaultResult_TXT As String = "5041_SBGBGF\"

    isDefault_FLAG = True

    kategorie_INT = workSlot_ARG.Workslotindex

    VBA.FileSystem.ChDir (oneDriveRoot_PFAD & "\")
    'VBA.FileSystem.ChDir (rawfile_PFAD)
    resultvalue_TXT = Dir(oneDriveRoot_PFAD & "\" & "504" & "*", vbDirectory)

    resultvalue_TXT = "5041_SBGBGF\"
    'If (fullStaging_PFAD = "") And (workSlot_ARG.Domain_TXT(1) <> "") Then
    If (workSlot_ARG.domain_TXT(1) <> "") And Not workSlotOnly_ARG Then
        resultvalue_TXT = resultvalue_TXT & workSlot_ARG.domain_TXT(1)
    End If

Errorhandler:

    domainStaging_PFAD = resultvalue_TXT
    
    If debug_FLAG Or Err.Number <> 0 Or isDefault_FLAG Then

        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.Err.Number=" & Err.Number
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.Err.Description=" & Err.Description
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.isDefault_FLAG=" & isDefault_FLAG
        Debug.Print "edv_FilesystemFolder_wrapper.domainStaging_PFAD.oneDriveRoot_PFAD=" & oneDriveRoot_PFAD
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

Debug.Print "edv_FilesystemFolder_wrapper.Class_Initialize"
Debug.Print "edv_FilesystemFolder_wrapper.debug_FLAG=" & debug_FLAG
Debug.Print "edv_FilesystemFolder_wrapper.stagingRoot_PFAD=" & stagingRoot_PFAD
Debug.Print "edv_FilesystemFolder_wrapper.oneDriveRoot_PFAD=" & oneDriveRoot_PFAD
Debug.Print "edv_FilesystemFolder_wrapper.stagingFilesCount_INT=" & stagingFilesCount_INT
Debug.Print "edv_FilesystemFolder_wrapper.fullStaging_PFAD=" & fullStaging_PFAD

End Sub

Private Sub Class_Terminate()

    Debug.Print "edv_FilesystemFolder_wrapper.Class_Terminate"
    Debug.Print "edv_FilesystemFolder_wrapper.stagingFilesCount_INT=" & stagingFilesCount_INT
End Sub

