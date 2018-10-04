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
Private Const staging_PFAD As String = "C:\Users\Public\Documents\"
Private fullStaging_PFAD As String
Private stagingFile_NAME(28) As String
Private stagingFilesCount_INT As Integer

'Private Const debug_FLAG As Boolean = False
Private Const debug_FLAG As Boolean = True

Public Sub debug_me()

    Debug.Print "edv_FilesystemFolder_wrapper.debug_me"
    

End Sub


Function IsExisting(path_ARG As String) As Boolean

On Error GoTo Errorhandler

    Dim resultvalue_FLAG As Boolean
    
    resultvalue_FLAG = False

    fullStaging_PFAD = VBA.FileSystem.Dir(path_ARG, vbNormal)
    
    resultvalue_FLAG = True
    
Errorhandler:
    
    Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.Err.Number=" & Err.Number
    Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.Err.Description=" & Err.Description
    Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.fullStaging_PFAD=" & fullStaging_PFAD
    Debug.Print "edv_FilesystemFolder_wrapper.IsExisting.path_ARG=" & path_ARG

End Function

Function IsActivated(path_ARG As String) As Boolean
Dim resultvalue_DBL As Double

    resultvalue_DBL = Shell("C:\Windows\explorer.exe " & path_ARG, vbMaximizedFocus)
    IsActivated = True
    
End Function



Function FullStaging_PATH(workSlot_ARG As edv_Workslot_handler) As String
On Error GoTo Errorhandler

'ein FilesystemFolder_wrapper muss wissen
'ob er nach dem Workslotprefix
'oder nach den Kategorien
'suchen muss
'wenn er einen Dateipfad sucht.

    If fullStaging_PFAD = "" Then
        fullStaging_PFAD = staging_PFAD & staging_PATH(workSlot_ARG)
    End If
    
    'If (fullStaging_PFAD = "") And (workSlot_ARG.Domain_TXT(1) <> "") Then
    If (workSlot_ARG.domain_TXT(1) <> "") Then
        fullStaging_PFAD = fullStaging_PFAD & "\" & workSlot_ARG.domain_TXT(1)
    End If
        
    FullStaging_PATH = fullStaging_PFAD
    
    If debug_FLAG Then
      
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH"
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.workSlot_ARG.Discriminator_TXT(1)=" & workSlot_ARG.discriminator_TXT(1)
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.workSlot_ARG.Domain_TXT(1)=" & workSlot_ARG.domain_TXT(1)
        Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.fullStaging_PFAD=" & fullStaging_PFAD
    
    End If
    
Errorhandler:
    
    Debug.Print "edv_FilesystemFolder_wrapper.FullStaging_PATH.debug_FLAG=" & debug_FLAG
    
End Function

Function staging_PATH(workSlot_ARG As edv_Workslot_handler) As String

On Error GoTo Errorhandler

Dim resultvalue_TXT As String
Dim directoryname_TXT As String


resultvalue_TXT = staging_PFAD

    VBA.FileSystem.ChDir (staging_PFAD)
    
    directoryname_TXT = VBA.FileSystem.Dir(staging_PFAD, vbDirectory)
   
    While directoryname_TXT <> ""
    
    directoryname_TXT = VBA.FileSystem.Dir()
    If workSlot_ARG.IsWorkslotprefix(directoryname_TXT) Then
        resultvalue_TXT = directoryname_TXT
        VBA.FileSystem.ChDir (directoryname_TXT)
        
    End If
    
    Wend
   
Errorhandler:

If debug_FLAG Then

    Debug.Print "edv_FilesystemFolder_wrapper.staging_PATH(workSlot_ARG As edv_Workslot_handler)"
    Debug.Print "edv_FilesystemFolder_wrapper.staging=" & resultvalue_TXT

End If
staging_PATH = resultvalue_TXT


End Function
Public Function stagingFiles_CNT(workSlot_ARG As edv_Workslot_handler) As Integer

Dim resultvalue_INT As Integer

On Error GoTo Errorhandler

    If fullStaging_PFAD = "" Then
        fullStaging_PFAD = staging_PFAD & staging_PATH(workSlot_ARG)
    End If
    
    VBA.FileSystem.ChDir (fullStaging_PFAD)
    
    stagingFile_NAME(resultvalue_INT) = VBA.FileSystem.Dir(fullStaging_PFAD & "\", vbNormal)
    
    While stagingFile_NAME(resultvalue_INT) <> ""
        resultvalue_INT = resultvalue_INT + 1
        stagingFile_NAME(resultvalue_INT) = VBA.FileSystem.Dir()
        stagingFiles_CNT = resultvalue_INT
    Wend
    
Errorhandler:

    Debug.Print "edv_FilesystemFolder_wrapper.stagingFiles_CNT.Err.Number=" & Err.Number
    Debug.Print "edv_FilesystemFolder_wrapper.stagingFiles_CNT.Err.Description=" & Err.Description
    Debug.Print "edv_FilesystemFolder_wrapper.stagingFiles_CNT.workSlot_ARG=" & workSlot_ARG.Workslotindex
    Debug.Print "edv_FilesystemFolder_wrapper.stagingFiles_CNT.fullStaging_PFAD=" & fullStaging_PFAD
    Debug.Print "edv_FilesystemFolder_wrapper.stagingFiles_CNT.stagingFile_NAME(resultvalue_INT)=" & stagingFile_NAME(resultvalue_INT)
    Debug.Print "edv_FilesystemFolder_wrapper.stagingFiles_CNT.resultvalue_INT=" & resultvalue_INT
    

End Function

Public Function stagingFileNameByIndex_TXT(index_ARG As Integer) As String

    stagingFileNameByIndex_TXT = stagingFile_NAME(index_ARG)

End Function


Function IsBatchRun(domain_ARG As String, scriptname_ARG As String) As Boolean
On Error GoTo Errorhandler

Dim path_var As String
Dim resultvalue_DBL As Double


path_var = ""
path_var = path_var & staging_PFAD
path_var = path_var & domain_ARG & "\"
path_var = path_var & scriptname_ARG & " "
'path_var = path_var & "160912 "
'path_var = path_var & "160815"

    'Debug.Print "requires sft.exe"
    resultvalue_DBL = Shell(path_var, vbMaximizedFocus)
    IsBatchRun = True
    
Errorhandler:

If Err.Number <> 0 Then
    Debug.Print Err.Description
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

Private Function testFileSystem()

'Dim logInfo As New System.IO.DirectoryInfo

'If my.Computer.FileSystem.DirectoryExists("C:\backup\logs") Then
'    logInfo = my.Computer.FileSystem.GetDirectoryInfo _
'    ("C:\backup\logs")
'End If

End Function


Private Sub Class_Initialize()

Debug.Print "edv_FilesystemFolder_wrapper.Class_Initialize"
Debug.Print "edv_FilesystemFolder_wrapper.debug_FLAG=" & debug_FLAG
Debug.Print "edv_FilesystemFolder_wrapper.stagingFilesCount_INT=" & stagingFilesCount_INT

    fullStaging_PFAD = ""

End Sub

Private Sub Class_Terminate()

    Debug.Print "edv_FilesystemFolder_wrapper.Class_Terminate"
    Debug.Print "edv_FilesystemFolder_wrapper.stagingFilesCount_INT=" & stagingFilesCount_INT
End Sub
