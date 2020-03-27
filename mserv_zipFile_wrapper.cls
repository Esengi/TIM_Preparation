VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mserv_zipFile_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem muß VBA Filesystem und
Rem Shell können
Rem Warum hier der Pfad zu "Program Files" auf Deutsch genannt werden MUSS
Rem (sonst funktioniert das nicht)
Rem ist mir noch immer nicht völlig verständlich
    Const classVersion_DBL As Double = 2.4
    Const programm_PFAD As String = "C:\Programme\7-zip\"
    Const programm_NAME As String = "7z.exe "
    Const debug_FLAG As Boolean = True
    'Const debug_FLAG As Boolean = False
Rem Variablen
Rem 1) Pfad zur Ablage der gezippten Datei
Rem die Stringvariable dient dem einfachen Cachen des Pfadnamens
Private zipFile_PFAD As String
Private zipFolder_OBJ As New edv_FilesystemFolder_wrapper
Rem 2) Pfad zur Ablage der ungezippten Datei
Private rawfile_PFAD As String
Private rawFolder_OBJ As New edv_FilesystemFolder_wrapper
Rem 3) Name der gezippten Datei
Private zipfile_NAME As String

 
Sub debug_me()
   
    Debug.Print "mserv_zipFile_wrapper.debug_FLAG" & debug_FLAG

End Sub



Function GetUnzippedTextfile(zipFolder_ARG As edv_FilesystemFolder_wrapper, zipFileName_ARG As String, rawFolder_ARG As edv_FilesystemFolder_wrapper) As Double
Rem wrapper, der mit 7zip unter Windows arbeitet
Rem und von der Logik der Dateinamen nichts, aber auch gar nichts wissen muß
On Error GoTo Errorhandler

    Dim zipContent_TRUNK As String
    Dim command_NAME As String
    Dim shellcommand_TXT As String
    Dim resultvalue_DBL As Double
    
    resultvalue_DBL = -9999
    
    zipContent_TRUNK = " *CALL*"
    
    Set zipFolder_OBJ = zipFolder_ARG
    Set rawFolder_OBJ = rawFolder_ARG
    zipfile_NAME = zipFileName_ARG
    resultvalue_DBL = -9998
    
    VBA.FileSystem.ChDir (zipFilePfad_TXT)
    resultvalue_DBL = -9997
    
Rem Wildcards or filenames with spaces must be quoted:
Rem    "Dir\Program files\*"
Rem     Dir\"Program files"\*
Rem from https://sevenzip.osdn.jp/chm/cmdline/syntax.htm
 'In zipFilePfad_TXT steckt ein Fehler!
    command_NAME = "e " & """" & zipFilePfad_TXT
    command_NAME = command_NAME & zipfile_NAME & """"
    
Rem -oC: MUSS zusammengeschrieben werden!

    command_NAME = command_NAME & " -o" & """" & rawFilePfad_TXT & """" & zipContent_TRUNK
    'command_NAME = command_NAME & " -o" & """" & rawFilePfad_TXT & """"
    resultvalue_DBL = -9996
    
    shellcommand_TXT = programm_PFAD & programm_NAME & command_NAME
    resultvalue_DBL = -9995

    
    If zipfile_NAME <> "" Then
        resultvalue_DBL = Shell(shellcommand_TXT, vbMaximizedFocus)
    End If
    resultvalue_DBL = -9994
       
Errorhandler:

    If Err.Number <> 0 Or debug_FLAG Then

        Debug.Print "mserv_zipFile_wrapper.GetUnzippedTextfile.Err.Number=" & Err.Number
        Debug.Print "mserv_zipFile_wrapper.GetUnzippedTextfile.Err.Description=" & Err.Description
        Debug.Print "mserv_zipFile_wrapper.GetUnzippedTextfile.debug_FLAG=" & debug_FLAG
        Debug.Print "mserv_zipFile_wrapper.GetUnzippedTextfile.resultvalue_DBL=" & resultvalue_DBL
        Debug.Print "mserv_zipFile_wrapper.GetUnzippedTextfile.zipFile_NAME=" & zipfile_NAME
        Debug.Print "mserv_zipFile_wrapper.GetUnzippedTextfile.rawFile_PFAD=" & rawfile_PFAD
        Debug.Print "mserv_zipFile_wrapper.GetUnzippedTextfile.shellcommand_TXT=" & shellcommand_TXT
    
    End If

GetUnzippedTextfile = resultvalue_DBL

End Function

Public Function zippedTextFile() As String

    zippedTextFile = zipfile_NAME

End Function


Public Function IsZipFileName(Optional zipFileName_ARG As String = "") As Boolean

On Error GoTo Errorhandler

    IsZipFileName = False
    If (zipFileName_ARG <> "" And zipfile_NAME = "") Then zipfile_NAME = zipFileName_ARG
    IsZipFileName = True
    
Errorhandler:

End Function

Private Function zipFile_SFX() As String

    zipFile_SFX = "gz"

End Function


Private Function rawFile_SFX() As String

    rawFile_SFX = "txt"

End Function

Private Function zipFilePfad_TXT() As String

On Error GoTo Errorhandler

Rem Wildcards or filenames with spaces must be quoted:
Rem    "Dir\Program files\*"
Rem     Dir\"Program files"\*
Rem from https://sevenzip.osdn.jp/chm/cmdline/syntax.htm

    
    If zipFile_PFAD = "" Then
        zipFile_PFAD = zipFolder_OBJ.staging_PATH
     End If
    
    zipFilePfad_TXT = zipFile_PFAD
    If Right(zipFile_PFAD, 1) <> "\" Then zipFilePfad_TXT = zipFile_PFAD & "\"
    
Errorhandler:

    If Err.Number <> 0 Then
    
        zipFilePfad_TXT = "<path not found>"
        Debug.Print "mserv_zipFile_wrapper.Err.Number=" & Err.Number
        Deubg.Print "mserv_zipFile_wrapper.Err.Description=" & Err.Description
   
    End If
    
    If debug_FLAG Then
        
        Debug.Print "mserv_zipFile_wrapper.debug_FLAG=" & debug_FLAG
        Debug.Print "mserv_zipFile_wrapper.zipFilePfad_TXT=" & zipFilePfad_TXT
        Debug.Print "mserv_zipFile_wrapper.zipFilePfad_TXT.Right(zipFile_PFAD, 1)=" & Right(zipFile_PFAD, 1)
    
    End If
    
    
End Function


Private Function rawFilePfad_TXT() As String

    If rawfile_PFAD = "" Then
    
        rawfile_PFAD = rawFolder_OBJ.staging_PATH
        
    End If
 
    rawFilePfad_TXT = rawfile_PFAD

End Function

Private Sub Class_Initialize()

    If debug_FLAG Then

        Debug.Print "mserv_zipFile_wrapper"
        Debug.Print "mserv_zipFile_wrapper.debug_FLAG=" & debug_FLAG
        Debug.Print "mserv_zipFile_wrapper.classVersion_DBL=" & classVersion_DBL
        Debug.Print "mserv_zipFile_wrapper.GetUnzippedTextfile.zipFile_NAME=" & zipfile_NAME
    
    End If

    zipfile_NAME = ""

End Sub

Private Sub Class_Terminate()

    If debug_FLAG Then

        Debug.Print "mserv_zipFile_wrapper.Class_Terminate"
        Debug.Print "mserv_zipFile_wrapper.debug_FLAG=" & debug_FLAG
        Debug.Print "mserv_zipFile_wrapper.zipFile_PFAD=" & zipFile_PFAD
    
    End If

End Sub
