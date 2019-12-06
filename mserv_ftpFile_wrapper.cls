VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mserv_ftpFile_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private zipFolder_OBJ As edv_FilesystemFolder_wrapper
Private filesystem_OBJ As Object

Private zipFileFolder_PATH As String

Private Const scriptfile_PATH As String = "C:\Users\gruenern\OneDrive for Business\304_CT2020_FixedVoice\sftp_script.txt"
Private Const scriptfile_NAME As String = "sftp_script.txt"
Private Const debug_FLAG As Boolean = True
Private Const version_DBL As Double = 3

Public Sub debug_me()

    Call IsGet("190805", "")

End Sub


Function IsGet(eventsource_ARG As String, cleanuptarget_ARG As String) As Boolean

Const scriptfile_PFAD As String = "C:\Users\Public\Documents"

    Call appendScript(eventsource_ARG, cleanuptarget_ARG)
    Call runShell("C:\Progra~2\WinSCP\WinSCP.com /script=""C:\Users\gruenern\ONEDRI~1\304_CT2020_FixedVoice\sftp_script.txt""")

End Function

Private Function appendScript(eventsource_ARG As String, cleanuptarget_ARG As String) As Boolean

On Error GoTo Errorhandler

'Dim filesystem_OBJ As Object
Dim file_OBJ As Object
'   wird ins initialize verlegt
    'Set filesystem_OBJ = CreateObject("Scripting.FileSystemObject")
    Set file_OBJ = filesystem_OBJ.CreateTextFile(scriptfile_PATH)
    
    zipFileFolder_PATH = "C:\Users\Public\Documents\304_CT2020_FixedVoice\zipfiles\"
    
file_OBJ.WriteLine "open sftp://FIXBENCH-INT-001@transfer.de.pri.o2.com"
file_OBJ.WriteLine "cd /IN/WEEKLY"
file_OBJ.WriteLine "get " & eventsource_ARG & "0000_RawData_CB_TE_WEEK.zip " & zipFileFolder_PATH & "" & eventsource_ARG & "0000_RawData_CB_TE_WEEK.zip"
'file_OBJ.WriteLine "get " & eventsource_ARG & "0000_RawData_LA_TE_WEEK.zip C:\Users\Public\Documents\02_FN_Benchmark\" & eventsource_ARG & "0000_RawData_LA_TE_WEEK.zip"
'file_OBJ.WriteLine "get " & eventsource_ARG & "0000_RawData_WB_TE_WEEK.zip C:\Users\Public\Documents\02_FN_Benchmark\" & eventsource_ARG & "0000_RawData_WB_TE_WEEK.zip"

'file_OBJ.WriteLine "rm " & cleanuptarget_ARG & "*.*"
file_OBJ.WriteLine "exit"
file_OBJ.Close
Set file_OBJ = Nothing
Set filesystem_OBJ = Nothing

Errorhandler:

Debug.Print Err.Number
Debug.Print Err.Description

End Function


Private Function runShell(textline_ARG As String) As Boolean
Rem soll das eigentliche Betriebssystem vom Script abschirmen
Rem hier reicht ein einfaches "Shell", da ein Programm mit Parametern aufgerufen werden soll.
Dim resultvalue_DBL As Double

    resultvalue_DBL = Shell(textline_ARG, vbMaximizedFocus)

End Function

Private Sub Class_Initialize()

    Debug.Print debug_FLAG
    Set filesystem_OBJ = CreateObject("Scripting.FileSystemObject")

End Sub

Private Sub Class_Terminate()

    Set filesystem_OBJ = Nothing
    Set zipFolder_OBJ = Nothing
    
End Sub
