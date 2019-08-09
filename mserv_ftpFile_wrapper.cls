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

Public Sub debug_me()

    Call IsGet("", "")

End Sub


Function IsGet(eventsource_ARG As String, cleanuptarget_ARG As String) As Boolean

Const scriptfile_PFAD As String = "C:\Users\Public\Documents"

    Call appendScript(eventsource_ARG, cleanuptarget_ARG)
    Call runShell("C:\Progra~2\WinSCP\WinSCP.com /script=""C:\Users\gruenern\ONEDRI~1\5051_CALL_Accessibility\sftp_script.txt""")

End Function

Private Function appendScript(eventsource_ARG As String, cleanuptarget_ARG As String) As Boolean

Dim filesystem_OBJ As Object
Dim file_OBJ As Object

    Set filesystem_OBJ = CreateObject("Scripting.FileSystemObject")
    Set file_OBJ = filesystem_OBJ.CreateTextFile("C:\Users\gruenern\OneDrive for Business\5051_CALL_Accessibility\sftp_script.txt")
    
file_OBJ.WriteLine "open sftp://FIXBENCH-INT-001@transfer.de.pri.o2.com"
file_OBJ.WriteLine "cd /IN/WEEKLY"
file_OBJ.WriteLine "get " & eventsource_ARG & "0000_RawData_CB_TE_WEEK.zip C:\Users\gruenern\ONEDRI~1\5051_CALL_Accessibility\zipfiles\" & eventsource_ARG & "0000_RawData_CB_TE_WEEK.zip"
'file_OBJ.WriteLine "get " & eventsource_ARG & "0000_RawData_LA_TE_WEEK.zip C:\Users\Public\Documents\02_FN_Benchmark\" & eventsource_ARG & "0000_RawData_LA_TE_WEEK.zip"
'file_OBJ.WriteLine "get " & eventsource_ARG & "0000_RawData_WB_TE_WEEK.zip C:\Users\Public\Documents\02_FN_Benchmark\" & eventsource_ARG & "0000_RawData_WB_TE_WEEK.zip"

'file_OBJ.WriteLine "rm " & cleanuptarget_ARG & "*.*"
file_OBJ.WriteLine "exit"
file_OBJ.Close
Set file_OBJ = Nothing
Set filesystem_OBJ = Nothing

End Function





Private Function runShell(textline_ARG As String) As Boolean
Rem soll das eigentliche Betriebssystem vom Script abschirmen
Rem hier reicht ein einfaches "Shell", da ein Programm mit Parametern aufgerufen werden soll.
Dim resultvalue_DBL As Double

    resultvalue_DBL = Shell(textline_ARG, vbMaximizedFocus)

End Function

