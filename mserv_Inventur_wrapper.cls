VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mserv_Inventur_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private kategorie_PFAD As String
Private domain_PFAD As String
Private shellcommand_TXT As String

Public debug_FLAG As Boolean



Public Sub debug_me()

    Debug.Print "mserv_Inventur_wrapper.debug_me=" & inventurTextfile_URI

End Sub

Public Function inventurTextfile_URI() As String
On Error GoTo Errorhandler

Const cmdexe_PATH As String = "C:\Windows\System32\cmd.exe /C "
Const command_TXT As String = "dir /C /S "
Const inventurdatei_NAME As String = "inventur.txt"
Dim resultvalue_TXT As String
Dim kategorie_INT As Integer

    resultvalue_TXT = "<unknown>"
    shellcommand_TXT = ""
    kategorie_INT = 504

    shellcommand_TXT = shellcommand_TXT & cmdexe_PATH & " "
    shellcommand_TXT = shellcommand_TXT & command_TXT
    shellcommand_TXT = shellcommand_TXT & archivverzeichnis_NAME(kategorie_INT)
    shellcommand_TXT = shellcommand_TXT & "\*.*"
    shellcommand_TXT = shellcommand_TXT & " > "
    shellcommand_TXT = shellcommand_TXT & archivverzeichnis_NAME
    shellcommand_TXT = shellcommand_TXT & "\"
    shellcommand_TXT = shellcommand_TXT & inventurdatei_NAME

    Call runShell(shellcommand_TXT)
    resultvalue_TXT = archivverzeichnis_NAME
    resultvalue_TXT = resultvalue_TXT & "\"
    resultvalue_TXT = resultvalue_TXT & inventurdatei_NAME
    
Errorhandler:

    If debug_FLAG Then
        Debug.Print "mserv_Inventur_wrapper.inventurTextfile_URI.shellcommand_TXT=" & shellcommand_TXT
        Debug.Print "mserv_Inventur_wrapper.inventurTextfile_URI.Err.Number=" & Err.Number
        Debug.Print "mserv_Inventur_wrapper.inventurTextfile_URI.Err.Description=" & Err.Description
    End If
    
    inventurTextfile_URI = resultvalue_TXT

End Function



Public Function unarchivedDocument_TXT() As String

    shellcommand_TXT = "dir /a:a"
    'shellcommand_TXT = "dir /a:-a"
    'shellcommand_TXT = "attrib -A inventarisiere.bat"
    'shellcommand_TXT = "attrib +A inventarisiere.bat"

    runsShell (shellcommand_TXT)

End Function

Private Function archivverzeichnis_NAME(Optional kategorie_ARG As Integer = 0) As String
On Error GoTo Errorhandler
    
    Const dokumentenverzeichnis_NAME As String = "C:\Users\Public\Documents"

    Dim domain_TXT As String
    Dim discriminator_TXT As String
    Dim resultvalue_TXT As String
    Dim IsMissingKategoriepfad_FLAG As Boolean
    
    IsMissingKategoriepfad_FLAG = True
    
    resultvalue_TXT = dokumentenverzeichnis_NAME & "\" & kategorie_PFAD & "\" & domain_PFAD
    
    If kategorie_PFAD <> "<unknown>" Then IsMissingKategoriepfad_FLAG = False
    
    If IsMissingKategoriepfad_FLAG Then
 
        domain_TXT = "SBGBGF"
        discriminator_TXT = "<>"
    
        VBA.FileSystem.ChDir (dokumentenverzeichnis_NAME)
        kategorie_PFAD = Dir(kategorie_ARG & "*", vbDirectory)
        VBA.FileSystem.ChDir (kategorie_PFAD)
        domain_PFAD = Dir(domain_TXT & "*", vbDirectory)
           
        resultvalue_TXT = dokumentenverzeichnis_NAME & "\" & kategorie_PFAD & "\" & domain_PFAD
    
    End If
    
Errorhandler:

    Debug.Print "mserv_Inventur_wrapper.archivverzeichnis_NAME.resultvalue_TXT=" & resultvalue_TXT
    archivverzeichnis_NAME = resultvalue_TXT
    
    
End Function

Private Function runShell(textline_ARG As String) As Boolean
Rem soll das eigentliche Betriebssystem vom Script abschirmen
Dim resultvalue_DBL As Double

    resultvalue_DBL = Shell(textline_ARG, vbMaximizedFocus)

End Function

Private Sub Class_Initialize()

    debug_FLAG = True
    Debug.Print "mserv_Inventur_wrapper.Class_Initialize"
    Debug.Print "mserv_Inventur_wrapper.debug_FLAG=" & debug_FLAG
    kategorie_PFAD = "<unknown>"
    domain_PFAD = "<unknown>"
    
End Sub

Private Sub Class_Terminate()

    kategorie_PFAD = "<unknown>"
    domain_PFAD = "<unknown>"
    

End Sub
