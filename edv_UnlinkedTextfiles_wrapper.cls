VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_UnlinkedTextfiles_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Sub debug_me()

 'Call test_indicateFile("07161620")
 Call test_concatText

End Sub

Public Function IsExisting() As Boolean

    IsExisting = True

End Function


Public Function test_indicateFile(dateTimePrefix_ARG As String) As Boolean
'wer ruft das hier schon auf?

Dim filesystem_OBJ As New Scripting.FileSystemObject
Dim textfile_OBJ As TextStream
Dim indexedTextFile_OBJ As TextStream

Dim dateTimeText_IDX As String

Dim scanner_PATH As String
Dim unIndexedFile_NAME As String
Dim indexedFile_NAME As String
Dim currentLine_TXT As String
Dim line_IDX As Integer

    scanner_PATH = "C:\Users\Public\Documents\106\"
    
    dateTimeText_IDX = dateTimePrefix_ARG
    
    unIndexedFile_NAME = dateTimeText_IDX & ".txt"
    indexedFile_NAME = dateTimeText_IDX & "_IDX.txt"
    line_IDX = 1

    Set textfile_OBJ = filesystem_OBJ.opentextfile(scanner_PATH & unIndexedFile_NAME)
    Set indexedTextFile_OBJ = filesystem_OBJ.CreateTextFile(scanner_PATH & indexedFile_NAME)
       
    Do
          
        currentLine_TXT = textfile_OBJ.ReadLine
        
        If currentLine_TXT <> "" Then
            Call indexedTextFile_OBJ.WriteLine(dateTimeText_IDX & ":" & line_IDX & ":" & currentLine_TXT)
            line_IDX = line_IDX + 1
        End If
         
    Loop Until textfile_OBJ.AtEndOfStream

End Function


Sub test_concatText()

'Library Scripting
'    C:\Windows\SysWOW64\scrrun.dll
'    Microsoft Scripting Runtime


On Error GoTo Errorhandler

Dim filesystem_OBJ As New Scripting.FileSystemObject
Dim textfile_OBJ As TextStream

Dim scanner_PATH As String
Dim indexedFile_NAME As String


    
    'methode über vba , nur bis zu einer bestimmten Dateigröße anwendbar
    

Debug.Print "indexedFile_NAME=" & indexedFile_NAME

    Do Until indexedFile_NAME = ""
    
        Set textfile_OBJ = filesystem_OBJ.opentextfile(scanner_PATH & indexedFile_NAME)
        
Debug.Print "indexedFile_NAME=" & indexedFile_NAME
        
        c02 = c02 & vbCrLf & textfile_OBJ.readall
        indexedFile_NAME = Dir
    Loop

    CreateObject("scripting.filesystemobject").CreateTextFile(scanner_PATH & "Textdatenbank.txt").Write c02


Errorhandler:

Debug.Print "Err.Number=" & Err.Number
Debug.Print "Err.Description=" & Err.Description
Debug.Print "scanner_PATH=" & scanner_PATH


End Sub


