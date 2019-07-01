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

Dim fileSystem_OBJ As New Scripting.FileSystemObject
Dim textFile_OBJ As TextStream
Dim indexedTextFile_OBJ As TextStream

Dim dateTimeText_IDX As String

Dim scanner_PATH As String
Dim unIndexedFile_NAME As String
Dim indexedFile_NAME As String
Dim currentLine_TXT As String
Dim line_IDX As Integer

    scanner_PATH = "C:\Users\Public\Documents\106\"
    
    dateTimeText_IDX = "27061353"
    
    unIndexedFile_NAME = dateTimeText_IDX & ".txt"
    indexedFile_NAME = dateTimeText_IDX & "_IDX.txt"
    line_IDX = 1

    Set textFile_OBJ = fileSystem_OBJ.opentextfile(scanner_PATH & unIndexedFile_NAME)
    Set indexedTextFile_OBJ = fileSystem_OBJ.CreateTextFile(scanner_PATH & indexedFile_NAME)
       
    Do
          
        currentLine_TXT = textFile_OBJ.ReadLine
        
        If currentLine_TXT <> "" Then
            Call indexedTextFile_OBJ.WriteLine(dateTimeText_IDX & ":" & line_IDX & ":" & currentLine_TXT)
            line_IDX = line_IDX + 1
        End If
         
    Loop Until textFile_OBJ.AtEndOfStream

End Sub




Sub test_concatText()

'Library Scripting
'    C:\Windows\SysWOW64\scrrun.dll
'    Microsoft Scripting Runtime


On Error GoTo Errorhandler

Dim fileSystem_OBJ As New Scripting.FileSystemObject
Dim textFile_OBJ As TextStream

Dim scanner_PATH As String
Dim indexedFile_NAME As String


    'scanner_PATH = "C:\Users\gruenern\OneDrive for Business\5041_SBGBGF\textfiles\"
    'indexedFile_NAME = Dir(scanner_PATH & "*.txt")
    
    scanner_PATH = "C:\Users\Public\Documents\106\"
    unIndexedFile_NAME = "18060612.txt"

Debug.Print "indexedFile_NAME=" & indexedFile_NAME

    Do Until indexedFile_NAME = ""
    
        Set textFile_OBJ = fileSystem_OBJ.opentextfile(scanner_PATH & indexedFile_NAME)
        
Debug.Print "indexedFile_NAME=" & indexedFile_NAME
        
        c02 = c02 & vbCrLf & textFile_OBJ.readall
        indexedFile_NAME = Dir
    Loop

    CreateObject("scripting.filesystemobject").CreateTextFile(scanner_PATH & "Textdatenbank.txt").Write c02


Errorhandler:

Debug.Print "Err.Number=" & Err.Number
Debug.Print "Err.Description=" & Err.Description
Debug.Print "scanner_PATH=" & scanner_PATH


End Sub


