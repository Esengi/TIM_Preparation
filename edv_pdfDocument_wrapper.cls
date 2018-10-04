VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_pdfDocument_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private pdfDocumentFolder_OBJ As New edv_FilesystemFolder_wrapper
Private logging_OBJ As New mserv_Journal_wrapper
Private workSlot_OBJ As New edv_Workslot_handler

Private countedDocuments_INT As Integer


Public Sub debug_me()

    Const workSlotNumber_INT As Integer = 5

    Debug.Print "edv_pdfDocument_wrapper.debug_me"
    
    'C:\Users\Public\Documents\205_DIAMTR_Messages
    'nehmen wir zum testen
    Debug.Print "edv_pdfDocument_wrapper.workSlot_OBJ.Workslotindex(vbTuesday, " & workSlotNumber_INT & ")=" & workSlot_OBJ.Workslotindex(vbTuesday, workSlotNumber_INT)
    Debug.Print "edv_pdfDocument_wrapper.pdfDocumentFile_NAME=" & pdfDocumentFile_NAME
    
    

End Sub



Public Function pdfDocumentFile_NAME() As String
    Dim resultvalue_TXT As String
    Dim pdfFile_IDX As Integer
    
    For pdfFile_IDX = 0 To 7
    
        resultvalue_TXT = pdfDocumentFolder_OBJ.stagingFileNameByIndex_TXT(pdfFile_IDX)
        
        If Right(resultvalue_TXT, 3) = "pdf" Then Debug.Print "edv_pdfDocument_wrapper.pdfDocumentFile_NAME.Right(resultvalue_TXT, 3)=""pdf"""
      
        Debug.Print "edv_pdfDocument_wrapper.pdfDocumentFile_NAME.resultvalue_TXT=" & resultvalue_TXT
        
     Next pdfFile_IDX
   
    
    pdfDocumentFile_NAME = resultvalue_TXT
End Function




Private Sub Class_Initialize()

    'Das bringt an der Stelle noch nichts, da der Workslot noch gar nicht bekannt ist!
    'Set pdfDocumentFolder_OBJ = pdfObjectFull_PATH
    countedDocuments_INT = 0
    
    Debug.Print "edv_PowerPoint_wrapper.pdfObjectFull_PATH.Class_Initialize"
    'Debug.Print "edv_PowerPoint_wrapper.pdfObjectFull_PATH.countedDocuments_INT=" & countedDocuments_INT

End Sub
