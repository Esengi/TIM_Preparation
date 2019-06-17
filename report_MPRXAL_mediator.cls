VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "report_MPRXAL_mediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private outlookMail_OBJ As New edv_OutlookFolder_wrapper
Private targetFolder_OBJ As New edv_FilesystemFolder_wrapper
Private workSlot_OBJ As New edv_Workslot_handler


Sub debug_me()

    
    
    

    '

End Sub

Public Function extractionStatus_TXT()

'email inhalt als text abspeichern
Call outlookMail_OBJ.TextfileImportFromWorkslot(workSlot_OBJ)
extractionStatus_TXT = "Done!"

End Function


Private Sub Class_Initialize()

    Call workSlot_OBJ.Workslotindex(vbFriday, 7)

End Sub


Private Sub Class_Terminate()

End Sub
