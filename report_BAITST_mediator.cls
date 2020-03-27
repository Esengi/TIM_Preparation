VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "report_BAITST_mediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private outlookMail_OBJ As New edv_OutlookMail_wrapper

Private communicationMatrix_TXT() As String
Private reminderList_FLAG() As Boolean


Sub debug_me()

    Call remindFriendly

End Sub

Public Function communicationMatrixEntry_TXT(index_ARG) As String

    communicationMatrixEntry_TXT = communicationMatrix_TXT(index_ARG)

End Function

Public Function updateReminderList_FLAG(index_ARG As Integer, value_ARG As Boolean) As Boolean

    reminderList_FLAG(index_ARG) = value_ARG

End Function

Public Function remindFriendly()

    For reminderList_IDX = 0 To 4
        If reminderList_FLAG(reminderList_IDX) Then Call outlookMail_OBJ.sendMessage("Friendly reminder: please update Status sheet in BAIT Notebook", communicationMatrix_TXT(reminderList_IDX))
    Next reminderList_IDX

End Function



Private Function messageBody_TXT() As String
    Dim oneNote_LNK As String

    oneNote_LNK = "https://telefonicacorp.sharepoint.com/sites/BAI.TMENQ/_layouts/15/Doc.aspx?sourcedoc={a21506db-8e0b-4c65-95db-6e0be43aa880}&action=edit&wd=target%28Status.one%7Cc9a7b553-926d-4436-af34-aff91c09511c%2FStatus%7Cb274c5ed-395d-4723-b12f-cf99312778a3%2F%29"

End Function



Private Function communicationMatrixInitialized_FLAG() As Boolean

    ReDim communicationMatrix_TXT(6)

communicationMatrix_TXT(0) = "norman.gruener@telefonica.com"
communicationMatrix_TXT(1) = "marcin.stolarzewicz@comarch.com "
communicationMatrix_TXT(2) = "ojdiez@minsait.com"
communicationMatrix_TXT(3) = "Steffen.jantke.external@telefonica.com"
communicationMatrix_TXT(4) = "akshita.chugh.external@telefonica.com"
communicationMatrix_TXT(5) = "rakesh.rajasekar.external@telefonica.com"

    ReDim reminderList_FLAG(6)

reminderList_FLAG(0) = True
reminderList_FLAG(1) = False
reminderList_FLAG(2) = False
reminderList_FLAG(3) = False
reminderList_FLAG(4) = False
reminderList_FLAG(5) = False

End Function
 

Private Sub Class_Initialize()

    Call communicationMatrixInitialized_FLAG

End Sub

Private Sub Class_Terminate()


End Sub
