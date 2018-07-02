VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_FN_TIM_Preparation 
   Caption         =   "TIM Preparation"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14685
   OleObjectBlob   =   "form_FN_TIM_Preparation.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "form_FN_TIM_Preparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    
Dim form_OBJ As New Form
Dim tim_OBJ As New report_TIM_FV_mediator
Dim benchmarkPerfMon_OBJ As New report_TITANS_mediator


Private Sub UserForm_Activate()

    WorkSlot_101.Caption = tim_OBJ.storeAvailabilityStatus_TXT(1)
    Workslot_103.Caption = tim_OBJ.storeNOInfomailStatus_TXT(1)
    Workslot_104.Caption = tim_OBJ.storeProblemreportStatus_TXT(1)
    WorkSlot_501.Caption = "501_"
    
    Rem color active workslot buttons
    If tim_OBJ.storeAvailabilityStatus_TXT(2) > 0 Then Workslot_103.BackColor = &H808000
    If tim_OBJ.storeNOInfomailStatus_TXT(2) > 0 Then Workslot_103.BackColor = &H808000
    If tim_OBJ.storeProblemreportStatus_TXT(2) > 0 Then Workslot_104.BackColor = &H808000
    
    Rem color inactive workslot buttons
    Workslot_105.BackColor = &HE0E0E0

End Sub


Private Sub WorkSlot_101_Click()

    Debug.Print "form_FN_TIM_Preparation.WorkSlot_101_Click.Now()=" & Now()
    Debug.Print "Dim tim_OBJ As New report_TIM_FV_mediator"
    Debug.Print "Call tim_OBJ.storeAvailabilityStatus_TXT(0)"
    
    Call tim_OBJ.storeAvailabilityStatus_TXT(0)

End Sub


Private Sub WorkSlot_102_Click()

    Call tim_OBJ.storeCCHWKLStatus_TXT
    
End Sub


Private Sub WorkSlot_103_Click()

    Call tim_OBJ.storeNOInfomailStatus_TXT

End Sub



Private Sub Workslot_104_Click()

    Call tim_OBJ.storeProblemreportStatus_TXT(0)

Rem ist inaktiv

End Sub


Private Sub Workslot_105_Click()

Rem ist inaktiv

End Sub

Private Sub WorkSlot_106_Click()

    Workslots_1xx.BackColor = &H80000006

End Sub

Private Sub WorkSlot_201_Click()

    Call ThisOutlookSession.store_205_DIAMTR_xls

End Sub

Private Sub WorkSlot_202_Click()



End Sub



Private Sub WorkSlotButton_205_Click()

    Call ThisOutlookSession.store_205_DIAMTR_xls

End Sub

Rem Freitag _______________________________________________________________

Private Sub WorkSlot_501_Click()

    Call ThisOutlookSession.store_504_PerformanceMonitoring
    'Call benchmarkPerfMon_OBJ.storePerfMonitStatus_TXT

End Sub
