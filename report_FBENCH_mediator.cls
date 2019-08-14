VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "report_FBENCH_mediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private workSlot305_OBJ As New edv_Workslot_handler
Private workSlot505_OBJ As New edv_Workslot_handler

Private unzipFile_OBJ As New mserv_zipFile_wrapper
Private ftpFile_OBJ As New mserv_ftpFile_wrapper

Rem XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Rem XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Rem XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Const moduleName_TXT As String = "report_FBENCH_mediator"
'

Sub debug_me()

    Call ftpFile_OBJ.IsGet("190805", "")

    Debug.Print extractionStatus_TXT

End Sub

Public Function extractionStatus_TXT() As String

Dim zipFIleFolder_TXT As String
Dim zipFileFolder_OBJ As New edv_FilesystemFolder_wrapper
Dim rawFileFolder_TXT As String
Dim rawFileFolder_OBJ As New edv_FilesystemFolder_wrapper

    zipFIleFolder_TXT = "C:\Users\gruenern\OneDrive for Business\5051_CALL_Accessibility\zipfiles"
    rawFileFolder_TXT = "C:\Users\gruenern\OneDrive for Business\5051_CALL_Accessibility\textfiles"
    
    'Call zipFileFolder_OBJ.IsExisting(zipFIleFolder_TXT)

    Debug.Print moduleName_TXT & ".extractionStatus_TXT.auskommentiert:'"
    Call zipFileFolder_OBJ.FullStaging_PATH(workSlot505_OBJ)
    
    Call rawFileFolder_OBJ.IsExisting(rawFileFolder_TXT)
    
    Debug.Print zipFileFolder_OBJ.Staging_PATH
    
    
    
    Call unzipFile_OBJ.GetUnzippedTextfile(zipFileFolder_OBJ, zipFileName_TXT(1), rawFileFolder_OBJ)

End Function

Private Function zipFileName_TXT(index_ARG As Integer) As String

On Error GoTo Errorhandler


    Dim resultvalue_TXT As String

    
    Const zipFileNameTrunk_TXT As String = "_RawData_CB_TE_WEEK.zip"
      
    resultvalue_TXT = datePrefix_TXT(index_ARG) & "0000" & zipFileNameTrunk_TXT
    zipFileName_TXT = resultvalue_TXT

Errorhandler:

    Debug.Print moduleName_TXT & ".zipFileName_TXT.resultvalue_TXT=" & resultvalue_TXT

End Function


Private Function datePrefix_TXT(weeksBack_ARG) As String

    Dim firstDayOfPreviousWeek As String
    Dim yearInFileName As String
    Dim monthInFileName As String
    Dim dayInFileName As String
    
    firstDayOfPreviousWeek = Now() - Weekday(Now(), vbMonday) - (weeksBack_ARG * 7) + 1
    
    monthInFileName = Month(firstDayOfPreviousWeek)
    dayInFileName = Day(firstDayOfPreviousWeek)

    datePrefix_TXT = Right(Year(firstDayOfPreviousWeek), 2)
    If monthInFileName < 10 Then datePrefix_TXT = datePrefix_TXT & "0"
    datePrefix_TXT = datePrefix_TXT & monthInFileName
    If dayInFileName < 10 Then datePrefix_TXT = datePrefix_TXT & "0"
    datePrefix_TXT = datePrefix_TXT & dayInFileName
    
    Debug.Print moduleName_TXT & ".datePrefix_TXT.firstDayOfPreviousWeek=" & firstDayOfPreviousWeek

End Function



Private Sub SetCodeWarningMessage(operationname_ARG As Integer)

    Call MsgBox("Programming error", vbOKOnly, operationname_ARG & " is deprecated since, please debug")
    
End Sub



Private Sub SetStatusmessage(statusmessage_ARG As String)

    
Errorhandler:

End Sub


Private Sub Class_Initialize()

    
Call workSlot505_OBJ.Workslotindex(vbFriday, 5)
Call workSlot505_OBJ.AddCategory(0, 1)
Call workSlot505_OBJ.AddDomain("zipfiles", 1)
'Call workSlot505_OBJ.AddDiscriminator("1", 0)

End Sub

Private Sub Class_Terminate()

Rem hier wäre der Destruktor

End Sub
