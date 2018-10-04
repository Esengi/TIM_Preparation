VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_Control_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fixednetVoiceLogging_OBJ As New edv_eventsource_wrapper

Private column_names(3) As String
Private domainindex_INT As Integer
Private pageIndexFromDB_INT As Integer
Private fileTypeFromDB_TXT As String

Private debug_FLAG As Boolean


Function debug_me() As String

    Dim debug_Workslot As New edv_Workslot_handler

    debug_me = "edv_Control_wrapper.debug_me"
    
    Call debug_Workslot.Workslotindex(vbMonday, 1)
    'Call debug_Workslot.Workslotindex(vbFriday, 2)
    
    'Debug.Print pageIndex_INT
    'Debug.Print fileType_TXT(debug_Workslot)
    'Debug.Print workslotpointer_INT("302")
    'Debug.Print slideindex_INT("102")
    
    'debug_me = IsLogged(debug_Workslot, "debug_me")
    
End Function

Function IsLogged(workSlot_ARG As edv_Workslot_handler, statusmessage_ARG As String) As Boolean
    Rem
    Dim inserter As New edv_SQL_wrapper

    Call inserter.AddInsertintostatement("logging ([Timestamp], ")
    Call inserter.AddSelectstatement("Now() AS myTimestamp, ", fixednetVoiceLogging_OBJ)
    Call inserter.AddInsertintostatement("Workslot , ")
    Call inserter.AddSelectstatement("""" & workSlot_ARG.Workslotindex & """ AS myWorkslot, ")
    '
    Call inserter.AddInsertintostatement("Filetype , ")
    Call inserter.AddSelectstatement("""" & fileTypeFromDB_TXT & """ AS myFiletype, ")
    '
    Call inserter.AddInsertintostatement("Pageindex, ")
    Call inserter.AddSelectstatement("""" & pageIndexFromDB_INT & """ AS myPageIndex, ")
    Call inserter.AddInsertintostatement("Statusmessage) ")
    Call inserter.AddSelectstatement("""" & statusmessage_ARG & """ AS myStatustext")
    
    IsLogged = inserter.IsExecuted
    
    If debug_FLAG Or Err.Number <> 0 Then
        Debug.Print "edv_control_wrapper.IsLogged.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_control_wrapper.IsLogged.fileTypeFromDB_TXT=" & fileTypeFromDB_TXT
        Debug.Print "edv_control_wrapper.IsLogged.pageIndexFromDB_INT=" & pageIndexFromDB_INT
        Debug.Print "edv_control_wrapper.IsLogged.statusmessage_ARG=" & statusmessage_ARG
    End If
    
End Function


Function slideindex_INT(Optional domainindex_ARG As Integer) As Integer

On Error GoTo Errorhandler
domainindex_INT = domainindex_ARG
Dim workslot_IDX As Integer

Dim slideindex_ARY(20) As Integer

slideindex_ARY(2) = 8
slideindex_ARY(10) = 7

Errorhandler:

If Err.Number <> 0 Then

    Debug.Print "edv_control_wrapper.slideindex_INT.domainindex_ARG=" & domainindex_ARG
    Debug.Print "edv_control_wrapper.slideindex_ARY=" & slideindex_ARY(workslotpointer_INT(domainindex_ARG & ""))

End If

slideindex_INT = slideindex_ARY(workslotpointer_INT(domainindex_ARG & ""))

End Function

Function chartindex_INT(Optional domainindex_ARG As Integer) As Integer

Dim chartindex_ARY(20) As Integer

chartindex_ARY(2) = 1
chartindex_ARY(10) = 2

chartindex_INT = chartindex_ARY(workslotpointer_INT(domainindex_ARG & ""))

End Function

Public Function addControlItem(index_ARG As Integer, title_ARG As String, value_ARG As Long) As Integer

    addControlItem = 1

End Function

Rem Funktion wird verwendet

'ist sie hier am richigen Platz?

Private Function workslotpointer_INT(domainindex_ARG As String) As Integer

    workslotpointer_INT = 0
    workslotpointer_INT = ((Left(domainindex_ARG, 1) - 1) * 4) + Mid(domainindex_ARG, 2, 2)
    
    Debug.Print "edv_Control_wrapper.workslotpointer_INT.domainindex_ARG=" & domainindex_ARG
    Debug.Print "edv_Control_wrapper.workslotpointer_INT=" & ((Left(domainindex_ARG, 1) - 1) * 4) + Mid(domainindex_ARG, 2, 2)
    
    
End Function

Public Function pageIndex_INT(workSlot_ARG As edv_Workslot_handler) As Integer
'wer ruft diese Funktion typischerweise auf?
'unlinkedRange_OBJ, getArrayFromPPT
'der workslot muss übergeben werden
On Error GoTo Errorhandler
    Dim reader As New edv_SQL_wrapper
    Dim myrecordset As New ADODB.Recordset
    Const target_table As String = "logging"
    Const debug_FLAG As Boolean = True
    Dim workslot_IDX As Integer
    'hier sind noch ordendlich Fehler enthalten

    workslot_IDX = 102
    workslot_IDX = workSlot_ARG.Workslotindex
    
    Call reader.AddSelectstatement("Last(" & target_table & ".Pageindex) AS lastPageIndex ", fixednetVoiceLogging_OBJ)
    Call reader.AddFromClause("" & target_table & " ")
    Call reader.AddGroupClause("" & target_table & ".Workslot ")
    Call reader.AddHavingClause("((" & target_table & ".Workslot)=" & workslot_IDX & ")")
 
    Set myrecordset = reader.queryResult
    'Set myField = myrecordset.Fields.Item("lastPageIndex")
    pageIndexFromDB_INT = myrecordset.Fields.Item("lastPageIndex")
    
Errorhandler:

If debug_FLAG Then

    Debug.Print "edv_Control_wrapper.pageIndex_INT.Now()=" & Now()
    Debug.Print "edv_Control_wrapper.pageIndex_INT.debug_FLAG=" & debug_FLAG
    Debug.Print "edv_Control_wrapper.pageIndex_INT.Err.Number=" & Err.Number
    Debug.Print "edv_Control_wrapper.pageIndex_INT.Err.Descriptionr=" & Err.Description
    Debug.Print "edv_Control_wrapper.pageIndex_INT.workslot_IDX=" & workSlot_ARG.Workslotindex
    Debug.Print "edv_Control_wrapper.pageIndex_INT.workslot_IDX=" & workslot_IDX
    Debug.Print "edv_Control_wrapper.pageIndex_INT.myrecordset.BOF=" & myrecordset.BOF
    Debug.Print "edv_Control_wrapper.pageIndex_INT.myrecordset.EOF=" & myrecordset.EOF
    Debug.Print "edv_Control_wrapper.pageIndex_INT.pageIndexFromDB_INT=" & pageIndexFromDB_INT
    
End If
    
    pageIndex_INT = pageIndexFromDB_INT

End Function

Public Function fileType_TXT(workSlot_ARG As edv_Workslot_handler) As String
'wer ruft diese Funktion typischerweise auf?
On Error GoTo Errorhandler

    Dim reader As New edv_SQL_wrapper
    Dim myrecordset As New ADODB.Recordset
    Const target_table As String = "logging"
    Const column_name As String = "Filetype"
    Const column_alias As String = "lastFileType"

    Dim workslot_IDX As Integer
    
    workslot_IDX = 101
    workslot_IDX = workSlot_ARG.Workslotindex
    
    Call reader.AddSelectstatement("Last(" & target_table & "." & column_name & ") AS " & column_alias & " ", fixednetVoiceLogging_OBJ)
    Call reader.AddFromClause("" & target_table & " ")
    Call reader.AddGroupClause("" & target_table & ".Workslot ")
    Call reader.AddHavingClause("((" & target_table & ".Workslot)=" & workslot_IDX & ")")
 
    Set myrecordset = reader.queryResult
    'Set myField = myrecordset.Fields.Item("lastPageIndex")
    fileTypeFromDB_TXT = myrecordset.Fields.Item(column_alias)
    
Errorhandler:

If debug_FLAG Then
    Debug.Print "edv_Control_wrapper.fileType_TXT.debug_FLAG=" & debug_FLAG
    Debug.Print "edv_Control_wrapper.fileType_TXT.Err.Number=" & Err.Number
    Debug.Print "edv_Control_wrapper.fileType_TXT.Err.Descriptionr=" & Err.Description
    Debug.Print "edv_Control_wrapper.fileType_TXT.workslot_IDX=" & workslot_IDX
    Debug.Print "edv_Control_wrapper.fileType_TXT.fileTypeFromDB_TXT=" & fileTypeFromDB_TXT
End If
    
    fileType_TXT = fileTypeFromDB_TXT

End Function


Private Sub Class_Initialize()

Debug.Print "edv_Control_wrapper.Class_Initialize.Now()=" & Now()

    'debug_FLAG = False
    debug_FLAG = True
    fileTypeFromDB_TXT = "00000"
    
    If debug_FLAG Then
        Debug.Print "edv_control_wrapper.Class_Initialize"
        Debug.Print "edv_control_wrapper.Class_Initialize.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_control_wrapper.fixednetVoiceLogging_OBJ.IsDataSource(""C:\Users\gruenern\OneDrive for Business\FN_VOIC-logging.accdb"")"
    End If
    Call fixednetVoiceLogging_OBJ.IsDataSource("C:\Users\gruenern\OneDrive for Business\FN_VOIC-logging.accdb")
    'fixednetVoiceLogging_OBJ

End Sub

Private Sub Class_Terminate()

    Debug.Print "edv_control_wrapper.Class_Terminate"

End Sub
