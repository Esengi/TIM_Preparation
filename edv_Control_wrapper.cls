VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_Control_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const sharepointRoot_PATH As String = "http://teamsite.de.pri.o2.com/teams/"

Private firstDayOfWeek_DAT As Date
Private meetingEventMonth_TXT As String
Rem Muss können:
Rem ein Zielverzeichnis ermitteln
Rem das macht targetFile_PATH
Rem einen Namen ermitteln
Rem das macht targetFile-NAME
Private targetFilePath_TXT As String

Private domainindex_INT As Integer

Private debug_FLAG As Boolean


Function debug_me() As String

    debug_me = "debug"
    Debug.Print workslotpointer_INT("102")
    Debug.Print workslotpointer_INT("302")
    Debug.Print slideindex_INT("102")
    Debug.Print enMonth_NAME(firstDayOfWeek)
    
End Function

Public Function targetFile_PATH(Optional domainindex_ARG As Integer) As String

    Const debug_FLAG As Boolean = True
    Rem hier ist leider die Schreibweise so gar nicht einheitlich...
    Const monthSeparator_CHR As String = "_"
    'Const monthSeparator_CHR As String = " "
        
    Dim reportTargetSubfolder_PATH As String

    On Error GoTo Errorhandler
    
    targetFilePath_TXT = sharepointRoot_PATH & "networks_servicemanagement/"
    targetFilePath_TXT = targetFilePath_TXT & "Freigegebene Dokumente/NPRM/"
    meetingEventMonth_TXT = ""

    meetingEventMonth_TXT = meetingEventMonth_TXT & Format(firstDayOfWeek(2), "yyyy") & "_" & Format(firstDayOfWeek(2), "mm")
    meetingEventMonth_TXT = meetingEventMonth_TXT & monthSeparator_CHR
    meetingEventMonth_TXT = meetingEventMonth_TXT & enMonth_NAME(firstDayOfWeek(2))
    
    targetFilePath_TXT = targetFilePath_TXT & meetingEventMonth_TXT & "/"
    targetFilePath_TXT = targetFilePath_TXT & Format((firstDayOfWeek_DAT - 1 + 3), "yyyymmdd") & "/"
    targetFile_PATH = targetFilePath_TXT
    
Errorhandler:

    If Err.Number <> 0 Or debug_FLAG Then
    
        Debug.Print "edv_Control_wrapper.targetFile_PATH.debug_FLAGr=" & debug_FLAG
        Debug.Print "edv_Control_wrapper.targetFile_PATH.Err.Number=" & Err.Number
        Debug.Print "edv_Control_wrapper.targetFile_PATH.targetFilePath_TXT=" & targetFilePath_TXT
        Debug.Print "edv_Control_wrapper.targetFile_PATH.meetingEventMonth_TXT=" & meetingEventMonth_TXT
        Debug.Print "edv_Control_wrapper.targetFile_PATH.monthSeparator_CHR=""" & monthSeparator_CHR & """"
        Debug.Print "edv_Control_wrapper.targetFile_PATH.targetFilePath_TXT=" & targetFilePath_TXT

    End If
    
End Function



Public Function targetFile_NAME(Optional domainindex_ARG As Integer) As String

    targetFile_NAME = Format((firstDayOfWeek_DAT - 1 + 3), "yyyy-mm-dd") & "-NPRM_Fixed-Voice_v0.9"

End Function

Function slideindex_INT(Optional domainindex_ARG As Integer) As Integer
domainindex_INT = domainindex_ARG
Dim workslot_IDX As Integer

Dim slideindex_ARY(20) As Integer

slideindex_ARY(2) = 8
slideindex_ARY(10) = 7

Debug.Print "edv_control_wrapper.slideindex_INT.domainindex_ARG=" & domainindex_ARG

slideindex_INT = slideindex_ARY(workslotpointer_INT(domainindex_ARG & ""))

End Function

Function chartindex_INT(Optional domainindex_ARG As Integer) As Integer

Dim chartindex_ARY(20) As Integer

chartindex_ARY(2) = 1
chartindex_ARY(10) = 2

chartindex_INT = chartindex_ARY(workslotpointer_INT(domainindex_ARG & ""))

End Function

Public Function firstDayOfWeek(Optional offset_ARG As Integer = 0) As Date
    
    firstDayOfWeek = firstDayOfWeek_DAT + offset_ARG

End Function

Private Function enMonth_NAME(date_ARG As Date) As String
    
    enMonth_NAME = Format(date_ARG, "mmm")
    enMonth_NAME = "May"
    
    Select Case Month(date_ARG)
    
        Case 1
            enMonth_NAME = "January"
        Case 2
            enMonth_NAME = "February"
            
        Case 3
            enMonth_NAME = "March"
            
        Case 5
            enMonth_NAME = "May"
    
        Case 6
            
            enMonth_NAME = "June"
    
        Case 7
    
            enMonth_NAME = "July"
            
        Case 8
    
            enMonth_NAME = "August"
            
        Case 9
    
            enMonth_NAME = "September"
            
        Case 10
    
            enMonth_NAME = "Oktober"
            
        Case 11
    
            enMonth_NAME = "November"
            
        Case 12
            Debug.Print "edv_Control_wrapper.enMonth_NAME.Month(date_ARG)=" & Month(date_ARG)
            enMonth_NAME = "December"
         Case Else
        
        Debug.Print "edv_Control_wrapper.Case Else"
        enMonth_NAME = CVErr(-10)
    
    End Select
    
Errorhandler:
 
 If Err.Number <> 0 Then
 
    Debug.Print Err.Number
    Debug.Print "edv_Control_wrapper.Err.Number=" & Err.Number
    Debug.Print "edv_Control_wrapper.Err.Description=" & Err.Description
    Debug.Print "edv_Control_wrapper.enMonth_NAME.Month(date_ARG)=" & Month(date_ARG)
    Debug.Print "edv_Control_wrapper.enMonth_NAME.enMonth_NAME=" & enMonth_NAME
    'Debug.Print "edv_Control_wrapper.enMonth_NAME.enMonth_NAME=" & enMonth_NAME
    'Debug.Print "edv_Control_wrapper.enMonth_NAME.Month(date_ARG)=" & Month(date_ARG)
 
 End If
    
End Function

Private Function workslotpointer_INT(domainindex_ARG As String) As Integer

    workslotpointer_INT = 0
    workslotpointer_INT = ((Left(domainindex_ARG, 1) - 1) * 4) + Mid(domainindex_ARG, 2, 2)
    targetFilePath_TXT = sharepointRoot_PATH & ""

End Function


Private Sub Class_Initialize()

    debug_FLAG = True

    firstDayOfWeek_DAT = Now() - Weekday(Now(), vbMonday) + 1
    
    If debug_FLAG Then
        Debug.Print "edv_Control_wrapper.Class_Initialize.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_Control_wrapper.Class_Initialize.Now()=" & Now()
        Debug.Print "edv_Control_wrapper.Class_Initialize.firstDayOfWeek_DAT=" & firstDayOfWeek_DAT
    End If

End Sub
