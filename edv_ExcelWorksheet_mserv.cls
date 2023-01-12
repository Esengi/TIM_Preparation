VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_ExcelWorksheet_mserv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const debug_FLAG As Boolean = True
Private Const module_NAME As String = "edv_ExcelWorksheet_mserv"
Private Const module_Version As Double = 0.3
Rem Microsoft Excel 16.0 Object Library
Private edv_Basis_OBJ As Excel.Application
Private edv_File_OBJ As Excel.Workbook
Private edv_Worksheet_OBJ As New Excel.Worksheet
Private edv_Cellrange_OBJ As Range
Private cells_OBJ As Range

Private sourcefile_PATH As String
Private sourcefile_NAME As String
Private sourcetsheet_NAME As String

Private debug_PFAD As String
Private statusmessage_TXT As String



Sub debug_me()

Dim debug_range_OBJ As Range
Dim debug_columns_INT As Integer
Dim debug_row_INT As Integer
Dim debug_value_TXT As String

On Error Resume Next
    
    
    sourcefile_PATH = "C:\Users\gruenern\OneDrive - Telefonica\Microsoft Teams-Chatdateien"
    Debug.Print "sourcefile_PATH=" & sourcefile_PATH
    sourcefile_NAME = "read_FLOW_REP_T_SAP_REPORT.xlsx"
    'sourcefile_NAME = "Tableaunutzer mit Login vor 12 Monaten.xlsx"
    Debug.Print "sourcefile_NAME=" & sourcefile_NAME
    sourcetsheet_NAME = "read_FLOW_REP_T_SAP_REPORT"
    Debug.Print module_NAME & ".debug_me.sourcetsheet_NAME=" & sourcetsheet_NAME

    
    If debug_FLAG Then
    
    
    End If
    
    Set debug_range_OBJ = Get_Cellrange
            
    debug_columns_INT = 1
    debug_row_INT = 1
    debug_value_TXT = debug_range_OBJ.Cells(debug_row_INT, debug_columns_INT).Value
        
    Debug.Print module_NAME & ".debug_me.debug_value_TXT=" & debug_value_TXT
    
    
    Call edv_File_OBJ.Sheets.Add
    
    Sheets("Tabelle1").Select
    edv_File_OBJ.SaveAs FileName:= _
        "https://telefonicacorp-my.sharepoint.com/personal/norman_gruener_telefonica_com/Documents/Microsoft%20Teams-Chatdateien/Tabelle1.txt" _
        , FileFormat:=xlText, CreateBackup:=False
    'edv_File_OBJ.SaveAs
    edv_File_OBJ.Close
    
End Sub

Public Function Statusmessage() As String

    Statusmessage = statusmessage_TXT

End Function






Private Function Get_Cellrange() As Range

Dim resultvalue_OBJ As Range

On Error GoTo Errorhandler

    Set edv_File_OBJ = edv_Basis_OBJ.Workbooks.Open(sourcefile_PATH & "\" & sourcefile_NAME)
    
    If debug_FLAG Then
        Debug.Print module_NAME & ".edv_File_OBJ.Name=" & edv_File_OBJ.Name
        Debug.Print module_NAME & ".edv_File_OBJ.Path=" & edv_File_OBJ.Path
        Debug.Print module_NAME & ".edv_File_OBJ.Worksheets.Count=" & edv_File_OBJ.Worksheets.Count
        Debug.Print module_NAME & ".edv_File_OBJ.Worksheets.Item(1)=" & edv_File_OBJ.Worksheets.Item(1).Name
    End If
 
    Set edv_Worksheet_OBJ = edv_File_OBJ.Worksheets(sourcetsheet_NAME)
    
    If debug_FLAG Then
        Debug.Print module_NAME & ".edv_Worksheet_OBJ.Name=" & edv_Worksheet_OBJ.Name
        Debug.Print module_NAME & ".edv_Worksheet_OBJ.Cells.CountLarge=" & edv_Worksheet_OBJ.Cells.CountLarge
    End If
    
    Set resultvalue_OBJ = edv_Worksheet_OBJ.Cells

Errorhandler:

    If Err.Number <> 0 Or debug_FLAG Then
        'Debug.Print module_NAME & ".Get_Cellrange.resultvalue_OBJ.Name=" & resultvalue_OBJ.Name
        'Set resultvalue_OBJ = Nothing
    End If

Set Get_Cellrange = resultvalue_OBJ

End Function

Private Sub Class_Initialize()

    Set edv_Basis_OBJ = CreateObject("Excel.Application")
    Debug.Print module_NAME & ".Class_Initialize.module_Version=" & module_Version
    
End Sub

Private Sub Class_Terminate()
 
    edv_File_OBJ.Close
    Set edv_Basis_OBJ = Nothing
 
End Sub


