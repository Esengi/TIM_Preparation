VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_Worksheet_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim excelapplication_OBJ As Object
'Dim worksheet_OBJ As New worksheet
    
   Private Const excelapplication_PFAD As String = "C:\Program Files (x86)\Microsoft Office\Office15\"
   Private Const excelapplication_NAME As String = "EXCEL.EXE"

Sub debug_me()

    Rem first i would like to get the values that i want to compare with the mail content from a Excel sheet.
    
    Rem therefore i need an Object, that can handle Excel, so i declare the Object:
    Dim excelapplication_OBJ As Object
    Rem i say that it shall be of the type Excel.Application:
    Set excelapplication_OBJ = CreateObject("Excel.Application")
    Rem i open the workbook:
    excelapplication_OBJ.Application.Workbooks.Open ("C:\Users\gruenern\Desktop\XING.xlsx")
    Rem i want to work with a worksheet:
    Dim excelworksheet_OBJ As Object
    Rem i use the first (and only) worksheet. If i had more in my workbook, this would be more complicated
    Set excelworksheet_OBJ = excelapplication_OBJ.Worksheets.Item(1)
    Call excelworksheet_OBJ.Activate
    Rem i want to store the content of a cell range in a way i can use it for searches.
    Rem therfore i declare an array of Strings
    Dim searchstring_TXT(3) As String
    Rem in order to walk through that, i also need a index or pointer variable
    Dim searchindex_INT As Integer
    For searchindex_INT = 1 To 3
        Rem i know that my values are in the "A" column in lines 1 - 3:
        searchstring_TXT(searchindex_INT) = excelworksheet_OBJ.Range("A" & searchindex_INT)
        
    Next searchindex_INT
End Sub


Function LaunchExcel(Optional excelfile_ARG As String) As Double

    LaunchExcel = Shell(excelapplication_PFAD & excelapplication_NAME & " " & excelfile_ARG, vbMaximizedFocus)

End Function

Function IsRefreshed() As Boolean
    Dim connection_NAME As String
    
    'connection_NAME = chartData_OBJ.Workbook.Connections(1).Name
    'Call chartData_OBJ.Workbook.Connections(1).Refresh
    
    
End Function

Private Sub Class_Initialize()

    Set excelapplication_OBJ = CreateObject("Excel.Application")

End Sub
