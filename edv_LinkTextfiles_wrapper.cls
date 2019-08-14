VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_LinkTextfiles_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const defaultDatabase_PATH As String = "C:\Users\Public"
Private Const defaultDatabase_NAME As String = "NQ2015-Eventsource.accdb"
Private Const debug_PFAD As String = "C:\Users\Public\Documents\503_ORIONW\AMAZONDE\AMAZONDE-import.accdb"

Private workSlot_OBJ As New edv_Workslot_handler
'Private eventsource_OBJ As Access.Application
'Private workspace_OBJ As Workspace
'Private database_OBJ As Database
'Private linktable_OBJ As TableDef

Private databaseFull_PATH As String
Private sourcefile_NAME As String
Private sourcetable_NAME As String
Private definedTablesCount_INT As Integer
Private formatstring_TXT As String
Private accdbValid_TXT As String


Sub debug_me()

    ' debugTable_NAME As String = "AMAZONDE-import"
    'todo: hier debuggen
    Call workSlot_OBJ.Workslotindex(vbTuesday, 1)
    
    Const debugTable_NAME As String = "170314.csv"
    Const debugPath_NAME As String = "C:\Users\Public\Documents\201_TITAN_connectreports\"
    
    databaseFull_PATH = "C:\Users\Public\Documents\201_TITAN_connectreports\TITANS-import.accdb"
    
    Debug.Print GetTextlinkName(workSlot_OBJ, "C:\Users\Public\Documents\201_TITAN_connectreports", debugTable_NAME, "170314")
    
    'Debug.Print "TableDefIndex(" & debugTable_NAME & ")=" & TableDefIndex(debugTable_NAME)

End Sub


Function GetTextlinkName(workSlot_ARG As edv_Workslot_handler, pfad_ARG As String, datei_ARG As String, TextlinkTableName_ARG As String) As String
Rem workslot_OBJ: der Workslot
Rem pfad_ARG
Rem erzeugt einen Link zu einer Textdatei und gibt den Namen der verlinkten Datei zurück
Rem kann unter Windows 8 gelegentlich zu Performanceproblemen führen
    On Error GoTo Errorhandler
    
    Dim connectstring_TXT As String
    Dim statusmessage_TXT As String
    Dim resultvalue_TXT As String
    Dim template_OBJ As Recordset
    
    sourcefile_NAME = datei_ARG
    sourcetable_NAME = TextlinkTableName_ARG
    
    Dim newtables_CNT As Integer
    
    resultvalue_TXT = "<unknown>"
    connectstring_TXT = "Text;"
    Rem Textimportstring
    connectstring_TXT = connectstring_TXT & "DSN=" & linkSpecification & ";"
    Rem Format
    Rem der formatstring wird durch die Prozedur SetFormat eingestellt
    Rem Initialwert ist immer "delimited"
    connectstring_TXT = connectstring_TXT & "FMT=" & formatstring_TXT & ";"
    Rem Header
    connectstring_TXT = connectstring_TXT & "HDR=NO;"
    Rem IMEX (was auch immer das ist)
    connectstring_TXT = connectstring_TXT & "IMEX=" & 2 & ";"
    Rem Filetype
    connectstring_TXT = connectstring_TXT & "CharacterSet=850;"
    connectstring_TXT = connectstring_TXT & "ACCDB=YES;"
    connectstring_TXT = connectstring_TXT & "DATABASE=" & pfad_ARG
    Rem Index darf auch "0" sein, wenn die allererste Tabelle gefunden wird!
    If TableDefIndex(TextlinkTableName_ARG) >= 0 Then
        Set database_OBJ = workspace_OBJ.OpenDatabase(databaseFull_PATH)
        database_OBJ.TableDefs.Delete (TextlinkTableName_ARG)
        database_OBJ.Close
    End If
    
    Set database_OBJ = workspace_OBJ.OpenDatabase(databaseFull_PATH)
    
    Set linktable_OBJ = database_OBJ.CreateTableDef(TextlinkTableName_ARG)
    
    linktable_OBJ.SourceTableName = sourcefile_NAME
    linktable_OBJ.Connect = connectstring_TXT
    statusmessage_TXT = "<appending linktable_OBJ>"
    database_OBJ.TableDefs.Append linktable_OBJ
    '
    database_OBJ.TableDefs.Refresh
    '
    newtables_CNT = database_OBJ.TableDefs.Count
    '
    database_OBJ.Close
    '
Errorhandler:
    If Err.Number <> 0 Then
            Debug.Print Err.Number & " " & Err.Description
            Debug.Print statusmessage_TXT
            resultvalue_TXT = "<error" & Err.Number & ">"
            database_OBJ.Close
        Else
            resultvalue_TXT = sourcetable_NAME
    End If
    
    GetTextlinkName = resultvalue_TXT
    
End Function

Private Function dataBaseFullPath() As String

    Debug.Print workSlot_OBJ.Workslotindex
    dataBaseFullPath = "C:\Users\Public\Documents\201_TITAN_connectreports\TITANS-import.accdb"

End Function

Private Function linkSpecification() As String
Rem
Const importspec_NAME As String = "MSysIMEXSpecs"
Dim importspec_SQL As New edv_SQL_wrapper
Dim importfields_SQL As New edv_SQL_wrapper
Dim specification_OBJ As ImportExportSpecification
Dim resultvalue_TXT As String
Dim var_TXT As String

    Debug.Print workSlot_OBJ.Workslotindex

    resultvalue_TXT = "<unknown>"
    resultvalue_TXT = "4"

    importspec_SQL.AddSelectstatement ("*")
    importspec_SQL.AddSelectstatement ("")
    importspec_SQL.AddFromClause (importspec_NAME)
    'Call importspec_SQL.SubmitQueryStatement("link:Importspezifications", 3)

    importfields_SQL.AddSelectstatement ("* ")
    importfields_SQL.AddFromClause ("MSysIMEXColumns")
    
    'database_NAME = Application.CurrentDb.Name
    'specification_OBJ = CurrentProject.ImportExportSpecifications(3)
    linkSpecification = resultvalue_TXT
   

End Function

Private Function TableDefIndex(tableName_ARG As String) As Integer
On Error GoTo Errorhandler

Dim resultvalue_INT As Integer
Dim isTableExisting_FLAG As Boolean
    
    isTableExisting_FLAG = False
    resultvalue_INT = 0

    Set database_OBJ = workspace_OBJ.OpenDatabase(databaseFull_PATH)
    
    definedTablesCount_INT = database_OBJ.TableDefs.Count
    For resultvalue_INT = 0 To (definedTablesCount_INT - 1)
            If database_OBJ.TableDefs(resultvalue_INT).Name = tableName_ARG Then
                isTableExisting_FLAG = True
                Exit For
            End If
    Next resultvalue_INT
    
Errorhandler:

If Err.Number <> 0 Or Not isTableExisting_FLAG Then

    resultvalue_INT = -1

End If

    database_OBJ.Close
    TableDefIndex = resultvalue_INT

End Function

Private Function WorkspaceByUser() As Workspace
Rem Fehlerbehandlung weicht absichtlich von sonstigen Gepflogenheiten ab!

On Error Resume Next

    DefaultType = dbUseJet
    Dim dev_user_OBJ As User
    Debug.Print "In case of User level seucrity, please enter password!"
    Set WorkspaceByUser = Nothing
    Set WorkspaceByUser = eventsource_OBJ.DBEngine.CreateWorkspace("Eventsource_WSP", "Administrator", "acc")
    Rem das hier ist absicht. Der erste Befehl schlägt zwar fehl, aber der zweite geht durch
    Set WorkspaceByUser = eventsource_OBJ.DBEngine.CreateWorkspace("Eventsource_WSP", "admin", "acc")
      
    If Err.Number <> 0 Then
        Debug.Print "Err.Number=" & Err.Number
        Debug.Print "Err.Description=" & Err.Description
        Debug.Print "WorkspaceByUser.Name=" & WorkspaceByUser.Name
    End If
 
End Function

Sub SetFormat(formatstring_ARG As String)

    formatstring_TXT = formatstring_ARG

End Sub

Private Sub Class_Initialize()

    databaseFull_PATH = defaultDatabase_PATH & "\" & defaultDatabase_NAME
    
    If debug_PFAD <> "" Then
        databaseFull_PATH = debug_PFAD
        Debug.Print databaseFull_PATH
    End If
    
    Set eventsource_OBJ = CreateObject("Access.Application")
    Set workspace_OBJ = WorkspaceByUser
    
    definedTablesCount_INT = 0
    accdbValid_TXT = "YES"
    'Das hier sollte vom textfileimport_wrapper geliefert werden!
    formatstring_TXT = "Delimited"

End Sub

Private Sub Class_Terminate()

On Error Resume Next


End Sub
