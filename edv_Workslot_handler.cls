VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_Workslot_handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Const version_INT As Integer = 2
'Const debug_FLAG As Boolean = True
Const debug_FLAG As Boolean = False

Private domain_OBJ As New edv_Domain_handler

'Die hierarchie durch gegenseitige Komposition errichten
Private domainMatrix_TXT(10) As String
Private discriminatorMatrix_TXT(28) As String
Rem die Kategorie ist bestandteil des Workslots. In einem Workslot wird an
Rem einer kategorie gearbeitetz
Rem und nur an einer, daher macht eine Kategorienmatrix keinen Sinn.
Private category_INT As Integer
Rem Matrix der Discriminatoren
'Private discriminatorMatrix_OBJ(28) As edv_Discriminator_handler
Rem Matrix der Domänen
'Private domainMatrix_OBJ(28) As edv_Domain_handler
Rem Komposition ergibt sich schon daraus, daß control
Rem ein "wrapper" ist und workslot ein "handler"
Rem daraus ergibt sich auch, dass der Workslot Handler nicht aus der Architektur entfallen darf
Rem sondern eine Rolle in der Komposition, dem zusammenbau
Rem der darunterliegenden Klassen hat
'Private control_OBJ As New edv_Control_wrapper

Private Weekday_INT As Integer
Private Timeslot_INT As Integer
Rem was macht dieser Int?
Private domains_CNT As Integer
Private Discriminator_CNT As Integer
Private PresetDiscriminator_INT As Integer

Public Sub debug_me()

        Debug.Print "edv_Workslot_handler.Class_Initialize.AddCategory=" & Workslotindex(0)
        Debug.Print "edv_Workslot_handler.Class_Initialize.AddCategory=" & AddCategory(5, 12)
        Call AddDiscriminator(AB, 0)

End Sub

Public Function IsLogged(statusmessage_ARG As String) As Boolean

    IsLogged = control_OBJ.IsLogged(Me, statusmessage_ARG)

End Function

Public Function AddDomain(domain_ARG As String, preset_ARG As Integer) As Integer
On Error GoTo Errorhandler
Rem die Domain kann eine Zeichenfolge aus Buchstaben sein
Rem oder eine Zahl (als zweizeichige Folge abgespeichert)
Rem die der "Serviceprovider ID" entspricht
    domains_CNT = domains_CNT + 1
    domainMatrix_TXT(domains_CNT) = domain_ARG
    AddDomain = domains_CNT
     
Errorhandler:

    Debug.Print "edv_Workslot_handler.AddDomain.debug_FLAG=" & debug_FLAG
    If Err.Number <> 0 Or debug_FLAG Then
        Debug.Print "edv_Workslot_handler.AddDomain.domain_ARG=" & domain_ARG
        Debug.Print "edv_Workslot_handler.domains_CNT=" & domains_CNT
    End If

End Function


Public Function AddDiscriminator(discriminator_ARG As String, preset_ARG As Integer) As Integer
'todo: in eine neue Klasse domain_handler auslagern
'todo: wer verwendet das zur Zeit?
        
    discriminatorMatrix_TXT(Discriminator_CNT) = discriminator_ARG
    
    Discriminator_CNT = Discriminator_CNT + 1
    PresetDiscriminator_INT = PresetDiscriminator_INT + preset_ARG
    AddDiscriminator = Discriminator_CNT
    
End Function

Public Function AddCategory(Optional grossArea_ARG As Integer, Optional subCategory_ARG As Integer, Optional preset_ARG As Integer) As Integer
Rem Funktion und Logik dürfen bis auf weiteres erst mal in der Klasse bleiben
Rem die grossArea ist sowieso zweistellig (geht von 01 bis etwa 25)
Rem die subCategory ist ebenfalls zweistellig, hat aber eine logische unterteilung
Rem nach der Dezimalstelle

Dim resultvalue_INT As Integer
    
    resultvalue_INT = 0
    resultvalue_INT = (((grossArea_ARG * 100) + subCategory_ARG))

    category_INT = resultvalue_INT

    AddCategory = resultvalue_INT

End Function


'Public Function addControl(control_ARG As edv_Control_wrapper) As Integer
'Dim resultvalue_INT As Integer
'
'On Error GoTo Errorhandler
'
'    resultvalue_INT = 0
'
'    control_OBJ = control_ARG
'
'    resultvalue_INT = 1
'
'Errorhandler:
'
'    addControl = resultvalue_INT
'
'
'End Function


Public Function ControlItem_TXT(index_ARG As Integer) As String
On Error GoTo Errorhandler
    Dim resultvalue_TXT(10) As String

    
    Rem title
    resultvalue_TXT(0) = ""
    Rem date
    resultvalue_TXT(1) = ""
    Rem target file type

    resultvalue_TXT(2) = "txt"
    Rem raw attachment type
    Debug.Print "//TODO: edv_Workslot_handler.ControlItem_TXT resultvalue_TXT(3) = control_OBJ.fileType_TXT(Me) "
    resultvalue_TXT(3) = control_OBJ.fileType_TXT(Me)
    resultvalue_TXT(3) = "xls"
    Rem raw attachment searchstring
    resultvalue_TXT(4) = "sbg_dsltest"

    Rem raw attachment seachstring postition
    resultvalue_TXT(5) = "1"

    'control_OBJ.fileType_TXT (This)

    ControlItem_TXT = resultvalue_TXT(index_ARG)
    
Errorhandler:

If Err.Number <> 0 Or debug_FLAG Then
    
    Debug.Print "edv_Workslot_handler.ControlItem_TXT.Err.Number=" & Err.Number
    Debug.Print "edv_Workslot_handler.ControlItem_TXT.Err.Description=" & Err.Description
    Debug.Print "edv_Workslot_handler.ControlItem_TXT.debug_FLAG=" & debug_FLAG
    Debug.Print "edv_Workslot_handler.ControlItem_TXT.resultvalue_TXT(2) = ""txt"""
    Debug.Print "edv_Workslot_handler.ControlItem_TXT.resultvalue_TXT(3)=" & resultvalue_TXT(3)
    Debug.Print "edv_Workslot_handler.ControlItem_TXT.resultvalue_TXT(4) = ""sbg_dsltest"""
    Debug.Print "edv_Workslot_handler.ControlItem_TXT.resultvalue_TXT(5)=" & resultvalue_TXT(5)
        
End If

End Function

Public Function domain_TXT(index_ARG As Integer) As String
'sollte idealerweise eine Instanz der Klasse edv_Domain_handler sein
    domain_TXT = domainMatrix_TXT(index_ARG)
    
End Function

Public Function Discriminator_TXT(index_ARG As Integer) As String
'sollte idealerweise eine Instanz der Klasse edv_Discriminator_handler sein
    Discriminator_TXT = discriminatorMatrix_TXT(index_ARG)
    
End Function

Public Function PresetFile_CNT() As Integer
    
    PresetFile_CNT = PresetDiscriminator_INT

End Function

Function Workslotindex(Optional weekday_ARG As VbDayOfWeek = 0, Optional timeslot_ARG As Integer = 0) As Integer
On Error GoTo Errorhandler
Dim resultvalue_INT As Integer

'Const debug_FLAG As Boolean = True
Const debug_FLAG As Boolean = False
    
    resultvalue_INT = 0
    
    If weekday_ARG <> 0 Then
        Weekday_INT = weekday_ARG - 1
        Timeslot_INT = timeslot_ARG
    End If
    resultvalue_INT = (100 * Weekday_INT) + Timeslot_INT
    
    
Errorhandler:

If (Err.Number <> 0) Or debug_FLAG Then
    
    Debug.Print "edv_Workslot_handler.Workslotindex.IsCalledFrom=?"
    Debug.Print "edv_Workslot_handler.Workslotindex.Weekday_INT=" & Weekday_INT
    Debug.Print "edv_Workslot_handler.Workslotindex.Timeslot_INT=" & Timeslot_INT
    
End If

Workslotindex = resultvalue_INT

End Function


Public Function IsCategoryPrefix(name_ARG As String) As Boolean

Dim Categoryprefix As String
On Error GoTo Errorhandler
Rem category_INT ist ein Integer (na sowas9

    Categoryprefix = "080201"

    IsCategoryPrefix = (Left(name_ARG, 6) = Categoryprefix)
    
Errorhandler:

    If debug_FLAG Or Err.Number <> 0 Then

        Debug.Print "edv_Workslot_handler.IsCategoryPrefix.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_Workslot_handler.category_INT=" & category_INT
 
    End If

End Function

Function IsWorkslotprefix(name_ARG As String) As Boolean
On Error GoTo Errorhandler

Dim Workslotprefix As String
Rem Initialisierungen
    IsWorkslotprefix = False
    Workslotprefix = Weekday_INT
    
    
    If Timeslot_INT < 10 Then
        Workslotprefix = Workslotprefix & "0"
    End If
    
    Workslotprefix = Workslotprefix & Timeslot_INT
Rem finale Wertzuweisung
    IsWorkslotprefix = (Left(name_ARG, 3) = Workslotprefix)

Errorhandler:

    If debug_FLAG Or Err.Number <> 0 Then

        Debug.Print "edv_Workslot_handler.IsWorkslotprefix.debug_FLAG=" & debug_FLAG
        Debug.Print "edv_Workslot_handler.IsWorkslotprefix.Err.Number=" & Err.Number
        Debug.Print "edv_Workslot_handler.IsWorkslotprefix.IsCalledFrom=?"
        Debug.Print "edv_Workslot_handler.IsWorkslotprefix.name_ARG=" & name_ARG
        Debug.Print "edv_Workslot_handler.IsWorkslotprefix.Weekday_INT=" & Weekday_INT
        Debug.Print "edv_Workslot_handler.IsWorkslotprefix.Workslotprefix=" & Workslotprefix
        Debug.Print "edv_Workslot_handler.IsWorkslotprefix.Timeslot_INT=" & Timeslot_INT
        Debug.Print "edv_Workslot_handler.IsWorkslotprefix.IsWorkslotprefix=" & (Left(name_ARG, 3) = Workslotprefix)
    
    End If

End Function

'Private Function workslotControl() As edv_Control_wrapper
'
'    Dim resultvalue_OBJ As New edv_Control_wrapper
'
'End Function


Private Sub Class_Initialize()

    Weekday_INT = 0
    Timeslot_INT = 0
    category_INT = 0
    domains_CNT = 0
    Discriminator_CNT = 0
    PresetDiscriminator_INT = 0
    
    Debug.Print "edv_Workslot_handler.debug_FLAG=" & debug_FLAG
    If debug_FLAG Then
        Debug.Print "edv_Workslot_handler.Class_Initialize.Weekday_INT=" & Weekday_INT
        Debug.Print "edv_Workslot_handler.Class_Initialize.Timeslot_INT=" & Timeslot_INT
        Debug.Print "edv_Workslot_handler.Class_Initialize.category_INT=" & category_INT
    End If
  
End Sub

Private Sub Class_Terminate()


    Debug.Print "edv_Workslot_handler.Class_Terminate"
    Debug.Print "edv_Workslot_handler.Class_Terminate.debug_FLAG=" & debug_FLAG

End Sub
