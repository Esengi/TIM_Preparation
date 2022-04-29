VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_WB_TSH_mediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem Verwendete Microservices
Private excelsheet_OBJ As New edv_LinkExcelSheet_mserv
Private createWorkbenchtimesheets_OBJ As New edv_SQL_wrapper

Private timesheetName_TXT As String
Private timefileNAME_TXT As String
Private calendarWeek_TXT As String

Rem den domänenspezifischen Pfad sollte dieser Mediator kennen und wissen. Er wird auf diesem Pfad arbeiten
Private Const domain_PATH As String = "C:\Users\gruenern\OneDrive - Telefonica\00 Budget Actual Planning\IT-Demand\Workbench_Process_4_Timesheets\"
Private Const defaultDatabase_PATH As String = "C:\Users\gruenern\OneDrive - Telefonica\00 Budget Actual Planning\IT-Demand"
Private Const subdomain_PATH As String = "\Workbench_Process_4_Timesheets\"
Private Const defaultDatabase_NAME As String = "REPORTING_Projects.accdb"
Private Const module_NAME As String = "edv_WB_TSH_mediator"
Private Const moduleVersion_DBL As Double = 0.8
Private Const debug_FLAG As Boolean = True




Sub debug_me()

    Rem 1) Bestimme den Namen und die Herkunftsdatei des neuen Timesheets
    
    calendarWeek_TXT = "CW16"
    timesheetName_TXT = calendarWeek_TXT & "$"
    timefileNAME_TXT = domain_PATH & "WIPRO_Workbench_BA_ASA_April_2022 (1)"
      
    If debug_FLAG Then
        Debug.Print module_NAME & ".calendarWeek_TXT=" & calendarWeek_TXT
        Debug.Print module_NAME & ".timesheetName_TXT=" & timesheetName_TXT
        Debug.Print module_NAME & ".timefileNAME_TXT=" & timefileNAME_TXT
    End If
    
    Rem 2) verlinke timesheet als Access Tabelle
    Call linkTimesheet
    Rem 2.5) definiere abfrage
    Call importTimesheet
    Rem 3) führe definierte Abfragen aus
    Rem 4) wiederhole bis keine neuen timesheets mehr vorhanden sind



'-----------------------------------------------------------------------------------------------------------

End Sub

Private Sub linkTimesheet()

    Call excelsheet_OBJ.LinkExcelfile(timesheetName_TXT, timefileNAME_TXT)

End Sub


Private Sub importTimesheet()

     Const execution_IDX As Integer = 3

    Rem 3) aktualisiere WorkbenchTimesheets

    Call createWorkbenchtimesheets_OBJ.AddInsertintostatement("WorkbenchTimesheets ( Week_TXT, Role, PO, Phase, Budgettype, SR, [SR Title], MD_ACTUAL, Budget, Hours, [Role Occupant] ) ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement(" """ & calendarWeek_TXT & "_2022"" AS Timesheet_Week, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F1, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F2, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F3, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F4, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F5, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F7, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F13, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F14, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F15, ")
    Call createWorkbenchtimesheets_OBJ.AddSelectstatement("[Workbench Timesheet].F16 ")
    Call createWorkbenchtimesheets_OBJ.AddFromClause("[Workbench Timesheet] ")
    Call createWorkbenchtimesheets_OBJ.AddWhereClause("((([Workbench Timesheet].F2) Is Not Null))")
    Call createWorkbenchtimesheets_OBJ.SubmitQueryStatement("create:Workbenchtimesheets", execution_IDX)

End Sub


'read WorkbenchTimesheet for Book:


'SELECT [mapping:TimesheetRoles].RoleInFlow_TXT,
'DateAdd("ww",(Mid([Week_TXT],3,2)*1),#12/31/2020#) AS StartDate_DAT,
'Sum(WorkbenchTimesheets.Hours) AS SummevonHours,
'"SR_" & Right([SR],6)*1 & "-" & IIf([Phase]="DEV",11,3) AS IssueKey,
'Year([StartDate_DAT]) & "-" & IIf(Month([StartDate_DAT])<10,"0","") & Month([StartDate_DAT]) & "-" & IIf(Day([StartDate_DAT])<10,"0","") & Day([StartDate_DAT]) AS StartDate,
'Year([StartDate_DAT]) & "-" & IIf(Month([StartDate_DAT])<10,"0","") & Month([StartDate_DAT]) & "-" & IIf(Day([StartDate_DAT])<10,"0","") & Day([StartDate_DAT]) AS EndDate,
'Fix([Hours]) & "." & (10^IIf((Len([Hours]-Fix([Hours]))-2)<1,0,Len([Hours]-Fix([Hours]))-2))*[Hours]-Fix([Hours]) AS TimeSpent,
'"" AS Comment
'From WorkbenchTimesheets
'LEFT JOIN [mapping:TimesheetRoles] ON WorkbenchTimesheets.Role = [mapping:TimesheetRoles].RoleInTimesheet_TXT
'WHERE ((([MD_ACTUAL] * 1) > 0) And ((WorkbenchTimesheets.SR) Is Not Null))
'GROUP BY [mapping:TimesheetRoles].RoleInFlow_TXT, DateAdd("ww",(Mid([Week_TXT],3,2)*1),#12/31/2020#), "SR_" & Right([SR],6)*1 & "-" & IIf([Phase]="DEV",11,3), Year([StartDate_DAT]) & "-" & IIf(Month([StartDate_DAT])<10,"0","") & Month([StartDate_DAT]) & "-" & IIf(Day([StartDate_DAT])<10,"0","") & Day([StartDate_DAT]), Fix([Hours]) & "." & (10^IIf((Len([Hours]-Fix([Hours]))-2)<1,0,Len([Hours]-Fix([Hours]))-2))*[Hours]-Fix([Hours]), Year([StartDate_DAT]) & "-" & IIf(Month([StartDate_DAT])<10,"0","") & Month([StartDate_DAT]) & "-" & IIf(Day([StartDate_DAT])<10,"0","") & Day([StartDate_DAT]);


Private Sub Class_Initialize()

    If debug_FLAG Then
        Debug.Print "debug_FLAG=" & debug_FLAG
        Debug.Print "module_NAME=" & module_NAME
        Debug.Print "moduleVersion_DBL=" & moduleVersion_DBL
        Debug.Print "domain_PATH=" & domain_PATH
        Debug.Print "subdomain_PATH=" & subdomain_PATH
        Debug.Print "defaultDatabase_PATH=" & defaultDatabase_PATH
        Debug.Print "defaultDatabase_NAME=" & defaultDatabase_NAME
    End If

End Sub

Private Sub Class_Terminate()

    If debug_FLAG Then
            Debug.Print module_NAME & ".Class_Terminate"
    End If

End Sub
