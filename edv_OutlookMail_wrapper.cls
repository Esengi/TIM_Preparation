VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_OutlookMail_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem wird verwendet
Private user_NAME As String
Private mySpace As Outlook.NameSpace
Rem diese Variable sollte im initialize gesetzt werden
Private myRootfolder As MAPIFolder
Private myWorkfolder As MAPIFolder
Private objMail As MailItem
Rem --------------Konstanten----------------------------------------
'Private Const debug_FLAG As Boolean = True
Private Const debug_FLAG As Boolean = False
Rem öffentliche Variablen
Public moduleversion_INT As Integer



Public Sub debug_me()

    Call sendMessage("testmessage", "norman.gruener@telefonica.com")


End Sub


Sub sendMessage(successmessage_ARG As String, recepientName_ARG As String)
On Error GoTo Errorhandler
    Rem Send a mail using the DoCmd.SendObject method
    
    Rem i need a Object for handling the new mail:
    Dim newmail_OBJ As Object
    
    Rem i declare the Object to be a Mail Item:
    Set newmail_OBJ = Application.CreateItem(olMailItem)
    

    Rem i set the important properties:
    newmail_OBJ.To = recepientName_ARG
    newmail_OBJ.Subject = successmessage_ARG
    'Set newmail_OBJ.BodyFormat = oIFormtPlain
    newmail_OBJ.Body = successmessage_ARG = successmessage_ARG & Chr(10)
    
    
    Rem i display it (optional, not necessary):
    'Call newmail_OBJ.Display
    
    Rem i send it:
    Call newmail_OBJ.Send
Errorhandler:
    
End Sub


Private Sub Class_Initialize()

On Error GoTo Errorhandler

    Rem setzen der public variablen
    moduleversion_INT = 4
    newMailInFolder_INT = -1
    
    Rem setzen der Steuerungsvariablen
    
    Rem setzen objektspezifischer variablen
    user_NAME = "norman.gruener@telefonica.com"
    Set mySpace = Application.GetNamespace("MAPI")
Errorhandler:

End Sub

Private Sub Class_Terminate()

End Sub
