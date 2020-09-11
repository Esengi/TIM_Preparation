VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_Domain_handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private discriminator_OBJ As New edv_Discriminator_handler

Private domainShortCode_TXT As String

Private Discriminator_CNT As Integer

Public Function AddShortcode(code_ARG As String) As Boolean

    domainShortCode_TXT = code_ARG

    AddShortcode = True

End Function

Public Function AddDiscriminator(code_ARG As String) As Boolean

    Call discriminator_OBJ.ShortCode(code_ARG)
    AddDiscriminator = True
    
    
'todo: in eine neue Klasse domain_handler auslagern
'todo: wer verwendet das zur Zeit?
'TITANS_mdediator.Class_Initialize
        
    'discriminatorMatrix_TXT(Discriminator_CNT) = discriminator_ARG
    
    Discriminator_CNT = Discriminator_CNT + 1
    'PresetDiscriminator_INT = PresetDiscriminator_INT + preset_ARG
End Function


Public Function control_TXT(controlName_ARG As String) As String

    control_TXT = domainShortCode_TXT

End Function

Private Sub Class_Initialize()

    Discriminator_CNT = 0

End Sub
