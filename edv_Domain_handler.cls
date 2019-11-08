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

Public Function AddDiscriminator(code_ARG As String) As Boolean

    Call discriminator_OBJ.ShortCode(code_ARG)
    AddDiscriminator = True

End Function
