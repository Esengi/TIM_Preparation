VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "edv_Discriminator_handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private stagingIndex_IDX As Integer
Private timeIndex_DAT As Date
Private shortCode_TXT As String

Public Function ShortCode(Optional code_ARG As String = "") As String

    ShortCode = shortCode_TXT

End Function


Function AddTimeIndex(time_ARG As Date) As Boolean

    AddTimeIndex = True

End Function
