VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mserv_onenote_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Add the following references (adjust to our office version):
'
' - Microsoft OneNote 14.0 Object Library
' - Microsoft XML, v6.0
Private OneNote As OneNote.Application
Private doc As IXMLDOMDocument2
Private oneNotePagesXml As String


Sub debug_me()

    Call ListOneNotePages


End Sub


Sub ListOneNotePages()
    ' Original example is from http://code.msdn.microsoft.com/office/onenote-2010-retrieve-data-023e69c0
    ' License: Apache 2.0
    ' Modified to get all pages & content instead of the notebook list

    ' Connect to OneNote 2010.
    ' OneNote will be started if it's not running.
    
    
    Set OneNote = New OneNote.Application
    'Call OneNote.OpenHierarchy(
    'Call OneNote.GetHierarchyParent("", "")
    ' Get the XML that represents the OneNote pages
    ' oneNotePagesXml gets filled in with an XML document providing information
    ' about all OneNote pages.
    ' You want all the data. Thus you provide an empty string
    ' for the bstrStartNodeID parameter.
    Call OneNote.GetHierarchy("", hsPages, oneNotePagesXml, xs2013)

    ' Use the MSXML Library to parse the XML.
    
    Set doc = New MSXML2.DOMDocument60

    If doc.LoadXML(oneNotePagesXml) Then
        ' Find all the Page nodes in the one namespace.
        Dim nodes As MSXML2.IXMLDOMNodeList
        Set nodes = doc.DocumentElement.SelectNodes("//one:Page")

        Dim node As MSXML2.IXMLDOMNode
        Dim pageName As String
        Dim sectionName As String
        Dim pageContent As String
        Dim temp As String
        ' Walk the collection of Pages.
        ' Read attribute values and write them
        ' out to the Immediate window of your VBA host.
        For Each node In nodes
            pageName = node.Attributes.getNamedItem("name").Text
            Debug.Print "Page name: "; vbCrLf & " " & pageName

            Call OneNote.GetPageContent(GetAttributeValueFromNode(node, "ID"), pageContent, piBasic)
            Debug.Print " content: " & pageContent

        Next
    Else
        MsgBox "OneNote 2010 XML Data failed to load."
    End If

End Sub


Private Function GetAttributeValueFromNode(node As MSXML2.IXMLDOMNode, attributeName As String) As String
    If node.Attributes.getNamedItem(attributeName) Is Nothing Then
        GetAttributeValueFromNode = "Not found."
    Else
        GetAttributeValueFromNode = node.Attributes.getNamedItem(attributeName).Text
    End If
End Function
