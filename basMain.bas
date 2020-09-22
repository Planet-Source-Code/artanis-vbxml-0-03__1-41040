Attribute VB_Name = "basMain"
Option Explicit

Public XML As New vbXML

Public Sub InitXML()
frmMain.lvColors.ListItems.Clear
frmMain.lvMain.ListItems.Clear
frmMain.txtXML.Text = ""

' Open the XML file (vbxml.xml)
XML.OpenXML App.Path & "/vbxml.xml", oxFile

' Set the textbox's fore and back colors using the GetColorValue method
With frmMain.txtXML
    .ForeColor = XML.GetColorValue("/main/colors/tbox/fore")
    .BackColor = XML.GetColorValue("/main/colors/tbox/back")
    
    ' Load in the full XML file
    .Text = XML.XML
End With

' Set the fore and back colors of lvColors
With frmMain.lvColors
    .ForeColor = XML.GetColorValue("/main/colors/lview/fore")
    .BackColor = XML.GetColorValue("/main/colors/lview/back")
End With

' Set the fore and back colors of lvMain
With frmMain.lvMain
    .ForeColor = XML.GetColorValue("/main/colors/lview/fore")
    .BackColor = XML.GetColorValue("/main/colors/lview/back")
End With

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' Begin loading in the nodes (into the lv)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Dim i As Long

' Start with the <main> node
With frmMain.lvMain
For i = 0 To XML.NodeCount("/main") - 1
    .ListItems.Add , , XML.GetChildName("/main", i)
Next i
End With

' Now get children under the <colors> node
With frmMain.lvColors
For i = 0 To XML.NodeCount("/main/colors") - 1
    .ListItems.Add , , XML.GetChildName("/main/colors", i)
Next i
End With

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' Load in the color values for editing
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
With frmMain
    .txtText = XML.ReadNode("/main/colors/tbox/fore") & " - " & XML.ReadNode("/main/colors/tbox/back")
    .txtList = XML.ReadNode("/main/colors/lview/fore") & " - " & XML.ReadNode("/main/colors/lview/back")
End With
End Sub
