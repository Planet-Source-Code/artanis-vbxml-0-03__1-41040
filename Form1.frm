VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "vbXML 0.03 Compatible Example"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "RGB Color Codes"
      Height          =   1815
      Left            =   7200
      TabIndex        =   7
      Top             =   3600
      Width           =   3015
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Color Codes"
         Default         =   -1  'True
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtList 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "ListView Code: (Fore - Back)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "TextBox Code: (Fore - Back)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parent Nodes Under <main>"
      Height          =   3375
      Left            =   7200
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin MSComctlLib.ListView lvColors 
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2143
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2143
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Nodes under <main>"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Nodes under <colors>"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Full XML Code"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtXML 
         Appearance      =   0  'Flat
         Height          =   4935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
' Write TextBox values
XML.WriteNode "/main/colors/tbox/fore", Left(txtText, InStr(txtText, " - ") - 1)
XML.WriteNode "/main/colors/tbox/back", Mid(txtText, InStr(txtText, " - ") + 3)

' Write ListView values
XML.WriteNode "/main/colors/lview/fore", Left(txtList, InStr(txtList, " - ") - 1)
XML.WriteNode "/main/colors/lview/back", Mid(txtList, InStr(txtList, " - ") + 3)

' Reload the XML File after saving
XML.Save App.Path & "\vbxml.xml"
InitXML
End Sub

Private Sub Form_Load()
'---------------------------------------------------
' vbXML uses the MSXML feature XPath.  Here is a quick
' example of how I would query a node (used in
' ReadNode, ReadNodeXML, and WriteNode, and many others):
' To access a node:
'   "/parent1/childofparent1/childofparent2"
' Here is a quick explination:
' I am using test.xml for this example.  For ease of
' this example I have pasted the contents here:
'
'  <test>
'     <text>
'         <hello>Hello</hello>
'         <bye>Goodbye</bye>
'     </text>
'  </test>
'
' To access the <hello> node, you would use this
' query:
'   "/test/text/hello"
'
' To access the <bye> node, you would use this query:
'   "/test/text/bye"
'
' In the last example (where we queried the <bye> node)
' "/test/" is parent1, "/text/" is childofparent1,
' and "/bye" is childofparent2
'
' Take note: you can have multiple child nodes
' (I dont know the exact count)
'---------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' This example is an XML editor.  It is meant to run
' with the vbXML class wrapper for Visual Basic
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

' See InitXML (In the basMain module) for the XML loading
InitXML
End Sub

Private Sub lvNodes_BeforeLabelEdit(Cancel As Integer)

End Sub
