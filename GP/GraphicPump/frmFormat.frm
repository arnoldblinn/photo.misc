VERSION 5.00
Begin VB.Form frmFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton rb_Profile 
      Caption         =   "Profile"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton b_FormatSettings 
      Caption         =   "Format Settings..."
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.OptionButton rb_Custom 
      Caption         =   "Custom"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cmb_FormatProfiles 
      Height          =   315
      ItemData        =   "frmFormat.frx":0000
      Left            =   1680
      List            =   "frmFormat.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem --------------------------------------------
Rem Boilerplate functionality for every form
Rem
Dim gXMLDomNode As IXMLDOMNode
Dim gXMLDomNodeClone As IXMLDOMNode
Dim gResult As Boolean
Dim gForm As Form

Public Property Get Result() As Boolean
    Result = gResult
End Property

Public Property Set XMLDomNode(objXMLDomNode As IXMLDOMNode)
    Set gXMLDomNode = objXMLDomNode
    
    Set gXMLDomNodeClone = gXMLDomNode.cloneNode(True)
        
    Call XMLToUI
    Set gForm = Me
End Property

Public Property Get XMLDomNode() As IXMLDOMNode
    Set XMLDomNode = gXMLDomNode
End Property

Private Sub b_OK_Click()
    If UIToXML() = False Then
        Exit Sub
    End If
    
    gResult = True
    Set gXMLDomNode = gXMLDomNodeClone
    gForm.Hide
End Sub

Private Sub b_Cancel_Click()
    gResult = False
    gForm.Hide
End Sub

Rem -----------------------------------------------------
Rem Stub functions same in every form but specific to this form
Rem
Private Function XMLToUI()
    If gXMLDomNodeClone.childNodes(XMLIFormat_Name).Text = "" Then
        rb_Custom.value = True
        rb_Profile.value = False
        cmb_FormatProfiles.Enabled = False
        b_FormatSettings.Enabled = True
    Else
        rb_Custom.value = False
        rb_Profile.value = True
        cmb_FormatProfiles.Enabled = True
        b_FormatSettings.Enabled = False
        
        cmb_FormatProfiles.Text = gXMLDomNodeClone.childNodes(XMLIFormat_Name).Text
    End If
    
End Function

Private Function UIToXML() As Boolean
    Dim i As Integer
    
    If rb_Custom.value = True Then
        gXMLDomNodeClone.childNodes(XMLIFormat_Name).Text = ""
    Else
        If cmb_FormatProfiles.Text = "" Then
            MsgBox "Format must be selected.", vbExclamation, "Error"
            UIToXML = False
            Exit Function
        End If
        gXMLDomNodeClone.childNodes(XMLIFormat_Name).Text = cmb_FormatProfiles.Text
        
        For i = 0 To g_objXMLDomNodeFormatProfiles.childNodes.length - 1
            Dim strText As String
            strText = g_objXMLDomNodeFormatProfiles.childNodes.Item(i).childNodes(XMLIFormatProfile_Name).Text
            If strText = cmb_FormatProfiles.Text Then
                Call gXMLDomNodeClone.replaceChild(g_objXMLDomNodeFormatProfiles.childNodes(i).childNodes(XMLIFormatProfile_Settings).cloneNode(True), gXMLDomNodeClone.childNodes(XMLIFormat_Settings))
            End If
        Next
    End If
    
    UIToXML = True
End Function


Rem ------------------------------------
Rem Code specific to this form
Rem
Private Sub refreshui()
    Dim i As Integer
    
    Rem Remove all the current items
    For i = 0 To cmb_FormatProfiles.ListCount - 1
        cmb_FormatProfiles.RemoveItem (0)
    Next
    
    Rem Initialize the form from the xml state
    For i = 0 To g_objXMLDomNodeFormatProfiles.childNodes.length - 1
        Dim strText As String
        strText = g_objXMLDomNodeFormatProfiles.childNodes.Item(i).childNodes(XMLIFormatProfile_Name).Text
        cmb_FormatProfiles.AddItem (strText)
    Next
    
    If g_objXMLDomNodeFormatProfiles.childNodes.length = 0 Then
        Call rb_Custom_Click
        rb_Profile.Enabled = False
    Else
        rb_Profile.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call refreshui
End Sub

Private Sub rb_Custom_Click()
    rb_Custom.value = True
    rb_Profile.value = False
    cmb_FormatProfiles.Enabled = False
    b_FormatSettings.Enabled = True
End Sub

Private Sub rb_Profile_Click()
    rb_Custom.value = False
    rb_Profile.value = True
    cmb_FormatProfiles.Enabled = True
    b_FormatSettings.Enabled = False
End Sub

Private Sub b_FormatSettings_Click()
    Set frmFormatSettings.XMLDomNode = gXMLDomNodeClone.childNodes(XMLIFormat_Settings)
    frmFormatSettings.Show (vbModal)
    If frmFormatSettings.Result = True Then
        Call gXMLDomNodeClone.replaceChild(frmFormatSettings.XMLDomNode, gXMLDomNodeClone.childNodes(XMLIFormat_Settings))
    End If
    Unload frmFormatSettings
End Sub

