VERSION 5.00
Begin VB.Form frmFormatProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format Profile"
   ClientHeight    =   1836
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3360
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1836
   ScaleWidth      =   3360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_Settings 
      Caption         =   "Settings..."
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox f_Name 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFormatProfile"
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
Dim gIndex As Integer

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
    f_Name.Text = gXMLDomNodeClone.childNodes(XMLIFormatProfile_Name).Text
    
End Function

Private Function UIToXML() As Boolean
    Dim i As Integer
    
    Rem Verify that the name isn't being used (e.g. unique)
    For i = 0 To g_objXMLDomNodeFormatProfiles.childNodes.length - 1
        If i <> gIndex Then
            If f_Name.Text = g_objXMLDomNodeFormatProfiles.childNodes(i).childNodes(XMLIFormatProfile_Name).Text Then
                MsgBox "Format name must be unique.", vbExclamation, "Error"
                UIToXML = False
                Exit Function
            End If
        End If
    Next
    
    gXMLDomNodeClone.childNodes(XMLIFormatProfile_Name).Text = f_Name.Text
    
    UIToXML = True
End Function

Rem ---------------------------------------------------
Rem Functionality specific to this form
Rem
Public Property Let FormatIndex(i As Integer)
    gIndex = i
End Property

Private Sub b_Settings_Click()
    Set frmFormatSettings.XMLDomNode = gXMLDomNodeClone.childNodes(XMLIFormatProfile_Settings)
    frmFormatSettings.Show (vbModal)
    If frmFormatSettings.Result = True Then
        Call gXMLDomNodeClone.replaceChild(frmFormatSettings.XMLDomNode, gXMLDomNodeClone.childNodes(XMLIFormatProfile_Settings))
    End If
    Unload frmFormatSettings
    
End Sub

Private Sub f_name_gotfocus()
    f_Name.SelStart = 0
    f_Name.SelLength = Len(f_Name.Text)
End Sub

