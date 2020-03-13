VERSION 5.00
Begin VB.Form frmTask 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   Icon            =   "frmTask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5805
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_Source 
      Caption         =   "Source..."
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton b_Formatting 
      Caption         =   "Formatting..."
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton b_Schedule 
      Caption         =   "Schedule..."
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton b_Destination 
      Caption         =   "Destination..."
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   5535
      Begin VB.TextBox f_Result 
         Height          =   735
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label l_LastRun 
         Caption         =   "Label9"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label l_Result 
         Caption         =   "Result"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Last Run"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox f_Name 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   90
      Width           =   3495
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Task Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmTask"
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
    Dim objXMLDomNodeStatus As IXMLDOMNode
    Dim dtLastRun As Date
    
    f_Name.Text = gXMLDomNodeClone.childNodes(XMLITask_Name).Text
    
    Set objXMLDomNodeStatus = gXMLDomNodeClone.childNodes(XMLITask_Status)
    
    If objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).Text = "" Then
        l_LastRun.Caption = "Task has never been run."
        l_Result.Visible = False
        f_Result.Visible = False
    Else
        dtLastRun = ISODateToDate(objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).Text)
        l_LastRun.Caption = CStr(dtLastRun)
        If CInt(objXMLDomNodeStatus.childNodes(XMLIStatus_Failed).Text) = 0 Then
            l_Result.Caption = "Succeded"
            f_Result.Visible = False
        Else
            l_Result.Caption = "Failed"
            f_Result.Text = objXMLDomNodeStatus.childNodes(XMLIStatus_Reason).Text
        End If
            
    End If
            
End Function

Private Function UIToXML() As Boolean
    gXMLDomNodeClone.childNodes(XMLITask_Name).Text = f_Name.Text
    
    UIToXML = True
End Function


Rem ---------------------------------------------------
Rem Functionality specific to this form
Rem

Private Sub b_Source_Click()
    Set frmSource.XMLDomNode = gXMLDomNodeClone.childNodes(XMLITask_Source)
    frmSource.Show (vbModal)
    If frmSource.Result = True Then
        Call gXMLDomNodeClone.replaceChild(frmSource.XMLDomNode, gXMLDomNodeClone.childNodes(XMLITask_Source))
    End If
    Unload frmSource
End Sub

Private Sub b_Schedule_Click()
    Set frmSchedule.XMLDomNode = gXMLDomNodeClone.childNodes(XMLITask_Schedule)
    frmSchedule.Show (vbModal)
    If frmSchedule.Result = True Then
        Call gXMLDomNodeClone.replaceChild(frmSchedule.XMLDomNode, gXMLDomNodeClone.childNodes(XMLITask_Schedule))
    End If
    Unload frmSchedule
End Sub

Private Sub b_Formatting_Click()
    Set frmFormat.XMLDomNode = gXMLDomNodeClone.childNodes(XMLITask_Format)
    frmFormat.Show (vbModal)
    If frmFormat.Result = True Then
        Call gXMLDomNodeClone.replaceChild(frmFormat.XMLDomNode, gXMLDomNodeClone.childNodes(XMLITask_Format))
    End If
    Unload frmFormat
End Sub


Private Sub b_Destination_Click()
    Set frmDestination.XMLDomNode = gXMLDomNodeClone.childNodes(XMLITask_Destination)
    frmDestination.Show (vbModal)
    If frmDestination.Result = True Then
        Call gXMLDomNodeClone.replaceChild(frmDestination.XMLDomNode, gXMLDomNodeClone.childNodes(XMLITask_Destination))
    End If
    Unload frmDestination
End Sub

Private Sub f_name_gotfocus()
    f_Name.SelStart = 0
    f_Name.SelLength = Len(f_Name.Text)
End Sub

