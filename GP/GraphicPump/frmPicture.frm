VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPicture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture"
   ClientHeight    =   1692
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5892
   ControlBox      =   0   'False
   Icon            =   "frmPicture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1692
   ScaleWidth      =   5892
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox f_File 
      Height          =   315
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   660
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton b_Browse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox f_Name 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   90
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "File/URL"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const L_SUPPORT_GIFLZW = 1
Const k_UnlockKey_GIF_V12 = "sg8Z2XkjL"

Rem ----------------------------------------
Rem Boilerplate for every form
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

Rem --------------------------------
Rem Functionality specific to this form
Rem
Rem
Private Sub XMLToUI()
    f_Name.Text = gXMLDomNodeClone.childNodes(XMLIPicture_Name).Text
    f_File.Text = Replace(gXMLDomNodeClone.childNodes(XMLIPicture_URI).Text, "%sd%", g_strExeDir)
End Sub

Private Function UIToXML() As Boolean
    If f_Name.Text = "" Then
        MsgBox "Name is required.", vbExclamation, "Error"
        UIToXML = False
        Exit Function
    End If
    
    If f_File.Text = "" Then
        MsgBox "File/URL is required.", vbExclamation, "Error"
        UIToXML = False
        Exit Function
    End If
    gXMLDomNodeClone.childNodes(XMLIPicture_Name).Text = f_Name.Text
    gXMLDomNodeClone.childNodes(XMLIPicture_URI).Text = f_File.Text
    UIToXML = True
End Function

Private Sub b_Browse_Click()
    Dim fso As Scripting.FileSystemObject
    Dim strFileName As String
    
On Error GoTo error
    
    Set fso = New Scripting.FileSystemObject
        
    strFileName = f_File.Text
    
    If fso.FileExists(strFileName) Then
        CommonDialog1.FileName = strFileName
        CommonDialog1.InitDir = strFileName
    Else
        CommonDialog1.InitDir = CurDir
    End If
    CommonDialog1.Filter = "JPEG files (*.jpg)|*.jpg|BMP Files (*.bmp)|*.bmp|Photoshop FIles (*.psd)|*.psd|Kodak Photo CD (*.pcd)|*.pcd|TIFF files (*.tif)|*.tif|GIF files (*.gif)|*.gif|All files|*.*"
    
    CommonDialog1.ShowSave
        
    f_File.Text = CommonDialog1.FileName
        
    Exit Sub
error:
    MsgBox "Error accessing file", vbExclamation, "Error"
End Sub

Private Sub f_File_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim vfn As Variant
    If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
        
        For Each vfn In Data.Files
            f_File.Text = vfn
        Next
    End If
End Sub

Private Sub f_File_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub f_name_gotfocus()
    f_Name.SelStart = 0
    f_Name.SelLength = Len(f_Name.Text)
End Sub

Private Sub f_File_GotFocus()
    f_File.SelStart = 0
    f_File.SelLength = Len(f_File.Text)
End Sub

