VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   ControlBox      =   0   'False
   Icon            =   "frmSource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8460
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame frame1 
      Caption         =   "Source"
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   8175
      Begin VB.ComboBox cmb_Type 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   3135
      End
      Begin VB.Frame frame_AlbumURI 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   735
         Left            =   3600
         TabIndex        =   16
         Top             =   240
         Width           =   4335
         Begin VB.CommandButton b_Browse 
            Caption         =   "Browse..."
            Height          =   375
            Left            =   3240
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox f_AlbumURI 
            Height          =   315
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   12
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Album File/URL"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Frame frame_Album 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4095
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Width           =   4455
         Begin VB.TextBox f_AlbumName 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   3135
         End
         Begin VB.CommandButton b_Import 
            Caption         =   "Import..."
            Height          =   375
            Left            =   2400
            TabIndex        =   9
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton b_Export 
            Caption         =   "Export..."
            Height          =   375
            Left            =   720
            TabIndex        =   8
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton b_Delete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   3240
            TabIndex        =   5
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton b_MoveDown 
            Caption         =   "Move Down"
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CommandButton b_MoveUp 
            Caption         =   "Move Up"
            Height          =   375
            Left            =   3240
            TabIndex        =   7
            Top             =   3000
            Width           =   1095
         End
         Begin VB.CommandButton b_Edit 
            Caption         =   "Edit..."
            Height          =   375
            Left            =   3240
            TabIndex        =   4
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton b_New 
            Caption         =   "New..."
            Height          =   375
            Left            =   3240
            TabIndex        =   3
            Top             =   1080
            Width           =   1095
         End
         Begin VB.ListBox l_List 
            Height          =   2400
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   2
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "Album Name"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "List of Image Files/URLs"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   2775
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   3480
         X2              =   3480
         Y1              =   240
         Y2              =   4320
      End
   End
End
Attribute VB_Name = "frmSource"
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
    Dim j As Integer
        
    cmb_Type.ListIndex = CInt(gXMLDomNodeClone.childNodes(XMLISource_Type).Text)
    
    f_AlbumURI.Text = Replace(gXMLDomNodeClone.childNodes(XMLISource_AlbumURI).Text, "%sd%", g_strExeDir)
    f_AlbumName.Text = gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_Name).Text
        
    Rem Remove all the current items
    For j = 0 To l_List.ListCount - 1
        l_List.RemoveItem (0)
    Next
    
    Rem Add the new items
    For j = 0 To gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes.length - 1
        Call l_List.AddItem(gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes.Item(j).childNodes.Item(XMLIPicture_Name).Text & " : " & Replace(gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes.Item(j).childNodes.Item(XMLIPicture_URI).Text, "%sd%", g_strExeDir), j)
    Next
    
    b_Edit.Enabled = False
    b_Delete.Enabled = False
    b_MoveDown.Enabled = False
    b_MoveUp.Enabled = False
    
End Function

Private Function UIToXML() As Boolean
    Dim i As Integer
    If cmb_Type.ListIndex = XMLISource_Type_Album Then
        Rem Set the type as picturelist
        gXMLDomNodeClone.childNodes(XMLISource_Type).Text = XMLISource_Type_Album
        gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_Name).Text = f_AlbumName.Text
        
        Rem Clear the data from the unset type
        gXMLDomNodeClone.childNodes(XMLISource_AlbumURI).Text = ""
    ElseIf cmb_Type.ListIndex = XMLISource_Type_AlbumURI Then
        Rem Set the type as albumuri, and get the uri
        gXMLDomNodeClone.childNodes(XMLISource_Type).Text = XMLISource_Type_AlbumURI
        gXMLDomNodeClone.childNodes(XMLISource_AlbumURI).Text = f_AlbumURI.Text
        
        Rem Clear the data from the unset type
        gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_Name).Text = ""
        
        For i = 0 To gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes.length - 1
            Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).removeChild(gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).firstChild)
        Next
    End If
    
    UIToXML = True
End Function


Rem -----------------------------------------------------
Rem Functions specific to this form
Rem
Private Sub f_AlbumName_GotFocus()
    f_AlbumName.SelStart = 0
    f_AlbumName.SelLength = Len(f_AlbumName.Text)
End Sub

Private Sub f_AlbumURI_GotFocus()
    f_AlbumURI.SelStart = 0
    f_AlbumURI.SelLength = Len(f_AlbumURI.Text)
End Sub

Private Sub Form_Load()
    Call cmb_Type.AddItem("List of Image Files/URLs", XMLISource_Type_Album)
    Call cmb_Type.AddItem("Album File/URL", XMLISource_Type_AlbumURI)
End Sub

Private Sub cmb_Type_Click()
    frame_Album.Visible = (cmb_Type.ListIndex = XMLISource_Type_Album)
    frame_AlbumURI.Visible = (cmb_Type.ListIndex = XMLISource_Type_AlbumURI)
    gXMLDomNodeClone.childNodes(XMLISource_Type).Text = cmb_Type.ListIndex
End Sub

Private Sub b_Browse_Click()
    Dim fso As Scripting.FileSystemObject
    Dim strFileName As String
    
On Error GoTo error
    Set fso = New Scripting.FileSystemObject
        
    strFileName = f_AlbumURI.Text
    If fso.FileExists(strFileName) Then
        CommonDialog1.FileName = strFileName
        CommonDialog1.InitDir = strFileName
    Else
        CommonDialog1.InitDir = CurDir
    End If

    CommonDialog1.Filter = "Album Files (*.gpa)|*.gpa"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        f_AlbumURI.Text = CommonDialog1.FileName
    End If
    
    Exit Sub
error:
    MsgBox "Error accessing file", vbExclamation, "Error"
    
End Sub

Private Sub b_MoveUp_Click()
    Dim i As Long
    Dim tempXMLDomNode
        
    i = l_List.ListIndex
    
    If i > 0 Then
        
        Set tempXMLDomNode = gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes(i)
        Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).removeChild(tempXMLDomNode)
        Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).insertBefore(tempXMLDomNode, gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes(i - 1))
        XMLToUI
        l_List.ListIndex = i - 1
    End If
End Sub

Private Sub b_MoveDown_Click()
    Dim i As Long
    Dim tempXMLDomNode
        
    i = l_List.ListIndex
    
    If i < l_List.ListCount - 1 Then
        
        Set tempXMLDomNode = gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes(i)
        Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).removeChild(tempXMLDomNode)
        Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).insertBefore(tempXMLDomNode, gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes(i + 1))
        XMLToUI
        l_List.ListIndex = i + 1
    End If
End Sub

Private Sub b_New_Click()
    Dim objXMLDom As MSXML.DOMDocument
    Set objXMLDom = New MSXML.DOMDocument
    objXMLDom.loadXML ("<Picture><Name>New Image</Name><URI></URI></Picture>")
    
    Set frmPicture.XMLDomNode = objXMLDom.childNodes.Item(0)
    frmPicture.Show (vbModal)
    If frmPicture.Result = True Then
        Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).appendChild(frmPicture.XMLDomNode)
        Call XMLToUI
        l_List.ListIndex = l_List.ListCount - 1
    End If
    Unload frmPicture
End Sub

Private Sub b_Delete_Click()
    Dim i As Integer
    i = l_List.ListIndex
    Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).removeChild(gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes.Item(i))
    XMLToUI
    If l_List.ListCount > 0 Then
        If i > l_List.ListCount - 1 Then
            l_List.ListIndex = l_List.ListCount - 1
        Else
            l_List.ListIndex = i
        End If
    End If
End Sub

Private Sub b_Edit_Click()
    Dim i As Integer
    i = l_List.ListIndex
    Set frmPicture.XMLDomNode = gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes(i)
    frmPicture.Show (vbModal)
    If frmPicture.Result = True Then
        Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).replaceChild(frmPicture.XMLDomNode, gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes(l_List.ListIndex))
        Call XMLToUI
        l_List.ListIndex = i
    End If
    Unload frmPicture
End Sub


Private Sub l_List_DblClick()
    Call b_Edit_Click
End Sub

Private Sub l_List_Click()
    Rem Start by turning everything off
    b_Edit.Enabled = False
    b_Delete.Enabled = False
    b_MoveDown.Enabled = False
    b_MoveUp.Enabled = False
    
    Rem Now start to enable buttons
    If l_List.ListIndex <> -1 Then
        b_Edit.Enabled = True
        b_Delete.Enabled = True

        If l_List.ListIndex > 0 Then
            b_MoveUp.Enabled = True
        End If
        
        If l_List.ListIndex < l_List.ListCount - 1 Then
            b_MoveDown.Enabled = True
        End If
    End If
End Sub

Private Sub l_List_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        Dim vfn

        For Each vfn In Data.Files
            Dim objXMLDom As MSXML.DOMDocument
            Set objXMLDom = New MSXML.DOMDocument
            objXMLDom.loadXML ("<Picture><Name>" & HTMLEncode(vfn) & "</Name><URI>" & HTMLEncode(vfn) & "</URI></Picture>")
            Call gXMLDomNodeClone.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).appendChild(objXMLDom.childNodes(0))
        Next vfn
        
        XMLToUI
        l_List.ListIndex = l_List.ListCount - 1
    End If
End Sub

Private Sub l_List_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy And Effect
        Exit Sub
    End If
    Effect = vbDropEffectNone
End Sub


Private Sub b_Export_Click()
On Error GoTo error:
    CommonDialog1.InitDir = CurDir
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Graphic Pump Album Files (*.gpa)|*.gpa"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Call SaveDOMToFile(CommonDialog1.FileName, gXMLDomNodeClone.childNodes(XMLISource_Album).XML)
    End If
    Exit Sub
error:
    MsgBox "An error occured exporting the album file", vbExclamation, "Error"
End Sub


Private Sub b_Import_Click()
    Dim i As Long
    Dim objXMLAlbum As IXMLDOMNode
    Dim strError As String
    Dim fso As Scripting.FileSystemObject
    Dim objXMLDocument As MSXML.DOMDocument
    
    strError = ""
    
On Error GoTo error:
    Set fso = New Scripting.FileSystemObject
    
    Rem Show the import dialog
    CommonDialog1.InitDir = CurDir
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Graphic Pump Album Files (*.gpa)|*.gpa"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.ShowOpen
    
    Rem Perform the import
    If CommonDialog1.FileName <> "" Then
        
        If fso.FileExists(CommonDialog1.FileName) Then
        
            Set objXMLDocument = LoadDOMFromFile(CommonDialog1.FileName, "Album", "1")
            If objXMLDocument Is Nothing Then
                GoTo error
            End If
        
            Rem Get the album
            Set objXMLAlbum = objXMLDocument.childNodes.Item(0)
        
            Rem Replace the album
            Call gXMLDomNodeClone.replaceChild(objXMLAlbum, gXMLDomNodeClone.childNodes(XMLISource_Album))
        
            Rem Refesh the UI from the new XML
            XMLToUI
            l_List.ListIndex = l_List.ListCount - 1
        Else
            strError = "The .gpa file does not exist"
            GoTo error
        End If
    End If
    Exit Sub
error:
    MsgBox "An error occured importing the album file. " & strError, vbExclamation, "Error"
            
End Sub


