VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDestination 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destination"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   ControlBox      =   0   'False
   Icon            =   "frmDestination.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destination"
      Height          =   2772
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   8415
      Begin VB.ComboBox cmb_Type 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox f_FileTemplate 
         Height          =   315
         Left            =   5880
         TabIndex        =   7
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Frame frame_Directory 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   4575
         Begin VB.CheckBox chk_DirectoryDelete 
            Height          =   255
            Left            =   2160
            TabIndex        =   6
            Top             =   600
            Width           =   2415
         End
         Begin VB.CommandButton b_Browse 
            Caption         =   "Browse..."
            Height          =   375
            Left            =   3480
            TabIndex        =   5
            Top             =   60
            Width           =   975
         End
         Begin VB.TextBox f_Directory 
            Height          =   315
            Left            =   840
            TabIndex        =   4
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Delete Contents"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Directory"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.Frame frame_DigiFrame 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1095
         Left            =   3720
         TabIndex        =   13
         Top             =   240
         Width           =   4575
         Begin VB.CommandButton b_Detect 
            Caption         =   "Detect"
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   60
            Width           =   975
         End
         Begin VB.ComboBox cmb_Media 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   600
            Width           =   2655
         End
         Begin VB.ComboBox cmb_Port 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Label7 
            Caption         =   "Media"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Port"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Label l_DigiFrameNote 
         Caption         =   $"frmDestination.frx":0442
         Height          =   615
         Left            =   3720
         TabIndex        =   17
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "File Name (e.g. file%i%.jpg)"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   1380
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   3480
         X2              =   3480
         Y1              =   240
         Y2              =   2640
      End
   End
End
Attribute VB_Name = "frmDestination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub b_Detect_Click()
    Dim objDigiFrame As clsDigiFrame
    Dim fConnected As Boolean
    Dim iPort As Integer
    Dim iCard As Integer
    
On Error GoTo error
    Screen.MousePointer = 11
    Set objDigiFrame = New clsDigiFrame
    
    If objDigiFrame.AutoDetect(iPort, iCard) = False Then
        GoTo error
    End If
    
    cmb_Port.ListIndex = iPort
    cmb_Media.ListIndex = iCard
    
    Screen.MousePointer = 0
            
    Exit Sub
    
error:
    Screen.MousePointer = 0
    
    MsgBox "Unable to detect frame.", vbExclamation, "Error"

End Sub

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
Rem Stub functions filled in by this form
Rem
Private Sub XMLToUI()
    cmb_Type.ListIndex = CInt(gXMLDomNodeClone.childNodes(XMLIDestination_Type).Text)
    f_Directory.Text = Replace(gXMLDomNodeClone.childNodes(XMLIDestination_Directory).Text, "%sd%", g_strExeDir)
    chk_DirectoryDelete.value = CInt(gXMLDomNodeClone.childNodes(XMLIDestination_DirectoryDelete).Text)
    f_FileTemplate.Text = gXMLDomNodeClone.childNodes(XMLIDestination_FileTemplate).Text
                
    cmb_Port.ListIndex = CInt(gXMLDomNodeClone.childNodes(XMLIDestination_DigiFramePort).Text)
    cmb_Media.ListIndex = CInt(gXMLDomNodeClone.childNodes(XMLIDestination_DigiFrameMedia).Text)
    
End Sub

Private Function ValidateFile(strFileName As String, strExt As String, fIndex As Boolean, fFixedSize As Boolean) As Boolean
    ValidateFile = True
    
    If fIndex And InStr(1, LCase(strFileName), "%i%") = 0 Then
        ValidateFile = False
    ElseIf Right(LCase(strFileName), 4) <> strExt Then
        ValidateFile = False
    ElseIf Len(strFileName) > 11 Or InStr(strFileName, " ") <> 0 Then
        ValidateFile = False
    ElseIf fFixedSize = True And Len(strFileName) <> 11 Then
        ValidateFile = False
    End If
    
End Function

Private Function ValidateDirectory(strDirectoryName As String)
    Dim fso As Scripting.FileSystemObject
    Dim errNum
    
    Set fso = New Scripting.FileSystemObject
    
    If fso.FolderExists(Replace(strDirectoryName, "%sd%", g_strExeDir)) = False Then
        ValidateDirectory = False
    Else
        ValidateDirectory = True
    End If
End Function

Private Function UIToXML() As Boolean
    
    If cmb_Type.ListIndex = XMLIDestination_Type_Directory Then
        gXMLDomNodeClone.childNodes(XMLIDestination_Type).Text = XMLIDestination_Type_Directory
        
        If ValidateDirectory(f_Directory.Text) = False Then
            MsgBox "Invalid directory", vbExclamation, "Error"
            UIToXML = False
            Exit Function
        End If
        gXMLDomNodeClone.childNodes(XMLIDestination_Directory).Text = f_Directory.Text
        gXMLDomNodeClone.childNodes(XMLIDestination_DirectoryDelete).Text = chk_DirectoryDelete.value
        gXMLDomNodeClone.childNodes(XMLIDestination_FileTemplate).Text = f_FileTemplate.Text
        If ValidateFile(f_FileTemplate.Text, ".jpg", True, False) = False Then
            MsgBox "The file name must have a .jpg extension, must not contain spaces, and must contain %i% (for the file index).", vbExclamation, "Error"
            UIToXML = False
            Exit Function
        End If
    Else
        gXMLDomNodeClone.childNodes(XMLIDestination_Type).Text = XMLIDestination_Type_DigiFrame
        gXMLDomNodeClone.childNodes(XMLIDestination_DigiFramePort).Text = cmb_Port.ListIndex
        gXMLDomNodeClone.childNodes(XMLIDestination_DigiFrameMedia).Text = cmb_Media.ListIndex
        If ValidateFile(f_FileTemplate.Text, ".jpg", True, True) = False Then
            MsgBox "The file name must have a .jpg extension, must not contain spaces, and must contain four characters and %i% (for the file index). In other words, in the form 'file%i%.jpg'.", vbExclamation, "Error"
            UIToXML = False
            Exit Function
        End If
    End If
    
    UIToXML = True
End Function

Private Sub f_Directory_GotFocus()
    f_Directory.SelStart = 0
    f_Directory.SelLength = Len(f_Directory.Text)
End Sub

Private Sub f_FileTemplate_GotFocus()
    f_FileTemplate.SelStart = 0
    f_FileTemplate.SelLength = Len(f_FileTemplate.Text)
End Sub

Rem --------------------------------
Rem Functionality specific to this form
Rem
Private Sub Form_Load()
    Call cmb_Port.AddItem("COM Port 1", XMLIDestination_DigiFramePort_COM1)
    Call cmb_Port.AddItem("COM Port 2", XMLIDestination_DigiFramePort_COM2)
    Call cmb_Port.AddItem("COM Port 3", XMLIDestination_DigiFramePort_COM3)
    Call cmb_Port.AddItem("COM Port 4", XMLIDestination_DigiFramePort_COM4)
    
    Call cmb_Media.AddItem("Compact Flash", XMLIDestination_DigiFrameMedia_CompactFlash)
    Call cmb_Media.AddItem("SmartMedia", XMLIDestination_DigiFrameMedia_SmartMedia)
    
    Call cmb_Type.AddItem("Directory", XMLIDestination_Type_Directory)
    Call cmb_Type.AddItem("Digi-Frame", XMLIDestination_Type_DigiFrame)
End Sub

Private Sub b_Browse_Click()
    Dim strDirectoryName As String
    
    strDirectoryName = Replace(f_Directory.Text, "%sd%", g_strExeDir)
    frmDirectory.Path = strDirectoryName
    frmDirectory.Show (vbModal)
    If frmDirectory.Path <> "" Then
        f_Directory.Text = frmDirectory.Path
    End If
    Unload frmDirectory
End Sub

Private Sub cmb_Type_Click()
    frame_Directory.Visible = (cmb_Type.ListIndex = XMLIDestination_Type_Directory)
    frame_DigiFrame.Visible = (cmb_Type.ListIndex = XMLIDestination_Type_DigiFrame)
    l_DigiFrameNote.Visible = (cmb_Type.ListIndex = XMLIDestination_Type_DigiFrame)
End Sub
