VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSave 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pump Image"
   ClientHeight    =   3744
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4992
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3744
   ScaleWidth      =   4992
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton b_BrowseSource 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox f_Source 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   180
      Width           =   2295
   End
   Begin VB.TextBox f_Filename 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CommandButton b_Formats 
      Caption         =   "Formats..."
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ComboBox cmb_Destination 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.ComboBox cmb_Format 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   660
      Width           =   2295
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "Pump"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame frame_DigiFrame 
      BorderStyle     =   0  'None
      Caption         =   "Digi-Frame"
      Height          =   1212
      Left            =   0
      TabIndex        =   18
      Top             =   1560
      Width           =   4935
      Begin VB.ComboBox cmb_Port 
         Height          =   288
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cmb_Media 
         Height          =   288
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton b_Detect 
         Caption         =   "Detect"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   153
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Port"
         Height          =   252
         Left            =   240
         TabIndex        =   20
         Top             =   276
         Width           =   1212
      End
      Begin VB.Label Label7 
         Caption         =   "Media"
         Height          =   252
         Left            =   240
         TabIndex        =   19
         Top             =   756
         Width           =   1332
      End
   End
   Begin VB.Frame frame_Directory 
      BorderStyle     =   0  'None
      Caption         =   "Directory"
      Height          =   1095
      Left            =   0
      TabIndex        =   16
      Top             =   1680
      Width           =   4935
      Begin VB.CheckBox chk_Overwrite 
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton b_Browse 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   60
         Width           =   975
      End
      Begin VB.TextBox f_Directory 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Overwrite"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Directory"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Label l_Source 
      Caption         =   "Label8"
      Height          =   255
      Left            =   1200
      TabIndex        =   23
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Source"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Filename"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2820
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Destination"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Format"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem Source file for the save operation
Dim m_strSource As String
Dim m_iIndex As Integer

Rem Property to set the source file for the save operation
Public Property Let SourceFile(strSource As String)
    m_strSource = strSource
End Property

Private Sub refresh_formats(strSelection As String)
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    
    For i = 0 To cmb_Format.ListCount - 1
        cmb_Format.RemoveItem (0)
    Next
    
    For i = 0 To g_objXMLDomNodeFormatProfiles.childNodes.length - 1
        Call cmb_Format.AddItem(g_objXMLDomNodeFormatProfiles.childNodes(i).childNodes(XMLIFormatProfile_Name).Text)
        
        If strSelection = g_objXMLDomNodeFormatProfiles.childNodes(i).childNodes(XMLIFormatProfile_Name).Text Then
            j = i
        End If
    Next
    
    cmb_Format.ListIndex = j
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

Private Sub b_BrowseSource_Click()
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    If fso.FileExists(f_Source.Text) Then
        CommonDialog1.FileName = f_Source.Text
        CommonDialog1.InitDir = f_Source.Text
    End If
    CommonDialog1.Filter = "JPEG files (*.jpg)|*.jpg|BMP Files (*.bmp)|*.bmp|Photoshop FIles (*.psd)|*.psd|Kodak Photo CD (*.pcd)|*.pcd|TIFF files (*.tif)|*.tif|GIF files (*.gif)|*.gif|All files|*.*"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        f_Source.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub b_Cancel_Click()
    Unload Me
End Sub

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

Private Sub b_Formats_Click()
    frmFormatProfiles.Show (vbModal)
    Unload frmFormatProfiles
    Call refresh_formats(cmb_Format.Text)
End Sub

Private Sub b_OK_Click()
    Dim strStatus As String
    Dim objXMLDomNodeFormatProfileSettings As IXMLDOMNode
    
    Rem Validate the format
    If cmb_Format.ListIndex = -1 Then
        strStatus = "You must select a format profile for the image."
        GoTo error
    End If
    Set objXMLDomNodeFormatProfileSettings = g_objXMLDomNodeFormatProfiles.childNodes(cmb_Format.ListIndex).childNodes(XMLIFormatProfile_Settings)
    
    Rem Validate that we have a file name
    If f_Filename.Text = "" Then
        strStatus = "You must select a name for the file"
        GoTo error
    ElseIf Right(f_Filename.Text, 4) <> ".jpg" Then
        strStatus = "You must specify a .jpg file"
        GoTo error
    ElseIf Len(f_Filename.Text) > 12 Or Len(f_Filename.Text) < 5 Then
        strStatus = "You must specify a valid .jpg file name"
        GoTo error
    End If
    
    Dim strFile As String
    strFile = f_Filename.Text
    
    Rem Validate that we have settings for the directory
    If cmb_Destination.ListIndex = XMLIDestination_Type_Directory Then
        If f_Directory.Text = "" Then
            strStatus = "You must enter a directory for the file."
            GoTo error
        End If
        
        If chk_Overwrite.value = 0 Then
            Dim fso As Scripting.FileSystemObject
            Dim strFullFileName As String
            
            Set fso = New Scripting.FileSystemObject
            
            strFullFileName = DirectoryFileToFullPath(f_Directory.Text, f_Filename.Text)
            
            If fso.FileExists(strFullFileName) Then
                strStatus = "File '" & strFullFileName & "' already exists"
                GoTo error
            End If
        End If
    Else
        If Len(strFile) < 12 Then
            strStatus = "File name for Digi-Frame must be of form file0000.jpg (8.3)"
            GoTo error
        End If
    End If
    
    Rem Validate that we have a file to push
    Dim strSource As String
    strSource = m_strSource
    If m_strSource = "" Then
        strSource = f_Source.Text
    End If
    If strSource = "" Then
        strStatus = "You must enter a file to process."
        GoTo error
    End If
    
    strStatus = RunFile(strSource, frmMain.frame_Status, frmMain.l_Status, frmMain.p_Progress, objXMLDomNodeFormatProfileSettings, cmb_Destination.ListIndex, cmb_Port.ListIndex, cmb_Media.ListIndex, f_Directory.Text, (chk_Overwrite.value = 1), strFile)
    
error:
    If strStatus <> "" Then
        MsgBox strStatus, vbExclamation, "Error"
    Else
        Rem If we ran OK, save the settings from this execution for the next time we are run
        Call RegWrite("Port ", CStr(cmb_Port.ListIndex))
        Call RegWrite("Media", CStr(cmb_Media.ListIndex))
        Call RegWrite("Overwrite", CStr(chk_Overwrite.value))
        Call RegWrite("Destination", CStr(cmb_Destination.ListIndex))
        Call RegWrite("Format", CStr(cmb_Format.Text))
        Call RegWrite("Directory", f_Directory.Text)
        m_iIndex = m_iIndex + 1
        If m_iIndex > 9999 Then
            m_iIndex = 0
        End If
        Call RegWrite("FileCounter", CStr(m_iIndex))
        
        Unload Me
    End If
        
End Sub



Private Sub Form_Load()
    If m_strSource = "" Then
        l_Source.Visible = False
    Else
        f_Source.Visible = False
        b_BrowseSource.Visible = False
        l_Source.Caption = FullPathToFile(m_strSource)
    End If
          
    Rem Initialize the formats
    Call refresh_formats(RegReadStringDefault("Format", ""))

    Call cmb_Port.AddItem("COM Port 1", XMLIDestination_DigiFramePort_COM1)
    Call cmb_Port.AddItem("COM Port 2", XMLIDestination_DigiFramePort_COM2)
    Call cmb_Port.AddItem("COM Port 3", XMLIDestination_DigiFramePort_COM3)
    Call cmb_Port.AddItem("COM Port 4", XMLIDestination_DigiFramePort_COM4)
    cmb_Port.ListIndex = 0
    
    Call cmb_Media.AddItem("Compact Flash", XMLIDestination_DigiFrameMedia_CompactFlash)
    Call cmb_Media.AddItem("SmartMedia", XMLIDestination_DigiFrameMedia_SmartMedia)
    cmb_Media.ListIndex = 0
    
    Call cmb_Destination.AddItem("Directory", XMLIDestination_Type_Directory)
    Call cmb_Destination.AddItem("Digi-Frame", XMLIDestination_Type_DigiFrame)
    cmb_Destination.ListIndex = 0
    frame_DigiFrame.Visible = False
    frame_Directory.Visible = True
    
    chk_Overwrite.value = RegReadIntDefault("Overwrite", 1)
    cmb_Port.ListIndex = RegReadIntDefault("Port", 0)
    cmb_Media.ListIndex = RegReadIntDefault("Media", 0)
    cmb_Destination.ListIndex = RegReadIntDefault("Destination", 0)
    f_Directory.Text = RegReadStringDefault("Directory", CurDir)
    m_iIndex = RegReadIntDefault("FileCounter", 0)
    
    f_Filename.Text = "file" & PadNumber(i) & ".jpg"
    
    If m_strSource <> "" Then
        f_Filename.TabIndex = 0
        b_OK.TabIndex = 1
        b_Cancel.TabIndex = 2
    End If
    
End Sub

Private Sub cmb_Destination_Click()
    frame_Directory.Visible = (cmb_Destination.ListIndex = XMLIDestination_Type_Directory)
    frame_DigiFrame.Visible = (cmb_Destination.ListIndex = XMLIDestination_Type_DigiFrame)
End Sub

Private Sub f_Filename_GotFocus()
    f_Filename.SelStart = 0
    f_Filename.SelLength = Len(f_Filename.Text)
End Sub

Private Sub f_Directory_GotFocus()
    f_Directory.SelStart = 0
    f_Directory.SelLength = Len(f_Directory.Text)
End Sub

