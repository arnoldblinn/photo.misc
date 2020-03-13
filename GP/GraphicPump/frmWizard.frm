VERSION 5.00
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Task Wizard"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton b_Back 
      Caption         =   "<< Back"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton b_Finish 
      Caption         =   "Finish"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton b_Next 
      Caption         =   "Next >>"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame frame_Destination 
      Caption         =   "Select Destination"
      Height          =   3735
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cmb_Destination 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Select what the images are being used for from the list below.  "
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame frame_DigiFrame 
      Caption         =   "Select Your Digi-Frame Options"
      Height          =   3735
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton b_Detect 
         Caption         =   "Detect"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox cmb_Media 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   4215
      End
      Begin VB.ComboBox cmb_Port 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Select the storage media on the Digi-Frame for the images"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Select the serial port the Digi-Frame is connected to"
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame frame_Finish 
      Caption         =   "Finish"
      Height          =   3735
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   5055
      Begin VB.Label Label3 
         Caption         =   "Congratulations!! To finish defining your task, click Finish."
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   $"frmWizard.frx":0000
         Height          =   1095
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label8 
         Caption         =   "To execute your task, click ""Run"" from the main task dialog."
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   4335
      End
   End
   Begin VB.Frame frame_Source 
      Caption         =   "Select Images to Move"
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton b_MoveDown 
         Caption         =   "Move Down"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton b_MoveUp 
         Caption         =   "Move Up"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton b_Delete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton b_Edit 
         Caption         =   "Edit..."
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton b_New 
         Caption         =   "New..."
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ListBox l_List 
         Height          =   2205
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   0
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Click Move Up or Move Down to reorder your images."
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Click New to add a new image file or image web address to the task, or drag a file into the list of images."
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame frame_Directory 
      Caption         =   "Select a Directory"
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton b_Browse 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox f_Directory 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label l_DirectoryLabel 
         Caption         =   "Directory"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label l_DirectoryDetail 
         Caption         =   "Select a directory"
         Height          =   2295
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gResult As Boolean
Dim gForm As Form
Dim gastrNames()
Dim gastrImages()
Dim gCount As Integer
Dim gstrXML As String
Dim gstrName As String
Dim giRotate As Integer
Dim giPad As Integer
Dim giWidth As Integer
Dim giHeight As Integer
Dim giThumbnail As Integer
Dim giThumbWidth As Integer
Dim giThumbHeight As Integer

Rem State of the wizard
Dim gState As Integer

Rem Wizard State constants
Const WS_Destination = 0
Const WS_DigiFrame = 1
Const WS_Directory = 2
Const WS_Source = 3
Const WS_Finish = 4

Rem Wizard Destinations
Const WD_Directory = 0
Const WD_ScreenSaver = 1
Const WD_DigiFrame560 = 2
Const WD_DigiFrame390 = 3

Public Property Get XML() As String
    XML = gstrXML
End Property

Public Property Get Result() As Boolean
    Result = gResult
End Property

Private Sub b_Browse_Click()
    Dim strDirectoryName As String
    
    strDirectoryName = f_Directory.Text
    frmDirectory.Path = strDirectoryName
    frmDirectory.Show (vbModal)
    If frmDirectory.Path <> "" Then
        f_Directory.Text = frmDirectory.Path
    End If
    Unload frmDirectory
End Sub

Private Sub b_Detect_Click()
    Dim objDigiFrame As clsDigiFrame
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

Private Sub b_Finish_Click()
    Dim strStatus As String
    Dim strSource As String
    Dim strSchedule As String
    Dim strFormat As String
    Dim strDestination As String
    Dim strXML As String
    Dim i As Integer
    Dim strDirectory As String
    Dim iPort As Integer
    Dim iMedia As Integer
    Dim iType As Integer
    
    
    Rem Build the status
    strStatus = "<Status><LastRun></LastRun><Failed>0</Failed><Reason></Reason></Status>"
    
    Rem Build the source
    strSource = "<Source><Type>0</Type><Album Version=""1""><Name>" & HTMLEncode(gstrName) & "</Name><PictureList>"
    For i = 0 To gCount - 1
        strSource = strSource & "<Picture><Name>" & HTMLEncode(gastrNames(i)) & "</Name><URI>" & HTMLEncode(gastrImages(i)) & "</URI></Picture>"
    Next
    strSource = strSource & "</PictureList></Album><AlbumURI></AlbumURI></Source>"
    strSchedule = "<Schedule><Type>0</Type><Disable>0</Disable><Connect>1</Connect><Hours>0</Hours><Minutes>0</Minutes><Weekday>0</Weekday><Monthday>1</Monthday></Schedule>"
    strFormat = "<Format><Name></Name><FormatSettings><Width>" & giWidth & "</Width><Height>" & giHeight & "</Height><Grow>1</Grow><Shrink>1</Shrink><Rotate>" & giRotate & "</Rotate><RotateDirection>0</RotateDirection><Pad>" & giPad & "</Pad><PadColor>0</PadColor><VerticalAlign>1</VerticalAlign><HorizontalAlign>1</HorizontalAlign><Margins>0</Margins><TopMargin>5</TopMargin><LeftMargin>5</LeftMargin><RightMargin>5</RightMargin><BottomMargin>5</BottomMargin><MarginColor>" & RGB(255, 255, 255) & "</MarginColor><Compression>2</Compression><Thumbnail>" & giThumbnail & "</Thumbnail><ThumbWidth>" & giThumbWidth & "</ThumbWidth><ThumbHeight>" & giThumbHeight & "</ThumbHeight></FormatSettings></Format>"
        
    If cmb_Destination.ListIndex = WD_DigiFrame560 Or cmb_Destination.ListIndex = WD_DigiFrame390 Then
        iType = XMLIDestination_Type_DigiFrame
    ElseIf cmb_Destination.ListIndex = WD_ScreenSaver Then
        iType = XMLIDestination_Type_Directory
    Else
        iType = XMLIDestination_Type_Directory
    End If
    
    strDestination = "<Destination><Type>" & iType & "</Type><Directory>" & HTMLEncode(f_Directory.Text) & "</Directory><DirectoryDelete>0</DirectoryDelete><FileTemplate>file%i%.jpg</FileTemplate><DigiFramePort>" & cmb_Port.ListIndex & "</DigiFramePort><DigiFrameMedia>" & cmb_Media.ListIndex & "</DigiFrameMedia></Destination>"
    
    gstrXML = "<Task Version=""1""><Name>" & HTMLEncode(gstrName) & "</Name>" & strStatus & strSource & strSchedule & strFormat & strDestination & "</Task>"
    
    gResult = True
    gForm.Hide
End Sub

Private Sub b_Cancel_Click()
    gResult = False
    gForm.Hide
End Sub

Private Function SetState(iState As Integer)
    Rem Validate the old state
    If gState = WS_Directory Then
        Dim fso As Scripting.FileSystemObject
        Set fso = New Scripting.FileSystemObject
        If Not fso.FolderExists(f_Directory.Text) Then
            MsgBox "Directory does not exist", vbExclamation, "Error"
            Exit Function
        End If
    End If
    
    Rem If we are leaving the destination state, then set some interesting globals
    giThumbnail = 0
    giThumbWidth = 0
    giThumbHeight = 0
    If cmb_Destination.ListIndex = WD_Directory Then
        giWidth = 640
        giHeight = 480
        giRotate = 0
        giPad = 0
        gstrName = "Images"
        l_DirectoryDetail.Caption = "Please select a destination directory for your images."
    ElseIf cmb_Destination.ListIndex = WD_ScreenSaver Then
        giWidth = 400
        giHeight = 400
        giRotate = 0
        giPad = 0
        gstrName = "Screen Saver Images"
        l_DirectoryDetail.Caption = "Please select a destination directory for your Screen Saver images. This can be any directory on the disk, but should match the directory being used by the current Screen Saver." & vbCr & vbCr
        l_DirectoryDetail.Caption = l_DirectoryDetail.Caption & "Any Screen Saver that displays images in a directory on the hard disk can be used, including the Graphic Pump Screen Saver installed with this application." & vbCr & vbCr
    ElseIf cmb_Destination.ListIndex = WD_DigiFrame560 Then
        giWidth = 640
        giHeight = 480
        giRotate = 0
        giPad = 0
        gstrName = "Digi-Frame 560"
    ElseIf cmb_Destination.ListIndex = WD_DigiFrame390 Then
        giWidth = 320
        giHeight = 240
        giRotate = 0
        giPad = 0
        gstrName = "Digi-Frame 390"
    End If
    
    l_DirectoryLabel.Caption = "Directory for the " & gstrName

    Rem Set the new state
    frame_Destination.Visible = (iState = WS_Destination)
    frame_Source.Visible = (iState = WS_Source)
    frame_Directory.Visible = (iState = WS_Directory)
    frame_DigiFrame.Visible = (iState = WS_DigiFrame)
    frame_Finish.Visible = (iState = WS_Finish)
    
    Rem Enable/Disable the buttons as necessary
    b_Back.Enabled = (iState <> WS_Destination)
    b_Next.Visible = (iState <> WS_Finish)
    b_Next.Default = b_Next.Enabled
    b_Finish.Visible = (iState = WS_Finish)
    b_Finish.Default = b_Finish.Visible
    gState = iState
End Function

Private Sub b_Back_Click()
        
    If gState = WS_Finish Then
        SetState (WS_Source)
    ElseIf gState = WS_Source Then
        If cmb_Destination.ListIndex = WD_DigiFrame560 Or cmb_Destination.ListIndex = WD_DigiFrame390 Then
            SetState (WS_DigiFrame)
        Else
            SetState (WS_Directory)
        End If
    Else
        SetState (WS_Destination)
    End If
End Sub


Private Sub b_Next_Click()
    If gState = WS_Destination Then
        If cmb_Destination.ListIndex = WD_DigiFrame560 Or cmb_Destination.ListIndex = WD_DigiFrame390 Then
            SetState (WS_DigiFrame)
        Else
            SetState (WS_Directory)
        End If
    ElseIf gState = WS_Directory Or gState = WS_DigiFrame Then
        SetState (WS_Source)
    ElseIf gState = WS_Source Then
        SetState (WS_Finish)
    End If
End Sub


Private Sub f_Directory_GotFocus()
    f_Directory.SelStart = 0
    f_Directory.SelLength = Len(f_Directory.Text)
End Sub

Private Sub Form_Load()
    
    frame_Destination.Visible = True
    frame_Source.Visible = False
    frame_Directory.Visible = False
    frame_DigiFrame.Visible = False
    frame_Finish.Visible = False
    b_Finish.Visible = False
    b_Back.Enabled = False
    
    Call cmb_Destination.AddItem("Simple Directory of Images", WD_Directory)
    Call cmb_Destination.AddItem("Screen Saver", WD_ScreenSaver)
    Call cmb_Destination.AddItem("Digi-Frame 560", WD_DigiFrame560)
    Call cmb_Destination.AddItem("Digi-Frame 390", WD_DigiFrame390)
    
    cmb_Destination.ListIndex = 0
    
    Call cmb_Port.AddItem("COM Port 1", XMLIDestination_DigiFramePort_COM1)
    Call cmb_Port.AddItem("COM Port 2", XMLIDestination_DigiFramePort_COM2)
    Call cmb_Port.AddItem("COM Port 3", XMLIDestination_DigiFramePort_COM3)
    Call cmb_Port.AddItem("COM Port 4", XMLIDestination_DigiFramePort_COM4)
    cmb_Port.ListIndex = XMLIDestination_DigiFramePort_COM1
    
    Call cmb_Media.AddItem("Compact Flash", XMLIDestination_DigiFrameMedia_CompactFlash)
    Call cmb_Media.AddItem("SmartMedia", XMLIDestination_DigiFrameMedia_SmartMedia)
    cmb_Media.ListIndex = XMLIDestination_DigiFrameMedia_CompactFlash
    
    f_Directory.Text = CurDir
    
    gState = WS_Destination
    gCount = 0
    
    Set gForm = Me
End Sub

Private Sub b_New_Click()
    Dim objXMLDom As MSXML.DOMDocument
    Set objXMLDom = New MSXML.DOMDocument
    objXMLDom.loadXML ("<Picture><Name>New Image</Name><URI></URI></Picture>")
    
    Set frmPicture.XMLDomNode = objXMLDom.childNodes.Item(0)
    frmPicture.Show (vbModal)
    If frmPicture.Result = True Then
        ReDim Preserve gastrNames(gCount + 1)
        ReDim Preserve gastrImages(gCount + 1)
        gastrNames(gCount) = frmPicture.XMLDomNode.childNodes(XMLIPicture_Name).Text
        gastrImages(gCount) = frmPicture.XMLDomNode.childNodes(XMLIPicture_URI).Text
        gCount = gCount + 1
    
        Call PopulateList
    
        l_List.ListIndex = l_List.ListCount - 1
    End If
    Unload frmPicture
    
End Sub

Private Sub b_Edit_Click()
    Dim objXMLDom As MSXML.DOMDocument
    Set objXMLDom = New MSXML.DOMDocument
    Dim strName As String
    Dim strImage As String
    Dim i As Integer
    
    i = l_List.ListIndex
    
    strName = gastrNames(i)
    strImage = gastrImages(i)
    objXMLDom.loadXML ("<Picture><Name>" & HTMLEncode(strName) & "</Name><URI>" & HTMLEncode(strImage) & "</URI></Picture>")
    
    Set frmPicture.XMLDomNode = objXMLDom.childNodes.Item(0)
    frmPicture.Show (vbModal)
    If frmPicture.Result = True Then
        gastrNames(i) = frmPicture.XMLDomNode.childNodes(XMLIPicture_Name).Text
        gastrImages(i) = frmPicture.XMLDomNode.childNodes(XMLIPicture_URI).Text
    
        Call PopulateList
    
        l_List.ListIndex = i
    End If
    Unload frmPicture
End Sub

Private Sub b_MoveUp_Click()
    Dim i As Integer
    Dim strTempName As String
    Dim strTempImage As String
    
    i = l_List.ListIndex
    strTempName = gastrNames(i)
    strTempImage = gastrImages(i)
    gastrNames(i) = gastrNames(i - 1)
    gastrImages(i) = gastrImages(i - 1)
    gastrNames(i - 1) = strTempName
    gastrImages(i - 1) = strTempImage
    
    Call PopulateList
    
    l_List.ListIndex = i - 1
End Sub

Private Sub b_MoveDown_Click()
    Dim i As Integer
    Dim strTempName As String
    Dim strTempImage As String
    
    i = l_List.ListIndex
    strTempName = gastrNames(i)
    strTempImage = gastrImages(i)
    gastrNames(i) = gastrNames(i + 1)
    gastrImages(i) = gastrImages(i + 1)
    gastrNames(i + 1) = strTempName
    gastrImages(i + 1) = strTempImage
    
    Call PopulateList
    
    l_List.ListIndex = i + 1
End Sub

Private Sub b_Delete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = l_List.ListIndex
    For i = j To l_List.ListCount - 2
        gastrNames(i) = gastrNames(i + 1)
        gastrImages(i) = gastrImages(i + 1)
    Next
    
    gCount = gCount - 1
    
    Call PopulateList
    
    If l_List.ListCount <= j Then
        l_List.ListIndex = l_List.ListCount - 1
    Else
        l_List.ListIndex = j
    End If
End Sub

Private Sub PopulateList()
    Dim i As Integer
    
    b_Edit.Enabled = False
    b_Delete.Enabled = False
    b_MoveDown.Enabled = False
    b_MoveUp.Enabled = False
    
    For i = 0 To l_List.ListCount - 1
        l_List.RemoveItem (0)
    Next
    
    For i = 0 To gCount - 1
        l_List.AddItem (gastrNames(i) & " : " & gastrImages(i))
    Next

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

Private Sub l_List_DblClick()
    Call b_Edit_Click
End Sub

Private Sub l_List_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        Dim vfn

        For Each vfn In Data.Files
            ReDim Preserve gastrNames(gCount + 1)
            ReDim Preserve gastrImages(gCount + 1)
            
            gastrNames(gCount) = vfn
            gastrImages(gCount) = vfn
            
            gCount = gCount + 1
        Next vfn
        
        Call PopulateList
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
