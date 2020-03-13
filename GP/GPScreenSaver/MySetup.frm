VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMySetup 
   Caption         =   "Graphic Pump Saver - Setup"
   ClientHeight    =   5220
   ClientLeft      =   585
   ClientTop       =   2490
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton b_About 
      Caption         =   "About..."
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Flip"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   4575
      Begin VB.TextBox f_FlipInterval 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Time (ms)"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.CommandButton b_Directory 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox f_Image 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton b_BackgroundColor 
      Caption         =   "Change..."
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movement"
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   4575
      Begin VB.TextBox f_VerticalDelta 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox f_HorizontalDelta 
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox f_MoveInterval 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Vertical Delta (pixels)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Horizontal Delta (pixels)"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Speed (ms)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Image/Path"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Background Color"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1920
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmMySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gVerticalDelta As Long
Dim gHorizontalDelta As Long
Dim gMoveInterval As Long
Dim gFlipInterval As Long
Dim gBackgroundColor As Long
Dim gDirectory As String
Dim gFile As String


Private Sub b_About_Click()
    frmAbout.Show vbModal
End Sub

Private Sub b_BackgroundColor_Click()
    CommonDialog1.Color = gBackgroundColor
    
    CommonDialog1.Flags = cdlCCRGBInit
    
    CommonDialog1.ShowColor
    
    gBackgroundColor = CommonDialog1.Color
    
    Shape1.FillColor = gBackgroundColor
End Sub


Private Sub b_Directory_Click()
    
    frmMyImagePath.ImagePath = f_Image.Text
    
    Call frmMyImagePath.Show(vbModal)
    
    f_Image.Text = frmMyImagePath.ImagePath
End Sub

Private Sub Form_Load()

    Rem Get values from registry
    gVerticalDelta = CLng(GetSetting("MySaver", "Options", "VerticalDelta", 2))
    gHorizontalDelta = CLng(GetSetting("MySaver", "Options", "HorizontalDelta", 2))
    gMoveInterval = CLng(GetSetting("MySaver", "Options", "MoveInterval", 100))
    gFlipInterval = CLng(GetSetting("MySaver", "Options", "FlipInterval", 10000))
    gBackgroundColor = CLng(GetSetting("MySaver", "Options", "BackgroundColor", RGB(0, 0, 0)))
    gDirectory = GetSetting("MySaver", "Options", "Directory", "")
    gFile = GetSetting("MySaver", "Options", "File", "")
    
    Rem Initialize the controls
    f_MoveInterval.Text = gMoveInterval
    f_HorizontalDelta.Text = gHorizontalDelta
    f_VerticalDelta.Text = gVerticalDelta
    Shape1.FillColor = gBackgroundColor
    If gDirectory <> "" Then
        f_Image.Text = gDirectory
    ElseIf gFile <> "" Then
        f_Image.Text = gFile
    End If
    f_FlipInterval.Text = gFlipInterval
    
    Rem Center this form
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.Show
End Sub

Private Sub cmdCancel_Click()
    'Close the dialog box without saving changes
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim fso As Scripting.FileSystemObject
    
    On Error Resume Next
    gMoveInterval = CLng(f_MoveInterval.Text)
    If Err.Number <> 0 Or gMoveInterval < 100 Or gMoveInterval > 50000 Then
        
        MsgBox "Speed must be a number between 100 and 50000", vbExclamation, "Error"
        f_MoveInterval.SetFocus
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error Resume Next
    gHorizontalDelta = CLng(f_HorizontalDelta.Text)
    If Err.Number <> 0 Or gHorizontalDelta < 0 Or gHorizontalDelta > 30 Then
        
        MsgBox "Horizontal Delta must be a number between 0 and 30", vbExclamation, "Error"
        f_HorizontalDelta.SetFocus
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error Resume Next
    gVerticalDelta = CLng(f_VerticalDelta.Text)
    If Err.Number <> 0 Or gVerticalDelta < 0 Or gVerticalDelta > 30 Then
        
        MsgBox "Vertical Delta must be a number between 0 and 30", vbExclamation, "Error"
        f_VerticalDelta.SetFocus
        Exit Sub
    End If
    On Error GoTo 0
    
    Set fso = New Scripting.FileSystemObject
    If fso.FolderExists(f_Image.Text) Then
        If Right(f_Image.Text, 1) = "\" Then
            gDirectory = f_Image.Text
        Else
            gDirectory = f_Image.Text & "\"
        End If
    ElseIf fso.FileExists(f_Image.Text) Then
        If Right(f_Image.Text, 4) = ".jpg" Then
            gDirectory = ""
            gFile = f_Image.Text
        Else
            MsgBox "Image must be either a directory or a .jpg file"
            Exit Sub
        End If
    Else
        MsgBox "Image must be either a directory of a .jpg file"
        Exit Sub
    End If
    
    On Error Resume Next
    gFlipInterval = CLng(f_FlipInterval.Text)
    If Err.Number <> 0 Or gFlipInterval < 100 Or gFlipInterval > 50000 Then
        
        MsgBox "Flip must be a number between 100 and 50000", vbExclamation, "Error"
        f_FlipInterval.SetFocus
        Exit Sub
    End If
    On Error GoTo 0
    
    Rem Save the settings
    Call SaveSetting("MySaver", "Options", "VerticalDelta", CStr(gVerticalDelta))
    Call SaveSetting("MySaver", "Options", "HorizontalDelta", CStr(gHorizontalDelta))
    Call SaveSetting("MySaver", "Options", "MoveInterval", CStr(gMoveInterval))
    Call SaveSetting("MySaver", "Options", "FlipInterval", CStr(gFlipInterval))
    Call SaveSetting("MySaver", "Options", "BackgroundColor", CStr(gBackgroundColor))
    Call SaveSetting("MySaver", "Options", "Directory", CStr(gDirectory))
    Call SaveSetting("MySaver", "OPtions", "File", CStr(gFile))
    
    Rem Close the dialog
    Unload Me
End Sub

