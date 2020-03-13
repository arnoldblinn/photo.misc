VERSION 5.00
Object = "{00120003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "LTOCX12N.OCX"
Begin VB.Form frmMySaver 
   BorderStyle     =   0  'None
   Caption         =   "Preview"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin LEADLib.LEAD LEAD1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
      _Version        =   65539
      _ExtentX        =   1296
      _ExtentY        =   1296
      _StockProps     =   229
      ScaleHeight     =   49
      ScaleWidth      =   49
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
   Begin VB.Timer TimerNotify 
      Interval        =   5
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer tmrExitNotify 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   5400
   End
End
Attribute VB_Name = "frmMySaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MySaver.frm
Option Explicit

Const gDebug = True

Rem Configuration
Dim gVerticalDelta As Integer
Dim gHorizontalDelta As Integer
Dim gMoveInterval As Long
Dim gFlipInterval As Long
Dim gBackgroundColor As Long

Const L_SUPPORT_GIFLZW = 1
Const k_UnlockKey_GIF_V12 = "sg8Z2XkjL"

Rem Configuration for the files
Dim gDirectory As String
Dim gFile As String

Rem Pseudo-constants
Dim gScreenHeight As InvertedTextFlags
Dim gScreenWidth As Integer

Rem State
Dim gMouseX As Single
Dim gMouseY As Single
Dim gLastMove As Long
Dim gLastFlip As Long
Dim gFreezeHorizontal As Boolean
Dim gFreezeVertical As Boolean
Dim gFreezeFlip As Boolean
Dim gFileNames() As String
Dim gFileCount As Long
Dim gCurrentFile As Long

Rem API function to get the system tickcount
Private Declare Function GetTickCount _
    Lib "kernel32" ( _
) As Long

Rem API function to hide/show the mouse pointer
Private Declare Function ShowCursor _
Lib "user32" ( _
    ByVal bShow As Long _
) As Long

Rem API function to signal activity to system
Private Declare Function SystemParametersInfo _
Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    ByRef lpvParam As Any, _
    ByVal fuWinIni As Long _
) As Long

Rem Constant for API function
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const EFX_TEXTALIGN_HCENTER_VCENTER = 4

Private Sub Form_Load()
    Dim lngRet As Long
    Dim objFSO As FileSystemObject
    Dim objFolder As Folder
    Dim objFile As File
    
    Rem Initialize state and pseudo constants
    gScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    gScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    gMouseX = 0
    gMouseY = 0
    gLastMove = GetTickCount()
    gLastFlip = GetTickCount()
    gFreezeHorizontal = False
    gFreezeVertical = False
    
    Rem Initialize stuff from the "configuration"
    gVerticalDelta = CLng(GetSetting("MySaver", "Options", "VerticalDelta", 5))
    gHorizontalDelta = CLng(GetSetting("MySaver", "Options", "HorizontalDelta", 5))
    gMoveInterval = CLng(GetSetting("MySaver", "Options", "MoveInterval", 100))
    gFlipInterval = CLng(GetSetting("MySaver", "Options", "FlipInterval", 10000))
    gBackgroundColor = CLng(GetSetting("MySaver", "Options", "BackgroundColor", RGB(0, 0, 0)))
    gDirectory = GetSetting("MySaver", "Options", "Directory", "")
    gFile = GetSetting("MySaver", "Options", "File", "")
    
    Rem Initialize the lead control's position
    Call LEAD1.UnlockSupport(L_SUPPORT_GIFLZW, k_UnlockKey_GIF_V12)
    LEAD1.Top = 0
    LEAD1.Left = 0
    
    Rem Get the image/images to display
    Set objFSO = New FileSystemObject
    If gDirectory <> "" And Not IsNull(gDirectory) And objFSO.FolderExists(gDirectory) Then
                
        Set objFolder = objFSO.GetFolder(gDirectory)
        
        gFileCount = 0
        For Each objFile In objFolder.Files
        
            If LCase(Right(objFile.Name, 4)) = ".jpg" Then
                gFileCount = gFileCount + 1
                ReDim Preserve gFileNames(gFileCount)
                gFileNames(gFileCount - 1) = gDirectory + objFile.Name
            End If
        Next
        
        If gFileCount > 0 Then
            gCurrentFile = 0
            Call SwapImage(gFileNames(0))
        Else
            Call SwapImage("")
            gFreezeFlip = True
        End If
    ElseIf gFile <> "" And Not IsNull(gFile) And objFSO.FileExists(gFile) Then
        Call SwapImage(gFile)
        gFreezeFlip = True
    Else
        Call SwapImage("")
        gFreezeFlip = True
    End If

    Rem Prepare to run
    lngRet = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 0, ByVal 0&, 0)
    
    If gblnShow = True Then
        Me.WindowState = vbMaximized
    End If
    
    Me.BackColor = gBackgroundColor
    If gblnShow = True And gDebug = False Then
        ShowCursor (False)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Using End here appears to prevent memory leaks
    End
End Sub

Private Sub Form_Click()
    'Quit if mouse is clicked, unless in preview mode
    If gblnShow = True Then
        tmrExitNotify.Enabled = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Quit if any key is pressed, unless in preview mode
    If gblnShow = True Then
        tmrExitNotify.Enabled = True
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static sngTimer As Single
    
    'Bail out quickly if in preview mode
    If gblnShow = False Then
        Exit Sub
    End If
    
    If gMouseX = 0 And gMouseY = 0 Then
        gMouseX = X
        gMouseY = Y
    End If
    
    If gMouseX = X And gMouseY = Y Then
        Exit Sub
    End If
    
    'Quit any time after first .25 seconds
    If sngTimer = 0 Then
        sngTimer = Timer
    ElseIf Timer > sngTimer + 0.25 Then
        tmrExitNotify.Enabled = True
    End If
End Sub



Private Sub LEAD1_KeyDown(KeyCode As Integer, Shift As Integer)
    If gblnShow = True Then
        tmrExitNotify.Enabled = True
    End If
End Sub

Private Sub SwapImage(strImage As String)
    Dim oldWidth As Integer
    Dim oldHeight As Integer
    
    If LEAD1.Bitmap Then
        oldWidth = LEAD1.BitmapWidth
        oldHeight = LEAD1.BitmapHeight
    Else
        oldWidth = 0
        oldHeight = 0
    End If
    
    Rem Load the new image
    If strImage = "" Then
        Call LEAD1.CreateBitmap(300, 200, 24)
        LEAD1.TextTop = 0
        LEAD1.TextLeft = 0
        LEAD1.TextWidth = 300
        LEAD1.TextHeight = 200
        LEAD1.DrawFontColor = RGB(255, 255, 255)
        LEAD1.DrawPersistence = True
    
    
        LEAD1.TextAlign = EFX_TEXTALIGN_HCENTER_VCENTER
        LEAD1.Font.Name = "Verdana,Sans-Serif"
        LEAD1.Font.Size = 9
        LEAD1.Font.Weight = 675
            
        LEAD1.Fill RGB(0, 0, 0)
        Call LEAD1.DrawText("No files configured to display", 0)
    Else
        Call LEAD1.Load(strImage, 0, 0, 1)
    End If
    
    
    Rem Size the control
    LEAD1.Width = LEAD1.BitmapWidth
    LEAD1.Height = LEAD1.BitmapHeight
    If oldWidth <> 0 And oldHeight <> 0 Then
        LEAD1.Left = LEAD1.Left + (oldWidth - LEAD1.Width) / 2
        LEAD1.Top = LEAD1.Top + (oldHeight - LEAD1.Height) / 2
    Else
        LEAD1.Left = (gScreenWidth - LEAD1.Width) / 2
        LEAD1.Top = (gScreenHeight - LEAD1.Height) / 2
    End If
    
    Rem Make sure we are positioned OK
    If gHorizontalDelta <> 0 And LEAD1.BitmapWidth < gScreenWidth Then
        If LEAD1.Left + LEAD1.Width > gScreenWidth Then
            LEAD1.Left = gScreenWidth - LEAD1.Width
        End If
        
        gFreezeHorizontal = False
    Else
        LEAD1.Left = (gScreenWidth - LEAD1.Width) / 2
        gFreezeHorizontal = True
    End If
    
    If gVerticalDelta <> 0 And LEAD1.BitmapHeight < gScreenHeight Then
        If LEAD1.Top + LEAD1.Height > gScreenHeight Then
            LEAD1.Top = gScreenHeight - LEAD1.Height
        End If
        gFreezeVertical = False
    Else
        LEAD1.Top = (gScreenHeight - LEAD1.Height) / 2
        gFreezeVertical = True
    End If
End Sub

Private Sub TimerNotify_Timer()
    Dim Right
    Dim Bottom
    
    Rem Determine if we need to swap images
    If gFreezeFlip = False And gLastFlip + gFlipInterval < GetTickCount() Then
        If gCurrentFile = gFileCount - 1 Then
            gCurrentFile = 0
        Else
            gCurrentFile = gCurrentFile + 1
        End If
        Call SwapImage(gFileNames(gCurrentFile))
        
        gLastFlip = GetTickCount()
    End If
    
    Rem Determine if we need to move
    If gLastMove + gMoveInterval < GetTickCount() Then
        gLastMove = GetTickCount()
                
        If gFreezeHorizontal = False And LEAD1.Width < gScreenWidth Then
            Right = LEAD1.Left + LEAD1.Width
            If Right + gHorizontalDelta > gScreenWidth Then
                Right = gScreenWidth - gHorizontalDelta + (gScreenWidth - Right)
                LEAD1.Left = Right - LEAD1.Width
                gHorizontalDelta = -1 * gHorizontalDelta
            ElseIf LEAD1.Left + gHorizontalDelta < 0 Then
                gHorizontalDelta = -1 * gHorizontalDelta
                LEAD1.Left = gHorizontalDelta - LEAD1.Left
            Else
                LEAD1.Left = LEAD1.Left + gHorizontalDelta
            End If
        End If
    
        If gFreezeVertical = False And LEAD1.Height < gScreenHeight Then
            Bottom = LEAD1.Top + LEAD1.Height
            If Bottom + gVerticalDelta > gScreenHeight Then
                Bottom = gScreenHeight - gVerticalDelta + (gScreenHeight - Bottom)
                LEAD1.Top = Bottom - LEAD1.Height
                gVerticalDelta = -1 * gVerticalDelta
            ElseIf LEAD1.Top + gVerticalDelta < 0 Then
                gVerticalDelta = -1 * gVerticalDelta
                LEAD1.Top = gVerticalDelta - LEAD1.Top
            Else
                LEAD1.Top = LEAD1.Top + gVerticalDelta
            End If
        End If
    End If
End Sub

Private Sub tmrExitNotify_Timer()
    Dim lngRet As Long
    
    'Tell system that screen saver is done
    lngRet = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 1, ByVal 0&, 0)
    
    'Time to quit
     If gblnShow = True And gDebug = False Then
        ShowCursor (True)
     End If

    End
End Sub

