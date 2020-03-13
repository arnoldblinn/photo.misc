VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C155360-3CD1-11D0-B17A-E18E3EAC3833}#1.0#0"; "SI_COMM.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphic Pump"
   ClientHeight    =   1800
   ClientLeft      =   48
   ClientTop       =   612
   ClientWidth     =   6372
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6372
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -120
      TabIndex        =   2
      Top             =   0
      Width           =   8175
   End
   Begin VB.Frame frame_Status 
      Caption         =   "Execution Status"
      Height          =   1575
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin MSComctlLib.ProgressBar p_Progress 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3975
         _ExtentX        =   7006
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label l_Status 
         Caption         =   "No tasks currently running"
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   3855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   5880
      Top             =   1320
   End
   Begin SI_COMMLib.SI_COMM SI_COMM1 
      Left            =   3840
      Top             =   1440
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ModemErrorMsg   =   "ERROR"
      ModemConnectSuccessMsg=   "CONNECT"
      ModemConnectFailureMsg=   "NO CARRIER"
      ModemLineBusyMsg=   "BUSY"
      ModemNoDialToneMsg=   "NO DIALTONE"
      ModemVerboseCmd =   "V1Q0"
      ModemAnswerCmd  =   "S0="
      ModemCarrierSpeedMsg=   "CARRIER"
      ModemInitCmd    =   "V1Q0"
      ModemToneDialCmd=   "DT"
      ModemPulseDialCmd=   "DD"
      ModemResetCmd   =   "Z"
      ModemEscCmd     =   "++"
      ModemSuccessMsg =   "OK"
      ModemCommandPrefix=   "AT"
      ModemCommandSuffix=   "Chr$(0D)"
      ModemHangUpCmd  =   "H"
      ModemCmdStringPaceTime=   1
      ModemEscStringPaceTime=   2
      ModemResponseTime=   2
      ModemConnectTime=   30
      ModemEchoTime   =   2
      ModemResetTime  =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   312
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Image PumpIconNormal 
      Height          =   384
      Left            =   5160
      Picture         =   "frmMain.frx":0442
      Top             =   1320
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1545
      Left            =   120
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":0884
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image PumpIconActive 
      Height          =   384
      Left            =   4560
      Picture         =   "frmMain.frx":C0A3
      Top             =   1320
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPumpFile 
         Caption         =   "&Pump Image"
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTray 
         Caption         =   "&Send To Tray"
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit1 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFormat 
         Caption         =   "&Format Profiles"
      End
      Begin VB.Menu mnuTasks 
         Caption         =   "&Tasks"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewStatus 
         Caption         =   "View &Status"
      End
      Begin VB.Menu mnuAlwaysTop 
         Caption         =   "&Always on Top"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&Registraion/About"
      End
   End
   Begin VB.Menu mnuPopupMenu 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigure 
         Caption         =   "Configure"
      End
      Begin VB.Menu mnuBreak4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Rem Width and height of full form

Rem UI State
Dim gState As Integer

Rem Various states the application can be in
Const STATE_Idle = 0
Const STATE_Edit = 2
Const STATE_Run = 3

Rem Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA


Private Function make_visible()

    If Me.Left < 0 Then
        Me.Left = 0
    End If
    If Me.Left + Me.Width > Screen.Width Then
        Me.Left = Screen.Width - Me.Width
    End If
    
    If Me.Top < 0 Then
        Me.Top = 0
    End If
    If Me.Top + Me.Height > Screen.Height Then
        Me.Top = Screen.Height - Me.Height
    End If
    
End Function




Rem ---------------------------------------------------
Rem Form_Load
Rem
Rem Initializes the object
Rem
Private Sub Form_Load()
    
    Rem Set the individual values of the NOTIFYICONDATA data type.
    nid.cbSize = Len(nid)
    nid.hwnd = frmMain.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    Rem nid.uCallBackMessage = WM_DROPFILES
    nid.hIcon = Me.Icon
    nid.szTip = "Graphic Pump" & vbNullChar

    Rem Call the Shell_NotifyIcon function to add the icon to the taskbar
    Shell_NotifyIcon NIM_ADD, nid
    
    Rem Restore the show state
    Dim strShowStatus As String
    strShowStatus = RegReadStringDefault("ShowStatus", "1")
    If strShowStatus = "1" Then
        Call show_status
    Else
        Call hide_status
    End If
    
    Rem Restore the position on the screen
    Me.Left = RegReadIntDefault("Left", (Screen.Width - Me.Width) / 2)
    Me.Top = RegReadIntDefault("Top", (Screen.Height - Me.Height) / 2)
    
    Rem Ensure that the window is visible on the screen
    Call make_visible
    
    Rem Restore the "always on top" attribute
    If RegReadIntDefault("AlwaysTop", 0) = 1 Then
        Call set_topmost
    End If
                
    Rem Hide ourselves
    Me.Hide
    gState = STATE_Idle

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gState <> STATE_Idle Then
        Cancel = 1
        Exit Sub
    End If
    
    If UnloadMode = 0 Then
    
     If mnuAlwaysTop.Checked Then
            Call clear_topmost
            frmExit.Show vbModal
            Call set_topmost
        Else
            frmExit.Show vbModal
        End If
        If frmExit.Result = 0 Then
            Cancel = 1
            Me.Hide
        ElseIf frmExit.Result = 1 Then
            Cancel = 0
        Else
            Cancel = 1
        End If
    End If
    
    Unload frmExit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rem Delete the added icon from the taskbar status area when the program ends.
    Shell_NotifyIcon NIM_DELETE, nid
    
    Rem Save the current position
    Call RegWrite("Top", CStr(Me.Top))
    Call RegWrite("Left", CStr(Me.Left))
End Sub


Private Sub Command1_Click()
    Call mnuPumpFile_Click
End Sub

Private Sub Image1_Click()
    Call mnuPumpFile_Click
End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim vfn As Variant
    
    If Data.GetFormat(vbCFFiles) Then
        If Data.Files.Count = 1 Then
            For Each vfn In Data.Files
                frmSave.SourceFile = vfn
                
                frmMain.SetFocus
                
                Call PostMessage(Me.Command1.hwnd, WM_LBUTTONDOWN, 0, &H20002)
                Call PostMessage(Me.Command1.hwnd, WM_LBUTTONUP, 0, &H20002)
                Rem Call PostMessage(Me.hwnd, 0, 0, 0)
            Next
        End If
    End If

End Sub

Private Sub Image1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    If Data.GetFormat(vbCFFiles) Then
        If Data.Files.Count = 1 Then
            Effect = vbDropEffectCopy And Effect
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Rem -----------------------------------------------------------
Rem set_topmost
Rem clear_topmost
Rem mnuAlwaysTop_Click
Rem
Rem Functions to deal with the topmost window
Rem
Private Sub set_topmost()
    SetWindowPos hwnd, conHwndTopmost, 0, 0, 0, 0, conSwpNoActivate Or conSwpShowWindow Or conSwpNoSize Or conSwpNoMove
    mnuAlwaysTop.Checked = True
End Sub

Private Sub clear_topmost()
    SetWindowPos hwnd, conHwndNoTopmost, 0, 0, 0, 0, conSwpNoActivate Or conSwpShowWindow Or conSwpNoSize Or conSwpNoMove
    mnuAlwaysTop.Checked = False
End Sub

Private Sub mnuAlwaysTop_Click()
    If Not mnuAlwaysTop.Checked Then
        Call set_topmost
        Call RegWrite("AlwaysTop", "1")
    Else
        Call clear_topmost
        Call RegWrite("AlwaysTop", "0")
    End If
End Sub

Rem --------------------------------------------------
Rem mnuAbout_Click
Rem
Rem Brings up the dialog box for viewing the about box information
Rem
Private Sub mnuAbout_Click()
    Dim strRegCode As String
    Dim strRegName As String
    
    gState = STATE_Edit
    
    If mnuAlwaysTop.Checked Then
        Call clear_topmost
        frmAbout.Show (vbModal)
        Call set_topmost
    Else
        frmAbout.Show (vbModal)
    End If
    
    Unload frmAbout

    gState = STATE_Idle

End Sub


Rem ---------------------------------------------------
Rem mnuFormat_Click
Rem
Rem Function to bring up the format profiles
Rem
Private Sub mnuFormat_Click()

    On Error GoTo error
    
    gState = STATE_Edit
    
    If mnuAlwaysTop.Checked Then
        Call clear_topmost
        frmFormatProfiles.Show (vbModal)
        Call set_topmost
    Else
        frmFormatProfiles.Show (vbModal)
    End If
    
    Unload frmFormatProfiles
    
    gState = STATE_Idle
    Exit Sub
error:
    On Error GoTo 0
    MsgBox "An error occured editing the format profiles.", vbExclamation, "Error"
    gState = STATE_Idle
End Sub

Rem ----------------------------------------------------
Rem mnuTasks_Click
Rem
Rem Responds to the menu command to edit the task list
Rem
Private Sub mnuTasks_Click()
    On Error GoTo error
    
    gState = STATE_Edit
    
    If mnuAlwaysTop.Checked Then
        Call clear_topmost
        frmTasks.Show (vbModal)
        Call set_topmost
    Else
        frmTasks.Show (vbModal)
    End If
    
    Unload frmTasks
    
    gState = STATE_Idle
    
    Exit Sub
error:
    On Error GoTo 0
    MsgBox "An error occured editing the tasks.", vbExclamation, "Error"
    gState = STATE_Idle

End Sub


Rem ------------------------------------------------
Rem mnuPumpFile_Click
Rem
Rem Responds to the menu command to pump an individual file
Rem
Private Sub mnuPumpFile_Click()
    On Error GoTo error
    gState = STATE_Edit
    
    If mnuAlwaysTop.Checked Then
        Call clear_topmost
        frmSave.Show (vbModal)
        Call set_topmost
    Else
        frmSave.Show (vbModal)
    End If
    
    Unload frmSave
    
    gState = STATE_Idle
    Exit Sub
error:
    MsgBox "There was an error saving this file", vbExclamation, "Error"
    gState = STATE_Idle
    

End Sub

Rem ---------------------------------------------------
Rem mnuTray_Click
Rem
Rem Menu command to send the main window to the system tray
Rem
Private Sub mnuTray_Click()
    Me.Hide
End Sub

Rem ----------------------------------------------------
Rem mnuConfigure_Click
Rem
Rem Menu command to open up the main window from the system tray
Rem
Private Sub mnuConfigure_Click()
    Me.Show
    Me.SetFocus
End Sub

Rem ----------------------------------------------------
Rem mnuExit1_Click
Rem mnuExit2_Click
Rem
Rem Two menus for the exit command to exit and terminate the application
Rem
Private Sub mnuExit1_Click()
    Unload Me
    End
End Sub

Private Sub mnuExit2_Click()
    Unload Me
    End
End Sub

Rem --------------------------------------------------------
Rem hide_status
Rem show_status
Rem mnuViewStatus_Click
Rem
Rem Functions to deal with the show/hide status window
Rem
Private Sub hide_status()
    mnuViewStatus.Checked = False
    
    frame_Status.Visible = False
    Me.Width = 6440 - frame_Status.Width
    Me.Height = 2425
End Sub

Private Sub show_status()
    mnuViewStatus.Checked = True
    
    frame_Status.Visible = True
    Me.Width = 6440
    Me.Height = 2425
End Sub

Private Sub mnuViewStatus_Click()
    If mnuViewStatus.Checked = True Then
        Call hide_status
    Else
        Call show_status
    End If
    Call make_visible
End Sub

Rem -------------------------------------------------------------
Rem Timer1_Timer
Rem
Rem Timer function to run the various tasks on their schedule
Rem
Private Sub Timer1_Timer()
    Dim i As Integer
    Dim objXMLDomNodeSchedule As IXMLDOMNode
    Dim objXMLDomNodeStatus As IXMLDOMNode
    Dim dtNow As Date
    Dim dtLastRun As Date
    Dim iType As Integer
    Dim iHours, iMinutes, iMonthday, iWeekday As Integer
    Dim fRunTask As Boolean
    Dim fTasksRun As Boolean
    Dim iDisable As Integer
    Dim iConnect As Integer
    
On Error GoTo error
    fTasksRun = False
    
    Rem Only run tasks if we are idle
    If gState <> STATE_Idle Then
        Exit Sub
    End If
    
    Rem Put ourselves into the running state
    gState = STATE_Run
    
    Rem Get the current time
    dtNow = Now
    
    Rem Initialize the form from the xml state
    For i = 0 To g_objXMLDomNodeTasks.childNodes.length - 1
        
        Rem Get the status
        Set objXMLDomNodeStatus = g_objXMLDomNodeTasks.childNodes(i).childNodes(XMLITask_Status)
        
        Rem Get the schedule
        Set objXMLDomNodeSchedule = g_objXMLDomNodeTasks.childNodes(i).childNodes(XMLITask_Schedule)
        
        Rem Get the attributes of this particular task
        iType = CInt(objXMLDomNodeSchedule.childNodes(XMLISchedule_Type).Text)
        iHours = CInt(objXMLDomNodeSchedule.childNodes(XMLISchedule_Hours).Text)
        iMinutes = CInt(objXMLDomNodeSchedule.childNodes(XMLISchedule_Minutes).Text)
        iMonthday = CInt(objXMLDomNodeSchedule.childNodes(XMLISchedule_Monthday).Text)
        iWeekday = CInt(objXMLDomNodeSchedule.childNodes(XMLISchedule_Weekday).Text)
        iConnect = CInt(objXMLDomNodeSchedule.childNodes(XMLISchedule_Connect).Text)
        iDisable = CInt(objXMLDomNodeSchedule.childNodes(XMLISchedule_Disable).Text)
        iConnect = CInt(objXMLDomNodeSchedule.childNodes(XMLISchedule_Connect).Text)
        
        
        If IsNull(objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).nodeTypedValue) Then
            dtLastRun = 0
        Else
            dtLastRun = objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).nodeTypedValue
        End If
        
        Rem See if we need to run this task
        fRunTask = False
        
        Rem If the task is disabled or we ran it within the last 2 minutes don't run now
        If iDisable = 0 And DateDiff("n", dtLastRun, dtNow) > 1 Then
        
            If iType = XMLISchedule_Type_Hourly Then
                If DatePart("n", dtNow) = iMinutes Then
                    fRunTask = True
                ElseIf DateDiff("n", dtLastRun, dtNow) > 60 Then
                    fRunTask = True
                End If
            ElseIf iType = XMLISchedule_Type_Daily Then
                If DatePart("n", dtNow) = iMinutes And DatePart("h", dtNow) = iHours Then
                    fRunTask = True
                ElseIf DateDiff("d", dtLastRun, dtNow) > 1 Then
                    fRunTask = True
                End If
            ElseIf iType = XMLISchedule_Type_Weekly Then
                If DatePart("n", dtNow) = iMinutes And DatePart("h", dtNow) = iHours And DatePart("w", dtNow) = iWeekday + 1 Then
                    fRunTask = True
                ElseIf DateDiff("d", dtLastRun, dtNow) > 7 Then
                    fRunTask = True
                End If
            ElseIf iType = XMLISchedule_Type_Monthly Then
                If DatePart("n", dtNow) = iMinutes And DatePart("h", dtNow) = iHours And DatePart("d", dtNow) = iMonthday Then
                    fRunTask = True
                ElseIf DateDiff("m", dtLastRun, dtNow) > 1 Then
                    fRunTask = True
                End If
            End If
            
            If fRunTask = True Then
                StartAnimate
                Call RunTask(g_objXMLDomNodeTasks.childNodes(i), frame_Status, l_Status, p_Progress, iConnect)
                fTasksRun = True
                EndAnimate
            End If
        End If
    Next
            
    gState = STATE_Idle
    
    Exit Sub
error:
    gState = STATE_Idle
End Sub

Rem -----------------------------------------------------------------
Rem Form_MouseMove
Rem
Rem Windows callback for mouse move to process the clicks on the task
Rem icon
Rem
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Rem Event occurs when the mouse pointer is within the rectangular
    Rem boundaries of the icon in the taskbar status area.
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONUP
            If gState = STATE_Idle Then
                mnuConfigure_Click
            End If
        Case WM_LBUTTONDBLCLK
        Case WM_LBUTTONDOWN
        Case WM_RBUTTONDOWN
        Case WM_RBUTTONUP
            If gState = STATE_Idle Then
                Call PopupMenu(mnuPopupMenu)
            End If
        Case WM_RBUTTONDBLCLK
    End Select
End Sub

Rem -----------------------------------------------------
Rem StartAnimate
Rem EndAnimate
Rem
Rem Functions to start/stop the icon status into the "animated" (pumping) state
Rem
Public Sub StartAnimate()
    Me.Icon = PumpIconActive.Picture
    nid.hIcon = Me.Icon
    Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Public Sub EndAnimate()
    Me.Icon = PumpIconNormal.Picture
    nid.hIcon = Me.Icon
    Shell_NotifyIcon NIM_MODIFY, nid
End Sub

