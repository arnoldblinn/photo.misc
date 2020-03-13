VERSION 5.00
Object = "{0C155360-3CD1-11D0-B17A-E18E3EAC3833}#1.0#0"; "SI_COMM.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin SI_COMMLib.SI_COMM SI_COMM1 
      Left            =   4800
      Top             =   2640
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
   Begin VB.CommandButton Command4 
      Caption         =   "Put c:\fun0.jpg"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Execute"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   5775
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As clsDigiFrame

    
Private Sub Command1_Click()
    Label1.Caption = Replace(f.Command(Text1.Text), vbCr, "")
    Label1.Refresh
End Sub

Private Sub Command2_Click()
    
    f.Connect
End Sub

Private Sub Command3_Click()
    f.Disconnect
End Sub

Private Sub Command4_Click()
    Label1.Caption = f.PutFile("file0000.jpg", "c:\fun0.jpg", True)
End Sub

Private Sub Form_Load()
    
    Set f = New clsDigiFrame
    
    f.Port = 0
    f.Card = 0
                
End Sub
