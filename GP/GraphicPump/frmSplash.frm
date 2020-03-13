VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5688
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5688
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox p_gp 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   -120
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   2772
      ScaleWidth      =   5772
      TabIndex        =   0
      Top             =   -120
      Width           =   5772
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2040
      Top             =   2760
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Delay(s As Integer)
    Dim dtStart As Date
    
    dtStart = Now
    While DateDiff("s", dtStart, Now) < s
    Wend

End Sub

Private Sub Form_Load()
        
    p_gp.Top = 0
    p_gp.Left = 0
    p_gp.Width = 503 * Screen.TwipsPerPixelX
    p_gp.Height = 231 * Screen.TwipsPerPixelY
        
    Me.Width = p_gp.Width
    Me.Height = p_gp.Height
    Me.Left = (Screen.Width - p_gp.Width) / 2
    Me.Top = (Screen.Height - p_gp.Height) / 2
End Sub

Private Sub Timer1_Timer()
    frmSplash.Hide
End Sub
