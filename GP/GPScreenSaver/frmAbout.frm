VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Graphic Pump Saver"
   ClientHeight    =   3024
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   4248
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2081.006
   ScaleMode       =   0  'User
   ScaleWidth      =   3986.274
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "http://www.graphicpump.com"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright 2000 by Graphic Pump.  All Rights Reserved."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":0000
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Graphic Pump Screen Saver"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label l_Version 
      Caption         =   "Version"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1245
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub


Private Sub Form_Load()
    Me.Caption = "About Graphic Pump Screen Saver"
    l_Version.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
