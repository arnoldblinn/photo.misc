VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   3420
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8052
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox p_gp 
      BorderStyle     =   0  'None
      Height          =   2052
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   171
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   5
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox p_vc 
      BorderStyle     =   0  'None
      Height          =   2052
      Left            =   120
      Picture         =   "frmAbout.frx":19206
      ScaleHeight     =   2052
      ScaleWidth      =   3372
      TabIndex        =   6
      Top             =   240
      Width           =   3372
   End
   Begin VB.Label Label2 
      Caption         =   "http://www.graphicpump.com"
      Height          =   252
      Left            =   3600
      TabIndex        =   8
      Top             =   840
      Width           =   3612
   End
   Begin VB.Label l_About2 
      Caption         =   "For more information, to upgrade to a newer version, or to suggest new features visit http://www.graphicpump.com"
      Height          =   612
      Left            =   3600
      TabIndex        =   7
      Top             =   2160
      Width           =   4212
   End
   Begin VB.Label l_Version 
      Caption         =   "Version"
      Height          =   228
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   3972
   End
   Begin VB.Label l_Title 
      Caption         =   "Graphic Pump"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3576
      TabIndex        =   3
      Top             =   120
      Width           =   3996
   End
   Begin VB.Label l_About 
      Caption         =   $"frmAbout.frx":31FCA
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   4248
   End
   Begin VB.Label l_Copyright 
      Caption         =   "Copyright 2000-20001 by Graphic Pump. All rights reserved"
      Height          =   252
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   4212
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_Cancel_Click()
    Me.Hide
End Sub

Private Sub b_OK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    l_Version.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

