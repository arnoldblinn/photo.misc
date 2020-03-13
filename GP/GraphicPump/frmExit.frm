VERSION 5.00
Begin VB.Form frmExit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Close"
   ClientHeight    =   1596
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1596
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton b_Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton b_Minimize 
      Caption         =   "Minimize"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Do you wish to minimize the application and leave it running on the system tray, or do you wish to exit the application?"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: frmExit.frm
Rem
Rem Description:
Rem     Contains code for the exit dialog for the application.
Rem
Rem -------------------------------------------------------------

Dim gResult As Integer

Public Property Get Result() As Integer
    Result = gResult
End Property

Private Sub b_Cancel_Click()
    gResult = 2
    Me.Hide
End Sub

Private Sub b_Exit_Click()
    gResult = 1
    Me.Hide
End Sub

Private Sub b_Minimize_Click()
    gResult = 0
    Me.Hide
End Sub

