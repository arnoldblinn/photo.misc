VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   5016
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8052
   ControlBox      =   0   'False
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox p_gp 
      BorderStyle     =   0  'None
      Height          =   2052
      Left            =   120
      Picture         =   "frmRegister.frx":0442
      ScaleHeight     =   171
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   11
      Top             =   240
      Width           =   3255
   End
   Begin VB.TextBox f_Name 
      Height          =   288
      Left            =   3720
      TabIndex        =   0
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox f_Code 
      Height          =   288
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   3495
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Register Later"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.PictureBox p_vc 
      BorderStyle     =   0  'None
      Height          =   2052
      Left            =   120
      Picture         =   "frmRegister.frx":19206
      ScaleHeight     =   2052
      ScaleWidth      =   3372
      TabIndex        =   12
      Top             =   240
      Width           =   3372
   End
   Begin VB.Label Label2 
      Caption         =   "http://www.graphicpump.com"
      Height          =   252
      Left            =   3600
      TabIndex        =   14
      Top             =   840
      Width           =   3612
   End
   Begin VB.Label l_About2 
      Caption         =   "Label2"
      Height          =   972
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   7692
   End
   Begin VB.Label l_Version 
      Caption         =   "Version"
      Height          =   228
      Left            =   3600
      TabIndex        =   10
      Top             =   360
      Width           =   3972
   End
   Begin VB.Label l_Title 
      Caption         =   "Graphic Pump"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3576
      TabIndex        =   9
      Top             =   120
      Width           =   3996
   End
   Begin VB.Label l_About 
      Caption         =   $"frmRegister.frx":31FCA
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   3600
      TabIndex        =   8
      Top             =   1440
      Width           =   4248
   End
   Begin VB.Label l_Copyright 
      Caption         =   "Copyright 2000-20001 by Graphic Pump. All rights reserved"
      Height          =   252
      Left            =   3600
      TabIndex        =   7
      Top             =   600
      Width           =   4212
   End
   Begin VB.Label l_Register 
      Caption         =   "This is an unregistered copy. To register, visit our web site at http://www.graphicpump.com"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   7812
   End
   Begin VB.Label l_Name 
      Caption         =   "Name"
      Height          =   252
      Left            =   2400
      TabIndex        =   5
      Top             =   3720
      Width           =   852
   End
   Begin VB.Label l_Code 
      Caption         =   "Code"
      Height          =   252
      Left            =   2400
      TabIndex        =   3
      Top             =   4200
      Width           =   852
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gResult As Boolean
Dim gstrName As String
Dim gstrCode As String
Dim gfRegistered As Boolean

Public Property Get Result() As Boolean
    Result = gResult
End Property

Public Property Get RegName() As String
    RegName = gstrName
End Property

Public Property Let RegName(strNewValue As String)
    gstrName = strNewValue
End Property

Public Property Get RegCode() As String
    RegCode = gstrCode
End Property

Public Property Let RegCode(strNewValue As String)
    gstrCode = strNewValue
End Property

Private Sub b_Cancel_Click()
    gResult = False
    Me.Hide
End Sub

Private Sub b_OK_Click()
    If gfRegistered = True Then
        gResult = False
        Me.Hide
    ElseIf Len(f_Name.Text) < 3 Or ValidateKey(f_Name.Text, f_Code.Text) = False Then
        MsgBox "Invalid registration name/code.", vbExclamation, "Error"
    Else
        gstrName = f_Name.Text
        gstrCode = f_Code.Text
        gResult = True
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    
    If ValidateKey(gstrName, gstrCode) = True Then
        gfRegistered = True
        f_Name.Visible = False
        l_Name.Visible = False
        f_Code.Visible = False
        l_Code.Visible = False
        l_About.Caption = "The Graphic Pump automatically moves images from the Internet or a computer to a Screen Saver, a Pocket PC, or a Digital Picture Frame."
        l_About2.Caption = "This copy is registered to " & gstrName
        l_Register.Caption = "For more information or to register more copies visit our web site at http://www.graphicpump.com"
        b_Cancel.Visible = False
        b_OK.Left = (Me.ScaleWidth - b_OK.Width) / 2
    Else
        gfRegistered = False
        f_Name.Visible = True
        f_Code.Visible = True
        l_Register.Caption = "For more information or to register this copy visit our web site at http://www.graphicpump.com"
        
        Rem Show the right bitmap
        If giVideoChip = 1 Then
            p_vc.Visible = True
            p_gp.Visible = False
            l_About.Caption = "abc"
            l_Title.Caption = "VideoChip Picture Generator"
            l_About.Caption = "The Video Chip Picture Generator automatically moves images from the Internet or a computer to your VideoChip Photo Wallet."
            l_About2.Caption = "The VideoChip Picture Generator is powered by the Graphic Pump application, and is fully functional for the VideoChip Photo Wallet.  To take advantage of more advanced features such as album file support, scheduled tasks, advanced formatting, and other output targets you must upgrade to the Graphic Pump by registering your copy."
        Else
            p_vc.Visible = False
            p_gp.Visible = True
            l_About.Caption = "The Graphic Pump automatically moves images from the Internet or a computer to a Screen Saver, a Pocket PC, or a Digital Picture Frame."
            l_About2.Caption = "This version is fully functional, but places a small text message in the graphic images pumped by the application.  To remove this message, you must register your copy."
        End If
    End If
End Sub

Private Sub Form_Load()
    l_Version.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

