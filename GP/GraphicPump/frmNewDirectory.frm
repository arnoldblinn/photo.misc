VERSION 5.00
Begin VB.Form frmNewDirectory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directory"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox f_NewDirectory 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label l_Directory 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Type in the name of a sub directory to create.  The current directory is:"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmNewDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_Cancel_Click()
    l_Directory.Caption = ""
End Sub

Private Sub b_OK_Click()
    Dim fso As Scripting.FileSystemObject
    Dim strNewDirectory As String
    Dim errNum As Integer
    
    Set fso = New Scripting.FileSystemObject
    
    If f_NewDirectory.Text <> "" Then
        If Right(l_Directory.Caption, 1) = "\" Then
            strNewDirectory = l_Directory.Caption & frmNewDirectory.f_NewDirectory.Text
        Else
            strNewDirectory = l_Directory.Caption & "\" & frmNewDirectory.f_NewDirectory.Text
        End If
        
        On Error Resume Next
        fso.CreateFolder (strNewDirectory)
        errNum = Err.Number
        On Error GoTo 0
        If errNum <> 0 Then
            MsgBox "Unable to create directory.", vbExclamation, "Error"
            Exit Sub
        End If
        
    End If

    l_Directory.Caption = strNewDirectory
    Me.Hide
End Sub

Private Sub f_NewDirectory_GotFocus()
    f_NewDirectory.SelStart = 0
    f_NewDirectory.SelLength = Len(f_NewDirectory.Text)
End Sub

