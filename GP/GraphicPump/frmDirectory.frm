VERSION 5.00
Begin VB.Form frmDirectory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "frmDirectory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox f_Path 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton b_New 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Directory"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Drive"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   660
      Width           =   975
   End
End
Attribute VB_Name = "frmDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Property Let Path(strValue As String)
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    
    If fso.FolderExists(strValue) Then
        f_Path = strValue
        Dir1.Path = strValue
    End If
End Property

Public Property Get Path() As String
    Path = f_Path.Text
End Property

Private Sub b_New_Click()
    
    frmNewDirectory.l_Directory.Caption = Dir1.Path
    frmNewDirectory.Show (vbModal)
    If frmNewDirectory.l_Directory.Caption <> "" Then
        Dir1.Path = frmNewDirectory.l_Directory.Caption
    End If
    
    Unload frmNewDirectory
    
End Sub

Private Sub b_OK_Click()
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    
    If fso.FolderExists(f_Path.Text) = False Then
        On Error Resume Next
        fso.CreateFolder (f_Path.Text)
        errNum = Err.Number
        On Error GoTo 0
        If errNum <> 0 Then
            MsgBox "Invalid directory", vbExclamation, "Error"
            Exit Sub
        End If
        
        ChDir (f_Path.Text)
    End If
    Me.Hide
End Sub

Private Sub b_Cancel_Click()
    f_Path.Text = ""
    Me.Hide
End Sub

Private Sub Dir1_Change()
    f_Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo done
    Dir1.Path = Drive1.Drive
done:
End Sub

Private Sub f_Path_GotFocus()
    f_Path.SelStart = 0
    f_Path.SelLength = Len(f_Path.Text)
End Sub

Private Sub Form_Load()
    f_Path = CurDir
    Dir1.Path = CurDir
End Sub

