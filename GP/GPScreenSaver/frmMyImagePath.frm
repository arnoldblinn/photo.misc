VERSION 5.00
Begin VB.Form frmMyImagePath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image Path"
   ClientHeight    =   3150
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.OptionButton o_All 
      Caption         =   "All jpg files in this directory"
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.OptionButton o_File 
      Caption         =   "Selected file only"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   2880
      TabIndex        =   1
      Top             =   1485
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmMyImagePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrImagePath As String

Private Sub b_OK_Click()
    If o_All.Value = True Then
        If Dir1.Path = "" Or File1.ListCount = 0 Then
            MsgBox "Please select a directory with .jpg files"
            Exit Sub
        End If
        
        If Right(Dir1.Path, 1) = "\" Then
            mstrImagePath = Dir1.Path
        Else
            mstrImagePath = Dir1.Path & "\"
        End If
    Else
        If File1.ListIndex = -1 Then
            MsgBox "Please select a file", vbApplicationModal, "Error"
            Exit Sub
        End If
        If Right(Dir1.Path, 1) = "\" Then
            mstrImagePath = Dir1.Path & File1.FileName
        Else
            mstrImagePath = Dir1.Path & "\" & File1.FileName
        End If
    End If
    
    frmMyImagePath.Hide
End Sub

Private Sub b_Cancel_Click()
    frmMyImagePath.Hide
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


Public Property Let ImagePath(strValue As String)
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    If fso.FolderExists(strValue) = False And fso.FileExists(strValue) = False Then
        mstrImagePath = ""
    Else
        mstrImagePath = strValue
    End If
End Property

Public Property Get ImagePath() As String
    ImagePath = mstrImagePath
End Property

Private Sub Form_Initialize()
    mstrImagePath = ""
End Sub

Private Sub Form_Activate()
    Dim p As Long
    Dim dirPath As String
    Dim fName As String
    
    File1.Pattern = "*.jpg"
    
    If mstrImagePath = "" Then
        Drive1.Drive = Left(CurDir, 2)
        Dir1.Path = CurDir
        o_All.Value = True
        File1.Enabled = False
    Else
        Drive1.Drive = Left(mstrImagePath, 2)
    
        If Right(mstrImagePath, 4) = ".jpg" Then
        
            p = InStrRev(mstrImagePath, "\")
            dirPath = Left(mstrImagePath, p)
            fName = Right(mstrImagePath, Len(mstrImagePath) - p)
        
            Dir1.Path = dirPath
            File1.Path = dirPath
            o_All.Value = False
            o_File.Value = True
            File1.Enabled = True
        Else
            Dir1.Path = mstrImagePath
            o_All.Value = True
            File1.Enabled = False
        End If
    End If
End Sub

Private Sub o_All_Click()
    If o_All.Value = True Then
        File1.Enabled = False
        File1.ListIndex = -1
    Else
        File1.Enabled = True
    End If
End Sub

Private Sub o_File_Click()
    If o_All.Value = True Then
        File1.ListIndex = -1
        File1.Enabled = False
    Else
        File1.Enabled = True
    End If
End Sub

