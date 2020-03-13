Attribute VB_Name = "MainEntry"
Option Explicit

Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: Main.bas
Rem
Rem Description:
Rem     Contains code for the main entry point for the application.
Rem

Rem Global settings
Global g_strExeDir As String

Public Sub Main()
    Dim strCmd As String
    Dim fSplash As Boolean
    Dim fShow As Boolean
    Dim aArgs() As String
    Dim j As Integer
    
On Error GoTo error

    Rem Initialize the global settings
    g_strExeDir = CurDir & "\"
    
    Rem Initialize local settings
    fSplash = True
    fShow = False
    
    Rem Get the command line
    strCmd = UCase(Trim(Command))
    
    Rem Get the arguments
    aArgs = Split(strCmd, "/")
    
    Rem Process the arguments
    Dim l As Integer
    For j = 0 To UBound(aArgs)
        Dim strTemp As String
        strTemp = aArgs(j)
        If strTemp <> "" Then
        
            If Left(strTemp, 3) = "VB " Then
                g_strExeDir = "c:\personal\GraphicPump\"
            ElseIf Left(strTemp, 2) = "S " Then
                fSplash = False
            ElseIf Left(strTemp, 2) = "O " Then
                fShow = True
            End If
        End If
    Next
     
    ChDir (g_strExeDir)
    
    Rem We should only ever run one task window
    If App.PrevInstance Then
        Exit Sub
    End If
    
    Rem Get the schema, and load the various global states
    g_strSchema = "x-schema:" & g_strExeDir & "pumpschema.xml"
    Set g_objXMLDomNodeTasks = LoadTasks()
    Set g_objXMLDomNodeFormatProfiles = LoadFormatProfiles()
    
    Rem Show the startup screen unless told not to
    If fSplash = True Then
        frmSplash.Show (vbModal)
        Unload frmSplash
    End If
        
    Load frmMain
    
    Rem Open the task window automatically if requested
    If fShow = True Then
        frmMain.Show
    End If
    
    Exit Sub
    
error:

    MsgBox "An internal error occured.", vbCritical, "Error"
    
End Sub
