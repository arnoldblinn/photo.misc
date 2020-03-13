Attribute VB_Name = "Util"
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: util.bas
Rem
Rem Description:
Rem     Contains general utility functions
Rem

Public Function HTMLEncode(ByVal strInput As String) As String
    HTMLEncode = strInput
    
    HTMLEncode = Replace(HTMLEncode, "&", "&amp;")
    HTMLEncode = Replace(HTMLEncode, ">", "&gt;")
    HTMLEncode = Replace(HTMLEncode, "<", "&lt;")

End Function

Public Function PadNumber(i As Integer) As String
    If i = 0 Then
        PadNumber = "0000"
    ElseIf i < 10 Then
        PadNumber = "000" & i
    ElseIf i < 100 Then
        PadNumber = "00" & i
    ElseIf i < 1000 Then
        PadNumber = "0" & i
    Else
        PadNumber = i
    End If
End Function

Public Function FullPathToFile(strFullPath As String) As String
    FullPathToFile = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Public Function FullPathToDirectory(strFullPath As String) As String
    FullPathToDirectory = Left(strFullPath, InStrRev(strFullPath, "\"))
End Function

Public Function DirectoryFileToFullPath(strDirectory As String, strFile As String) As String
    If Right(strDirectory, 1) = "\" Then
        DirectoryFileToFullPath = strDirectory & strFile
    Else
        DirectoryFileToFullPath = strDirectory & "\" & strFile
    End If
End Function


