Attribute VB_Name = "Registry"
Option Explicit

Public Function RegWrite(strKey As String, strValue As String)
    Dim WshShell As Object
    Dim strRoot As String
    
    strRoot = "HKEY_LOCAL_MACHINE\SOFTWARE\GraphicPump\"

    Set WshShell = CreateObject("WScript.Shell")
    
    Call WshShell.RegWrite(strRoot & strKey, strValue)

End Function

Public Function RegReadIntDefault(strKey As String, iDefault As Integer) As Integer

On Error GoTo error
    Dim WshShell As Object
    Dim strRoot As String
    
    strRoot = "HKEY_LOCAL_MACHINE\SOFTWARE\GraphicPump\"
    Set WshShell = CreateObject("WScript.Shell")
    
    RegReadIntDefault = CInt(WshShell.RegRead(strRoot & strKey))
    
    Exit Function
error:
    RegReadIntDefault = iDefault

End Function


Public Function RegReadStringDefault(strKey As String, strDefault As String) As String

On Error GoTo error
    Dim WshShell As Object
    Dim strRoot As String
    
    strRoot = "HKEY_LOCAL_MACHINE\SOFTWARE\GraphicPump\"
    Set WshShell = CreateObject("WScript.Shell")
    
    RegReadStringDefault = WshShell.RegRead(strRoot & strKey)
    
    Exit Function
error:
    RegReadStringDefault = strDefault

End Function
