Attribute VB_Name = "WebSite"
Option Explicit
Rem ------------------------------------------------------------------------
Rem GetPageHTML
Rem
Rem Gets HTML for a page
Rem
Function GetPageHTML(ByVal strIndexName As String, ByVal strPageName As String, ByVal strPrevPageName As String, ByVal strNextPageName As String, ByVal strImageName As String, ByVal strImageFileName As String) As String
    Dim html As String
    Dim links As String
        
    html = "<html><title>" & HTMLEncode(strImageName) & "</title><body>"
        
    links = ""
    If strIndexName <> "" Then
        links = links & "<a href=" & strIndexName & "><font class=verdana size=-1>Index</font></a>"
    Else
        links = links & "<font class=verdana size=-1>Index</font>"
    End If
    links = links & "&nbsp;|&nbsp;"
    If strPrevPageName <> "" Then
        links = links & "<a href=" & strPrevPageName & "><font class=verdana size=-1>Previous</font></a>"
    Else
        links = links & "<font class=verdana size=-1>Previous</font>"
    End If
    links = links & "&nbsp;|&nbsp;"
    If strNextPageName <> "" Then
        links = links & "<a href=" & strNextPageName & "><font class=verdana size=-1>Next</font></a>"
    Else
        links = links & "<font class=verdana size=-1>Next</font>"
    End If
                   
    html = html & links
    html = html & "<font class=verdana size=-1>&nbsp;&nbsp;<b>" & HTMLEncode(strImageName) & "</b></font><br>"
    html = html & "<img src=" & strImageFileName & ">"
    html = html & "<br>" & links
    html = html & "</body></html>"
                  
    
    GetPageHTML = html
    
End Function

Function SaveText(strFullFile As String, strText As String, fso As Scripting.FileSystemObject) As Boolean
    Dim objTextStream As Scripting.TextStream
    
On Error GoTo 0
    SaveText = False
    
    Set objTextStream = fso.OpenTextFile(strFullFile, ForWriting, True)
    objTextStream.Write (strText)
    
    SaveText = True
    
error:
    On Error GoTo 0
End Function

Rem ------------------------------------------------------------------------
Rem MakeWebSite
Rem
Rem Makes a web site from images in the directory
Rem
Function MakeWebSite(strDirectory As String, fDeleteOld As Boolean, strAlbumName As String, strPageTemplate As String, strIndexTemplate As String) As String
    Dim strStatus As String
    Dim folder As Scripting.folder
    Dim strPageHTML As String
    Dim strIndexHTML As String
    Dim fso As Scripting.FileSystemObject
    Dim iCount As Integer

On Error GoTo error
    Set fso = CreateObject("scripting.filesystemobject")
    Set folder = fso.GetFolder(strDirectory)
    
    If fDeleteOld Then
        
        Dim file As Scripting.file
        Dim strLeftPage, strRightPage As String
        Dim iLocation As Integer
                        
        iLocation = InStr(1, strPageTemplate, "%i%")
        strLeftPage = UCase(Left(strPageTemplate, iLocation - 1))
        strRightPage = UCase(Right(strPageTemplate, Len(strPageTemplate) - (iLocation + 2)))
                    
        
        For Each file In folder.Files
            If UCase(Left(file.name, Len(strLeftPage))) = strLeftPage And UCase(Right(file.name, Len(strRightPage))) = strRightPage Then
                Call file.Delete(True)
            ElseIf UCase(strIndexTemplate) = UCase(file.name) Then
                Call file.Delete(True)
            End If
        Next
    End If
    
    iCount = 0
    For Each file In folder.Files
        If Right(file.name, 3) = "jpg" Then
            iCount = iCount + 1
        End If
    Next
    
    strIndexHTML = "<html><title>" & HTMLEncode(strAlbumName) & "</title><body><h2><font class=verdana>" & HTMLEncode(strAlbumName) & "</font></h2><ul>"
    
    Dim i As Integer
    Dim strPageName, strPrevPageName, strNextPageName As String
    Dim strImageName, strImageFileName As String
    i = 0
    For Each file In folder.Files
        If Right(file.name, 3) = "jpg" Then
            strImageFileName = file.name
            strImageName = "File: " & file.name
            
            strPageName = Replace(strPageTemplate, "%i%", PadNumber(i))
            
            If i > 0 Then
                strPrevPageName = Replace(strPageTemplate, "%i%", PadNumber(i - 1))
            Else
                strPrevPageName = ""
            End If
            
            If i < iCount - 1 Then
                strNextPageName = Replace(strPageTemplate, "%i%", PadNumber(i + 1))
            Else
                strNextPageName = ""
            End If
            
            strPageHTML = GetPageHTML(strIndexTemplate, strPageName, strPrevPageName, strNextPageName, strImageName, strImageFileName)
            If SaveText(strDirectory & strPageName, strPageHTML, fso) = False Then
                strStatus = "Could not save " & strPageName
                GoTo error
            End If
            strIndexHTML = strIndexHTML & "<li><a href=" & strPageName & "><font class=verdana size=-1>" & HTMLEncode(strImageName) & "</a></font></li>"
            i = i + 1
        End If
    Next
    
    strIndexHTML = strIndexHTML & "</ul></body></html>"
    If SaveText(strDirectory & strIndexTemplate, strIndexHTML, fso) = False Then
        strStatus = "Could not save index.htm"
        GoTo error
    End If
    
done:
        
    MakeWebSite = ""
    
    Exit Function
    
error:
    MakeWebSite = strStatus
        
End Function
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        


