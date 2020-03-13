Attribute VB_Name = "TrayIcon"
Const CB_OP_ADD = 768
Const CB_DST_0 = 32
Const CB_DST_1 = 48
Const FILE_GIF = 2
Const FILE_JFIF = 10
Const ROTATE_RESIZE = 1
Const k_UnlockKey_GIF = "ou82Ohmy1"
Const L_SUPPORT_GIFLZW = 1
Const k_Compression = 42


Function GetXML(xmlURL As String, directory As String)
    Dim objXMLDocument As MSXML.DOMDocument
    Dim objXMLNodeList As MSXML.IXMLDOMNodeList
    Dim objXMLDomNode As MSXML.IXMLDOMNode
    Dim objXMLDomNamedNodeMap As MSXML.IXMLDOMNamedNodeMap
    Dim i
    Dim url As String
    Dim title As String
    Dim html As String
    Dim links As String
    Dim fso As Scripting.FileSystemObject
    Dim objTextStream As Scripting.TextStream
    
    Set fso = New Scripting.FileSystemObject
    
    Set objXMLDocument = New MSXML.DOMDocument
    
    objXMLDocument.Load (xmlURL)
    
    Set objXMLNodeList = objXMLDocument.getElementsByTagName("Picture")
    
    For i = 0 To objXMLNodeList.length - 1
        url = objXMLNodeList.Item(i).Attributes.getNamedItem("url").text
        title = objXMLNodeList.Item(i).Attributes.getNamedItem("title").text
        
        html = "<html><title>" & title & "</title><body>"
        
        links = ""
        If i = 0 Then
            links = "<font class=verdana size=-1>Previous</font>&nbsp;|&nbsp;<a href=page" & i + 1 & ".htm><font class=verdana size=-1>Next</font></a>"
        ElseIf i = objXMLNodeList.length - 1 Then
            links = "<a href=page" & i - 1 & ".htm><font class=verdana size=-1>Previous</font></a>&nbsp;|&nbsp;<font class=verdana size=-1>Next</font>"
        Else
            links = "<a href=page" & i - 1 & ".htm><font class=verdana size=-1 size=-1>Previous</font></a>&nbsp;|&nbsp;<a href=page" & i + 1 & ".htm><font class=verdana size=-1>Next</font></a>"
        End If
                
        html = html & links
        html = html & "<font class=verdana size=-1>&nbsp;&nbsp;<b>" & title & "</b></font><br>"
        html = html & "<img src=file" & i & ".jpg>"
        html = html & "<br>" & links
        html = html & "</body></html>"
        
        Set objTextStream = fso.OpenTextFile(directory & "page" & i & ".htm", ForWriting, True)
        objTextStream.Write (html)
                                                  
        Call SaveImage(url, directory & "file" & i & ".jpg", 250, 250, 5, 5, 5, 5, RGB(0, 0, 0), 1, 1, False, True)
    Next
    
    MsgBox "Done"
    
End Function

Function SaveImage(refURL As String, file As String, desiredWidth As Integer, desiredHeight As Integer, topMargin, bottomMargin, leftMargin, rightMargin, colorMargin, vAlign, hAlign, grow, fRotate) As Variant
    Dim SWPage As Object
    Dim lead1 As Object
    Dim lead2 As Object
    Dim fso As Object
    Dim sourceWidth
    Dim sourceHeight
    Dim targetWidth
    Dim targetHeight
    Dim fileData
    Dim numbits 'Local variable to store the bit depth of the source image
    Dim format 'Local variable to store the format of the source image
    Dim ext
    Dim offsetTop, offsetLeft
    
    Rem Create some helper objects
    Set SWPage = CreateObject("Sidewalk.SWPage")
    Set lead1 = CreateObject("LEAD.LeadCtrl.120")
    Set lead2 = CreateObject("LEAD.LeadCtrl.120")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Call lead1.UnlockSupport(L_SUPPORT_GIFLZW, k_UnlockKey_GIF)
    Call lead2.UnlockSupport(L_SUPPORT_GIFLZW, k_UnlockKey_GIF)
    
    On Error GoTo errorhandler
    If InStr(UCase(refURL), "HTTP") = 0 Then
        If InStr(UCase(refURL), "FTP") = 0 Then
            refURL = "http://" & refURL
        End If
    End If
       
    fileData = SWPage.GetURLData(refURL, 1)
    
    Call SWPage.SaveFile(file, 1, fileData)
    
    Rem Try and load the file we got
    lead1.Load file, 0, 0, 1
    'lead1.LoadMemory fileData, 0, 0, 1
    
    format = lead1.InfoFormat
    numbits = lead1.InfoBits
    
    Rem GIF files must be saved as jpgs to work properly
    If format = FILE_GIF Or numbits <> 24 Then
        lead1.save file, FILE_JFIF, 24, 2, 0
        lead1.Load file, 0, 0, 1
    End If
  
    sourceWidth = lead1.BitmapWidth
    sourceHeight = lead1.BitmapHeight
    
    If fRotate And ((sourceWidth < sourceHeight And desiredWidth > desiredHeight) Or (sourceWidth > sourceHeight And desiredWidth < desiredHeight)) Then
        Call lead1.rotate(-9000, ROTATE_RESIZE, RGB(0, 0, 0))
        sourceWidth = lead1.BitmapWidth
        sourceHeight = lead1.BitmapHeight
    End If
            
    Rem Check if we need to change aspect ratio
    If sourceWidth < desiredWidth And sourceHeight < desiredHeight Then
        targetWidth = sourceWidth
        targetHeight = sourceHeight
    ElseIf sourceWidth / desiredWidth > sourceHeight / desiredHeight Then
        targetWidth = desiredWidth
        targetHeight = desiredWidth * sourceHeight / sourceWidth
    Else
        targetHeight = desiredHeight
        targetWidth = desiredHeight * sourceWidth / sourceHeight
    End If
    
    lead1.Size targetWidth, targetHeight, 0
        

    If grow = False Then
        desiredWidth = targetWidth
        desiredHeight = targetHeight
    End If
    lead2.CreateBitmap desiredWidth + leftMargin + rightMargin, desiredHeight + topMargin + bottomMargin, lead1.InfoBits
    lead2.Fill colorMargin
    
    If vAlign = 0 Then
        offsetTop = 0
    ElseIf vAlign = 1 Then
        offsetTop = (desiredHeight - targetHeight) / 2
    ElseIf vAlign = 2 Then
        offsetTop = (desiredHeight - targetHeight)
    End If
    
    If hAlign = 0 Then
        offsetLeft = 0
    ElseIf hAlign = 1 Then
        offsetLeft = (desiredWidth - targetWidth) / 2
    ElseIf hAlign = 2 Then
        offsetLeft = (desiredWidth - targetWidth)
    End If
    
    Call lead2.combine(leftMargin + (desiredWidth - targetWidth) / 2, topMargin + offsetTop, targetWidth, targetHeight, lead1.Bitmap, 0, 0, (CB_OP_ADD + CB_DST_0))
    lead2.save file, FILE_JFIF, numbits, k_Compression, 0
    
    SaveImage = True
    
    Exit Function

errorhandler:
    SaveImage = Err.Description
    On Error Resume Next
        If fso.FileExists(file) Then
            fso.DeleteFile (file)
        End If
    Exit Function
    On Error GoTo 0
End Function

