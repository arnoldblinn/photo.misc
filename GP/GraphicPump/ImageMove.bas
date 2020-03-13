Attribute VB_Name = "ImageMove"
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: ImageMove.bas
Rem
Rem Description:
Rem     Contains code for executing the graphic pump job.
Rem
Rem -------------------------------------------------------------

Rem Constants for the lead toolkit
Const CB_OP_ADD = 768
Const CB_DST_0 = 32
Const CB_DST_1 = 48
Const FILE_JFIF = 10
Const FILE_LEAD1JFIF = 21
Const ROTATE_RESIZE = 1
Const k_Compression = 42
Const FILE_GIF = 3
Const L_SUPPORT_GIFLZW = 1
Const k_UnlockKey_GIF_V12 = "sg8Z2XkjL"
Const RESIZE_RESAMPLE = 2

Public Function RunTask(objXMLDomNodeTask As IXMLDOMNode, statusFrame As Frame, label As label, progress As ProgressBar, iConnect As Integer)
    Dim objXMLDomNodeFormatProfileSettings As IXMLDOMNode
    Dim objXMLDomNodeDestination As IXMLDOMNode
    Dim objXMLDomNodeStatus As IXMLDOMNode
    Dim objXMLDomNodeSource As IXMLDOMNode
    Dim strStatus As String
    Dim objXMLDomNodePictureList As MSXML.IXMLDOMNodeList
    Dim i As Integer
    Dim lDestinationType As Long
    Dim strDestinationDirectory As String
    Dim lSourceType As Long
    Dim strAlbumName As String
    Dim fso As Scripting.FileSystemObject
    Dim objDigiFrame As clsDigiFrame
    Dim strSource As String, strDestination As String
    Dim fFrameConnect, fTempDirCreated As Boolean
    Dim strImageName As String
    Dim fOnline As Boolean
    Dim fCanConnect As Boolean
    Dim fTaskConnected As Boolean
    Dim lConnectedState As Long
    Dim strFileTemplate As String, strPageTemplate As String, strIndexTemplate As String
    Dim fDirectoryDelete As Boolean
    Dim iNumSteps As Integer
    Dim strPageNum As String
    Dim iPort As Integer
    Dim iCard As Integer
    Dim strSourceImageName As String

On Error GoTo error
    Rem Get interesting data about the task we are running
    Set objXMLDomNodeStatus = objXMLDomNodeTask.childNodes(XMLITask_Status)
    Set objXMLDomNodeSource = objXMLDomNodeTask.childNodes(XMLITask_Source)
    Set objXMLDomNodeFormatProfileSettings = objXMLDomNodeTask.childNodes(XMLITask_Format).childNodes(XMLIFormat_Settings)
    Set objXMLDomNodeDestination = objXMLDomNodeTask.childNodes(XMLITask_Destination)
        
    Rem Important local variables for cleanup later
    fFrameConnect = False
    fTempDirCreated = False
    fTaskConnected = False
    strStatus = ""
    
    Rem Create a file system object for use later
    Set fso = New Scripting.FileSystemObject
    
    Rem Create a temporary directory for saving files from the web into
    If fso.FolderExists(g_strExeDir & "temp") Then
        Call fso.DeleteFolder(g_strExeDir & "temp", True)
    End If
    Call fso.CreateFolder(g_strExeDir & "temp")
    fTempDirCreated = True
    
    Rem Determine if we are online
    lConnectedState = 0
    fCanConnect = True
    fOnline = InternetGetConnectedState(lConnectedState, 0)
    If fOnline = False Then
        fOnline = False
        If (lConnectedState And INTERNET_CONNECTION_MODEM) = False Then
            fCanConnect = False
        End If
    End If
                
    Rem Initialize feedback values
    frmMain.StartAnimate
    statusFrame.Caption = "Execution Status: " & objXMLDomNodeTask.childNodes(XMLITask_Name).Text
    statusFrame.Refresh
    label.Visible = True
    Screen.MousePointer = 11
    label.Caption = "Getting album file"
    label.Refresh
    progress.value = 0
            
    Rem Get the source
    lSourceType = CLng(objXMLDomNodeSource.childNodes(XMLISource_Type).Text)

    Rem Get the list of pictures to process
    If lSourceType = XMLISource_Type_Album Then

        Rem The picture list is simply in the passed in configuration
        Set objXMLDomNodePictureList = objXMLDomNodeSource.childNodes(XMLISource_Album).childNodes(XMLIAlbum_PictureList).childNodes

        strAlbumName = objXMLDomNodeSource.childNodes(XMLISource_Album).childNodes(XMLIAlbum_Name).Text

    ElseIf lSourceType = XMLISource_Type_AlbumURI Then
        Dim objXMLDocument As MSXML.DOMDocument
        Dim strXMLFile As String
    
        strXMLFile = objXMLDomNodeSource.childNodes(XMLISource_AlbumURI).Text
        strXMLFile = Replace(strXMLFile, "%sd%", g_strExeDir)
    
        Rem See if we need to go online
        If fOnline = False And (Left(strXMLFile, 5) = "http:" Or Left(strXMLFile, 4) = "ftp:") Then
            If iConnect = 1 Then
                If fCanConnect Then
                    If InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0) = False Then
                        strStatus = "Could not connect to Internet"
                        GoTo done
                    End If
                Else
                    strStatus = "Cannot connect to Internet."
                    GoTo done
                End If
            Else
                strStatus = "Not connected to Internet."
                GoTo done
            End If
        
            fOnline = True
            fTaskConnected = True
        End If
    
        Rem Get the data
        strStatus = "Error getting album file: " & strXMLFile
        If Left(strXMLFile, 5) = "http:" Or Left(strXMLFile, 4) = "ftp:" Then
            Dim bData() As Byte
            Dim intFile As Integer
        
            Rem Get the URL as binary data
            intFile = FreeFile()
            bData() = frmControls.Inet1.OpenURL(strXMLFile, icByteArray)

            Rem Save to the file
            Open g_strExeDir & "temp\temp.gpa" For Binary Access Write As #intFile
            Put #intFile, , bData()
            Close #intFile
            strXMLFile = g_strExeDir & "temp\temp.gpa"
        End If
    
        Rem Get the picture list
        Set objXMLDocument = LoadDOMFromFile(strXMLFile, "Album", "1")
        If objXMLDocument Is Nothing Then
            GoTo done
        End If
    
        strStatus = ""
        
        Set objXMLDomNodePictureList = objXMLDocument.getElementsByTagName("Picture")
        strAlbumName = objXMLDocument.childNodes(0).childNodes(XMLIAlbum_Name).Text
    End If

    Rem Get information about the destination
    lDestinationType = CLng(objXMLDomNodeDestination.childNodes(XMLIDestination_Type).Text)
    fDirectoryDelete = CBool(objXMLDomNodeDestination.childNodes(XMLIDestination_DirectoryDelete).Text)
    strFileTemplate = objXMLDomNodeDestination.childNodes(XMLIDestination_FileTemplate).Text
    
    Rem Progress counter is two * the number of picture + connect to frame + disconnect from frame + shutdown + init
    iNumSteps = 1 ' Initialize
    iNumSteps = iNumSteps + objXMLDomNodePictureList.length ' Fetch of each picture
    If lDestinationType = XMLIDestination_Type_DigiFrame Then
        iNumSteps = iNumSteps + 1 ' Connect to frame
        iNumSteps = iNumSteps + objXMLDomNodePictureList.length ' Send each picture to frame
    Else
        iNumSteps = iNumSteps + 1 ' Copy of images
    End If
    iNumSteps = iNumSteps + 1 ' Deinitialize
    
    Rem Set the number of steps, and indicate where we are
    progress.max = iNumSteps
    progress.value = 1
    
    Rem Get each picture and save to a temporary directory
    For i = 0 To objXMLDomNodePictureList.length - 1
        
        Rem Get the name of the image
        strImageName = objXMLDomNodePictureList.Item(i).childNodes(XMLIPicture_Name).Text
        strSource = objXMLDomNodePictureList.Item(i).childNodes(XMLIPicture_URI).Text
        strSource = Replace(strSource, "%sd%", g_strExeDir)
                
        Rem Update our progress counter
        label.Caption = "Getting Image """ & strImageName & ": " & strSource & """"
        label.Refresh
        progress.value = progress.value + 1
        
        Rem See if we need to go online
        If fOnline = False And (Left(strSource, 5) = "http:" Or Left(strSource, 4) = "ftp:") Then
            If iConnect = 1 Then
                If fCanConnect Then
                    If InternetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0) = False Then
                        strStatus = "Could not connect to Internet."
                        GoTo done
                    End If
                Else
                    strStatus = "Unabled to connect to Internet."
                End If
            Else
                strStatus = "Not connected to Internet."
                GoTo done
            End If
            
            fOnline = True
            fTaskConnected = True
        End If
        
        Rem Save the image in the right format
        If SaveImageXML(g_strExeDir & "temp\", "file" & i & ".jpg", strSource, objXMLDomNodeFormatProfileSettings) = False Then
            strStatus = "Problem processing image " & strSource
            GoTo done
        End If
    Next
    
    Rem Hang up if the task forced the connection
    If fOnline And fTaskConnected Then
        InternetAutodialHangup (0)
        fOnline = False
        fTaskConnected = False
    End If
        
    Rem If the destination is the digiframe then open the connection
    If lDestinationType = XMLIDestination_Type_DigiFrame Then
    
        Rem Update our progress counter
        label.Caption = "Connecting to Digi-Frame"
        label.Refresh
        progress.value = progress.value + 1
       
        Rem Create the frame object and initialize
        Set objDigiFrame = New clsDigiFrame
                    
        Rem Set up the port and card to write to
        iPort = CInt(objXMLDomNodeDestination.childNodes(XMLIDestination_DigiFramePort).Text)
        iCard = CInt(objXMLDomNodeDestination.childNodes(XMLIDestination_DigiFrameMedia).Text)
        objDigiFrame.Port = iPort
        objDigiFrame.Card = iCard
        
        fFrameConnect = objDigiFrame.Connect
        If fFrameConnect = False Then
            strStatus = "Could not connect to Digi-Frame."
            GoTo done
        End If
    Else
        
        label.Caption = "Copying files to directory."
        label.Refresh
        progress.value = progress.value + 1
        
        strDestinationDirectory = objXMLDomNodeDestination.childNodes(XMLIDestination_Directory).Text
        strDestinationDirectory = Replace(strDestinationDirectory, "%sd%", g_strExeDir)
        
        Rem Verify that there is a destination directory
        If fso.FolderExists(strDestinationDirectory) = False Then
            strStatus = "Unable to access directory """ & strDestinationDirectory & """"
            GoTo done
        End If
        
        If fDirectoryDelete Then
            Dim folder As Scripting.folder
            Dim file As Scripting.file
            Dim strLeftFile, strRightFile As String
            Dim iLocation As Integer
            
            iLocation = InStr(1, strFileTemplate, "%i%")
            strLeftFile = UCase(Left(strFileTemplate, iLocation - 1))
            strRightFile = UCase(Right(strFileTemplate, Len(strFileTemplate) - (iLocation + 2)))
                        
            Set folder = fso.GetFolder(strDestinationDirectory)
            For Each file In folder.Files
                If UCase(Left(file.Name, Len(strLeftFile))) = strLeftFile And UCase(Right(file.Name, Len(strRightFile))) = strRightFile Then
                    Call file.Delete(True)
                End If
            Next
        End If
        
        If Right(strDestinationDirectory, 1) <> "\" Then
            strDestinationDirectory = strDestinationDirectory & "\"
        End If
    End If
            
    For i = 0 To objXMLDomNodePictureList.length - 1
    
        strPageNum = PadNumber(i)
    
        strImageName = objXMLDomNodePictureList.Item(i).childNodes(XMLIPicture_Name).Text
        strSourceImageName = objXMLDomNodePictureList.Item(i).childNodes(XMLIPicture_URI).Text
        strSourceImageName = Replace(strSourceImageName, "%sd%", g_strExeDir)
        strSource = g_strExeDir & "temp\file" & i & ".jpg"
        
        If lDestinationType = XMLIDestination_Type_DigiFrame Then
            
            label.Caption = "Loading to frame """ & strImageName & ": " & strSourceImageName & """"
            label.Refresh
            progress.value = progress.value + 1
                                               
            If objDigiFrame.PutFile(Replace(strFileTemplate, "%i%", strPageNum), strSource, True) = False Then
                strStatus = "Could not write to Digi-Frame image """ & strImageName & ": " & strSourceImageName & """"
                GoTo done
            End If
        Else
            strDestination = strDestinationDirectory & Replace(strFileTemplate, "%i%", strPageNum)
            Call fso.CopyFile(strSource, strDestination, True)
        End If
    Next
    
    Rem Tell the user we are finishing up
    label.Caption = "Finishing Task"
    label.Refresh
    progress.value = progress.value + 1
    
    Rem We are successful
    strStatus = ""
    
    Rem Store the status of running the job
    Set objXMLDomNodeStatus = objXMLDomNodeTask.childNodes(XMLITask_Status)
    objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).Text = DateToISODate(Now)
    If strStatus <> "" Then
        objXMLDomNodeStatus.childNodes(XMLIStatus_Failed).Text = "1"
        objXMLDomNodeStatus.childNodes(XMLIStatus_Reason).Text = strStatus
    Else
        objXMLDomNodeStatus.childNodes(XMLIStatus_Failed).Text = "0"
        objXMLDomNodeStatus.childNodes(XMLIStatus_Reason).Text = ""
    End If
    
    Rem Clear the status frame
    
    
    GoTo done
    
error:
    strStatus = "Internal error"
        
done:
    On Error GoTo 0
    
    Rem Close the connection to the frame
    If fFrameConnect Then
        Call objDigiFrame.Disconnect
        fFrameConnect = False
    End If
        
    Rem Delete the temp directory
    If fTempDirCreated Then
        Call fso.DeleteFolder(g_strExeDir & "temp", True)
        fTempDirCreated = False
    End If
    
    Rem Hang up if the task forced the connection
    If fOnline And fTaskConnected Then
        InternetAutodialHangup (0)
        fOnline = False
        fTaskConnected = False
    End If
    
    Rem Store the job status
    objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).Text = DateToISODate(Now)
    objXMLDomNodeStatus.childNodes(XMLIStatus_Reason).Text = strStatus
    If strStatus = "" Then
        objXMLDomNodeStatus.childNodes(XMLIStatus_Failed).Text = "0"
    Else
        objXMLDomNodeStatus.childNodes(XMLIStatus_Failed).Text = "1"
    End If
    
    Rem Reset all the feedback controls
    progress.value = 0
    progress.max = 100
    label.Caption = "No tasks currently running"
    label.Visible = False
    Screen.MousePointer = 0
    statusFrame.Caption = "Execution Status"
    statusFrame.Refresh
    frmMain.EndAnimate
    
    Rem Return the status text
    RunTask = strStatus
End Function

Public Function RunFile(strSource As String, statusFrame As Frame, label As label, progress As ProgressBar, objXMLDomNodeFormatProfileSettings As IXMLDOMNode, iType As Integer, iPort As Integer, iMedia As Integer, strDirectory As String, fOverwrite As Boolean, strFile As String) As String
    Dim fResult As Boolean
    Dim fOnline As Boolean
    Dim fCanConnect As Boolean
    Dim fTaskConnected As Boolean
    Dim strStatus As String
    Dim fTempDirCreated As Boolean
    Dim fFrameConnected As Boolean
    Dim objDigiFrame As clsDigiFrame
    Dim lConnectedState As Long
    Dim fso As Scripting.FileSystemObject
    
On Error GoTo error

    Rem Initialize interesting local variables for cleanup
    fTaskConnected = False
    fTempDirCreated = False
    fTempDirCreated = False
    fFrameConnected = False
    
    Rem Initialize feedback
    frmMain.StartAnimate
    statusFrame.Caption = "Execution Status: Pump Image"
    label.Caption = "Initializing"
    label.Refresh
    label.Visible = True
    progress.value = 0
    progress.max = 3
    Screen.MousePointer = 11
    
    strStatus = "There was an error processing the file"
    
    Rem Connect to the web if necessary
    If Left(strSource, 5) = "http:" Or Left(strSource, 4) = "ftp:" Then
        Rem Determine if we are online
        lConnectedState = 0
        fCanConnect = True
        fOnline = InternetGetConnectedState(lConnectedState, 0)
        If fOnline = False Then
            fOnline = False
            If (lConnectedState And INTERNET_CONNECTION_MODEM) = False Then
                fCanConnect = False
            End If
        End If

        If fOnline = False And fCanConnect = False Then
            strStatus = "Unable to connect to Internet."
            GoTo done
        End If
    
        If InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0) = False Then
            strStatus = "There was a problem connecting to the Internet."
            GoTo done
        End If
            
        fTaskConnected = True
    End If
            
    Rem Create a file system object for use later
    Set fso = New Scripting.FileSystemObject
    
    Rem Create a temporary directory for saving files into
    If fso.FolderExists(g_strExeDir & "temp2") Then
        Call fso.DeleteFolder(g_strExeDir & "temp2", True)
    End If
    Call fso.CreateFolder(g_strExeDir & "temp2")
    fTempDirCreated = True
    
    Rem Update the progress
    label.Caption = "Saving image"
    label.Refresh
    progress.value = progress.value + 1
    
    Rem Save the image into the temporary directory with the proper formatting
    fResult = SaveImageXML(g_strExeDir & "temp2", "temp.jpg", strSource, objXMLDomNodeFormatProfileSettings)
    If fResult = False Then
        strStatus = "There was a problem processing the file"
        GoTo done
    End If
            
    If iType = XMLIDestination_Type_Directory Then
        
        label.Caption = "Copying image"
        label.Refresh
        progress.value = progress.value + 1
        
        Rem We just need to copy the image to our destination
        Call fso.CopyFile(g_strExeDir & "temp2\temp.jpg", DirectoryFileToFullPath(strDirectory, strFile))
    Else
        label.Caption = "Saving image to Digi-Frame"
        label.Refresh
        progress.value = progress.value + 1
        
        Rem Open the digi-frame
        Set objDigiFrame = New clsDigiFrame
        objDigiFrame.Port = iPort
        objDigiFrame.Card = iMedia
        
        fFrameConnected = objDigiFrame.Connect
        If fFrameConnected = False Then
            strStatus = "There was a problem connecting to the frame"
            GoTo done
        End If
    
        Rem Put the file onto the digi-frame
        If objDigiFrame.PutFile(strFile, g_strExeDir & "temp2\temp.jpg", True) = False Then
            strStatus = "There was a problem copying the image onto the frame"
            GoTo done
        End If
    End If
    
    label.Caption = "Finishing"
    label.Refresh
    progress.value = progress.value + 1
    
    strStatus = ""
    
    GoTo done
    
error:
    strStatus = "An internal error occured"
    
done:
    On Error GoTo 0
    
    Rem Close the connection to the frame
    If fFrameConnected Then
        Call objDigiFrame.Disconnect
        fFrameConnected = False
    End If
    
    Rem Delete the temp directory
    If fTempDirCreated Then
        Call fso.DeleteFolder(g_strExeDir & "temp2", True)
        fTempDirCreated = False
    End If
    
    Rem Disconnect if we connected
    If fOnline And fTaskConnected Then
        InternetAutodialHangup (0)
        fTaskConnected = False
        fOnline = False
    End If
    
    Rem Reset the status/feedback
    statusFrame.Caption = "Execution Status"
    statusFrame.Refresh
    progress.value = 0
    progress.max = 100
    label.Caption = "No tasks currently running"
    label.Visible = False
    Screen.MousePointer = 0
    frmMain.EndAnimate
    
    RunFile = strStatus
End Function

Rem ------------------------------------------------------------
Rem SaveImageXML
Rem
Rem Saves an image according to the XML format passed in
Rem
Function SaveImageXML(strDirectory As String, strFile As String, ByVal refURL As String, objXMLDomNodeFormatProfileSettings As IXMLDOMNode) As Boolean
    Dim lWidth, lHeight, lLeftMargin, lRightMargin, lTopMargin, lBottomMargin As Long
    Dim lPadColor, lMarginColor, lVerticalAlign, lHorizontalAlign As Long
    Dim fMargins, fPad, fGrow, fShrink As Boolean
    Dim fThumbnail As Boolean
    Dim lThumbWidth, lThumbHeight As Long
    Dim fRotate As Boolean
    Dim iRotate As Integer
    Dim lCompression As Long
    
    Rem Get information about the formatting
    lWidth = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Width).Text)
    lHeight = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Height).Text)
    fThumbnail = CBool(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Thumbnail).Text)
    lThumbWidth = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_ThumbWidth).Text)
    lThumbHeight = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_ThumbHeight).Text)
    fMargins = CBool(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Margins).Text)
    lLeftMargin = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_LeftMargin).Text)
    lRightMargin = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Rightmargin).Text)
    lTopMargin = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_TopMargin).Text)
    lBottomMargin = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_BottomMargin).Text)
    lMarginColor = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_MarginColor).Text)
    lVerticalAlign = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_VerticalAlign).Text)
    lHorizontalAlign = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_HorizontalAlign).Text)
    fGrow = CBool(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Grow).Text)
    fShrink = CBool(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Shrink).Text)
    fRotate = CBool(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Rotate).Text)
    If fRotate = True Then
        If CInt(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_RotateDirection).Text) = XMLIFormatSettings_RotateDirection_CW Then
            iRotate = 90
        Else
            iRotate = 270
        End If
    Else
        iRotate = 0
    End If
    fPad = CBool(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Pad).Text)
    lPadColor = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_PadColor).Text)
    lCompression = CLng(objXMLDomNodeFormatProfileSettings.childNodes(XMLIFormatSettings_Compression).Text)
        
    SaveImageXML = SaveImageVar(strDirectory, strFile, refURL, lWidth, lHeight, lTopMargin, lBottomMargin, lLeftMargin, lRightMargin, lMarginColor, lVerticalAlign, lHorizontalAlign, fGrow, fShrink, iRotate, fPad, lPadColor, lCompression, fThumbnail, lThumbWidth, lThumbHeight)
End Function


Rem ------------------------------------------------------------
Rem SaveImageVar
Rem
Rem This is the main function for saving a file/URL into a new file based on the formatting rules
Rem
Function SaveImageVar(strDirectory As String, strFile As String, ByVal refURL As String, ByVal desiredWidth As Integer, ByVal desiredHeight As Integer, topMargin, bottomMargin, leftMargin, rightMargin, colorMargin, vAlign, hAlign, fGrow, fShrink, iRotate, fPad, colorPad, lCompression, fThumbnail, lThumbWidth, lThumbHeight) As Boolean
    Dim LEAD1 As Object ' Image
    Dim LEAD2 As Object ' Margin
    Dim LEAD3 As Object ' Padding
    Dim sourceWidth, sourceHeight
    Dim framedWidth, framedHeight
    Dim stampw, stamph
    Dim fileData
    Dim numbits
    Dim Format
    Dim ext
    Dim offsetTop, offsetLeft
    Dim fso
    Dim fDidGrow, fDidShrink, fDidRotate As Boolean
    Dim strFullFileName As String
    Dim bData() As Byte
    Dim intFile As Integer
    Dim Inet1 As Inet
    
    Rem Set Inet1 = New Inet
    
On Error GoTo error
    strFullFileName = DirectoryFileToFullPath(strDirectory, strFile)
    
    fDidGrow = False
    fDidShrink = False
    fDidRotate = False
    
    Rem Create some helper objects
    Rem Set LEAD1 = CreateObject("LEAD.LeadCtrl.120")
    Rem Set LEAD2 = CreateObject("LEAD.LeadCtrl.120")
    Rem Set LEAD3 = CreateObject("LEAD.LeadCtrl.120")
    Set LEAD1 = frmControls.LEAD1
    Set LEAD2 = frmControls.LEAD2
    Set LEAD3 = frmControls.LEAD3
    Call LEAD1.UnlockSupport(L_SUPPORT_GIFLZW, k_UnlockKey_GIF_V12)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
       
    Rem Load the file
    If Left(refURL, 5) = "http:" Or Left(refURL, 4) = "ftp:" Then
        Rem Get the URL as binary data
        intFile = FreeFile()
        bData() = frmControls.Inet1.OpenURL(refURL, icByteArray)
    
        Rem Save to the file
        Open strFullFileName For Binary Access Write As #intFile
        Put #intFile, , bData()
        Close #intFile
                
        Rem Try and load the file we got
        LEAD1.Load strFullFileName, 0, 0, 1
    Else
        
        LEAD1.Load refURL, 0, 0, 1
    End If
    
    Rem Get some general information aabout the source image
    Format = LEAD1.InfoFormat
    numbits = LEAD1.InfoBits
    sourceWidth = LEAD1.BitmapWidth
    sourceHeight = LEAD1.BitmapHeight
    
    Rem GIF files must be saved as jpgs to work properly
    If Format = FILE_GIF Or numbits <> 24 Then
        LEAD1.save strFullFileName, FILE_LEAD1JFIF, 24, 2, 0
        LEAD1.Load strFullFileName, 0, 0, 1
    End If
  
    Rem Rotate the image for optimal display
    If iRotate <> 0 And ((sourceWidth < sourceHeight And desiredWidth > desiredHeight) Or (sourceWidth > sourceHeight And desiredWidth < desiredHeight)) Then
        Call LEAD1.Rotate(iRotate * 100, ROTATE_RESIZE, RGB(0, 0, 0))
        sourceWidth = LEAD1.BitmapWidth
        sourceHeight = LEAD1.BitmapHeight
        
        fDidRotate = True
    End If
    
    Rem When checking our scaling and growing, use the margin
    sourceWidth = sourceWidth + leftMargin + rightMargin
    sourceHeight = sourceHeight + topMargin + bottomMargin
            
    Rem Check if we need to change aspect ratio
    If sourceWidth < desiredWidth And sourceHeight < desiredHeight Then
        If fGrow = False Then
            framedWidth = sourceWidth
            framedHeight = sourceHeight
        ElseIf sourceWidth / desiredWidth > sourceHeight / desiredHeight And fShrink = True Then
            framedWidth = desiredWidth
            framedHeight = desiredWidth * sourceHeight / sourceWidth
            fDidGrow = True
        Else
            framedHeight = desiredHeight
            framedWidth = desiredHeight * sourceWidth / sourceHeight
            fDidGrow = True
        End If
    ElseIf fShrink = True Then
        If sourceWidth / desiredWidth > sourceHeight / desiredHeight And fShrink = True Then
            framedWidth = desiredWidth
            framedHeight = desiredWidth * sourceHeight / sourceWidth
            fDidShrink = True
        Else
            framedHeight = desiredHeight
            framedWidth = desiredHeight * sourceWidth / sourceHeight
            fDidShrink = True
        End If
    Else
        framedWidth = sourceWidth
        framedHeight = sourceHeight
    End If
    
    LEAD1.Size framedWidth, framedHeight, RESIZE_RESAMPLE
        
    Rem Create the bitmap containing the framed image and the margin and print the image into it
    Call LEAD2.CreateBitmap(framedWidth, framedHeight, LEAD1.InfoBits)
    Call LEAD2.Fill(colorMargin)
    Call LEAD2.Combine(leftMargin, topMargin, framedWidth - (leftMargin + rightMargin), framedHeight - (topMargin + bottomMargin), LEAD1.Bitmap, 0, 0, (CB_OP_ADD + CB_DST_0))
            
    Rem Now deal with the padding
    If fPad = True And (fDidGrow Or fDidShrink) Then
        Call LEAD3.CreateBitmap(desiredWidth, desiredHeight, LEAD1.InfoBits)
        
        Call LEAD3.Fill(colorPad)
        
        If vAlign = 0 Then
            offsetTop = 0
        ElseIf vAlign = 1 Then
            offsetTop = (desiredHeight - framedHeight) / 2
        ElseIf vAlign = 2 Then
            offsetTop = (desiredHeight - framedHeight)
        End If
    
        If hAlign = 0 Then
            offsetLeft = 0
        ElseIf hAlign = 1 Then
            offsetLeft = (desiredWidth - framedWidth) / 2
        ElseIf hAlign = 2 Then
            offsetLeft = (desiredWidth - framedWidth)
        End If
        
        Call LEAD3.Combine(offsetLeft, offsetTop, framedWidth, framedHeight, LEAD2.Bitmap, 0, 0, (CB_OP_ADD + CB_DST_0))
        
        Rem Deal with the thumbnail or regular save
        If fThumbnail Then
            If LEAD3.BitmapWidth / lThumbWidth > LEAD3.BitmapHeight / lThumbHeight Then
                stampw = lThumbWidth
                stamph = lThumbWidth * LEAD3.BitmapHeight / LEAD3.BitmapWidth
            Else
                stamph = lThumbHeight
                stampw = lThumbHeight * LEAD3.BitmapWidth / LEAD3.BitmapHeight
            End If
            Call LEAD3.SaveWithStamp(strFullFileName, FILE_LEAD1JFIF, 24, lCompression, stampw, stamph, 24)
        Else
            Call LEAD3.save(strFullFileName, FILE_LEAD1JFIF, 24, lCompression, 0)
        End If
    Else
        If fThumbnail Then
            
            If LEAD2.BitmapWidth / lThumbWidth > LEAD2.BitmapHeight / lThumbHeight Then
                stampw = lThumbWidth
                stamph = lThumbWidth * LEAD2.BitmapHeight / LEAD2.BitmapWidth
            Else
                stamph = lThumbHeight
                stampw = lThumbHeight * LEAD2.BitmapWidth / LEAD2.BitmapHeight
            End If
            Call LEAD2.SaveWithStamp(strFullFileName, FILE_LEAD1JFIF, 24, lCompression, stampw, stamph, 24)
        Else
            Call LEAD2.save(strFullFileName, FILE_LEAD1JFIF, 24, lCompression, 0)
        End If
    End If
    
    SaveImageVar = True
    
    Unload frmControls
    
    Exit Function
error:
    Unload frmControls
    SaveImageVar = False
    
End Function

