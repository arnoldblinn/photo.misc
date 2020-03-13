VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTasks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tasks"
   ClientHeight    =   4056
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4056
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton b_Import 
      Caption         =   "Import..."
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton b_Export 
      Caption         =   "Export..."
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox l_Tasks 
      Height          =   3312
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton b_Wizard 
      Caption         =   "Wizard..."
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton b_Run 
      Caption         =   "Run"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton b_Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton b_Edit 
      Caption         =   "Edit..."
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton b_New 
      Caption         =   "New..."
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem ------------------------------------------
Rem XMLChanged
Rem
Rem Called when there is a change to the XML that needs to be
Rem reflected in the UI and on disk
Rem
Private Function XMLChanged(iListIndex As Integer)
    Dim i As Integer
    
    l_Tasks.ListIndex = -1
    
    Rem Remove all the current items
    For i = 0 To l_Tasks.ListCount - 1
        l_Tasks.RemoveItem (0)
    Next
    
    Rem Initialize the form from the xml state
    For i = 0 To g_objXMLDomNodeTasks.childNodes.length - 1
        Dim strText As String
        Dim strStatus As String
        Dim dtLastRun As Date
        Dim objXMLDomNodeStatus As MSXML.IXMLDOMNode
        
        strText = g_objXMLDomNodeTasks.childNodes.Item(i).childNodes(XMLITask_Name).Text
        Set objXMLDomNodeStatus = g_objXMLDomNodeTasks.childNodes.Item(i).childNodes(XMLITask_Status)
        
        If objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).Text = "" Then
            strText = strText & ": Never run"
        Else
            dtLastRun = ISODateToDate(objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).Text)
            If objXMLDomNodeStatus.childNodes(XMLIStatus_Failed).Text = "1" Then
                strText = strText & ": Failed on " & dtLastRun
            Else
                strText = strText & ": Ran on " & dtLastRun
            End If
        End If
        
        l_Tasks.AddItem (strText)
    Next
    
    If l_Tasks.ListCount > 0 Then
        If iListIndex >= l_Tasks.ListCount Then
            l_Tasks.ListIndex = l_Tasks.ListCount - 1
        Else
            l_Tasks.ListIndex = iListIndex
        End If
    Else
        b_Edit.Enabled = False
        b_Delete.Enabled = False
        b_Run.Enabled = False
        b_Export.Enabled = False
    End If
End Function


Private Sub ExecuteEdit()
        
    Set frmTask.XMLDomNode = g_objXMLDomNodeTasks.childNodes.Item(l_Tasks.ListIndex)
    
    frmTask.Show (vbModal)
    
    If frmTask.Result = True Then
        Call g_objXMLDomNodeTasks.replaceChild(frmTask.XMLDomNode, g_objXMLDomNodeTasks.childNodes.Item(l_Tasks.ListIndex))
        Call XMLChanged(l_Tasks.ListIndex)

    End If
    
    Unload frmTask
End Sub

Private Sub b_Delete_Click()
    Call ExecuteDelete
End Sub

Private Sub b_Edit_Click()
    Call ExecuteEdit
End Sub

Private Sub b_Export_Click()
    Call ExecuteExport
End Sub

Private Sub b_Import_Click()
    Call ExecuteImport
End Sub

Private Sub b_New_Click()
    Call ExecuteNew
End Sub

Private Sub b_OK_Click()
    Rem Save the XML
    Call SaveTasks
    
    Me.Hide
End Sub

Private Sub b_Run_Click()
    Call ExecuteRun
End Sub

Private Sub b_Wizard_Click()
    Call ExecuteWizard
End Sub

Private Sub l_Tasks_Click()
    Dim fSelected As Boolean
    If l_Tasks.ListIndex = -1 Then
        fSelected = False
    Else
        fSelected = True
    End If
        
    b_Edit.Enabled = fSelected
    b_Delete.Enabled = fSelected
    b_Run.Enabled = fSelected
    b_Export.Enabled = fSelected
End Sub

Private Sub l_Tasks_DblClick()
    Call ExecuteEdit
End Sub


Private Sub ExecuteRun()
    
    Call RunTask(g_objXMLDomNodeTasks.childNodes.Item(l_Tasks.ListIndex), frmMain.frame_Status, frmMain.l_Status, frmMain.p_Progress, 1)
    
    Call ExecuteStatus
    
    XMLChanged (l_Tasks.ListIndex)
End Sub

Private Sub ExecuteStatus()
    Dim objXMLDomNodeStatus As IXMLDOMNode
    Dim dtLastRun As Date
    Dim strStatus As String
    Dim strFailed As String
    Dim strMessage As String
    Dim strReason As String
    

    Set objXMLDomNodeStatus = g_objXMLDomNodeTasks.childNodes.Item(l_Tasks.ListIndex).childNodes(XMLITask_Status)
    
    If objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).Text = "" Then
        MsgBox "Task has never been run.", vbInformation, "Status"
    Else
        
        dtLastRun = ISODateToDate(objXMLDomNodeStatus.childNodes(XMLIStatus_LastRun).Text)
        If CInt(objXMLDomNodeStatus.childNodes(XMLIStatus_Failed).Text) = 0 Then
            strFailed = "Succeded"
            strReason = ""
        Else
            strFailed = "Failed"
            strStatus = objXMLDomNodeStatus.childNodes(XMLIStatus_Reason).Text
            strReason = vbCr & "    " & "Reason: " & vbTab & strStatus
        End If
        
        MsgBox "Task Status" & vbCr & vbCr & "    " & "Date run: " & vbTab & dtLastRun & vbCr & "    " & "Result:     " & vbTab & strFailed & strReason, vbInformation, "Status"
    End If
End Sub

Private Sub ExecuteWizard()
    Dim objXMLDom As MSXML.DOMDocument
        
    frmWizard.Show (vbModal)
    
    If frmWizard.Result = True Then
        Set objXMLDom = LoadDOMFromString(frmWizard.XML, "Task", "1")
        Call g_objXMLDomNodeTasks.appendChild(objXMLDom.childNodes(0))
        Call XMLChanged(g_objXMLDomNodeTasks.childNodes.length - 1)
    End If
    
    Unload frmWizard
    
End Sub

Private Sub ExecuteNew()
    Dim strDestination As String
    Dim strFormat As String
    Dim strSchedule As String
    Dim strSource As String
    Dim strStatus As String
    Dim strXML As String
    Dim objXMLDom As MSXML.DOMDocument
    Dim iFormat As Integer
    Dim iWidth As Integer
    Dim iHeight As Integer
    Dim iThumbnail As Integer
    Dim iThumbWidth As Integer
    Dim iThumbHeight As Integer
    Dim iRotate As Integer
    
    Set objXMLDom = New DOMDocument
        
    iWidth = 640
    iHeight = 480
    iThumbnail = 0
    iThumbWidth = 0
    iThumbHeight = 0
    iRotate = 0

    strStatus = "<Status><LastRun></LastRun><Failed>0</Failed><Reason></Reason></Status>"
    strSource = "<Source><Type>0</Type><Album Version=""1""><Name>New Album</Name><PictureList></PictureList></Album><AlbumURI></AlbumURI></Source>"
    strSchedule = "<Schedule><Type>0</Type><Disable>0</Disable><Connect>1</Connect><Hours>0</Hours><Minutes>0</Minutes><Weekday>0</Weekday><Monthday>1</Monthday></Schedule>"
    strFormat = "<Format><Name></Name><FormatSettings><Width>" & iWidth & "</Width><Height>" & iHeight & "</Height><Grow>1</Grow><Shrink>1</Shrink><Rotate>" & iRotate & "</Rotate><RotateDirection>0</RotateDirection><Pad>1</Pad><PadColor>0</PadColor><VerticalAlign>1</VerticalAlign><HorizontalAlign>1</HorizontalAlign><Margins>1</Margins><TopMargin>5</TopMargin><LeftMargin>5</LeftMargin><RightMargin>5</RightMargin><BottomMargin>5</BottomMargin><MarginColor>" & RGB(255, 255, 255) & "</MarginColor><Compression>52</Compression><Thumbnail>" & iThumbnail & "</Thumbnail><ThumbWidth>" & iThumbWidth & "</ThumbWidth><ThumbHeight>" & iThumbHeight & "</ThumbHeight></FormatSettings></Format>"
    strDestination = "<Destination><Type>0</Type><Directory>" & g_strExeDir & "</Directory><DirectoryDelete>0</DirectoryDelete><FileTemplate>file%i%.jpg</FileTemplate><DigiFramePort>0</DigiFramePort><DigiFrameMedia>0</DigiFrameMedia></Destination>"
    strXML = "<Task Version=""1""><Name>New Task</Name>" & strStatus & strSource & strSchedule & strFormat & strDestination & "</Task>"
    
    Set objXMLDom = LoadDOMFromString(strXML, "Task", "1")
    Set frmTask.XMLDomNode = objXMLDom.childNodes.Item(0)
            
    frmTask.Show (vbModal)
    
    If frmTask.Result = True Then
        Call g_objXMLDomNodeTasks.appendChild(frmTask.XMLDomNode)
        Call XMLChanged(g_objXMLDomNodeTasks.childNodes.length - 1)
    End If
    
    Unload frmTask
End Sub


Rem --------------------------------------------------------------------------------
Rem ExecuteDelete()
Rem
Rem Delete item
Rem
Private Sub ExecuteDelete()
                            
    Call g_objXMLDomNodeTasks.removeChild(g_objXMLDomNodeTasks.childNodes.Item(l_Tasks.ListIndex))
        
    Call XMLChanged(l_Tasks.ListIndex)
End Sub

Rem --------------------------------------------------------
Rem ExecuteExport
Rem
Rem Code to handle importing a task
Rem
Private Sub ExecuteExport()
    
    Rem Get the file to export to
    CommonDialog1.InitDir = CurDir
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Graphic Pump Task File (*.gpt)|*.gpt"
    CommonDialog1.ShowSave
    
    Rem If a file was specified then write the exported data to the file
    If CommonDialog1.FileName <> "" Then
        Call SaveDOMToFile(CommonDialog1.FileName, g_objXMLDomNodeTasks.childNodes.Item(l_Tasks.ListIndex).XML)
    End If
    
End Sub


Rem --------------------------------------------------------
Rem ExecuteImport
Rem
Rem Code to handle importing a task
Rem
Private Sub ExecuteImport()
    Dim objXMLDocumentTask As MSXML.DOMDocument
    Dim fso As Scripting.FileSystemObject
    Dim strError As String
    
    strError = ""
            
    Rem Use the common dialog to open the task (gpt) file
    CommonDialog1.Filter = "Graphic Pump Task File (*.gpt)|*.gpt"
    CommonDialog1.InitDir = CurDir
    CommonDialog1.FileName = ""
    CommonDialog1.Flags = cdlOFNHideReadOnly
    
    CommonDialog1.ShowOpen
    
    Set fso = New Scripting.FileSystemObject
    
    Rem If we specified a file, do the import
    If CommonDialog1.FileName <> "" Then
    
        If fso.FileExists(CommonDialog1.FileName) Then
        
            Rem Import the file
            Set objXMLDocumentTask = LoadDOMFromFile(CommonDialog1.FileName, "Task", "1")
            
            Rem Validate that it imported fine
            If objXMLDocumentTask Is Nothing Then
                MsgBox "There was an error importing the task.", vbExclamation, "Error"
                Exit Sub
            End If
                    
            objXMLDocumentTask.childNodes(0).childNodes(XMLITask_Status).childNodes(XMLIStatus_LastRun).Text = ""
            objXMLDocumentTask.childNodes(0).childNodes(XMLITask_Status).childNodes(XMLIStatus_Reason).Text = ""
        
            Rem Add the new task
            Call g_objXMLDomNodeTasks.appendChild(objXMLDocumentTask.childNodes(0))
        
            Rem Refresh the XML from the UI
            XMLChanged (g_objXMLDomNodeTasks.childNodes.length)
        Else
            MsgBox "The .gpt file does not exist.", vbExclamation, "Error"
            Exit Sub
        End If
    End If
End Sub



Private Sub Form_Load()
    Rem Populate the UI
    Call XMLChanged(0)
End Sub
