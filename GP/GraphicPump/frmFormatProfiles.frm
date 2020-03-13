VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFormatProfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format Profiles"
   ClientHeight    =   3228
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4464
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3228
   ScaleWidth      =   4464
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   2760
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton b_Import 
      Caption         =   "Import..."
      Height          =   372
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton b_Export 
      Caption         =   "Export..."
      Height          =   372
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   972
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton b_Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton b_Edit 
      Caption         =   "Edit..."
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton b_New 
      Caption         =   "New..."
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.ListBox l_Profiles 
      Height          =   2352
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmFormatProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function refreshui(iListIndex As Integer)
    Dim i As Integer
    
    l_Profiles.ListIndex = -1
    
    For i = 0 To l_Profiles.ListCount - 1
        l_Profiles.RemoveItem (0)
    Next
    
    For i = 0 To g_objXMLDomNodeFormatProfiles.childNodes.length - 1
        l_Profiles.AddItem (g_objXMLDomNodeFormatProfiles.childNodes(i).childNodes(XMLIFormatProfile_Name).Text)
    Next
    
    If l_Profiles.ListCount > 0 Then
        If iListIndex >= l_Profiles.ListCount Then
            l_Profiles.ListIndex = l_Profiles.ListCount - 1
        Else
            l_Profiles.ListIndex = iListIndex
        End If
    Else
        b_Edit.Enabled = False
        b_Delete.Enabled = False
        b_Export.Enabled = False
    End If
    
End Function

Private Sub b_Delete_Click()
    Dim i As Integer
    Dim j As Integer
    Dim strName As String
    
    Rem Get the selected item
    i = l_Profiles.ListIndex
    If i = -1 Then
        Exit Sub
    End If
    
    Rem Go through the jobs and see if any of them reference this format
    Rem If they do, change the type and copy the settings
    strName = g_objXMLDomNodeFormatProfiles.childNodes(i).childNodes(XMLIFormatProfile_Name).Text
    For j = 0 To g_objXMLDomNodeTasks.childNodes.length - 1
        If strName = g_objXMLDomNodeTasks.childNodes(j).childNodes(XMLITask_Format).childNodes(XMLIFormat_Name).Text Then
            g_objXMLDomNodeTasks.childNodes(j).childNodes(XMLITask_Format).childNodes(XMLIFormat_Name).Text = ""
            Call g_objXMLDomNodeTasks.childNodes(j).childNodes(XMLITask_Format).replaceChild(g_objXMLDomNodeFormatProfiles.childNodes(i).childNodes(XMLIFormatProfile_Settings).cloneNode(True), g_objXMLDomNodeTasks.childNodes(j).childNodes(XMLITask_Format).childNodes(XMLIFormat_Settings))
        End If
    Next
    
    Call g_objXMLDomNodeFormatProfiles.removeChild(g_objXMLDomNodeFormatProfiles.childNodes.Item(i))
    
    Call refreshui(i)
    
    Rem Save the XML
    Call SaveTasks
    Call SaveFormatProfiles
End Sub

Private Sub b_Edit_Click()

    Set frmFormatProfile.XMLDomNode = g_objXMLDomNodeFormatProfiles.childNodes.Item(l_Profiles.ListIndex)
    frmFormatProfile.FormatIndex = l_Profiles.ListIndex
    frmFormatProfile.Show (vbModal)
    If frmFormatProfile.Result = True Then
    
        Dim strOldName As String
        Dim strNewName As String
        Dim i As Integer
        Dim objXMLDomNodeFormatProfileSettings As IXMLDOMNode
        
        strOldName = g_objXMLDomNodeFormatProfiles.childNodes.Item(l_Profiles.ListIndex).childNodes(XMLIFormatProfile_Name).Text
        
        strNewName = frmFormatProfile.XMLDomNode.childNodes(XMLIFormatProfile_Name).Text
        Set objXMLDomNodeFormatProfileSettings = frmFormatProfile.XMLDomNode.childNodes(XMLIFormatProfile_Settings)
        
        Rem Check all the tasks.  Any job that used this format type should update its name
        Rem and settings
        For i = 0 To g_objXMLDomNodeTasks.childNodes.length - 1
            If g_objXMLDomNodeTasks.childNodes(i).childNodes(XMLITask_Format).childNodes(XMLIFormat_Name).Text = strOldName Then
                g_objXMLDomNodeTasks.childNodes(i).childNodes(XMLITask_Format).childNodes(XMLIFormat_Name).Text = strNewName
                Call g_objXMLDomNodeTasks.childNodes(i).childNodes(XMLITask_Format).replaceChild(objXMLDomNodeFormatProfileSettings.cloneNode(True), g_objXMLDomNodeTasks.childNodes(i).childNodes(XMLITask_Format).childNodes(XMLIFormat_Settings))
            End If
        Next
            
        Call g_objXMLDomNodeFormatProfiles.replaceChild(frmFormatProfile.XMLDomNode, g_objXMLDomNodeFormatProfiles.childNodes.Item(l_Profiles.ListIndex))
        
        Call refreshui(l_Profiles.ListIndex)
        
        Call SaveFormatProfiles
        Call SaveTasks
    End If
    
    Unload frmFormatProfile
        
End Sub

Private Sub b_Export_Click()
    Rem Get the file to export to
    CommonDialog1.InitDir = CurDir
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Graphic Pump Format Profile File (*.gpf)|*.gpf"
    CommonDialog1.ShowSave
    
    Rem If a file was specified then write the exported data to the file
    If CommonDialog1.FileName <> "" Then
        Call SaveDOMToFile(CommonDialog1.FileName, g_objXMLDomNodeFormatProfiles.childNodes.Item(l_Profiles.ListIndex).XML)
    End If
End Sub

Private Sub b_New_Click()
    Dim strFormatSettings As String
    Dim objXMLDom As MSXML.DOMDocument
    
    strFormatSettings = "<FormatSettings><Width>640</Width><Height>480</Height><Grow>1</Grow><Shrink>1</Shrink><Rotate>0</Rotate><RotateDirection>0</RotateDirection><Pad>1</Pad><PadColor>0</PadColor><VerticalAlign>1</VerticalAlign><HorizontalAlign>1</HorizontalAlign><Margins>1</Margins><TopMargin>5</TopMargin><LeftMargin>5</LeftMargin><RightMargin>5</RightMargin><BottomMargin>5</BottomMargin><MarginColor>" & RGB(255, 255, 255) & "</MarginColor><Compression>52</Compression><Thumbnail>0</Thumbnail><ThumbWidth>0</ThumbWidth><ThumbHeight>0</ThumbHeight></FormatSettings>"
    Set objXMLDom = New DOMDocument
    Call objXMLDom.loadXML("<FormatProfile Version=""1""><Name>New Format Profile</Name>" & strFormatSettings & "</FormatProfile>")
    
    Set frmFormatProfile.XMLDomNode = objXMLDom.childNodes.Item(0)
    frmFormatProfile.FormatIndex = -1
    frmFormatProfile.Show (vbModal)
    If frmFormatProfile.Result = True Then
        Call g_objXMLDomNodeFormatProfiles.appendChild(frmFormatProfile.XMLDomNode)
        
        Call refreshui(g_objXMLDomNodeFormatProfiles.childNodes.length)
        
        Call SaveFormatProfiles
    End If
    
    Unload frmFormatProfile
    
End Sub

Private Sub b_OK_Click()
    Me.Hide
End Sub

Private Sub b_Import_Click()
    Dim objXMLDocumentFormatProfile As MSXML.DOMDocument
    Dim fso As Scripting.FileSystemObject
    Dim strError As String
    
    strError = ""
            
    Rem Use the common dialog to open the format profile file (gpf) file
    CommonDialog1.Filter = "Graphic Pump Format Profile File (*.gpf)|*.gpf"
    CommonDialog1.InitDir = CurDir
    CommonDialog1.FileName = ""
    CommonDialog1.Flags = cdlOFNHideReadOnly
    
    CommonDialog1.ShowOpen
    
    Set fso = New Scripting.FileSystemObject
    
    Rem If we specified a file, do the import
    If CommonDialog1.FileName <> "" Then
    
        If fso.FileExists(CommonDialog1.FileName) Then
        
            Rem Import the file
            Set objXMLDocumentFormatProfile = LoadDOMFromFile(CommonDialog1.FileName, "FormatProfile", "1")
            
            Rem Validate that it imported fine
            If objXMLDocumentFormatProfile Is Nothing Then
                MsgBox "There was an error importing the format profile.", vbExclamation, "Error"
                Exit Sub
            End If
            
            Rem Now see if the name is unique
            Dim i As Integer
            Dim j As Integer
            Dim strName As String
            j = 1
            strName = objXMLDocumentFormatProfile.childNodes(0).childNodes(XMLIFormatProfile_Name).Text
            i = 0
            While i < g_objXMLDomNodeFormatProfiles.childNodes.length
                If strName = g_objXMLDomNodeFormatProfiles.childNodes(i).childNodes(XMLIFormatProfile_Name).Text Then
                    strName = objXMLDocumentFormatProfile.childNodes(0).childNodes(XMLIFormatProfile_Name).Text + " " + CStr(j)
                    j = j + 1
                    i = 0
                Else
                    i = i + 1
                End If
            Wend
            
            objXMLDocumentFormatProfile.childNodes(0).childNodes(XMLIFormatProfile_Name).Text = strName
                    
            Rem Add the new format profile
            Call g_objXMLDomNodeFormatProfiles.appendChild(objXMLDocumentFormatProfile.childNodes(0))
        
            Rem Refresh the XML from the UI
            Call refreshui(g_objXMLDomNodeFormatProfiles.childNodes.length)
        Else
            MsgBox "The .gpf file does not exist.", vbExclamation, "Error"
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    Call refreshui(0)
End Sub

Private Sub l_Profiles_Click()
    Dim fSelected As Boolean
    If l_Profiles.ListIndex = -1 Then
        fSelected = False
    Else
        fSelected = True
    End If
    
    b_Edit.Enabled = fSelected
    b_Delete.Enabled = fSelected
    b_Export.Enabled = fSelected
End Sub

Private Sub l_Profiles_DblClick()
    Call b_Edit_Click
End Sub
