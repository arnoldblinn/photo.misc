VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFormatSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format Settings"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   ControlBox      =   0   'False
   Icon            =   "frmFormatSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Rotate"
      Height          =   1095
      Left            =   120
      TabIndex        =   34
      Top             =   3720
      Width           =   3495
      Begin VB.CheckBox chk_Rotate 
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmb_RotateDirection 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   660
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Rotate to Size"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label l_RotateDirection 
         Caption         =   "Direction"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Resizing"
      Height          =   1095
      Left            =   120
      TabIndex        =   35
      Top             =   2520
      Width           =   3495
      Begin VB.CheckBox chk_Grow 
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chk_Shrink 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Grow to Size"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Shrink to Size"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Padding"
      Height          =   2295
      Left            =   3720
      TabIndex        =   33
      Top             =   120
      Width           =   3495
      Begin VB.ComboBox cmb_Vertical 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cmb_Horizontal 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox chk_Pad 
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton b_PadColor 
         Caption         =   "Change..."
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Pad to Size"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label l_Vertical 
         Caption         =   "Vertical Align"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label l_Horizontal 
         Caption         =   "Horizontal Align"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Shape shape_PadColor 
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1440
         Shape           =   1  'Square
         Top             =   720
         Width           =   495
      End
      Begin VB.Label l_PadColor 
         Caption         =   "Pad Color"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   5760
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Margins"
      Height          =   3135
      Left            =   3720
      TabIndex        =   16
      Top             =   2520
      Width           =   3495
      Begin VB.CheckBox chk_Margins 
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton b_MarginColor 
         Caption         =   "Change..."
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox f_BottomMargin 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox f_RightMargin 
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox f_LeftMargin 
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox f_TopMargin 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Margins"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape shape_MarginColor 
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1440
         Shape           =   1  'Square
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label l_MarginColor 
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label l_Bottom 
         Caption         =   "Bottom"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label l_Top 
         Caption         =   "Top"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.Label l_Right 
         Caption         =   "Right"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label l_Left 
         Caption         =   "Left"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "JPEG Settings"
      Height          =   735
      Left            =   120
      TabIndex        =   36
      Top             =   4920
      Width           =   3495
      Begin VB.ComboBox cmb_Compression 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label l_Compression 
         Caption         =   "Compression"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Target Size"
      Height          =   2295
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   3495
      Begin VB.CheckBox chk_Thumbnail 
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox f_tWidth 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox f_tHeight 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox f_Width 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox f_Height 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Thumbnail"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label l_tWidth 
         Caption         =   "Thumb Width"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label l_tHeight 
         Caption         =   "Thumb Height"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label l_Height 
         Caption         =   "Height"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label l_Width 
         Caption         =   "Width"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   300
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmFormatSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem Boilerplate for every form
Dim gXMLDomNode As IXMLDOMNode
Dim gXMLDomNodeClone As IXMLDOMNode
Dim gResult As Boolean
Dim gForm As Form

Public Property Get Result() As Boolean
    Result = gResult
End Property

Public Property Set XMLDomNode(objXMLDomNode As IXMLDOMNode)
    Set gXMLDomNode = objXMLDomNode
    
    Set gXMLDomNodeClone = gXMLDomNode.cloneNode(True)
    
    Call XMLToUI
    Set gForm = Me
End Property

Public Property Get XMLDomNode() As IXMLDOMNode
    Set XMLDomNode = gXMLDomNode
End Property

Private Sub b_OK_Click()
    If UIToXML() = False Then
        Exit Sub
    End If
    
    gResult = True
    Set gXMLDomNode = gXMLDomNodeClone
    gForm.Hide
End Sub

Private Sub b_Cancel_Click()
    gResult = False
    gForm.Hide
End Sub


Rem XML To UI and UI to XML functions. Change, but every form has them
Private Function XMLToUI()
    f_Width.Text = gXMLDomNodeClone.childNodes(XMLIFormatSettings_Width).Text
    f_Height.Text = gXMLDomNodeClone.childNodes(XMLIFormatSettings_Height).Text
    chk_Grow.value = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_Grow).Text)
    chk_Shrink.value = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_Shrink).Text)
    chk_Rotate.value = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_Rotate).Text)
    cmb_RotateDirection.ListIndex = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_RotateDirection).Text)
    chk_Pad.value = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_Pad).Text)
    shape_PadColor.FillColor = CLng(gXMLDomNodeClone.childNodes(XMLIFormatSettings_PadColor).Text)
    cmb_Vertical.ListIndex = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_VerticalAlign).Text)
    cmb_Horizontal.ListIndex = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_HorizontalAlign).Text)
    chk_Margins.value = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_Margins).Text)
    f_LeftMargin.Text = gXMLDomNodeClone.childNodes(XMLIFormatSettings_LeftMargin).Text
    f_RightMargin.Text = gXMLDomNodeClone.childNodes(XMLIFormatSettings_Rightmargin).Text
    f_TopMargin.Text = gXMLDomNodeClone.childNodes(XMLIFormatSettings_TopMargin).Text
    f_BottomMargin.Text = gXMLDomNodeClone.childNodes(XMLIFormatSettings_BottomMargin).Text
    shape_MarginColor.FillColor = CLng(gXMLDomNodeClone.childNodes(XMLIFormatSettings_MarginColor).Text)
    cmb_Compression.ListIndex = (CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_Compression).Text) - 2) / 50
    chk_Thumbnail.value = CInt(gXMLDomNodeClone.childNodes(XMLIFormatSettings_Thumbnail).Text)
    f_tWidth.Text = gXMLDomNodeClone.childNodes(XMLIFormatSettings_ThumbWidth).Text
    f_tHeight.Text = gXMLDomNodeClone.childNodes(XMLIFormatSettings_ThumbHeight).Text
End Function

Private Function UIToXML() As Boolean
    UIToXML = False
    
    Rem Validate the values
    If ValidateInt(f_Width.Text, "Width", 50, 2000) = False Then Exit Function
    If ValidateInt(f_Height.Text, "Height", 50, 2000) = False Then Exit Function
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Width).Text = f_Width.Text
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Height).Text = f_Height.Text
    
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Margins).Text = chk_Margins.value
    If f_LeftMargin.Enabled Then
        If ValidateInt(f_LeftMargin.Text, "Left Margin", 0, 40) = False Then Exit Function
        If ValidateInt(f_RightMargin.Text, "Right Margin", 0, 40) = False Then Exit Function
        If ValidateInt(f_TopMargin.Text, "Top Margin", 0, 40) = False Then Exit Function
        If ValidateInt(f_BottomMargin.Text, "Bottom Margin", 0, 40) = False Then Exit Function
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_LeftMargin).Text = f_LeftMargin.Text
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_Rightmargin).Text = f_RightMargin.Text
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_TopMargin).Text = f_TopMargin.Text
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_BottomMargin).Text = f_BottomMargin.Text
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_MarginColor).Text = shape_MarginColor.FillColor
    End If
    
    Rem Set into the XML
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Grow).Text = chk_Grow.value
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Shrink).Text = chk_Shrink.value
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Rotate).Text = chk_Rotate.value
    If chk_Rotate.value = 1 Then
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_RotateDirection).Text = cmb_RotateDirection.ListIndex
    End If
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Pad).Text = chk_Pad.value
    If chk_Pad.value = 1 Then
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_PadColor).Text = shape_PadColor.FillColor
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_VerticalAlign).Text = cmb_Vertical.ListIndex
        gXMLDomNodeClone.childNodes(XMLIFormatSettings_HorizontalAlign).Text = cmb_Horizontal.ListIndex
    End If
    
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Compression).Text = cmb_Compression.ListIndex * 50 + 2
    
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_Thumbnail).Text = chk_Thumbnail.value
        
    Rem Validate that the thumbnail isn't too big for a .jpg file
    If chk_Thumbnail.value = 1 Then
        If ValidateInt(f_tWidth.Text, "Thumbnail Width", 10, 255) = False Then Exit Function
        If ValidateInt(f_tHeight.Text, "Thumbnail Height", 10, 255) = False Then Exit Function
    
        If Int(f_tWidth) * Int(f_tHeight) * 3 + 10 > 65535 Then
            MsgBox "Total size of thumbnail is too great.  Please reduce width or height.", vbExclamation, "Error"
            Exit Function
        End If
    End If
        
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_ThumbWidth).Text = f_tWidth.Text
    gXMLDomNodeClone.childNodes(XMLIFormatSettings_ThumbHeight).Text = f_tHeight.Text
        
    UIToXML = True
End Function

Rem -------------------
Rem Functionality specific to this form
Rem
Private Sub process_check()

    Rem Handle the enabling of the width/height controls
    l_Width.Enabled = True
    l_Height.Enabled = True
    f_Width.Enabled = True
    f_Height.Enabled = True
    chk_Thumbnail.Enabled = True
        
    Rem Handle the enabling of the thumbnail controls
    If chk_Thumbnail.value = 0 Then
        f_tWidth.Enabled = False
        f_tWidth = 0
        f_tHeight = 0
        f_tHeight.Enabled = False
        l_tWidth.Enabled = False
        l_tHeight.Enabled = False
    Else
        f_tWidth.Enabled = True
        f_tHeight.Enabled = True
        l_tWidth.Enabled = True
        l_tHeight.Enabled = True
    End If

    Rem Handle the enabling of the margin controls
    If chk_Margins = 0 Then
        l_Top.Enabled = False
        l_Bottom.Enabled = False
        l_Right.Enabled = False
        l_Left.Enabled = False
        l_MarginColor.Enabled = False
        shape_MarginColor.FillStyle = 7
        f_TopMargin.Enabled = False
        f_TopMargin.Text = "5"
        f_LeftMargin.Enabled = False
        f_LeftMargin.Text = "5"
        f_RightMargin.Enabled = False
        f_RightMargin.Text = "5"
        f_BottomMargin.Enabled = False
        f_BottomMargin.Text = "5"
        b_MarginColor.Enabled = False
    Else
        l_Top.Enabled = True
        l_Bottom.Enabled = True
        l_Right.Enabled = True
        l_Left.Enabled = True
        l_MarginColor.Enabled = True
        shape_MarginColor.FillStyle = 0
        f_TopMargin.Enabled = True
        f_LeftMargin.Enabled = True
        f_RightMargin.Enabled = True
        f_BottomMargin.Enabled = True
        b_MarginColor.Enabled = True
    End If
    
    Rem Handle the enabling of the rotate controls
    If chk_Rotate.value = 0 Then
        cmb_RotateDirection.Enabled = False
        l_RotateDirection.Enabled = False
    Else
        cmb_RotateDirection.Enabled = True
        l_RotateDirection.Enabled = True
    End If
        
    Rem Handle the enabling of the pad option and controls
    If chk_Grow.value = 0 And chk_Shrink.value = 0 Then
        chk_Pad.Enabled = False
    Else
        chk_Pad.Enabled = True
    End If
    
    If chk_Pad.value = 0 Or (chk_Grow.value = 0 And chk_Shrink.value = 0) Then
        l_PadColor.Enabled = False
        b_PadColor.Enabled = False
        l_Vertical.Enabled = False
        l_Horizontal.Enabled = False
        cmb_Vertical.Enabled = False
        cmb_Horizontal.Enabled = False
        shape_PadColor.FillStyle = 7
    Else
        l_PadColor.Enabled = True
        b_PadColor.Enabled = True
        l_Vertical.Enabled = True
        l_Horizontal.Enabled = True
        cmb_Vertical.Enabled = True
        cmb_Horizontal.Enabled = True
        shape_PadColor.FillStyle = 0
    End If
End Sub


Private Sub cmb_Type_Click()
    process_check
End Sub

Private Sub chk_Grow_Click()
    process_check
End Sub

Private Sub chk_Margins_Click()
    process_check
End Sub

Private Sub chk_Thumbnail_Click()
    process_check
End Sub

Private Sub chk_Pad_Click()
    process_check
End Sub

Private Sub chk_Rotate_Click()
    process_check
End Sub

Private Sub chk_Shrink_Click()
    process_check
End Sub

Private Sub f_TopMargin_GotFocus()
    f_TopMargin.SelStart = 0
    f_TopMargin.SelLength = Len(f_TopMargin.Text)
End Sub
Private Sub f_BottomMargin_GotFocus()
    f_BottomMargin.SelStart = 0
    f_BottomMargin.SelLength = Len(f_BottomMargin.Text)
End Sub
Private Sub f_RightMargin_GotFocus()
    f_RightMargin.SelStart = 0
    f_RightMargin.SelLength = Len(f_RightMargin.Text)
End Sub
Private Sub f_LeftMargin_GotFocus()
    f_LeftMargin.SelStart = 0
    f_LeftMargin.SelLength = Len(f_LeftMargin.Text)
End Sub

Private Sub f_Width_GotFocus()
    f_Width.SelStart = 0
    f_Width.SelLength = Len(f_Width.Text)
End Sub

Private Sub f_Height_GotFocus()
    f_Height.SelStart = 0
    f_Height.SelLength = Len(f_Height.Text)
End Sub

Private Sub f_tWidth_GotFocus()
    f_tWidth.SelStart = 0
    f_tWidth.SelLength = Len(f_tWidth.Text)
End Sub

Private Sub f_tHeight_GotFocus()
    f_tHeight.SelStart = 0
    f_tHeight.SelLength = Len(f_tHeight.Text)
End Sub

Private Sub Form_Load()
    Call cmb_RotateDirection.AddItem("Clockwise", 0)
    Call cmb_RotateDirection.AddItem("Counter Clockwise", 1)
    
    Call cmb_Vertical.AddItem("Align Top", 0)
    Call cmb_Vertical.AddItem("Align Center", 1)
    Call cmb_Vertical.AddItem("Align Bottom", 2)
    
    Call cmb_Horizontal.AddItem("Align Left", 0)
    Call cmb_Horizontal.AddItem("Align Center", 1)
    Call cmb_Horizontal.AddItem("Align Right", 2)
    
    Call cmb_Compression.AddItem("Highest Quality ", 0)
    Call cmb_Compression.AddItem("Good Quality", 1)
    Call cmb_Compression.AddItem("OK Quality", 2)
    Call cmb_Compression.AddItem("OK Compression", 3)
    Call cmb_Compression.AddItem("Good Compression", 4)
    Call cmb_Compression.AddItem("Highest Compression", 5)
                    
End Sub

Private Sub b_PadColor_Click()
    CommonDialog1.Color = shape_PadColor.FillColor
    CommonDialog1.ShowColor
    shape_PadColor.FillColor = CommonDialog1.Color
End Sub

Private Sub b_MarginColor_Click()
    CommonDialog1.Color = shape_MarginColor.FillColor
    CommonDialog1.ShowColor
    shape_MarginColor.FillColor = CommonDialog1.Color
End Sub

