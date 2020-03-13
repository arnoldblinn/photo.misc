VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSchedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ControlBox      =   0   'False
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7740
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton b_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Schedule"
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7455
      Begin VB.ComboBox cmb_Type 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox chk_Disable 
         Caption         =   "Disable"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox chk_Connect 
         Caption         =   "Force Connection if necessary"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Frame frame_Hourly 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1215
         Left            =   3720
         TabIndex        =   23
         Top             =   360
         Width           =   3255
         Begin ComCtl2.DTPicker time_Hourly 
            Height          =   315
            Left            =   480
            TabIndex        =   1
            Top             =   120
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "mm"
            Format          =   22675459
            UpDown          =   -1  'True
            CurrentDate     =   36806
         End
         Begin VB.Label Label1 
            Caption         =   "At"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "minutes past the hour"
            Height          =   255
            Left            =   1200
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame frame_Daily 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1095
         Left            =   3720
         TabIndex        =   20
         Top             =   360
         Width           =   3495
         Begin ComCtl2.DTPicker time_Daily 
            Height          =   315
            Left            =   480
            TabIndex        =   4
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22675458
            CurrentDate     =   36806
         End
         Begin VB.Label Label3 
            Caption         =   "At"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "every day"
            Height          =   255
            Left            =   2280
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame frame_Weekly 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1215
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   3495
         Begin VB.ComboBox c_Weekday 
            Height          =   315
            ItemData        =   "frmSchedule.frx":0442
            Left            =   840
            List            =   "frmSchedule.frx":045B
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   600
            Width           =   1935
         End
         Begin ComCtl2.DTPicker time_Weekly 
            Height          =   315
            Left            =   480
            TabIndex        =   2
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22675458
            CurrentDate     =   36806
         End
         Begin VB.Label Label5 
            Caption         =   "At"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label6 
            Caption         =   "every week"
            Height          =   255
            Left            =   2280
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "on"
            Height          =   255
            Left            =   480
            TabIndex        =   17
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame frame_Monthly 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1215
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   3495
         Begin VB.TextBox f_Monthday 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   600
            Width           =   735
         End
         Begin ComCtl2.DTPicker time_Monthly 
            Height          =   315
            Left            =   480
            TabIndex        =   5
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22675458
            CurrentDate     =   36806
         End
         Begin VB.Label Label8 
            Caption         =   "At"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "of every month"
            Height          =   255
            Left            =   2040
            TabIndex        =   14
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "on day"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   3480
         X2              =   3480
         Y1              =   240
         Y2              =   2760
      End
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem --------------------------------------------
Rem Boilerplate functionality for every form
Rem
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

Rem -----------------------------------------------------
Rem Stub functions same in every form but specific to this form
Rem
Private Function XMLToUI()
    Dim t As Integer
    Dim iHours, iMinutes As Long
    
    t = CInt(gXMLDomNodeClone.childNodes(XMLISchedule_Type).Text)
    cmb_Type.ListIndex = t
    
    iHours = CInt(gXMLDomNodeClone.childNodes(XMLISchedule_Hours).Text)
    iMinutes = CInt(gXMLDomNodeClone.childNodes(XMLISchedule_Minutes).Text)
    
    time_Daily.Hour = iHours
    time_Daily.Minute = iMinutes
    time_Weekly.Hour = iHours
    time_Weekly.Minute = iMinutes
    time_Monthly.Hour = iHours
    time_Monthly.Minute = iMinutes
    
    c_Weekday.ListIndex = CInt(gXMLDomNodeClone.childNodes(XMLISchedule_Weekday).Text)
    f_Monthday.Text = gXMLDomNodeClone.childNodes(XMLISchedule_Monthday).Text
    
    chk_Disable.value = CInt(gXMLDomNodeClone.childNodes(XMLISchedule_Disable).Text)
    chk_Connect.value = CInt(gXMLDomNodeClone.childNodes(XMLISchedule_Connect).Text)
End Function

Private Function UIToXML() As Boolean
    Dim t As Integer
    Dim iHours, iMinutes As Integer
    Dim iWeekday As Integer
    Dim iMonthday As Integer
    
    iHours = 0
    iMinutes = 0
    iWeekday = 0
    iMonthday = 1
    t = cmb_Type.ListIndex
    If t = XMLISchedule_Type_Hourly Then
        iMinutes = time_Hourly.Minute
    ElseIf t = XMLISchedule_Type_Daily Then
        iMinutes = time_Daily.Minute
        iHours = time_Daily.Hour
    ElseIf t = XMLISchedule_Type_Weekly Then
        iMinutes = time_Weekly.Minute
        iHours = time_Weekly.Hour
        iWeekday = c_Weekday.ListIndex
    ElseIf t = XMLISchedule_Type_Monthly Then
        iMinutes = time_Monthly.Minute
        iHours = time_Monthly.Hour
        If ValidateInt(f_Monthday, "Day of month", 1, 28) = False Then
            Exit Function
        End If
        iMonthday = CInt(f_Monthday.Text)
    End If
           
    gXMLDomNodeClone.childNodes(XMLISchedule_Type).Text = CStr(t)
    gXMLDomNodeClone.childNodes(XMLISchedule_Hours).Text = CStr(iHours)
    gXMLDomNodeClone.childNodes(XMLISchedule_Minutes).Text = CStr(iMinutes)
    gXMLDomNodeClone.childNodes(XMLISchedule_Weekday).Text = CStr(iWeekday)
    gXMLDomNodeClone.childNodes(XMLISchedule_Monthday).Text = CStr(iMonthday)
    
    gXMLDomNodeClone.childNodes(XMLISchedule_Disable).Text = chk_Disable.value
    gXMLDomNodeClone.childNodes(XMLISchedule_Connect).Text = chk_Connect.value
    
    UIToXML = True
End Function

Rem ---------------------------------------------------
Rem Functionality specific to this form
Rem
Private Sub show_frame(t As Integer)
    frame_Hourly.Visible = CBool(t = XMLISchedule_Type_Hourly)
    frame_Daily.Visible = CBool(t = XMLISchedule_Type_Daily)
    frame_Weekly.Visible = CBool(t = XMLISchedule_Type_Weekly)
    frame_Monthly.Visible = CBool(t = XMLISchedule_Type_Monthly)
    
    chk_Connect.Visible = CBool(t <> XMLISchedule_Type_None)
    chk_Disable.Visible = CBool(t <> XMLISchedule_Type_None)
End Sub

Private Sub chk_Disable_Click()
    Dim fEnabled As Boolean
    
    fEnabled = Not CBool(chk_Disable.value)
    chk_Connect.Enabled = fEnabled
    cmb_Type.Enabled = fEnabled
    
    time_Hourly.Enabled = fEnabled
    time_Daily.Enabled = fEnabled
    time_Weekly.Enabled = fEnabled
    time_Monthly.Enabled = fEnabled
    
    c_Weekday.Enabled = fEnabled
    f_Monthday.Enabled = fEnabled
    
    Label1.Enabled = fEnabled
    Label2.Enabled = fEnabled
    Label3.Enabled = fEnabled
    Label4.Enabled = fEnabled
    Label5.Enabled = fEnabled
    Label6.Enabled = fEnabled
    Label7.Enabled = fEnabled
    Label8.Enabled = fEnabled
    Label9.Enabled = fEnabled
    Label10.Enabled = fEnabled
    
End Sub

Private Sub f_Monthday_GotFocus()
    f_Monthday.SelStart = 0
    f_Monthday.SelLength = Len(f_Monthday.Text)
End Sub

Private Sub Form_Load()
    Call c_Weekday.AddItem("Sunday", XMLISchedule_Weekday_Sunday)
    Call c_Weekday.AddItem("Monday", XMLISchedule_Weekday_Monday)
    Call c_Weekday.AddItem("Tuesday", XMLISchedule_Weekday_Tuesday)
    Call c_Weekday.AddItem("Wednesday", XMLISchedule_Weekday_Wednesday)
    Call c_Weekday.AddItem("Thursday", XMLISchedule_Weekday_Thursday)
    Call c_Weekday.AddItem("Friday", XMLISchedule_Weekday_Friday)
    Call c_Weekday.AddItem("Saturday", XMLISchedule_Weekday_Saturday)
    
    Call cmb_Type.AddItem("None", XMLISchedule_Type_None)
    Call cmb_Type.AddItem("Hourly", XMLISchedule_Type_Hourly)
    Call cmb_Type.AddItem("Daily", XMLISchedule_Type_Daily)
    Call cmb_Type.AddItem("Weekly", XMLISchedule_Type_Weekly)
    Call cmb_Type.AddItem("Monthly", XMLISchedule_Type_Monthly)
    
End Sub

Private Sub cmb_Type_Click()
    show_frame (cmb_Type.ListIndex)
End Sub
