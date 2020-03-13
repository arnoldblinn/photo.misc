Attribute VB_Name = "modMySaver"
'MySaver.bas
Option Explicit

'Rectangle data structure
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Constants for some API functions
Private Const WS_CHILD = &H40000000
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const HWND_TOP = 0&
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

'--- API functions
Private Declare Function GetClientRect _
Lib "user32" ( _
    ByVal hwnd As Long, _
    lpRect As RECT _
) As Long

Private Declare Function GetWindowLong _
Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long _
) As Long

Private Declare Function SetWindowLong _
Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

Private Declare Function SetWindowPos _
Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long _
) As Long

Private Declare Function SetParent _
Lib "user32" ( _
    ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long _
) As Long

'Global Show/Preview flag
Public gblnShow As Boolean

'Module level variables
Private mlngDisplayHwnd As Long
Private recDisplay As RECT
    
'Starting point....
Public Sub Main()
    Dim strCmd As String
    Dim strTwo As String
    Dim lngStyle As Long
    Dim lngPreviewHandle As Long
    Dim lngParam As Long
    'Process the command line
    strCmd = UCase(Trim(Command))
    strTwo = Left(strCmd, 2)
    Select Case strTwo
    'Preview screen saver in small display window
    Case "/P"
        'Get HWND of display window
        mlngDisplayHwnd = Val(Mid(strCmd, 4))
        'Get display rectangle dimensions
        GetClientRect mlngDisplayHwnd, recDisplay
        'Load form for preview
        gblnShow = False
        Load frmMySaver
        'Get HWND for display form
        lngPreviewHandle = frmMySaver.hwnd
        'Get current window style
        lngStyle = GetWindowLong(lngPreviewHandle, GWL_STYLE)
        'Append "WS_CHILD" style to the current window style
        lngStyle = lngStyle Or WS_CHILD
        'Add new style to display window
        SetWindowLong lngPreviewHandle, GWL_STYLE, lngStyle
        'Set display window as parent window
        SetParent lngPreviewHandle, mlngDisplayHwnd
        'Save the parent hWnd in the display form's window structure.
        SetWindowLong lngPreviewHandle, GWL_HWNDPARENT, mlngDisplayHwnd
        'Preview screensaver in the window...
        SetWindowPos lngPreviewHandle, _
            HWND_TOP, 0&, 0&, recDisplay.Right, recDisplay.Bottom, _
            SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
        Exit Sub
    'Allow user to setup screen saver
    Case "/C"
        Load frmMySetup
        Exit Sub
    'Password - not implemented here
    Case "/A"
        MsgBox "No password is necessary for this Screen Saver", _
                vbInformation, "Password Information"
        Exit Sub
    'Show screen saver in normal full screen mode
    Case "/S"
        gblnShow = True
        Load frmMySaver
        frmMySaver.Show
        Exit Sub
    'Unknown command line parameters
    Case Else
        Exit Sub
    End Select
End Sub


