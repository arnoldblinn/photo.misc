Attribute VB_Name = "NotifyIcon"
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: NotifyIcon.bas
Rem
Rem Description:
Rem     Contains the function defines and contanst for dealing with
Rem     the winapi for manipulating icons on the system tray.
Rem
Rem -------------------------------------------------------------

Rem Declare a user-defined variable to pass to the Shell_NotifyIcon
Rem function.
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
 End Type

Rem Declare the constants for the API function. These constants can be
Rem found in the header file Shellapi.h.

Rem The following constants are the messages sent to the
Rem Shell_NotifyIcon function to add, modify, or delete an icon from the
Rem taskbar status area.
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Rem The following constants are the flags that indicate the valid
Rem members of the NOTIFYICONDATA data type.
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Rem Declare the API function call.
Public Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Rem The following constant is the message sent when a mouse event occurs
Rem within the rectangular boundaries of the icon in the taskbar status
Rem area.
Public Const WM_MOUSEMOVE = &H200
Public Const WM_DROPFILES = &H233

Rem The following constants are used to determine the mouse input on the
Rem the icon in the taskbar status area.

Rem Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Rem Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

' Declaration of a Windows routine.
' This statement should be placed in the module.
Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, ByVal _
  hWndInsertAfter As Long, ByVal X As Long, ByVal Y As _
  Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags _
  As Long) As Long

' Set some constant values (from WIN32API.TXT).
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40
Public Const conSwpNoSize = &H1
Public Const conSwpNoMove = &H2

Public Declare Function PostMessage Lib "user32" _
   Alias "PostMessageA" _
   (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long
   
Public Const WH_KEYBOARD = 2
Public Const KBH_MASK = &H20000000




