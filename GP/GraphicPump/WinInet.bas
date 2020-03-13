Attribute VB_Name = "WinInet"
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: wininet.bas
Rem
Rem Description:
Rem     Contains defines for functions and constants in the wininet
Rem     dll.
Rem
Rem -------------------------------------------------------------

Public Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Public Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Public Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4

Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8

Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lFlags As Long, ByVal lReserved As Long) As Boolean
Declare Function InternetAutodial Lib "wininet.dll" (ByVal lFlags As Long, ByVal lReserved As Long) As Boolean
Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal lReserved As Long) As Boolean
