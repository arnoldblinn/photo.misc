VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDigiFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: clsDigiFrame.cls
Rem
Rem Description:
Rem     Contains code for talking to the digi-frame.
Rem
Rem -------------------------------------------------------------

Rem Global variables
Dim giPort As Integer
Dim giCard As Integer
Dim gfOnline As Boolean
Dim gfFrame As Boolean
Dim SI_COMM1 As SI_COMM

Rem ---------------------------------------------
Rem Delay
Rem
Rem Utility function to delay for the passed in number of seconds
Rem
Private Sub Delay(s As Integer)
    Dim dtStart As Date
    
    dtStart = Now
    While DateDiff("s", dtStart, Now) < s
    Wend

End Sub

Rem -------------------------------------------
Rem WaitPrompt
Rem
Rem Will wait for the frame to respond with its prompt, returning
Rem the string from the frame.
Rem
Private Function WaitPrompt() As String
    Dim dtStart As Date
    Dim c As Integer
    Dim strResult As String
        
    Rem Wait a few seconds before getting data
    Call Delay(3)
    
    Rem Wait for the frame to respond with it's prompt
    dtStart = Now
    
    strResult = ""
    While DateDiff("s", dtStart, Now) < 15
        Call SI_COMM1.GetByte(c)
        strResult = strResult & CStr(Chr(c))
        If Right(strResult, 3) = "DF>" Then
            GoTo done
        End If
    Wend
    
    GoTo error
    
done:
    Call Delay(3)
    
    WaitPrompt = strResult
    Exit Function
    
error:
    WaitPrompt = ""
End Function

Rem ----------------------------------------------
Rem Class_Initialize
Rem
Rem Called to initialize the class.  Because of the implementation
Rem of the serial package, it must have a pointer to the OCX created
Rem in the frmTasks
Rem
Private Sub Class_Initialize()
    gfOnline = False
    gfFrame = False
    Set SI_COMM1 = frmMain.SI_COMM1
End Sub

Private Sub Class_Terminate()
    Call Disconnect
End Sub

Rem ----------------------------------------------
Rem Property: Port
Rem
Rem This is the COM port with which the digi frame is connected to.  Possible values
Rem are 0 (port 1) through 3 (PORT 4)
Rem
Public Property Let Port(iPort As Integer)
    giPort = iPort
End Property

Rem ----------------------------------------------
Rem Property: Card
Rem
Rem This is the card with which the digi frame should communicate with. 0 is
Rem compact flash, 1 is smart media
Rem
Public Property Let Card(iCard As Integer)
    giCard = iCard
End Property

Rem ----------------------------------------------
Rem Connect
Rem
Rem Will connect to the digi-frame using the current port.
Rem
Public Function Connect() As Boolean
    Dim strData As String
    Connect = False
    
    If gfOnline = True Then
        Exit Function
    End If
On Error GoTo error
    
    SI_COMM1.BaudRate = Baud1152k
    SI_COMM1.Parity = ParityNone
    SI_COMM1.StopBits = StopBit1
    SI_COMM1.WordLength = WordLen8
    SI_COMM1.FlowControl = NoHandShaking
    SI_COMM1.PortNumber = giPort
    
    SI_COMM1.CardType = CardWinApi
    SI_COMM1.PortOpen = True
    
    SI_COMM1.OnParityErr = False
    SI_COMM1.OnFrameErr = False
    SI_COMM1.OnOverRunErr = False
    SI_COMM1.rxAlertCount = 0
    SI_COMM1.AsyncMode = False
    gfOnline = True

    Rem Let's try to talk to the frame. This can fail, so we try a few times....
    Dim i As Integer
    i = 0
    While i < 3
    
        Delay (2)
        SI_COMM1.PutString (vbCr)
        
        strData = WaitPrompt
        If strData <> "" Then
            gfFrame = True
            Connect = True
            Exit Function
        End If
        

        i = i + 1
    Wend

error:
    Connect = False
    Exit Function
End Function

Rem ----------------------------------------------
Rem Disconnect
Rem
Rem Will disconnect the clsDigiFrame from the digital picture frame
Rem
Public Function Disconnect() As Boolean
    
On Error GoTo error
    
    If gfFrame Then
        SI_COMM1.PutString ("QUIT" & vbCr)
    
        Call Delay(3)
        
        gfFrame = False
    End If
    
    If gfOnline Then
        SI_COMM1.PortOpen = False
    
        gfOnline = False
    End If
    
    Exit Function
    
error:
    Exit Function
End Function

Rem ----------------------------------------------
Rem PutFile
Rem
Rem Will put the passed into file onto the frame as the passed in
Rem file name.  Note for the file to actually be guaranteed to be this name
Rem once put onto the frame the frame must have firmware version k or better.
Rem
Public Function PutFile(strFrameFile As String, strPCFile As String, fUnlink As Boolean) As Boolean
    Dim strCommand As String
    Dim strCardCommand As String
    Dim strUnlinkDrive As String
    Dim iResult
        
On Error GoTo error

    If giCard = 0 Then
        strCardCommand = "CF"
        strUnlinkDrive = "C"
    Else
        strCardCommand = "SS"
        strUnlinkDrive = "S"
    End If
        
    If fUnlink Then
        
        If "" = Command("UNLINK " & strUnlinkDrive & "/" & strFrameFile) Then
            GoTo error
        End If
        
        If "" = Command("UNLINK " & strUnlinkDrive & "/" & Left(strFrameFile, Len(strFrameFile) - 3) & ".JXX") Then
            GoTo error
        End If
    End If

    strCommand = strCardCommand & "PUT " & strFrameFile & vbCr
        
    Dim i As Integer
    
    i = 0
    While i < 3
        SI_COMM1.PutString (strCommand)
        Delay (1)
        iResult = SI_COMM1.SendFile(strPCFile, Xmodem)
        If iResult = 0 Then
            PutFile = True
            Exit Function
        End If
        
        i = i + 1
        Delay (2)
    Wend
    
error:
    PutFile = False
    Exit Function
            
End Function

Rem ----------------------------------------------
Rem GetFile
Rem
Rem Currently unimplemented
Rem
Public Function GetFile(strFrameFile, strPCFile) As Boolean
    GetFile = False
End Function

Rem -----------------------------------------------
Rem AutoDetect
Rem
Rem Will detect the frame, returning the port and card ids for the
Rem currently attached frame.
Rem
Rem Returns true if detected, false if not
Rem
Public Function AutoDetect(ByRef iPort As Integer, ByRef iCard As Integer) As Boolean
    Dim i As Integer
    
    AutoDetect = False
    
    For i = 0 To 3
        giPort = i
        If Connect() = True Then
            iPort = i

            If InStr(Command("CFDIR"), "NO CARD") = 0 Then
                iCard = 0
                AutoDetect = True
            Else
                If InStr(Command("SSDIR"), "NO CARD") = 0 Then
                    iCard = 1
                    AutoDetect = True
                End If
            End If
            Disconnect
            Exit Function
        End If
    Next
    
    AutoDetect = False
End Function



Rem -----------------------------------------------
Rem Command
Rem
Rem Will pass a command to the frame, returning the
Rem data back from the frame.
Rem
Public Function Command(strCommand As String) As String
    Dim strData As String
    
On Error GoTo error
    SI_COMM1.PutString (strCommand & vbCr)
    
    strData = WaitPrompt()
    
    Command = strData
    
    Exit Function
error:
    Command = ""
    Exit Function
End Function

