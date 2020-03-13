Attribute VB_Name = "Validate"
Option Explicit

Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: Validate.bas
Rem
Rem Description:
Rem     Contains basic validation and conversion functions
Rem
Rem -------------------------------------------------------------

Rem -------------------------------------------------------------
Rem ISODateToDate
Rem
Rem Converts an ISO formatted date string into a date variant.
Rem Will throw an exception if there is a failure in the conversion.
Rem
Public Function ISODateToDate(strISODate As String) As Date
    Dim Y, m, d, h, mi, s As Integer
    Dim a1, a2, a3
    
    If strISODate = "" Then
        ISODateToDate = CDate(0)
        Exit Function
    End If
    
    a1 = Split(strISODate, "T")
    a2 = Split(a1(0), "-")
    a3 = Split(a1(1), ":")
    Y = CInt(a2(0))
    m = CInt(a2(1))
    d = CInt(a2(2))
    h = CInt(a3(0))
    mi = CInt(a3(1))
    s = CInt(a3(2))

    ISODateToDate = DateSerial(Y, m, d) + TimeSerial(h, mi, s)
End Function

Rem -----------------------------------------------------------
Rem PadDatePart
Rem
Rem Functions similar to datepart, but will pad the resulting numbers
Rem with leading zeros as necessary
Rem
Public Function PadDatePart(strFormat As String, dtDate As Date) As String
    If DatePart(strFormat, dtDate) = 0 Then
        PadDatePart = "00"
    ElseIf DatePart(strFormat, dtDate) < 10 Then
        PadDatePart = "0" & DatePart(strFormat, dtDate)
    Else
        PadDatePart = DatePart(strFormat, dtDate)
    End If
End Function

Rem -----------------------------------------------------------
Rem DateToISODate
Rem
Rem Will return an ISO date string for a passed in date variant.
Rem Really doesn't have much of a way to fail.
Rem
Public Function DateToISODate(dtDate As Date) As String
    
    DateToISODate = DatePart("yyyy", dtDate) & "-" & PadDatePart("m", dtDate) & "-" & PadDatePart("d", dtDate) & "T" & PadDatePart("h", dtDate) & ":" & PadDatePart("n", dtDate) & ":" & PadDatePart("s", dtDate)
    
End Function

Rem ----------------------------------------------------------
Rem ValidateString
Rem
Rem Will validate the passed in value string as meeting the passed in
Rem criteria (min = minumum length, max = maximum length)
Rem
Rem Will display an error message box on failure and return false.
Rem
Public Function ValidateString(value As String, label As String, min As Integer, max As Integer) As Boolean
    If Len(value) < min Or Len(value) > max Then
        GoTo error
    End If
    
    ValidateString = True
    Exit Function
    
error:
    MsgBox "The value for " & label & " must be between " & min & " and " & max & " characters.", vbExclamation, "Error"
    ValidateString = False

End Function

Rem ----------------------------------------------------------
Rem ValidateInt
Rem
Rem Will valiate the passed in value string as meeting the passed in
Rem criteria (min = minumum value, max = maximum value)
Rem
Rem Will display an error message box on failure and return false.
Rem
Public Function ValidateInt(value As String, label As String, min As Integer, max As Integer) As Boolean
    Dim X As Long
    Dim errNumber As Long
    On Error Resume Next
    X = CLng(value)
    errNumber = Err.Number
    On Error GoTo 0
    If errNumber <> 0 Then
        GoTo error
    End If
    If X < min Or X > max Then
        GoTo error
    End If
    
    ValidateInt = True
    Exit Function
    
error:
    MsgBox "The value for " & label & " must be a number between " & min & " and " & max & ".", vbExclamation, "Error"
    ValidateInt = False
End Function
