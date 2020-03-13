Attribute VB_Name = "DOM"
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: ImageMove.bas
Rem
Rem Description:
Rem     Contains code loading and saving pieces of XML to and
Rem     from a DOM object.
Rem
Rem     Valid for <Tasks>, <Album>, and <Task>
Rem
Rem -------------------------------------------------------------

Rem Derived global settings
Global g_strSchema As String

Rem The two main xml documents containing all the tasks and all the saved formats
Global g_objXMLDomNodeTasks As IXMLDOMNode
Global g_objXMLDomNodeFormatProfiles As IXMLDOMNode

Rem Functions to save the tasks and formats
Public Sub SaveTasks()
    Call SaveDOMToFile(g_strExeDir & "gptasks.xml", g_objXMLDomNodeTasks.XML)
End Sub

Public Sub SaveFormatProfiles()
    Call SaveDOMToFile(g_strExeDir & "gpformats.xml", g_objXMLDomNodeFormatProfiles.XML)
End Sub

Public Function LoadTasks() As IXMLDOMNode
    Dim objXMLDocumentWrapper As MSXML.DOMDocument
    Set objXMLDocumentWrapper = LoadDOMFromFile(g_strExeDir & "gptasks.xml", "Tasks", "1")
    
    Set LoadTasks = objXMLDocumentWrapper.childNodes.Item(0)
End Function

Public Function LoadFormatProfiles() As IXMLDOMNode
    Dim objXMLDocumentWrapper As MSXML.DOMDocument
    Set objXMLDocumentWrapper = LoadDOMFromFile(g_strExeDir & "gpformats.xml", "FormatProfiles", "1")
    
    Set LoadFormatProfiles = objXMLDocumentWrapper.childNodes.Item(0)
End Function

Public Function LoadDOMFromString(strXML As String, strTag As String, strVersion As String) As MSXML.DOMDocument
    Dim fso As Scripting.FileSystemObject
    Dim objTextStream As TextStream
    Dim strTempName As String
    Dim strSaveXML As String
    Dim objDomDocument As MSXML.DOMDocument
    
    Set objDomDocument = New MSXML.DOMDocument
    
        
    Rem I've found on some machines the DOM simply can't deal with a schema from a path
    strSaveXML = Replace(strXML, "<" & strTag, "<" & strTag & " xmlns=""" & g_strSchema & """ ")
    
    Set fso = New Scripting.FileSystemObject
    strTempName = fso.GetTempName
    Set objTextStream = fso.OpenTextFile(strTempName, ForWriting, True)
    objTextStream.Write (strSaveXML)
    objTextStream.Close
    
    objDomDocument.Load (strTempName)
    
    fso.DeleteFile (strTempName)
    
    Rem Validate the object
    If objDomDocument.parseError.errorCode <> 0 Then
        Set objDomDocument = Nothing
    ElseIf objDomDocument.childNodes.Item(0).nodeName <> strTag Then
        Set objDomDocument = Nothing
    ElseIf (strVersion <> "") Then
        If CInt(objDomDocument.childNodes.Item(0).attributes(1).nodeValue) > CInt(strVersion) Then
            Set objDomDocument = Nothing
        End If
    End If
    
    Set LoadDOMFromString = objDomDocument
End Function

Public Function LoadDOMFromFile(strFile As String, strTag As String, strVersion As String) As MSXML.DOMDocument
    Dim fso As Scripting.FileSystemObject
    Dim objTextStream As Scripting.TextStream
    Dim strXML As String
    Dim strBlank As String
    
    Rem Form the "blank" version of the data (which may not be valid)
    If strVersion <> "" Then
        strBlank = "<" & strTag & " Version=""" & strVersion & """></" & strTag & ">"
    Else
        strBlank = "<" & strTag & "></" & strTag & ">"
    End If
    
    Rem Open the file
    Set fso = New Scripting.FileSystemObject
    Set objTextStream = fso.OpenTextFile(strFile, ForReading, True)
    
    If Not objTextStream.AtEndOfStream Then
        strXML = objTextStream.ReadAll
        If strXML = "" Then
            strXML = strBlank
        End If
    Else
        strXML = strBlank
    End If
            
    Set LoadDOMFromFile = LoadDOMFromString(strXML, strTag, strVersion)
End Function

Public Function SaveDOMToFile(strFile As String, strXML As String)
    Dim fso As Scripting.FileSystemObject
    Dim objTextStream As TextStream
    Dim X As String
        
    Set fso = New Scripting.FileSystemObject
    Set objTextStream = fso.OpenTextFile(strFile, ForWriting, True)
    X = Replace(strXML, " xmlns=""" & g_strSchema & """", "")
    X = Replace(X, "xmlns=""""", "")
    Call objTextStream.Write(X)
End Function




