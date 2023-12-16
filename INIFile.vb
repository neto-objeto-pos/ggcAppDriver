'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Ini File reader/writer Object
'
' Copyright 2012 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 12/28/2012 12:05 pm ]
'      Start translating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports System.IO

Public Class INIFile
   ' API functions
   Private Declare Ansi Function GetPrivateProfileString _
     Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
     (ByVal lpApplicationName As String, _
     ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As System.Text.StringBuilder, _
     ByVal nSize As Integer, ByVal lpFileName As String) _
     As Integer
   Private Declare Ansi Function WritePrivateProfileString _
     Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
     (ByVal lpApplicationName As String, _
     ByVal lpKeyName As String, ByVal lpString As String, _
     ByVal lpFileName As String) As Integer
   Private Declare Ansi Function GetPrivateProfileInt _
     Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
     (ByVal lpApplicationName As String, _
     ByVal lpKeyName As String, ByVal nDefault As Integer, _
     ByVal lpFileName As String) As Integer
   Private Declare Ansi Function FlushPrivateProfileString _
     Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
     (ByVal lpApplicationName As Integer, _
     ByVal lpKeyName As Integer, ByVal lpString As Integer, _
     ByVal lpFileName As String) As Integer

   Private p_sFileName As String

   ' Constructor, accepting a filename
   Public Sub New(Optional ByVal Filename As String = "")
      p_sFileName = Filename
   End Sub

   ' Read-only filename property
   Public Property FileName As String
      Get
         Return p_sFileName
      End Get
      Set(ByVal value As String)
         p_sFileName = value
      End Set
   End Property

   Public Function IsFileExist() As Boolean
      Return File.Exists(p_sFileName)
   End Function

   Public Function GetTextValue(ByVal fsSection As String, _
                                ByVal fsKey As String, _
                                Optional ByVal fsDefault As String = "") As String
      ' Returns a string from your INI file
      Dim lnCharCount As Integer
      Dim loResult As New System.Text.StringBuilder(256)

      lnCharCount = GetPrivateProfileString(fsSection, fsKey, _
         fsDefault, loResult, loResult.Capacity, p_sFileName)
      If lnCharCount > 0 Then
         Return Left(loResult.ToString, lnCharCount)
      Else
         Return ""
      End If
   End Function

   Public Function GetNumericValue(ByVal fsSection As String, _
                                   ByVal fsKey As String, _
                                   Optional ByVal fsDefault As Integer = 0) As Integer
      ' Returns an integer from your INI file
      Return GetPrivateProfileInt(fsSection, fsKey, fsDefault, p_sFileName)
   End Function

   Public Function GetBooleanValue(ByVal fsSection As String, _
                                   ByVal fsKey As String, _
                                   Optional ByVal fsDefault As Boolean = False) As Boolean
      ' Returns a boolean from your INI file
      Return (GetPrivateProfileInt(fsSection, fsKey, CInt(fsDefault), p_sFileName) = 1)
   End Function

   Public Sub SetStringValue(ByVal fsSection As String, ByVal fsKey As String, ByVal fsValue As String)
      ' Writes a string to your INI file
      WritePrivateProfileString(fsSection, fsKey, fsValue, p_sFileName)
      Flush()
   End Sub

   Public Sub SetNumericValue(ByVal fsSection As String, ByVal fsKey As String, ByVal fsValue As Integer)
      ' Writes an integer to your INI file
      SetStringValue(fsSection, fsKey, CStr(fsValue))
      Flush()
   End Sub

   Public Sub SetBooleanValue(ByVal fsSection As String, ByVal fsKey As String, ByVal fsValue As Boolean)
      ' Writes a boolean to your INI file
      SetStringValue(fsSection, fsKey, CStr(CInt(fsValue)))
      Flush()
   End Sub

   Private Sub Flush()
      ' Stores all the cached changes to your INI file
      FlushPrivateProfileString(0, 0, 0, p_sFileName)
   End Sub
End Class
