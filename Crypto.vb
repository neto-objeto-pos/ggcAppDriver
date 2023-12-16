'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Encryption Object
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
'  XerSys [ 12/28/2012 12:06 pm ]
'      Start translating of this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Public Class Crypto
   Private p_sInBuffer As String
   Private p_sOutBuffer As String
    Private p_sSignature As String

   Public Property InBuffer As String
      Get
         Return p_sInBuffer
      End Get
      Set(ByVal value As String)
         p_sInBuffer = value
      End Set
   End Property

   Public Property OutBuffer As String
      Get
         Return p_sOutBuffer
      End Get
      Set(ByVal value As String)
         p_sOutBuffer = value
      End Set
   End Property

   Public Property Signature() As String
      Get
         Return p_sSignature
      End Get
      Set(ByVal value As String)
         p_sSignature = value
      End Set
   End Property

   Public Sub New()
      p_sSignature = "Beyond the boundary of imagination, what"
   End Sub

   Public Sub New(ByVal fsSignature As String)
      p_sSignature = fsSignature
   End Sub

   Public Function Encrypt() As String
      Call RC4()
      Return p_sOutBuffer
   End Function

   Public Function Decrypt() As String
      Call RC4()
      Return p_sOutBuffer
   End Function

   Private Sub RC4()
      Dim S(0 To 255) As Byte, K(0 To 255) As Byte, i As Long
      Dim j As Long, temp As Byte, y As Byte, t As Long, x As Long
      Dim Outp As String

      For i = 0 To 255
         S(i) = i
      Next

      j = 1
      For i = 0 To 255
         If j > Len(p_sSignature) Then j = 1
         K(i) = Asc(Mid$(p_sSignature, j, 1))
         j = j + 1
      Next i

      j = 0
      For i = 0 To 255
         j = (j + S(i) + K(i)) Mod 256
         temp = S(i)
         S(i) = S(j)
         S(j) = temp
      Next i

      i = 0
      j = 0
      Outp = ""
      For x = 1 To Len(InBuffer)
         i = (i + 1) Mod 256
         j = (j + S(i)) Mod 256
         temp = S(i)
         S(i) = S(j)
         S(j) = temp
         t = (S(i) + (S(j) Mod 256)) Mod 256
         y = S(t)

         Outp = Outp & Chr(Asc(Mid(InBuffer, x, 1)) Xor y)
      Next
      p_sOutBuffer = Outp
   End Sub
End Class
