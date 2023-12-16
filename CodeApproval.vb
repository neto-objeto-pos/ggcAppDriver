'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Code Approval Object
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
' Note: This object is a translation of the VB6 project - ggcCodeApproval, to VB.Net.
' ==========================================================================================
'  Kalyptus [ 04/29/2016 02:20 pm ]
'      Started creating this object.
'  Kalyptus [ 06/10/2016 02:26 pm ]
'      For the past 2-3 weeks we have added the following Code Approvals:
'           pxeMCIssuance As String = "MI"
'           pxeMCClusteringDelivery As String = "CD"
'           pxeFSEPActivation As String = "FA"
'           pxeFSEXActivation As String = "FX"
'           pxeMPDiscount As String = "MD"
'           pxePreApproved As String = "PA"
'   Note: As I try to analyze the categories of different approval code, their comonalities
'         seems divided into the following categories:
'           DAY2DAY
'           MANUALX
'           Misc is based on the Reference No of the transaction either [trans no/doc no]
'               REFER01
'           Misc is based on Client Name
'               FLLNME1
'           Just like FLLNME1 but with no branch request, and date 
'               DATE001 
'               DATE003
'               DATE060
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Public Class CodeApproval
    Shared randm As New Random()
    Private poRawxxx As xObject
    Private poResult As xObject
    Private pbEncode As Boolean
    Private pbReferx As Boolean

    Public Const pxeDay2Day As String = "DT"
    Public Const pxeManualLog As String = "ML"
    Public Const pxeForgot2Log As String = "FL"
    Public Const pxeBusinessTrip As String = "OB"
    Public Const pxeBusinessTripWLog As String = "OL"
    Public Const pxeLeave As String = "LV"
    Public Const pxeOvertime As String = "OT"
    Public Const pxeShift As String = "SH"
    Public Const pxeDayOff As String = "DO"
    Public Const pxeTardiness As String = "TD"
    Public Const pxeUnderTime As String = "UD"
    Public Const pxeCreditInvestigation As String = "CI"
    Public Const pxeCreditApplication As String = "CA"
    Public Const pxeWholeSaleDiscount As String = "WD"      'Spareparts
    Public Const pxeCashBalance As String = "CB"
    Public Const pxeOfficeRebate As String = "R1"
    Public Const pxeFieldRebate As String = "R2"
    Public Const pxePartsDiscount As String = "SI"          'Spareparts 
    Public Const pxeMCDiscount As String = "DR"
    Public Const pxeSPPurcDelivery As String = "PD"         'Spareparts
    Public Const pxeIssueORNotPR As String = "OR"
    Public Const pxeIssueORNotSI As String = "OX"
    Public Const pxeAdditional As String = "RS"
    Public Const pxeBiyahingFiesta As String = "BF"
    Public Const pxeTeleMktg As String = "TM"
    Public Const pxeMCIssuance As String = "MI"
    Public Const pxeMCClusteringDelivery As String = "CD"
    Public Const pxeFSEPActivation As String = "FA"         'Spareparts
    Public Const pxeFSEXActivation As String = "FX"         'Spareparts
    Public Const pxeMPDiscount As String = "MD"
    Public Const pxePreApproved As String = "PA"
    Public Const pxeJobOrderWOGCard As String = "JG"        'Spareparts

    Public ReadOnly Property Result() As String
        Get
            If pbEncode Then
                If poRawxxx.Branch <> "" Then
                    'Result = MMMIBBSSDD
                    Return poResult.Misc & poResult.IssuedBy & poResult.Branch & poResult.XSystem & poResult.XDate
                Else
                    'Result = MMMISSDDDD
                    Return poResult.Misc & poResult.IssuedBy & poResult.XSystem & poResult.XDate
                End If
            Else
                Return ""
            End If
        End Get
    End Property

    Public WriteOnly Property Branch() As String
        Set(ByVal fsBranchCD As String)
            poRawxxx.Branch = fsBranchCD
        End Set
    End Property

    Public WriteOnly Property DateRequested() As Date
        Set(ByVal fdRequestxx As Date)
            poRawxxx.XDate = fdRequestxx
        End Set
    End Property

    Public WriteOnly Property IssuedBy() As String
        Set(ByVal fcIssuedBy As String)
            poRawxxx.IssuedBy = fcIssuedBy
        End Set
    End Property

    Public WriteOnly Property XSystem() As String
        Set(ByVal fsSystem As String)
            poRawxxx.XSystem = fsSystem
        End Set
    End Property

    Public WriteOnly Property MiscInfo() As String
        Set(ByVal fsValue As String)
            poRawxxx.Misc = fsValue
        End Set
    End Property

    Public WriteOnly Property IsByRef() As Boolean
        Set(ByVal fbValue As Boolean)
            pbReferx = fbValue
        End Set
    End Property

    Public Function Encode() As Boolean
        pbEncode = False

        'Verify System Approval Requested
        If poRawxxx.XSystem = "" Then
            MsgBox("Invalid System Approval Request detected!", vbCritical + vbOKOnly, "Verification")
            Return False
        End If

        'Verify date requested
        If poRawxxx.XDate = "" Then
            MsgBox("Invalid Date Requested detected!", vbCritical + vbOKOnly, "Verification")
            Return False
        End If

        'Verify issuing Department/Person
        If poRawxxx.IssuedBy = "" Then
            MsgBox("Invalid Issuing Department/Person detected!", vbCritical + vbOKOnly, "Verification")
            Return False
        End If

        Select Case poRawxxx.XSystem
            Case CodeApproval.pxeManualLog
                'Misc should be the binary equivalent of the periods approved...
                If Not IsNumeric(poRawxxx.Misc) Then
                    MsgBox("Invalid Reference Number detected!", vbCritical + vbOKOnly, "Verification")
                    Return False
                End If

                poResult.Misc = iRandom(0, 9) & PadLeft(Hex(poRawxxx.Misc), 2, "0")
            Case CodeApproval.pxeDay2Day
                'Misc should be the time the request was issued...
                poResult.Misc = Chr(iRandom(65, 90)) & PadLeft(Hex(Val(poRawxxx.Misc) + 70), 2, "0")
            Case CodeApproval.pxeOfficeRebate, _
                 CodeApproval.pxeFieldRebate, _
                 CodeApproval.pxeMCDiscount, _
                 CodeApproval.pxePartsDiscount, _
                 CodeApproval.pxeSPPurcDelivery, _
                 CodeApproval.pxeIssueORNotPR, _
                 CodeApproval.pxeIssueORNotSI, _
                 CodeApproval.pxeMCIssuance, _
                 CodeApproval.pxeMPDiscount, _
                 CodeApproval.pxeJobOrderWOGCard

                If poRawxxx.XSystem <> CodeApproval.pxeJobOrderWOGCard Then
                    'Misc should be the reference number of the transaction approved...
                    If Not IsNumeric(poRawxxx.Misc) Then
                        MsgBox("Invalid Reference Number detected!", vbCritical + vbOKOnly, "Verification")
                        Return False
                    End If
                End If

                poResult.Misc = PadLeft(Hex(TotalStr(poRawxxx.Misc)), 3, "0")
            Case CodeApproval.pxeForgot2Log, _
                 CodeApproval.pxeBusinessTrip, _
                 CodeApproval.pxeBusinessTripWLog, _
                 CodeApproval.pxeLeave, _
                 CodeApproval.pxeOvertime, _
                 CodeApproval.pxeShift, _
                 CodeApproval.pxeDayOff, _
                 CodeApproval.pxeTardiness, _
                 CodeApproval.pxeUnderTime, _
                 CodeApproval.pxeCreditInvestigation, _
                 CodeApproval.pxeCreditApplication, _
                 CodeApproval.pxeCashBalance, _
                 CodeApproval.pxeMCClusteringDelivery, _
                 CodeApproval.pxeFSEPActivation, _
                 CodeApproval.pxeFSEXActivation

                'FIRST 30 characters of fullname with the following format:
                '   LASTNAME, FIRSTNAME(SFX) MIDDNAME
                poResult.Misc = PadLeft(Hex(TotalStr(LCase(Left(poRawxxx.Misc, 30)))), 3, "0")
            Case CodeApproval.pxeAdditional, CodeApproval.pxeBiyahingFiesta, CodeApproval.pxeTeleMktg, CodeApproval.pxePreApproved
                poResult.Misc = PadLeft(Hex(TotalStr(LCase(Left(poRawxxx.Misc, 30)))), 3, "0")
            Case Else
                MsgBox("Invalid System Approval Request detected!", vbCritical + vbOKOnly, "Verification")
                Return False
        End Select

        If poRawxxx.Branch <> "" Then
            poResult.Branch = PadLeft(Hex(TotalStr(Mid(poRawxxx.Branch, 2))), 2, "0")
            poResult.XDate = PadLeft(Hex(Month(CDate(poRawxxx.XDate)) + Day(CDate(poRawxxx.XDate)) + Int(Format(CDate(poRawxxx.XDate), "yy"))), 2, "0")
        Else
            poResult.XDate = Hex(PadLeft(Day(CDate(poRawxxx.XDate)), 2, "0") & PadLeft(Month(CDate(poRawxxx.XDate)), 2, "0") & Right(Year(CDate(poRawxxx.XDate)), 2))
        End If

        poResult.XSystem = PadLeft(Hex(TotalStr(poRawxxx.XSystem)), 2, "0")
        poResult.IssuedBy = poRawxxx.IssuedBy

        pbEncode = True
        Return True
    End Function

    Public Function Equal(fsCode1 As String, fsCode2 As String) As Integer
        'Equal = -100

        'Length is not equal to 10
        If Len(fsCode1) <> 10 Then Return -100
        If Len(fsCode2) <> 10 Then Return -100

        'Convert to uppercase the code to be checked
        fsCode1 = UCase(fsCode1)
        fsCode2 = UCase(fsCode2)

        'Requesting branch is different from the given code
        If Not Mid(fsCode1, 5, 2) = Mid(fsCode2, 5, 2) Then Return -100

        'System approval request is different from the given code
        If Not Mid(fsCode1, 7, 2) = Mid(fsCode2, 7, 2) Then Return -100

        'Date requested is different from the given code
        If Not Mid(fsCode1, 9, 2) = Mid(fsCode2, 9, 2) Then Return -100

        'Issuing Department/Person is different from the given code
        If Not Mid(fsCode1, 4, 1) = Mid(fsCode2, 4, 1) Then Return -100

        Select Case Mid(fsCode1, 7, 2) 'System Approval Request
            Case PadLeft(Hex(TotalStr(CodeApproval.pxeDay2Day)), 2, "0")
                'Misc Info is different from the given code
                'New Issued - Old Issued => If <=0 Then INVALID SINCE we need a new code
                Equal = Val(Mid(fsCode2, 2, 2)) - Val(Mid(fsCode1, 2, 2))
                Exit Function
            Case PadLeft(Hex(TotalStr(CodeApproval.pxeManualLog)), 2, "0")
                'Misc Info is different from the given code
                If Not Mid(fsCode1, 2, 2) = Mid(fsCode2, 2, 2) Then Return -100
            Case Else
                'Misc Info is different from the given code
                If Not Mid(fsCode1, 1, 3) = Mid(fsCode2, 1, 3) Then Return -100
        End Select

        Return 0
    End Function

    'fsCode2 is the Approval Code given by the authorized personel/department
    Function Equalx(fsCode1 As String, fsCode2 As String) As Integer
        Dim ldDate1 As Date
        Dim ldDate2 As Date
        Dim lsDatex As String

        'Equalx = -100

        'Length is not equal to 10
        If Len(fsCode1) <> 10 Then Return -100
        If Len(fsCode2) <> 10 Then Return -100

        'Convert to uppercase the code to be checked
        fsCode1 = UCase(fsCode1)
        fsCode2 = UCase(fsCode2)

        'Issuing Department/Person is different from the given code
        If Not Mid(fsCode1, 4, 1) = Mid(fsCode2, 4, 1) Then Return -100

        'System approval request is different from the given code
        If Not Mid(fsCode1, 5, 2) = Mid(fsCode2, 5, 2) Then Return -100

        'Misc Information/Name
        If Not Mid(fsCode1, 1, 3) = Mid(fsCode2, 1, 3) Then Return -100

        'Check date
        'Date has this format: DDMMYY
        lsDatex = PadLeft(CLng("&H" & Mid(fsCode2, 7)), 6, "0")

        ldDate1 = Mid(lsDatex, 3, 2) & "/" & Mid(lsDatex, 1, 2) & "/" & Mid(lsDatex, 5, 2)

        lsDatex = PadLeft(CLng("&H" & Mid(fsCode1, 7)), 5, "0")
        ldDate2 = Mid(lsDatex, 3, 2) & "/" & Mid(lsDatex, 1, 2) & "/" & Mid(lsDatex, 5, 2)

        Select Case Mid(fsCode1, 5, 2)
            Case PadLeft(Hex(TotalStr(CodeApproval.pxeTeleMktg)), 2, "0"), PadLeft(Hex(TotalStr(CodeApproval.pxePreApproved)), 2, "0")
                If ldDate1 >= ldDate2 And ldDate1 <= DateAdd("D", 60, ldDate2) Then
                    Return 0
                End If
            Case PadLeft(Hex(TotalStr(CodeApproval.pxeBiyahingFiesta)), 2, "0")
                'Biyaheng Fiesta should be dated from the start of Biyaheng Fiesta
                If ldDate1 >= ldDate2 And ldDate1 <= DateAdd("D", 3, ldDate2) Then
                    Return 0
                End If
            Case PadLeft(Hex(TotalStr(CodeApproval.pxeAdditional)), 2, "0")
                If ldDate1 = ldDate2 Then
                    Return 0
                End If
        End Select
        Return -100
    End Function

    Private Function TotalStr(ByVal value As String) As Int32
        Dim lnTotal As Int32
        value = Replace(value, " ", "")
        value = Replace(value, ",", "")

        Dim asciis As Byte() = System.Text.Encoding.ASCII.GetBytes(value)

        Dim i As Int32
        For i = 0 To asciis.Length - 1
            lnTotal = lnTotal + CByte(asciis(i))
        Next

        Return lnTotal
    End Function

    Private Function Bin2Dec(Num As String) As Long
        If (IsNumeric(Num)) Then
            Return Convert.ToInt32(Num, 2)
        Else
            Return 0
        End If
    End Function

    Private Function Dec2Bin(ByVal Num As Long) As String
        If (IsNumeric(Num)) Then
            Return Convert.ToInt32(Num, 2)
        Else
            Return "0"
        End If
    End Function

    Private Function iRandom(ByVal from As Long, ByVal thru As Long) As Integer
        Return randm.Next(from, thru)
    End Function

    Private Function PadLeft(ByVal fsStr As String, ByVal fnLen As Integer, ByVal fsPad As String) As String
        Return fsStr.PadLeft(fnLen, fsPad)
    End Function

    Private Class xObject
        Public XSystem As String  'System Approval Requested
        Public Branch As String   'Requesting Branch
        Public IssuedBy As String 'Issuing Department/Person
        Public XDate As String
        Public Misc As String
    End Class

    Public Sub New()
        poRawxxx = New xObject
        poResult = New xObject
    End Sub
End Class
