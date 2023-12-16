'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Control Array Utility Object
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
'  Mac™ [ 04/13/2013 10:25 am ]
'      Start revising this object.
'      This object is based from kepler77's code in his article on www.codeproject.com -
'           The Closest thing to a VB6 control array in VB.NET
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports System.Windows.Forms

Public Class ControlArrayUtils
    'Converts same type of controls on a form to a control 
    'array by using the notation ControlName_1, ControlName_2, 
    'where the _ can be replaced by any separator string

    Public Shared Function getControlArray(ByVal frm As Windows.Forms.Form, _
         ByVal controlName As String, _
         Optional ByVal separator As String = "") As System.Array

        Dim i As Short
        Dim startOfIndex As Short
        Dim alist As New ArrayList
        Dim ctl As System.Windows.Forms.Control
        '        Dim ctrls() As System.Windows.Forms.Control
        Dim controlType As System.Type
        Dim strSuffix As String
        Dim maxIndex As Short = -1 'Default

        'Loop through all controls, looking for 
        'controls with the matching name pattern
        'Find the highest indexed control

        For Each ctl In frm.Controls
            startOfIndex = ctl.Name.ToLower.IndexOf(controlName.ToLower & separator)
            If startOfIndex = 0 Then
                strSuffix = ctl.Name.Substring(controlName.Length)
                'Check that the suffix is an
                ' integer (index of the array)
                If IsInteger(strSuffix) Then
                    If Val(strSuffix) > maxIndex Then _
                       maxIndex = Val(strSuffix) 'Find the highest 
                    'indexed Element
                End If
            End If
        Next ctl

        'Add to the list of controls in correct order
        If maxIndex > -1 Then

            For i = 0 To maxIndex
                Dim aControl As Control = _
                  getControlFromName(frm, controlName, i, separator)
                If Not (aControl Is Nothing) Then
                    'Save the object Type (uses the last 
                    'control found as the Type)
                    controlType = aControl.GetType
                End If
                alist.Add(aControl)
            Next
        End If
        Return alist.ToArray(controlType)
    End Function

    'Converts any type of like named controls on a form 
    'to a control array by using the notation ControlName_1, 
    'ControlName_2, where the _ can be replaced by any 
    'separator string

    Public Shared Function getMixedControlArray(ByVal frm As Windows.Forms.Form, ByVal controlName As String, _
       Optional ByVal separator As String = "") As Control()

        Dim i As Short
        Dim startOfIndex As Short
        Dim alist As New ArrayList
        '        Dim controlType As System.Type
        Dim ctl As System.Windows.Forms.Control
        '        Dim ctrls() As System.Windows.Forms.Control
        Dim strSuffix As String
        Dim maxIndex As Short = -1 'Default

        'Loop through all controls, looking for controls 
        'with the matching name pattern
        'Find the highest indexed control

        For Each ctl In frm.Controls
            startOfIndex = ctl.Name.ToLower.IndexOf(controlName.ToLower & separator)
            If startOfIndex = 0 Then
                strSuffix = ctl.Name.Substring(controlName.Length)
                'Check that the suffix is an integer 
                '(index of the array)
                If IsInteger(strSuffix) Then
                    If Val(strSuffix) > maxIndex Then _
                       maxIndex = Val(strSuffix) 'Find the highest 
                    'indexed Element
                End If
            End If
        Next ctl

        'Add to the list of controls in correct order
        If maxIndex > -1 Then
            For i = 0 To maxIndex
                Dim aControl As Control = getControlFromName(frm, _
                                         controlName, i, separator)
                alist.Add(aControl)
            Next
        End If
        Return alist.ToArray(GetType(Control))
    End Function

    Private Shared Function getControlFromName(ByRef frm As Windows.Forms.Form, _
           ByVal controlName As String, ByVal index As Short, _
           ByVal separator As String) As System.Windows.Forms.Control
        controlName = controlName & separator & index
        For Each ctl As Control In frm.Controls
            If String.Compare(ctl.Name, controlName, True) = 0 Then
                Return ctl
            End If
        Next ctl
        Return Nothing 'Could not find this control by name
    End Function

    Private Shared Function IsInteger(ByVal Value As String) As Boolean
        If Value = "" Then Return False
        For Each chr As Char In Value
            If Not Char.IsDigit(chr) Then
                Return False
            End If
        Next
        Return True
    End Function
End Class