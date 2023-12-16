'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Application Driver Object
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
'  XerSys [ 12/27/2012 11:13 am ]
'      Start translating of this object.
'      This object is based from the Application driver originally written in vb6.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports System.Environment
Imports System.IO
Imports MySql.Data.MySqlClient
Imports ADODB

Public Class GRider
    Public Const xeLogical_NO As String = "0"
    Public Const xeLogical_YES As String = "1"

    Private Const xsSignature As String = "08220326"
    Private Const xsMODULENAME = "clsAppDriver"

    Private p_oConn As New MySqlConnection
    Private p_oTrans As MySqlTransaction

    Private p_oSC As New MySqlCommand

    ' system environment variables
    Private p_sClientID As String
    Private p_sClientNm As String
    Private p_sAddressx As String
    Private p_sTownName As String
    Private p_sZippCode As String
    Private p_sProvName As String
    Private p_sTelNoxxx As String
    Private p_sFaxNoxxx As String
    Private p_sApproved As String
    Private p_sSysAdmin As String
    Private p_sProdctID As String
    Private p_sProdctNm As String
    Private p_sNetWarex As String
    Private p_sMachinex As String
    Private p_sApplPath As String
    Private p_sReptPath As String
    Private p_sImgePath As String
    Private p_dSysDatex As Date
    Private p_nNetError As Long
    Private p_sBranchCd As String
    Private p_dCapturex As Date
    Private p_dLicencex As Date
    Private p_sComptrID As String
    Private p_sMenuName As String

    Private p_sDatabase As String
    Private p_sServerNm As String
    Private p_sServerPs As String
    Private p_sServerUs As String

    ' system-user variables
    Private p_sUserIDxx As String
    Private p_sLogNamex As String
    Private p_sUserName As String
    Private p_nUserLevl As Integer
    Private p_sLogNoxxx As String
    Private p_sEmployNo As String

    ' branch variables
    Private p_sBranchNm As String
    Private p_cMainOffc As Byte
    Private p_cWareHous As Byte

    Private p_sCompName As String
    Private p_sProcName As String

    Private p_oMDIMainx As Object

    Private pbErrorLog As Boolean
    Private pnTransact As Integer
    Private psModlName As String
    Private psSQL As String

    Private p_sEmployID As String
    Private p_sEmpLevID As String

    Private pbChkErrCt As Boolean
    Private p_nHexCrypt As Integer

    Public Function LoadEnv() As Boolean
        p_oConn = doConnect()
        If IsNothing(p_oConn) Then
            Return False
        End If

        p_sCompName = Environment.MachineName
        Return loadConfig()
    End Function

    Public Overridable Function LogUser() As Boolean
        Dim loLogIn As New frmLogin
        Dim loDT As New DataTable

        Dim lnCtr As Integer = 0
        Dim lbLogIn As Boolean = False

        With loLogIn
            ' Assign the screen

            Do
                loLogIn.txtServer.Text = p_sServerNm
                loLogIn.lblBranchName.Text = p_sBranchNm
                loLogIn.lblAddress.Text = p_sAddressx
                loLogIn.lblProvince.Text = p_sProvName
                loLogIn.ShowDialog()

                ' Check log name and password
                If loLogIn.Cancelled = True Then
                    Return False
                End If

                p_oSC.CommandText = getSQ_User()
                p_oSC.Parameters.Clear()
                p_oSC.Parameters.AddWithValue("?sLogNamex", Encrypt(loLogIn.txtUsername.Text, xsSignature))
                p_oSC.Parameters.AddWithValue("?sPassword", Encrypt(loLogIn.txtPassword.Text, xsSignature))
                Debug.Print(p_oSC.CommandText)
                loDT = ExecuteQuery(p_oSC)

                If loDT.Rows.Count = 0 Then
                    MsgBox("User Does Not Exist!" & vbCrLf & "Verify log name and/or password.", vbCritical, "Warning")
                    lnCtr += 1
                Else

                    If Not isUserActive(loDT) Then
                        lnCtr = 0
                    Else
                        lbLogIn = True
                    End If
                End If
            Loop Until lbLogIn Or lnCtr = 3
        End With

        If lbLogIn Then
            p_sUserIDxx = loDT.Rows(0).Item("sUserIDxx")
            p_sUserName = loDT.Rows(0).Item("sUserName")
            p_sLogNamex = loDT.Rows(0).Item("sLogNamex")
            p_nUserLevl = loDT.Rows(0).Item("nUserLevl")
            p_sEmployNo = loDT.Rows(0).Item("sEmployNo")
        End If
        Return lbLogIn
    End Function

    Public Overridable Function LogUser(ByVal fsUserID As String) As Boolean
        Dim loDT As New DataTable

        p_oSC.CommandText = getSQ_UserByID()
        p_oSC.Parameters.Clear()
        p_oSC.Parameters.AddWithValue("?sUserIDxx", fsUserID)
        loDT = ExecuteQuery(p_oSC)


        If loDT.Rows.Count = 0 Then
            MsgBox("User Does Not Exist!" & vbCrLf & "Verify log name and/or password.", vbCritical, "Warning")
            Return False
        Else
            If Not isUserActive(loDT) Then
                Return False
            End If
        End If

        p_sUserIDxx = loDT.Rows(0).Item("sUserIDxx")
        p_sUserName = loDT.Rows(0).Item("sUserName")
        p_sLogNamex = loDT.Rows(0).Item("sLogNamex")
        p_nUserLevl = loDT.Rows(0).Item("nUserLevl")
        p_sEmployNo = loDT.Rows(0).Item("sEmployNo")
        Return True
    End Function

    Public Function BeginTransaction() As Boolean
        If Not p_oTrans Is Nothing Then
            MsgBox("Transaction Exists!" & vbCrLf & _
                   "Can not create new transaction!", vbCritical, "Warning")
            Return False
        End If

        If p_oConn.Ping = False Then
            p_oConn = doConnect()
            If IsNothing(p_oConn) Then
                MsgBox("Invalid connection detected!" & vbCrLf & _
                       "Can not create new conection!", vbCritical, "Warning")
                Return False
            End If
        End If

        Try
            p_oTrans = p_oConn.BeginTransaction()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw ex
        End Try

        Return False
    End Function

    Public Function CommitTransaction() As Boolean
        If p_oTrans Is Nothing Then
            Return False
        End If

        Try
            p_oTrans.Commit()
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw ex
        End Try

        p_oTrans = Nothing
        Return True
    End Function

    Public Function RollBackTransaction() As Boolean
        If p_oTrans Is Nothing Then
            Return False
        End If

        Try
            p_oTrans.Rollback()
        Catch ex As MySqlException
            MsgBox(ex.Message)
            Throw ex
        End Try

        p_oTrans = Nothing
        Return True
    End Function

    Public Function Execute(ByVal sQuery As String, _
                        ByVal sTableNme As String) As Integer
        Return Execute(sQuery, sTableNme, "", "")
    End Function

    Public Function Execute(ByVal sQuery As String, _
                            ByVal sTableNme As String, _
                            ByVal sBranchCd As String) As Integer
        Return Execute(sQuery, sTableNme, sBranchCd, "")
    End Function

    'Since error is handled, an error in Statement also returns 0 
    'thus we can't possibly catch 
    Public Function Execute(ByVal sQuery As String, _
                            ByVal sTableNme As String, _
                            ByVal sBranchCd As String, _
                            ByVal sDestinat As String) As Integer

        Dim lnAffected As Integer
        Dim lbCreate As Boolean
        Dim lsSQL As String

        lbCreate = p_oTrans Is Nothing
        If lbCreate Then
            If p_oConn.Ping = False Then
                p_oConn = doConnect()
                If IsNothing(p_oConn) Then
                    MsgBox("Invalid connection detected!" & vbCrLf & _
                           "Can not create new conection!", vbCritical, "Warning")
                    Return 0
                End If
            End If

            p_oTrans = p_oConn.BeginTransaction
        End If

        If sQuery = "" Then
            Return 0
        End If

        If sBranchCd = "" Then
            sBranchCd = p_sBranchCd
        End If
        lsSQL = "INSERT INTO xxxReplicationLog" & _
                        " SET sTransNox = " & strParm(GetNextCode("xxxReplicationLog", "sTransNox", True, _
                              p_oConn, True, p_sBranchCd)) & _
                        ", sBranchCd = " & strParm(sBranchCd) & _
                        ", sStatemnt = " & strParm(sQuery) & _
                        ", sTableNme = " & strParm(sTableNme) & _
                        ", sDestinat = " & strParm(sDestinat) & _
                        ", sModified = " & strParm(p_sUserIDxx) & _
                        ", dEntryDte = " & datetimeParm(SysDate()) & _
                        ", dModified = " & datetimeParm(SysDate())
        Debug.Print(lsSQL)
        Try
            If ExecuteActionQuery(lsSQL) <= 0 Then
                If lbCreate Then
                    p_oTrans.Rollback()
                    p_oTrans = Nothing
                End If
                Return 0
            End If
        Catch ex As Exception
            If lbCreate Then
                p_oTrans.Rollback()
                p_oTrans = Nothing
            End If
            MsgBox(ex.Message)
            Return 0
        End Try

        Try
            lnAffected = ExecuteActionQuery(sQuery)
        Catch ex As Exception
            If lbCreate Then
                p_oTrans.Rollback()
                p_oTrans = Nothing
            End If
            MsgBox(ex.Message)
            Return 0
        End Try

        If lbCreate Then
            p_oTrans.Commit()
            p_oTrans = Nothing
        End If
        Return lnAffected
    End Function

    Public Function ExecuteQuery(ByVal fsSQLCmd As String) As DataTable
        Dim loDA As New MySqlDataAdapter
        Dim loDataTable As New DataTable

        Try
            loDA.SelectCommand = New MySqlCommand(fsSQLCmd, p_oConn)
        Catch ex As MySqlException
            MsgBox(ex.Message)
            Throw ex
        End Try

        Debug.Print(fsSQLCmd)
        Console.WriteLine(fsSQLCmd)

        Try
            loDA.Fill(loDataTable)
        Catch ex As Exception
            loDA.SelectCommand = New MySqlCommand(fsSQLCmd, p_oConn)

            'loDA.Fill(loDataTable)
        End Try

        Return loDataTable
    End Function

    Public Function ExecuteQuery(ByRef fsMySQLCmd As MySqlCommand) As DataTable
        Dim loDataTable As New DataTable
        Dim loDA As New MySqlDataAdapter

        Try
            loDA.SelectCommand = fsMySQLCmd
        Catch ex As MySqlException
            MsgBox(ex.Message)
            Throw ex
        End Try

        Try
            loDA.Fill(loDataTable)
        Catch ex As Exception
            loDA.SelectCommand = fsMySQLCmd
            loDA.Fill(loDataTable)
        End Try

        Return loDataTable
    End Function

    Public Function SaveEvent(ByVal sEventIDx As String, _
                                ByVal sRemarksx As String, _
                                ByVal sSerialNo As String) As Boolean
        'event log
        Dim lsSQLEvent = "INSERT INTO Event_Master SET" & _
                                "  sTransNox = " & strParm(GetNextCode("Event_Master", "sTransNox", True, Connection, True, BranchCode)) & _
                                ", sEventIDx = " & strParm(sEventIDx) & _
                                ", sRemarksx = " & strParm(sRemarksx) & _
                                ", sUserIDxx = " & strParm(UserID) & _
                                ", sSerialNo = " & strParm(sSerialNo) & _
                                ", sComptrNm = " & strParm(Environment.MachineName) & _
                                ", dModified = " & datetimeParm(getSysDate)
        Execute(lsSQLEvent, "lsSQLEvent")

        Return True
    End Function

    Public Function getConfiguration(ByVal ConfigCd As String, Optional ByVal BranchCd As String = "") As String
        Dim loDT As DataTable
        Dim lsSQL As String

        If BranchCd = "" Then
            lsSQL = "SELECT sValuexxx" & _
                     " FROM xxxOtherConfig" & _
                     " WHERE sProdctID = " & strParm(p_sProdctID) & _
                        " AND sConfigId = " & strParm(ConfigCd)
        Else
            lsSQL = "SELECT " & ConfigCd & " sValuexxx" & _
                     " FROM xxxOtherInfo a" & _
                        " LEFT JOIN xxxSysClient b" & _
                           " ON a.sClientID = b.sClientID" & _
                     " WHERE sBranchCd = " & strParm(BranchCd)
        End If

        Try
            loDT = ExecuteQuery(lsSQL)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return ""
        End Try

        With loDT
            If .Rows.Count <= 0 Then
                Return ""
            End If

            Return .Rows(0).Item("sValuexxx")
        End With
    End Function

    Function doConnect() As MySqlConnection
        Dim loINI As New INIFile
        Dim loCrypt As New Crypto
        Dim loConn As New MySqlConnection

        Dim lsUserName As String, lsPassword As String
        Dim lsConnStr As String, lsPortNmbr As String

        loINI.FileName = Environ("windir") & "\GRider.ini"

        If Not loINI.IsFileExist() Then
            MsgBox("Invalid Config File Detected!" & vbCrLf & "Verify your argument then try Again!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical)
            Return Nothing
        End If

        p_nHexCrypt = loINI.GetTextValue(p_sProdctID, "CryptType")
        lsUserName = Decrypt(loINI.GetTextValue(p_sProdctID, "UserName"), xsSignature)
        lsPassword = Decrypt(loINI.GetTextValue(p_sProdctID, "Password"), xsSignature)
        lsPortNmbr = loINI.GetTextValue(p_sProdctID, "Port")
        p_sClientID = loINI.GetTextValue(p_sProdctID, "ClientID")

        'get the database info here...
        p_sDatabase = loINI.GetTextValue(p_sProdctID, "Database")
        p_sServerNm = loINI.GetTextValue(p_sProdctID, "ServerName")
        p_sServerPs = lsPassword
        p_sServerUs = lsUserName

        lsConnStr = "Data Source=" & p_sServerNm & ";" & _
                    "Database=" & p_sDatabase & ";" & _
                    "User=" & lsUserName & ";" & _
                    "Password=" & lsPassword & ";" & _
                    "Convert Zero Datetime=True;"
        ' "MultipleActiveResultSets=True"
        ' Include this to allow data transfer from mysql to adodb object

        Debug.Print(lsConnStr)
        loConn.ConnectionString = lsConnStr

        Try
            loConn.Open()
        Catch err As MySqlException
            MsgBox("Err #:" & err.Number & vbCrLf & _
                   "Description:" & err.Message)
            Return Nothing
        End Try
        Return loConn
    End Function

    Private Function loadConfig() As Boolean
        Dim loDT As DataTable

        p_oSC.Connection = p_oConn
        p_oSC.CommandText = getSQ_LoadConfig()
        Debug.Print(p_sProdctID & " " & p_sClientID & " " & p_sCompName)
        p_oSC.Parameters.AddWithValue("?sProdctID", p_sProdctID)
        p_oSC.Parameters.AddWithValue("?sClientID", p_sClientID)
        p_oSC.Parameters.AddWithValue("?sComptrNm", p_sCompName)
        loDT = ExecuteQuery(p_oSC)

        If loDT.Rows.Count = 0 Then
            MsgBox("Application is Not Registered!", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        If String.IsNullOrEmpty(loDT.Rows(0).Item("sComptrNm").ToString) Then
            MsgBox("Computer is Not Registered to Use the Selected Sytem!", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        ' check error count
        p_nNetError = loDT.Rows(0).Item("nNetError")
        If pbChkErrCt Then
            If p_nNetError > 100 And p_nNetError < 200 Then
                MsgBox("Application has Reached 100 ApplicatiOn Error!" & vbCrLf & _
                       "Please Inform the GCC-SEG for this Application to Avoid Further Damages!", MsgBoxStyle.Critical, "Warning")
            ElseIf p_nNetError > 200 Then
                MsgBox("Maximum Error Limit has been Reached!" & vbCrLf & _
                          "The Application will Locked to Avert Further Damages!", MsgBoxStyle.Critical, "Warning")
                Return False
            End If
        End If

        p_sClientNm = loDT.Rows(0).Item("sClientNm")
        p_sAddressx = loDT.Rows(0).Item("sAddressx")
        p_sTownName = loDT.Rows(0).Item("sTownName")
        p_sZippCode = loDT.Rows(0).Item("sZippCode")
        p_sProvName = loDT.Rows(0).Item("sProvName")
        p_sTelNoxxx = loDT.Rows(0).Item("sTelNoxxx")
        p_sFaxNoxxx = loDT.Rows(0).Item("sFaxNoxxx")
        p_sApproved = loDT.Rows(0).Item("sApproved")
        p_sBranchCd = loDT.Rows(0).Item("sBranchCd")
        p_sProdctNm = loDT.Rows(0).Item("sProdctNm")
        p_sApplPath = loDT.Rows(0).Item("sApplPath")
        p_sReptPath = loDT.Rows(0).Item("sReptPath")
        p_sImgePath = loDT.Rows(0).Item("sImgePath")
        p_sSysAdmin = loDT.Rows(0).Item("sSysAdmin")
        p_sNetWarex = loDT.Rows(0).Item("sNetWarex")
        p_sMachinex = loDT.Rows(0).Item("sMachinex")
        p_dSysDatex = loDT.Rows(0).Item("dSysDatex").ToString
        p_dLicencex = loDT.Rows(0).Item("dLicencex").ToString
        p_nNetError = loDT.Rows(0).Item("nNetError")
        p_sBranchNm = loDT.Rows(0).Item("sBranchNm")
        p_cWareHous = loDT.Rows(0).Item("cWareHous")
        p_cMainOffc = loDT.Rows(0).Item("cMainOffc")

        If Not isSignatureOK(p_sMachinex, p_sNetWarex, p_sSysAdmin) Then
            Return False
        End If

        Return True
    End Function

    Private Function isSignatureOK(ByVal fsMachineX As String, _
                                   ByVal fsNetWareX As String, _
                                   ByVal fsSysAdmin As String) As Boolean
        Dim loDT As DataTable

        fsMachineX = Decrypt(fsMachineX, xsSignature)
        fsNetWareX = Decrypt(fsNetWareX, xsSignature)
        fsSysAdmin = Decrypt(fsSysAdmin, xsSignature)

        p_oSC.CommandText = getSQ_Signature()
        p_oSC.Parameters.Clear()
        p_oSC.Parameters.AddWithValue("?sUserID1", fsMachineX)
        p_oSC.Parameters.AddWithValue("?sUserID2", fsNetWareX)
        loDT = ExecuteQuery(p_oSC)

        If loDT.Rows.Count <> 2 Then
            MsgBox("Unregistered Copy of " & p_sProdctNm & " Detected!", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        Return True
    End Function

    Private Function isUserActive(ByRef loDT As DataTable) As Boolean
        Dim lnCtr As Integer = 0
        Dim lbMember As Boolean = False

        ' check membership of user
        If loDT.Rows(0).Item("cUserType").Equals(0) Then
            For lnCtr = 0 To loDT.Rows.Count - 1
                If loDT.Rows(0).Item("sProdctID").Equals(p_sProdctID) Then
                    Exit For
                    lbMember = True
                End If
            Next
        Else
            lbMember = True
        End If

        If Not lbMember Then
            MsgBox("User is not a member of this application!!!" & vbCrLf & _
               "Application used is not allowed!!!", vbCritical, "Warning")
        End If

        ' check user status
        If loDT.Rows(0).Item("cUserStat").Equals(xeUserStatus.Suspended) Then
            MsgBox("User is currently suspended!!!" & vbCrLf & _
                     "Application used is not allowed!!!", vbCritical, "Warning")
            Return False
        End If
        Return True
    End Function

    Public ReadOnly Property Connection As MySqlConnection
        Get
            Return p_oConn
        End Get
    End Property

    Public ReadOnly Property Signature As String
        Get
            Return xsSignature
        End Get
    End Property

    Public ReadOnly Property ClientID As String
        Get
            Return p_sClientID
        End Get
    End Property

    Public ReadOnly Property ClientName As String
        Get
            Return p_sClientNm
        End Get
    End Property

    Public ReadOnly Property Address As String
        Get
            Return p_sAddressx
        End Get
    End Property

    Public ReadOnly Property TownCity As String
        Get
            Return p_sTownName
        End Get
    End Property

    Public ReadOnly Property ZippCode As String
        Get
            Return p_sZippCode
        End Get
    End Property

    Public ReadOnly Property Province As String
        Get
            Return p_sProvName
        End Get
    End Property

    Public ReadOnly Property TelNumber As String
        Get
            Return p_sTelNoxxx
        End Get
    End Property

    Public ReadOnly Property FaxNumber As String
        Get
            Return p_sFaxNoxxx
        End Get
    End Property

    Public ReadOnly Property Approved As String
        Get
            Return p_sApproved
        End Get
    End Property

    Public ReadOnly Property SysAdmin As String
        Get
            Return p_sSysAdmin
        End Get
    End Property

    Public ReadOnly Property ProductID As String
        Get
            Return p_sProdctID
        End Get
    End Property

    Public ReadOnly Property ProductName As String
        Get
            Return p_sProdctNm
        End Get
    End Property

    Public ReadOnly Property AppPath As String
        Get
            Return p_sApplPath
        End Get
    End Property

    Public ReadOnly Property SysDate As Date
        Get
            Return getSysDate()
        End Get
    End Property

    Public ReadOnly Property NetError As Long
        Get
            Return p_nNetError
        End Get
    End Property

    Public ReadOnly Property BranchCode As String
        Get
            Return p_sBranchCd
        End Get
    End Property

    Public ReadOnly Property BranchName As String
        Get
            Return p_sBranchNm
        End Get
    End Property

    Public ReadOnly Property UserID As String
        Get
            UserID = p_sUserIDxx
        End Get
    End Property

    Public ReadOnly Property LogName As String
        Get
            LogName = p_sLogNamex
        End Get
    End Property

    Public ReadOnly Property UserName As String
        Get
            UserName = p_sUserName
        End Get
    End Property

    Public ReadOnly Property UserLevel As Integer
        Get
            UserLevel = p_nUserLevl
        End Get
    End Property

    Public ReadOnly Property ServerName As String
        Get
            If p_nUserLevl >= xeUserRights.MANAGER Then
                ServerName = p_sServerNm
            Else
                ServerName = ""
            End If
        End Get
    End Property

    Public ReadOnly Property Database As String
        Get
            If p_nUserLevl >= xeUserRights.MANAGER Then
                Database = p_sDatabase
            Else
                Database = ""
            End If
        End Get
    End Property

    Public ReadOnly Property EmployID As String
        Get
            Return p_sEmployID
        End Get
    End Property

    Public ReadOnly Property EmployeeLevel As String
        Get
            Return p_sEmpLevID
        End Get
    End Property

    Public ReadOnly Property ServerUser As String
        Get
            If p_nUserLevl >= xeUserRights.MANAGER Then
                ServerUser = p_sServerUs
            Else
                ServerUser = ""
            End If
        End Get
    End Property

    Public ReadOnly Property ServerPassword As String
        Get
            If p_nUserLevl >= xeUserRights.MANAGER Then
                ServerPassword = p_sServerPs
            Else
                ServerPassword = ""
            End If
        End Get
    End Property

    Public ReadOnly Property EmployNo As String
        Get
            EmployNo = p_sEmployNo
        End Get
    End Property


    Public Property MDI As Object
        Get
            MDI = p_oMDIMainx
        End Get
        Set(ByVal value As Object)
            p_oMDIMainx = value
        End Set
    End Property

    'Functions used to make sure that connection objection is working properly...
    '+++++++++++++++++++++++++++++++++++++++
    Public Function isConnected() As Boolean
        If Not IsNothing(p_oConn) Then
            Return p_oConn.Ping
        End If
        Return False
    End Function

    Public Function Reconnect() As Boolean
        p_oConn = doConnect()
        If IsNothing(p_oConn) Then
            MsgBox("Invalid connection detected!" & vbCrLf & _
                   "Can not create new conection!", vbCritical, "Warning")
            Return False
        End If
        Return True
    End Function
    '+++++++++++++++++++++++++++++++++++++++

    Public Function ExecuteActionQuery(ByVal fsSQLCmd As String) As Integer
        Dim loSC As New MySqlCommand(fsSQLCmd, p_oConn)
        Dim lnAffected As Integer

        Try
            lnAffected = loSC.ExecuteNonQuery()
        Catch ex As MySqlException
            MsgBox(ex.Message)
            Throw ex

            Return 0
        End Try

        Return lnAffected
    End Function

    Public Function getSysDate() As Date
        Dim loDT As DataTable

        If Not isConnected() Then
            If Not Reconnect() Then
                Return "1900-01-01"
            End If
        End If

        loDT = ExecuteQuery("SELECT SYSDATE()")

        Return loDT(0).Item(0).ToString
    End Function

    Public Function Encrypt(ByVal Code As String, Optional ByVal Signature As String = "") As String
        Dim loCrypt As New Crypto
        Dim xValue As String
        If Signature <> "" Then
            loCrypt.Signature = Signature
        End If

        loCrypt.InBuffer = Code
        loCrypt.Encrypt()
        xValue = loCrypt.OutBuffer

        If (p_nHexCrypt = 1) Then
            xValue = StringToHex(xValue)
        End If

        Encrypt = xValue
    End Function

    Public Function Decrypt(ByVal Code As String, Optional ByVal Signature As String = "") As String
        Dim loCrypt As New Crypto

        If Signature <> "" Then
            loCrypt.Signature = Signature
        End If

        If (p_nHexCrypt = 1) Then
            Code = HexToString(Code)
        End If

        loCrypt.InBuffer = Code
        loCrypt.Decrypt()
        Return loCrypt.OutBuffer
    End Function

    Private Function getSQ_LoadConfig() As String
        Return "SELECT a.sClientID" & _
           ", a.sClientNm" & _
           ", a.sAddressx" & _
           ", a.sTownName" & _
           ", a.sZippCode" & _
           ", a.sProvName" & _
           ", a.sTelNoxxx" & _
           ", a.sFaxNoxxx" & _
           ", a.sApproved" & _
           ", a.sBranchCd" & _
           ", b.sProdctNm" & _
           ", c.sApplPath" & _
           ", c.sReptPath" & _
           ", c.sImgePath" & _
           ", c.sSysAdmin" & _
           ", c.sNetWarex" & _
           ", c.sMachinex" & _
           ", c.dSysDatex" & _
           ", c.dLicencex" & _
           ", c.nNetError" & _
           ", d.sComptrNm" & _
           ", e.sBranchNm" & _
           ", e.cWareHous" & _
           ", e.cMainOffc" & _
         " FROM xxxSysClient a" & _
             ", xxxSysObject b" & _
             ", xxxSysApplication c" & _
                  " LEFT JOIN xxxSysWorkStation d" & _
                     " ON c.sClientID = d.sClientID" & _
                        " AND d.sComptrNm = ?sComptrNm" & _
            ", Branch e" & _
         " WHERE c.sClientID = a.sClientID" & _
            " AND c.sProdctID = b.sProdctID" & _
            " AND a.sBranchCd = e.sBranchCd" & _
            " AND a.sClientID = ?sClientID" & _
            " AND b.sProdctID = ?sProdctID"
    End Function

    Private Function getSQ_Signature() As String
        Return "SELECT" & _
           "  sUserIDxx" & _
           ", sUserName" & _
           ", sLogNamex" & _
           " FROM xxxSysUser" & _
           " WHERE sUserIDxx IN (?sUserID1, ?sUserID2)"
    End Function

    Private Function getSQ_UserByID() As String
        Return "SELECT sUserIDxx" & _
              ", sLogNamex" & _
              ", sPassword" & _
              ", sUserName" & _
              ", nUserLevl" & _
              ", cUserType" & _
              ", sProdctID" & _
              ", cUserStat" & _
              ", nSysError" & _
              ", cLogStatx" & _
              ", cLockStat" & _
              ", cAllwLock" & _
              ", sEmployNo" & _
           " FROM xxxSysUser" & _
           " WHERE sUserIDxx = ?sUserIDxx"
    End Function

    Private Function getSQ_User() As String
        Return "SELECT sUserIDxx" & _
              ", sLogNamex" & _
              ", sPassword" & _
              ", sUserName" & _
              ", nUserLevl" & _
              ", cUserType" & _
              ", sProdctID" & _
              ", cUserStat" & _
              ", nSysError" & _
              ", cLogStatx" & _
              ", cLockStat" & _
              ", cAllwLock" & _
              ", sEmployNo" & _
           " FROM xxxSysUser" & _
           " WHERE sLogNamex = ?sLogNamex" & _
              " AND sPassword = ?sPassword"
    End Function

    Public Sub New(ByVal fsProdctID As String)
        p_sProdctID = fsProdctID
    End Sub

#Region "Error Log"
    Public Sub ErrorLog(ByVal Message As String)
        Dim sw As StreamWriter = Nothing

        Try
            Dim sLogFormat As String = DateTime.Now.ToShortDateString().ToString() & " " & DateTime.Now.ToLongTimeString().ToString() & " ==> "
            Dim sPathName As String = "xxxErrorLog_"

            Dim sYear As String = DateTime.Now.Year.ToString()
            Dim sMonth As String = DateTime.Now.Month.ToString()
            Dim sDay As String = DateTime.Now.Day.ToString()

            Dim sErrorTime As String = sDay & "-" & sMonth & "-" & sYear

            sw = New StreamWriter(sPathName & sErrorTime & ".txt", True)

            sw.WriteLine(sLogFormat & Message)

            sw.Flush()
        Catch ex As Exception
        Finally
            If sw IsNot Nothing Then
                sw.Dispose()
                sw.Close()
            End If
        End Try
    End Sub
#End Region

End Class
