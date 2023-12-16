Imports MySql.Data.MySqlClient
Imports ADODB
Imports System.Windows.Forms
Imports System.Reflection
Imports System.Threading
Imports System.Text.RegularExpressions

Public Module ModMain
    'mac 2020-07-23
    Public Function RMJExecute(ByVal WorkingDIR As String, ByVal ProcessPath As String, ByVal FileName As String) As Boolean
        Dim objProcess As System.Diagnostics.Process
        Try
            objProcess = New System.Diagnostics.Process()
            objProcess.StartInfo.WorkingDirectory = WorkingDIR
            objProcess.StartInfo.FileName = ProcessPath
            objProcess.StartInfo.Arguments = FileName
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            objProcess.Start()

            'Wait until the process passes back an exit code 
            objProcess.WaitForExit()

            'Free resources associated with this process
            objProcess.Close()
        Catch
            Return False
        End Try

        Return True
    End Function

    'mac 2020-07-23
    Public Function DeleteFile(ByVal fsPath As String) As Boolean
        If System.IO.File.Exists(fsPath) = True Then
            System.IO.File.Delete(fsPath)
            Return True
        End If

        Return True
    End Function

    'mac 2020-07-23
    Public Function ReadFile(ByVal fsPath As String) As String
        If System.IO.File.Exists(fsPath) = False Then Return ""

        Debug.Print(My.Computer.FileSystem.ReadAllText(fsPath))
        Return My.Computer.FileSystem.ReadAllText(fsPath)
    End Function

    Public Function Encrypt(ByVal Code As String, Optional ByVal Signature As String = "") As String
        Dim loCrypt As New Crypto

        If Signature <> "" Then
            loCrypt.Signature = Signature
        End If

        loCrypt.InBuffer = Code
        loCrypt.Encrypt()
        Return loCrypt.OutBuffer
    End Function

    Public Function Decrypt(ByVal Code As String, Optional ByVal Signature As String = "") As String
        Dim loCrypt As New Crypto

        If Signature <> "" Then
            loCrypt.Signature = Signature
        End If

        loCrypt.InBuffer = Code
        loCrypt.Decrypt()
        Return loCrypt.OutBuffer
    End Function

    Public Function AddCondition(ByVal Query As String, ByVal Condition As String) As String
        Dim lnPos As Integer
        Dim lsOrder As String

        ' check first if Condition contains a valid value
        If Condition = "" Then
            Return Query
        End If

        lsOrder = ""
        ' parse first the group by clause
        lnPos = InStr(Query, "GROUP BY", vbTextCompare)
        If lnPos > 0 Then
            lsOrder = Trim(Mid(Query, lnPos))
            Query = Trim(Left(Query, lnPos - 1))
        Else
            lnPos = InStr(Query, "HAVING", vbTextCompare)
            If lnPos > 0 Then
                lsOrder = Trim(Mid(Query, lnPos))
                Query = Trim(Left(Query, lnPos - 1))
            Else
                lnPos = InStr(Query, "ORDER BY", vbTextCompare)
                If lnPos > 0 Then
                    lsOrder = Trim(Mid(Query, lnPos))
                    Query = Trim(Left(Query, lnPos - 1))
                Else
                    lnPos = InStr(Query, "UNION", vbTextCompare)
                    If lnPos > 0 Then
                        'lsOrder = Trim(Mid(Query, lnPos))
                        'Query = Trim(Left(Query, lnPos - 1))

                        Return AddCondition(Trim(Left(Query, lnPos - 1)), Condition) & " UNION " & AddCondition(Trim(Mid(Query, lnPos + 5)), Condition)

                    End If
                End If
            End If
        End If

        lnPos = InStr(Query, "WHERE", vbTextCompare)
        If lnPos > 0 Then
            Query = Trim(Query) & " AND " & Condition
        Else
            Query = Trim(Query) & " WHERE " & Condition
        End If

        Return Trim(Query) & " " & lsOrder
    End Function

    Public Function AddCondition(ByVal Query As String, ByVal Find As String, ByVal Replacement As String) As String
        Return Replace(Query, Find, Replacement, , , CompareMethod.Text)
    End Function

    Public Function GetCodeApproval(ByVal oAppDriver As GRider, _
                                    ByRef sApprovalCde As String, _
                                    ByRef sApproveID As String, _
                                    ByRef sApproveName As String) As Boolean


        Dim loForm As frmCodeApproval
        Dim lnCtr As Integer
        Dim lbLogIn As Boolean

        GetCodeApproval = False

        loForm = New frmCodeApproval

        lnCtr = 0
        lbLogIn = False
        Do
            loForm.AppDriver = oAppDriver
            loForm.AppPath = oAppDriver.AppPath
            loForm.ShowDialog()
            If loForm.Cancel = True Then
                lnCtr = 2
            Else
                If Mid(UCase(loForm.CodeApproval), 4, 1) <> loForm.IssueeType Then
                    MsgBox("Invalid Approval Parameter Detected." & vbCrLf & vbCrLf & _
                             "Kindly verify your entry.", vbCritical, "Approval Error")
                Else
                    lbLogIn = True
                End If
            End If
            lnCtr = lnCtr + 1
        Loop Until (lbLogIn = True) Or (lnCtr = 3)

        If lbLogIn = False Then GoTo endProc

        sApprovalCde = UCase(loForm.CodeApproval)
        sApproveID = loForm.UserID
        sApproveName = loForm.Issuee

        GetCodeApproval = True

endProc:
        loForm.Close()
        loForm = Nothing

        Exit Function
    End Function

    Public Function isValidApproveCode( _
        ByVal fsSysReq As String, _
        ByVal fsBranch As String, _
        ByVal fsIssuee As String, _
        ByVal fsDatexx As String, _
        ByVal fsMiscxx As String, _
        ByVal fsCdeGvn As String _
        ) As Boolean

        Dim loCode As CodeApproval
        loCode = New CodeApproval

        isValidApproveCode = False

        loCode.XSystem = fsSysReq
        loCode.Branch = fsBranch
        loCode.IssuedBy = fsIssuee
        loCode.DateRequested = fsDatexx
        loCode.MiscInfo = fsMiscxx

        If loCode.Encode Then
            Debug.Print(loCode.Result)
            If fsBranch <> "" Then
                If loCode.Equal(loCode.Result, fsCdeGvn) >= 0 Then
                    isValidApproveCode = True
                End If
            Else
                If loCode.Equalx(loCode.Result, fsCdeGvn) >= 0 Then
                    isValidApproveCode = True
                End If
            End If
        End If

        Return isValidApproveCode

    End Function

    Public Function GetNextCode(ByVal Table As String, _
                                ByVal Field As String, _
                                ByVal YearFormat As Boolean, _
                                ByVal Connection As MySqlConnection) As String

        Return GetNextCode(Table, Field, YearFormat, Connection, False, "")
    End Function

    Public Function GetNextCode(ByVal Table As String, _
                                ByVal Field As String, _
                                ByVal YearFormat As Boolean, _
                                ByVal Connection As MySqlConnection, _
                                ByVal ByBranch As Boolean) As String
        Return GetNextCode(Table, Field, YearFormat, Connection, ByBranch, "")
    End Function

    Public Function GetNextCode(ByVal Table As String, _
                                ByVal Field As String, _
                                ByVal YearFormat As Boolean, _
                                ByVal Connection As MySqlConnection, _
                                ByVal ByBranch As Boolean, _
                                ByVal Branch As String) As String
        Dim loDA As New MySqlDataAdapter
        Dim loDT As New DataTable
        Dim loDS As New DataSet
        Dim lsField As String
        Dim lsSQL As String
        Dim lnCounter As Integer
        Dim lnCode As Long
        Dim lnLen As Long
        Dim ldDate As Date
        Dim lsStr As String = ""

        If ByBranch = True And Branch = "" Then Return ""

        lsField = IIf(ByBranch, Branch, "")
        If YearFormat Then
            loDT = ExecuteQuery("SELECT SYSDATE()", Connection)
            If loDT.Rows.Count = 0 Then Return ""

            ldDate = loDT.Rows(0).Item(0).ToString
            lsField = lsField & Format(ldDate, "yy")
            lnCounter = Len(lsField)
            Debug.Print(lsField, lnCounter)
        Else
            lnCounter = Len(lsField)
        End If

        lsSQL = "SELECT " & Field & _
                 " FROM " & Table & _
                 " WHERE " & Field & " LIKE " & strParm(lsField & "%") & _
                 " ORDER BY " & Field & " DESC LIMIT 1"

        Try
            loDA.SelectCommand = New MySqlCommand(lsSQL, Connection)
        Catch ex As MySqlException
            MsgBox(ex.Message)
            Throw ex
        End Try

        loDT.Clear()
        loDA.Fill(loDT)
        If loDT.Rows.Count = 0 Then
            lsSQL = ""

            loDA.FillSchema(loDS, SchemaType.Source)
            lnLen = loDS.Tables(0).Columns(0).MaxLength
            lnCode = 1

            lsSQL = IIf(ByBranch, Branch, "")
            If YearFormat Then
                lsSQL = lsSQL & Format(ldDate, "yy")
                lnCounter = Len(Branch) + 2
            Else
                lnCounter = Len(lsSQL)
            End If
        Else
            lsSQL = loDT.Rows(0).Item(Field)
            lnLen = Len(lsSQL)
            lnCode = CLng(Mid(lsSQL, lnCounter + 1)) + 1
        End If

        If lsSQL = "" Then
            lnCode = 1
            'lnCode = CLng(Mid(lsSQL, lnCounter + 1)) + 1
        Else
            lsSQL = IIf(ByBranch, Branch, "")
            If YearFormat Then
                lsSQL = lsSQL & Format(ldDate, "yy")
                lnCounter = Len(Branch) + 2
            Else
                lnCounter = Len(lsSQL)
            End If
        End If

        If lsSQL = "" Then
            Return Format(lnCode, lsStr.PadRight(lnLen, "0"))
        Else
            Return Left(lsSQL, lnCounter) & Format(lnCode, lsStr.PadRight(lnLen - lnCounter, "0"))
        End If

    End Function

    Public Function strParm(ByVal Value As String) As String
        Value = Replace(Value, Chr(34), Chr(34) & Chr(34))
        Value = Replace(Value, "\", "\\")
        Value = "'" & Replace(Value, "'", "''") & "'"
        Return Value
    End Function

    Public Function dateParm(ByVal Value As Date) As String
        dateParm = "'" & Format(Value, "yyyy-MM-dd") & "'"
    End Function

    Public Function datetimeParm(ByVal Value As Date) As String
        datetimeParm = "'" & Format(Value, "yyyy-MM-dd H:mm:ss") & "'"
    End Function

    Public Function KwikSearch(ByRef foAppDriver As GRider _
                   , ByVal fsSource As String _
                   , ByVal fsSearch As String _
                   , Optional ByVal fsColName As String = "" _
                   , Optional ByVal fsColTitle As String = "" _
                   , Optional ByVal fsColFormat As String = "" _
                   , Optional ByVal fsFieldName As String = "" _
                   , Optional ByVal fnDefCol As Integer = 1) As DataRow
        Dim loLookup As Lookup

        loLookup = New Lookup(foAppDriver, fsSource, fsSearch, fsColName, fsColTitle, fsColFormat, fsFieldName, fnDefCol)
        loLookup.ShowDialog()

        If loLookup.Cancelled Then Return Nothing

        Return loLookup.SelectedRow
        loLookup.Dispose()
    End Function

    Public Function KwikSearch(ByRef foAppDriver As GRider _
                   , ByVal fsSource As String _
                   , ByVal fbShowCondition As Boolean _
                   , ByVal fsSearch As String _
                   , Optional ByVal fsColName As String = "" _
                   , Optional ByVal fsColTitle As String = "" _
                   , Optional ByVal fsColFormat As String = "" _
                   , Optional ByVal fsFieldName As String = "" _
                   , Optional ByVal fnDefCol As Integer = 1) As DataRow
        Dim loLookup As Lookup

        loLookup = New Lookup(foAppDriver, fsSource, fsSearch, fsColName, fsColTitle, fsColFormat, fsFieldName, fnDefCol)
        loLookup.ShowCondition = fbShowCondition
        loLookup.ShowDialog()

        If loLookup.Cancelled Then Return Nothing

        Return loLookup.SelectedRow
        loLookup.Dispose()
    End Function

    Public Function KwikSearch(ByRef foAppDriver As GRider _
                   , ByVal fsSource As String _
                   , ByVal fsSearch As String _
                   , ByVal fbDisplay As Boolean _
                   , Optional ByVal fsColName As String = "" _
                   , Optional ByVal fsColTitle As String = "" _
                   , Optional ByVal fsColFormat As String = "" _
                   , Optional ByVal fsFieldName As String = "" _
                   , Optional ByVal fnDefCol As Integer = 1) As DataRow
        Dim loLookup As Lookup

        loLookup = New Lookup(foAppDriver, fsSource, fsSearch, fsColName, fsColTitle, fsColFormat, fsFieldName, fnDefCol)
        loLookup.ShowDialog()

        If loLookup.Cancelled Then Return Nothing

        Return loLookup.SelectedRow
        loLookup.Dispose()
    End Function

    Public Function ExecuteQuery(ByVal fsSQLCmd As String, ByRef fsConnection As MySqlConnection) As DataTable
        Dim loDT As New DataTable
        Dim loDA As New MySqlDataAdapter

        Try
            loDA.SelectCommand = New MySqlCommand(fsSQLCmd, fsConnection)
        Catch ex As MySqlException
            MsgBox(ex.Message)
            Throw ex
        End Try

        loDA.Fill(loDT)
        Return loDT
    End Function

    Public Function ExecuteQuery(ByRef fsMySQLCmd As MySqlCommand, ByRef fsConnection As MySqlConnection) As DataTable
        Dim loDT As New DataTable
        Dim loDA As New MySqlDataAdapter

        Try
            loDA.SelectCommand = fsMySQLCmd
        Catch ex As MySqlException
            MsgBox(ex.Message)
            Throw ex
        End Try

        loDA.Fill(loDT)
        Return loDT
    End Function

    Public Function ExecuteActionQuery(ByVal fsSQLCmd As String, ByRef fsConnection As MySqlConnection) As Integer
        Dim loSC As New MySqlCommand(fsSQLCmd, fsConnection)
        Dim lnAffected As Integer

        Try
            lnAffected = loSC.ExecuteNonQuery()
        Catch ex As MySqlException
            MsgBox(ex.Message)
            Throw ex
        End Try

        Return lnAffected
    End Function

    Public Function GetSplitedName(ByVal lsName As String) As Object
        Dim lasName() As String
        Dim lsLName As String, lsFName As String, lsMName As String
        Dim lnCtr As Integer

        lsLName = ""
        lsFName = ""
        lsMName = ""

        If lsName = "" Then
            ReDim lasName(0)
            lsLName = ""
        Else
            lasName = Split(Trim(lsName), ",")
            lsLName = Trim(UCase(Left(lasName(0), 1)) & Mid(lasName(0), 2))
        End If

        If UBound(lasName) > 0 Then
            If Trim(lasName(1)) <> "" Then
                lasName = Split(Trim(lasName(1)), " ")
                lsFName = Trim(UCase(Left(lasName(0), 1)) & Mid(lasName(0), 2))
                If UBound(lasName) > 0 Then
                    lsMName = Trim(UCase(Left(lasName(UBound(lasName)), 1)) & Mid(lasName(UBound(lasName)), 2))
                    For lnCtr = 1 To UBound(lasName) - 1
                        lsFName = lsFName & " " & Trim(UCase(Left(lasName(lnCtr), 1)) & Mid(lasName(lnCtr), 2))
                    Next
                End If
            End If
        End If

        ReDim lasName(2)
        lasName(0) = lsLName
        lasName(1) = lsFName
        lasName(2) = lsMName

        Return lasName
    End Function

    'Translated from VB6
    'KALYPTUS - 2013.04.02 
    Public Function ADO2SQL( _
                       ByVal foDta As DataTable _
                     , ByVal fsTableNme As String _
                     , Optional ByVal fsFilter As String = "" _
                     , Optional ByVal fsModified As String = "" _
                     , Optional ByVal fdModified As String = "" _
                     , Optional ByVal fsExcluded As String = "") As String
        Dim lsSQL As String
        Dim loDta As DataTable
        Dim foRS As DataTable
        'If this is an update and no changes has been detected then exit immediately the function

        loDta = New DataTable
        foRS = foDta.Copy
        If fsFilter <> "" Then
            loDta = foRS.GetChanges
            If IsNothing(loDta) Then
                Return ""
            End If
            foRS.RejectChanges()
        End If

        lsSQL = ""

        If fsModified <> "" Then fsExcluded = fsExcluded & "»sModified"
        If fdModified <> "" Then fsExcluded = fsExcluded & "»dModified"

        For lnCtr = 0 To foRS.Columns.Count - 1 Step 1
            If InStr(fsExcluded, foRS.Columns(lnCtr).ColumnName) = 0 Then
                If fsFilter = "" Then
                    lsSQL = lsSQL & ", " & AddField(foRS.Columns(lnCtr).ColumnName, FieldParam(foRS.Columns(lnCtr).DataType.Name, foRS(0).Item(lnCtr)))
                Else
                    If IsDBNull(foRS(0).Item(lnCtr)) Then
                        If IsDBNull(loDta(0).Item(lnCtr)) = False Then
                            lsSQL = lsSQL & ", " & AddField(loDta.Columns(lnCtr).ColumnName, FieldParam(loDta.Columns(lnCtr).DataType.Name, loDta(0).Item(lnCtr)))
                        End If
                    ElseIf IsDBNull(loDta(0).Item(lnCtr)) Then
                        If IsDBNull(foRS(0).Item(lnCtr)) = False Then
                            lsSQL = lsSQL & ", " & AddField(loDta.Columns(lnCtr).ColumnName, FieldParam(loDta.Columns(lnCtr).DataType.Name, loDta(0).Item(lnCtr)))
                        End If
                    ElseIf foRS(0).Item(lnCtr) <> loDta(0).Item(lnCtr) Then
                        lsSQL = lsSQL & ", " & AddField(loDta.Columns(lnCtr).ColumnName, FieldParam(loDta.Columns(lnCtr).DataType.Name, loDta(0).Item(lnCtr)))
                    End If
                End If
            End If
        Next

        If lsSQL <> "" Then
            If fsModified <> "" Then lsSQL = lsSQL & ", sModified = " & strParm(fsModified)
            If fdModified <> "" Then lsSQL = lsSQL & ", dModified = " & datetimeParm(fdModified)

            If fsFilter = "" Then
                lsSQL = "INSERT INTO " & fsTableNme & " SET " & Mid(lsSQL, 3)
            Else
                lsSQL = "UPDATE " & fsTableNme & " SET " & Mid(lsSQL, 3) & " WHERE " & fsFilter
            End If
        End If

        Return lsSQL
    End Function

    'Translated from VB6
    'KALYPTUS - 2013.04.02 
    Public Function AddField(ByVal fsFieldName As String, ByVal fvValue As Object) As String
        AddField = fsFieldName & " = " & fvValue
    End Function

    ' XerSys- 2013.05.22
    '   Allows client to specify current row for multi-row processing
    Public Function ADO2SQL( _
                       ByVal foDta As DataTable _
                     , ByVal fnRow As Integer _
                     , ByVal fsTableNme As String _
                     , Optional ByVal fsFilter As String = "" _
                     , Optional ByVal fsModified As String = "" _
                     , Optional ByVal fdModified As String = "" _
                     , Optional ByVal fsExcluded As String = "") As String
        Dim lsSQL As String
        Dim loDta As DataTable
        Dim loRS As DataTable
        Dim lbEdit As Boolean

        loDta = New DataTable
        loRS = New DataTable

        loRS = foDta.Clone()
        loRS.ImportRow(foDta(fnRow))

        'loDta = New DataTable
        'loDta = foDta.Clone
        'For lnctr = 0 To foDta.Rows.Count - 1
        '    loDta.ImportRow(foDta.Rows(lnctr))
        'Next

        ' instead of basing the update mode from the parameter, 
        '   check it from the row status
        loDta = loRS.GetChanges(DataRowState.Added)
        If IsNothing(loDta) Then
            lbEdit = True

            ' edit mode
            If fsFilter = "" Then Return ""

            loDta = loRS.GetChanges
            If IsNothing(loDta) Then
                Return ""
            End If
            loRS.RejectChanges()
        End If

        lsSQL = ""

        If fsModified <> "" Then fsExcluded = fsExcluded & "»sModified"
        If fdModified <> "" Then fsExcluded = fsExcluded & "»dModified"

        For lnCtr = 0 To loRS.Columns.Count - 1 Step 1
            If InStr(fsExcluded, loRS.Columns(lnCtr).ColumnName) = 0 Then
                If lbEdit Then
                    If IsDBNull(loRS(0).Item(lnCtr)) Then
                        If IsDBNull(loDta(0).Item(lnCtr)) = False Then
                            lsSQL = lsSQL & ", " & AddField(loDta.Columns(lnCtr).ColumnName, FieldParam(loDta.Columns(lnCtr).DataType.Name, loDta(0).Item(lnCtr)))
                        End If
                    ElseIf IsDBNull(loDta(0).Item(lnCtr)) Then
                        If IsDBNull(loRS(lnCtr).Item(lnCtr)) = False Then
                            lsSQL = lsSQL & ", " & AddField(loDta.Columns(lnCtr).ColumnName, FieldParam(loDta.Columns(lnCtr).DataType.Name, loDta(0).Item(lnCtr)))
                        End If
                    ElseIf loRS(0).Item(lnCtr) <> loDta(0).Item(lnCtr) Then
                        lsSQL = lsSQL & ", " & AddField(loDta.Columns(lnCtr).ColumnName, FieldParam(loDta.Columns(lnCtr).DataType.Name, loDta(0).Item(lnCtr)))
                    End If
                Else
                    lsSQL = lsSQL & ", " & AddField(loRS.Columns(lnCtr).ColumnName, FieldParam(loRS.Columns(lnCtr).DataType.Name, loRS(0).Item(lnCtr)))
                End If
            End If
        Next

        If lsSQL <> "" Then
            If fsModified <> "" Then lsSQL = lsSQL & ", sModified = " & strParm(fsModified)
            If fdModified <> "" Then lsSQL = lsSQL & ", dModified = " & datetimeParm(fdModified)

            If lbEdit Then
                lsSQL = "UPDATE " & fsTableNme & " SET " & Mid(lsSQL, 3) & " WHERE " & fsFilter
            Else
                lsSQL = "INSERT INTO " & fsTableNme & " SET " & Mid(lsSQL, 3)
            End If
        End If

        Return lsSQL
    End Function

    'Translated from VB6
    'KALYPTUS - 2013.04.02 
    Public Function FieldParam(ByVal lnFieldType As String, ByVal lvFieldValue As Object) As String
        Select Case lnFieldType
            Case "String", "Char" ' string
                If IsDBNull(lvFieldValue) Then
                    FieldParam = "NULL"
                Else
                    FieldParam = strParm(lvFieldValue)
                End If
            Case "SByte""Byte", "Int16", "Int32", "Int64", "UInt16", "UInt32", "UInt64"      ' numeric without decimal point
                If Not IsNumeric(lvFieldValue) = True Or IsDBNull(lvFieldValue) = True Then
                    'set the value to 0 if value is empty/null
                    FieldParam = 0
                Else
                    FieldParam = lvFieldValue
                End If
            Case "Decimal", "Double", "Single"             ' numeric with decimal point
                If Not IsNumeric(lvFieldValue) = True Or IsDBNull(lvFieldValue) = True Then
                    'set the value to 0 if value is empty/null
                    FieldParam = 0.0
                Else
                    FieldParam = lvFieldValue
                End If
            Case "DateTime" ' datetime
                If IsDBNull(lvFieldValue) Or Not IsDate(lvFieldValue) Then
                    FieldParam = "NULL"
                Else
                    FieldParam = datetimeParm(lvFieldValue)
                End If
            Case "TimeSpan" 'time
                MsgBox(Not IsDate(lvFieldValue))
                If IsDBNull(lvFieldValue) Or Not IsDate(lvFieldValue) Then
                    FieldParam = "NULL"
                Else
                    FieldParam = "'" & Format(lvFieldValue, "hh:mm:ss") & "'"
                End If
            Case Else
                If IsDBNull(lvFieldValue) Then
                    FieldParam = "NULL"
                Else
                    FieldParam = lvFieldValue
                End If
        End Select
    End Function

    'Translated from VB6
    'KALYPTUS - 2013.04.02 
    Public Function IFNull(ByVal Value As Object, Optional ByVal Repl As Object = Nothing) As Object
        IFNull = IIf(IsDBNull(Value), Repl, IIf(IsNothing(Value), "", Value))
    End Function

    'MAC [04-16-13]
    Public Function TitleCase(Optional ByVal Value As String = "") As String
        Return StrConv(Value, VbStrConv.ProperCase)
    End Function
    Public Function LCase(Optional ByVal Value As String = "") As String
        Return StrConv(Value, VbStrConv.Lowercase)
    End Function
    Public Function UCase(Optional ByVal Value As String = "") As String
        Return StrConv(Value, VbStrConv.Uppercase)
    End Function

    'MAC [04-25-13]
    Public Function StringFormat(ByVal sValue As String,
                              Optional ByVal sFormat As String = "",
                              Optional ByVal sSeparator As String = "") As String
        Dim lasFormat() As String
        Dim lsString As String = ""

        If sSeparator = "" And sSeparator.Length <> 1 Then GoTo endProc
        If sFormat = "" Then GoTo endproc

        If InStr(sFormat, sSeparator, CompareMethod.Text) = 0 Then GoTo endproc

        lasFormat = Split(sFormat, sSeparator)

        For lnCtr = 0 To UBound(lasFormat)
            lsString = lsString + Strings.Left(sValue, lasFormat(lnCtr).Length) & _
                        IIf(lnCtr = UBound(lasFormat), "", sSeparator)
            sValue = Strings.Right(sValue, sValue.Length - lasFormat(lnCtr).Length)
        Next

endProc:
        Return lsString
    End Function

    'dissect a given string for SQL-IN Clause use...
    'eg. strDissect("10") returns "'1', '0'"
    Public Function strDissect(ByVal fsValue As String) As String
        Dim lnCtr As Int32
        Dim lsStr As String

        If Len(fsValue) > 0 Then
            lsStr = strParm(Mid(fsValue, 1, 1))
            If Len(fsValue) > 1 Then
                For lnCtr = 2 To Len(fsValue)
                    lsStr = lsStr & "," & strParm(Mid(fsValue, lnCtr, 1))
                Next
            End If
        Else
            lsStr = "''"
        End If

        Return lsStr
    End Function

    Public Sub showModalForm(ByVal foForm As Form, ByVal foMdi As Form)
        foForm.MdiParent = foMdi
        foForm.MaximizeBox = False
        foForm.TopMost = True
        foForm.Show()
    End Sub

    Public Function isValidEmail(ByVal emailAddress As String) As Boolean
        ' Dim email As New Regex("^(?<user>[^@]+)@(?<host>.+)$")
        Dim email As New Regex("([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})")
        If email.IsMatch(emailAddress) Then
            Return True
        Else
            Return False
        End If
    End Function

    'FROM CR Inventory
    'This method can handle all events using EventHandler
    Public Sub grpEventHandler(ByVal foParent As Control, ByVal foType As Type, ByVal fsGroupNme As String, ByVal fsEvent As String, ByVal foAddress As EventHandler)
        Dim loTxt As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = foType Then
                'Handle events for this controls only
                If LCase(Mid(loTxt.Name, 1, Len(fsGroupNme))) = LCase(fsGroupNme) Then
                    If foType = GetType(TextBox) Then
                        Dim loObj = DirectCast(loTxt, TextBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(CheckBox) Then
                        Dim loObj = DirectCast(loTxt, CheckBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(Button) Then
                        Dim loObj = DirectCast(loTxt, Button)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    End If
                End If 'LCase(Mid(loTxt.Name, 1, 8)) = "txtfield"
            Else
                If loTxt.HasChildren Then
                    Call grpEventHandler(loTxt, foType, fsGroupNme, fsEvent, foAddress)
                End If
            End If
        Next 'loTxt In loControl.Controls
    End Sub

    'This method can handle all events using CancelEventHandler
    Public Sub grpCancelHandler(ByVal foParent As Control, ByVal foType As Type, ByVal fsGroupNme As String, ByVal fsEvent As String, ByVal foAddress As System.ComponentModel.CancelEventHandler)
        Dim loTxt As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = foType Then
                'Handle events for this controls only
                If LCase(Mid(loTxt.Name, 1, Len(fsGroupNme))) = LCase(fsGroupNme) Then
                    If foType = GetType(TextBox) Then
                        Dim loObj = DirectCast(loTxt, TextBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(CheckBox) Then
                        Dim loObj = DirectCast(loTxt, CheckBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(Button) Then
                        Dim loObj = DirectCast(loTxt, Button)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    End If
                End If 'LCase(Mid(loTxt.Name, 1, 8)) = "txtfield"
            Else
                If loTxt.HasChildren Then
                    Call grpCancelHandler(loTxt, foType, fsGroupNme, fsEvent, foAddress)
                End If
            End If
        Next 'loTxt In loControl.Controls
    End Sub

    'This method can handle all events using KeyEventHandler
    Public Sub grpKeyHandler(ByVal foParent As Control, ByVal foType As Type, ByVal fsGroupNme As String, ByVal fsEvent As String, ByVal foAddress As KeyEventHandler)
        Dim loTxt As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = foType Then
                'Handle events for this controls only
                If LCase(Mid(loTxt.Name, 1, Len(fsGroupNme))) = LCase(fsGroupNme) Then
                    If foType = GetType(TextBox) Then
                        Dim loObj = DirectCast(loTxt, TextBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(CheckBox) Then
                        Dim loObj = DirectCast(loTxt, CheckBox)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    ElseIf foType = GetType(Button) Then
                        Dim loObj = DirectCast(loTxt, Button)
                        Dim loEvent As EventInfo = foType.GetEvent(fsEvent)
                        loEvent.AddEventHandler(loObj, foAddress)
                    End If
                End If 'LCase(Mid(loTxt.Name, 1, 8)) = "txtfield"
            Else
                If loTxt.HasChildren Then
                    Call grpKeyHandler(loTxt, foType, fsGroupNme, fsEvent, foAddress)
                End If
            End If
        Next 'loTxt In loControl.Controls
    End Sub

    'This method can handle all events using EventHandler
    Public Function FindTextBox(ByVal foParent As Control, ByVal fsName As String) As Control
        Dim loTxt As Control
        Static loRet As Control
        For Each loTxt In foParent.Controls
            If loTxt.GetType = GetType(TextBox) Then
                'Handle events for this controls only
                If LCase(loTxt.Name) = LCase(fsName) Then
                    loRet = loTxt
                End If
            Else
                If loTxt.HasChildren Then
                    Call FindTextBox(loTxt, fsName)
                End If
            End If
        Next 'loTxt In loControl.Controls

        Return loRet
    End Function

    Public Sub SortTable(ByRef foDTTable As DataTable, ByVal fsColumn As String)
        Dim oDTCopy As New DataTable
        oDTCopy = foDTTable.Copy()
        foDTTable.Clear()
        Dim aDRRows() As DataRow
        aDRRows = oDTCopy.Select("", fsColumn + " asc")
        Dim oRow As DataRow
        For Each oRow In aDRRows
            foDTTable.ImportRow(oRow)
        Next
        oDTCopy.Clear()
    End Sub

    Public Function HexToString(ByVal HexToStr As String) As String
        Dim strTemp As String
        Dim strReturn As String = ""
        Dim I As Long
        For I = 1 To Len(HexToStr) Step 2
            strTemp = Chr(Val("&H" & Mid$(HexToStr, I, 2)))
            strReturn = strReturn & strTemp
        Next I
        HexToString = strReturn
    End Function

    Public Function StringToHex(ByVal StrToHex As String) As String
        Dim strTemp As String
        Dim strReturn As String = ""
        Dim I As Long
        For I = 1 To Len(StrToHex)
            strTemp = Hex$(Asc(Mid$(StrToHex, I, 1)))
            If Len(strTemp) = 1 Then strTemp = "0" & strTemp
            strReturn = strReturn & strTemp
        Next I
        StringToHex = strReturn
    End Function

    Public Enum xeUserStatus
        SUSPENDED = 0
        ACTIVE = 1
    End Enum

    Public Enum xeLogical
        NO = 0
        YES = 1
    End Enum

    Public Enum xeUserRights
        DATAENTRY = 1
        SUPERVISOR = 2
        MANAGER = 4
        AUDIT = 8
        SYSADMIN = 16
        SYSOWNER = 32
        ENGINEER = 64
        SYSMASTER = 128
    End Enum

    Public Enum xeRecordStat
        RECORD_EMPTY = 0
        RECORD_NEW = 1
        RECORD_EXISTING = 2
        RECORD_DELETED = 3
    End Enum

    Public Enum xeTranStat
        TRANS_OPEN = 0
        TRANS_CLOSED = 1
        TRANS_POSTED = 2
        TRANS_CANCELLED = 3
        TRANS_UNKNOWN = 4
    End Enum

    Public Enum xeAccountStat
        ACTIVE = 0
        CLOSED = 1
        DEAD = 2
        IMPOUNDED = 3
        DISCARDED = 4

    End Enum

    Public Enum xeEditMode
        MODE_UNKNOWN = -1
        MODE_READY = 0
        MODE_ADDNEW = 1
        MODE_UPDATE = 2
        MODE_DELETE = 3
    End Enum

    'MAC (07-08-2013)
    Public Enum xeRakingMode
        PROB = 0
        RnF = 1
        MNGR = 2
    End Enum

    Public Enum MCLocation
        xeLocWarehouse = 0
        xeLocBranch = 1
        xeLocSupplier = 2
        xeLocCustomer = 3
        xeLocUnknown = 4
        xeLocServiceCenter = 5
    End Enum

    Public Const xsDATE_SHORT As String = "MM-dd-yyyy"
    Public Const xsDATE_MEDIUM As String = "MMM dd, yyyy"
    Public Const xsDATE_LONG As String = "MMMM dd, yyyy"
    Public Const xsDATE_TIME As String = "yyyy-MM-dd hh:mm:ss"
    Public Const xsNULL_DATE As String = "1900-01-01"
    Public Const xsDECIMAL As String = "#,##0.00"
    Public Const xsINTEGER As String = "#,##0"
End Module
