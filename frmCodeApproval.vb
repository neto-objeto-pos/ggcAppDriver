Imports System.Windows.Forms

Public Class frmCodeApproval

    Private pnLoadx As Integer
    Private poControl As Control

    Private pbCancel As Boolean
    Private psAppPath As String
    Private p_oAppDrivr As GRider

    Private p_sUserIDxx As String
    Private p_sIssueexx As String
    Private p_cIssueexx As String
    Private p_sCodeAprv As String


    Public ReadOnly Property UserID As String
        Get
            UserID = p_sUserIDxx
        End Get
    End Property

    Public ReadOnly Property Issuee As String
        Get
            Issuee = p_sIssueexx
        End Get
    End Property

    Public ReadOnly Property IssueeType As String
        Get
            IssueeType = p_cIssueexx
        End Get
    End Property

    Public ReadOnly Property CodeApproval As String
        Get
            CodeApproval = p_sCodeAprv
        End Get
    End Property

    Public WriteOnly Property AppDriver As GRider
        Set(value As GRider)
            p_oAppDrivr = value
        End Set
    End Property

    Public WriteOnly Property AppPath As String
        Set(value As String)
            psAppPath = value
        End Set
    End Property

    Public ReadOnly Property Cancel As Boolean
        Get
            Cancel = pbCancel
        End Get
    End Property

    Private Sub frmCodeApproval_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
            Dim loTxt As Control

            loTxt = Nothing
            If TypeOf poControl Is TextBox Then
                loTxt = CType(poControl, System.Windows.Forms.TextBox)

                '*********************
                Dim loIndex As Integer
                loIndex = Val(Mid(loTxt.Name, 9))

                If loTxt.Name = "txtIssuee" Then
                    Call getIssuee(txtIssuee.Text, e.KeyCode = Keys.F3)
                End If
                '*********************
            End If

            If TypeOf poControl Is TextBox Then
                SelectNextControl(loTxt, True, True, True, True)
            End If
        End If
    End Sub

    Private Sub txtApprovalCode_Validate(Cancel As Boolean)
        p_sCodeAprv = txtApprovalCode.Text
    End Sub

    Private Sub getIssuee(ByVal fsValue As String, ByVal fbSearch As Boolean)
        Dim lsSQL As String
        Dim lsConditn As String

        If fsValue = "" And fbSearch = False Then
            p_sUserIDxx = ""
            p_sIssueexx = ""
            p_cIssueexx = ""
            GoTo endProc
        End If

        lsSQL = "SELECT" & _
                       "  a.sUserIDxx" & _
                       ", c.sCompnyNm sIssueexx" & _
                       ", b.sEmpLevID" & _
                       ", b.sDeptIDxx" & _
               " FROM xxxSysUser a" & _
                   ", Employee_Master001 b" & _
                       " LEFT JOIN Client_Master c ON b.sEmployID = c.sClientID" & _
               " WHERE a.sEmployNo = b.sEmployID" & _
               " GROUP BY sEmployID"

        lsConditn = ""
        If fbSearch Then
            If fsValue <> "" Then
                lsConditn = "c.sCompnyNm LIKE " & strParm(fsValue & "%")
            End If
        Else
            lsConditn = "c.sCompnyNm = " & strParm(fsValue)
        End If

        If lsConditn <> "" Then
            lsSQL = AddCondition(lsSQL, lsConditn)
        End If

        Debug.Print(lsSQL)

        Dim loDta As DataTable
        loDta = p_oAppDrivr.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_sUserIDxx = ""
            p_sIssueexx = ""
            p_cIssueexx = ""
            GoTo endProc
        ElseIf loDta.Rows.Count = 1 Then
            p_sUserIDxx = loDta(0).Item("sUserIDxx")
            p_sIssueexx = loDta(0).Item("sIssueexx")

            If loDta(0).Item("sEmpLevID") = "4" Then
                p_cIssueexx = "0"
            Else
                Select Case loDta(0).Item("sDeptIDxx")
                    Case "021"   'Human Capital Management
                        p_cIssueexx = "1"
                    Case "022"   'Credit Support Services
                        p_cIssueexx = "2"
                    Case "034"   'Compliance Management
                        p_cIssueexx = "3"
                    Case "025"   'Marketing & Promotions
                        p_cIssueexx = "4"
                    Case "027"    'After Sales Management
                        p_cIssueexx = "5"
                    Case "035"   'Telemarketing
                        p_cIssueexx = "6"
                    Case "024"   'Supply Chain Management
                        p_cIssueexx = "7"
                    Case "026"   'Management Information Systems
                        p_cIssueexx = "X"
                    Case Else
                        p_cIssueexx = ""
                End Select
            End If
        Else

            Dim loRow As DataRow = KwikSearch(p_oAppDrivr _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sUserIDxx»sIssueexx" _
                                             , "ID»Issuee", _
                                             , "a.sUserIDxx»c.sCompnyNm" _
                                             , IIf(fbSearch, 1, 0))

            If IsNothing(loRow) Then
                p_sUserIDxx = ""
                p_sIssueexx = ""
                p_cIssueexx = ""
                GoTo endProc
            Else

                p_sUserIDxx = loRow(0).Item("sUserIDxx")
                p_sIssueexx = loRow(0).Item("sIssueexx")

                If loRow(0).Item("sEmpLevID") = "4" Then
                    p_cIssueexx = "0"
                Else
                    Select Case loRow(0).Item("sDeptIDxx")
                        Case "021"   'Human Capital Management
                            p_cIssueexx = "1"
                        Case "022"   'Credit Support Services
                            p_cIssueexx = "2"
                        Case "034"   'Compliance Management
                            p_cIssueexx = "3"
                        Case "025"   'Marketing & Promotions
                            p_cIssueexx = "4"
                        Case "027"    'After Sales Management
                            p_cIssueexx = "5"
                        Case "035"   'Telemarketing
                            p_cIssueexx = "6"
                        Case "024"   'Supply Chain Management
                            p_cIssueexx = "7"
                        Case "026"   'Management Information Systems
                            p_cIssueexx = "X"
                        Case Else
                            p_cIssueexx = ""
                    End Select
                End If
            End If
        End If

endProc:
        txtIssuee.Text = p_sIssueexx
    End Sub

    Private Sub cmdOkay_Click(sender As System.Object, e As System.EventArgs) Handles cmdOkay.Click
        p_sCodeAprv = txtApprovalCode.Text
        pbCancel = False
        Me.Hide()
    End Sub

    Private Sub frmCodeApproval_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        pbCancel = True
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        pbCancel = True
        Me.Hide()
    End Sub

    Private Sub txtApprovalCode_GotFocus(sender As Object, e As System.EventArgs) Handles txtApprovalCode.GotFocus
        poControl = sender
    End Sub

    Private Sub txtIssuee_GotFocus(sender As Object, e As System.EventArgs) Handles txtIssuee.GotFocus
        poControl = sender
    End Sub

    Private Sub cmdOkay_GotFocus(sender As Object, e As System.EventArgs) Handles cmdOkay.GotFocus
        poControl = sender
    End Sub

    Private Sub cmdCancel_GotFocus(sender As Object, e As System.EventArgs) Handles cmdCancel.GotFocus
        poControl = sender
    End Sub
End Class


