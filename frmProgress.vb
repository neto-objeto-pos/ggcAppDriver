Imports Org.Mentalis.Multimedia

Public Class frmProgress
    Private b_isCancel As Boolean
    Private p_sPstnPth As String

    Public ReadOnly Property isCancelled() As Boolean
        Get
            Return b_isCancel
        End Get
    End Property

    Public Property MaxValue As Long
        Get
            Return Me.prgBar.Maximum
        End Get
        Set(ByVal value As Long)
            Me.prgBar.Maximum = value
            Me.prgBar.Value = 0
        End Set
    End Property

    Public Property PistonInfo As String
        Get
            Return p_sPstnPth
        End Get
        Set(ByVal value As String)
            p_sPstnPth = value
        End Set
    End Property

    Public Sub ShowTitle(ByVal fsValue As String)
        Me.lblTitle.Text = fsValue
    End Sub

    Public Sub ShowProcess(ByVal fsValue As String)
        Me.lblProcess.Text = fsValue
        Me.prgBar.Value = Me.prgBar.Value + 1
    End Sub

    Public Sub ShowSuccess()
        Me.Close()
        MsgBox("Process was successfull!", vbOKOnly, "Information")
    End Sub

    Public Sub AbortProcess()
        Me.Close()
        MsgBox("Process was aborted!", vbOKOnly, "Warning")
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        If MsgBox("Do you really want to cancel this process?", vbYesNo, "Warning") = vbYes Then
            Me.Hide()
            b_isCancel = True
        End If
    End Sub

    Private Sub frmProgress_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'Dim v As New VideoFile(PistonInfo, Me.Panel1)
        'v.Repeat = True
        'v.Play()
    End Sub
End Class