Imports System.Windows.Forms
Imports System.Drawing

Public Class frmUserDisc
    Private p_bCancelled As Boolean = True
    Private p_sLogNamex As String
    Private p_sPassword As String

    Public ReadOnly Property Cancelled As Boolean
        Get
            Cancelled = p_bCancelled
        End Get
    End Property

    Public ReadOnly Property LogName As String
        Get
            Return p_sLogNamex
        End Get
    End Property

    Public ReadOnly Property Password As String
        Get
            Return p_sPassword
        End Get
    End Property

    Private Sub frmUserDisc_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Call initForm()
    End Sub

    Private Sub frmUserDisc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                p_bCancelled = True
                Me.Hide()
            Case Keys.Return, Keys.Down
                SetNextFocus()
            Case Keys.Up
                SetPreviousFocus()
        End Select
    End Sub

    Private Sub frmUserDisc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call initForm()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub initForm()
        With Me
            .Location = New Point(507, 90)

            .txtUserName.Text = ""
            .txtPassword.Text = ""
            .txtUserName.Focus()
        End With
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdButton1.Click, cmdButton2.Click
        Select Case DirectCast(sender, Button).Name
            Case cmdButton1.Name 'ok
                p_sLogNamex = txtUserName.Text
                p_sPassword = txtPassword.Text

                p_bCancelled = False
                Me.Hide()
            Case cmdButton2.Name 'cancel
                p_bCancelled = True
                Me.Hide()
        End Select
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub PreventFlicker()
        With Me
            .SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            .SetStyle(ControlStyles.UserPaint, True)
            .SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            .UpdateStyles()
        End With
    End Sub

End Class