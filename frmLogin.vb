Imports System.Windows.Forms

Public Class frmLogin
   Private p_bCancelled As Boolean = False

   Public ReadOnly Property Cancelled As Boolean
      Get
         Cancelled = p_bCancelled
      End Get
   End Property

   Private Sub frmLogin_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
      Call initForm()
   End Sub

   Private Sub frmLogin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
      MsgBox(sender.ToString)

      If e.Alt Or e.Control Then
         Return
      End If

      Select Case e.KeyCode
         Case Keys.Return
            SendKeys.Send("{TAB}")
         Case Keys.F3
            ' call for search method
      End Select
   End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Call initForm()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

   Private Sub initForm()
        With Me
            .txtUsername.Text = ""
            .txtPassword.Text = ""
            .txtUsername.Focus()
        End With
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
       Handles cmdButton1.Click, cmdButton2.Click

        p_bCancelled = DirectCast(sender, Button).Name.Equals(cmdButton2.Name)
        Select Case DirectCast(sender, Button).Name
            Case cmdButton1.Name 'ok
                Me.Hide()
            Case cmdButton2.Name 'cancel
                Me.Hide()
        End Select
    End Sub
End Class
