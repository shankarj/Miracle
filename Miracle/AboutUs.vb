Public Class AboutUs

    Private Sub AboutUs_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Control Then
            MsgBox("Shankar")

        End If
    End Sub

    Private Sub AboutUs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Focus()
    End Sub


End Class