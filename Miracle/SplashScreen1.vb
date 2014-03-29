Public NotInheritable Class SplashScreen1

 
    Dim l As Integer = 0
    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        l += 1
        If l = 3 Then
            Dim ShowSplash As New FrontForm
            ShowSplash.Show()
            Me.Close()
        End If
    End Sub
End Class
