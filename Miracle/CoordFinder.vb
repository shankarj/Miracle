Public Class CoordFinder
    Dim rectcord As Integer = 1
    Dim COORDS(4) As Integer
    Dim StartLoc As Point
    Dim EndLoc As Point
    Dim MoveLoc As Point

    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown

        Try

            If RadioButton1.Checked = True Then
                If Not rectcord > 2 Then
                    If rectcord = 1 Then
                        COORDS(0) = Label3.Text
                        COORDS(1) = Label4.Text
                        Panel1.Size = New Size(0, 0)
                        StartLoc = New Point(e.X + 245, e.Y + 73)
                        Panel1.Location = StartLoc
                        Panel1.Visible = True
                    End If

                    If rectcord = 2 Then
                        COORDS(2) = Label3.Text
                        COORDS(3) = Label4.Text
                        TextBox1.Text = COORDS(0) & "," & COORDS(1) & "," & COORDS(2) & "," & COORDS(3)
                        Panel1.Visible = False
                        Panel1.Location = New Point(0, 0)
                        rectcord = 0
                        COORDS(0) = 0
                        COORDS(1) = 0
                        COORDS(2) = 0
                        COORDS(3) = 0
                    End If

                    rectcord += 1
                Else
                    rectcord += 1
                End If
            End If
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        Label3.Text = e.X
        Label4.Text = e.Y

        If Panel1.Visible = True Then
            Panel1.Size = New Size((e.X + 245) - StartLoc.X, (e.Y + 73) - StartLoc.Y)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Clipboard.SetText(TextBox1.Text)
        MsgBox("Co-ordinates Copied.", MsgBoxStyle.OkOnly, "Miracle")

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click


    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        MsgBox("This is not available in this version.", MsgBoxStyle.Critical, "Miracle")

    End Sub

    Private Sub PictureBox1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseWheel

    End Sub

    Private Sub CoordFinder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class