Public Class NewProj

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim FileNamed As String
        FolderBrowserDialog1.ShowDialog()
        FileNamed = FolderBrowserDialog1.SelectedPath
        TextBox2.Text = FileNamed
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If Not TextBox1.Text = Nothing And Not TextBox2.Text = Nothing Then
                IO.Directory.CreateDirectory(TextBox2.Text & "\" & TextBox1.Text)
                IO.Directory.CreateDirectory(TextBox2.Text & "\" & TextBox1.Text & "\Resources")
                IO.Directory.CreateDirectory(TextBox2.Text & "\" & TextBox1.Text & "\Resources\Common")
                ProjectPath = TextBox2.Text & "\" & TextBox1.Text & "\"
                Form1.ListBox2.Items.Add("Main.html")
                Form1.Text = TextBox1.Text & " - Miracle (Main.html)"
                ProjectName = TextBox1.Text
                Dim Writes As New IO.StreamWriter(ProjectPath & ProjectName & ".mirproj", False)
                Writes.WriteLine(ProjectName)
                Writes.WriteLine(ProjectPath)
                Writes.WriteLine("con##1")
                Writes.WriteLine("Main.html")
                Writes.Close()
                Me.Close()
                Form1.TabControl1.Visible = True
                Form1.Label1.Visible = True
                Form1.Label2.Visible = True
                Form1.ListBox2.Visible = True
                Form1.ListBox3.Visible = True
                Form1.MenuStrip1.Visible = True
                Form1.RichTextBox1.SaveFile(ProjectPath & "Main.html", 1)
                CurrentPage = "Main.html"
            Else
                MsgBox("Please enter all details.", MsgBoxStyle.Exclamation, "Miracle")
            End If
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub
End Class