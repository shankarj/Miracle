Public Class FrontForm

    Private Sub FrontForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        SingleWebPage = False
        Dim NextForm As New NewProj
        NextForm.Show()
        Me.Close()
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        SingleWebPage = False
        Try
            Dim FileNamed As String

            With OpenFileDialog1
                .Title = "Open"
                .InitialDirectory = "C:\"
                .FileName = Nothing
                .Filter = "Miracle Projects (*.mirproj) | *.mirproj"
                .ShowDialog()
                FileNamed = .FileName
            End With

            Dim Reads As New IO.StreamReader(FileNamed)
            ProjectName = Reads.ReadLine
            ProjectPath = Reads.ReadLine
            While Not Reads.EndOfStream
                If Reads.ReadLine = "con##1" Then
                    Dim TextToAdd As String = Reads.ReadLine
                    If Not TextToAdd = Nothing Then
                        Form1.ListBox2.Items.Add(TextToAdd)
                    End If
                Else
                    Dim TextToAdd As String = Reads.ReadLine
                    If Not TextToAdd = Nothing Then
                        Form1.ListBox3.Items.Add(TextToAdd)
                    End If
                End If
            End While

            Reads.Close()
            Form1.TabControl1.Visible = True
            Form1.Label1.Visible = True
            Form1.Label2.Visible = True
            Form1.ListBox2.Visible = True
            Form1.ListBox3.Visible = True
            Form1.MenuStrip1.Visible = True
            Form1.RichTextBox1.LoadFile(ProjectPath & "Main.html", RichTextBoxStreamType.PlainText)
            Form1.Text = ProjectName & " - Miracle (Main.html)"
            CurrentPage = "Main.html"

            Dim k As Int16
            k = (Form1.Size.Height - 90)
            Form1.TabControl1.Size = New Size(Form1.TabControl1.Width, k)



            Me.Close()
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try


    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        End
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Try
            Dim FileNamed As String

            With OpenFileDialog1
                .Title = "Open"
                .InitialDirectory = "C:\"
                .FileName = Nothing
                .Filter = "HTML Pages (*.html)|*.html|HTM Pages (*.htm)|*.htm"
                .ShowDialog()
                FileNamed = .FileName
            End With

            Form1.TabControl1.Visible = True
            Form1.Label1.Visible = True
            Form1.Label2.Visible = True
            Form1.ListBox2.Visible = True
            Form1.ListBox3.Visible = True
            Form1.MenuStrip1.Visible = True
            Form1.RichTextBox1.LoadFile(FileNamed, RichTextBoxStreamType.PlainText)
            Form1.Text = "Miracle (" & FileNamed & ")"
            SingleWebPage = True
            SinglePagePath = FileNamed
            Me.Close()

        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub
End Class