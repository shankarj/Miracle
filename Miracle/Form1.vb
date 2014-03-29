Imports System.Drawing.Imaging
Public Class Form1
    Dim saved As Boolean
    Dim LineNo As Integer = 1
    Dim o As New MediaPlayer.MediaPlayer
    
    Private Sub PreviewPageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If Not SingleWebPage = True Then
                If saved = False Then

                End If

                TabControl1.SelectTab(1)

                Try
                    WebBrowser1.Url = New Uri(ProjectPath & CurrentPage)
                Catch ex As Exception

                End Try

                Label3.Text = "Title : " & CurrentPage
            Else
                TabControl1.SelectTab(1)

                Try
                    WebBrowser1.Url = New Uri(SinglePagePath)
                    Label3.Text = "Title : " & SinglePagePath
                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub RichTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Try
            If e.KeyValue = 188 Then
                Dim o As New System.Drawing.Point(RichTextBox1.GetPositionFromCharIndex(RichTextBox1.SelectionStart))
                Panel1.Visible = True
                Panel1.Location = New System.Drawing.Point(o.X + 15, o.Y + 23)
                ListBox1.Focus()
            End If



            If e.KeyData = Keys.Back Then
                Panel1.Visible = False
            End If

            If e.KeyCode = Keys.Enter Then
                LineNo += 1
                ListBox4.Items.Add(LineNo)


            End If


            If e.KeyCode = Keys.Escape Then
                Panel1.Visible = False
            End If
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub LoadAttrib(ByVal TheText As String)

    End Sub

    Private Sub ListBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Dim l As String

            l = ListBox1.SelectedItem.ToString

            If l.Contains("!") Then
                l = ListBox1.SelectedItem.ToString & "-->"
            ElseIf l.Contains("hr") Then
                l = ListBox1.SelectedItem.ToString & ">"
            ElseIf l.Contains("br") Then
                l = ListBox1.SelectedItem.ToString & ">"
            ElseIf l.Contains("=") Then
                l = ListBox1.SelectedItem.ToString & "/>"
            Else
                l = ListBox1.SelectedItem.ToString & ">" & vbCrLf & vbCrLf & "</" & ListBox1.SelectedItem.ToString & ">"
            End If
            Panel1.Visible = False


            Clipboard.SetText(l)
            RichTextBox1.Paste()

            RichTextBox1.Focus()

            For index = LineNo To RichTextBox1.Lines.Count - 1
                LineNo += 1
                ListBox4.Items.Add(LineNo)
            Next


            If RichTextBox1.Focused = False Then
                RichTextBox1.Focus()

            End If
        End If


        If e.KeyCode = Keys.Escape Then
            Panel1.Visible = False
            RichTextBox1.Focus()
        End If

        If e.KeyCode = Keys.Back Then
            Panel1.Visible = False
            RichTextBox1.Focus()
        End If


    End Sub

    Private Sub ListBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            Dim l As String = ListBox1.SelectedItem.ToString & ">" & vbCrLf & vbCrLf & "</" & ListBox1.SelectedItem.ToString & ">"
            Panel1.Visible = False
            Clipboard.SetText(l)
            RichTextBox1.Paste()
            RichTextBox1.Focus()

            If RichTextBox1.Focused = False Then
                RichTextBox1.Focus()
            End If

        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub


    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    'Dim bmpScreenshot As Bitmap = New Bitmap(WebBrowser1.Width, WebBrowser1.Height, PixelFormat.Format32bppPArgb)

    '' Create a graphics object from the bitmap
    'Dim gfxScreenshot As Graphics = Graphics.FromImage(bmpScreenshot)

    '' Take a screen shot of just the button2 control
    'Dim k As New Point(WebBrowser1.Location.X, WebBrowser1.Location.Y + 50)

    'gfxScreenshot.CopyFromScreen(Me.PointToScreen(k), New Point(0, 0), WebBrowser1.Size, CopyPixelOperation.SourceCopy)

    '' Save the screenshot
    'bmpScreenshot.Save("D:\Panel1.jpg", ImageFormat.Jpeg)

    'End Sub


    Private Sub CodeViewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TabControl1.SelectTab(0)

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        NotifyIcon1.Dispose()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub


    Private Sub WebPageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If Not SingleWebPage = True Then
                Dim NewPageName As String
                NewPageName = InputBox("Please enter the new page name.", "Miracle", )
                If Not NewPageName = Nothing Then
                    FileSave(ProjectPath & CurrentPage, RichTextBox1)
                    RichTextBox1.Text = Nothing
                    CurrentPage = NewPageName & ".html"
                    FileSave(ProjectPath & CurrentPage, RichTextBox1)
                    ListBox2.Items.Add(CurrentPage)
                    Dim Writes As New IO.StreamWriter(ProjectPath & ProjectName & ".mirproj", True)
                    Writes.WriteLine("con##1")
                    Writes.WriteLine(CurrentPage)
                    Writes.Close()
                End If
            Else
                MsgBox("Please Create a Project to add multiple files.", MsgBoxStyle.Information, "Miracle")

            End If
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub FileSave(ByVal Path As String, ByVal TextboxName As RichTextBox)
        Try
            TextBox2.Text = TextboxName.Text
            Dim Saver As New IO.StreamWriter(Path)
            Saver.WriteLine(TextBox2.Text)
            Saver.Close()
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub ListBox2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        On Error Resume Next
        If Not ListBox2.SelectedItem.ToString = Nothing Then
            FileSave(ProjectPath & CurrentPage, RichTextBox1)
            RichTextBox1.LoadFile(ProjectPath & ListBox2.SelectedItem.ToString, RichTextBoxStreamType.PlainText)
            CurrentPage = ListBox2.SelectedItem.ToString
            Me.Text = ProjectName & " - Miracle (" & ListBox2.SelectedItem.ToString & ")"
        End If
    End Sub

    Private Sub PrieviewInBrowserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not SingleWebPage = True Then
            System.Diagnostics.Process.Start(ProjectPath & CurrentPage)
        Else
            System.Diagnostics.Process.Start(SinglePagePath)
        End If
    End Sub

    Private Sub ListBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next
        Dim k As Int16
        RichTextBox1.Focus()
        k = RichTextBox1.GetFirstCharIndexFromLine(ListBox4.SelectedItem.ToString - 1)
        RichTextBox1.SelectionStart = k
        RichTextBox1.Select(k, RichTextBox1.Lines(ListBox4.SelectedItem.ToString - 1).Length)
    End Sub

    Private Sub RichTextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Try
            If RichTextBox1.Lines(RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart)) = "<" Then
                Dim o As New System.Drawing.Point(RichTextBox1.GetPositionFromCharIndex(RichTextBox1.SelectionStart))
                Panel1.Visible = True
                Panel1.Location = New System.Drawing.Point(o.X + 15, o.Y + 23)
                ListBox1.Focus()
            End If
        Catch ex As Exception
        End Try
    End Sub


    Private Sub RichTextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        saved = False

        If LineNo > RichTextBox1.Lines.Count Then
            ListBox4.Items.Clear()
            LineNo = 0
        End If

        For index = LineNo To RichTextBox1.Lines.Count - 1
            LineNo += 1
            ListBox4.Items.Add(LineNo)
        Next


    End Sub

    Private Sub SaveAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If Not SingleWebPage = True Then
                NotifyIcon1.ShowBalloonTip(1000, "Miracle", ProjectName & " Saved.", ToolTipIcon.Info)

                FileSave(ProjectPath & CurrentPage, RichTextBox1)
            Else
                NotifyIcon1.ShowBalloonTip(1000, "Miracle", SinglePagePath & " Saved.", ToolTipIcon.Info)
                FileSave(SinglePagePath, RichTextBox1)
            End If
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub OpenWebsiteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub HTMLProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SingleWebPage = False
        Dim NextForm As New NewProj
        NextForm.Show()
        TabControl1.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        ListBox2.Visible = False
        ListBox3.Visible = False
        MenuStrip1.Visible = False
        RichTextBox1.Text = Nothing

    End Sub

    Private Sub ProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

            ListBox2.Items.Clear()
            ListBox3.Items.Clear()

            If Not SingleWebPage = True Then
                If saved = False Then
                    Dim Result As MsgBoxResult
                    Result = MsgBox("Current Project not Saved. Save now ?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Miracle")

                    If Result = MsgBoxResult.Yes Then
                        FileSave(ProjectPath & CurrentPage, RichTextBox1)
                    End If
                End If
            End If

            RichTextBox1.Text = Nothing

            Dim FileNamed As String

            With OpenFileDialog1
                .Title = "Open"
                .InitialDirectory = "C:\"
                .FileName = Nothing
                .Filter = "Miracle Projects (*.mirproj) | *.mirproj"
                .ShowDialog()
                FileNamed = .FileName
            End With
            If Not FileNamed = Nothing Then
                Dim Reads As New IO.StreamReader(FileNamed)
                ProjectName = Reads.ReadLine
                ProjectPath = Reads.ReadLine
                While Not Reads.EndOfStream
                    If Reads.ReadLine = "con##1" Then
                        Dim TextToAdd As String = Reads.ReadLine
                        If Not TextToAdd = Nothing Then
                            ListBox2.Items.Add(TextToAdd)
                        End If
                    ElseIf Reads.ReadLine = "res##1" Then
                        Dim TextToAdd As String = Reads.ReadLine
                        If Not TextToAdd = Nothing Then
                            ListBox3.Items.Add(TextToAdd)
                        End If
                    End If
                End While

                Reads.Close()
                RichTextBox1.LoadFile(ProjectPath & "Main.html", RichTextBoxStreamType.PlainText)
                Me.Text = ProjectName & " - Miracle (Main.html)"
                CurrentPage = "Main.html"
            End If

        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()

        End Try

    End Sub

    Private Sub SplitContainer1_Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub ChooseEnvironmentColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub EnvironmentColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With ColorDialog1
            .ShowDialog()
            TabPage2.BackColor = ColorDialog1.Color
        End With
    End Sub

    Private Sub TextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With ColorDialog1
            .ShowDialog()
            RichTextBox1.ForeColor = ColorDialog1.Color
        End With
    End Sub

    Private Sub BackGroundToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With ColorDialog1
            .ShowDialog()
            RichTextBox1.BackColor = ColorDialog1.Color
        End With
    End Sub

    Private Sub WebPageToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If Not SingleWebPage = False Then
                Dim FileNamed As String

                With OpenFileDialog1
                    .Title = "Open"
                    .InitialDirectory = "C:\"
                    .FileName = Nothing
                    .Filter = "HTML Pages (*.html)|*.html"
                    .ShowDialog()
                    FileNamed = .FileName
                End With

                RichTextBox1.LoadFile(FileNamed, RichTextBoxStreamType.PlainText)
                Text = "Miracle (" & FileNamed & ")"
                SingleWebPage = True
                SinglePagePath = FileNamed
            Else
                MsgBox("Close project before opening a new web page.", MsgBoxStyle.Critical, "Miracle")

            End If
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub AddAPageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If Not SingleWebPage = True Then
                Dim FileNamed As String
                Dim PageName As String
                With OpenFileDialog1
                    .Filter = "HTML Files (*.html)|*.html"
                    .FileName = Nothing
                    .ShowDialog()
                    FileNamed = .FileName
                    PageName = .SafeFileName
                End With

                IO.File.Copy(FileNamed, ProjectPath & PageName)
                CurrentPage = PageName
                Dim Writes As New IO.StreamWriter(ProjectPath & ProjectName & ".mirproj", True)
                Writes.WriteLine("con##1")
                Writes.WriteLine(CurrentPage)
                ListBox2.Items.Add(CurrentPage)
                Writes.Close()
                RichTextBox1.LoadFile(ProjectPath & CurrentPage, RichTextBoxStreamType.PlainText)
            Else
                MsgBox("Please create a project to add multiple files.", MsgBoxStyle.Information, "Miracle")
            End If

        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub PictureToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If Not SingleWebPage = True Then
            Try

                Dim PicName As String
                Dim FileNamed As String

                With OpenFileDialog1
                    .Filter = "Multimedia Files|*.*"
                    .FileName = Nothing
                    .ShowDialog()
                    FileNamed = .FileName
                    PicName = .SafeFileName
                End With

                IO.File.Copy(FileNamed, ProjectPath & "\Resources\Common\" & PicName)

                ListBox3.Items.Add(PicName)
                Dim Writes As New IO.StreamWriter(ProjectPath & ProjectName & ".mirproj", True)
                Writes.WriteLine("res##1")
                Writes.WriteLine(PicName)
                Writes.Close()
            Catch ex As Exception
                Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
                Writer.WriteLine(ex.Message)
                Writer.Close()
            End Try
        End If

    End Sub


    Private Sub ResourceViewerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            TabControl1.SelectTab(4)
            ListBox5.Items.Clear()

            For index = 0 To ListBox3.Items.Count - 1
                ListBox5.Items.Add(ListBox3.Items(index))
            Next
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try
    End Sub

    Private Sub ListBox5_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        On Error Resume Next
        PictureBox1.Image = Image.FromFile(ProjectPath & "Resources\Common\" & ListBox5.SelectedItem.ToString)
    End Sub


    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PictureBox1.SizeMode = PictureBoxSizeMode.Normal

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next
        o.FileName = ProjectPath & "Resources\Common\" & ListBox5.SelectedItem.ToString
        o.Play()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        o.Stop()
    End Sub

    Private Sub ListBox5_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

    End Sub

    Private Sub ListBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub GeneralAnalysisToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If Not SingleWebPage = True Then
                TabControl1.SelectTab(2)
                Label4.Text = "General Analysis Data for Page : " & CurrentPage
                Dim Obj As New AnalyticsModule
                Obj.Analyze(ProjectPath & CurrentPage, RichTextBox2, ProgressBar1)
                GenAnalysis = True
            Else
                TabControl1.SelectTab(2)
                Label4.Text = "General Analysis Data for Page : " & CurrentPage
                Dim Obj As New AnalyticsModule
                Obj.Analyze(SinglePagePath, RichTextBox2, ProgressBar1)
                GenAnalysis = True
            End If
        Catch ex As Exception
            Dim Writer As New IO.StreamWriter(Environment.CurrentDirectory & "\errorrep.txt")
            Writer.WriteLine(ex.Message)
            Writer.Close()
        End Try

    End Sub

  
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        If Not ListBox5.SelectedItem.ToString = Nothing Then
            Dim o As New CoordFinder
            o.Show()
            o.PictureBox1.Image = Image.FromFile(ProjectPath & "\Resources\Common\" & ListBox5.SelectedItem.ToString)
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ListBox5.Items.Clear()

        For index = 0 To ListBox3.Items.Count - 1
            ListBox5.Items.Add(ListBox3.Items(index))
        Next

    End Sub

    Private Sub LearnHTMLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PasteLocationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteLocationToolStripMenuItem.Click
        Clipboard.SetText(ProjectPath & CurrentPage)
        RichTextBox1.Paste()
    End Sub

    Private Sub ListBox3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Clipboard.SetText(ProjectPath & "Common\Resources\" & ListBox3.SelectedItem.ToString)
        RichTextBox1.Paste()

    End Sub

    Private Sub ListBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If saved = False Then
            Dim Result As MsgBoxResult
            Result = MsgBox("Current Project not Saved. Save now ?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Miracle")

            If Result = MsgBoxResult.Yes Then
                FileSave(ProjectPath & CurrentPage, RichTextBox1)
            End If

            End
        Else
            End
        End If
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)

    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RichTextBox1.Cut()

    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RichTextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RichTextBox1.Paste()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RichTextBox1.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RichTextBox1.Redo()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim k As New AboutUs

        k.Show()

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class
