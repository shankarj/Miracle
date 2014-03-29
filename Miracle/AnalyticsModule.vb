Public Class AnalyticsModule

    Public Function Analyze(ByVal ContentLocation As String, ByVal OutputArea As RichTextBox, ByVal ProgressIndicator As ProgressBar) As Exception
        Try
            Dim LineToAnalayze As String
            Dim LineCount As Integer = 0
            Dim TempText As New RichTextBox

            PerformanceCounts = 0

            ProgressIndicator.Value = 0

            TempText.LoadFile(ContentLocation, RichTextBoxStreamType.PlainText)
            ProgressIndicator.Maximum = TempText.Lines.Count

            OutputArea.Text = Nothing
            OutputArea.Text &= vbCrLf & vbCrLf
       
        Try
            Dim ContentReader As New IO.StreamReader(ContentLocation)

            While Not ContentReader.EndOfStream

                LineToAnalayze = ContentReader.ReadLine
                LineCount += 1

                If LineToAnalayze.Contains("<html>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a html document on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<body>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a body section on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<head>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a head section on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<title>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a title of document on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<img") = True Then
                    OutputArea.Text &= vbCrLf & "Added image on line no : " & LineCount
                    PerformanceCounts += 1
                ElseIf LineToAnalayze.Contains("<a href") = True And LineToAnalayze.Contains("#") And LineToAnalayze.Contains(".html") Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Added a link to a section in another page on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<a href") = True And LineToAnalayze.Contains("#") Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Added a link to a section in same page on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<a href") = True And LineToAnalayze.Contains(".html") Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Added a link to another page on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<a href") = True Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Added a link to a website on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<br>") = True Then
                    OutputArea.Text &= vbCrLf & "Given a line break on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<code") = True Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Source code listing on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<div") = True Then
                    OutputArea.Text &= vbCrLf & "Structure or block of text is formatted on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<hr>") = True Then
                    OutputArea.Text &= vbCrLf & "Given a horizontal rule on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<p>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a paragragh on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<pre>") = True Then
                    OutputArea.Text &= vbCrLf & "Prefomatted the text on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<span") = True Then
                    OutputArea.Text &= vbCrLf & "Inline formatting on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<strong>") Or LineToAnalayze.Contains("<b>") Then
                    OutputArea.Text &= vbCrLf & "Bold text on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<sub>") = True Then
                    OutputArea.Text &= vbCrLf & "Subscript text on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<sup>") = True Then
                    OutputArea.Text &= vbCrLf & "Superscript text on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<em>") Or LineToAnalayze.Contains("<i>") Then
                    OutputArea.Text &= vbCrLf & "Italics text on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<h1>") Or LineToAnalayze.Contains("<h2>") Or LineToAnalayze.Contains("<h3>") Or LineToAnalayze.Contains("<h4>") Or LineToAnalayze.Contains("<h5>") Or LineToAnalayze.Contains("<h6>") Then
                    OutputArea.Text &= vbCrLf & "Created text with preformated headings on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<frame") = True Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Created a frame on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<frameset") = True Then
                    OutputArea.Text &= vbCrLf & "Created a frame set on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<iframe") = True Then
                    OutputArea.Text &= vbCrLf & "Created a inline frame on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<ul>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a unordered list on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<ol>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a ordered list on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<li>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a list on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<dd>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a definition on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<dl>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a definition list on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<dt>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a definition term on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<table>") = True Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Created a table on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<tr>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a table row on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<th>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a table heading on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<td>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a table column on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<thead>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a table heading  on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<tfoot>") = True Then
                    OutputArea.Text &= vbCrLf & "Created a table footer on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<caption>") = True Then
                    OutputArea.Text &= vbCrLf & "Specified a table caption on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<tbody>") = True Then
                    OutputArea.Text &= vbCrLf & "Specified a table body on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<td colspan") = True Then
                    OutputArea.Text &= vbCrLf & "Specified table column span on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<td rowspan") = True Then
                    OutputArea.Text &= vbCrLf & "Specified table column span on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<th colspan") = True Then
                    OutputArea.Text &= vbCrLf & "Specified columns table header cell span on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<form") = True Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Created a form on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<eldset") = True Then
                    OutputArea.Text &= vbCrLf & "Created a group of related items on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<option") = True Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Created a Menu item in a select box on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<select") = True Then
                    OutputArea.Text &= vbCrLf & "Created a Drop-down menu on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<textarea") = True Then
                    PerformanceCounts += 1
                    OutputArea.Text &= vbCrLf & "Created a Multi-row text area on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<input type") = True Then
                    PerformanceCounts += 5
                    OutputArea.Text &= vbCrLf & "Created a Form element on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<tt") = True Then
                    OutputArea.Text &= vbCrLf & "Created a inline monospacettext on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<cite") = True Then
                    OutputArea.Text &= vbCrLf & "Created a citation in italic text on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<font") = True Then
                    OutputArea.Text &= vbCrLf & "Specified text emphasized on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<blockquote") = True Then
                    OutputArea.Text &= vbCrLf & "Created a indent text at both the margins on line no : " & LineCount
                ElseIf LineToAnalayze.Contains("<!--") = True Then
                    OutputArea.Text &= vbCrLf & "Created a comment section on line no : " & LineCount

                End If
                ProgressIndicator.Value += 1

            End While

            ProgressIndicator.Value = TempText.Lines.Count

            Catch ex As Exception
                Return ex
            End Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Miracle")
        End Try

    End Function

    
End Class




