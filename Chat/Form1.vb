Imports Google.API.Search

Public Class Form1

    Dim bookmarks As New List(Of String)
    Dim descriptions As New List(Of String)
    Dim commands As New List(Of String)

    Dim todos As New List(Of String)
    Dim markings As New List(Of Boolean)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ProcessCommand(TextBox1.Text)
        TextBox1.Clear()
    End Sub

    Private Sub ProcessCommand(ByVal command As String)
        commands.Add(command)
        Dim num As Integer = command.Length
        If num < 4 Then
            RichTextBox1.AppendText("Invalid command.  Command must start with igor." + vbNewLine)
            RichTextBox1.ScrollToCaret()
            Return
        Else
            If command.Length = 4 + 1 + 4 Then
                If command = "igor time" Then
                    RichTextBox1.AppendText(TimeOfDay.ToString("h:mm:ss tt") + vbNewLine)
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If

            If command.Length > 11 Then
                If command.Substring(0, 11) = "igor google" Then
                    Dim client As Google.API.Search.GwebSearchClient = New Google.API.Search.GwebSearchClient("www.jdchrom.com")
                    Dim results As List(Of IWebResult) = client.Search(command.Substring(12), 32)
                    RichTextBox1.AppendText("Title: " + results(0).Title + vbNewLine) 'results(0).Content
                    RichTextBox1.AppendText("Content: " + results(0).Content + vbNewLine)
                    RichTextBox1.AppendText("URL: " + results(0).Url + vbNewLine)
                    Return
                End If
            End If
            If command.Length = 4 + 1 + 7 Then
                If command = "igor history" Then
                    If commands.Count > 10 Then
                        For i As Integer = commands.Count - 11 To commands.Count - 1
                            RichTextBox1.AppendText(commands(i) + vbNewLine)
                        Next
                    Else
                        For i As Integer = 0 To commands.Count - 1
                            RichTextBox1.AppendText(commands(i) + vbNewLine)
                        Next
                    End If
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length > 4 + 1 + 10 - 1 Then
                If command.Substring(0, 4 + 1 + 10 - 1) = "igor calculate" Then
                    RichTextBox1.AppendText(Evaluate(command.Substring(15)).ToString() + vbNewLine)
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length > 9 Then
                If command.Substring(0, 9) = "igor play" Then
                    Dim webAddress As String = command.Substring(10)
                    Process.Start(webAddress)
                    RichTextBox1.AppendText("Playing video." + vbNewLine)
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length > 14 Then
                If command.Substring(0, 14) = "igor todo mark" Then
                    Dim index As Integer = todos.IndexOf(command.Substring(15))
                    If index <> -1 Then
                        markings(index) = True
                        RichTextBox1.AppendText("Todo item marked." + vbNewLine)
                    Else
                        RichTextBox1.AppendText("Todo item not found." + vbNewLine)
                    End If
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length = 14 Then
                If command.Substring(0, 14) = "igor todo list" Then
                    If todos.Count > 0 Then
                        For i As Integer = 0 To todos.Count - 1
                            RichTextBox1.AppendText(markings(i).ToString + vbTab + todos(i) + vbNewLine)
                        Next
                    Else
                        RichTextBox1.AppendText("Todo list is empty." + vbNewLine)
                    End If
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length > 9 Then
                If command.Substring(0, 9) = "igor todo" Then
                    todos.Add(command.Substring(10))
                    markings.Add(False)
                    RichTextBox1.AppendText("Todo item added." + vbNewLine)
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length > 4 + 1 + 7 Then
                If command.Substring(0, 4 + 1 + 7) = "igor history" Then
                    Dim filteredCommands As New List(Of String)
                    Dim filter As String = command.Substring(4 + 1 + 7 + 1)
                    For i As Integer = 0 To commands.Count - 1
                        If commands(i).Length >= filter.Length Then
                            If commands(i).Substring(0, filter.Length) = filter Then
                                filteredCommands.Add(commands(i))
                            End If
                        End If
                    Next

                    RichTextBox1.AppendText("Filtered commands: " + vbNewLine)
                    If filteredCommands.Count > 10 Then
                        For i As Integer = filteredCommands.Count - 11 To filteredCommands.Count - 1
                            RichTextBox1.AppendText(filteredCommands(i) + vbNewLine)
                        Next
                    Else
                        If filteredCommands.Count > 0 Then
                            For i As Integer = 0 To filteredCommands.Count - 1
                                RichTextBox1.AppendText(filteredCommands(i) + vbNewLine)
                            Next
                        Else
                            RichTextBox1.AppendText("No commands found using that filter." + vbNewLine)

                        End If
                    End If
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length >= 24 Then
                If command.Substring(0, 24) = "igor bookmark delete all" Then
                    bookmarks.Clear()
                    descriptions.Clear()
                    RichTextBox1.AppendText("All bookmarks deleted." + vbNewLine)
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length >= 21 Then
                If command.Substring(0, 20) = "igor bookmark delete" Then
                    Dim i As Integer = bookmarks.IndexOf(command.Substring(21))
                    If (i = -1) Then
                        RichTextBox1.AppendText("Could not find specified bookmark." + vbNewLine)
                    Else
                        bookmarks.RemoveAt(i)
                        bookmarks.RemoveAt(i)
                        RichTextBox1.AppendText("Successfully deleted." + vbNewLine)
                    End If
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length >= 18 Then
                If command.Substring(0, 18) = "igor bookmark list" Then
                    If bookmarks.Count > 0 Then
                        RichTextBox1.AppendText("Bookmarks:" + vbNewLine)
                        For i As Integer = 0 To bookmarks.Count - 1
                            RichTextBox1.AppendText(bookmarks(i) + " " + descriptions(i) + vbNewLine)
                        Next
                    Else
                        RichTextBox1.AppendText("Bookmark list is empty." + vbNewLine)
                    End If
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
            If command.Length >= 14 Then
                If command.Substring(0, 13) = "igor bookmark" Then
                    Dim str As String = command.Substring(14)
                    Dim url_and_desc As String() = command.Substring(14).Split(New Char() {" "c})
                    If url_and_desc.Length < 2 Then
                        RichTextBox1.AppendText("Invalid syntax.  Correct syntax is: igor bookmark <url> <desc>" + vbNewLine)
                    Else
                        bookmarks.Add(url_and_desc(0))
                        Dim len As Integer = url_and_desc(0).Length
                        descriptions.Add(command.Substring(14 + url_and_desc(0).Length))
                        RichTextBox1.AppendText("Successfully added bookmark and description." + vbNewLine)
                    End If
                    RichTextBox1.ScrollToCaret()
                    Return
                End If
            End If
        End If
    End Sub
End Class
