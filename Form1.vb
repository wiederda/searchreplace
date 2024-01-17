Imports System.Reflection

<Assembly: AssemblyVersion("1.0.0.1")>
<Assembly: AssemblyFileVersion("1.0.0.1")>
Public Class Form1
    Dim sFile As New s_File
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Wort, das du finden und markieren möchtest
        Dim wordToFind As String = TextBox1.Text

        ' Suche nach allen Vorkommen des Worts und markiere sie in der RichTextBox
        If TextBox1.Text <> "" Then
            MarkiereAlleWorteInRichTextBox(RichTextBox1, wordToFind)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Wort, das du finden und ersetzen möchtest
        Dim wordToFind As String = TextBox1.Text

        ' Wort, durch das du ersetzen möchtest
        Dim replacementWord As String = TextBox2.Text

        ' Ersetze alle Vorkommen des Worts in der RichTextBox
        If TextBox1.Text <> "" And TextBox2.Text <> "" Then
            ErsetzeAlleWorteInRichTextBox(RichTextBox1, wordToFind, replacementWord)
        End If
    End Sub

    Private Sub MarkiereAlleWorteInRichTextBox(richTextBox As RichTextBox, word As String)
        ' Aktuelle Position des Cursors speichern
        Dim currentPosition As Integer = richTextBox.SelectionStart

        ' Startposition des gesuchten Wortes
        Dim startIndex As Integer = richTextBox.Text.IndexOf(word)

        ' Durchlaufe die RichTextBox, bis alle Vorkommen des Worts markiert wurden
        While startIndex <> -1
            ' Markiere das gefundene Wort
            richTextBox.Select(startIndex, word.Length)
            richTextBox.SelectionBackColor = Color.Yellow

            ' Suche nach dem nächsten Vorkommen des Worts
            startIndex = richTextBox.Text.IndexOf(word, startIndex + 1)
        End While

        ' Setze den Cursor wieder auf die ursprüngliche Position zurück
        richTextBox.Select(currentPosition, 0)
        ' Setze die Hintergrundfarbe zurück
        richTextBox.SelectionBackColor = richTextBox.BackColor
    End Sub

    Private Sub ErsetzeAlleWorteInRichTextBox(richTextBox As RichTextBox, wordToFind As String, replacementWord As String)
        ' Speichere die aktuelle Position des Cursors
        Dim currentPosition As Integer = richTextBox.SelectionStart

        ' Durchlaufe die RichTextBox, bis alle Vorkommen des Worts ersetzt wurden
        Dim startIndex As Integer = 0
        Do
            ' Startposition des gesuchten Wortes
            startIndex = richTextBox.Text.IndexOf(wordToFind, startIndex)

            ' Wenn das Wort gefunden wurde, ersetze es
            If startIndex <> -1 Then
                richTextBox.Select(startIndex, wordToFind.Length)
                richTextBox.SelectedText = replacementWord

                ' Setze den Startindex auf das Ende des ersetzen Worts
                startIndex += replacementWord.Length
            End If
        Loop While startIndex <> -1

        ' Setze den Cursor wieder auf die ursprüngliche Position zurück
        richTextBox.Select(currentPosition, 0)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Opacity = 0
        Dim args() As String = Environment.GetCommandLineArgs

        If args.Length > 1 Then
            If args(1) <> "" And args(2) <> "" And args(3) <> "" Then
                RichTextBox1.Text = ""
                RichTextBox1.Text = IO.File.ReadAllText(args(1))
                ErsetzeAlleWorteInRichTextBox(RichTextBox1, args(2), args(3))
                sFile.WriteFile(Application.StartupPath & "\Replace.log", Format(Now, "dd.MM.yyyy HH:mm" & args(1) & " " & args(2) & " " & args(3)), True)
                IO.File.WriteAllText(args(1), RichTextBox1.Text)
                Close()
            End If
        Else
            Text = Text & " " & Application.ProductVersion
            Opacity = 100
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        RichTextBox1.Text = IO.File.ReadAllText(sFile.OpenFileDialog("C:\", "*.* (All Files)|*.*"))
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        IO.File.WriteAllText(sFile.SaveFileDialog("C:\", "*.txt (Textfiles|*.txt"), RichTextBox1.Text)
    End Sub
End Class
