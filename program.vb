Imports System.IO

Public Class FileCompare

    Dim fileContent = String.Empty
    Dim fileName = String.Empty
    Dim file1ContentList As New List(Of String)
    Dim file2ContentList As New List(Of String)
    Dim line As String = Nothing
    Dim button1Clicked As Boolean = False
    Dim button2Clicked As Boolean = False
    Dim differenceFound As Boolean = False
    Dim fd As OpenFileDialog = New OpenFileDialog()                               'Opens the Brose files window

    Public Sub File1_Click(sender As Object, e As EventArgs) Handles File1.Click  'Handles the click of the file1 button
        RichTextBox1.Clear()
        button1Clicked = True
        fd.Title = "Open File Browser"
        fd.InitialDirectory = "C:\"
        fd.Filter = "Text Files (*.txt)|*.txt|XML files (*.xml)|*.xml|SQL Files (*.sql)|*.sql"
        fd.FilterIndex = 3
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then                                 'This condition executes when the file is selected
            fileName = fd.FileName                                                'fd.FileName holds the entire path of the file
            fileName = fileName.Substring(fileName.LastIndexOf("\") + 1)          'This would filter out only the fileName
            RichTextBox1.Text = fileName
            Dim reader = File.OpenText(fd.FileName)
            While (reader.Peek <> -1)                                             'WHen not a blank line
                line = reader.ReadLine()
                If Not String.IsNullOrEmpty(line) Then
                    If EnableTrimming.Checked = True Then                         'This condition executes When the trimming option is selected
                        file1ContentList.Add(line.Trim)
                    Else
                        file1ContentList.Add(line)
                    End If
                End If
            End While
        End If
    End Sub

    Public Sub File2_Click(sender As Object, e As EventArgs) Handles File2.Click  'This function performs the sames tasks as of the file1 button
        RichTextBox2.Clear()
        button2Clicked = True
        Dim fileContent = String.Empty
        Dim fileName = String.Empty
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = "Open File Browser"
        fd.Filter = "Text Files (*.txt)|*.txt|XML files (*.xml)|*.xml|SQL Files (*.sql)|*.sql"
        fd.FilterIndex = 3
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then
            fileName = fd.FileName
            fileName = fileName.Substring(fileName.LastIndexOf("\") + 1)
            RichTextBox2.Text = fileName
            Dim reader = File.OpenText(fd.FileName)
            While (reader.Peek <> -1)
                line = reader.ReadLine
                If Not String.IsNullOrEmpty(line) Then
                    If EnableTrimming.Checked = True Then
                        file2ContentList.Add(line.Trim)
                    Else
                        file2ContentList.Add(line)
                    End If
                End If
            End While
        End If
    End Sub

    Public Sub Compare_Click(sender As Object, e As EventArgs) Handles Compare.Click                'This button is clicked once both the files have been uploaded
        RichTextBox3.Clear()
        If EnableTrimming.Checked = True Then
            For i As Integer = 0 To file1ContentList.Count - 1
                file1ContentList.Item(i) = file1ContentList.Item(i).Trim
            Next
            For i As Integer = 0 To file2ContentList.Count - 1
                file2ContentList.Item(i) = file2ContentList.Item(i).Trim
            Next

        End If
        If (button1Clicked = True And button2Clicked = True) Then
            StartComparison()                                                                       'This function has performs the comparison task
        End If
        button1Clicked = False
        button2Clicked = False
    End Sub

    Public Sub StartComparison()
        Dim smallestList As Integer
        Dim largestList As Integer
        Dim matchedIndexes As New List(Of Integer)
        Dim n As Integer = 0
        If (file1ContentList.Count < file2ContentList.Count) Then
            smallestList = file1ContentList.Count
            largestList = file2ContentList.Count
        Else
            smallestList = file2ContentList.Count
            largestList = file1ContentList.Count
        End If
        For i As Integer = 0 To smallestList - 1
            Dim indexes1 As New List(Of Integer)
            Dim indexes2 As New List(Of Integer)
            Dim string1 As String = file1ContentList.Item(i)
            Dim string2 As String = file2ContentList.Item(i)

            Dim charMismatch As Boolean = False
            If (file1ContentList.Item(i) <> file2ContentList.Item(i)) Then
                If (n Mod 2 = 0) Then
                    RichTextBox3.SelectionColor = Color.Blue
                    RichTextBox4.SelectionColor = Color.Blue
                Else
                    RichTextBox3.SelectionColor = Color.Brown
                    RichTextBox4.SelectionColor = Color.Brown
                End If

                If (string1.Length < string2.Length) Then
                    For j As Integer = 0 To string1.Length - 1
                        If (charMismatch = False) Then
                            If Not matchedIndexes.Contains(j) Then
                                If (string1(j) <> string2(j)) Then
                                    If Not indexes2.Contains(j) Then
                                        indexes2.Add(j)
                                    End If
                                    charMismatch = True
                                End If
                            ElseIf (matchedIndexes.Contains(j)) Then
                                charMismatch = True
                            End If
                        End If
                        If (charMismatch = True) Then
                            For k As Integer = j + 1 To string2.Length - 1
                                If (charMismatch = True) Then
                                    If Not matchedIndexes.Contains(k) Then
                                        If (string1(j) = string2(k)) Then
                                            If indexes2.Contains(k) Then
                                                indexes2.Remove(indexes2.IndexOf(k))
                                            End If
                                            If Not matchedIndexes.Contains(k) Then
                                                matchedIndexes.Add(k)
                                            End If
                                            charMismatch = False
                                        ElseIf string1(j) <> string2(k) Then
                                            If Not indexes2.Contains(k) Then
                                                indexes2.Add(k)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        charMismatch = False
                    Next
                    If matchedIndexes.Count > 0 Then
                        For m As Integer = 0 To matchedIndexes.Count - 1
                            If indexes2.Contains(matchedIndexes(m)) Then
                                indexes2.RemoveAt(indexes2.IndexOf(matchedIndexes(m)))
                            End If
                        Next
                        matchedIndexes.Clear()
                    End If
                ElseIf (string2.Length < string1.Length) Then
                    For j As Integer = 0 To string2.Length - 1
                        If (charMismatch = False) Then
                            If Not matchedIndexes.Contains(j) Then
                                If (string1(j) <> string2(j)) Then
                                    If Not indexes1.Contains(j) Then
                                        indexes1.Add(j)
                                    End If
                                    charMismatch = True
                                End If
                            ElseIf matchedIndexes.Contains(j) Then
                                charMismatch = True
                            End If
                        End If
                        If (charMismatch = True) Then
                            For k As Integer = j + 1 To string1.Length - 1
                                If (charMismatch = True) Then
                                    If Not matchedIndexes.Contains(k) Then
                                        If (string1(k) = string2(j)) Then
                                            If (indexes1.Contains(k)) Then
                                                indexes1.RemoveAt(indexes1.IndexOf(k))
                                            End If
                                            If Not matchedIndexes.Contains(k) Then
                                                matchedIndexes.Add(k)
                                            End If
                                            charMismatch = False
                                        ElseIf (string1(k) <> string2(j)) Then
                                            If Not indexes1.Contains(k) Then
                                                indexes1.Add(k)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        charMismatch = False
                    Next
                    If matchedIndexes.Count > 0 Then
                        For m As Integer = 0 To matchedIndexes.Count - 1
                            If indexes1.Contains(matchedIndexes(m)) Then
                                indexes1.RemoveAt(indexes1.IndexOf(matchedIndexes(m)))
                            End If
                        Next
                        matchedIndexes.Clear()
                    End If
                End If
                charMismatch = False
                For j As Integer = 0 To string1.Length - 1
                    If (indexes1.Contains(j)) Then
                        RichTextBox3.SelectionBackColor = Color.LightCoral
                        RichTextBox3.AppendText(string1(j))
                    Else
                        RichTextBox3.SelectionBackColor = Color.Transparent
                        RichTextBox3.AppendText(string1(j))
                    End If
                Next
                RichTextBox3.AppendText(Environment.NewLine)
                RichTextBox3.AppendText(Environment.NewLine)
                For j As Integer = 0 To string2.Length - 1
                    If (indexes2.Contains(j)) Then
                        RichTextBox4.SelectionBackColor = Color.LightCoral
                        RichTextBox4.AppendText(string2(j))
                    Else
                        RichTextBox4.SelectionBackColor = Color.Transparent
                        RichTextBox4.AppendText(string2(j))
                    End If
                Next
                RichTextBox4.AppendText(Environment.NewLine)
                RichTextBox4.AppendText(Environment.NewLine)
                differenceFound = True
                n = n + 1
                indexes1.Clear()
                indexes2.Clear()
            End If
        Next

        If (differenceFound = False) Then
            RichTextBox3.Text = "No Difference Found"
            RichTextBox4.Text = "No Difference Found"
        End If

    End Sub

    Public Sub Reset_Click(sender As Object, e As EventArgs) Handles Reset.Click
        button1Clicked = False
        button2Clicked = False
        RichTextBox1.Clear()
        RichTextBox2.Clear()
        RichTextBox3.Clear()
        RichTextBox4.Clear()
        file1ContentList.Clear()
        file2ContentList.Clear()
        EnableTrimming.Checked = False
        fd.Reset()
    End Sub

    ' Public Sub SaveFile_Click(sender As Object, e As EventArgs)      'This function to be used to save the comparison output in a file (Changes to be done)
    '     Dim sfd As SaveFileDialog = New SaveFileDialog()
    '     sfd.InitialDirectory = "C:\"
    '     sfd.Title = "Save file"
    '     sfd.DefaultExt = "txt"
    '     sfd.Filter = "TXT Files (*.txt)|*.txt|SQL Files (*.sql)|*.sql|XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
    '     sfd.FilterIndex = 1
    '     sfd.RestoreDirectory = True
    '     If (sfd.ShowDialog() = DialogResult.OK) Then
    '         Dim objWriter As StreamWriter = New StreamWriter(sfd.FileName)
    '         objWriter.Write(RichTextBox3.Text)
    '         objWriter.Close()
    '     End If
    ' End Sub
End Class
