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

    Public Sub File1_Click(sender As Object, e As EventArgs) Handles File1.Click  'Handles the click of the file1 button
        RichTextBox1.Clear()
        button1Clicked = True

        Dim fd As OpenFileDialog = New OpenFileDialog()                             'Opens the Brose files window
        fd.Title = "Open File Browser"
        fd.InitialDirectory = "C:\"
        fd.Filter = "Text Files (*.txt)|*.txt|XML files (*.xml)|*.xml|SQL Files (*.sql)|*.sql"
        fd.FilterIndex = 3
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then
            fileName = fd.FileName
            fileName = fileName.Substring(fileName.LastIndexOf("\") + 1)
            RichTextBox1.Text = fileName
            Dim reader = File.OpenText(fd.FileName)
            While (reader.Peek <> -1)
                line = reader.ReadLine()
                If Not String.IsNullOrEmpty(line) Then
                    If EnableTrimming.Checked = True Then
                        file1ContentList.Add(line.Trim)
                    Else
                        file1ContentList.Add(line)
                    End If
                End If
            End While
        End If
    End Sub

    Public Sub File2_Click(sender As Object, e As EventArgs) Handles File2.Click
        RichTextBox2.Clear()
        button2Clicked = True
        Dim fileContent = String.Empty
        Dim fileName = String.Empty
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = "Open File Browser"
        ' fd.InitialDirectory = "C:\"
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

    Public Sub Compare_Click(sender As Object, e As EventArgs) Handles Compare.Click
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
            StartComparison()
        End If
        button1Clicked = False
        button2Clicked = False
    End Sub

    Public Sub StartComparison()
        Dim smallestList As Integer
        Dim largestList As Integer
        If (file1ContentList.Count < file2ContentList.Count) Then
            smallestList = file1ContentList.Count
        Else
            smallestList = file2ContentList.Count
        End If

        If (file1ContentList.Count > file2ContentList.Count) Then
            largestList = file1ContentList.Count
        Else
            largestList = file2ContentList.Count
        End If
        For i As Integer = 0 To smallestList - 1
            If (file1ContentList.Item(i) <> file2ContentList.Item(i)) Then
                RichTextBox3.AppendText(file1ContentList.Item(i) + " |---| " + file2ContentList.Item(i) & Environment.NewLine)
                differenceFound = True
            End If
        Next
        If (file1ContentList.Count <> file2ContentList.Count) Then
            If (file1ContentList.Count = largestList) Then
                differenceFound = True
                For i As Integer = smallestList To largestList - 1
                    RichTextBox3.AppendText(file1ContentList.Item(i) + " |---|  |------------- No File2 Content for/from this line----------|" & Environment.NewLine)

                Next
            ElseIf (file2ContentList.Count = largestList) Then
                differenceFound = True
                For i As Integer = smallestList To largestList - 1
                    RichTextBox3.AppendText(" |------No File1 Content for/from this line-------|  |---|" + file2ContentList.Item(i) & Environment.NewLine)
                Next
            End If
        End If
        If (differenceFound = False) Then
            RichTextBox3.Text = "No Difference Found"
        End If

    End Sub


    Public Sub Reset_Click(sender As Object, e As EventArgs) Handles Reset.Click
        button1Clicked = False
        button2Clicked = False
        RichTextBox1.Clear()
        RichTextBox2.Clear()
        RichTextBox3.Clear()
    End Sub

    Public Sub SaveFile_Click(sender As Object, e As EventArgs) Handles SaveFile.Click

        Dim sfd As SaveFileDialog = New SaveFileDialog()
        sfd.InitialDirectory = "C:\"
        sfd.Title = "Save file"
        sfd.DefaultExt = "txt"
sfd.Filter = "Text Files (*.txt)|*.txt|SQL Files (*.sql)|*.sql|XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
        sfd.FilterIndex = 1
        sfd.RestoreDirectory = True
        If (sfd.ShowDialog() = DialogResult.OK) Then
                Dim objWriter As StreamWriter = New StreamWriter(sfd.FileName)
                objWriter.Write(RichTextBox3.Text)
                objWriter.Close()
            End If
    End Sub

End Class
