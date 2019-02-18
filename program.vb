Imports System.IO

Public Class FileCompare
    Dim file1ContentList As New List(Of String)
    Dim file2ContentList As New List(Of String)
    Dim line As String = Nothing
    Dim button1Clicked As Boolean = False
    Dim button2Clicked As Boolean = False
    Dim differenceFound As Boolean = False

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles File1.Click
        RichTextBox1.Clear()
        button1Clicked = True
        Dim fileContent = String.Empty
        Dim fileName = String.Empty
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = "Open File Browser"
        fd.InitialDirectory = "C:\"
        fd.Filter = "Text Files (*.txt)|*.txt|XML files (*.xml)|*.xml"
        fd.FilterIndex = 1
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
                        file1ContentList.Add(line)
                    Else
                        file1ContentList.Add(line.Trim)
                    End If
                End If
            End While
        End If
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles File2.Click
        RichTextBox2.Clear()
        button2Clicked = True
        Dim fileContent = String.Empty
        Dim fileName = String.Empty
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = "Open File Browser"
        ' fd.InitialDirectory = "C:\"
        fd.Filter = "Text Files (*.txt)|*.txt|XML files (*.xml)|*.xml"
        fd.FilterIndex = 1
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
                        file1ContentList.Add(line)
                    Else
                        file1ContentList.Add(line.Trim)
                    End If
                End If
            End While
        End If
    End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Compare.Click
        If (button1Clicked = True And button2Clicked = True) Then
            startComparison()
        End If
        button1Clicked = False
        button2Clicked = False
    End Sub

    Public Sub startComparison()
        Dim smallestList As Integer
        Dim largestList As Integer
        Dim maxLength As Integer
        Dim noOfSpaces As Integer
        Dim file1ListWithAppendedSpaces As String
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

        maxLength = findMaximumLength()
        For i As Integer = 0 To smallestList - 1
            If (file1ContentList.Item(i) <> file2ContentList.Item(i)) Then
                noOfSpaces = maxLength - file1ContentList.Item(i).Length
                file1ListWithAppendedSpaces = file1ContentList.Item(i)
                For j As Integer = 1 To noOfSpaces
                    file1ListWithAppendedSpaces = file1ListWithAppendedSpaces + " "
                Next
                RichTextBox3.AppendText(file1ContentList.Item(i) + " |---| " + file2ContentList.Item(i) & Environment.NewLine)
                differenceFound = True
                file1ListWithAppendedSpaces = ""
            End If
        Next
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
        If (differenceFound = False) Then
            RichTextBox3.Text = "No Difference Found"
        End If
    End Sub

    Function findMaximumLength() As Integer
        Dim maxLength As Integer = 0
        For i As Integer = 0 To file1ContentList.Count - 1
            If (file1ContentList.Item(i).Length > maxLength) Then
                maxLength = file1ContentList.Item(i).Length
            End If
        Next
        Return maxLength
    End Function

    Public Sub Button4_Click(sender As Object, e As EventArgs) Handles Reset.Click
        button1Clicked = False
        button2Clicked = False
        RichTextBox1.Clear()
        RichTextBox2.Clear()
        RichTextBox3.Clear()
    End Sub


End Class
