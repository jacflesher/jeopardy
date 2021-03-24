Public Class Form1

#Region "Declaration"
    Dim folderpath As String = ""
    Public question As String = ""
#End Region

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        FolderBrowserDialog1.ShowDialog()
        folderpath = FolderBrowserDialog1.SelectedPath.ToString
    End Sub

    Private Function QuestionPrompt(column As String, points As String, buttonNumber As Integer)

        question = My.Computer.FileSystem.ReadAllText(folderpath + "\" & column & "-" & points & ".txt")
        Dim questionForm As Question = New Question
        questionForm.ShowDialog()
        If questionForm.submit_clicked = True Then
            TallyScore(questionForm.team, points, questionForm.ans, buttonNumber)
        End If
        Return Nothing
    End Function
    Private Function TallyScore(team As String, points As String, ans As String, buttonNumber As Integer)
        If team.Contains("1") AndAlso ans = "Correct" Then
            Label1.Text = CInt(Label1.Text) + points
        ElseIf team.Contains("2") AndAlso ans = "Correct" Then
            Label2.Text = CInt(Label2.Text) + points
        ElseIf team.Contains("3") AndAlso ans = "Correct" Then
            Label3.Text = CInt(Label3.Text) + points
        ElseIf team.Contains("1") AndAlso ans = "Incorrect" Then
            Label1.Text = CInt(Label1.Text) - points
        ElseIf team.Contains("2") AndAlso ans = "Incorrect" Then
            Label2.Text = CInt(Label2.Text) - points
        ElseIf team.Contains("3") AndAlso ans = "Incorrect" Then
            Label3.Text = CInt(Label3.Text) - points
        End If
        deactivateQuestionButton(buttonNumber, ans)
        Return Nothing
    End Function

    Private Function deactivateQuestionButton(buttonNumber As Integer, ans As String)

        If ans = "Incorrect" Then
            Dim response As MsgBoxResult = MsgBox("Incorrect answers do not trigger the question button to deactivate. Do you wish to overrule and deactivate the question button now?", MsgBoxStyle.YesNo, "Deactivate Question Button")
            If response = MsgBoxResult.Yes Then
                disableButton(buttonNumber)
            Else
                Return Nothing
            End If
        Else
            disableButton(buttonNumber)
        End If

        Return Nothing
    End Function

    Private Function disableButton(buttonNumber As String)
        If buttonNumber = 1 Then
            Button1.Enabled = False
        ElseIf buttonNumber = 2 Then
            Button2.Enabled = False
        ElseIf buttonNumber = 3 Then
            Button3.Enabled = False
        ElseIf buttonNumber = 4 Then
            Button4.Enabled = False
        ElseIf buttonNumber = 5 Then
            Button5.Enabled = False
        ElseIf buttonNumber = 6 Then
            Button6.Enabled = False
        ElseIf buttonNumber = 7 Then
            Button7.Enabled = False
        ElseIf buttonNumber = 8 Then
            Button8.Enabled = False
        ElseIf buttonNumber = 9 Then
            Button9.Enabled = False
        ElseIf buttonNumber = 10 Then
            Button10.Enabled = False
        ElseIf buttonNumber = 11 Then
            Button11.Enabled = False
        ElseIf buttonNumber = 12 Then
            Button12.Enabled = False
        ElseIf buttonNumber = 13 Then
            Button13.Enabled = False
        ElseIf buttonNumber = 14 Then
            Button14.Enabled = False
        ElseIf buttonNumber = 15 Then
            Button15.Enabled = False
        ElseIf buttonNumber = 16 Then
            Button16.Enabled = False
        ElseIf buttonNumber = 17 Then
            Button17.Enabled = False
        ElseIf buttonNumber = 18 Then
            Button18.Enabled = False
        ElseIf buttonNumber = 19 Then
            Button19.Enabled = False
        ElseIf buttonNumber = 20 Then
            Button20.Enabled = False
        ElseIf buttonNumber = 21 Then
            Button21.Enabled = False
        ElseIf buttonNumber = 22 Then
            Button22.Enabled = False
        ElseIf buttonNumber = 23 Then
            Button23.Enabled = False
        ElseIf buttonNumber = 24 Then
            Button24.Enabled = False
        ElseIf buttonNumber = 25 Then
            Button25.Enabled = False
        End If
        Return Nothing
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If folderpath <> "" Then
            QuestionPrompt("C1", 200, 1)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If folderpath <> "" Then
            QuestionPrompt("C1", 400, 10)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If folderpath <> "" Then
            QuestionPrompt("C1", 600, 15)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If folderpath <> "" Then
            QuestionPrompt("C1", 800, 20)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        If folderpath <> "" Then
            QuestionPrompt("C1", 1000, 25)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If folderpath <> "" Then
            QuestionPrompt("C2", 200, 2)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If folderpath <> "" Then
            QuestionPrompt("C2", 400, 9)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If folderpath <> "" Then
            QuestionPrompt("C2", 600, 14)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        If folderpath <> "" Then
            QuestionPrompt("C2", 800, 19)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        If folderpath <> "" Then
            QuestionPrompt("C2", 1000, 24)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If folderpath <> "" Then
            QuestionPrompt("C3", 200, 3)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If folderpath <> "" Then
            QuestionPrompt("C3", 400, 8)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If folderpath <> "" Then
            QuestionPrompt("C3", 600, 13)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If folderpath <> "" Then
            QuestionPrompt("C3", 800, 18)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        If folderpath <> "" Then
            QuestionPrompt("C3", 1000, 23)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If folderpath <> "" Then
            QuestionPrompt("C4", 200, 4)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If folderpath <> "" Then
            QuestionPrompt("C4", 400, 7)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If folderpath <> "" Then
            QuestionPrompt("C4", 600, 12)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If folderpath <> "" Then
            QuestionPrompt("C4", 800, 17)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If folderpath <> "" Then
            QuestionPrompt("C4", 1000, 22)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If folderpath <> "" Then
            QuestionPrompt("C5", 200, 5)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If folderpath <> "" Then
            QuestionPrompt("C5", 400, 6)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If folderpath <> "" Then
            QuestionPrompt("C5", 600, 11)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If folderpath <> "" Then
            QuestionPrompt("C5", 800, 16)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If folderpath <> "" Then
            QuestionPrompt("C5", 1000, 21)
        Else
            MsgBox("Please choose a directory that contains your question files")
        End If
    End Sub
End Class
