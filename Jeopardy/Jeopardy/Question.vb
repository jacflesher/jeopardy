Public Class Question
    Public submit_clicked As Boolean = False
    Public team As String = ""
    Public ans As String = ""


    Private Sub Question_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = Form1.question

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sb As Text.StringBuilder = New Text.StringBuilder
        If ComboBox1.Text = "" Then
            sb.AppendLine(">No team has been selected")
        End If
        If ComboBox2.Text = "" Then
            sb.AppendLine(">No answer result have been entered")
        End If
        If sb.ToString <> "" Then
            MsgBox("Errors: " & sb.ToString * vbNewLine & "No points have been tallied due to errors. Please try again.")
        Else
            team = ComboBox1.Text
            ans = ComboBox2.Text
            submit_clicked = True
            Me.Close()
        End If

    End Sub


End Class