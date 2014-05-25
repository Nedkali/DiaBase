Public Class UserMessaging

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub UserMessaging_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Left = Form1.Left + 200
        Me.Top = Form1.Top + 300

        TextBox1.Text = Mymessages
        Mymessages = ""
        TextBox1.TextAlign = HorizontalAlignment.Center

        My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background) 'plays d2 thunk alert sound
    End Sub
End Class