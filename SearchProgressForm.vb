Public Class SearchProgressForm

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub SearchProgressForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Left = Form1.Left + 160
        Me.Top = Form1.Top + 100
    End Sub
End Class