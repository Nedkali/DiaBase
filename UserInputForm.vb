Public Class UserInputForm
    'clears User Input Textbox
    Private Sub UserInputForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UserInputTEXTBOX.Text = Nothing
    End Sub

    'YES BUTTON PRESS
    Private Sub UserInputYesBUTTON_Click(sender As Object, e As EventArgs) Handles UserInputYesBUTTON.Click
        If UserInputTEXTBOX.Text <> "" Then DialogResult = Windows.Forms.DialogResult.Yes : Me.Close() Else UserInputTEXTBOX.Select()
    End Sub

    'NO BUTTON PRESS
    Private Sub UserInputNoBUTTON_Click(sender As Object, e As EventArgs) Handles UserInputNoBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.No : Me.Close()
    End Sub
End Class