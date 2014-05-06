Public Class ItemImageSelector

    Private Sub ItemImageSelector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PictureBox1.Load(Application.StartupPath & "\Skins\" + ItemImageList(NumericUpDown1.Value) + ".jpg")
        NumericUpDown1.Value = 1
    End Sub

    Private Sub AddSelectImageCancelBUTTON_Click(sender As Object, e As EventArgs) Handles AddSelectImageCancelBUTTON.Click
        Me.Close()
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        PictureBox1.Load(Application.StartupPath & "\Skins\" + ItemImageList(NumericUpDown1.Value) + ".jpg")
    End Sub

    Private Sub AddSelectImageBUTTON_Click(sender As Object, e As EventArgs) Handles AddSelectImageBUTTON.Click

        Dim temp As String = Convert.ToString(NumericUpDown1.Value)
        EditItemForm.EditItemImageTEXTBOX.Text = temp
        Me.Close()
    End Sub
End Class