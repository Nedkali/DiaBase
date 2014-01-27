Public Class ClosingAppForm

    Private Sub AppCloseConfirmBUTTON_Click(sender As Object, e As EventArgs) Handles AppCloseConfirmBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.Yes

    End Sub

    Private Sub AppCloseCancelBUTTONButton1_Click(sender As Object, e As EventArgs) Handles AppCloseCancelBUTTONButton1.Click
        DialogResult = Windows.Forms.DialogResult.No
    End Sub

   
    Private Sub ClosingAppForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background) ' Plays d2 thunk sound
        'My.Computer.Audio.Play(My.Resources.Baalfadeout, AudioPlayMode.Background) 'plays baal laugh sound
    End Sub
End Class