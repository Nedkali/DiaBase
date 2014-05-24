Public Class ClosingAppForm
    Private Sub AppCloseConfirmBUTTON_Click(sender As Object, e As EventArgs) Handles AppCloseConfirmBUTTON.Click
        saveonexit = SaveDatabaseCHECKBOX.Checked
        backuponexit = BackupDatabaseCHRCKBOX.Checked
        DialogResult = Windows.Forms.DialogResult.Yes
    End Sub

    Private Sub AppCloseCancelBUTTONButton1_Click(sender As Object, e As EventArgs) Handles AppCloseCancelBUTTONButton1.Click
        DialogResult = Windows.Forms.DialogResult.No
    End Sub

    Private Sub ClosingAppForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        SaveDatabaseCHECKBOX.Checked = saveonexit
        BackupDatabaseCHRCKBOX.Checked = backuponexit
        My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background) ' Plays d2 thunk sound
    End Sub

End Class
