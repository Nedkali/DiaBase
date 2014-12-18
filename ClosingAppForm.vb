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

    'SET SAVE DBASE WHEN BACKUP DBASE IS SELECTED ON CLOSING APP CONFIRMATION
    Private Sub BackupDatabaseCHRCKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles BackupDatabaseCHRCKBOX.CheckedChanged
        'A new backup on exit system duplicates the saved file and changes extenseion as opposed to saving a whole new file.
        'becase of this the backup the file must also be saved to the autochecks the save box when backuop box is checked. ROB - Rev 29
        If BackupDatabaseCHRCKBOX.Checked = True Then SaveDatabaseCHECKBOX.Checked = True
    End Sub

    'KEEPS SAVE DBASE SET TO TRUE WHILE BACKUP DATABASE IS SELECTED (BACKUP NOW JUST COPIES THE SAVED DATAFILE (EASIER)
    Private Sub SaveDatabaseCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles SaveDatabaseCHECKBOX.CheckedChanged
        If SaveDatabaseCHECKBOX.Checked = False And BackupDatabaseCHRCKBOX.Checked = True Then SaveDatabaseCHECKBOX.Checked = True
    End Sub
End Class
