Public Class Settings

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles CancelDefaultsBUTTON.Click
        If PictureBox1.Visible = False Then
            Mymessages = "Must set correct Etal folder" : MyMessageBox()
            Return
        End If
        If PictureBox2.Visible = False Then
            Mymessages = "Must Set Database file" : MyMessageBox()
            Return
        End If

        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles EtalPathBrowseBUTTON.Click
        FolderBrowserDialog1.ShowDialog()
        EtalPathTEXTBOX.Text = FolderBrowserDialog1.SelectedPath
        SettingsChecker()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles DefaultDatabaseBrowseBUTTON.Click
        'todo - set a variable and set to selected Data base file name
        Dim openFileDialog1 As New OpenFileDialog()
        DatabaseFileTEXTBOX.Text = ""
        PictureBox2.Visible = False
        openFileDialog1.InitialDirectory = Application.StartupPath + "\DataBase"
        openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True
        openFileDialog1.ShowDialog()
        DatabaseFileTEXTBOX.Text = openFileDialog1.FileName : DatabaseFileTEXTBOX.Refresh()
        DataBasePath = openFileDialog1.FileName
        DefaultDatabaseFile = DataBasePath
        SettingsChecker()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles SaveDefaultsBUTTON.Click

        If PictureBox1.Visible = False Then
            Mymessages = "Must set correct Etal folder" : MyMessageBox()
            Return
        End If
        If PictureBox2.Visible = False Then
            Mymessages = "Must Set Database file" : MyMessageBox()
            Return
        End If
        SaveConfigFile() 'branch to save config sub in module1

           Me.Close()

    End Sub

    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = Form1.Left + 150
        Me.Top = Form1.Top + 100
        EtalPathTEXTBOX.Text = EtalPath
        DatabaseFileTEXTBOX.Text = DefaultDatabaseFile
        NumericUpDown1.Value = TimerMins
        CheckBox3.Checked = KeepPassPrivate
        AutoBackupImportsCHECKBOX.Checked = AutoBackups
        BackupOnEditsCHECKBOX.Checked = EditBackups
        DupeCheckBox1.Checked = RemoveDupeMule
        DisableSearchProgressBarCHECKBOX.Checked = HideSearchPopup
        SettingsChecker()


    End Sub

    Private Sub SettingsChecker()
        Me.PictureBox1.Visible = False
        PictureBox2.Visible = False
        If (My.Computer.FileSystem.DirectoryExists(String.Concat(Me.EtalPathTEXTBOX.Text, "\scripts\AMS\MuleInventory"))) Then Me.PictureBox1.Visible = True
        If My.Computer.FileSystem.FileExists(DatabaseFileTEXTBOX.Text) = True Then PictureBox2.Visible = True
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles DupeCheckBox1.CheckedChanged
        If DupeCheckBox1.Checked = True Then
            RemoveDupeMule = True
        End If
    End Sub
End Class