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
        OpenFileDialog1.InitialDirectory = Application.StartupPath + "\DataBase"
        OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 2
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.ShowDialog()
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

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath + "\Settings.cfg", False)
        file.WriteLine(EtalPathTEXTBOX.Text)
        file.WriteLine(DatabaseFileTEXTBOX.Text)
        file.WriteLine(NumericUpDown1.Value)
        file.WriteLine(CheckBox3.Checked) : KeepPassPrivate = CheckBox3.CheckState '<---------- Display Password Fix rev 12 (AussieHack)
        file.WriteLine(AutoBackupImportsCHECKBOX.Checked) '                                     added this for auto backup on import setting
        file.WriteLine(BackupOnEditsCHECKBOX.Checked)   '                                       added this for auto backup on edits setting
        file.WriteLine(DupeCheckBox1.Checked) : RemoveDupeMule = DupeCheckBox1.CheckState
        file.Close()
        EtalPath = EtalPathTEXTBOX.Text
        TimerMins = NumericUpDown1.Value
        TimerSecs = TimerMins * 60
        Me.Close()

        '------------------------------------------------------------------------------------------------------------------------------------------------
        'Apply Search Progress checkbox Setting To Global Config bool "ShowSearchProgress" 
        'THIS SETTING CANT BE SAVED YET (not included in config file yet) AND BY DEFAULT IS SET TO SHOW THE PROGRESS BAR FORM WHEN SEARCHING
        If DisableSearchProgressBarCHECKBOX.Checked = True Then ShowSearchProgress = False
        If DisableSearchProgressBarCHECKBOX.Checked = False Then ShowSearchProgress = True
        '------------------------------------------------------------------------------------------------------------------------------------------------

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
        SettingsChecker()
        'APPLY THE TRIAL "DONT DISPLAY SEARCH PROGRESS BAR" CHECKBOX FROM ITS GLOBAL CONFIG VAR (NOT SAVED IN ACTUAL CONFIG FILE YET) - REV 28
        If ShowSearchProgress = True Then DisableSearchProgressBarCHECKBOX.Checked = False
        If ShowSearchProgress = False Then DisableSearchProgressBarCHECKBOX.Checked = True

    End Sub

    Private Sub SettingsChecker()
        If My.Computer.FileSystem.DirectoryExists(EtalPathTEXTBOX.Text & "\scripts\AMS\MuleInventory") = True Then
            PictureBox1.Visible = True
        End If
        If My.Computer.FileSystem.DirectoryExists(EtalPathTEXTBOX.Text & "\scripts\AMS\MuleInventory") = False Then
            PictureBox1.Visible = False
        End If

        If My.Computer.FileSystem.FileExists(DatabaseFileTEXTBOX.Text) = True Then PictureBox2.Visible = True
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles DupeCheckBox1.CheckedChanged
        If DupeCheckBox1.Checked = True Then
            RemoveDupeMule = True
        End If
    End Sub
End Class