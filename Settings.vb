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
        DatabaseFileTEXTBOX.Text = openFileDialog1.FileName
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
        file.WriteLine(CheckBox3.Checked)
        file.WriteLine(AutoBackupImportsCHECKBOX.Checked) ' added this for auto backup on import setting
        file.WriteLine(BackupOnEditsCHECKBOX.Checked)   ' added this for auto backup on edits setting
        file.Close()
        EtalPath = EtalPathTEXTBOX.Text
        DataBaseFile = DatabaseFileTEXTBOX.Text
        TimerMins = NumericUpDown1.Value
        TimerSecs = TimerMins * 60
        Me.Close()

    End Sub

    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = Form1.Left + 150
        Me.Top = Form1.Top + 100
        EtalPathTEXTBOX.Text = EtalPath
        DatabaseFileTEXTBOX.Text = DatabaseFile
        NumericUpDown1.Value = TimerMins
        CheckBox3.Checked = KeepPassPrivate
        AutoBackupImportsCHECKBOX.Checked = AutoBackups
        BackupOnEditsCHECKBOX.Checked = EditBackups
        SettingsChecker()

    End Sub

    Private Sub SettingsChecker()

        If My.Computer.FileSystem.DirectoryExists(EtalPathTEXTBOX.Text) = True Then
            Dim temp = Split(EtalPathTEXTBOX.Text, "\")
            If temp(temp.Length - 1) = "D2NT" Then PictureBox1.Visible = True
        End If

        If My.Computer.FileSystem.FileExists(DatabaseFileTEXTBOX.Text) = True Then PictureBox2.Visible = True
    End Sub

    
End Class