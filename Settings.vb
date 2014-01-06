Public Class Settings

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles CancelDefaultsBUTTON.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles EtalPathBrowseBUTTON.Click
        EtalPathTEXTBOX.Text = ""
        PictureBox1.Visible = False
        DatabaseFileTEXTBOX.Text = ""
        'todo - set a variable and set to selected folder
        FolderBrowserDialog1.ShowDialog()
        Dim temp As Array
        temp = Split(FolderBrowserDialog1.SelectedPath, "\")
        If temp(temp.Length - 1) = "D2NT" Then EtalPathTEXTBOX.Text = FolderBrowserDialog1.SelectedPath
        If temp(temp.Length - 1) <> "D2NT" Then
            MessageBox.Show("not correct folder", "Error")
            Return
        End If
        If EtalPathTEXTBOX.Text <> Nothing Then PictureBox1.Visible = True

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
        If OpenFileDialog1.FileName <> Nothing Then DatabaseFileTEXTBOX.Text = OpenFileDialog1.FileName
        If DatabaseFileTEXTBOX.Text <> Nothing Then PictureBox2.Visible = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles SaveDefaultsBUTTON.Click

        If PictureBox1.Visible = False And PictureBox2.Visible = False Then
            MessageBox.Show("Settings incomplete", "Error")
            Return
        End If

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath + "\Settings.cfg", False)
        file.WriteLine(EtalPathTEXTBOX.Text)
        file.WriteLine(DatabaseFileTEXTBOX.Text)
        file.WriteLine(NumericUpDown1.Value)
        file.WriteLine(CheckBox3.Checked)
        file.Close()
        EtalPath = EtalPathTEXTBOX.Text
        DataBaseFile = DatabaseFileTEXTBOX.Text
        TimerMins = NumericUpDown1.Value
        Me.Close()

    End Sub

    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        
        If DatabaseFileTEXTBOX.Text <> Nothing Then PictureBox2.Visible = True
        If EtalPathTEXTBOX.Text <> Nothing Then PictureBox1.Visible = True


    End Sub
End Class