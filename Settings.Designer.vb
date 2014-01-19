<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Settings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Settings))
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SaveDefaultsBUTTON = New System.Windows.Forms.Button()
        Me.DefaultDatabaseBrowseBUTTON = New System.Windows.Forms.Button()
        Me.EtalPathBrowseBUTTON = New System.Windows.Forms.Button()
        Me.CancelDefaultsBUTTON = New System.Windows.Forms.Button()
        Me.DatabaseFileTEXTBOX = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.EtalPathTEXTBOX = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.AutoBackupCHECKBOX = New System.Windows.Forms.CheckBox()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(407, 188)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 13)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "Minuites"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.WindowText
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.Label4.Location = New System.Drawing.Point(332, 160)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(121, 13)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "Time between file scans"
        '
        'SaveDefaultsBUTTON
        '
        Me.SaveDefaultsBUTTON.Location = New System.Drawing.Point(326, 240)
        Me.SaveDefaultsBUTTON.Name = "SaveDefaultsBUTTON"
        Me.SaveDefaultsBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.SaveDefaultsBUTTON.TabIndex = 27
        Me.SaveDefaultsBUTTON.Text = "Save"
        Me.SaveDefaultsBUTTON.UseVisualStyleBackColor = True
        '
        'DefaultDatabaseBrowseBUTTON
        '
        Me.DefaultDatabaseBrowseBUTTON.Location = New System.Drawing.Point(377, 115)
        Me.DefaultDatabaseBrowseBUTTON.Name = "DefaultDatabaseBrowseBUTTON"
        Me.DefaultDatabaseBrowseBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.DefaultDatabaseBrowseBUTTON.TabIndex = 26
        Me.DefaultDatabaseBrowseBUTTON.Text = "Browse"
        Me.DefaultDatabaseBrowseBUTTON.UseVisualStyleBackColor = True
        '
        'EtalPathBrowseBUTTON
        '
        Me.EtalPathBrowseBUTTON.Location = New System.Drawing.Point(377, 53)
        Me.EtalPathBrowseBUTTON.Name = "EtalPathBrowseBUTTON"
        Me.EtalPathBrowseBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.EtalPathBrowseBUTTON.TabIndex = 25
        Me.EtalPathBrowseBUTTON.Text = "Browse"
        Me.EtalPathBrowseBUTTON.UseVisualStyleBackColor = True
        '
        'CancelDefaultsBUTTON
        '
        Me.CancelDefaultsBUTTON.Location = New System.Drawing.Point(407, 240)
        Me.CancelDefaultsBUTTON.Name = "CancelDefaultsBUTTON"
        Me.CancelDefaultsBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.CancelDefaultsBUTTON.TabIndex = 24
        Me.CancelDefaultsBUTTON.Text = "Cancel"
        Me.CancelDefaultsBUTTON.UseVisualStyleBackColor = True
        '
        'DatabaseFileTEXTBOX
        '
        Me.DatabaseFileTEXTBOX.Location = New System.Drawing.Point(74, 115)
        Me.DatabaseFileTEXTBOX.Name = "DatabaseFileTEXTBOX"
        Me.DatabaseFileTEXTBOX.Size = New System.Drawing.Size(275, 20)
        Me.DatabaseFileTEXTBOX.TabIndex = 22
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.Label1.Location = New System.Drawing.Point(71, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(125, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Etal Folder - eg C:\D2NT"
        '
        'EtalPathTEXTBOX
        '
        Me.EtalPathTEXTBOX.Location = New System.Drawing.Point(74, 53)
        Me.EtalPathTEXTBOX.Name = "EtalPathTEXTBOX"
        Me.EtalPathTEXTBOX.Size = New System.Drawing.Size(275, 20)
        Me.EtalPathTEXTBOX.TabIndex = 20
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.WindowText
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.Label3.Location = New System.Drawing.Point(71, 90)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 13)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Default Database"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(335, 185)
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(56, 20)
        Me.NumericUpDown1.TabIndex = 33
        Me.NumericUpDown1.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.BackColor = System.Drawing.SystemColors.WindowText
        Me.CheckBox3.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.CheckBox3.Location = New System.Drawing.Point(42, 162)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(212, 17)
        Me.CheckBox3.TabIndex = 34
        Me.CheckBox3.Text = "Make Mule Account Password Invisible"
        Me.CheckBox3.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.WindowText
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(42, 49)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(23, 24)
        Me.PictureBox1.TabIndex = 35
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.SystemColors.WindowText
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(42, 111)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(23, 24)
        Me.PictureBox2.TabIndex = 36
        Me.PictureBox2.TabStop = False
        Me.PictureBox2.Visible = False
        '
        'AutoBackupCHECKBOX
        '
        Me.AutoBackupCHECKBOX.AutoSize = True
        Me.AutoBackupCHECKBOX.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.AutoBackupCHECKBOX.ForeColor = System.Drawing.Color.White
        Me.AutoBackupCHECKBOX.Location = New System.Drawing.Point(42, 193)
        Me.AutoBackupCHECKBOX.Name = "AutoBackupCHECKBOX"
        Me.AutoBackupCHECKBOX.Size = New System.Drawing.Size(216, 17)
        Me.AutoBackupCHECKBOX.TabIndex = 37
        Me.AutoBackupCHECKBOX.Text = "Backup Database Before Importng Logs"
        Me.AutoBackupCHECKBOX.UseVisualStyleBackColor = False
        '
        'Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.ClientSize = New System.Drawing.Size(511, 295)
        Me.Controls.Add(Me.AutoBackupCHECKBOX)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.SaveDefaultsBUTTON)
        Me.Controls.Add(Me.DefaultDatabaseBrowseBUTTON)
        Me.Controls.Add(Me.EtalPathBrowseBUTTON)
        Me.Controls.Add(Me.CancelDefaultsBUTTON)
        Me.Controls.Add(Me.DatabaseFileTEXTBOX)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.EtalPathTEXTBOX)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Settings"
        Me.Text = "Settings"
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents SaveDefaultsBUTTON As System.Windows.Forms.Button
    Friend WithEvents DefaultDatabaseBrowseBUTTON As System.Windows.Forms.Button
    Friend WithEvents EtalPathBrowseBUTTON As System.Windows.Forms.Button
    Friend WithEvents CancelDefaultsBUTTON As System.Windows.Forms.Button
    Friend WithEvents DatabaseFileTEXTBOX As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents EtalPathTEXTBOX As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents AutoBackupCHECKBOX As System.Windows.Forms.CheckBox
End Class
