<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ClosingAppForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ClosingAppForm))
        Me.BackupDatabaseCHRCKBOX = New System.Windows.Forms.CheckBox()
        Me.SaveDatabaseCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.AppCloseCancelBUTTONButton1 = New System.Windows.Forms.Button()
        Me.AppCloseConfirmBUTTON = New System.Windows.Forms.Button()
        Me.CloseAppInfoLABEL = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'BackupDatabaseCHRCKBOX
        '
        Me.BackupDatabaseCHRCKBOX.AutoSize = True
        Me.BackupDatabaseCHRCKBOX.BackColor = System.Drawing.Color.Black
        Me.BackupDatabaseCHRCKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.BackupDatabaseCHRCKBOX.Location = New System.Drawing.Point(38, 169)
        Me.BackupDatabaseCHRCKBOX.Name = "BackupDatabaseCHRCKBOX"
        Me.BackupDatabaseCHRCKBOX.Size = New System.Drawing.Size(112, 17)
        Me.BackupDatabaseCHRCKBOX.TabIndex = 0
        Me.BackupDatabaseCHRCKBOX.Text = "Backup Database"
        Me.BackupDatabaseCHRCKBOX.UseVisualStyleBackColor = False
        '
        'SaveDatabaseCHECKBOX
        '
        Me.SaveDatabaseCHECKBOX.AutoSize = True
        Me.SaveDatabaseCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.SaveDatabaseCHECKBOX.Checked = True
        Me.SaveDatabaseCHECKBOX.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SaveDatabaseCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.SaveDatabaseCHECKBOX.Location = New System.Drawing.Point(38, 146)
        Me.SaveDatabaseCHECKBOX.Name = "SaveDatabaseCHECKBOX"
        Me.SaveDatabaseCHECKBOX.Size = New System.Drawing.Size(100, 17)
        Me.SaveDatabaseCHECKBOX.TabIndex = 1
        Me.SaveDatabaseCHECKBOX.Text = "Save Database"
        Me.SaveDatabaseCHECKBOX.UseVisualStyleBackColor = False
        '
        'AppCloseCancelBUTTONButton1
        '
        Me.AppCloseCancelBUTTONButton1.BackColor = System.Drawing.Color.DimGray
        Me.AppCloseCancelBUTTONButton1.BackgroundImage = CType(resources.GetObject("AppCloseCancelBUTTONButton1.BackgroundImage"), System.Drawing.Image)
        Me.AppCloseCancelBUTTONButton1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.AppCloseCancelBUTTONButton1.ForeColor = System.Drawing.Color.BurlyWood
        Me.AppCloseCancelBUTTONButton1.Location = New System.Drawing.Point(285, 166)
        Me.AppCloseCancelBUTTONButton1.Name = "AppCloseCancelBUTTONButton1"
        Me.AppCloseCancelBUTTONButton1.Size = New System.Drawing.Size(75, 23)
        Me.AppCloseCancelBUTTONButton1.TabIndex = 2
        Me.AppCloseCancelBUTTONButton1.Text = "Cancel"
        Me.AppCloseCancelBUTTONButton1.UseVisualStyleBackColor = False
        '
        'AppCloseConfirmBUTTON
        '
        Me.AppCloseConfirmBUTTON.BackColor = System.Drawing.Color.DimGray
        Me.AppCloseConfirmBUTTON.BackgroundImage = CType(resources.GetObject("AppCloseConfirmBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.AppCloseConfirmBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.AppCloseConfirmBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.AppCloseConfirmBUTTON.Location = New System.Drawing.Point(193, 166)
        Me.AppCloseConfirmBUTTON.Name = "AppCloseConfirmBUTTON"
        Me.AppCloseConfirmBUTTON.Size = New System.Drawing.Size(86, 23)
        Me.AppCloseConfirmBUTTON.TabIndex = 3
        Me.AppCloseConfirmBUTTON.Text = "Exit"
        Me.AppCloseConfirmBUTTON.UseVisualStyleBackColor = False
        '
        'CloseAppInfoLABEL
        '
        Me.CloseAppInfoLABEL.BackColor = System.Drawing.Color.Black
        Me.CloseAppInfoLABEL.ForeColor = System.Drawing.Color.White
        Me.CloseAppInfoLABEL.Location = New System.Drawing.Point(35, 62)
        Me.CloseAppInfoLABEL.Name = "CloseAppInfoLABEL"
        Me.CloseAppInfoLABEL.Size = New System.Drawing.Size(325, 73)
        Me.CloseAppInfoLABEL.TabIndex = 4
        Me.CloseAppInfoLABEL.Text = resources.GetString("CloseAppInfoLABEL.Text")
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Black
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(35, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(325, 23)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "SAVE AND BACKUP BEFORE EXITING"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ClosingAppForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.d2graphicselect
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(397, 224)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CloseAppInfoLABEL)
        Me.Controls.Add(Me.AppCloseConfirmBUTTON)
        Me.Controls.Add(Me.AppCloseCancelBUTTONButton1)
        Me.Controls.Add(Me.SaveDatabaseCHECKBOX)
        Me.Controls.Add(Me.BackupDatabaseCHRCKBOX)
        Me.Name = "ClosingAppForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Close DiaBase Confirmation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BackupDatabaseCHRCKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents SaveDatabaseCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents AppCloseCancelBUTTONButton1 As System.Windows.Forms.Button
    Friend WithEvents AppCloseConfirmBUTTON As System.Windows.Forms.Button
    Friend WithEvents CloseAppInfoLABEL As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
