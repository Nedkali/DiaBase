<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OpenContainingDatabase
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(OpenContainingDatabase))
        Me.YesNoCONFIRMBUTTON = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SaveFirstCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.BackupFirstCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.OpenContainingLABEL = New System.Windows.Forms.Label()
        Me.YesNoHeaderLABEL = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'YesNoCONFIRMBUTTON
        '
        Me.YesNoCONFIRMBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.YesNoCONFIRMBUTTON.BackgroundImage = CType(resources.GetObject("YesNoCONFIRMBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.YesNoCONFIRMBUTTON.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.YesNoCONFIRMBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.YesNoCONFIRMBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.YesNoCONFIRMBUTTON.Location = New System.Drawing.Point(310, 197)
        Me.YesNoCONFIRMBUTTON.Name = "YesNoCONFIRMBUTTON"
        Me.YesNoCONFIRMBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.YesNoCONFIRMBUTTON.TabIndex = 2
        Me.YesNoCONFIRMBUTTON.Text = "Cancel"
        Me.YesNoCONFIRMBUTTON.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Button1.BackgroundImage = CType(resources.GetObject("Button1.BackgroundImage"), System.Drawing.Image)
        Me.Button1.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.Color.BurlyWood
        Me.Button1.Location = New System.Drawing.Point(229, 197)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Continue"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'SaveFirstCHECKBOX
        '
        Me.SaveFirstCHECKBOX.AutoSize = True
        Me.SaveFirstCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.SaveFirstCHECKBOX.Checked = True
        Me.SaveFirstCHECKBOX.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SaveFirstCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.SaveFirstCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.SaveFirstCHECKBOX.Location = New System.Drawing.Point(23, 183)
        Me.SaveFirstCHECKBOX.Name = "SaveFirstCHECKBOX"
        Me.SaveFirstCHECKBOX.Size = New System.Drawing.Size(137, 17)
        Me.SaveFirstCHECKBOX.TabIndex = 9
        Me.SaveFirstCHECKBOX.Text = "Save Current Database"
        Me.SaveFirstCHECKBOX.UseVisualStyleBackColor = False
        '
        'BackupFirstCHECKBOX
        '
        Me.BackupFirstCHECKBOX.AutoSize = True
        Me.BackupFirstCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.BackupFirstCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.BackupFirstCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.BackupFirstCHECKBOX.Location = New System.Drawing.Point(23, 203)
        Me.BackupFirstCHECKBOX.Name = "BackupFirstCHECKBOX"
        Me.BackupFirstCHECKBOX.Size = New System.Drawing.Size(149, 17)
        Me.BackupFirstCHECKBOX.TabIndex = 10
        Me.BackupFirstCHECKBOX.Text = "Backup Current Database"
        Me.BackupFirstCHECKBOX.UseVisualStyleBackColor = False
        '
        'OpenContainingLABEL
        '
        Me.OpenContainingLABEL.BackColor = System.Drawing.Color.Black
        Me.OpenContainingLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpenContainingLABEL.ForeColor = System.Drawing.Color.White
        Me.OpenContainingLABEL.Location = New System.Drawing.Point(26, 55)
        Me.OpenContainingLABEL.Name = "OpenContainingLABEL"
        Me.OpenContainingLABEL.Size = New System.Drawing.Size(353, 117)
        Me.OpenContainingLABEL.TabIndex = 11
        Me.OpenContainingLABEL.Text = "Set at runtime"
        '
        'YesNoHeaderLABEL
        '
        Me.YesNoHeaderLABEL.BackColor = System.Drawing.Color.Black
        Me.YesNoHeaderLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YesNoHeaderLABEL.ForeColor = System.Drawing.Color.BurlyWood
        Me.YesNoHeaderLABEL.Location = New System.Drawing.Point(21, 23)
        Me.YesNoHeaderLABEL.Name = "YesNoHeaderLABEL"
        Me.YesNoHeaderLABEL.Size = New System.Drawing.Size(356, 23)
        Me.YesNoHeaderLABEL.TabIndex = 12
        Me.YesNoHeaderLABEL.Text = "Confirm To Open Containing Database"
        Me.YesNoHeaderLABEL.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'OpenContainingDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(409, 247)
        Me.ControlBox = False
        Me.Controls.Add(Me.YesNoHeaderLABEL)
        Me.Controls.Add(Me.OpenContainingLABEL)
        Me.Controls.Add(Me.BackupFirstCHECKBOX)
        Me.Controls.Add(Me.SaveFirstCHECKBOX)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.YesNoCONFIRMBUTTON)
        Me.MaximumSize = New System.Drawing.Size(425, 285)
        Me.MinimumSize = New System.Drawing.Size(425, 285)
        Me.Name = "OpenContainingDatabase"
        Me.Text = "Open Containing Database"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents YesNoCONFIRMBUTTON As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents SaveFirstCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents BackupFirstCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents OpenContainingLABEL As System.Windows.Forms.Label
    Friend WithEvents YesNoHeaderLABEL As System.Windows.Forms.Label
End Class
