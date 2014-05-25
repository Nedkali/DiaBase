<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D2StyleOpenFileDialog
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
        Me.components = New System.ComponentModel.Container()
        Me.SavedDatabasesLISTBOX = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DatabaseFilenameCOMBOBOX = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.NewDatabaseLABEL = New System.Windows.Forms.Label()
        Me.OpenDatabaseBUTTON = New System.Windows.Forms.Button()
        Me.NewDatabaseCancelBUTTON = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.SaveBeforeOpeningCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.OpenDatabaseCONTEXTMENU = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.OpenSelectedDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RenameSelectedDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteSelectedDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenDatabaseCONTEXTMENU.SuspendLayout()
        Me.SuspendLayout()
        '
        'SavedDatabasesLISTBOX
        '
        Me.SavedDatabasesLISTBOX.BackColor = System.Drawing.Color.Black
        Me.SavedDatabasesLISTBOX.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.SavedDatabasesLISTBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SavedDatabasesLISTBOX.ForeColor = System.Drawing.Color.White
        Me.SavedDatabasesLISTBOX.FormattingEnabled = True
        Me.SavedDatabasesLISTBOX.ItemHeight = 16
        Me.SavedDatabasesLISTBOX.Location = New System.Drawing.Point(52, 107)
        Me.SavedDatabasesLISTBOX.Name = "SavedDatabasesLISTBOX"
        Me.SavedDatabasesLISTBOX.Size = New System.Drawing.Size(388, 256)
        Me.SavedDatabasesLISTBOX.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Black
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(27, 425)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Database File"
        '
        'DatabaseFilenameCOMBOBOX
        '
        Me.DatabaseFilenameCOMBOBOX.BackColor = System.Drawing.Color.Black
        Me.DatabaseFilenameCOMBOBOX.ForeColor = System.Drawing.Color.White
        Me.DatabaseFilenameCOMBOBOX.FormattingEnabled = True
        Me.DatabaseFilenameCOMBOBOX.Location = New System.Drawing.Point(122, 423)
        Me.DatabaseFilenameCOMBOBOX.Name = "DatabaseFilenameCOMBOBOX"
        Me.DatabaseFilenameCOMBOBOX.Size = New System.Drawing.Size(337, 21)
        Me.DatabaseFilenameCOMBOBOX.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Black
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(36, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(423, 23)
        Me.Label2.TabIndex = 880
        Me.Label2.Text = "SELECT DATABASE TO OPEN"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'NewDatabaseLABEL
        '
        Me.NewDatabaseLABEL.BackColor = System.Drawing.Color.Black
        Me.NewDatabaseLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewDatabaseLABEL.ForeColor = System.Drawing.Color.SeaShell
        Me.NewDatabaseLABEL.Location = New System.Drawing.Point(36, 60)
        Me.NewDatabaseLABEL.Name = "NewDatabaseLABEL"
        Me.NewDatabaseLABEL.Size = New System.Drawing.Size(423, 41)
        Me.NewDatabaseLABEL.TabIndex = 881
        Me.NewDatabaseLABEL.Text = "To open a saved database please select the database from the list or enter the da" & _
    "tabase name into the text box."
        '
        'OpenDatabaseBUTTON
        '
        Me.OpenDatabaseBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.OpenDatabaseBUTTON.BackgroundImage = Global.DiaBase.My.Resources.Resources.back1
        Me.OpenDatabaseBUTTON.FlatAppearance.BorderSize = 2
        Me.OpenDatabaseBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.OpenDatabaseBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.OpenDatabaseBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.OpenDatabaseBUTTON.Location = New System.Drawing.Point(305, 458)
        Me.OpenDatabaseBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.OpenDatabaseBUTTON.Name = "OpenDatabaseBUTTON"
        Me.OpenDatabaseBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.OpenDatabaseBUTTON.TabIndex = 884
        Me.OpenDatabaseBUTTON.Text = "Open"
        Me.OpenDatabaseBUTTON.UseVisualStyleBackColor = False
        '
        'NewDatabaseCancelBUTTON
        '
        Me.NewDatabaseCancelBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.NewDatabaseCancelBUTTON.BackgroundImage = Global.DiaBase.My.Resources.Resources.back1
        Me.NewDatabaseCancelBUTTON.FlatAppearance.BorderSize = 2
        Me.NewDatabaseCancelBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.NewDatabaseCancelBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.NewDatabaseCancelBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.NewDatabaseCancelBUTTON.Location = New System.Drawing.Point(388, 458)
        Me.NewDatabaseCancelBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.NewDatabaseCancelBUTTON.Name = "NewDatabaseCancelBUTTON"
        Me.NewDatabaseCancelBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.NewDatabaseCancelBUTTON.TabIndex = 883
        Me.NewDatabaseCancelBUTTON.Text = "Cancel"
        Me.NewDatabaseCancelBUTTON.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.BurlyWood
        Me.Label3.Location = New System.Drawing.Point(52, 103)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(388, 3)
        Me.Label3.TabIndex = 887
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.BurlyWood
        Me.Label4.Location = New System.Drawing.Point(52, 362)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(388, 3)
        Me.Label4.TabIndex = 888
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.BurlyWood
        Me.Label5.Location = New System.Drawing.Point(121, 423)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(338, 3)
        Me.Label5.TabIndex = 889
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.BurlyWood
        Me.Label6.Location = New System.Drawing.Point(121, 441)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(339, 3)
        Me.Label6.TabIndex = 890
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.BurlyWood
        Me.Label7.Location = New System.Drawing.Point(121, 423)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(3, 20)
        Me.Label7.TabIndex = 891
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.BurlyWood
        Me.Label8.Location = New System.Drawing.Point(457, 423)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(3, 20)
        Me.Label8.TabIndex = 892
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.BurlyWood
        Me.Label9.Location = New System.Drawing.Point(52, 103)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(3, 262)
        Me.Label9.TabIndex = 893
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.BurlyWood
        Me.Label10.Location = New System.Drawing.Point(437, 103)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(3, 262)
        Me.Label10.TabIndex = 894
        '
        'SaveBeforeOpeningCHECKBOX
        '
        Me.SaveBeforeOpeningCHECKBOX.AutoSize = True
        Me.SaveBeforeOpeningCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.SaveBeforeOpeningCHECKBOX.Checked = True
        Me.SaveBeforeOpeningCHECKBOX.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SaveBeforeOpeningCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.SaveBeforeOpeningCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaveBeforeOpeningCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.SaveBeforeOpeningCHECKBOX.Location = New System.Drawing.Point(52, 378)
        Me.SaveBeforeOpeningCHECKBOX.Name = "SaveBeforeOpeningCHECKBOX"
        Me.SaveBeforeOpeningCHECKBOX.Size = New System.Drawing.Size(306, 20)
        Me.SaveBeforeOpeningCHECKBOX.TabIndex = 895
        Me.SaveBeforeOpeningCHECKBOX.Text = "Save The Current Database Before Openening"
        Me.SaveBeforeOpeningCHECKBOX.UseVisualStyleBackColor = False
        '
        'OpenDatabaseCONTEXTMENU
        '
        Me.OpenDatabaseCONTEXTMENU.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenSelectedDatabaseToolStripMenuItem, Me.RenameSelectedDatabaseToolStripMenuItem, Me.DeleteSelectedDatabaseToolStripMenuItem})
        Me.OpenDatabaseCONTEXTMENU.Name = "OpenDatabaseCONTEXTMENU"
        Me.OpenDatabaseCONTEXTMENU.Size = New System.Drawing.Size(227, 70)
        '
        'OpenSelectedDatabaseToolStripMenuItem
        '
        Me.OpenSelectedDatabaseToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.OpenSelectedDatabaseToolStripMenuItem.Name = "OpenSelectedDatabaseToolStripMenuItem"
        Me.OpenSelectedDatabaseToolStripMenuItem.Size = New System.Drawing.Size(226, 22)
        Me.OpenSelectedDatabaseToolStripMenuItem.Text = "Open Selected Database"
        '
        'RenameSelectedDatabaseToolStripMenuItem
        '
        Me.RenameSelectedDatabaseToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.RenameSelectedDatabaseToolStripMenuItem.Name = "RenameSelectedDatabaseToolStripMenuItem"
        Me.RenameSelectedDatabaseToolStripMenuItem.Size = New System.Drawing.Size(226, 22)
        Me.RenameSelectedDatabaseToolStripMenuItem.Text = "Rename Selected Database"
        '
        'DeleteSelectedDatabaseToolStripMenuItem
        '
        Me.DeleteSelectedDatabaseToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.DeleteSelectedDatabaseToolStripMenuItem.Name = "DeleteSelectedDatabaseToolStripMenuItem"
        Me.DeleteSelectedDatabaseToolStripMenuItem.Size = New System.Drawing.Size(226, 22)
        Me.DeleteSelectedDatabaseToolStripMenuItem.Text = "Delete Selected Database"
        '
        'D2StyleOpenFileDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.BigSettings
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(498, 528)
        Me.ControlBox = False
        Me.Controls.Add(Me.SaveBeforeOpeningCHECKBOX)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.OpenDatabaseBUTTON)
        Me.Controls.Add(Me.NewDatabaseCancelBUTTON)
        Me.Controls.Add(Me.NewDatabaseLABEL)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DatabaseFilenameCOMBOBOX)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SavedDatabasesLISTBOX)
        Me.MaximumSize = New System.Drawing.Size(506, 555)
        Me.MinimumSize = New System.Drawing.Size(506, 555)
        Me.Name = "D2StyleOpenFileDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Open A Saved Database"
        Me.OpenDatabaseCONTEXTMENU.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SavedDatabasesLISTBOX As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DatabaseFilenameCOMBOBOX As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents NewDatabaseLABEL As System.Windows.Forms.Label
    Friend WithEvents OpenDatabaseBUTTON As System.Windows.Forms.Button
    Friend WithEvents NewDatabaseCancelBUTTON As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents SaveBeforeOpeningCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents OpenDatabaseCONTEXTMENU As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents OpenSelectedDatabaseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RenameSelectedDatabaseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteSelectedDatabaseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
