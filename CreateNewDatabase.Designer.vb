<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreateNewDatabase
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
        Me.NewDatabaseLABEL = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.NewDatabaseTEXTBOX = New System.Windows.Forms.TextBox()
        Me.NewDatabaseCancelBUTTON = New System.Windows.Forms.Button()
        Me.NewDatabaseCreateBUTTON = New System.Windows.Forms.Button()
        Me.NewDatabaseAutoOpenCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'NewDatabaseLABEL
        '
        Me.NewDatabaseLABEL.BackColor = System.Drawing.Color.Black
        Me.NewDatabaseLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewDatabaseLABEL.ForeColor = System.Drawing.Color.SeaShell
        Me.NewDatabaseLABEL.Location = New System.Drawing.Point(37, 57)
        Me.NewDatabaseLABEL.Name = "NewDatabaseLABEL"
        Me.NewDatabaseLABEL.Size = New System.Drawing.Size(326, 45)
        Me.NewDatabaseLABEL.TabIndex = 1
        Me.NewDatabaseLABEL.Text = "To create a new database please enter its unique name into the textbox below..."
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.BurlyWood
        Me.Label4.Location = New System.Drawing.Point(362, 121)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(3, 20)
        Me.Label4.TabIndex = 873
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.BurlyWood
        Me.Label26.Location = New System.Drawing.Point(34, 121)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(3, 20)
        Me.Label26.TabIndex = 872
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(34, 139)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(331, 3)
        Me.Label2.TabIndex = 871
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.BurlyWood
        Me.Label17.Location = New System.Drawing.Point(34, 120)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(331, 3)
        Me.Label17.TabIndex = 870
        '
        'NewDatabaseTEXTBOX
        '
        Me.NewDatabaseTEXTBOX.BackColor = System.Drawing.Color.Black
        Me.NewDatabaseTEXTBOX.ForeColor = System.Drawing.Color.White
        Me.NewDatabaseTEXTBOX.Location = New System.Drawing.Point(35, 121)
        Me.NewDatabaseTEXTBOX.Name = "NewDatabaseTEXTBOX"
        Me.NewDatabaseTEXTBOX.Size = New System.Drawing.Size(329, 20)
        Me.NewDatabaseTEXTBOX.TabIndex = 869
        '
        'NewDatabaseCancelBUTTON
        '
        Me.NewDatabaseCancelBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.NewDatabaseCancelBUTTON.BackgroundImage = Global.DiaBase.My.Resources.Resources.back1
        Me.NewDatabaseCancelBUTTON.FlatAppearance.BorderSize = 2
        Me.NewDatabaseCancelBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.NewDatabaseCancelBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.NewDatabaseCancelBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.NewDatabaseCancelBUTTON.Location = New System.Drawing.Point(309, 194)
        Me.NewDatabaseCancelBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.NewDatabaseCancelBUTTON.Name = "NewDatabaseCancelBUTTON"
        Me.NewDatabaseCancelBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.NewDatabaseCancelBUTTON.TabIndex = 874
        Me.NewDatabaseCancelBUTTON.Text = "Cancel"
        Me.NewDatabaseCancelBUTTON.UseVisualStyleBackColor = False
        '
        'NewDatabaseCreateBUTTON
        '
        Me.NewDatabaseCreateBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.NewDatabaseCreateBUTTON.BackgroundImage = Global.DiaBase.My.Resources.Resources.back1
        Me.NewDatabaseCreateBUTTON.FlatAppearance.BorderSize = 2
        Me.NewDatabaseCreateBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.NewDatabaseCreateBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.NewDatabaseCreateBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.NewDatabaseCreateBUTTON.Location = New System.Drawing.Point(226, 194)
        Me.NewDatabaseCreateBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.NewDatabaseCreateBUTTON.Name = "NewDatabaseCreateBUTTON"
        Me.NewDatabaseCreateBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.NewDatabaseCreateBUTTON.TabIndex = 875
        Me.NewDatabaseCreateBUTTON.Text = "Create"
        Me.NewDatabaseCreateBUTTON.UseVisualStyleBackColor = False
        '
        'NewDatabaseAutoOpenCHECKBOX
        '
        Me.NewDatabaseAutoOpenCHECKBOX.AutoSize = True
        Me.NewDatabaseAutoOpenCHECKBOX.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.NewDatabaseAutoOpenCHECKBOX.Checked = True
        Me.NewDatabaseAutoOpenCHECKBOX.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NewDatabaseAutoOpenCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.NewDatabaseAutoOpenCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewDatabaseAutoOpenCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.NewDatabaseAutoOpenCHECKBOX.Location = New System.Drawing.Point(34, 157)
        Me.NewDatabaseAutoOpenCHECKBOX.Name = "NewDatabaseAutoOpenCHECKBOX"
        Me.NewDatabaseAutoOpenCHECKBOX.Size = New System.Drawing.Size(314, 20)
        Me.NewDatabaseAutoOpenCHECKBOX.TabIndex = 876
        Me.NewDatabaseAutoOpenCHECKBOX.Text = "Save The Current Database And Open This One"
        Me.NewDatabaseAutoOpenCHECKBOX.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Black
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(37, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(326, 23)
        Me.Label1.TabIndex = 877
        Me.Label1.Text = "ENTER DATABASE NAME"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'CreateNewDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(407, 250)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.NewDatabaseAutoOpenCHECKBOX)
        Me.Controls.Add(Me.NewDatabaseCreateBUTTON)
        Me.Controls.Add(Me.NewDatabaseCancelBUTTON)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.NewDatabaseTEXTBOX)
        Me.Controls.Add(Me.NewDatabaseLABEL)
        Me.MaximumSize = New System.Drawing.Size(423, 288)
        Me.MinimumSize = New System.Drawing.Size(423, 288)
        Me.Name = "CreateNewDatabase"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Create A New Database File"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents NewDatabaseLABEL As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents NewDatabaseTEXTBOX As System.Windows.Forms.TextBox
    Friend WithEvents NewDatabaseCancelBUTTON As System.Windows.Forms.Button
    Friend WithEvents NewDatabaseCreateBUTTON As System.Windows.Forms.Button
    Friend WithEvents NewDatabaseAutoOpenCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
