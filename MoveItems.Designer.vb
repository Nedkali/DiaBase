<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MoveItems
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MoveItems))
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.DatabaseFilenameTEXTBOX = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.NewDatabaseLABEL = New System.Windows.Forms.Label()
        Me.MoveItemsExportBUTTON = New System.Windows.Forms.Button()
        Me.MoveItemsCancelBUTTON = New System.Windows.Forms.Button()
        Me.DeleteItemsCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.OpenDatabaseCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.SavedDatabasesLISTBOX = New System.Windows.Forms.ListBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.BurlyWood
        Me.Label4.Location = New System.Drawing.Point(415, 339)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(3, 20)
        Me.Label4.TabIndex = 878
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.BurlyWood
        Me.Label26.Location = New System.Drawing.Point(42, 338)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(3, 20)
        Me.Label26.TabIndex = 877
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(42, 356)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(376, 3)
        Me.Label2.TabIndex = 876
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.BurlyWood
        Me.Label17.Location = New System.Drawing.Point(42, 337)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(376, 3)
        Me.Label17.TabIndex = 875
        '
        'DatabaseFilenameTEXTBOX
        '
        Me.DatabaseFilenameTEXTBOX.BackColor = System.Drawing.Color.Black
        Me.DatabaseFilenameTEXTBOX.ForeColor = System.Drawing.Color.White
        Me.DatabaseFilenameTEXTBOX.Location = New System.Drawing.Point(43, 338)
        Me.DatabaseFilenameTEXTBOX.Name = "DatabaseFilenameTEXTBOX"
        Me.DatabaseFilenameTEXTBOX.Size = New System.Drawing.Size(374, 20)
        Me.DatabaseFilenameTEXTBOX.TabIndex = 874
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Black
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(94, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(442, 23)
        Me.Label1.TabIndex = 879
        Me.Label1.Text = "ENTER DESTINATION DATABASE"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'NewDatabaseLABEL
        '
        Me.NewDatabaseLABEL.BackColor = System.Drawing.Color.Black
        Me.NewDatabaseLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewDatabaseLABEL.ForeColor = System.Drawing.Color.SeaShell
        Me.NewDatabaseLABEL.Location = New System.Drawing.Point(39, 69)
        Me.NewDatabaseLABEL.Name = "NewDatabaseLABEL"
        Me.NewDatabaseLABEL.Size = New System.Drawing.Size(560, 74)
        Me.NewDatabaseLABEL.TabIndex = 880
        Me.NewDatabaseLABEL.Text = resources.GetString("NewDatabaseLABEL.Text")
        '
        'MoveItemsExportBUTTON
        '
        Me.MoveItemsExportBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.MoveItemsExportBUTTON.BackgroundImage = Global.DiaBase.My.Resources.Resources.back1
        Me.MoveItemsExportBUTTON.FlatAppearance.BorderSize = 2
        Me.MoveItemsExportBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.MoveItemsExportBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.MoveItemsExportBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.MoveItemsExportBUTTON.Location = New System.Drawing.Point(439, 334)
        Me.MoveItemsExportBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.MoveItemsExportBUTTON.Name = "MoveItemsExportBUTTON"
        Me.MoveItemsExportBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.MoveItemsExportBUTTON.TabIndex = 882
        Me.MoveItemsExportBUTTON.Text = "Export"
        Me.MoveItemsExportBUTTON.UseVisualStyleBackColor = False
        '
        'MoveItemsCancelBUTTON
        '
        Me.MoveItemsCancelBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.MoveItemsCancelBUTTON.BackgroundImage = Global.DiaBase.My.Resources.Resources.back1
        Me.MoveItemsCancelBUTTON.FlatAppearance.BorderSize = 2
        Me.MoveItemsCancelBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.MoveItemsCancelBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.MoveItemsCancelBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.MoveItemsCancelBUTTON.Location = New System.Drawing.Point(522, 334)
        Me.MoveItemsCancelBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.MoveItemsCancelBUTTON.Name = "MoveItemsCancelBUTTON"
        Me.MoveItemsCancelBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.MoveItemsCancelBUTTON.TabIndex = 881
        Me.MoveItemsCancelBUTTON.Text = "Cancel"
        Me.MoveItemsCancelBUTTON.UseVisualStyleBackColor = False
        '
        'DeleteItemsCHECKBOX
        '
        Me.DeleteItemsCHECKBOX.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.DeleteItemsCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.DeleteItemsCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeleteItemsCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.DeleteItemsCHECKBOX.Location = New System.Drawing.Point(347, 190)
        Me.DeleteItemsCHECKBOX.Name = "DeleteItemsCHECKBOX"
        Me.DeleteItemsCHECKBOX.Size = New System.Drawing.Size(249, 52)
        Me.DeleteItemsCHECKBOX.TabIndex = 883
        Me.DeleteItemsCHECKBOX.Text = "Do Not Delete Exported Items From The Source Database"
        Me.DeleteItemsCHECKBOX.UseVisualStyleBackColor = False
        '
        'OpenDatabaseCHECKBOX
        '
        Me.OpenDatabaseCHECKBOX.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.OpenDatabaseCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.OpenDatabaseCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpenDatabaseCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.OpenDatabaseCHECKBOX.Location = New System.Drawing.Point(347, 131)
        Me.OpenDatabaseCHECKBOX.Name = "OpenDatabaseCHECKBOX"
        Me.OpenDatabaseCHECKBOX.Size = New System.Drawing.Size(249, 60)
        Me.OpenDatabaseCHECKBOX.TabIndex = 884
        Me.OpenDatabaseCHECKBOX.Text = "Open Desination Database Once Items Are Exported"
        Me.OpenDatabaseCHECKBOX.UseVisualStyleBackColor = False
        '
        'SavedDatabasesLISTBOX
        '
        Me.SavedDatabasesLISTBOX.BackColor = System.Drawing.Color.Black
        Me.SavedDatabasesLISTBOX.ForeColor = System.Drawing.Color.White
        Me.SavedDatabasesLISTBOX.FormattingEnabled = True
        Me.SavedDatabasesLISTBOX.Location = New System.Drawing.Point(42, 148)
        Me.SavedDatabasesLISTBOX.Name = "SavedDatabasesLISTBOX"
        Me.SavedDatabasesLISTBOX.Size = New System.Drawing.Size(288, 173)
        Me.SavedDatabasesLISTBOX.TabIndex = 885
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.BurlyWood
        Me.Label3.Location = New System.Drawing.Point(44, 148)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(285, 3)
        Me.Label3.TabIndex = 886
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.BurlyWood
        Me.Label5.Location = New System.Drawing.Point(44, 318)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(285, 3)
        Me.Label5.TabIndex = 887
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.BurlyWood
        Me.Label6.Location = New System.Drawing.Point(327, 148)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(3, 173)
        Me.Label6.TabIndex = 888
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.BurlyWood
        Me.Label7.Location = New System.Drawing.Point(42, 148)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(3, 173)
        Me.Label7.TabIndex = 889
        '
        'MoveItems
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(634, 403)
        Me.ControlBox = False
        Me.Controls.Add(Me.NewDatabaseLABEL)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.SavedDatabasesLISTBOX)
        Me.Controls.Add(Me.OpenDatabaseCHECKBOX)
        Me.Controls.Add(Me.DeleteItemsCHECKBOX)
        Me.Controls.Add(Me.MoveItemsExportBUTTON)
        Me.Controls.Add(Me.MoveItemsCancelBUTTON)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.DatabaseFilenameTEXTBOX)
        Me.Name = "MoveItems"
        Me.Text = "Export Items To A Second Database"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents DatabaseFilenameTEXTBOX As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents NewDatabaseLABEL As System.Windows.Forms.Label
    Friend WithEvents MoveItemsExportBUTTON As System.Windows.Forms.Button
    Friend WithEvents MoveItemsCancelBUTTON As System.Windows.Forms.Button
    Friend WithEvents DeleteItemsCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents OpenDatabaseCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents SavedDatabasesLISTBOX As System.Windows.Forms.ListBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
End Class
