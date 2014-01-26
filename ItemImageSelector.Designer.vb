<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ItemImageSelector
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
        Me.AddSelectImageBUTTON = New System.Windows.Forms.Button()
        Me.AddSelectImageCancelBUTTON = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Label128 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AddSelectImageBUTTON
        '
        Me.AddSelectImageBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.AddSelectImageBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.AddSelectImageBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.AddSelectImageBUTTON.Location = New System.Drawing.Point(211, 100)
        Me.AddSelectImageBUTTON.Name = "AddSelectImageBUTTON"
        Me.AddSelectImageBUTTON.Size = New System.Drawing.Size(75, 24)
        Me.AddSelectImageBUTTON.TabIndex = 11
        Me.AddSelectImageBUTTON.Text = "Select"
        Me.AddSelectImageBUTTON.UseVisualStyleBackColor = False
        '
        'AddSelectImageCancelBUTTON
        '
        Me.AddSelectImageCancelBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.AddSelectImageCancelBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.AddSelectImageCancelBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.AddSelectImageCancelBUTTON.Location = New System.Drawing.Point(211, 147)
        Me.AddSelectImageCancelBUTTON.Name = "AddSelectImageCancelBUTTON"
        Me.AddSelectImageCancelBUTTON.Size = New System.Drawing.Size(75, 24)
        Me.AddSelectImageCancelBUTTON.TabIndex = 10
        Me.AddSelectImageCancelBUTTON.Text = "Cancel"
        Me.AddSelectImageCancelBUTTON.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.ControlText
        Me.PictureBox1.BackgroundImage = Global.DiaBase.My.Resources.Resources.ImageBackground
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox1.Location = New System.Drawing.Point(52, 39)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Padding = New System.Windows.Forms.Padding(11, 9, 0, 0)
        Me.PictureBox1.Size = New System.Drawing.Size(106, 132)
        Me.PictureBox1.TabIndex = 17
        Me.PictureBox1.TabStop = False
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.BackColor = System.Drawing.Color.Black
        Me.NumericUpDown1.ForeColor = System.Drawing.Color.White
        Me.NumericUpDown1.Location = New System.Drawing.Point(211, 39)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {658, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(75, 20)
        Me.NumericUpDown1.TabIndex = 18
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label128
        '
        Me.Label128.BackColor = System.Drawing.Color.BurlyWood
        Me.Label128.Location = New System.Drawing.Point(211, 38)
        Me.Label128.Name = "Label128"
        Me.Label128.Size = New System.Drawing.Size(75, 2)
        Me.Label128.TabIndex = 870
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(211, 58)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 2)
        Me.Label1.TabIndex = 871
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.BurlyWood
        Me.Label8.Location = New System.Drawing.Point(210, 38)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(2, 22)
        Me.Label8.TabIndex = 873
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(284, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(2, 22)
        Me.Label2.TabIndex = 874
        Me.Label2.Text = "0"
        '
        'ItemImageSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.d2graphicselect
        Me.ClientSize = New System.Drawing.Size(350, 220)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label128)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.AddSelectImageBUTTON)
        Me.Controls.Add(Me.AddSelectImageCancelBUTTON)
        Me.MaximumSize = New System.Drawing.Size(366, 258)
        Me.MinimumSize = New System.Drawing.Size(366, 258)
        Me.Name = "ItemImageSelector"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select An Item Image"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents AddSelectImageBUTTON As System.Windows.Forms.Button
    Friend WithEvents AddSelectImageCancelBUTTON As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label128 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
