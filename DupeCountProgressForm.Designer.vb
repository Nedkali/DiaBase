<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DupeCountProgressForm
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.DupePROGRESSBAR = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(400, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(2, 24)
        Me.Label2.TabIndex = 158
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.BurlyWood
        Me.Label19.Location = New System.Drawing.Point(39, 39)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(2, 23)
        Me.Label19.TabIndex = 157
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(39, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label1.Size = New System.Drawing.Size(363, 2)
        Me.Label1.TabIndex = 156
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.BurlyWood
        Me.Label7.Location = New System.Drawing.Point(39, 60)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label7.Size = New System.Drawing.Size(363, 2)
        Me.Label7.TabIndex = 155
        '
        'DupePROGRESSBAR
        '
        Me.DupePROGRESSBAR.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.DupePROGRESSBAR.ForeColor = System.Drawing.Color.BurlyWood
        Me.DupePROGRESSBAR.Location = New System.Drawing.Point(40, 40)
        Me.DupePROGRESSBAR.Name = "DupePROGRESSBAR"
        Me.DupePROGRESSBAR.Size = New System.Drawing.Size(361, 20)
        Me.DupePROGRESSBAR.TabIndex = 152
        '
        'DupeCountProgressForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.mbox
        Me.ClientSize = New System.Drawing.Size(441, 110)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.DupePROGRESSBAR)
        Me.MaximumSize = New System.Drawing.Size(457, 148)
        Me.MinimumSize = New System.Drawing.Size(457, 148)
        Me.Name = "DupeCountProgressForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Checking For Duped Items, Please Wait..."
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents DupePROGRESSBAR As System.Windows.Forms.ProgressBar
End Class
