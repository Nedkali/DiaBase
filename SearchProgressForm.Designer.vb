<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SearchProgressForm
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
        Me.SearchPROGRESSBAR = New System.Windows.Forms.ProgressBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'SearchPROGRESSBAR
        '
        Me.SearchPROGRESSBAR.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SearchPROGRESSBAR.ForeColor = System.Drawing.Color.BurlyWood
        Me.SearchPROGRESSBAR.Location = New System.Drawing.Point(43, 32)
        Me.SearchPROGRESSBAR.Name = "SearchPROGRESSBAR"
        Me.SearchPROGRESSBAR.Size = New System.Drawing.Size(361, 20)
        Me.SearchPROGRESSBAR.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(90, 65)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(270, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Searching Database For Matches, Please Wait..."
        '
        'SearchProgressForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.mbox
        Me.ClientSize = New System.Drawing.Size(441, 110)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SearchPROGRESSBAR)
        Me.Name = "SearchProgressForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Search In Progress... "
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SearchPROGRESSBAR As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
