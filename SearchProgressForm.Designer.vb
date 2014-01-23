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
        Me.SearchProgressLABEL1 = New System.Windows.Forms.Label()
        Me.SearchProgressLABEL2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'SearchPROGRESSBAR
        '
        Me.SearchPROGRESSBAR.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SearchPROGRESSBAR.ForeColor = System.Drawing.Color.BurlyWood
        Me.SearchPROGRESSBAR.Location = New System.Drawing.Point(43, 33)
        Me.SearchPROGRESSBAR.Name = "SearchPROGRESSBAR"
        Me.SearchPROGRESSBAR.Size = New System.Drawing.Size(361, 20)
        Me.SearchPROGRESSBAR.TabIndex = 0
        '
        'SearchProgressLABEL1
        '
        Me.SearchProgressLABEL1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.SearchProgressLABEL1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SearchProgressLABEL1.ForeColor = System.Drawing.Color.BurlyWood
        Me.SearchProgressLABEL1.Location = New System.Drawing.Point(43, 60)
        Me.SearchProgressLABEL1.Name = "SearchProgressLABEL1"
        Me.SearchProgressLABEL1.Size = New System.Drawing.Size(152, 15)
        Me.SearchProgressLABEL1.TabIndex = 2
        Me.SearchProgressLABEL1.Text = "0 Matches"
        Me.SearchProgressLABEL1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SearchProgressLABEL2
        '
        Me.SearchProgressLABEL2.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.SearchProgressLABEL2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SearchProgressLABEL2.ForeColor = System.Drawing.Color.BurlyWood
        Me.SearchProgressLABEL2.Location = New System.Drawing.Point(201, 60)
        Me.SearchProgressLABEL2.Name = "SearchProgressLABEL2"
        Me.SearchProgressLABEL2.Size = New System.Drawing.Size(203, 15)
        Me.SearchProgressLABEL2.TabIndex = 3
        Me.SearchProgressLABEL2.Text = "Searching 0 Of 0 Item Records"
        Me.SearchProgressLABEL2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'SearchProgressForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.mbox
        Me.ClientSize = New System.Drawing.Size(441, 110)
        Me.ControlBox = False
        Me.Controls.Add(Me.SearchProgressLABEL2)
        Me.Controls.Add(Me.SearchProgressLABEL1)
        Me.Controls.Add(Me.SearchPROGRESSBAR)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SearchProgressForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Searching Database, Please Wait...."
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SearchPROGRESSBAR As System.Windows.Forms.ProgressBar
    Friend WithEvents SearchProgressLABEL1 As System.Windows.Forms.Label
    Friend WithEvents SearchProgressLABEL2 As System.Windows.Forms.Label
End Class
