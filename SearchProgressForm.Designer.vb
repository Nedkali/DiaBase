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
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'SearchPROGRESSBAR
        '
        Me.SearchPROGRESSBAR.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SearchPROGRESSBAR.ForeColor = System.Drawing.Color.BurlyWood
        Me.SearchPROGRESSBAR.Location = New System.Drawing.Point(40, 33)
        Me.SearchPROGRESSBAR.Name = "SearchPROGRESSBAR"
        Me.SearchPROGRESSBAR.Size = New System.Drawing.Size(361, 20)
        Me.SearchPROGRESSBAR.TabIndex = 0
        '
        'SearchProgressLABEL1
        '
        Me.SearchProgressLABEL1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.SearchProgressLABEL1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SearchProgressLABEL1.ForeColor = System.Drawing.Color.BurlyWood
        Me.SearchProgressLABEL1.Location = New System.Drawing.Point(38, 60)
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
        Me.SearchProgressLABEL2.Location = New System.Drawing.Point(200, 60)
        Me.SearchProgressLABEL2.Name = "SearchProgressLABEL2"
        Me.SearchProgressLABEL2.Size = New System.Drawing.Size(203, 15)
        Me.SearchProgressLABEL2.TabIndex = 3
        Me.SearchProgressLABEL2.Text = "Searching 0 Of 0 Item Records"
        Me.SearchProgressLABEL2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.BurlyWood
        Me.Label7.Location = New System.Drawing.Point(39, 53)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label7.Size = New System.Drawing.Size(363, 2)
        Me.Label7.TabIndex = 134
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(39, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label1.Size = New System.Drawing.Size(363, 2)
        Me.Label1.TabIndex = 135
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.BurlyWood
        Me.Label19.Location = New System.Drawing.Point(39, 31)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(2, 23)
        Me.Label19.TabIndex = 150
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(401, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(2, 24)
        Me.Label2.TabIndex = 151
        '
        'SearchProgressForm
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
        Me.Controls.Add(Me.SearchProgressLABEL2)
        Me.Controls.Add(Me.SearchProgressLABEL1)
        Me.Controls.Add(Me.SearchPROGRESSBAR)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SearchProgressForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Checking For Matches, Please Wait...."
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SearchPROGRESSBAR As System.Windows.Forms.ProgressBar
    Friend WithEvents SearchProgressLABEL1 As System.Windows.Forms.Label
    Friend WithEvents SearchProgressLABEL2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
