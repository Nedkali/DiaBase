<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class YesNoD2Style
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(YesNoD2Style))
        Me.YesNoCANCELBUTTON = New System.Windows.Forms.Button()
        Me.YesNoCONFIRMBUTTON = New System.Windows.Forms.Button()
        Me.YesNoMessageLABEL = New System.Windows.Forms.Label()
        Me.YesNoHeaderLABEL = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'YesNoCANCELBUTTON
        '
        Me.YesNoCANCELBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.YesNoCANCELBUTTON.BackgroundImage = CType(resources.GetObject("YesNoCANCELBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.YesNoCANCELBUTTON.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.YesNoCANCELBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.YesNoCANCELBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.YesNoCANCELBUTTON.Location = New System.Drawing.Point(319, 214)
        Me.YesNoCANCELBUTTON.Name = "YesNoCANCELBUTTON"
        Me.YesNoCANCELBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.YesNoCANCELBUTTON.TabIndex = 0
        Me.YesNoCANCELBUTTON.Text = "Cancel"
        Me.YesNoCANCELBUTTON.UseVisualStyleBackColor = False
        '
        'YesNoCONFIRMBUTTON
        '
        Me.YesNoCONFIRMBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.YesNoCONFIRMBUTTON.BackgroundImage = CType(resources.GetObject("YesNoCONFIRMBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.YesNoCONFIRMBUTTON.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.YesNoCONFIRMBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.YesNoCONFIRMBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.YesNoCONFIRMBUTTON.Location = New System.Drawing.Point(222, 214)
        Me.YesNoCONFIRMBUTTON.Name = "YesNoCONFIRMBUTTON"
        Me.YesNoCONFIRMBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.YesNoCONFIRMBUTTON.TabIndex = 1
        Me.YesNoCONFIRMBUTTON.Text = "Confirm"
        Me.YesNoCONFIRMBUTTON.UseVisualStyleBackColor = False
        '
        'YesNoMessageLABEL
        '
        Me.YesNoMessageLABEL.BackColor = System.Drawing.Color.Black
        Me.YesNoMessageLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YesNoMessageLABEL.ForeColor = System.Drawing.Color.White
        Me.YesNoMessageLABEL.Location = New System.Drawing.Point(39, 76)
        Me.YesNoMessageLABEL.Name = "YesNoMessageLABEL"
        Me.YesNoMessageLABEL.Size = New System.Drawing.Size(353, 124)
        Me.YesNoMessageLABEL.TabIndex = 2
        Me.YesNoMessageLABEL.Text = "YesNoMessageLABEL - Set at runtime"
        '
        'YesNoHeaderLABEL
        '
        Me.YesNoHeaderLABEL.BackColor = System.Drawing.Color.Black
        Me.YesNoHeaderLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YesNoHeaderLABEL.ForeColor = System.Drawing.Color.BurlyWood
        Me.YesNoHeaderLABEL.Location = New System.Drawing.Point(36, 38)
        Me.YesNoHeaderLABEL.Name = "YesNoHeaderLABEL"
        Me.YesNoHeaderLABEL.Size = New System.Drawing.Size(356, 23)
        Me.YesNoHeaderLABEL.TabIndex = 3
        Me.YesNoHeaderLABEL.Text = "YesNoHeaderLABEL - Set at runtime"
        Me.YesNoHeaderLABEL.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YesNoD2Style
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.d2graphicselect
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(431, 275)
        Me.ControlBox = False
        Me.Controls.Add(Me.YesNoHeaderLABEL)
        Me.Controls.Add(Me.YesNoMessageLABEL)
        Me.Controls.Add(Me.YesNoCONFIRMBUTTON)
        Me.Controls.Add(Me.YesNoCANCELBUTTON)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "YesNoD2Style"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Universal Confirmation Box - Set Text At Runtime"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents YesNoCANCELBUTTON As System.Windows.Forms.Button
    Friend WithEvents YesNoCONFIRMBUTTON As System.Windows.Forms.Button
    Friend WithEvents YesNoMessageLABEL As System.Windows.Forms.Label
    Friend WithEvents YesNoHeaderLABEL As System.Windows.Forms.Label
End Class
