<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Label3 = New System.Windows.Forms.Label()
        Me.AllItemsInDatabaseListBox = New System.Windows.Forms.ListBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.EditExistingItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SearchOperatorLABEL = New System.Windows.Forms.Label()
        Me.SearchOperatorCOMBOBOX = New System.Windows.Forms.ComboBox()
        Me.SearchValueNUMERICUPDWN = New System.Windows.Forms.NumericUpDown()
        Me.SearchValueLABEL = New System.Windows.Forms.Label()
        Me.SearchWordCOMBOBOX = New System.Windows.Forms.ComboBox()
        Me.SearchFieldCOMBOBOX = New System.Windows.Forms.ComboBox()
        Me.SearchBUTTON = New System.Windows.Forms.Button()
        Me.ItemsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.AddNewItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.DeleteItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SetDefaultsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CloseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.SearchLABEL = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.OpenDatabaseDIALOG = New System.Windows.Forms.OpenFileDialog()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.ListboxTABCONTROL = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.SearchLISTBOX = New System.Windows.Forms.ListBox()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.SearchValueNUMERICUPDWN, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ListboxTABCONTROL.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.Label3.Location = New System.Drawing.Point(466, 39)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 13)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Logging Status"
        '
        'AllItemsInDatabaseListBox
        '
        Me.AllItemsInDatabaseListBox.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.AllItemsInDatabaseListBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.AllItemsInDatabaseListBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AllItemsInDatabaseListBox.ForeColor = System.Drawing.SystemColors.MenuBar
        Me.AllItemsInDatabaseListBox.FormattingEnabled = True
        Me.AllItemsInDatabaseListBox.Location = New System.Drawing.Point(3, 3)
        Me.AllItemsInDatabaseListBox.Name = "AllItemsInDatabaseListBox"
        Me.AllItemsInDatabaseListBox.Size = New System.Drawing.Size(367, 417)
        Me.AllItemsInDatabaseListBox.TabIndex = 22
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ToolStripProgressBar1.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(150, 16)
        Me.ToolStripProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ToolStripStatusLabel1.ForeColor = System.Drawing.SystemColors.Control
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(129, 17)
        Me.ToolStripStatusLabel1.Text = "Time untill next update"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1, Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 720)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(794, 22)
        Me.StatusStrip1.TabIndex = 20
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(15, 17)
        Me.ToolStripStatusLabel2.Text = "<"
        '
        'EditExistingItemToolStripMenuItem
        '
        Me.EditExistingItemToolStripMenuItem.Name = "EditExistingItemToolStripMenuItem"
        Me.EditExistingItemToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.EditExistingItemToolStripMenuItem.Text = "Edit"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.Button2.Location = New System.Drawing.Point(232, 116)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 122
        Me.Button2.Text = "Re-Search"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'SearchOperatorLABEL
        '
        Me.SearchOperatorLABEL.AutoSize = True
        Me.SearchOperatorLABEL.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.SearchOperatorLABEL.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SearchOperatorLABEL.Location = New System.Drawing.Point(96, 101)
        Me.SearchOperatorLABEL.Name = "SearchOperatorLABEL"
        Me.SearchOperatorLABEL.Size = New System.Drawing.Size(48, 13)
        Me.SearchOperatorLABEL.TabIndex = 121
        Me.SearchOperatorLABEL.Text = "Operator"
        '
        'SearchOperatorCOMBOBOX
        '
        Me.SearchOperatorCOMBOBOX.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SearchOperatorCOMBOBOX.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.SearchOperatorCOMBOBOX.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SearchOperatorCOMBOBOX.FormattingEnabled = True
        Me.SearchOperatorCOMBOBOX.Items.AddRange(New Object() {"Equal To", "Not Equal To", "Greater Than", "Less Than"})
        Me.SearchOperatorCOMBOBOX.Location = New System.Drawing.Point(99, 118)
        Me.SearchOperatorCOMBOBOX.Name = "SearchOperatorCOMBOBOX"
        Me.SearchOperatorCOMBOBOX.Size = New System.Drawing.Size(115, 21)
        Me.SearchOperatorCOMBOBOX.TabIndex = 120
        Me.SearchOperatorCOMBOBOX.Text = "Equal To"
        '
        'SearchValueNUMERICUPDWN
        '
        Me.SearchValueNUMERICUPDWN.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SearchValueNUMERICUPDWN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SearchValueNUMERICUPDWN.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SearchValueNUMERICUPDWN.Location = New System.Drawing.Point(30, 119)
        Me.SearchValueNUMERICUPDWN.Name = "SearchValueNUMERICUPDWN"
        Me.SearchValueNUMERICUPDWN.Size = New System.Drawing.Size(56, 20)
        Me.SearchValueNUMERICUPDWN.TabIndex = 118
        '
        'SearchValueLABEL
        '
        Me.SearchValueLABEL.AutoSize = True
        Me.SearchValueLABEL.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.SearchValueLABEL.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SearchValueLABEL.Location = New System.Drawing.Point(26, 101)
        Me.SearchValueLABEL.Name = "SearchValueLABEL"
        Me.SearchValueLABEL.Size = New System.Drawing.Size(34, 13)
        Me.SearchValueLABEL.TabIndex = 119
        Me.SearchValueLABEL.Text = "Value"
        '
        'SearchWordCOMBOBOX
        '
        Me.SearchWordCOMBOBOX.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SearchWordCOMBOBOX.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.SearchWordCOMBOBOX.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SearchWordCOMBOBOX.FormattingEnabled = True
        Me.SearchWordCOMBOBOX.Items.AddRange(New Object() {"Unique", "Set", "RuneWord"})
        Me.SearchWordCOMBOBOX.Location = New System.Drawing.Point(179, 56)
        Me.SearchWordCOMBOBOX.Name = "SearchWordCOMBOBOX"
        Me.SearchWordCOMBOBOX.Size = New System.Drawing.Size(221, 21)
        Me.SearchWordCOMBOBOX.TabIndex = 5
        '
        'SearchFieldCOMBOBOX
        '
        Me.SearchFieldCOMBOBOX.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SearchFieldCOMBOBOX.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.SearchFieldCOMBOBOX.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SearchFieldCOMBOBOX.FormattingEnabled = True
        Me.SearchFieldCOMBOBOX.Items.AddRange(New Object() {"Item Name", "Item Base", "Item Quality", "Item Defense", "RuneWord", "Chance To Block", "One Hand Damage Max", "One Hand Damage Min", "Two Hand Damage Max", "Two Hand Damage Min", "Throw Damage Max", "Throw Damage Min", "Required Level", "Required Strength", "Required Dexterity", "Attack Class", "Attack Speed", "Unique Attributes", "Mule Name", "Mule Account", "Pickit Bot", "Pickit Area", "User Reference"})
        Me.SearchFieldCOMBOBOX.Location = New System.Drawing.Point(29, 56)
        Me.SearchFieldCOMBOBOX.Name = "SearchFieldCOMBOBOX"
        Me.SearchFieldCOMBOBOX.Size = New System.Drawing.Size(140, 21)
        Me.SearchFieldCOMBOBOX.TabIndex = 3
        Me.SearchFieldCOMBOBOX.Text = "Item Name"
        '
        'SearchBUTTON
        '
        Me.SearchBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SearchBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.SearchBUTTON.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SearchBUTTON.Location = New System.Drawing.Point(325, 116)
        Me.SearchBUTTON.Name = "SearchBUTTON"
        Me.SearchBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.SearchBUTTON.TabIndex = 1
        Me.SearchBUTTON.Text = "Search"
        Me.SearchBUTTON.UseVisualStyleBackColor = False
        '
        'ItemsToolStripMenuItem
        '
        Me.ItemsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem2, Me.ToolStripSeparator2, Me.AddNewItemToolStripMenuItem, Me.EditExistingItemToolStripMenuItem, Me.ToolStripSeparator1, Me.DeleteItemToolStripMenuItem})
        Me.ItemsToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.ItemsToolStripMenuItem.Name = "ItemsToolStripMenuItem"
        Me.ItemsToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.ItemsToolStripMenuItem.Text = "Items"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(196, 22)
        Me.ToolStripMenuItem2.Text = "Import Mule Logs Now"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(193, 6)
        '
        'AddNewItemToolStripMenuItem
        '
        Me.AddNewItemToolStripMenuItem.Name = "AddNewItemToolStripMenuItem"
        Me.AddNewItemToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.AddNewItemToolStripMenuItem.Text = "Add"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(193, 6)
        '
        'DeleteItemToolStripMenuItem
        '
        Me.DeleteItemToolStripMenuItem.Name = "DeleteItemToolStripMenuItem"
        Me.DeleteItemToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.DeleteItemToolStripMenuItem.Text = "Delete"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.Label2.Location = New System.Drawing.Point(605, 186)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Item Details"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.Label1.Location = New System.Drawing.Point(64, 186)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 13)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Data Base List"
        '
        'SetDefaultsToolStripMenuItem
        '
        Me.SetDefaultsToolStripMenuItem.Name = "SetDefaultsToolStripMenuItem"
        Me.SetDefaultsToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.SetDefaultsToolStripMenuItem.Text = "Set Defaults"
        '
        'SettingsToolStripMenuItem
        '
        Me.SettingsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SetDefaultsToolStripMenuItem})
        Me.SettingsToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SettingsToolStripMenuItem.Name = "SettingsToolStripMenuItem"
        Me.SettingsToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.SettingsToolStripMenuItem.Text = "Settings"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(139, 22)
        Me.ToolStripMenuItem1.Text = "Backup Files"
        '
        'CloseToolStripMenuItem
        '
        Me.CloseToolStripMenuItem.Name = "CloseToolStripMenuItem"
        Me.CloseToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.CloseToolStripMenuItem.Text = "Save"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.OpenToolStripMenuItem.Text = "Open"
        '
        'ItemFileToolStripMenuItem
        '
        Me.ItemFileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenToolStripMenuItem, Me.CloseToolStripMenuItem, Me.NewToolStripMenuItem1, Me.ToolStripMenuItem1, Me.ExitToolStripMenuItem})
        Me.ItemFileToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.ItemFileToolStripMenuItem.Name = "ItemFileToolStripMenuItem"
        Me.ItemFileToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.ItemFileToolStripMenuItem.Text = "Database"
        '
        'NewToolStripMenuItem1
        '
        Me.NewToolStripMenuItem1.Name = "NewToolStripMenuItem1"
        Me.NewToolStripMenuItem1.Size = New System.Drawing.Size(139, 22)
        Me.NewToolStripMenuItem1.Text = "New"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ItemFileToolStripMenuItem, Me.ItemsToolStripMenuItem, Me.SettingsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(794, 24)
        Me.MenuStrip1.TabIndex = 14
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox2.BackColor = System.Drawing.SystemColors.MenuText
        Me.RichTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox2.CausesValidation = False
        Me.RichTextBox2.ForeColor = System.Drawing.SystemColors.Highlight
        Me.RichTextBox2.Location = New System.Drawing.Point(473, 358)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.ReadOnly = True
        Me.RichTextBox2.Size = New System.Drawing.Size(270, 347)
        Me.RichTextBox2.TabIndex = 15
        Me.RichTextBox2.Text = ""
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.ForeColor = System.Drawing.SystemColors.MenuBar
        Me.TextBox2.Location = New System.Drawing.Point(189, 186)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(81, 13)
        Me.TextBox2.TabIndex = 25
        '
        'SearchLABEL
        '
        Me.SearchLABEL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SearchLABEL.AutoSize = True
        Me.SearchLABEL.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.SearchLABEL.ForeColor = System.Drawing.SystemColors.ControlLight
        Me.SearchLABEL.Location = New System.Drawing.Point(26, 39)
        Me.SearchLABEL.Name = "SearchLABEL"
        Me.SearchLABEL.Size = New System.Drawing.Size(64, 13)
        Me.SearchLABEL.TabIndex = 26
        Me.SearchLABEL.Text = "Item Search"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.ControlText
        Me.PictureBox1.BackgroundImage = Global.D2Items.My.Resources.Resources.ImageBackground
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox1.Location = New System.Drawing.Point(560, 224)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Padding = New System.Windows.Forms.Padding(8, 6, 0, 0)
        Me.PictureBox1.Size = New System.Drawing.Size(103, 128)
        Me.PictureBox1.TabIndex = 16
        Me.PictureBox1.TabStop = False
        '
        'OpenDatabaseDIALOG
        '
        Me.OpenDatabaseDIALOG.FileName = "OpenFileDialog1"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.ForeColor = System.Drawing.SystemColors.MenuBar
        Me.RichTextBox1.Location = New System.Drawing.Point(469, 55)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RichTextBox1.Size = New System.Drawing.Size(298, 95)
        Me.RichTextBox1.TabIndex = 27
        Me.RichTextBox1.Text = ""
        '
        'ListboxTABCONTROL
        '
        Me.ListboxTABCONTROL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.ListboxTABCONTROL.Controls.Add(Me.TabPage1)
        Me.ListboxTABCONTROL.Controls.Add(Me.TabPage2)
        Me.ListboxTABCONTROL.Location = New System.Drawing.Point(26, 247)
        Me.ListboxTABCONTROL.Name = "ListboxTABCONTROL"
        Me.ListboxTABCONTROL.SelectedIndex = 0
        Me.ListboxTABCONTROL.Size = New System.Drawing.Size(381, 449)
        Me.ListboxTABCONTROL.TabIndex = 123
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.AllItemsInDatabaseListBox)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(373, 423)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Item List"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.SearchLISTBOX)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(373, 423)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Search List"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'SearchLISTBOX
        '
        Me.SearchLISTBOX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SearchLISTBOX.FormattingEnabled = True
        Me.SearchLISTBOX.Location = New System.Drawing.Point(3, 3)
        Me.SearchLISTBOX.Name = "SearchLISTBOX"
        Me.SearchLISTBOX.Size = New System.Drawing.Size(367, 417)
        Me.SearchLISTBOX.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.D2Items.My.Resources.Resources.D2Data
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(794, 742)
        Me.Controls.Add(Me.ListboxTABCONTROL)
        Me.Controls.Add(Me.SearchValueLABEL)
        Me.Controls.Add(Me.SearchOperatorLABEL)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.SearchValueNUMERICUPDWN)
        Me.Controls.Add(Me.SearchWordCOMBOBOX)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.SearchLABEL)
        Me.Controls.Add(Me.SearchFieldCOMBOBOX)
        Me.Controls.Add(Me.SearchOperatorCOMBOBOX)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.SearchBUTTON)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.RichTextBox2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(810, 800)
        Me.MinimumSize = New System.Drawing.Size(810, 736)
        Me.Name = "Form1"
        Me.Text = "DiaBase"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.SearchValueNUMERICUPDWN, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ListboxTABCONTROL.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents AllItemsInDatabaseListBox As System.Windows.Forms.ListBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents EditExistingItemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddNewItemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents SetDefaultsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SettingsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CloseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents DeleteItemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SearchBUTTON As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents SearchOperatorLABEL As System.Windows.Forms.Label
    Friend WithEvents SearchOperatorCOMBOBOX As System.Windows.Forms.ComboBox
    Friend WithEvents SearchValueNUMERICUPDWN As System.Windows.Forms.NumericUpDown
    Friend WithEvents SearchValueLABEL As System.Windows.Forms.Label
    Friend WithEvents SearchWordCOMBOBOX As System.Windows.Forms.ComboBox
    Friend WithEvents SearchFieldCOMBOBOX As System.Windows.Forms.ComboBox
    Friend WithEvents SearchLABEL As System.Windows.Forms.Label
    Friend WithEvents OpenDatabaseDIALOG As System.Windows.Forms.OpenFileDialog
    Friend WithEvents NewToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListboxTABCONTROL As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents SearchLISTBOX As System.Windows.Forms.ListBox


End Class
