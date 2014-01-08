
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load 'check if required folders exist on startup and create if necessary
        If My.Computer.FileSystem.DirectoryExists(Application.StartupPath + "\DataBase\Backup") = False Then
            My.Computer.FileSystem.CreateDirectory(Application.StartupPath + "\DataBase\Backup")
            Mymessages = "Database & Backup Folder created" : MyMessageBox()
        End If

        If My.Computer.FileSystem.DirectoryExists(Application.StartupPath + "\Archive") = False Then
            My.Computer.FileSystem.CreateDirectory(Application.StartupPath + "\Archive")
            Mymessages = "Archive Folder created" : MyMessageBox()
        End If
        If My.Computer.FileSystem.FileExists(Application.StartupPath + "\DataBase\Default.txt") = False Then
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath + "\DataBase\Default.txt", False)
            file.Close()
            Mymessages = "Default Data base file created" : MyMessageBox()

        End If
        If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Settings.cfg") = False Then
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath + "\Settings.cfg", False)
            file.WriteLine("C\D2NT")
            file.WriteLine(Application.StartupPath + "\DataBase\Default.txt")
            file.WriteLine("30")
            file.WriteLine("True")
            file.WriteLine("False")
            file.Close()
            Mymessages = "Settings file created" : MyMessageBox()

        End If
    End Sub
    'Stuff to run after program starts
    Private Sub Form1_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        LoadConfigFile()

        OpenDatabaseRoutine(DataBaseFile)

        StartTimer()
        If LoggerRunning = False Then RichTextBox1.Text = "AutoLogging is Idle" & vbCrLf
        If My.Computer.FileSystem.DirectoryExists(EtalPath) = False Then
            Dim MySettings As New Settings
            MySettings.Show()
        End If
    End Sub
    Public WithEvents ImportTimer As New System.Windows.Forms.Timer()

    Sub StartTimer()
        Timercount = 0
        TimerSecs = TimerMins * 60
        ToolStripProgressBar1.Maximum = 100
        ImportTimer.Interval = 1000
        ImportTimer.Start()

    End Sub

    Private Sub TimerEventProcessor(ByVal myObject As Object, ByVal MyEventArgs As EventArgs) Handles ImportTimer.Tick

        Timercount = Timercount + 1
        If Timercount > TimerSecs Then
            Timercount = 0
            ImportTimer.Stop()
            RichTextBox1.Text = "Checking for new Log Files"
            LoggerRunning = True
            ImportLogFiles()
            LoggerRunning = False
            StartTimer()
            RichTextBox1.AppendText("AutoLogging is Idle" & vbCrLf)
        End If

        Dim Timerprogress As Integer = Math.Round((Timercount / TimerSecs) * 100)
        ToolStripProgressBar1.Value = Timerprogress
        ToolStripStatusLabel2.Text = "< " & Math.Ceiling((TimerSecs - Timercount) / 60) & " mins"
    End Sub

    'Open New Database Button Handler
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        If LoggerRunning = True Then ' dont try to open file if auto logger is working - will cause app crash
            Mymessages = "Please wait File in use, Logger Running" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop() '        stop timer b4 form opens
        OpenDatabaseDIALOG.Title = "Open Existing Database" '                           set dialog title
        OpenDatabaseDIALOG.InitialDirectory = Application.StartupPath + "\DataBase\" '  set initial dir
        OpenDatabaseDIALOG.Filter = ".txt|*.txt"
        OpenDatabaseDIALOG.FileName = "Default.txt"

        If OpenDatabaseDIALOG.ShowDialog() = DialogResult.Cancel Then
            If Button3.Text = "Timer Stop" Then ImportTimer.Start() '       restart timer if not paused
            Return '          without this if user clicks cancel app crashes
        End If

        SearchLISTBOX.Items.Clear() '                                                   clean out old search matches
        PictureBox1.BackgroundImage = D2Items.My.Resources.Resources.ImageBackground '  clean out old image                 [NOT WORKING]
        RichTextBox2.Text = "" '                                                        clean out old item stats
        RichTextBox1.Text = "" '                                                        clean out logging window

        Dim DatabaseFilename As String = OpenDatabaseDIALOG.FileName

        OpenDatabaseRoutine(DatabaseFilename) ' Routine puts saved items ito object arrays as ItemObject class collection

   
        Display_Items() '                       Routine Populates all items listbox with, um, all items obviously :)
        If Button3.Text = "Timer Stop" Then ImportTimer.Start() '       restart timer if not paused
    End Sub

    'Populates all items listbox routine
    Public Sub Display_Items()
        AllItemsInDatabaseListBox.Items.Clear() 'insure old list is definetly clear b4 repoulating
        For x = 0 To Objects.Count - 1
            AllItemsInDatabaseListBox.Items.Add(Objects(x).ItemName)
        Next
        TextBox2.Text = Objects.Count & " Items"

        If AllItemsInDatabaseListBox.Items.Count > 0 Then AllItemsInDatabaseListBox.SelectedIndex = 0
    End Sub

    'OPENES THE NEW ITEM FORM TO MANUALLY ADD A NEW ITEM TO THE DATABASE
    Private Sub AddNewItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddNewItemToolStripMenuItem.Click
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If

        ImportTimer.Stop() '        stop timer b4 form opens
        AddItemForm.ShowDialog()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start() '       restart timer after form closes
    End Sub

    'OPENS THE EDIT EXISTING ITEM FORM TO MANUALLY EDIT AN ITEMS FIELDS
    Private Sub EditExistingItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditExistingItemToolStripMenuItem.Click
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()
        If AllItemsInDatabaseListBox.SelectedIndex > -1 Then EditItemForm.ShowDialog()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'dISPLAYS THE SETTINGS FOM SELECTED FROM PULLDOWN MENUES
    Private Sub SetDefaultsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetDefaultsToolStripMenuItem.Click
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()
        Settings.Show()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub


    Private Sub AllItemsInDatabaseListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AllItemsInDatabaseListBox.SelectedIndexChanged
        Dim RowNumber As Integer = AllItemsInDatabaseListBox.SelectedIndex
        If RowNumber = -1 Then Return ' do nothing
        If RowNumber >= 0 Then
            RichTextBox2.Text = ""  'clears form has been some overlap of listings occur - wierd behaviour

            Dim DisplayType As String = Objects(RowNumber).ItemQuality
            Dim count As Integer = Objects(RowNumber).ItemQuality.Length
            If DisplayType = "Normal" Or DisplayType = "Superior" Then
                If Objects(RowNumber).ItemBase = "Rune" Then
                    RichTextBox2.SelectionColor = Color.Orange
                    RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
                End If
                If Objects(RowNumber).ItemName.IndexOf("Rune") = -1 Then
                    RichTextBox2.SelectionColor = Color.White
                    RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
                End If
            End If
            If DisplayType = "Magic" Then
                RichTextBox2.SelectionColor = Color.DodgerBlue
                RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
            End If
            If DisplayType = "Rare" Then
                RichTextBox2.SelectionColor = Color.Goldenrod
                RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
            End If
            If DisplayType = "Set" Then
                RichTextBox2.SelectionColor = Color.Green
                RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
            End If
            If DisplayType = "Unique" Then
                RichTextBox2.SelectionColor = Color.BurlyWood
                RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
            End If


            count = RichTextBox2.TextLength
            'from this point we want to add white text for basic item info
            If Objects(RowNumber).Defense <> Nothing Then RichTextBox2.AppendText("Defense: " & Objects(RowNumber).Defense & vbCrLf)
            If Objects(RowNumber).ChanceToBlock <> Nothing Then RichTextBox2.AppendText("ChanceToBlock: " & Objects(RowNumber).ChanceToBlock & vbCrLf)
            If Objects(RowNumber).OneHandDamageMax <> Nothing Then RichTextBox2.AppendText("OneHandDamage: " & Objects(RowNumber).OneHandDamageMin & " To " & Objects(RowNumber).OneHandDamageMax & vbCrLf)
            If Objects(RowNumber).TwoHandDamageMax <> Nothing Then RichTextBox2.AppendText("TwoHandDamage: " & Objects(RowNumber).TwoHandDamageMin & " To " & Objects(RowNumber).TwoHandDamageMax & vbCrLf)
            If Objects(RowNumber).DurabilityMin <> Nothing Then RichTextBox2.AppendText("Durability: " & Objects(RowNumber).DurabilityMin & " of " & Objects(RowNumber).DurabilityMax & vbCrLf)
            If Objects(RowNumber).RequiredStrength <> Nothing Then RichTextBox2.AppendText("Required Strength: " & Objects(RowNumber).RequiredStrength & vbCrLf)
            If Objects(RowNumber).RequiredDexterity <> Nothing Then RichTextBox2.AppendText("Required Dexterity: " & Objects(RowNumber).RequiredDexterity & vbCrLf)
            If Objects(RowNumber).RequiredLevel <> Nothing Then RichTextBox2.AppendText("Required Level: " & Objects(RowNumber).RequiredLevel & vbCrLf)

            Dim count2 As Integer = RichTextBox2.TextLength - count
            RichTextBox2.Select(count, count2)
            RichTextBox2.SelectionColor = Color.White
            'from this point we dont need to add colors - Default Blue
            If Objects(RowNumber).Stat1 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat1 & vbCrLf)
            If Objects(RowNumber).Stat2 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat2 & vbCrLf)
            If Objects(RowNumber).Stat3 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat3 & vbCrLf)
            If Objects(RowNumber).Stat4 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat4 & vbCrLf)
            If Objects(RowNumber).Stat5 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat5 & vbCrLf)
            If Objects(RowNumber).Stat6 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat6 & vbCrLf)
            If Objects(RowNumber).Stat7 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat7 & vbCrLf)
            If Objects(RowNumber).Stat8 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat8 & vbCrLf)
            If Objects(RowNumber).Stat9 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat9 & vbCrLf)
            If Objects(RowNumber).Stat10 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat10 & vbCrLf)
            If Objects(RowNumber).Stat11 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat11 & vbCrLf)
            If Objects(RowNumber).Stat12 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat12 & vbCrLf)
            If Objects(RowNumber).Stat13 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat13 & vbCrLf)
            If Objects(RowNumber).Stat14 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat14 & vbCrLf)
            If Objects(RowNumber).Stat15 <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).Stat15 & vbCrLf)


            RichTextBox2.SelectionStart = RichTextBox2.TextLength 'ensure we are at end of document
            count = RichTextBox2.GetLineFromCharIndex(RichTextBox2.TextLength)
            For x = 0 To 20 - count
                RichTextBox2.AppendText(vbCrLf) ' a few line feeds so we can display muleaccount etc info near bottom
            Next

            RichTextBox2.AppendText("Mule Account: " & Objects(RowNumber).MuleAccount & vbCrLf)
            RichTextBox2.AppendText("Mule Name: " & Objects(RowNumber).MuleName & vbCrLf)
            RichTextBox2.AppendText("Mule Pass: " & Objects(RowNumber).MulePass & vbCrLf)

        End If
        RichTextBox2.SelectAll()
        RichTextBox2.SelectionAlignment = HorizontalAlignment.Center

        PictureBox1.Load("Skins\" + ItemImageList(Objects(RowNumber).ItemImage) + ".jpg")


    End Sub



    Private Sub DeleteItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteItemToolStripMenuItem.Click
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        Dim RowNumber As Integer = AllItemsInDatabaseListBox.SelectedIndex
        If RowNumber = -1 Then Return ' do nothing
        If RowNumber >= 0 Then
            Objects.RemoveAt(RowNumber)
            AllItemsInDatabaseListBox.Items.Clear()
            Display_Items()
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        RichTextBox1.Text = "Checking for new Logs" & vbCrLf
        ImportLogFiles()
        RichTextBox1.AppendText("AutoLogging is Idle")
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        If Objects.Count < 1 Then Return ' nothing to save
        Dim result = MessageBox.Show("Are you sure?", "Save File", MessageBoxButtons.YesNo)
        If result = Windows.Forms.DialogResult.No Then Return
        ImportTimer.Stop()
        SaveItems()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub
    Private Sub SaveItems()

        Try

            Dim LogWriter = My.Computer.FileSystem.OpenTextFileWriter(DataBaseFile, False)

            For x = 0 To Objects.Count - 1

                LogWriter.WriteLine("--------------------")
                LogWriter.WriteLine(Objects(x).ItemName)
                LogWriter.WriteLine(Objects(x).ItemBase)
                LogWriter.WriteLine(Objects(x).ItemQuality)
                LogWriter.WriteLine(Objects(x).RequiredCharacter)
                LogWriter.WriteLine(Objects(x).EtherealItem)
                LogWriter.WriteLine(Objects(x).Sockets)
                LogWriter.WriteLine(Objects(x).RuneWord)
                LogWriter.WriteLine(Objects(x).ThrowDamageMin)
                LogWriter.WriteLine(Objects(x).ThrowDamageMax)
                LogWriter.WriteLine(Objects(x).OneHandDamageMin)
                LogWriter.WriteLine(Objects(x).OneHandDamageMax)
                LogWriter.WriteLine(Objects(x).TwoHandDamageMin)
                LogWriter.WriteLine(Objects(x).TwoHandDamageMax)
                LogWriter.WriteLine(Objects(x).Defense)
                LogWriter.WriteLine(Objects(x).ChanceToBlock)
                LogWriter.WriteLine(Objects(x).QuantityMin)
                LogWriter.WriteLine(Objects(x).QuantityMax)
                LogWriter.WriteLine(Objects(x).DurabilityMin)
                LogWriter.WriteLine(Objects(x).DurabilityMax)
                LogWriter.WriteLine(Objects(x).RequiredStrength)
                LogWriter.WriteLine(Objects(x).RequiredDexterity)
                LogWriter.WriteLine(Objects(x).RequiredLevel)
                LogWriter.WriteLine(Objects(x).AttackClass)
                LogWriter.WriteLine(Objects(x).AttackSpeed)
                LogWriter.WriteLine(Objects(x).Stat1)
                LogWriter.WriteLine(Objects(x).Stat2)
                LogWriter.WriteLine(Objects(x).Stat3)
                LogWriter.WriteLine(Objects(x).Stat4)
                LogWriter.WriteLine(Objects(x).Stat5)
                LogWriter.WriteLine(Objects(x).Stat6)
                LogWriter.WriteLine(Objects(x).Stat7)
                LogWriter.WriteLine(Objects(x).Stat8)
                LogWriter.WriteLine(Objects(x).Stat9)
                LogWriter.WriteLine(Objects(x).Stat10)
                LogWriter.WriteLine(Objects(x).Stat11)
                LogWriter.WriteLine(Objects(x).Stat12)
                LogWriter.WriteLine(Objects(x).Stat13)
                LogWriter.WriteLine(Objects(x).Stat14)
                LogWriter.WriteLine(Objects(x).Stat15)
                LogWriter.WriteLine(Objects(x).MuleName)
                LogWriter.WriteLine(Objects(x).MuleAccount)
                LogWriter.WriteLine(Objects(x).MulePass)
                LogWriter.WriteLine(Objects(x).PickitBot)
                LogWriter.WriteLine(Objects(x).UserReference)
                LogWriter.WriteLine(Objects(x).ItemImage)

            Next

            LogWriter.Close()

        Catch ex As Exception
            Mymessages = "File Write Error" : MyMessageBox()
        End Try

    End Sub


    'Refreshes selected database when selected from menu bar
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click

        Dim BackupPath = Application.StartupPath & "\Database\Backup\"
        Dim temp As String = ""
        Dim myarray = Split(DataBaseFile, ".txt", 0)
        Dim tempname = myarray(0) & ".bak"
        myarray = Split(tempname, "\")
        tempname = myarray(myarray.Length - 1)

        If My.Computer.FileSystem.FileExists(BackupPath & tempname) = True Then
            My.Computer.FileSystem.DeleteFile(BackupPath & tempname)
        End If
        My.Computer.FileSystem.CopyFile(DataBasePath & DataBaseFile, BackupPath & tempname)
    End Sub



    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        Help.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        ImportTimer.Stop()
        ItemImageSelector.ShowDialog()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    Private Sub SearchBUTTON_Click(sender As Object, e As EventArgs) Handles SearchBUTTON.Click

        SearchRoutine()

    End Sub


    Private Sub SearchFieldCOMBOBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SearchFieldCOMBOBOX.SelectedIndexChanged

        'turns off string search textbox for value only searches
        SearchWordCOMBOBOX.Enabled = True
        If UCase(SearchFieldCOMBOBOX.Text) = "ONE HAND DAMAGE MAX" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "ONE HAND DAMAGE MIN" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "TWO HAND DAMAGE MAX" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "TWO HAND DAMAGE MIN" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "THROW  DAMAGE MAX" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "THROW  DAMAGE MIN" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "REQUIRED STRENGTH" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "REQUIRED DEXTERITY" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "DEFENSE" Then SearchWordCOMBOBOX.Enabled = False
        If UCase(SearchFieldCOMBOBOX.Text) = "CHANCE TO BLOCK" Then SearchWordCOMBOBOX.Enabled = False


        'ADD ALL ITEM NAMES SEARCHED SINCE STARTUP TO WORD DROP DOWN WHEN ITEM NAME IS SELECTED

        If UCase(SearchFieldCOMBOBOX.Text) = "ITEM NAME" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each item In ItemNamePulldownList
                SearchWordCOMBOBOX.Items.Add(item)
            Next
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL ITEM BASE ENTRYS WHEN ITEM BASE IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "ITEM BASE" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.ItemBase) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.ItemBase)
            Next
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL ITEM QUALITY ENTRYS WHEN ITEM QUALITY IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "ITEM QUALITY" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.ItemQuality) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.ItemQuality)
            Next
        End If





    End Sub


    Private Sub SearchLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SearchLISTBOX.SelectedIndexChanged

        AllItemsInDatabaseListBox.SelectedIndex = SearchReferenceList(SearchLISTBOX.SelectedIndex)
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If Button3.Text = "Timer Stop" Then
            ImportTimer.Stop()
            Button3.Text = "Timer Start"
            RichTextBox1.AppendText("Timer Stopped" & vbCrLf)
            Return
        End If
        If Button3.Text = "Timer Start" Then
            ImportTimer.Start()
            Button3.Text = "Timer Stop"
            RichTextBox1.AppendText("Timer Running" & vbCrLf)
        End If
    End Sub
End Class
