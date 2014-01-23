Imports System.IO

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load 'check if required folders exist on startup and create if necessary
        My.Computer.Audio.Play(My.Resources.BigDLaugh, AudioPlayMode.Background)


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

        SearchFieldCOMBOBOX.Text = "Item Name"
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
            RichTextBox1.Text = "Checking for new Log Files" & vbCrLf
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
        PictureBox1.BackgroundImage = DiaBase.My.Resources.Resources.ImageBackground '  clean out old image                 [NOT WORKING]
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
        ListboxTABCONTROL.SelectTab(0) ' ensure listboxTABCONTROL itset to all items list for repopuleate
        TextBox2.Text = Objects.Count & " - Total Items"

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


            ' MuleInfoRICHTEXTBOX.Clear() 'clearc mule info textbox
            MuleAccountTextbox.Clear()
            MuleNameTextbox.Clear()
            MulePassTextbox.Clear()


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
            RichTextBox2.AppendText(vbCrLf) ' add a spacing line between item name and item stats (looks neater) ROBS EDIT

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
            

            'ROBS EDIT - includes attack class in the main stat display  as opposed to the unique attibutes block?
            If Objects(RowNumber).AttackClass <> Nothing Then RichTextBox2.AppendText(Objects(RowNumber).AttackClass & " Class") : If Objects(RowNumber).AttackSpeed <> Nothing Then RichTextBox2.AppendText(" - " & Objects(RowNumber).AttackSpeed & vbCrLf) Else RichTextBox2.AppendText(vbCrLf)


            RichTextBox2.AppendText(vbCrLf) ' add a spacing line between item stats and unique attributes (looks neater) ROBS EDIT


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


            MuleAccountTextbox.Text = Objects(RowNumber).MuleAccount
            MuleNameTextbox.Text = Objects(RowNumber).MuleName
            MulePassTextbox.Text = Objects(RowNumber).MulePass


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
        ImportTimer.Stop()
        ImportLogFiles()
        RichTextBox1.AppendText("AutoLogging is Idle")
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        If Objects.Count < 1 Then Return ' nothing to save
        YesNoD2Style.Text = "Save Database"
        YesNoD2Style.YesNoHeaderLABEL.Text = ""
        YesNoD2Style.YesNoHeaderLABEL.TextAlign = ContentAlignment.MiddleCenter
        YesNoD2Style.YesNoMessageLABEL.Text = "This will Overwrite the Database" & vbCrLf & vbCrLf & vbCrLf & "Are you sure you?"
        YesNoD2Style.ShowDialog()
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.No Then Return

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
        YesNoD2Style.Text = "Confirm Database Backup"
        YesNoD2Style.YesNoHeaderLABEL.Text = "PLEASE READ BEFORE REFRESHING BACKUP"
        YesNoD2Style.YesNoMessageLABEL.Text = "This will permantly destroy the current database backup file. Doing this will mean you will no longer be able to resore the current backup file to a complete database." & vbCrLf & "Are you sure you wish to continue?"
        YesNoD2Style.ShowDialog()
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.Yes Then BackupDatabase()
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

        'This next 'if then' checks that the selected search operator is valid if its not set to default equal to, in upper case the emphasise the fix to the user

        'NOTED: Checks for the other comboboxes entry vaildity are IN EACH ENTRY CHANGED EVENT HANDLER. The Search button will remain disabled until the 
        'minimum ammount of search criteria has been entered. This insures the search routine wont crash due to user input error (i hope)

        If UCase(SearchOperatorCOMBOBOX.Text) = "EQUAL TO" Or UCase(SearchOperatorCOMBOBOX.Text) = "NOT EQUAL TO" Or UCase(SearchOperatorCOMBOBOX.Text) = "GREATER THAN" Or UCase(SearchOperatorCOMBOBOX.Text) = "LESS THAN" Then
            SearchRoutine() ' If all is good and search seems valid then branch to the search routine sub in Module1
        End If
        If SearchLISTBOX.Items.Count > 0 Then SearchLISTBOX.Focus() Else AllItemsInDatabaseListBox.Focus() ' ensure focus is returned to the relevant item list after search

    End Sub


    Private Sub SearchFieldCOMBOBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SearchFieldCOMBOBOX.SelectedIndexChanged
        'QUICKLY SETUP SOME SEARCH STUFF FIRST 

        'THIS ADDS ALL ITEM NAME SEARCHED AND MATCHED THIS SESSION (SAVED IN ItemNamePullDownList) TO THE SEARCH STRING  
        'COMBOBOX DROPDOWN. THIS SYSTEM IS ONLY USED FOR THE  ITEM NAME FIELD... 
        'As There is no point adding the entire Item Name list to the dropdown this wil be a compromise of sorts
        'The reference array is not saved so a new list is created each runtime session (could add save to file option as improvment idea l8r)

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

        'POPULATE WORD SEARCH DROPDOWN WITH ALL USER REFERENCE ENTRYS WHEN USER REF IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "USER REFERENCE" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.UserReference) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.UserReference)
            Next
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL MULE ACCOUNT ENTRYS WHEN MULE ACCOUNT IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "MULE ACCOUNT" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.MuleAccount) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.MuleAccount)
            Next
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL MULE NAME ENTRYS WHEN MULE NAME IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "MULE NAME" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.MuleName) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.MuleName)
            Next
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL MULE PASS ENTRYS WHEN MULE PASS IS SELECTED FOR SEARCH
        'BUT ONLY WHEN HIDE PASWORDS CHECKBOX IN SETTINGS IS UNCHECKED. A BLANK DROPDOWN IS RETURNED IF HIDE PASSWORDS IS CHECKED
        If UCase(SearchFieldCOMBOBOX.Text) = "MULE PASS" Then
            SearchWordCOMBOBOX.Items.Clear()
            If Settings.CheckBox3.Checked = False Then
                For Each ItemObjectItem As ItemObjects In Objects
                    If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.MulePass) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.MulePass)
                Next
            End If
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL ATTACK CLASS ENTRYS WHEN ATTACK CLASS IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "ATTACK CLASS" Then
            SearchWordCOMBOBOX.Items.Clear()
            If Settings.CheckBox3.Checked = False Then
                For Each ItemObjectItem As ItemObjects In Objects
                    If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.AttackClass) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.AttackClass)
                Next
            End If
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL ATTACK SPEED ENTRYS WHEN ATTACK SPEED IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "ATTACK SPEED" Then
            SearchWordCOMBOBOX.Items.Clear()
            If Settings.CheckBox3.Checked = False Then
                For Each ItemObjectItem As ItemObjects In Objects
                    If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.AttackSpeed) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.AttackSpeed)
                Next
            End If
        End If

        'ENTER SEARCH CRITERIA AND READY TO SEARCH!!!!!!!!!!!!!!!!!!

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

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        My.Computer.Audio.Play(My.Resources.diablotaunt1, AudioPlayMode.Background)
        PictureBox2.Visible = False
    End Sub

    'This alters the search button from search to refine depending on the chercked state of the refine checkbox
    Private Sub RefineSearchCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles RefineSearchCHECKBOX.CheckedChanged
        If RefineSearchCHECKBOX.Checked = True Then SearchBUTTON.Text = "Refine Now"
        If RefineSearchCHECKBOX.Checked = False Then SearchBUTTON.Text = "Search Now"

    End Sub

    'this button selects tab page 0 and colors button to show all items listbox on page 0 and displays total items value to textbox2
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListboxTABCONTROL.SelectTab(0)
        Button2.BackColor = Color.DimGray
        Button1.BackColor = Color.Black
        TextBox2.Text = AllItemsInDatabaseListBox.Items.Count & " - Total Items"


    End Sub

    'this button selects tab page 1 and colors button to show all search listbox on page 1 and displays total search matches value to textbox2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListboxTABCONTROL.SelectTab(1)
        Button1.BackColor = Color.DimGray
        Button2.BackColor = Color.Black
        TextBox2.Text = SearchLISTBOX.Items.Count & " - Total Matches"

    End Sub


    'items and matches nuber textbox
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub RestoreBackupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestoreBackupToolStripMenuItem.Click
        'SET D2 DIALOG TITLE AND MESSAGES
        YesNoD2Style.Text = "Restore Database From Backup"
        YesNoD2Style.YesNoHeaderLABEL.Text = "PLEASE READ BEFORE RESTORING BACKUP"
        YesNoD2Style.YesNoHeaderLABEL.TextAlign = ContentAlignment.MiddleCenter
        YesNoD2Style.YesNoMessageLABEL.Text = "You are about to permanently delete the current database and replace it with its backup file (if one exists)." & vbCrLf & vbCrLf & "Only continue if you are sure you know what you are doing as severe data loss may occur if you dont backup regularly." & vbCrLf & "Most record issues can be fixed manually by editing the faulty item record with the Item / Edit Item function." & vbCrLf & vbCrLf & "Are you sure you wish to restore the backup file?"
        YesNoD2Style.ShowDialog()

        'Restore form backup
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.Yes Then

            'ty for build backup filename routine Ned - im um borrowing it haha heehe!     
            ' MessageBox.Show("do it") '<----------------------------------------------------------------------------------ROBS DEBUG
            Dim BackupPath = Application.StartupPath & "\Database\Backup\"
            Dim temp As String = ""
            Dim myarray = Split(DatabaseFile, ".txt", 0)
            Dim tempname = myarray(0) & ".bak"
            myarray = Split(tempname, "\")
            tempname = myarray(myarray.Length - 1)

            If My.Computer.FileSystem.FileExists(BackupPath & tempname) = True Then

                'found a backup file and copying it over to replace current database file, then reload it
                My.Computer.FileSystem.DeleteFile(DatabaseFile) 'delete old dbase file
                My.Computer.FileSystem.CopyFile(BackupPath & tempname, DatabaseFile, True) ' copy over new dbase file and rename it

                OpenDatabaseRoutine(DatabaseFile) 'refresh new database file to the lists
            Else
                'There is not backup file for this database so cant restore it
            End If
        End If

        'cancel backup rtrstoration
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.No Then
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ClosingAppForm.ShowDialog()
        'this end without backing up or saving
        If ClosingAppForm.DialogResult = Windows.Forms.DialogResult.No Then e.Cancel = True
        If ClosingAppForm.DialogResult = Windows.Forms.DialogResult.OK And ClosingAppForm.SaveDatabaseCHECKBOX.Checked = False And ClosingAppForm.BackupDatabaseCHRCKBOX.Checked = False Then End
        If ClosingAppForm.DialogResult = Windows.Forms.DialogResult.OK And ClosingAppForm.SaveDatabaseCHECKBOX.Checked = True Then SaveItems()

        If ClosingAppForm.DialogResult = Windows.Forms.DialogResult.OK And ClosingAppForm.SaveDatabaseCHECKBOX.Checked = True Then BackupDatabase()


SkipExit:
        ClosingAppForm.Close()

    End Sub










  
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub NewToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem1.Click
        YesNoD2Style.Text = "Confirm New Database"
        YesNoD2Style.YesNoHeaderLABEL.Text = "PLEASE READ BEFOR CREATING A NEW DATABASE"
        YesNoD2Style.YesNoMessageLABEL.Text = "The Beat version does not support multiple databases. Creating a new database will destroy the current one." & vbCrLf & "Are you sure you want to contine?"
        YesNoD2Style.ShowDialog()
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.Yes Then

            If My.Computer.FileSystem.FileExists(Application.StartupPath + "\DataBase\Default.txt") = False Then

                My.Computer.FileSystem.DeleteFile(Application.StartupPath + "\DataBase\Default.txt")
            Else

                Dim file As System.IO.StreamWriter
                file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath + "\DataBase\Default.txt", False)
                file.Close()
                Mymessages = "Default Data base file created" : MyMessageBox()
                OpenDatabaseRoutine(Databasefile) 'refresh new database file to the lists


            End If

        End If


SkipNewDatabase:
    End Sub
End Class
