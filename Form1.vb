Imports System.IO
Imports System.Drawing.Text

Public Class Form1

    'FORM1 LOAD EVENT - PLAYS AUDIO LAUGH AND GETS FILE CONFIG VALUES AND SETS UP APPLICATION ELEMENTS AND OPENS THE DEFAULT DATABASE
    'ROUTINE ALSO SETS UP REQURED FILES ECT ON FIRST RUN (INSTALLS APP)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load 'check if required folders exist on startup and create if necessary
        My.Computer.Audio.Play(My.Resources.BigDLaugh, AudioPlayMode.Background)

        '------------------------------------------------------------------------------------------------------------------
        'Set Version And revision number in form 1 titlebar text - Version and Revision number set in module1 variable list
        Me.Text = "DiaBASE Beta - Version " & VersionNumber & "  Revision " & RevisionNumber
        '------------------------------------------------------------------------------------------------------------------

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
            file.WriteLine("False") 'added for backup before imports 
            file.WriteLine("False") 'added for backup before item edits 
            file.WriteLine("False") 'added for remove mule dupe
            file.Close()
            Mymessages = "Settings file created" : MyMessageBox()
        End If

        'Next bit setup up diablo 2 heading text and game text true type fonts (.ttf) 
        'Applying these fonts here and now will overwrite any value set in the from designer properties window
        'Setup pfc as our font collestion label, then assign the .ttf font fileas the font to use (should be in extras folder)

        'This adds Diablo2 Heading Text font as 0
        If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Extras\DiabloFont1.ttf") = True Then
            pfc.AddFontFile(Application.StartupPath + "\Extras\DiabloFont1.ttf")

            'Fancy Headers (16 point size)
            Label1.Font = New Font(pfc.Families(0), 16, FontStyle.Regular)  'Item Lists Header
            Label2.Font = New Font(pfc.Families(0), 16, FontStyle.Regular)  'Details Header
            Label3.Font = New Font(pfc.Families(0), 16, FontStyle.Regular)  'Autologging Header
            Label29.Font = New Font(pfc.Families(0), 16, FontStyle.Regular) 'Item and Mule Search

            'BUTTONS (9 point size)
            ListControlTabBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            SearchListControlTabBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            TradesListControlTabBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            SearchBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
               Button3.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)

            'DialogResult Form Buttons (9 point size)
            YesNoD2Style.YesNoCONFIRMBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            YesNoD2Style.YesNoCANCELBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ClosingAppForm.AppCloseCancelBUTTONButton1.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ClosingAppForm.AppCloseConfirmBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            CreateNewDatabase.NewDatabaseCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            CreateNewDatabase.NewDatabaseCreateBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            MoveItems.MoveItemsCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            MoveItems.MoveItemsExportBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            Settings.SaveDefaultsBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            Settings.CancelDefaultsBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            Settings.EtalPathBrowseBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            Settings.DefaultDatabaseBrowseBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            Settings.Label4.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            UserInputForm.UserInputNoBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            UserInputForm.UserInputYesBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            D2StyleOpenFileDialog.NewDatabaseCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            D2StyleOpenFileDialog.OpenDatabaseBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            UserMessaging.Button1.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)

            'DIALOG WINDOW HEADERS (12 point size)
            YesNoD2Style.YesNoHeaderLABEL.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            ClosingAppForm.Label1.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            CreateNewDatabase.Label1.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            D2StyleOpenFileDialog.Label2.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            MoveItems.Label1.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            Settings.Label1.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            Settings.Label3.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            UserInputForm.UserInputHeaderLABEL.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)

            'EDIT AND ADD ITEM FORMS BUTTONS
            EditItemForm.EditItemImageBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            EditItemForm.EditItemCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            EditItemForm.EditItemSaveChangesBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            EditItemForm.EditItemUndoChangesBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            AddItemForm.AddItemCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            AddItemForm.AddItemClearAllBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            AddItemForm.AddItemImageBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            AddItemForm.AddItemSaveBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)

            ItemImageSelector.AddSelectImageBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ItemImageSelector.AddSelectImageCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
        End If
    End Sub

    'Stuff to run after program starts
    Private Sub Form1_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        LoadConfigFile()
        OpenDatabaseRoutine(Databasefile)
        Me.CurrentDatabaseLABEL.Text = Replace(My.Computer.FileSystem.GetName(Databasefile), ".txt", "")

        StartTimer()
        If LoggerRunning = False Then RichTextBox1.Text = "AutoLogging is Idle" & vbCrLf
        If My.Computer.FileSystem.DirectoryExists(EtalPath) = False Then
            Dim MySettings As New Settings
            MySettings.Show()
        End If

        SearchFieldCOMBOBOX.Text = "Item Name"
        AllItemsInDatabaseListBox.Select() ' focuses control on main listbox on startup
    End Sub

    'DEFINES THE ImportTimer AS A GLOBAL SYSTEM TIMER (ONLY FOR AUTO IMPORTS)
    Public WithEvents ImportTimer As New System.Windows.Forms.Timer()

    'START IMPORT TIMER ROUTINE
    Sub StartTimer()
        Timercount = 0
        TimerSecs = TimerMins * 60
        ToolStripProgressBar1.Maximum = 100
        ImportTimer.Interval = 1000
        ImportTimer.Start()

    End Sub

    'IMPORT DELAY COUNTER - HANDLES IMPORT DLEAY AUTOLOGGING SYSTEM - BRANCES TO IMPORT ROUTINE WHEN DELAY IS UP
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
        D2StyleOpenFileDialog.ShowDialog()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start() '       restart timer if not paused
    End Sub

    'Populates all items listbox routine
    Public Sub Display_Items()
        AllItemsInDatabaseListBox.Items.Clear() 'insure old list is definetly clear b4 repoulating
        For x = 0 To Objects.Count - 1
            AllItemsInDatabaseListBox.Items.Add(Objects(x).ItemName)
        Next
        ListboxTABCONTROL.SelectTab(0) ' ensure listboxTABCONTROL itset to all items list for repopuleate
        ItemTallyTEXTBOX.Text = Objects.Count & " - Total Items"

        If AllItemsInDatabaseListBox.Items.Count > 0 Then AllItemsInDatabaseListBox.SelectedIndex = 0
    End Sub

    'OPENES THE NEW ITEM FORM TO MANUALLY ADD A NEW ITEM TO THE DATABASE
    Private Sub AddNewItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddNewItemToolStripMenuItem.Click
        AddNewItem() 'add item form is opened in this next sub
    End Sub

    'OPENS THE EDIT EXISTING ITEM FORM TO MANUALLY EDIT AN ITEMS FIELDS
    Private Sub EditExistingItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditExistingItemToolStripMenuItem.Click
        EditItem() 'edit item form is opened in this next sub
    End Sub

    'ALL ITEMS LISTBOX SELECTED INDEX CHANGE HANDLER
    Private Sub AllItemsInDatabaseListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AllItemsInDatabaseListBox.SelectedIndexChanged
        If AllItemsInDatabaseListBox.SelectedIndex <> -1 Then
            Dim RowNumber As Integer = AllItemsInDatabaseListBox.SelectedIndex
            If RowNumber = -1 Then Return ' do nothing
            If RowNumber >= 0 Then
                RichTextBox2.Text = ""  'clears form has been some overlap of listings occur - wierd behaviour

                ' MuleInfoRICHTEXTBOX.Clear() 'clears mule info textbox
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

                    If Objects(RowNumber).ItemBase = "Quest" Then
                        RichTextBox2.SelectionColor = Color.Orange
                        RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
                    End If

                    If Objects(RowNumber).ItemName.IndexOf("Rune") = -1 And Objects(RowNumber).ItemBase <> "Quest" Then
                        RichTextBox2.SelectionColor = Color.White
                        RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
                    End If
                End If
                If DisplayType = "Magic" Then
                    RichTextBox2.SelectionColor = Color.DodgerBlue
                    RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
                End If
                If DisplayType = "Rare" Then
                    RichTextBox2.SelectionColor = Color.Yellow
                    RichTextBox2.SelectedText = Objects(RowNumber).ItemName & vbCrLf
                End If
                If DisplayType = "Crafted" Then
                    RichTextBox2.SelectionColor = Color.DarkGoldenrod
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

                'converts pass to '***' with function if hide pass is set
                If KeepPassPrivate = True Then MulePassTextbox.Text = HidePass(Objects(RowNumber).MulePass) Else MulePassTextbox.Text = Objects(RowNumber).MulePass

            End If
            RichTextBox2.SelectAll()
            RichTextBox2.SelectionAlignment = HorizontalAlignment.Center
            PictureBox1.Load("Skins\" + ItemImageList(Objects(RowNumber).ItemImage) + ".jpg")
        End If
    End Sub

    'DELETE ITEM BUTTON PRESS HANDLER - BRANCHES TO DeleteItem ROUTINE
    Private Sub DeleteItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteItemToolStripMenuItem.Click
        DeleteItem()
    End Sub

    'IMPORT NOW FUNCTION BUTTON PRESS HANDLER - ACTRIVATES THE IMPORT ROUTINE RIGHT NOW REGARDLESS OF IMPORT DELAY VALUE
    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        RichTextBox1.Text = "Checking for new Logs" & vbCrLf
        ImportTimer.Stop()
        ImportLogFiles()
        RichTextBox1.AppendText("AutoLogging is Idle")
        Timercount = 0
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'SAVE DATABASE CONFIRMATION DIALOG ROUTINE
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        If Objects.Count < 1 Then Return ' nothing to save
        YesNoD2Style.Text = "Save Database"
        YesNoD2Style.YesNoHeaderLABEL.Text = "CONFIRM SAVE"
        YesNoD2Style.YesNoHeaderLABEL.TextAlign = ContentAlignment.MiddleCenter
        YesNoD2Style.YesNoMessageLABEL.Text = "You are about to update the " & My.Computer.FileSystem.GetName(Databasefile) & " database file." & vbCrLf & vbCrLf & _
                                              "Select " & Chr(34) & "Confirm" & Chr(34) & " to replace the old database file or" & Chr(34) & "Cancel" & Chr(34) & " to abort.."
        YesNoD2Style.ShowDialog()
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.No Then Return

        ImportTimer.Stop()
        SaveItems()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'SAVES THE CURRENT DATABASE TO FILE
    Public Sub SaveItems()
        Try
            Dim LogWriter = My.Computer.FileSystem.OpenTextFileWriter(Databasefile, False)
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
        YesNoD2Style.Text = "Backup Database"
        YesNoD2Style.YesNoHeaderLABEL.Text = "CONFIRM BACKUP"
        YesNoD2Style.YesNoMessageLABEL.Text = "You are about to backup the " & My.Computer.FileSystem.GetName(Databasefile) & " database file. This will replace any backup file that already exists for this database." & vbCrLf & vbCrLf & _
                                              "Select " & Chr(34) & "Confirm" & Chr(34) & " to replace the old backup file or " & Chr(34) & "Cancel" & Chr(34) & " to abort.."
        YesNoD2Style.ShowDialog()
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.Yes Then Module1.BackupDatabase()
    End Sub

    'SHOW HELP FILE
    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        Help.ShowDialog()
    End Sub

    'PAUSE IMPORT TIMER HANDLER
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        ImportTimer.Stop()
        ItemImageSelector.ShowDialog()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'SEARCH / REFINE NOW BUTTON CLICK HANDLER
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
                If item <> Nothing Then SearchWordCOMBOBOX.Items.Add(item)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'populate for unique attributes block FOR THE SAME REASONS AS FOR THE ITEM NAME FIELD ABOVE?/\
        If UCase(SearchFieldCOMBOBOX.Text) = "UNIQUE ATTRIBUTES" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each item In UniqueAttribsPulldownList
                If item <> Nothing Then SearchWordCOMBOBOX.Items.Add(item)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL ITEM BASE ENTRYS WHEN ITEM BASE IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "ITEM BASE" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If ItemObjectItem.ItemBase <> Nothing Then If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.ItemBase) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.ItemBase)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL ITEM QUALITY ENTRYS WHEN ITEM QUALITY IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "ITEM QUALITY" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If ItemObjectItem.ItemQuality <> Nothing Then If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.ItemQuality) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.ItemQuality)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL USER REFERENCE ENTRYS WHEN USER REF IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "USER REFERENCE" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each item In UserReferencePulldownList
                If item <> Nothing Then SearchWordCOMBOBOX.Items.Add(item)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL MULE ACCOUNT ENTRYS WHEN MULE ACCOUNT IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "MULE ACCOUNT" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If ItemObjectItem.MuleAccount <> Nothing Then If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.MuleAccount) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.MuleAccount)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL MULE NAME ENTRYS WHEN MULE NAME IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "MULE NAME" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If ItemObjectItem.MuleName <> Nothing Then If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.MuleName) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.MuleName)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL MULE PASS ENTRYS WHEN MULE PASS IS SELECTED FOR SEARCH
        'BUT ONLY WHEN HIDE PASWORDS CHECKBOX IN SETTINGS IS UNCHECKED. A BLANK DROPDOWN IS RETURNED IF HIDE PASSWORDS IS CHECKED
        If UCase(SearchFieldCOMBOBOX.Text) = "MULE PASS" Then
            SearchWordCOMBOBOX.Items.Clear()
            If Settings.CheckBox3.Checked = False Then
                For Each ItemObjectItem As ItemObjects In Objects
                    If ItemObjectItem.MulePass <> Nothing Then If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.MulePass) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.MulePass)
                Next
                SearchWordCOMBOBOX.Select()
            End If
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL ATTACK CLASS ENTRYS WHEN ATTACK CLASS IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "ATTACK CLASS" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If ItemObjectItem.AttackClass <> Nothing Then If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.AttackClass) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.AttackClass)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL ATTACK SPEED ENTRYS WHEN ATTACK SPEED IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "ATTACK SPEED" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If ItemObjectItem.AttackSpeed <> Nothing Then If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.AttackSpeed) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.AttackSpeed)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'POPULATE WORD SEARCH DROPDOWN WITH ALL RUNEWORD ENTRYS WHEN RUNEWORD IS SELECTED FOR SEARCH
        If UCase(SearchFieldCOMBOBOX.Text) = "RUNEWORD" Then
            SearchWordCOMBOBOX.Items.Clear()
            For Each ItemObjectItem As ItemObjects In Objects
                If ItemObjectItem.RuneWord <> Nothing Then If SearchWordCOMBOBOX.Items.Contains(ItemObjectItem.RuneWord) = False Then SearchWordCOMBOBOX.Items.Add(ItemObjectItem.RuneWord)
            Next
            SearchWordCOMBOBOX.Select()
        End If

        'Clear out the word pulldowns if var searches apply also clears out word search entry and select value box ready for input
        If UCase(SearchFieldCOMBOBOX.Text) = "ONE HAND DAMAGE MIN" Or UCase(SearchFieldCOMBOBOX.Text) = "ONE HAND DAMAGE MAX" Or UCase(SearchFieldCOMBOBOX.Text) = "TWO HAND DAMAGE MIN" Or UCase(SearchFieldCOMBOBOX.Text) = "TWO HAND DAMAGE MAX" _
             Or UCase(SearchFieldCOMBOBOX.Text) = "THROW DAMAGE MIN" Or UCase(SearchFieldCOMBOBOX.Text) = "THROW DAMAGE MAX" Or UCase(SearchFieldCOMBOBOX.Text) = "REQUIRED LEVEL" Or UCase(SearchFieldCOMBOBOX.Text) = "REQUIRED STRENGTH" _
              Or UCase(SearchFieldCOMBOBOX.Text) = "REQUIRED DEXTERITY" Or UCase(SearchFieldCOMBOBOX.Text) = "CHANCE TO BLOCK" Or UCase(SearchFieldCOMBOBOX.Text) = "ITEM DEFENSE" Then SearchWordCOMBOBOX.Items.Clear() : SearchWordCOMBOBOX.Text = "" : SearchValueNUMERICUPDWN.Select()
    End Sub

    'INDEX CHANGED HANDLER FOR THE SEARCH LISTBOX - SELECTS THE SAME ITEM IN THE ALL ITEMS LISTBOX SO STATS WILL DISPLAY FOR THE SELECTED SEARCHed ITEM
    Private Sub SearchLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SearchLISTBOX.SelectedIndexChanged
        AllItemsInDatabaseListBox.SelectedIndex = -1 'Clears mutiple selection(s) from AllItemsListbox otherwise it only seems to foucus on the topmost item 

        'FIX POTENTIAL BUG FOR WHEN NOTHING IS SELECTED
        If SearchLISTBOX.SelectedIndex <> -1 Then
            AllItemsInDatabaseListBox.SelectedIndex = SearchReferenceList(SearchLISTBOX.SelectedIndex)
            'IF NOTING IS SELECTED IN THE SEARCH LIST DO THIS
        Else
            'CLEAN OUT OLD ITEM STATS
            PictureBox1.Image = Nothing
            RichTextBox2.Text = Nothing
            MuleAccountTextbox.Text = Nothing
            MulePassTextbox.Text = Nothing
            MuleNameTextbox.Text = Nothing
        End If
    End Sub

    'PAUSE AND RESTART AUTO IMPORT TIMER HANDLER
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ListControlTabBUTTON.Click
        ListboxTABCONTROL.SelectTab(0)
        ListControlTabBUTTON.BackColor = Color.DimGray
        SearchListControlTabBUTTON.BackColor = Color.Black
        TradesListControlTabBUTTON.BackColor = Color.Black
        ItemTallyTEXTBOX.Text = AllItemsInDatabaseListBox.Items.Count & " - Total Items"
    End Sub

    'this button selects tab page 1 and colors button to show all search listbox on page 1 and displays total search matches value to textbox2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles SearchListControlTabBUTTON.Click
        ListboxTABCONTROL.SelectTab(1)
        SearchListControlTabBUTTON.BackColor = Color.DimGray
        ListControlTabBUTTON.BackColor = Color.Black
        TradesListControlTabBUTTON.BackColor = Color.Black
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Total Matches"
    End Sub

    'selects user list tab
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles TradesListControlTabBUTTON.Click
        ListboxTABCONTROL.SelectTab(2)
        SearchListControlTabBUTTON.BackColor = Color.Black
        ListControlTabBUTTON.BackColor = Color.Black
        TradesListControlTabBUTTON.BackColor = Color.DimGray

        'SHORT ROUTINE TO COUNT TRADE ITEMS IN RICHTEXT3 BY COUNTING THE GAPS BETWEEN THE ITEMS (SUBTRACTS 1 DUE TO EMPTY LINE AT END OF TEXT) 
        Dim TradeItemCounter As Integer = 0
        For Each item In RichTextBox3.Lines
            If item = Nothing Then TradeItemCounter = TradeItemCounter + 1
        Next
        If TradeItemCounter = 0 Then TradeItemCounter = 1
        ItemTallyTEXTBOX.Text = (TradeItemCounter - 1) & " - Trade Entries"
    End Sub

    ' This restores the database from its backup file (if there is one ofc)
    Private Sub RestoreBackupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestoreBackupToolStripMenuItem.Click
        'SET D2 DIALOG TITLE AND MESSAGES
        YesNoD2Style.Text = "Restore Database From Backup"
        YesNoD2Style.YesNoHeaderLABEL.Text = "CONFIRM TO RESTORE BACKUP"
        YesNoD2Style.YesNoHeaderLABEL.TextAlign = ContentAlignment.MiddleCenter
        YesNoD2Style.YesNoMessageLABEL.Text = "You are about to permanently delete the " & Chr(34) + CurrentDatabaseLABEL.Text + Chr(34) & " database and replace it with its backup file (if one exists)." & vbCrLf & vbCrLf & _
                                              "Only continue if you are sure about what you are doing as severe data loss may occur if you havn't backed up regularly." & vbCrLf & vbCrLf & _
                                              "Select " & Chr(34) & "Confirm" & Chr(34) & " to restore the backup file or " & Chr(34) & "Cancel" & Chr(34) & " to abort.."
        YesNoD2Style.ShowDialog()

        'On confirmation Restore form backup
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.Yes Then

            Dim BackupPath = Application.StartupPath & "\Database\Backup\"
            Dim temp As String = ""
            Dim myarray = Split(Databasefile, ".txt", 0)
            Dim tempname = myarray(0) & ".bak"
            myarray = Split(tempname, "\")
            tempname = myarray(myarray.Length - 1)
            If My.Computer.FileSystem.FileExists(BackupPath & tempname) = True Then

                'found a backup file and copying it over to replace current database file, then reload it
                My.Computer.FileSystem.DeleteFile(Databasefile) 'delete old dbase file
                My.Computer.FileSystem.CopyFile(BackupPath & tempname, Databasefile, True) ' copy over new dbase file and rename it
                OpenDatabaseRoutine(Databasefile) 'refresh new database file to the lists

                'SELECT MAIN LIST BOX AFTER BACKUP RESTORE
                ListControlTabBUTTON.BackColor = Color.DimGray
                SearchListControlTabBUTTON.BackColor = Color.Black
                TradesListControlTabBUTTON.BackColor = Color.Black
                ListboxTABCONTROL.SelectTab(0)

            Else
                'There is not backup file for this database so cant restore it
            End If
        End If

        'cancel backup restoration
        If YesNoD2Style.DialogResult = Windows.Forms.DialogResult.No Then
        End If
    End Sub

    'Closing Form1 (exit Application) Event Handler. Throws To exit confirmation message. also handles saves and backups if set to do so
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ClosingAppForm.ShowDialog()
        If ClosingAppForm.DialogResult = Windows.Forms.DialogResult.No Then e.Cancel = True : Return
        If ClosingAppForm.DialogResult = Windows.Forms.DialogResult.Yes Then
            If saveonexit = True Then SaveItems()
            If backuponexit = True Then Module1.BackupDatabase()
        End If

    End Sub

    'This is the actual exit command to close app which triggers above handler (do it this way so handler activates no matter where app is closed from)
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    'CREATE NEW DATABASE OPTION FOR MULTIPLE DATABASES-------> NEEDS TO BE RESOLVED YET <-------------------------------------------[FINISH THIS ALREADY ROB]
    Private Sub NewToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem1.Click
        If LoggerRunning = True Then ' dont try to create file if auto logger is working - will cause app crash
            Mymessages = "Please wait File in use, Logger Running" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()                                      'stop timer b4 new database creation
        CreateNewDatabase.ShowDialog()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start() 'restart timer but only if its not set to pause
    End Sub

    'DISPLAYS THE SETTINGS FROM SELECTED FROM PULLDOWN MENUES
    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()
        Settings.Show()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()

    End Sub

    'This Does The Sort Database Command From Pull Down "items" Menu
    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Objects.Sort(Function(x, y) x.ItemName.CompareTo(y.ItemName))
        Display_Items()
    End Sub

    'ALL ITEMS LISTBOX CONTEXT MENU STRIP DISPLAY ROUTINE
    Private Sub AllItemsInDatabaseListBox_MouseDown(sender As Object, e As MouseEventArgs) Handles AllItemsInDatabaseListBox.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If AllItemsInDatabaseListBox.SelectedIndex > -1 Then
                If Not String.IsNullOrEmpty(AllItemsInDatabaseListBox.Text) Then
                    Me.ItemListboxCONTEXTMENUSTRIP.Show(Control.MousePosition)
                End If
            End If
        Else
        End If
    End Sub

    'Search Listbox Context MenuStrip display routine 
    Private Sub SearchLISTBOX_MouseDown(sender As Object, e As MouseEventArgs) Handles SearchLISTBOX.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If SearchLISTBOX.Items.Count > 0 Then
                Me.SearchListboxCONTEXTMENUSTRIP.Show(Control.MousePosition)
            End If
        End If
    End Sub

    'trade entries Context MenuStrip display routine 
    Private Sub RichTextBox3_MouseDown(sender As Object, e As MouseEventArgs) Handles RichTextBox3.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If RichTextBox3.Text <> Nothing Then
                Me.TradesCONTEXTMENUSTRIP.Show(Control.MousePosition)
            End If
        End If
    End Sub

    'BUTTON CLICK HANDLERS FOR ALL ITEMS CONTEXT POPUP MENU OPTIONS

    'Add New Item  Menu Option All Items Context Menu
    Private Sub AddItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddItemToolStripMenuItem.Click
        AddNewItem()
    End Sub

    'Edit Item Menu Option All Items Context Menu
    Private Sub EditItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditItemToolStripMenuItem.Click
        EditItem()
    End Sub

    'Delete Item(s) Menu Option
    Private Sub DeleteItemToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DeleteItemToolStripMenuItem1.Click
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()
        DeleteItem()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    Private Sub SortToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SortToolStripMenuItem.Click
        Objects.Sort(Function(x, y) x.ItemName.CompareTo(y.ItemName))
        Display_Items() ' refresh list after sorting
    End Sub

    'CLEARS THE SEARCH LIST SELECTIONS (HIGHLIGHTED ITEMS IN SEARCHLIST)
    Private Sub ClearSearchListToolStripMenuItem_Click(sender As Object, e As EventArgs)
        SearchLISTBOX.SelectedItem = -1
    End Sub

    'Action subs relating to mouse click events redirected to here
    '*************************************************************
    Private Sub DeleteItem()
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()
        Dim FocusOnExit As Integer = AllItemsInDatabaseListBox.SelectedIndex

        'check for backup on edits set to true if so then backup now
        If Settings.BackupOnEditsCHECKBOX.Checked = True Then Module1.BackupDatabase()
        If AllItemsInDatabaseListBox.SelectedIndices.Count > 0 Then
            For index = AllItemsInDatabaseListBox.SelectedIndices.Count - 1 To 0 Step -1
                Dim a As Integer = AllItemsInDatabaseListBox.SelectedIndices(index)
                AllItemsInDatabaseListBox.Items.RemoveAt(a)
                Objects.RemoveAt(a)
                'Removes Item From Search Listbox & Ref List
                Dim count As Integer = SearchReferenceList.Count - 1
                For Each item In SearchReferenceList

                    If SearchReferenceList(count) = a Then
                        SearchReferenceList.RemoveAt(count)
                        SearchLISTBOX.Items.RemoveAt(count)

                        For x = a To SearchReferenceList.Count - 1
                            SearchReferenceList(x) = SearchReferenceList(x) - 1
                        Next
                        Exit For
                    End If
                    count = count - 1
                Next
            Next
            ItemTallyTEXTBOX.Text = AllItemsInDatabaseListBox.Items.Count & " - Total Items"

            'SET THE DELETED OBJECT LOCATION IN THE LIST AS THE HIGHLIGHTED ITEM ON RETURN FROM DETETE
            If FocusOnExit >= (AllItemsInDatabaseListBox.Items.Count) Then FocusOnExit = AllItemsInDatabaseListBox.Items.Count - 1
            If AllItemsInDatabaseListBox.Items.Count = 1 Then FocusOnExit = 0
            If AllItemsInDatabaseListBox.Items.Count = 0 Then FocusOnExit = -1
            ClearStats()
            AllItemsInDatabaseListBox.SelectedIndex = FocusOnExit
            If Button3.Text = "Timer Stop" Then ImportTimer.Start()
            Return
        End If
    End Sub

    'Deletes items form search list by selecting each item and deleting it from the main list
    Private Sub DeleteItemsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteItemsToolStripMenuItem.Click
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()
        DeleteSearchItems() 'branch to searchdete sub
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'EDIT ITEM BUTTON CLICK HANDLER - BRANCHES TO EDIT ITEM FORM
    Private Sub EditItem()
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()
        If AllItemsInDatabaseListBox.SelectedIndex > -1 Then EditItemForm.ShowDialog()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'ADD NEW ITEM BUTTON CLICK HANDLER - BRANCHES TO ADD ITEM FORM
    Private Sub AddNewItem()
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop() '        stop timer b4 form opens
        AddItemForm.ShowDialog()
        If Button3.Text = "Timer Stop" Then ImportTimer.Start() '       restart timer after form closes

    End Sub

    'SENDS HIGHLIGHTED ITEMS TO THE TRADE LIST
    Private Sub AddToUserListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddToUserListToolStripMenuItem.Click

        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If
        ImportTimer.Stop()

        If AllItemsInDatabaseListBox.SelectedIndices.Count > 0 Then
            Dim a As Integer = 0
            Dim count As Integer = 0
            DupeCountProgressForm.Show() : DupeCountProgressForm.DupePROGRESSBAR.Value = 0 'show and reset progress bar
            For index = 0 To AllItemsInDatabaseListBox.SelectedIndices.Count - 1

                'CALCULATE PROGRESS BAR
                DupeCountProgressForm.DupePROGRESSBAR.Value = Int((count / AllItemsInDatabaseListBox.Items.Count) * 100)
                count = count + 1
                a = AllItemsInDatabaseListBox.SelectedIndices(index)

                Dim Temp = Objects(a).ItemName
                If Objects(a).ItemBase = "Rune" Or Objects(a).ItemBase = "Gem" Or Objects(a).ItemName.IndexOf("Token") > -1 Or Objects(a).ItemName.IndexOf("Key of") > -1 Or Objects(a).ItemName.IndexOf("Essence") > -1 Then
                    If Objects(a).ItemName.IndexOf("Token") > -1 Then Temp = "Token"
                    RichTextBox3.AppendText(Temp & vbCrLf & vbCrLf)
                Else
                    SendToTradeList(a)
                End If
            Next
            AllItemsInDatabaseListBox.SelectedIndex = -1
        End If

        DupeCountProgressForm.Close()
        DupesList()

        'SET TRADELIST HIGHLIGHT AND SELECT TRADE LIST TAB
        ListControlTabBUTTON.BackColor = Color.Black
        SearchListControlTabBUTTON.BackColor = Color.Black
        TradesListControlTabBUTTON.BackColor = Color.DimGray
        ListboxTABCONTROL.SelectTab(2)

        'SHORT ROUTINE TO COUNT TRADE ITEMS IN RICHTEXT3 BY COUNTING THE GAPS BETWEEN THE ITEMS (SUBTRACTS 1 DUE TO EMPTY LINE AT END OF TEXT) 
        Dim TradeItemCounter As Integer = 0
        For Each item In RichTextBox3.Lines
            If item = Nothing Then TradeItemCounter = TradeItemCounter + 1
        Next
        If TradeItemCounter = 0 Then TradeItemCounter = 1
        ItemTallyTEXTBOX.Text = (TradeItemCounter - 1) & " - Trade Entries"
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'ADD THE SELECTED ITEM TO THE TRADE LIST FROM SEARCH LIST
    Private Sub AddItemToTradeListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddItemToTradeListToolStripMenuItem.Click

        'CHECK FOR LOGGER ACTIVE
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If

        ImportTimer.Stop()
        If SearchLISTBOX.Items.Count > 0 Then

            'ADD THE SELECTED ITEM
            Dim a = SearchReferenceList(SearchLISTBOX.SelectedIndex)
            SendToTradeList(a)
        End If
        AllItemsInDatabaseListBox.SelectedIndex = -1
        DupesList()

        'SET TRADELIST HIGHLIGHT AND SELECT TRADE LIST TAB
        ListControlTabBUTTON.BackColor = Color.Black
        SearchListControlTabBUTTON.BackColor = Color.Black
        TradesListControlTabBUTTON.BackColor = Color.DimGray
        ListboxTABCONTROL.SelectTab(2)

        'SHORT ROUTINE TO COUNT TRADE ITEMS IN RICHTEXT3 BY COUNTING THE GAPS BETWEEN THE ITEMS (SUBTRACTS 1 DUE TO EMPTY LINE AT END OF TEXT) 
        Dim TradeItemCounter As Integer = 0
        For Each item In RichTextBox3.Lines
            If item = Nothing Then TradeItemCounter = TradeItemCounter + 1
        Next
        If TradeItemCounter = 0 Then TradeItemCounter = 1
        ItemTallyTEXTBOX.Text = (TradeItemCounter - 1) & " - Trade Entries"
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'ADD ALL ITEMS TO TRADE LIST FROM SEARCH LIST
    Private Sub AddAllItemsToUserListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddAllItemsToUserListToolStripMenuItem.Click

        'CHECK FOR LOGGER ACTIVE
        If LoggerRunning = True Then
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If

        ImportTimer.Stop()
        If SearchLISTBOX.Items.Count > 0 Then
            RichTextBox3.Clear()
            'ADDS ALL THE ITEMS
            Dim Counter As Integer = 0
            Dim count As Integer = 0
            DupeCountProgressForm.Show() : DupeCountProgressForm.DupePROGRESSBAR.Value = 0 'show and reset progress bar

            For Each ITEM In SearchLISTBOX.Items

                'CALCULATE PROGRESS BAR
                DupeCountProgressForm.DupePROGRESSBAR.Value = Int((count / SearchLISTBOX.Items.Count) * 100)
                count = count + 1

                Dim a = SearchReferenceList(Counter)
                Dim Temp = Objects(a).ItemName
                If Objects(a).ItemBase = "Rune" Or Objects(a).ItemBase = "Gem" Or Objects(a).ItemName.IndexOf("Token") > -1 Or Objects(a).ItemName.IndexOf("Key of") > -1 Or Objects(a).ItemName.IndexOf("Essence") > -1 Then

                    If Objects(a).ItemName.IndexOf("Token") > -1 Then Temp = "Token"
                    RichTextBox3.AppendText(Temp & vbCrLf & vbCrLf)
                Else
                    SendToTradeList(a)
                End If
                Counter = Counter + 1
            Next
            AllItemsInDatabaseListBox.SelectedIndex = -1
        End If
        DupesList()
        DupeCountProgressForm.Close()

        'SET TRADELIST HIGHLIGHT AND SELECT TRADE LIST TAB
        ListControlTabBUTTON.BackColor = Color.Black
        SearchListControlTabBUTTON.BackColor = Color.Black
        TradesListControlTabBUTTON.BackColor = Color.DimGray
        ListboxTABCONTROL.SelectTab(2)

        'SHORT ROUTINE TO COUNT TRADE ITEMS IN RICHTEXT3 BY COUNTING THE GAPS BETWEEN THE ITEMS (SUBTRACTS 1 DUE TO EMPTY LINE AT END OF TEXT) 
        Dim TradeItemCounter As Integer = 0
        For Each item In RichTextBox3.Lines
            If item = Nothing Then TradeItemCounter = TradeItemCounter + 1
        Next
        If TradeItemCounter = 0 Then TradeItemCounter = 1
        ItemTallyTEXTBOX.Text = (TradeItemCounter - 1) & " - Trade Entries"
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()

    End Sub

    Private Sub DupesList()
        Dim arr() As String = RichTextBox3.Text.Split(Chr(10))
        Dim count(UBound(arr)) As Integer

        For i = 0 To UBound(arr) - 1 'find duplicates and delete
            count(i) = 1
            For x = (i + 1) To UBound(arr)
                If arr(x) = "" Then Continue For
                If arr(i) = arr(x) Then arr(x) = "" : count(i) += 1
            Next
        Next
        RichTextBox3.Clear() 'clear list
        For x = 0 To UBound(arr) ' re - sort and put back
            If count(x) > 1 Then arr(x) = arr(x) & " (" & count(x) & ")"
        Next
        Array.Sort(arr)

        For x = 0 To UBound(arr) ' re - sort and put back
            If arr(x) <> "" Then
                RichTextBox3.AppendText(arr(x) & vbCrLf & vbCrLf)
            End If
        Next
    End Sub

    'clears the trade RICHTEXT on items menu bar selection
    Private Sub ClearTradeListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearTradeListToolStripMenuItem.Click
        RichTextBox3.Clear()
    End Sub

    'clears the trade RICHTEXT on CONTEXT MENU SELECTION
    Private Sub ClearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
        ItemTallyTEXTBOX.Text = "0 - Trade Entries " : RichTextBox3.Clear()
    End Sub

    'COPY ENTIRE TRADE RICHTEXT TO CLIPBOARD
    'If any text is selected only the selection will be copied otherwise the entire text will be copied

    Private Sub CopyToClipboardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToClipboardToolStripMenuItem.Click
        My.Computer.Clipboard.Clear()
        If RichTextBox3.SelectedText = Nothing Then
            My.Computer.Clipboard.SetText(RichTextBox3.Text)
        Else
            My.Computer.Clipboard.SetText(RichTextBox3.SelectedText)
        End If
    End Sub

    'APPEND ENTIRE TRADE RICHTEXT TO CLIPBOARD
    'If any text is selected only the selection will be appended otherwise the entire text will be appended

    Private Sub AppendToClipboardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AppendToClipboardToolStripMenuItem.Click
        If RichTextBox3.SelectedText = Nothing Then
            My.Computer.Clipboard.SetText(My.Computer.Clipboard.GetText & RichTextBox3.Text)
        Else
            My.Computer.Clipboard.SetText(My.Computer.Clipboard.GetText & RichTextBox3.SelectedText)
        End If

    End Sub

    'starts search routine from enter keypress   
    Private Sub SearchBUTTON_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchBUTTON.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            'This next 'if then' checks that the selected search operator is valid if its not set to default equal to, in upper case the emphasise the fix to the user

            'NOTED: Checks for the other comboboxes entry vaildity are IN EACH ENTRY CHANGED EVENT HANDLER. The Search button will remain disabled until the 
            'minimum ammount of search criteria has been entered. This insures the search routine wont crash due to user input error (i hope)
            If UCase(SearchOperatorCOMBOBOX.Text) = "EQUAL TO" Or UCase(SearchOperatorCOMBOBOX.Text) = "NOT EQUAL TO" Or UCase(SearchOperatorCOMBOBOX.Text) = "GREATER THAN" Or UCase(SearchOperatorCOMBOBOX.Text) = "LESS THAN" Then
                SearchRoutine() ' If all is good and search seems valid then branch to the search routine sub in Module1
            End If
            If SearchLISTBOX.Items.Count > 0 Then SearchLISTBOX.Focus() Else AllItemsInDatabaseListBox.Focus() ' ensure focus is returned to the relevant item list after search
        End If
    End Sub

    'SWITCHES FOCUS FROM STRING ENTRY TEXTBOX TO SEARCH BUTTON ON ENTER KEYPRESS
    Private Sub SearchWordCOMBOBOX_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchWordCOMBOBOX.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If UCase(SearchOperatorCOMBOBOX.Text) = "EQUAL TO" Or UCase(SearchOperatorCOMBOBOX.Text) = "NOT EQUAL TO" Or UCase(SearchOperatorCOMBOBOX.Text) = "GREATER THAN" Or UCase(SearchOperatorCOMBOBOX.Text) = "LESS THAN" Then
                SearchRoutine() ' If all is good and search seems valid then branch to the search routine sub in Module1
            End If
            If SearchLISTBOX.Items.Count > 0 Then SearchLISTBOX.Focus() Else AllItemsInDatabaseListBox.Focus() ' ensure focus is returned to the relevant item list after search
        End If

    End Sub

    'SWITCHES FOCUS FROM  VALUE ENTRY NUMERIC UP DOWN BOX TO SEARCH BUTTON ON ENTER KEYPRESS
    Private Sub SearchValueNUMERICUPDWN_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchValueNUMERICUPDWN.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            SearchBUTTON.Select()
        End If
    End Sub

    'CLEANS OUT OLD ITEM STATISTICS
    Sub ClearStats()
        PictureBox1.Image = Nothing
        PictureBox1.BackgroundImage = DiaBase.My.Resources.Resources.ImageBackground '  clean out old image
        RichTextBox2.Text = Nothing '                                                   clean out old item stats
        RichTextBox1.Text = Nothing '                                                   clean out logging window
        MuleAccountTextbox.Text = Nothing '                                             clean out mule Account
        MulePassTextbox.Text = Nothing '                                                clean out mule Pass
        MuleNameTextbox.Text = Nothing '                                                clean out mule Name
    End Sub

    'MOVES SELECECTED ITEMS IN SEARCH MATCHES LIST OVER TO A NEW DATABASE OR CREATE A NEW ONE (NOS SURE WHAT WORDING WAS BETTER: "EXPORT ITEM(S)" OR "MOVE ITEM(S)"
    Private Sub ExportSelectedItemsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportSelectedItemsToolStripMenuItem.Click

        If LoggerRunning = True Then 'Check for logger running .... as usual lol
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If

        ImportTimer.Stop()
        If SearchLISTBOX.SelectedIndex <> -1 Then MoveItems.ShowDialog() ' MOVE/EXPORT ROUTINES ARE ALL IN THE "MoveItems" CODE FORM
        If Button3.Text = "Timer Stop" Then ImportTimer.Start()
    End Sub

    'DELETES SELECTED ITEMS FROM THE SEARCH LIST
    Sub DeleteSearchItems()
        Dim a As Integer
        Dim b As Integer
        Dim FocusOnExit As Integer = SearchLISTBOX.SelectedIndex

        For index = SearchLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
            a = SearchLISTBOX.SelectedIndices(index)
            b = SearchReferenceList(a)
            SearchLISTBOX.Items.RemoveAt(a)
            AllItemsInDatabaseListBox.Items.RemoveAt(b)
            Objects.RemoveAt(b)
            SearchReferenceList.RemoveAt(a)
            For x = a To SearchReferenceList.Count - 1
                SearchReferenceList(x) = SearchReferenceList(x) - 1
            Next
        Next
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Total Items"

        'SET THE DELETED OBJECT LOCATION IN THE LIST AS THE HIGHLIGHTED ITEM ON RETURN FROM DETETE
        If FocusOnExit >= (SearchLISTBOX.Items.Count) Then FocusOnExit = SearchLISTBOX.Items.Count - 1
        If SearchLISTBOX.Items.Count = 1 Then FocusOnExit = 0
        If SearchLISTBOX.Items.Count = 0 Then FocusOnExit = -1
        ClearStats()
        SearchLISTBOX.SelectedIndex = FocusOnExit
    End Sub

   'REMOVE MULTIPLE ITEMS FROM SEARCH ITEM LIST - DOES NOT DELETE ITMES ONLY REMOVES THEM FROM THE SEARCH LIST
    Private Sub RemoveSelectedItemssFromSearchListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveSelectedItemssFromSearchListToolStripMenuItem.Click
        If LoggerRunning = True Then '                                                                      CHECK FOR RUNNING AUTOLOG ROUTINE
            Mymessages = "Please wait Import in progress" : MyMessageBox()
            Return
        End If

        Dim FocusOnExit As Integer = SearchLISTBOX.SelectedIndex
        Dim a As Integer
        For index = SearchLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
            a = SearchLISTBOX.SelectedIndices(index)
            SearchLISTBOX.Items.RemoveAt(a)
            SearchReferenceList.RemoveAt(a)
        Next
        SearchLISTBOX.SelectedItem = -1

        'KEEPS FocusOnExit WITHIN LISTBOX BOUNDARYS
        If FocusOnExit >= (SearchLISTBOX.Items.Count) Then FocusOnExit = SearchLISTBOX.Items.Count - 1
        If SearchLISTBOX.Items.Count = 1 Then FocusOnExit = 0
        If SearchLISTBOX.Items.Count = 0 Then FocusOnExit = -1
        ClearStats()
        SearchLISTBOX.SelectedIndex = FocusOnExit
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Total Items"

        If Button3.Text = "Timer Stop" Then ImportTimer.Start() '                                              RESTART LOGGER
    End Sub

    'CLEARS ALL ITEMS OUT OF THE SEARCH LIST
    Private Sub ClearSearchListToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ClearSearchListToolStripMenuItem.Click
        SearchLISTBOX.Items.Clear()
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Total Items"
    End Sub


End Class



