Imports System.IO

Public Class D2StyleOpenFileDialog
    'close open database form
    Private Sub NewDatabaseCancelBUTTON_Click(sender As Object, e As EventArgs) Handles NewDatabaseCancelBUTTON.Click
        Me.Close()
    End Sub

    'Setup open database form - load event handler
    Private Sub D2StyleOpenFileDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SavedDatabasesLISTBOX.Items.Clear()
        DatabaseFilenameCOMBOBOX.Text = Nothing
        Dim AllSavedDatabasesFileNames As Array = Nothing
        AllSavedDatabasesFileNames = Directory.GetFiles(Application.StartupPath + "\Database\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

        For Each item In AllSavedDatabasesFileNames
            SavedDatabasesLISTBOX.Items.Add(Replace(item, ".txt", ""))
        Next

        'refresh the open database combobox drop down with file alread successfully opened in this session (stored in OpenDatabaseDropDown collection)
        For Each item In OpenDatabaseDropDown
            DatabaseFilenameCOMBOBOX.Items.Add(item)
        Next

        DatabaseFilenameCOMBOBOX.Select()
    End Sub

    'KEEPS COMBOBX SELECTED
    Private Sub SavedDatabasesLISTBOX_MouseDown(sender As Object, e As MouseEventArgs) Handles SavedDatabasesLISTBOX.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Right Then
            If SavedDatabasesLISTBOX.SelectedIndex = -1 Then
                OpenDatabaseCONTEXTMENU.Items(0).Enabled = False
                OpenDatabaseCONTEXTMENU.Items(1).Enabled = False
                OpenDatabaseCONTEXTMENU.Items(2).Enabled = False
            Else
                OpenDatabaseCONTEXTMENU.Items(0).Enabled = True
                OpenDatabaseCONTEXTMENU.Items(1).Enabled = True
                OpenDatabaseCONTEXTMENU.Items(2).Enabled = True
            End If

            'If Not String.IsNullOrEmpty(SavedDatabasesLISTBOX.Text) Then
            Me.OpenDatabaseCONTEXTMENU.Show(Control.MousePosition)
        End If
        'End If
        'Else
        'End If
        DatabaseFilenameCOMBOBOX.Select()
        DatabaseFilenameCOMBOBOX.SelectionLength = 0
    End Sub

    'PUTS SELECTED ITEM INTO TEXTBOX AND KEEPS COMBOBX SELECTED
    Private Sub SavedDatabasesLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SavedDatabasesLISTBOX.SelectedIndexChanged
        DatabaseFilenameCOMBOBOX.Text = SavedDatabasesLISTBOX.SelectedItem
        DatabaseFilenameCOMBOBOX.Select()
        DatabaseFilenameCOMBOBOX.SelectionLength = 0
    End Sub

    'OPEN DATABASE BUTTON PRESS
    Private Sub MoveItemsBUTTON_Click(sender As Object, e As EventArgs) Handles OpenDatabaseBUTTON.Click
        OpenSavedDatabase()
    End Sub

    'OPEN DATABASE CONTEXT MENU BUTTON PRESS
    Private Sub OpenSelectedDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenSelectedDatabaseToolStripMenuItem.Click
        OpenSavedDatabase()
    End Sub

    'LOADS THE SELECTED DATABASE IF IT EXISTS
    Sub OpenSavedDatabase()
        Dim OpenError As Boolean = False
        Try
            If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt") = True Then

                'VERIFYS THE SELECTED is really a diabase data file by looking for the exact "--------------------" item spacer (20 x -) on the first line
                Dim FileVerify = My.Computer.FileSystem.OpenTextFileReader(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt")
                Dim TempLine = FileVerify.ReadLine
                If TempLine <> "--------------------" And TempLine <> Nothing Then
                    'VERIFICATION FAILED - CLOSE FORM
                    Mymessages = "DiaBase File Verification Error." : MyMessageBox()
                    SavedDatabasesLISTBOX.SelectedIndex = -1
                    DatabaseFilenameCOMBOBOX.Text = Nothing
                    DatabaseFilenameCOMBOBOX.Select()
                    OpenError = True 'Error flag to skip irrelevant code when an error occurs
                End If

                FileVerify.Close()
                If OpenError = True Then GoTo ErrorSkipPoint

                'CHECKS THE SAVE CURRENT CHECKBOX AND SAVES THE CURRENT DATABASE BEFORE LOADING THE NEW ONE IF SET TO DO SO (DEFAULT IT CHECKED)
                If SaveBeforeOpeningCHECKBOX.Checked = True Then

                    Databasefile = Application.StartupPath + "\Database\" + Form1.CurrentDatabaseLABEL.Text + ".txt" 'UPDATE THE MAIN VAR FOR CURRNET DATABASE FILE AND PATH
                    Form1.SaveItems() 'Branch to save routine to save the current database first

                    'Clean out old items from the last loaded database
                    Form1.SearchLISTBOX.Items.Clear()
                    Form1.RichTextBox3.Text = Nothing
                    Form1.ClearStats()

                    'Branch to routine to load the selected dbase and display first item and set current database label in top right of form1
                    OpenDatabaseRoutine(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt")
                    Form1.Display_Items()
                    Form1.CurrentDatabaseLABEL.Text = DatabaseFilenameCOMBOBOX.Text
                    Databasefile = Application.StartupPath + "\Database\" + SavedDatabasesLISTBOX.SelectedItem + ".txt"

                    'setlect main list and prep form1
                    Form1.ListboxTABCONTROL.SelectTab(0)
                    Form1.ListControlTabBUTTON.BackColor = Color.DimGray
                    Form1.SearchListControlTabBUTTON.BackColor = Color.Black
                    Form1.TradesListControlTabBUTTON.BackColor = Color.Black
                    Form1.ItemTallyTEXTBOX.Text = Form1.AllItemsInDatabaseListBox.Items.Count & " - Total Items"

                    'after successfull save this puts the database name into the combobox dropdown menu if its not already there
                    If OpenDatabaseDropDown.Contains(DatabaseFilenameCOMBOBOX.Text) = False Then OpenDatabaseDropDown.Add(DatabaseFilenameCOMBOBOX.Text)
                    Me.Close() ' close form once file has been verified

                End If
            Else
                'If the entered file does not exist then do this
                DatabaseFilenameCOMBOBOX.Text = Nothing
                SavedDatabasesLISTBOX.SelectedItem = Nothing
                DatabaseFilenameCOMBOBOX.Select()
                OpenError = True
            End If
            If OpenError = True Then GoTo ErrorSkipPoint 'Error flag to skip irrelevant code when an error occurs

            'catch open file errors exception handler
        Catch ex As Exception
            Mymessages = "File Error Opening Saved Database." : MyMessageBox()
        End Try
ErrorSkipPoint:
    End Sub

    'RENAMES THE SELECTED FILE IN OPEN DTATBASE LISTBOX - SavedDatabasesLISTBOX CONTEXT MENU
    Private Sub RenameSelectedDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameSelectedDatabaseToolStripMenuItem.Click

        UserInputForm.Text = "Rename The Selected Database"
        UserInputForm.UserInputHeaderLABEL.Text = "ENTER NEW DATABASE NAME"
        UserInputForm.UserImputMessageLABEL.Text = "To rename the " + SavedDatabasesLISTBOX.SelectedItem + " database file please enter its new unique name into the text box"
        UserInputForm.UserInputYesBUTTON.Text = "Rename"
        UserInputForm.UserInputNoBUTTON.Text = "Cancel"
        UserInputForm.UserInputTEXTBOX.Text = SavedDatabasesLISTBOX.SelectedItem
        UserInputForm.UserInputTEXTBOX.Select()

        DialogResult = UserInputForm.ShowDialog

        If DialogResult = Windows.Forms.DialogResult.Yes Then
            'RENAMES THE DATABASE AND ITS BACKUP AND ITS PROGRAM REFERENCES IF ITS CURRENTLY LOADED  <--------------- R E N A M E   D A T A B A S E    H E R E  
            Try
                'UPDATE THE MAIN DATABASE FILE RENAME
                If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt") = True Then
                    My.Computer.FileSystem.RenameFile(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt", UserInputForm.UserInputTEXTBOX.Text + ".txt")
                End If

                'UPDATE THE DATABASES BACKUP FILE RENAME(IF IT EXISTS)
                If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".bak") = True Then
                    My.Computer.FileSystem.RenameFile(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".bak", UserInputForm.UserInputTEXTBOX.Text + ".bak")
                End If

                'Update default database setting in settings and its var if its the one that was just renamed
                If Settings.DatabaseFileTEXTBOX.Text = Application.StartupPath + "\Database\" + SavedDatabasesLISTBOX.SelectedItem + ".txt" Then

                    Settings.DatabaseFileTEXTBOX.Text = Application.StartupPath + "\Database\" + UserInputForm.UserInputTEXTBOX.Text + ".txt"
                    Settings.DatabaseFileTEXTBOX.Text = Application.StartupPath + "\Database\" + UserInputForm.UserInputTEXTBOX.Text + ".txt"
                    Databasefile = Application.StartupPath + "\Database\" + UserInputForm.UserInputTEXTBOX.Text + ".txt"

                End If

                'Update current database label in top right corner of form1 if its the one that was just renamed
                If Form1.CurrentDatabaseLABEL.Text = SavedDatabasesLISTBOX.SelectedItem Then
                    Form1.CurrentDatabaseLABEL.Text = UserInputForm.UserInputTEXTBOX.Text
                End If

                'update renamed database in the saved database listbox
                Dim AllSavedDatabasesFileNames As Array = Nothing
                AllSavedDatabasesFileNames = Directory.GetFiles(Application.StartupPath + "\Database\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

                For Each item In AllSavedDatabasesFileNames
                    SavedDatabasesLISTBOX.Items.Add(Replace(item, ".txt", ""))
                Next

                'Repopulate saved database list after namechange
                SavedDatabasesLISTBOX.Items.Clear() : AllSavedDatabasesFileNames = Nothing
                AllSavedDatabasesFileNames = Directory.GetFiles(Application.StartupPath + "\Database\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

                For Each item In AllSavedDatabasesFileNames
                    SavedDatabasesLISTBOX.Items.Add(Replace(item, ".txt", ""))
                Next

            Catch ex As Exception
                Mymessages = "File Error Renaming Saved Database." : MyMessageBox()
            End Try
        Else

            If DialogResult = Windows.Forms.DialogResult.No Then

                'Cancels any selections in listbox and or textbox
                DatabaseFilenameCOMBOBOX.Text = Nothing
                SavedDatabasesLISTBOX.SelectedItem = Nothing
                DatabaseFilenameCOMBOBOX.Select()
            End If
        End If
        DialogResult = Nothing
    End Sub

    'DELETES THE SELECTED DATABASE
    Private Sub DeleteSelectedDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteSelectedDatabaseToolStripMenuItem.Click
        Dim DeleteError As Boolean = False

        'CHECK THE SELECTED DATABASE IS NOT THE DEFAULT ONE
        If Settings.DatabaseFileTEXTBOX.Text = Application.StartupPath + DatabaseFilenameCOMBOBOX.Text + ".TXT" Then
            Mymessages = "You Cannot Delete The Default Database." : MyMessageBox()
            DeleteError = True
        End If
        If DeleteError = True Then GoTo SkipDeleteOnError

        'CHECK THE SELECTED DATABASE IS NOT THE CURRENT ONE
        If DatabaseFilenameCOMBOBOX.Text = Form1.CurrentDatabaseLABEL.Text Then
            Mymessages = "You Cannot Delete The Current Database." : MyMessageBox()
            DeleteError = True
        End If
        If DeleteError = True Then GoTo SkipDeleteOnError

        YesNoD2Style.Text = "Delete Database File"
        YesNoD2Style.YesNoHeaderLABEL.Text = "CONFIRM TO DELETE DATABASE"
        YesNoD2Style.YesNoMessageLABEL.Text = "To delete the " + SavedDatabasesLISTBOX.SelectedItem + " database file please press the confirm button" & vbCrLf & vbCrLf + "NOTE: This action is irreversable and will permantly delete all items within the selected database as well as its backup file."
        DialogResult = YesNoD2Style.ShowDialog
        If DialogResult = Windows.Forms.DialogResult.Yes Then

            'DELETES THE ACTUAL DATABASE HERE
            Try

                'DELETE BACKUP
                If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".bak") = True Then
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".bak")
                End If

                'DELETE DATABASE FILE
                If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt") = True Then
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt")
                End If

                'DELETE ERROR ERROR EXCEPTION
            Catch ex As Exception
                Mymessages = "File Error Deleting Database." : MyMessageBox()
                DeleteError = True
            End Try

            If DeleteError = True Then GoTo SkipDeleteOnError

            'Repopulate saved database list after namechange
            SavedDatabasesLISTBOX.Items.Clear() : Dim AllSavedDatabasesFileNames As Array
            AllSavedDatabasesFileNames = Directory.GetFiles(Application.StartupPath + "\Database\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

            For Each item In AllSavedDatabasesFileNames
                SavedDatabasesLISTBOX.Items.Add(Replace(item, ".txt", ""))
            Next

SkipDeleteOnError:

            'Cancels any selections in listbox and or textbox
            DatabaseFilenameCOMBOBOX.Text = Nothing
            SavedDatabasesLISTBOX.SelectedItem = Nothing
            DatabaseFilenameCOMBOBOX.Select()
        End If
        DialogResult = Nothing
    End Sub
End Class