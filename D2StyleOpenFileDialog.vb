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

        'Adds the default database name to the file name combobox drop down list if its not already there
        If DatabaseFilenameCOMBOBOX.Items.Contains(Form1.CurrentDatabaseLABEL.Text) = False Then
            DatabaseFilenameCOMBOBOX.Items.Add(Form1.CurrentDatabaseLABEL.Text)
        End If

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
    Private Sub MoveItemsBUTTON_Click(sender As Object, e As EventArgs) Handles MoveItemsBUTTON.Click
        OpenSavedDatabase()
    End Sub

    'OPEN DATABASE CONTEXT MENU BUTTON PRESS
    Private Sub OpenSelectedDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenSelectedDatabaseToolStripMenuItem.Click
        OpenSavedDatabase()
    End Sub


    'LOADS THE SELECTED DATABASE IF IT EXISTS
    Sub OpenSavedDatabase()
        Try
            If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt") = True Then

                'VERIFYS THE SELECTED is really a diabase data file by looking for the exact "--------------------" item spacer (20 x -) on the first line
                Dim FileVerify = My.Computer.FileSystem.OpenTextFileReader(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt")
                If FileVerify.ReadLine <> "--------------------" Then
                    'VERIFICATION FAILED - CLOSE FORM
                    Mymessages = "DiaBase File Verification Error." : MyMessageBox()
                    FileVerify.Close() : Me.Close()
                End If

                FileVerify.Close()
                'Checks open destination database checkbox then save this database and load the destionation database if nessicary
                If SaveBeforeOpeningCHECKBOX.Checked = True Then

                    Databasefile = Application.StartupPath + "\Database\" + Form1.CurrentDatabaseLABEL.Text + ".txt"
                    Form1.SaveItems() 'Branch to save routine to save the current database first

                    'Clean out old items from the last loaded database
                    Form1.SearchLISTBOX.Items.Clear()
                    Form1.RichTextBox3.Text = Nothing
                    Form1.ClearStats()

                    'Branch to routine to load the selected dbase and display first item and set current database label in top right of form1
                    OpenDatabaseRoutine(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt")
                    Form1.Display_Items()
                    Form1.CurrentDatabaseLABEL.Text = DatabaseFilenameCOMBOBOX.Text

                    'after successfull save this puts the database name into the combobox dropdown menu if its not already there
                    If DatabaseFilenameCOMBOBOX.Items.Contains(DatabaseFilenameCOMBOBOX.Text) = False Then DatabaseFilenameCOMBOBOX.Items.Add(DatabaseFilenameCOMBOBOX.Text)
                    Me.Close() ' close form once file has been verified

                End If
            Else
                'If the entered file does not exist
                DatabaseFilenameCOMBOBOX.Text = Nothing
            End If

            'catch errors exception handler
        Catch ex As Exception
            Mymessages = "File Error Opening Saved Database." : MyMessageBox()
        End Try
    End Sub

    Private Sub RenameSelectedDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameSelectedDatabaseToolStripMenuItem.Click

        UserInputForm.Text = "Rename The Selected Database"
        UserInputForm.UserInputHeaderLABEL.Text = "ENTER NEW DATABASE NAME"
        UserInputForm.UserImputMessageLABEL.Text = "To rename the selected database please enter its new unique name into the text box"
        UserInputForm.UserInputYesBUTTON.Text = "Rename"
        UserInputForm.UserInputNoBUTTON.Text = "Cancel"

        DialogResult = UserInputForm.ShowDialog

        If DialogResult = Windows.Forms.DialogResult.Yes Then
            'RENAMES THE DATABASE HERE
            Try
                If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt") = True Then
                    My.Computer.FileSystem.RenameFile(Application.StartupPath + "\Database\" + DatabaseFilenameCOMBOBOX.Text + ".txt", UserInputForm.UserInputTEXTBOX.Text + ".txt")

                End If

                'Update default database setting if its the one that was just renamed - STILL TO DO!

                'Update current database label if its the one that was just renamed - STILL TO DO!

            Catch ex As Exception
                Mymessages = "File Error Renaming Saved Database." : MyMessageBox()
            End Try
        Else
        End If
    End Sub
End Class