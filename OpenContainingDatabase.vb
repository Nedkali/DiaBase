Public Class OpenContainingDatabase

   
    Private Sub YesNoCONFIRMBUTTON_Click(sender As Object, e As EventArgs) Handles YesNoCONFIRMBUTTON.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

        'check if save database checkbox is checked
     

        Databasefile = Application.StartupPath & "\DataBase\" & Form1.CurrentDatabaseLABEL.Text & ".txt"
        If BackupFirstCHECKBOX.Checked = True Then Module1.BackupDatabase()
        If SaveFirstCHECKBOX.Checked = True Then Form1.SaveItems()


        'Now open the database containing the selected item (only in user list)
        'First check file exists
        Try

            If My.Computer.FileSystem.FileExists(Application.StartupPath + "\DataBase\" & UserListDatabase(Form1.UserLISTBOX.SelectedIndex).DatabaseFileName + ".txt") = True Then

                ReadData(Application.StartupPath + "\DataBase\" + UserListDatabase(Form1.UserLISTBOX.SelectedIndex).DatabaseFileName + ".txt", False)
                Form1.Display_Items()

                Databasefile = Application.StartupPath & "\DataBase\" & UserListDatabase(Form1.UserLISTBOX.SelectedIndex).DatabaseFileName + ".txt"
                Form1.CurrentDatabaseLABEL.Text = UserListDatabase(Form1.UserLISTBOX.SelectedIndex).DatabaseFileName


                'Noew set everything back up (listboxes, selected tabs, context menues, ect) for the newly opened database...
                
                Form1.ListboxTABCONTROL.SelectTab(0)
                Form1.ListControlTabBUTTON.BackColor = Color.DimGray
                Form1.SearchListControlTabBUTTON.BackColor = Color.Black
                Form1.TradesListControlTabBUTTON.BackColor = Color.Black
                Form1.UserRefControlTabBUTTON.BackColor = Color.Black
                Form1.ItemTallyTEXTBOX.Text = Form1.AllItemsInDatabaseListBox.Items.Count & " - Total Items"

                'ATTEMPTS TO MATCH THE USER LIST ITEM WITH THE SELECTED MAIN LIST ITEM AFTER OPENING CONTAINING DBASE / FUNCTION NEEDS LOTS OF WORK
                Form1.AllItemsInDatabaseListBox.SelectedIndex = -1
                Form1.AllItemsInDatabaseListBox.SelectedIndex = LocateUserListItem(Form1.UserLISTBOX.SelectedIndex) 'FUNCTION IS IN MODULE1

                'ELSE IF A MATCH CANT BE FOUND THEN DEFAULT SELECT THE FIRST ITEM IF THERE IS ONE ELSE SELECT NOTHING (MUST MAKE ALL CONTEXT FUNCTIONS SELECTION SENSITIVE!!!!)
                If Form1.AllItemsInDatabaseListBox.SelectedIndex = -1 And Form1.AllItemsInDatabaseListBox.Items.Count > 0 Then Form1.AllItemsInDatabaseListBox.SelectedIndex = 0

                'enables or disables the open containing database function is user list contet menu 
                If Form1.UserLISTBOX.SelectedItems.Count = 1 Then
                    If UserListDatabase(Form1.UserLISTBOX.SelectedIndex).DatabaseFileName <> Form1.CurrentDatabaseLABEL.Text Then Form1.OpenContainingDatabaseToolStripMenuItem.Enabled = True Else Form1.OpenContainingDatabaseToolStripMenuItem.Enabled = False
                Else
                    Form1.OpenContainingDatabaseToolStripMenuItem.Enabled = False
                End If

                'enables or disables send to trade list depending if current database holds the selected item
                If Form1.UserLISTBOX.SelectedIndex <> -1 Then If Form1.CurrentDatabaseLABEL.Text = UserListDatabase(Form1.UserLISTBOX.SelectedIndex).DatabaseFileName Then Form1.SendToTradeListToolStripMenuItem1.Enabled = True Else Form1.SendToTradeListToolStripMenuItem1.Enabled = False

                'enables or disables delete depending if current database holds the selected item
                If Form1.UserLISTBOX.SelectedIndex <> -1 Then If Form1.CurrentDatabaseLABEL.Text = UserListDatabase(Form1.UserLISTBOX.SelectedIndex).DatabaseFileName Then Form1.DeleteItemsToolStripMenuItem1.Enabled = True Else Form1.DeleteItemsToolStripMenuItem1.Enabled = False

                'enables or disables export items depending if current database holds the selected item
                If Form1.UserLISTBOX.SelectedIndex <> -1 Then If Form1.CurrentDatabaseLABEL.Text = UserListDatabase(Form1.UserLISTBOX.SelectedIndex).DatabaseFileName Then Form1.ExportItemsToolStripMenuItem.Enabled = True Else Form1.ExportItemsToolStripMenuItem.Enabled = False

                'hide containing database field
                Form1.ItemsDatabaseHeadingTEXTBOX.Hide()
                Form1.ItemsDatabaseFileNameTEXTBOX.Hide()

            End If

        Catch ex As Exception
            'when shit hits fan - app can continue from this ok
            Mymessages = "Cannot Open The Containing Database." : MyMessageBox()

        End Try

    End Sub
End Class