Imports System.IO
'EXPORTS SELECTED ITEMS IN THE SEARCH LIST TO A SECOND DATABASE
Public Class MoveItems
    'Close move item Form
    Private Sub NewDatabaseCancelBUTTON_Click(sender As Object, e As EventArgs) Handles MoveItemsCancelBUTTON.Click
        Me.Close()
    End Sub

    'setup move items form - load form event handler
    Private Sub MoveItems_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DatabaseFilenameTEXTBOX.Text = Nothing
        SavedDatabasesLISTBOX.Items.Clear()
        Dim AllSavedDatabasesFileNames As Array = Nothing

        'GETS SAVED DATABASE FILES FROM DATBASE DIRECTORY AND USES THEM TO POPULATE LISTBOX
        AllSavedDatabasesFileNames = Directory.GetFiles(Application.StartupPath + "\Database\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

        For Each item In AllSavedDatabasesFileNames
            SavedDatabasesLISTBOX.Items.Add(Replace(item, ".txt", ""))
        Next
    End Sub

    'auto selects the textbox and enters listbox selection into textbox as its selected in the listbox (keeps textbox always selected)
    Private Sub SavedDatabasesLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SavedDatabasesLISTBOX.SelectedIndexChanged
        DatabaseFilenameTEXTBOX.Text = SavedDatabasesLISTBOX.SelectedItem
        DatabaseFilenameTEXTBOX.Select()
        DatabaseFilenameTEXTBOX.SelectionLength = 0
    End Sub

    'Stops list box from selecting when nothing in the textbox is selected (keeps textbox always selected)
    Private Sub SavedDatabasesLISTBOX_MouseDown(sender As Object, e As MouseEventArgs) Handles SavedDatabasesLISTBOX.MouseDown
        DatabaseFilenameTEXTBOX.Select()
        DatabaseFilenameTEXTBOX.SelectionLength = 0
    End Sub

    Private Sub MoveItemsBUTTON_Click(sender As Object, e As EventArgs) Handles MoveItemsExportBUTTON.Click

        'Just incase noobs include path and or file extension this should successfully take them back out
        My.Computer.FileSystem.GetName(DatabaseFilenameTEXTBOX.Text)
        DatabaseFilenameTEXTBOX.Text = Replace(DatabaseFilenameTEXTBOX.Text, ".txt", "")

        'Open destination database ready for appending selected items to it.
        If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Database\" + DatabaseFilenameTEXTBOX.Text + ".txt") = False Then

            'creates a new database file if it doesnt already exist
            Dim NewDatabase = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath + "\Database\" + DatabaseFilenameTEXTBOX.Text + ".txt", False)
            NewDatabase.Close()
        End If

        'Writes selected files onto the end of the destination database
        Try
            Dim AppendDatabase = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath + "\Database\" + DatabaseFilenameTEXTBOX.Text + ".txt", True)

            For Each item In Form1.SearchLISTBOX.SelectedIndices
                AppendDatabase.WriteLine("--------------------")
                AppendDatabase.WriteLine(Objects(item).ItemName)
                AppendDatabase.WriteLine(Objects(item).ItemBase)
                AppendDatabase.WriteLine(Objects(item).ItemQuality)
                AppendDatabase.WriteLine(Objects(item).RequiredCharacter)
                AppendDatabase.WriteLine(Objects(item).EtherealItem)
                AppendDatabase.WriteLine(Objects(item).Sockets)
                AppendDatabase.WriteLine(Objects(item).RuneWord)
                AppendDatabase.WriteLine(Objects(item).ThrowDamageMin)
                AppendDatabase.WriteLine(Objects(item).ThrowDamageMax)
                AppendDatabase.WriteLine(Objects(item).OneHandDamageMin)
                AppendDatabase.WriteLine(Objects(item).OneHandDamageMax)
                AppendDatabase.WriteLine(Objects(item).TwoHandDamageMin)
                AppendDatabase.WriteLine(Objects(item).TwoHandDamageMax)
                AppendDatabase.WriteLine(Objects(item).Defense)
                AppendDatabase.WriteLine(Objects(item).ChanceToBlock)
                AppendDatabase.WriteLine(Objects(item).QuantityMin)
                AppendDatabase.WriteLine(Objects(item).QuantityMax)
                AppendDatabase.WriteLine(Objects(item).DurabilityMin)
                AppendDatabase.WriteLine(Objects(item).DurabilityMax)
                AppendDatabase.WriteLine(Objects(item).RequiredStrength)
                AppendDatabase.WriteLine(Objects(item).RequiredDexterity)
                AppendDatabase.WriteLine(Objects(item).RequiredLevel)
                AppendDatabase.WriteLine(Objects(item).AttackClass)
                AppendDatabase.WriteLine(Objects(item).AttackSpeed)
                AppendDatabase.WriteLine(Objects(item).Stat1)
                AppendDatabase.WriteLine(Objects(item).Stat2)
                AppendDatabase.WriteLine(Objects(item).Stat3)
                AppendDatabase.WriteLine(Objects(item).Stat4)
                AppendDatabase.WriteLine(Objects(item).Stat5)
                AppendDatabase.WriteLine(Objects(item).Stat6)
                AppendDatabase.WriteLine(Objects(item).Stat7)
                AppendDatabase.WriteLine(Objects(item).Stat8)
                AppendDatabase.WriteLine(Objects(item).Stat9)
                AppendDatabase.WriteLine(Objects(item).Stat10)
                AppendDatabase.WriteLine(Objects(item).Stat11)
                AppendDatabase.WriteLine(Objects(item).Stat12)
                AppendDatabase.WriteLine(Objects(item).Stat13)
                AppendDatabase.WriteLine(Objects(item).Stat14)
                AppendDatabase.WriteLine(Objects(item).Stat15)
                AppendDatabase.WriteLine(Objects(item).MuleName)
                AppendDatabase.WriteLine(Objects(item).MuleAccount)
                AppendDatabase.WriteLine(Objects(item).MulePass)
                AppendDatabase.WriteLine(Objects(item).PickitBot)
                AppendDatabase.WriteLine(Objects(item).UserReference)
                AppendDatabase.WriteLine(Objects(item).ItemImage)

            Next
            AppendDatabase.Close()

            'Error trapping exception for when shit hits fan
        Catch ex As Exception
            Mymessages = "Error Exporting Items. Restore Backup May Be Required" : MyMessageBox()
        End Try

        'checks delete items checkbox and delete selected items if nessicary
        If DeleteItemsCHECKBOX.Checked = False Then
            Form1.DeleteSearchItems()
        End If

        'checks open destination database checkbox then save this database and load the destionation database if nessicary
        If OpenDatabaseCHECKBOX.Checked = True Then
            Form1.SaveItems() 'branch to save routine to save source dbase before loading destination dbase NOTE: I MAY PUT A CHECKBOX IN FOR THIS

            'clean out old items from the last loaded database
            Form1.SearchLISTBOX.Items.Clear()
            Form1.RichTextBox3.Text = Nothing
            Form1.ClearStats()

            'load up destination one and display first item and set current database label in top right of form1
            Databasefile = Application.StartupPath + "\Database\" + DatabaseFilenameTEXTBOX.Text + ".txt"
            OpenDatabaseRoutine(Databasefile)
            Form1.Display_Items()
            Form1.CurrentDatabaseLABEL.Text = DatabaseFilenameTEXTBOX.Text

            Form1.ListboxTABCONTROL.SelectTab(0)
            Form1.ListControlTabBUTTON.BackColor = Color.DimGray
            Form1.SearchListControlTabBUTTON.BackColor = Color.Black
            Form1.TradesListControlTabBUTTON.BackColor = Color.Black
            Form1.ItemTallyTEXTBOX.Text = Form1.AllItemsInDatabaseListBox.Items.Count & " - Total Items"

        End If
        Me.Close()
    End Sub
End Class