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
        If DatabaseFilenameTEXTBOX.Text <> "" Then
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
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).ItemName)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).ItemBase)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).ItemQuality)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).RequiredCharacter)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).EtherealItem)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Sockets)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).RuneWord)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).ThrowDamageMin)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).ThrowDamageMax)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).OneHandDamageMin)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).OneHandDamageMax)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).TwoHandDamageMin)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).TwoHandDamageMax)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Defense)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).ChanceToBlock)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).QuantityMin)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).QuantityMax)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).DurabilityMin)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).DurabilityMax)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).RequiredStrength)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).RequiredDexterity)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).RequiredLevel)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).AttackClass)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).AttackSpeed)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat1)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat2)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat3)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat4)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat5)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat6)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat7)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat8)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat9)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat10)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat11)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat12)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat13)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat14)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).Stat15)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).MuleName)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).MuleAccount)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).MulePass)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).PickitBot)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).UserReference)
                    AppendDatabase.WriteLine(Objects(SearchReferenceList(item)).ItemImage)

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
                SearchReferenceList.Clear()
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

                'ensure dabase filename is hidden for focusing on main list box
                Form1.ItemsDatabaseHeadingTEXTBOX.Hide()
                Form1.ItemsDatabaseFileNameTEXTBOX.Hide()


            End If
            Me.Close()
        Else
            Mymessages = "You Must Supply A Database Name" : MyMessageBox()
        End If
    End Sub
End Class