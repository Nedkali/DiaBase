Public Class EditItemForm
    'Sub AddEditItemComboboxItems(ByVal ItemIndexNumberToRefrence)
    'End Sub
    Public EditedFields As List(Of String) = New List(Of String)
    Public UpdatingField As Boolean = False

    Private Sub EditItemForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        iEdit = Form1.AllItemsInDatabaseListBox.SelectedIndex
        EditedFields.Clear()

        'POPULATE EDIT COMBOBOXES
        UpdatingField = True
        EditItemMuleNameCOMBOBOX.Items.Clear() 'Next 5 lines clears out all existing combobox items
        EditItemMuleAccountCOMBOBOX.Items.Clear()
        EditItemMulePassCOMBOBOX.Items.Clear()
        EditItemAttackClassCOMBOBOX.Items.Clear()
        EditItemBaseCOMBOBOX.Items.Clear()
        EditItemQualityCOMBOBOX.Items.Clear()
        EditItemPickitBotCOMBOBOX.Items.Clear()

        For Each ItemObjectItem As ItemObjects In Objects
            If ItemObjectItem.MuleName <> "" Then
                If EditItemMuleNameCOMBOBOX.Items.Contains(ItemObjectItem.MuleName) = False Then EditItemMuleNameCOMBOBOX.Items.Add(ItemObjectItem.MuleName)
            End If
            If ItemObjectItem.MuleAccount <> "" Then
                If EditItemMuleAccountCOMBOBOX.Items.Contains(ItemObjectItem.MuleAccount) = False Then EditItemMuleAccountCOMBOBOX.Items.Add(ItemObjectItem.MuleAccount)
            End If
            If ItemObjectItem.MulePass <> "" Then
                If EditItemMulePassCOMBOBOX.Items.Contains(ItemObjectItem.MulePass) = False Then EditItemMulePassCOMBOBOX.Items.Add(ItemObjectItem.MulePass)
            End If
            If ItemObjectItem.AttackClass <> "" Then
                If EditItemAttackClassCOMBOBOX.Items.Contains(ItemObjectItem.AttackClass) = False Then EditItemAttackClassCOMBOBOX.Items.Add(ItemObjectItem.AttackClass)
            End If
            If ItemObjectItem.ItemBase <> "" Then
                If EditItemBaseCOMBOBOX.Items.Contains(ItemObjectItem.ItemBase) = False Then EditItemBaseCOMBOBOX.Items.Add(ItemObjectItem.ItemBase)
            End If
            If ItemObjectItem.ItemQuality <> "" Then
                If EditItemQualityCOMBOBOX.Items.Contains(ItemObjectItem.ItemQuality) = False Then EditItemQualityCOMBOBOX.Items.Add(ItemObjectItem.ItemQuality)
            End If
            If ItemObjectItem.PickitBot <> "" Then
                If EditItemPickitBotCOMBOBOX.Items.Contains(ItemObjectItem.PickitBot) = False Then EditItemPickitBotCOMBOBOX.Items.Add(ItemObjectItem.PickitBot)
            End If
        Next
        UpdatingField = False
    End Sub

    Private Sub EditItemCancelBUTTON_Click(sender As Object, e As EventArgs) Handles EditItemCancelBUTTON.Click
        Me.Close()
    End Sub

    Private Sub EditItemForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        AutoCompleteForm()
    End Sub

    Private Sub EditItemUndoChangesBUTTON_Click(sender As Object, e As EventArgs) Handles EditItemUndoChangesBUTTON.Click
        Clear_Form()
        AutoCompleteForm()
    End Sub

    'TAB PRESS EVENT HANDLER - Controlls tabbing textbox foucus for edit form
    Private Sub AllTextBoxes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles EditItemNameTEXTBOX.KeyDown, EditItemBaseCOMBOBOX.KeyDown, _
      EditItemEtherealItemCHECKBOX.KeyDown, EditItemSocketsCOMBOBOX.KeyDown, EditItemThrowDamageMinTEXTBOX.KeyDown, EditItemThrowDamageMaxTEXTBOX.KeyDown, EditItemOneHandDamageMaxTEXTBOX.KeyDown, _
     EditItemOneHandDamageMinTEXTBOX.KeyDown, EditItemAttackClassLABEL.KeyDown, EditItemTwoHandDamageMinTEXTBOX.KeyDown, EditItemTwoHandDamageMaxTEXTBOX.KeyDown, EditItemDefenseTEXTBOX.KeyDown, EditItemChanceToBlockTEXTBOX.KeyDown, _
      EditItemQuantityMinTEXTBOX.KeyDown, EditItemQuantityMaxTEXTBOX.KeyDown, EditItemDurabilityMinTEXTBOX.KeyDown, EditItemDurabilityMaxTEXTBOX.KeyDown, EditItemRequiredStrengthTEXTBOX.KeyDown, _
      EditItemRequiredDexterityTEXTBOX.KeyDown, EditItemRequiredLevelTEXTBOX.KeyDown, EditItemAttackClassCOMBOBOX.KeyDown, EditItemAttackSpeedCOMBOBOX.KeyDown, EditItemStat1TEXTBOX.KeyDown, _
      EditItemStat2TEXTBOX.KeyDown, EditItemStat3TEXTBOX.KeyDown, EditItemStat4TEXTBOX.KeyDown, EditItemStat5TEXTBOX.KeyDown, EditItemStat6TEXTBOX.KeyDown, EditItemStat7TEXTBOX.KeyDown, _
      EditItemStat8TEXTBOX.KeyDown, EditItemStat9TEXTBOX.KeyDown, EditItemStat10TEXTBOX.KeyDown, EditItemStat11TEXTBOX.KeyDown, EditItemStat12TEXTBOX.KeyDown, EditItemStat13TEXTBOX.KeyDown, EditItemStat14TEXTBOX.KeyDown, EditItemStat15TEXTBOX.KeyDown, EditItemMuleNameCOMBOBOX.KeyDown, _
      EditItemMuleAccountCOMBOBOX.KeyDown, EditItemMulePassCOMBOBOX.KeyDown, EditItemUserReferenceTEXTBOX.KeyDown, EditItemPickitBotCOMBOBOX.KeyDown

        'checks For Keypress and switches focus up (SHIFT-TAB) and down (TAB) through textboxes
        If e.KeyCode = Keys.Tab Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(True)
        End If
    End Sub






    'Fill out form with current item information
    Private Sub AutoCompleteForm()
        UpdatingField = True
        EditItemNameTEXTBOX.Text = Objects(iEdit).ItemName
        EditItemBaseCOMBOBOX.Text = Objects(iEdit).ItemBase
        EditItemQualityCOMBOBOX.Text = Objects(iEdit).ItemQuality
        EditItemRunewordTEXTBOX.Text = Objects(iEdit).RuneWord
        EditItemSocketsCOMBOBOX.Text = Objects(iEdit).Sockets
        If Objects(iEdit).EtherealItem = True Then EditItemEtherealItemCHECKBOX.Checked = True
        If Objects(iEdit).EtherealItem = False Then EditItemEtherealItemCHECKBOX.Checked = False
        EditItemThrowDamageMinTEXTBOX.Text = Objects(iEdit).ThrowDamageMin
        EditItemThrowDamageMaxTEXTBOX.Text = Objects(iEdit).ThrowDamageMax
        EditItemOneHandDamageMinTEXTBOX.Text = Objects(iEdit).OneHandDamageMin
        EditItemOneHandDamageMaxTEXTBOX.Text = Objects(iEdit).OneHandDamageMax
        EditItemTwoHandDamageMinTEXTBOX.Text = Objects(iEdit).TwoHandDamageMin
        EditItemTwoHandDamageMaxTEXTBOX.Text = Objects(iEdit).TwoHandDamageMax
        EditItemQuantityMinTEXTBOX.Text = Objects(iEdit).QuantityMin
        EditItemQuantityMaxTEXTBOX.Text = Objects(iEdit).QuantityMax
        EditItemDurabilityMinTEXTBOX.Text = Objects(iEdit).DurabilityMin
        EditItemDurabilityMaxTEXTBOX.Text = Objects(iEdit).DurabilityMax
        EditItemDefenseTEXTBOX.Text = Objects(iEdit).Defense
        EditItemChanceToBlockTEXTBOX.Text = Objects(iEdit).ChanceToBlock
        EditItemRequiredStrengthTEXTBOX.Text = Objects(iEdit).RequiredStrength
        EditItemRequiredDexterityTEXTBOX.Text = Objects(iEdit).RequiredDexterity
        EditItemRequiredLevelTEXTBOX.Text = Objects(iEdit).RequiredLevel
        EditItemAttackClassCOMBOBOX.Text = Objects(iEdit).AttackClass
        EditItemAttackSpeedCOMBOBOX.Text = Objects(iEdit).AttackSpeed
        EditItemImageTEXTBOX.Text = Objects(iEdit).ItemImage
        EditItemUserReferenceTEXTBOX.Text = Objects(iEdit).UserReference
        EditItemStat1TEXTBOX.Text = Objects(iEdit).Stat1
        EditItemStat2TEXTBOX.Text = Objects(iEdit).Stat2
        EditItemStat3TEXTBOX.Text = Objects(iEdit).Stat3
        EditItemStat4TEXTBOX.Text = Objects(iEdit).Stat4
        EditItemStat5TEXTBOX.Text = Objects(iEdit).Stat5
        EditItemStat6TEXTBOX.Text = Objects(iEdit).Stat6
        EditItemStat7TEXTBOX.Text = Objects(iEdit).Stat7
        EditItemStat8TEXTBOX.Text = Objects(iEdit).Stat8
        EditItemStat9TEXTBOX.Text = Objects(iEdit).Stat9
        EditItemStat10TEXTBOX.Text = Objects(iEdit).Stat10
        EditItemStat11TEXTBOX.Text = Objects(iEdit).Stat11
        EditItemStat12TEXTBOX.Text = Objects(iEdit).Stat12
        EditItemStat13TEXTBOX.Text = Objects(iEdit).Stat13
        EditItemStat14TEXTBOX.Text = Objects(iEdit).Stat14
        EditItemStat15TEXTBOX.Text = Objects(iEdit).Stat15
        EditItemMuleNameCOMBOBOX.Text = Objects(iEdit).MuleName
        EditItemMuleAccountCOMBOBOX.Text = Objects(iEdit).MuleAccount
        If KeepPassPrivate = True Then EditItemMulePassCOMBOBOX.Text = HidePass(Objects(iEdit).MulePass) Else EditItemMulePassCOMBOBOX.Text = Objects(iEdit).MulePass
        'EditItemMulePassCOMBOBOX.Text = Objects(iEdit).MulePass
        EditItemPickitBotCOMBOBOX.Text = Objects(iEdit).PickitBot
        UpdatingField = False


    End Sub

    'Clear form some items need correcting still TODO
    Private Sub Clear_Form()
        UpdatingField = True
        EditItemNameTEXTBOX.Clear()
        EditItemBaseCOMBOBOX.Text = ""
        EditItemQualityCOMBOBOX.Text = ""
        EditItemRunewordTEXTBOX.Clear()
        EditItemSocketsCOMBOBOX.Text = ""
        EditItemEtherealItemCHECKBOX.Checked = False
        EditItemThrowDamageMinTEXTBOX.Clear()
        EditItemThrowDamageMaxTEXTBOX.Clear()
        EditItemOneHandDamageMinTEXTBOX.Clear()
        EditItemOneHandDamageMaxTEXTBOX.Clear()
        EditItemTwoHandDamageMinTEXTBOX.Clear()
        EditItemTwoHandDamageMaxTEXTBOX.Clear()
        EditItemQuantityMinTEXTBOX.Clear()
        EditItemQuantityMaxTEXTBOX.Clear()
        EditItemDurabilityMinTEXTBOX.Clear()
        EditItemDurabilityMaxTEXTBOX.Clear()
        EditItemDefenseTEXTBOX.Clear()
        EditItemChanceToBlockTEXTBOX.Clear()
        EditItemRequiredStrengthTEXTBOX.Clear()
        EditItemRequiredDexterityTEXTBOX.Clear()
        EditItemRequiredLevelTEXTBOX.Clear()
        EditItemAttackClassCOMBOBOX.Text = ""
        EditItemAttackSpeedCOMBOBOX.Text = ""
        EditItemImageTEXTBOX.Clear()
        EditItemUserReferenceTEXTBOX.Clear()
        EditItemStat1TEXTBOX.Clear()
        EditItemStat2TEXTBOX.Clear()
        EditItemStat3TEXTBOX.Clear()
        EditItemStat4TEXTBOX.Clear()
        EditItemStat5TEXTBOX.Clear()
        EditItemStat6TEXTBOX.Clear()
        EditItemStat7TEXTBOX.Clear()
        EditItemStat8TEXTBOX.Clear()
        EditItemStat9TEXTBOX.Clear()
        EditItemStat10TEXTBOX.Clear()
        EditItemStat11TEXTBOX.Clear()
        EditItemStat12TEXTBOX.Clear()
        EditItemStat13TEXTBOX.Clear()
        EditItemStat14TEXTBOX.Clear()
        EditItemStat15TEXTBOX.Clear()
        EditItemMuleNameCOMBOBOX.Text = ""
        EditItemMuleAccountCOMBOBOX.Text = ""
        EditItemMulePassCOMBOBOX.Text = ""
        EditItemPickitBotCOMBOBOX.Text = ""
        UpdatingField = False
    End Sub

    'Update object's information based on user input
    Private Sub EditItemSaveChangesBUTTON_Click(sender As Object, e As EventArgs) Handles EditItemSaveChangesBUTTON.Click

        If EditItemNameTEXTBOX.Text = Nothing Then
            Mymessages = "Item must have a Name" : UserMessaging.ShowDialog()
            Clear_Form()
            AutoCompleteForm()
            Return
        End If

        'check for backup on edits set to true if so the backup now
        'UpdatingField = True
        If Settings.BackupOnEditsCHECKBOX.Checked = True Then Module1.BackupDatabase()
        Dim ConfirmResult As Windows.Forms.DialogResult
        Dim FirstSelected As Integer = Nothing
        Dim AllSelected As List(Of String) = New List(Of String)

        If Form1.AllItemsInDatabaseListBox.SelectedItems.Count > 1 And EditedFields.Count > 0 Then

            'get all selected items for reselection at end of process
            For Each item In Form1.AllItemsInDatabaseListBox.SelectedIndices
                AllSelected.Add(item)
            Next


            'confirmation popup box
            YesNoD2Style.Text = "Confirm To Copy Edits To All Selected Items"
            YesNoD2Style.YesNoHeaderLABEL.Text = "Copy Edits To All Selected Items"
            YesNoD2Style.YesNoMessageLABEL.Text = "Select " & Chr(34) & "Yes" & Chr(34) & " to copy edited field values to all currently selected items. " & vbCrLf & vbCrLf & "NOTE: This will permanently replace the edited field values for all other items currently selected." & vbCrLf & vbCrLf & "Select " & Chr(34) & "No" & Chr(34) & " to only edit the first selected item."
            YesNoD2Style.YesNoCANCELBUTTON.Text = "No"
            YesNoD2Style.YesNoCONFIRMBUTTON.Text = "Yes"

            'confirm results...
            ConfirmResult = YesNoD2Style.ShowDialog
            YesNoD2Style.YesNoCANCELBUTTON.Text = "Cancel"
            YesNoD2Style.YesNoCONFIRMBUTTON.Text = "Confirm"
            If ConfirmResult = Windows.Forms.DialogResult.Yes Then

                'apply edits to all selected items
                For count = 0 To Form1.AllItemsInDatabaseListBox.SelectedItems.Count - 1 Step 1
                    If count = 0 Then FirstSelected = Form1.AllItemsInDatabaseListBox.SelectedIndex ' get first selected item for reselection at end of process
                    iEdit = Form1.AllItemsInDatabaseListBox.SelectedIndex + count
                    For Each item In EditedFields
                        If item = "EditItemNameTEXTBOX" Then Objects(iEdit).ItemName = EditItemNameTEXTBOX.Text
                        If item = "EditItemBaseCOMBOBOX" Then Objects(iEdit).ItemBase = EditItemBaseCOMBOBOX.Text
                        If item = "EditItemQualityCOMBOBOX" Then Objects(iEdit).ItemQuality = EditItemQualityCOMBOBOX.Text
                        If item = "EditItemRunewordTEXTBOX" Then Objects(iEdit).RuneWord = EditItemRunewordTEXTBOX.Text
                        If item = "EditItemSocketsCOMBOBOX" Then Objects(iEdit).Sockets = EditItemSocketsCOMBOBOX.Text
                        If item = "EditItemEtherealItemCHECKBOX" And EditItemEtherealItemCHECKBOX.Checked = True Then Objects(iEdit).EtherealItem = "true"
                        If item = "EditItemEtherealItemCHECKBOX" And EditItemEtherealItemCHECKBOX.Checked = False Then Objects(iEdit).EtherealItem = "false"
                        If item = "EditItemThrowDamageMinTEXTBOX" Then Objects(iEdit).ThrowDamageMin = EditItemThrowDamageMinTEXTBOX.Text
                        If item = "EditItemThrowDamageMaxTEXTBOX" Then Objects(iEdit).ThrowDamageMax = EditItemThrowDamageMaxTEXTBOX.Text
                        If item = "EditItemOneHandDamageMinTEXTBOX" Then Objects(iEdit).OneHandDamageMin = EditItemOneHandDamageMinTEXTBOX.Text
                        If item = "EditItemOneHandDamageMaxTEXTBOX" Then Objects(iEdit).OneHandDamageMax = EditItemOneHandDamageMaxTEXTBOX.Text
                        If item = "EditItemTwoHandDamageMinTEXTBOX" Then Objects(iEdit).TwoHandDamageMin = EditItemTwoHandDamageMinTEXTBOX.Text
                        If item = "EditItemTwoHandDamageMaxTEXTBOX" Then Objects(iEdit).TwoHandDamageMax = EditItemTwoHandDamageMaxTEXTBOX.Text
                        If item = "EditItemQuantityMinTEXTBOX" Then Objects(iEdit).QuantityMin = EditItemQuantityMinTEXTBOX.Text
                        If item = "EditItemQuantityMaxTEXTBOX" Then Objects(iEdit).QuantityMax = EditItemQuantityMaxTEXTBOX.Text
                        If item = "EditItemDurabilityMinTEXTBOX" Then Objects(iEdit).DurabilityMin = EditItemDurabilityMinTEXTBOX.Text
                        If item = "EditItemDurabilityMaxTEXTBOX" Then Objects(iEdit).DurabilityMax = EditItemDurabilityMaxTEXTBOX.Text
                        If item = "EditItemDefenseTEXTBOX" Then Objects(iEdit).Defense = EditItemDefenseTEXTBOX.Text
                        If item = "EditItemChanceToBlockTEXTBOX" Then Objects(iEdit).ChanceToBlock = EditItemChanceToBlockTEXTBOX.Text
                        If item = "EditItemRequiredStrengthTEXTBOX" Then Objects(iEdit).RequiredStrength = EditItemRequiredStrengthTEXTBOX.Text
                        If item = "EditItemRequiredDexterityTEXTBOX" Then Objects(iEdit).RequiredDexterity = EditItemRequiredDexterityTEXTBOX.Text
                        If item = "EditItemRequiredLevelTEXTBOX" Then Objects(iEdit).RequiredLevel = EditItemRequiredLevelTEXTBOX.Text
                        If item = "EditItemAttackClassCOMBOBOX" Then Objects(iEdit).AttackClass = EditItemAttackClassCOMBOBOX.Text
                        If item = "EditItemAttackSpeedCOMBOBOX" Then Objects(iEdit).AttackSpeed = EditItemAttackSpeedCOMBOBOX.Text
                        If item = "EditItemStat1TEXTBOX" Then Objects(iEdit).Stat1 = EditItemStat1TEXTBOX.Text
                        If item = "EditItemStat2TEXTBOX" Then Objects(iEdit).Stat2 = EditItemStat2TEXTBOX.Text
                        If item = "EditItemStat3TEXTBOX" Then Objects(iEdit).Stat3 = EditItemStat3TEXTBOX.Text
                        If item = "EditItemStat4TEXTBOX" Then Objects(iEdit).Stat4 = EditItemStat4TEXTBOX.Text
                        If item = "EditItemStat5TEXTBOX" Then Objects(iEdit).Stat5 = EditItemStat5TEXTBOX.Text
                        If item = "EditItemStat6TEXTBOX" Then Objects(iEdit).Stat6 = EditItemStat6TEXTBOX.Text
                        If item = "EditItemStat7TEXTBOX" Then Objects(iEdit).Stat7 = EditItemStat7TEXTBOX.Text
                        If item = "EditItemStat8TEXTBOX" Then Objects(iEdit).Stat8 = EditItemStat8TEXTBOX.Text
                        If item = "EditItemStat9TEXTBOX" Then Objects(iEdit).Stat9 = EditItemStat9TEXTBOX.Text
                        If item = "EditItemStat10TEXTBOX" Then Objects(iEdit).Stat10 = EditItemStat10TEXTBOX.Text
                        If item = "EditItemStat11TEXTBOX" Then Objects(iEdit).Stat11 = EditItemStat11TEXTBOX.Text
                        If item = "EditItemStat12TEXTBOX" Then Objects(iEdit).Stat12 = EditItemStat12TEXTBOX.Text
                        If item = "EditItemStat13TEXTBOX" Then Objects(iEdit).Stat13 = EditItemStat13TEXTBOX.Text
                        If item = "EditItemStat14TEXTBOX" Then Objects(iEdit).Stat14 = EditItemStat14TEXTBOX.Text
                        If item = "EditItemStat15TEXTBOX" Then Objects(iEdit).Stat15 = EditItemStat15TEXTBOX.Text
                        If item = "EditItemMuleNameCOMBOBOX" Then Objects(iEdit).MuleName = EditItemMuleNameCOMBOBOX.Text
                        If item = "EditItemMuleAccountCOMBOBOX" Then Objects(iEdit).MuleAccount = EditItemMuleAccountCOMBOBOX.Text
                        If item = "EditItemMulePassCOMBOBOX" Then Objects(iEdit).MulePass = EditItemMulePassCOMBOBOX.Text
                        If item = "EditItemPickitBotCOMBOBOX" Then Objects(iEdit).PickitBot = EditItemPickitBotCOMBOBOX.Text
                        If item = "EditItemImageTEXTBOX" Then Objects(iEdit).ItemImage = EditItemImageTEXTBOX.Text
                        If item = "EditItemUserReferenceTEXTBOX" Then Objects(iEdit).UserReference = EditItemUserReferenceTEXTBOX.Text
                        'If item = "" Then Objects(iEdit). = .Text
                    Next
                Next

                'reselects items to refresh item stats display
                Form1.AllItemsInDatabaseListBox.SelectedIndex = -1
                For Each item In AllSelected
                    Form1.AllItemsInDatabaseListBox.SetSelected(item, True)
                Next
                'apply edits to only 1 selected it or when only 1 item is selected from here down
            Else
                If ConfirmResult = Windows.Forms.DialogResult.No Then
                    UpdateItemToDatabase()
                    'reselects item to refresh stats display
                    Form1.AllItemsInDatabaseListBox.SelectedIndex = -1
                    For Each item In AllSelected
                        Form1.AllItemsInDatabaseListBox.SetSelected(item, True)
                    Next
                End If
            End If
        Else
            If Form1.AllItemsInDatabaseListBox.SelectedItems.Count = 1 Then
                UpdateItemToDatabase()
                'reselects item to refresh stats display
                Form1.AllItemsInDatabaseListBox.SelectedIndex = -1
                Form1.AllItemsInDatabaseListBox.SelectedIndex = iEdit
            End If
        End If
        Me.Close()
    End Sub

    'Save routine for single selection edits and when edits are only applied to the first selected item
    Sub UpdateItemToDatabase()
        Objects(iEdit).ItemName = EditItemNameTEXTBOX.Text
        Objects(iEdit).ItemBase = EditItemBaseCOMBOBOX.Text
        Objects(iEdit).ItemQuality = EditItemQualityCOMBOBOX.Text
        Objects(iEdit).RuneWord = EditItemRunewordTEXTBOX.Text
        Objects(iEdit).Sockets = EditItemSocketsCOMBOBOX.Text
        If EditItemEtherealItemCHECKBOX.Checked = False Then Objects(iEdit).EtherealItem = False
        If EditItemEtherealItemCHECKBOX.Checked = True Then Objects(iEdit).EtherealItem = True
        Objects(iEdit).ThrowDamageMin = EditItemThrowDamageMinTEXTBOX.Text
        Objects(iEdit).ThrowDamageMax = EditItemThrowDamageMaxTEXTBOX.Text
        Objects(iEdit).OneHandDamageMin = EditItemOneHandDamageMinTEXTBOX.Text
        Objects(iEdit).OneHandDamageMax = EditItemOneHandDamageMaxTEXTBOX.Text
        Objects(iEdit).TwoHandDamageMin = EditItemTwoHandDamageMinTEXTBOX.Text
        Objects(iEdit).TwoHandDamageMax = EditItemTwoHandDamageMaxTEXTBOX.Text
        Objects(iEdit).QuantityMin = EditItemQuantityMinTEXTBOX.Text
        Objects(iEdit).QuantityMax = EditItemQuantityMaxTEXTBOX.Text
        Objects(iEdit).DurabilityMin = EditItemDurabilityMinTEXTBOX.Text
        Objects(iEdit).DurabilityMax = EditItemDurabilityMaxTEXTBOX.Text
        Objects(iEdit).Defense = EditItemDefenseTEXTBOX.Text
        Objects(iEdit).ChanceToBlock = EditItemChanceToBlockTEXTBOX.Text
        Objects(iEdit).RequiredStrength = EditItemRequiredStrengthTEXTBOX.Text
        Objects(iEdit).RequiredDexterity = EditItemRequiredDexterityTEXTBOX.Text
        Objects(iEdit).RequiredLevel = EditItemRequiredLevelTEXTBOX.Text
        Objects(iEdit).AttackClass = EditItemAttackClassCOMBOBOX.ValueMember
        Objects(iEdit).AttackSpeed = EditItemAttackSpeedCOMBOBOX.ValueMember
        Objects(iEdit).ItemImage = EditItemImageTEXTBOX.Text
        Objects(iEdit).UserReference = EditItemUserReferenceTEXTBOX.Text
        Objects(iEdit).Stat1 = EditItemStat1TEXTBOX.Text
        Objects(iEdit).Stat2 = EditItemStat2TEXTBOX.Text
        Objects(iEdit).Stat3 = EditItemStat3TEXTBOX.Text
        Objects(iEdit).Stat4 = EditItemStat4TEXTBOX.Text
        Objects(iEdit).Stat5 = EditItemStat5TEXTBOX.Text
        Objects(iEdit).Stat6 = EditItemStat6TEXTBOX.Text
        Objects(iEdit).Stat7 = EditItemStat7TEXTBOX.Text
        Objects(iEdit).Stat8 = EditItemStat8TEXTBOX.Text
        Objects(iEdit).Stat9 = EditItemStat9TEXTBOX.Text
        Objects(iEdit).Stat10 = EditItemStat10TEXTBOX.Text
        Objects(iEdit).Stat11 = EditItemStat11TEXTBOX.Text
        Objects(iEdit).Stat12 = EditItemStat12TEXTBOX.Text
        Objects(iEdit).Stat13 = EditItemStat13TEXTBOX.Text
        Objects(iEdit).Stat14 = EditItemStat14TEXTBOX.Text
        Objects(iEdit).Stat15 = EditItemStat15TEXTBOX.Text
        Objects(iEdit).MuleName = EditItemMuleNameCOMBOBOX.Text
        Objects(iEdit).MuleAccount = EditItemMuleAccountCOMBOBOX.Text
        Objects(iEdit).MulePass = EditItemMulePassCOMBOBOX.Text
        Objects(iEdit).PickitBot = EditItemPickitBotCOMBOBOX.Text

        Form1.AllItemsInDatabaseListBox.Items.RemoveAt(iEdit)                           'Clears previous item's name from form1 itemlistbox
        Form1.AllItemsInDatabaseListBox.Items.Insert(iEdit, (Objects(iEdit).ItemName))  'Updates new details/name in form1 itemlistbox

    End Sub


    Private Sub EditItemImageBUTTON_Click(sender As Object, e As EventArgs) Handles EditItemImageBUTTON.Click
        ItemImageSelector.ShowDialog()
    End Sub

    Private Sub CheckForFieldChanges(sender As Object, e As EventArgs) Handles EditItemNameTEXTBOX.TextChanged, EditItemBaseCOMBOBOX.TextChanged, EditItemQualityCOMBOBOX.TextChanged, EditItemRunewordTEXTBOX.TextChanged, _
        EditItemSocketsCOMBOBOX.TextChanged, EditItemEtherealItemCHECKBOX.CheckedChanged, EditItemThrowDamageMinTEXTBOX.TextChanged, EditItemThrowDamageMaxTEXTBOX.TextChanged, EditItemOneHandDamageMinTEXTBOX.TextChanged, _
        EditItemOneHandDamageMaxTEXTBOX.TextChanged, EditItemTwoHandDamageMinTEXTBOX.TextChanged, EditItemTwoHandDamageMaxTEXTBOX.TextChanged, EditItemQuantityMinTEXTBOX.TextChanged, EditItemQuantityMaxTEXTBOX.TextChanged, _
        EditItemDurabilityMinTEXTBOX.TextChanged, EditItemDefenseTEXTBOX.TextChanged, EditItemChanceToBlockTEXTBOX.TextChanged, EditItemRequiredStrengthTEXTBOX.TextChanged, EditItemRequiredDexterityTEXTBOX.TextChanged, _
        EditItemRequiredLevelTEXTBOX.TextChanged, EditItemAttackClassCOMBOBOX.TextChanged, EditItemAttackSpeedCOMBOBOX.TextChanged, EditItemStat1TEXTBOX.TextChanged, EditItemStat2TEXTBOX.TextChanged, EditItemStat3TEXTBOX.TextChanged, _
        EditItemStat4TEXTBOX.TextChanged, EditItemStat5TEXTBOX.TextChanged, EditItemStat6TEXTBOX.TextChanged, EditItemStat7TEXTBOX.TextChanged, EditItemStat8TEXTBOX.TextChanged, EditItemStat9TEXTBOX.TextChanged, EditItemStat10TEXTBOX.TextChanged, _
        EditItemStat11TEXTBOX.TextChanged, EditItemStat12TEXTBOX.TextChanged, EditItemStat13TEXTBOX.TextChanged, EditItemStat14TEXTBOX.TextChanged, EditItemStat15TEXTBOX.TextChanged, EditItemMuleNameCOMBOBOX.TextChanged, _
        EditItemMuleAccountCOMBOBOX.TextChanged, EditItemMulePassCOMBOBOX.TextChanged, EditItemPickitBotCOMBOBOX.TextChanged, EditItemImageTEXTBOX.TextChanged, EditItemUserReferenceTEXTBOX.TextChanged

        If UpdatingField = False Then
            Dim ControlName As String = CType(sender, Control).Name.ToString()
            If EditedFields.contains(ControlName) = False Then EditedFields.add(ControlName)
        End If
    End Sub
End Class