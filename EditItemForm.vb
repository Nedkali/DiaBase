Public Class EditItemForm

    Private Sub EditItemForm_Load(sender As Object, e As EventArgs) Handles Me.Load

        iEdit = Form1.AllItemsInDatabaseListBox.SelectedIndex

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
        EditItemMulePassCOMBOBOX.Text = Objects(iEdit).MulePass
        EditItemPickitBotCOMBOBOX.Text = Objects(iEdit).PickitBot



    End Sub

    'Clear form some items need correcting still TODO
    Private Sub Clear_Form()
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
        If Settings.BackupOnEditsCHECKBOX.Checked = True Then BackupDatabase()





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

        Me.Close()

    End Sub

    Private Sub EditItemImageBUTTON_Click(sender As Object, e As EventArgs) Handles EditItemImageBUTTON.Click
        ItemImageSelector.ShowDialog()
    End Sub
End Class