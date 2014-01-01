Public Class AddItemForm

    Private Sub AddItemClearAllBUTTON_Click(sender As Object, e As EventArgs) Handles AddItemClearAllBUTTON.Click
        ClearAddForm()

    End Sub

    'HANDLES [TAB] AND [ENTER] Controlls tabbing textbox foucus for Add form
    Private Sub AllTextBoxes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles AddItemNameTEXTBOX.KeyDown, AddItemBaseCOMBOBOX.KeyDown, AddItemQualityCOMBOBOX.KeyDown, _
       AddItemSocketsCOMBOBOX.KeyDown, AddItemRunewordTEXTBOX.KeyDown, AddItemEtherealItemCHECKBOX.KeyDown, AddItemThrowDamageMinTEXTBOX.KeyDown, AddItemThrowDamageMaxTEXTBOX.KeyDown, AddItemOneHandDamageMinTEXTBOX.KeyDown, AddItemOneHandDamageMaxTEXTBOX.KeyDown, _
       AddItemTwoHandDamageMinTEXTBOX.KeyDown, AddItemTwoHandDamageMaxTEXTBOX.KeyDown, AddItemDefenseTEXTBOX.KeyDown, AddItemChanceToBlockTEXTBOX.KeyDown, _
       AddItemQuantityMinTEXTBOX.KeyDown, AddItemQuantityMaxTEXTBOX.KeyDown, AddItemDurabilityMinTEXTBOX.KeyDown, AddItemDurabilityMaxTEXTBOX.KeyDown, AddItemRequiredStrengthTEXTBOX.KeyDown, AddItemRequiredDexterityTEXTBOX.KeyDown, AddItemRequiredLevelTEXTBOX.KeyDown, AddItemAttackClassCOMBOBOX.KeyDown, AddItemAttackSpeedCOMBOBOX.KeyDown, AddItemStat1TEXTBOX.KeyDown, _
       AddItemStat2TEXTBOX.KeyDown, AddItemStat3TEXTBOX.KeyDown, AddItemStat4TEXTBOX.KeyDown, AddItemStat5TEXTBOX.KeyDown, AddItemStat6TEXTBOX.KeyDown, AddItemStat7TEXTBOX.KeyDown, _
       AddItemStat8TEXTBOX.KeyDown, AddItemStat9TEXTBOX.KeyDown, AddItemStat10TEXTBOX.KeyDown, AddItemStat11TEXTBOX.KeyDown, AddItemStat12TEXTBOX.KeyDown, AddItemMuleNameCOMBOBOX.KeyDown, _
       AddItemStat13TEXTBOX.KeyDown, AddItemStat14TEXTBOX.KeyDown, AddItemStat15TEXTBOX.KeyDown, _
       AddItemMuleAccountCOMBOBOX.KeyDown, AddItemMulePassCOMBOBOX.KeyDown, AddItemUserReferenceTEXTBOX.KeyDown, AddItemPickitBotCOMBOBOX.KeyDown

        'Checks For TAB Keypress and Switches Focus Up (SHIFT-TAB) and Down (TAB) Through Textboxes
        If e.KeyCode = Keys.Tab Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(True)
        End If
    End Sub


    'PUTS ALL EDITED QUICK SELECT ITEMS INTO EDIT COMBOBOX DROPDOWS 
    Sub PopulateDropDowns()
        'Dim Count As Integer = 0 'Record tracking var

        'CLEARS OUT OLD DROPDOWN ITEMS
        'AddItemMuleNameCOMBOBOX.Items.Clear()
        ' AddItemMuleAccountCOMBOBOX.Items.Clear()
        'AddItemMulePassCOMBOBOX.Items.Clear()
        'AddItemPickitBotCOMBOBOX.Items.Clear()
        'AddItemBaseCOMBOBOX.Items.Clear()
        'AddItemAttackClassCOMBOBOX.Items.Clear()


        'Checks if the item is included it the pulldown already, if not it adds it
        'For Each Type In ItemObjects(Count).ItemName
        'If AddItemMuleNameCOMBOBOX.Items.Contains(MuleName(Count)) = False Then AddItemMuleNameCOMBOBOX.Items.Add(MuleName(Count))
        'If AddItemMuleAccountCOMBOBOX.Items.Contains(MuleAccount(Count)) = False Then AddItemMuleAccountCOMBOBOX.Items.Add(MuleAccount(Count))
        'If AddItemMulePassCOMBOBOX.Items.Contains(MulePass(Count)) = False Then AddItemMulePassCOMBOBOX.Items.Add(MulePass(Count))
        'If AddItemPickitAreaCOMBOBOX.Items.Contains(PickitArea(Count)) = False Then AddItemPickitAreaCOMBOBOX.Items.Add(PickitArea(Count))
        'If AddItemPickitBotCOMBOBOX.Items.Contains(PickitBot(Count)) = False Then AddItemPickitBotCOMBOBOX.Items.Add(PickitBot(Count))
        'If AddItemBaseCOMBOBOX.Items.Contains(ItemClass(Count)) = False Then AddItemBaseCOMBOBOX.Items.Add(ItemClass(Count))
        'If AddItemAttackClassCOMBOBOX.Items.Contains(AttackClass(Count)) = False Then AddItemAttackClassCOMBOBOX.Items.Add(AttackClass(Count))
        'Count = Count + 1
        ' Next

        'REMOVES ANY POSSIBLE NULL ENTRIES FROM EDIT DROPDOWNS (Nulls are used to hold a field value that is not in use by the item)
        'If AddItemMuleAccountCOMBOBOX.Items.Contains("") = True Then AddItemMuleAccountCOMBOBOX.Items.Remove("")
        'If AddItemMuleNameCOMBOBOX.Items.Contains("") = True Then AddItemMuleNameCOMBOBOX.Items.Remove("")
        'If AddItemMulePassCOMBOBOX.Items.Contains("") = True Then AddItemMulePassCOMBOBOX.Items.Remove("")
        'If AddItemPickitBotCOMBOBOX.Items.Contains("") = True Then AddItemPickitBotCOMBOBOX.Items.Remove("")
        'If AddItemAttackClassCOMBOBOX.Items.Contains("") = True Then AddItemAttackClassCOMBOBOX.Items.Remove("")
        'If AddItemBaseCOMBOBOX.Items.Contains("") = True Then AddItemBaseCOMBOBOX.Items.Remove("")
    End Sub



    Private Sub ClearAddForm()
        AddItemNameTEXTBOX.Clear()
        AddItemBaseCOMBOBOX.Text = ""
        AddItemQualityCOMBOBOX.Text = ""
        AddItemRunewordTEXTBOX.Clear()
        AddItemSocketsCOMBOBOX.Text = ""
        AddItemEtherealItemCHECKBOX.Checked = False
        AddItemThrowDamageMinTEXTBOX.Clear()
        AddItemThrowDamageMaxTEXTBOX.Clear()
        AddItemOneHandDamageMinTEXTBOX.Clear()
        AddItemOneHandDamageMaxTEXTBOX.Clear()
        AddItemTwoHandDamageMinTEXTBOX.Clear()
        AddItemTwoHandDamageMaxTEXTBOX.Clear()
        AddItemQuantityMinTEXTBOX.Clear()
        AddItemQuantityMaxTEXTBOX.Clear()
        AddItemDurabilityMinTEXTBOX.Clear()
        AddItemDurabilityMaxTEXTBOX.Clear()
        AddItemDefenseTEXTBOX.Clear()
        AddItemChanceToBlockTEXTBOX.Clear()
        AddItemRequiredStrengthTEXTBOX.Clear()
        AddItemRequiredDexterityTEXTBOX.Clear()
        AddItemRequiredLevelTEXTBOX.Clear()
        AddItemAttackClassCOMBOBOX.Text = ""
        AddItemAttackSpeedCOMBOBOX.Text = ""
        AddItemImageTEXTBOX.Clear()
        AddItemUserReferenceTEXTBOX.Clear()
        AddItemStat1TEXTBOX.Clear()
        AddItemStat2TEXTBOX.Clear()
        AddItemStat3TEXTBOX.Clear()
        AddItemStat4TEXTBOX.Clear()
        AddItemStat5TEXTBOX.Clear()
        AddItemStat6TEXTBOX.Clear()
        AddItemStat7TEXTBOX.Clear()
        AddItemStat8TEXTBOX.Clear()
        AddItemStat9TEXTBOX.Clear()
        AddItemStat10TEXTBOX.Clear()
        AddItemStat11TEXTBOX.Clear()
        AddItemStat12TEXTBOX.Clear()
        AddItemStat13TEXTBOX.Clear()
        AddItemStat14TEXTBOX.Clear()
        AddItemStat15TEXTBOX.Clear()
        AddItemMuleNameCOMBOBOX.Text = ""
        AddItemMuleAccountCOMBOBOX.Text = ""
        AddItemMulePassCOMBOBOX.Text = ""
        AddItemPickitBotCOMBOBOX.Text = ""

    End Sub
    Private Sub AddItemCancelBUTTON_Click(sender As Object, e As EventArgs) Handles AddItemCancelBUTTON.Click
        Me.Close()

    End Sub

    Private Sub AddItemSaveBUTTON_Click(sender As Object, e As EventArgs) Handles AddItemSaveBUTTON.Click
        If AddItemNameTEXTBOX.Text = "" Then
            Mymessages = "Must enter an Item Name" : MyMessageBox()
            Return
        End If

        Dim count As Integer = Objects.Count
        Dim AddObject As New ItemObjects

        AddObject.ItemName = AddItemNameTEXTBOX.Text
        AddObject.ItemType = AddItemBaseCOMBOBOX.Text
        AddObject.ItemClass = AddItemQualityCOMBOBOX.Text
        AddObject.RuneWord = AddItemRunewordTEXTBOX.Text
        AddObject.Sockets = AddItemSocketsCOMBOBOX.Text
        If AddItemEtherealItemCHECKBOX.Checked = False Then AddObject.EtherealItem = False
        If AddItemEtherealItemCHECKBOX.Checked = True Then AddObject.EtherealItem = True
        AddObject.ThrowDamageMin = AddItemThrowDamageMinTEXTBOX.Text
        AddObject.ThrowDamageMax = AddItemThrowDamageMaxTEXTBOX.Text
        AddObject.OneHandDamageMin = AddItemOneHandDamageMinTEXTBOX.Text
        AddObject.OneHandDamageMax = AddItemOneHandDamageMaxTEXTBOX.Text
        AddObject.TwoHandDamageMin = AddItemTwoHandDamageMinTEXTBOX.Text
        AddObject.TwoHandDamageMax = AddItemTwoHandDamageMaxTEXTBOX.Text
        AddObject.QuantityMin = AddItemQuantityMinTEXTBOX.Text
        AddObject.QuantityMax = AddItemQuantityMaxTEXTBOX.Text
        AddObject.DurabilityMin = AddItemDurabilityMinTEXTBOX.Text
        AddObject.DurabilityMax = AddItemDurabilityMaxTEXTBOX.Text
        AddObject.Defense = AddItemDefenseTEXTBOX.Text
        AddObject.ChanceToBlock = AddItemChanceToBlockTEXTBOX.Text
        AddObject.RequiredStrength = AddItemRequiredStrengthTEXTBOX.Text
        AddObject.RequiredDexterity = AddItemRequiredDexterityTEXTBOX.Text
        AddObject.RequiredLevel = AddItemRequiredLevelTEXTBOX.Text
        AddObject.AttackClass = AddItemAttackClassCOMBOBOX.ValueMember
        AddObject.AttackSpeed = AddItemAttackSpeedCOMBOBOX.ValueMember
        AddObject.ItemImage = AddItemImageTEXTBOX.Text
        AddObject.UserReference = AddItemUserReferenceTEXTBOX.Text
        AddObject.Stat1 = AddItemStat1TEXTBOX.Text
        AddObject.Stat2 = AddItemStat2TEXTBOX.Text
        AddObject.Stat3 = AddItemStat3TEXTBOX.Text
        AddObject.Stat4 = AddItemStat4TEXTBOX.Text
        AddObject.Stat5 = AddItemStat5TEXTBOX.Text
        AddObject.Stat6 = AddItemStat6TEXTBOX.Text
        AddObject.Stat7 = AddItemStat7TEXTBOX.Text
        AddObject.Stat8 = AddItemStat8TEXTBOX.Text
        AddObject.Stat9 = AddItemStat9TEXTBOX.Text
        AddObject.Stat10 = AddItemStat10TEXTBOX.Text
        AddObject.Stat11 = AddItemStat11TEXTBOX.Text
        AddObject.Stat12 = AddItemStat12TEXTBOX.Text
        AddObject.Stat13 = AddItemStat13TEXTBOX.Text
        AddObject.Stat14 = AddItemStat14TEXTBOX.Text
        AddObject.Stat15 = AddItemStat15TEXTBOX.Text
        AddObject.MuleName = AddItemMuleNameCOMBOBOX.Text
        AddObject.MuleAccount = AddItemMuleAccountCOMBOBOX.Text
        AddObject.MulePass = AddItemMulePassCOMBOBOX.Text
        AddObject.PickitBot = AddItemPickitBotCOMBOBOX.Text

        If AddObject.ItemImage = "" Then AddObject.ItemImage = 0 'causes app crash if null
        Objects.Add(AddObject)
        Form1.AllItemsInDatabaseListBox.Items.Add(Objects(count).ItemName)

        If AddAnotherCHECKBOX.Checked = False Then Me.Close()
        ClearAddForm()

    End Sub

    'Load Add Form Event Handler - Initiates every time add form is Loaded
    Private Sub AddItemForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulateDropDowns()
    End Sub

    Private Sub AddItemImageBUTTON_Click(sender As Object, e As EventArgs) Handles AddItemImageBUTTON.Click
        ItemImageSelector.ShowDialog()
    End Sub
End Class