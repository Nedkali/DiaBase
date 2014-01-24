Module Search
    Sub SearchRoutineTest()


        If Form1.RefineSearchCHECKBOX.Checked = True Then

            RefineSearchReferenceList.Clear()
            For Each item In SearchReferenceList
                RefineSearchReferenceList.Add(item) ' no idea what this is for - refine what? doesnt seem to do anything
            Next

        End If

        Form1.SearchLISTBOX.Items.Clear()                       'Clear out old search matches
        SearchReferenceList.Clear()                             'Clear Out old Item Search Reference List (Holds location in database of each matched item)

        SearchProgressForm.Show() : SearchProgressForm.SearchPROGRESSBAR.Value = 0

        Select Case Form1.SearchFieldCOMBOBOX.Text
            Case "Item Name"
                SearchItemName()
            Case "Item Base"
                SearchItemBase()
            Case "Item Quality"
                SearchItemQuality()
            Case "Item Defense"
                SearchItemDefense()
            Case "RuneWord"
                SearchRuneWord()
            Case "Chance To Block"
                SearchChanceToBlock()
            Case "One Hand Damage Max"
                SearchOneHandDamageMax()
            Case "One Hand Damage Min"
                SearchOneHandDamageMin()
            Case "Two Hand Damage Max"
                SearchTwoHandDamageMax()
            Case "Two Hand Damage Min"
                SearchTwoHandDamageMin()
            Case "Throw Damage Max"
                SearchThrowDamageMax()
            Case "Throw Damage Min"
                SearchThrowDamageMin()
            Case "Required Level"
                SearchRequiredLevel()
            Case "Required Strength"
                SearchRequiredStrength()
            Case "Required Dexterity"
                SearchRequiredDexterity()
            Case "Attack Class"
                SearchAttackClass()
            Case "Attack Speed"
                SearchAttackSpeed()
            Case "Unique Attributes"
                SearchUniqueAttributes()
            Case "Mule Name"
                SearchMuleName()
            Case "Mule Account"
                SearchMuleAccount()
            Case "Mule Pass"
                SearchMulePass()
            Case "User Reference"
                SearchUserReference()

        End Select
        SearchProgressForm.Close()
    End Sub



    Sub SearchItemName()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If Form1.ExactMatchCHECKBOX.Checked = False Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If UCase(ItemObjectItem.ItemName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    If UCase(ItemObjectItem.ItemName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If UCase(ItemObjectItem.ItemName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                    If UCase(ItemObjectItem.ItemName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                End If
            End If


            If Form1.ExactMatchCHECKBOX.Checked = True Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemName) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemName) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemName) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                    If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemName) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                End If
            End If


        Next

    End Sub



    Sub SearchItemBase()

        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If Form1.ExactMatchCHECKBOX.Checked = False Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If UCase(ItemObjectItem.ItemBase).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                    If UCase(ItemObjectItem.ItemBase).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If UCase(ItemObjectItem.ItemBase).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '   NotEqualTo 
                    If UCase(ItemObjectItem.ItemBase).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '   NotEqualTo 
                End If
            End If

            'EXACT MATCH SEARCH FOR ITEM BASE
            If Form1.ExactMatchCHECKBOX.Checked = True Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemBase) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                    If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemBase) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemBase) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '  NotEqualTo 
                    If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemBase) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '  NotEqualTo 
                End If
            End If
        Next

    End Sub


    Sub SearchItemQuality()

        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If UCase(ItemObjectItem.ItemQuality).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    If UCase(ItemObjectItem.ItemQuality).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If UCase(ItemObjectItem.ItemQuality).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                    If UCase(ItemObjectItem.ItemQuality).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                End If
            End If

            'EXACT MATCH SEARCH FOR ITEM Quality
            If Form1.ExactMatchCHECKBOX.Checked = True Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemQuality) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemQuality) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemQuality) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                    If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemQuality) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                End If
            End If
        Next

    End Sub

    Sub SearchAttackClass()

        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ATTACK CLASS" Then

                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.AttackClass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                        If UCase(ItemObjectItem.AttackClass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                    End If
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.AttackClass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                        If UCase(ItemObjectItem.AttackClass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    End If
                End If

                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.AttackClass) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.AttackClass) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.AttackClass) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.AttackClass) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                    End If
                End If
            End If
        Next
    End Sub

    Sub SearchAttackSpeed()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ATTACK SPEED" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.AttackSpeed).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                        If UCase(ItemObjectItem.AttackSpeed).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.AttackSpeed).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                        If UCase(ItemObjectItem.AttackSpeed).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR ATTACK SPEED
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.AttackSpeed) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.AttackSpeed) And Form1.RefineSearchCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.AttackSpeed) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.AttackSpeed) And Form1.RefineSearchCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                    End If
                End If
            End If
        Next
    End Sub

    Sub SearchMuleAccount()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "MULE ACCOUNT" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.MuleAccount).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                        If UCase(ItemObjectItem.MuleAccount).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.MuleAccount).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                        If UCase(ItemObjectItem.MuleAccount).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR MULE ACCOUNT NAME
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MuleAccount) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MuleAccount) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.MuleAccount) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.MuleAccount) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                    End If
                End If

            End If
        Next
    End Sub

    Sub SearchMuleName()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "MULE NAME" Then

                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.MuleName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                        If UCase(ItemObjectItem.MuleName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.MuleName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '   NotEqualTo 
                        If UCase(ItemObjectItem.MuleName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '   NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR MULE NAME
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MuleName) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MuleName) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> UCase(ItemObjectItem.MuleName) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.MuleName) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '  NotEqualTo 
                    End If
                End If
            End If
        Next
    End Sub

    Sub SearchMulePass()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "MULE PASS" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.MulePass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                        If UCase(ItemObjectItem.MulePass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.MulePass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '   NotEqualTo 
                        If UCase(ItemObjectItem.MulePass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '   NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR MULE NAME
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MulePass) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MulePass) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> UCase(ItemObjectItem.MulePass) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.MulePass) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '  NotEqualTo 
                    End If

                End If

            End If
        Next
    End Sub

    Sub SearchUserReference()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)


            If Form1.ExactMatchCHECKBOX.Checked = False Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If UCase(ItemObjectItem.UserReference).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    If UCase(ItemObjectItem.UserReference).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If UCase(ItemObjectItem.UserReference).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                    If UCase(ItemObjectItem.UserReference).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                End If

            End If

            'EXACT MATCH SEARCH FOR USER REFERENCE
            If Form1.ExactMatchCHECKBOX.Checked = True Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Form1.SearchWordCOMBOBOX.Text = ItemObjectItem.UserReference And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    If Form1.SearchWordCOMBOBOX.Text = ItemObjectItem.UserReference And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Form1.SearchWordCOMBOBOX.Text <> ItemObjectItem.UserReference And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                    If Form1.SearchWordCOMBOBOX.Text <> ItemObjectItem.UserReference And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                End If
            End If

        Next
    End Sub

    Sub SearchRuneWord()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.ExactMatchCHECKBOX.Checked = False Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If UCase(ItemObjectItem.RuneWord).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    If UCase(ItemObjectItem.RuneWord).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If UCase(ItemObjectItem.RuneWord).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                    If UCase(ItemObjectItem.RuneWord).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' NotEqualTo 
                End If
            End If

            'EXACT MATCH SEARCH FOR USER REFERENCE
            If Form1.ExactMatchCHECKBOX.Checked = True Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Form1.SearchWordCOMBOBOX.Text = ItemObjectItem.RuneWord And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                    If Form1.SearchWordCOMBOBOX.Text = ItemObjectItem.RuneWord And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                End If

                If Form1.SearchWordCOMBOBOX.Text <> ItemObjectItem.RuneWord And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Form1.SearchWordCOMBOBOX.Text <> ItemObjectItem.RuneWord And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 

            End If
        Next

    End Sub

    Sub SearchRequiredLevel()

        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.RequiredLevel) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
                If Val(ItemObjectItem.RequiredLevel) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.RequiredLevel) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
                If Val(ItemObjectItem.RequiredLevel) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.RequiredLevel) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
                If Val(ItemObjectItem.RequiredLevel) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.RequiredLevel) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
                If Val(ItemObjectItem.RequiredLevel) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchRequiredStrength()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.RequiredStrength) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
                If Val(ItemObjectItem.RequiredStrength) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.RequiredStrength) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
                If Val(ItemObjectItem.RequiredStrength) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.RequiredStrength) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
                If Val(ItemObjectItem.RequiredStrength) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.RequiredStrength) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
                If Val(ItemObjectItem.RequiredStrength) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchRequiredDexterity()

        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.RequiredDexterity) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
                If Val(ItemObjectItem.RequiredDexterity) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.RequiredDexterity) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
                If Val(ItemObjectItem.RequiredDexterity) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.RequiredDexterity) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
                If Val(ItemObjectItem.RequiredDexterity) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.RequiredDexterity) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
                If Val(ItemObjectItem.RequiredDexterity) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchOneHandDamageMax()

        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.OneHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                If Val(ItemObjectItem.OneHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.OneHandDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Val(ItemObjectItem.OneHandDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.OneHandDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' GreaterThan
                If Val(ItemObjectItem.OneHandDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.OneHandDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    LessThan 
                If Val(ItemObjectItem.OneHandDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchOneHandDamageMin()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.OneHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                If Val(ItemObjectItem.OneHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.OneHandDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Val(ItemObjectItem.OneHandDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.OneHandDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' GreaterThan
                If Val(ItemObjectItem.OneHandDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.OneHandDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    LessThan 
                If Val(ItemObjectItem.OneHandDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchTwoHandDamageMax()

        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.OneHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                If Val(ItemObjectItem.OneHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.OneHandDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Val(ItemObjectItem.OneHandDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.OneHandDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' GreaterThan
                If Val(ItemObjectItem.OneHandDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.OneHandDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    LessThan 
                If Val(ItemObjectItem.OneHandDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchTwoHandDamageMin()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)

            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.TwoHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                If Val(ItemObjectItem.TwoHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.TwoHandDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Val(ItemObjectItem.TwoHandDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.TwoHandDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' GreaterThan
                If Val(ItemObjectItem.TwoHandDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.TwoHandDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    LessThan 
                If Val(ItemObjectItem.TwoHandDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchThrowDamageMin()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.ThrowDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                If Val(ItemObjectItem.TwoHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.ThrowDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Val(ItemObjectItem.ThrowDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.ThrowDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' GreaterThan
                If Val(ItemObjectItem.ThrowDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.ThrowDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    LessThan 
                If Val(ItemObjectItem.ThrowDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchThrowDamageMax()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.ThrowDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                If Val(ItemObjectItem.TwoHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.ThrowDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Val(ItemObjectItem.ThrowDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.ThrowDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' GreaterThan
                If Val(ItemObjectItem.ThrowDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.ThrowDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    LessThan 
                If Val(ItemObjectItem.ThrowDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub


    Sub SearchChanceToBlock()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.ChanceToBlock) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                If Val(ItemObjectItem.ChanceToBlock) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.ChanceToBlock) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Val(ItemObjectItem.ChanceToBlock) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.ChanceToBlock) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' GreaterThan
                If Val(ItemObjectItem.ChanceToBlock) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.ChanceToBlock) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    LessThan 
                If Val(ItemObjectItem.ChanceToBlock) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub

    Sub SearchItemDefense()
        Dim count = -1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            ProgressBar1(count)
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Val(ItemObjectItem.Defense) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     EqualTo
                If Val(ItemObjectItem.Defense) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '         EqualTo
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Val(ItemObjectItem.Defense) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) 'NotEqualTo 
                If Val(ItemObjectItem.Defense) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    NotEqualTo 
            End If
            If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                If Val(ItemObjectItem.Defense) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) ' GreaterThan
                If Val(ItemObjectItem.Defense) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '     GreaterThan
            End If

            If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                If Val(ItemObjectItem.Defense) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '    LessThan 
                If Val(ItemObjectItem.Defense) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : ItemMatched(count) '        LessThan 
            End If
        Next

    End Sub



    'this backsup the current database from menu selection and when closing app
    Sub BackupDatabase()
        Dim BackupPath = Application.StartupPath & "\Database\Backup\"
        Dim temp As String = ""
        Dim myarray = Split(Databasefile, ".txt", 0)
        Dim tempname = myarray(0) & ".bak"
        myarray = Split(tempname, "\")
        tempname = myarray(myarray.Length - 1)

        If My.Computer.FileSystem.FileExists(BackupPath & tempname) = True Then
            My.Computer.FileSystem.DeleteFile(BackupPath & tempname)
        End If
        My.Computer.FileSystem.CopyFile(DataBasePath & Databasefile, BackupPath & tempname)

    End Sub
    Sub SearchUniqueAttributes()
        Form1.SearchLISTBOX.Items.Clear()
        SearchReferenceList.Clear()
        Dim temp As String = ""
        Dim MyValue = -1


        For count = 0 To Objects.Count - 1
            ProgressBar1(count)
            If Objects(count).Stat1 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat1.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat1) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat1).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat1) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat2 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat2.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat2) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat2).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat2) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat3 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat3.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat3) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat3).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat3) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat4 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat4.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat4) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat4).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat4) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat5 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat5.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat5) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat5).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat5) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat6 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat6.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat6) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat6).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat6) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat7 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat7.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat7) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat7).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat7) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat8 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat8.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat8) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat8).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat8) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat9 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat9.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat9) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat9).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat9) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat10 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat10.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat10) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat10).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat10) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat11 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat11.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat11) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat11).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat11) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat12 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat12.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat12) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat12).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat12) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat13 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat13.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat13) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat13).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat13) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat14 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat14.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat14) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat14).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat14) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat15 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat15.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyValue = getvalue(Objects(count).Stat15) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat15).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyValue = getvalue(Objects(count).Stat15) : MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
            End If

        Next



        If SearchReferenceList.Count > 0 Then
            Form1.SearchLISTBOX.SelectedIndex = 0
            Form1.ListboxTABCONTROL.SelectTab(1)
            Form1.Button1.BackColor = Color.DimGray
            Form1.Button2.BackColor = Color.Black
            Form1.TextBox2.Text = Form1.SearchLISTBOX.Items.Count & " - Total Matches"
        End If


    End Sub


    Sub MyDecipher(ByVal count, ByVal xval)

        If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
            If xval = Form1.SearchValueNUMERICUPDWN.Value Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
        End If
        If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
            If xval > Form1.SearchValueNUMERICUPDWN.Value Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
        End If
        If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
            If xval < Form1.SearchValueNUMERICUPDWN.Value Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
        End If
        If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
            If xval <> Form1.SearchValueNUMERICUPDWN.Value Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
        End If

    End Sub


    Function getvalue(ByVal temp) As Integer
        Dim myvalue As Integer = 0
        temp = Replace(temp, "+", "")
        temp = Replace(temp, "%", "")
        temp = Replace(temp, ":", "")
        temp = Replace(temp, "-", "")
        Dim myarray = temp.Split(" ")
        For x = 0 To myarray.Length - 1
            If IsNumeric(myarray(x)) Then
                myvalue = myarray(x)
                Return myvalue
            End If
        Next
        Return myvalue
    End Function
    Sub ItemMatched(ByVal count)

        Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName)
        SearchReferenceList.Add(count)

    End Sub

    Sub ProgressBar1(ByRef count)
        If SearchProgressForm.Enabled = False Then Return
        SearchProgressForm.SearchPROGRESSBAR.Value = Int((count / Form1.AllItemsInDatabaseListBox.Items.Count) * 100)
        SearchProgressForm.SearchProgressLABEL1.Text = Form1.SearchLISTBOX.Items.Count & " Matches"
        SearchProgressForm.SearchProgressLABEL2.Text = "Searching " & count & " of " & Form1.AllItemsInDatabaseListBox.Items.Count & " Item Records..."
        ' SearchProgressForm.Refresh()
    End Sub
End Module
