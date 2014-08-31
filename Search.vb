Module Search
    Sub SearchRoutine()
        RefineSearchReferenceList.Clear()                       'Clear out old refine list
        If Form1.SearchLISTBOX.Items.Count > 0 Then             'Can only run a refined search if items exist already in the search lisbox
            For Each item In SearchReferenceList
                RefineSearchReferenceList.Add(item)             'Create a reference list of each items object location in database that have 
                '                                               'Already been matched and are right now still in the matched search list....
                '                                               'Refine searches will use this list to look for matches as opposed all items list
            Next
        End If

        SearchReferenceList.Clear()                                 'Clear Out old Item Search Reference List (Holds location in database of each matched item)
        Form1.SearchLISTBOX.Items.Clear()                           'Clear out old search matches
        SearchProgressForm.SearchPROGRESSBAR.Value = 0
        '--------------------------------------------------------------------------------------------------------------------------------------------
        'Rob added if/thens here to trial skip search progress bar to speed up big database searches idea
        If ShowSearchProgress = True Then SearchProgressForm.Show() 'Search Progress Bar Form
        If ShowSearchProgress = False Then Form1.ItemTallyTEXTBOX.Text = "Searching..." : Form1.ItemTallyTEXTBOX.Refresh()

        SearchProgressForm.Refresh() 'Not totally sure why but gold border around progress bar wont work unless i refresh the form again here and now..???
        '--------------------------------------------------------------------------------------------------------------------------------------------

        'SELECT CASE FOR THE FIELD BEING SEARCHED - NOTE: UNIQUE ATTRIBUTES HAS ITS OWN SEARCH ENGINE SUB ROUTINE
        Select Case Form1.SearchFieldCOMBOBOX.Text

            Case "Item Name"
                'Normal Searches - Item Name
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count) 'Rob added "if/Then" on this line to trial skipping search progress bar idea
                        TextDecipher(count, Objects(count).ItemName)
                    Next
                End If

                'Refining Searches - Item Name
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count) 'Rob added "if/Then" on this line to trial skipping search progress bar idea
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).ItemName)
                    Next
                End If

            Case "Item Base"
                'Normal Searches - Item Base
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count) 'Rob added "if/Then" on this line to trial skipping search progress bar idea
                        TextDecipher(count, Objects(count).ItemBase)
                    Next
                End If

                'Refining Searches - Item Base
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count) 'Rob added "if/Then" on this line to trial skipping search progress bar idea
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).ItemBase)
                    Next
                End If

            Case "Item Quality"
                'Normal Searches - Quality
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count) 'Rob added "if/Then" on this line to trial skipping search progress bar idea
                        TextDecipher(count, Objects(count).ItemQuality)
                    Next
                End If

                'Refining Searches - Quality
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count) 'Rob added "if/Then" on this line to trial skipping search progress bar idea
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).ItemQuality)
                    Next
                End If

            Case "Item Defense"
                'Normal Searches - Defense
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(count).Defense) > 0 Then MyDecipher(count, Val(Objects(count).Defense))
                    Next
                End If

                'Refining Searches - Defense
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(RefineSearchReferenceList(count)).Defense) > 0 Then MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).Defense))
                    Next
                End If



            Case "RuneWord"
                'Normal Searches - Runeword
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        'If Objects(count).RuneWord = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count)
                        TextDecipher(count, Objects(count).RuneWord) ' ROB BUG FIX - REV 28-  Made runeword search string based (TextDeciperSUB) instead of boolean based to avoid crashes from RuneWord fields that dont have a valid true or false value
                    Next
                End If

                'Refining Searches - Runeword
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        'If Objects(RefineSearchReferenceList(count)).RuneWord = True Then Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count))
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).RuneWord) ' ROB BUG FIX - REV 28 - Made runeword search string based (TextDeciperSUB) instead of boolean based to avoid crashes from RuneWord fields that dont have a vaild true or false value
                    Next
                End If

            Case "Chance To Block"
                'Normal Searches - Runeword
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(count).ChanceToBlock))
                    Next
                End If

                'Refining Searches - Runeword
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).ChanceToBlock))
                    Next
                End If

            Case "One Hand Damage Max"
                'Normal Searches - One Hand Damage Max
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(count).OneHandDamageMax) > 0 Then MyDecipher(count, Val(Objects(count).OneHandDamageMax))
                    Next
                End If

                'Refining Searches - One Hand Damage Max
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(RefineSearchReferenceList(count)).OneHandDamageMax) > 0 Then MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).OneHandDamageMax))
                    Next
                End If

            Case "One Hand Damage Min"
                'Normal Searches - One Hand Damage Min
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(count).OneHandDamageMin) > 0 Then MyDecipher(count, Val(Objects(count).OneHandDamageMin))
                    Next
                End If

                'Refining Searches - One Hand Damage Min
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(RefineSearchReferenceList(count)).OneHandDamageMin) > 0 Then MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).OneHandDamageMin))
                    Next
                End If

            Case "Two Hand Damage Max"
                'Normal Searches - Two Hand Damage Max
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(count).TwoHandDamageMax) > 0 Then MyDecipher(count, Val(Objects(count).TwoHandDamageMax))
                    Next
                End If

                'Refining Searches - Two Hand Damage Max
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(RefineSearchReferenceList(count)).TwoHandDamageMax) > 0 Then MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).TwoHandDamageMax))
                    Next
                End If

            Case "Two Hand Damage Min"
                'Normal Searches - Two Hand Damage Min
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(count).TwoHandDamageMin) > 0 Then MyDecipher(count, Val(Objects(count).TwoHandDamageMin))
                    Next
                End If
                'Refining Searches - Two Hand Damage Min
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(RefineSearchReferenceList(count)).TwoHandDamageMin) > 0 Then MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).TwoHandDamageMin))
                    Next
                End If

            Case "Throw Damage Max"
                'Normal Searches - Throw Damage Max
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(count).ThrowDamageMax) > 0 Then MyDecipher(count, Val(Objects(count).ThrowDamageMax))
                    Next
                End If
                'Refining Searches - Throw Damage Max
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(RefineSearchReferenceList(count)).ThrowDamageMax) > 0 Then MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).ThrowDamageMax))
                    Next
                End If

            Case "Throw Damage Min"
                'Normal Searches - Throw Damage Min
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(count).ThrowDamageMin) > 0 Then MyDecipher(count, Val(Objects(count).ThrowDamageMin))
                    Next
                End If

                'Refiening Searches - Throw Damage Min
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        If Val(Objects(RefineSearchReferenceList(count)).ThrowDamageMin) > 0 Then MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).ThrowDamageMin))
                    Next
                End If

            Case "Required Level"
                'Normal Searches - Required Level
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(count).RequiredLevel))
                    Next
                End If

                'Refining Searches - Required Level
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).RequiredLevel))
                    Next
                End If

            Case "Sockets"
                'Normal Searches - Sockets
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(count).Sockets))
                    Next
                End If

                'Refining Searches - Sockets
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).Sockets))
                    Next
                End If

            Case "Required Strength"
                'Normal Searches - Required Strength
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(count).RequiredStrength))
                    Next
                End If

                'Refining Searches - Required Strength
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).RequiredStrength))
                    Next
                End If

            Case "Required Dexterity"
                'Normal Searches - Required Dextreity
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(count).RequiredDexterity))
                    Next
                End If

                'Refining Searches - Required Dextreity
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        MyDecipher(count, Val(Objects(RefineSearchReferenceList(count)).RequiredDexterity))
                    Next
                End If

            Case "Attack Class"
                'Normal Searches - Attack Class
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(count).AttackClass)
                    Next
                End If

                'Refining Searches - Attack Class
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).AttackClass)
                    Next
                End If

            Case "Attack Speed"
                'Normal Searches - Attack Speed
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(count).AttackSpeed)
                    Next
                End If

                'Refining Searches - Attack Speed
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).AttackSpeed)
                    Next
                End If

            Case "Mule Name"
                'Normal Searches - Mule Name
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(count).MuleName)
                    Next
                End If

                'Refining Searches - Mule Name
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).MuleName)
                    Next
                End If

            Case "Mule Account"
                'Normal Searches - Mule Account
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(count).MuleAccount)
                    Next
                End If

                'Refining Searches - Mule Account
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).MuleAccount)
                    Next
                End If

            Case "Mule Pass"
                'Normal Searches - Mule Account Password
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(count).MulePass)
                    Next
                End If

                'Refining Searches - Mule Account Password
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).MulePass)
                    Next
                End If

            Case "User Reference"
                'Normal Searches - Users Reference Field
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    For count = 0 To Objects.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(count).UserReference)
                    Next
                End If

                'Refining Searches - Users Reference Field
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    For count = 0 To RefineSearchReferenceList.Count - 1
                        If ShowSearchProgress = True Then ProgressBar1(count)
                        TextDecipher(count, Objects(RefineSearchReferenceList(count)).UserReference)
                    Next
                End If

            Case "Unique Attributes"
                'Seperate Routine For Searching Unique Attribs Block
                SearchUniqueAttributes()

        End Select
        SearchProgressForm.Close()

        If SearchReferenceList.Count > 0 Then
            Form1.SearchLISTBOX.SelectedIndex = 0
            Form1.ListboxTABCONTROL.SelectTab(1)
            Form1.SearchListControlTabBUTTON.BackColor = Color.DimGray
            Form1.ListControlTabBUTTON.BackColor = Color.Black
            Form1.TradesListControlTabBUTTON.BackColor = Color.Black
         
            'POPULATES SEARCHES QUICK SELECTION DROP DOWN COLLECTIONS after successful search entries (Item Name, User Reference, Unique Attribs Strings) <-----------ROBS REV20
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ITEM NAME" And Form1.SearchWordCOMBOBOX.Text <> "" And ItemNamePulldownList.Contains(Form1.SearchWordCOMBOBOX.Text) = False Then ItemNamePulldownList.Add(Form1.SearchWordCOMBOBOX.Text) : Form1.SearchWordCOMBOBOX.Items.Add(Form1.SearchWordCOMBOBOX.Text)
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "UNIQUE ATTRIBUTES" And Form1.SearchWordCOMBOBOX.Text <> "" And UniqueAttribsPulldownList.Contains(Form1.SearchWordCOMBOBOX.Text) = False Then UniqueAttribsPulldownList.Add(Form1.SearchWordCOMBOBOX.Text) : Form1.SearchWordCOMBOBOX.Items.Add(Form1.SearchWordCOMBOBOX.Text)
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "USER REFERENCE" And Form1.SearchWordCOMBOBOX.Text <> "" And UserReferencePulldownList.Contains(Form1.SearchWordCOMBOBOX.Text) = False Then UserReferencePulldownList.Add(Form1.SearchWordCOMBOBOX.Text) : Form1.SearchWordCOMBOBOX.Items.Add(Form1.SearchWordCOMBOBOX.Text)
        End If
        Form1.ItemTallyTEXTBOX.Text = Form1.SearchLISTBOX.Items.Count & " - Total Matches"

    End Sub

    Sub TextDecipher(ByVal count, ByVal txtval)

        'EQUAL TO - NOT EXACT MATCH - NORMAL
        If Form1.ExactMatchCHECKBOX.Checked = False Then
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    If txtval = Nothing Then txtval = "" 'To avoid null reference errors 
                    If LCase(txtval).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then
                        If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Exit Sub
                        Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                    End If
                End If

                'EQUAL TO - NOT EXACT MATCH - REFINED
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    If txtval = Nothing Then txtval = "" 'To avoid null reference errors 
                    If LCase(txtval).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then
                        If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Exit Sub
                        Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Return
                    End If
                End If
            End If

            'NOT EQUAL TO - NOT EXACT MATCH - NORMAL
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    If txtval = Nothing Then txtval = "" 'To avoid null reference errors 
                    If LCase(txtval).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 Then
                        If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                        Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                    End If
                End If

                'NOT EQUAL TO - NOT EXACT MATCH - REFINED
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    If txtval = Nothing Then txtval = "" 'To avoid null reference errors 
                    If LCase(txtval).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 Then
                        If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Return
                        Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Return
                    End If
                End If
            End If
        End If

        'EQUAL TO - EXACT MATCH - NORMAL
        If Form1.ExactMatchCHECKBOX.Checked = True Then
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    If txtval = Nothing Then txtval = "" 'To avoid null reference errors 
                    If txtval = Form1.SearchWordCOMBOBOX.Text Then
                        If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Exit Sub
                        Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                    End If
                End If

                'EQUAL TO - EXACT MATCH - REFINED
                If Form1.RefineSearchCHECKBOX.Checked = True Then
                    If txtval = Nothing Then txtval = "" 'To avoid null reference errors 
                    If txtval = Form1.SearchWordCOMBOBOX.Text Then
                        If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Exit Sub
                        Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Return
                    End If
                End If
            End If

            'NOT EQUAL TO - EXACT MATCH - NORMAL
            If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                If Form1.RefineSearchCHECKBOX.Checked = False Then
                    If txtval = Nothing Then txtval = "" 'To avoid null reference errors 
                    If txtval = Form1.SearchWordCOMBOBOX.Text Then
                        If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                        Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                    End If
                End If
            End If

            'NOT EQUAL TO - EXACT MATCH - REFINED
            If Form1.RefineSearchCHECKBOX.Checked = True Then
                If txtval = Nothing Then txtval = "" 'To avoid null reference errors 
                If txtval = Form1.SearchWordCOMBOBOX.Text Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Return
                End If
            End If
        End If

    End Sub

    Sub SearchUniqueAttributes() 'SEARCHES UNIQUE ATTRIBUTES
        'NORMAL SEARCH FOR UNIQUE ATTRIBUTES BLOCK
        Dim NotEqualToCounter As Integer = 0

        If Form1.RefineSearchCHECKBOX.Checked = False Then
            For count = 0 To Objects.Count - 1
                If ShowSearchProgress = True Then ProgressBar1(count)

                'STAT1
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat1 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat1)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat1 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat1 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 'If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat1).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat1)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat1).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat1).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT2
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat2 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat2)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat2 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat2 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat2).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat2)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat2).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat2).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT3
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat3 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat3)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat3 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat3 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat3).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat3)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat3).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat3).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT4
                 If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat4 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat4)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat4 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat4 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat4).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat4)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat4).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat4).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT5
                 If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat5 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat5)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat5 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat5 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat5).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat5)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat5).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat5).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT6
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat6 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat6)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat6 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat6 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat6).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat6)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat6).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat6).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT7
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat7 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat7)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat7 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat7 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 'If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat7).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat7)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat7).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat7).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT8
               If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat8 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat8)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat8 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat8 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat8).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat8)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat8).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat8).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT9
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat9 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat9)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat9 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat9 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat9).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat9)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat9).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat9).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT10
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat10 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat10)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat10 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat10 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat10).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat10)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat10).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat10).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT11
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat11 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat11)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat11 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat11 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat11).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat11)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat11).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat11).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT12
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat12 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat12)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat12 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat12 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat12).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat12)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat12).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat12).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT13
                 If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat13 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat13)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat13 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat13 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat13).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat13)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat13).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat13).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT14
                 If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat14 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat14)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat14 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat14 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat14).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat14)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat14).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat14).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'STAT15
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING  FOR VALUE SEARCHES
                    If Objects(count).Stat15 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat15)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(count).Stat15 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(count).Stat15 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(count).Stat15).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(count).Stat15)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(count).Stat15).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(count).Stat15).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Continue For
                End If

                'APPLYS NOT EQUAL TO STRING SEARCH MATCH RESULT IF ALL STATS ARE NOT EQUAL TO SEARCH STRING
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" And NotEqualToCounter = 15 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count)
                NotEqualToCounter = 0
            Next
        End If

        'REFINED SEARCH FOR UNIQUE ATTRIBUTES BLOCK
        If Form1.RefineSearchCHECKBOX.Checked = True Then
            For count = 0 To RefineSearchReferenceList.Count - 1
                If ShowSearchProgress = True Then ProgressBar1(count)

                'STAT1
                 If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat1 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat1)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat1 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat1 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat1).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat1)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat1).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat1).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT2
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat2 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat2)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat2 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat2 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat2).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat2)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat2).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat2).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT3
                 If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat3 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat3)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat3 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat3 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat3).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat3)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat3).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat3).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT4
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat4 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat4)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat4 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat4 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat4).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat4)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat4).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat4).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT5
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat5 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat5)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat5 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat5 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat5).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat5)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat5).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat5).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT6
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat6 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat6)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat6 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat6 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat6).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat6)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat6).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat6).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT7
               If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat7 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat7)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat7 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat7 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat7).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat7)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat7).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat7).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT8
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat8 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat8)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat8 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat8 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat8).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat8)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat8).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat8).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT9
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat9 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat9)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat9 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat9 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat9).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat9)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat9).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat9).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT10
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat10 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat10)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat10 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat10 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat10).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat10)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat10).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat10).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT11
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat11 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat11)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat11 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat11 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat11).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat11)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat11).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat11).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT12
               If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat12 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat12)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat12 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat12 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat12).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat12)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat12).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat12).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT13
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat13 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat13)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat13 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat13 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat13).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat13)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat13).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat13).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT14
                If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat14 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat14)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat14 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat14 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat14).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat14)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat14).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat14).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'STAT15
                 If Form1.ExactMatchCHECKBOX.Checked = True Then
                    'VALUE AND STRING SEARCH - EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If Objects(RefineSearchReferenceList(count)).Stat15 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat15)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat15 = Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If Objects(RefineSearchReferenceList(count)).Stat15 <> Form1.SearchWordCOMBOBOX.Text And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    'VALUE AND STRING SEARCH - NOT EXACT MATCH - ASSUMES EQUAL TO STRING FOR VALUE SEARCHES
                    If LCase(Objects(RefineSearchReferenceList(count)).Stat15).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value > 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else MyDecipher(count, getvalue(Objects(RefineSearchReferenceList(count)).Stat15)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat15).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For

                    'STRING ONLY SEARCHES - NOT EXACT MATCH - NOT EQUAL TO
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then If LCase(Objects(RefineSearchReferenceList(count)).Stat15).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.SearchValueNUMERICUPDWN.Value = 0 Then NotEqualToCounter = NotEqualToCounter + 1 ' If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Continue For
                End If

                'APPLYS NOT EQUAL TO STRING SEARCH MATCH IF ALL STATS ARE NOT EQUAL TO STRING
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" And NotEqualToCounter = 15 Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Continue For Else Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count))
                NotEqualToCounter = 0

            Next
        End If
    End Sub

    'Compares Extracted String Value With Entered Search Value Relevant To The Selected Operator (= <> > <)
    Sub MyDecipher(ByVal count, ByVal xval)
        If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
            'Normal Searches - Equal To
            If Form1.RefineSearchCHECKBOX.Checked = False Then
                If xval = Form1.SearchValueNUMERICUPDWN.Value Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                End If
            End If

            'Refining Searches - Equal To
            If Form1.RefineSearchCHECKBOX.Checked = True Then
                If xval = Form1.SearchValueNUMERICUPDWN.Value Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Return
                End If
            End If
        End If

        If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
            'Normal Searches - Not Equal To
            If Form1.RefineSearchCHECKBOX.Checked = False Then
                If xval <> Form1.SearchValueNUMERICUPDWN.Value Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                End If
            End If

            'Refining Searches - Not Equal To
            If Form1.RefineSearchCHECKBOX.Checked = True Then
                If xval <> Form1.SearchValueNUMERICUPDWN.Value Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Return
                End If
            End If
        End If

        If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
            'Normal Searches - Greater Than
            If Form1.RefineSearchCHECKBOX.Checked = False Then
                If xval > Form1.SearchValueNUMERICUPDWN.Value Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                End If
            End If

            'Refining Searches - Greater Than
            If Form1.RefineSearchCHECKBOX.Checked = True Then
                If xval > Form1.SearchValueNUMERICUPDWN.Value Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Return
                End If
            End If
        End If

        If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
            'Normal Searches - Less Than
            If Form1.RefineSearchCHECKBOX.Checked = False Then
                If xval < Form1.SearchValueNUMERICUPDWN.Value Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                End If
            End If

            'Refining Searches - Less Than
            If Form1.RefineSearchCHECKBOX.Checked = True Then
                If xval < Form1.SearchValueNUMERICUPDWN.Value Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(RefineSearchReferenceList(count)).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(RefineSearchReferenceList(count)).ItemName) : SearchReferenceList.Add(RefineSearchReferenceList(count)) : Return
                End If
            End If
        End If
    End Sub

    'PULLS INTEGER 
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

    'runs search progress bar
    Sub ProgressBar1(ByRef count)
        If SearchProgressForm.Enabled = False Then Return
        SearchProgressForm.SearchPROGRESSBAR.Value = Int((count / Form1.AllItemsInDatabaseListBox.Items.Count) * 100)

        SearchProgressForm.SearchProgressLABEL1.Text = Form1.SearchLISTBOX.Items.Count & " Matches"
        SearchProgressForm.SearchProgressLABEL1.Refresh()

        SearchProgressForm.SearchProgressLABEL2.Text = "Searching " & count & " of " & Form1.AllItemsInDatabaseListBox.Items.Count & " Item Records..."
        SearchProgressForm.SearchProgressLABEL2.Refresh()
    End Sub
End Module
