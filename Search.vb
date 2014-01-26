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
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).ItemName)
                Next
            Case "Item Base"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).ItemBase)
                Next
            Case "Item Quality"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).ItemQuality)
                Next
            Case "Item Defense"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    If Val(Objects(count).Defense) > 0 Then MyDecipher(count, Val(Objects(count).Defense))
                Next
            Case "RuneWord"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    If Objects(count).RuneWord = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count)
                Next
            Case "Chance To Block"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    MyDecipher(count, Val(Objects(count).ChanceToBlock))
                Next
            Case "One Hand Damage Max"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    If Val(Objects(count).OneHandDamageMax) > 0 Then MyDecipher(count, Val(Objects(count).OneHandDamageMax))
                Next
            Case "One Hand Damage Min"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    If Val(Objects(count).OneHandDamageMin) > 0 Then MyDecipher(count, Val(Objects(count).OneHandDamageMin))
                Next
            Case "Two Hand Damage Max"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    If Val(Objects(count).TwoHandDamageMax) > 0 Then MyDecipher(count, Val(Objects(count).TwoHandDamageMax))
                Next
            Case "Two Hand Damage Min"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    If Val(Objects(count).TwoHandDamageMin) > 0 Then MyDecipher(count, Val(Objects(count).TwoHandDamageMin))
                Next
            Case "Throw Damage Max"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    If Val(Objects(count).ThrowDamageMax) > 0 Then MyDecipher(count, Val(Objects(count).ThrowDamageMax))
                Next
            Case "Throw Damage Min"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    If Val(Objects(count).ThrowDamageMin) > 0 Then MyDecipher(count, Val(Objects(count).ThrowDamageMin))
                Next
            Case "Required Level"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    MyDecipher(count, Val(Objects(count).RequiredLevel))
                Next
            Case "Required Strength"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    MyDecipher(count, Val(Objects(count).RequiredStrength))
                Next
            Case "Required Dexterity"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    MyDecipher(count, Val(Objects(count).RequiredDexterity))
                Next
            Case "Attack Class"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).AttackClass)
                Next
            Case "Attack Speed"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).AttackSpeed)
                Next
            Case "Unique Attributes"
                SearchUniqueAttributes()
            Case "Mule Name"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).MuleName)
                Next
            Case "Mule Account"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).MuleAccount)
                Next
            Case "Mule Pass"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).MulePass)
                Next
            Case "User Reference"
                For count = 0 To Objects.Count - 1
                    ProgressBar1(count)
                    TextDecipher(count, Objects(count).UserReference)
                Next

        End Select

        SearchProgressForm.Close()

        If SearchReferenceList.Count > 0 Then
            Form1.SearchLISTBOX.SelectedIndex = 0
            Form1.ListboxTABCONTROL.SelectTab(1)
            Form1.Button1.BackColor = Color.DimGray
            Form1.Button2.BackColor = Color.Black
            Form1.TextBox2.Text = Form1.SearchLISTBOX.Items.Count & " - Total Matches"
        End If
    End Sub

    Sub TextDecipher(ByVal count, ByVal txtval)

        If Form1.ExactMatchCHECKBOX.Checked = True Then
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If txtval.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                End If
            End If
            If Form1.SearchOperatorCOMBOBOX.Text <> "Equal To" Then
                If txtval.IndexOf(Form1.SearchWordCOMBOBOX.Text) = -1 Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                End If
            End If
        End If

        If Form1.ExactMatchCHECKBOX.Checked = False Then
            If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                If LCase(txtval).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                End If

            End If
            If Form1.SearchOperatorCOMBOBOX.Text <> "Equal To" Then
                If LCase(txtval).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) = -1 Then
                    If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                    Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
                End If
            End If
        End If

    End Sub


    Sub SearchUniqueAttributes()
        Form1.SearchLISTBOX.Items.Clear()
        SearchReferenceList.Clear()

        For count = 0 To Objects.Count - 1
            ProgressBar1(count)
            If Objects(count).Stat1 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat1.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat1)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat1).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat1)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat2 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat2.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat2)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat2).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat2)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat3 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat3.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat3)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat3).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat3)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat4 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat4.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat4)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat4).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat4)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat5 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat5.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat5)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat5).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat5)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat6 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat6.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat6)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat6).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat6)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat7 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat7.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat7)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat7).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat7)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat8 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat8.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat8)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat8).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat8)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat9 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat9.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat9)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat9).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat9)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat10 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat10.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat10)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat10).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat10)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat11 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat11.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat11)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat11).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat11)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat12 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat12.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat12)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat12).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat12)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat13 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat13.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat13)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat13).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat13)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat14 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat14.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat14)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat14).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat14)) : Continue For 'skip rest we have found it
            End If

            If Objects(count).Stat15 = "" Then Continue For ' optimize seach time
            If Form1.ExactMatchCHECKBOX.Checked = True Then ' case sensitive search
                If Objects(count).Stat15.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat15)) : Continue For 'skip rest we have found it
            End If
            If Form1.ExactMatchCHECKBOX.Checked = False Then
                If LCase(Objects(count).Stat15).IndexOf(LCase(Form1.SearchWordCOMBOBOX.Text)) > -1 Then MyDecipher(count, getvalue(Objects(count).Stat15)) : Continue For 'skip rest we have found it
            End If

        Next


    End Sub


    Sub MyDecipher(ByVal count, ByVal xval)

        If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
            If xval = Form1.SearchValueNUMERICUPDWN.Value Then
                If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
            End If
        End If
        If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
            If xval > Form1.SearchValueNUMERICUPDWN.Value Then
                If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
            End If
        End If

        If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
            If xval < Form1.SearchValueNUMERICUPDWN.Value Then
                If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
            End If

        End If
        If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
            If xval <> Form1.SearchValueNUMERICUPDWN.Value Then
                If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(Objects(count).ItemName) = True Then Return
                Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
            End If

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
