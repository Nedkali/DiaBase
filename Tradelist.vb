Imports System.Text.RegularExpressions

Module Tradelist


    Public Sub SendToTradeList(ByVal x)
        Dim temp As String = ""

        If Objects(x).ItemBase = "Rune" Then
            Dim rune = Objects(x).ItemName.Replace(" Rune", "")
            Form1.RichTextBox3.AppendText(rune & vbCrLf & vbCrLf)
            Return
        End If

        '***********************************************
        'Unique items
        '***********************************************
        If Objects(x).ItemName = "The Stone of Jordan Ring" Then Form1.RichTextBox3.AppendText("Soj" & vbCrLf & vbCrLf) : Return
        If Objects(x).ItemName = "Mara's Kaleidoscope Amulet" Then
            temp = "Mara's " & Objects(x).Stat3
            GoTo Abbrev
        End If

        If Objects(x).ItemBase = "Small Charm" And Objects(x).ItemQuality = "Unique" Then
            If Objects(x).Stat2 = "" Then Form1.RichTextBox3.AppendText("Anni Unid" & vbCrLf & vbCrLf) : Return
            temp = "Anni "
            Dim temp1 = Regex.Replace(Objects(x).Stat2, "[^0-9]", "") & " " & Regex.Replace(Objects(x).Stat3, "[^0-9]", "") & " " & Regex.Replace(Objects(x).Stat4, "[^0-9]", "")
            Form1.RichTextBox3.AppendText(temp & temp1 & vbCrLf & vbCrLf) : Return
        End If

        If Objects(x).ItemName = "Andariel's Visage Demonhead" Then
            temp = "Andies " & Objects(x).Stat4 & " " & Objects(x).Stat6
            If Objects(x).EtherealItem = True Then temp = temp & " Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Crown of Ages Corona" Then
            temp = "CoA " & "Def" & Objects(x).Defense
            If Objects(x).Stat1 <> "Indestructible" Then temp = temp & Objects(x).Stat1
            temp = temp & " Soc" & Objects(x).Sockets
            If Objects(x).EtherealItem = True Then temp = temp & " Eth"
            GoTo Abbrev
        End If


        If Objects(x).ItemName = "Arachnid Mesh Spiderweb Sash" Then
            temp = "Arach  Def" & Objects(x).Defense
            GoTo Abbrev
        End If

        If Objects(x).ItemName = "Chance Guards Chain Gloves" Then
            temp = "Chancies  Mf" & Objects(x).Stat5
            GoTo Abbrev
        End If

        If Objects(x).ItemName = "Dracul's Grasp Vampirebone Gloves" Then
            temp = "Dracs  " & Objects(x).Stat2 & " " & Objects(x).Stat5
            temp = temp.Replace("Strength", "Str") : temp = temp.Replace("Life stolen per hit", "Loh")
            Form1.RichTextBox3.AppendText(temp & vbCrLf & vbCrLf) : Return
        End If

        If Objects(x).ItemName = "Eschuta's Temper Eldritch Orb" Then temp = "Eschuta's" : GoTo Stats




        '***********************************************
        'Specific items
        '***********************************************
        If Objects(x).ItemName = "Token of Absolution" Then Form1.RichTextBox3.AppendText("Token" & vbCrLf & vbCrLf) : Return



        '***********************************************
        'Rare/magic/white items
        '***********************************************
        temp = Objects(x).ItemName & ", " 'sets the default name

        Select Case (Objects(x).ItemBase)
            Case "Armor", "Helm", "Belt", "Shield", "Boots", "Gloves"
                temp = temp & "Def " & Objects(x).Defense
            Case "Small Charm"
                temp = "SC"
            Case "Large Charm"
                temp = "LC"
            Case "Grand Charm"
                temp = "GC"
        End Select

Stats:
        If Objects(x).Stat1 <> "" Then temp = temp & ", " & Objects(x).Stat1
        If Objects(x).Stat2 <> "" Then temp = temp & ", " & Objects(x).Stat2
        If Objects(x).Stat3 <> "" Then temp = temp & ", " & Objects(x).Stat3
        If Objects(x).Stat4 <> "" Then temp = temp & ", " & Objects(x).Stat4
        If Objects(x).Stat5 <> "" Then temp = temp & ", " & Objects(x).Stat5
        If Objects(x).Stat6 <> "" Then temp = temp & ", " & Objects(x).Stat6
        If Objects(x).Stat7 <> "" Then temp = temp & ", " & Objects(x).Stat7
        If Objects(x).Stat8 <> "" Then temp = temp & ", " & Objects(x).Stat8
        If Objects(x).Stat9 <> "" Then temp = temp & ", " & Objects(x).Stat9
        If Objects(x).Stat10 <> "" Then temp = temp & ", " & Objects(x).Stat10
        If Objects(x).Stat11 <> "" Then temp = temp & ", " & Objects(x).Stat11
        If Objects(x).Stat12 <> "" Then temp = temp & ", " & Objects(x).Stat13
        If Objects(x).Stat11 <> "" Then temp = temp & ", " & Objects(x).Stat14
        If Objects(x).Stat12 <> "" Then temp = temp & ", " & Objects(x).Stat15

Abbrev:
        temp = temp.Replace("Lightning Resistance", "Lr")
        temp = temp.Replace("Lightning Resist", "Lr")
        temp = temp.Replace("Cold Resistance", "Cr")
        temp = temp.Replace("Cold Resist", "Cr")
        temp = temp.Replace("Fire Resistance", "Fr")
        temp = temp.Replace("Fire Resist", "Fr")
        temp = temp.Replace("Poison Resistance", "Pr")
        temp = temp.Replace("Poison Resist", "Pr")
        temp = temp.Replace("to Lightning Skill Damage", "to Lit dmg")
        temp = temp.Replace("Sorceress Skill Levels", "Sorc Skills")
        temp = temp.Replace("to Dexterity", "to Dex")
        temp = temp.Replace("Better Chance of Getting Magic Items", "Mf")
        temp = temp.Replace("Ethereal (Cannot be Repaired)", "eth")
        temp = temp.Replace("Extra Gold from Monsters", "Gf")
        temp = temp.Replace("All Resistances", "Res All")
        temp = temp.Replace("to Experience Gained", "Xp")
        temp = temp.Replace("Faster Cast Rate", "fcr")
        temp = temp.Replace("Enhanced Defense", "Edef")
        temp = temp.Replace("Faster Hit Recovery", "Fhr")
        temp = temp.Replace("Faster Run/Walk", "Frw")
        temp = temp.Replace("Attack Rating", "Ar")
        temp = temp.Replace("Faster Block Rate", "Fbr")
        temp = temp.Replace("Increased Attack Speed", "Ias")
        temp = temp.Replace("Strength", "Str")
        temp = temp.Replace("Life stolen per hit", "LoH")
        temp = temp.Replace("Regenerate Mana", "Mana Regen")
        temp = temp.Replace("Life stolen per hit", "Loh")
        temp = temp.Replace("damage", "dmg")
        temp = temp.Replace("Maximum Damage", "Max dmg")
        temp = temp.Replace("Defense", "Def")
        temp = temp.Replace("Fire Skill Damage", "Fire dmg")


        temp = temp.Replace("Socketed", "Soc")
        temp = temp.Replace("Unidentified", "Unid")
        temp = temp.Replace("Unique", "Uniq")
        temp = temp.Replace("Ethereal", "Eth")


        temp = temp.Replace("+", "")
        temp = temp.Replace("(", "")
        temp = temp.Replace(")", "")
        temp = temp.Replace("Sorceress", "Sorc")
        temp = temp.Replace("Paladin", "Pala")
        temp = temp.Replace("Necromancer", "Nec")
        temp = temp.Replace("Barbarian", "Barb")
        'temp = temp.Replace("Druid Only", "Druid") ' not worth trimming
        temp = temp.Replace("Assasin", "Sin")
        temp = temp.Replace("Amazon", "Zon")
        temp = temp.Replace("Only", "")
        temp = temp.Replace(", ,", ",") ' couple extra "," slip in for some reason
        temp = temp.Replace(",,", ",") ' couple extra "," slip in for some reason

        Form1.RichTextBox3.AppendText(temp & vbCrLf & vbCrLf)
    End Sub


End Module
