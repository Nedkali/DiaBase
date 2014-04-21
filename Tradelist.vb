Imports System.Text.RegularExpressions

Module Tradelist


    Public Sub SendToTradeList(ByVal x)
        Dim temp As String = ""

        '***********************************************
        'Unique items
        '***********************************************
        'Rings
        '***********************************************
        If Objects(x).ItemName = "The Stone of Jordan Ring" Then Form1.RichTextBox3.AppendText("Soj" & vbCrLf & vbCrLf) : Return
        If Objects(x).ItemName = "Dwarf Star Ring" Then Form1.RichTextBox3.AppendText(Objects(x).ItemName & vbCrLf & vbCrLf) : Return
        If Objects(x).ItemName = "Raven Frost Ring" Then temp = "Raven Frost, " & Objects(x).Stat1 & " " & Objects(x).Stat3 : GoTo Abbrev
        If Objects(x).ItemName = "Nagelring Ring" Then temp = "Nagel Ring, " & Objects(x).Stat4 : GoTo Abbrev
        If Objects(x).ItemName = "Nature's Peace Ring" Then temp = "Nature's Peace Ring, " : GoTo Abbrev
        If Objects(x).ItemName = "Manald Heal Ring" Then temp = "Manald Heal Ring, " : GoTo Abbrev

        '***********************************************
        'Amulets
        '***********************************************
        If Objects(x).ItemName = "Mara's Kaleidoscope Amulet" Then temp = "Mara's, " & Objects(x).Stat3 : GoTo Abbrev
        If Objects(x).ItemName = "Tal Rasha's Adjudication Amulet" Then temp = "Tal's  Amulet" : GoTo Abbrev
        If Objects(x).ItemName = "The Eye of Etlich Amulet" Then temp = "Etlich Amulet, " & Objects(x).Stat3 : GoTo Abbrev
        If Objects(x).ItemName = "The Rising Sun Amulet" Then temp = "Rising Sun Amulet, " : GoTo Abbrev
        If Objects(x).ItemName = "The Mahim-Oak Curio Amulet" Then temp = "Mahim-Oak Amulet, " : GoTo Abbrev
        If Objects(x).ItemName = "Seraph's Hymn Amulet" Then temp = "Seraph's Hymn Amulet, " : GoTo Abbrev

        '***********************************************
        'Annihilus
        '***********************************************
        If Objects(x).ItemBase = "Small Charm" And Objects(x).ItemQuality = "Unique" Then
            If Objects(x).Stat2 = "" Then Form1.RichTextBox3.AppendText("Anni Unid" & vbCrLf & vbCrLf) : Return
            temp = "Anni, "
            Dim temp1 = Regex.Replace(Objects(x).Stat2, "[^0-9]", "") & " " & Regex.Replace(Objects(x).Stat3, "[^0-9]", "") & " " & Regex.Replace(Objects(x).Stat4, "[^0-9]", "")
            Form1.RichTextBox3.AppendText(temp & temp1 & vbCrLf & vbCrLf) : Return
        End If

        '***********************************************
        'Helms
        '***********************************************
        If Objects(x).ItemName = "Andariel's Visage Demonhead" Then
            temp = "Andies, " & Objects(x).Stat4 & " " & Objects(x).Stat6 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Crown of Ages Corona" Then
            temp = "CoA, Def" & Objects(x).Defense
            If Objects(x).Stat1 <> "Indestructible" Then temp = temp & Objects(x).Stat1
            temp = temp & " Soc" & Objects(x).Sockets : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Tal Rasha's Horadric Crest Death Mask" Then
            temp = "Tals Death Mask, Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Harlequin Crest Shako" Then
            temp = "Shako, Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Veil of Steel Spired Helm" Then
            temp = "Veil of Steel, Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Peasant Crown War Hat" Then
            temp = "Peasant Hat , Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Kira's Guardian Tiara" Then
            temp = "Kira's , Def " & Objects(x).Defense & " " & Objects(x).Stat3 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If


        '***********************************************
        'Belts
        '***********************************************
        If Objects(x).ItemName = "Arachnid Mesh Spiderweb Sash" Then
            temp = "Arach, Def" & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Thundergod's Vigor War Belt" Then
            temp = "Thundergod's, Def" & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Verdungo's Hearty Cord Mithril Coil" Then
            temp = "Dungo's, Def" & Objects(x).Defense & " " & Objects(x).Stat3
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Tal Rasha's Fine-Spun Cloth Mesh Belt" Then
            temp = "Tal's Belt, Def" & Objects(x).Defense & " " & Objects(x).Stat4 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "String of Ears Demonhide Sash" Then
            temp = "String of Ears, Def" & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Goldwrap Heavy Belt" Then temp = "Goldwrap, Def" : GoTo Stats

        '***********************************************
        'Gloves
        '***********************************************
        If Objects(x).ItemName = "Chance Guards Chain Gloves" Then
            temp = "Chancies, Mf" & Objects(x).Stat5 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName.IndexOf("Magefist") > -1 Then
            temp = "Magefist, Def" & Objects(x).Defense & " " & Objects(x).Stat4 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Dracul's Grasp Vampirebone Gloves" Then
            temp = "Dracs, " & Objects(x).Stat2 & " " & Objects(x).Stat5 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Venom Grip Demonhide Gloves" Then
            temp = "Venom Grip, Def" & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Lava Gout Crusader Gauntlets" Then
            temp = "Lava Gout Crusader Gauntlets, Def" & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Lava Gout Battle Gauntlets" Then
            temp = "Lava Gout Battle Gauntlets, Def" & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If


        '***********************************************Lava Gout Battle Gauntlets
        'Boots
        '***********************************************
        If Objects(x).ItemName = "War Traveler Battle Boots" Then
            temp = "War Travs, " & Objects(x).Stat8 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Waterwalk Sharkskin Boots" Then
            temp = "Waterwalks, " & Objects(x).Stat5 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Sandstorm Trek Scarabshell Boots" Then
            temp = "Treks, Def" & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Marrowwalk Boneweave Boots" Then
            temp = "Marrowwalk, Def" & Objects(x).Defense & " " & Objects(x).Stat2 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Gore Rider War Boots" Then
            temp = "Gore Rider's, Def" & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If


        '***********************************************
        'Armor
        '***********************************************
        If Objects(x).ItemName = "Skin of the Vipermagi Serpentskin Armor" Then
            temp = "Vipermagi  " & Objects(x).Stat4 : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Tal Rasha's Guardianship Lacquered Plate" Then
            temp = "Tal's Armor Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Shaftstop Mesh Armor" Then
            temp = "Shaftstop Armor Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If


        '***********************************************
        ' Runeword Armor
        '***********************************************
        If Objects(x).ItemName.IndexOf("Chains of Honor") > -1 And Objects(x).ItemBase = "Armor" Then
            temp = Objects(x).ItemName & " Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName.IndexOf("Enigma") > -1 And Objects(x).ItemBase = "Armor" Then
            temp = Objects(x).ItemName & " Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName.IndexOf("Treachery") > -1 And Objects(x).ItemBase = "Armor" Then
            temp = Objects(x).ItemName & " Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If


        '***********************************************
        'Shields
        '***********************************************
        If Objects(x).ItemName.IndexOf("Herald of Zakarum") > -1 And Objects(x).ItemQuality = "Unique" Then
            temp = "Hoz " & " Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName.IndexOf("Steelclash Kite Shield") > -1 Then
            temp = "Steelclash " & " Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If


        '***********************************************
        ' Runeword Shields
        '***********************************************
        If Objects(x).ItemName.IndexOf("Splendor") > -1 And Objects(x).ItemBase = "Shield" Then
            temp = Objects(x).ItemName & " Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName = "Lidless Wall Grim Shield" Then
            temp = "Lidless Shield" & " Def " & Objects(x).Defense : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If



        '***********************************************
        'Weapons
        '***********************************************
        If Objects(x).ItemName = "Eschuta's Temper Eldritch Orb" Then
            temp = "Eschuta's" : GoTo Stats
        End If
        If Objects(x).ItemName = "Tal Rasha's Lidless Eye Swirling Crystal" Then
            temp = "Tal's Lidless, " & Objects(x).Stat3 & Objects(x).Stat4 & Objects(x).Stat5
            temp = temp.Replace("(Sorceress Only)", "") : If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If

        '***********************************************
        ' Runeword Weapons
        '***********************************************
        If Objects(x).ItemName.IndexOf("Insight") > -1 And Objects(x).ItemBase = "polearm" Then
            temp = Objects(x).ItemName & " " : temp = temp & Objects(x).Stat2
            If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If
        If Objects(x).ItemName.IndexOf("Spirit") > -1 And Objects(x).RuneWord = True And Objects(x).ItemBase = "Sword" Then
            temp = Objects(x).ItemName & " " : temp = temp & Objects(x).Stat3
            If Objects(x).EtherealItem = True Then temp = temp & "Eth"
            GoTo Abbrev
        End If

        '***********************************************
        ' non Specific items
        '***********************************************
        If Objects(x).ItemName.IndexOf("Hellfire Torch") > -1 Then
            If Objects(x).Stat2.IndexOf("Ama") > -1 Then temp = "Zon "
            If Objects(x).Stat2.IndexOf("Druid") > -1 Then temp = "Druid "
            If Objects(x).Stat2.IndexOf("Pal") > -1 Then temp = "Pala "
            If Objects(x).Stat2.IndexOf("Sorc") > -1 Then temp = "Sorc "
            If Objects(x).Stat2.IndexOf("Necro") > -1 Then temp = "Necro "
            If Objects(x).Stat2.IndexOf("Assa") > -1 Then temp = "Sin "
            temp = temp + "Torch " & Objects(x).Stat3 & " " & Objects(x).Stat4
            temp = temp.Replace(" to all Attributes", " ")
            temp = temp.Replace(" All Resistances", " ")
            temp = temp.Replace("+", "")
            Form1.RichTextBox3.AppendText(temp & vbCrLf & vbCrLf) : Return
        End If

        If Objects(x).ItemName = "Rainbow Facet Jewel" Then temp = "Rainbow Facet  " & Objects(x).Stat3 & " " & Objects(x).Stat4 : GoTo Abbrev

        '***********************************************

        'Rare/magic/white items
        '***********************************************
        temp = Objects(x).ItemName & ", " 'sets the default name

        Select Case (Objects(x).ItemBase)
            Case "Armor", "Helm", "Belt", "Shield", "Boots", "Gloves"
                If IsNumeric(Objects(x).Defense) Then temp = temp & "Def " & Objects(x).Defense
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
        If Objects(x).Stat12 <> "" Then temp = temp & ", " & Objects(x).Stat12
        If Objects(x).Stat13 <> "" Then temp = temp & ", " & Objects(x).Stat13
        If Objects(x).Stat14 <> "" Then temp = temp & ", " & Objects(x).Stat14
        If Objects(x).Stat15 <> "" Then temp = temp & ", " & Objects(x).Stat15

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
        temp = temp.Replace("Paladin Skill Levels", "Pala Skills")
        temp = temp.Replace("Amazon Skill Levels", "Zon Skills")
        temp = temp.Replace("Necromancer Skill Levels", "Nec Skills")
        temp = temp.Replace("Barbarian Skill Levels", "Barb Skills")
        temp = temp.Replace("Druid Skill Levels", "Druid Skills")
        temp = temp.Replace("to Dexterity", "to Dex")
        temp = temp.Replace("Better Chance of Getting Magic Items", "Mf")
        temp = temp.Replace("Ethereal (Cannot be Repaired)", "Eth")
        temp = temp.Replace("Extra Gold from Monsters", "Gf")
        temp = temp.Replace("All Resistances", "Res All")
        temp = temp.Replace("to Experience Gained", "Xp")
        temp = temp.Replace("Faster Cast Rate", "fcr")
        temp = temp.Replace("Enhanced Defense", "Ed")
        temp = temp.Replace("Faster Hit Recovery", "Fhr")
        temp = temp.Replace("Faster Run/Walk", "Frw")
        temp = temp.Replace("Attack Rating", "Ar")
        temp = temp.Replace("Faster Block Rate", "Fbr")
        temp = temp.Replace("Increased Attack Speed", "Ias")
        temp = temp.Replace("Strength", "Str")
        temp = temp.Replace("Life stolen per hit", "Loh")
        temp = temp.Replace("Regenerate Mana", "Mana Regen")
        temp = temp.Replace("Life stolen per hit", "Loh")
        temp = temp.Replace("Meditation Aura When Equipped", "Med")
        temp = temp.Replace("Chains of Honor", "COH")
        temp = temp.Replace("Chance of Crushing Blow", "Cb")

        temp = temp.Replace("Maximum Damage", "Max dmg")
        temp = temp.Replace("Defense", "Def")
        temp = temp.Replace("Fire Skill Damage", "Fire dmg")
        temp = temp.Replace("Increased Chance of Blocking", "Icb")
        temp = temp.Replace("Fire Ball", "Fb")
        temp = temp.Replace("Magic Damage Reduced", "Mdr")
        temp = temp.Replace("Replenish", "Rep")
        temp = temp.Replace("Damage", "Dmg")
        temp = temp.Replace("Character", "Char")
        temp = temp.Replace("Level", "Lvl")
        temp = temp.Replace("damage", "dmg")
        temp = temp.Replace("Lightning", "Lit")
        temp = temp.Replace("lightning", "Lit")
        temp = temp.Replace("Poison and Bone", "PnB")
        temp = temp.Replace("poison", "Psn")
        temp = temp.Replace("Poison", "Psn")
        temp = temp.Replace("seconds", "Secs")
        temp = temp.Replace("Minimum", "Min")
        temp = temp.Replace("Maximum", "Max")
        temp = temp.Replace("Stamina", "Stam")
        temp = temp.Replace("Defensive", "Def")
        temp = temp.Replace("Skeleton", "Skel")

        temp = temp.Replace("Socketed", "Socs")
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
        temp = temp.Replace("Assassin", "Sin")
        temp = temp.Replace("Amazon", "Zon")
        temp = temp.Replace("Only", "")
        temp = temp.Replace(", ,", ",") ' couple extra "," slip in for some reason
        temp = temp.Replace(",,", ",") ' couple extra "," slip in for some reason

        temp = temp.Replace("Adds", "+") ' leave at end of filters/abbreviations


        Form1.RichTextBox3.AppendText(temp & vbCrLf & vbCrLf)
    End Sub



End Module
