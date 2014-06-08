Imports System.Text.RegularExpressions
Module Tradelist
    Public Sub SendToTradeList(ByVal x)
        Dim temp As String = ""

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
        ' torch
        '***********************************************
        If Objects(x).ItemName.IndexOf("Hellfire Torch") > -1 Then
            If Objects(x).Stat2.IndexOf("Ama") > -1 Then temp = "Zon "
            If Objects(x).Stat2.IndexOf("Druid") > -1 Then temp = "Druid "
            If Objects(x).Stat2.IndexOf("Barb") > -1 Then temp = "Barb "
            If Objects(x).Stat2.IndexOf("Pal") > -1 Then temp = "Pala "
            If Objects(x).Stat2.IndexOf("Sorceress") > -1 Then temp = "Sorc "
            If Objects(x).Stat2.IndexOf("Necro") > -1 Then temp = "Necro "
            If Objects(x).Stat2.IndexOf("Assa") > -1 Then temp = "Sin "
            temp = temp + "Torch " & Objects(x).Stat3 & " " & Objects(x).Stat4
            temp = temp.Replace(" to all Attributes", " ")
            temp = temp.Replace(" All Resistances", " ")
            temp = temp.Replace("+", "")
            Form1.RichTextBox3.AppendText(temp & vbCrLf & vbCrLf) : Return
        End If
        '***********************************************
        ' Gheeds
        '***********************************************
        If Objects(x).ItemName = "Gheed's Fortune Grand Charm" Then
            temp = "Gheeds " & Objects(x).Stat1 & " " & Objects(x).Stat2 & " " & Objects(x).Stat3
            temp = temp.Replace("% Extra Gold from Monsters", "")
            temp = temp.Replace("% Better Chance of Getting Magic Items", "")
            temp = temp.Replace("Reduces all Vendor Prices", "")
            temp = temp.Replace("%", "")
            Form1.RichTextBox3.AppendText(temp & vbCrLf & vbCrLf) : Return
        End If

        ' if Identified & set item go to set function
        If Objects(x).ItemQuality = "Set" And Objects(x).Stat1.IndexOf("Unid") = -1 Then
            temp = Set_items(x)
            If Objects(x).Sockets <> "" Then temp = temp & ", Socs " & Objects(x).Sockets
            If Objects(x).EtherealItem = True Then temp = temp & ", Eth"
        End If

        ' if Identified & set item go to Unique function
        If Objects(x).ItemQuality = "Unique" And Objects(x).Stat1.IndexOf("Unid") = -1 Then
            temp = Uniq_items(x) : If Objects(x).Sockets <> "" Then temp = temp & ", Socs " & Objects(x).Sockets
            If Objects(x).EtherealItem = True Then temp = temp & ", Eth"
        End If


        If Objects(x).RuneWord = "True" Then temp = RuneWord_items(x) : If Objects(x).EtherealItem = True Then temp = temp & ", Eth"


        '***********************************************
        'Rare/magic/white items
        '***********************************************
        If temp = "" Then
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
            temp = temp & Stats_items(x)
        End If

        'filters -> abbreviations
        temp = temp.Replace("Magic Damage Reduced by", "Mdr")
        temp = temp.Replace("Damage Reduced by", "Dr")
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
        temp = temp.Replace("Increase Maximum Durability", "Edur")
        temp = temp.Replace("Enhanced Damage", "Edmg")
        temp = temp.Replace("Faster Hit Recovery", "Fhr")
        temp = temp.Replace("Faster Run/Walk", "Frw")
        temp = temp.Replace("Attack Rating", "Ar")
        temp = temp.Replace("Faster Block Rate", "Fbr")
        temp = temp.Replace("Increased Attack Speed", "Ias")
        temp = temp.Replace("Strength", "Str")
        temp = temp.Replace("Vitality", "Vit")
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

    Function Stats_items(ByRef x)
        Dim temp As String = ""
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

        Return temp
    End Function

    Function RuneWord_items(ByRef x)
        Dim temp As String = ""

        '***********************************************
        ' Runeword Armor
        '***********************************************
        If Objects(x).ItemBase = "Armor" Then
            Return Objects(x).ItemName & " Def " & Objects(x).Defense

        End If

        '***********************************************
        ' Runeword Shields
        '***********************************************
        If Objects(x).ItemBase = "Shield" Or Objects(x).ItemBase = "Auric Shield" Then
            If Objects(x).ItemName.IndexOf("Spirit") > -1 Then Return Objects(x).ItemName & " Def " & Objects(x).Defense & ", " & Objects(x).Stat2
            If Objects(x).ItemName.IndexOf("Splendor") > -1 Then Return Objects(x).ItemName & " Def " & Objects(x).Defense
        End If

        '***********************************************
        ' Runeword Weapons
        '***********************************************
        If Objects(x).ItemBase = "polearm" Or Objects(x).ItemBase = "Sword" Or Objects(x).ItemBase = "Mace" Or Objects(x).ItemBase = "Axe" Then
            If Objects(x).ItemName.IndexOf("Beast") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat10
            If Objects(x).ItemName.IndexOf("Breath") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat10 & ", " & Objects(x).Stat11
            If Objects(x).ItemName.IndexOf("Insight") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat2
            If Objects(x).ItemName.IndexOf("Spirit") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat3
            If Objects(x).ItemName.IndexOf("Oak") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat11
            If Objects(x).ItemName.IndexOf("Arms") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat6 & ", " & Objects(x).Stat7 & ", " & Objects(x).Stat8
            If Objects(x).ItemName.IndexOf("Grief") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat2
            If Objects(x).ItemName.IndexOf("Infinity") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat6
            If Objects(x).ItemName.IndexOf("Last Wish") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat8
            If Objects(x).ItemName.IndexOf("Obedience") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat12
            If Objects(x).ItemName.IndexOf("Oath") > -1 Then
                If Objects(x).Stat2 = "Indestructible" Then Return Objects(x).ItemName & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat10
                Return Objects(x).ItemName & Objects(x).Stat6 & ", " & Objects(x).Stat11
            End If
            If Objects(x).ItemName.IndexOf("Justice") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat6

        End If

        If Objects(x).ItemBase = "Staff" Then
            If Objects(x).ItemName.IndexOf("Memory") > -1 Then Return Objects(x).ItemName & ", " & Objects(x).Stat6 & ", " & Objects(x).Stat7

        End If

        Return Objects(x).ItemName & " not found"
    End Function


    Function Uniq_items(ByRef x)
        Dim temp As String = ""
        '***********************************************
        'Unique items
        If Objects(x).ItemName = "Rainbow Facet Jewel" Then Return "Rainbow Facet  " & Objects(x).Stat3 & ", " & Objects(x).Stat4
        '***********************************************
        'Rings
        '***********************************************
        If Objects(x).ItemBase = "Ring" Then
            If Objects(x).ItemName = "Bul-Kathos' Wedding Band Ring" Then Return "BK Ring " & Objects(x).Stat2
            If Objects(x).ItemName = "Nagelring Ring" Then Return "Nagel Ring, " & Objects(x).Stat4
            If Objects(x).ItemName = "The Stone of Jordan Ring" Then Return "Soj"
            If Objects(x).ItemName = "Carrion Wind Ring" Then Return "Nagel Ring, " & Objects(x).Stat3
            If Objects(x).ItemName = "Dwarf Star Ring" Then Return Objects(x).ItemName
            If Objects(x).ItemName = "Manald Heal Ring" Then Return "Manald Heal Ring, "
            If Objects(x).ItemName = "Raven Frost Ring" Then Return "Raven Frost, " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Wisp Projector Ring" Then Return "Wisp Ring, " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Nature's Peace Ring" Then Return "Nature's Peace Ring, " & Objects(x).Stat3
        End If
        '***********************************************
        'Amulets
        '***********************************************
        If Objects(x).ItemBase = "Amulet" Then
            If Objects(x).ItemName = "Atma's Scarab Amulet" Then Return "Atma's Amulet, "
            If Objects(x).ItemName = "Cresent Moon Amulet" Then Return "Cresent Amulet, "
            If Objects(x).ItemName = "Highlord's Wrath Amulet" Then Return "Highlord's Amulet, "
            If Objects(x).ItemName = "Mara's Kaleidoscope Amulet" Then Return "Mara's, " & Objects(x).Stat3
            If Objects(x).ItemName = "Metalgrid Amulet" Then Return "Metal Grid Amulet, "
            If Objects(x).ItemName = "Nokozan Relic Amulet" Then Return "Nokozan Relic, "
            If Objects(x).ItemName = "Saracen's Chance Amulet" Then Return "Saracen's Amulet, "
            If Objects(x).ItemName = "Seraph's Hymn Amulet" Then Return "Seraph's Hymn Amulet, "
            If Objects(x).ItemName = "The Cat's Eye Amulet" Then Return "Cat's Eye Amulet, "
            If Objects(x).ItemName = "The Eye of Etlich Amulet" Then Return "Etlich Amulet, " & Objects(x).Stat3
            If Objects(x).ItemName = "The Mahim-Oak Curio Amulet" Then Return "Mahim-Oak Amulet, "
            If Objects(x).ItemName = "The Rising Sun Amulet" Then Return "Rising Sun Amulet, "
        End If
        '***********************************************
        ' Unique Helms
        '***********************************************

        If Objects(x).ItemBase = "Helm" Or Objects(x).ItemBase = "Circlet" Or Objects(x).ItemBase = "Primal Helm" Then
            If Objects(x).ItemName = "Andariel's Visage Demonhead" Then Return "Andies, " & Objects(x).Stat4 & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Arreat's Face Slayer Guard" Then Return "Arreats, " & Objects(x).Stat4 & " " & Objects(x).Stat5
            If Objects(x).ItemName = "Arreat's Face Guardian Crown" Then Return "Arreats, " & Objects(x).Stat4 & " " & Objects(x).Stat5
            If Objects(x).ItemName = "Biggin's Bonnet Cap" Then Return "Biggins, " & Objects(x).Stat3
            If Objects(x).ItemName = "Blackhorn's Face Death Mask" Then Return "Blackhorns, " & Objects(x).Stat3
            If Objects(x).ItemName = "Cerebus' Bite Blood Spirit" Then Return "Cerebus, " & Objects(x).Stat3
            If Objects(x).ItemName = "Coif of Glory Helm" Then Return "Coif of Glory, " & Objects(x).Stat2
            If Objects(x).ItemName = "Crown of Ages Corona" Then Return "COA, " & Objects(x).Defense & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Crown of Thieves Corona" Then Return "Crown of Thieves, " & Objects(x).Defense & ", " & Objects(x).Stat1
            If Objects(x).ItemName = "Darksight Helm Bassinet" Then Return "Darksight, " & Objects(x).Stat4
            If Objects(x).ItemName = "Demonhorn's Edge Destroyer Helm" Then Return "Demonhorn, " & Objects(x).Stat1 & ", " & Objects(x).Stat2 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Duskdeep Full Helm" Then Return "Duskdeep, " & Objects(x).Defense
            If Objects(x).ItemName = "Giant Skull Bone Visage" Then Return "Giant Skull, " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Griffon's Eye Diadem" Then Return "Griffon's Eye, " & Objects(x).Defense & ", " & Objects(x).Stat3 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Halaberd's Reign Conqueror Crown" Then Return "Halaberd, " & Objects(x).Defense & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Harlequin Crest Shako" Then Return "Shako, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Howltusk Great Helm" Then Return "Howltusk, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Jalal's Mane Totemic Mask" Then Return "Jalal's, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Kira's Guardian Tiara" Then Return "Kira's, Def " & Objects(x).Defense & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Kira's Guardian Diadem" Then Return "Kira's, Def " & Objects(x).Defense & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Nightwing's Veil Spired Helm" Then Return "Nightwings, Def " & Objects(x).Defense & ", " & Objects(x).Stat2 & ", " & Objects(x).Stat3 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Peasant Crown War Hat" Then Return "Peasant Hat , Def " & Objects(x).Defense
            If Objects(x).ItemName = "Peasant Crown Shako" Then Return "Peasant Shako , Def " & Objects(x).Defense
            If Objects(x).ItemName = "Ravenlore Sky Spirit" Then Return "Ravenlore, Def " & Objects(x).Defense & ", " & Objects(x).Stat2 & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Rockstopper Sallet" Then Return "Rockstopper, Def " & Objects(x).Defense & ", " & Objects(x).Stat3 & " " & Objects(x).Stat4 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Spirit Keeper Earth Spirit" Then Return "Spirit Keeper, Def " & Objects(x).Defense & ", " & Objects(x).Stat1 & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Steelskull Casque" Then Return "Steelskull, Def " & Objects(x).Defense & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Steel Shade Armet" Then Return "Steel Shade, Def " & Objects(x).Defense & ", " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "The Face of Horror Mask" Then Return "Face of Horror, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Undead Crown Crown" Then Return "Undead Crown, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Valkyrie Wing Winged Helm" Then Return "Valkyrie, Def " & Objects(x).Defense & ", " & Objects(x).Stat1
            If Objects(x).ItemName = "Vampire Gaze Grim Helm" Then Return "Vampire Gaze, Def " & Objects(x).Defense & ", " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Vampire Gaze Bone Visage" Then Return "Vampire Gaze, Def " & Objects(x).Defense & ", " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Veil of Steel Spired Helm" Then Return "Veil of Steel, Def " & Objects(x).Defense & ", " & Objects(x).Stat3 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Wolfhowl Fury Visor" Then Return "Wolfhowl, " & Stats_items(x)
            If Objects(x).ItemName = "Wormskull Bone Helm" Then Return "Wolfhowl, Def " & Objects(x).Defense
        End If
        '***********************************************
        'Belts
        '***********************************************
        If Objects(x).ItemBase = "Belt" Then
            If Objects(x).ItemName = "Arachnid Mesh Spiderweb Sash" Then Return "Arach, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Bladebuckle Plated Belt" Then Return "Arach, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Gloom's Trap Mesh Belt" Then Return "Gloom's Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Goldwrap Heavy Belt" Then Return "Goldwrap, Def " & Objects(x).Defense & ", " & Stats_items(x)
            If Objects(x).ItemName = "Goldwrap Troll Belt" Then Return "Goldwrap TB, " & Objects(x).Defense & ", " & Stats_items(x)
            If Objects(x).ItemName = "Lenymo Spiderweb Sash" Then Return "Lenymo Sash, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Nightsmoke" Then Return "Nightsmoke, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Nosferatu's Coil Vampirefang Belt" Then Return "Nosferatu's, Def " & Objects(x).Defense & ", " & Objects(x).Stat2
            If Objects(x).ItemName = "Razortail Vampirefang Belt" Then Return "Razortail, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Snakecord Light Belt" Then Return "Snakecord, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Snowclash Battle Belt" Then Return "Snakecord, Def " & Objects(x).Defense & ", " & Objects(x).Stat1
            If Objects(x).ItemName = "String of Ears Demonhide Sash" Then Return "String of Ears, " & Stats_items(x)
            If Objects(x).ItemName = "Thundergod's Vigor War Belt" Then Return "Thundergod's Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Thundergod's Vigor Colossus Girdle" Then Return "Thundergod's Girdle, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Verdungo's Hearty Cord Mithril Coil" Then Return "Dungo's, Def " & Objects(x).Defense & ", " & Objects(x).Stat3 & ", " & Objects(x).Stat4
        End If

        '***********************************************
        'Gloves
        '***********************************************
        If Objects(x).ItemBase = "Gloves" Then
            If Objects(x).ItemName = "Bloodfist Heavy Gloves" Then Return "Bloodfist, " & Objects(x).Stat5
            If Objects(x).ItemName = "Bloodfist Vampirebone Gloves" Then Return "Bloodfist VG, " & Objects(x).Stat5
            If Objects(x).ItemName = "Chance Guards Chain Gloves" Then Return "Chancies, " & Objects(x).Stat5
            If Objects(x).ItemName = "Chance Guards Vambraces" Then Return "Chancies Upp'd, " & Objects(x).Stat5
            If Objects(x).ItemName = "Dracul's Grasp Vampirebone Gloves" Then Return "Dracs, " & Objects(x).Stat2 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Frostburn Gauntlets" Then Return "Frostburn, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Ghoulhide Heavy Bracers" Then Return "Ghoulhide, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Gravepalm Sharkskin Gloves" Then Return "Gravepalm, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hellmouth War Gauntlets" Then Return "Hellmouth, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Magefist Light Gauntlets" Then Return "Magefist, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Magefist Crusader Gauntlets" Then Return "Magefist Crusader, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Magefist Battle Gauntlets" Then Return "Magefist Battle, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Soul Drainer Vambraces" Then Return "Soul Drainer, Def " & Objects(x).Defense & ", " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Steelrend Ogre Gauntlets" Then Return "Steelrend, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "The Hand of Brock" Then Return "Hand of Brock, " & Objects(x).Stat3
            If Objects(x).ItemName = "Venom Grip Demonhide Gloves" Then Return "Venom Grip, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lava Gout Crusader Gauntlets" Then Return "Lava Gout Crusader Gauntlets, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lava Gout Battle Gauntlets" Then Return "Lava Gout Battle Gauntlets, Def " & Objects(x).Defense & ", " & Objects(x).Stat3
        End If

        '***********************************************
        'Boots
        '***********************************************
        If Objects(x).ItemBase = "Boots" Then
            If Objects(x).ItemName = "Goblin Toe Light Plated Boots" Then Return "Goblin Toe, Def " & Objects(x).Defense & ", " & Objects(x).Stat8
            If Objects(x).ItemName = "Goblin Toe Mirrored Boots" Then Return "Goblin Toe, Def " & Objects(x).Defense & ", " & Objects(x).Stat8
            If Objects(x).ItemName = "Gore Rider War Boots" Then Return "Gore Rider's, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Gore Rider Myrmidon Greaves" Then Return "Gore Rider's Upp'd, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Gorefoot Heavy Boots" Then Return "Gorefoot, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hotspur Boots" Then Return "Hotspur, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hotspur Wyrmhide Boots" Then Return "Hotspur, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Marrowwalk Boneweave Boots" Then Return "Marrowwalk, Def " & Objects(x).Defense & ", " & Objects(x).Stat2
            If Objects(x).ItemName = "Sandstorm Trek Scarabshell Boots" Then Return "Treks, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Shadow Dancer Myrmidon Greaves" Then Return "Shadow Dancer, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Silkweave Mesh Boots" Then Return "Silkweave, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Silkweave Boneweave Boots" Then Return "Silkweave, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Tearhaunch Greaves" Then temp = "Tearhaunch, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Treads of Cthon Chain Boots" Then Return "Treads of Cthon, Def " & Objects(x).Defense
            If Objects(x).ItemName = "War Traveler Battle Boots" Then Return "War Travs, Def " & Objects(x).Defense & ", " & Objects(x).Stat8
            If Objects(x).ItemName = "Waterwalk Sharkskin Boots" Then Return "Waterwalks, Def " & Objects(x).Defense & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Waterwalk Scarabshell Boots" Then Return "Waterwalks, Def " & Objects(x).Defense & ", " & Objects(x).Stat5
        End If

        '***********************************************
        ' Unique Armor
        '***********************************************
        If Objects(x).ItemBase = "Armor" Then
            If Objects(x).ItemName = "Arkaine's Valor Balrog Skin" Then Return "Arkaine's Valor, Def " & Objects(x).Defense & ", " & Objects(x).Stat1
            If Objects(x).ItemName = "Atma's Wail Embossed Plate" Then Return "Atma's Wail, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Black Hades Chaos Armor" Then Return "Black Hades, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Blinkbat's Form Leather Armor" Then Return "Blinkbat's, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Boneflesh Plate Mail" Then Return "Boneflesh, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Corpsemourn Ornate Plate" Then Return "Corpsemourn, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Crow Caw Tigulated Mail" Then Return "Crow Caw, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Crow Caw Loricated Mail" Then Return "Crow Caw, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Darkglow Ring Mail" Then Return "Darkglow, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Duriel's Shell Cuirass" Then Return "Duriel's Shell, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Duriel's Shell Great Hauberk" Then Return "Duriel's Shell, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Goldskin Full Plate Mail" Then Return "Goldskin, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Goldskin Shadow Plate" Then Return "Goldskin, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Greyform Quilted Armor" Then Return "Greyform, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Guardian Angel Templar Coat" Then Return "Guardian Angel, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Guardian Angel Hellforge Plate" Then Return "Guardian Angel, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hawkmail Scale Mail" Then Return "Hawkmail, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hawkmail Loricated Mail" Then Return "Hawkmail, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Heavenly Garb Light Plate" Then Return "Heavenly Garb, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Iceblink Splint Mail" Then Return "Iceblink, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Leviathan Kraken Shell" Then Return "Leviathan, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Ormus' Robes Dusk Shroud" Then Return "Ormus Robes, " & Stats_items(x)
            If Objects(x).ItemName = "Que-Hegan's Wisdom Mage Plate" Then Return "Que-Hegan's, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Rattlecage Gothic Plate" Then Return "Rattlecage, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Rockfleece Field Plate" Then Return "Rockfleece, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Shaftstop Mesh Armor" Then Return "Shaftstop Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Shaftstop Boneweave" Then Return "Shaftstop bw Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Silks of the Victor Ancient Armor" Then Return "Silks of the Victor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Skin of the Flayed One Demonhide Armor" Then Return "Skin of the Flayed One, Def " & Objects(x).Defense & ", " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Skin of the Flayed One Scarab Husk" Then Return "Skin of the Flayed One, Def " & Objects(x).Defense & ", " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Skin of the Vipermagi Serpentskin Armor" Then Return "Vipermagi,  " & Objects(x).Stat4
            If Objects(x).ItemName = "Skin of the Vipermagi Wyrmhide" Then Return "Vipermagi, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Skullder's Ire Russet Armor" Then Return "Skullders,  " & Objects(x).Defense
            If Objects(x).ItemName = "Skullder's Ire Balrog Skin" Then Return "Skullders,  " & Objects(x).Defense
            If Objects(x).ItemName = "Sparking Mail Chain Mail" Then Return "Sparking Mail,  " & Objects(x).Defense
            If Objects(x).ItemName = "Spirit Forge Diamond Mail" Then Return "Spirit Forge,  " & Objects(x).Defense
            If Objects(x).ItemName = "Steel Carapace Shadow Plate" Then Return "Steel Carapace,  " & Objects(x).Defense
            If Objects(x).ItemName = "Templar's Might Sacred Armor" Then Return "Templar's Might,  " & Objects(x).Defense & ", " & Objects(x).Stat1 & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "The Centurion Hard Leather Armor" Then Return "The Centurion,  " & Objects(x).Defense
            If Objects(x).ItemName = "The Gladiator's Bane Wire Fleece" Then Return "Gladiator's Bane,  " & Objects(x).Defense
            If Objects(x).ItemName = "The Spirit Shroud Ghost Armor" Then Return "Spirit Shroud,  " & Objects(x).Defense
            If Objects(x).ItemName = "Toothrow Sharktooth Armor" Then Return "Toothrow,  " & Objects(x).Defense
            If Objects(x).ItemName = "Twitchthroe Studded Leather" Then Return "Twitchthroe SL,  " & Objects(x).Defense
            If Objects(x).ItemName = "Twitchthroe Wire Fleece" Then Return "Twitchthroe wf,  " & Objects(x).Defense
            If Objects(x).ItemName = "Tyrael's Might Sacred Armor" Then Return "Tyrael's Might,  " & Objects(x).Defense & ", " & Objects(x).Stat6 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Venom Ward Breast Plate" Then Return "Venom Ward,  " & Objects(x).Defense
            If Objects(x).ItemName = "Venom Ward Great Hauberk" Then Return "Venom Ward,  " & Objects(x).Defense
        End If

        '***********************************************
        'Shields
        '***********************************************
        If Objects(x).ItemBase = "Shield" Or Objects(x).ItemBase = "Auric Shield" Or Objects(x).ItemBase = "Voodoo Heads" Then
            If Objects(x).ItemName = "Alma Negra Sacred Rondache" Then Return "Alma Negra, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Blackoak Shield Luna" Then Return "Blackoak Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Boneflame Succubus Skull" Then Return "Boneflame, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Bverrit Keep Tower Shield" Then Return "Bverrit Keep, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Darkforce Spawn Bloodlord Skull" Then Return "Darkforce, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Dragonscale Zakarum Shield" Then Return "Dragonscale, Def " & Objects(x).Defense & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Gerke's Sanctuary Pavise" Then Return "Gerke's Sanctuary, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Head Hunter's Glory Troll Nest" Then Return "Head Hunter's Glory, Def " & Objects(x).Defense & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Herald of Zakarum Gilded Shield" Then Return "Hoz, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Herald of Zakarum Zakarum Shield" Then Return "Hoz, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Homunculus Hierophant Trophy" Then Return "Homunculus, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lance Guard Barbed Shield" Then Return "Lance Guard, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lance Guard Blade Barrier" Then Return "Lance Guard, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lidless Wall Grim Shield" Then Return "Lidless Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lidless Wall Troll Nest" Then Return "Lidless TN, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Medusa's Gaze Aegis" Then Return "Medusa's Gaze, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Moser's Blessed Circle Round Shield" Then Return "Moser's Blessed Circle, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Pelta Lunata Buckler" Then Return "Pelta Lunata, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Spike Thorn Blade Barrier" Then Return "Pelta Lunata, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Spirit Ward Ward" Then Return "Spirit Ward, Def " & Objects(x).Defense & ", " & Objects(x).Stat3 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Steelclash Kite Shield" Then Return "Steelclash, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Stormguild Large Shield" Then Return "Stormguild, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Stormshield Monarch" Then Return "Stormshield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Swordback Hold Spiked Shield" Then Return "Swordback Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Swordback Hold Blade Barrier" Then Return "Swordback Barrier, Def " & Objects(x).Defense
            If Objects(x).ItemName = "The Ward Gothic Shield" Then Return "The Ward, Def " & Objects(x).Defense & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Tiamat's Rebuke Dragon Shield" Then Return "Tiamat's Rebuke, Def " & Objects(x).Defense & ", " & Objects(x).Stat8
            If Objects(x).ItemName = "Umbral Disk Small Shield" Then Return "Umbral Disk, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Visceratuant Defender" Then Return "Visceratuant, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Wall of the Eyeless Bone Shield" Then Return "Wall of the Eyeless, Def " & Objects(x).Defense
        End If

        '***********************************************
        'Weapons Axes
        '***********************************************
        If Objects(x).ItemBase = "Axe" Then
            If Objects(x).ItemName = "Axe of Fechmar" Then Return "Fechmars Axe, " & Objects(x).Stat1
            If Objects(x).ItemName = "Bladebone Double Axe" Then Return "Bladebone, " & Objects(x).Stat2
            If Objects(x).ItemName = "Boneslayer Blade Gothic Axe" Then Return "Boneslayer, " & Objects(x).Stat1
            If Objects(x).ItemName = "Brainhew Great Axe" Then Return "Brainhew, " & Objects(x).Stat1 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Butcher's Pupil Cleaver" Then Return "Butcher's Pupil, " & Objects(x).Stat3
            If Objects(x).ItemName = "Cranebeak War Spike" Then Return "Cranebeak, " & Objects(x).Stat2 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Death Cleaver Berserker Axe" Then Return "Death Cleaver, " & Objects(x).Stat2 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Deathspade Axe" Then Return "Deathspade, " & Objects(x).Stat1
            If Objects(x).ItemName = "Ethereal Edge Silver-edged Axe" Then Return "Ethereal Edge, " & Objects(x).Stat3 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Executioner's Justice Glorious Axe" Then Return "Executioner's Justice, " & Objects(x).Stat3
            If Objects(x).ItemName = "Gimmershred Flying Axe" Then Return "Gimmershred, " & Objects(x).Stat2
            If Objects(x).ItemName = "Goreshovel Broad Axe" Then Return "Goreshovel, " & Objects(x).Stat2
            If Objects(x).ItemName = "Hellslayer Decapitator" Then Return "Hellslayer"
            If Objects(x).ItemName = "Homongous Giant Axe" Then Return "Homongous, " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Islestrike Twin Axe" Then Return "Islestrike, " & Objects(x).Stat2
            If Objects(x).ItemName = "Lacerator Winged Axe" Then Return "Lacerator, " & Objects(x).Stat3
            If Objects(x).ItemName = "Messerschmidt's Reaver Champion Axe" Then Return "Messerschmidt's"
            If Objects(x).ItemName = "Pompeii's Wrath War Spike" Then Return "Pompeii Spike, " & Objects(x).Stat2
            If Objects(x).ItemName = "Rakescar War Axe" Then Return "Rakescar, " & Objects(x).Stat2
            If Objects(x).ItemName = "Razor's Edge Tomahawk" Then Return "Razor's Edge, " & Objects(x).Stat2
            If Objects(x).ItemName = "Rune Master Ettin Axe" Then Return "Rune Master, " & Objects(x).Stat1
            If Objects(x).ItemName = "Skull Splitter Military Pick" Then Return "Skull Splitter, " & Objects(x).Stat3
            If Objects(x).ItemName = "Stormrider Tabar" Then Return "Stormrider"
            If Objects(x).ItemName = "The Chieftain Battle Axe" Then Return "The Chieftain, " & Objects(x).Stat4
            If Objects(x).ItemName = "The Gnasher Hand Axe" Then Return "The Gnasher, " & Objects(x).Stat1
            If Objects(x).ItemName = "The Minotaur Ancient Axe" Then Return "The Minotaur, " & Objects(x).Stat1 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "The Scalper Francisca" Then Return "The Scalper, " & Objects(x).Stat2 & ", " & Objects(x).Stat4
        End If

        '***********************************************
        'Weapons Bows
        '***********************************************
        If Objects(x).ItemBase = "Bow" Then
            If Objects(x).ItemName = "Blastbark Long War Bow" Then Return "Blastbark, " & Objects(x).Stat2
            If Objects(x).ItemName = "Blood Raven's Charge Matriarchal Bow" Then Return "Blood Raven's Charge, " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Cliffkiller Large Siege Bow" Then Return "Cliffkiller, " & Objects(x).Stat2
            If Objects(x).ItemName = "Eaglehorn Crusader Bow" Then Return "Eaglehorn, " & Objects(x).Stat2
            If Objects(x).ItemName = "Goldstrike Arch Gothic Bow" Then Return "Goldstrike, " & Objects(x).Stat3 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Hellclap Short War Bow" Then Return "Hellclap, " & Objects(x).Stat3 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Kuko Shakaku Cedar Bow" Then Return "Kuko Shakaku, " & Objects(x).Stat4
            If Objects(x).ItemName = "Lycander's Aim Ceremonial Bow" Then Return "Lycander's, " & Objects(x).Stat4 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Magewrath Rune Bow" Then Return "Magewrath, " & Objects(x).Stat2 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Pluckeye Short Bow" Then Return "Pluckeye"
            If Objects(x).ItemName = "Raven Claw Long Bow" Then Return "Raven Claw, " & Objects(x).Stat2
            If Objects(x).ItemName = "Riphook Razor Bow" Then Return "Riphook RB, " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Riphook Blade Bow" Then Return "Riphook BB, " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Rogue's Bow Composite Bow" Then Return "Rogue's Bow, " & Objects(x).Stat2
            If Objects(x).ItemName = "Stormstrike Short Battle Bow" Then Return "Stormstrike, " & Objects(x).Stat2
            If Objects(x).ItemName = "Widowmaker Ward Bow" Then Return "Widowmaker, " & Objects(x).Stat2 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Windforce Hydra Bow" Then Return "Windforce, " & Objects(x).Stat4
            If Objects(x).ItemName = "Witchwild String Short Siege Bow" Then Return "Witchwild, " & Objects(x).Stat3
            If Objects(x).ItemName = "Witherstring Hunter's Bow" Then Return "Witherstring, " & Objects(x).Stat1
            If Objects(x).ItemName = "Wizendraw Long Battle Bow" Then Return "Wizendraw, " & Objects(x).Stat3
        End If

        '***********************************************
        'Weapons Claws
        '***********************************************
        If Objects(x).ItemBase = "Assassin Claw" Then
            If Objects(x).ItemName = "Bartuc's Cut-Throat Greater Talons" Then Return "Bartuc's, " & Objects(x).Stat4 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Firelizard's Talons Feral Claws" Then Return "Firelizard's, " & Objects(x).Stat3 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Jade Talon Wrist Sword" Then Return "Jade Talon, " & Objects(x).Stat4 & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Shadow Killer Battle Cestus" Then Return "Shadow Killer, " & Objects(x).Stat3
            If Objects(x).ItemName = "Shadow Killer Battle Cestus" Then Return "Shadow Killer, " & Objects(x).Stat3


        End If

        '***********************************************
        'Weapons Crossbows
        '***********************************************
        If Objects(x).ItemBase = "Crossbow" Then
            If Objects(x).ItemName = "Buriza-Do Kyanon Ballista" Then Return "Buriza-Do Kyanon, " & Objects(x).Stat3
            If Objects(x).ItemName = "Buriza-Do Kyanon Colossus Crossbow" Then Return "Buriza-Do Kyanon, " & Objects(x).Stat3
            If Objects(x).ItemName = "Doomslinger Repeating Crossbow" Then Return "Doomslinger, " & Objects(x).Stat4
            If Objects(x).ItemName = "Hellcast Heavy Crossbow" Then Return "Hellcast, " & Objects(x).Stat3
            If Objects(x).ItemName = "Hellrack Colossus Crossbow" Then Return "Hellrack, " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Ichorsting Crossbow" Then Return "Ichorsting"
            If Objects(x).ItemName = "Leadcrow Light Crossbow" Then Return "Leadcrow"
        End If
        '***********************************************
        'Weapons Daggers
        '***********************************************
        If Objects(x).ItemBase = "Dagger" Or Objects(x).ItemBase = "Throwing Knife" Then
            If Objects(x).ItemName = "Blackbog's Sharp Cinquedeas" Then Return "Blackbog's "
            If Objects(x).ItemName = "Deathbit Battle Dart" Then Return "Deathbit, " & Objects(x).Stat1 & ", " & Objects(x).Stat2 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Fleshripper Fanged Knife" Then Return "Fleshripper, " & Objects(x).Stat1
            If Objects(x).ItemName = "Ghostflame Legend Spike" Then Return "Ghostflame, " & Objects(x).Stat2 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Gull Dagger" Then Return "Gull Dagger"
            If Objects(x).ItemName = "Heart Carver Rondel" Then Return "Heart Carver, " & Objects(x).Stat1
            If Objects(x).ItemName = "Spectral Shard Blade" Then Return "Spectral Shard"
            If Objects(x).ItemName = "Spineripper Poignard" Then Return "Spineripper, " & Objects(x).Stat3
            If Objects(x).ItemName = "The Diggler Dirk" Then Return "The Diggler"
            If Objects(x).ItemName = "The Jade Tan Do Kris" Then Return "The Jade Tan Do, " & Objects(x).Stat1
            If Objects(x).ItemName = "Wizardspike" Then Return "Wizardspike"
            If Objects(x).ItemName = "Warshrike Winged Knife" Then Return "Warshrike, " & Objects(x).Stat4
        End If
        '***********************************************
        'Weapons Javelins
        '***********************************************
        If Objects(x).ItemBase = "Javelin" Or Objects(x).ItemBase = "Amazon Javelin" Then
            If Objects(x).ItemName = "Demon's Arch Balrog Spear" Then Return "Demon's Arch, " & Objects(x).Stat2 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Gargoyle's Bite Winged Harpoon" Then Return "Gargoyle's Bite, " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Thunderstroke Matriarchal Javelin" Then Return "Thunderstroke, " & Objects(x).Stat2 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Titan's Revenge Ceremonial Javelin" Then Return "Titan's Revenge CJ, " & Objects(x).Stat4 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Titan's Revenge Matriarchal Javelin" Then Return "Titan's Revenge MJ, " & Objects(x).Stat4 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Wraith Flight Ghost Glaive" Then Return "Wraith Flight, " & Objects(x).Stat1 & ", " & Objects(x).Stat2
        End If

        '***********************************************
        'Weapons Maces
        '***********************************************
        If Objects(x).ItemBase = "Mace" Or Objects(x).ItemBase = "Club" Then
            If Objects(x).ItemName = "Baezil's Vortex Knout" Then Return "Baezil's Vortex, " & Objects(x).Stat3
            If Objects(x).ItemName = "Baranar's Star Devil Star" Then Return "Baranar's Star"
            If Objects(x).ItemName = "Bloodrise Morning Star" Then Return "Bloodrise"
            If Objects(x).ItemName = "Bloodtree Stump War Club" Then Return "Bloodtree Stump, " & Objects(x).Stat2
            If Objects(x).ItemName = "Bonesnap Maul" Then Return "Bonesnap, " & Objects(x).Stat1
            If Objects(x).ItemName = "Crushflange Mace" Then Return "Crushflangep, " & Objects(x).Stat1
            If Objects(x).ItemName = "Demon Limb Tyrant Club" Then Return "Demon Limb, " & Objects(x).Stat1 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Earth Shifter Thunder Maul" Then Return "Earth Shifter, " & Objects(x).Stat5
            If Objects(x).ItemName = "Earthshaker" Then Return "Earthshaker"
            If Objects(x).ItemName = "Felloak Club" Then Return "Felloak, " & Objects(x).Stat1
            If Objects(x).ItemName = "Fleshrender Barbed Club" Then Return "Fleshrender, " & Objects(x).Stat3
            If Objects(x).ItemName = "Horizon's Tornado Scourge" Then Return "Horizon's Tornado, " & Objects(x).Stat3
            If Objects(x).ItemName = "Ironstone War Hammer" Then Return "Ironstone, " & Objects(x).Stat1 & ", " & Objects(x).Stat2
            If Objects(x).ItemName = "Schaefer's Hammer Legendary Mallet" Then Return "Schaefer's Hammer, " & Objects(x).Stat4
            If Objects(x).ItemName = "Nord's Tenderizer Truncheon" Then Return "Nord's Tenderizer, " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Steeldriver Great Maul" Then Return "Steeldriver, " & Objects(x).Stat2
            If Objects(x).ItemName = "Stone Crusher Legendary Mallet" Then Return "Stone Crusher, " & Objects(x).Stat1 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Stormlash Scourge" Then Return "Stormlash, " & Objects(x).Stat4 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Stoutnail Spiked Club" Then Return "Stoutnail"
            If Objects(x).ItemName = "The Cranium Basher Thunder Maul" Then Return "The Cranium Basher, " & Objects(x).Stat4
            If Objects(x).ItemName = "The Gavel of Pain Martel de Fer" Then Return "The Gavel of Pain, " & Objects(x).Stat4
            If Objects(x).ItemName = "The General's Tan Do Li Ga Flail" Then Return "The General's Flail, " & Objects(x).Stat2
            If Objects(x).ItemName = "The General's Tan Do Li Ga Scourge" Then Return "The General's Scourge, " & Objects(x).Stat2
            If Objects(x).ItemName = "Windhammer Ogre Maul" Then Return "Windhammer, " & Objects(x).Stat3
        End If


        '***********************************************
        'Weapons orbs
        '***********************************************
        If Objects(x).ItemBase = "Orb" Then
            If Objects(x).ItemName = "Death's Fathom Dimensional Shard" Then Return "Death's Fathom, " & Objects(x).Stat3 & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Eschuta's Temper Eldritch Orb" Then Return "Eschuta's, " & Objects(x).Stat1 & ", " & Objects(x).Stat3 & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "The Oculus Swirling Crystal" Then Return "Oculus"
        End If

        '***********************************************
        'Weapons Polearms
        '***********************************************
        If Objects(x).ItemBase = "polearm" Then
            If Objects(x).ItemName = "Bonehew Ogre Axe" Then Return "Bonehew, " & Objects(x).Stat3
            If Objects(x).ItemName = "Dimoak's Hew " Then Return "Dimoak's Hew, "
            If Objects(x).ItemName = "Pierre Tombale Couant Partizan" Then Return "Pierre Tombale, " & Objects(x).Stat3
            If Objects(x).ItemName = "Soul Harvest Scythe" Then Return "Soul Harvest, " & Objects(x).Stat1
            If Objects(x).ItemName = "Steelgoad Voulge" Then Return "Steelgoad, " & Objects(x).Stat1
            If Objects(x).ItemName = "Stormspire Giant Thresher" Then Return "Stormspire, " & Objects(x).Stat5
            If Objects(x).ItemName = "The Battlebranch Poleaxe" Then Return "The Battlebranch, " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "The Grim Reaper War Scythe" Then Return "The Grim Reaper, " & Objects(x).Stat3
            If Objects(x).ItemName = "The Reaper's Toll Thresher" Then Return "Reaper's Toll, " & Objects(x).Stat2 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Tomb Reaver Cryptic Axe" Then Return "Tomb Reaver, " & Objects(x).Stat2 & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat8
            If Objects(x).ItemName = "Woestave Halberd" Then Return "Woestave, " & Objects(x).Stat1
        End If

        '***********************************************
        'Weapons Sceptors
        '***********************************************
        If Objects(x).ItemBase = "Scepter" Then
            If Objects(x).ItemName = "Astreon's Iron Ward Caduceus" Then Return "Astreon's Iron Ward, " & Objects(x).Stat1 & ", " & Objects(x).Stat3 & ", " & Objects(x).Stat9
            If Objects(x).ItemName = "Hand of Blessed Light Divine Sceptor" Then Return "Hand of Blessed Light, " & Objects(x).Stat3
            If Objects(x).ItemName = "Heaven's Light Mighty Scepter" Then Return "Heaven's Light, " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Knell Striker Scepter" Then Return "Knell Striker, " & Objects(x).Stat1
            If Objects(x).ItemName = "Rusthandle Grand Scepter" Then Return "Rusthandle, " & Objects(x).Stat2 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Stormeye War Scepter" Then Return "Stormeye, " & Objects(x).Stat1 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "The Fetid Sprinkler Holy Water Sprinkler" Then Return "The Fetid Sprinkler, " & Objects(x).Stat4 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "The Redeemer Mighty Scepter" Then Return "The Redeemer, " & Objects(x).Stat2 & ", " & Objects(x).Stat6 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Zakarum Hand Rune Scepter" Then Return "Zakarum Hand, " & Objects(x).Stat3
        End If

        '***********************************************
        'Weapons Spears
        '***********************************************
        If Objects(x).ItemBase = "Spear" Then
            If Objects(x).ItemName = "Arioc's Needle Hyperion Spear" Then Return "Arioc's Needle, " & Objects(x).Stat1 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Bloodthief Brandistock" Then Return "Bloodthief, " & Objects(x).Stat1 & ", " & Objects(x).Stat2
            If Objects(x).ItemName = "Hone Sundan Yari" Then Return "Hone Sundan, " & Objects(x).Stat1
            If Objects(x).ItemName = "Kelpie Snare Fuscina" Then Return "Kelpie Snare, " & Objects(x).Stat1
            If Objects(x).ItemName = "Lance of Yaggai Spetum" Then Return "Lance of Yaggai"
            If Objects(x).ItemName = "Lycander's Flank Ceremonial Spike" Then Return "Lycander's, " & Objects(x).Stat4 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Razortine Trident" Then Return "Razortine, " & Objects(x).Stat2
            If Objects(x).ItemName = "Soulfeast Tine War Fork" Then Return "Soulfeast, " & Objects(x).Stat1 & ", " & Objects(x).Stat2
            If Objects(x).ItemName = "Spire of Honor Lance" Then Return "Spire of Honor, " & Objects(x).Stat3
            If Objects(x).ItemName = "Steel Pillar War Pike" Then Return "Steel Pillar, " & Objects(x).Stat3 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Stoneraven Matriarchal Spear" Then Return "Stoneraven, " & Objects(x).Stat2 & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "The Dragon Chang Spear" Then Return "The Dragon Chang"
            If Objects(x).ItemName = "The Tannr Gorerod Pike" Then Return "Tannr Gorerod, " & Objects(x).Stat1
            If Objects(x).ItemName = "Viperfork Mancatcher" Then Return "Viperfork, " & Objects(x).Stat3 & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat6
        End If

        '***********************************************
        'Weapons Staves
        '***********************************************
        If Objects(x).ItemBase = "Staff" Then
            If Objects(x).ItemName = "Bone Ash" Then Return "Bone Ash"
            If Objects(x).ItemName = "Chromatic Ire" Then Return "Chromatic Ire, " & Objects(x).Stat5 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Mang Song's Lesson Archon Staff" Then Return "Mang Song's Lesson, " & Objects(x).Stat3 & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Ondal's Wisdom Elder Staff" Then Return "Ondal's Wisdom, " & Objects(x).Stat1 & ", " & Objects(x).Stat3 & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Razorswitch Io Staff" Then Return "Razorswitch"
            If Objects(x).ItemName = "Ribcracker Quarterstaff" Then Return "Ribcracker, " & Objects(x).Stat3
            If Objects(x).ItemName = "Serpent Lord Long Staff" Then Return "Serpent Lord, " & Objects(x).Stat1
            If Objects(x).ItemName = "Skull Collector Rune Staff" Then Return "Skull Collector"
            If Objects(x).ItemName = "Spire of Lazarus Gnarled Staff" Then Return "Spire of Lazarus"
            If Objects(x).ItemName = "The Iron Jang Bong War Staff" Then Return "The Iron Jang Bong"
            If Objects(x).ItemName = "The Salamander Battle Staff" Then Return "The Salamander"
            If Objects(x).ItemName = "Warpspear Gothic Staff" Then Return "Warpspear"

        End If

        '***********************************************
        'Weapons Swords
        '***********************************************
        If Objects(x).ItemBase = "Sword" Then
            If Objects(x).ItemName = "Azurewrath Phase Blade" Then Return "Azurewrath, " & Objects(x).Stat4 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Blacktongue Bastard Sword" Then Return "Blacktongue, " & Objects(x).Stat1
            If Objects(x).ItemName = "Blade of Ali Baba Tulwar" Then Return "Blade of Ali Baba, " & Objects(x).Stat1 & ", " & Objects(x).Stat2
            If Objects(x).ItemName = "Blade of Ali Baba Hydra Edge" Then Return "Blade of Ali Baba, " & Objects(x).Stat1 & ", " & Objects(x).Stat2
            If Objects(x).ItemName = "Blood Crescent Scimitar" Then Return "Blood Crescent, " & Objects(x).Stat2
            If Objects(x).ItemName = "Bloodletter Gladius" Then Return "Bloodletter, " & Objects(x).Stat6 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Coldsteel Eye Cutlass" Then Return "Coldsteel Eye, " & Objects(x).Stat2
            If Objects(x).ItemName = "Crainte Vomir Espondon" Then Return "Crainte Vomir, " & Objects(x).Stat3
            If Objects(x).ItemName = "Culwen's Point War Sword" Then Return "Culwen's Point, " & Objects(x).Stat4
            If Objects(x).ItemName = "Djinn Slayer Ataghan" Then Return "Djinn Slayer, " & Objects(x).Stat1 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Doombringer Champion Sword" Then Return "Doombringer, " & Objects(x).Stat3 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Flamebellow Balrog Blade" Then Return "Flamebellow, " & Objects(x).Stat3 & ", " & Objects(x).Stat6 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "Frostwind Cryptic Sword" Then Return "Frostwind, " & Objects(x).Stat2 & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Ginter's Rift Dimensional Blade" Then Return "Ginter's Rift, " & Objects(x).Stat2 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Gleamscythe Falchion" Then Return "Gleamscythe, " & Objects(x).Stat2
            If Objects(x).ItemName = "Griswold's Edge Broad Sword" Then Return "Griswold's Edge, " & Objects(x).Stat2 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Headstriker Battle Sword" Then Return "Headstriker"
            If Objects(x).ItemName = "Hellplague Long Sword" Then Return "Hellplague, " & Objects(x).Stat2
            If Objects(x).ItemName = "Hexfire Shamshir" Then Return "Hexfire, " & Objects(x).Stat2
            If Objects(x).ItemName = "Kinemil's Awl Giant Sword" Then Return "Kinemil's Awl, " & Objects(x).Stat1 & ", " & Objects(x).Stat2
            If Objects(x).ItemName = "Lightsabre Phase Blade" Then Return "Lightsabre, " & Objects(x).Stat3 & ", " & Objects(x).Stat8
            If Objects(x).ItemName = "Plague Bearer Rune Sword" Then Return "Plague Bearer"
            If Objects(x).ItemName = "Ripsaw Flamberge" Then Return "Ripsaw, " & Objects(x).Stat1
            If Objects(x).ItemName = "Rixot's Keen" Then Return "Rixot's Keen" 'unfinnished
            If Objects(x).ItemName = "Shadowfang Two-Handed Sword" Then Return "Shadowfang"
            If Objects(x).ItemName = "Skewer of Krintiz Sabre" Then Return "Skewer of Krintiz"
            If Objects(x).ItemName = "Soulflay Claymore" Then Return "Soulflay, " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Swordguard Executioner Sword" Then Return "Swordguard, " & Objects(x).Stat3 & ", " & Objects(x).Stat7
            If Objects(x).ItemName = "The Atlantean Ancient Sword" Then Return "The Atlantean, " & Objects(x).Stat2
            If Objects(x).ItemName = "The Grandfather Colossus Blade" Then Return "The Grandfather, " & Objects(x).Stat2
            If Objects(x).ItemName = "The Patriarch Great Sword" Then Return "The Patriarch, " & Objects(x).Stat1
            If Objects(x).ItemName = "The Vile Husk Tusk Sword" Then Return "The Vile Husk, " & Objects(x).Stat2
            If Objects(x).ItemName = "Todesfaelle Flamme Zweihander" Then Return "Todesfaelle Flamme, " & Objects(x).Stat2
        End If

        '***********************************************
        'Weapons Wands
        '***********************************************
        If Objects(x).ItemBase = "Wand" Then
            If Objects(x).ItemName = "Arm of King Leoric Tomb Wand" Then Return "Arm of King Leoric"
            If Objects(x).ItemName = "Blackhand Key Grave Wand" Then Return "Blackhand Key"
            If Objects(x).ItemName = "Boneshade Lich Wand" Then Return "Boneshade, " & Stats_items(x)
            If Objects(x).ItemName = "Carin Shard Petrified Wand" Then Return "Carin Shard"
            If Objects(x).ItemName = "Death's Web Unearthed Wand" Then Return "Death's Web, " & Objects(x).Stat2 & ", " & Objects(x).Stat3
            If Objects(x).ItemName = "Gravespine Bone Wand" Then Return "Gravespine, " & Objects(x).Stat6
            If Objects(x).ItemName = "Maelstrom Yew Wand" Then Return "Maelstrom, " & Objects(x).Stat3 & ", " & Objects(x).Stat4 & ", " & Objects(x).Stat5 & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Suicide Branch Burnt Wand" Then Return "Suicide Branch"
            If Objects(x).ItemName = "Torch of Iro Wand" Then Return "Torch of Iro"
            If Objects(x).ItemName = "Ume's Lament Grim Wand" Then Return "Ume's Lament"
        End If

        Return Objects(x).ItemName & Stats_items(x) & " [Not found]"
    End Function


    Function Set_items(ByVal x)
        Dim temp As String = ""
        'MsgBox("Set Items")
        '***********************************************
        'Amulets
        '***********************************************
        If Objects(x).ItemBase = "Amulet" Then
            If Objects(x).ItemName = "Tal Rasha's Adjudication Amulet" Then Return "Tal's  Amulet"
            Return Objects(x).ItemName
        End If

        '***********************************************
        'Rings
        '***********************************************
        If Objects(x).ItemBase = "Ring" Then
            Return Objects(x).ItemName
        End If

        '***********************************************
        ' Set Helms
        '***********************************************
        If Objects(x).ItemBase = "Helm" Or Objects(x).ItemBase = "Circlet" Or Objects(x).ItemBase = "Primal Helm" Then
            If Objects(x).ItemName = "Aldur's Stony Gaze Hunter's Guise" Then Return "Aldur's Guise, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Arcanna's Head Skull Cap" Then Return "Arcanna's Cap, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Berserker's Headgear Helm" Then Return "Berserker's Helm, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Cathan's Visage Mask" Then Return "Cathan's Mask, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Griswold's Valor Corona" Then Return "Griswold's Corona, Def " & Objects(x).Defense & " " & Objects(x).Stat1 & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Guillaume's Face Winged Helm" Then Return "Guillaume's Helm, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hwanin's Splendor Grand Crown" Then Return "Hwanin's Crown, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Immortal King's Will Avenger Guard" Then Return "IK Helm, Def " & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Infernal Cranium Cap" Then Return "Infernal Cap, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Iratha's Coil Crown" Then Return "Iratha's Crown, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Isenhart's Horns Full Helm" Then Return "Isenhart's Helm, Def " & Objects(x).Defense
            If Objects(x).ItemName = "M'avina's True Sight Diadem" Then Return "M'avina's Diadem, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Milabrega's Diadem Crown" Then Return "Milabrega's Crown, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Naj's Circlet Circlet" Then Return "Naj's Circlet, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Natalya's Totem Grim Helm" Then Return "Natalya's Helm, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Ondal's Almighty Spired Helm" Then Return "Ondal's Helm, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sander's Paragon Cap" Then Return "Sander's Cap, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sazabi's Mental Sheath Basinet" Then Return "Sazabi's Basinet, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sigon's Visor Great Helm" Then Return "Sigon's Helm, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Tal Rasha's Horadric Crest Death Mask" Then Return "Tals Mask, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Tancred's Skull Bone Helm" Then Return "Tancred's Helm, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Trang-Ouls' Guise Bone Visage" Then Return "Trang-Ouls Visage, Def " & Objects(x).Defense
        End If

        '***********************************************
        ' Set Armor
        '***********************************************
        If Objects(x).ItemBase = "Armor" Then
            If Objects(x).ItemName = "Aldur's Deception Shadow Plate" Then Return "Aldur's, Armor Def " & Objects(x).Defense & ", " & Objects(x).Stat6
            If Objects(x).ItemName = "Angelic Mantle Ring Mail" Then Return "Angelic Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Arcanna's Flesh Light Plate" Then Return "Arcanna's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Arctic Furs Quilted Armor" Then Return "Arctic Furs Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Berserker's Hauberk Splint Mail" Then Return "Beserker's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Cathan's Mesh Chain Mail" Then Return "Cathan's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Dark Adherent Dusk Shroud" Then Return "Dark Adherent Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Griswold's Heart Ornate Plate" Then Return "Griswold's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Haemosu's Adamant Cuirass" Then Return "Haemosu's Cuirass, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hwanin's Refuge Tigulated Mail" Then Return "Hwanin's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Immortal King's Soul Cage Sacred Armor" Then Return "IK Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Isenhart's Case Breast Plate" Then Return "Isenhart's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "M'avina's Embrace Kraken Shell" Then Return "M'avina's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Milabrega's Robe Ancient Armor" Then Return "Milabrega's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Naj's Light Plate Hellforge Plate" Then Return "Naj's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Natalya's Shadow Loricated Mail" Then Return "Natalya's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sazabi's Ghost Liberator Balrog Skin" Then Return "Sazabi's Armor, Def " & Objects(x).Defense & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Sigon's Shelter Gothic Plate" Then Return "Sigon's Armor, Def " & Objects(x).Defense & " " & Objects(x).Stat5
            If Objects(x).ItemName = "Tal Rasha's Guardianship Lacquered Plate" Then Return "Tal's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Tancred's Spine Full Plate Mail" Then Return "Tancred's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Trang-Oul's Scales Chaos Armor" Then Return "Trang-Oul's Armor, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Vidala's Ambush Leather Armor" Then Return "Vidala's, Def " & Objects(x).Defense
        End If

        '***********************************************
        ' Set Gloves
        '***********************************************
        If Objects(x).ItemBase = "Shield" Or Objects(x).ItemBase = "Voodoo Heads" Then
            If Objects(x).ItemName = "Civerb's Ward Large Shield" Then Return "Civerb's Shield, Def " & Objects(x).Defense '******
            If Objects(x).ItemName = "Cleglaw's Claw Small Shield" Then Return "Cleglaw's Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Griswold's Honor Vortex Shield" Then Return "Griswold's Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hsarus' Iron Fist Buckler" Then Return "Hsarus' Shield, Def " & Objects(x).Defense '******
            If Objects(x).ItemName = "Isenhart's Parry Gothic Shield" Then Return "Isenhart's Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Milabrega's Orb Kite Shield" Then Return "Milabrega's Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sigon's Guard Tower Shield" Then Return "Sigon's Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Taebaek's Glory Ward" Then Return "Taebaek's Shield, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Trang-Oul's Wing Cantor Trophy" Then Return "Trang-Oul's Trophy, Def " & Objects(x).Defense & " " & Objects(x).Stat5
            If Objects(x).ItemName = "Whitstan's Guard Round Shield" Then Return "Whitstan's Shield, Def " & Objects(x).Defense
        End If

        '***********************************************
        ' Set Gloves
        '***********************************************
        If Objects(x).ItemBase = "Gloves" Then
            If Objects(x).ItemName = "Arctic Mitts Light Gauntlets" Then Return "Arctic Mitts, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Cleglaw's Pincers Chain Gloves" Then Return "Cleglaw's Gloves, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Death's Hand Leather Gloves" Then Return "Death's Hand  Gloves, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Immortal King's Forge War Gauntlets" Then Return "IK Gauntlets, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Iratha's Cuff Light Gauntlets" Then Return "Iratha's Gauntlets, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Laying of Hands Bramble Mitts" Then Return "LoH Mitts, Def " & Objects(x).Defense
            If Objects(x).ItemName = "M'avina's Icy Clutch Battle Gauntlets" Then Return "M'avina's Gauntlets, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Magnus' Skin Sharkskin Gloves" Then Return "Magnus Gloves, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sander's Taboo Heavy Gloves" Then Return "Sander's Gloves, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sigon's Gage Gauntlets" Then Return "Sigon's Gloves, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Trang-Oul's Claws Heavy Bracers" Then Return "Trang-Oul's Bracers, Def " & Objects(x).Defense
        End If

        '***********************************************
        ' Set Belts
        '***********************************************
        If Objects(x).ItemBase = "Belt" Then
            If Objects(x).ItemName = "Arctic Binding Light Belt" Then Return "Arctic Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Credendum Mithril Coil" Then Return "Credendum Mithril Coil, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Death's Guard Sash" Then Return "Death's Guard Sash, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hsarus' Iron Stay Belt" Then Return "Hsarus' Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Hwanin's Blessing Belt" Then Return "Hwanin's Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Immortal King's Detail War Belt" Then Return "IK Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Infernal Sign Heavy Belt" Then Return "Infernal Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Iratha's Cord Heavy Belt" Then Return "Iratha's Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "M'avina's Tenet Sharkskin Belt" Then Return "M'avina's Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sigon's Wrap Plated Belt" Then Return "Sigon's Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Tal Rasha's Fine-Spun Cloth Mesh Belt" Then Return "Tal's Belt, Def" & Objects(x).Defense & ", " & Objects(x).Stat4
            If Objects(x).ItemName = "Trang-Oul's Girth Troll Belt" Then Return "Trang-Oul's Belt, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Wilhelm's Pride Battle Belt" Then Return "Wilhelm's Belt, Def " & Objects(x).Defense
        End If

        '***********************************************
        ' Set Boots
        '***********************************************
        If Objects(x).ItemBase = "Boots" Then
            If Objects(x).ItemName = "Aldur's Advance Battle Boots" Then Return "Aldur's Boots, Def " & Objects(x).Defense & ", " & Objects(x).Stat5
            If Objects(x).ItemName = "Hsarus' Iron Heel Chain Boots" Then Return "Hsarus' Boots, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Immortal King's Pillar War Boots" Then Return "IK Boots, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Natalya's Soul Mesh Boots" Then Return "Natalya's Boots, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Rite of Passage Demonhide Boots" Then Return "Rite Of Passage Boots, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sander's Riprap Heavy Boots" Then Return "Sander's Boots, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Sigon's Sabot Greaves" Then Return "Sigon's Greaves, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Tancred's Hobnails Boots" Then Return "Tancred's Boots, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Vidala's Fetlock Light Plated Boots" Then Return "Vidala's Boots, Def " & Objects(x).Defense
        End If

        If Objects(x).ItemBase = "Axe" Then
            If Objects(x).ItemName = "Berserker's Hatchet Double Axe" Then Return "Berserker's Axe"
            If Objects(x).ItemName = "Tancred's Crowbill Military Pick" Then Return "Tancred's Pick"
        End If

        If Objects(x).ItemBase = "Bow" Or Objects(x).ItemBase = "Amazon Bow" Then
            If Objects(x).ItemName = "Arctic Horn Short War Bow" Then Return "Arctic Bow"
            If Objects(x).ItemName = "M'avina's Caster Grand Matron Bow" Then Return "M'avina's Bow"
            If Objects(x).ItemName = "Vidala's Barb Long Battle Bow" Then Return "Vidala's Bow"
        End If


        If Objects(x).ItemBase = "Claw" Then
            If Objects(x).ItemName = "Natalya's Mark Scissors Suwayyah" Then Return "Natalya's Mark"
        End If

        If Objects(x).ItemBase = "Mace" Or Objects(x).ItemBase = "Hammer" Then
            If Objects(x).ItemName = "Aldur's Rhythm Jagged Star" Then Return "Aldur's Star"
            If Objects(x).ItemName = "Dangoon's Teaching Reinforced Mace" Then Return "Dangoon's Mace"
            If Objects(x).ItemName = "Immortal King's Stone Crusher Ogre Maul" Then Return "IK Maul, " & Objects(x).Stat7
        End If

        If Objects(x).ItemBase = "Orb" Then
            If Objects(x).ItemName = "Tal Rasha's Lidless Eye Swirling Crystal" Then Return "Tal's Lidless"
        End If

        If Objects(x).ItemBase = "polearm" Then
            If Objects(x).ItemName = "Hwanin's Justice Bill" Then Return "Hwanin's Justice"
        End If

        If Objects(x).ItemBase = "Staff" Then
            If Objects(x).ItemName = "Arcanna's Deathwand War Staff" Then Return "Arcanna's Staff"
            If Objects(x).ItemName = "Cathan's Rule Battle Staff" Then Return "Cathan's Staff"
            If Objects(x).ItemName = "Naj's Puzzler Elder Staff" Then Return "Naj's Staff"
        End If

        If Objects(x).ItemBase = "Sword" Then
            If Objects(x).ItemName = "Angelic Sickle Sabre" Then Return "Angelic Sabre"
            If Objects(x).ItemName = "Bul-Kathos' Sacred Charge Colossus Blade" Then Return "BK Blade"
            If Objects(x).ItemName = "Bul-Kathos' Tribal Guardian Mythical Sword" Then Return "Bk Sword"
            If Objects(x).ItemName = "Cleglaw's Tooth Long Sword" Then Return "Cleglaw's Sword"
            If Objects(x).ItemName = "Death's Touch War Sword" Then Return "Death's Touch Sword"
            If Objects(x).ItemName = "Isenhart's Lightbrand Broad Sword" Then Return "Isenhart's Sword"
            If Objects(x).ItemName = "Sazabi's Cobalt Redeemer Cryptic Sword" Then Return "Sazabi's Sword"
        End If

        If Objects(x).ItemBase = "Scepter" Then
            If Objects(x).ItemName = "Civerb's Cudgel Grand Scepter" Then Return "Civerb's Scepter"
            If Objects(x).ItemName = "Milabrega's Rod War Scepter" Then Return "Milabrega's Scepter"
        End If

        If Objects(x).ItemName = "Infernal Torch Grim Wand" Then Return "Infernal Wand"
        If Objects(x).ItemName = "Sander's Superstition Bone Wand" Then Return "Sander's Wand"


        Return Objects(x).ItemName & Stats_items(x) & " [Not found]"

    End Function

End Module
