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


        ' if Identified & set item go to set function
        If Objects(x).ItemQuality = "Set" And Objects(x).Stat1.IndexOf("Unid") = -1 Then
            temp = Set_items(x)
            If Objects(x).Sockets <> "" Then temp = temp & ", Socs " & Objects(x).Sockets
            If Objects(x).EtherealItem = True Then temp = temp & ", Eth"
        End If

        ' if Identified & set item go to Unique function
        If Objects(x).ItemQuality = "Unique" And Objects(x).Stat1.IndexOf("Unid") = -1 Then temp = Uniq_items(x) : If Objects(x).EtherealItem = True Then temp = temp & " Eth"

        If Objects(x).RuneWord = "True" Then temp = RuneWord_items(x) : If Objects(x).EtherealItem = True Then temp = temp & " Eth"


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
            temp = Objects(x).ItemName & " Def " & Objects(x).Defense
            'If Objects(x).ItemName.IndexOf("Enigma") > -1  Then
            'If Objects(x).ItemName.IndexOf("Treachery") > -1 Then
            'If Objects(x).ItemName.IndexOf("Fortitude") > -1 Then

            Return temp
        End If

        '***********************************************
        ' Runeword Shields
        '***********************************************
        If Objects(x).ItemBase = "Shield" Then
            If Objects(x).ItemName.IndexOf("Spirit") > -1 Then temp = Objects(x).ItemName & " Def " & Objects(x).Defense & Objects(x).Stat2 : Return temp

            temp = Objects(x).ItemName & " Def " & Objects(x).Defense
            Return temp
        End If

        '***********************************************
        ' Runeword Weapons
        '***********************************************
        If Objects(x).ItemBase = "polearm" Or Objects(x).ItemBase = "Sword" Or Objects(x).ItemBase = "Mace" Or Objects(x).ItemBase = "Axe" Then
            If Objects(x).ItemName.IndexOf("Insight") > -1 Then temp = Objects(x).ItemName & " " & Objects(x).Stat2
            If Objects(x).ItemName.IndexOf("Spirit") > -1 Then temp = Objects(x).ItemName & " " & Objects(x).Stat3
            If Objects(x).ItemName.IndexOf("Oak") > -1 Then temp = Objects(x).ItemName & " " & Objects(x).Stat12
            If Objects(x).ItemName.IndexOf("Arms") > -1 Then temp = Objects(x).ItemName & " " & Objects(x).Stat6 & " " & Objects(x).Stat7 & " " & Objects(x).Stat8
            If Objects(x).ItemName.IndexOf("Grief") > -1 Then temp = Objects(x).ItemName & " " & Objects(x).Stat2
            If Objects(x).ItemName.IndexOf("Infinity") > -1 Then temp = Objects(x).ItemName & " " & Objects(x).Stat6
            If Objects(x).ItemName.IndexOf("Last Wish") > -1 Then temp = Objects(x).ItemName & " " & Objects(x).Stat8
            Return temp
        End If

        Return temp
    End Function


    Function Uniq_items(ByRef x)
        Dim temp As String = ""
        '***********************************************
        'Unique items
        If Objects(x).ItemName = "Rainbow Facet Jewel" Then Return "Rainbow Facet  " & Objects(x).Stat3 & " " & Objects(x).Stat4
        '***********************************************
        'Rings
        '***********************************************
        If Objects(x).ItemName = "Bul-Kathos' Wedding Band Ring" Then Return "BK Ring " & Objects(x).Stat2
        If Objects(x).ItemName = "Nagelring Ring" Then Return "Nagel Ring, " & Objects(x).Stat4
        If Objects(x).ItemName = "The Stone of Jordan Ring" Then Return "Soj"
        If Objects(x).ItemName = "Carrion Wind Ring" Then Return "Nagel Ring, " & Objects(x).Stat3
        If Objects(x).ItemName = "Dwarf Star Ring" Then Return Objects(x).ItemName
        If Objects(x).ItemName = "Manald Heal Ring" Then Return "Manald Heal Ring, "
        If Objects(x).ItemName = "Raven Frost Ring" Then Return "Raven Frost, " & Objects(x).Stat1 & " " & Objects(x).Stat3
        If Objects(x).ItemName = "Wisp Projector Ring" Then Return "Wisp Ring, " & Objects(x).Stat2 & " " & Objects(x).Stat3
        If Objects(x).ItemName = "Nature's Peace Ring" Then Return "Nature's Peace Ring, " & Objects(x).Stat3

        '***********************************************
        'Amulets
        '***********************************************
        If Objects(x).ItemName = "Atma's Scarab Amulet" Then Return "Atma's Amulet, "
        If Objects(x).ItemName = "Cresent Moon Amulet" Then Return "Cresent Amulet, "
        If Objects(x).ItemName = "Highlord's Wrath Amulet" Then Return "Highlord's Amulet, "
        If Objects(x).ItemName = "Mara's Kaleidoscope Amulet" Then Return "Mara's, " & Objects(x).Stat3
        If Objects(x).ItemName = "Metal Grid Amulet" Then Return "Metal Grid Amulet, "
        If Objects(x).ItemName = "Nokozan Relic" Then Return "Nokozan Relic, "
        If Objects(x).ItemName = "Saracen's Chance Amulet" Then Return "Saracen's Amulet, "
        If Objects(x).ItemName = "Seraph's Hymn Amulet" Then Return "Seraph's Hymn Amulet, "
        If Objects(x).ItemName = "The Cat's Eye Amulet" Then Return "Cat's Eye Amulet, "
        If Objects(x).ItemName = "The Eye of Etlich Amulet" Then Return "Etlich Amulet, " & Objects(x).Stat3
        If Objects(x).ItemName = "The Mahim-Oak Curio Amulet" Then Return "Mahim-Oak Amulet, "
        If Objects(x).ItemName = "The Rising Sun Amulet" Then Return "Rising Sun Amulet, "

        '***********************************************
        ' Unique Helms
        '***********************************************

        If Objects(x).ItemBase = "Helm" Then
            If Objects(x).ItemName = "Andariel's Visage Demonhead" Then Return "Andies, " & Objects(x).Stat4 & " " & Objects(x).Stat6
            If Objects(x).ItemName = "Harlequin Crest Shako" Then Return "Shako, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Veil of Steel Spired Helm" Then Return "Veil of Steel, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Peasant Crown War Hat" Then Return "Peasant Hat , Def " & Objects(x).Defense
            If Objects(x).ItemName = "Kira's Guardian Tiara" Then Return "Kira's , Def " & Objects(x).Defense & " " & Objects(x).Stat3

            If Objects(x).ItemName = "Crown of Ages Corona" Then
                temp = "CoA, Def" & Objects(x).Defense
                If Objects(x).Stat1 <> "Indestructible" Then temp = temp & Objects(x).Stat1
                Return temp & " Soc" & Objects(x).Sockets
            End If

            Return temp
        End If
        '***********************************************
        'Belts
        '***********************************************
        If Objects(x).ItemBase = "Belt" Then
            If Objects(x).ItemName = "Arachnid Mesh Spiderweb Sash" Then Return "Arach, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Thundergod's Vigor War Belt" Then Return "Thundergod's, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Verdungo's Hearty Cord Mithril Coil" Then Return "Dungo's, Def " & Objects(x).Defense & " " & Objects(x).Stat3
            If Objects(x).ItemName = "String of Ears Demonhide Sash" Then Return "String of Ears, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Goldwrap Heavy Belt" Then Return "Goldwrap, " & Stats_items(x)

            Return temp
        End If

        '***********************************************
        'Gloves
        '***********************************************
        If Objects(x).ItemBase = "Gloves" Then
            If Objects(x).ItemName = "Chance Guards Chain Gloves" Then Return "Chancies, " & Objects(x).Stat5
            If Objects(x).ItemName.IndexOf("Magefist") > -1 Then Return "Magefist, Def " & Objects(x).Defense & " " & Objects(x).Stat4
            If Objects(x).ItemName = "Dracul's Grasp Vampirebone Gloves" Then Return "Dracs, " & Objects(x).Stat2 & " " & Objects(x).Stat5
            If Objects(x).ItemName = "Venom Grip Demonhide Gloves" Then Return "Venom Grip, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lava Gout Crusader Gauntlets" Then Return "Lava Gout Crusader Gauntlets, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lava Gout Battle Gauntlets" Then Return "Lava Gout Battle Gauntlets, Def " & Objects(x).Defense
            Return temp
        End If

        '***********************************************Lava Gout Battle Gauntlets
        'Boots
        '***********************************************
        If Objects(x).ItemBase = "Boots" Then
            If Objects(x).ItemName = "War Traveler Battle Boots" Then temp = "War Travs, Def " & Objects(x).Defense & " " & Objects(x).Stat8
            If Objects(x).ItemName = "Waterwalk Sharkskin Boots" Then temp = "Waterwalks, Def " & Objects(x).Defense & " " & Objects(x).Stat5
            If Objects(x).ItemName = "Sandstorm Trek Scarabshell Boots" Then temp = "Treks, Def " & Objects(x).Defense
            If Objects(x).ItemName = "Marrowwalk Boneweave Boots" Then temp = "Marrowwalk, Def " & Objects(x).Defense & " " & Objects(x).Stat2
            If Objects(x).ItemName = "Gore Rider War Boots" Then temp = "Gore Rider's, Def " & Objects(x).Defense
            Return temp
        End If

        '***********************************************
        ' Unique Armor
        '***********************************************
        If Objects(x).ItemBase = "Armor" Then
            If Objects(x).ItemName = "Skin of the Vipermagi Serpentskin Armor" Then temp = "Vipermagi  " & Objects(x).Stat4
            If Objects(x).ItemName = "Shaftstop Mesh Armor" Then temp = "Shaftstop Armor Def " & Objects(x).Defense

            Return temp
        End If

        '***********************************************
        'Shields
        '***********************************************
        If Objects(x).ItemBase = "Shield" Or Objects(x).ItemBase = "Auric Shield" Then
            If Objects(x).ItemName = "Herald of Zakarum Gilded Shield" Then Return "Hoz " & " Def " & Objects(x).Defense
            If Objects(x).ItemName = "Steelclash Kite Shield" Then Return "Steelclash " & " Def " & Objects(x).Defense
            If Objects(x).ItemName = "Lidless Wall Grim Shield" Then Return "Lidless Shield" & " Def " & Objects(x).Defense

            Return temp
        End If

        '***********************************************
        'Weapons orbs
        '***********************************************
        If Objects(x).ItemName = "Eschuta's Temper Eldritch Orb" Then Return "Eschuta's" & Stats_items(x)
        If Objects(x).ItemName = "Tal Rasha's Lidless Eye Swirling Crystal" Then Return "Tal's Lidless, " & Objects(x).Stat3 & Objects(x).Stat4 & Objects(x).Stat5


        Return temp
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
            Return Objects(x).ItemName & " Not Listed"
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
            Return Objects(x).ItemName & " Not Listed"
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
            If Objects(x).ItemName = "Whitstan's Guard Round Shield" Then Return "Whitstan's Guard Round Shield, Def " & Objects(x).Defense
            Return Objects(x).ItemName & " Not Listed"
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
            Return Objects(x).ItemName & " Not Listed"
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
            Return Objects(x).ItemName & " Not Listed"
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
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemBase = "Axe" Then
            If Objects(x).ItemName = "Berserker's Hatchet Double Axe" Then Return "Berserker's Axe"
            If Objects(x).ItemName = "Tancred's Crowbill Military Pick" Then Return "Tancred's Pick"
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemBase = "Bow" Then
            If Objects(x).ItemName = "Arctic Horn Short War Bow" Then Return "Arctic Bow"
            If Objects(x).ItemName = "M'avina's Caster Grand Matron Bow" Then Return "M'avina's Bow"
            If Objects(x).ItemName = "Vidala's Barb Long Battle Bow" Then Return "Vidala's Bow"
            Return Objects(x).ItemName & " Not Listed"
        End If


        If Objects(x).ItemBase = "Claw" Then
            If Objects(x).ItemName = "Natalya's Mark Scissors Suwayyah" Then Return "Natalya's Mark"
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemBase = "Mace" Or Objects(x).ItemBase = "Hammer" Then
            If Objects(x).ItemName = "Aldur's Rhythm Jagged Star" Then Return "Aldur's Star"
            If Objects(x).ItemName = "Dangoon's Teaching Reinforced Mace" Then Return "Dangoon's Mace"
            If Objects(x).ItemName = "Immortal King's Stone Crusher Ogre Maul" Then Return "IK Maul, " & Objects(x).Stat7
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemBase = "Orb" Then
            If Objects(x).ItemName = "Tal Rasha's Lidless Eye Swirling Crystal" Then Return "Tal's Lidless"
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemBase = "polearm" Then
            If Objects(x).ItemName = "Hwanin's Justice Bill" Then Return "Hwanin's Justice"
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemBase = "Staff" Then
            If Objects(x).ItemName = "Arcanna's Deathwand War Staff" Then Return "Arcanna's Staff"
            If Objects(x).ItemName = "Cathan's Rule Battle Staff" Then Return "Cathan's Staff"
            If Objects(x).ItemName = "Naj's Puzzler Elder Staff" Then Return "Naj's Staff"
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemBase = "Sword" Then
            If Objects(x).ItemName = "Angelic Sickle Sabre" Then Return "Angelic Sabre"
            If Objects(x).ItemName = "Bul-Kathos' Sacred Charge Colossus Blade" Then Return "BK Blade"
            If Objects(x).ItemName = "Bul-Kathos' Tribal Guardian Mythical Sword" Then Return "Bk Sword"
            If Objects(x).ItemName = "Cleglaw's Tooth Long Sword" Then Return "Cleglaw's Sword"
            If Objects(x).ItemName = "Death's Touch War Sword" Then Return "Death's Touch Sword"
            If Objects(x).ItemName = "Isenhart's Lightbrand Broad Sword" Then Return "Isenhart's Sword"
            If Objects(x).ItemName = "Sazabi's Cobalt Redeemer Cryptic Sword" Then Return "Sazabi's Sword"
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemBase = "Scepter" Then
            If Objects(x).ItemName = "Civerb's Cudgel Grand Scepter" Then Return "Civerb's Scepter"
            If Objects(x).ItemName = "Milabrega's Rod War Scepter" Then Return "Milabrega's Scepter"
            Return Objects(x).ItemName & " Not Listed"
        End If

        If Objects(x).ItemName = "Infernal Torch Grim Wand" Then Return "Infernal Wand"
        If Objects(x).ItemName = "Sander's Superstition Bone Wand" Then Return "Sander's Wand"
        Return Objects(x).ItemName & " Not Listed"

    End Function

End Module
