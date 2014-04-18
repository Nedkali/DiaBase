Imports System.Text.RegularExpressions

Module Tradelist


    Public Sub SendToTradeList(ByVal x)
        Dim temp As String = ""

        If Objects(x).ItemBase = "Rune" Then
            Dim rune = Objects(x).ItemName.Replace(" Rune", "")
            Form1.RichTextBox3.AppendText(rune & vbCrLf)
            Return
        End If

        temp = Objects(x).ItemName & ", " 'sets the default name


        If Objects(x).ItemName = "The Stone of Jordan Ring" Then Form1.RichTextBox3.AppendText("Soj" & vbCrLf) : Return
        If Objects(x).ItemName = "Mara's Kaleidoscope Amulet" Then
            temp = "Mara's " & Objects(x).Stat3
            temp = temp.Replace("All Resistances", "") : temp = temp.Replace("+", "")
            Form1.RichTextBox3.AppendText(temp & vbCrLf) : Return
        End If
        If Objects(x).ItemName = "Token of Absolution" Then Form1.RichTextBox3.AppendText("Token" & vbCrLf) : Return



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


        If Objects(x).Stat1 <> Nothing Then temp = temp & ", " & Objects(x).Stat1
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

        temp = temp.Replace("Lightning Resist", "Lr")
        temp = temp.Replace("Cold Resist", "Cr")
        temp = temp.Replace("Fire Resist", "Fr")
        temp = temp.Replace("Poison Resist", "Pr")
        temp = temp.Replace("to Dexterity", "to Dex")
        temp = temp.Replace("Better Chance of Getting Magic Items", "Mf")
        temp = temp.Replace("Ethereal (Cannot be Repaired)", "eth")
        temp = temp.Replace("Extra Gold from Monsters", "Gf")
        temp = temp.Replace("All Resistances", "Res All")
        temp = temp.Replace("to Experience Gained", "XP")
        temp = temp.Replace("Faster Cast Rate", "fcr")
        temp = temp.Replace("Enhanced Defense", "Edef")
        temp = temp.Replace("Faster Hit Recovery", "Fhr")
        temp = temp.Replace("Faster Run/Walk", "Frw")
        temp = temp.Replace("Attack Rating", "Ar")
        temp = temp.Replace("Faster Block Rate", "Fbr")
        temp = temp.Replace("Increased Attack Speed", "Ias")


        temp = temp.Replace("Regenerate Mana", "Mana Regen")

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



        Form1.RichTextBox3.AppendText(temp & vbCrLf)
    End Sub


End Module
