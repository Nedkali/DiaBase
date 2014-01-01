Imports System.IO



Module AutoLogger


    Public Sub ImportLogFiles()


        'these variables maybe better placed in module1 ????
        MuleLogPath = EtalPath + "\scripts\AMS\MuleInventory\"
        DataBasePath = Application.StartupPath + "\DataBase\"
        MuleDataPath = EtalPath + "\scripts\AMS\MuleLogs\"
        ArchiveFolder = Application.StartupPath + "\Archive\"

        'Check folders exist and path information is correct

        If My.Computer.FileSystem.DirectoryExists(MuleDataPath) = False Then
            MessageBox.Show("Error with Logs Path", "Folder  Error")
            Return
        End If

        If My.Computer.FileSystem.DirectoryExists(DataBasePath) = False Then
            MessageBox.Show("Error with Database Path", "Folder  Error")
            Return
        End If
        If My.Computer.FileSystem.DirectoryExists(MuleLogPath) = False Then
            MessageBox.Show("Error with Etal Path", "Folder  Error")
            Return
        End If

        If My.Computer.FileSystem.DirectoryExists(ArchiveFolder) = False Then
            MessageBox.Show("Error with Backup Path", "Folder  Error")
            Return
        End If

        ' Check Log folder for files to process
        GetLogFiles()
        If LogFilesList.Count = 0 Then
            Form1.RichTextBox1.AppendText("No Logs to import" & vbCrLf)
            Return 'If There Are no Log Files - exit
        End If


        Form1.RichTextBox1.Text = "Logs to import = " & LogFilesList.Count & vbCrLf
        Dim Pretotal = Objects.Count
        GetmuleaccountFiles()

        ProcessLogFiles() 'moved rest to a separate sub to cut down on coded lines in a single sub

        If Objects.Count - Pretotal > 0 Then SaveLoggedItems(Pretotal)
        If Objects.Count - Pretotal = 0 Then MessageBox.Show("logged accounts = " & Objects.Count - Pretotal, "check")

    End Sub
    Sub GetLogFiles()

        Dim Tally As Integer = 0
        LogFilesList.Clear()
        Dim LogFiles As String() = (Directory.GetFiles(MuleLogPath, "*")).Select(Function(p) Path.GetFileName(p)).ToArray() ' gets file and crops path

        Do Until Tally = LogFiles.Count
            If LogFiles(Tally).IndexOf("delete me.txt") = -1 Then                   'ignore the delete me file (if its still there)
                If Replace(LogFiles(Tally), ".txt", "").IndexOf("_") > -1 Then
                    LogFilesList.Add(LogFiles(Tally))                               'ADD LOG FILE NAME TO LogFilesList()
                End If
            End If
            Tally = Tally + 1
        Loop

    End Sub
    'SUB GETS ALL _muleaccount.txt fileas from etals mulelogs dir and puts em in PassFiles()
    Sub GetmuleaccountFiles()
        Dim Tally As Integer = 0
        PassFiles.Clear()

        Dim AllFiles As String() = (Directory.GetFiles(MuleDataPath, "*")).Select(Function(p) Path.GetFileName(p)).ToArray() ' gets file and crops path
        For Each item In AllFiles
            If AllFiles(Tally).IndexOf("_muleaccount.txt") > -1 Then PassFiles.Add(AllFiles(Tally))
            Tally = Tally + 1
        Next

    End Sub

    Private Sub ProcessLogFiles()
        Dim temp As String = ""
        Dim Tally As Integer = 0
        Dim mycounter As Integer = 0
        Dim found As Boolean = True
        Dim myarray As Array
        Dim thislogmulename As String
        Dim thislogmuleacc As String
        Dim thislogpass As String
        Do Until Tally = LogFilesList.Count

            If My.Computer.FileSystem.FileExists(MuleLogPath & LogFilesList(Tally)) = True Then 'Verify the log Exists
                Dim LogFile = My.Computer.FileSystem.OpenTextFileReader(MuleLogPath & LogFilesList(Tally))



                thislogmuleacc = LogFile.ReadLine 'these lines should exist for each log
                thislogmulename = LogFile.ReadLine
                LogFile.ReadLine()                            'Blank line
                thislogpass = GetMulePass(thislogmuleacc)
                If KeepPassPrivate = True Then thislogpass = "***********"

                Do
                    Dim NewObject As New ItemObjects

                    NewObject.MuleAccount = thislogmuleacc
                    NewObject.MuleName = thislogmulename
                    NewObject.MulePass = thislogpass

                    For x = 0 To 4 '     just in case of extra blank lines
                        If LogFile.EndOfStream = True Then Exit Do
                        temp = LogFile.ReadLine()
                        If temp <> "" Then Exit For
                    Next
                    NewObject.ItemName = temp 'these 5 lines should exist for each item
                    temp = LogFile.ReadLine() : myarray = Split(temp, " ")
                    NewObject.ItemType = myarray(2)
                    temp = LogFile.ReadLine() : myarray = Split(temp, " ")
                    NewObject.ItemClass = myarray(2)
                    temp = LogFile.ReadLine() : myarray = Split(temp, " ")
                    NewObject.ItemImage = myarray(2)
                    NewObject.RuneWord = LogFile.ReadLine()

                    While LogFile.EndOfStream = False   'attempt to read item added information and exit if end of stream/file
                        temp = LogFile.ReadLine()

                        If temp = Nothing Then Exit While 'breaks while loop if we find blank line then move onto next item
                        myarray = Split(temp, ": ", 0)  ' gathers most basic stats

                        Select Case myarray(0) + ":"
                            Case "Throw Damage:"
                                Dim temp1 = myarray(myarray.Length - 1)
                                temp1 = Replace(temp1, "to ", "")
                                myarray = Split(temp1, " ", 0)
                                NewObject.ThrowDamageMin = myarray(0)
                                NewObject.ThrowDamageMax = myarray(myarray.Length - 1)

                            Case "One-Hand Damage:"
                                Dim temp1 = myarray(myarray.Length - 1)
                                temp1 = Replace(temp1, "to ", "")
                                myarray = Split(temp1, " ", 0)
                                NewObject.OneHandDamageMin = myarray(0)
                                NewObject.OneHandDamageMax = myarray(myarray.Length - 1)

                            Case "Two-Hand Damage:"
                                Dim temp1 = myarray(myarray.Length - 1)
                                temp1 = Replace(temp1, "to ", "")
                                myarray = Split(temp1, " ", 0)
                                NewObject.TwoHandDamageMin = myarray(0)
                                NewObject.TwoHandDamageMax = myarray(myarray.Length - 1)

                            Case "Quantity:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.QuantityMin = myarray(myarray.Length - 1)

                            Case "Durability:"
                                Dim temp1 = myarray(myarray.Length - 1)
                                temp1 = Replace(temp1, "of ", "")
                                myarray = Split(temp1, " ", 0)
                                NewObject.DurabilityMin = myarray(0)
                                NewObject.DurabilityMax = myarray(myarray.Length - 1)

                            Case "Defense:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.Defense = myarray(myarray.Length - 1)

                            Case "Chance to Block:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.ChanceToBlock = myarray(myarray.Length - 1)

                            Case "Required Strength:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.RequiredStrength = myarray(myarray.Length - 1)

                            Case "Required Level:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.RequiredLevel = myarray(myarray.Length - 1)

                            Case "Required Dexterity:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.RequiredDexterity = myarray(myarray.Length - 1)

                            Case Else
                                found = False


                        End Select

                        ' if not basic item stat then we need to set it to stat1 stat2 etc
                        If found = False Then
                            If temp.IndexOf("Class - ") <> -1 Then
                                myarray = Split(temp, "Class - ", 0)
                                Dim temp1 = myarray(myarray.Length - 1)
                                NewObject.AttackSpeed = temp1
                                NewObject.AttackClass = NewObject.ItemType
                            End If

                            If temp.IndexOf("Socketed") <> -1 Then
                                myarray = Split(temp, "Socketed (", 0)
                                Dim temp1 = myarray(myarray.Length - 1)
                                myarray = Split(temp1, ")", 0)
                                NewObject.Sockets = myarray(0)
                            End If

                            If temp.IndexOf("Ethereal") <> -1 Then
                                NewObject.EtherealItem = True
                            End If

                            If temp = "(Paladin Only)" Then
                                NewObject.RequiredCharacter = "Paladin"
                                found = True
                            End If

                            If temp = "(Sorceress Only)" Then
                                NewObject.RequiredCharacter = "Sorceress"
                                found = True
                            End If

                            If temp = "(Necromancer Only)" Then
                                NewObject.RequiredCharacter = "Necromancer"
                                found = True
                            End If

                            If temp = "(Amazon Only)" Then
                                NewObject.RequiredCharacter = "Amazon"
                                found = True
                            End If

                            If temp = "(Assassin Only)" Then
                                NewObject.RequiredCharacter = "Assassin"
                                found = True
                            End If

                            If temp = "(Druid Only)" Then
                                NewObject.RequiredCharacter = "Druid"
                                found = True
                            End If


                            If found = False And NewObject.Stat1 = "" Then NewObject.Stat1 = temp : found = True
                            If found = False And NewObject.Stat2 = "" Then NewObject.Stat2 = temp : found = True
                            If found = False And NewObject.Stat3 = "" Then NewObject.Stat3 = temp : found = True
                            If found = False And NewObject.Stat4 = "" Then NewObject.Stat4 = temp : found = True
                            If found = False And NewObject.Stat5 = "" Then NewObject.Stat5 = temp : found = True
                            If found = False And NewObject.Stat6 = "" Then NewObject.Stat6 = temp : found = True
                            If found = False And NewObject.Stat7 = "" Then NewObject.Stat7 = temp : found = True
                            If found = False And NewObject.Stat8 = "" Then NewObject.Stat8 = temp : found = True
                            If found = False And NewObject.Stat9 = "" Then NewObject.Stat9 = temp : found = True
                            If found = False And NewObject.Stat10 = "" Then NewObject.Stat10 = temp : found = True
                            If found = False And NewObject.Stat11 = "" Then NewObject.Stat11 = temp : found = True
                            If found = False And NewObject.Stat12 = "" Then NewObject.Stat12 = temp : found = True
                            If found = False And NewObject.Stat13 = "" Then NewObject.Stat13 = temp : found = True
                            If found = False And NewObject.Stat14 = "" Then NewObject.Stat14 = temp : found = True
                            If found = False And NewObject.Stat15 = "" Then NewObject.Stat15 = temp : found = True


                        End If


                    End While
                    If NewObject.ItemName.IndexOf("Rune") > -1 Then
                        temp = GetRunes(NewObject.ItemName)
                        myarray = Split(temp, ",")
                        NewObject.RequiredLevel = myarray(0)
                        NewObject.Stat1 = myarray(1)
                        NewObject.Stat2 = myarray(2)
                        NewObject.Stat3 = myarray(3)
                        NewObject.Stat4 = myarray(4)
                        NewObject.Stat5 = myarray(5)
                        NewObject.Stat6 = myarray(6)
                        NewObject.Stat7 = myarray(7)
                        NewObject.Stat8 = myarray(8)
                        NewObject.Stat9 = myarray(9)
                        NewObject.Stat10 = myarray(10)
                        NewObject.Stat11 = myarray(11)

                    End If
                    If NewObject.ItemType = "Gem" Then
                        temp = GetGemsStats(NewObject.ItemName)
                        myarray = Split(temp, ",")
                        NewObject.RequiredLevel = myarray(0)
                        NewObject.Stat1 = myarray(1)
                        NewObject.Stat2 = myarray(2)
                        NewObject.Stat3 = myarray(3)
                        NewObject.Stat4 = myarray(4)
                    End If

                    If NewObject.ItemName.IndexOf("Token of Absolution") <> -1 Then
                        NewObject.ItemName = "Token of Absolution"
                        NewObject.Stat1 = "Right-click to reset Stat/Skill Points"
                    End If

                    NewObject.PickitBot = "" 'Still to fix these lines maybe just use current time/date
                    Objects.Add(NewObject)

                Loop Until LogFile.EndOfStream



                LogFile.Close()


                Dim filecheck = My.Computer.FileSystem.FileExists(ArchiveFolder & LogFilesList(Tally))
                Dim filecount As Integer = 0
                If filecheck = False Then
                    My.Computer.FileSystem.MoveFile(MuleLogPath & LogFilesList(Tally), ArchiveFolder & LogFilesList(Tally))
                End If

                If filecheck = True Then
                    myarray = Split(LogFilesList(Tally), ".txt", 0)
                    Dim tempname = myarray(0)
                    Do Until filecheck = False
                        temp = tempname & filecount & ".txt"
                        filecheck = My.Computer.FileSystem.FileExists(ArchiveFolder & temp)
                        filecount = filecount + 1
                    Loop
                    My.Computer.FileSystem.MoveFile(MuleLogPath & LogFilesList(Tally), ArchiveFolder & temp)
                End If


            End If

            Tally = Tally + 1
        Loop

    End Sub
    Function GetMulePass(ByVal accname)
        Dim counter As Integer = 0
        Dim temp1 As String = ""
        Dim password As String = ""

        Form1.RichTextBox1.AppendText("Account name searching for = " & accname & vbCrLf)
        Form1.RichTextBox1.AppendText("Search Folder = " & MuleDataPath & PassFiles(counter) & vbCrLf)

        For Each item In PassFiles

            If My.Computer.FileSystem.FileExists(MuleDataPath & PassFiles(counter)) = False Then
                Form1.RichTextBox1.AppendText("File doesnt Exist " & vbCrLf)
                Return password
            End If

            Dim ReadPassFiles = My.Computer.FileSystem.OpenTextFileReader(MuleDataPath & PassFiles(counter))
            Form1.RichTextBox1.AppendText("checking file " & PassFiles(counter) & vbCrLf)
            While ReadPassFiles.EndOfStream = False
                temp1 = ReadPassFiles.ReadLine()
                If temp1.IndexOf(accname) <> -1 Then
                    Dim AccAndPass = temp1.Split(New [Char]() {"/"c})
                    password = AccAndPass(1)     'Get mule acc pass from _muleaccount.txt
                End If
            End While
            ReadPassFiles.Close()

            counter = counter + 1
        Next
        Form1.RichTextBox1.AppendText("Password found = " & password & vbCrLf)
        Form1.RichTextBox1.ScrollToCaret()
        Return password
    End Function

    Private Sub SaveLoggedItems(ByVal itemstart)
        ' MessageBox.Show(DataBaseFile, "Check")
        ' Return
        Try

            Dim LogWriter = My.Computer.FileSystem.OpenTextFileWriter(DataBaseFile, True)

            For x = itemstart To Objects.Count - 1

                LogWriter.WriteLine("--------------------")
                LogWriter.WriteLine(Objects(x).ItemName)
                LogWriter.WriteLine(Objects(x).ItemType)
                LogWriter.WriteLine(Objects(x).ItemClass)
                LogWriter.WriteLine(Objects(x).RequiredCharacter)
                LogWriter.WriteLine(Objects(x).EtherealItem)
                LogWriter.WriteLine(Objects(x).Sockets)
                LogWriter.WriteLine(Objects(x).RuneWord)
                LogWriter.WriteLine(Objects(x).ThrowDamageMin)
                LogWriter.WriteLine(Objects(x).ThrowDamageMax)
                LogWriter.WriteLine(Objects(x).OneHandDamageMin)
                LogWriter.WriteLine(Objects(x).OneHandDamageMax)
                LogWriter.WriteLine(Objects(x).TwoHandDamageMin)
                LogWriter.WriteLine(Objects(x).TwoHandDamageMax)
                LogWriter.WriteLine(Objects(x).Defense)
                LogWriter.WriteLine(Objects(x).ChanceToBlock)
                LogWriter.WriteLine(Objects(x).QuantityMin)
                LogWriter.WriteLine(Objects(x).QuantityMax)
                LogWriter.WriteLine(Objects(x).DurabilityMin)
                LogWriter.WriteLine(Objects(x).DurabilityMax)
                LogWriter.WriteLine(Objects(x).RequiredStrength)
                LogWriter.WriteLine(Objects(x).RequiredDexterity)
                LogWriter.WriteLine(Objects(x).RequiredLevel)
                LogWriter.WriteLine(Objects(x).AttackClass)
                LogWriter.WriteLine(Objects(x).AttackSpeed)
                LogWriter.WriteLine(Objects(x).Stat1)
                LogWriter.WriteLine(Objects(x).Stat2)
                LogWriter.WriteLine(Objects(x).Stat3)
                LogWriter.WriteLine(Objects(x).Stat4)
                LogWriter.WriteLine(Objects(x).Stat5)
                LogWriter.WriteLine(Objects(x).Stat6)
                LogWriter.WriteLine(Objects(x).Stat7)
                LogWriter.WriteLine(Objects(x).Stat8)
                LogWriter.WriteLine(Objects(x).Stat9)
                LogWriter.WriteLine(Objects(x).Stat10)
                LogWriter.WriteLine(Objects(x).Stat11)
                LogWriter.WriteLine(Objects(x).Stat12)
                LogWriter.WriteLine(Objects(x).Stat13)
                LogWriter.WriteLine(Objects(x).Stat14)
                LogWriter.WriteLine(Objects(x).Stat15)
                LogWriter.WriteLine(Objects(x).MuleName)
                LogWriter.WriteLine(Objects(x).MuleAccount)
                LogWriter.WriteLine(Objects(x).MulePass)
                LogWriter.WriteLine(Objects(x).PickitBot)
                LogWriter.WriteLine(Objects(x).UserReference)
                LogWriter.WriteLine(Objects(x).ItemImage)

            Next

            LogWriter.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "File Write Error")

        End Try

        Form1.TextBox2.Text = Objects.Count & " Items"
    End Sub
    Function GetRunes(ByVal runename)
        Dim runestats As String = ""
        Select Case runename

            Case "Eld Rune"
                runestats = "11, Weapons:, +50 To Attack Rate, +175% Damage To Undead, Armour:, +15% Stamina Drain, Helmets:, +15% Stamina Drain, Shields:, Increased Chance Of Blocking, , , , "

            Case "El Rune"
                runestats = "11, Weapons:, +1 To Light Radius, 50 To Attack Rating,  Armour:, +50 To Attack Rate Vs Undead, +175% Damage Vs Undead, Helmets:, +1 To Light Radius, +15 To Defense, Shields:, +1 to Light Radius, +15 To Defense"

            Case "Tir Rune"
                runestats = "13, Weapons:, +2 Mana After Kill, Armour:, +2 Mana After Kill, Helmets:, +2 Mana After Kill, Shields:, +2 Mana After Kill, , , , "

            Case "Nef Rune"
                runestats = "13, Weapons:, Knockback, Armour:, +30 Defense Vs Missiles, Helmets:, +30 Defense Vs Missiles, Shields:, +30 Defense Vs Missiles, , , "

            Case "Eth Rune"
                runestats = "15, Weapons:, -25% To Targets Defense, Armour:, Regenerate Mana 15%, Helmets:, Regenerate Mana 15%, Shields:, Regenerate Mana 15%, , , "

            Case "Ith Rune"
                runestats = "15, Weapons:, -25% To Targets Defense, Armour:, Regenerate Mana 15%, Helmets:, Regenerate Mana 15%, Shields:, Regenerate Mana 15%, , , "

            Case "Tal Rune"
                runestats = "17, Weapons:, +19 Poison Damage for 2 seconds, Armour:, Poison Resistance +30%, Helmets:, Poison Resistance +30%, Shields:, Regenerate Mana 15%, , , , , "

            Case "Ral Rune"
                runestats = "19, Weapons:, Adds 5 To 30 Fire Damage, Armour:, Fire Resistance +30%, Helmets:, Fire Resistance +30%, Shields:, Fire Resistance +30%, , , , , , , "

            Case "Ort Rune"
                runestats = "19, Weapons:, Adds 1 To 50 Light Damage, Armour:, Lightning Resistance +30%, Helmets:, Lightning Resistance +30%, Shields:, Lightning Resistance +30%, , , , , "

            Case "Thul Rune"
                runestats = "21, Weapons:, Adds 3 To 14 Cold Damage, Armour:, Cold Resistance +30%, Helmets:, Cold Resistance +30%, Shields:, Cold Resistance +30%, , , ,"

            Case "Amn Rune"
                runestats = "23, Weapons:, +7% Life Stolen Per Hit, Armour:, Attacker Takes Damage Of 14, Helmets:, Attacker Takes Damage Of 14, Shields:, Attacker Takes Damage Of 14, , , "

            Case "Sol Rune"
                runestats = "25, Weapons:, +9 To Minimum Damage, Armour:, Damage Reduced By 7, Helmets:, Damage Reduced By 7, Shields:, Damage Reduced By 7, , , "

            Case "Shael Rune"
                runestats = "27, Weapons:, Increased Attack Speed, Armour:, Faster Hit Recovery, Helmets:, Faster Hit Recovery, Shields:, Faster Block Rate, , , "

            Case "Shae Rune"
                runestats = "29, Weapons:, Increased Attack Speed, Armour:, Faster Hit Recovery, Helmets:, Faster Hit Recovery, Shields:, Faster Block Rate, , , "

            Case "Dol Rune"
                runestats = "31, Weapons:, +32% Monster Flees, Armour:, +7 Replenish Life, Helmets:, +7 Replenish Life, Shields:, +7 Replenish Life, , , "

            Case "Hel Rune"
                runestats = "1, Weapons:, -20 Requirements, Armour:, -15 Requirements, Helmets:, -15 Requirements, Shields:, -15 Requirements, , , "

            Case "Po Rune"
                runestats = "35, Weapons:, +10 To Vitality, Armour:, +10 To Vitality, Helmets:, +10 To Vitality, Shields:, +10 To Vitality, , , "

            Case "Io Rune"
                runestats = "35, Weapons:, +10 To Vitality, Armour:, +10 To Vitality, Helmets:, +10 To Vitality, Shields:, +10 To Vitality, , , "

            Case "Lum Rune"
                runestats = "37, +10 To Energy, Armour:, +10 To Energy, Helmets:, +10 To Energy, Shields:, +10 To Energy, , , , "

            Case "Ko Rune"
                runestats = "39, Weapons:, +10 To Dexterity, Armour:, +10 To Dexterity, Helmets:, +10 To Dexterity, Shields:, +10 To Dexterity, , , "

            Case "Fal Rune"
                runestats = "41, Weapons:, +10 To Strength, Armour:, +10 To Strength, Helmets:, +10 To Strength, Shields:, +10 To Strength, , , "

            Case "Lem Rune"
                runestats = "43, Weapons:, +75% Extra Gold From Slain Monsters, Armour:, +50% Extra Gold From Slain Monsters, Helmets:, +50% Extra Gold From Slain Monsters, Shields:, +50% Extra Gold From Slain Monsters, , , "

            Case "Pul Rune"
                runestats = "45, Weapons:, +100 To Attack Rating Against Demons, +175% Damage Against Demons , Armour:, +30% Enhanced Defense, Helmets:, +30% Enhanced Defense, Shields:, +30% Enhanced Defense, ,"


            Case "Um Rune"
                runestats = "47, Weapons:, +25% Chance Of Open Wounds, Armour:, +20 To All Resistances, Helmets:, +10 To All Resistances, Shields:, +20 To All Resistances, , , "

            Case "Mal Rune"
                runestats = "49, Weapons:, Prevent Monster Heal, Armour:, -7 To Magic Damage Received, Helmets:, -7 To Magic Damage Received, Shields:, -7 To Magic Damage Received, , , "

            Case "Ist Rune"
                runestats = "51, Weapons:, +30% To Magic Items, Armour:, +25% To Magic Find, Helmets:, +25% To Magic Items, Shields:, +25% To Magic Items, , , "

            Case "Gul Rune"
                runestats = "53, Weapons:, +20% To Attack Rating, Armour:, +3% To Max Poison Resist, Helmets:, +3% To Max Poison Resist, Shields:, +3% To Max Poison Resist, , , "

            Case "Vex Rune"
                runestats = "55, Weapons:, +7% Mana Stolen Per Hit, Armour:, +3% To Max Fire Resist, Helmets:, +3% To Max Fire Resist, Shields:, +3% To Max Fire Resist, , , "

            Case "Ohm Rune"
                runestats = "57, Weapons:, +50% Enhanced Damege, Armour:, +3% To Max Cold Resist, Helmets:, +3% To Max Cold Resist, Shields:, +3% To Max Cold Resist, , , "

            Case "Lo Rune"
                runestats = "59, Weapons:, +20% Chance Of Deadly Strike, Armour:, +3% To Max Lightning Resist, Helmets:, +3% To Max Lightning Resist, Shields:, +3% To Max Lightning Resist, , , "

            Case "Sur Rune"
                runestats = "61, Weapons:, Hit Blinds Target, Armour:, +5% To Max Mana, Helmets:, +5% To Max Mana, Shields:, +50 To Max Mana, , , "

            Case "Ber Rune"
                runestats = "63, Weapons:, +20% Chance Of Crushing Blow, Armour:, -8% To Damage Received, Helmets:, -8% To Damage Received, Shields:, -8% To Damage Received, , , "

            Case "Jo Rune"
                runestats = "65, Weapons:, Slows Target By 25%, Armour:, +5% To Max Life, Helmets:, +5% To Max Life, Shields:, +50 To Life, , , "

            Case "Cho Rune"
                runestats = "65, Weapons:, Slows Target By 25%, Armour:, +5% To Max Life, Helmets:, +5% To Max Life, Shields:, +50 To Life, , , "

            Case "Cham Rune"
                runestats = "67, Weapons:, Freeze Target, Armour:, Cannot Be Frozen, Helmets:, Cannot Be Frozen, Shields:, Cannot Be Frozen, , , "

            Case "Zod Rune"
                runestats = "69, Weapons:, Indestructable, Armour:, Indestructable, Helmets:, Indestructable, Shields:, Indestructable, , , "

        End Select


        Return runestats
    End Function
    Function GetGemsStats(ByVal gemname)

        Dim gemstats As String = ""
        Select Case gemname

            Case "Perfect Diamond"
                gemstats = "18,Weapons: +68% Damage to Undead, Armor: +100 to Attack Rating, Helms: +100 to Attack Rating, Shields: All Resistances 19"

            Case "Perfect Ruby"
                gemstats = "18,Weapons: Adds 16-20 Fire Damage, Armor: +38 to Life, Helms: +38 to Life, Shields: Fire Resist +40%"

            Case "Perfect Sapphire"
                gemstats = "18,Weapons: Adds 10-14 Cold Damage, Armor: +38 to Mana, Helms: +38 to Mana, Shields: Cold Resist +40%"

            Case "Perfect Emerald"
                gemstats = "18,Weapons: +100 Poison Damage over 7 Seconds, Armor: +10 to Dexterity, Helms: +10 to Dexterity, Shields: Poison Resist +40%"

            Case "Perfect Topaz"
                gemstats = "18,Weapons: Adds 10-14 Lightning Damage, Armor: 24% BetterChance of Finding Magic Items, Helms: 24% BetterChance of Finding Magic Items, Shields: Lightning Resist +40%"

            Case "Perfect Amethyst"
                gemstats = "18,Weapons: +160 To Attack Rating, Armor: +10 To Strength, Helms: +10 To Strength, Shields: +30 Defense"

            Case "Perfect Skull"
                gemstats = "18,Weapons: 4% Life Stolen Per Hit, 4% Mana Stolen Per Hit, Armor: Regenerate Mana 19%, Replenish life +5, Helms: Regenerate Mana 19%, Replenish life +5, Shields: Attacker Takes Damage of 20"
        End Select

        If gemstats = "" Then gemstats = "18, Weapons: N/A, Armor: N/A, Helms: N/A, Shields: N/A"
        Return gemstats
    End Function
End Module
