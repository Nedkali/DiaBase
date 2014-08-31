Imports System.IO
Module AutoLogger
    Public Sub ImportLogFiles()

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

        ' This Backsup the database if logfiles exist to import (above routine checks this is true) and is set to do so in settings auto backup checkbox
        ' Routine currently will only backup the selected database - linking will have to be considered also b4 routine is complete <---------- [FINISH THIS ROB]

        If Settings.AutoBackupImportsCHECKBOX.Checked = True Then

            'get Backup directory path
            Dim BackupPath = Application.StartupPath + "\DataBase\Backup\" ' set path to backup folder

            'get name of database name fron databasefile path string
            Dim DBaseName = Databasefile.Replace(Application.StartupPath + "\DataBase\", "")

            'if an old backup of this database exists here already then this deletes it
            If My.Computer.FileSystem.FileExists(BackupPath + DBaseName) = True Then My.Computer.FileSystem.DeleteFile(BackupPath + DBaseName)

            'This saves the new database backup file
            My.Computer.FileSystem.CopyFile(Databasefile, BackupPath + DBaseName)
            Form1.RichTextBox1.AppendText(DBaseName & " Backup Successful" & vbCrLf) '  logger event window entry backup message, REMOVE THIS IF YOU LIKE
        End If

        Form1.RichTextBox1.AppendText("Logs to import = " & LogFilesList.Count & vbCrLf)
        Pretotal = Objects.Count

        ProcessLogFiles() 'moved rest to a separate sub to cut down on coded lines in a single sub
        Form1.SaveItems()
        Form1.Display_Items()

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

    Sub GetmuleaccountFiles()
        Dim Tally As Integer = 0
        PassFiles.Clear()
        Dim AllFiles As String() = (Directory.GetFiles(MuleDataPath, "*")).Select(Function(p) Path.GetFileName(p)).ToArray() ' gets file and crops path
        For Each item In AllFiles
            If AllFiles(Tally).IndexOf("_muleaccount.txt") > -1 Then PassFiles.Add(AllFiles(Tally))
            Tally = Tally + 1
        Next

    End Sub

    Function GetMulePass(ByVal accname)
        GetmuleaccountFiles()
        Dim counter As Integer = 0
        Dim temp As String = ""
        Dim returnstring As String = "Unknown"
        Dim temp1 As Array
        Form1.RichTextBox1.AppendText("Looking up " & accname & " information" & vbCrLf)
        For Each item In PassFiles

            Dim ReadPassFiles = My.Computer.FileSystem.OpenTextFileReader(MuleDataPath & PassFiles(counter))
            While ReadPassFiles.EndOfStream = False
                temp = ReadPassFiles.ReadLine()
                If temp.IndexOf(accname) <> -1 Then
                    returnstring = PassFiles(counter).Replace("_muleaccount.txt", "")
                    temp1 = temp.Split("/")
                    returnstring = returnstring + "," + temp1(1)
                    ReadPassFiles.Close()
                    Return returnstring
                    Exit For
                End If
            End While
            counter = counter + 1
        Next
        Return (",")
    End Function

    Private Sub ProcessLogFiles()
        Dim temp As String = ""
        Dim Tally As Integer = 0
        Dim mycounter As Integer = 0
        Dim found As Boolean = True
        Dim myarray As Array

        Do Until Tally = LogFilesList.Count
            If My.Computer.FileSystem.FileExists(MuleLogPath & LogFilesList(Tally)) = True Then 'Verify the log Exists
                Dim LogFile = My.Computer.FileSystem.OpenTextFileReader(MuleLogPath & LogFilesList(Tally))

                Dim thispickbot = LogFile.ReadLine() 'these lines should exist for each log
                Dim thislogmuleacc = LogFile.ReadLine()
                Dim thislogmulename = LogFile.ReadLine()
                Dim thislogpass = LogFile.ReadLine()
                LogFile.ReadLine()

                If (thispickbot = "Unknown" Or thislogpass = "Unknown") Then 'lets try to find them
                    Dim result As String = GetMulePass(thislogmuleacc)
                    myarray = result.Split(",")
                    If (myarray(0) <> "") Then
                        thispickbot = myarray(0)
                    End If
                    If (myarray(1) <> "") Then
                        thislogpass = myarray(1)
                    End If
                End If

                '------------------------------------------------------------------------------------------------------------------------------------------------
                'ROB DEBUG MESSAGE: \/\/ This has thrown an error for Sig. Seems to remove the dupes ok but then if import fails and database is then
                'saved the dupes remain removed. I think this is causing the error on a second attampt to re-log aborterd imports as the dupe was removed 
                'first import attempt so crashes when it attempts to re-remove. So I added Try command. NOT TOTALLY SURE ABOUT THIS THOUGH, JUST A HUNCH?

                'Maybe we could move this down a little to after the log imports successfully, could possibly be a better way to fix issue?

                If RemoveDupeMule = True Then                                               'Adding in remove duplicated mule option (Ned)
                    Try
                        For mc = Objects.Count - 1 To 0 Step -1
                            'MsgBox("Curent index = " & mc)
                            If Pretotal > 0 And Objects(mc).MuleName = thislogmulename Then
                                Objects.RemoveAt(mc)
                                Pretotal = Pretotal - 1
                                If Form1.AllItemsInDatabaseListBox.Items.Count > 0 Then     'Remove item from list if there
                                    Form1.AllItemsInDatabaseListBox.Items.RemoveAt(mc)
                                End If
                            End If
                        Next
                    Catch ex As Exception
                        'Mymessages = "Could Not Removed Duped Mule..." : MyMessageBox()    'Fail Message to notify user - Is probably not really nessicary
                        '                                                                   'saying something that wasnt there wasnt removed, lol.
                        '                                                                   'Will Comment This Out For Now :)
                    End Try
                End If
                '------------------------------------------------------------------------------------------------------------------------------------------------

                Do
                    Dim NewObject As New ItemObjects
                    NewObject.PickitBot = thispickbot
                    NewObject.MuleAccount = thislogmuleacc
                    NewObject.MuleName = thislogmulename
                    NewObject.MulePass = thislogpass

                    For x = 0 To 4 '     just in case of extra blank lines
                        If LogFile.EndOfStream = True Then Exit Do
                        temp = LogFile.ReadLine()
                        If temp <> "" Then Exit For
                    Next
                    NewObject.ItemName = temp 'these 5 lines should exist for each item
                    temp = LogFile.ReadLine()
                    NewObject.ItemBase = temp.Replace("Item Base ", "")
                    temp = LogFile.ReadLine() : myarray = Split(temp, " ")
                    NewObject.ItemQuality = myarray(2)
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
                                NewObject.AttackClass = NewObject.ItemBase
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
                            ' check for fix for item class
                            If temp.IndexOf("Class") > -1 Then
                                myarray = temp.Split(" ")
                                NewObject.AttackClass = myarray(0)
                                myarray = temp.Split("- ")
                                NewObject.AttackSpeed = LTrim(myarray(1))
                                Continue While
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

                    ' fixes area - correcting item imports
                    If NewObject.ItemName.IndexOf("Large Charm") > -1 Then NewObject.ItemBase = "Large Charm" ' torches misread as medium ????? weird maybe valid ???
                    If NewObject.ItemName.IndexOf("Grand Charm") > -1 Then NewObject.ItemBase = "Grand Charm" ' this also as above

                    If NewObject.ItemBase = "Rune" Then
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
                    If NewObject.ItemBase = "Gem" Then
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

                    ''check pickitbot var
                    'NewObject.PickitBot = thispickbot
                    Objects.Add(NewObject)

                Loop Until LogFile.EndOfStream
                LogFile.Close()
                My.Computer.FileSystem.MoveFile(MuleLogPath & LogFilesList(Tally), ArchiveFolder & LogFilesList(Tally), True)
            End If
            Tally = Tally + 1
        Loop
    End Sub

    Private Sub SaveLoggedItems(ByVal itemstart)

        Form1.RichTextBox1.AppendText("Saving to file " & Databasefile & vbCrLf)
        Form1.RichTextBox1.AppendText("Items to be saved = " & Objects.Count - itemstart & vbCrLf)
        Dim count = 0
        Try
            Dim LogWriter = My.Computer.FileSystem.OpenTextFileWriter(Databasefile, True)
            For x = itemstart To Objects.Count - 1
                LogWriter.WriteLine("--------------------")
                LogWriter.WriteLine(Objects(x).ItemName)
                LogWriter.WriteLine(Objects(x).ItemBase)
                LogWriter.WriteLine(Objects(x).ItemQuality)
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
                count = count + 1
            Next
            LogWriter.Close()

        Catch ex As Exception
            Form1.RichTextBox1.AppendText("File Write Error" & vbCrLf)

        End Try
        Form1.RichTextBox1.AppendText("Items Saved = " & count & vbCrLf)
        For y = itemstart To Objects.Count - 1
            Form1.AllItemsInDatabaseListBox.Items.Add(Objects(y).ItemName)
        Next
        Form1.ItemTallyTEXTBOX.Text = Objects.Count & " Items"
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

            Case "Jah Rune"
                runestats = "65, Weapons:, Ignore Target's Defense, Armour:, Increase Maximum Life 5%, Helmets:, Increase Maximum Life 5%, Shields:, +50 To Life, , , "
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

            Case "Chipped Diamond"
                gemstats = "1,Weapons: +28% Damage to Undead, Armor: +20 to Attack Rating, Helms: +20 to Attack Rating, Shields: All Resistances 6"

            Case "Chipped Ruby"
                gemstats = "1,Weapons: Adds 3-4 Fire Damage, Armor: +10 to Life, Helms: +10 to Life, Shields: Fire Resist +12%"

            Case "Chipped Sapphire"
                gemstats = "1,Weapons: Adds 1-3 Cold Damage, Armor: +10 to Mana, Helms: +10 to Mana, Shields: Cold Resist +12%"

            Case "Chipped Emerald"
                gemstats = "1,Weapons: +10 Poison Damage over 3 Seconds, Armor: +3 to Dexterity, Helms: +3 to Dexterity, Shields: Poison Resist +12%"

            Case "Chipped Topaz"
                gemstats = "1,Weapons: Adds 1-8 Lightning Damage, Armor: 9% BetterChance of Finding Magic Items, Helms: 9% BetterChance of Finding Magic Items, Shields: Lightning Resist +12%"

            Case "Chipped Amethyst"
                gemstats = "1,Weapons: +40 To Attack Rating, Armor: +3 To Strength, Helms: +3 To Strength, Shields: +8 Defense"

            Case "Chipped Skull"
                gemstats = "1,Weapons: 2% Life Stolen Per Hit, 1% Mana Stolen Per Hit, Armor: Regenerate Mana 8%, Replenish life +2, Helms: Regenerate Mana 8%, Replenish life +2, Shields: Attacker Takes Damage of 4"

            Case "Flawed Diamond"
                gemstats = "5,Weapons: +34% Damage to Undead, Armor: +40 to Attack Rating, Helms: +40 to Attack Rating, Shields: All Resistances 8"

            Case "Flawed Ruby"
                gemstats = "5,Weapons: Adds 5-8 Fire Damage, Armor: +17 to Life, Helms: +17 to Life, Shields: Fire Resist +16%"

            Case "Flawed Sapphire"
                gemstats = "5,Weapons: Adds 3-5 Cold Damage, Armor: +17 to Mana, Helms: +17 to Mana, Shields: Cold Resist +16%"

            Case "Flawed Emerald"
                gemstats = "5,Weapons: +20 Poison Damage over 4 Seconds, Armor: +4 to Dexterity, Helms: +4 to Dexterity, Shields: Poison Resist +16%"

            Case "Flawed Topaz"
                gemstats = "5,Weapons: Adds 1-14 Lightning Damage, Armor: 13% BetterChance of Finding Magic Items, Helms: 13% BetterChance of Finding Magic Items, Shields: Lightning Resist +16%"

            Case "Flawed Amethyst"
                gemstats = "5,Weapons: +60 To Attack Rating, Armor: +4 To Strength, Helms: +4 To Strength, Shields: +12 Defense"

            Case "Flawed Skull"
                gemstats = "5,Weapons: 2% Life Stolen Per Hit, 2% Mana Stolen Per Hit, Armor: Regenerate Mana 8%, Replenish life +3, Helms: Regenerate Mana 8%, Replenish life +3, Shields: Attacker Takes Damage of 8"

            Case "Diamond"
                gemstats = "12,Weapons: +44% Damage to Undead, Armor: +60 to Attack Rating, Helms: +60 to Attack Rating, Shields: All Resistances 11"

            Case "Ruby"
                gemstats = "12,Weapons: Adds 8-12 Fire Damage, Armor: +24 to Life, Helms: +24 to Life, Shields: Fire Resist +22%"

            Case "Sapphire"
                gemstats = "12,Weapons: Adds 4-7 Cold Damage, Armor: +24 to Mana, Helms: +24 to Mana, Shields: Cold Resist +22%"

            Case "Emerald"
                gemstats = "12,Weapons: +40 Poison Damage over 5 Seconds, Armor: +6 to Dexterity, Helms: +6 to Dexterity, Shields: Poison Resist +22%"

            Case "Topaz"
                gemstats = "12,Weapons: Adds 1-22 Lightning Damage, Armor: 16% BetterChance of Finding Magic Items, Helms: 16% BetterChance of Finding Magic Items, Shields: Lightning Resist +22%"

            Case "Amethyst"
                gemstats = "12,Weapons: +80 To Attack Rating, Armor: +6 To Strength, Helms: +6 To Strength, Shields: +18 Defense"

            Case "Skull"
                gemstats = "12,Weapons: 3% Life Stolen Per Hit, 2% Mana Stolen Per Hit, Armor: Regenerate Mana 12%, Replenish life +3, Helms: Regenerate Mana 12%, Replenish life +3, Shields: Attacker Takes Damage of 12"

            Case "Flawless Diamond"
                gemstats = "15,Weapons: +54% Damage to Undead, Armor: +80 to Attack Rating, Helms: +80 to Attack Rating, Shields: All Resistances 14"

            Case "Flawless Ruby"
                gemstats = "15,Weapons: Adds 10-16 Fire Damage, Armor: +31 to Life, Helms: +31 to Life, Shields: Fire Resist +28%"

            Case "Flawless Sapphire"
                gemstats = "15,Weapons: Adds 6-10 Cold Damage, Armor: +31 to Mana, Helms: +31 to Mana, Shields: Cold Resist +28%"

            Case "Flawless Emerald"
                gemstats = "15,Weapons: +60 Poison Damage over 6 Seconds, Armor: +8 to Dexterity, Helms: +8 to Dexterity, Shields: Poison Resist +28%"

            Case "Flawless Topaz"
                gemstats = "15,Weapons: Adds 1-30 Lightning Damage, Armor: 20% BetterChance of Finding Magic Items, Helms: 20% BetterChance of Finding Magic Items, Shields: Lightning Resist +28%"

            Case "Flawless Amethyst"
                gemstats = "15,Weapons: +100 To Attack Rating, Armor: +8 To Strength, Helms: +8 To Strength, Shields: +24 Defense"

            Case "Flawless Skull"
                gemstats = "15,Weapons: 3% Life Stolen Per Hit, 3% Mana Stolen Per Hit, Armor: Regenerate Mana 12%, Replenish life +4, Helms: Regenerate Mana 12%, Replenish life +4, Shields: Attacker Takes Damage of 16"

            Case "Perfect Diamond"
                gemstats = "18,Weapons: +68% Damage to Undead, Armor: +100 to Attack Rating, Helms: +100 to Attack Rating, Shields: All Resistances 19"

            Case "Perfect Ruby"
                gemstats = "18,Weapons: Adds 15-20 Fire Damage, Armor: +38 to Life, Helms: +38 to Life, Shields: Fire Resist +40%"

            Case "Perfect Sapphire"
                gemstats = "18,Weapons: Adds 10-14 Cold Damage, Armor: +38 to Mana, Helms: +38 to Mana, Shields: Cold Resist +40%"

            Case "Perfect Emerald"
                gemstats = "18,Weapons: +100 Poison Damage over 7 Seconds, Armor: +10 to Dexterity, Helms: +10 to Dexterity, Shields: Poison Resist +40%"

            Case "Perfect Topaz"
                gemstats = "18,Weapons: Adds 1-40 Lightning Damage, Armor: 24% BetterChance of Finding Magic Items, Helms: 24% BetterChance of Finding Magic Items, Shields: Lightning Resist +40%"

            Case "Perfect Amethyst"
                gemstats = "18,Weapons: +150 To Attack Rating, Armor: +10 To Strength, Helms: +10 To Strength, Shields: +30 Defense"

            Case "Perfect Skull"
                gemstats = "18,Weapons: 4% Life Stolen Per Hit, 3% Mana Stolen Per Hit, Armor: Regenerate Mana 19%, Replenish life +5, Helms: Regenerate Mana 19%, Replenish life +5, Shields: Attacker Takes Damage of 20"

        End Select
        If gemstats = "" Then gemstats = "18, Weapons: N/A, Armor: N/A, Helms: N/A, Shields: N/A"
        Return gemstats
    End Function
End Module
