Module FileReadWrite

    'Load up the app configuration Values from \InstallDir\Settings.cfg file
    Sub LoadConfigFile()
        Dim file As System.IO.StreamReader
        file = My.Computer.FileSystem.OpenTextFileReader(Application.StartupPath + "\Settings.cfg")
        EtalPath = file.ReadLine() : Settings.EtalPathTEXTBOX.Text = EtalPath
        Databasefile = file.ReadLine() : Settings.DatabaseFileTEXTBOX.Text = Databasefile : DefaultDatabaseFile = Databasefile
        TimerMins = file.ReadLine() : Settings.NumericUpDown1.Value = TimerMins
        KeepPassPrivate = file.ReadLine() : If KeepPassPrivate = True Then Settings.CheckBox3.Checked = True
        AutoBackups = file.ReadLine() : If AutoBackups = True Then Settings.AutoBackupImportsCHECKBOX.Checked = True ' added this for auto backup setting
        EditBackups = file.ReadLine() : If AutoBackups = True Then Settings.AutoBackupImportsCHECKBOX.Checked = True ' added this for backup befor saving item edits
        RemoveDupeMule = file.ReadLine()
        Dim temp = file.ReadLine()
        If temp = "False" Then HideSearchPopup = False
        If temp = "True" Then HideSearchPopup = True
        If HideSearchPopup = True Then Settings.DisableSearchProgressBarCHECKBOX.Checked = True : ShowSearchProgress = False
        If HideSearchPopup = False Then Settings.DisableSearchProgressBarCHECKBOX.Checked = False : ShowSearchProgress = True
        file.Close()
    End Sub

    ' Below new database file read to object items improvement to deal with file read errors and utilise older database items
    ' may need further coding to have a single read file function
    ' oldfile option will be removed upon app rewrite
    Public Function ReadData(fName, oldfile)

        If My.Computer.FileSystem.FileExists(fName) = False Then
            Mymessages = "File doesn't exist"
            MyMessageBox()
            Return False
        End If

        Dim Reader = My.Computer.FileSystem.OpenTextFileReader(fName)

        Try ' catch any error during item importing
            Do
                If Reader.EndOfStream = True Then Exit Do
                Reader.ReadLine()
                If Reader.EndOfStream = True Then Exit Do
                Dim NewObject As New ItemObjects
                NewObject.ItemName = Reader.ReadLine
                If oldfile = True Then NewObject.Ilevel = "0"
                If oldfile = False Then NewObject.Ilevel = Reader.ReadLine
                NewObject.ItemBase = Reader.ReadLine
                NewObject.ItemQuality = Reader.ReadLine
                NewObject.RequiredCharacter = Reader.ReadLine
                NewObject.EtherealItem = Reader.ReadLine
                NewObject.Sockets = Reader.ReadLine
                NewObject.RuneWord = Reader.ReadLine
                NewObject.ThrowDamageMin = Reader.ReadLine
                NewObject.ThrowDamageMax = Reader.ReadLine
                NewObject.OneHandDamageMin = Reader.ReadLine
                NewObject.OneHandDamageMax = Reader.ReadLine
                NewObject.TwoHandDamageMin = Reader.ReadLine
                NewObject.TwoHandDamageMax = Reader.ReadLine
                NewObject.Defense = Reader.ReadLine
                NewObject.ChanceToBlock = Reader.ReadLine
                NewObject.QuantityMin = Reader.ReadLine
                NewObject.QuantityMax = Reader.ReadLine
                NewObject.DurabilityMin = Reader.ReadLine
                NewObject.DurabilityMax = Reader.ReadLine
                NewObject.RequiredStrength = Reader.ReadLine
                NewObject.RequiredDexterity = Reader.ReadLine
                NewObject.RequiredLevel = Reader.ReadLine
                NewObject.AttackClass = Reader.ReadLine
                NewObject.AttackSpeed = Reader.ReadLine
                NewObject.Stat1 = Reader.ReadLine
                NewObject.Stat2 = Reader.ReadLine
                NewObject.Stat3 = Reader.ReadLine
                NewObject.Stat4 = Reader.ReadLine
                NewObject.Stat5 = Reader.ReadLine
                NewObject.Stat6 = Reader.ReadLine
                NewObject.Stat7 = Reader.ReadLine
                NewObject.Stat8 = Reader.ReadLine
                NewObject.Stat9 = Reader.ReadLine
                NewObject.Stat10 = Reader.ReadLine
                NewObject.Stat11 = Reader.ReadLine
                NewObject.Stat12 = Reader.ReadLine
                NewObject.Stat13 = Reader.ReadLine
                NewObject.Stat14 = Reader.ReadLine
                NewObject.Stat15 = Reader.ReadLine
                NewObject.MuleName = Reader.ReadLine
                NewObject.MuleAccount = Reader.ReadLine
                NewObject.MulePass = Reader.ReadLine
                NewObject.PickitBot = Reader.ReadLine
                If oldfile = True Then NewObject.ItemRealm = "Unknown"
                If oldfile = False Then NewObject.ItemRealm = Reader.ReadLine
                NewObject.UserReference = Reader.ReadLine
                NewObject.ItemImage = Reader.ReadLine
                Objects.Add(NewObject)
            Loop Until Reader.EndOfStream
            Reader.Close()


        Catch ex As Exception
            Reader.Close()

            Return False
        End Try

        Objects.Sort(Function(x, y) x.ItemName.CompareTo(y.ItemName)) 'sort list alphabetically after assigning objects
        Form1.Display_Items() ' may need to move this following specific function calls ???

        Return True

    End Function


End Module
