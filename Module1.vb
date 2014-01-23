Module Module1
    Public Objects As List(Of ItemObjects) = New List(Of ItemObjects) ' sets up new list of class items for database


    'DataBase variables
    Public iEdit As Integer         'used in item edit form to locate item(array) number being edited
    Public EtalPath As String
    Public DataBasePath As String
    Public TimerMins As Integer
    Public Timercount As Integer
    Public TimerSecs As Integer
    Public Mymessages As String
    ' needs to be set to true when logger is doing log reads/imports and then set to false when completed
    ' Need this to prevent reading database while logger maybe trying to access it - logger will need priority
    ' or will most likely trigger an exception and crash Program
    Public LoggerRunning As Boolean = False
    Public KeepPassPrivate As Boolean = True

    'Logging variables
    Public MuleLogPath As String
    Public Databasefile As String
    Public MuleDataPath As String
    Public ArchiveFolder As String
    Public AutoBackups As String
    Public EditBackups As String
    Public MyCounter As Array
    Public SearchReferenceList As List(Of String) = New List(Of String)
    Public RefineSearchReferenceList As List(Of String) = New List(Of String)
    Public StringMatches As List(Of String) = New List(Of String)
    Public IntegerMatches As List(Of String) = New List(Of String)
    Dim GetAllIntegers As List(Of String) = New List(Of String)


    Public ItemNamePulldownList As List(Of String) = New List(Of String)
    Public UniqueAttribsPulldownList As List(Of String) = New List(Of String)
    Public LogFilesList As List(Of String) = New List(Of String)    'Holds all logs found in log directory
    Public PassFiles As List(Of String) = New List(Of String)       'Holds all _muleaccount.txt file used to get mule pass and mule account
    Public LogType As List(Of String) = New List(Of String)

    Public Sub MyMessageBox()
        UserMessaging.ShowDialog()
    End Sub


    'This Reads the database from file and puts it in the object database (moved from form1 to allow it to be used at startup to load default database)
    Sub OpenDatabaseRoutine(DatabaseFile)
        Form1.SearchLISTBOX.Items.Clear()               ' clears out search lists and other stat windows
        Form1.MuleAccountTextbox.Text = ""
        Form1.MuleNameTextbox.Text = ""
        Form1.MulePassTextbox.Text = ""
        Form1.RichTextBox2.Text = ""
        Form1.PictureBox1.Image = Nothing
        Form1.AllItemsInDatabaseListBox.Items.Clear()   ' clears items listed
        If Objects.Count > 0 Then                       ' had to be done this way - havent figured out a better way for now
            Objects.RemoveRange(1, Objects.Count - 1)
            Objects.RemoveAt(0)
        End If

        Dim Reader = My.Computer.FileSystem.OpenTextFileReader(DatabaseFile)

        Do
            If Reader.EndOfStream = True Then Exit Do
            Reader.ReadLine()
            If Reader.EndOfStream = True Then Exit Do
            Dim NewObject As New ItemObjects

            NewObject.ItemName = Reader.ReadLine
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
            NewObject.UserReference = Reader.ReadLine
            NewObject.ItemImage = Reader.ReadLine

            Objects.Add(NewObject)
        Loop Until Reader.EndOfStream
        Reader.Close()

        Form1.display_items()

    End Sub

    Sub LoadConfigFile()
        Dim file As System.IO.StreamReader
        file = My.Computer.FileSystem.OpenTextFileReader(Application.StartupPath + "\Settings.cfg")
        EtalPath = file.ReadLine() : Settings.EtalPathTEXTBOX.Text = EtalPath
        DataBaseFile = file.ReadLine() : Settings.DatabaseFileTEXTBOX.Text = DataBaseFile
        TimerMins = file.ReadLine() : Settings.NumericUpDown1.Value = TimerMins
        KeepPassPrivate = file.ReadLine() : If KeepPassPrivate = True Then Settings.CheckBox3.Checked = True
        AutoBackups = file.ReadLine() : If AutoBackups = True Then Settings.AutoBackupImportsCHECKBOX.Checked = True ' added this for auto backup setting
        EditBackups = file.ReadLine() : If AutoBackups = True Then Settings.AutoBackupImportsCHECKBOX.Checked = True ' added this for backup befor saving item edits
        file.Close()
    End Sub
    Public Function ItemImageList(sender As Integer)
        Dim item As String
        Select Case Int(sender)
            Case 0
                item = "hax"
            Case 1
                item = "axe"
            Case 2
                item = "2ax"
            Case 3
                item = "mpi"
            Case 4
                item = "wax"
            Case 5
                item = "lax"
            Case 6
                item = "bax"
            Case 7
                item = "btx"
            Case 8
                item = "gax"
            Case 9
                item = "gix"
            Case 10
                item = "wnd"
            Case 11
                item = "ywn"
            Case 12
                item = "bwn"
            Case 13
                item = "gwn"
            Case 14
                item = "clb"
            Case 15
                item = "scp"
            Case 16
                item = "gsc"
            Case 17
                item = "wsp"
            Case 18
                item = "spc"
            Case 19
                item = "mac"
            Case 20
                item = "mst"
            Case 21
                item = "fla"
            Case 22
                item = "whm"
            Case 23
                item = "mau"
            Case 24
                item = "gma"
            Case 25
                item = "ssd"
            Case 26
                item = "scm"
            Case 27
                item = "sbr"
            Case 28
                item = "flc"
            Case 29
                item = "crs"
            Case 30
                item = "bsd"
            Case 31
                item = "lsd"
            Case 32
                item = "wsd"
            Case 33
                item = "2hs"
            Case 34
                item = "clm"
            Case 35
                item = "gis"
            Case 36
                item = "bsw"
            Case 37
                item = "flb"
            Case 38
                item = "gsd"
            Case 39
                item = "dgr"
            Case 40
                item = "dir"
            Case 41
                item = "kri"
            Case 42
                item = "bld"
            Case 43
                item = "tkf"
            Case 44
                item = "tax"
            Case 45
                item = "bkf"
            Case 46
                item = "bal"
            Case 47
                item = "jav"
            Case 48
                item = "pil"
            Case 49
                item = "ssp"
            Case 50
                item = "glv"
            Case 51
                item = "tsp"
            Case 52
                item = "spr"
            Case 53
                item = "tri"
            Case 54
                item = "brn"
            Case 55
                item = "spt"
            Case 56
                item = "pik"
            Case 57
                item = "bar"
            Case 58
                item = "vou"
            Case 59
                item = "scy"
            Case 60
                item = "pax"
            Case 61
                item = "hal"
            Case 62
                item = "wsc"
            Case 63
                item = "sst"
            Case 64
                item = "lst"
            Case 65
                item = "cst"
            Case 66
                item = "bst"
            Case 67
                item = "wst"
            Case 68
                item = "sbw"
            Case 69
                item = "hbw"
            Case 70
                item = "lbw"
            Case 71
                item = "cbw"
            Case 72
                item = "sbb"
            Case 73
                item = "lbb"
            Case 74
                item = "swb"
            Case 75
                item = "lwb"
            Case 76
                item = "lxb"
            Case 77
                item = "mxb"
            Case 78
                item = "hxb"
            Case 79
                item = "rxb"
            Case 80
                item = "gps"
            Case 81
                item = "ops"
            Case 82
                item = "gpm"
            Case 83
                item = "opm"
            Case 84
                item = "gpl"
            Case 85
                item = "opl"
            Case 86
                item = "nia"
            Case 87
                item = "g33"
            Case 88
                item = "leg"
            Case 89
                item = "hdm"
            Case 90
                item = "hfh"
            Case 91
                item = "hst"
            Case 92
                item = "msf"
            Case 93
                item = "9ha"
            Case 94
                item = "9ax"
            Case 95
                item = "92a"
            Case 96
                item = "9mp"
            Case 97
                item = "9wa"
            Case 98
                item = "9la"
            Case 99
                item = "9ba"
            Case 100
                item = "9bt"
            Case 101
                item = "9ga"
            Case 102
                item = "9gi"
            Case 103
                item = "9wn"
            Case 104
                item = "9yw"
            Case 105
                item = "9bw"
            Case 106
                item = "9gw"
            Case 107
                item = "9cl"
            Case 108
                item = "9sc"
            Case 109
                item = "9qs"
            Case 110
                item = "9ws"
            Case 111
                item = "9sp"
            Case 112
                item = "9ma"
            Case 113
                item = "9mt"
            Case 114
                item = "9fl"
            Case 115
                item = "9wh"
            Case 116
                item = "9m9"
            Case 117
                item = "9gm"
            Case 118
                item = "9ss"
            Case 119
                item = "9sm"
            Case 120
                item = "9sb"
            Case 121
                item = "9fc"
            Case 122
                item = "9cr"
            Case 123
                item = "9bs"
            Case 124
                item = "9ls"
            Case 125
                item = "9wd"
            Case 126
                item = "92h"
            Case 127
                item = "9cm"
            Case 128
                item = "9gs"
            Case 129
                item = "9b9"
            Case 130
                item = "9fb"
            Case 131
                item = "9gd"
            Case 132
                item = "9dg"
            Case 133
                item = "9di"
            Case 134
                item = "9kr"
            Case 135
                item = "9bl"
            Case 136
                item = "9tk"
            Case 137
                item = "9ta"
            Case 138
                item = "9bk"
            Case 139
                item = "9b8"
            Case 140
                item = "9ja"
            Case 141
                item = "9pi"
            Case 142
                item = "9s9"
            Case 143
                item = "9gl"
            Case 144
                item = "9ts"
            Case 145
                item = "9sr"
            Case 146
                item = "9tr"
            Case 147
                item = "9br"
            Case 148
                item = "9st"
            Case 149
                item = "9p9"
            Case 150
                item = "9b7"
            Case 151
                item = "9vo"
            Case 152
                item = "9s8"
            Case 153
                item = "9pa"
            Case 154
                item = "9h9"
            Case 155
                item = "9wc"
            Case 156
                item = "8ss"
            Case 157
                item = "8ls"
            Case 158
                item = "8cs"
            Case 159
                item = "8bs"
            Case 160
                item = "8ws"
            Case 161
                item = "8sb"
            Case 162
                item = "8hb"
            Case 163
                item = "8lb"
            Case 164
                item = "8cb"
            Case 165
                item = "8s8"
            Case 166
                item = "8l8"
            Case 167
                item = "8sw"
            Case 168
                item = "8lw"
            Case 169
                item = "8lx"
            Case 170
                item = "8mx"
            Case 171
                item = "8hx"
            Case 172
                item = "8rx"
            Case 173
                item = "qf1"
            Case 174
                item = "qf2"
            Case 175
                item = "ktr"
            Case 176
                item = "wrb"
            Case 177
                item = "axf"
            Case 178
                item = "ces"
            Case 179
                item = "clw"
            Case 180
                item = "btl"
            Case 181
                item = "skr"
            Case 182
                item = "9ar"
            Case 183
                item = "9wb"
            Case 184
                item = "9xf"
            Case 185
                item = "9cs"
            Case 186
                item = "9lw"
            Case 187
                item = "9tw"
            Case 188
                item = "9qr"
            Case 189
                item = "7ar"
            Case 190
                item = "7wb"
            Case 191
                item = "7xf"
            Case 192
                item = "7cs"
            Case 193
                item = "7lw"
            Case 194
                item = "7tw"
            Case 195
                item = "7qr"
            Case 196
                item = "7ha"
            Case 197
                item = "7ax"
            Case 198
                item = "72a"
            Case 199
                item = "7mp"
            Case 200
                item = "7wa"
            Case 201
                item = "7la"
            Case 202
                item = "7ba"
            Case 203
                item = "7bt"
            Case 204
                item = "7ga"
            Case 205
                item = "7gi"
            Case 206
                item = "7wn"
            Case 207
                item = "7yw"
            Case 208
                item = "7bw"
            Case 209
                item = "7gw"
            Case 210
                item = "7cl"
            Case 211
                item = "7sc"
            Case 212
                item = "7qs"
            Case 213
                item = "7ws"
            Case 214
                item = "7sp"
            Case 215
                item = "7ma"
            Case 216
                item = "7mt"
            Case 217
                item = "7fl"
            Case 218
                item = "7wh"
            Case 219
                item = "7m7"
            Case 220
                item = "7gm"
            Case 221
                item = "7ss"
            Case 222
                item = "7sm"
            Case 223
                item = "7sb"
            Case 224
                item = "7fc"
            Case 225
                item = "7cr"
            Case 226
                item = "7bs"
            Case 227
                item = "7ls"
            Case 228
                item = "7wd"
            Case 229
                item = "72h"
            Case 230
                item = "7cm"
            Case 231
                item = "7gs"
            Case 232
                item = "7b7"
            Case 233
                item = "7fb"
            Case 234
                item = "7gd"
            Case 235
                item = "7dg"
            Case 236
                item = "7di"
            Case 237
                item = "7kr"
            Case 238
                item = "7bl"
            Case 239
                item = "7tk"
            Case 240
                item = "7ta"
            Case 241
                item = "7bk"
            Case 242
                item = "7b8"
            Case 243
                item = "7ja"
            Case 244
                item = "7pi"
            Case 245
                item = "7s7"
            Case 246
                item = "7gl"
            Case 247
                item = "7ts"
            Case 248
                item = "7sr"
            Case 249
                item = "7tr"
            Case 250
                item = "7br"
            Case 251
                item = "7st"
            Case 252
                item = "7p7"
            Case 253
                item = "7o7"
            Case 254
                item = "7vo"
            Case 255
                item = "7s8"
            Case 256
                item = "7pa"
            Case 257
                item = "7h7"
            Case 258
                item = "7wc"
            Case 259
                item = "6ss"
            Case 260
                item = "6ls"
            Case 261
                item = "6cs"
            Case 262
                item = "6bs"
            Case 263
                item = "6ws"
            Case 264
                item = "6sb"
            Case 265
                item = "6hb"
            Case 266
                item = "6lb"
            Case 267
                item = "6cb"
            Case 268
                item = "6s7"
            Case 269
                item = "6l7"
            Case 270
                item = "6sw"
            Case 271
                item = "6lw"
            Case 272
                item = "6lx"
            Case 273
                item = "6mx"
            Case 274
                item = "6hx"
            Case 275
                item = "6rx"
            Case 276
                item = "ob1"
            Case 277
                item = "ob2"
            Case 278
                item = "ob3"
            Case 279
                item = "ob4"
            Case 280
                item = "ob5"
            Case 281
                item = "am1"
            Case 282
                item = "am2"
            Case 283
                item = "am3"
            Case 284
                item = "am4"
            Case 285
                item = "am5"
            Case 286
                item = "ob6"
            Case 287
                item = "ob7"
            Case 288
                item = "ob8"
            Case 289
                item = "ob9"
            Case 290
                item = "oba"
            Case 291
                item = "am6"
            Case 292
                item = "am7"
            Case 293
                item = "am8"
            Case 294
                item = "am9"
            Case 295
                item = "ama"
            Case 296
                item = "obb"
            Case 297
                item = "obc"
            Case 298
                item = "obd"
            Case 299
                item = "obe"
            Case 300
                item = "obf"
            Case 301
                item = "amb"
            Case 302
                item = "amc"
            Case 303
                item = "amd"
            Case 304
                item = "ame"
            Case 305
                item = "amf"
            Case 306
                item = "cap"
            Case 307
                item = "skp"
            Case 308
                item = "hlm"
            Case 309
                item = "fhl"
            Case 310
                item = "ghm"
            Case 311
                item = "crn"
            Case 312
                item = "msk"
            Case 313
                item = "qui"
            Case 314
                item = "lea"
            Case 315
                item = "hla"
            Case 316
                item = "stu"
            Case 317
                item = "rng"
            Case 318
                item = "scl"
            Case 319
                item = "chn"
            Case 320
                item = "brs"
            Case 321
                item = "spl"
            Case 322
                item = "plt"
            Case 323
                item = "fld"
            Case 324
                item = "gth"
            Case 325
                item = "ful"
            Case 326
                item = "aar"
            Case 327
                item = "ltp"
            Case 328
                item = "buc"
            Case 329
                item = "sml"
            Case 330
                item = "lrg"
            Case 331
                item = "kit"
            Case 332
                item = "tow"
            Case 333
                item = "gts"
            Case 334
                item = "lgl"
            Case 335
                item = "vgl"
            Case 336
                item = "mgl"
            Case 337
                item = "tgl"
            Case 338
                item = "hgl"
            Case 339
                item = "lbt"
            Case 340
                item = "vbt"
            Case 341
                item = "mbt"
            Case 342
                item = "tbt"
            Case 343
                item = "hbt"
            Case 344
                item = "lbl"
            Case 345
                item = "vbl"
            Case 346
                item = "mbl"
            Case 347
                item = "tbl"
            Case 348
                item = "hbl"
            Case 349
                item = "bhm"
            Case 350
                item = "bsh"
            Case 351
                item = "spk"
            Case 352
                item = "xap"
            Case 353
                item = "xkp"
            Case 354
                item = "xlm"
            Case 355
                item = "xhl"
            Case 356
                item = "xhm"
            Case 357
                item = "xrn"
            Case 358
                item = "xsk"
            Case 359
                item = "xui"
            Case 360
                item = "xea"
            Case 361
                item = "xla"
            Case 362
                item = "xtu"
            Case 363
                item = "xng"
            Case 364
                item = "xcl"
            Case 365
                item = "xhn"
            Case 366
                item = "xrs"
            Case 367
                item = "xpl"
            Case 368
                item = "xlt"
            Case 369
                item = "xld"
            Case 370
                item = "xth"
            Case 371
                item = "xul"
            Case 372
                item = "xar"
            Case 373
                item = "xtp"
            Case 374
                item = "xuc"
            Case 375
                item = "xml"
            Case 376
                item = "xrg"
            Case 377
                item = "xit"
            Case 378
                item = "xow"
            Case 379
                item = "xts"
            Case 380
                item = "xlg"
            Case 381
                item = "xvg"
            Case 382
                item = "xmg"
            Case 383
                item = "xtg"
            Case 384
                item = "xhg"
            Case 385
                item = "xlb"
            Case 386
                item = "xvb"
            Case 387
                item = "xmb"
            Case 388
                item = "xtb"
            Case 389
                item = "xhb"
            Case 390
                item = "zlb"
            Case 391
                item = "zvb"
            Case 392
                item = "zmb"
            Case 393
                item = "ztb"
            Case 394
                item = "zhb"
            Case 395
                item = "xh9"
            Case 396
                item = "xsh"
            Case 397
                item = "xpk"
            Case 398
                item = "dr1"
            Case 399
                item = "dr2"
            Case 400
                item = "dr3"
            Case 401
                item = "dr4"
            Case 402
                item = "dr5"
            Case 403
                item = "ba1"
            Case 404
                item = "ba2"
            Case 405
                item = "ba3"
            Case 406
                item = "ba4"
            Case 407
                item = "ba5"
            Case 408
                item = "pa1"
            Case 409
                item = "pa2"
            Case 410
                item = "pa3"
            Case 411
                item = "pa4"
            Case 412
                item = "pa5"
            Case 413
                item = "ne1"
            Case 414
                item = "ne2"
            Case 415
                item = "ne3"
            Case 416
                item = "ne4"
            Case 417
                item = "ne5"
            Case 418
                item = "ci0"
            Case 419
                item = "ci1"
            Case 420
                item = "ci2"
            Case 421
                item = "ci3"
            Case 422
                item = "uap"
            Case 423
                item = "ukp"
            Case 424
                item = "ulm"
            Case 425
                item = "uhl"
            Case 426
                item = "uhm"
            Case 427
                item = "urn"
            Case 428
                item = "usk"
            Case 429
                item = "uui"
            Case 430
                item = "uea"
            Case 431
                item = "ula"
            Case 432
                item = "utu"
            Case 433
                item = "ung"
            Case 434
                item = "ucl"
            Case 435
                item = "uhn"
            Case 436
                item = "urs"
            Case 437
                item = "upl"
            Case 438
                item = "ult"
            Case 439
                item = "uld"
            Case 440
                item = "uth"
            Case 441
                item = "uul"
            Case 442
                item = "uar"
            Case 443
                item = "utp"
            Case 444
                item = "uuc"
            Case 445
                item = "uml"
            Case 446
                item = "urg"
            Case 447
                item = "uit"
            Case 448
                item = "uow"
            Case 449
                item = "uts"
            Case 450
                item = "ulg"
            Case 451
                item = "uvg"
            Case 452
                item = "umg"
            Case 453
                item = "utg"
            Case 454
                item = "uhg"
            Case 455
                item = "ulb"
            Case 456
                item = "uvb"
            Case 457
                item = "umb"
            Case 458
                item = "utb"
            Case 459
                item = "uhb"
            Case 460
                item = "ulc"
            Case 461
                item = "uvc"
            Case 462
                item = "umc"
            Case 463
                item = "utc"
            Case 464
                item = "uhc"
            Case 465
                item = "uh9"
            Case 466
                item = "ush"
            Case 467
                item = "upk"
            Case 468
                item = "dr6"
            Case 469
                item = "dr7"
            Case 470
                item = "dr8"
            Case 471
                item = "dr9"
            Case 472
                item = "dra"
            Case 473
                item = "ba6"
            Case 474
                item = "ba7"
            Case 475
                item = "ba8"
            Case 476
                item = "ba9"
            Case 477
                item = "baa"
            Case 478
                item = "pa6"
            Case 479
                item = "pa7"
            Case 480
                item = "pa8"
            Case 481
                item = "pa9"
            Case 482
                item = "paa"
            Case 483
                item = "ne6"
            Case 484
                item = "ne7"
            Case 485
                item = "ne8"
            Case 486
                item = "ne9"
            Case 487
                item = "nea"
            Case 488
                item = "drb"
            Case 489
                item = "drc"
            Case 490
                item = "drd"
            Case 491
                item = "dre"
            Case 492
                item = "drf"
            Case 493
                item = "bab"
            Case 494
                item = "bac"
            Case 495
                item = "bad"
            Case 496
                item = "bae"
            Case 497
                item = "baf"
            Case 498
                item = "pab"
            Case 499
                item = "pac"
            Case 500
                item = "pad"
            Case 501
                item = "pae"
            Case 502
                item = "paf"
            Case 503
                item = "neb"
            Case 504
                item = "neg"
            Case 505
                item = "ned"
            Case 506
                item = "nee"
            Case 507
                item = "nef"
            Case 508
                item = "nia"
            Case 509
                item = "nia"
            Case 510
                item = "nia"
            Case 511
                item = "nia"
            Case 512
                item = "nia"
            Case 513
                item = "vps"
            Case 514
                item = "yps"
            Case 515
                item = "rvs"
            Case 516
                item = "rvl"
            Case 517
                item = "wms"
            Case 518
                item = "tbk"
            Case 519
                item = "ibk"
            Case 520
                item = "amu"
            Case 521
                item = "vip"
            Case 522
                item = "rin"
            Case 523
                item = "nia"
            Case 524
                item = "bks"
            Case 525
                item = "bkd"
            Case 526
                item = "aqv"
            Case 527
                item = "tch"
            Case 528
                item = "cqv"
            Case 529
                item = "tsc"
            Case 530
                item = "isc"
            Case 531
                item = "hrt"
            Case 532
                item = "nia"
            Case 533
                item = "nia"
            Case 534
                item = "nia"
            Case 535
                item = "nia"
            Case 536
                item = "nia"
            Case 537
                item = "nia"
            Case 538
                item = "nia"
            Case 539
                item = "nia"
            Case 540
                item = "nia"
            Case 541
                item = "nia"
            Case 542
                item = "nia"
            Case 543
                item = "key"
            Case 544
                item = "nia"
            Case 545
                item = "xyz"
            Case 546
                item = "j34"
            Case 547
                item = "g34"
            Case 548
                item = "bbb"
            Case 549
                item = "box"
            Case 550
                item = "tr1"
            Case 551
                item = "mss"
            Case 552
                item = "ass"
            Case 553
                item = "qey"
            Case 554
                item = "qhr"
            Case 555
                item = "qbr"
            Case 556
                item = "ear"
            Case 557
                item = "gcv"
            Case 558
                item = "gfv"
            Case 559
                item = "gsv"
            Case 560
                item = "gzv"
            Case 561
                item = "gpv"
            Case 562
                item = "gcy"
            Case 563
                item = "gfy"
            Case 564
                item = "gsy"
            Case 565
                item = "gly"
            Case 566
                item = "gpy"
            Case 567
                item = "gcb"
            Case 568
                item = "gfb"
            Case 569
                item = "gsb"
            Case 570
                item = "glb"
            Case 571
                item = "gpb"
            Case 572
                item = "gcg"
            Case 573
                item = "gfg"
            Case 574
                item = "gsg"
            Case 575
                item = "glg"
            Case 576
                item = "gpg"
            Case 577
                item = "gcr"
            Case 578
                item = "gfr"
            Case 579
                item = "gsr"
            Case 580
                item = "glr"
            Case 581
                item = "gpr"
            Case 582
                item = "gcw"
            Case 583
                item = "gfw"
            Case 584
                item = "gsw"
            Case 585
                item = "glw"
            Case 586
                item = "gpw"
            Case 587
                item = "hp1"
            Case 588
                item = "hp2"
            Case 589
                item = "hp3"
            Case 590
                item = "hp4"
            Case 591
                item = "hp5"
            Case 592
                item = "mp1"
            Case 593
                item = "mp2"
            Case 594
                item = "mp3"
            Case 595
                item = "mp4"
            Case 596
                item = "mp5"
            Case 597
                item = "skc"
            Case 598
                item = "skf"
            Case 599
                item = "sku"
            Case 600
                item = "skl"
            Case 601
                item = "skz"
            Case 602
                item = "nia"
            Case 603
                item = "cm1"
            Case 604
                item = "cm2"
            Case 605
                item = "cm3"
            Case 606
                item = "nia"
            Case 607
                item = "nia"
            Case 608
                item = "nia"
            Case 609
                item = "nia"
            Case 610
                item = "r01"
            Case 611
                item = "r02"
            Case 612
                item = "r03"
            Case 613
                item = "r04"
            Case 614
                item = "r05"
            Case 615
                item = "r06"
            Case 616
                item = "r07"
            Case 617
                item = "r08"
            Case 618
                item = "r09"
            Case 619
                item = "r10"
            Case 620
                item = "r11"
            Case 621
                item = "r12"
            Case 622
                item = "r13"
            Case 623
                item = "r14"
            Case 624
                item = "r15"
            Case 625
                item = "r16"
            Case 626
                item = "r17"
            Case 627
                item = "r18"
            Case 628
                item = "r19"
            Case 629
                item = "r20"
            Case 630
                item = "r21"
            Case 631
                item = "r22"
            Case 632
                item = "r23"
            Case 633
                item = "r24"
            Case 634
                item = "r25"
            Case 635
                item = "r26"
            Case 636
                item = "r27"
            Case 637
                item = "r28"
            Case 638
                item = "r29"
            Case 639
                item = "r30"
            Case 640
                item = "r31"
            Case 641
                item = "r32"
            Case 642
                item = "r33"
            Case 643
                item = "jew"
            Case 644
                item = "ice"
            Case 645
                item = "nia"
            Case 646
                item = "tr2"
            Case 647
                item = "pk1"
            Case 648
                item = "pk2"
            Case 649
                item = "pk3"
            Case 650
                item = "dhn"
            Case 651
                item = "bey"
            Case 652
                item = "mbr"
            Case 653
                item = "toa"
            Case 654
                item = "tes"
            Case 655
                item = "ceh"
            Case 656
                item = "bet"
            Case 657
                item = "fed"
            Case 658
                item = "std"

            Case Else
                item = "nia"


        End Select

        Return item


    End Function

    Sub SearchRoutine()
        Searchtest() : Return

        'This builds the RefineSearchReferenceList array which is a copy of the SearchReferenceList which is deleted before search
        'its used for refined searches... It verifys the item being tested for a match already existed as a match in the last search thus refining the old matched results
        If Form1.RefineSearchCHECKBOX.Checked = True Then

            RefineSearchReferenceList.Clear()
            For Each item In SearchReferenceList
                RefineSearchReferenceList.Add(item)
            Next

        End If

        Dim MATCHEDUNIQUEATTRIBS = 0                           'Flag to add susscessful user attribs searches into quick edit pulldown
        Dim MATCHEDITEMNAME = 0                                 'FLAG USED TO ADD SUCCESSFUL ITEM NAME SEARCHES TO PULLDOWN
        Dim count = -1                                          'Used to track items database reference number
        Form1.SearchLISTBOX.Items.Clear()                       'Clear out old search matches
        SearchReferenceList.Clear()                             'Clear Out old Item Search Reference List (Holds location in database of each matched item)
        StringMatches.Clear() 'for unique attribs block searches
        IntegerMatches.Clear() 'for unique attribs block searches




        'SEARCH ROUTINE STARTS RIGHT HERE --->

        'SEARCH LIMITATIONS:
        'Exact match considers character casing as well as the string itself.   STRING MATCH MUST BE EXACT TO PASS
        'Not exact match ignores casing and redundant characters in the string. STRING MATCH MUST EXIST WITHIN THE TEST STRING REGARDLESS OF CASING TO PASS
        'Hide dupes only considers the items name when testing for duplicates
        'Refine search only considers items already matched (exists in the search match list already) when testing for matches

        '----------------------------------------------------------------------------------------------------------------------------------------------
        'DEBUGGING - CAN DELETE THIS BUT KEEP FOR NOW FOR DEBUGG (use as var tags pinned on references below for step by step bebugging)
        'Dim A = UCase(Form1.SearchFieldCOMBOBOX.Text)
        'Dim B = UCase(Form1.SearchWordCOMBOBOX.Text)
        'Dim C As String
        'Dim D As String
        '----------------------------------------------------------------------------------------------------------------------------------------------
        Dim StartSearch = 1
        For Each ItemObjectItem As ItemObjects In Objects       'Sequentually reference through all Object entries (iterates entire database)
            count = count + 1                                   ' move item locatino refrence counter onto next item number

            '----------------------------------------------------------------------------------------------------------------------------------------------
            'DEBUGGING - CAN DELETE THIS BUT KEEP FOR NOW FOR HANDY DEBUGG (use these here as pinned var tags in step by step bebugging)
            'A = UCase(Form1.SearchFieldCOMBOBOX.Text)
            'B = UCase(Form1.SearchWordCOMBOBOX.Text)
            'C = UCase(ItemObjectItem.ItemName)
            'D = UCase(Form1.SearchWordCOMBOBOX.Text).IndexOf(UCase(ItemObjectItem.ItemName))
            '----------------------------------------------------------------------------------------------------------------------------------------------

            'This Handles the Search Progress Bar Popup Window
            If StartSearch = 1 Then SearchProgressForm.Show() : SearchProgressForm.SearchPROGRESSBAR.Value = 0 : StartSearch = 0
            If SearchProgressForm.Enabled = False Then Exit For
            SearchProgressForm.SearchPROGRESSBAR.Value = Int((count / Form1.AllItemsInDatabaseListBox.Items.Count) * 100)
            SearchProgressForm.SearchProgressLABEL1.Text = Form1.SearchLISTBOX.Items.Count & " Matches"
            SearchProgressForm.SearchProgressLABEL2.Text = "Searching " & count & " of " & Form1.AllItemsInDatabaseListBox.Items.Count & " Item Records..."
            SearchProgressForm.Refresh()

            '****************************************************
            '*   S T R I N G      S E A R C H E S     O N L Y   *
            '****************************************************

            'Search for Word In item name (Equal to / Not Equal to)
            '------------------------------------------------------

            'NOT EXACT MATCH SEARCH FOR ITEM NAME
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ITEM NAME" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.ItemName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched '     EqualTo
                        If UCase(ItemObjectItem.ItemName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : MATCHEDITEMNAME = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.ItemName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                        If UCase(ItemObjectItem.ItemName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : MATCHEDITEMNAME = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched ' NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR ITEM NAME 
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemName) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched '     EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemName) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemName) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched 'NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemName) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDITEMNAME = 1 : GoTo ItemMatched 'NotEqualTo 
                    End If
                End If
            End If

            'Search For Word In ItemBase (EqualTo / NotEqualTo)
            '--------------------------------------------------

            'NOT EXACT MATCH SEARCH FOR ITEM BASE 
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ITEM BASE" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.ItemBase).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.ItemBase).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.ItemBase).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.ItemBase).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '   NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR ITEM BASE
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemBase) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemBase) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemBase) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemBase) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '  NotEqualTo 
                    End If
                End If
            End If

            'Search For Word In ItemQuality (EqualTo / NotEqualTo)
            '-----------------------------------------------------

            'NOT EXACT MATCH SEARCH FOR ITEM Quality 
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ITEM QUALITY" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.ItemQuality).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If UCase(ItemObjectItem.ItemQuality).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.ItemQuality).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                        If UCase(ItemObjectItem.ItemQuality).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR ITEM Quality
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemQuality) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.ItemQuality) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If
                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemQuality) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.ItemQuality) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    End If
                End If

            End If

            'Search For Word In AttackClass (EqualTo / NotEqualTo)
            '-----------------------------------------------------

            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ATTACK CLASS" Then

                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.AttackClass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                        If UCase(ItemObjectItem.AttackClass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                    End If
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.AttackClass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If UCase(ItemObjectItem.AttackClass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If
                End If

                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.AttackClass) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.AttackClass) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.AttackClass) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.AttackClass) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    End If
                End If
            End If

            'Search For Word In AttackSpeed (EqualTo / NotEqualTo)
            '-----------------------------------------------------

            'NOT EXACT MATCH SEARCH FOR ATTACK SPEED
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ATTACK SPEED" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.AttackSpeed).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If UCase(ItemObjectItem.AttackSpeed).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.AttackSpeed).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                        If UCase(ItemObjectItem.AttackSpeed).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR ATTACK SPEED
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.AttackSpeed) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.AttackSpeed) And Form1.RefineSearchCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.AttackSpeed) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.AttackSpeed) And Form1.RefineSearchCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    End If
                End If
            End If

            'Search For Word In MuleAccount (EqualTo / NotEqualTo)
            '-----------------------------------------------------

            'NOT EXACT MATCH SEARCH FOR MULE ACCOUNT NAME 
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "MULE ACCOUNT" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.MuleAccount).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If UCase(ItemObjectItem.MuleAccount).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.MuleAccount).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                        If UCase(ItemObjectItem.MuleAccount).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR MULE ACCOUNT NAME
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MuleAccount) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MuleAccount) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.MuleAccount) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.MuleAccount) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    End If
                End If

            End If

            'Search For Word In MuleName (EqualTo / NotEqualTo)
            '-----------------------------------------------------                                                                                                                                                                                               

            'NOT EXACT MATCH SEARCH FOR MULE NAME                                                                                                                                                                                                           
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "MULE NAME" Then

                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.MuleName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.MuleName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.MuleName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.MuleName).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '   NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR MULE NAME
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MuleName) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MuleName) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> UCase(ItemObjectItem.MuleName) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.MuleName) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '  NotEqualTo 
                    End If
                End If
            End If

            'Search For Word In MulePass (EqualTo / NotEqualTo)
            '--------------------------------------------------

            'NOT EXACT MATCH SEARCH FOR MULE NAME
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "MULE PASS" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.MulePass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.MulePass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.MulePass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.MulePass).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '   NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR MULE NAME
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MulePass) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.MulePass) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> UCase(ItemObjectItem.MulePass) And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.MulePass) And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '  NotEqualTo 
                    End If

                End If

            End If

            'Search For Word In UserReference (EqualTo / NotEqualTo)
            '-------------------------------------------------------

            'NOT EXACT MATCH SEARCH FOR USER REFERENCE
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "USER REFERENCE" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.UserReference).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If UCase(ItemObjectItem.UserReference).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.UserReference).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                        If UCase(ItemObjectItem.UserReference).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                    End If

                End If

                'EXACT MATCH SEARCH FOR USER REFERENCE
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If Form1.SearchWordCOMBOBOX.Text = ItemObjectItem.UserReference And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If Form1.SearchWordCOMBOBOX.Text = ItemObjectItem.UserReference And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If Form1.SearchWordCOMBOBOX.Text <> ItemObjectItem.UserReference And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                        If Form1.SearchWordCOMBOBOX.Text <> ItemObjectItem.UserReference And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    End If
                End If
            End If

            'Search For Word In UserReference (EqualTo / NotEqualTo)
            '-------------------------------------------------------

            'NOT EXACT MATCH SEARCH FOR USER REFERENCE
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "RUNEWORD" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.RuneWord).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If UCase(ItemObjectItem.RuneWord).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.RuneWord).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                        If UCase(ItemObjectItem.RuneWord).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' NotEqualTo 
                    End If
                End If

                'EXACT MATCH SEARCH FOR USER REFERENCE
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If Form1.SearchWordCOMBOBOX.Text = ItemObjectItem.RuneWord And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                        If Form1.SearchWordCOMBOBOX.Text = ItemObjectItem.RuneWord And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    End If

                    If Form1.SearchWordCOMBOBOX.Text <> ItemObjectItem.RuneWord And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Form1.SearchWordCOMBOBOX.Text <> ItemObjectItem.RuneWord And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 

                End If
            End If

            '*******************************************************
            '*   I N T E G E R      S E A R C H E S      O N L Y   *
            '*******************************************************

            'Exact match checkbox sets the search routine to ignore character case sets and any redundant text when testing matches. 
            'Therefore it really has no relevance when searching integers and so is ignored in the following integer search routines

            'Search For Integer In RequiredLevel (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------

            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "REQUIRED LEVEL" Then

                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.RequiredLevel) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                    If Val(ItemObjectItem.RequiredLevel) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.RequiredLevel) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                    If Val(ItemObjectItem.RequiredLevel) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.RequiredLevel) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                    If Val(ItemObjectItem.RequiredLevel) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.RequiredLevel) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                    If Val(ItemObjectItem.RequiredLevel) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In Required Strength (EqualTo / NotEqualTo / GreaterThan / LessThan)

            'Refine Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "REQUIRED STRENGTH" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.RequiredStrength) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                    If Val(ItemObjectItem.RequiredStrength) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.RequiredStrength) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                    If Val(ItemObjectItem.RequiredStrength) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.RequiredStrength) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                    If Val(ItemObjectItem.RequiredStrength) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.RequiredStrength) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                    If Val(ItemObjectItem.RequiredStrength) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In Required Dexterity (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------

            'Refine Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "REQUIRED DEXTERITY" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.RequiredDexterity) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                    If Val(ItemObjectItem.RequiredDexterity) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.RequiredDexterity) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                    If Val(ItemObjectItem.RequiredDexterity) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.RequiredDexterity) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                    If Val(ItemObjectItem.RequiredDexterity) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.RequiredDexterity) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                    If Val(ItemObjectItem.RequiredDexterity) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In One Hand Damage Max (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------------

            'Refined Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ONE HAND DAMAGE MAX" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.OneHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    If Val(ItemObjectItem.OneHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.OneHandDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Val(ItemObjectItem.OneHandDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.OneHandDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' GreaterThan
                    If Val(ItemObjectItem.OneHandDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.OneHandDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    LessThan 
                    If Val(ItemObjectItem.OneHandDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In One Hand Damage Max (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------------

            'Refined Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ONE HAND DAMAGE MIN" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.OneHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    If Val(ItemObjectItem.OneHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.OneHandDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Val(ItemObjectItem.OneHandDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.OneHandDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' GreaterThan
                    If Val(ItemObjectItem.OneHandDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.OneHandDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    LessThan 
                    If Val(ItemObjectItem.OneHandDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In Two Hand Damage Max (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------------

            'Refined Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "TWO HAND DAMAGE MAX" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.OneHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    If Val(ItemObjectItem.OneHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.OneHandDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Val(ItemObjectItem.OneHandDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.OneHandDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' GreaterThan
                    If Val(ItemObjectItem.OneHandDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.OneHandDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    LessThan 
                    If Val(ItemObjectItem.OneHandDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In Two Hand Damage Min (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------------

            'Refined Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "TWO HAND DAMAGE MIN" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.TwoHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    If Val(ItemObjectItem.TwoHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.TwoHandDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Val(ItemObjectItem.TwoHandDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.TwoHandDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' GreaterThan
                    If Val(ItemObjectItem.TwoHandDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.TwoHandDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    LessThan 
                    If Val(ItemObjectItem.TwoHandDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In Throw Damage Min (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------------

            'Refined Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "THROW DAMAGE MIN" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.ThrowDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    If Val(ItemObjectItem.TwoHandDamageMin) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.ThrowDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Val(ItemObjectItem.ThrowDamageMin) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.ThrowDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' GreaterThan
                    If Val(ItemObjectItem.ThrowDamageMin) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.ThrowDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    LessThan 
                    If Val(ItemObjectItem.ThrowDamageMin) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In Throw Damage Max (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------------

            'Refined Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "THROW DAMAGE MAX" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.ThrowDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    If Val(ItemObjectItem.TwoHandDamageMax) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.ThrowDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Val(ItemObjectItem.ThrowDamageMax) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.ThrowDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' GreaterThan
                    If Val(ItemObjectItem.ThrowDamageMax) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.ThrowDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    LessThan 
                    If Val(ItemObjectItem.ThrowDamageMax) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In Chance To Block (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------------

            'Refined Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "CHANCE TO BLOCK" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.ChanceToBlock) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    If Val(ItemObjectItem.ChanceToBlock) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.ChanceToBlock) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Val(ItemObjectItem.ChanceToBlock) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.ChanceToBlock) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' GreaterThan
                    If Val(ItemObjectItem.ChanceToBlock) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.ChanceToBlock) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    LessThan 
                    If Val(ItemObjectItem.ChanceToBlock) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            'Search For Integer In Defense (EqualTo / NotEqualTo / GreaterThan / LessThan)
            '-----------------------------------------------------------------------------------------

            'Refined Search
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "ITEM DEFENSE" Then
                If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                    If Val(ItemObjectItem.Defense) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     EqualTo
                    If Val(ItemObjectItem.Defense) = Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '         EqualTo
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                    If Val(ItemObjectItem.Defense) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched 'NotEqualTo 
                    If Val(ItemObjectItem.Defense) <> Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    NotEqualTo 
                End If
                If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
                    If Val(ItemObjectItem.Defense) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched ' GreaterThan
                    If Val(ItemObjectItem.Defense) > Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '     GreaterThan
                End If

                If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
                    If Val(ItemObjectItem.Defense) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '    LessThan 
                    If Val(ItemObjectItem.Defense) < Form1.SearchValueNUMERICUPDWN.Value And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : GoTo ItemMatched '        LessThan 
                End If
            End If

            '*****************************************************************************
            '*  S E A R C H    U N I Q U E     A T T R I B S   S T R I N G    M A T C H  *
            '*****************************************************************************

            'not exact
            If UCase(Form1.SearchFieldCOMBOBOX.Text) = "UNIQUE ATTRIBUTES" Then
                If Form1.ExactMatchCHECKBOX.Checked = False Then

                    'Refined Searches Not Exact
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If UCase(ItemObjectItem.Stat1).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat2).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat3).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat4).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat5).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat6).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat7).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat8).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat9).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat10).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat11).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat12).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat13).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat14).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat15).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo

                        If UCase(ItemObjectItem.Stat1).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat2).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat3).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat4).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat5).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat6).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat7).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat8).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat9).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat10).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat11).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat12).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat13).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat14).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If UCase(ItemObjectItem.Stat15).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) > -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If UCase(ItemObjectItem.Stat1).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat2).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat3).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat4).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat5).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat6).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat7).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat8).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat9).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat10).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat11).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat12).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat13).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat14).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat15).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = True And RefineSearchReferenceList.Contains(count) = True Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 

                        If UCase(ItemObjectItem.Stat1).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat2).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat3).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat4).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat5).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat6).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat7).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat8).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat9).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat10).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat11).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat12).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat13).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat14).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                        If UCase(ItemObjectItem.Stat15).IndexOf(UCase(Form1.SearchWordCOMBOBOX.Text)) = -1 And Form1.RefineSearchCHECKBOX.Checked = False Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '   NotEqualTo 
                    End If
                End If

                'EXACT
                If Form1.ExactMatchCHECKBOX.Checked = True Then

                    'Refined Searches Exact
                    If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat1) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat2) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat3) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat4) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat5) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat6) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat7) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat8) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat9) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat10) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat11) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat12) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat13) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat14) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat15) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo

                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat1) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat2) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat3) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat4) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat5) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat6) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat7) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat8) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat9) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat10) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat11) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat12) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat13) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat14) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                        If (Form1.SearchWordCOMBOBOX.Text) = (ItemObjectItem.Stat15) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '       EqualTo
                    End If

                    If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat1) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat2) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat3) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat4) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat5) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat6) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat7) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat8) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat9) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat10) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat11) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat12) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat13) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat14) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat15) And RefineSearchReferenceList.Contains(count) = True Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 

                        'New Searches Exact

                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat1) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat2) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat3) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat4) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat5) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat6) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat7) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat8) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat9) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat10) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat11) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat12) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat13) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat14) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                        If (Form1.SearchWordCOMBOBOX.Text) <> (ItemObjectItem.Stat15) Then If Form1.HideDuplicatesCHECKBOX.Checked = True And Form1.SearchLISTBOX.Items.Contains(ItemObjectItem.ItemName) = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched Else If Form1.HideDuplicatesCHECKBOX.Checked = False Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : MATCHEDUNIQUEATTRIBS = 1 : GoTo ItemMatched '  NotEqualTo 
                    End If
                End If

            End If







ItemMatched:  ' Jump point to avoid redundant routine once a match has been found (also a potential time saver for searching much larger dbases) 




            '   Next bit puts successfull item name search words into the searchword quickedit pulldown and a array collection list for reference 
            '   when repopulating dropdown on Item Name reselection. 

            '   The Flag Var MATCHEDITEMNAME = 1 or 0 is assigned in above search name routine only. And again only if a match is found. 
            '   NOTE: This is ONLY used for item name searches (as it would be pointless to put the entire database item name list into the Item Name quick edit dropdown)

            If MATCHEDITEMNAME = 1 And Form1.SearchWordCOMBOBOX.Items.Contains(Form1.SearchWordCOMBOBOX.Text) = False Then Form1.SearchWordCOMBOBOX.Items.Add(Form1.SearchWordCOMBOBOX.Text) : ItemNamePulldownList.Add(Form1.SearchWordCOMBOBOX.Text) : MATCHEDITEMNAME = 0

            If MATCHEDUNIQUEATTRIBS = 1 And Form1.SearchWordCOMBOBOX.Items.Contains(Form1.SearchWordCOMBOBOX.Text) = False Then Form1.SearchWordCOMBOBOX.Items.Add(Form1.SearchWordCOMBOBOX.Text) : UniqueAttribsPulldownList.Add(Form1.SearchWordCOMBOBOX.Text) : MATCHEDUNIQUEATTRIBS = 0



        Next ' Search all itmes loop point (for next)

        SearchProgressForm.Close()




        'if search was successful and matches exist this selects the first match in the searchlist and focuses on the search list tab page to show matches
        If Form1.SearchLISTBOX.Items.Count > 0 Then
            Form1.SearchLISTBOX.SelectedIndex = 0
            Form1.ListboxTABCONTROL.SelectTab(1)
            Form1.Button1.BackColor = Color.DimGray
            Form1.Button2.BackColor = Color.Black
            Form1.TextBox2.Text = Form1.SearchLISTBOX.Items.Count & " - Total Matches"
        Else
            'if there are no matches then show the all item list 

            If Form1.RefineSearchCHECKBOX.Checked = True Then
                For Each item In RefineSearchReferenceList
                    Form1.SearchLISTBOX.Items.Add(Objects(item).ItemName)
                Next

            Else

                Form1.TextBox2.Text = Form1.AllItemsInDatabaseListBox.Items.Count & " - Total Items"
                Form1.ListboxTABCONTROL.SelectTab(0)
                Form1.Button1.BackColor = Color.Black
                Form1.Button2.BackColor = Color.DimGray
            End If

        End If
        Form1.SearchBUTTON.Enabled = True 'Finally re-enable the search button and were done, sooo lets get the hell outta here... phew!
    End Sub



    'this backsup the current database from menu selection and when closing app
    Sub BackupDatabase()
        Dim BackupPath = Application.StartupPath & "\Database\Backup\"
        Dim temp As String = ""
        Dim myarray = Split(DatabaseFile, ".txt", 0)
        Dim tempname = myarray(0) & ".bak"
        myarray = Split(tempname, "\")
        tempname = myarray(myarray.Length - 1)

        If My.Computer.FileSystem.FileExists(BackupPath & tempname) = True Then
            My.Computer.FileSystem.DeleteFile(BackupPath & tempname)
        End If
        My.Computer.FileSystem.CopyFile(DataBasePath & DatabaseFile, BackupPath & tempname)

    End Sub
    Sub Searchtest()
        Form1.SearchLISTBOX.Items.Clear()
        SearchReferenceList.Clear()
        Dim temp As String = ""
        Dim MyValue = -1

        If Form1.SearchFieldCOMBOBOX.Text = "Unique Attributes" Then

            For count = 0 To Objects.Count - 1

                If Objects(count).Stat1.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then
                    MyValue = getvalue(Objects(count).Stat1)
                    MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
                End If

                If Objects(count).Stat2.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then
                    MyValue = getvalue(Objects(count).Stat2)
                    MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
                End If

                If Objects(count).Stat3.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then
                    MyValue = getvalue(Objects(count).Stat3)
                    MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
                End If

                If Objects(count).Stat4.IndexOf(Form1.SearchWordCOMBOBOX.Text) > -1 Then
                    MyValue = getvalue(Objects(count).Stat4)
                    MyDecipher(count, MyValue) : Continue For 'skip rest we have found it
                End If

            Next



        End If


        If SearchReferenceList.Count > 0 Then
            Form1.SearchLISTBOX.SelectedIndex = 0
            Form1.ListboxTABCONTROL.SelectTab(1)
            Form1.Button1.BackColor = Color.DimGray
            Form1.Button2.BackColor = Color.Black
            Form1.TextBox2.Text = Form1.SearchLISTBOX.Items.Count & " - Total Matches"
        End If


    End Sub
    Sub MyDecipher(ByVal count, ByVal xval)

        If Form1.SearchOperatorCOMBOBOX.Text = "Equal To" Then
            If xval = Form1.SearchValueNUMERICUPDWN.Value Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
        End If
        If Form1.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
            If xval > Form1.SearchValueNUMERICUPDWN.Value Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
        End If
        If Form1.SearchOperatorCOMBOBOX.Text = "Less Than" Then
            If xval < Form1.SearchValueNUMERICUPDWN.Value Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
        End If
        If Form1.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
            If xval <> Form1.SearchValueNUMERICUPDWN.Value Then Form1.SearchLISTBOX.Items.Add(Objects(count).ItemName) : SearchReferenceList.Add(count) : Return
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

End Module
