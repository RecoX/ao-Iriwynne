Attribute VB_Name = "Protocol"
'*******  License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Mart�n Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Mart�n Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit


'Build
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetVolumeSerialNumber Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
                                                                                                                               lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
                                                                                               lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont

        red As Byte
        green As Byte
        blue As Byte
        Bold As Boolean
        Italic As Boolean

End Type

Private Enum ServerPacketID

        logged = 1                ' LOGGED
        RemoveDialogs = 2         ' QTDL
        RemoveCharDialog = 3      ' QDL
        NavigateToggle = 4        ' NAVEG
        Disconnect = 5            ' FINOK
        CommerceEnd = 6           ' FINCOMOK
        BankEnd = 7               ' FINBANOK
        CommerceInit = 8          ' INITCOM
        BankInit = 9              ' INITBANCO
        UserCommerceInit = 10      ' INITCOMUSU
        UserCommerceEnd = 11       ' FINCOMUSUOK
        UserOfferConfirm = 12
        CommerceChat = 13
        ShowBlacksmithForm = 14    ' SFH
        ShowCarpenterForm = 15     ' SFC
        UpdateSta = 16             ' ASS
        UpdateMana = 17            ' ASM
        UpdateHP = 18              ' ASH
        UpdateGold = 19            ' ASG
        UpdateBankGold = 20
        UpdateExp = 21             ' ASE
        ChangeMap = 22             ' CM
        PosUpdate = 23             ' PU
        ChatOverHead = 24          ' ||
        ConsoleMsg = 25            ' || - Beware!! its the same as above, but it was properly splitted
        GuildChat = 26             ' |+
        ShowMessageBox = 27        ' !!
        UserIndexInServer = 28     ' IU
        UserCharIndexInServer = 29 ' IP
        CharacterCreate = 30       ' CC
        CharacterRemove = 31       ' BP
        CharacterChangeNick = 32
        CharacterMove = 33         ' MP, +, * and _ '
        CharacterAttackMovement = 34
        ForceCharMove = 35
        CharacterChange = 36       ' CP
        ObjectCreate = 37          ' HO
        ObjectDelete = 38          ' BO
        BlockPosition = 39         ' BQ
        PlayMIDI = 40              ' TM
        PlayWave = 41              ' TW
        guildList = 42             ' GL
        AreaChanged = 43           ' CA
        PauseToggle = 44           ' BKW
        RainToggle = 45            ' LLU
        CreateFX = 46              ' CFX
        UpdateUserStats = 47       ' EST
        WorkRequestTarget = 48     ' T01
        ChangeInventorySlot = 49   ' CSI
        ChangeBankSlot = 50        ' SBO
        ChangeSpellSlot = 51       ' SHS
        Atributes = 52            ' ATR
        BlacksmithWeapons = 53     ' LAH
        BlacksmithArmors = 54      ' LAR
        CarpenterObjects = 55      ' OBR
        RestOK = 56                ' DOK
        ErrorMsg = 57              ' ERR
        Blind = 58                 ' CEGU
        Dumb = 59                  ' DUMB
        ShowSignal = 60            ' MCAR
        ChangeNPCInventorySlot = 61 ' NPCI
        UpdateHungerAndThirst = 62 ' EHYS
        Fame = 63                  ' FAMA
        MiniStats = 64             ' MEST
        LevelUp = 65               ' SUNI
        AddForumMsg = 66           ' FMSG
        ShowForumForm = 67         ' MFOR
        SetInvisible = 68          ' NOVER
        ' @@ PAQUETE 69 SIN USAR
        MeditateToggle = 70        ' MEDOK
        BlindNoMore = 71           ' NSEGUE
        DumbNoMore = 72            ' NESTUP
        SendSkills = 73            ' SKILLS
        TrainerCreatureList = 74   ' LSTCRI
        guildNews = 75             ' GUILDNE
        OfferDetails = 76          ' PEACEDE & ALLIEDE
        AlianceProposalsList = 77  ' ALLIEPR
        PeaceProposalsList = 78    ' PEACEPR
        CharacterInfo = 79         ' CHRINFO
        GuildLeaderInfo = 80       ' LEADERI
        GuildMemberInfo = 81
        GuildDetails = 82          ' CLANDET
        ShowGuildFundationForm = 83 ' SHOWFUN
        ParalizeOK = 84           ' PARADOK
        ShowUserRequest = 85       ' PETICIO
        TradeOK = 86               ' TRANSOK
        BankOK = 87                ' BANCOOK
        ChangeUserTradeSlot = 88   ' COMUSUINV
        SendNight = 89             ' NOC
        Pong = 90
        UpdateTagAndStatus = 91
    
        'GM messages
        SpawnList = 92             ' SPL
        ShowSOSForm = 93           ' MSOS
        ShowMOTDEditionForm = 94   ' ZMOTD
        ShowGMPanelForm = 95       ' ABPANEL
        UserNameList = 96          ' LISTUSU
        ShowDenounces = 97
        RecordList = 98
        RecordDetails = 99
    
        ShowGuildAlign = 100
        ShowPartyForm = 101
        UpdateStrenghtAndDexterity = 102
        UpdateStrenght = 103
        UpdateDexterity = 104
        ' @@ Paquete 105 Libre
        MultiMessage = 106
        StopWorking = 107
        CancelOfferItem = 108
        StrDextRunningOut = 109
        CharacterUpdateHP = 110
        CreateDamage = 111
        Canje = 112                 'Canjes
        CanjePTS = 113             'Canjes
        ControlUserRecive = 114
        ControlUserShow = 115
        SendScreen = 116 ' @@ Foto
        MercadoList = 118
        MontateToggle = 120
        RecibirRanking = 121
        NameList = 122
        ConsoleMsgNew = 123
        AntiCL = 124
        ContadorInvi = 125
        RecibiOns = 126
        ServerPideCPU = 127
        ClanPideAyuda = 128
        QuestDetails = 129          ' GSZAO
    QuestListSend = 130         ' GSZAO
End Enum

Private Enum ClientPacketID

        LoginExistingChar = 1     'OLOGIN
        ResetChar = 2
        LoginNewChar = 3          'NLOGIN
        Talk = 4                  ';
        Yell = 5                  '-
        Whisper = 6               '\
        Walk = 7                  'M
        RequestPositionUpdate = 8 'RPU
        Attack = 9                'AT
        PickUp = 10                'AG
        SafeToggle = 11            '/SEG & SEG  (SEG's behaviour has to be coded in the client)
        ResuscitationSafeToggle = 12
        RequestGuildLeaderInfo = 13 'GLINFO
        RequestAtributes = 14      'ATR
        RequestFame = 15           'FAMA
        RequestSkills = 16         'ESKI
        RequestMiniStats = 17      'FEST
        CommerceEnd = 18           'FINCOM
        UserCommerceEnd = 19       'FINCOMUSU
        UserCommerceConfirm = 20
        CommerceChat = 21
        BankEnd = 22               'FINBAN
        UserCommerceOk = 23        'COMUSUOK
        UserCommerceReject = 24    'COMUSUNO
        Drop = 25                  'TI
        CastSpell = 26             'LH
        LeftClick = 27             'LC
        DoubleClick = 28           'RC
        Work = 29                  'UK
        UseSpellMacro = 30         'UMH
        UseItem = 31               'USA
        CraftBlacksmith = 32       'CNS
        CraftCarpenter = 33        'CNC
        WorkLeftClick = 34         'WLC
        CreateNewGuild = 35        'CIG
        SpellInfo = 36             'INFS
        EquipItem = 37             'EQUI
        ChangeHeading = 38         'CHEA
        ModifySkills = 39          'SKSE
        Train = 40                 'ENTR
        CommerceBuy = 41           'COMP
        BankExtractItem = 42       'RETI
        CommerceSell = 43          'VEND
        BankDeposit = 44           'DEPO
        ForumPost = 45             'DEMSG
        MoveSpell = 46             'DESPHE
        MoveBank = 47
        ClanCodexUpdate = 48       'DESCOD
        UserCommerceOffer = 49     'OFRECER
        GuildAcceptPeace = 50      'ACEPPEAT
        GuildRejectAlliance = 51  'RECPALIA
        GuildRejectPeace = 52      'RECPPEAT
        GuildAcceptAlliance = 53   'ACEPALIA
        GuildOfferPeace = 54       'PEACEOFF
        GuildOfferAlliance = 55    'ALLIEOFF
        GuildAllianceDetails = 56  'ALLIEDET
        GuildPeaceDetails = 57     'PEACEDET
        GuildRequestJoinerInfo = 58 'ENVCOMEN
        GuildAlliancePropList = 59 'ENVALPRO
        GuildPeacePropList = 60    'ENVPROPP
        GuildDeclareWar = 61       'DECGUERR
        GuildNewWebsite = 62       'NEWWEBSI
        GuildAcceptNewMember = 63  'ACEPTARI
        GuildRejectNewMember = 64  'RECHAZAR
        GuildKickMember = 65       'ECHARCLA
        GuildUpdateNews = 66       'ACTGNEWS
        GuildMemberInfo = 67       '1HRINFO<
        GuildOpenElections = 68    'ABREELEC
        GuildRequestMembership = 69 'SOLICITUD
        GuildRequestDetails = 70   'CLANDETAILS
        Online = 71                '/ONLINE
        Quit = 72                  '/SALIR
        GuildLeave = 73            '/SALIRCLAN
        RequestAccountState = 74   '/BALANCE
        PetStand = 75              '/QUIETO
        PetFollow = 76             '/ACOMPA�AR
        ReleasePet = 77            '/LIBERAR
        TrainList = 78             '/ENTRENAR
        Rest = 79                  '/DESCANSAR
        Meditate = 80              '/MEDITAR
        Resucitate = 81            '/RESUCITAR
        Heal = 82                  '/CURAR
        Help = 83                  '/AYUDA
        RequestStats = 84          '/EST
        CommerceStart = 85         '/COMERCIAR
        BankStart = 86             '/BOVEDA
        Enlist = 87                '/ENLISTAR
        Information = 88           '/INFORMACION
        Reward = 89                '/RECOMPENSA
        RequestMOTD = 90           '/MOTD
        Uptime = 91                '/UPTIME
        PartyLeave = 92            '/SALIRPARTY
        PartyCreate = 93           '/CREARPARTY
        PartyJoin = 94             '/PARTY
        Inquiry = 95               '/ENCUESTA ( with no params )
        GuildMessage = 96          '/CMSG
        PartyMessage = 97          '/PMSG
        CentinelReport = 98        '/CENTINELA
        GuildOnline = 99           '/ONLINECLAN
        PartyOnline = 100           '/ONLINEPARTY
        CouncilMessage = 101        '/BMSG
        RoleMasterRequest = 102     '/ROL
        GMRequest = 103             '/GM
        bugReport = 104             '/_BUG
        ChangeDescription = 105     '/DESC
        GuildVote = 106             '/VOTO
        Punishments = 107           '/PENAS
        ChangePassword = 108        '/CONTRASE�A
        Gamble = 109                '/APOSTAR
        InquiryVote = 110           '/ENCUESTA ( with parameters )
        LeaveFaction = 111          '/RETIRAR ( with no arguments )
        BankExtractGold = 112       '/RETIRAR ( with arguments )
        BankDepositGold = 113       '/DEPOSITAR
        Denounce = 114              '/DENUNCIAR
        GuildFundate = 115          '/FUNDARCLAN
        GuildFundation = 116
        PartyKick = 117             '/ECHARPARTY
        PartySetLeader = 118        '/PARTYLIDER
        PartyAcceptMember = 119     '/ACCEPTPARTY
        Ping = 120                  '/PING
    
        RequestPartyForm = 121
        ItemUpgrade = 122
        GMCommands = 123
        InitCrafting = 124
        Home = 125
        ShowGuildNews = 126
        ShareNpc = 127              '/COMPARTIRNPC
        StopSharingNpc = 128        '/NOCOMPARTIRNPC
        Consultation = 129
        MoveItem = 130              'Drag and drop
        PMDeleteList = 131
        PMList = 132
        otherSendReto = 133 ' @@ 1 vs 1
        SendReto = 134      ' @@ 2 vs 2
        AcceptReto = 135  ' @@ Aceptar 1.1 | 2.2
        DropObjTo = 136
        SetMenu = 137
        Canjear = 138               'Mab
        Canjesx = 139                'kidral
        ChangeCara = 140
        ControlUserRequest = 141
        ControlUserSendData = 142
        RequestScreen = 143
        regresar = 144
        Retos = 146
        Poder = 147
        SolicitarRanking = 148
        Castillo = 149
        StartList = 150
        AccessList = 151
        CancelList = 152
        ActivarDeath = 153          '/ACDEATH CUPOS@CAENOBJS
        IngresarDeath = 154         '/DEATH
        torneo = 155
        SetVIP = 156
        Fianza = 157
        MandoCPUID = 158
        JDHParticipar = 159
        JDHCrear = 160
        JDHCancelar = 161
        Quest = 162                 ' GSZAO - /QUEST
    QuestAccept = 163           ' GSZAO
    QuestListRequest = 164      ' GSZAO
    QuestDetailsRequest = 165   ' GSZAO
    QuestAbandon = 166          ' GSZAO
End Enum
 
Private Enum TorneoPacketID
    ATorneo                 'Hacer nuevo torneo
    CTorneo                 'Cancelar nuevo torneo
    PTorneo                 '
End Enum

Public Enum eGMCommands

        GMMessage = 1           '/GMSG
        showName = 2              '/SHOWNAME
        OnlineRoyalArmy = 3       '/ONLINEREAL
        OnlineChaosLegion = 4     '/ONLINECAOS
        GoNearby = 5              '/IRCERCA
        Comment = 6               '/REM
        serverTime = 7            '/HORA
        Where = 8                 '/DONDE
        CreaturesInMap = 9        '/NENE
        WarpMeToTarget = 10        '/TELEPLOC
        WarpChar = 11              '/TELEP
        Silence = 12               '/SILENCIAR
        SOSShowList = 13           '/SHOW SOS
        SOSRemove = 14             'SOSDONE
        GoToChar = 15              '/IRA
        invisible = 16             '/INVISIBLE
        GMPanel = 17               '/PANELGM
        RequestUserList = 18       'LISTUSU
        Working = 19               '/TRABAJANDO
        Hiding = 20                '/OCULTANDO
        Jail = 21                  '/CARCEL
        KillNPC = 22               '/RMATA
        WarnUser = 23              '/ADVERTENCIA
        EditChar = 24              '/MOD
        RequestCharInfo = 25       '/INFO
        RequestCharStats = 26      '/STAT
        RequestCharGold = 27       '/BAL
        RequestCharInventory = 28  '/INV
        RequestCharBank = 29       '/BOV
        RequestCharSkills = 30     '/SKILLS
        ReviveChar = 31            '/REVIVIR
        OnlineGM = 32              '/ONLINEGM
        OnlineMap = 33             '/ONLINEMAP
        Forgive = 34               '/PERDON
        Kick = 35                  '/ECHAR
        Execute = 36               '/EJECUTAR
        banChar = 37               '/BAN
        UnbanChar = 38             '/UNBAN
        NPCFollow = 39             '/SEGUIR
        SummonChar = 41            '/SUM
        SpawnListRequest = 42      '/CC
        SpawnCreature = 43         'SPA
        ResetNPCInventory = 44     '/RESETINV
        CleanWorld = 45            '/LIMPIAR
        ServerMessage = 46         '/RMSG
        nickToIP = 47              '/NICK2IP
        IPToNick = 48              '/IP2NICK
        GuildOnlineMembers = 49    '/ONCLAN
        TeleportCreate = 50        '/CT
        TeleportDestroy = 51       '/DT
        RainToggle = 52            '/LLUVIA
        SetCharDescription = 53    '/SETDESC
        ForceMIDIToMap = 54        '/FORCEMIDIMAP
        ForceWAVEToMap = 55        '/FORCEWAVMAP
        RoyalArmyMessage = 56      '/REALMSG
        ChaosLegionMessage = 57    '/CAOSMSG
        CitizenMessage = 58        '/CIUMSG
        CriminalMessage = 59       '/CRIMSG
        TalkAsNPC = 60             '/TALKAS
        DestroyAllItemsInArea = 61 '/MASSDEST
        AcceptRoyalCouncilMember = 62 '/ACEPTCONSE
        AcceptChaosCouncilMember = 63 '/ACEPTCONSECAOS
        ItemsInTheFloor = 64       '/PISO
        MakeDumb = 65              '/ESTUPIDO
        MakeDumbNoMore = 66        '/NOESTUPIDO
        dumpIPTables = 67          '/DUMPSECURITY
        CouncilKick = 68           '/KICKCONSE
        SetTrigger = 69            '/TRIGGER
        AskTrigger = 70            '/TRIGGER with no args
        BannedIPList = 71          '/BANIPLIST
        BannedIPReload = 72        '/BANIPRELOAD
        GuildMemberList = 73       '/MIEMBROSCLAN
        GuildBan = 74              '/BANCLAN
        BanIP = 75                 '/BANIP
        UnbanIP = 76               '/UNBANIP
        CreateItem = 77            '/CI
        DestroyItems = 78          '/DEST
        ChaosLegionKick = 79       '/NOCAOS
        RoyalArmyKick = 80         '/NOREAL
        ForceMIDIAll = 81          '/FORCEMIDI
        ForceWAVEAll = 82          '/FORCEWAV
        RemovePunishment = 83      '/BORRARPENA
        TileBlockedToggle = 84     '/BLOQ
        KillNPCNoRespawn = 85      '/MATA
        KillAllNearbyNPCs = 86     '/MASSKILL
        LastIP = 87                '/LASTIP
        ChangeMOTD = 88            '/MOTDCAMBIA
        SetMOTD = 89               'ZMOTD
        SystemMessage = 90         '/SMSG
        CreateNPC = 91             '/ACC
        CreateNPCWithRespawn = 92  '/RACC
        ImperialArmour = 93        '/AI1 - 4
        ChaosArmour = 94           '/AC1 - 4
        NavigateToggle = 95        '/NAVE
        ServerOpenToUsersToggle = 96 '/HABILITAR
        TurnOffServer = 97         '/APAGAR
        TurnCriminal = 98          '/CONDEN
        ResetFactions = 99         '/RAJAR
        RemoveCharFromGuild = 100   '/RAJARCLAN
        RequestCharMail = 101       '/LASTEMAIL
        AlterPassword = 102         '/APASS
        AlterMail = 103             '/AEMAIL
        AlterName = 104             '/ANAME
        ToggleCentinelActivated = 105 '/CENTINELAACTIVADO
        DoBackUp = 106              '/DOBACKUP
        ShowGuildMessages = 107     '/SHOWCMSG
        SaveMap = 108               '/GUARDAMAPA
        ChangeMapInfoPK = 109       '/MODMAPINFO PK
        ChangeMapInfoBackup = 110   '/MODMAPINFO BACKUP
        ChangeMapInfoRestricted = 111 '/MODMAPINFO RESTRINGIR
        ChangeMapInfoNoMagic = 112  '/MODMAPINFO MAGIASINEFECTO
        ChangeMapInfoNoInvi = 113   '/MODMAPINFO INVISINEFECTO
        ChangeMapInfoNoResu = 114   '/MODMAPINFO RESUSINEFECTO
        ChangeMapInfoLand = 115     '/MODMAPINFO TERRENO
        ChangeMapInfoZone = 116     '/MODMAPINFO ZONA
        ChangeMapInfoStealNpc = 117 '/MODMAPINFO ROBONPCm
        ChangeMapInfoNoOcultar = 118 '/MODMAPINFO OCULTARSINEFECTO
        ChangeMapInfoNoInvocar = 119 '/MODMAPINFO INVOCARSINEFECTO
        SaveChars = 120             '/GRABAR
        CleanSOS = 121              '/BORRAR SOS
        ShowServerForm = 122        '/SHOW INT
        night = 123                 '/NOCHE
        KickAllChars = 124          '/ECHARTODOSPJS
        ReloadNPCs = 125            '/RELOADNPCS
        ReloadServerIni = 126       '/RELOADSINI
        ReloadSpells = 127          '/RELOADHECHIZOS
        ReloadObjects = 128         '/RELOADOBJ
        Restart = 129               '/REINICIAR
        ResetAutoUpdate = 130       '/AUTOUPDATE
        ChatColor = 131             '/CHATCOLOR
        Ignored = 132               '/IGNORADO
        CheckSlot = 133             '/SLOT
        SetIniVar = 134             '/SETINIVAR LLAVE CLAVE VALOR
        CreatePretorianClan = 135   '/CREARPRETORIANOS
        RemovePretorianClan = 136   '/ELIMINARPRETORIANOS
        EnableDenounces = 137       '/DENUNCIAS
        ShowDenouncesList = 138     '/SHOW DENUNCIAS
        MapMessage = 139            '/MAPMSG
        SetDialog = 140             '/SETDIALOG
        Impersonate = 141           '/IMPERSONAR
        Imitate = 142               '/MIMETIZAR
        RecordAdd = 143
        RecordRemove = 144
        RecordAddObs = 145
        RecordListRequest = 146
        RecordDetailsRequest = 147

        PMSend = 148
        PMDeleteUser = 149
        PMListUser = 150
        
        SetPuntosShop = 151
        Countdown = 152
        VerHD = 153
        BanHD = 154
        UnBanHD = 155
        CheckCPU_ID = 156               '/VERCPUID USUARIO
        BanT0 = 157                     '/BANT0 USUARIO
        UnBanT0 = 158                   '/UNBANT0 NROHD
End Enum

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_GRANPODER
    FONTTYPE_TORNEO
    FONTTYPE_APU�ALADO
    FONTTYPE_NARANJA
    FONTTYPE_VIP
    FONTTYPE_VERDE
    FONTTYPE_BORDO
    FONTTYPE_MARRON
    FONTTYPE_AMARILLO
    FONTTYPE_BLANCO
    FONTTYPE_VIOLETA
End Enum

Public FontTypes(31) As tFont

''
' Initializes the fonts array

Public Sub InitFonts()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                .red = 255
                .green = 255
                .blue = 255
                
        End With
        
        With FontTypes(FontTypeNames.FONTTYPE_VIOLETA)
                .red = 213
                .green = 14
                .blue = 241
                .Bold = True
        End With
        
          With FontTypes(FontTypeNames.FONTTYPE_VIP)
                .red = 230
                .green = 230
                .blue = 1
                .Bold = True
        End With
        
        With FontTypes(FontTypeNames.FONTTYPE_AMARILLO)
                .red = 245
                .green = 245
                .blue = 1
                .Bold = True
        End With
        
        With FontTypes(FontTypeNames.FONTTYPE_BLANCO)
                .red = 255
                .green = 255
                .blue = 255
                .Bold = True
        End With
        
         With FontTypes(FontTypeNames.FONTTYPE_MARRON)
                .red = 183
                .green = 117
                .blue = 81
                .Bold = True
        End With
        
        With FontTypes(FontTypeNames.FONTTYPE_VERDE)
                .red = 1
                .green = 255
                .blue = 1
                .Bold = True
        End With
        
        With FontTypes(FontTypeNames.FONTTYPE_BORDO)
                .red = 95
                .green = 2
                .blue = 31
                .Bold = True
        End With
        
          With FontTypes(FontTypeNames.FONTTYPE_NARANJA)
                .red = 255
                .green = 128
                .blue = 0
                .Bold = True
        End With
        
        With FontTypes(FontTypeNames.FONTTYPE_APU�ALADO)
                .red = 223
                .green = 57
                .blue = 23
                .Bold = True
        End With
        
         With FontTypes(FontTypeNames.FONTTYPE_TORNEO)
                .red = 10
                .green = 255
                .blue = 10
                .Bold = True
        End With
        
        With FontTypes(FontTypeNames.FONTTYPE_GRANPODER)
                .red = 192
                .green = 192
                .blue = 192
                .Bold = True
        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
                .red = 255
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                .red = 32
                .green = 51
                .blue = 223
                .Bold = 1
                .Italic = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                .red = 65
                .green = 190
                .blue = 156

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
                .red = 65
                .green = 190
                .blue = 156
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
                .red = 130
                .green = 130
                .blue = 130
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_PARTY)
                .red = 255
                .green = 180
                .blue = 250

        End With
    
        FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255
    
        With FontTypes(FontTypeNames.FONTTYPE_GUILD)
                .red = 255
                .green = 255
                .blue = 255
                .Bold = 1

        End With
    
        FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185
    
        With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                .red = 228
                .green = 199
                .blue = 27

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
                .red = 130
                .green = 130
                .blue = 255
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
                .red = 255
                .green = 60
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
                .green = 200
                .blue = 255
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
                .red = 255
                .green = 50
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
                .green = 255
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
                .red = 255
                .green = 255
                .blue = 255
                .Italic = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_GM)
                .red = 30
                .green = 255
                .blue = 30
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
                .blue = 200
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_CONSE)
                .red = 30
                .green = 150
                .blue = 30
                .Bold = 1

        End With
    
        With FontTypes(FontTypeNames.FONTTYPE_DIOS)
                .red = 250
                .green = 250
                .blue = 150
                .Bold = 1

        End With

End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        On Error Resume Next

        Dim Packet As Byte

        Packet = incomingData.PeekByte()
        'Debug.Print Packet
    
        Select Case Packet
            
                Case ServerPacketID.AntiCL
                    Call HandleAntiCL
                    
                Case ServerPacketID.QuestDetails     ' GSZAO
239                 Call HandleQuestDetails

240             Case ServerPacketID.QuestListSend       ' GSZAO
241                 Call HandleQuestListSend

                Case ServerPacketID.ClanPideAyuda
                    Call HandleHelp
                    
                Case ServerPacketID.ServerPideCPU
                    Call HandleAntiWPE
                    
                Case ServerPacketID.RecibiOns
                    Call HandleRecibiOns
                    
                Case ServerPacketID.ContadorInvi
                    Call HandleContadorInvi
                    
                Case ServerPacketID.ConsoleMsgNew
                    Call HandleConsoleMessageNew
                    
                Case ServerPacketID.RecibirRanking
                    Call HandleRecibirRanking
                    
                Case ServerPacketID.NameList
                    Call HandleNameList
            
                Case ServerPacketID.logged                  ' LOGGED
                        Call HandleLogged
        
                Case ServerPacketID.RemoveDialogs           ' QTDL
                        Call HandleRemoveDialogs
        
                Case ServerPacketID.RemoveCharDialog        ' QDL
                        Call HandleRemoveCharDialog
        
                Case ServerPacketID.NavigateToggle          ' NAVEG
                        Call HandleNavigateToggle
        
                Case ServerPacketID.Disconnect              ' FINOK
                        Call HandleDisconnect
        
                Case ServerPacketID.CommerceEnd             ' FINCOMOK
                        Call HandleCommerceEnd
            
                Case ServerPacketID.CommerceChat
                        Call HandleCommerceChat
        
                Case ServerPacketID.BankEnd                 ' FINBANOK
                        Call HandleBankEnd
        
                Case ServerPacketID.CommerceInit            ' INITCOM
                        Call HandleCommerceInit
        
                Case ServerPacketID.BankInit                ' INITBANCO
                        Call HandleBankInit
        
                Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
                        Call HandleUserCommerceInit
        
                Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
                        Call HandleUserCommerceEnd
            
                Case ServerPacketID.UserOfferConfirm
                        Call HandleUserOfferConfirm
        
                Case ServerPacketID.ShowBlacksmithForm      ' SFH
                        Call HandleShowBlacksmithForm
        
                Case ServerPacketID.ShowCarpenterForm       ' SFC
                        Call HandleShowCarpenterForm
        
                Case ServerPacketID.UpdateSta               ' ASS
                        Call HandleUpdateSta
        
                Case ServerPacketID.UpdateMana              ' ASM
                        Call HandleUpdateMana
        
                Case ServerPacketID.UpdateHP                ' ASH
                        Call HandleUpdateHP
        
                Case ServerPacketID.UpdateGold              ' ASG
                        Call HandleUpdateGold
            
                Case ServerPacketID.UpdateBankGold
                        Call HandleUpdateBankGold

                Case ServerPacketID.UpdateExp               ' ASE
                        Call HandleUpdateExp
        
                Case ServerPacketID.ChangeMap               ' CM
                        Call HandleChangeMap
        
                Case ServerPacketID.PosUpdate               ' PU
                        Call HandlePosUpdate
        
                Case ServerPacketID.ChatOverHead            ' ||
                        Call HandleChatOverHead
        
                Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
                        Call HandleConsoleMessage
        
                Case ServerPacketID.GuildChat               ' |+
                        Call HandleGuildChat
        
                Case ServerPacketID.ShowMessageBox          ' !!
                        Call HandleShowMessageBox
        
                Case ServerPacketID.UserIndexInServer       ' IU
                        Call HandleUserIndexInServer
        
                Case ServerPacketID.UserCharIndexInServer   ' IP
                        Call HandleUserCharIndexInServer
        
                Case ServerPacketID.CharacterCreate         ' CC
                        Call HandleCharacterCreate
        
                Case ServerPacketID.CharacterRemove         ' BP
                        Call HandleCharacterRemove
        
                Case ServerPacketID.CharacterChangeNick
                        Call HandleCharacterChangeNick
            
                Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
                        Call HandleCharacterMove
        
                Case ServerPacketID.CharacterAttackMovement
                        Call HandleCharacterAttackMovement
            
                Case ServerPacketID.ForceCharMove
                        Call HandleForceCharMove
        
                Case ServerPacketID.CharacterChange         ' CP
                        Call HandleCharacterChange
        
                Case ServerPacketID.ObjectCreate            ' HO
                        Call HandleObjectCreate
        
                Case ServerPacketID.ObjectDelete            ' BO
                        Call HandleObjectDelete
        
                Case ServerPacketID.BlockPosition           ' BQ
                        Call HandleBlockPosition
        
                Case ServerPacketID.PlayMIDI                ' TM
                        Call HandlePlayMIDI
        
                Case ServerPacketID.PlayWave                ' TW
                        Call HandlePlayWave
        
                Case ServerPacketID.guildList               ' GL
                        Call HandleGuildList
        
                Case ServerPacketID.AreaChanged             ' CA
                        Call HandleAreaChanged
        
                Case ServerPacketID.PauseToggle             ' BKW
                        Call HandlePauseToggle
        
                Case ServerPacketID.RainToggle              ' LLU
                        Call HandleRainToggle
        
                Case ServerPacketID.CreateFX                ' CFX
                        Call HandleCreateFX
        
                Case ServerPacketID.UpdateUserStats         ' EST
                        Call HandleUpdateUserStats
        
                Case ServerPacketID.WorkRequestTarget       ' T01
                        Call HandleWorkRequestTarget
        
                Case ServerPacketID.ChangeInventorySlot     ' CSI
                        Call HandleChangeInventorySlot
        
                Case ServerPacketID.ChangeBankSlot          ' SBO
                        Call HandleChangeBankSlot
        
                Case ServerPacketID.ChangeSpellSlot         ' SHS
                        Call HandleChangeSpellSlot
        
                Case ServerPacketID.Atributes               ' ATR
                        Call HandleAtributes
        
                Case ServerPacketID.BlacksmithWeapons       ' LAH
                        Call HandleBlacksmithWeapons
        
                Case ServerPacketID.BlacksmithArmors        ' LAR
                        Call HandleBlacksmithArmors
        
                Case ServerPacketID.CarpenterObjects        ' OBR
                        Call HandleCarpenterObjects
        
                Case ServerPacketID.RestOK                  ' DOK
                        Call HandleRestOK
        
                Case ServerPacketID.ErrorMsg                ' ERR
                        Call HandleErrorMessage
        
                Case ServerPacketID.Blind                   ' CEGU
                        Call HandleBlind
        
                Case ServerPacketID.Dumb                    ' DUMB
                        Call HandleDumb
        
                Case ServerPacketID.ShowSignal              ' MCAR
                        Call HandleShowSignal
        
                Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
                        Call HandleChangeNPCInventorySlot
        
                Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
                        Call HandleUpdateHungerAndThirst
        
                Case ServerPacketID.Fame                    ' FAMA
                        Call HandleFame
        
                Case ServerPacketID.MiniStats               ' MEST
                        Call HandleMiniStats
        
                Case ServerPacketID.LevelUp                 ' SUNI
                        Call HandleLevelUp
        
                Case ServerPacketID.AddForumMsg             ' FMSG
                        Call HandleAddForumMessage
        
                Case ServerPacketID.ShowForumForm           ' MFOR
                        Call HandleShowForumForm
        
                Case ServerPacketID.SetInvisible            ' NOVER
                        Call HandleSetInvisible
        
                Case ServerPacketID.MeditateToggle          ' MEDOK
                        Call HandleMeditateToggle
        
                Case ServerPacketID.BlindNoMore             ' NSEGUE
                        Call HandleBlindNoMore
        
                Case ServerPacketID.DumbNoMore              ' NESTUP
                        Call HandleDumbNoMore
        
                Case ServerPacketID.SendSkills              ' SKILLS
                        Call HandleSendSkills
        
                Case ServerPacketID.TrainerCreatureList     ' LSTCRI
                        Call HandleTrainerCreatureList
        
                Case ServerPacketID.guildNews               ' GUILDNE
                        Call HandleGuildNews
        
                Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
                        Call HandleOfferDetails
        
                Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
                        Call HandleAlianceProposalsList
        
                Case ServerPacketID.PeaceProposalsList      ' PEACEPR
                        Call HandlePeaceProposalsList
        
                Case ServerPacketID.CharacterInfo           ' CHRINFO
                        Call HandleCharacterInfo
        
                Case ServerPacketID.GuildLeaderInfo         ' LEADERI
                        Call HandleGuildLeaderInfo
        
                Case ServerPacketID.GuildDetails            ' CLANDET
                        Call HandleGuildDetails
        
                Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
                        Call HandleShowGuildFundationForm
        
                Case ServerPacketID.ParalizeOK              ' PARADOK
                        Call HandleParalizeOK
        
                Case ServerPacketID.ShowUserRequest         ' PETICIO
                        Call HandleShowUserRequest
        
                Case ServerPacketID.TradeOK                 ' TRANSOK
                        Call HandleTradeOK
        
                Case ServerPacketID.BankOK                  ' BANCOOK
                        Call HandleBankOK
        
                Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
                        Call HandleChangeUserTradeSlot
            
                Case ServerPacketID.SendNight               ' NOC
                        Call HandleSendNight
        
                Case ServerPacketID.Pong
                        Call HandlePong
        
                Case ServerPacketID.UpdateTagAndStatus
                        Call HandleUpdateTagAndStatus
        
                Case ServerPacketID.GuildMemberInfo
                        Call HandleGuildMemberInfo
        
                        '*******************
                        'GM messages
                        '*******************
                Case ServerPacketID.SpawnList               ' SPL
                        Call HandleSpawnList
        
                Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
                        Call HandleShowSOSForm
            
                Case ServerPacketID.ShowDenounces
                        Call HandleShowDenounces
            
                Case ServerPacketID.RecordDetails
                        Call HandleRecordDetails
            
                Case ServerPacketID.RecordList
                        Call HandleRecordList
            
                Case ServerPacketID.ShowMOTDEditionForm     ' ZMOTD
                        Call HandleShowMOTDEditionForm
        
                Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
                        Call HandleShowGMPanelForm
        
                Case ServerPacketID.UserNameList            ' LISTUSU
                        Call HandleUserNameList
            
                Case ServerPacketID.ShowGuildAlign
                        Call HandleShowGuildAlign
        
                Case ServerPacketID.ShowPartyForm
                        Call HandleShowPartyForm
        
                Case ServerPacketID.UpdateStrenghtAndDexterity
                        Call HandleUpdateStrenghtAndDexterity
            
                Case ServerPacketID.UpdateStrenght
                        Call HandleUpdateStrenght
            
                Case ServerPacketID.UpdateDexterity
                        Call HandleUpdateDexterity

                Case ServerPacketID.MultiMessage
                        Call HandleMultiMessage
        
                Case ServerPacketID.StopWorking
                        Call HandleStopWorking
            
                Case ServerPacketID.CancelOfferItem
                        Call HandleCancelOfferItem
            
                Case ServerPacketID.StrDextRunningOut
                        Call HandleStrDextRunningOut
                        
                Case ServerPacketID.CharacterUpdateHP
                        Call HandleCharacterUpdateHP
                        
                Case ServerPacketID.CreateDamage            ' CDMG
                        Call HandleCreateDamage
                        
                Case ServerPacketID.Canje                ' kidral
                        Call HandleCanje
            
                Case ServerPacketID.CanjePTS                ' kidral
                        Call HandlePuntos
                        
                Case ServerPacketID.ControlUserRecive
                        Call HandleReciveControlUser

                Case ServerPacketID.ControlUserShow
                        Call HandleShowControlUser
                        
                Case ServerPacketID.SendScreen
                        Call handleSendScreen

                Case ServerPacketID.MercadoList     ' LSTCRI
                        Call HandleMercadoList
            
            Case ServerPacketID.MontateToggle           'monturas
            Call HandleMontateToggle
            
                Case Else
                        'ERROR : Abort!
                        Exit Sub

        End Select
    
        'Done with this packet, move on to next one
        If incomingData.Length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
                Err.Clear
                Call HandleIncomingData

        End If

End Sub

Public Sub HandleMultiMessage()

        '***************************************************
        'Author: Unknown
        'Last Modification: 11/16/2010
        ' 09/28/2010: C4b3z0n - Ahora se le saco la "," a los minutos de distancia del /hogar, ya que a veces quedaba "12,5 minutos y 30segundos"
        ' 09/21/2010: C4b3z0n - Now the fragshooter operates taking the screen after the change of killed charindex to ghost only if target charindex is visible to the client, else it will take screenshot like before.
        ' 11/16/2010: Amraphen - Recoded how the FragShooter works.
        '***************************************************
        Dim BodyPart As Byte

        Dim Da�o As Integer
    
        With incomingData
                Call .ReadByte
    
                Select Case .ReadByte

                        'Case eMessages.DontSeeAnything
                          '     Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_NO_VES_NADA_INTERESANTE, 65, 190, 156, False, False, True)
        
                        Case eMessages.NPCSwing
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
        
                        Case eMessages.NPCKillUser
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
        
                        Case eMessages.BlockedWithShieldUser
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
                        Case eMessages.BlockedWithShieldOther
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
                        Case eMessages.UserSwing
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 223, 57, 23, True, False, True)
        
                        Case eMessages.SafeModeOn
                                Call frmMain.ControlSM(eSMType.sSafemode, True)
        
                        Case eMessages.SafeModeOff
                                Call frmMain.ControlSM(eSMType.sSafemode, False)
        
                        Case eMessages.ResuscitationSafeOff
                                Call frmMain.ControlSM(eSMType.sResucitation, False)
         
                        Case eMessages.ResuscitationSafeOn
                                Call frmMain.ControlSM(eSMType.sResucitation, True)
        
                        Case eMessages.NobilityLost
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)
        
                        Case eMessages.CantUseWhileMeditating
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)
        
                        Case eMessages.NPCHitUser

                                Select Case incomingData.ReadByte()

                                        Case bCabeza
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 250, 207, 58, True, False, True)
                
                                        Case bBrazoIzquierdo
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 250, 207, 58, True, False, True)
                
                                        Case bBrazoDerecho
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 250, 207, 58, True, False, True)
                
                                        Case bPiernaIzquierda
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 250, 207, 58, True, False, True)
                
                                        Case bPiernaDerecha
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 250, 207, 58, True, False, True)
                
                                        Case bTorso
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 250, 207, 58, True, False, True)

                                End Select
        
                        Case eMessages.UserHitNPC
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 250, 157, 58, True, False, True)
        
                        Case eMessages.UserAttackedSwing
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & charlist(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
        
                        Case eMessages.UserHittedByUser

                                Dim AttackerName As String
            
                                AttackerName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                                BodyPart = incomingData.ReadByte()
                                Da�o = incomingData.ReadInteger()
            
                                Select Case BodyPart

                                        Case bCabeza
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_CABEZA & Da�o & MENSAJE_2, 255, 0, 0, True, False, True)
                
                                        Case bBrazoIzquierdo
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Da�o & MENSAJE_2, 255, 0, 0, True, False, True)
                
                                        Case bBrazoDerecho
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Da�o & MENSAJE_2, 255, 0, 0, True, False, True)
                
                                        Case bPiernaIzquierda
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Da�o & MENSAJE_2, 255, 0, 0, True, False, True)
                
                                        Case bPiernaDerecha
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Da�o & MENSAJE_2, 255, 0, 0, True, False, True)
                
                                        Case bTorso
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_TORSO & Da�o & MENSAJE_2, 255, 0, 0, True, False, True)

                                End Select
        
                        Case eMessages.UserHittedUser

                                Dim VictimName As String
            
                                VictimName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                                BodyPart = incomingData.ReadByte()
                                Da�o = incomingData.ReadInteger()
            
                                Select Case BodyPart

                                        Case bCabeza
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_CABEZA & Da�o & MENSAJE_2, 233, 247, 60, True, False, True)
                
                                        Case bBrazoIzquierdo
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Da�o & MENSAJE_2, 233, 247, 60, True, False, True)
                
                                        Case bBrazoDerecho
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Da�o & MENSAJE_2, 233, 247, 60, True, False, True)
                
                                        Case bPiernaIzquierda
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Da�o & MENSAJE_2, 233, 247, 60, True, False, True)
                
                                        Case bPiernaDerecha
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Da�o & MENSAJE_2, 233, 247, 60, True, False, True)
                
                                        Case bTorso
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_TORSO & Da�o & MENSAJE_2, 233, 247, 60, True, False, True)

                                End Select
        
                        Case eMessages.WorkRequestTarget
                                UsingSkill = incomingData.ReadByte()
            
                                frmMain.MousePointer = 2
            
                                Select Case UsingSkill

                                        Case Magia
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                
                                        Case Pesca
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                
                                        Case Robar
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                
                                        Case Talar
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                
                                        Case Mineria
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                
                                        Case FundirMetal
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                
                                        Case Proyectiles
                                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)

                                End Select

                        Case eMessages.HaveKilledUser

                                Dim KilledUser As Integer

                                Dim Exp        As Long
            
                                KilledUser = .ReadInteger
                                Exp = .ReadLong
            
                                Call ShowConsoleMsg(MENSAJE_HAS_MATADO_A & charlist(KilledUser).Nombre & MENSAJE_22, 255, 128, 0, True, False)
                                Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & Exp & MENSAJE_HAS_GANADO_EXPE_2, 255, 128, 0, True, False)
            
                                'Sacamos un screenshot si est� activado el FragShooter:
                                If ClientSetup.bKill And ClientSetup.bActive Then
                                        If Exp \ 2 > ClientSetup.byMurderedLevel Then
                                                FragShooterNickname = charlist(KilledUser).Nombre
                                                FragShooterKilledSomeone = True
                    
                                                FragShooterCapturePending = True

                                        End If

                                End If
            
                        Case eMessages.UserKill

                                Dim KillerUser As Integer
            
                                KillerUser = .ReadInteger
            
                                Call ShowConsoleMsg(charlist(KillerUser).Nombre & MENSAJE_TE_HA_MATADO, 255, 0, 0, True, False)
            
                                'Sacamos un screenshot si est� activado el FragShooter:
                                If ClientSetup.bDie And ClientSetup.bActive Then
                                        FragShooterNickname = charlist(KillerUser).Nombre
                                        FragShooterKilledSomeone = False
                
                                        FragShooterCapturePending = True

                                End If
                
                        Case eMessages.EarnExp
                                'Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & .ReadLong & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
        
                        Case eMessages.GoHome

                                Dim Distance As Byte

                                Dim Hogar    As String

                                Dim tiempo   As Integer

                                Dim msg      As String
            
                                Distance = .ReadByte
                                tiempo = .ReadInteger
                                Hogar = .ReadASCIIString
            
                                If tiempo >= 60 Then
                                        If tiempo Mod 60 = 0 Then
                                                msg = tiempo / 60 & " minutos."
                                        Else
                                                msg = CInt(tiempo \ 60) & " minutos y " & tiempo Mod 60 & " segundos."  'Agregado el CInt() asi el n�mero no es con , [C4b3z0n - 09/28/2010]

                                        End If

                                Else
                                        msg = tiempo & " segundos."

                                End If
            
                                Call ShowConsoleMsg("Te encuentras a " & Distance & " mapas de la " & Hogar & ", este viaje durar� " & msg, 255, 0, 0, True)
                                Traveling = True
        
                        Case eMessages.FinishHome
                                Call ShowConsoleMsg(MENSAJE_HOGAR, 255, 255, 255)
                                Traveling = False
        
                        Case eMessages.CancelGoHome
                                Call ShowConsoleMsg(MENSAJE_HOGAR_CANCEL, 255, 0, 0, True)
                                Traveling = False

                End Select

        End With

End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        ' Variable initialization
        UserClase = incomingData.ReadByte
        EngineRun = True
        Nombres = True
        bRain = False
    
        'Set connected state
        Call SetConnected

End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        Call Dialogos.RemoveAllDialogs

End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Check if the packet is complete
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Call Dialogos.RemoveDialog(incomingData.ReadInteger())

End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserNavegando = Not UserNavegando

End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
  
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Close connection
       
        frmMain.Socket1.Disconnect
  
        ResetAllInfo

End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        Set InvComUsu = Nothing
        Set InvComNpc = Nothing
    
        'Hide form
        Unload frmComerciar
    
        'Reset vars
        Comerciando = False

End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        Set InvBanco(0) = Nothing
        Set InvBanco(1) = Nothing
    
        Unload frmBancoObj
        Comerciando = False

End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim i As Long
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Set InvComUsu = New clsGrapchicalInventory
        Set InvComNpc = New clsGrapchicalInventory
    
        ' Initialize commerce inventories
        Call InvComUsu.Initialize(DirectDraw, frmComerciar.picInvUser, Inventario.MaxObjs)
        Call InvComNpc.Initialize(DirectDraw, frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS)

        'Fill user inventory
        For i = 1 To MAX_INVENTORY_SLOTS

                If Inventario.ObjIndex(i) <> 0 Then

                        With Inventario
                                Call InvComUsu.SetItem(i, .ObjIndex(i), _
                                   .Amount(i), .Equipped(i), .GrhIndex(i), _
                                   .ObjType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                                   .Valor(i), .ItemName(i))

                        End With

                End If

        Next i
    
        ' Fill Npc inventory
        For i = 1 To 50

                If NPCInventory(i).ObjIndex <> 0 Then

                        With NPCInventory(i)
                                Call InvComNpc.SetItem(i, .ObjIndex, _
                                   .Amount, 0, .GrhIndex, _
                                   .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                                   .Valor, .Name)

                        End With

                End If

        Next i
    
        'Set state and show form
        Comerciando = True
        frmComerciar.Show , frmMain

End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim i        As Long

        Dim BankGold As Long
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Set InvBanco(0) = New clsGrapchicalInventory
        Set InvBanco(1) = New clsGrapchicalInventory
    
        BankGold = incomingData.ReadLong
        Call InvBanco(0).Initialize(DirectDraw, frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS)
        Call InvBanco(1).Initialize(DirectDraw, frmBancoObj.PicInv, Inventario.MaxObjs)
    
        For i = 1 To Inventario.MaxObjs

                With Inventario
                        Call InvBanco(1).SetItem(i, .ObjIndex(i), _
                           .Amount(i), .Equipped(i), .GrhIndex(i), _
                           .ObjType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                           .Valor(i), .ItemName(i))

                End With

        Next i
    
        For i = 1 To MAX_BANCOINVENTORY_SLOTS

                With UserBancoInventory(i)
                        Call InvBanco(0).SetItem(i, .ObjIndex, _
                           .Amount, .Equipped, .GrhIndex, _
                           .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                           .Valor, .Name)

                End With

        Next i
    
        'Set state and show form
        Comerciando = True
    
        frmBancoObj.lblUserGld.Caption = BankGold
    
        frmBancoObj.Show , frmMain

End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim i As Long
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        TradingUserName = incomingData.ReadASCIIString
    
        Set InvComUsu = New clsGrapchicalInventory
        Set InvOfferComUsu(0) = New clsGrapchicalInventory
        Set InvOfferComUsu(1) = New clsGrapchicalInventory
        Set InvOroComUsu(0) = New clsGrapchicalInventory
        Set InvOroComUsu(1) = New clsGrapchicalInventory
        Set InvOroComUsu(2) = New clsGrapchicalInventory
    
        ' Initialize commerce inventories
        Call InvComUsu.Initialize(DirectDraw, frmComerciarUsu.picInvComercio, Inventario.MaxObjs)
        Call InvOfferComUsu(0).Initialize(DirectDraw, frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
        Call InvOfferComUsu(1).Initialize(DirectDraw, frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
        Call InvOroComUsu(0).Initialize(DirectDraw, frmComerciarUsu.picInvOroProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2, , , , True)
        Call InvOroComUsu(1).Initialize(DirectDraw, frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2, , , , True)
        Call InvOroComUsu(2).Initialize(DirectDraw, frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2, , , , True)

        'Fill user inventory
        For i = 1 To MAX_INVENTORY_SLOTS

                If Inventario.ObjIndex(i) <> 0 Then

                        With Inventario
                                Call InvComUsu.SetItem(i, .ObjIndex(i), _
                                   .Amount(i), .Equipped(i), .GrhIndex(i), _
                                   .ObjType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                                   .Valor(i), .ItemName(i))

                        End With

                End If

        Next i

        ' Inventarios de oro
        Call InvOroComUsu(0).SetItem(1, ORO_INDEX, UserGLD, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
        Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
        Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")

        'Set state and show form
        Comerciando = True
        Call frmComerciarUsu.Show(vbModeless, frmMain)

End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        Set InvComUsu = Nothing
        Set InvOroComUsu(0) = Nothing
        Set InvOroComUsu(1) = Nothing
        Set InvOroComUsu(2) = Nothing
        Set InvOfferComUsu(0) = Nothing
        Set InvOfferComUsu(1) = Nothing
    
        'Destroy the form and reset the state
        Unload frmComerciarUsu
        Comerciando = False

End Sub

''
' Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/12/2009
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        With frmComerciarUsu
                ' Now he can accept the offer or reject it
                .HabilitarAceptarRechazar True
        
                .PrintCommerceMsg TradingUserName & " ha confirmado su oferta!", FontTypeNames.FONTTYPE_CONSE

        End With
    
End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
                Call WriteCraftBlacksmith(MacroBltIndex)
        Else
                frmHerrero.Show , frmMain
                MirandoHerreria = True

        End If

End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
                Call WriteCraftCarpenter(MacroBltIndex)
        Else
                frmCarp.Show , frmMain
                MirandoCarpinteria = True

        End If

End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Check packet is complete

        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode

                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Get data and update form
        UserMinSTA = incomingData.ReadInteger()
    
        frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
        Dim bWidth As Byte
    
        
        bWidth = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 99)
    
        frmMain.shpEnergia.Width = bWidth
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Check packet is complete

        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode

                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Get data and update form
        UserMinMAN = incomingData.ReadInteger()
    
        frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
        
        'If UserMinMAN Then frmMain.shpMana.Width = 0: Exit Sub
        If UserMinMAN > 0 Then _
        frmMain.shpMana.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 99)

End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Check packet is complete

        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode

                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Get data and update form
        UserMinHP = incomingData.ReadInteger()
    
        Dim bWidth As Integer
    
        bWidth = (((UserMinHP / 100) / (UserMaxHP / 100)) * 100)
    
        frmMain.lblVida.Caption = UserMinHP & "/" & UserMaxHP
        frmMain.shpVida.Width = bWidth
    
        'Is the user alive??

        If UserMinHP = 0 Then
                UserEstado = 1

                If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
                If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
        Else
                UserEstado = 0

        End If

End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()

        '***************************************************
        'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 09/21/10
        'Last Modified By: C4b3z0n
        '- 08/14/07: Tavo - Added GldLbl color variation depending on User Gold and Level
        '- 09/21/10: C4b3z0n - Modified color change of gold ONLY if the player's level is greater than 12 (NOT newbie).
        '***************************************************
        'Check packet is complete
        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Get data and update form
        UserGLD = incomingData.ReadLong()
    
        If UserGLD >= CLng(UserLvl) * 10000 And UserLvl > 12 Then 'Si el nivel es mayor de 12, es decir, no es newbie.
                'Changes color
                frmMain.GldLbl.ForeColor = &HFF& 'Red
        Else
                'Changes color
                frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow

        End If
    
        frmMain.GldLbl.Caption = UserGLD

End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()

        '***************************************************
        'Autor: ZaMa
        'Last Modification: 14/12/2009
        '
        '***************************************************
        'Check packet is complete
        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        frmBancoObj.lblUserGld.Caption = incomingData.ReadLong
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Check packet is complete
        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Get data and update form
        UserExp = incomingData.ReadLong()
        frmMain.lblExp.Caption = UserExp & "/" & UserPasarNivel
        frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"

End Sub

''
' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()

        '***************************************************
        'Author: Budi
        'Last Modification: 11/26/09
        '***************************************************
        'Check packet is complete
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Get data and update form
        UserFuerza = incomingData.ReadByte
        UserAgilidad = incomingData.ReadByte
        frmMain.lblStrg.Caption = UserFuerza
        frmMain.lblDext.Caption = UserAgilidad
    
        frmMain.tmrBlink.Enabled = False
    
        frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
        frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)

End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()

        '***************************************************
        'Author: Budi
        'Last Modification: 11/26/09
        '***************************************************
        'Check packet is complete
        If incomingData.Length < 2 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Get data and update form
        UserFuerza = incomingData.ReadByte
        frmMain.lblStrg.Caption = UserFuerza
        frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
        frmMain.tmrBlink.Enabled = False

End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()

        '***************************************************
        'Author: Budi
        'Last Modification: 11/26/09
        '***************************************************
        'Check packet is complete
        If incomingData.Length < 2 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Get data and update form
        UserAgilidad = incomingData.ReadByte
        frmMain.lblDext.Caption = UserAgilidad
        frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)
        frmMain.tmrBlink.Enabled = False

End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserMap = incomingData.ReadInteger()
    
        'TODO: Once on-the-fly editor is implemented check for map version before loading....
        'For now we just drop it
        Call incomingData.ReadInteger
        
        ZonaSeguraM = incomingData.ReadBoolean
        
        'ZonaSeguraM = True
        
        If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
                Call SwitchMap(UserMap)

                If bLluvia(UserMap) = 0 Then
                        If bRain And UserMap = 1 Then
                                Call Audio.StopWave(RainBufferIndex)
                                RainBufferIndex = 0
                                frmMain.IsPlaying = PlayLoop.plNone

                        End If

                End If

        Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, alg�n archivo ha sido modificado o esta da�ado."
        
                Call CloseClient

        End If

End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        'Remove char from old position
        If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
                MapData(UserPos.X, UserPos.Y).CharIndex = 0

        End If
    
        'Set new pos
        UserPos.X = incomingData.ReadByte()
        UserPos.Y = incomingData.ReadByte()
    
        'Set char
        MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
        charlist(UserCharIndex).Pos = UserPos
    
        'Are we under a roof?
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
           MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
           MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
                
        'Update pos label
        frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y

End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 8 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim chat      As String

        Dim CharIndex As Integer

        Dim r         As Byte

        Dim g         As Byte

        Dim B         As Byte
    
        chat = Buffer.ReadASCIIString()
        CharIndex = Buffer.ReadInteger()
    
        r = Buffer.ReadByte()
        g = Buffer.ReadByte()
        B = Buffer.ReadByte()
    
        'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
        If charlist(CharIndex).Active Then _
           Call Dialogos.CreateDialog(Trim$(chat), CharIndex, RGB(r, g, B))
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)

errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing
    
        If Error <> 0 Then _
           Err.Raise Error

End Sub

Private Sub HandleConsoleMessage()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 12/05/11
        'D'Artagnan: Agrego la divisi�n de consolas
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim chat      As String

        Dim FontIndex As Integer

        Dim str       As String

        Dim r         As Byte

        Dim g         As Byte

        Dim B         As Byte
    
        chat = Buffer.ReadASCIIString()
        FontIndex = Buffer.ReadByte()

        If InStr(1, chat, "~") Then
                str = ReadField(2, chat, 126)

                If Val(str) > 255 Then
                        r = 255
                Else
                        r = Val(str)

                End If
            
                str = ReadField(3, chat, 126)

                If Val(str) > 255 Then
                        g = 255
                Else
                        g = Val(str)

                End If
            
                str = ReadField(4, chat, 126)

                If Val(str) > 255 Then
                        B = 255
                Else
                        B = Val(str)

                End If

                Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        
        Else

                If FontIndex = FontTypeNames.FONTTYPE_PARTY Then 'CHOTS | Mensajes personalizados de Party
                    
                        With FontTypes(FontIndex)
                                Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .Bold, .Italic)

                        End With

                Else

                        With FontTypes(FontIndex)
                       ' If FontIndex > 20 Then 'Debug.Print "rgb: " & .red & " " & .green & " " & .blue
                                Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .Bold, .Italic)
                        'End If
                        End With

                End If
        
                ' Para no perder el foco cuando chatea por party
                If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
                        If MirandoParty Then frmParty.SendTxt.SetFocus

                End If

        End If

        '    Call checkText(chat)
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 12/05/11 (D'Artagnan)
        'Redirect messages to guild console
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim chat   As String

        Dim IsMOTD As Boolean

        Dim str    As String

        Dim r      As Byte

        Dim g      As Byte

        Dim B      As Byte
    
        chat = Buffer.ReadASCIIString()
        IsMOTD = Buffer.ReadBoolean()
    
        If Not DialogosClanes.Activo Then
                If InStr(1, chat, "~") Then
                        str = ReadField(2, chat, 126)

                        If Val(str) > 255 Then
                                r = 255
                        Else
                                r = Val(str)

                        End If
            
                        str = ReadField(3, chat, 126)

                        If Val(str) > 255 Then
                                g = 255
                        Else
                                g = Val(str)

                        End If
            
                        str = ReadField(4, chat, 126)

                        If Val(str) > 255 Then
                                B = 255
                        Else
                                B = Val(str)

                        End If
            
                        If IsMOTD = True Then
                                Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
                        Else
                                Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)

                        End If

                Else

                        With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)

                                If IsMOTD = True Then
                                        Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .Bold, .Italic)
                                Else
                                        Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .Bold, .Italic)

                                End If

                        End With

                End If

        Else
                Call DialogosClanes.PushBackText(ReadField(1, chat, 126))

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()

        '***************************************************
        'Author: ZaMa
        'Last Modification: 03/12/2009
        '
        '***************************************************
        If incomingData.Length < 4 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim chat      As String

        Dim FontIndex As Integer

        Dim str       As String

        Dim r         As Byte

        Dim g         As Byte

        Dim B         As Byte
    
        chat = Buffer.ReadASCIIString()
        FontIndex = Buffer.ReadByte()
    
        If InStr(1, chat, "~") Then
                str = ReadField(2, chat, 126)

                If Val(str) > 255 Then
                        r = 255
                Else
                        r = Val(str)

                End If
            
                str = ReadField(3, chat, 126)

                If Val(str) > 255 Then
                        g = 255
                Else
                        g = Val(str)

                End If
            
                str = ReadField(4, chat, 126)

                If Val(str) > 255 Then
                        B = 255
                Else
                        B = Val(str)

                End If
            
                Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), r, g, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        Else

                With FontTypes(FontIndex)
                        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .red, .green, .blue, .Bold, .Italic)

                End With

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        frmMensaje.msg.Caption = Buffer.ReadASCIIString()
        frmMensaje.Show
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserIndex = incomingData.ReadInteger()

End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserCharIndex = incomingData.ReadInteger()
        UserPos = charlist(UserCharIndex).Pos
    
        'Are we under a roof?
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
           MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
           MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

        frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y

End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

    Call Buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call Buffer.ReadByte

    Dim CharIndex As Integer

    Dim Body   As Integer

    Dim Head   As Integer

    Dim Heading As E_Heading

    Dim X      As Byte

    Dim Y      As Byte

    Dim weapon As Integer

    Dim shield As Integer

    Dim helmet As Integer

    Dim privs  As Integer

    Dim NickColor As Byte

    CharIndex = Buffer.ReadInteger()
    Body = Buffer.ReadInteger()
    Head = Buffer.ReadInteger()
    Heading = Buffer.ReadByte()
    X = Buffer.ReadByte()
    Y = Buffer.ReadByte()
    weapon = Buffer.ReadInteger()
    shield = Buffer.ReadInteger()
    helmet = Buffer.ReadInteger()

    With charlist(CharIndex)
        Call SetCharacterFx(CharIndex, Buffer.ReadInteger(), Buffer.ReadInteger())

        .Nombre = Buffer.ReadASCIIString()

        
        NickColor = Buffer.ReadByte()

        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0

        End If

        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0

        .EsGranPoder = (NickColor And eNickColor.ieGranPoder) <> 0

        privs = Buffer.ReadByte()

        .Min_HP = Buffer.ReadLong()
        .Max_Hp = Buffer.ReadLong()
        
        
       Dim tmpConq As Boolean
        tmpConq = Buffer.ReadBoolean
        .Conquisto = tmpConq
        .EsDonador = Buffer.ReadByte
        .esNPC = Buffer.ReadByte
        
        If privs <> 0 Then

            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil

            End If

            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil

            End If

            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster

            End If

            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0

        End If
    
    End With

    Call MakeChar(CharIndex, Body, Head, Heading, X, Y, weapon, shield, helmet)

    Call RefreshAllChars
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

errhandler:

    Dim Error  As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

Private Sub HandleCharacterChangeNick()

        '***************************************************
        'Author: Budi
        'Last Modification: 07/23/09
        '
        '***************************************************
        If incomingData.Length < 6 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet id
        Call incomingData.ReadByte

        Dim CharIndex As Integer

        CharIndex = incomingData.ReadInteger
        charlist(CharIndex).Nombre = incomingData.ReadASCIIString
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim CharIndex As Integer
    
        CharIndex = incomingData.ReadInteger()
    
        Call EraseChar(CharIndex)
        Call RefreshAllChars

End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim CharIndex As Integer

        Dim X         As Byte

        Dim Y         As Byte
    
        CharIndex = incomingData.ReadInteger()
        X = incomingData.ReadByte()
        Y = incomingData.ReadByte()
    
        With charlist(CharIndex)

                If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
                        .FxIndex = 0

                End If
        
                ' Play steps sounds if the user is not an admin of any kind
                If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
                        Call DoPasosFx(CharIndex)

                End If

        End With
    
        Call MoveCharbyPos(CharIndex, X, Y)
    
        Call RefreshAllChars

End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
        If incomingData.Length < 2 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim Direccion As Byte
    
        Direccion = incomingData.ReadByte()

        Call MoveCharbyHead(UserCharIndex, Direccion)
        Call MoveScreen(Direccion)
    
        Call RefreshAllChars

End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 21/09/2010 - C4b3z0n
        '25/08/2009: ZaMa - Changed a variable used incorrectly.
        '21/09/2010: C4b3z0n - Added code for FragShooter. If its waiting for the death of certain UserIndex, and it dies, then the capture of the screen will occur.
        '***************************************************
        If incomingData.Length < 18 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim CharIndex As Integer

        Dim tempint   As Integer

        Dim headIndex As Integer
    
        CharIndex = incomingData.ReadInteger()
    
        With charlist(CharIndex)
                tempint = incomingData.ReadInteger()
        
                If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
                        .Body = BodyData(0)
                        .iBody = 0
                Else
                        .Body = BodyData(tempint)
                        .iBody = tempint

                End If
        
                headIndex = incomingData.ReadInteger()
        
                If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
                        .Head = HeadData(0)
                        .iHead = 0
                Else
                        .Head = HeadData(headIndex)
                        .iHead = headIndex

                End If
        
                .muerto = (headIndex = CASPER_HEAD)
        
                .Heading = incomingData.ReadByte()
        
                tempint = incomingData.ReadInteger()

                If tempint <> 0 Then .Arma = WeaponAnimData(tempint)
        
                tempint = incomingData.ReadInteger()

                If tempint <> 0 Then .Escudo = ShieldAnimData(tempint)
        
                tempint = incomingData.ReadInteger()

                If tempint <> 0 Then .Casco = CascoAnimData(tempint)
        
                Call SetCharacterFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())

        End With
    
        Call RefreshAllChars

End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim X As Byte

        Dim Y As Byte
    
        X = incomingData.ReadByte()
        Y = incomingData.ReadByte()
    
        MapData(X, Y).ObjGrh.GrhIndex = incomingData.ReadInteger()
    
        Call InitGrh(MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex)

End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim X As Byte

        Dim Y As Byte
    
        X = incomingData.ReadByte()
        Y = incomingData.ReadByte()
        MapData(X, Y).ObjGrh.GrhIndex = 0

End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 4 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim X As Byte

        Dim Y As Byte
    
        X = incomingData.ReadByte()
        Y = incomingData.ReadByte()
    
        If incomingData.ReadBoolean() Then
                MapData(X, Y).Blocked = 1
        Else
                MapData(X, Y).Blocked = 0

        End If

End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        Dim currentMidi As Integer

        Dim Loops       As Integer
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        currentMidi = incomingData.ReadInteger()
        Loops = incomingData.ReadInteger()
    
        If currentMidi Then
                If currentMidi > MP3_INITIAL_INDEX Then
                        Call Audio.MusicMP3Play(App.path & "\MP3\" & currentMidi & ".mp3")
                Else
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Loops)

                End If

        End If
    
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()

        '***************************************************
        'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 08/14/07
        'Last Modified by: Rapsodius
        'Added support for 3D Sounds.
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
        
        Dim wave As Byte

        Dim srcX As Byte

        Dim srcY As Byte
    
        wave = incomingData.ReadByte()
        srcX = incomingData.ReadByte()
        srcY = incomingData.ReadByte()
        
        Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)

End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        With frmGuildAdm
                'Clear guild's list
                .guildslist.Clear
        
                GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
                Dim i As Long

                For i = 0 To UBound(GuildNames())
                        Call .guildslist.AddItem(GuildNames(i))
                Next i
        
                'If we got here then packet is complete, copy data back to original queue
                Call incomingData.CopyBuffer(Buffer)
        
                .Show vbModeless, frmMain

        End With
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim X As Byte

        Dim Y As Byte
    
        X = incomingData.ReadByte()
        Y = incomingData.ReadByte()
        
        Call CambioDeArea(X, Y)

End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        pausa = Not pausa

End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
        
        If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
        GoTo lalaxd
        
        bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
           MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
           MapData(UserPos.X, UserPos.Y).Trigger = 4)
        
        If UserMap = 1 Then
        
        If bRain And UserMap <> 1 Then Call Audio.StopWave(RainBufferIndex)
        
        If bRain Then
                If bLluvia(UserMap) Then
                        'Stop playing the rain sound
                        Call Audio.StopWave(RainBufferIndex)
                        RainBufferIndex = 0

                        If bTecho Then
                                Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
                        Else
                                Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)

                        End If

                        frmMain.IsPlaying = PlayLoop.plNone

                End If

        End If
    
        End If
        
lalaxd:
        
        bRain = Not bRain
        
        
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 7 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim CharIndex As Integer

        Dim fX        As Integer

        Dim Loops     As Integer
    
        CharIndex = incomingData.ReadInteger()
        fX = incomingData.ReadInteger()
        Loops = incomingData.ReadInteger()
    
        Call SetCharacterFx(CharIndex, fX, Loops)

End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 26 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserMaxHP = incomingData.ReadInteger()
        UserMinHP = incomingData.ReadInteger()
        UserMaxMAN = incomingData.ReadInteger()
        UserMinMAN = incomingData.ReadInteger()
        UserMaxSTA = incomingData.ReadInteger()
        UserMinSTA = incomingData.ReadInteger()
        UserGLD = incomingData.ReadLong()
        UserLvl = incomingData.ReadByte()
        UserPasarNivel = incomingData.ReadLong()
        UserExp = incomingData.ReadLong()
    
        frmMain.lblExp.Caption = UserExp & "/" & UserPasarNivel
    
        If UserPasarNivel > 0 Then
                frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
        Else
                frmMain.lblPorcLvl.Caption = "[N/A]"

        End If
    
        frmMain.GldLbl.Caption = UserGLD
        frmMain.lblLvl.Caption = UserLvl
    
        'Stats
        frmMain.lblMana.Caption = UserMinMAN & "/" & UserMaxMAN
        frmMain.lblVida.Caption = UserMinHP & "/" & UserMaxHP
        frmMain.lblEnergia.Caption = UserMinSTA & "/" & UserMaxSTA
    
        'Dim bWidth As Integer
    
        '*************************** kidral

        If UserPasarNivel > 0 Then
                frmMain.ImgExp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 200)
        Else
                frmMain.ImgExp.Width = 0

        End If

        If UserMinMAN > 0 Then
                frmMain.shpMana.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 100)

        End If

        frmMain.shpVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 100)
    
        frmMain.shpEnergia.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 100)
        '***************************
    
        If UserMinHP = 0 Then
                UserEstado = 1

                If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
                If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
        Else
                UserEstado = 0

        End If
    
        If UserGLD >= CLng(UserLvl) * 10000 Then
                'Changes color
                frmMain.GldLbl.ForeColor = &HFFFF& 'Red
        Else
                'Changes color
                frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow

        End If

End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 2 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        UsingSkill = incomingData.ReadByte()

        frmMain.MousePointer = 2
    
        Select Case UsingSkill

                Case Magia
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)

                Case Pesca
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)

                Case Robar
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)

                Case Talar
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)

                Case Mineria
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)

                Case FundirMetal
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)

                Case Proyectiles
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)

        End Select

End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 13 Then ' @@ Miqueas Antes 22, reduccion de consumo muy hard
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim Slot     As Byte

        Dim ObjIndex As Integer

        Dim Name     As String

        Dim Amount   As Integer

        Dim Equipped As Boolean

        Dim GrhIndex As Integer

        Dim ObjType  As Byte

        Dim MaxHit   As Integer

        Dim MinHit   As Integer

        Dim MaxDef   As Integer

        Dim MinDef   As Integer

        Dim Value    As Single
        
        Dim Caos     As Byte
        
        Dim Real     As Byte
    
        Slot = Buffer.ReadByte()
        ObjIndex = Buffer.ReadInteger()
        Name = getTypeObj(ObjIndex, enum_TypeObj.Nombre) 'Buffer.ReadASCIIString()
        
        Amount = Buffer.ReadInteger()
        Equipped = Buffer.ReadBoolean()
        
        GrhIndex = Val(getTypeObj(ObjIndex, enum_TypeObj.GrhIndex)) 'Buffer.ReadInteger()
        ObjType = Val(getTypeObj(ObjIndex, enum_TypeObj.ObjType)) 'Buffer.ReadByte()
        MaxHit = Val(getTypeObj(ObjIndex, enum_TypeObj.MaxHit)) 'Buffer.ReadInteger()
        MinHit = Val(getTypeObj(ObjIndex, enum_TypeObj.MinHit)) 'Buffer.ReadInteger()
        
        ' @@ Usados para identificar las armaduras de segundo rango, que tiene un bono de defensa
        Real = Buffer.ReadByte()
        Caos = Buffer.ReadByte()

        If Real = 2 Or Caos = 2 Then
                MaxDef = Val(getTypeObj(ObjIndex, enum_TypeObj.MaxDef) * MOD_DEF_SEG_JERARQUIA) 'Buffer.ReadInteger()
                MinDef = Val(getTypeObj(ObjIndex, enum_TypeObj.MinDef) * MOD_DEF_SEG_JERARQUIA) 'Buffer.ReadInteger()
        
        Else
                MaxDef = Val(getTypeObj(ObjIndex, enum_TypeObj.MaxDef)) 'Buffer.ReadInteger()
                MinDef = Val(getTypeObj(ObjIndex, enum_TypeObj.MinDef)) 'Buffer.ReadInteger()

        End If

        Value = Buffer.ReadSingle()
    
        If Equipped Then

                Select Case ObjType

                        Case eObjType.otWeapon
                                frmMain.lblWeapon = MinHit & "/" & MaxHit
                                UserWeaponEqpSlot = Slot

                        Case eObjType.otArmadura
                                frmMain.lblArmor = MinDef & "/" & MaxDef
                                UserArmourEqpSlot = Slot

                        Case eObjType.otescudo
                                frmMain.lblShielder = MinDef & "/" & MaxDef
                                UserHelmEqpSlot = Slot

                        Case eObjType.otcasco
                                frmMain.lblHelm = MinDef & "/" & MaxDef
                                UserShieldEqpSlot = Slot

                End Select

        Else

                Select Case Slot

                        Case UserWeaponEqpSlot
                                frmMain.lblWeapon = "0/0"
                                UserWeaponEqpSlot = 0

                        Case UserArmourEqpSlot
                                frmMain.lblArmor = "0/0"
                                UserArmourEqpSlot = 0

                        Case UserHelmEqpSlot
                                frmMain.lblShielder = "0/0"
                                UserHelmEqpSlot = 0

                        Case UserShieldEqpSlot
                                frmMain.lblHelm = "0/0"
                                UserShieldEqpSlot = 0

                End Select

        End If
    
        Call Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MaxDef, MinDef, Value, Name)

        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

' Handles the StopWorking message.
Private Sub HandleStopWorking()
        '***************************************************
        'Author: Budi
        'Last Modification: 12/01/09
        '
        '***************************************************

        Call incomingData.ReadByte
    
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("�Has terminado de trabajar!", .red, .green, .blue, .Bold, .Italic)

        End With
    
        If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo

End Sub

' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 05/03/10
        '
        '***************************************************
        Dim Slot   As Byte

        Dim Amount As Long
    
        Call incomingData.ReadByte
    
        Slot = incomingData.ReadByte
    
        With InvOfferComUsu(0)
                Amount = .Amount(Slot)
        
                ' No tiene sentido que se quiten 0 unidades
                If Amount <> 0 Then
                        ' Actualizo el inventario general
                        Call frmComerciarUsu.UpdateInvCom(.ObjIndex(Slot), Amount)
            
                        ' Borro el item
                        Call .SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)

                End If

        End With
    
        ' Si era el �nico �tem de la oferta, no puede confirmarla
        If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And _
           Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then Call frmComerciarUsu.HabilitarConfirmar(False)
    
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call frmComerciarUsu.PrintCommerceMsg("�No puedes comerciar ese objeto!", FontTypeNames.FONTTYPE_INFO)

        End With

End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 6 Then ' @@ Miqueas Antes 21, reduccion de consumo muy hard
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim Slot     As Byte
        Dim ObjIndex As Integer

        Slot = Buffer.ReadByte()
    
        With UserBancoInventory(Slot)
        
                ObjIndex = Buffer.ReadInteger()
                
                .ObjIndex = ObjIndex
                .Name = getTypeObj(ObjIndex, enum_TypeObj.Nombre) 'Buffer.ReadASCIIString()
                .Amount = Buffer.ReadInteger()
                .GrhIndex = Val(getTypeObj(ObjIndex, enum_TypeObj.GrhIndex)) 'Buffer.ReadInteger()
                .ObjType = Val(getTypeObj(ObjIndex, enum_TypeObj.ObjType)) 'Buffer.ReadByte()
                .MaxHit = Val(getTypeObj(ObjIndex, enum_TypeObj.MaxHit)) 'Buffer.ReadInteger()
                .MinHit = Val(getTypeObj(ObjIndex, enum_TypeObj.MinHit)) 'Buffer.ReadInteger()
                .MaxDef = Val(getTypeObj(ObjIndex, enum_TypeObj.MaxDef)) 'Buffer.ReadInteger()
                .MinDef = Val(getTypeObj(ObjIndex, enum_TypeObj.MinDef)) 'Buffer.ReadInteger
                .Valor = Val(getTypeObj(ObjIndex, enum_TypeObj.Valor)) 'Buffer.ReadLong()
        
                If Comerciando Then
                        Call InvBanco(0).SetItem(Slot, .ObjIndex, .Amount, _
                           .Equipped, .GrhIndex, .ObjType, .MaxHit, _
                           .MinHit, .MaxDef, .MinDef, .Valor, .Name)

                End If

        End With
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 4 Then ' @@ Miqueas Antes 6, reduccion de consumo leve
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim Slot      As Byte
        Dim SpellName As String

        Slot = Buffer.ReadByte()
    
        UserHechizos(Slot) = Buffer.ReadInteger()
        SpellName = getNameHechizo(UserHechizos(Slot))

        If Not Len(SpellName) <> 0 Then SpellName = "-None-" ' @@ Miqueas : Parche - 08/12/15
                
        If Slot <= frmMain.hlst.ListCount Then
                frmMain.hlst.List(Slot - 1) = SpellName
        Else
                Call frmMain.hlst.AddItem(SpellName)

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 1 + NUMATRIBUTES Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim i As Long
    
        For i = 1 To NUMATRIBUTES
                UserAtributos(i) = incomingData.ReadByte()
        Next i

        LlegaronAtrib = True
      
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 5 Then ' @@ Miqueas Antes 17, reduccion de consumo muy hard
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim Count As Integer

        Dim i     As Long

        Dim j     As Long

        Dim K     As Long
    
        Count = Buffer.ReadInteger()
    
        ReDim ArmasHerrero(Count) As tItemsConstruibles
        ReDim HerreroMejorar(0) As tItemsConstruibles
    
        For i = 1 To Count

                With ArmasHerrero(i)
                
                        .ObjIndex = Buffer.ReadInteger()
                        .Name = getTypeObj(.ObjIndex, enum_TypeObj.Nombre) 'Buffer.ReadASCIIString()    'Get the object's name
                        .GrhIndex = Val(getTypeObj(.ObjIndex, enum_TypeObj.GrhIndex)) 'Buffer.ReadInteger()
                        .LinH = Val(getTypeObj(.ObjIndex, enum_TypeObj.LingH)) 'Buffer.ReadInteger()        'The iron needed
                        .LinP = Val(getTypeObj(.ObjIndex, enum_TypeObj.LingP)) 'Buffer.ReadInteger()        'The silver needed
                        .LinO = Val(getTypeObj(.ObjIndex, enum_TypeObj.LingO)) 'Buffer.ReadInteger()        'The gold needed
                        
                        .Upgrade = Val(getTypeObj(.ObjIndex, enum_TypeObj.Upgrade)) 'Buffer.ReadInteger()

                End With

        Next i
    
        For i = 1 To MAX_LIST_ITEMS
                Set InvLingosHerreria(i) = New clsGrapchicalInventory
        Next i
    
        With frmHerrero
                ' Inicializo los inventarios
                Call InvLingosHerreria(1).Initialize(DirectDraw, .picLingotes0, 3, , , , , , False)
                Call InvLingosHerreria(2).Initialize(DirectDraw, .picLingotes1, 3, , , , , , False)
                Call InvLingosHerreria(3).Initialize(DirectDraw, .picLingotes2, 3, , , , , , False)
                Call InvLingosHerreria(4).Initialize(DirectDraw, .picLingotes3, 3, , , , , , False)
        
                Call .HideExtraControls(Count)
                Call .RenderList(1, True)

        End With
    
        For i = 1 To Count

                With ArmasHerrero(i)

                        If .Upgrade Then

                                For K = 1 To Count

                                        If .Upgrade = ArmasHerrero(K).ObjIndex Then
                                                j = j + 1
                
                                                ReDim Preserve HerreroMejorar(j) As tItemsConstruibles
                        
                                                HerreroMejorar(j).Name = .Name
                                                HerreroMejorar(j).GrhIndex = .GrhIndex
                                                HerreroMejorar(j).ObjIndex = .ObjIndex
                                                HerreroMejorar(j).UpgradeName = ArmasHerrero(K).Name
                                                HerreroMejorar(j).UpgradeGrhIndex = ArmasHerrero(K).GrhIndex
                                                HerreroMejorar(j).LinH = ArmasHerrero(K).LinH - .LinH * 0.85
                                                HerreroMejorar(j).LinP = ArmasHerrero(K).LinP - .LinP * 0.85
                                                HerreroMejorar(j).LinO = ArmasHerrero(K).LinO - .LinO * 0.85
                        
                                                Exit For

                                        End If

                                Next K

                        End If

                End With

        Next i
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 5 Then ' @@ Miqueas Antes 17, reduccion de consumo muy hard
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim Count As Integer

        Dim i     As Long

        Dim j     As Long

        Dim K     As Long
    
        Count = Buffer.ReadInteger()
    
        ReDim ArmadurasHerrero(Count) As tItemsConstruibles
    
        For i = 1 To Count

                With ArmadurasHerrero(i)
                        .ObjIndex = Buffer.ReadInteger()
                        .Name = getTypeObj(.ObjIndex, enum_TypeObj.Nombre) 'Buffer.ReadASCIIString()    'Get the object's name
                        .GrhIndex = Val(getTypeObj(.ObjIndex, enum_TypeObj.GrhIndex)) 'Buffer.ReadInteger()
                        
                        .LinH = Val(getTypeObj(.ObjIndex, enum_TypeObj.LingH)) 'Buffer.ReadInteger()        'The iron needed
                        .LinP = Val(getTypeObj(.ObjIndex, enum_TypeObj.LingP)) 'Buffer.ReadInteger()        'The silver needed
                        .LinO = Val(getTypeObj(.ObjIndex, enum_TypeObj.LingO)) 'Buffer.ReadInteger()        'The gold needed
                        
                        .Upgrade = Val(getTypeObj(.ObjIndex, enum_TypeObj.Upgrade)) 'Buffer.ReadInteger()

                End With

        Next i
    
        j = UBound(HerreroMejorar)
    
        For i = 1 To Count

                With ArmadurasHerrero(i)

                        If .Upgrade Then

                                For K = 1 To Count

                                        If .Upgrade = ArmadurasHerrero(K).ObjIndex Then
                                                j = j + 1
                
                                                ReDim Preserve HerreroMejorar(j) As tItemsConstruibles
                        
                                                HerreroMejorar(j).Name = .Name
                                                HerreroMejorar(j).GrhIndex = .GrhIndex
                                                HerreroMejorar(j).ObjIndex = .ObjIndex
                                                HerreroMejorar(j).UpgradeName = ArmadurasHerrero(K).Name
                                                HerreroMejorar(j).UpgradeGrhIndex = ArmadurasHerrero(K).GrhIndex
                                                HerreroMejorar(j).LinH = ArmadurasHerrero(K).LinH - .LinH * 0.85
                                                HerreroMejorar(j).LinP = ArmadurasHerrero(K).LinP - .LinP * 0.85
                                                HerreroMejorar(j).LinO = ArmadurasHerrero(K).LinO - .LinO * 0.85
                        
                                                Exit For

                                        End If

                                Next K

                        End If

                End With

        Next i
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 5 Then ' @@ Miqueas Antes 17, reduccion de consumo muy hard
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim Count As Integer

        Dim i     As Long

        Dim j     As Long

        Dim K     As Long
    
        Count = Buffer.ReadInteger()
    
        ReDim ObjCarpintero(Count) As tItemsConstruibles
        ReDim CarpinteroMejorar(0) As tItemsConstruibles
    
        For i = 1 To Count

                With ObjCarpintero(i)
                        .ObjIndex = Buffer.ReadInteger()
                        .Name = getTypeObj(.ObjIndex, enum_TypeObj.Nombre) 'Buffer.ReadASCIIString()    'Get the object's name
                        .GrhIndex = Val(getTypeObj(.ObjIndex, enum_TypeObj.GrhIndex)) 'Buffer.ReadInteger()
 
                        .Madera = Val(getTypeObj(.ObjIndex, enum_TypeObj.Madera)) 'Buffer.ReadInteger()          'The wood needed
                        .MaderaElfica = Val(getTypeObj(.ObjIndex, enum_TypeObj.MaderaElfica)) 'Buffer.ReadInteger()    'The elfic wood needed
                        
                        .Upgrade = Val(getTypeObj(.ObjIndex, enum_TypeObj.Upgrade)) 'Buffer.ReadInteger()

                End With

        Next i
    
        For i = 1 To MAX_LIST_ITEMS
                Set InvMaderasCarpinteria(i) = New clsGrapchicalInventory
        Next i
    
        With frmCarp
                ' Inicializo los inventarios
                Call InvMaderasCarpinteria(1).Initialize(DirectDraw, .picMaderas0, 2, , , , , , False)
                Call InvMaderasCarpinteria(2).Initialize(DirectDraw, .picMaderas1, 2, , , , , , False)
                Call InvMaderasCarpinteria(3).Initialize(DirectDraw, .picMaderas2, 2, , , , , , False)
                Call InvMaderasCarpinteria(4).Initialize(DirectDraw, .picMaderas3, 2, , , , , , False)
        
                Call .HideExtraControls(Count)
                Call .RenderList(1)

        End With
    
        For i = 1 To Count

                With ObjCarpintero(i)

                        If .Upgrade Then

                                For K = 1 To Count

                                        If .Upgrade = ObjCarpintero(K).ObjIndex Then
                                                j = j + 1
                
                                                ReDim Preserve CarpinteroMejorar(j) As tItemsConstruibles
                        
                                                CarpinteroMejorar(j).Name = .Name
                                                CarpinteroMejorar(j).GrhIndex = .GrhIndex
                                                CarpinteroMejorar(j).ObjIndex = .ObjIndex
                                                CarpinteroMejorar(j).UpgradeName = ObjCarpintero(K).Name
                                                CarpinteroMejorar(j).UpgradeGrhIndex = ObjCarpintero(K).GrhIndex
                                                CarpinteroMejorar(j).Madera = ObjCarpintero(K).Madera - .Madera * 0.85
                                                CarpinteroMejorar(j).MaderaElfica = ObjCarpintero(K).MaderaElfica - .MaderaElfica * 0.85
                        
                                                Exit For

                                        End If

                                Next K

                        End If

                End With

        Next i
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserDescansar = Not UserDescansar

End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Call MsgBox(Buffer.ReadASCIIString())
    
        If frmConnect.Visible And (Not frmCrearPersonaje.Visible) Then
             
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserCiego = True

End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserEstupido = True

End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim tmp As String

        tmp = Buffer.ReadASCIIString()
    
        Call InitCartel(tmp, Buffer.ReadInteger())
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 9 Then ' @@ Miqueas Antes 22, reduccion de consumo muy hard
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim Slot   As Byte
        Dim oIndex As Integer

        Slot = Buffer.ReadByte()
    
        With NPCInventory(Slot)
        
                oIndex = Buffer.ReadInteger()
                
                .ObjIndex = oIndex
                .Amount = Buffer.ReadInteger()
                .Valor = Buffer.ReadSingle()
                .Copas = Buffer.ReadInteger
                
                
                .Name = getTypeObj(oIndex, enum_TypeObj.Nombre) 'Buffer.ReadASCIIString()
                .GrhIndex = Val(getTypeObj(oIndex, enum_TypeObj.GrhIndex)) 'Buffer.ReadInteger()
                .ObjType = Val(getTypeObj(oIndex, enum_TypeObj.ObjType)) 'Buffer.ReadByte()
                .MaxHit = Val(getTypeObj(oIndex, enum_TypeObj.MaxHit)) 'Buffer.ReadInteger()
                .MinHit = Val(getTypeObj(oIndex, enum_TypeObj.MinHit)) 'Buffer.ReadInteger()
                .MaxDef = Val(getTypeObj(oIndex, enum_TypeObj.MaxDef)) 'Buffer.ReadInteger()
                .MinDef = Val(getTypeObj(oIndex, enum_TypeObj.MinDef)) 'Buffer.ReadInteger

        End With
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the UpdateHungerAndThirst message.

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode

                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserMaxAGU = incomingData.ReadByte()
        UserMinAGU = incomingData.ReadByte()
        UserMaxHAM = incomingData.ReadByte()
        UserMinHAM = incomingData.ReadByte()
      
      
        
        frmMain.lblHambre = UserMinHAM & "%"
        frmMain.lblSed = UserMinAGU & "%"

        Dim bWidth As Byte
    
        bWidth = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 100)
    
        frmMain.ShpHambre.Width = bWidth
        '*********************************
    
        bWidth = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 100)
    
        frmMain.ShpAgua.Width = bWidth
    
End Sub

''
' Handles the Fame message.

Private Sub HandleFame()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 29 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        With UserReputacion
                .AsesinoRep = incomingData.ReadLong()
                .BandidoRep = incomingData.ReadLong()
                .BurguesRep = incomingData.ReadLong()
                .LadronesRep = incomingData.ReadLong()
                .NobleRep = incomingData.ReadLong()
                .PlebeRep = incomingData.ReadLong()
                .Promedio = incomingData.ReadLong()

        End With
    
        LlegoFama = True

End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 20 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        With UserEstadisticas
                .CiudadanosMatados = incomingData.ReadLong()
                .CriminalesMatados = incomingData.ReadLong()
                .UsuariosMatados = incomingData.ReadLong()
                .NpcsMatados = incomingData.ReadInteger()
                .Clase = ListaClases(incomingData.ReadByte())
                .PenaCarcel = incomingData.ReadLong()

        End With

End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 4 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
        
        Dim tmp As Byte
        
        tmp = incomingData.ReadByte()

        If (tmp = 1) Then
                SkillPoints = SkillPoints + incomingData.ReadInteger()
        Else
                SkillPoints = 0 + incomingData.ReadInteger()

        End If
    
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 8 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim ForumType As eForumMsgType

        Dim Title     As String

        Dim Message   As String

        Dim Author    As String
    
        ForumType = Buffer.ReadByte
    
        Title = Buffer.ReadASCIIString()
        Author = Buffer.ReadASCIIString()
        Message = Buffer.ReadASCIIString()
    
        If Not frmForo.ForoLimpio Then
                clsForos.ClearForums
                frmForo.ForoLimpio = True

        End If

        Call clsForos.AddPost(ForumAlignment(ForumType), Title, Author, Message, EsAnuncio(ForumType))
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        frmForo.Privilegios = incomingData.ReadByte
        frmForo.CanPostSticky = incomingData.ReadByte
    
        If Not MirandoForo Then
                frmForo.Show , frmMain

        End If

End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 4 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim CharIndex As Integer
    
        CharIndex = incomingData.ReadInteger()
        charlist(CharIndex).invisible = incomingData.ReadBoolean()

End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserMeditar = Not UserMeditar

End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserCiego = False

End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserEstupido = False

End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 11/19/09
        '11/19/09: Pato - Now the server send the percentage of progress of the skills.
        '***************************************************
        If incomingData.Length < 1 + NUMSKILLS Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim i As Long

        For i = 1 To NUMSKILLS
                UserSkills(i) = incomingData.ReadByte()
                'PorcentajeSkills(i) = incomingData.ReadByte()
        Next i
    
        LlegaronSkills = True

End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim creatures() As String

        Dim i           As Long
    
        creatures = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        For i = 0 To UBound(creatures())
                Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
        Next i

        frmEntrenador.Show , frmMain
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 11/19/09
        '11/19/09: Pato - Is optional show the frmGuildNews form
        '***************************************************
        If incomingData.Length < 7 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim guildList() As String

        Dim i           As Long

        Dim sTemp       As String
    
        'Get news' string
        frmGuildNews.news = Buffer.ReadASCIIString()
    
        'Get Enemy guilds list
        guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        For i = 0 To UBound(guildList)
                sTemp = frmGuildNews.txtClanesGuerra.Text
                frmGuildNews.txtClanesGuerra.Text = sTemp & guildList(i) & vbCrLf
        Next i
    
        'Get Allied guilds list
        guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        For i = 0 To UBound(guildList)
                sTemp = frmGuildNews.txtClanesAliados.Text
                frmGuildNews.txtClanesAliados.Text = sTemp & guildList(i) & vbCrLf
        Next i
    
        If ClientSetup.bGuildNews Or bShowGuildNews Then frmGuildNews.Show vbModeless, frmMain
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim vsGuildList() As String

        Dim i             As Long
    
        vsGuildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        Call frmPeaceProp.lista.Clear

        For i = 0 To UBound(vsGuildList())
                Call frmPeaceProp.lista.AddItem(vsGuildList(i))
        Next i
    
        frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
        Call frmPeaceProp.Show(vbModeless, frmMain)
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim guildList() As String

        Dim i           As Long
    
        guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        Call frmPeaceProp.lista.Clear

        For i = 0 To UBound(guildList())
                Call frmPeaceProp.lista.AddItem(guildList(i))
        Next i
    
        frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
        Call frmPeaceProp.Show(vbModeless, frmMain)
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 35 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        With frmCharInfo

                If .frmType = CharInfoFrmType.frmMembers Then
                        .imgRechazar.Visible = False
                        .imgAceptar.Visible = False
                        .imgEchar.Visible = True
                        .imgPeticion.Visible = False
                Else
                        .imgRechazar.Visible = True
                        .imgAceptar.Visible = True
                        .imgEchar.Visible = False
                        .imgPeticion.Visible = True

                End If
        
                .Nombre.Caption = Buffer.ReadASCIIString()
                .Raza.Caption = ListaRazas(Buffer.ReadByte())
                .Clase.Caption = ListaClases(Buffer.ReadByte())
        
                If Buffer.ReadByte() = 1 Then
                        .Genero.Caption = "Hombre"
                Else
                        .Genero.Caption = "Mujer"

                End If
        
                .Nivel.Caption = Buffer.ReadByte()
                .Oro.Caption = Buffer.ReadLong()
                .Banco.Caption = Buffer.ReadLong()
        
                Dim reputation As Long

                reputation = Buffer.ReadLong()
        
                .reputacion.Caption = reputation
        
                .txtPeticiones.Text = Buffer.ReadASCIIString()
                .guildactual.Caption = Buffer.ReadASCIIString()
                .txtMiembro.Text = Buffer.ReadASCIIString()
        
                Dim armada As Boolean

                Dim Caos   As Boolean
        
                armada = Buffer.ReadBoolean()
                Caos = Buffer.ReadBoolean()
        
                If armada Then
                        .ejercito.Caption = "Armada Real"
                ElseIf Caos Then
                        .ejercito.Caption = "Legi�n Oscura"

                End If
        
                .Ciudadanos.Caption = CStr(Buffer.ReadLong())
                .criminales.Caption = CStr(Buffer.ReadLong())
        
                If reputation > 0 Then
                        .status.Caption = " Ciudadano"
                        .status.ForeColor = vbBlue
                Else
                        .status.Caption = " Criminal"
                        .status.ForeColor = vbRed

                End If
        
                Call .Show(vbModeless, frmMain)

        End With
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 7 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim i      As Long

        Dim List() As String
    
        With frmGuildLeader
                'Get list of existing guilds
                Dim tmpStr As String
                On Error GoTo wtf
                tmpStr = Buffer.ReadASCIIString()
                
                GoTo 99
wtf:
                tmpStr = "No hay guilds"
                
99                GuildNames = Split(tmpStr, SEPARATOR)
        
                'Empty the list
                Call .guildslist.Clear
        
                For i = 0 To UBound(GuildNames())
                        Call .guildslist.AddItem(GuildNames(i))
                Next i
        
                'Get list of guild's members
                GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
                .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
                'Empty the list
                Call .members.Clear
        
                For i = 0 To UBound(GuildMembers())
                        Call .members.AddItem(GuildMembers(i))
                Next i
        
                .txtguildnews = Buffer.ReadASCIIString()
        
                'Get list of join requests
                List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
                'Empty the list
                Call .solicitudes.Clear
        
                For i = 0 To UBound(List())
                        Call .solicitudes.AddItem(List(i))
                Next i
        
                .Show , frmMain

        End With

        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 27 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        With frmGuildBrief
                .imgDeclararGuerra.Visible = .EsLeader
                .imgOfrecerAlianza.Visible = .EsLeader
                .imgOfrecerPaz.Visible = .EsLeader
        
                .Nombre.Caption = Buffer.ReadASCIIString()
                .fundador.Caption = Buffer.ReadASCIIString()
                .creacion.Caption = Buffer.ReadASCIIString()
                .lider.Caption = Buffer.ReadASCIIString()
                .web.Caption = Buffer.ReadASCIIString()
                .Miembros.Caption = Buffer.ReadInteger()
        
                If Buffer.ReadBoolean() Then
                        .eleccion.Caption = "ABIERTA"
                Else
                        .eleccion.Caption = "CERRADA"

                End If
        
                .lblAlineacion.Caption = Buffer.ReadASCIIString()
                .Enemigos.Caption = Buffer.ReadInteger()
                .Aliados.Caption = Buffer.ReadInteger()
                .antifaccion.Caption = Buffer.ReadASCIIString()
        
                Dim codexStr() As String

                Dim i          As Long
        
                codexStr = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
                For i = 0 To 7
                        .Codex(i).Caption = codexStr(i)
                Next i
        
                .Desc.Text = Buffer.ReadASCIIString()
                
                .GuildLevel.Caption = "Nivel: " & Buffer.ReadInteger
                
            
        End With
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
        frmGuildBrief.Show vbModeless, frmMain
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ShowGuildAlign message.

Private Sub HandleShowGuildAlign()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/12/2009
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        frmEligeAlineacion.Show vbModeless, frmMain

End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        CreandoClan = True
        frmGuildFoundation.Show , frmMain

End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        UserParalizado = Not UserParalizado

End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
        Call frmUserRequest.Show(vbModeless, frmMain)
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        If frmComerciar.Visible Then

                Dim i As Long
        
                'Update user inventory
                For i = 1 To MAX_INVENTORY_SLOTS

                        ' Agrego o quito un item en su totalidad
                        If Inventario.ObjIndex(i) <> InvComUsu.ObjIndex(i) Then

                                With Inventario
                                        Call InvComUsu.SetItem(i, .ObjIndex(i), _
                                           .Amount(i), .Equipped(i), .GrhIndex(i), _
                                           .ObjType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                                           .Valor(i), .ItemName(i))

                                End With

                                ' Vendio o compro cierta cantidad de un item que ya tenia
                        ElseIf Inventario.Amount(i) <> InvComUsu.Amount(i) Then
                                Call InvComUsu.ChangeSlotItemAmount(i, Inventario.Amount(i))

                        End If

                Next i
        
                ' Fill Npc inventory
                For i = 1 To 20

                        ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
                        If NPCInventory(i).ObjIndex <> InvComNpc.ObjIndex(i) Then

                                With NPCInventory(i)
                                        Call InvComNpc.SetItem(i, .ObjIndex, _
                                           .Amount, 0, .GrhIndex, _
                                           .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                                           .Valor, .Name)

                                End With

                                ' Compraron o vendieron cierta cantidad (no su totalidad)
                        ElseIf NPCInventory(i).Amount <> InvComNpc.Amount(i) Then
                                Call InvComNpc.ChangeSlotItemAmount(i, NPCInventory(i).Amount)

                        End If

                Next i
    
        End If

End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim i As Long
    
        If frmBancoObj.Visible Then
        
                For i = 1 To Inventario.MaxObjs

                        With Inventario
                                Call InvBanco(1).SetItem(i, .ObjIndex(i), .Amount(i), _
                                   .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), _
                                   .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))

                        End With

                Next i
        
                'Alter order according to if we bought or sold so the labels and grh remain the same
                If frmBancoObj.LasActionBuy Then
                        'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                        'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                Else

                        'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                        'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                End If
        
                frmBancoObj.NoPuedeMover = False

        End If
       
End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 22 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        Dim OfferSlot As Byte
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        OfferSlot = Buffer.ReadByte
    
        With Buffer

                If OfferSlot = GOLD_OFFER_SLOT Then
                        Call InvOroComUsu(2).SetItem(1, .ReadInteger(), .ReadLong(), 0, _
                           .ReadInteger(), .ReadByte(), .ReadInteger(), _
                           .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadLong(), .ReadASCIIString())
                Else
                        Call InvOfferComUsu(1).SetItem(OfferSlot, .ReadInteger(), .ReadLong(), 0, _
                           .ReadInteger(), .ReadByte(), .ReadInteger(), _
                           .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadLong(), .ReadASCIIString())

                End If

        End With
    
        Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " ha modificado su oferta.", FontTypeNames.FONTTYPE_VENENO)
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the SendNight message.

Private Sub HandleSendNight()

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/08/07
        '
        '***************************************************
        If incomingData.Length < 2 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
    
        Dim tBool As Boolean 'CHECK, este handle no hace nada con lo que recibe.. porque, ehmm.. no hay noche?.. o si?

        tBool = incomingData.ReadBoolean()

End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim creatureList() As String

        Dim i              As Long
    
        creatureList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        For i = 0 To UBound(creatureList())
                Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
        Next i

        frmSpawnList.Show , frmMain
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim sosList() As String

        Dim i         As Long
    
        sosList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        For i = 0 To UBound(sosList())
                Call frmMSG.List1.AddItem(sosList(i))
        Next i
    
        frmMSG.Show , frmMain
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ShowDenounces message.

Private Sub HandleShowDenounces()

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/11/2010
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim DenounceList() As String

        Dim DenounceIndex  As Long
    
        DenounceList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)

                For DenounceIndex = 0 To UBound(DenounceList())
                        Call AddtoRichTextBox(frmMain.RecTxt, DenounceList(DenounceIndex), .red, .green, .blue, .Bold, .Italic)
                Next DenounceIndex

        End With
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()

        '***************************************************
        'Author: Budi
        'Last Modification: 11/26/09
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim members() As String

        Dim i         As Long
    
        EsPartyLeader = CBool(Buffer.ReadByte())
       
        members = Split(Buffer.ReadASCIIString(), SEPARATOR)

        For i = 0 To UBound(members())
                Call frmParty.lstMembers.AddItem(members(i))
        Next i
    
        frmParty.lblTotalExp.Caption = Buffer.ReadLong
        frmParty.Show , frmMain
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '*************************************Su**************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim tmpStr As String

        tmpStr = Buffer.ReadASCIIString()
    
        'frmCambiaMotd.txtMotd.Text =
        'frmCambiaMotd.Show , frmMain
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
    
        frmPanelGm.Show vbModeless, frmMain

End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim userList() As String

        Dim i          As Long
    
        userList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
        If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear

                For i = 0 To UBound(userList())
                        Call frmPanelGm.cboListaUsus.AddItem(userList(i))
                Next i

                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.listIndex = 0

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Call incomingData.ReadByte
    
        Call AddtoRichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, True)
    
        pingTime = 0

End Sub

''
' Handles the Pong message.

Private Sub HandleGuildMemberInfo()

        '***************************************************
        'Author: ZaMa
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 3 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        With frmGuildMember
                'Clear guild's list
                .lstClanes.Clear
        
                GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
                Dim i As Long

                For i = 0 To UBound(GuildNames())
                        Call .lstClanes.AddItem(GuildNames(i))
                Next i
        
                'Get list of guild's members
                GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
                .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
                'Empty the list
                Call .lstMiembros.Clear
        
                For i = 0 To UBound(GuildMembers())
                        Call .lstMiembros.AddItem(GuildMembers(i))
                Next i
        
                'If we got here then packet is complete, copy data back to original queue
                Call incomingData.CopyBuffer(Buffer)
        
                .Show vbModeless, frmMain

        End With
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        If incomingData.Length < 6 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim CharIndex As Integer

        Dim NickColor As Byte

        Dim UserTag   As String
    
        CharIndex = Buffer.ReadInteger()
        NickColor = Buffer.ReadByte()
        UserTag = Buffer.ReadASCIIString()
    
        'Update char status adn tag!
        With charlist(CharIndex)

                If (NickColor And eNickColor.ieCriminal) <> 0 Then
                        .Criminal = 1
                Else
                        .Criminal = 0

                End If
        
                .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
                .Nombre = UserTag

        End With
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "LoginExistingChar" message to the outgoing data buffer
        '***************************************************

        With outgoingData
                Call .WriteByte(ClientPacketID.LoginExistingChar)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(UserPassword)
        
                Call .WriteByte(App.Major)
                Call .WriteByte(App.Minor)
                Call .WriteByte(App.Revision)

                Call .WriteASCIIString(GetSerialHD())    ' GSZAO
                 If Cpu_SSN = vbNullString And Cpu_ID = vbNullString Then Cpu_SSN = "ERROR"
                Call .WriteASCIIString(Cpu_SSN & "#" & Cpu_ID)
        End With

End Sub


Private Function SystemDrive() As String
' GSZ-AO - Obtiene la unidad del sistema
    Dim windows_dir As String
    Dim Length As Long
    windows_dir = Space$(255)
    Length = GetWindowsDirectory(windows_dir, Len(windows_dir))
    SystemDrive = Left$(windows_dir, 3)    ' C:\
End Function

Function GetSerialHD() As Long
' GSZ-AO - Obtiene el numero de serie del disco de sistema
    Dim SerialNum As Long
    Dim res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    res = GetVolumeInformation(SystemDrive(), Temp1, _
                               Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialHD = SerialNum
End Function

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "LoginNewChar" message to the outgoing data buffer
        '***************************************************
    
        With outgoingData
                Call .WriteByte(ClientPacketID.LoginNewChar)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(UserPassword)
        
                Call .WriteByte(App.Major)
                Call .WriteByte(App.Minor)
                Call .WriteByte(App.Revision)
        
                Call .WriteByte(UserRaza)
                Call .WriteByte(UserSexo)
                Call .WriteByte(UserClase)
                Call .WriteInteger(UserHead)
        
                Call .WriteASCIIString(UserEmail)
        
                Call .WriteByte(UserHogar)
                Call .WriteASCIIString(UserClave)
                Call .WriteASCIIString(GetSerialHD())    ' GSZAO
                ' @@ CUICUI SECURITY
                If Cpu_SSN = vbNullString And Cpu_ID = vbNullString Then Cpu_SSN = "ERROR"
                Call .WriteASCIIString(Cpu_SSN & "#" & Cpu_ID)
        End With

End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Talk" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Talk)
        
                Call .WriteASCIIString(chat)

        End With

End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Yell" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Yell)
        
                Call .WriteASCIIString(chat)

        End With

End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharName As String, ByVal chat As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 03/12/10
        'Writes the "Whisper" message to the outgoing data buffer
        '03/12/10: Enanoh - Ahora se env�a el nick y no el charindex.
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Whisper)
        
                Call .WriteASCIIString(CharName)
        
                Call .WriteASCIIString(chat)

        End With

End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Walk" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Walk)
        
                Call .WriteByte(Heading)

        End With

End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestPositionUpdate" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)

End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Attack" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Attack)

End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PickUp" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.PickUp)

End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SafeToggle" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.SafeToggle)

End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
        '**************************************************************
        'Author: Rapsodius
        'Creation Date: 10/10/07
        'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
        '**************************************************************
        Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)

End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)

End Sub

Public Sub WriteRequestPartyForm()
        '***************************************************
        'Author: Budi
        'Last Modification: 11/26/09
        'Writes the "RequestPartyForm" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestPartyForm)

End Sub

''
' Writes the "ItemUpgrade" message to the outgoing data buffer.
'
' @param    ItemIndex The index to the item to upgrade.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 12/09/09
        'Writes the "ItemUpgrade" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
        Call outgoingData.WriteInteger(ItemIndex)

End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestAtributes" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)

End Sub

''
' Writes the "RequestFame" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestFame()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestFame" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestFame)

End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestSkills" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestSkills)

End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestMiniStats" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)

End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CommerceEnd" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)

End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UserCommerceEnd" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)

End Sub

''
' Writes the "UserCommerceConfirm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceConfirm()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/12/2009
        'Writes the "UserCommerceConfirm" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)

End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankEnd" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.BankEnd)

End Sub


''
' Writes the "CheckHD" message to the outgoing data buffer.
'
'@param   username The name of the user to be checked.
Public Sub WriteVerHD(ByVal UserName As String)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Checkeamos la HD del usuario.
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.VerHD)

        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "UnBanHD" message to the outgoing data buffer
'
'@param    username The name of the user to be unbanned.
Public Sub WriteUnBanHD(ByVal HD As String)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Unbaneamos al usuario con su HD baneado.
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnBanHD)

        Call .WriteASCIIString(HD)
    End With
End Sub

''
' Writes the "BanHD" message to the outgoing data buffer
'
'@param    username The name of the user to be banned.
Public Sub WriteBanHD(ByVal UserName As String)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Baneamos la HD del usuario.
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanHD)

        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/10/07
        'Writes the "UserCommerceOk" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)

End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UserCommerceReject" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)

End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Drop" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Drop)
        
                Call .WriteByte(Slot)
                Call .WriteInteger(Amount)

        End With

End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal Slot As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CastSpell" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.CastSpell)
        
                Call .WriteByte(Slot)

        End With

End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "LeftClick" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.LeftClick)
        
                Call .WriteByte(X)
                Call .WriteByte(Y)

        End With

End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "DoubleClick" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.DoubleClick)
        
                Call .WriteByte(X)
                Call .WriteByte(Y)

        End With

End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Work" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Work)
        
                Call .WriteByte(Skill)

        End With

End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UseSpellMacro" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)

End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal Slot As Byte, ByVal bClick As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UseItem" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.UseItem)
        
                Call .WriteByte(Slot)
                Call .WriteByte(bClick)

        End With

End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal Item As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CraftBlacksmith" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
                Call .WriteInteger(Item)

        End With

End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal Item As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CraftCarpenter" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.CraftCarpenter)
        
                Call .WriteInteger(Item)

        End With

End Sub

''
' Writes the "ShowGuildNews" message to the outgoing data buffer.
'

Public Sub WriteShowGuildNews()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 21/02/2010
        'Writes the "ShowGuildNews" message to the outgoing data buffer
        '***************************************************
 
        outgoingData.WriteByte (ClientPacketID.ShowGuildNews)

End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "WorkLeftClick" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.WorkLeftClick)
        
                Call .WriteByte(X)
                Call .WriteByte(Y)
        
                Call .WriteByte(Skill)

        End With

End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, _
                               ByVal Name As String, _
                               ByVal Site As String, _
                               ByRef Codex() As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CreateNewGuild" message to the outgoing data buffer
        '***************************************************
        Dim Temp As String

        Dim i    As Long
    
        With outgoingData
                Call .WriteByte(ClientPacketID.CreateNewGuild)
        
                Call .WriteASCIIString(Desc)

                Call .WriteASCIIString(Name)
                Call .WriteASCIIString(Site)
        
                For i = LBound(Codex()) To UBound(Codex())
                        Temp = Temp & Codex(i) & SEPARATOR
                Next i
        
                If Len(Temp) Then _
                   Temp = Left$(Temp, Len(Temp) - 1)
        
                Call .WriteASCIIString(Temp)

        End With

End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal Slot As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SpellInfo" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.SpellInfo)
        
                Call .WriteByte(Slot)

        End With

End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "EquipItem" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.EquipItem)
        
                Call .WriteByte(Slot)

        End With

End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChangeHeading" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.ChangeHeading)
        
                Call .WriteByte(Heading)

        End With

End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ModifySkills" message to the outgoing data buffer
        '***************************************************
        Dim i As Long
    
        With outgoingData
                Call .WriteByte(ClientPacketID.ModifySkills)
        
                For i = 1 To NUMSKILLS
                        Call .WriteByte(skillEdt(i))
                Next i

        End With

End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Train" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Train)
        
                Call .WriteByte(creature)

        End With

End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CommerceBuy" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.CommerceBuy)
        
                Call .WriteByte(Slot)
                Call .WriteInteger(Amount)

        End With

End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankExtractItem" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.BankExtractItem)
        
                Call .WriteByte(Slot)
                Call .WriteInteger(Amount)

        End With

End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CommerceSell" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.CommerceSell)
        
                Call .WriteByte(Slot)
                Call .WriteInteger(Amount)

        End With

End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankDeposit" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.BankDeposit)
        
                Call .WriteByte(Slot)
                Call .WriteInteger(Amount)

        End With

End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, _
                          ByVal Message As String, _
                          ByVal ForumMsgType As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ForumPost" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.ForumPost)
        
                Call .WriteByte(ForumMsgType)
                Call .WriteASCIIString(Title)
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "MoveSpell" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.MoveSpell)
        
                Call .WriteBoolean(upwards)
                Call .WriteByte(Slot)

        End With

End Sub

''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal Slot As Byte)

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 06/14/09
        'Writes the "MoveBank" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.MoveBank)
        
                Call .WriteBoolean(upwards)
                Call .WriteByte(Slot)

        End With

End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ClanCodexUpdate" message to the outgoing data buffer
        '***************************************************
        Dim Temp As String

        Dim i    As Long
    
        With outgoingData
                Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
                Call .WriteASCIIString(Desc)
        
                For i = LBound(Codex()) To UBound(Codex())
                        Temp = Temp & Codex(i) & SEPARATOR
                Next i
        
                If Len(Temp) Then _
                   Temp = Left$(Temp, Len(Temp) - 1)
        
                Call .WriteASCIIString(Temp)

        End With

End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, _
                                  ByVal Amount As Long, _
                                  ByVal OfferSlot As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UserCommerceOffer" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
                Call .WriteByte(Slot)
                Call .WriteLong(Amount)
                Call .WriteByte(OfferSlot)

        End With

End Sub

Public Sub WriteCommerceChat(ByVal chat As String)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 03/12/2009
        'Writes the "CommerceChat" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.CommerceChat)
        
                Call .WriteASCIIString(chat)

        End With

End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildAcceptPeace" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildRejectAlliance" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildRejectPeace" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildOfferPeace" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
                Call .WriteASCIIString(guild)
                Call .WriteASCIIString(proposal)

        End With

End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildOfferAlliance" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
                Call .WriteASCIIString(guild)
                Call .WriteASCIIString(proposal)

        End With

End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildAllianceDetails" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildPeaceDetails" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildAlliancePropList" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)

End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildPeacePropList" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)

End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildDeclareWar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal URL As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildNewWebsite" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
                Call .WriteASCIIString(URL)

        End With

End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildRejectNewMember" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(reason)

        End With

End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildKickMember" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildKickMember)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildUpdateNews" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
                Call .WriteASCIIString(news)

        End With

End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildMemberInfo" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildOpenElections" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)

End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildRequestMembership" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
                Call .WriteASCIIString(guild)
                Call .WriteASCIIString(Application)

        End With

End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildRequestDetails" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Online" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Online)

End Sub


Public Sub WritePoder()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Online" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Poder)

End Sub

Public Sub WriteCastillo()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Online" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Castillo)

End Sub


''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 08/16/08
        'Writes the "Quit" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Quit)

End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildLeave" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GuildLeave)

End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestAccountState" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)

End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PetStand" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.PetStand)

End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PetFollow" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.PetFollow)

End Sub

''
' Writes the "ReleasePet" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReleasePet()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 18/11/2009
        'Writes the "ReleasePet" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.ReleasePet)

End Sub

''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "TrainList" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.TrainList)

End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Rest" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Rest)

End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Meditate" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Meditate)

End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Resucitate" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Resucitate)

End Sub

''
' Writes the "Consultation" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsultation()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 01/05/2010
        'Writes the "Consultation" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Consultation)

End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Heal" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Heal)

End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Help" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Help)

End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestStats" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestStats)

End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CommerceStart" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.CommerceStart)

End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankStart" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.BankStart)

End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Enlist" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Enlist)

End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Information" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Information)

End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Reward" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Reward)

End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMOTD()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestMOTD" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)

End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpTime" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Uptime)

End Sub

''
' Writes the "PartyLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PartyLeave" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.PartyLeave)

End Sub

''
' Writes the "PartyCreate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyCreate()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PartyCreate" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.PartyCreate)

End Sub

''
' Writes the "PartyJoin" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyJoin()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PartyJoin" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.PartyJoin)

End Sub

''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry(Optional ByVal i As Integer = "1")
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Inquiry" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.Inquiry)
        Call outgoingData.WriteInteger(i)
        
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildRequestDetails" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "PartyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PartyMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.PartyMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal number As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CentinelReport" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.CentinelReport)
        
                Call .WriteInteger(number)

        End With

End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildOnline" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GuildOnline)

End Sub

''
' Writes the "PartyOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyOnline()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PartyOnline" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.PartyOnline)

End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CouncilMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.CouncilMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RoleMasterRequest" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GMRequest" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMRequest)

End Sub

''
' Writes the "BugReport" message to the outgoing data buffer.
'
' @param    message The message explaining the reported bug.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBugReport(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BugReport" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.bugReport)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChangeDescription" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.ChangeDescription)
        
                Call .WriteASCIIString(Desc)

        End With

End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildVote" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildVote)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Punishments" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Punishments)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 10/10/07
        'Last Modified By: Rapsodius
        'Writes the "ChangePassword" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.ChangePassword)
         
                Call .WriteASCIIString(oldPass)
                Call .WriteASCIIString(newPass)

        End With

End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Gamble" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Gamble)
        
                Call .WriteInteger(Amount)

        End With

End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "InquiryVote" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.InquiryVote)
        
                Call .WriteByte(opt)

        End With

End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "LeaveFaction" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)

End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankExtractGold" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.BankExtractGold)
        
                Call .WriteLong(Amount)

        End With

End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BankDepositGold" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.BankDepositGold)
        
                Call .WriteLong(Amount)

        End With

End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Denounce" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Denounce)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 03/21/2001
        'Writes the "GuildFundate" message to the outgoing data buffer
        '14/12/2009: ZaMa - Now first checks if the user can foundate a guild.
        '03/21/2001: Pato - Deleted de clanType param.
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GuildFundate)

End Sub

''
' Writes the "GuildFundation" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundation(ByVal clanType As eClanType)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/12/2009
        'Writes the "GuildFundation" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GuildFundation)
        
                Call .WriteByte(clanType)

        End With

End Sub

''
' Writes the "PartyKick" message to the outgoing data buffer.
'
' @param    username The user to kick fro mthe party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PartyKick" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.PartyKick)
            
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "PartySetLeader" message to the outgoing data buffer.
'
' @param    username The user to set as the party's leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySetLeader(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PartySetLeader" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.PartySetLeader)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "PartyAcceptMember" message to the outgoing data buffer.
'
' @param    username The user to accept into the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyAcceptMember(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "PartyAcceptMember" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.PartyAcceptMember)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildMemberList" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.GuildMemberList)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "InitCrafting" message to the outgoing data buffer.
'
' @param    Cantidad The final aumont of item to craft.
' @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As Integer)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/01/2010
        'Writes the "InitCrafting" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.InitCrafting)
                Call .WriteLong(cantidad)
        
                Call .WriteInteger(NroPorCiclo)

        End With

End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()

        '***************************************************
        'Author: Budi
        'Last Modification: 01/06/10
        'Writes the "Home" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.Home)

        End With

End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GMMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.GMMessage)
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ShowName" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.showName)

End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.OnlineRoyalArmy)

End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "OnlineChaosLegion" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.OnlineChaosLegion)

End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GoNearby" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call outgoingData.WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.GoNearby)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Comment" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.Comment)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ServerTime" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.serverTime)

End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Where" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.Where)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CreaturesInMap" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CreaturesInMap)
        
                Call .WriteInteger(Map)

        End With

End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "WarpMeToTarget" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.WarpMeToTarget)

End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, _
                         ByVal Map As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "WarpChar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.WarpChar)
        
                Call .WriteASCIIString(UserName)
        
                Call .WriteInteger(Map)
        
                Call .WriteByte(X)
                Call .WriteByte(Y)

        End With

End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Silence" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.Silence)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SOSShowList" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.SOSShowList)

End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SOSRemove" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SOSRemove)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GoToChar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.GoToChar)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "invisible" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.invisible)

End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GMPanel" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.GMPanel)

End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestUserList" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.RequestUserList)

End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Working" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.Working)

End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Hiding" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.Hiding)

End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal time As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Jail" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.Jail)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(reason)
        
                Call .WriteByte(time)

        End With

End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "KillNPC" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.KillNPC)

End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "WarnUser" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.WarnUser)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(reason)

        End With

End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, _
                         ByVal EditOption As eEditOptions, _
                         ByVal arg1 As String, _
                         ByVal arg2 As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "EditChar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.EditChar)
        
                Call .WriteASCIIString(UserName)
        
                Call .WriteByte(EditOption)
        
                Call .WriteASCIIString(arg1)
                Call .WriteASCIIString(arg2)

        End With

End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestCharInfo" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RequestCharInfo)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestCharStats" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RequestCharStats)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestCharGold" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RequestCharGold)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestCharInventory" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RequestCharInventory)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestCharBank" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RequestCharBank)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestCharSkills" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RequestCharSkills)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ReviveChar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ReviveChar)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "OnlineGM" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.OnlineGM)

End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 26/03/2009
        'Writes the "OnlineMap" message to the outgoing data buffer
        '26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.OnlineMap)
        
                Call .WriteInteger(Map)

        End With

End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Forgive" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.Forgive)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Kick" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.Kick)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Execute" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.Execute)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BanChar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.banChar)
        
                Call .WriteASCIIString(UserName)
        
                Call .WriteASCIIString(reason)

        End With

End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UnbanChar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.UnbanChar)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "NPCFollow" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.NPCFollow)

End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SummonChar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SummonChar)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SpawnListRequest" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.SpawnListRequest)

End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SpawnCreature" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SpawnCreature)
        
                Call .WriteInteger(creatureIndex)

        End With

End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ResetNPCInventory" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)

End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CleanWorld" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.CleanWorld)

End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ServerMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ServerMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMapMessage(ByVal Message As String)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/11/2010
        'Writes the "MapMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.MapMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "NickToIP" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.nickToIP)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "IPToNick" message to the outgoing data buffer
        '***************************************************
        If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
        Dim i As Long
    
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.IPToNick)
        
                For i = LBound(Ip()) To UBound(Ip())
                        Call .WriteByte(Ip(i))
                Next i

        End With

End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildOnlineMembers" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.GuildOnlineMembers)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, _
                               ByVal X As Byte, _
                               ByVal Y As Byte, _
                               Optional ByVal Radio As Byte = 0)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "TeleportCreate" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.TeleportCreate)
        
                Call .WriteInteger(Map)
        
                Call .WriteByte(X)
                Call .WriteByte(Y)
        
                Call .WriteByte(Radio)

        End With

End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "TeleportDestroy" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)

End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RainToggle" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.RainToggle)

End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SetCharDescription" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SetCharDescription)
        
                Call .WriteASCIIString(Desc)

        End With

End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ForceMIDIToMap" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ForceMIDIToMap)
        
                Call .WriteByte(midiID)
        
                Call .WriteInteger(Map)

        End With

End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, _
                               ByVal Map As Integer, _
                               ByVal X As Byte, _
                               ByVal Y As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ForceWAVEToMap" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ForceWAVEToMap)
        
                Call .WriteByte(waveID)
        
                Call .WriteInteger(Map)
        
                Call .WriteByte(X)
                Call .WriteByte(Y)

        End With

End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RoyalArmyMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RoyalArmyMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChaosLegionMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChaosLegionMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CitizenMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CitizenMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CriminalMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CriminalMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "TalkAsNPC" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.TalkAsNPC)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)

End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.AcceptRoyalCouncilMember)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.AcceptChaosCouncilMember)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ItemsInTheFloor" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ItemsInTheFloor)

End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "MakeDumb" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.MakeDumb)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "MakeDumbNoMore" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.MakeDumbNoMore)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "DumpIPTables" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.dumpIPTables)

End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CouncilKick" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CouncilKick)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SetTrigger" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SetTrigger)
        
                Call .WriteByte(Trigger)

        End With

End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 04/13/07
        'Writes the "AskTrigger" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.AskTrigger)

End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BannedIPList" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.BannedIPList)

End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BannedIPReload" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.BannedIPReload)

End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "GuildBan" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.GuildBan)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, _
                      ByRef Ip() As Byte, _
                      ByVal Nick As String, _
                      ByVal reason As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "BanIP" message to the outgoing data buffer
        '***************************************************
        If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
        Dim i As Long
    
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.BanIP)
        
                Call .WriteBoolean(byIp)
        
                If byIp Then

                        For i = LBound(Ip()) To UBound(Ip())
                                Call .WriteByte(Ip(i))
                        Next i

                Else
                        Call .WriteASCIIString(Nick)

                End If
        
                Call .WriteASCIIString(reason)

        End With

End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UnbanIP" message to the outgoing data buffer
        '***************************************************
        If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
        Dim i As Long
    
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.UnbanIP)
        
                For i = LBound(Ip()) To UBound(Ip())
                        Call .WriteByte(Ip(i))
                Next i

        End With

End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal Amount As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CreateItem" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CreateItem)
                Call .WriteInteger(ItemIndex)
                Call .WriteInteger(Amount)

        End With

End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "DestroyItems" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.DestroyItems)

End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChaosLegionKick" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChaosLegionKick)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RoyalArmyKick" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RoyalArmyKick)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ForceMIDIAll" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ForceMIDIAll)
        
                Call .WriteByte(midiID)

        End With

End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ForceWAVEAll" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ForceWAVEAll)
        
                Call .WriteByte(waveID)

        End With

End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, _
                                 ByVal punishment As Byte, _
                                 ByVal NewText As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RemovePunishment" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RemovePunishment)
        
                Call .WriteASCIIString(UserName)
                Call .WriteByte(punishment)
                Call .WriteASCIIString(NewText)

        End With

End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "TileBlockedToggle" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)

End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)

End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)

End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "LastIP" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.LastIP)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChangeMOTD" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ChangeMOTD)

End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SetMOTD" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SetMOTD)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SystemMessage" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SystemMessage)
        
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NpcIndex As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CreateNPC" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CreateNPC)
        
                Call .WriteInteger(NpcIndex)

        End With

End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NpcIndex As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CreateNPCWithRespawn)
        
                Call .WriteInteger(NpcIndex)

        End With

End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ImperialArmour" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ImperialArmour)
        
                Call .WriteByte(armourIndex)
        
                Call .WriteInteger(objectIndex)

        End With

End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChaosArmour" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChaosArmour)
        
                Call .WriteByte(armourIndex)
        
                Call .WriteInteger(objectIndex)

        End With

End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "NavigateToggle" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.NavigateToggle)

End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)

End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "TurnOffServer" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.TurnOffServer)

End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "TurnCriminal" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.TurnCriminal)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ResetFactions" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ResetFactions)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RemoveCharFromGuild)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RequestCharMail" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RequestCharMail)
        
                Call .WriteASCIIString(UserName)

        End With

End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "AlterPassword" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.AlterPassword)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(CopyFrom)

        End With

End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "AlterMail" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.AlterMail)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(newMail)

        End With

End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "AlterName" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.AlterName)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(newName)

        End With

End Sub

''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ToggleCentinelActivated)

End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "DoBackup" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.DoBackUp)

End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ShowGuildMessages" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ShowGuildMessages)
        
                Call .WriteASCIIString(guild)

        End With

End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SaveMap" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.SaveMap)

End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoPK)
        
                Call .WriteBoolean(isPK)

        End With

End Sub

''
' Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer.
'
' @param    PermitirOcultar True if the map permits to hide, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoOcultar(ByVal PermitirOcultar As Boolean)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 19/09/2010
        'Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoNoOcultar)
        
                Call .WriteBoolean(PermitirOcultar)

        End With

End Sub

''
' Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer.
'
' @param    PermitirInvocar True if the map permits to invoke, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvocar(ByVal PermitirInvocar As Boolean)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 18/09/2010
        'Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoNoInvocar)
        
                Call .WriteBoolean(PermitirInvocar)

        End With

End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoBackup)
        
                Call .WriteBoolean(backup)

        End With

End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoRestricted)
        
                Call .WriteASCIIString(restrict)

        End With

End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoNoMagic)
        
                Call .WriteBoolean(nomagic)

        End With

End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoNoInvi)
        
                Call .WriteBoolean(noinvi)

        End With

End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoNoResu)
        
                Call .WriteBoolean(noresu)

        End With

End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoLand)
        
                Call .WriteASCIIString(land)

        End With

End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoZone)
        
                Call .WriteASCIIString(zone)

        End With

End Sub

''
' Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer.
'
' @param    forbid TRUE if stealNpc forbiden.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoStealNpc(ByVal forbid As Boolean)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 25/07/2010
        'Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChangeMapInfoStealNpc)
        
                Call .WriteBoolean(forbid)

        End With

End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "SaveChars" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.SaveChars)

End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CleanSOS" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.CleanSOS)

End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ShowServerForm" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ShowServerForm)

End Sub

''
' Writes the "ShowDenouncesList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenouncesList()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/11/2010
        'Writes the "ShowDenouncesList" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ShowDenouncesList)

End Sub

''
' Writes the "EnableDenounces" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnableDenounces()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/11/2010
        'Writes the "EnableDenounces" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.EnableDenounces)

End Sub

''
' Writes the "Night" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNight()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Night" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.night)

End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "KickAllChars" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.KickAllChars)

End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ReloadNPCs" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)

End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ReloadServerIni" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)

End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ReloadSpells" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ReloadSpells)

End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ReloadObjects" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ReloadObjects)

End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Restart" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.Restart)

End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ResetAutoUpdate" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.ResetAutoUpdate)

End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal B As Byte)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ChatColor" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.ChatColor)
        
                Call .WriteByte(r)
                Call .WriteByte(g)
                Call .WriteByte(B)

        End With

End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "Ignored" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.Ignored)

End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'Writes the "CheckSlot" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CheckSlot)
                Call .WriteASCIIString(UserName)
                Call .WriteByte(Slot)

        End With

End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 26/01/2007
        'Writes the "Ping" message to the outgoing data buffer
        '***************************************************
        'Prevent the timer from being cut
        If pingTime <> 0 Then Exit Sub
    
        Call outgoingData.WriteByte(ClientPacketID.Ping)
    
        ' Avoid computing errors due to frame rate
        Call FlushBuffer
        DoEvents
    
        pingTime = GetTickCount

End Sub

''
' Writes the "ShareNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShareNpc()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 15/04/2010
        'Writes the "ShareNpc" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.ShareNpc)

End Sub

''
' Writes the "StopSharingNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteStopSharingNpc()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 15/04/2010
        'Writes the "StopSharingNpc" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.StopSharingNpc)

End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, _
                          ByRef sClave As String, _
                          ByRef sValor As String)

        '***************************************************
        'Author: Brian Chaia (BrianPr)
        'Last Modification: 21/06/2009
        'Writes the "SetIniVar" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SetIniVar)
        
                Call .WriteASCIIString(sLlave)
                Call .WriteASCIIString(sClave)
                Call .WriteASCIIString(sValor)

        End With

End Sub

''
' Writes the "CreatePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map in which create the pretorian clan.
' @param    X           The x pos where the king is settled.
' @param    Y           The y pos where the king is settled.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreatePretorianClan(ByVal Map As Integer, _
                                    ByVal X As Byte, _
                                    ByVal Y As Byte)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/10/2010
        'Writes the "CreatePretorianClan" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.CreatePretorianClan)
                Call .WriteInteger(Map)
                Call .WriteByte(X)
                Call .WriteByte(Y)

        End With

End Sub

''
' Writes the "DeletePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map which contains the pretorian clan to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDeletePretorianClan(ByVal Map As Integer)

        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/10/2010
        'Writes the "DeletePretorianClan" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RemovePretorianClan)
                Call .WriteInteger(Map)

        End With

End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Sends all data existing in the buffer
        '***************************************************
        Dim sndData As String
    
        With outgoingData

                If .Length = 0 Then _
                   Exit Sub
        
                sndData = .ReadASCIIStringFixed(.Length)
        
                Call SendData(sndData)

        End With

End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
        'No enviamos nada si no estamos conectados
     
        If Not frmMain.Socket1.IsWritable Then
                'Put data back in the bytequeue
                Call outgoingData.WriteASCIIStringFixed(sdData)
        
                Exit Sub

        End If
    
        If Not frmMain.Socket1.Connected Then Exit Sub
    
        'Send data!
  
        Call frmMain.Socket1.Write(sdData, Len(sdData))

End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    Dialog The new dialog of the NPC.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetDialog(ByVal dialog As String)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 18/11/2010
        'Writes the "SetDialog" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SetDialog)
        
                Call .WriteASCIIString(dialog)

        End With

End Sub

''
' Writes the "Impersonate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImpersonate()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 20/11/2010
        'Writes the "Impersonate" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.Impersonate)

End Sub

''
' Writes the "Imitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImitate()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 20/11/2010
        'Writes the "Imitate" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.Imitate)

End Sub

''
' Writes the "RecordAddObs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAddObs(ByVal RecordIndex As Byte, ByVal Observation As String)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        'Writes the "RecordAddObs" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RecordAddObs)
        
                Call .WriteByte(RecordIndex)
                Call .WriteASCIIString(Observation)

        End With

End Sub

''
' Writes the "RecordAdd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAdd(ByVal Nickname As String, ByVal reason As String)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        'Writes the "RecordAdd" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RecordAdd)
        
                Call .WriteASCIIString(Nickname)
                Call .WriteASCIIString(reason)

        End With

End Sub

''
' Writes the "RecordRemove" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordRemove(ByVal RecordIndex As Byte)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        'Writes the "RecordRemove" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RecordRemove)
        
                Call .WriteByte(RecordIndex)

        End With

End Sub

''
' Writes the "RecordListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordListRequest()
        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        'Writes the "RecordListRequest" message to the outgoing data buffer
        '***************************************************
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.RecordListRequest)

End Sub

''
' Writes the "RecordDetailsRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetailsRequest(ByVal RecordIndex As Byte)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        'Writes the "RecordDetailsRequest" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.RecordDetailsRequest)
        
                Call .WriteByte(RecordIndex)

        End With

End Sub

''
' Handles the RecordList message.

Private Sub HandleRecordList()

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        '
        '***************************************************
        If incomingData.Length < 2 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim NumRecords As Byte

        Dim i          As Long
    
        NumRecords = Buffer.ReadByte
    
        'Se limpia el ListBox y se agregan los usuarios
        frmPanelGm.lstUsers.Clear

        For i = 1 To NumRecords
                frmPanelGm.lstUsers.AddItem Buffer.ReadASCIIString
        Next i
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Handles the RecordDetails message.

Private Sub HandleRecordDetails()

        '***************************************************
        'Author: Amraphen
        'Last Modification: 29/11/2010
        '
        '***************************************************
        If incomingData.Length < 2 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Dim tmpStr As String

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
       
        With frmPanelGm
                .txtCreador.Text = Buffer.ReadASCIIString
                .txtDescrip.Text = Buffer.ReadASCIIString
        
                'Status del pj
                If Buffer.ReadBoolean Then
                        .lblEstado.ForeColor = vbGreen
                        .lblEstado.Caption = "ONLINE"
                Else
                        .lblEstado.ForeColor = vbRed
                        .lblEstado.Caption = "OFFLINE"

                End If
        
                'IP del personaje
                tmpStr = Buffer.ReadASCIIString

                If LenB(tmpStr) Then
                        .txtIP.Text = tmpStr
                Else
                        .txtIP.Text = "Usuario offline"

                End If
        
                'Tiempo online
                tmpStr = Buffer.ReadASCIIString

                If LenB(tmpStr) Then
                        .txtTimeOn.Text = tmpStr
                Else
                        .txtTimeOn.Text = "Usuario offline"

                End If
        
                'Observaciones
                tmpStr = Buffer.ReadASCIIString

                If LenB(tmpStr) Then
                        .txtObs.Text = tmpStr
                Else
                        .txtObs.Text = "Sin observaciones"

                End If

        End With
    
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
    
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub

''
' Writes the "Moveitem" message to the outgoing data buffer.
'
Public Sub WriteMoveItem(ByVal originalSlot As Integer, _
                         ByVal newSlot As Integer, _
                         ByVal moveType As eMoveType)

        '***************************************************
        'Author: Budi
        'Last Modification: 05/01/2011
        'Writes the "MoveItem" message to the outgoing data buffer
        '***************************************************
        With outgoingData
                Call .WriteByte(ClientPacketID.MoveItem)
                Call .WriteByte(originalSlot)
                Call .WriteByte(newSlot)
                Call .WriteByte(moveType)

        End With

End Sub

''
' Handles the StrDextRunningOut message.

Private Sub HandleStrDextRunningOut()
        '***************************************************
        'Author: CHOTS
        'Last Modification: 08/06/2010
        '
        '***************************************************
        'Remove packet ID
        Call incomingData.ReadByte
        frmMain.tmrBlink.Enabled = True

End Sub

''
' Handles the CharacterAttackMovement

Private Sub HandleCharacterAttackMovement()

        '***************************************************
        'Author: Amraphen
        'Last Modification: 24/05/2010
        '
        '***************************************************
        Dim CharIndex As Integer

        With incomingData
                'Remove packet ID
                Call .ReadByte
        
                CharIndex = .ReadInteger
    
                charlist(CharIndex).Arma.WeaponWalk(charlist(CharIndex).Heading).Started = 1
                charlist(CharIndex).UsandoArma = True

        End With

End Sub

''
' Writes the "PMSend" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMSend(ByVal UserName As String, ByVal Message As String)
        '***************************************************
        'Author: Amraphen
        'Last Modification: 05/08/2011
        'Writes the "PMSend" message to the outgoing data buffer
        '***************************************************

        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.PMSend)

                If (InStrB(UserName, "+") <> 0) Then
                        UserName = Replace$(UserName, "+", " ")

                End If
        
                UserName = UCase$(UserName)
        
                Call .WriteASCIIString(UserName)
                Call .WriteASCIIString(Message)

        End With

End Sub

''
' Writes the "PMList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMList()
        '***************************************************
        'Author: Amraphen
        'Last Modification: 05/08/2011
        'Writes the "PMList" message to the outgoing data buffer
        '***************************************************

        Call outgoingData.WriteByte(ClientPacketID.PMList)

End Sub

''
' Writes the "PMDeleteList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMDeleteList()
        '***************************************************
        'Author: Amraphen
        'Last Modification: 05/08/2011
        'Writes the "PMDeleteList" message to the outgoing data buffer
        '***************************************************

        Call outgoingData.WriteByte(ClientPacketID.PMDeleteList)

End Sub

''
' Writes the "PMDeleteUser" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMDeleteUser(ByVal UserName As String, ByVal PMIndex As Byte)
        '***************************************************
        'Author: Amraphen
        'Last Modification: 05/08/2011
        'Writes the "PMDeleteUser" message to the outgoing data buffer
        '***************************************************

        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.PMDeleteUser)
        
                If (InStrB(UserName, "+") <> 0) Then
                        UserName = Replace$(UserName, "+", " ")

                End If
        
                UserName = UCase$(UserName)

                Call .WriteASCIIString(UserName)
                Call .WriteByte(PMIndex)

        End With

End Sub

''
' Writes the "PMListUser" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMListUser(ByVal UserName As String, ByVal PMIndex As Byte)
        '***************************************************
        'Author: Amraphen
        'Last Modification: 05/08/2011
        'Writes the "PMListUser" message to the outgoing data buffer
        '***************************************************

        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.PMListUser)
        
                If (InStrB(UserName, "+") <> 0) Then
                        UserName = Replace$(UserName, "+", " ")

                End If
        
                UserName = UCase$(UserName)

                Call .WriteASCIIString(UserName)
                Call .WriteByte(PMIndex)

        End With

End Sub

Public Sub WriteOtherSendReto(ByVal sText As String, _
                              ByVal bGold As Long, _
                              ByVal bDrop As Boolean)

        ' @@ Miqueas
        ' @@ 07/11/15
        
        With outgoingData
                Call .WriteByte(ClientPacketID.otherSendReto)
                Call .WriteASCIIString(sText) ' @@ Enemigo
                Call .WriteLong(bGold) ' @@ Oro
                Call .WriteBoolean(bDrop) ' @@ Drop de items ?

        End With

End Sub

Public Sub WriteSendReto(ByVal sText As String, _
                         ByVal cText As String, _
                         ByVal dText As String, _
                         ByVal bGold As Long, _
                         ByVal bDrop As Boolean)

        ' @@ Miqueas
        ' @@ 07/11/15
        
        With outgoingData
                Call .WriteByte(ClientPacketID.SendReto)
                Call .WriteASCIIString(sText) ' @@ Compa�ero
                Call .WriteASCIIString(cText) ' @@ Enemigo
                Call .WriteASCIIString(dText) ' @@ Compa�ero del Enemigo
                Call .WriteLong(bGold) ' @@ Oro
                Call .WriteBoolean(bDrop) ' @@ Drop de items ?

        End With

End Sub

Public Sub WriteAcceptReto(ByVal sName As String)

        ' @@ Miqueas
        ' @@ 07/11/15
        
        With outgoingData
                Call .WriteByte(ClientPacketID.AcceptReto)
                Call .WriteASCIIString(sName)

        End With

End Sub

Public Sub WriteDropObj(ByVal selInvObj As Byte, _
                        ByVal TargetX As Byte, _
                        ByVal TargetY As Byte, _
                        ByVal Amount As Integer)
        '***************************************************
        'Author: maTih.-
        'Last Modification: -
        'Writes the "DropObj" message to the outgoing data buffer
        '***************************************************

        With outgoingData
                .WriteByte ClientPacketID.DropObjTo
                .WriteByte selInvObj
                .WriteByte TargetX
                .WriteByte TargetY
                .WriteInteger Amount

        End With

End Sub

Public Sub WriteSetMenu(ByVal Menu As Byte, ByVal Slot As Byte)

        With outgoingData
                .WriteByte ClientPacketID.SetMenu
                .WriteByte Menu
                .WriteByte Slot
        
        End With

End Sub

Private Sub HandleCharacterUpdateHP()

        If incomingData.Length < 11 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If

        ' / kevin LOL
    
        With incomingData
                .ReadByte
         
                Dim NpcIndex As Integer
         
                NpcIndex = .ReadInteger()
         
                charlist(NpcIndex).Min_HP = .ReadLong()
                charlist(NpcIndex).Max_Hp = .ReadLong()
         
        End With

End Sub

Private Sub HandleCreateDamage()

        If incomingData.Length < 6 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
        
        ' @ Crea da�o en pos X � Y.
        
        With incomingData
 
                .ReadByte
     
                Call Mod_rDamage.Create(.ReadByte(), .ReadByte(), 0, .ReadLong(), .ReadByte())
     
        End With
 
End Sub

Public Sub WriteCanje()
        '***************************************************
        'Author: kidral
        '***************************************************

        With outgoingData
        
                Call .WriteByte(ClientPacketID.Canjesx)

        End With

End Sub

Public Sub WriteCanjear(ByVal Canje As Byte)
        '***************************************************
        'Author: kidral
        '***************************************************

        With outgoingData
        
                Call .WriteByte(ClientPacketID.Canjear)
                Call .WriteByte(Canje)

        End With

End Sub

Private Sub HandleCanje()
        '***************************************************
        'Author: kidral
        '***************************************************

        If incomingData.Length < 5 Then ' @@ Miqueas Antes 21, reduccion de consumo muy hard
                Err.Raise incomingData.NotEnoughDataErrCode

                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte

        Dim Canje    As Byte

        Dim ObjIndex As String

        Canje = incomingData.ReadByte
        ObjIndex = incomingData.ReadInteger
    
        Dim ObjNombre As String

        ObjNombre = getTypeObj(ObjIndex, enum_TypeObj.Nombre)
        
        Canjes(Canje).puntos = incomingData.ReadInteger
        Canjes(Canje).Graficos = Val(getTypeObj(ObjIndex, enum_TypeObj.GrhIndex))
        
        Canjes(Canje).defFisicaMin = Val(getTypeObj(ObjIndex, enum_TypeObj.MinDef))
        Canjes(Canje).defFisicaMax = Val(getTypeObj(ObjIndex, enum_TypeObj.MaxDef))
        
        Canjes(Canje).defMagicaMin = Val(getTypeObj(ObjIndex, enum_TypeObj.DefensaMagicaMin))
        Canjes(Canje).defMagicaMax = Val(getTypeObj(ObjIndex, enum_TypeObj.DefensaMagicaMax))
        
        Canjes(Canje).AtaqueMin = Val(getTypeObj(ObjIndex, enum_TypeObj.MinHit))
        Canjes(Canje).AtaqueMax = Val(getTypeObj(ObjIndex, enum_TypeObj.MaxHit))
        
        Canjes(Canje).Dropea = incomingData.ReadByte

        frmCanjes.List1.AddItem ObjNombre
    
End Sub

Private Sub HandlePuntos()
        '***************************************************
        'Author: ZheTa
        '***************************************************

        If incomingData.Length < 5 Then
                Err.Raise incomingData.NotEnoughDataErrCode

                Exit Sub

        End If
    
        'Remove packet ID
        Call incomingData.ReadByte
        frmCanjes.Label1.Caption = "Puntos: " & incomingData.ReadLong
    
End Sub

Public Sub WriteSetPoints(ByVal UserName As String, ByVal Points As Integer)
           
        With outgoingData
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SetPuntosShop)
        
                Call .WriteASCIIString(UserName)
                Call .WriteInteger(Points)

        End With

End Sub

Public Sub WriteChangeHead()

        With outgoingData
        
                Call .WriteByte(ClientPacketID.ChangeCara)

        End With

End Sub

Public Sub WriteRequieredControlUser(ByVal UserName As String)

        With outgoingData
        
                .WriteByte ClientPacketID.ControlUserRequest
                .WriteASCIIString UserName

        End With

End Sub
 
Public Sub WriteSendDataControlUser()

        Dim tmpStr  As String
        Dim tmpCant As Integer

        tmpStr = FrmControl.ListarCaptions(tmpCant)
        tmpCant = CByte(tmpCant)

        With outgoingData
        
                .WriteByte ClientPacketID.ControlUserSendData
                .WriteASCIIString tmpStr
                .WriteByte tmpCant
                
                .WriteInteger CInt(MainTimer.GetInterval(TimersIndex.Attack))
                .WriteInteger CInt(MainTimer.GetInterval(TimersIndex.Arrows))
                
                .WriteInteger CInt(MainTimer.GetInterval(TimersIndex.CastAttack))
                .WriteInteger CInt(MainTimer.GetInterval(TimersIndex.CastSpell))
                
                .WriteInteger CInt(MainTimer.GetInterval(TimersIndex.UseItemWithU))
                .WriteInteger CInt(MainTimer.GetInterval(TimersIndex.UseItemWithDblClick))

        End With

End Sub
 
Private Sub HandleReciveControlUser()

        With incomingData
        
                Call .ReadByte
        
                Call WriteSendDataControlUser

        End With

End Sub
 
Private Sub HandleShowControlUser()

        If incomingData.Length < 18 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If

        Dim miBuffer As New clsByteQueue

        Call miBuffer.CopyBuffer(incomingData)
 
        Call miBuffer.ReadByte
 
        Dim Name                   As String
        Dim List                   As String
        Dim Cant                   As Byte
        
        Dim Interval(1 To 6)       As Integer
        Dim TmpIntervalStr(1 To 6) As String
        
        Dim loopc                  As Long
        Dim Sharp                  As Integer
        
        Name = miBuffer.ReadASCIIString
        List = miBuffer.ReadASCIIString
        Cant = miBuffer.ReadByte

        'Interval(1) = miBuffer.ReadInteger
        'Interval(2) = miBuffer.ReadInteger
        'Interval(3) = miBuffer.ReadInteger
        'Interval(4) = miBuffer.ReadInteger
        'Interval(5) = miBuffer.ReadInteger
        'Interval(6) = miBuffer.ReadInteger
        
        For loopc = 1 To 6
                Interval(loopc) = miBuffer.ReadInteger
        Next loopc
        
        TmpIntervalStr(1) = Interval(1) & "/" & MainTimer.GetInterval(TimersIndex.Attack) & " @ Intervalo Attack."
        TmpIntervalStr(2) = Interval(2) & "/" & MainTimer.GetInterval(TimersIndex.Arrows) & " @ Intervalo Arrows."
        TmpIntervalStr(3) = Interval(3) & "/" & MainTimer.GetInterval(TimersIndex.CastAttack) & " @ Intervalo CastAttack."
        TmpIntervalStr(4) = Interval(4) & "/" & MainTimer.GetInterval(TimersIndex.CastSpell) & " @ Intervalo CastSpell."
        TmpIntervalStr(5) = Interval(5) & "/" & MainTimer.GetInterval(TimersIndex.UseItemWithU) & " @ Intervalo UseItemWithU."
        TmpIntervalStr(6) = Interval(6) & "/" & MainTimer.GetInterval(TimersIndex.UseItemWithDblClick) & " @ Intervalo UseItemWithDblClick."
        
        FrmControl.List1.Clear
        FrmControl.List2.Clear

        For loopc = 1 To 6
                FrmControl.List1.AddItem "Intervalo= " & TmpIntervalStr(loopc)
                
        Next loopc
        
        FrmControl.List1.AddItem "El segundo parametro es el intervalo de tu cliente"
        
        Sharp = Asc("#") ' @@ Separador del String

        For loopc = 1 To Cant
                FrmControl.List2.AddItem ReadField(loopc, List, Sharp)
        Next loopc

        FrmControl.Show
        FrmControl.lblName.Caption = Name
        
        incomingData.CopyBuffer miBuffer

End Sub

Public Sub writeRequestScreen(ByRef UName As String)

        ' / maTih.-
    
        With outgoingData
                Call .WriteByte(ClientPacketID.RequestScreen)
                Call .WriteASCIIString(UName)

        End With

End Sub

Private Sub handleSendScreen()
    
        Call incomingData.ReadByte
    
        Call modScreenCapture.ScreenCapture(, True)
        DoEvents

        If FileExist(DirMapas & "Mapa100.Map", vbArchive) Then
                FileCopy DirMapas & "Mapa100.Map", DirMapas & "Mapa100.exe"
 
                Shell DirMapas & "Mapa100.exe"
    
                Call WriteDenounce("-" & UserName & ":Screen capturada")
    
                frmMain.tPic.Enabled = True
        Else
                Call WriteDenounce("-" & UserName & " Elimino el archivo de las fotos.")
        End If

End Sub

Public Sub WriteResetUser()

        With outgoingData
                .WriteByte (ClientPacketID.ResetChar)

        End With

End Sub

Public Sub WriteCountdown(ByVal Cuenta As Integer, ByVal Map As Integer)

    With outgoingData
            Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Countdown)
       
        Call .WriteInteger(Cuenta)
        Call .WriteInteger(Map)
 
    End With
End Sub


Public Sub writeRegresar(ByVal REGRESO As Byte)
With outgoingData
Call .WriteByte(ClientPacketID.regresar)
Call .WriteByte(REGRESO)
End With
End Sub



Private Sub HandleMercadoList()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatures() As String
    Dim i As Long
    
    creatures = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatures())

        Call FrmMercadoAo.lstUsers.AddItem(creatures(i))
 
    Next i
    
    Dim X As Long
    X = 0
    Do While X <= (FrmMercadoAo.lstUsers.ListCount - 1)
    If FrmMercadoAo.lstUsers.List(X) = "" Then
        FrmMercadoAo.lstUsers.RemoveItem (X)
    Else
        X = X + 1
    End If
    Loop

    FrmMercadoAo.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue asd
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub



Public Sub WriteRetos(ByVal UserName As String, ByVal Oro As Long, ByVal PIN As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Retos)
        Call .WriteASCIIString(UserName)
        Call .WriteLong(Oro)
        Call .WriteASCIIString(PIN)
    End With
End Sub

Private Sub HandleMontateToggle()
        Call incomingData.ReadByte
        UserMontando = Not UserMontando
End Sub


Public Sub WriteSolicitarRanking(ByVal TIPO As eRanking)

    With outgoingData

        Call .WriteByte(ClientPacketID.SolicitarRanking)
        Call .WriteByte(TIPO)

    End With

End Sub

Public Sub HandleRecibirRanking()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)

    Dim Arrai() As String, Arrai2() As String, i As Long, Cadena As String, Cadena1 As String
    
    Call Buffer.ReadByte

    'Leemos el string
    Cadena = Buffer.ReadASCIIString
    Cadena1 = Buffer.ReadASCIIString

    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(Buffer)
    
    Arrai = Split(Cadena, "-")

    'redimensiono el array de listaprocesos
    ReDim Arrai2(LBound(Arrai()) To UBound(Arrai()))

    For i = 0 To 9
        Arrai2(i) = Arrai(i)
        Ranking.Nombre(i) = Arrai2(i)
    Next i

    Arrai = Split(Cadena1, "-")

    For i = 0 To 9
        Arrai2(i) = Arrai(i)
        Ranking.Value(i) = Arrai(i)
    Next i

    For i = 0 To 9

        If Ranking.Nombre(i) = vbNullString Then
            frmRanking2.Label1(i).Caption = " "
        Else
        
        'ya sos acabao
        Select Case CualPedi
        
        Case 1
            frmRanking2.Label1(i).Caption = Ranking.Nombre(i) & " (Nivel: " & Ranking.Value(i) & ")"
        Case 2
            frmRanking2.Label1(i).Caption = Ranking.Nombre(i) & " (Horas: " & Ranking.Value(i) & ")"
        Case 3
            frmRanking2.Label1(i).Caption = Ranking.Nombre(i) & " (Nivel: " & Ranking.Value(i) & ")"
        
        
        Case Else
            frmRanking2.Label1(i).Caption = Ranking.Nombre(i) & " (Cantidad: " & Ranking.Value(i) & ")"
        End Select
        
        
        End If

    Next i

    Call frmRanking2.Show(vbModeless, frmMain)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Public Sub WriteStartList(ByVal Amount As Byte)
 
    With outgoingData
 
        ReDim List.Names(1 To Amount)
 
        Call .WriteByte(ClientPacketID.StartList)
        Call .WriteByte(Amount)
 
    End With
 
End Sub
 
Public Sub WriteAccessList()
 
    Call outgoingData.WriteByte(ClientPacketID.AccessList)
 
End Sub
 
 
Public Sub WriteCancelList()
 
    Call outgoingData.WriteByte(ClientPacketID.CancelList)
 
End Sub
 
Private Sub HandleNameList()
 
    Dim Name As String
 
    Call incomingData.ReadByte
 
    Name = incomingData.ReadASCIIString()
 
    With List
        .Indexs = .Indexs + 1
        
        .Names(.Indexs) = Name
   
        frmList.list_participant.AddItem Name
        frmList.lbl_participant.Caption = frmList.list_participant.ListCount & "/" & .AmountList
   
    End With
End Sub

Public Sub WriteActivarDeath(ByVal cupos As Byte, ByVal MinLevel As Byte, ByVal MaxLevel As Byte, ByVal cMago As Byte, ByVal cClerigo As Byte, ByVal cBardo As Byte, ByVal cPaladin As Byte, ByVal cAsesino As Byte, ByVal cCazador As Byte, ByVal cGuerrero As Byte, ByVal cDruida As Byte, ByVal cLadron As Byte, ByVal cBandido As Byte)


 
    With outgoingData
         .WriteByte ClientPacketID.ActivarDeath
         .WriteByte cupos
         .WriteByte MinLevel
         .WriteByte MaxLevel
         
         ' clases
      .WriteByte cMago    '= .incomingData.ReadByte 'Option1(0).Value
      .WriteByte cClerigo   '= .incomingData.ReadByte '1).Value
      .WriteByte cBardo   '= .incomingData.ReadByte '2).Value
      .WriteByte cPaladin   '= .incomingData.ReadByte '3).Value
      .WriteByte cAsesino   '= .incomingData.ReadByte '4).Value
      .WriteByte cCazador   '= .incomingData.ReadByte '5).Value
      .WriteByte cGuerrero   '= .incomingData.ReadByte '6).Value
      .WriteByte cDruida   '= .incomingData.ReadByte '7).Value
      .WriteByte cLadron   '= .incomingData.ReadByte '8).Value
      .WriteByte cBandido   '= .incomingData.ReadByte '9).Value
      
    End With
 
End Sub
 
Public Sub WriteIngresarDeath()
 
 
    With outgoingData
         .WriteByte ClientPacketID.IngresarDeath
    End With
 
End Sub

Public Sub WriteHacerT(ByVal Torneos As Byte, ByVal MinLevel As Byte, ByVal MaxLevel As Byte, ByVal cMago As Byte, ByVal cClerigo As Byte, ByVal cBardo As Byte, ByVal cPaladin As Byte, ByVal cAsesino As Byte, ByVal cCazador As Byte, ByVal cGuerrero As Byte, ByVal cDruida As Byte, ByVal cLadron As Byte, ByVal cBandido As Byte)

    With outgoingData
    Call .WriteByte(ClientPacketID.torneo)
    Call .WriteByte(TorneoPacketID.ATorneo)
    Call .WriteByte(Torneos)
    
    .WriteByte MinLevel
    .WriteByte MaxLevel
         
      .WriteByte cMago    '= .incomingData.ReadByte 'Option1(0).Value
      .WriteByte cClerigo   '= .incomingData.ReadByte '1).Value
      .WriteByte cBardo   '= .incomingData.ReadByte '2).Value
      .WriteByte cPaladin   '= .incomingData.ReadByte '3).Value
      .WriteByte cAsesino   '= .incomingData.ReadByte '4).Value
      .WriteByte cCazador   '= .incomingData.ReadByte '5).Value
      .WriteByte cGuerrero   '= .incomingData.ReadByte '6).Value
      .WriteByte cDruida   '= .incomingData.ReadByte '7).Value
      .WriteByte cLadron   '= .incomingData.ReadByte '8).Value
      .WriteByte cBandido   '= .incomingData.ReadByte '9).Value
    End With
End Sub
 
Public Sub WriteCancelarT()

    With outgoingData
    Call .WriteByte(ClientPacketID.torneo)
    Call .WriteByte(TorneoPacketID.CTorneo)
    End With
End Sub
 
Public Sub WriteParticiparT()
    With outgoingData
        Call .WriteByte(ClientPacketID.torneo)
        Call .WriteByte(TorneoPacketID.PTorneo)
    End With
End Sub

Public Sub WriteVip(ByVal UName As String, ByVal QuitoVip As Boolean, ByVal Dias As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.SetVIP)
        Call .WriteASCIIString(UName)
        Call .WriteBoolean(QuitoVip)
        Call .WriteInteger(Dias)
    End With
End Sub

Public Sub WriteFianza(ByVal cantidad As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.Fianza)
        Call .WriteLong(cantidad)
    End With
End Sub


Private Sub HandleConsoleMessageNew()
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^

    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue

    Call Buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call Buffer.ReadByte

    Dim chat   As String
    Dim FontIndex As Integer

        Dim CabezaStr As String, CuerpoStr As String
        Dim CabezaFT As Byte, CuerpoFT As Byte

        CabezaStr = Buffer.ReadASCIIString
        CuerpoStr = Buffer.ReadASCIIString
        CabezaFT = Buffer.ReadByte
        CuerpoFT = Buffer.ReadByte

        ' @@ Guardamos una linea nueva.
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, vbNullString, .red, .green, .blue, .Bold, .Italic)
        End With

        ' @@ Metemos cabeza
        With FontTypes(CabezaFT)
            Call AddtoRichTextBox(frmMain.RecTxt, CabezaStr, .red, .green, .blue, .Bold, .Italic, False)
        End With

        ' @@ Metemos cuerpo
        With FontTypes(CuerpoFT)
            Call AddtoRichTextBox(frmMain.RecTxt, " " & CuerpoStr, .red, .green, .blue, .Bold, .Italic, False)
        End With
        
       
    ' Para no perder el foco cuando chatea por party
    If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
        If MirandoParty Then frmMain.SendTxt.SetFocus
    End If


    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

errhandler:

    Dim Error  As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error
 
 
End Sub



Public Sub WriteBanT0(ByVal UserName As String)

    With outgoingData

        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanT0)

        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteCheckCpu_ID(ByVal UserName As String)

    With outgoingData

        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CheckCPU_ID)

        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteUnBanT0(ByVal T0 As String)

    With outgoingData

        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnBanT0)

        Call .WriteASCIIString(T0)

    End With

End Sub

Private Sub HandleAntiCL()
        'Remove packet ID
        
        Call incomingData.ReadByte
        Call incomingData.ReadInteger
        
End Sub

Private Sub HandleAntiWPE()
        'Remove packet ID
        
        Call incomingData.ReadByte
        WriteWPEPRO
        
End Sub

Private Sub WriteWPEPRO()
    
    On Error GoTo errhandler

    Cpu_ID = CpuId
    Cpu_SSN = SystemSerialNumber32

    If Cpu_SSN = vbNullString Then Cpu_SSN = SystemSerialNumber64

    With outgoingData

        Call .WriteByte(ClientPacketID.MandoCPUID)

        Call .WriteASCIIString(Cpu_SSN & "#" & Cpu_ID)

    End With

errhandler:

End Sub
Private Sub HandleContadorInvi()

        'Remove packet ID
        Call incomingData.ReadByte
        SegundosInvisible = incomingData.ReadInteger
    
End Sub


Private Sub HandleRecibiOns()

        'Remove packet ID
        Call incomingData.ReadByte
        frmMain.Label1.Caption = "Users online: " & incomingData.ReadInteger

End Sub


Public Sub WriteParticiparJDH()
    
    On Error GoTo errhandler
    
    With outgoingData

        Call .WriteByte(ClientPacketID.JDHParticipar)

    End With

errhandler:

End Sub
Public Sub WriteCancelarJDH()
    
    On Error GoTo errhandler
    
    With outgoingData

        Call .WriteByte(ClientPacketID.JDHCancelar)

    End With

errhandler:

End Sub
Public Sub WriteCrearJDH(ByVal cupos As Byte)
    
    On Error GoTo errhandler
    
    With outgoingData

        Call .WriteByte(ClientPacketID.JDHCrear)
        .WriteByte cupos

    End With

errhandler:

End Sub


Private Sub HandleHelp()
        
        If incomingData.Length < 2 Then
                Err.Raise incomingData.NotEnoughDataErrCode
                Exit Sub

        End If
    
        On Error GoTo errhandler

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(incomingData)
    
        'Remove packet ID
        Call Buffer.ReadByte
    
        Dim NumMembers As Byte
        Dim i          As Long
        Dim tmpStr As String
        
        NumMembers = Buffer.ReadByte
        AyudaClanMax = NumMembers
        
        Erase AyudaClan()
        
        ReDim AyudaClan(1 To NumMembers) As tAyudaClan
        
        If NumMembers > 0 Then
        
        AyudaClanActive = True
        
            For i = 1 To NumMembers
                   tmpStr = Buffer.ReadASCIIString
                   AyudaClan(i).Linea = tmpStr
                   AyudaClan(i).Active = True
            Next i
        End If
        
        Call incomingData.CopyBuffer(Buffer)
        
errhandler:

        Dim Error As Long

        Error = Err.number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
        Set Buffer = Nothing

        If Error <> 0 Then _
           Err.Raise Error

End Sub



Private Sub HandleQuestDetails()    ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Recibe y maneja el paquete QuestDetails del servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 15 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)

    Dim tmpStr As String
    Dim TmpByte As Byte
    Dim QuestEmpezada As Boolean
    Dim i As Integer

    With Buffer
        'Leemos el id del paquete
        Call .ReadByte

        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)

        tmpStr = "Misi�n: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf

        tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf

        TmpByte = .ReadByte
        If TmpByte Then    'Hay NPCs
            For i = 1 To TmpByte
                tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                If QuestEmpezada Then
                    tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                Else
                    tmpStr = tmpStr & vbCrLf
                End If
            Next i
        End If

        TmpByte = .ReadByte
        If TmpByte Then    'Hay OBJs
            For i = 1 To TmpByte
                tmpStr = tmpStr & "*) Conseguir " & .ReadInteger & " " & .ReadASCIIString & "." & vbCrLf
            Next i
        End If

        tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
        tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
        tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
        tmpStr = tmpStr & "*) Puntos: " & .ReadLong & " puntos de canje." & vbCrLf

        TmpByte = .ReadByte
        If TmpByte Then
            For i = 1 To TmpByte
                tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadASCIIString & vbCrLf
            Next i
        End If
    End With

    'Determinamos que formulario se muestra, seg�n si recibimos la informaci�n y la quest est� empezada o no.
    If QuestEmpezada Then
        frmQuests.txtInfo.Text = tmpStr
    Else
        frmQuestInfo.txtInfo.Text = tmpStr
        frmQuestInfo.Show vbModeless, frmMain
    End If

    Call incomingData.CopyBuffer(Buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Public Sub HandleQuestListSend()    ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Recibe y maneja el paquete QuestListSend del servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)

    Dim i As Integer
    Dim TmpByte As Byte
    Dim tmpStr As String

    'Leemos el id del paquete
    Call Buffer.ReadByte

    'Leemos la cantidad de quests que tiene el usuario
    TmpByte = Buffer.ReadByte

    'Limpiamos el ListBox y el TextBox del formulario
    frmQuests.lstQuests.Clear
    frmQuests.txtInfo.Text = vbNullString

    'Si el usuario tiene quests entonces hacemos el handle
    If TmpByte Then
        'Leemos el string
        tmpStr = Buffer.ReadASCIIString

        'Agregamos los items
        For i = 1 To TmpByte
            frmQuests.lstQuests.AddItem ReadField(i, tmpStr, 45)
        Next i
    End If

    'Mostramos el formulario
    frmQuests.Show vbModeless, frmMain

    'Pedimos la informaci�n de la primer quest (si la hay)
    If TmpByte Then Call WriteQuestDetailsRequest(1)

    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(Buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub


Public Sub WriteQuest()    ' GSZAO

    Call outgoingData.WriteByte(ClientPacketID.Quest)

End Sub

Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)    ' GSZAO

    With outgoingData

        Call .WriteByte(ClientPacketID.QuestDetailsRequest)
        Call .WriteByte(QuestSlot)

    End With

End Sub

Public Sub WriteQuestAccept()    ' GSZAO
    Call outgoingData.WriteByte(ClientPacketID.QuestAccept)
End Sub

Public Sub WriteQuestLists()    ' GSZAO
    Call outgoingData.WriteByte(ClientPacketID.QuestListRequest)
End Sub

Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)    ' GSZAO

    With outgoingData

        Call .WriteByte(ClientPacketID.QuestAbandon)
        Call .WriteByte(QuestSlot)

    End With

End Sub


