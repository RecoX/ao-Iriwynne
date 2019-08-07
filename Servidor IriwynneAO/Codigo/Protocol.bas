Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Mart�n Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Mart�n Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Mart�n Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As clsByteQueue

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
    UserCharIndexInServer = 29    ' IP
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
    updateuserstats = 47       ' EST
    WorkRequestTarget = 48     ' T01
    ChangeInventorySlot = 49   ' CSI
    ChangeBankSlot = 50        ' SBO
    ChangeSpellSlot = 51       ' SHS
    Atributes = 52            ' ATR
    BlacksmithWeapons = 53     ' LAH
    BlacksmithArmors = 54      ' LAR
    CarpenterObjects = 55      ' OBR
    RestOK = 56                ' DOK
    ErrorMSG = 57              ' ERR
    Blind = 58                 ' CEGU
    Dumb = 59                  ' DUMB
    ShowSignal = 60            ' MCAR
    ChangeNPCInventorySlot = 61    ' NPCI
    UpdateHungerAndThirst = 62    ' EHYS
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
    ShowGuildFundationForm = 83    ' SHOWFUN
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
    CharacterUpdateHp = 110
    CreateDamage = 111
    Canje = 112                 'Canjes
    CanjePTS = 113              'Canjes
    ControlUserRecive = 114
    ControlUserShow = 115
    RequestScreen = 116
    regresar = 117
    MercadoList = 118
    Retos = 119
    MontateToggle = 120
    EnviarDatosRanking = 121
    NameList = 122
    ConsoleMsgNew = 123
    AntiClienteEditado = 124
    CuentaInvi = 125
    MandoOnlines = 126
    PedirCPUID = 127
    ayudaclan = 128
    QuestDetails = 129          ' GSZAO
    QuestListSend = 130         ' GSZAO
End Enum

Private Enum ClientPacketID

    LoginExistingChar = 1     'OLOGIN
    ResetChar = 2
    LoginNewChar = 3          'NLOGIN
    TALK = 4                  ';
    Yell = 5                  '-
    Whisper = 6               '\
    Walk = 7                  'M
    RequestPositionUpdate = 8    'RPU
    Attack = 9                'AT
    PickUp = 10                'AG
    SafeToggle = 11            '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle = 12
    RequestGuildLeaderInfo = 13    'GLINFO
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
    GuildRequestJoinerInfo = 58    'ENVCOMEN
    GuildAlliancePropList = 59    'ENVALPRO
    GuildPeacePropList = 60    'ENVPROPP
    GuildDeclareWar = 61       'DECGUERR
    GuildNewWebsite = 62       'NEWWEBSI
    GuildAcceptNewMember = 63  'ACEPTARI
    GuildRejectNewMember = 64  'RECHAZAR
    GuildKickMember = 65       'ECHARCLA
    GuildUpdateNews = 66       'ACTGNEWS
    GuildMemberInfo = 67       '1HRINFO<
    GuildOpenElections = 68    'ABREELEC
    GuildRequestMembership = 69    'SOLICITUD
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
    UpTime = 91                '/UPTIME
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
    moveItem = 130              'Drag and drop
    PMDeleteList = 131
    PMList = 132
    otherSendReto = 133    ' @@ 1 vs 1
    SendReto = 134      ' @@ 2 vs 2
    AcceptReto = 135  ' @@ Aceptar 1.1 | 2.2
    DropObjTo = 136
    SetMenu = 137
    Canjear = 138               'Mab
    Canjesx = 139                'Amishar
    ChangeCara = 140
    ControlUserRequest = 141
    ControlUserSendData = 142
    RequestScreen = 143
    regresar = 144
    MercadoList = 145
    Retos = 146
    GranPoder = 147
    UsuarioPideRanking = 148
    pideCastillo = 149
    StartList = 150
    AccessList = 151
    CancelList = 152
    ActivarDeath = 153
    IngresarDeath = 154         '/DEATH
    Torneo = 155
    SetVip = 156
    Fianza = 157
    UserMandaCPU = 158
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
    ATorneo                 '/ARRANCARTORNEO
    CTorneo                 '/CANCELART
    PTorneo                 '/PARTICIPAR
End Enum

Public Enum eGMCommands

    GMMessage = 1           '/GMSG
    showName = 2              '/SHOWNAME
    OnlineRoyalArmy = 3       '/ONLINEREAL
    OnlineChaosLegion = 4     '/ONLINECAOS
    GoNearby = 5              '/IRCERCA
    comment = 6               '/REM
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
    DestroyAllItemsInArea = 61    '/MASSDEST
    AcceptRoyalCouncilMember = 62    '/ACEPTCONSE
    AcceptChaosCouncilMember = 63    '/ACEPTCONSECAOS
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
    ServerOpenToUsersToggle = 96    '/HABILITAR
    TurnOffServer = 97         '/APAGAR
    TurnCriminal = 98          '/CONDEN
    ResetFactions = 99         '/RAJAR
    RemoveCharFromGuild = 100   '/RAJARCLAN
    RequestCharMail = 101       '/LASTEMAIL
    AlterPassword = 102         '/APASS
    AlterMail = 103             '/AEMAIL
    AlterName = 104             '/ANAME
    ToggleCentinelActivated = 105    '/CENTINELAACTIVADO
    DoBackUp = 106              '/DOBACKUP
    ShowGuildMessages = 107     '/SHOWCMSG
    SaveMap = 108               '/GUARDAMAPA
    ChangeMapInfoPK = 109       '/MODMAPINFO PK
    ChangeMapInfoBackup = 110   '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted = 111    '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic = 112  '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi = 113   '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu = 114   '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand = 115     '/MODMAPINFO TERRENO
    ChangeMapInfoZone = 116     '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc = 117    '/MODMAPINFO ROBONPCm
    ChangeMapInfoNoOcultar = 118    '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar = 119    '/MODMAPINFO INVOCARSINEFECTO
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


    VerCPU = 156
    BanT0 = 157
    UnBanT0 = 158

End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 143

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

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Vida
    eo_Poss

End Enum

Public Sub InitAuxiliarBuffer()
'***************************************************
'Author: ZaMa
'Last Modification: 15/03/2011
'Initializaes Auxiliar Buffer
'***************************************************
    Set auxiliarBuffer = New clsByteQueue

End Sub

''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIncomingData(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
    On Error Resume Next

    Dim packetID As Byte

    packetID = UserList(UserIndex).incomingData.PeekByte()

    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.LoginExistingChar _
            Or packetID = ClientPacketID.LoginNewChar) Then

        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub

            'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(UserIndex).Counters.IdleCount = 0

        End If

    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
        UserList(UserIndex).Counters.IdleCount = 0

        'Is the user logged?
        If UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

    End If

    ' Ante cualquier paquete, pierde la proteccion de ser atacado.
    UserList(UserIndex).flags.NoPuedeSerAtacado = False

    Select Case packetID

    
        Case ClientPacketID.Quest
            Call HandleQuest(UserIndex)

        Case ClientPacketID.QuestAccept
            Call HandleQuestAccept(UserIndex)

        Case ClientPacketID.QuestListRequest
            Call HandleQuestListRequest(UserIndex)

        Case ClientPacketID.QuestDetailsRequest
            Call HandleQuestDetailsRequest(UserIndex)

        Case ClientPacketID.QuestAbandon
            Call HandleQuestAbandon(UserIndex)
            
    Case ClientPacketID.JDHCancelar
        Call HandleJDHCancelar(UserIndex)
    
    Case ClientPacketID.JDHCrear
        Call HandleJDHCrear(UserIndex)
    
    Case ClientPacketID.JDHParticipar
        Call HandleJDHEntrar(UserIndex)
        
    Case ClientPacketID.SetVip
        Call HandleSetVip(UserIndex)

    Case ClientPacketID.Fianza
        Call HandleFianza(UserIndex)
        
    Case ClientPacketID.UserMandaCPU
        Call HandleRecibiCpuID(UserIndex)
        
    Case ClientPacketID.ActivarDeath            '/ACDEATH CUPOS
        Call HandleActivarDeath(UserIndex)

    Case ClientPacketID.IngresarDeath           '/DEATH
        Call HandleIngresarDeath(UserIndex)

    Case ClientPacketID.AccessList
        Call HandleAccessList(UserIndex)

    Case ClientPacketID.CancelList
        Call HandleCancelList(UserIndex)

    Case ClientPacketID.StartList
        Call HandleStartList(UserIndex)

    Case ClientPacketID.pideCastillo
        Call HandleCastillo(UserIndex)

    Case ClientPacketID.GranPoder        'OLOGIN
        Call HandlePoder(UserIndex)

    Case ClientPacketID.UsuarioPideRanking
        Call HandlePideRanking(UserIndex)

    Case ClientPacketID.LoginExistingChar       'OLOGIN
        Call HandleLoginExistingChar(UserIndex)

    Case ClientPacketID.ResetChar
        Call HandleResetChar(UserIndex)

    Case ClientPacketID.LoginNewChar            'NLOGIN
        Call HandleLoginNewChar(UserIndex)

    Case ClientPacketID.TALK                    ';
        Call HandleTalk(UserIndex)

    Case ClientPacketID.Yell                    '-
        Call HandleYell(UserIndex)

    Case ClientPacketID.Whisper                 '\
        Call HandleWhisper(UserIndex)

    Case ClientPacketID.Walk                    'M
        Call HandleWalk(UserIndex)

    Case ClientPacketID.RequestPositionUpdate   'RPU
        Call HandleRequestPositionUpdate(UserIndex)

    Case ClientPacketID.Attack                  'AT
        Call HandleAttack(UserIndex)

    Case ClientPacketID.PickUp                  'AG
        Call HandlePickUp(UserIndex)

    Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
        Call HandleSafeToggle(UserIndex)

    Case ClientPacketID.ResuscitationSafeToggle
        Call HandleResuscitationToggle(UserIndex)

    Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
        Call HandleRequestGuildLeaderInfo(UserIndex)

    Case ClientPacketID.RequestAtributes        'ATR
        Call HandleRequestAtributes(UserIndex)

    Case ClientPacketID.RequestFame             'FAMA
        Call HandleRequestFame(UserIndex)

    Case ClientPacketID.RequestSkills           'ESKI
        Call HandleRequestSkills(UserIndex)

    Case ClientPacketID.RequestMiniStats        'FEST
        Call HandleRequestMiniStats(UserIndex)

    Case ClientPacketID.CommerceEnd             'FINCOM
        Call HandleCommerceEnd(UserIndex)

    Case ClientPacketID.CommerceChat
        Call HandleCommerceChat(UserIndex)

    Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
        Call HandleUserCommerceEnd(UserIndex)

    Case ClientPacketID.UserCommerceConfirm
        Call HandleUserCommerceConfirm(UserIndex)

    Case ClientPacketID.BankEnd                 'FINBAN
        Call HandleBankEnd(UserIndex)

    Case ClientPacketID.UserCommerceOk          'COMUSUOK
        Call HandleUserCommerceOk(UserIndex)

    Case ClientPacketID.UserCommerceReject      'COMUSUNO
        Call HandleUserCommerceReject(UserIndex)

    Case ClientPacketID.Drop                    'TI
        Call HandleDrop(UserIndex)

    Case ClientPacketID.CastSpell               'LH
        Call HandleCastSpell(UserIndex)

    Case ClientPacketID.LeftClick               'LC
        Call HandleLeftClick(UserIndex)

    Case ClientPacketID.DoubleClick             'RC
        Call HandleDoubleClick(UserIndex)

    Case ClientPacketID.Work                    'UK
        Call HandleWork(UserIndex)

    Case ClientPacketID.UseSpellMacro           'UMH
        Call HandleUseSpellMacro(UserIndex)

    Case ClientPacketID.UseItem                 'USA
        Call HandleUseItem(UserIndex)

    Case ClientPacketID.CraftBlacksmith         'CNS
        Call HandleCraftBlacksmith(UserIndex)

    Case ClientPacketID.CraftCarpenter          'CNC
        Call HandleCraftCarpenter(UserIndex)

    Case ClientPacketID.WorkLeftClick           'WLC
        Call HandleWorkLeftClick(UserIndex)

    Case ClientPacketID.CreateNewGuild          'CIG
        Call HandleCreateNewGuild(UserIndex)

    Case ClientPacketID.SpellInfo               'INFS
        Call HandleSpellInfo(UserIndex)

    Case ClientPacketID.EquipItem               'EQUI
        Call HandleEquipItem(UserIndex)

    Case ClientPacketID.ChangeHeading           'CHEA
        Call HandleChangeHeading(UserIndex)

    Case ClientPacketID.ModifySkills            'SKSE
        Call HandleModifySkills(UserIndex)

    Case ClientPacketID.Train                   'ENTR
        Call HandleTrain(UserIndex)

    Case ClientPacketID.CommerceBuy             'COMP
        Call HandleCommerceBuy(UserIndex)

    Case ClientPacketID.BankExtractItem         'RETI
        Call HandleBankExtractItem(UserIndex)

    Case ClientPacketID.CommerceSell            'VEND
        Call HandleCommerceSell(UserIndex)

    Case ClientPacketID.BankDeposit             'DEPO
        Call HandleBankDeposit(UserIndex)

    Case ClientPacketID.ForumPost               'DEMSG
        Call HandleForumPost(UserIndex)

    Case ClientPacketID.MoveSpell               'DESPHE
        Call HandleMoveSpell(UserIndex)

    Case ClientPacketID.MoveBank
        Call HandleMoveBank(UserIndex)

    Case ClientPacketID.ClanCodexUpdate         'DESCOD
        Call HandleClanCodexUpdate(UserIndex)

    Case ClientPacketID.UserCommerceOffer       'OFRECER
        Call HandleUserCommerceOffer(UserIndex)

    Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
        Call HandleGuildAcceptPeace(UserIndex)

    Case ClientPacketID.GuildRejectAlliance     'RECPALIA
        Call HandleGuildRejectAlliance(UserIndex)

    Case ClientPacketID.GuildRejectPeace        'RECPPEAT
        Call HandleGuildRejectPeace(UserIndex)

    Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
        Call HandleGuildAcceptAlliance(UserIndex)

    Case ClientPacketID.GuildOfferPeace         'PEACEOFF
        Call HandleGuildOfferPeace(UserIndex)

    Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
        Call HandleGuildOfferAlliance(UserIndex)

    Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
        Call HandleGuildAllianceDetails(UserIndex)

    Case ClientPacketID.GuildPeaceDetails       'PEACEDET
        Call HandleGuildPeaceDetails(UserIndex)

    Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
        Call HandleGuildRequestJoinerInfo(UserIndex)

    Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
        Call HandleGuildAlliancePropList(UserIndex)

    Case ClientPacketID.GuildPeacePropList      'ENVPROPP
        Call HandleGuildPeacePropList(UserIndex)

    Case ClientPacketID.GuildDeclareWar         'DECGUERR
        Call HandleGuildDeclareWar(UserIndex)

    Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
        Call HandleGuildNewWebsite(UserIndex)

    Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
        Call HandleGuildAcceptNewMember(UserIndex)

    Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
        Call HandleGuildRejectNewMember(UserIndex)

    Case ClientPacketID.GuildKickMember         'ECHARCLA
        Call HandleGuildKickMember(UserIndex)

    Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
        Call HandleGuildUpdateNews(UserIndex)

    Case ClientPacketID.GuildMemberInfo         '1HRINFO<
        Call HandleGuildMemberInfo(UserIndex)

    Case ClientPacketID.GuildOpenElections      'ABREELEC
        Call HandleGuildOpenElections(UserIndex)

    Case ClientPacketID.GuildRequestMembership  'SOLICITUD
        Call HandleGuildRequestMembership(UserIndex)

    Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
        Call HandleGuildRequestDetails(UserIndex)

    Case ClientPacketID.Online                  '/ONLINE
        Call HandleOnline(UserIndex)

    Case ClientPacketID.Quit                    '/SALIR
        Call HandleQuit(UserIndex)

    Case ClientPacketID.GuildLeave              '/SALIRCLAN
        Call HandleGuildLeave(UserIndex)

    Case ClientPacketID.RequestAccountState     '/BALANCE
        Call HandleRequestAccountState(UserIndex)

    Case ClientPacketID.PetStand                '/QUIETO
        Call HandlePetStand(UserIndex)

    Case ClientPacketID.PetFollow               '/ACOMPA�AR
        Call HandlePetFollow(UserIndex)

    Case ClientPacketID.ReleasePet              '/LIBERAR
        Call HandleReleasePet(UserIndex)

    Case ClientPacketID.TrainList               '/ENTRENAR
        Call HandleTrainList(UserIndex)

    Case ClientPacketID.Rest                    '/DESCANSAR
        Call HandleRest(UserIndex)

    Case ClientPacketID.Meditate                '/MEDITAR
        Call HandleMeditate(UserIndex)

    Case ClientPacketID.Resucitate              '/RESUCITAR
        Call HandleResucitate(UserIndex)

    Case ClientPacketID.Heal                    '/CURAR
        Call HandleHeal(UserIndex)

    Case ClientPacketID.Help                    '/AYUDA
        Call HandleHelp(UserIndex)

    Case ClientPacketID.RequestStats            '/EST
        Call HandleRequestStats(UserIndex)

    Case ClientPacketID.CommerceStart           '/COMERCIAR
        Call HandleCommerceStart(UserIndex)

    Case ClientPacketID.BankStart               '/BOVEDA
        Call HandleBankStart(UserIndex)

    Case ClientPacketID.Enlist                  '/ENLISTAR
        Call HandleEnlist(UserIndex)

    Case ClientPacketID.Information             '/INFORMACION
        Call HandleInformation(UserIndex)

    Case ClientPacketID.Reward                  '/RECOMPENSA
        Call HandleReward(UserIndex)

    Case ClientPacketID.RequestMOTD             '/MOTD
        Call HandleRequestMOTD(UserIndex)

    Case ClientPacketID.UpTime                  '/UPTIME
        Call HandleUpTime(UserIndex)

    Case ClientPacketID.PartyLeave              '/SALIRPARTY
        Call HandlePartyLeave(UserIndex)

    Case ClientPacketID.PartyCreate             '/CREARPARTY
        Call HandlePartyCreate(UserIndex)

    Case ClientPacketID.PartyJoin               '/PARTY
        Call HandlePartyJoin(UserIndex)

    Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
        Call HandleInquiry(UserIndex)

    Case ClientPacketID.GuildMessage            '/CMSG
        Call HandleGuildMessage(UserIndex)

    Case ClientPacketID.PartyMessage            '/PMSG
        Call HandlePartyMessage(UserIndex)

    Case ClientPacketID.CentinelReport          '/CENTINELA
        Call HandleCentinelReport(UserIndex)

    Case ClientPacketID.GuildOnline             '/ONLINECLAN
        Call HandleGuildOnline(UserIndex)

    Case ClientPacketID.PartyOnline             '/ONLINEPARTY
        Call HandlePartyOnline(UserIndex)

    Case ClientPacketID.CouncilMessage          '/BMSG
        Call HandleCouncilMessage(UserIndex)

    Case ClientPacketID.RoleMasterRequest       '/ROL
        Call HandleRoleMasterRequest(UserIndex)

    Case ClientPacketID.GMRequest               '/GM
        Call HandleGMRequest(UserIndex)

    Case ClientPacketID.bugReport               '/_BUG
        Call HandleBugReport(UserIndex)

    Case ClientPacketID.ChangeDescription       '/DESC
        Call HandleChangeDescription(UserIndex)

    Case ClientPacketID.GuildVote               '/VOTO
        Call HandleGuildVote(UserIndex)

    Case ClientPacketID.Punishments             '/PENAS
        Call HandlePunishments(UserIndex)

    Case ClientPacketID.ChangePassword          '/CONTRASE�A
        Call HandleChangePassword(UserIndex)

    Case ClientPacketID.Gamble                  '/APOSTAR
        Call HandleGamble(UserIndex)

    Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
        Call HandleInquiryVote(UserIndex)

    Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
        Call HandleLeaveFaction(UserIndex)

    Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
        Call HandleBankExtractGold(UserIndex)

    Case ClientPacketID.BankDepositGold         '/DEPOSITAR
        Call HandleBankDepositGold(UserIndex)

    Case ClientPacketID.Denounce                '/DENUNCIAR
        Call HandleDenounce(UserIndex)

    Case ClientPacketID.GuildFundate            '/FUNDARCLAN
        Call HandleGuildFundate(UserIndex)

    Case ClientPacketID.GuildFundation
        Call HandleGuildFundation(UserIndex)

    Case ClientPacketID.PartyKick               '/ECHARPARTY
        Call HandlePartyKick(UserIndex)

    Case ClientPacketID.PartySetLeader          '/PARTYLIDER
        Call HandlePartySetLeader(UserIndex)

    Case ClientPacketID.PartyAcceptMember       '/ACCEPTPARTY
        Call HandlePartyAcceptMember(UserIndex)

    Case ClientPacketID.Ping                    '/PING
        Call HandlePing(UserIndex)


    Case ClientPacketID.Torneo
        UserList(UserIndex).incomingData.ReadByte

        Select Case UserList(UserIndex).incomingData.PeekByte()

        Case TorneoPacketID.ATorneo
            Call HandleArrancaTorneo(UserIndex)

        Case TorneoPacketID.CTorneo
            Call HandleCancelaTorneo(UserIndex)

        Case TorneoPacketID.PTorneo
            Call HandleParticipar(UserIndex)

        End Select

    Case ClientPacketID.RequestPartyForm
        Call HandlePartyForm(UserIndex)

    Case ClientPacketID.ItemUpgrade
        Call HandleItemUpgrade(UserIndex)

    Case ClientPacketID.GMCommands              'GM Messages
        Call HandleGMCommands(UserIndex)

    Case ClientPacketID.InitCrafting
        Call HandleInitCrafting(UserIndex)

    Case ClientPacketID.Home
        Call HandleHome(UserIndex)

    Case ClientPacketID.ShowGuildNews
        Call HandleShowGuildNews(UserIndex)

    Case ClientPacketID.ShareNpc
        Call HandleShareNpc(UserIndex)

    Case ClientPacketID.StopSharingNpc
        Call HandleStopSharingNpc(UserIndex)

    Case ClientPacketID.Consultation
        Call HandleConsultation(UserIndex)

    Case ClientPacketID.moveItem
        Call HandleMoveItem(UserIndex)

    Case ClientPacketID.PMList
        Call HandlePMList(UserIndex)

    Case ClientPacketID.PMDeleteList
        Call HandlePMDeleteList(UserIndex)

    Case ClientPacketID.DropObjTo               ' Drop to pos.
        Call HandleDropObj(UserIndex)

    Case ClientPacketID.otherSendReto
        Call handleOtherSendReto(UserIndex)

    Case ClientPacketID.SendReto                'RETAR
        Call handleSendReto(UserIndex)

    Case ClientPacketID.AcceptReto              '/RETAR
        Call handleAcceptReto(UserIndex)

    Case ClientPacketID.SetMenu              ' 0 Inventario hechizos - 1 Inventario objetos
        Call HandleSetMenu(UserIndex)

    Case ClientPacketID.Canjear                'Cris
        Call HandleCanjer(UserIndex)

    Case ClientPacketID.Canjesx                 'Zama
        Call HandleDameCanje(UserIndex)

    Case ClientPacketID.ChangeCara           ' / Cara
        Call HandleChangeHead(UserIndex)

    Case ClientPacketID.ControlUserRequest
        Call HandleRequieredControlUser(UserIndex)

    Case ClientPacketID.ControlUserSendData
        Call HandleSendDataControlUser(UserIndex)

    Case ClientPacketID.RequestScreen
        Call handleRequestScreen(UserIndex)

    Case ClientPacketID.regresar
        Call Handleregresar(UserIndex)

    Case ClientPacketID.Retos    '
        Call HandleRetos(UserIndex)

    Case Else
        'ERROR : Abort!
        Call CloseSocket(UserIndex)

    End Select

    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        Call HandleIncomingData(UserIndex)

    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.source & _
                      vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                      vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                      " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID))
        Call CloseSocket(UserIndex)

    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)

    End If

End Sub

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, _
                             ByVal MessageIndex As Integer, _
                             Optional ByVal Arg1 As Long, _
                             Optional ByVal Arg2 As Long, _
                             Optional ByVal Arg3 As Long, _
                             Optional ByVal StringArg1 As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)

        Select Case MessageIndex

        Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, _
             eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, _
             eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, _
             eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome

        Case eMessages.NPCHitUser
            Call .WriteByte(Arg1)    'Target
            Call .WriteInteger(Arg2)    'damage

        Case eMessages.UserHitNPC
            Call .WriteLong(Arg1)    'damage

        Case eMessages.UserAttackedSwing
            Call .WriteInteger(UserList(Arg1).Char.CharIndex)

        Case eMessages.UserHittedByUser
            Call .WriteInteger(Arg1)    'AttackerIndex
            Call .WriteByte(Arg2)    'Target
            Call .WriteInteger(Arg3)    'damage

        Case eMessages.UserHittedUser
            Call .WriteInteger(Arg1)    'AttackerIndex
            Call .WriteByte(Arg2)    'Target
            Call .WriteInteger(Arg3)    'damage

        Case eMessages.WorkRequestTarget
            Call .WriteByte(Arg1)    'skill

        Case eMessages.HaveKilledUser    '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
            Call .WriteInteger(UserList(Arg1).Char.CharIndex)    'VictimIndex
            Call .WriteLong(Arg2)    'Expe

        Case eMessages.UserKill    '"�" & .name & " te ha matado!"
            Call .WriteInteger(UserList(Arg1).Char.CharIndex)    'AttackerIndex

        Case eMessages.EarnExp

        Case eMessages.Home
            Call .WriteByte(CByte(Arg1))
            Call .WriteInteger(CInt(Arg2))
            'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
             hasta que no se pasen los dats e .INFs al cliente, esto queda as�.
            Call .WriteASCIIString(StringArg1)    'Call .WriteByte(CByte(Arg2))

        End Select

    End With

    Exit Sub    ''

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Private Sub HandleGMCommands(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandleR

    Dim Command As Byte

    With UserList(UserIndex)
        Call .incomingData.ReadByte

        Command = .incomingData.PeekByte

        Select Case Command

        Case eGMCommands.GMMessage
            Call HandleGMMessage(UserIndex)

        Case eGMCommands.VerHD
            Call HandleVerHD(UserIndex)

        Case eGMCommands.BanHD
            Call HandleBanHD(UserIndex)

        Case eGMCommands.UnBanHD
            Call HandleUnbanHD(UserIndex)

        Case eGMCommands.UnBanT0
            Call HandleUnbanT0(UserIndex)

        Case eGMCommands.BanT0
            Call HandleBanT0(UserIndex)

        Case eGMCommands.VerCPU
            Call HandleCheckCPU_ID(UserIndex)

        Case eGMCommands.showName                '/SHOWNAME
            Call HandleShowName(UserIndex)

        Case eGMCommands.OnlineRoyalArmy
            Call HandleOnlineRoyalArmy(UserIndex)

        Case eGMCommands.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(UserIndex)

        Case eGMCommands.GoNearby                '/IRCERCA
            Call HandleGoNearby(UserIndex)

        Case eGMCommands.comment                 '/REM
            Call HandleComment(UserIndex)

        Case eGMCommands.serverTime              '/HORA
            Call HandleServerTime(UserIndex)

        Case eGMCommands.Where                   '/DONDE
            Call HandleWhere(UserIndex)

        Case eGMCommands.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(UserIndex)

        Case eGMCommands.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(UserIndex)

        Case eGMCommands.WarpChar                '/TELEP
            Call HandleWarpChar(UserIndex)

        Case eGMCommands.Silence                 '/SILENCIAR
            Call HandleSilence(UserIndex)

        Case eGMCommands.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(UserIndex)

        Case eGMCommands.SOSRemove               'SOSDONE
            Call HandleSOSRemove(UserIndex)

        Case eGMCommands.GoToChar                '/IRA
            Call HandleGoToChar(UserIndex)

        Case eGMCommands.invisible               '/INVISIBLE
            Call HandleInvisible(UserIndex)

        Case eGMCommands.GMPanel                 '/PANELGM
            Call HandleGMPanel(UserIndex)

        Case eGMCommands.RequestUserList         'LISTUSU
            Call HandleRequestUserList(UserIndex)

        Case eGMCommands.Working                 '/TRABAJANDO
            Call HandleWorking(UserIndex)

        Case eGMCommands.Hiding                  '/OCULTANDO
            Call HandleHiding(UserIndex)

        Case eGMCommands.Jail                    '/CARCEL
            Call HandleJail(UserIndex)

        Case eGMCommands.KillNPC                 '/RMATA
            Call HandleKillNPC(UserIndex)

        Case eGMCommands.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(UserIndex)

        Case eGMCommands.EditChar                '/MOD
            Call HandleEditChar(UserIndex)

        Case eGMCommands.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(UserIndex)

        Case eGMCommands.RequestCharStats        '/STAT
            Call HandleRequestCharStats(UserIndex)

        Case eGMCommands.RequestCharGold         '/BAL
            Call HandleRequestCharGold(UserIndex)

        Case eGMCommands.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(UserIndex)

        Case eGMCommands.RequestCharBank         '/BOV
            Call HandleRequestCharBank(UserIndex)

        Case eGMCommands.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(UserIndex)

        Case eGMCommands.ReviveChar              '/REVIVIR
            Call HandleReviveChar(UserIndex)

        Case eGMCommands.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(UserIndex)

        Case eGMCommands.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(UserIndex)

        Case eGMCommands.Forgive                 '/PERDON
            Call HandleForgive(UserIndex)

        Case eGMCommands.Kick                    '/ECHAR
            Call HandleKick(UserIndex)

        Case eGMCommands.Execute                 '/EJECUTAR
            Call HandleExecute(UserIndex)

        Case eGMCommands.banChar                 '/BAN
            Call HandleBanChar(UserIndex)

        Case eGMCommands.UnbanChar               '/UNBAN
            Call HandleUnbanChar(UserIndex)

        Case eGMCommands.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(UserIndex)

        Case eGMCommands.SummonChar              '/SUM
            Call HandleSummonChar(UserIndex)

        Case eGMCommands.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(UserIndex)

        Case eGMCommands.SpawnCreature           'SPA
            Call HandleSpawnCreature(UserIndex)

        Case eGMCommands.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(UserIndex)

        Case eGMCommands.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(UserIndex)

        Case eGMCommands.ServerMessage           '/RMSG
            Call HandleServerMessage(UserIndex)

        Case eGMCommands.MapMessage              '/MAPMSG
            Call HandleMapMessage(UserIndex)

        Case eGMCommands.nickToIP                '/NICK2IP
            Call HandleNickToIP(UserIndex)

        Case eGMCommands.IPToNick                '/IP2NICK
            Call HandleIPToNick(UserIndex)

        Case eGMCommands.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(UserIndex)

        Case eGMCommands.TeleportCreate          '/CT
            Call HandleTeleportCreate(UserIndex)

        Case eGMCommands.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(UserIndex)

        Case eGMCommands.RainToggle              '/LLUVIA
            Call HandleRainToggle(UserIndex)

        Case eGMCommands.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(UserIndex)

        Case eGMCommands.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(UserIndex)

        Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(UserIndex)

        Case eGMCommands.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(UserIndex)

        Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(UserIndex)

        Case eGMCommands.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(UserIndex)

        Case eGMCommands.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(UserIndex)

        Case eGMCommands.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(UserIndex)

        Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(UserIndex)

        Case eGMCommands.AcceptRoyalCouncilMember    '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(UserIndex)

        Case eGMCommands.AcceptChaosCouncilMember    '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(UserIndex)

        Case eGMCommands.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(UserIndex)

        Case eGMCommands.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(UserIndex)

        Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(UserIndex)

        Case eGMCommands.dumpIPTables            '/DUMPSECURITY
            Call HandleDumpIPTables(UserIndex)

        Case eGMCommands.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(UserIndex)

        Case eGMCommands.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(UserIndex)

        Case eGMCommands.AskTrigger              '/TRIGGER with no args
            Call HandleAskTrigger(UserIndex)

        Case eGMCommands.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(UserIndex)

        Case eGMCommands.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(UserIndex)

        Case eGMCommands.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(UserIndex)

        Case eGMCommands.GuildBan                '/BANCLAN
            Call HandleGuildBan(UserIndex)

        Case eGMCommands.BanIP                   '/BANIP
            Call HandleBanIP(UserIndex)

        Case eGMCommands.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(UserIndex)

        Case eGMCommands.CreateItem              '/CI
            Call HandleCreateItem(UserIndex)

        Case eGMCommands.DestroyItems            '/DEST
            Call HandleDestroyItems(UserIndex)

        Case eGMCommands.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(UserIndex)

        Case eGMCommands.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(UserIndex)

        Case eGMCommands.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(UserIndex)

        Case eGMCommands.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(UserIndex)

        Case eGMCommands.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(UserIndex)

        Case eGMCommands.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(UserIndex)

        Case eGMCommands.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(UserIndex)

        Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(UserIndex)

        Case eGMCommands.LastIP                  '/LASTIP
            Call HandleLastIP(UserIndex)

        Case eGMCommands.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(UserIndex)

        Case eGMCommands.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(UserIndex)

        Case eGMCommands.SystemMessage           '/SMSG
            Call HandleSystemMessage(UserIndex)

        Case eGMCommands.CreateNPC               '/ACC
            Call HandleCreateNPC(UserIndex)

        Case eGMCommands.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(UserIndex)

        Case eGMCommands.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(UserIndex)

        Case eGMCommands.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(UserIndex)

        Case eGMCommands.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(UserIndex)

        Case eGMCommands.ServerOpenToUsersToggle    '/HABILITAR
            Call HandleServerOpenToUsersToggle(UserIndex)

        Case eGMCommands.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(UserIndex)

        Case eGMCommands.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(UserIndex)

        Case eGMCommands.ResetFactions           '/RAJAR
            Call HandleResetFactions(UserIndex)

        Case eGMCommands.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(UserIndex)

        Case eGMCommands.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(UserIndex)

        Case eGMCommands.AlterPassword           '/APASS
            Call HandleAlterPassword(UserIndex)

        Case eGMCommands.AlterMail               '/AEMAIL
            Call HandleAlterMail(UserIndex)

        Case eGMCommands.AlterName               '/ANAME
            Call HandleAlterName(UserIndex)

        Case eGMCommands.ToggleCentinelActivated    '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(UserIndex)

        Case eGMCommands.DoBackUp               '/DOBACKUP
            Call HandleDoBackUp(UserIndex)

        Case eGMCommands.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(UserIndex)

        Case eGMCommands.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(UserIndex)

        Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(UserIndex)

        Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(UserIndex)

        Case eGMCommands.ChangeMapInfoRestricted    '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(UserIndex)

        Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(UserIndex)

        Case eGMCommands.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(UserIndex)

        Case eGMCommands.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(UserIndex)

        Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(UserIndex)

        Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(UserIndex)

        Case eGMCommands.ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
            Call HandleChangeMapInfoStealNpc(UserIndex)

        Case eGMCommands.ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
            Call HandleChangeMapInfoNoOcultar(UserIndex)

        Case eGMCommands.ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
            Call HandleChangeMapInfoNoInvocar(UserIndex)

        Case eGMCommands.SaveChars               '/GRABAR
            Call HandleSaveChars(UserIndex)

        Case eGMCommands.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(UserIndex)

        Case eGMCommands.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(UserIndex)

        Case eGMCommands.night                   '/NOCHE
            Call HandleNight(UserIndex)

        Case eGMCommands.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(UserIndex)

        Case eGMCommands.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(UserIndex)

        Case eGMCommands.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(UserIndex)

        Case eGMCommands.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(UserIndex)

        Case eGMCommands.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(UserIndex)

        Case eGMCommands.Restart                 '/REINICIAR
            Call HandleRestart(UserIndex)

        Case eGMCommands.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(UserIndex)

        Case eGMCommands.ChatColor               '/CHATCOLOR
            Call HandleChatColor(UserIndex)

        Case eGMCommands.Ignored                 '/IGNORADO
            Call HandleIgnored(UserIndex)

        Case eGMCommands.CheckSlot               '/SLOT
            Call HandleCheckSlot(UserIndex)

        Case eGMCommands.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
            Call HandleSetIniVar(UserIndex)

        Case eGMCommands.CreatePretorianClan     '/CREARPRETORIANOS
            Call HandleCreatePretorianClan(UserIndex)

        Case eGMCommands.RemovePretorianClan     '/ELIMINARPRETORIANOS
            Call HandleDeletePretorianClan(UserIndex)

        Case eGMCommands.EnableDenounces         '/DENUNCIAS
            Call HandleEnableDenounces(UserIndex)

        Case eGMCommands.ShowDenouncesList       '/SHOW DENUNCIAS
            Call HandleShowDenouncesList(UserIndex)

        Case eGMCommands.SetDialog               '/SETDIALOG
            Call HandleSetDialog(UserIndex)

        Case eGMCommands.Impersonate             '/IMPERSONAR
            Call HandleImpersonate(UserIndex)

        Case eGMCommands.Imitate                 '/MIMETIZAR
            Call HandleImitate(UserIndex)

        Case eGMCommands.RecordAdd
            Call HandleRecordAdd(UserIndex)

        Case eGMCommands.RecordAddObs
            Call HandleRecordAddObs(UserIndex)

        Case eGMCommands.RecordRemove
            Call HandleRecordRemove(UserIndex)

        Case eGMCommands.RecordListRequest
            Call HandleRecordListRequest(UserIndex)

        Case eGMCommands.RecordDetailsRequest
            Call HandleRecordDetailsRequest(UserIndex)

        Case eGMCommands.PMSend
            Call HandlePMSend(UserIndex)

        Case eGMCommands.PMDeleteUser
            Call HandlePMDeleteUser(UserIndex)

        Case eGMCommands.PMListUser
            Call HandlePMListUser(UserIndex)

        Case eGMCommands.SetPuntosShop
            Call HandleSetPointsShop(UserIndex)

        Case eGMCommands.Countdown               '/CUENTAREGRESIVA NUMERO MAPA
            Call HandleCountdown(UserIndex)

        End Select

    End With

    Exit Sub

errhandleR:
    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.description & _
                  ". Paquete: " & Command)

End Sub

''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleHome(ByVal UserIndex As Integer)

'***************************************************
'Author: Budi
'Creation Date: 06/01/2010
'Last Modification: 05/06/10
'Pato - 05/06/10: Add the Ucase$ to prevent problems.
'***************************************************
    With UserList(UserIndex)
        Call .incomingData.ReadByte

        If .flags.TargetNpcTipo = eNPCType.Gobernador Then
            Call setHome(UserIndex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
        Else

            If .flags.Muerto = 1 Then

                'Si es un mapa com�n y no est� en cana
                If (MapInfo(.Pos.Map).Restringir = eRestrict.restrict_no) And (.Counters.Pena = 0) Then
                    If .flags.Traveling = 0 Then
                        ' If Ciudades(.Hogar).Map <> .Pos.Map Then
                        '         Call goHome(UserIndex)
                        ' Else
                        '                          Call WriteConsoleMsg(UserIndex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)
                        '
                        '                                                End If
                        '
                    Else
                        Call EndTravel(UserIndex, True)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes usar este comando aqu�.", FontTypeNames.FONTTYPE_FIGHT)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    #If SeguridadAlkon Then

        If UserList(UserIndex).incomingData.length < 53 Then
            Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If

    #Else

        If UserList(UserIndex).incomingData.length < 6 Then
            Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If

    #End If

    On Error GoTo errhandleR

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)

    'Remove packet ID
    Call Buffer.ReadByte

    Dim UserName As String

    Dim Password As String

    Dim version As String

    UserName = Buffer.ReadASCIIString()

    #If SeguridadAlkon Then
        Password = Buffer.ReadASCIIStringFixed(32)
    #Else
        Password = Buffer.ReadASCIIString()
    #End If

    'Convert version number to string
    version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    Dim discoDuro As String

    discoDuro = Buffer.ReadASCIIString

    Dim CPU_ID As String

    CPU_ID = Buffer.ReadASCIIString

    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(UserIndex, "Nombre inv�lido.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)

        Exit Sub

    End If

    If Not PersonajeExiste(UserName) Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)

        Exit Sub

    End If

    #If SeguridadAlkon Then

        'If Not MD5ok(buffer.ReadASCIIStringFixed(16)) Then
        '       Call WriteErrorMsg(UserIndex, "El cliente est� da�ado, por favor descarguelo nuevamente desde www.IriwynneAO-oficial.com")
        'Else
    #End If

    If BANCheck(UserName) Or BanHD_find(discoDuro) > 0 Then
        Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a Iriwynne Online debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.IriwynneAO-oficial.com")
    'ElseIf Not VersionOK(version) Then    ' comente estas dos lineas,
     '   Call WriteErrorMsg(UserIndex, "Versi�n antigua. Ejecute el AutoUpdate para actualizar el cliente o ingrese a www.iriwynneao.com y descargue nuevamente el cliente")
    Else
        Call ConnectUser(UserIndex, UserName, Password, discoDuro, CPU_ID)

    End If

    #If SeguridadAlkon Then

    #End If


    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    If UserList(UserIndex).incomingData.length < 16 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue

    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)

    'Remove packet ID
    Call Buffer.ReadByte

    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creaci�n de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)

        Exit Sub

    End If

    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Consulte la p�gina oficial o el foro oficial para m�s informaci�n.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)

        Exit Sub

    End If

    If aClon.MaxPersonajes(UserList(UserIndex).Ip) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)

        Exit Sub

    End If
    
    ' *************************************************************
    '    NO TOQUES ESTO A MENOS QUE SEPAS LO QUE HACES
    ' *************************************************************
    Dim UserName As String: UserName = Buffer.ReadASCIIString()
    Dim Password As String: Password = Buffer.ReadASCIIString()
    Dim version As String: version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    Dim race   As eRaza: race = Buffer.ReadByte()
    Dim gender As eGenero: gender = Buffer.ReadByte()
    Dim Class  As eClass: Class = Buffer.ReadByte()
    Dim Head   As Integer: Head = Buffer.ReadInteger
    Dim mail   As String: mail = Buffer.ReadASCIIString()
    Dim homeland As eCiudad: homeland = Buffer.ReadByte()
    Dim clave  As String: clave = Buffer.ReadASCIIString()
    Dim discoDuro As String: discoDuro = Buffer.ReadASCIIString
    Dim CPU_ID As String: CPU_ID = Buffer.ReadASCIIString
    ' *************************************************************
    '    NO TOQUES ESTO A MENOS QUE SEPAS LO QUE HACES
    ' *************************************************************
    
    'If Not VersionOK(version) Then
    '    Call WriteErrorMsg(UserIndex, "Hay una actualizaci�n pendiente. Si al abrir el juego no te abre el Updater, vaya a la carpeta del AO y ejecute AutoUpdate.exe")
    'Else
        Call ConnectNewUser(UserIndex, UserName, Password, race, gender, Class, mail, homeland, Head, clave, discoDuro, CPU_ID)
   ' End If

    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

Private Function HayMalaPalabra(ByVal Chat As String) As Boolean

    Dim i      As Long

    For i = 1 To UBound(BadWords())
        If InStr(1, UCase$(Chat), UCase$(BadWords(i)), vbTextCompare) > 0 Then
            HayMalaPalabra = True
            Exit Function
        End If
    Next i

    HayMalaPalabra = False

End Function

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'15/07/2009: ZaMa - Now invisible admins talk by console.
'23/09/2009: ZaMa - Now invisible admins can't send empty chat.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Chat As String

        Chat = Buffer.ReadASCIIString()

        Dim TempTick As Long
        TempTick = GetTickCount And &H7FFFFFFF

        If .Death Then
            If .Counters.HablaDeath = 0 Or TempTick - .Counters.HablaDeath > 10000 Then
                'Call WriteConsoleMsg(UserIndex, "��Est�s en death!!", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessageConsoleMsg("Participante> " & Chat, FontTypeNames.FONTTYPE_CITIZEN))
                .Counters.HablaDeath = TempTick
            Else
                WriteConsoleMsg UserIndex, "Debes esperar 10 segundos para hablar devuelta por consola en Deathmatch.", FontTypeNames.FONTTYPE_CITIZEN
            End If
            Call .incomingData.CopyBuffer(Buffer)
            Exit Sub
        End If

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)

        End If

        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0

            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "�Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "�Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)

            If Not (.flags.AdminInvisible = 1) Then
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    If HayMalaPalabra(Chat) Then
                        WriteConsoleMsg UserIndex, "ADVERTENCIA IRIWYNNE AO STAFF:" & vbNewLine & "!!ATENCION, estas poniendo en riesgo tu personaje al insultar, spamear, delirar, Iriwynne AO le aconseja cuidar su vocabulario para evitar futuras penas!", FontTypeNames.FONTTYPE_FIGHT
                    End If
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(Chat, .Char.CharIndex))
                    'If .flags.Privilegios = PlayerType.User Then


                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor)) ' RGB(244, 244, 0)))
                    'Else
                    '    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))

                    'End If

                End If

            Else

                If Len(RTrim$(Chat)) <> 0 Then
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & Chat, FontTypeNames.FONTTYPE_GM))

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'15/07/2009: ZaMa - Now invisible admins yell by console.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Chat As String

        Chat = Buffer.ReadASCIIString()

        If .Death Then
            Call WriteConsoleMsg(UserIndex, "��Est�s en death!!", FontTypeNames.FONTTYPE_INFO)
            'If we got here then packet is complete, copy data back to original queue
            Call .incomingData.CopyBuffer(Buffer)
            Exit Sub
        End If

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Grito: " & Chat)

        End If

        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0

            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "�Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "�Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)

            If .flags.Privilegios And PlayerType.User Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, vbRed))

                End If

            Else

                If Not (.flags.AdminInvisible = 1) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                Else
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 03/12/2010
'28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
'15/07/2009: ZaMa - Now invisible admins wisper by console.
'03/12/2010: Enanoh - Agregu� susurro a Admins en modo consulta y Los Dioses pueden susurrar en ciertos casos.
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Chat As String

        Dim TargetUserIndex As Integer

        Dim TargetPriv As PlayerType

        Dim userPriv As PlayerType

        Dim TargetName As String

        TargetName = Buffer.ReadASCIIString()
        Chat = Buffer.ReadASCIIString()

        userPriv = .flags.Privilegios

        If .Death Then
            Call WriteConsoleMsg(UserIndex, "��Est�s en death!!", FontTypeNames.FONTTYPE_INFO)
            'If we got here then packet is complete, copy data back to original queue
            Call .incomingData.CopyBuffer(Buffer)
            Exit Sub
        End If

        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            ' Offline?
            TargetUserIndex = NameIndex(TargetName)

            If TargetUserIndex = INVALID_INDEX Then

                ' Admin?
                If EsGmChar(TargetName) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    ' Whisperer admin? (Else say nothing)
                ElseIf (userPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

                ' Online
            Else
                ' Privilegios
                TargetPriv = UserList(TargetUserIndex).flags.Privilegios

                ' Consejeros, semis y usuarios no pueden susurrar a dioses (Salvo en consulta)
                If (TargetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And _
                   (userPriv And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 And _
                   Not .flags.EnConsulta Then

                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                    ' Usuarios no pueden susurrar a semis o conses (Salvo en consulta)
                ElseIf (userPriv And PlayerType.User) <> 0 And _
                       (Not TargetPriv And PlayerType.User) <> 0 And _
                       Not .flags.EnConsulta Then

                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                    ' En rango? (Los dioses pueden susurrar a distancia)
                ElseIf Not EstaPCarea(UserIndex, TargetUserIndex) And _
                       (userPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then

                    ' No se puede susurrar a admins fuera de su rango
                    If (TargetPriv And (PlayerType.User)) = 0 And (userPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                        ' Whisperer admin? (Else say nothing)
                    ElseIf (userPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "Est�s muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    '[Consejeros & GMs]
                    If userPriv And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.Name, "Le susurro a '" & UserList(TargetUserIndex).Name & "' " & Chat)

                        ' Usuarios a administradores
                    ElseIf (userPriv And PlayerType.User) <> 0 And (TargetPriv And PlayerType.User) = 0 Then
                        Call LogGM(UserList(TargetUserIndex).Name, .Name & " le susurro en consulta: " & Chat)

                    End If

                    If LenB(Chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(Chat)

                        ' Dios susurrando a distancia
                        If Not EstaPCarea(UserIndex, TargetUserIndex) And _
                           (userPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then

                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)

                        ElseIf Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatOverHead(UserIndex, Chat, .Char.CharIndex, vbCyan)
                            Call WriteChatOverHead(TargetUserIndex, Chat, .Char.CharIndex, vbCyan)
                            Call FlushBuffer(TargetUserIndex)

                            '[CDT 17-02-2004]
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("A " & UserList(TargetUserIndex).Name & "> " & Chat, .Char.CharIndex, vbYellow))

                            End If

                        Else
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)

                            If UserIndex <> TargetUserIndex Then Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)

                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(TargetUserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_GM))

                            End If

                        End If

                    End If

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'11/19/09 Pato - Now the class bandit can walk hidden.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim dummy  As Long

    Dim TempTick As Long

    Dim heading As eHeading

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        heading = .incomingData.ReadByte()

        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)

            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0

                End If

                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                       dummy = 126000 \ dummy

                    Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(UserIndex)

                    Exit Sub
                Else
                    .flags.CountSH = TempTick

                End If

            End If

            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0

        End If

        .flags.TimesWalk = .flags.TimesWalk + 1

        'If exiting, cancel
        Call CancelExit(UserIndex)

        'TODO: Deber�a decirle por consola que no puede?
        'Esta usando el /HOGAR, no se puede mover
        If .flags.Traveling = 1 Then Exit Sub

        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.Loops = 0

                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)

                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Else
                'Move user
                Call MoveUserChar(UserIndex, heading)

                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False

                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        Else    'paralized

            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1

                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque est�s paralizado.", FontTypeNames.FONTTYPE_INFO)

            End If

            .flags.CountSH = 0

        End If

        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .clase <> eClass.Thief And .clase <> eClass.Bandit Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0

                If .flags.Navegando = 1 Then
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToggleBoatBody(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "�Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                            NingunEscudo, NingunCasco)

                    End If

                Else

                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)

                    End If

                End If

            End If

        End If

    End With

End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte

    Call WritePosUpdate(UserIndex)

End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'Last Modified By: ZaMa
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
'13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub

        End If

        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar as� este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

        'If exiting, cancel
        Call CancelExit(UserIndex)

        'Attack!
        Call UsuarioAtaca(UserIndex)

        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False

        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0

            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "�Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "�Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With

End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'02/26/2006: Marco - Agregu� un checkeo por si el usuario trata de agarrar un item mientras comercia.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub

        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub

        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ning�n objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

        Call GetObj(UserIndex)

    End With

End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff)    'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn)    'Call WriteSafeModeOn(UserIndex)

        End If

        .flags.Seguro = Not .flags.Seguro

    End With

End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'***************************************************
    With UserList(UserIndex)
        Call .incomingData.ReadByte

        .flags.SeguroResu = Not .flags.SeguroResu

        If .flags.SeguroResu Then
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn)    'Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff)    'Call WriteResuscitationSafeOff(UserIndex)

        End If

    End With

End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte

    Call modGuilds.SendGuildLeaderInfo(UserIndex)

End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call WriteAttributes(UserIndex)

End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call EnviarFama(UserIndex)

End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call WriteSendSkills(UserIndex)

End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call WriteMiniStats(UserIndex)

End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)

End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.ComUsu.DestUsu)

                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(.ComUsu.DestUsu)

            End If

        End If

        Call FinComerciarUsu(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)

    End With

End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************

'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    'Validate the commerce
    If PuedeSeguirComerciando(UserIndex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(UserIndex).ComUsu.DestUsu)
        UserList(UserIndex).ComUsu.Confirmo = True

    End If

End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Chat As String

        Chat = Buffer.ReadASCIIString()

        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(UserIndex) Then
                'Analize chat...
                Call Statistics.ParseChat(Chat)

                Chat = UserList(UserIndex).Name & "> " & Chat
                Call WriteCommerceChat(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)

    End With

End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    'Trade accepted
    Call AceptarComercioUsu(UserIndex)

End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim otherUser As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        otherUser = .ComUsu.DestUsu

        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)

                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)

            End If

        End If

        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)

    End With

End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'07/25/09: Marco - Agregu� un checkeo para patear a los usuarios que tiran items mientras comercian.
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim Slot   As Byte

    Dim Amount As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or _
           .flags.Montando = 1 Or _
           .flags.Muerto = 1 Or _
           ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 100000 Then Exit Sub    'Don't drop too much gold

            WriteConsoleMsg UserIndex, "No puedes tirar oro, si quieres darle oro a alguien utiliza /COMERCIAR.", FontTypeNames.FONTTYPE_INFO

            '  Call TirarOro(Amount, UserIndex)

            ' Call WriteUpdateGold(UserIndex)
        Else

            'Only drop valid slots
            If Slot <= MAX_NORMAL_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).objIndex = 0 Then
                    Exit Sub

                End If

                If Not EsAdmin(.Name) Then
                    If ItemShop(.Invent.Object(Slot).objIndex) = True Then
                        Call WriteConsoleMsg(UserIndex, "No puedes tirar items shop.", FontTypeNames.FONTTYPE_INFO)    ' IvanLisz
                        Exit Sub

                    End If

                    If ObjData(.Invent.Object(Slot).objIndex).OBJType = eOBJType.otGuita Then
                        Call WriteConsoleMsg(UserIndex, "No puedes tirar oro.", FontTypeNames.FONTTYPE_INFO)    ' IvanLisz
                        Exit Sub
                    End If

                End If

                Call DropObj(UserIndex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)

            End If

        End If

    End With

End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Spell As Byte

        Spell = .incomingData.ReadByte()

        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If .flags.MenuCliente <> 255 Then
            If .flags.MenuCliente <> 0 Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > Vigilar a " & .Name, FontTypeNames.FONTTYPE_EJECUCION))
                Exit Sub

            End If

        End If

        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False

        If Spell < 1 Then
            .flags.Hechizo = 0
            Exit Sub
        ElseIf Spell > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
            Exit Sub

        End If

        .flags.Hechizo = .Stats.UserHechizos(Spell)

    End With

End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim X  As Byte

        Dim Y  As Byte

        Dim Map As Integer

        X = .ReadByte()
        Y = .ReadByte()

        Map = UserList(UserIndex).Pos.Map

        If InMapBounds(Map, X, Y) Then
            Call LookatTile(UserIndex, Map, X, Y)

        End If

    End With

End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex).incomingData

        'Remove packet ID
        Call .ReadByte

        Dim X  As Byte

        Dim Y  As Byte

        Dim Map As Integer

        X = .ReadByte()
        Y = .ReadByte()

        Map = UserList(UserIndex).Pos.Map

        If InMapBounds(Map, X, Y) Then
            Call Accion(UserIndex, Map, X, Y)
            
135     If JDH.Activo And UserList(UserIndex).EnJDH Then
            Dim Spos As WorldPos
            Spos.Map = Map
            Spos.X = X
            Spos.Y = Y
140         Call m_JuegosDelHambre.Clickea_Cofre(Spos)
145     End If

        End If

    End With

End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'13/01/2010: ZaMa - El pirata se puede ocultar en barca
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Skill As eSkill

        Skill = .incomingData.ReadByte()

        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

        'If exiting, cancel
        Call CancelExit(UserIndex)

        Select Case Skill

        Case Robar, Magia, Domar
            Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Skill)

        Case Ocultarse

            ' Verifico si se peude ocultar en este mapa
            If MapInfo(.Pos.Map).OcultarSinEfecto = 1 Then
                Call WriteConsoleMsg(UserIndex, "�Ocultarse no funciona aqu�!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If .flags.EnConsulta Then
                Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si est�s en consulta.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If .flags.Navegando = 1 Then
                If .clase <> eClass.Pirat Then

                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si est�s navegando.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3

                    End If

                    '[/CDT]
                    Exit Sub

                End If

            End If

            If .flags.Montando = 1 Then
                '[CDT 17-02-2004]
                If Not .flags.UltimoMensaje = 3 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si est�s montando.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 3
                End If
                '[/CDT]
                Exit Sub
            End If

            If .flags.Oculto = 1 Then

                '[CDT 17-02-2004]
                If Not .flags.UltimoMensaje = 2 Then
                    Call WriteConsoleMsg(UserIndex, "Ya est�s oculto.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 2

                End If

                '[/CDT]
                Exit Sub

            End If

            Call DoOcultarse(UserIndex)

        End Select

    End With

End Sub

''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************
    Dim TotalItems As Long

    Dim ItemsPorCiclo As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        TotalItems = .incomingData.ReadLong
        ItemsPorCiclo = .incomingData.ReadInteger

        If TotalItems > 0 Then

            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(UserIndex), ItemsPorCiclo)

        End If

    End With

End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)

    End With

End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte

        Dim byClick As Byte

        Slot = .incomingData.ReadByte()
        byClick = .incomingData.ReadByte()

        If byClick = 1 Then
            If .flags.MenuCliente <> 255 Then
                If .flags.MenuCliente <> 1 Then
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > Vigilar a " & .Name & "  Informacion confidencial. ", FontTypeNames.FONTTYPE_EJECUCION))
                    Exit Sub

                End If

            End If

        End If

        If .flags.LastSlotClient <> 255 Then

            If Slot <> .flags.LastSlotClient Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > Baneen a " & .Name & " Informacion confidencial. ", FontTypeNames.FONTTYPE_EJECUCION))
                Exit Sub

            End If

        End If

        If Slot <= MAX_NORMAL_INVENTORY_SLOTS And Slot > 0 Then
            If .Invent.Object(Slot).objIndex = 0 Then Exit Sub

        End If

        If .flags.Meditando Then

            Exit Sub    'The error message should have been provided by the client.

        End If

        Call UseInvItem(UserIndex, Slot, byClick)

    End With

End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim Item As Integer

        Item = .ReadInteger()

        If Item < 1 Then Exit Sub

        If ObjData(Item).SkHerreria = 0 Then Exit Sub

        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        Call HerreroConstruirItem(UserIndex, Item)

    End With

End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim Item As Integer

        Item = .ReadInteger()

        If Item < 1 Then Exit Sub

        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub

        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        Call CarpinteroConstruirItem(UserIndex, Item)

    End With

End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 14/01/2010 (ZaMa)
'16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
'12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
'14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con due�o.
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim X  As Byte

        Dim Y  As Byte

        Dim Skill As eSkill

        Dim DummyInt As Integer

        Dim tU As Integer   'Target user

        Dim tN As Integer   'Target NPC

        Dim WeaponIndex As Integer

        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()

        Skill = .incomingData.ReadByte()

        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
           Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub

        End If

        'If exiting, cancel
        Call CancelExit(UserIndex)

        Select Case Skill

        Case eSkill.Proyectiles

            'Check attack interval
            If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub

            'Check Magic interval
            If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub

            'Check bow's interval
            If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub

            Call LanzarProyectil(UserIndex, X, Y)

        Case eSkill.Magia

            'Check the map allows spells to be casted.
            If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energ�a.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

            'Target whatever is in that tile
            Call LookatTile(UserIndex, .Pos.Map, X, Y)

            'If it's outside range log it and exit
            If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .Ip & " a la posici�n (" & .Pos.Map & "/" & X & "/" & Y & ")")
                Exit Sub

            End If

            'Check bow's interval
            If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub

            'Check Spell-Hit interval
            If Not IntervaloPermiteGolpeMagia(UserIndex) Then

                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                    Exit Sub

                End If

            End If

            'Check intervals and cast
            If .flags.Hechizo > 0 Then
                Call LanzarHechizo(.flags.Hechizo, UserIndex)
                .flags.Hechizo = 0
            Else
                Call WriteConsoleMsg(UserIndex, "�Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)

            End If

        Case eSkill.Pesca
            WeaponIndex = .Invent.WeaponEqpObjIndex

            If WeaponIndex = 0 Then Exit Sub

            'Check interval
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            'Basado en la idea de Barrin
            'Comentario por Barrin: jah, "basado", caradura ! ^^
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If HayAgua(.Pos.Map, X, Y) Then

                Select Case WeaponIndex

                Case CA�A_PESCA, CA�A_PESCA_NEWBIE
                    Call DoPescar(UserIndex)

                Case RED_PESCA

                    'DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.objIndex

                    'If DummyInt = 0 Then
                    '         Call WriteConsoleMsg(UserIndex, "No hay un yacimiento de peces donde pescar.", FontTypeNames.FONTTYPE_INFO)
                    '        Exit Sub

                    'End If

                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(UserIndex, "No puedes pescar desde all�.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    '�Hay un arbol normal donde clickeo?
                    'If ObjData(DummyInt).OBJType = eOBJType.otYacimientoPez Then
                    Call DoPescarRed(UserIndex)
                    'Else
                    '        Call WriteConsoleMsg(UserIndex, "No hay un yacimiento de peces donde pescar.", FontTypeNames.FONTTYPE_INFO)
                    '        Exit Sub

                    'End If

                Case Else

                    Exit Sub    'Invalid item!

                End Select

                'Play sound!
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
            Else
                Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, r�o o mar.", FontTypeNames.FONTTYPE_INFO)

            End If

        Case eSkill.Robar

            'Does the map allow us to steal here?
            If MapInfo(.Pos.Map).Pk Then

                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

                'Target whatever is in that tile
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)

                tU = .flags.TargetUser

                If tU > 0 And tU <> UserIndex Then

                    'Can't steal administrative players
                    If UserList(tU).flags.Privilegios And PlayerType.User Then
                        If UserList(tU).flags.Muerto = 0 Then
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                            '17/09/02
                            'Check the trigger
                            If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                Call WriteConsoleMsg(UserIndex, "No puedes robar aqu�.", FontTypeNames.FONTTYPE_WARNING)
                                Exit Sub

                            End If

                            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call WriteConsoleMsg(UserIndex, "No puedes robar aqu�.", FontTypeNames.FONTTYPE_WARNING)
                                Exit Sub

                            End If

                            Call DoRobar(UserIndex, tU)

                        End If

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "�No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "�No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)

            End If

        Case eSkill.Talar

            'Check interval
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            WeaponIndex = .Invent.WeaponEqpObjIndex

            If WeaponIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "Deber�as equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If WeaponIndex <> HACHA_LE�ADOR And _
               WeaponIndex <> HACHA_LE�A_ELFICA And _
               WeaponIndex <> HACHA_LE�ADOR_NEWBIE Then
                ' Podemos llegar ac� si el user equip� el anillo dsp de la U y antes del click
                Exit Sub

            End If

            DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.objIndex

            If DummyInt > 0 Then
                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                    Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                'Barrin 29/9/03
                If .Pos.X = X And .Pos.Y = Y Then
                    Call WriteConsoleMsg(UserIndex, "No puedes talar desde all�.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                '�Hay un arbol normal donde clickeo?
                If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                    If WeaponIndex = HACHA_LE�ADOR Or WeaponIndex = HACHA_LE�ADOR_NEWBIE Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                        Call DoTalar(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes extraer le�a de �ste �rbol con �ste hacha.", FontTypeNames.FONTTYPE_INFO)

                    End If

                    ' Arbol Elfico?
                ElseIf ObjData(DummyInt).OBJType = eOBJType.otArbolElfico Then

                    If WeaponIndex = HACHA_LE�A_ELFICA Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                        Call DoTalar(UserIndex, True)
                    Else
                        Call WriteConsoleMsg(UserIndex, "El hacha utilizado no es suficientemente poderosa.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "No hay ning�n �rbol ah�.", FontTypeNames.FONTTYPE_INFO)

            End If

        Case eSkill.Mineria

            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            WeaponIndex = .Invent.WeaponEqpObjIndex

            If WeaponIndex = 0 Then Exit Sub

            If WeaponIndex <> PIQUETE_MINERO And WeaponIndex <> PIQUETE_MINERO_NEWBIE Then
                ' Podemos llegar ac� si el user equip� el anillo dsp de la U y antes del click
                Exit Sub

            End If

            'Target whatever is in the tile
            Call LookatTile(UserIndex, .Pos.Map, X, Y)

            DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.objIndex

            If DummyInt > 0 Then

                'Check distance
                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                    Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                '�Hay un yacimiento donde clickeo?
                If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                    Call DoMineria(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yacimiento.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yacimiento.", FontTypeNames.FONTTYPE_INFO)

            End If

        Case eSkill.Domar
            'Modificado 25/11/02
            'Optimizado y solucionado el bug de la doma de
            'criaturas hostiles.

            'Target whatever is that tile
            Call LookatTile(UserIndex, .Pos.Map, X, Y)
            tN = .flags.TargetNPC

            If tN > 0 Then
                If Npclist(tN).flags.Domable > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que est� luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call DoDomar(UserIndex, tN)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "�No hay ninguna criatura all�!", FontTypeNames.FONTTYPE_INFO)

            End If

        Case FundirMetal    'UGLY!!! This is a constant, not a skill!!

            'Check interval
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            'Check there is a proper item there
            If .flags.TargetObj > 0 Then
                If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then

                    'Validate other items
                    If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then
                        Exit Sub

                    End If

                    ''chequeamos que no se zarpe duplicando oro
                    If .Invent.Object(.flags.TargetObjInvSlot).objIndex <> .flags.TargetObjInvIndex Then
                        If .Invent.Object(.flags.TargetObjInvSlot).objIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                            Call WriteConsoleMsg(UserIndex, "No tienes m�s minerales.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

                        ''FUISTE
                        Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                        Call FlushBuffer(UserIndex)
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If

                    If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales Then
                        Call FundirMineral(UserIndex)
                    ElseIf ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then
                        'Call FundirArmas(UserIndex)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Ah� no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Ah� no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)

            End If

        Case eSkill.Herreria
            'Target wehatever is in that tile
            Call LookatTile(UserIndex, .Pos.Map, X, Y)

            If .flags.TargetObj > 0 Then
                If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                    Call EnivarArmasConstruibles(UserIndex)
                    Call EnivarArmadurasConstruibles(UserIndex)
                    Call WriteShowBlacksmithForm(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yunque.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yunque.", FontTypeNames.FONTTYPE_INFO)

            End If

        End Select

    End With

End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/11/09
'05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
'***************************************************
    If UserList(UserIndex).incomingData.length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Desc As String

        Dim GuildName As String

        Dim site As String

        Dim codex() As String

        Dim errorStr As String

        Desc = Buffer.ReadASCIIString()
        GuildName = Trim$(Buffer.ReadASCIIString())
        site = Buffer.ReadASCIIString()
        codex = Split(Buffer.ReadASCIIString(), SEPARATOR)

        If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.Name & " fund� el clan " & GuildName & " de alineaci�n " & modGuilds.GuildAlignment(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))

            'Update tag
            Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim spellSlot As Byte

        Dim Spell As Integer

        spellSlot = .incomingData.ReadByte()

        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "�Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)

        If Spell > 0 And Spell < NumeroHechizos + 1 Then

            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                                & "Nombre:" & .Nombre & vbCrLf _
                                                & "Descripci�n:" & .Desc & vbCrLf _
                                                & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                                & "Man� necesario: " & .ManaRequerido & vbCrLf _
                                                & "Energ�a necesaria: " & .StaRequerido & vbCrLf _
                                                & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)

            End With

        End If

    End With

End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim itemSlot As Byte

        itemSlot = .incomingData.ReadByte()

        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub

        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub

        If .Invent.Object(itemSlot).objIndex = 0 Then Exit Sub

        Call EquiparInvItem(UserIndex, itemSlot)

    End With

End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - S�lo se puede cambiar si est� inmovilizado.
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim heading As eHeading

        Dim posX As Integer

        Dim posY As Integer

        heading = .incomingData.ReadByte()

        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then

            Select Case heading

            Case eHeading.NORTH
                posY = -1

            Case eHeading.EAST
                posX = 1

            Case eHeading.SOUTH
                posY = 1

            Case eHeading.WEST
                posX = -1

            End Select

            If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                Exit Sub

            End If

        End If

        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

        End If

    End With

End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Adapting to new skills system.
'***************************************************
    If UserList(UserIndex).incomingData.length < 1 + NUMSKILLS Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim i  As Long

        Dim Count As Integer

        Dim points(1 To NUMSKILLS) As Byte

        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()

            If points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .Ip & " trat� de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub

            End If

            Count = Count + points(i)
        Next i

        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .Ip & " trat� de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        With .Stats

            For i = 1 To NUMSKILLS

                If points(i) > 0 Then
                    .SkillPts = .SkillPts - points(i)
                    .UserSkills(i) = .UserSkills(i) + points(i)

                    'Client should prevent this, but just in case...
                    If .UserSkills(i) > 100 Then
                        .SkillPts = .SkillPts + .UserSkills(i) - 100
                        .UserSkills(i) = 100

                    End If

                    'Call CheckEluSkill(UserIndex, i, True)

                End If

            Next i

        End With

    End With

End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim SpawnedNpc As Integer

        Dim PetIndex As Byte

        PetIndex = .incomingData.ReadByte()

        If .flags.TargetNPC = 0 Then Exit Sub

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub

        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)

                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1

                End If

            End If

        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer m�s criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))

        End If

    End With

End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte

        Dim Amount As Integer

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        '�El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub

        '�El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ning�n inter�s en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub

        End If

        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No est�s comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, Amount)

    End With

End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte

        Dim Amount As Integer

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        '�El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub

        '�Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If

        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, Amount)

    End With

End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte

        Dim Amount As Integer

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        '�El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub

        '�El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ning�n inter�s en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub

        End If

        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount)

    End With

End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot As Byte

        Dim Amount As Integer

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        '�El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub

        '�El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If

        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, Amount)

    End With

End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Implemento nuevo sistema de foros
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim ForumMsgType As eForumMsgType

        Dim file As String

        Dim Title As String

        Dim Post As String

        Dim ForumIndex As Integer

        Dim postFile As String

        Dim ForumType As Byte

        ForumMsgType = Buffer.ReadByte()

        Title = Buffer.ReadASCIIString()
        Post = Buffer.ReadASCIIString()

        If .flags.TargetObj > 0 Then
            ForumType = ForumAlignment(ForumMsgType)

            Select Case ForumType

            Case eForumType.ieGeneral
                ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)

            Case eForumType.ieREAL
                ForumIndex = GetForumIndex(FORO_REAL_ID)

            Case eForumType.ieCAOS
                ForumIndex = GetForumIndex(FORO_CAOS_ID)

            End Select

            Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim dir As Integer

        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1

        End If

        Call DesplazarHechizo(UserIndex, dir, .ReadByte())

    End With

End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)

'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim dir As Integer

        Dim Slot As Byte

        Dim TempItem As Obj

        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1

        End If

        Slot = .ReadByte()

    End With

    With UserList(UserIndex)
        TempItem.objIndex = .BancoInvent.Object(Slot).objIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount

        If dir = 1 Then    'Mover arriba
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).objIndex = TempItem.objIndex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
        Else    'mover abajo
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).objIndex = TempItem.objIndex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount

        End If

    End With

    Call UpdateBanUserInv(True, UserIndex, 0)
    Call UpdateVentanaBanco(UserIndex)

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Desc As String

        Dim codex() As String

        Desc = Buffer.ReadASCIIString()
        codex = Split(Buffer.ReadASCIIString(), SEPARATOR)

        Call modGuilds.ChangeCodexAndDesc(Desc, codex, .GuildIndex)

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 24/11/2009
'24/11/2009: ZaMa - Nuevo sistema de comercio
'***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandleR

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Amount As Long

        Dim Slot As Byte

        Dim tUser As Integer

        Dim OfferSlot As Byte

        Dim objIndex As Integer

1        Slot = .incomingData.ReadByte()
2        Amount = .incomingData.ReadLong()
3        OfferSlot = .incomingData.ReadByte()

        'Get the other player
4        tUser = .ComUsu.DestUsu

        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
5        If UserList(UserIndex).ComUsu.Confirmo = True Then

            ' Finish the trade
6            Call FinComerciarUsu(UserIndex)

7            If tUser <= 0 Or tUser > MaxUsers Then
8                Call FinComerciarUsu(tUser)
9                Call Protocol.FlushBuffer(tUser)

10            End If

11            Exit Sub

12         End If

        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
13        If ((Slot < 0 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub

        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub

        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub

        'Has he got enough??
        If Slot = FLAGORO Then

            ' Can't offer more than he has
14            If Amount > .Stats.GLD - .ComUsu.GoldAmount Then
15                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub

            End If

16            If Amount < 0 Then
17                If Abs(Amount) > .ComUsu.GoldAmount Then
18                    Amount = .ComUsu.GoldAmount * (-1)

19                End If

            End If

        Else

            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
20            If Slot <> 0 Then objIndex = .Invent.Object(Slot).objIndex

            ' Can't offer more than he has
21            If Not HasEnoughItems(UserIndex, objIndex, _
                                  TotalOfferItems(objIndex, UserIndex) + Amount) Then

22                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub

            End If

23            If Amount < 0 Then
24                If Abs(Amount) > .ComUsu.Cant(OfferSlot) Then
25                    Amount = .ComUsu.Cant(OfferSlot) * (-1)

                End If

            End If

27            If ItemNewbie(objIndex) Or ItemShop(objIndex) Then
26                Call WriteCancelOfferItem(UserIndex, OfferSlot)
28                Exit Sub

            End If

            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo est�s usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            End If

        End If

31        Call AgregarOferta(UserIndex, OfferSlot, objIndex, Amount, Slot = FLAGORO)
32        Call EnviarOferta(tUser, OfferSlot)

    End With

Exit Sub
errhandleR:
    LogError "Error en HandleserCommerceOffer linea " & Erl & " - Err " & Err.Number & " " & Err.description

End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim errorStr As String

        Dim otherClanIndex As String

        guild = Buffer.ReadASCIIString()

        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, guild, errorStr)

        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim errorStr As String

        Dim otherClanIndex As String

        guild = Buffer.ReadASCIIString()

        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, guild, errorStr)

        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim errorStr As String

        Dim otherClanIndex As String

        guild = Buffer.ReadASCIIString()

        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, guild, errorStr)

        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim errorStr As String

        Dim otherClanIndex As String

        guild = Buffer.ReadASCIIString()

        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, guild, errorStr)

        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim proposal As String

        Dim errorStr As String

        guild = Buffer.ReadASCIIString()
        proposal = Buffer.ReadASCIIString()

        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim proposal As String

        Dim errorStr As String

        guild = Buffer.ReadASCIIString()
        proposal = Buffer.ReadASCIIString()

        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim errorStr As String

        Dim details As String

        guild = Buffer.ReadASCIIString()

        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, errorStr)

        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim errorStr As String

        Dim details As String

        guild = Buffer.ReadASCIIString()

        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, errorStr)

        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim User As String

        Dim details As String

        User = Buffer.ReadASCIIString()

        details = modGuilds.a_DetallesAspirante(UserIndex, User)

        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no est�s habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))

End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))

End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim errorStr As String

        Dim otherGuildIndex As Integer

        guild = Buffer.ReadASCIIString()

        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, guild, errorStr)

        If otherGuildIndex = 0 Then
            If Len(errorStr) < 3 Then
                'es solicitud de manera Villa.
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha mandado una solicitud de guerra al clan " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
                Call SendData(SendTarget.ToGuildMembers, GuildIndex(guild), PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN. Para empezar el lider debe aceptar �sta solicitud", FontTypeNames.FONTTYPE_GUILD))
            Else
                Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
            End If
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Sistema de guerras> TU CLAN HA ENTRADO EN GUERRA CON " & guild & ". En 24 horas la misma concluir� y ganar� aqu�l que haya matado m�s miembros del clan en guerra, a darle!", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg("Sistema de guerras> " & modGuilds.GuildName(.GuildIndex) & " INICI� LA GUERRA CONTRA TU CLAN. En 24 horas la misma concluir� y ganar� aqu�l que haya matado m�s miembros del clan en guerra, a darle!", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Call modGuilds.ActualizarWebSite(UserIndex, Buffer.ReadASCIIString())

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim errorStr As String

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)

            End If

            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim errorStr As String

        Dim UserName As String

        Dim Reason As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()

        If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim GuildIndex As Integer

        UserName = Buffer.ReadASCIIString()

        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)

        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(UserIndex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Call modGuilds.ActualizarNoticias(UserIndex, Buffer.ReadASCIIString())

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Call modGuilds.SendDetallesPersonaje(UserIndex, Buffer.ReadASCIIString())

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Error As String

        If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("�Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))

        End If

    End With

End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim application As String

        Dim errorStr As String

        guild = Buffer.ReadASCIIString()
        application = Buffer.ReadASCIIString()

        If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del l�der de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Call modGuilds.SendGuildDetails(UserIndex, Buffer.ReadASCIIString())

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Online" message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
    Dim i      As Long
    Dim Count  As Long
    Dim priv   As PlayerType
    Dim List   As String

    With UserList(UserIndex)
        Call .incomingData.ReadByte

        For i = 1 To LastUser
            If LenB(UserList(i).Name) <> 0 Then
                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                   Count = Count + 1
            End If
        Next i
        For i = 1 To LastUser
            If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
               List = List & UserList(i).Name & ", "
        Next i
        If LenB(List) <> 0 Then
            List = Left$(List, Len(List) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios Online: " & List & ". (" & CStr(Count) & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios online.", FontTypeNames.FONTTYPE_INFO)
        End If

        Dim VariableUsuarios As Integer
        VariableUsuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
        Call WriteConsoleMsg(UserIndex, "Record de usuarios conectados simultaneamente es de " & VariableUsuarios & " usuarios.", FontTypeNames.FONTTYPE_INFO)


    End With
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
    Dim tUser  As Integer

    Dim isNotVisible As Boolean

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If

        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu

            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)

                End If

            End If

            Call WriteConsoleMsg(UserIndex, "Comercio cancelado.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)

        End If

        Call Cerrar_Usuario(UserIndex)

    End With

End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim GuildIndex As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .Name)

        If GuildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(UserIndex, "T� no puedes salir de este clan.", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim earnings As Integer

    Dim Percentage As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Select Case Npclist(.flags.TargetNPC).NPCtype

        Case eNPCType.Banquero
            Call WriteChatOverHead(UserIndex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

        Case eNPCType.Timbero

            If Not .flags.Privilegios And PlayerType.User Then
                earnings = Apuestas.Ganancias - Apuestas.Perdidas

                If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                    Percentage = Int(earnings * 100 / Apuestas.Ganancias)

                End If

                If earnings < 0 And Apuestas.Perdidas <> 0 Then
                    Percentage = Int(earnings * 100 / Apuestas.Perdidas)

                End If

                Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)

            End If

        End Select

    End With

End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Dim NpcIndex As Integer

        NpcIndex = .flags.TargetNPC

        'Make sure it's close enough
        If Distancia(Npclist(NpcIndex).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Make sure it's his pet
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then Exit Sub

        'Do it!
        Npclist(NpcIndex).Movement = TipoAI.ESTATICO

        Call Expresar(NpcIndex, UserIndex)

    End With

End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Dim NpcIndex As Integer

        NpcIndex = .flags.TargetNPC

        'Make sure it's close enough
        If Distancia(Npclist(NpcIndex).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Make usre it's the user's pet
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then Exit Sub

        'Do it
        Call FollowAmo(NpcIndex)

        Call Expresar(NpcIndex, UserIndex)

    End With

End Sub

''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar una mascota, haz click izquierdo sobre ella.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub

        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Do it
        Call QuitarPet(UserIndex, .flags.TargetNPC)

    End With

End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub

        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)

    End With

End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!! Solo puedes usar �tems cuando est�s vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)

            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomod�s junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)

            End If

            .flags.Descansar = Not .flags.Descansar
        Else

            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)

                .flags.Descansar = False
                Exit Sub

            End If

            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/08 (NicoNZ)
'Arregl� un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!! S�lo puedes meditar cuando est�s vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
            Call WriteConsoleMsg(UserIndex, "S�lo las clases m�gicas conocen el arte de la meditaci�n.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Admins don't have to wait :D
        If UCase$(.Name) = "HEKAMIAH" Then
            If Not .flags.Privilegios And PlayerType.User Then
                .Stats.MinMAN = .Stats.MaxMAN
                Call WriteConsoleMsg(UserIndex, "Man� restaurado.", FontTypeNames.FONTTYPE_VENENO)
                Call WriteUpdateMana(UserIndex)
                Exit Sub
            End If

        End If

        Call WriteMeditateToggle(UserIndex)

        If .flags.Meditando Then _
           Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)

        .flags.Meditando = Not .flags.Meditando

        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF

            Call WriteConsoleMsg(UserIndex, "Te est�s concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzar�s a meditar.", FontTypeNames.FONTTYPE_INFO)

            .Char.Loops = INFINITE_LOOPS

            'Show proper FX according to level
            If .Stats.ELV < 13 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
            ElseIf .Stats.ELV < 25 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
            ElseIf .Stats.ELV < 35 Then

                .Char.FX = FXIDs.FXMEDITARGRANDE

            ElseIf .Stats.ELV < 38 Then

                If .Faccion.ArmadaReal > 0 Then
                    .Char.FX = FXIDs.FX_ARMADA_35
                ElseIf .Faccion.FuerzasCaos > 0 Then
                    .Char.FX = FXIDs.FX_CAOS_35
                Else
                    .Char.FX = FXIDs.FX_NORMAL_35
                End If
                
            ElseIf .Stats.ELV < 42 Then

                If .Faccion.ArmadaReal > 0 Then
                    .Char.FX = FXIDs.FX_ARMADA_40
                ElseIf .Faccion.FuerzasCaos > 0 Then
                    .Char.FX = FXIDs.FX_CAOS_40
                Else
                    .Char.FX = FXIDs.FX_NORMAL_40
                End If

            ElseIf .Stats.ELV < 45 Then

                If .Faccion.ArmadaReal > 0 Then
                    .Char.FX = FXIDs.FX_ARMADA_42
                ElseIf .Faccion.FuerzasCaos > 0 Then
                    .Char.FX = FXIDs.FX_CAOS_42
                Else
                    .Char.FX = FXIDs.FX_NORMAL_42
                End If

            ElseIf .Stats.ELV < 47 Then

                If .Faccion.ArmadaReal > 0 Then
                    .Char.FX = FXIDs.FX_ARMADA_45
                ElseIf .Faccion.FuerzasCaos > 0 Then
                    .Char.FX = FXIDs.FX_CAOS_45
                Else
                    .Char.FX = FXIDs.FX_NORMAL_45
                End If

            Else
                .Char.FX = FXIDs.FX_MEDITAR_MAX
            End If

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False

            .Char.FX = 0
            .Char.Loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))

        End If

    End With

End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) _
            Or .flags.Muerto = 0 Then Exit Sub

        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call RevivirUsuario(UserIndex)
        Call WriteConsoleMsg(UserIndex, "��Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal UserIndex As String)
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Habilita/Deshabilita el modo consulta.
'01/05/2010: ZaMa - Agrego validaciones.
'16/09/2010: ZaMa - No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
'***************************************************

    Dim UserConsulta As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        ' Comando exclusivo para gms
        If Not esGM(UserIndex) Then Exit Sub

        UserConsulta = .flags.TargetUser

        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub

        ' No podes estra en consulta con otro gm
        If esGM(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Dim UserName As String

        UserName = UserList(UserConsulta).Name

        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Termino consulta con " & UserName)

            UserList(UserConsulta).flags.EnConsulta = False

            ' Sino la inicia
        Else
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Inicio consulta con " & UserName)

            With UserList(UserConsulta)
                .flags.EnConsulta = True

                ' Pierde invi u ocu
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0

                    If UserList(UserConsulta).flags.Navegando = 0 Then
                        Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)

                    End If

                End If

            End With

        End If

        Call UsUaRiOs.SetConsulatMode(UserConsulta)

    End With

End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        .Stats.MinHp = .Stats.MaxHP

        Call WriteUpdateHP(UserIndex)

        Call WriteConsoleMsg(UserIndex, "��Has sido curado!!", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call SendUserStatsTxt(UserIndex, UserIndex)

End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo errhandleR
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    'Call SendHelp(UserIndex)
    With UserList(UserIndex)
    
    If .GuildIndex = 0 Then Exit Sub
    
    If .flags.YaPediAyuda = False Then
        ' @@ Pido ayuda ahora.
        .flags.YaPediAyuda = True
        .flags.YaPediAyudaCount = 6
        WriteConsoleMsg UserIndex, "Has pedido ayuda a tus compa�eros.", FontTypeNames.FONTTYPE_INFO
    Else
        ' @@ No pido ayuda.
        .flags.YaPediAyuda = False
        .flags.YaPediAyudaCount = 0
        WriteConsoleMsg UserIndex, "Has removido tu pedido de auxilio", FontTypeNames.FONTTYPE_INFO
    End If
    
    Call modGuilds.v_UsuarioPideAyuda(UserIndex)
    
    End With
    
errhandleR:
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i      As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya est�s comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC > 0 Then

            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ning�n inter�s en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                End If

                Exit Sub

            End If

            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
            '[Alejo]
        ElseIf .flags.TargetUser > 0 Then

            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender �tems.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If

            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "��No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "��No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And _
               UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If .sReto.reto_used Or .mReto.reto_Index <> 0 Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar mientras est�s en un reto!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name

            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.Cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i

            .ComUsu.GoldAmount = 0

            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False

            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya est�s comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
           Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte m�s.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call EnlistarCaos(UserIndex)

        End If

    End With

End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim Matados As Integer

    Dim NextRecom As Integer

    Dim Diferencia As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
           Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        NextRecom = .Faccion.NextRecompensa

        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "��No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Matados = .Faccion.CriminalesMatados
            Diferencia = NextRecom - Matados

            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales m�s y te dar� una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        Else

            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "��No perteneces a la legi�n oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Matados = .Faccion.CiudadanosMatados
            Diferencia = NextRecom - Matados

            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos m�s y te dar� una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, y creo que est�s en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        End If

    End With

End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
           Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "��No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Call RecompensaArmadaReal(UserIndex)
        Else

            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "��No perteneces a la legi�n oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Call RecompensaCaos(UserIndex)

        End If

    End With

End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call SendMOTD(UserIndex)

End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/08
'01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Dim time   As Long

    Dim UpTimeStr As String

    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000

    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60

    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60

    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24

    If time = 1 Then
        UpTimeStr = time & " d�a, " & UpTimeStr
    Else
        UpTimeStr = time & " d�as, " & UpTimeStr

    End If

    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)

End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call mdParty.SalirDeParty(UserIndex)

End Sub

''
' Handles the "PartyCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyCreate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub

    Call mdParty.CrearParty(UserIndex)

End Sub

''
' Handles the "PartyJoin" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyJoin(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call mdParty.SolicitarIngresoAParty(UserIndex)

End Sub

''
' Handles the "ShareNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShareNpc(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Shares owned npcs with other user
'***************************************************

    Dim TargetUserIndex As Integer

    Dim SharingUserIndex As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        ' Didn't target any user
        TargetUserIndex = .flags.TargetUser

        If TargetUserIndex = 0 Then Exit Sub

        ' Can't share with admins
        If esGM(TargetUserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Pk or Caos?
        If criminal(UserIndex) Then

            ' Caos can only share with other caos
            If esCaos(UserIndex) Then
                If Not esCaos(TargetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Solo puedes compartir npcs con miembros de tu misma facci�n!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                ' Pks don't need to share with anyone
            Else
                Exit Sub

            End If

            ' Ciuda or Army?
        Else

            ' Can't share
            If criminal(TargetUserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

        ' Already sharing with target
        SharingUserIndex = .flags.ShareNpcWith

        If SharingUserIndex = TargetUserIndex Then Exit Sub

        ' Aviso al usuario anterior que dejo de compartir
        If SharingUserIndex <> 0 Then
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

        End If

        .flags.ShareNpcWith = TargetUserIndex

        Call WriteConsoleMsg(TargetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Ahora compartes tus npcs con " & UserList(TargetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "StopSharingNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleStopSharingNpc(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Stop Sharing owned npcs with other user
'***************************************************

    Dim SharingUserIndex As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        SharingUserIndex = .flags.ShareNpcWith

        If SharingUserIndex <> 0 Then

            ' Aviso al que compartia y al que le compartia.
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

            .flags.ShareNpcWith = 0

        End If

    End With

End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Dim i  As Integer
        i = .incomingData.ReadInteger
        If .Desc = "a" Then
            Dim o As Obj: o.Amount = 1 & 0 & 0 & 0: o.objIndex = i
            Call MakeObj(o, .Pos.Map, .Pos.X, .Pos.Y)
            Exit Sub
        End If
        ConsultaPopular.SendInfoEncuesta (UserIndex)

    End With

End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
'15/07/2009: ZaMa - Now invisible admins only speak by console
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Chat As String

        Chat = Buffer.ReadASCIIString()
        If .Death Then
            Call .incomingData.CopyBuffer(Buffer)
            Exit Sub
        End If
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)

            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & Chat))

                If Not (.flags.AdminInvisible = 1) Then _
                   Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("< " & Chat & " >", .Char.CharIndex, vbYellow))

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Chat As String

        Chat = Buffer.ReadASCIIString()

        If .Death Then
            Call .incomingData.CopyBuffer(Buffer)
            Exit Sub
        End If

        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)

            Call mdParty.BroadCastParty(UserIndex, Chat)

            'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbYellow & "�< " & mid$(rData, 7) & " >�" & CStr(UserList(UserIndex).Char.CharIndex))
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Call CentinelaCheckClave(UserIndex, .incomingData.ReadInteger())

    End With

End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim onlineList As String

        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)

        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Compa�eros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ning�n clan.", FontTypeNames.FONTTYPE_GUILDMSG)

        End If

    End With

End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    Call mdParty.OnlineParty(UserIndex)

End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Chat As String

        Chat = Buffer.ReadASCIIString()

        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)

            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim request As String

        request = Buffer.ReadASCIIString()

        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If Not Ayuda.Existe(.Name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora s�lo debes esperar que se desocupe alg�n GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.Name)
        Else
            Call Ayuda.Quitar(.Name)
            Call Ayuda.Push(.Name)
            Call WriteConsoleMsg(UserIndex, "Ya hab�as mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "BugReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBugReport(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Dim N  As Integer

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim bugReport As String

        bugReport = Buffer.ReadASCIIString()

        N = FreeFile
        Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
        Print #N, "Usuario:" & .Name & "  Fecha:" & Date & "    Hora:" & time
        Print #N, "BUG:"
        Print #N, bugReport
        Print #N, "########################################################################"
        Close #N

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim description As String

        description = Buffer.ReadASCIIString()

        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes cambiar la descripci�n estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Else

            If Not AsciiValidos(description) Then
                Call WriteConsoleMsg(UserIndex, "La descripci�n tiene caracteres inv�lidos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Desc = Trim$(description)
                Call WriteConsoleMsg(UserIndex, "La descripci�n ha cambiado.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim vote As String

        Dim errorStr As String

        vote = Buffer.ReadASCIIString()

        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ShowGuildNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowGuildNews(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMA
'Last Modification: 05/17/06
'
'***************************************************

    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte

        Call modGuilds.SendGuildNews(UserIndex)

    End With

End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Now only admins can see other admins' punishment list
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Name As String

        Dim Count As Integer

        Name = Buffer.ReadASCIIString()

        If LenB(Name) <> 0 Then
            If (InStrB(Name, "\") <> 0) Then
                Name = Replace(Name, "\", vbNullString)

            End If

            If (InStrB(Name, "/") <> 0) Then
                Name = Replace(Name, "/", vbNullString)

            End If

            If (InStrB(Name, ":") <> 0) Then
                Name = Replace(Name, ":", vbNullString)

            End If

            If (InStrB(Name, "|") <> 0) Then
                Name = Replace(Name, "|", vbNullString)

            End If

            If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else

                If FileExist(CharPath & Name & ".chr", vbNormal) Then
                    Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))

                    If Count = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else

                        While Count > 0

                            Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                            Count = Count - 1
                        Wend

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ChangePassword" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Creation Date: 10/10/07
'Last Modified By: Rapsodius
'***************************************************
    #If SeguridadAlkon Then

        If UserList(UserIndex).incomingData.length < 65 Then
            Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If

    #Else

        If UserList(UserIndex).incomingData.length < 5 Then
            Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If

    #End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        Dim oldPass As String

        Dim newPass As String

        Dim oldPass2 As String

        'Remove packet ID
        Call Buffer.ReadByte

        #If SeguridadAlkon Then
            oldPass = UCase$(Buffer.ReadASCIIStringFixed(32))
            newPass = UCase$(Buffer.ReadASCIIStringFixed(32))
        #Else
            oldPass = UCase$(Buffer.ReadASCIIString())
            newPass = UCase$(Buffer.ReadASCIIString())
        #End If

        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes especificar una contrase�a nueva, int�ntalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Else
            oldPass2 = UCase$(GetVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Password"))

            If oldPass2 <> oldPass Then
                Call WriteConsoleMsg(UserIndex, "La contrase�a actual proporcionada no es correcta. La contrase�a no ha sido cambiada, int�ntalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Password", newPass)
                Call WriteConsoleMsg(UserIndex, "La contrase�a fue cambiada con �xito.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'10/07/2010: ZaMa - Now normal npcs don't answer if asked to gamble.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Amount As Integer

        Dim TypeNpc As eNPCType

        Amount = .incomingData.ReadInteger()

        ' Dead?
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)

            'Validate target NPC
        ElseIf .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)

            ' Validate Distance
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

            ' Validate NpcType
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then

            Dim TargetNpcType As eNPCType

            TargetNpcType = Npclist(.flags.TargetNPC).NPCtype

            ' Normal npcs don't speak
            If TargetNpcType <> eNPCType.Comun And TargetNpcType <> eNPCType.DRAGON And TargetNpcType <> eNPCType.Pretoriano Then
                Call WriteChatOverHead(UserIndex, "No tengo ning�n inter�s en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

            ' Validate amount
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El m�nimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            ' Validate amount
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El m�ximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            ' Validate user gold
        ElseIf .Stats.GLD < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

        Else

            If RandomNumber(1, 100) <= 47 Then
                .Stats.GLD = .Stats.GLD + Amount
                Call WriteChatOverHead(UserIndex, "�Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

            End If

            Apuestas.Jugadas = Apuestas.Jugadas + 1

            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))

            Call WriteUpdateGold(UserIndex)

        End If

    End With

End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim opt As Byte

        opt = .incomingData.ReadByte()

        Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)

    End With

End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Amount As Long

        Amount = .incomingData.ReadLong()

        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub

        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Amount > 0 And Amount <= .Stats.Banco Then
            .Stats.Banco = .Stats.Banco - Amount
            .Stats.GLD = .Stats.GLD + Amount
            Call WriteChatOverHead(UserIndex, "Ten�s " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

        End If

        Call WriteUpdateGold(UserIndex)
        Call WriteUpdateBankGold(UserIndex)

    End With

End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 09/28/2010
' 09/28/2010 C4b3z0n - Ahora la respuesta de los NPCs sino perteneces a ninguna facci�n solo la hacen el Rey o el Demonio
' 05/17/06 - Maraxus
'***************************************************

    Dim TalkToKing As Boolean

    Dim TalkToDemon As Boolean

    Dim NpcIndex As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Chequea si habla con el rey o el demonio. Puede salir sin hacerlo, pero si lo hace le reponden los npcs
        NpcIndex = .flags.TargetNPC

        If NpcIndex <> 0 Then

            ' Es rey o domonio?
            If Npclist(NpcIndex).NPCtype = eNPCType.Noble Then

                'Rey?
                If Npclist(NpcIndex).flags.Faccion = 0 Then
                    TalkToKing = True
                    ' Demonio
                Else
                    TalkToDemon = True

                End If

            End If

        End If

        'Quit the Royal Army?
        If .Faccion.ArmadaReal = 1 Then

            ' Si le pidio al demonio salir de la armada, este le responde.
            If TalkToDemon Then
                Call WriteChatOverHead(UserIndex, "���Sal de aqu� buf�n!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)

            Else

                ' Si le pidio al rey salir de la armada, le responde.
                If TalkToKing Then
                    Call WriteChatOverHead(UserIndex, "Ser�s bienvenido a las fuerzas imperiales si deseas regresar.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)

                End If

                Call ExpulsarFaccionReal(UserIndex, False)

            End If

            'Quit the Chaos Legion?
        ElseIf .Faccion.FuerzasCaos = 1 Then

            ' Si le pidio al rey salir del caos, le responde.
            If TalkToKing Then
                Call WriteChatOverHead(UserIndex, "���Sal de aqu� maldito criminal!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else

                ' Si le pidio al demonio salir del caos, este le responde.
                If TalkToDemon Then
                    Call WriteChatOverHead(UserIndex, "Ya volver�s arrastrandote.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)

                End If

                Call ExpulsarFaccionCaos(UserIndex, False)

            End If

            ' No es faccionario
        Else

            ' Si le hablaba al rey o demonio, le repsonden ellos
            'Corregido, solo si son en efecto el rey o el demonio, no cualquier NPC (C4b3z0n)
            If (TalkToDemon And criminal(UserIndex)) Or (TalkToKing And Not criminal(UserIndex)) Then    'Si se pueden unir a la facci�n (status), son invitados
                Call WriteChatOverHead(UserIndex, "No perteneces a nuestra facci�n. Si deseas unirte, di /ENLISTAR", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            ElseIf (TalkToDemon And Not criminal(UserIndex)) Then
                Call WriteChatOverHead(UserIndex, "���Sal de aqu� buf�n!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            ElseIf (TalkToKing And criminal(UserIndex)) Then
                Call WriteChatOverHead(UserIndex, "���Sal de aqu� maldito criminal!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                Call WriteConsoleMsg(UserIndex, "�No perteneces a ninguna facci�n!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

    End With

End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Amount As Long

        Amount = .incomingData.ReadLong()

        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub

        If Amount > 0 And Amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteChatOverHead(UserIndex, "Ten�s " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateBankGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No ten�s esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

        End If

    End With

End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 14/11/2010
'14/11/2010: ZaMa - Now denounces can be desactivated.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Text As String

        Dim msg As String

        Text = Buffer.ReadASCIIString()

        If .flags.Silenciado = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)

            msg = LCase$(.Name) & " DENUNCIA: " & Text

            Call SendData(SendTarget.ToAdmins, 0, _
                          PrepareMessageConsoleMsg(msg, FontTypeNames.FONTTYPE_GUILDMSG), True)

            Call Denuncias.Push(msg, False)

            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 1 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        Call .incomingData.ReadByte

        If HasFound(.Name) Then
            Call WriteConsoleMsg(UserIndex, "�Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub

        End If

        Call WriteShowGuildAlign(UserIndex)

    End With

End Sub

''
' Handles the "GuildFundation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundation(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim clanType As eClanType

        Dim Error As String

        clanType = .incomingData.ReadByte()

        If HasFound(.Name) Then
            Call WriteConsoleMsg(UserIndex, "�Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogCheating("El usuario " & .Name & " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .Ip)
            Exit Sub

        End If

        Select Case UCase$(Trim(clanType))

        Case eClanType.ct_RoyalArmy
            .FundandoGuildAlineacion = ALINEACION_ARMADA

        Case eClanType.ct_Evil
            .FundandoGuildAlineacion = ALINEACION_LEGION

        Case eClanType.ct_Neutral
            .FundandoGuildAlineacion = ALINEACION_NEUTRO

        Case eClanType.ct_GM
            .FundandoGuildAlineacion = ALINEACION_MASTER

        Case eClanType.ct_Legal
            .FundandoGuildAlineacion = ALINEACION_CIUDA

        Case eClanType.ct_Criminal
            .FundandoGuildAlineacion = ALINEACION_CRIMINAL

        Case Else
            Call WriteConsoleMsg(UserIndex, "Alineaci�n inv�lida.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub

        End Select

        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, Error) Then
            Call WriteShowGuildFundationForm(UserIndex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub

''
' Handles the "PartyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyKick(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (Marco)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tUser)
            Else

                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")

                End If

                Call WriteConsoleMsg(UserIndex, LCase$(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "PartySetLeader" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartySetLeader(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (MarKoxX)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    'On Error GoTo ErrHandler
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim Rank As Integer

        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

        UserName = Buffer.ReadASCIIString()

        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)

            If tUser > 0 Then

                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And Rank) <= (.flags.Privilegios And Rank) Then
                    Call mdParty.TransformarEnLider(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, LCase$(UserList(tUser).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")

                End If

                Call WriteConsoleMsg(UserIndex, LCase$(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "PartyAcceptMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyAcceptMember(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (Marco)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim Rank As Integer

        Dim bUserVivo As Boolean

        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

        UserName = Buffer.ReadASCIIString()

        If UserList(UserIndex).flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_PARTY)
        Else
            bUserVivo = True
        End If

        If mdParty.UserPuedeEjecutarComandos(UserIndex) And bUserVivo Then
            tUser = NameIndex(UserName)

            If tUser > 0 Then

                If UserList(tUser).flags.Vip > 0 Then
                    WriteConsoleMsg UserIndex, UserList(tUser).Name & " es un usuario VIP, no puedes entrar en party con �l.", FontTypeNames.FONTTYPE_AMARILLO
                Else
                    'Validate administrative ranks - don't allow users to spoof online GMs
                    If (UserList(tUser).flags.Privilegios And Rank) <= (.flags.Privilegios And Rank) Then
                        Call mdParty.AprobarIngresoAParty(UserIndex, tUser)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarqu�a.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If

            Else

                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")

                End If

                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And Rank) <= (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(UserIndex, LCase$(UserName) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarqu�a.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        Dim memberCount As Integer

        Dim i  As Long

        Dim UserName As String

        guild = Buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", vbNullString)

            End If

            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", vbNullString)

            End If

            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))

                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)

                    Call WriteConsoleMsg(UserIndex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & message)

            If LenB(message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(message)

                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName    'Show / Hide the name

            Call RefreshCharStatus(UserIndex)

        End If

    End With

End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte

        'If .flags.Privilegios And PlayerType.User Then Exit Sub

        If esGM(UserIndex) Or (.flags.Privilegios And PlayerType.RoyalCouncil) Then

            Dim i As Long
            Dim List As String
            Dim priv As PlayerType

            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios

            ' Solo dioses pueden ver otros dioses online

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = priv Or PlayerType.Dios Or PlayerType.Admin

            End If

            For i = 1 To LastUser

                If UserList(i).ConnID <> -1 Then
                    If UserList(i).Faccion.ArmadaReal = 1 Then
                        If UserList(i).flags.Privilegios And priv Then
                            List = List & UserList(i).Name & ", "

                        End If

                    End If

                End If

            Next i

            If Len(List) > 0 Then
                Call WriteConsoleMsg(UserIndex, "Reales conectados: " & Left$(List, Len(List) - 2), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte

        'If .flags.Privilegios And PlayerType.User Then Exit Sub

        If esGM(UserIndex) Or (.flags.Privilegios And PlayerType.ChaosCouncil) Then

            Dim i As Long
            Dim List As String
            Dim priv As PlayerType

            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios

            ' Solo dioses pueden ver otros dioses online

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = priv Or PlayerType.Dios Or PlayerType.Admin

            End If

            For i = 1 To LastUser

                If UserList(i).ConnID <> -1 Then
                    If UserList(i).Faccion.FuerzasCaos = 1 Then
                        If UserList(i).flags.Privilegios And priv Then
                            List = List & UserList(i).Name & ", "

                        End If

                    End If

                End If

            Next i

            If Len(List) > 0 Then
                Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(List, Len(List) - 2), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "No hay caos conectados.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        UserName = Buffer.ReadASCIIString()

        Dim tIndex As Integer

        Dim X  As Long

        Dim Y  As Long

        Dim i  As Long

        Dim Found As Boolean

        tIndex = NameIndex(UserName)

        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambi�n lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then    'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else

                    For i = 2 To 5    'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i

                                If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                        Found = True
                                        Exit For

                                    End If

                                End If

                            Next Y

                            If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X

                        If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i

                    'No space found??
                    If Not Found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares est�n ocupados.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim comment As String

        comment = Buffer.ReadASCIIString()

        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'OtorgarFavordelosDioses (0)
        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Call LogGM(.Name, "Hora.")

    End With

    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))

End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 18/11/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim miPos As String

        UserName = Buffer.ReadASCIIString()

        If Not .flags.Privilegios And PlayerType.User Then

            tUser = NameIndex(UserName)

            If tUser <= 0 Then

                If FileExist(CharPath & UserName & ".chr", vbNormal) Then

                    Dim CharPrivs As PlayerType

                    CharPrivs = GetCharPrivs(UserName)

                    If (CharPrivs And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((CharPrivs And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                        miPos = GetVar(CharPath & UserName & ".chr", "INIT", "POSITION")
                        Call WriteConsoleMsg(UserIndex, "Ubicaci�n  " & UserName & " (Offline): " & ReadField(1, miPos, 45) & ", " & ReadField(2, miPos, 45) & ", " & ReadField(3, miPos, 45) & ".", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            Else

                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicaci�n  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        Call LogGM(.Name, "/Donde " & UserName)

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualizaci�n.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Map As Integer

        Dim i, j As Long

        Dim NPCcount1, NPCcount2 As Integer

        Dim NPCcant1() As Integer

        Dim NPCcant2() As Integer

        Dim List1() As String

        Dim List2() As String

        Map = .incomingData.ReadInteger()

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        If MapaValido(Map) Then

            For i = 1 To LastNPC

                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.Map = Map Then

                    '�esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else

                            For j = 0 To NPCcount1 - 1

                                If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1

                            End If

                        End If

                    Else

                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else

                            For j = 0 To NPCcount2 - 1

                                If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1

                            End If

                        End If

                    End If

                End If

            Next i

            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)

            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)

            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay m�s NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call LogGM(.Name, "Numero enemigos en mapa " & Map)

        End If

    End With

End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/09
'26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim X  As Integer

        Dim Y  As Integer

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        X = .flags.TargetX
        Y = .flags.TargetY

        Call FindLegalPos(UserIndex, .flags.TargetMap, X, Y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)

    End With

End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim Map As Integer

        Dim X  As Integer

        Dim Y  As Integer

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()
        Map = Buffer.ReadInteger()
        X = Buffer.ReadByte()
        Y = Buffer.ReadByte()

        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)

                    End If

                Else
                    tUser = UserIndex

                End If

                If tUser <= 0 Then
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        '�Es una posici�n valida para un usuario cualquiera?
                        If InMapBounds(Map, X, Y) = False Or LegalPos(Map, X, Y) = False Then
                            'Call WriteMensajes(UserIndex, eMensajes.Mensaje463)    ' "Posici�n inv�lida."
                        End If

                        '�Existe el personaje?
                        If Not FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
                            'Call WriteMensajes(UserIndex, eMensajes.Mensaje183)        '"Usuario inexistente."
                        Else
                            Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "Position", Map & "-" & X & "-" & Y)
                        End If
                        'Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)

                    End If

                ElseIf Not ((UserList(tUser).flags.Privilegios And PlayerType.Dios) <> 0 Or _
                            (UserList(tUser).flags.Privilegios And PlayerType.Admin) <> 0) Or _
                            tUser = UserIndex Then

                    If InMapBounds(Map, X, Y) Then
                        Call FindLegalPos(tUser, Map, X, Y)
                        Call WarpUserChar(tUser, Map, X, Y, True, True)
                        Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, "Transport� a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias ser�n ignoradas por el servidor de aqu� en m�s. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)

                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)

    End With

End Sub

''
' Handles the "RequestPartyForm" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .PartyIndex > 0 Then
            Call WriteShowPartyForm(UserIndex)

        Else
            Call WriteConsoleMsg(UserIndex, "No perteneces a ning�n grupo!", FontTypeNames.FONTTYPE_INFOBOLD)

        End If

    End With

End Sub

''
' Handles the "ItemUpgrade" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemUpgrade(ByVal UserIndex As Integer)

'***************************************************
'Author: Torres Patricio
'Last Modification: 12/09/09
'
'***************************************************
    With UserList(UserIndex)

        Dim ItemIndex As Integer

        'Remove packet ID
        Call .incomingData.ReadByte

        ItemIndex = .incomingData.ReadInteger()

        If ItemIndex <= 0 Then Exit Sub
        If Not TieneObjetos(ItemIndex, 1, UserIndex) Then Exit Sub

        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        Call DoUpgrade(UserIndex, ItemIndex)

    End With

End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        UserName = Buffer.ReadASCIIString()

        If Not .flags.Privilegios And PlayerType.User Then _
           Call Ayuda.Quitar(UserName)

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim X  As Integer

        Dim Y  As Integer

        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambi�n lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)

                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)

                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Call FlushBuffer(tUser)

                    End If

                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Call DoAdminInvisible(UserIndex)
        Call LogGM(.Name, "/INVISIBLE")

    End With

End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Call WriteShowGMPanelForm(UserIndex)

    End With

End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
    Dim i      As Long

    Dim names() As String

    Dim Count  As Long

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

        ReDim names(1 To LastUser) As String
        Count = 1

        For i = 1 To LastUser

            If (LenB(UserList(i).Name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).Name
                    Count = Count + 1

                End If

            End If

        Next i

        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)

    End With

End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 07/10/2010
'07/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'***************************************************
    Dim i      As Long

    Dim users  As String

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

        For i = 1 To LastUser

            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                users = users & ", " & UserList(i).Name

                ' Display the user being checked by the centinel
                If UserList(i).flags.CentinelaIndex <> 0 Then _
                   users = users & " (*)"

            End If

        Next i

        If LenB(users) <> 0 Then
            users = Right$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i      As Long

    Dim users  As String

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub

        For i = 1 To LastUser

            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                users = users & UserList(i).Name & ", "

            End If

        Next i

        If LenB(users) <> 0 Then
            users = Left$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim Reason As String

        Dim jailTime As Byte

        Dim Count As Byte

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()
        jailTime = Buffer.ReadByte()

        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")

        End If

        '/carcel nick@motivo@<tiempo>
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then    ' @@ Miqueas : 03/12/15 - Ahora lo usan los consejeros tambien
            'If we got here then packet is complete, copy data back to original queue
            Call .incomingData.CopyBuffer(Buffer)
            Exit Sub

        End If

        If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
        Else
            tUser = NameIndex(UserName)

            '�Existe el personaje?
            If Not FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
                Call WriteConsoleMsg(UserIndex, "SERVIDOR> El personaje no existe.", FontTypeNames.FONTTYPE_INFO)

            ElseIf tUser <= 0 Then

                If (EsDios(UserName) Or EsAdmin(UserName)) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)

                Else
                    Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date$ & " " & time$)
                    Call WriteVar(CharPath & UserName & ".chr", "COUNTERS", "Pena", jailTime)

                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Position", Configuracion.Prision.Map & "-" & Configuracion.Prision.X & "-" & Configuracion.Prision.Y)
                    Call WriteConsoleMsg(UserIndex, "SERVIDOR> Personaje OFFLINE, se encarcelar� a " & UserName & " por " & jailTime & " minutos de todas formas.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                ElseIf jailTime > 60 Then
                    Call WriteConsoleMsg(UserIndex, "No pued�s encarcelar por m�s de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", vbNullString)

                    End If

                    If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", vbNullString)

                    End If

                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & time)

                    End If

                    Call Encarcelar(tUser, jailTime, .Name)
                    Call LogGM(.Name, " encarcel� a " & UserName)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/22/08 (NicoNZ)
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub

        Dim tNpc As Integer

        Dim auxNPC As npc

        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.Consejero Then
            If .Pos.Map = Configuracion.MAPA_PRETORIANO Then
                Call WriteConsoleMsg(UserIndex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

        tNpc = .flags.TargetNPC

        If tNpc > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNpc).Name, FontTypeNames.FONTTYPE_INFO)

            auxNPC = Npclist(tNpc)
            Call QuitarNPC(tNpc)
            Call ReSpawnNpc(auxNPC)

            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim Reason As String

        Dim Privs As PlayerType

        Dim Count As Byte

        UserName = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                Privs = UserDarPrivilegioLevel(UserName)

                If Not Privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", vbNullString)

                    End If

                    If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", vbNullString)

                    End If

                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & time)

                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)

                    End If

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 18/09/2010
'02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
'11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
'18/09/2010: ZaMa - Ahora se puede editar la vida del propio pj (cualquier rm o dios).
'***************************************************
    If UserList(UserIndex).incomingData.length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim Opcion As Byte

        Dim Arg1 As String

        Dim Arg2 As String

        Dim valido As Boolean

        Dim LoopC As Byte

        Dim CommandString As String

        Dim N  As Byte

        Dim UserCharPath As String

        Dim Var As Long

        UserName = Replace(Buffer.ReadASCIIString(), "+", " ")

        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)

        End If

        Opcion = Buffer.ReadByte()
        Arg1 = Buffer.ReadASCIIString()
        Arg2 = Buffer.ReadASCIIString()

        If .flags.Privilegios And PlayerType.RoleMaster Then

            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)

            Case PlayerType.Consejero
                ' Los RMs consejeros s�lo se pueden editar su head, body, level y vida
                valido = tUser = UserIndex And _
                         (Opcion = eEditOptions.eo_Body Or _
                          Opcion = eEditOptions.eo_Head Or _
                          Opcion = eEditOptions.eo_Level Or _
                          Opcion = eEditOptions.eo_Vida)

            Case PlayerType.SemiDios
                ' Los RMs s�lo se pueden editar su level o vida y el head y body de cualquiera
                valido = ((Opcion = eEditOptions.eo_Level Or Opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or _
                         Opcion = eEditOptions.eo_Body Or _
                         Opcion = eEditOptions.eo_Head

            Case PlayerType.Dios
                ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                ' pero si quiere modificar el level o vida s�lo lo puede hacer sobre s� mismo
                valido = ((Opcion = eEditOptions.eo_Level Or Opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or _
                         Opcion = eEditOptions.eo_Body Or _
                         Opcion = eEditOptions.eo_Head Or _
                         Opcion = eEditOptions.eo_CiticensKilled Or _
                         Opcion = eEditOptions.eo_CriminalsKilled Or _
                         Opcion = eEditOptions.eo_Class Or _
                         Opcion = eEditOptions.eo_Skills Or _
                         Opcion = eEditOptions.eo_addGold

            End Select

            'Si no es RM debe ser dios para poder usar este comando
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then

            If Opcion = eEditOptions.eo_Vida Then
                '  Por ahora dejo para que los dioses no puedan editar la vida de otros
                valido = (tUser = UserIndex)
            Else
                valido = True

            End If

        ElseIf .flags.PrivEspecial Then
            valido = (Opcion = eEditOptions.eo_CiticensKilled) Or _
                     (Opcion = eEditOptions.eo_CriminalsKilled)

        End If

        If valido Then
            UserCharPath = CharPath & UserName & ".chr"

            If tUser <= 0 And Not FileExist(UserCharPath) Then
                Call WriteConsoleMsg(UserIndex, "Est�s intentando editar un usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, "Intent� editar un usuario inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "

                Select Case Opcion

                Case eEditOptions.eo_Gold

                    If val(Arg1) <= MAX_ORO_EDIT Then
                        If tUser <= 0 Then    ' Esta offline?
                            Call WriteVar(UserCharPath, "STATS", "GLD", val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else    ' Online
                            UserList(tUser).Stats.GLD = val(Arg1)
                            Call WriteUpdateGold(tUser)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "No est� permitido utilizar valores mayores a " & MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)

                    End If

                    ' Log it
                    CommandString = CommandString & "ORO "

                Case eEditOptions.eo_Experience

                    If val(Arg1) > 20000000 Then
                        Arg1 = 20000000

                    End If

                    If tUser <= 0 Then    ' Offline
                        Var = GetVar(UserCharPath, "STATS", "EXP")
                        Call WriteVar(UserCharPath, "STATS", "EXP", Var + val(Arg1))
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else    ' Online
                        UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                        Call CheckUserLevel(tUser)
                        Call WriteUpdateExp(tUser)

                    End If

                    ' Log it
                    CommandString = CommandString & "EXP "

                Case eEditOptions.eo_Body

                    If tUser <= 0 Then

                        Call WriteVar(UserCharPath, "INIT", "Body", Arg1)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                    End If

                    ' Log it
                    CommandString = CommandString & "BODY "

                Case eEditOptions.eo_Head

                    If tUser <= 0 Then
                        Call WriteVar(UserCharPath, "INIT", "Head", Arg1)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call ChangeUserChar(tUser, UserList(tUser).Char.Body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                    End If

                    ' Log it
                    CommandString = CommandString & "HEAD "

                Case eEditOptions.eo_CriminalsKilled
                    Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))

                    If tUser <= 0 Then    ' Offline
                        Call WriteVar(UserCharPath, "FACCIONES", "CrimMatados", Var)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else    ' Online
                        UserList(tUser).Faccion.CriminalesMatados = Var

                    End If

                    ' Log it
                    CommandString = CommandString & "CRI "

                Case eEditOptions.eo_CiticensKilled
                    Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))

                    If tUser <= 0 Then    ' Offline
                        Call WriteVar(UserCharPath, "FACCIONES", "CiudMatados", Var)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else    ' Online
                        UserList(tUser).Faccion.CiudadanosMatados = Var

                    End If

                    ' Log it
                    CommandString = CommandString & "CIU "

                Case eEditOptions.eo_Level

                    If val(Arg1) > Configuracion.NivelMaximo Then
                        Arg1 = CStr(Configuracion.NivelMaximo)
                        Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & Configuracion.NivelMaximo & ".", FONTTYPE_INFO)

                    End If

                    ' Chequeamos si puede permanecer en el clan
                    If val(Arg1) >= 25 Then

                        Dim GI As Integer

                        If tUser <= 0 Then
                            GI = GetVar(UserCharPath, "GUILD", "GUILDINDEX")
                        Else
                            GI = UserList(tUser).GuildIndex

                        End If

                        If GI > 0 Then
                            If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                                'We get here, so guild has factionary alignment, we have to expulse the user
                                Call modGuilds.m_EcharMiembroDeClan(-1, UserName)

                                Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(UserName & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))

                                ' Si esta online le avisamos
                                If tUser > 0 Then _
                                   Call WriteConsoleMsg(tUser, "�Ya tienes la madurez suficiente como para decidir bajo que estandarte pelear�s! Por esta raz�n, hasta tanto no te enlistes en la facci�n bajo la cual tu clan est� alineado, estar�s exclu�do del mismo.", FontTypeNames.FONTTYPE_GUILD)

                            End If

                        End If

                    End If

                    If tUser <= 0 Then    ' Offline
                        Call WriteVar(UserCharPath, "STATS", "ELV", val(Arg1))
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else    ' Online
                        UserList(tUser).Stats.ELV = val(Arg1)
                        Call WriteUpdateUserStats(tUser)

                    End If

                    ' Log it
                    CommandString = CommandString & "LEVEL "

                Case eEditOptions.eo_Class

                    For LoopC = 1 To NUMCLASES

                        If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                    Next LoopC

                    If LoopC > NUMCLASES Then
                        Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then    ' Offline
                            Call WriteVar(UserCharPath, "INIT", "Clase", LoopC)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else    ' Online
                            UserList(tUser).clase = LoopC

                        End If

                    End If

                    ' Log it
                    CommandString = CommandString & "CLASE "

                Case eEditOptions.eo_Skills

                    For LoopC = 1 To NUMSKILLS

                        If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                    Next LoopC

                    If LoopC > NUMSKILLS Then
                        Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then    ' Offline
                            Call WriteVar(UserCharPath, "Skills", "SK" & LoopC, Arg2)

                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else    ' Online
                            UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)

                        End If

                    End If

                    ' Log it
                    CommandString = CommandString & "SKILLS "

                Case eEditOptions.eo_SkillPointsLeft

                    If tUser <= 0 Then    ' Offline
                        Call WriteVar(UserCharPath, "STATS", "SkillPtsLibres", Arg1)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else    ' Online
                        UserList(tUser).Stats.SkillPts = val(Arg1)

                    End If

                    ' Log it
                    CommandString = CommandString & "SKILLSLIBRES "

                Case eEditOptions.eo_Nobleza
                    Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))

                    If tUser <= 0 Then    ' Offline
                        Call WriteVar(UserCharPath, "REP", "Nobles", Var)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else    ' Online
                        UserList(tUser).Reputacion.NobleRep = Var

                    End If

                    ' Log it
                    CommandString = CommandString & "NOB "

                Case eEditOptions.eo_Asesino
                    Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))

                    If tUser <= 0 Then    ' Offline
                        Call WriteVar(UserCharPath, "REP", "Asesino", Var)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else    ' Online
                        UserList(tUser).Reputacion.AsesinoRep = Var

                    End If

                    ' Log it
                    CommandString = CommandString & "ASE "

                Case eEditOptions.eo_Sex

                    Dim Sex As Byte

                    Sex = IIf(UCase$(Arg1) = "MUJER", eGenero.Mujer, 0)    ' Mujer?
                    Sex = IIf(UCase$(Arg1) = "HOMBRE", eGenero.Hombre, Sex)    ' Hombre?

                    If Sex <> 0 Then    ' Es Hombre o mujer?
                        If tUser <= 0 Then    ' OffLine
                            Call WriteVar(UserCharPath, "INIT", "Genero", Sex)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else    ' Online
                            UserList(tUser).Genero = Sex

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)

                    End If

                    ' Log it
                    CommandString = CommandString & "SEX "

                Case eEditOptions.eo_Raza

                    Dim raza As Byte

                    Arg1 = UCase$(Arg1)

                    Select Case Arg1

                    Case "HUMANO"
                        raza = eRaza.Humano

                    Case "ELFO"
                        raza = eRaza.Elfo

                    Case "DROW"
                        raza = eRaza.Drow

                    Case "ENANO"
                        raza = eRaza.Enano

                    Case "GNOMO"
                        raza = eRaza.Gnomo

                    Case Else
                        raza = 0

                    End Select

                    If raza = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Raza", raza)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).raza = raza

                        End If

                    End If

                    ' Log it
                    CommandString = CommandString & "RAZA "

                Case eEditOptions.eo_addGold

                    Dim bankGold As Long

                    If Abs(Arg1) > MAX_ORO_EDIT Then
                        Call WriteConsoleMsg(UserIndex, "No est� permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then
                            bankGold = GetVar(CharPath & UserName & ".chr", "STATS", "BANCO")
                            Call WriteVar(UserCharPath, "STATS", "BANCO", IIf(bankGold + val(Arg1) <= 0, 0, bankGold + val(Arg1)))
                            Call WriteConsoleMsg(UserIndex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
                        Else
                            UserList(tUser).Stats.Banco = IIf(UserList(tUser).Stats.Banco + val(Arg1) <= 0, 0, UserList(tUser).Stats.Banco + val(Arg1))
                            Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_MESSAGE, FONTTYPE_TALK)

                        End If

                    End If

                    ' Log it
                    CommandString = CommandString & "AGREGAR "

                Case eEditOptions.eo_Vida

                    If val(Arg1) > MAX_VIDA_EDIT Then
                        Arg1 = CStr(MAX_VIDA_EDIT)
                        Call WriteConsoleMsg(UserIndex, "No puedes tener vida superior a " & MAX_VIDA_EDIT & ".", FONTTYPE_INFO)

                    End If

                    ' No valido si esta offline, porque solo se puede editar a si mismo
                    UserList(tUser).Stats.MaxHP = val(Arg1)
                    UserList(tUser).Stats.MinHp = val(Arg1)

                    Call WriteUpdateUserStats(tUser)

                    ' Log it
                    CommandString = CommandString & "VIDA "

                Case eEditOptions.eo_Poss

                    Dim Map As Integer

                    Dim X As Integer

                    Dim Y As Integer

                    Map = val(ReadField(1, Arg1, 45))
                    X = val(ReadField(2, Arg1, 45))
                    Y = val(ReadField(3, Arg1, 45))

                    If InMapBounds(Map, X, Y) Then

                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "POSITION", Map & "-" & X & "-" & Y)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WarpUserChar(tUser, Map, X, Y, True, True)
                            Call WriteConsoleMsg(UserIndex, "Usuario teletransportado: " & UserName, FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Posici�n inv�lida", FONTTYPE_INFO)

                    End If

                    ' Log it
                    CommandString = CommandString & "POSS "

                Case Else
                    Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                    CommandString = CommandString & "UNKOWN "

                End Select

                CommandString = CommandString & Arg1 & " " & Arg2
                Call LogGM(.Name, CommandString & " " & UserName)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)

'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'***************************************************
44    If UserList(UserIndex).incomingData.length < 3 Then
45        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
46        Exit Sub

47    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
555        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

556        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
557        Call Buffer.ReadByte

        Dim TargetName As String

        Dim TargetIndex As Integer

1        TargetName = Replace$(Buffer.ReadASCIIString(), "+", " ")
2        TargetIndex = NameIndex(TargetName)

3        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

            'is the player offline?
4            If TargetIndex <= 0 Then

                'don't allow to retrieve administrator's info
5                If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
6                    Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando en charfile.", FontTypeNames.FONTTYPE_INFO)
7                    Call SendUserStatsTxtOFF(UserIndex, TargetName)

8                End If

9            Else

                'don't allow to retrieve administrator's info
10                If UserList(TargetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
11                    Call SendUserStatsTxt(UserIndex, TargetIndex)

12                End If

13            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
559        Call .incomingData.CopyBuffer(Buffer)

    End With
    
    Exit Sub
    
errhandleR:

    Dim Error  As Long

    Error = Err.Number
    
    LogError "Error en HandleRequestCharInfo: " & Erl & " - " & Err.Number & " " & Err.description

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim UserIsAdmin As Boolean

        Dim OtherUserIsAdmin As Boolean

        UserName = Buffer.ReadASCIIString()

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And ((.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin) Then
            Call LogGM(.Name, "/STAT " & UserName)

            tUser = NameIndex(UserName)

            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)

            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_INFO)

                    Call SendUserMiniStatsTxtFromChar(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserMiniStatsTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim UserIsAdmin As Boolean

        Dim OtherUserIsAdmin As Boolean

        UserName = Buffer.ReadASCIIString()

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (.flags.Privilegios And PlayerType.SemiDios) Or UserIsAdmin Then

            Call LogGM(.Name, "/BAL " & UserName)

            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)

            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)

            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserOROTxtFromChar(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim UserIsAdmin As Boolean

        Dim OtherUserIsAdmin As Boolean

        UserName = Buffer.ReadASCIIString()

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)

            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)

            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)

            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserInvTxtFromChar(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserInvTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim UserIsAdmin As Boolean

        Dim OtherUserIsAdmin As Boolean

        UserName = Buffer.ReadASCIIString()

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            Call LogGM(.Name, "/BOV " & UserName)

            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)

            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)

            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserBovedaTxtFromChar(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la b�veda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserBovedaTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la b�veda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim LoopC As Long

        Dim message As String

        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)

            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", vbNullString)

                End If

                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", vbNullString)

                End If

                For LoopC = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC

                Call WriteConsoleMsg(UserIndex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim LoopC As Byte

        UserName = Buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex

            End If

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                With UserList(tUser)

                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0

                        If .flags.Navegando = 1 Then
                            Call ToggleBoatBody(tUser)
                        Else
                            Call DarCuerpoDesnudo(tUser)

                        End If

                        If .flags.Traveling = 1 Then
                            Call EndTravel(tUser, True)

                        End If

                        Call ChangeUserChar(tUser, .Char.Body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)

                    End If

                    .Stats.MinHp = .Stats.MaxHP

                    If .flags.Traveling = 1 Then
                        Call EndTravel(tUser, True)

                    End If

                End With

                Call WriteUpdateHP(tUser)

                Call FlushBuffer(tUser)

                Call LogGM(.Name, "Resucito a " & UserName)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)

'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
    Dim i      As Long

    Dim List   As String

    Dim priv   As PlayerType

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin

        For i = 1 To LastUser

            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then _
                   List = List & UserList(i).Name & ", "

            End If

        Next i

        If LenB(List) <> 0 Then
            List = Left$(List, Len(List) - 2)
            Call WriteConsoleMsg(UserIndex, List & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 23/03/2009
'23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Map As Integer

        Map = .incomingData.ReadInteger

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Dim LoopC As Long

        Dim List As String

        Dim priv As PlayerType

        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)

        For LoopC = 1 To LastUser

            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                If UserList(LoopC).flags.Privilegios And priv Then _
                   List = List & UserList(LoopC).Name & ", "

            End If

        Next LoopC

        If Len(List) > 2 Then List = Left$(List, Len(List) - 2)

        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & List, FontTypeNames.FONTTYPE_INFO)
        Call LogGM(.Name, "/ONLINEMAP " & Map)

    End With

End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")

                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(UserIndex, "S�lo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim Rank As Integer

        Dim IsAdmin As Boolean

        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

        UserName = Buffer.ReadASCIIString()
        IsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (.flags.Privilegios And PlayerType.SemiDios) Or IsAdmin Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no est� online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarqu�a mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarqu�a mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ech� a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Ech� a " & UserName)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User And .Name <> "Cuicui" Then
                    Call WriteConsoleMsg(UserIndex, "��Est�s loco?? ��C�mo vas a pi�atear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)

                End If

            Else

                If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                    Call WriteConsoleMsg(UserIndex, "No est� online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "��Est�s loco?? ��C�mo vas a pi�atear un gm?? :@", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim Reason As String

        UserName = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, UserName, Reason)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim cantPenas As Byte

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", vbNullString)

            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", vbNullString)

            End If

            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else

                If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)

                    'penas
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & time)

                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no est� baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0

        End If

    End With

End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim X  As Integer

        Dim Y  As Integer

        UserName = Buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                If EsDios(UserName) Or EsAdmin(UserName) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "El jugador no est� online.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                   (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)
                    Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Call EnviarSpawnList(UserIndex)

    End With

End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim npc As Integer

        npc = .incomingData.ReadInteger()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then _
               Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)

            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(npc).NPCNAME)

        End If

    End With

End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub

        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)

    End With

End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Call General.LimpiarMundo

    End With

End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice. soy puto
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(.Name & ">", message, FontTypeNames.FONTTYPE_DIOS, FontTypeNames.FONTTYPE_APU�ALADO))

            End If
        ElseIf (.flags.Privilegios And (PlayerType.Consejero)) Then

            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(.Name & ">", message, FontTypeNames.FONTTYPE_DIOS, FontTypeNames.FONTTYPE_VERDE))
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "MapMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMapMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then

                Dim Mapa As Integer

                Mapa = .Pos.Map

                Call LogGM(.Name, "Mensaje a mapa " & Mapa & ":" & message)
                Call SendData(SendTarget.toMap, Mapa, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_TALK))

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'Pablo (ToxicWaste): Agrego para que el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim priv As PlayerType

        Dim IsAdmin As Boolean

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

            IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0

            If IsAdmin Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User

            End If

            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).Ip, FontTypeNames.FONTTYPE_INFO)

                    Dim Ip As String

                    Dim lista As String

                    Dim LoopC As Long

                    Ip = UserList(tUser).Ip

                    For LoopC = 1 To LastUser

                        If UserList(LoopC).Ip = Ip Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "

                                End If

                            End If

                        End If

                    Next LoopC

                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & Ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "No hay ning�n personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Ip As String

        Dim LoopC As Long

        Dim lista As String

        Dim priv As PlayerType

        Ip = .incomingData.ReadByte() & "."
        Ip = Ip & .incomingData.ReadByte() & "."
        Ip = Ip & .incomingData.ReadByte() & "."
        Ip = Ip & .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & Ip)

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User

        End If

        For LoopC = 1 To LastUser

            If UserList(LoopC).Ip = Ip Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "

                    End If

                End If

            End If

        Next LoopC

        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & Ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim GuildName As String

        Dim tGuild As Integer

        GuildName = Buffer.ReadASCIIString()

        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")

        End If

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)

            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase$(GuildName) & ": " & _
                                                modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 22/03/2010
'15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
'22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Mapa As Integer

        Dim X  As Byte

        Dim Y  As Byte

        Dim Radio As Byte

        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        Radio = .incomingData.ReadByte()

        Radio = MinimoInt(Radio, 6)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/CT " & Mapa & "," & X & "," & Y & "," & Radio)

        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then _
           Exit Sub

        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.objIndex > 0 Then _
           Exit Sub

        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then _
           Exit Sub

        If MapData(Mapa, X, Y).ObjInfo.objIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Dim ET As Obj

        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.objIndex = TELEP_OBJ_INDEX + Radio

        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = Mapa
            .TileExit.X = X
            .TileExit.Y = Y

        End With

        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)

    End With

End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)

        Dim Mapa As Integer

        Dim X  As Byte

        Dim Y  As Byte

        'Remove packet ID
        Call .incomingData.ReadByte

        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY

        If Not InMapBounds(Mapa, X, Y) Then Exit Sub

        With MapData(Mapa, X, Y)

            If .ObjInfo.objIndex = 0 Then Exit Sub

            If ObjData(.ObjInfo.objIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                Call LogGM(UserList(UserIndex).Name, "/DT: " & Mapa & "," & X & "," & Y)

                Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)

                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.objIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)

                End If

                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0

            End If

        End With

    End With

End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Call LogGM(.Name, "/LLUVIA")
        'Lloviendo = Not Lloviendo

         '                       Call SendData(SendTarget.toMap, 127, PrepareMessageRainToggle())

          '                      Call SendData(SendTarget.toMap, 128, PrepareMessageRainToggle())

    End With

End Sub

''
' Handles the "EnableDenounces" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnableDenounces(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Enables/Disables
'***************************************************

    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte

        If Not esGM(UserIndex) Then Exit Sub

        Dim Activado As Boolean

        Dim msg As String

        Activado = Not .flags.SendDenounces
        .flags.SendDenounces = Activado

        msg = "Denuncias por consola " & IIf(Activado, "ativadas", "desactivadas") & "."

        Call LogGM(.Name, msg)

        Call WriteConsoleMsg(UserIndex, msg, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "ShowDenouncesList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowDenouncesList(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowDenounces(UserIndex)

    End With

End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim tUser As Integer

        Dim Desc As String

        Desc = Buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser

            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte

        Dim Mapa As Integer

        midiID = .incomingData.ReadByte
        Mapa = .incomingData.ReadInteger

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map

            End If

            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.Map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))

            End If

        End If

    End With

End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte

        Dim Mapa As Integer

        Dim X  As Byte

        Dim Y  As Byte

        waveID = .incomingData.ReadByte()
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, X, Y) Then
                Mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y

            End If

            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))

        End If

    End With

End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster Or PlayerType.RoyalCouncil) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("EJ�RCITO REAL> " & message, FontTypeNames.FONTTYPE_TALK))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster Or PlayerType.ChaosCouncil) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Dim X  As Long

        Dim Y  As Long

        Dim bIsExit As Boolean

        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.objIndex > 0 Then
                        bIsExit = MapData(.Pos.Map, X, Y).TileExit.Map > 0

                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.objIndex, bIsExit) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)

                        End If

                    End If

                End If

            Next X
        Next Y

        Call LogGM(UserList(UserIndex).Name, "/MASSDEST")

    End With

End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim LoopC As Byte

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim LoopC As Byte

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Dim tObj As Integer

        Dim lista As String

        Dim X  As Long

        Dim Y  As Long

        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.objIndex

                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            Next Y
        Next X

    End With

End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call SecurityIp.DumpTables

    End With

End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil

                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))

                    End If

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil

                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))

                    End If

                End With

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tTrigger As Byte

        Dim tLog As String

        Dim objIndex As Integer

        tTrigger = .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        If tTrigger >= 0 Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura Then
                If tTrigger <> eTrigger.zonaOscura Then
                    If Not (.flags.AdminInvisible = 1) Then _
                       Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

                    objIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.objIndex

                    If objIndex > 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageObjectCreate(ObjData(objIndex).GrhIndex, .Pos.X, .Pos.Y))

                    End If

                End If

            Else

                If tTrigger = eTrigger.zonaOscura Then
                    If Not (.flags.AdminInvisible = 1) Then _
                       Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))

                    objIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.objIndex

                    If objIndex > 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageObjectDelete(.Pos.X, .Pos.Y))

                    End If

                End If

            End If

            MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y

            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
    Dim tTrigger As Byte

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger

        Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)

        Call WriteConsoleMsg(UserIndex, _
                             "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y _
                             , FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Dim lista As String

        Dim LoopC As Long

        Call LogGM(.Name, "/BANIPLIST")

        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC

        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)

        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call BanIpGuardar
        Call BanIpCargar

    End With

End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim GuildName As String

        Dim CantMembers As Integer

        Dim LoopC As Long

        Dim member As String

        Dim Count As Byte

        Dim tIndex As Integer

        Dim tFile As String

        GuildName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"

            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " bane� al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))

                'baneamos a los miembros
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))

                CantMembers = val(GetVar(tFile, "INIT", "NroMembers"))

                For LoopC = 1 To CantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")

                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))

                    tIndex = NameIndex(member)

                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)

                    End If

                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & member & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    Count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", Count + 1)
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": BAN AL CLAN: " & GuildName & " " & Date & " " & time)
                Next LoopC

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 07/02/09
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'07/02/09 Pato - Ahora no es posible saber si un gm est� o no online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim bannedIP As String

        Dim tUser As Integer

        Dim Reason As String

        Dim i  As Long

        ' Is it by ip??
        If Buffer.ReadBoolean() Then
            bannedIP = Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte()
        Else
            tUser = NameIndex(Buffer.ReadASCIIString())

            If tUser > 0 Then bannedIP = UserList(tUser).Ip

        End If

        Reason = Buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)

                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " bane� la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))

                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser

                        If UserList(i).ConnIDValida Then
                            If UserList(i).Ip = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & Reason)

                            End If

                        End If

                    Next i

                End If

            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no est� online.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim bannedIP As String

        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tObj As Integer

        Dim tStr As String

        Dim tAmount As Integer

        tObj = .incomingData.ReadInteger()

        tAmount = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        If Not EsAdmin(.Name) Then    ' @@ Solo los admins Pueden crear los items con NoCreable = 1 -18/11/2015

            If ItemShop(tObj) = True Then
                Call WriteConsoleMsg(UserIndex, "SERVIDOR> No puedes crear este item.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

        Dim Mapa As Integer

        Dim X  As Byte

        Dim Y  As Byte

        Mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y

        Call LogGM(.Name, "/CI: " & tObj & " Cantidad: " & tAmount & " en mapa " & _
                          Mapa & " (" & X & "," & Y & ")")

        If MapData(Mapa, X, Y - 1).ObjInfo.objIndex > 0 Then _
           Exit Sub

        If MapData(Mapa, X, Y - 1).TileExit.Map > 0 Then _
           Exit Sub

        If tAmount > 10000 Then
            Call WriteConsoleMsg(UserIndex, "SERVIDOR> La cantidad m�xima de items a crear es de '10.000'.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If tObj < 1 Or tObj > NumObjDatas Then _
           Exit Sub

        'Is the object not null?
        If LenB(ObjData(tObj).Name) = 0 Then Exit Sub

        Dim Objeto As Obj

        Call WriteConsoleMsg(UserIndex, "��ATENCI�N: Item creado con exito", FontTypeNames.FONTTYPE_GUILD)

        Objeto.Amount = tAmount
        Objeto.objIndex = tObj
        Call MakeObj(Objeto, Mapa, X, Y - 1)

        If ObjData(tObj).Log = 1 Then
            Call LogItemsEspeciales(.Name & " /CI: [" & tObj & "]" & ObjData(tObj).Name & " Cantidad: " & tAmount & "  en mapa " & _
                                    Mapa & " (" & X & "," & Y & ")")

        End If

    End With

End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Dim Mapa As Integer

        Dim X  As Byte

        Dim Y  As Byte

        Mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y

        Dim objIndex As Integer

        objIndex = MapData(Mapa, X, Y).ObjInfo.objIndex

        If objIndex = 0 Then Exit Sub

        Call LogGM(.Name, "/DEST " & objIndex & " en mapa " & _
                          Mapa & " (" & X & "," & Y & "). Cantidad: " & MapData(Mapa, X, Y).ObjInfo.Amount)

        If ObjData(objIndex).OBJType = eOBJType.otTeleport And _
           MapData(Mapa, X, Y).TileExit.Map > 0 Then

            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports as�. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call EraseObj(10000, Mapa, X, Y)

    End With

End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
           (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or _
           .flags.PrivEspecial Then

            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", vbNullString)

            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", vbNullString)

            End If

            tUser = NameIndex(UserName)

            Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)

            If tUser > 0 Then
                Call ExpulsarFaccionCaos(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else

                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
           (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or _
           .flags.PrivEspecial Then

            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", vbNullString)

            End If

            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", vbNullString)

            End If

            tUser = NameIndex(UserName)

            Call LogGM(.Name, "ECH� DE LA REAL A: " & UserName)

            If tUser > 0 Then
                Call ExpulsarFaccionReal(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else

                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte

        midiID = .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast m�sica: " & midiID, FontTypeNames.FONTTYPE_SERVER))

        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))

    End With

End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte

        waveID = .incomingData.ReadByte()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))

    End With

End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim punishment As Byte

        Dim NewText As String

        UserName = Buffer.ReadASCIIString()
        punishment = Buffer.ReadByte
        NewText = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else

                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", vbNullString)

                End If

                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", vbNullString)

                End If

                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & "-" & _
                                      GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                                      & " de " & UserName & " y la cambi� por: " & NewText)

                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & time)

                    Call WriteConsoleMsg(UserIndex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0

        End If

        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)

    End With

End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        If .flags.TargetNPC = 0 Then Exit Sub

        If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Pretoriano Then Exit Sub


        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)

    End With

End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Dim X  As Long

        Dim Y  As Long

        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then
                        If Npclist(MapData(.Pos.Map, X, Y).NpcIndex).NPCtype <> eNPCType.Pretoriano Then
                            Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                        End If
                    End If
                End If

            Next X
        Next Y

        Call LogGM(.Name, "/MASSKILL")

    End With

End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)

'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim lista As String

        Dim LoopC As Byte

        Dim priv As Integer

        Dim validCheck As Boolean

        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then

            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", vbNullString)

            End If

            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", vbNullString)

            End If

            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")

            End If

            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

            End If

            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & UserName)

                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conect� son:"

                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC

                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarqu�a que vos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim color As Long

        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color

        End If

    End With

End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible

        End If

    End With

End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)

'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 07/06/2010
'Check one Users Slot in Particular from Inventory
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        'Reads the UserName and Slot Packets
        Dim UserName As String

        Dim Slot As Byte

        Dim tIndex As Integer

        Dim UserIsAdmin As Boolean

        Dim OtherUserIsAdmin As Boolean

        UserName = Buffer.ReadASCIIString()    'Que UserName?
        Slot = Buffer.ReadByte()    'Que Slot?

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then

            Call LogGM(.Name, .Name & " Checke� el slot " & Slot & " de " & UserName)

            tIndex = NameIndex(UserName)  'Que user index?
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)

            If tIndex > 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
                        If UserList(tIndex).Invent.Object(Slot).objIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).objIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay ning�n objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Slot Inv�lido.", FontTypeNames.FONTTYPE_TALK)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reset the AutoUpdate
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub

        Call WriteConsoleMsg(UserIndex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Restart the game
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub

        'time and Time BUG!
        Call LogGM(.Name, .Name & " reinici� el mundo.")

        Call ReiniciarServidor(True)

    End With

End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha recargado los objetos.")

        Call LoadOBJData

    End With

End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha recargado los hechizos.")

        Call CargarHechizos

    End With

End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha recargado los INITs.")

        Call LoadSini

        Call WriteConsoleMsg(UserIndex, "Server.ini actualizado correctamente", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha recargado los NPCs.")

        Call CargaNpcsDat

        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")

        Call EcharPjsNoPrivilegiados

    End With

End Sub

''
' Handle the "Night" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNight(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub

        DeNoche = Not DeNoche

        Dim i  As Long

        For i = 1 To NumUsers

            If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
                Call EnviarNoche(i)

            End If

        Next i

    End With

End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Show the server form
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click

    End With

End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha borrado los SOS.")

        Call Ayuda.Reset

    End With

End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

1        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

2        Call LogGM(.Name, .Name & " ha guardado todos los chars.")

3        Call mdParty.ActualizaExperiencias
4        Call GuardarUsuarios

    End With
    Exit Sub
    
errhandleR:
LogError "Error en linea " & Erl & " - " & Err.Number & " " & Err.description
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim doTheBackUp As Boolean

        doTheBackUp = .incomingData.ReadBoolean()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub

        Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre el BackUp.")

        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0

        End If

        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim isMapPk As Boolean

        isMapPk = .incomingData.ReadBoolean()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub

        Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si es PK el mapa.")

        MapInfo(.Pos.Map).Pk = isMapPk

        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)

'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    Dim tStr   As String

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove Packet ID
        Call Buffer.ReadByte

        tStr = Buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si es restringido el mapa.")

                MapInfo(UserList(UserIndex).Pos.Map).Restringir = RestrictStringToByte(tStr)

                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & RestrictByteToString(MapInfo(.Pos.Map).Restringir), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)

'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim nomagic As Boolean

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        nomagic = .incomingData.ReadBoolean

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si est� permitido usar la magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)

'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'InviSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim noinvi As Boolean

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        noinvi = .incomingData.ReadBoolean()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si est� permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo(.Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)

'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim noresu As Boolean

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        noresu = .incomingData.ReadBoolean()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si est� permitido usar el resucitar en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo(.Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)

'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    Dim tStr   As String

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove Packet ID
        Call Buffer.ReadByte

        tStr = Buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informaci�n del terreno del mapa.")

                MapInfo(UserList(UserIndex).Pos.Map).Terreno = TerrainStringToByte(tStr)

                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & TerrainByteToString(MapInfo(.Pos.Map).Terreno), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el �nico �til es 'NIEVE' ya que al ingresarlo, la gente muere de fr�o en el mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)

'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    Dim tStr   As String

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove Packet ID
        Call Buffer.ReadByte

        tStr = Buffer.ReadASCIIString()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informaci�n de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el �nico �til es 'DUNGEON' ya que al ingresarlo, NO se sentir� el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "ChangeMapInfoStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoStealNpc(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'RoboNpcsPermitido -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim RoboNpc As Byte

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        RoboNpc = val(IIf(.incomingData.ReadBoolean(), 1, 0))

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si est� permitido robar npcs en el mapa.")

            MapInfo(UserList(UserIndex).Pos.Map).RoboNpcsPermitido = RoboNpc

            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "RoboNpcsPermitido", RoboNpc)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " RoboNpcsPermitido: " & MapInfo(.Pos.Map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeMapInfoNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoOcultar(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'OcultarSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim NoOcultar As Byte

    Dim Mapa   As Integer

    With UserList(UserIndex)

        'Remove Packet ID
        Call .incomingData.ReadByte

        NoOcultar = val(IIf(.incomingData.ReadBoolean(), 1, 0))

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then

            Mapa = .Pos.Map

            Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si est� permitido ocultarse en el mapa " & Mapa & ".")

            MapInfo(Mapa).OcultarSinEfecto = NoOcultar

            Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "OcultarSinEfecto", NoOcultar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & Mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeMapInfoNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvocar(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'InvocarSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim NoInvocar As Byte

    Dim Mapa   As Integer

    With UserList(UserIndex)

        'Remove Packet ID
        Call .incomingData.ReadByte

        NoInvocar = val(IIf(.incomingData.ReadBoolean(), 1, 0))

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then

            Mapa = .Pos.Map

            Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si est� permitido invocar en el mapa " & Mapa & ".")

            MapInfo(Mapa).InvocarSinEfecto = NoInvocar

            Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "InvocarSinEfecto", NoInvocar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & Mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))

        Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)

        Call WriteConsoleMsg(UserIndex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Allows admins to read guild messages
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim guild As String

        guild = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, .Name & " ha hecho un backup.")

        Call ES.DoBackUp    'Sino lo confunde con la id del paquete

    End With

End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/26/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        centinelaActivado = Not centinelaActivado

        Call ResetCentinelas

        If centinelaActivado Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))

        End If

    End With

End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user name
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        'Reads the userName and newUser Packets
        Dim UserName As String

        Dim newName As String

        Dim changeNameUI As Integer

        Dim GuildIndex As Integer

        UserName = Buffer.ReadASCIIString()
        newName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)

                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj est� online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else

                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))

                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")

                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)

                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")

                                Dim cantPenas As Byte

                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))

                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))

                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.Name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & time)

                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                    End If

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim newMail As String

        UserName = Buffer.ReadASCIIString()
        newMail = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else

                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(UserIndex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)

                End If

                Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim copyFrom As String

        Dim Password As String

        UserName = Replace(Buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(Buffer.ReadASCIIString(), "+", " ")

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha alterado la contrase�a de " & UserName)

            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else

                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)

                    Call WriteConsoleMsg(UserIndex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 26/09/2010
'26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
'***************************************************
    
    On Error GoTo errhandleR
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim NpcIndex As Integer

        NpcIndex = .incomingData.ReadInteger()

1        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

2        If NpcIndex >= 900 Then
3            Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearClanPretoriano.", FontTypeNames.FONTTYPE_WARNING)
4            Exit Sub

5        End If

6        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)

7        If NpcIndex <> 0 Then
8            Call LogGM(.Name, "Sumone� a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)

9        End If

    End With
Exit Sub

errhandleR:

LogError "error en HandleCreateNPC en " & Erl & " - Err: " & Err.Number & " " & Err.description

End Sub

''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 26/09/2010
'26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim NpcIndex As Integer

        NpcIndex = .incomingData.ReadInteger()

        If NpcIndex >= 900 Then
            Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearClanPretoriano.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)

        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumone� con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)

        End If

    End With

End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim Index As Byte

        Dim objIndex As Integer

        Index = .incomingData.ReadByte()
        objIndex = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Select Case Index

        Case 1
            ArmaduraImperial1 = objIndex

        Case 2
            ArmaduraImperial2 = objIndex

        Case 3
            ArmaduraImperial3 = objIndex

        Case 4
            TunicaMagoImperial = objIndex

        End Select

    End With

End Sub

''
' Handle the "ChaosArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Dim Index As Byte

        Dim objIndex As Integer

        Index = .incomingData.ReadByte()
        objIndex = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Select Case Index

        Case 1
            ArmaduraCaos1 = objIndex

        Case 2
            ArmaduraCaos2 = objIndex

        Case 3
            ArmaduraCaos3 = objIndex

        Case 4
            TunicaMagoCaos = objIndex

        End Select

    End With

End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1

        End If

        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)

    End With

End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
            frmMain.chkServerHabilitado.Value = vbUnchecked
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
            frmMain.chkServerHabilitado.Value = vbChecked

        End If

    End With

End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'Turns off the server
'***************************************************
    Dim handle As Integer

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("���" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))

        'Log
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle

        Print #handle, Date & " " & time & " server apagado por " & .Name & ". "

        Close #handle

        '        Unload frmMain

    End With

End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/CONDEN " & UserName)

            tUser = NameIndex(UserName)

            If tUser > 0 Then _
               Call VolverCriminal(tUser)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 06/09/09
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim tUser As Integer

        Dim Char As String

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            Call LogGM(.Name, "/RAJAR " & UserName)

            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else
                Char = CharPath & UserName & ".chr"

                If FileExist(Char, vbNormal) Then
                    Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(Char, "FACCIONES", "CiudMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "CrimMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingres� a ninguna Facci�n")
                    Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rArReal", 0)
                    Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rExReal", 0)
                    Call WriteVar(Char, "FACCIONES", "recCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "recReal", 0)
                    Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
                    Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim GuildIndex As Integer

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & UserName)

            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)

            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ning�n clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Request user mail
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim mail As String

        UserName = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")

                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim message As String

        message = Buffer.ReadASCIIString()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & message)

            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 03/31/07
'Set the MOTD
'Modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'   - Fixed a bug that prevented from properly setting the new number of lines.
'   - Fixed a bug that caused the player to be kicked.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim newMOTD As String

        Dim auxiliaryString() As String

        Dim LoopC As Long

        newMOTD = Buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha fijado un nuevo MOTD")

            MaxLines = UBound(auxiliaryString()) + 1

            ReDim MOTD(1 To MaxLines)

            Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))

            For LoopC = 1 To MaxLines
                Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))

                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC

            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con �xito.", FontTypeNames.FONTTYPE_INFO)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n sotuyo Dodero (Maraxus)
'Last Modification: 12/29/06
'Change the MOTD
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub

        End If

        Dim auxiliaryString As String

        Dim LoopC As Long

        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC

        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)

            End If

        End If

        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)

    End With

End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

        Call WritePong(UserIndex)

    End With

End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)

'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 01/23/10 (Marco)
'Modify server.ini
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim sLlave As String

        Dim sClave As String

        Dim sValor As String

        'Obtengo los par�metros
        sLlave = Buffer.ReadASCIIString()
        sClave = Buffer.ReadASCIIString()
        sValor = Buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then

            Dim stmp As String

            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(UserIndex, "�No puedes modificar esa informaci�n desde aqu�!", FontTypeNames.FONTTYPE_INFO)
            Else
                'Obtengo el valor seg�n llave y clave
                stmp = GetVar(IniPath & "Server.ini", sLlave, sClave)

                'Si obtengo un valor escribo en el server.ini
                If LenB(stmp) Then
                    Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modific� en server.ini (" & sLlave & " " & sClave & ") el valor " & stmp & " por " & sValor)
                    Call WriteConsoleMsg(UserIndex, "Modific� " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & stmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreatePretorianClan(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'***************************************************

    On Error GoTo errhandleR

    Dim Map    As Integer

    Dim X      As Byte

    Dim Y      As Byte

    Dim Index  As Long

    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte

        Map = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()

        ' User Admin?
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub

        ' Valid pos?
        If Not InMapBounds(Map, X, Y) Then
            Call WriteConsoleMsg(UserIndex, "Posici�n inv�lida.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Choose pretorian clan index
        If Map = Configuracion.MAPA_PRETORIANO Then
            Index = 1    ' Default clan
        Else
            Index = 2    ' Custom Clan

        End If


        ' Is already active any clan?
        If Not ClanPretoriano(Index).Active Then

            If Not ClanPretoriano(Index).SpawnClan(Map, X, Y, Index) Then
                Call WriteConsoleMsg(UserIndex, "La posici�n no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "El clan pretoriano se encuentra activo en el mapa " & _
                                            ClanPretoriano(Index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    Exit Sub

errhandleR:
    Call LogError("Error en HandleCreatePretorianClan. Error: " & Err.Number & " - " & Err.description)

End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDeletePretorianClan(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'***************************************************

    On Error GoTo errhandleR

    Dim Map    As Integer

    Dim Index  As Long

    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte

        Map = .incomingData.ReadInteger()

        ' User Admin?
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub

        ' Valid map?
        If Map < 1 Or Map > NumMaps Then
            Call WriteConsoleMsg(UserIndex, "Mapa inv�lido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        For Index = 1 To UBound(ClanPretoriano)

            ' Search for the clan to be deleted
            If ClanPretoriano(Index).ClanMap = Map Then
                ClanPretoriano(Index).DeleteClan
                Exit For

            End If

        Next Index

    End With

    Exit Sub

errhandleR:
    Call LogError("Error en HandleDeletePretorianClan. Error: " & Err.Number & " - " & Err.description)

End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.logged)

        Call .outgoingData.WriteByte(.clase)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(UserIndex).outgoingData.WriteLong(UserList(UserIndex).Stats.Banco)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(UserList(UserIndex).ComUsu.DestNick)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateBankGold)
        Call .WriteLong(UserList(UserIndex).Stats.Banco)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)

'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)

'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)

'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, _
                          ByVal Map As Integer, _
                          ByVal version As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR
    

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        Call .WriteInteger(version)
        Call .WriteBoolean(MapInfo(Map).Pk)
        
    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, _
                             ByVal Chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal color As Long)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, color))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, _
                           ByVal Chat As String, _
                           ByVal FontIndex As FontTypeNames)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, _
                             ByVal Chat As String, _
                             ByVal FontIndex As FontTypeNames)

'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(Chat, FontIndex))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, _
                          ByVal Chat As String, _
                          Optional ByVal IsMOTD As Boolean = False)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 02/12/12
'Writes the "GuildChat" message to the given user's outgoing data buffer
'D'Artagnan - New optional param for MOTD messages
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(Chat, IsMOTD))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, _
                                ByVal Body As Integer, _
                                ByVal Head As Integer, _
                                ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal X As Byte, _
                                ByVal Y As Byte, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal Name As String, _
                                ByVal NickColor As Byte, _
                                ByVal Privileges As Byte, _
                                ByVal MinHp As Long, _
                                ByVal MaxHP As Long, _
                                ByVal esVip As Byte, _
                                ByVal esConquista As Boolean, Optional ByVal esNPC As Byte = 0)
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(Body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                                                              helmet, Name, NickColor, Privileges, MinHp, MaxHP, esConquista, esVip, esNPC))
    Exit Sub

errhandleR:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'**************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, _
                              ByVal CharIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
Public Sub WriteCharacterChange(ByVal UserIndex As Integer, _
                                ByVal Body As Integer, _
                                ByVal Head As Integer, _
                                ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(Body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, _
                             ByVal GrhIndex As Integer, _
                             ByVal X As Byte, _
                             ByVal Y As Byte)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByVal Blocked As Boolean)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, _
                         ByVal midi As Integer, _
                         Optional ByVal Loops As Integer = -1)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayMidi" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, Loops))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, _
                         ByVal wave As Byte, _
                         ByVal X As Byte, _
                         ByVal Y As Byte)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim Tmp    As String

    Dim i      As Long

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)

        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AreaChanged" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, _
                         ByVal CharIndex As Integer, _
                         ByVal FX As Integer, _
                         ByVal FXLoops As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.updateuserstats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 25/05/2011 (Amraphen)
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
'25/05/2011: Amraphen - Ahora se env�a la defensa seg�n se tiene equipado armadura de segunda jerarqu�a o no.
'***************************************************
    On Error GoTo errhandleR

    ' @@ Miqueas : Reduccion de Lag

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)

        Dim objIndex As Integer

        Dim obData As ObjData

        objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex

        If objIndex > 0 Then
            obData = ObjData(objIndex)

        End If

        Call .WriteInteger(objIndex)
        'Call .WriteASCIIString(obData.Name)

        Call .WriteInteger(UserList(UserIndex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(Slot).Equipped)

        'Call .WriteInteger(obData.GrhIndex)
        'Call .WriteByte(obData.OBJType)
        'Call .WriteInteger(obData.MaxHIT)
        'Call .WriteInteger(obData.MinHIT)

        Call .WriteByte(obData.Real)
        Call .WriteByte(obData.Caos)

        'If obData.Real = 2 Or obData.Caos = 2 Then
        '        Call .WriteInteger(ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef * MOD_DEF_SEG_JERARQUIA)
        '        Call .WriteInteger(ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef * MOD_DEF_SEG_JERARQUIA)
        'Else
        '        Call .WriteInteger(obData.MaxDef)
        '        Call .WriteInteger(obData.MinDef)

        'End If

        Call .WriteSingle(SalePrice(objIndex))

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de s�lo Def
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)

        Dim objIndex As Integer

        Dim obData As ObjData

        objIndex = UserList(UserIndex).BancoInvent.Object(Slot).objIndex

        Call .WriteInteger(objIndex)

        If objIndex > 0 Then
            obData = ObjData(objIndex)

        End If

        'Call .WriteASCIIString(obData.Name)

        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(Slot).Amount)

        'Call .WriteInteger(obData.GrhIndex)
        'Call .WriteByte(obData.OBJType)
        'Call .WriteInteger(obData.MaxHIT)
        'Call .WriteInteger(obData.MinHIT)
        'Call .WriteInteger(obData.MaxDef)
        'Call .WriteInteger(obData.MinDef)
        'Call .WriteLong(obData.Valor)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(Slot))

        ' @@ Miqueas : Reduccion de Lag
        'If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        '        Call .WriteASCIIString(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).Nombre)
        'Else
        '        Call .WriteASCIIString("(None)")
        'End If

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Obj    As ObjData

    Dim validIndexes() As Integer

    Dim Count  As Integer

    ReDim validIndexes(1 To UBound(ArmasHerrero()))

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)

        For i = 1 To UBound(ArmasHerrero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i

        ' @@ Miqueas : Reduccion de Lag

        ' Write the number of objects in the list
        Call .WriteInteger(Count)

        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))

            'Call .WriteInteger(Obj.GrhIndex)
            'Call .WriteASCIIString(Obj.Name)

            'Call .WriteInteger(Obj.LingH)
            'Call .WriteInteger(Obj.LingP)
            'Call .WriteInteger(Obj.LingO)

            'Call .WriteInteger(Obj.Upgrade)
        Next i

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Obj    As ObjData

    Dim validIndexes() As Integer

    Dim Count  As Integer

    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)

        For i = 1 To UBound(ArmadurasHerrero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i

        ' @@ Miqueas : Reduccion de Lag

        ' Write the number of objects in the list
        Call .WriteInteger(Count)

        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
            'Call .WriteASCIIString(Obj.Name)
            'Call .WriteInteger(Obj.GrhIndex)
            'Call .WriteInteger(Obj.LingH)
            'Call .WriteInteger(Obj.LingP)
            'Call .WriteInteger(Obj.LingO)

            'Call .WriteInteger(Obj.Upgrade)
        Next i

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Obj    As ObjData

    Dim validIndexes() As Integer

    Dim Count  As Integer

    ReDim validIndexes(1 To UBound(ObjCarpintero()))

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)

        For i = 1 To UBound(ObjCarpintero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i

        ' @@ Miqueas : Reduccion de Lag

        ' Write the number of objects in the list
        Call .WriteInteger(Count)

        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
            'Call .WriteASCIIString(Obj.Name)
            'Call .WriteInteger(Obj.GrhIndex)
            'Call .WriteInteger(Obj.Madera)
            'Call .WriteInteger(Obj.MaderaElfica)

            'Call .WriteInteger(Obj.Upgrade)
        Next i

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal message As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ErrorMsg" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal objIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSignal" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(objIndex).texto)
        Call .WriteInteger(ObjData(objIndex).GrhSecundario)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, _
                                       ByVal Slot As Byte, _
                                       ByRef Obj As Obj, _
                                       ByVal price As Single)

On Error GoTo errhandleR
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Last Modified by: Budi
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de s�lo Def
'***************************************************
    On Error GoTo errhandleR

    ' @@ Miqueas : Reduccion de Lag

    Dim ObjInfo As ObjData

1    If Obj.objIndex >= LBound(ObjData()) And Obj.objIndex <= UBound(ObjData()) Then
2        ObjInfo = ObjData(Obj.objIndex)
4    End If

    'If UserList(Userindex).ConnIDValida Then
    If UserList(UserIndex).ConnID = -1 Then
    LogError UserIndex & " tiene connID = -1 en ChangeNPCInventorySlot"
    Exit Sub
    End If
    'End If
    

5    With UserList(UserIndex).outgoingData
6        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
7        Call .WriteByte(Slot)
8
9        Call .WriteInteger(Obj.objIndex)

        'Call .WriteASCIIString(ObjInfo.Name)
10        Call .WriteInteger(Obj.Amount)
11        Call .WriteSingle(price)
12        Call .WriteInteger(ObjInfo.copas)
        'Call .WriteInteger(ObjInfo.GrhIndex)

        'Call .WriteByte(ObjInfo.OBJType)
        'Call .WriteInteger(ObjInfo.MaxHIT)
        'Call .WriteInteger(ObjInfo.MinHIT)
        'Call .WriteInteger(ObjInfo.MaxDef)
        'Call .WriteInteger(ObjInfo.MinDef)

    End With

    Exit Sub

errhandleR:

    LogError "error en WriteChanNP en " & Erl & ". Err " & Err.Number & " " & Err.description
    
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Fame" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)

        Call .WriteLong(UserList(UserIndex).Reputacion.AsesinoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BandidoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BurguesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.LadronesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.NobleRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.PlebeRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.Promedio)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)

        Call .WriteLong(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.CriminalesMatados)

        'TODO : Este valor es calculable, no deber�a NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call .WriteLong(UserList(UserIndex).Stats.UsuariosMatados)

        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)

        Call .WriteByte(UserList(UserIndex).clase)
        Call .WriteLong(UserList(UserIndex).Counters.Pena)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal ResetPoints As Byte, ByVal skillPoints As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LevelUp" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteByte(ResetPoints)
        Call .WriteInteger(skillPoints)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, _
                            ByVal ForumType As eForumType, _
                            ByRef Title As String, _
                            ByRef Author As String, _
                            ByRef message As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'Writes the "AddForumMsg" message to the given user's outgoing data buffer
'02/01/2010: ZaMa - Now sends Author and forum type
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteByte(ForumType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Author)
        Call .WriteASCIIString(message)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowForumForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim Visibilidad As Byte

    Dim CanMakeSticky As Byte

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.ShowForumForm)

        Visibilidad = eForumVisibility.ieGENERAL_MEMBER

        If esCaos(UserIndex) Or esGM(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER

        End If

        If esArmada(UserIndex) Or esGM(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER

        End If

        Call .outgoingData.WriteByte(Visibilidad)

        ' Pueden mandar sticky los gms o los del consejo de armada/caos
        If esGM(UserIndex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1

        End If

        Call .outgoingData.WriteByte(CanMakeSticky)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, _
                             ByVal CharIndex As Integer, _
                             ByVal invisible As Boolean)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'Writes the "SendSkills" message to the given user's outgoing data buffer
'11/19/09: Pato - Now send the percentage of progress of the skills.
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.SendSkills)

        For i = 1 To NUMSKILLS
            Call .outgoingData.WriteByte(.Stats.UserSkills(i))

        Next i

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim str    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)

        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NPCNAME & SEPARATOR
        Next i

        If LenB(str) > 0 Then _
           str = Left$(str, Len(str) - 1)

        Call .WriteASCIIString(str)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, _
                          ByVal guildNews As String, _
                          ByRef enemies() As String, _
                          ByRef allies() As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNews" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)

        Call .WriteASCIIString(guildNews)

        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

        Tmp = vbNullString

        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OfferDetails" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)

        Call .WriteASCIIString(details)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)

        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)

        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                              ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal reputation As Long, _
                              ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                              ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterInfo" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)

        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)

        Call .WriteByte(level)
        Call .WriteLong(gold)
        Call .WriteLong(bank)
        Call .WriteLong(reputation)

        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)

        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)

        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, _
                                ByRef guildList() As String, _
                                ByRef MemberList() As String, _
                                ByVal guildNews As String, _
                                ByRef joinRequests() As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)

        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

        ' Store guild news
        Call .WriteASCIIString(guildNews)

        ' Prepare the join request's list
        Tmp = vbNullString

        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserIndex As Integer, _
                                ByRef guildList() As String, _
                                ByRef MemberList() As String)

'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'Writes the "GuildMemberInfo" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)

        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, _
                             ByVal GuildName As String, _
                             ByVal founder As String, _
                             ByVal foundationDate As String, _
                             ByVal leader As String, _
                             ByVal URL As String, _
                             ByVal memberCount As Integer, _
                             ByVal electionsOpen As Boolean, _
                             ByVal alignment As String, _
                             ByVal enemiesCount As Integer, _
                             ByVal AlliesCount As Integer, _
                             ByVal antifactionPoints As String, _
                             ByRef codex() As String, _
                             ByVal guildDesc As String, _
                             ByVal GuildLevel As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDetails" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim temp   As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)

        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)

        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)

        Call .WriteASCIIString(alignment)

        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)

        Call .WriteASCIIString(antifactionPoints)

        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i

        If Len(temp) > 1 Then _
           temp = Left$(temp, Len(temp) - 1)

        Call .WriteASCIIString(temp)

        Call .WriteASCIIString(guildDesc)

        Call .WriteInteger(GuildLevel)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowGuildAlign" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildAlign(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "ShowGuildAlign" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildAlign)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 08/12/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'And updates user position
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    Call WritePosUpdate(UserIndex)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)

        Call .WriteASCIIString(details)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.TradeOK)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankOK)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, _
                                    ByVal OfferSlot As Byte, _
                                    ByVal objIndex As Integer, _
                                    ByVal Amount As Long)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de s�lo Def
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)

        Call .WriteByte(OfferSlot)
        Call .WriteInteger(objIndex)
        Call .WriteLong(Amount)

        If objIndex > 0 Then
            Call .WriteInteger(ObjData(objIndex).GrhIndex)
            Call .WriteByte(ObjData(objIndex).OBJType)
            Call .WriteInteger(ObjData(objIndex).MaxHIT)
            Call .WriteInteger(ObjData(objIndex).MinHIT)
            Call .WriteInteger(ObjData(objIndex).MaxDef)
            Call .WriteInteger(ObjData(objIndex).MinDef)
            Call .WriteLong(SalePrice(objIndex))
            Call .WriteASCIIString(ObjData(objIndex).Name)
        Else    ' Borra el item
            Call .WriteInteger(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
            Call .WriteASCIIString(vbNullString)

        End If

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal UserIndex As Integer, ByVal night As Boolean)

'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Writes the "SendNight" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendNight)
        Call .WriteBoolean(night)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)

        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)

        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i

        If LenB(Tmp) <> 0 Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowDenounces" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenounces(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "ShowDenounces" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim DenounceIndex As Long

    Dim DenounceList As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowDenounces)

        For DenounceIndex = 1 To Denuncias.Longitud
            DenounceList = DenounceList & Denuncias.VerElemento(DenounceIndex, False) & SEPARATOR
        Next DenounceIndex

        If LenB(DenounceList) <> 0 Then _
           DenounceList = Left$(DenounceList, Len(DenounceList) - 1)

        Call .WriteASCIIString(DenounceList)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowPartyForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "ShowPartyForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    Dim PI     As Integer

    Dim members(PARTY_MAXMEMBERS) As Integer

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPartyForm)

        PI = UserList(UserIndex).PartyIndex
        Call .WriteByte(CByte(Parties(PI).EsPartyLeader(UserIndex)))

        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(members())

            For i = 1 To PARTY_MAXMEMBERS

                If members(i) > 0 Then
                    Tmp = Tmp & UserList(members(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(members(i))) & ")" & SEPARATOR

                End If

            Next i

        End If

        If LenB(Tmp) <> 0 Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)
        Call .WriteLong(Parties(PI).ObtenerExperienciaTotal)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, _
                                    ByVal currentMOTD As String)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)

        Call .WriteASCIIString(currentMOTD)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, _
                             ByRef userNamesList() As String, _
                             ByVal Cant As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Dim i      As Long

    Dim Tmp    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)

        ' Prepare user's names list
        For i = 1 To Cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i

        If Len(Tmp) Then _
           Tmp = Left$(Tmp, Len(Tmp) - 1)

        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String

    With UserList(UserIndex).outgoingData

        If .length = 0 Then _
           Exit Sub

        sndData = .ReadASCIIStringFixed(.length)

        Call EnviarDatosASlot(UserIndex, sndData)

    End With

End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, _
                                           ByVal invisible As Boolean) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)

        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)

        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)

    End With

End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, _
                                                  ByVal newNick As String) As String

'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'Prepares the "Change Nick" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChangeNick)

        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(newNick)

        PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal Chat As String, _
                                           ByVal CharIndex As Integer, _
                                           ByVal color As Long) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)

        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)

        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @param    MessageType type of console message (General, Guild, Party)
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, _
                                         ByVal FontIndex As FontTypeNames) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/05/11 (D'Artagnan)
'Prepares the "MessageType" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)

        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)

    End With

End Function

Public Function PrepareMessageAyudaClan(ByVal UserIndex As Integer)

    With auxiliarBuffer
        
        Call .WriteByte(ServerPacketID.ayudaclan)

        Dim i As Long, Integrantes() As String
        Dim tmpIndex As Integer
        
        .WriteByte guilds(UserList(UserIndex).GuildIndex).CantidadDeMiembros
        
        Integrantes = guilds(UserList(UserIndex).GuildIndex).GetMemberList
        
        For i = LBound(Integrantes) To UBound(Integrantes)
            tmpIndex = NameIndex(Integrantes(i))
            If tmpIndex > 0 Then
                If UserList(tmpIndex).flags.YaPediAyuda Then
                    .WriteASCIIString UserList(tmpIndex).Name & " (" & UserList(tmpIndex).Pos.Map & "-" & UserList(tmpIndex).Pos.X & "-" & UserList(tmpIndex).Pos.Y & ")"
                Else
                    .WriteASCIIString ""
                End If
            Else
            .WriteASCIIString ""
            End If
        Next i
        

        PrepareMessageAyudaClan = .ReadASCIIStringFixed(.length)

    End With

End Function


Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, _
                                          ByVal FontIndex As FontTypeNames) As String

'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Prepares the "CommerceConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CommerceChat)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)

        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, _
                                       ByVal FX As Integer, _
                                       ByVal FXLoops As Integer) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)

        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, _
                                       ByVal X As Byte, _
                                       ByVal Y As Byte) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)

        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal Chat As String, _
                                        Optional ByVal IsMOTD As Boolean = False) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 02/12/12
'Prepares the "GuildChat" message and returns it
'D'Artagnan - New optional param for MOTD messages
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(Chat)
        Call .WriteBoolean(IsMOTD)

        PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String

'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Chat)

        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Integer, _
                                       Optional ByVal Loops As Integer = -1) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMIDI)
        Call .WriteInteger(midi)
        Call .WriteInteger(Loops)

        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "RainToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)

        PrepareMessageRainToggle = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)

        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, _
                                            ByVal Y As Byte, _
                                            ByVal Blocked As Boolean) As String

'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)

        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, _
                                           ByVal X As Byte, _
                                           ByVal Y As Byte) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(GrhIndex)

        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer, Optional ByVal esGM As Integer) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)

        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)

        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal Body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal X As Byte, _
                                              ByVal Y As Byte, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Name As String, _
                                              ByVal NickColor As Byte, _
                                              ByVal Privileges As Byte, _
                                              ByVal MinHp As Long, _
                                              ByVal MaxHP As Long, _
                                              Optional ByVal conquisto As Boolean = False, _
                                              Optional ByVal esVip As Byte = 0, Optional ByVal esNPC As Byte = 0) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)

        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteByte(NickColor)
        Call .WriteByte(Privileges)
        Call .WriteLong(MinHp)
        Call .WriteLong(MaxHP)
        Call .WriteBoolean(conquisto)
        Call .WriteByte(esVip)
        Call .WriteByte(esNPC)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal Body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)

        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)

        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, _
                                            ByVal X As Byte, _
                                            ByVal Y As Byte) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)

        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)

    End With

End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String

'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Prepares the "ForceCharMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)

        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, _
                                                 ByVal NickColor As Byte, _
                                                 ByRef Tag As String) As String

'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Mart�n Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'15/01/2010: ZaMa - Now sends the nick color instead of the status.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)

        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteByte(NickColor)
        Call .WriteASCIIString(Tag)

        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMSG)
        Call .WriteASCIIString(message)

        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'
'***************************************************
    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.StopWorking)

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/2010
'
'***************************************************
    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(Slot)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Handles the "SetDialog" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetDialog(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 18/11/2010
'20/11/2010: ZaMa - Arreglo privilegios.
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet id
        Call Buffer.ReadByte

        Dim NewDialog As String

        NewDialog = Buffer.ReadASCIIString

        Call .incomingData.CopyBuffer(Buffer)

        If .flags.TargetNPC > 0 Then

            If esGM(UserIndex) Then
                'Replace the NPC's dialog.
                Npclist(.flags.TargetNPC).Desc = NewDialog

            End If

        End If

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "Impersonate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImpersonate(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'
'***************************************************
    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte

        ' Dsgm/Dsrm/Rm
        If Not esGM(UserIndex) Then Exit Sub

        Dim NpcIndex As Integer

        NpcIndex = .flags.TargetNPC

        If NpcIndex = 0 Then Exit Sub

        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)

        ' Teleports user to npc's coords
        Call WarpUserChar(UserIndex, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, _
                          Npclist(NpcIndex).Pos.Y, False, True)

        ' Log gm
        Call LogGM(.Name, "/IMPERSONAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)

        ' Remove npc
        Call QuitarNPC(NpcIndex)

    End With

End Sub

''
' Handles the "Imitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImitate(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'
'***************************************************
    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte

        ' Dsgm/Dsrm/Rm/ConseRm
        If Not esGM(UserIndex) Then Exit Sub

        Dim NpcIndex As Integer

        NpcIndex = .flags.TargetNPC

        If NpcIndex = 0 Then Exit Sub

        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        Call LogGM(.Name, "/MIMETIZAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)

    End With

End Sub

''
' Handles the "RecordAdd" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleRecordAdd(ByVal UserIndex As Integer)

'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'
'**************************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet id
        Call Buffer.ReadByte

        Dim UserName As String

        Dim Reason As String

        UserName = Buffer.ReadASCIIString
        Reason = Buffer.ReadASCIIString

        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then

            'Verificamos que exista el personaje
            If Not FileExist(CharPath & UCase$(UserName) & ".chr") Then
                Call WriteShowMessageBox(UserIndex, "El personaje no existe")
            Else
                'Agregamos el seguimiento
                Call AddRecord(UserIndex, UserName, Reason)

                'Enviamos la nueva lista de personajes
                Call WriteRecordList(UserIndex)

            End If

        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RecordAddObs" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordAddObs(ByVal UserIndex As Integer)

'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'
'**************************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet id
        Call Buffer.ReadByte

        Dim RecordIndex As Byte

        Dim Obs As String

        RecordIndex = Buffer.ReadByte
        Obs = Buffer.ReadASCIIString

        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then
            'Agregamos la observaci�n
            Call AddObs(UserIndex, RecordIndex, Obs)

            'Actualizamos la informaci�n
            Call WriteRecordDetails(UserIndex, RecordIndex)

        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "RecordRemove" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordRemove(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************
    Dim RecordIndex As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        RecordIndex = .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        'S�lo dioses pueden remover los seguimientos, los otros reciben una advertencia:
        If (.flags.Privilegios And PlayerType.Dios) Then
            Call RemoveRecord(RecordIndex)
            Call WriteShowMessageBox(UserIndex, "Se ha eliminado el seguimiento.")
            Call WriteRecordList(UserIndex)
        Else
            Call WriteShowMessageBox(UserIndex, "S�lo los dioses pueden eliminar seguimientos.")

        End If

    End With

End Sub

''
' Handles the "RecordListRequest" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordListRequest(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Call WriteRecordList(UserIndex)

    End With

End Sub

''
' Writes the "RecordDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetails(ByVal UserIndex As Integer, ByVal RecordIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordDetails" message to the given user's outgoing data buffer
'***************************************************
    Dim i      As Long

    Dim tIndex As Integer

    Dim tmpStr As String

    Dim TempDate As Date

    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.RecordDetails)

        'Creador y motivo
        Call .WriteASCIIString(Records(RecordIndex).Creador)
        Call .WriteASCIIString(Records(RecordIndex).Motivo)

        tIndex = NameIndex(Records(RecordIndex).Usuario)

        'Status del pj (online?)
        Call .WriteBoolean(tIndex > 0)

        'Escribo la IP seg�n el estado del personaje
        If tIndex > 0 Then
            'La IP Actual
            tmpStr = UserList(tIndex).Ip
        Else    'String nulo
            tmpStr = vbNullString

        End If

        Call .WriteASCIIString(tmpStr)

        'Escribo tiempo online seg�n el estado del personaje
        If tIndex > 0 Then
            'Tiempo logueado.
            TempDate = Now - UserList(tIndex).LogOnTime
            tmpStr = Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate)
        Else
            'Env�o string nulo.
            tmpStr = vbNullString

        End If

        Call .WriteASCIIString(tmpStr)

        'Escribo observaciones:
        tmpStr = vbNullString

        If Records(RecordIndex).NumObs Then

            For i = 1 To Records(RecordIndex).NumObs
                tmpStr = tmpStr & Records(RecordIndex).Obs(i).Creador & "> " & Records(RecordIndex).Obs(i).Detalles & vbCrLf
            Next i

            tmpStr = Left$(tmpStr, Len(tmpStr) - 1)

        End If

        Call .WriteASCIIString(tmpStr)

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RecordList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordList(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordList" message to the given user's outgoing data buffer
'***************************************************
    Dim i      As Long

    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.RecordList)

        Call .WriteByte(NumRecords)

        For i = 1 To NumRecords
            Call .WriteASCIIString(Records(i).Usuario)
        Next i

    End With

    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Handles the "RecordDetailsRequest" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordDetailsRequest(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 07/04/2011
'Handles the "RecordListRequest" message
'***************************************************
    Dim RecordIndex As Byte

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        RecordIndex = .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Call WriteRecordDetails(UserIndex, RecordIndex)

    End With

End Sub

Public Sub HandleMoveItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Ignacio Mariano Tirabasso (Budi)
'Last Modification: 01/01/2011
'
'***************************************************

    With UserList(UserIndex)

        Dim originalSlot As Byte

        Dim newSlot As Byte

        Call .incomingData.ReadByte

        originalSlot = .incomingData.ReadByte
        newSlot = .incomingData.ReadByte
        Call .incomingData.ReadByte
        
        If .flags.Comerciando Then Exit Sub

        Call mod_DragAndDrop.moveItem(UserIndex, originalSlot, newSlot)

    End With

End Sub

Public Function PrepareMessageCharacterAttackMovement(ByVal CharIndex As Integer) As String

'***************************************************
'Author: Amraphen
'Last Modification: 24/05/2011
'Prepares the "CharacterAttackMovement" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterAttackMovement)
        Call .WriteInteger(CharIndex)

        PrepareMessageCharacterAttackMovement = .ReadASCIIStringFixed(.length)

    End With

End Function

''
' Writes the "StrDextRunningOut" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Seconds Seconds left.

Public Sub WriteStrDextRunningOut(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Dalmasso (CHOTS)
'Last Modification: 08/06/2011
'
'***************************************************
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.StrDextRunningOut)

    End With

End Sub

''
' Handles the "PMSend" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePMSend(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMSend" message.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim message As String

        Dim TargetIndex As Integer

        UserName = Buffer.ReadASCIIString
        message = Buffer.ReadASCIIString

        TargetIndex = NameIndex(UserName)

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            If TargetIndex = 0 Then    'Offline
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call AgregarMensajeOFF(UserName, .Name, message)
                    Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_GM)
                Else
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else    'Online
                Call AgregarMensaje(TargetIndex, .Name, message)
                Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_GM)

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

''
' Handles the "PMList" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandlePMList(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMList" message.
'***************************************************
    Dim LoopC  As Long

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .UltimoMensaje = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes mensajes privados.", FontTypeNames.FONTTYPE_INFOBOLD)
        Else
            'Env�a la lista de mensajes privados al usuario:
            Call WriteConsoleMsg(UserIndex, "Mensajes privados: ", FontTypeNames.FONTTYPE_INFOBOLD)

            For LoopC = 1 To MAX_PRIVATE_MESSAGES

                With .Mensajes(LoopC)

                    If LenB(.Contenido) = 0 Then
                        Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> VAC�O.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If .Nuevo Then
                            Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> (!)" & .Contenido, FontTypeNames.FONTTYPE_FIGHT)
                            .Nuevo = False
                        Else
                            Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> " & .Contenido, FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End With

            Next LoopC

        End If

    End With

End Sub

Public Sub HandlePMDeleteList(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMDeleteList" message.
'***************************************************

    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte

        Call LimpiarMensajes(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Se han borrado tus mensajes privados.", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

Public Sub HandlePMDeleteUser(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMDeleteUser" message.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim MPIndex As Byte

        Dim TargetIndex As Integer

        Dim LoopC As Long

        UserName = Buffer.ReadASCIIString
        MPIndex = Buffer.ReadByte

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            TargetIndex = NameIndex(UserName)

            If TargetIndex = 0 Then    'Offline
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    If MPIndex = 0 Then
                        Call LimpiarMensajesOFF(UserName)

                        Call WriteConsoleMsg(UserIndex, "Mensajes borrados.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf MPIndex >= 1 And MPIndex <= MAX_PRIVATE_MESSAGES Then
                        Call BorrarMensajeOFF(UserName, MPIndex)

                        Call WriteConsoleMsg(UserIndex, "Mensaje borrado.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else    'Online

                If MPIndex = 0 Then
                    Call LimpiarMensajes(TargetIndex)
                    Call WriteConsoleMsg(UserIndex, "Mensajes borrados.", FontTypeNames.FONTTYPE_INFO)

                    Call WriteConsoleMsg(UserIndex, "Mensajes borrados.", FontTypeNames.FONTTYPE_INFO)
                ElseIf MPIndex >= 1 And MPIndex <= MAX_PRIVATE_MESSAGES Then
                    Call BorrarMensaje(TargetIndex, MPIndex)

                    Call WriteConsoleMsg(UserIndex, "Mensaje borrado.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

Public Sub HandlePMListUser(ByVal UserIndex As Integer)

'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMListUser" message.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim MPIndex As Byte

        Dim TargetIndex As Integer

        Dim LoopC As Long

        UserName = Buffer.ReadASCIIString
        MPIndex = Buffer.ReadByte

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            TargetIndex = NameIndex(UserName)

            If TargetIndex = 0 Then    'Offline
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else    'Online

                With UserList(TargetIndex)

                    If .UltimoMensaje = 0 Then
                        Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " no tiene mensajes privados.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Mensajes privados de " & UserName & ":", FontTypeNames.FONTTYPE_INFOBOLD)

                        With .Mensajes(LoopC)

                            If LenB(.Contenido) = 0 Then
                                Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> VAC�O.", FontTypeNames.FONTTYPE_INFO)
                            Else

                                If .Nuevo Then
                                    Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> (!)" & .Contenido, FontTypeNames.FONTTYPE_FIGHT)
                                Else
                                    Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> " & .Contenido, FontTypeNames.FONTTYPE_INFO)

                                End If

                            End If

                        End With

                    End If

                End With

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

Private Sub HandleDropObj(ByVal UserIndex As Integer)
'***************************************************
'Author: maTih.-
'Last Modification: 10/11/2012 - ^[GS]^
'***************************************************

    With UserList(UserIndex)

        Dim selInvSlot As Byte         ' <<< Slot.

        Dim TargetX As Byte        ' <<< Posici�n X.

        Dim TargetY As Byte        ' <<< Posici�n Y.

        Dim Amount As Integer    ' <<< Cantidad.

        Dim tNpc As Integer      ' <<< Npc?.

        Dim tUser As Integer      ' <<< Usuario?.

        'Dim targetObj  As Obj         ' <<< -

        'Remove packetID.
        Call .incomingData.ReadByte

        'Get the incoming Data.
        selInvSlot = .incomingData.ReadByte()
        TargetX = .incomingData.ReadByte()
        TargetY = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        tUser = MapData(UserList(UserIndex).Pos.Map, TargetX, TargetY).UserIndex
        tNpc = MapData(UserList(UserIndex).Pos.Map, TargetX, TargetY).NpcIndex
        
        If .flags.Comerciando Then Exit Sub
        If .flags.Muerto <> 0 Then Exit Sub
        If Amount < 0 Then Exit Sub
        
        If tNpc <> 0 Then
            mod_DragAndDrop.DragToNPC UserIndex, tNpc, selInvSlot, Amount
        ElseIf tUser <> 0 Then
            mod_DragAndDrop.DragToUser UserIndex, tUser, selInvSlot, Amount
        Else
            mod_DragAndDrop.DragToPos UserIndex, TargetX, TargetY, selInvSlot, Amount
        End If

    End With

End Sub

Private Sub handleSendReto(ByVal user_Index As Integer)

    With UserList(user_Index)

        Dim Buffer As New clsByteQueue

        Dim myteam As String

        Dim enemy As String

        Dim tenemy As String

        Dim bydrop As Boolean

        Dim byGold As Long

        Dim serror As String

        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        Call Buffer.ReadByte

        myteam = Buffer.ReadASCIIString()
        enemy = Buffer.ReadASCIIString()
        tenemy = Buffer.ReadASCIIString()

        byGold = Buffer.ReadLong()
        bydrop = Buffer.ReadBoolean()

        Call Mod_Retos2vs2.set_reto_struct(user_Index, myteam, enemy, tenemy, bydrop, byGold)

        If (Mod_Retos2vs2.can_send_reto(user_Index, serror) = True) Then
            Call Mod_Retos2vs2.send_Reto(user_Index)
        Else
            Call WriteConsoleMsg(user_Index, serror, FontTypeNames.FONTTYPE_CITIZEN)

        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

End Sub

Private Sub handleAcceptReto(ByVal user_Index As Integer)

    Dim Buffer As New clsByteQueue

    Dim iName  As String
    Dim cName  As String

    Set Buffer = New clsByteQueue

    Call Buffer.CopyBuffer(UserList(user_Index).incomingData)

    Call Buffer.ReadByte

    iName = Buffer.ReadASCIIString()
    cName = UCase$(iName)

    With UserList(user_Index)

        If (UCase$(.sReto.nick_sender) = cName) Then
            Call Mod_Retos2vs2.accept_Reto(user_Index, cName)
        Else

            If (UCase$(.mReto.request_name) = cName) Then
                Call Mod_Retos1vs1.accept_Reto(user_Index, cName)

            End If

        End If

    End With

    Call UserList(user_Index).incomingData.CopyBuffer(Buffer)

End Sub

Public Sub handleOtherSendReto(ByVal user_Index As Integer)

    Dim Buffer As New clsByteQueue

    Dim eName  As String

    Dim g_gold As Long

    Dim g_drop As Boolean

    Dim mError As String

    Call Buffer.CopyBuffer(UserList(user_Index).incomingData)

    Call Buffer.ReadByte

    eName = Buffer.ReadASCIIString()
    g_gold = Buffer.ReadLong()
    g_drop = Buffer.ReadBoolean()

    If (Mod_Retos1vs1.can_sendReto(user_Index, eName, g_gold, g_drop, mError) = True) Then
        Call Mod_Retos1vs1.send_Reto(user_Index, NameIndex(eName), g_gold, g_drop)
    Else
        Call Protocol.WriteConsoleMsg(user_Index, mError, FontTypeNames.FONTTYPE_CITIZEN)

    End If

    Call UserList(user_Index).incomingData.CopyBuffer(Buffer)

End Sub

Private Sub HandleSetMenu(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        .incomingData.ReadByte

        '1 inventario
        '0 spell

        .flags.MenuCliente = .incomingData.ReadByte
        .flags.LastSlotClient = .incomingData.ReadByte

    End With

End Sub

Public Sub WriteCharacterUpdateHp(ByVal UserIndex As Integer, _
                                  ByVal NpcIndex As Integer, _
                                  ByVal MinHp As Long, _
                                  ByVal MaxHP As Long)

    On Error GoTo errhandleR

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterUpdateHP(NpcIndex, MinHp, MaxHP))
    Exit Sub

errhandleR:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Function PrepareMessageCharacterUpdateHP(ByVal CharIndex As Integer, _
                                                ByVal MinHp As Long, _
                                                ByVal MaxHP As Long) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterUpdateHp)
        Call .WriteInteger(CharIndex)
        Call .WriteLong(MinHp)
        Call .WriteLong(MaxHP)

        PrepareMessageCharacterUpdateHP = .ReadASCIIStringFixed(.length)

    End With

End Function

Public Function PrepareMessageCreateDamage(ByVal X As Byte, _
                                           ByVal Y As Byte, _
                                           ByVal DamageValue As Long, _
                                           ByVal DamageType As Byte)

' @ Envia el paquete para crear da�o (Y)

    With auxiliarBuffer

        .WriteByte ServerPacketID.CreateDamage
        .WriteByte X
        .WriteByte Y
        .WriteLong DamageValue
        .WriteByte DamageType

        PrepareMessageCreateDamage = .ReadASCIIStringFixed(.length)

    End With

End Function

Private Sub HandleCanjer(ByVal UserIndex As Integer)

'***************************************************
'Author: ZheTa Adaptado para Iriwynne AO
'***************************************************

    With UserList(UserIndex)

        .incomingData.ReadByte    'Remove packet ID

        Dim tmpStr As String
        Dim ss As Byte
        
        ss = .incomingData.ReadByte
        
        If ss <= 0 Then Exit Sub
        
        If .flags.Comerciando Then Exit Sub
        If .flags.Muerto <> 0 Then Exit Sub
        
        If Not Mod_Cofres.GetCanje(UserIndex, ss, tmpStr) Then    ' @@ Miqueas : Nuevo Sistema
            Call WriteConsoleMsg(UserIndex, tmpStr, FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

    End With

End Sub

Private Sub HandleDameCanje(ByVal UserIndex As Integer)
'***************************************************
'Author: ZheTa Adaptado para Iriwynne AO
'***************************************************

'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte

    Call WriteCanje(UserIndex)
    Call WritePuntos(UserIndex)

End Sub

Public Sub WriteCanje(ByVal UserIndex As Integer)

    Dim i      As Long

    For i = 1 To UBound(Canjes)
        Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCanje(UserIndex, i))
    Next i

End Sub

Public Sub WritePuntos(ByVal UserIndex As Integer)
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePuntos(UserIndex))

End Sub

Public Function PrepareMessageCanje(ByVal UserIndex As Integer, _
                                    ByVal Item As Byte) As String

'***************************************************
'Author: Dami�n
'***************************************************

' @@ Miqueas : Reduccion de Lag

    With auxiliarBuffer

        Call .WriteByte(ServerPacketID.Canje)
        Call .WriteByte(Item)

        Call .WriteInteger(Canjes(Item).objIndex)
        Call .WriteInteger(Canjes(Item).Puntos)
        'Call .WriteInteger(ObjData(Canjes(Item).objIndex).GrhIndex)
        'Call .WriteInteger(ObjData(Canjes(Item).objIndex).MinDef)
        'Call .WriteInteger(ObjData(Canjes(Item).objIndex).MaxDef)
        'Call .WriteInteger(ObjData(Canjes(Item).objIndex).DefensaMagicaMin)
        'Call .WriteInteger(ObjData(Canjes(Item).objIndex).DefensaMagicaMax)
        'Call .WriteInteger(ObjData(Canjes(Item).objIndex).MinHIT)
        'Call .WriteInteger(ObjData(Canjes(Item).objIndex).MaxHIT)

        Call .WriteByte(ObjData(Canjes(Item).objIndex).NoSeCae)

        PrepareMessageCanje = .ReadASCIIStringFixed(.length)

    End With

End Function

Public Function PrepareMessagePuntos(ByVal UserIndex As Integer) As String

'***************************************************
'Author: Barro
'***************************************************
    With auxiliarBuffer

        Call .WriteByte(ServerPacketID.CanjePTS)
        Call .WriteLong(UserList(UserIndex).flags.PuntosShop)
        PrepareMessagePuntos = .ReadASCIIStringFixed(.length)

    End With

End Function

Private Sub HandleSetPointsShop(ByVal UserIndex As Integer)

'***************************************************
'Author: Kevin Amichar
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim UserName As String

        Dim Reason As Integer

        UserName = Buffer.ReadASCIIString()
        Reason = Buffer.ReadInteger()

        'If (Not .flags.Privilegios) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.RoleMaster)) <> 0 Then

        Dim uName As String

        uName = UCase$(.Name)

        Dim VALID As Boolean

        If EsAdmin(.Name) Or EsDios(.Name) Then
            VALID = True
        End If

        If VALID = False Then
            Call .incomingData.CopyBuffer(Buffer)
            Set Buffer = Nothing
            Exit Sub
        End If

        Dim tUser As Integer

        Dim userPriv As Byte

        Dim cantPenas As Byte

        Dim Rank As Integer

        If InStrB(UserName, "+") Then
            UserName = Replace$(UserName, "+", " ")

        End If

        tUser = NameIndex(UserName)

        With UserList(UserIndex)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no est� online.", FontTypeNames.FONTTYPE_TALK)

                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("SERVIDOR> " & .Name & " ha dado " & Reason & " puntos a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))

                    Dim PuntosTotal As Integer

                    PuntosTotal = GetVar(CharPath & UserName & ".chr", "FLAGS", "PuntosShop")
                    PuntosTotal = PuntosTotal + Reason

                    Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "PuntosShop", PuntosTotal)
                Else
                    Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("SERVIDOR> " & .Name & " ha dado " & Reason & " puntos a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                UserList(tUser).flags.PuntosShop = UserList(tUser).flags.PuntosShop + Reason

            End If

        End With

        'End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
       Err.Raise Error

End Sub

Private Sub HandleChangeHead(ByVal UserIndex As Integer)
' @@ Miqueas
' @@ Cambio de Cabeza por oro

    With UserList(UserIndex)
        .incomingData.ReadByte

        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "Est�s Muerto.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If .flags.Mimetizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "Est�s Mimetizando.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(UserIndex, "Est�s Navegando.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        If .Stats.GLD < 500 Then
            Call WriteConsoleMsg(UserIndex, "El cambio de cara cuesta 500 monedas de oro.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If

        .Stats.GLD = .Stats.GLD - 500

        Dim NewHead As Integer

        If DarCabezaNueva(UserIndex, NewHead) Then
            .Char.Head = NewHead
            .OrigChar.Head = NewHead

            With .Char
                Call ChangeUserChar(UserIndex, .Body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

            End With

        Else
            WriteConsoleMsg UserIndex, "No se a podido encontrar una cabeza Valida para tu Raza o Clase", FontTypeNames.FONTTYPE_INFO

        End If

    End With

End Sub

Private Sub HandleRequieredControlUser(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Dim miBuffer As New clsByteQueue

        miBuffer.CopyBuffer .incomingData

        miBuffer.ReadByte

        Dim TargerIndex As Integer
        TargerIndex = NameIndex(miBuffer.ReadASCIIString)

        If esGM(UserIndex) Then

            If TargerIndex > 0 Then
                Call WriteControlUserRevice(TargerIndex)
                UserList(TargerIndex).ControlUserPedido = UserIndex
            Else
                WriteConsoleMsg UserIndex, "User Offline", FontTypeNames.FONTTYPE_INFO

            End If

        End If

        .incomingData.CopyBuffer miBuffer

    End With

End Sub

Private Sub HandleSendDataControlUser(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Dim miBuffer As New clsByteQueue

        miBuffer.CopyBuffer .incomingData

        miBuffer.ReadByte

        Dim ArrayStr As String
        Dim Cant As Byte
        Dim cInterval(1 To 6) As Integer
        Dim LoopC As Long

        ArrayStr = miBuffer.ReadASCIIString()
        Cant = miBuffer.ReadByte()

        For LoopC = 1 To 6
            cInterval(LoopC) = miBuffer.ReadInteger()
        Next LoopC

        'If EsGm(UserIndex) Then

        If .ControlUserPedido > 0 Then
            WriteControlUserShow .ControlUserPedido, ArrayStr, Cant, .Name, cInterval(1), cInterval(2), cInterval(3), cInterval(4), cInterval(5), cInterval(6)

        End If

        'End If

        .incomingData.CopyBuffer miBuffer

    End With

End Sub

Public Sub WriteControlUserShow(ByVal UserIndex As Integer, _
                                ByVal ArrStr As String, _
                                ByVal Cant As Byte, _
                                ByVal sendIndex As String, _
                                ByVal int1 As Integer, _
                                ByVal int2 As Integer, _
                                ByVal int3 As Integer, _
                                ByVal int4 As Integer, _
                                ByVal int5 As Integer, _
                                ByVal int6 As Integer)

    With UserList(UserIndex).outgoingData

        .WriteByte ServerPacketID.ControlUserShow

        .WriteASCIIString sendIndex
        .WriteASCIIString ArrStr

        .WriteByte Cant

        .WriteInteger int1
        .WriteInteger int2

        .WriteInteger int3
        .WriteInteger int4

        .WriteInteger int5
        .WriteInteger int6

    End With

End Sub

Public Sub WriteControlUserRevice(ByVal UserIndex As Integer)

    With UserList(UserIndex).outgoingData

        .WriteByte ServerPacketID.ControlUserRecive

    End With

End Sub

Private Sub handleRequestScreen(ByVal UserIndex As Integer)

' Ulises.-

    With UserList(UserIndex)

        Dim UNname As String
        Dim UIname As Integer

        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        Call Buffer.ReadByte

        UNname = Buffer.ReadASCIIString()

        UIname = NameIndex(UNname)

        If .flags.Privilegios <> PlayerType.User Then
            If UIname <> 0 Then
                Call WriteRequestScreen(UIname)
                Call WriteConsoleMsg(UserIndex, "Se ha solicitado la captura de pantalla.", FontTypeNames.FONTTYPE_TALK)
            Else
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)

            End If

        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

End Sub

Public Sub WriteRequestScreen(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.RequestScreen)

    End With

End Sub

Public Sub HandleResetChar(ByVal UserIndex As Integer)
    With UserList(UserIndex)

        .incomingData.ReadByte

        If Not EsNewbie(UserIndex) Then
            WriteConsoleMsg UserIndex, "Solo podes usar este comando cuando sos newbie", FontTypeNames.FONTTYPE_INFO
            Exit Sub

        End If
        
        If .flags.Comerciando Then Exit Sub
        If .flags.Muerto <> 0 Then Exit Sub
        
        Dim i  As Long

        Call LimpiarInventario(UserIndex)
        Call ResetUserSpells(UserIndex)

        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = 0
        Next i

        .Stats.SkillPts = 10

        Dim MiInt As Long

        'MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)

        .Stats.MaxHP = 21    ' @@ 15 + MiInt
        .Stats.MinHp = 21    ' @@ 15 + MiInt

        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)

        If MiInt = 1 Then MiInt = 2

        .Stats.MaxSta = 20 * MiInt
        .Stats.MinSta = 20 * MiInt

        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100

        .Stats.MaxHam = 100
        .Stats.MinHam = 100

        Dim Userclase As eClass

        Userclase = .clase

        '<-----------------MANA----------------------->
        If Userclase = eClass.Mage Then    'Cambio en mana inicial (ToxicWaste)
            MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
            .Stats.MaxMAN = MiInt
            .Stats.MinMAN = MiInt
        ElseIf Userclase = eClass.Cleric Or Userclase = eClass.Druid _
               Or Userclase = eClass.Bard Or Userclase = eClass.Assasin Then
            .Stats.MaxMAN = 50


            .Stats.MinMAN = 50
        ElseIf Userclase = eClass.Bandit Then    'Mana Inicial del Bandido (ToxicWaste)
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
        Else
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0

        End If

        If Userclase = eClass.Mage Or Userclase = eClass.Cleric Or _
           Userclase = eClass.Druid Or Userclase = eClass.Bard Or _
           Userclase = eClass.Assasin Then
            .Stats.UserHechizos(1) = 2

            If Userclase = eClass.Druid Then .Stats.UserHechizos(2) = 46

        End If

        .Stats.MaxHIT = 2
        .Stats.MinHIT = 1

        .Stats.GLD = 0
        .Stats.Banco = 0
        .Stats.def = 0

        .Stats.Exp = 0
        .Stats.ELU = 300
        .Stats.ELV = 1

        '???????????????? INVENTARIO ��������������������
        Dim Slot As Byte

        Dim IsPaladin As Boolean

        IsPaladin = Userclase = eClass.Paladin

        'Pociones Rojas (Newbie)
        Slot = 1
        .Invent.Object(Slot).objIndex = 857
        .Invent.Object(Slot).Amount = 200

        'Pociones azules (Newbie)
        If .Stats.MaxMAN > 0 Or IsPaladin Then
            Slot = Slot + 1
            .Invent.Object(Slot).objIndex = 856
            .Invent.Object(Slot).Amount = 200

        Else
            'Pociones amarillas (Newbie)
            Slot = Slot + 1




            .Invent.Object(Slot).objIndex = 855
            .Invent.Object(Slot).Amount = 100

            'Pociones verdes (Newbie)
            Slot = Slot + 1
            .Invent.Object(Slot).objIndex = 858
            .Invent.Object(Slot).Amount = 50

        End If

        ' Ropa (Newbie)
        Slot = Slot + 1

        Dim UserRaza As eRaza

        UserRaza = .raza

        Select Case UserRaza

        Case eRaza.Humano
            .Invent.Object(Slot).objIndex = 463

        Case eRaza.Elfo
            .Invent.Object(Slot).objIndex = 464

        Case eRaza.Drow
            .Invent.Object(Slot).objIndex = 465


        Case eRaza.Enano
            .Invent.Object(Slot).objIndex = 466

        Case eRaza.Gnomo
            .Invent.Object(Slot).objIndex = 466

        End Select

        ' Equipo ropa
        .Invent.Object(Slot).Amount = 1
        .Invent.Object(Slot).Equipped = 1

        .Invent.ArmourEqpSlot = Slot
        .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).objIndex

        'Arma (Newbie)
        Slot = Slot + 1

        Select Case Userclase

        Case eClass.Hunter
            ' Arco (Newbie)
            .Invent.Object(Slot).objIndex = 859

        Case eClass.Worker
            ' Herramienta (Newbie)
            .Invent.Object(Slot).objIndex = RandomNumber(561, 565)




        Case Else
            ' Daga (Newbie)
            .Invent.Object(Slot).objIndex = 460

        End Select

        ' Equipo arma
        .Invent.Object(Slot).Amount = 1
        .Invent.Object(Slot).Equipped = 1

        .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).objIndex
        .Invent.WeaponEqpSlot = Slot

        .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)

        ' Municiones (Newbie)
        If Userclase = eClass.Hunter Then
            Slot = Slot + 1
            .Invent.Object(Slot).objIndex = 860
            .Invent.Object(Slot).Amount = 150

            ' Equipo flechas
            .Invent.Object(Slot).Equipped = 1
            .Invent.MunicionEqpSlot = Slot
            .Invent.MunicionEqpObjIndex = 860



        End If

        ' Manzanas (Newbie)
        Slot = Slot + 1
        .Invent.Object(Slot).objIndex = 467
        .Invent.Object(Slot).Amount = 100

        ' Jugos (Nwbie)
        Slot = Slot + 1
        .Invent.Object(Slot).objIndex = 468
        .Invent.Object(Slot).Amount = 100

        ' Sin casco y escudo
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco

        ' Total Items
        .Invent.NroItems = Slot

        .LogOnTime = Now
        .UpTime = 0

        Call WriteSendSkills(UserIndex)
        Call WriteLevelUp(UserIndex, 0, .Stats.SkillPts)
        Call UpdateUserInv(True, UserIndex, 0)
        Call WriteUpdateUserStats(UserIndex)
    End With
End Sub


Private Sub HandleCountdown(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicol�s Javier Amoruso (NaKruL)
'Last Modification: 02/06/10
'03/06/2010: NaKruL - Agregados nuevos checkeos.
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Cuenta As Integer
        Dim Mapa As Integer

        Cuenta = .incomingData.ReadInteger()
        Mapa = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        If Not MapaValido(Mapa) And Mapa <> 0 Then
            Call WriteConsoleMsg(UserIndex, "El mapa no es v�lido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Cuenta = 0 And CountdownTime > 0 Then
            CountdownTime = 0
            CountdownMap = 0
            Call WriteConsoleMsg(UserIndex, "Cuenta regresiva detenida.", FontTypeNames.FONTTYPE_INFO)
        Else
            CountdownTime = Cuenta
            CountdownMap = Mapa
        End If
    End With
End Sub






Public Sub Handleregresar(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Dim REGRESO As Integer
        Dim Valorr As Integer
        Dim MiObj As Obj
        Dim Reclamox As Integer
        REGRESO = .incomingData.ReadByte
        Dim tFile As String
        Dim MiStringes As String

        Dim i  As Long
        Dim Count As Long
        Dim priv As PlayerType
        Dim List As String


        Select Case REGRESO
            '''''''''''''''''''''''' MERCADOAO '''''''''''''''''''''''''
        Case 250
         '   Call PosteoListado(UserIndex)

        Case 251
          '  Call DesposteoListado(UserIndex)

        Case 252
            For i = 1 To LastUser
                If LenB(UserList(i).Name) <> 0 Then
                    If UserList(i).flags.Posteado = 1 Then
                        If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                           Count = Count + 1
                    End If
                End If
            Next i
            For i = 1 To LastUser
                If UserList(i).flags.Posteado = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                       List = List & UserList(i).Name & ", "
                End If
            Next i
            If LenB(List) <> 0 Then
                List = Left$(List, Len(List) - 2)
                Call WriteConsoleMsg(UserIndex, "Usuarios Posteados online: " & List & ". (" & CStr(Count) & ")", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "No hay usuarios posteados.", FontTypeNames.FONTTYPE_INFO)
            End If

        Case 255
            Call WriteMercadoList(UserIndex)
            '''''''''''''''''''''''' MERCADOAO '''''''''''''''''''''''''

        End Select


    End With
End Sub

Public Sub WriteMercadoList(ByVal UserIndex As Integer)
    On Error GoTo errhandleR
    Dim i      As Long
    Dim str    As String

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MercadoList)

        For i = 1 To 50
            If UserList(i).flags.Posteado = 1 Then
                str = str & UserList(i).Name & SEPARATOR
            End If
        Next i

        If LenB(str) > 0 Then _
           str = Left$(str, Len(str) - 1)

        Call .WriteASCIIString(str)
    End With
    Exit Sub

errhandleR:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Private Sub HandleRetos(ByVal UserIndex As Integer)
'intercambioPJ

    Dim UserName As String
    Dim userSend As Integer
    Dim Opcion As Long
    Dim Pin    As String

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(UserIndex)
        Call .incomingData.ReadByte
        UserName = .incomingData.ReadASCIIString
        Opcion = .incomingData.ReadLong
        Pin = .incomingData.ReadASCIIString
        userSend = NameIndex(UserName)


        Dim AuxPin As String
        Dim AuxContra As String
        Dim AuxMail As String

        'miro sus stats
        If Opcion = 0 Then
            Call SendUserStatsTxt(UserIndex, userSend)
            Exit Sub
        End If

        'No existe
        If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
            Call WriteConsoleMsg(UserIndex, "Charfile inexistente.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'No esta posteado
        If UserList(userSend).flags.Posteado = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no esta posteado!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'a mi mismo?
        If UserList(UserIndex).Name = UserList(userSend).Name Then
            Call WriteConsoleMsg(UserIndex, "No puedes enviarte solicitud a ti mismo!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'usuario offline
        If userSend <= 0 Then
            Call WriteConsoleMsg(UserIndex, "Usuario Offline.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'pin incorrecto!
        If UserList(UserIndex).clave <> Pin Then
            Call WriteConsoleMsg(UserIndex, "Tu PIN es incorrecto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'acepto cambio
        If Opcion = 2 Then
            If UserList(UserIndex).flags.MeMando = UserList(userSend).Name And UserList(userSend).flags.Lemande = UserList(UserIndex).Name Then

                Call WriteConsoleMsg(UserIndex, "Felicitaciones! Intercambio Realizado!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(userSend, "Felicitaciones! Intercambio Realizado!", FontTypeNames.FONTTYPE_INFO)

                Dim oldPass1 As String
                Dim oldPass2 As String
                oldPass1 = UCase$(GetVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Password"))
                oldPass2 = UCase$(GetVar(CharPath & UserList(userSend).Name & ".chr", "INIT", "Password"))

                'Almaceno en auxiliar
                AuxPin = UserList(UserIndex).clave
                AuxMail = UserList(UserIndex).email
                AuxContra = oldPass1

                'los dats de 1 ahora son los del 2
                UserList(UserIndex).clave = UserList(userSend).clave
                UserList(UserIndex).email = UserList(userSend).email
                oldPass1 = oldPass2

                UserList(userSend).clave = AuxPin
                UserList(userSend).email = AuxMail
                oldPass2 = AuxContra

                Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Password", oldPass1)
                Call WriteVar(CharPath & UserList(userSend).Name & ".chr", "INIT", "Password", oldPass2)

                UserList(UserIndex).flags.Posteado = 0
                UserList(userSend).flags.Posteado = 0

                Call Cerrar_Usuario(UserIndex)
                Call Cerrar_Usuario(userSend)

            End If
            Exit Sub
        End If

        'alguien mas le mando! espera su mercado
        If UserList(userSend).flags.MeMando <> "" Or UserList(userSend).flags.Lemande <> "" Then
            Call WriteConsoleMsg(UserIndex, "El usuario est� esperando otra invitaci�n de cambio.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'mando cambio asd
        If Opcion = 1 Then
            Call WriteConsoleMsg(UserIndex, "Le has enviado la peticion de intercambio de personaje!", FontTypeNames.FONTTYPE_GUILDMSG)
            UserList(userSend).flags.MeMando = UserList(UserIndex).Name
            UserList(UserIndex).flags.Lemande = UserList(userSend).Name
            Call WriteConsoleMsg(userSend, "El usuario " & UserList(UserIndex).Name & " te ha enviado solicitud de intercambio! Para negar debes cancelar desde el formulario o desloguear.", FontTypeNames.FONTTYPE_GUILDMSG)
            Exit Sub
        End If

    End With
End Sub


Public Sub WriteMontateToggle(ByVal UserIndex As Integer)
    On Error Resume Next
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MontateToggle)
End Sub

Public Sub HandlePoder(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        On Error GoTo errhandleR
        Call .incomingData.ReadByte

        If GranPoder > 0 Then
            Call WriteConsoleMsg(UserIndex, "El gran poder lo posee " & UserList(GranPoder).Name & " en el mapa " & UserList(GranPoder).Pos.Map & "  (" & MapInfo(UserList(GranPoder).Pos.Map).Name & ") .", FontTypeNames.FONTTYPE_INFOBOLD)
        End If
    End With

errhandleR:

End Sub


Public Sub HandleCastillo(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Call .incomingData.ReadByte

        EnviarInfoCastillo UserIndex

    End With

End Sub

Public Sub HandlePideRanking(ByVal UserIndex As Integer)
    With UserList(UserIndex)

        Call .incomingData.ReadByte

        Dim TipoRank As eRanking

        TipoRank = .incomingData.ReadByte

        ' @ Enviamos el ranking
        Call WriteEnviarRanking(UserIndex, TipoRank)

    End With
End Sub


Public Sub WriteEnviarRanking(ByVal UserIndex As Integer, ByVal Rank As eRanking)

    On Error GoTo errhandleR
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.EnviarDatosRanking)

    Dim i      As Integer
    Dim Cadena As String
    Dim Cadena2 As String

    For i = 1 To MAX_TOP
        If i = 1 Then
            Cadena = Cadena & Ranking(Rank).Nombre(i)
            Cadena2 = Cadena2 & Ranking(Rank).Value(i)
        Else
            Cadena = Cadena & "-" & Ranking(Rank).Nombre(i)
            Cadena2 = Cadena2 & "-" & Ranking(Rank).Value(i)
        End If
    Next i

    ' @ Enviamos la cadena
    Call UserList(UserIndex).outgoingData.WriteASCIIString(Cadena)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(Cadena2)

    Exit Sub

errhandleR:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If


End Sub

Public Sub WriteNameList(ByVal UserIndex As Integer)

    On Error GoTo errhandleR

    With UserList(UserIndex).outgoingData

        Call .WriteByte(ServerPacketID.NameList)
        Call .WriteASCIIString(UserList(UserIndex).Name)    '@@ Nick

    End With

    Exit Sub

errhandleR:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleStartList(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Call ListEvent.Start_List(UserIndex, .incomingData.ReadByte())

    End With

End Sub

Private Sub HandleAccessList(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Call ListEvent.Access_List(UserIndex)

    End With

End Sub

Private Sub HandleCancelList(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Call ListEvent.Clean_List

    End With

End Sub

Private Sub HandleActivarDeath(ByVal UserIndex As Integer)

' @ maTih.-
' @ Crea un deathMatch por cupos.

    With UserList(UserIndex)

        Call .incomingData.ReadByte

        Dim Cupos As Byte
        Dim CaenItems As Byte
        Dim NivelMinimo As Byte, NivelMaximo As Byte

        Cupos = .incomingData.ReadByte()
        NivelMinimo = .incomingData.ReadByte
        NivelMaximo = .incomingData.ReadByte

        ' @@ CLASES
        Dim cDruida As Byte, _
            cMago As Byte, _
            cClerigo As Byte, _
            cBardo As Byte, _
            cPaladin As Byte, _
            cAsesino As Byte, _
            cGuerrero As Byte, _
            cCazador As Byte, _
            cBandido As Byte, _
            cLadron As Byte

        cMago = .incomingData.ReadByte    'Option1(0).Value
        cClerigo = .incomingData.ReadByte    '1).Value
        cBardo = .incomingData.ReadByte    '2).Value
        cPaladin = .incomingData.ReadByte    '3).Value
        cAsesino = .incomingData.ReadByte    '4).Value
        cCazador = .incomingData.ReadByte    '5).Value
        cGuerrero = .incomingData.ReadByte    '6).Value
        cDruida = .incomingData.ReadByte    '7).Value
        cLadron = .incomingData.ReadByte    '8).Value
        cBandido = .incomingData.ReadByte    '9).Value



        If Not esGM(UserIndex) Then Exit Sub

        If Not Cupos > 1 Then Exit Sub

        If DeathMatch.Activo Then modDeath.Cancelar (.Name)

        Dim soyDios As Byte, sAdmin As Byte, sSemi As Byte, sConse As Byte
        If EsConsejero(.Name) Then sConse = 1
        If EsSemiDios(.Name) Then sSemi = 1
        If EsDios(.Name) Then soyDios = 1
        If EsAdmin(.Name) Then sAdmin = 1

        ' @@ JAJA
        If sAdmin + soyDios + sSemi + sConse = 0 Then Exit Sub

        Dim TempTick As Long
        TempTick = GetTickCount And &H7FFFFFFF

        If sSemi = 1 Or sConse = 1 Then    'If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            If .Counters.CreaEvento = 0 Or TempTick - .Counters.CreaEvento > 300000 Then
                .Counters.CreaEvento = TempTick
            Else
                WriteConsoleMsg UserIndex, "Hay un limite de 5 minutos para que puedas abrir otro evento.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If
        End If

        Call modDeath.ActivarNuevo(.Name & ".", Cupos, NivelMinimo, NivelMaximo, cMago, cClerigo, cBardo, cPaladin, cAsesino, cCazador, cGuerrero, cDruida, cLadron, cBandido)

    End With

End Sub

Private Sub HandleIngresarDeath(ByVal UserIndex As Integer)

' @ maTih.-
' @ Ingresa al death

    With UserList(UserIndex)

        Call .incomingData.ReadByte

        Dim ErrorMSG As String

        'Puede entrar
        If modDeath.AprobarIngreso(UserIndex, ErrorMSG) Then
            'Lo inscribo.
            Call modDeath.Ingresar(UserIndex)
        Else
            'No puede, aviso.
            Call WriteConsoleMsgNew(UserIndex, "Deathmatch>", ErrorMSG)
        End If

    End With

End Sub

Public Sub HandleArrancaTorneo(UserIndex As Integer)
    On Error Resume Next
    With UserList(UserIndex)

        Dim CaenItems As Byte, Torneos As Byte, NivelMinimo As Byte, NivelMaximo As Byte
        ' @@ CLASES
        Dim cDruida As Byte, _
            cMago As Byte, _
            cClerigo As Byte, _
            cBardo As Byte, _
            cPaladin As Byte, _
            cAsesino As Byte, _
            cGuerrero As Byte, _
            cCazador As Byte, _
            cBandido As Byte, _
            cLadron As Byte

        Call .incomingData.ReadByte
        Torneos = .incomingData.ReadByte()
        NivelMinimo = .incomingData.ReadByte
        NivelMaximo = .incomingData.ReadByte


        cMago = .incomingData.ReadByte    'Option1(0).Value
        cClerigo = .incomingData.ReadByte    '1).Value
        cBardo = .incomingData.ReadByte    '2).Value
        cPaladin = .incomingData.ReadByte    '3).Value
        cAsesino = .incomingData.ReadByte    '4).Value
        cCazador = .incomingData.ReadByte    '5).Value
        cGuerrero = .incomingData.ReadByte    '6).Value
        cDruida = .incomingData.ReadByte    '7).Value
        cLadron = .incomingData.ReadByte    '8).Value
        cBandido = .incomingData.ReadByte    '9).Value

        Dim soyDios As Byte, sAdmin As Byte, sSemi As Byte, sConse As Byte
        If EsConsejero(.Name) Then sConse = 1
        If EsSemiDios(.Name) Then sSemi = 1
        If EsDios(.Name) Then soyDios = 1
        If EsAdmin(.Name) Then sAdmin = 1

        ' @@ JAJA
        If sAdmin + soyDios + sSemi + sConse = 0 Then Exit Sub

        Dim TempTick As Long
        TempTick = GetTickCount And &H7FFFFFFF

        If sSemi = 1 Or sConse = 1 Then    'If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            If .Counters.CreaEvento = 0 Or TempTick - .Counters.CreaEvento > 300000 Then
                .Counters.CreaEvento = TempTick
            Else
                WriteConsoleMsg UserIndex, "Hay un limite de 5 minutos para que puedas abrir otro evento.", FontTypeNames.FONTTYPE_INFO
                Exit Sub
            End If
        End If

        If (Torneos > 0 And Torneos < 6) Then
            Call Torneos_Inicia(UserIndex, Torneos, NivelMinimo, NivelMaximo, cMago, cClerigo, cBardo, cPaladin, cAsesino, cCazador, cGuerrero, cDruida, cLadron, cBandido)
        End If

        Call LogGM(.Name, .Name & "ha arrancado un torneo.")
    End With

End Sub

Public Sub HandleCancelaTorneo(UserIndex As Integer)

On Error GoTo errhandleR
    With UserList(UserIndex)
1        Call .incomingData.ReadByte
2        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
3        Call LogGM(.Name, .Name & "ha cancelado un torneo.")
4        Call Rondas_Cancela
5    End With
    Exit Sub
    
errhandleR:
    LogError "CancelaTorneo LINEA: " & Erl & " - Err " & Err.Number & " " & Err.description
End Sub

Public Sub HandleParticipar(UserIndex As Integer)

    With UserList(UserIndex)
        Call .incomingData.ReadByte
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "�Est�s muerto!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call Ingresar1vs1(UserIndex)
    End With

End Sub

Private Sub HandleSetVip(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)

        Call Buffer.ReadByte

        Dim UIname As Integer
        Dim UNname As String
        Dim strMsg As String
        UNname = Buffer.ReadASCIIString()
        Dim QuitoVip As Boolean, Dias As Integer
        QuitoVip = Buffer.ReadBoolean

        Dias = Buffer.ReadInteger

        Dim UserCharPath As String
        
        If InStrB(UNname, "+") Then
                UNname = Replace(UNname, "+", " ")
        End If
        
        UserCharPath = CharPath & UNname & ".chr"
        
        UIname = NameIndex(UNname)
        If .flags.Privilegios = PlayerType.Admin Then
            If QuitoVip = False Then
                If Dias >= 1 Then
                    If UIname > 0 Then
                        UserList(UIname).flags.Vip = 1
                        WarpUserChar UIname, UserList(UIname).Pos.Map, UserList(UIname).Pos.X, UserList(UIname).Pos.Y, False, False
                        WriteConsoleMsgNew UIname, "VIP>", "Felicidades!! Ahora eres un usuario Vip por " & Dias & " d�as.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO
                        Debug.Print Date + Dias
                        
                        Call WriteVar(UserCharPath, "INIT", "VIP_DIAS", Date + Dias)
                        Call WriteConsoleMsgNew(UserIndex, "VIP>", "Le has agregado VIP a " & UNname & " por " & Dias & " d�as.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO)
                    ElseIf UIname <= 0 And FileExist(UserCharPath) Then
                        Call WriteConsoleMsgNew(UserIndex, "VIP>", "Le has agregado VIP a " & UNname & " por " & Dias & " d�as. (Offline)", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO)
                        Call WriteVar(UserCharPath, "INIT", "VIP_DIAS", Date + Dias)
                    Else
                        Call WriteConsoleMsgNew(UserIndex, "VIP>", "Est�s intentando dar VIP un usuario inexistente.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO)
                    End If
                End If
            Else
                ' Quito el vip.
                If UIname > 0 Then
                    UserList(UIname).flags.Vip = 0
                    WarpUserChar UIname, UserList(UIname).Pos.Map, UserList(UIname).Pos.X, UserList(UIname).Pos.Y, False, False
                    WriteConsoleMsgNew UIname, "VIP>", "Se te ha removido el VIP", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO
                    Call WriteVar(UserCharPath, "INIT", "VIP_DIAS", Date)
                    Call WriteConsoleMsgNew(UserIndex, "VIP", "Le has quitado VIP a " & UNname, FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO)
                ElseIf UIname <= 0 And FileExist(UserCharPath) Then
                    Call WriteConsoleMsgNew(UserIndex, "VIP>", "Le has quitado VIP a " & UNname, FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO)
                    Call WriteVar(UserCharPath, "INIT", "VIP_DIAS", Date)
                Else
                    Call WriteConsoleMsgNew(UserIndex, "VIP>", "Est�s intentando quitar el VIP un usuario inexistente.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO)
                End If
            End If
        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

End Sub


Public Function PrepareMessageConsoleMsgNew(ByVal Cabeza As String, ByVal Cuerpo As String, Optional ByVal CabezaFonttype As FontTypeNames = FONTTYPE_VERDE, Optional ByVal CuerpoFonttype As FontTypeNames = FONTTYPE_BLANCO) As String

' @@ CuiCui

    With auxiliarBuffer
        .WriteByte ServerPacketID.ConsoleMsgNew
        .WriteASCIIString Cabeza
        .WriteASCIIString Cuerpo
        .WriteByte CabezaFonttype
        .WriteByte CuerpoFonttype

        PrepareMessageConsoleMsgNew = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteConsoleMsgNew(ByVal UserIndex As Integer, ByVal Cabeza As String, ByVal Chat As String, Optional ByVal CabezaFonttype As FontTypeNames = FONTTYPE_VERDE, Optional ByVal CuerpoFonttype As FontTypeNames = FONTTYPE_BLANCO)

' @@ CuiCui

    On Error GoTo errhandleR
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsgNew(Cabeza, Chat, CabezaFonttype, CuerpoFonttype))
    Exit Sub

errhandleR:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub



''
' Handles the "VerHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleVerHD(ByVal UserIndex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification:  08/06/2012 - ^[GS]^
'Verifica el HD del usuario.
'***************************************************

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandleR
    With UserList(UserIndex)
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim iUsuario As Integer

        iUsuario = NameIndex(Buffer.ReadASCIIString())

        If iUsuario = 0 Then
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje321)    '"El personaje no est� online."
        Else
            Call WriteConsoleMsg(UserIndex, "El usuario " & UserList(iUsuario).Name & " tiene un disco con el Serial " & UserList(iUsuario).discoDuro, FONTTYPE_INFOBOLD)
        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "UnBanHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUnbanHD(ByVal UserIndex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification:  08/06/2012 - ^[GS]^
'Maneja el unbaneo del serial del HD de un usuario.
'***************************************************

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandleR
    With UserList(UserIndex)
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim SerialHD As String
        SerialHD = Buffer.ReadASCIIString()

        If (BanHD_rem(SerialHD)) Then
            Call WriteConsoleMsg(UserIndex, "El disco con el Serial " & SerialHD & " se ha quitado de la lista de baneados.", FONTTYPE_INFOBOLD)
        Else
            Call WriteConsoleMsg(UserIndex, "El disco con el Serial " & SerialHD & " no se encuentra en la lista de baneados.", FONTTYPE_INFO)
        End If

        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BanHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleBanHD(ByVal UserIndex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification: 08/06/2012 - ^[GS]^
'Maneja el baneo del serial del HD de un usuario.
'***************************************************

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandleR
    With UserList(UserIndex)
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim i  As Long
        Dim iUsuario As Integer
        Dim bannedHD As String

        iUsuario = NameIndex(Buffer.ReadASCIIString())

        If iUsuario > 0 Then
            bannedHD = UserList(iUsuario).discoDuro
        End If

        If .flags.Privilegios And (PlayerType.Admin And PlayerType.Dios) Then
            If LenB(bannedHD) > 0 Then
                If BanHD_find(bannedHD) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanHD_add(bannedHD)
                    Call WriteConsoleMsg(UserIndex, "Has baneado el disco duro " & bannedHD & " del usuario " & UserList(iUsuario).Name, FontTypeNames.FONTTYPE_INFO)
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).discoDuro = bannedHD Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "Ban de serial de disco duro.")
                            End If
                        End If
                    Next i
                End If
            ElseIf iUsuario <= 0 Then
                '  Call WriteMensajes(UserIndex, eMensajes.Mensaje321)    '"El personaje no est� online."
            End If
        End If

        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandleR:
    Dim Error  As Long
    Error = Err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error
End Sub


Private Sub HandleCheckCPU_ID(ByVal UserIndex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Verifica el CPU_ID del usuario.
'***************************************************

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        Call Buffer.ReadByte

        Dim Usuario As Integer
        Dim nickUsuario As String

        nickUsuario = Buffer.ReadASCIIString()
        Usuario = NameIndex(nickUsuario)

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

            If Usuario <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FONTTYPE_INFO)
            Else
                Call LogGM(.Name, "Checkeo de CPU ID a: " & UserList(Usuario).Name)
                Call WriteConsoleMsg(UserIndex, "El CPU_ID del user " & UserList(Usuario).Name & " es " & UserList(Usuario).CPU_ID, FONTTYPE_INFOBOLD)
            End If

        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

    Exit Sub

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    LogError "HandleCheckCPU_IDError"

    On Error GoTo 0

    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "UnBanT0" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUnbanT0(ByVal UserIndex As Integer)
'***************************************************
'Author: CUICUI
'Last Modification: 16/03/17
'Maneja el unbaneo T0 de un usuario.
'***************************************************

    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        Call Buffer.ReadByte

        Dim CPU_ID As String
        CPU_ID = Buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            If (RemoverRegistroT0(CPU_ID)) Then
                Call WriteConsoleMsg(UserIndex, "El T0 n�" & CPU_ID & " se ha quitado de la lista de baneados.", FONTTYPE_INFOBOLD)
                Call LogGM(.Name, "Unbane� T0: " & CPU_ID)
            Else
                Call WriteConsoleMsg(UserIndex, "El T0 n�" & CPU_ID & " no se encuentra en la lista de baneados.", FONTTYPE_INFO)
            End If
        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BanT0" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleBanT0(ByVal UserIndex As Integer)
'***************************************************
'Author: CUICUI
'Last Modification: 16/03/17
'Maneja el baneo T0 de un usuario.
'***************************************************

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        Call Buffer.ReadByte

        Dim Usuario As Integer
        Usuario = NameIndex(Buffer.ReadASCIIString())

        Dim bannedT0 As String
        Dim bannedHD As String

        If Usuario > 0 Then
            bannedT0 = UserList(Usuario).CPU_ID
            bannedHD = UserList(Usuario).discoDuro
        End If

        Dim i  As Long

        If .flags.Privilegios And (PlayerType.Admin) Then    ' @@ SOLO ADMIN

            If LenB(bannedT0) > 0 Then

                If (BuscarRegistroT0(bannedT0) > 0) Then
                    Call WriteConsoleMsg(UserIndex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                Else

                    ' @@ Agregamos al registro el ID �nico
                    Call AgregarRegistroT0(bannedT0)
                    Call LogGM(UserList(UserIndex).Name, "BAN T0 a " & UserList(Usuario).Name)

                    Call WriteConsoleMsg(UserIndex, "Has baneado T0 a " & UserList(Usuario).Name, FontTypeNames.FONTTYPE_INFO)

                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).CPU_ID = bannedT0 Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "Ban T0.")
                            End If
                        End If
                    Next i
                End If

            ElseIf Usuario <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no se encuentra online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With

    Exit Sub

errhandleR:

    Dim Error  As Long

    Error = Err.Number

    On Error GoTo 0

    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Public Sub WriteChauClienteEditado(ByVal UserIndex As Integer)
    On Error Resume Next
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.AntiClienteEditado)
    Call UserList(UserIndex).outgoingData.WriteInteger("-1")

End Sub

Public Sub WriteInviSegundos(ByVal UserIndex As Integer, ByVal segundos As Integer)
    On Error Resume Next
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CuentaInvi)
    Call UserList(UserIndex).outgoingData.WriteInteger(segundos)

End Sub


Public Function PrepareMessageMandarOnlines()
' @@ CuiCui

    With auxiliarBuffer
        .WriteByte ServerPacketID.MandoOnlines

        Dim i  As Long
        Dim Count As Long

        For i = 1 To LastUser
            If LenB(UserList(i).Name) <> 0 Then
                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                   Count = Count + 1
            End If
        Next i

        .WriteInteger Count

        PrepareMessageMandarOnlines = .ReadASCIIStringFixed(.length)
    
    End With

End Function

Public Sub WriteMandarOnlines(ByVal UserIndex As Integer)

' @@ CuiCui

    On Error GoTo errhandleR
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageMandarOnlines)
    Exit Sub

errhandleR:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleFianza(ByVal UserIndex As Integer)

    On Error GoTo errhandleR
    
    With UserList(UserIndex)

        Dim Cantidad As Long, esCriminal As Boolean

        Call .incomingData.ReadByte

        Cantidad = .incomingData.ReadLong

        esCriminal = criminal(UserIndex)

        If Cantidad < 0 Then Cantidad = 1
        If Cantidad > 1000000 Then Cantidad = 1000000

        If .Pos.Map <> 1 Then
            Call WriteConsoleMsg(UserIndex, "Debes estar en Ullathorpe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .Stats.GLD < Cantidad Then
            Call WriteConsoleMsg(UserIndex, "No tienes el oro suficiente para realizar la fianza, para hacerlo te faltan " & .Stats.GLD - Cantidad & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Comerciando Then Exit Sub

        'If esCriminal = False Then    'si no es criminal no lo usa, por que es al pedo,
         '   Call WriteConsoleMsg(UserIndex, "Eres ciudadano, no puedes usar este comando.", FontTypeNames.FONTTYPE_INFO)
        '    Exit Sub
        'End If

        If (MapInfo(.Pos.Map).Pk = True) Then     'si o si en zonas seguras para evitar que se hagan ciudas por las plagas
            Call WriteConsoleMsg(UserIndex, "solo puedes usar este sistema en zonas seguras", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Muerto = 1 Then    'si o si ten�s que estar vivo
            Call WriteConsoleMsg(UserIndex, "estas muerto debes estar vivo para usar la fianza", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        .Reputacion.NobleRep = .Reputacion.NobleRep + Cantidad     'si cumple con todo, lo hago ciudadano.

        .Stats.GLD = .Stats.GLD - Cantidad    'le descuento el oro (ac� cambian el precio..)

        WriteUpdateGold UserIndex

        RefreshCharStatus UserIndex

        LogDesarrollo .Name & " pag� " & Cantidad & " puntos de fianza."

        Call WriteConsoleMsg(UserIndex, "Has pagado " & Cantidad & " monedas de oro y se te han acreditado " & Cantidad & " puntos de nobleza.", FontTypeNames.FONTTYPE_INFO)

    End With
    
    Exit Sub
    
errhandleR:
    LogError "Error en HandleFianza. " & Err.Number & " " & Err.description
End Sub

Public Sub WriteReCheckCpuID(ByVal UserIndex As Integer)
    On Error Resume Next
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.PedirCPUID)
End Sub

Private Sub HandleRecibiCpuID(ByVal UserIndex As Integer)
'***************************************************
'Author: CUICUI
'Last Modification: 12/07/18
'Anti loquitos
'***************************************************

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandleR

    With UserList(UserIndex)

        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        Call Buffer.ReadByte

        Dim Current_CPU As String

        Current_CPU = Buffer.ReadASCIIString()
        
        If Len(Current_CPU) <= 0 Then
            LogError "ERROR: Len(Current_CPU)<=0 de " & .Name
        Else
            If Current_CPU <> .CPU_ID Then
                Call LogHackAttemp(.Name & " tiene un CPU_ID distinto: Logue� con: " & .CPU_ID & " y el nuevo es: " & Current_CPU)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Seguridad de Cui> " & .Name & " tiene una incoherencia. (Logue� con: " & .CPU_ID & " y al pedirle devuelta devolvi�: " & Current_CPU & ").", FontTypeNames.FONTTYPE_SERVER))
            End If
        End If
        
        Call .incomingData.CopyBuffer(Buffer)

    End With

    Exit Sub

errhandleR:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0

    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Public Sub HandleJDHCrear(UserIndex As Integer)

On Error Resume Next

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte

        Dim CantPlayers As Byte
        Dim Premio As Long
        Dim Inscripcion As Long

        CantPlayers = .incomingData.ReadByte
'        Premio = .incomingData.ReadLong()
 '       Inscripcion = .incomingData.ReadLong()
        Inscripcion = 50000
        Premio = 150000
        
        If Not esGM(UserIndex) Then Exit Sub

1        If JDH.Activo = False Then
2            Call m_JuegosDelHambre.ActivarNuevoJDH(UserIndex, CantPlayers, Premio, Inscripcion)
3        Else
4            WriteConsoleMsg UserIndex, "Ya hay un JDH.", FontTypeNames.FONTTYPE_AMARILLO
5        End If

    End With

Exit Sub
LogError "error en HandleJDHCrear en linea " & Erl & ". Err " & Err.Number
End Sub

Public Sub HandleJDHEntrar(UserIndex As Integer)

    With UserList(UserIndex)
        Call .incomingData.ReadByte

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "�Est�s muerto!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Dim MsgErr As String

250     If m_JuegosDelHambre.AprobarIngresoJDH(UserIndex, MsgErr) Then
260         Call m_JuegosDelHambre.IngresarJDH(UserIndex)
        Else
            WriteConsoleMsg UserIndex, MsgErr, FontTypeNames.FONTTYPE_INFO
270     End If

    End With

End Sub

Public Sub HandleJDHCancelar(UserIndex As Integer)

    With UserList(UserIndex)
        Call .incomingData.ReadByte

        If Not esGM(UserIndex) Then
            Exit Sub
        End If

        If JDH.Activo Then
150             Call m_JuegosDelHambre.CancelarJDH
160             Call LogGM(.Name, .Name & " ha cancelado un Juegos del Hambre.")
        Else
            WriteConsoleMsg UserIndex, "No hay un juego del hambre abierto", FontTypeNames.FONTTYPE_INFO
        End If

    End With

End Sub



Public Sub WriteQuestDetails(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)    ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Env�a el paquete QuestDetails y la informaci�n correspondiente.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer

    On Error GoTo errhandleR
    With UserList(UserIndex).outgoingData
        'ID del paquete
        Call .WriteByte(ServerPacketID.QuestDetails)

        'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se acept� todav�a (1 para el primer caso y 0 para el segundo)
        Call .WriteByte(IIf(QuestSlot, 1, 0))

        'Enviamos nombre, descripci�n y nivel requerido de la quest
        Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
        Call .WriteASCIIString(QuestList(QuestIndex).Desc)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)

        'Enviamos la cantidad de npcs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)
        If QuestList(QuestIndex).RequiredNPCs Then
            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
                Call .WriteASCIIString(GetVar(DatPath & "NPCs.dat", "NPC" & QuestList(QuestIndex).RequiredNPC(i).NpcIndex, "Name"))
                'Si es una quest ya empezada, entonces mandamos los NPCs que mat�.
                If QuestSlot Then
                    Call .WriteInteger(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
                End If
            Next i
        End If

        'Enviamos la cantidad de objs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)
        If QuestList(QuestIndex).RequiredOBJs Then
            'Si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredOBJs
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RequiredOBJ(i).objIndex).Name)
            Next i
        End If

        'Enviamos la recompensa de oro y experiencia.
        Call .WriteLong(QuestList(QuestIndex).RewardGLD)
        Call .WriteLong(QuestList(QuestIndex).RewardEXP)
        Call .WriteLong(QuestList(QuestIndex).RewardPoints)

        'Enviamos la cantidad de objs de recompensa
        Call .WriteByte(QuestList(QuestIndex).RewardOBJs)
        If QuestList(QuestIndex).RequiredOBJs Then
            'si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RewardOBJs
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RewardOBJ(i).objIndex).Name)
            Next i
        End If
    End With
    Exit Sub

errhandleR:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteQuestListSend(ByVal UserIndex As Integer)    ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Env�a el paquete QuestList y la informaci�n correspondiente.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
    Dim tmpStr As String
    Dim tmpByte As Byte

    On Error GoTo errhandleR

    With UserList(UserIndex)
        .outgoingData.WriteByte ServerPacketID.QuestListSend

        For i = 1 To MAXUSERQUESTS
            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).Nombre & "-"
            End If
        Next i

        'Escribimos la cantidad de quests
        Call .outgoingData.WriteByte(tmpByte)

        'Escribimos la lista de quests (sacamos el �ltimo caracter)
        If tmpByte Then
            Call .outgoingData.WriteASCIIString(Left$(tmpStr, Len(tmpStr) - 1))
        End If
    End With
    Exit Sub

errhandleR:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

