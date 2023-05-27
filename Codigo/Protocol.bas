Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
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
'The binary prtocol here used was designed by Juan Martin Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

#If False Then

    Dim Map, X, Y, n, Mapa, race, helmet, weapon, shield, Color, Value, errHandler, punishments, Length, obj, index As Variant

#End If

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer  As clsByteQueue

Private Enum ServerPacketID
    Logged = 1                  ' LOGGED
    RemoveDialogs               ' QTDL
    RemoveCharDialog            ' QDL
    NavigateToggle              ' NAVEG
    Disconnect                  ' FINOK
    CommerceEnd                 ' FINCOMOK
    BankEnd                     ' FINBANOK
    CommerceInit                ' INITCOM
    BankInit                    ' INITBANCO
    CommerceChat
    UpdateSta                   ' ASS
    UpdateMana                  ' ASM
    UpdateHP                    ' ASH
    UpdateGold                  ' ASG
    UpdateExp                   ' ASE
    ChangeMap                   ' CM
    PosUpdate                   ' PU
    ChatOverHead                ' ||
    ConsoleMsg                  ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat                   ' |+
    ShowMessageBox              ' !!
    UserIndexInServer           ' IU
    UserCharIndexInServer       ' IP
    CharacterCreate             ' CC
    CharacterRemove             ' BP
    CharacterChangeNick
    CharacterMove               ' MP, +, * and _ '
    ForceCharMove
    CharacterChange             ' CP
    HeadingChange
    ObjectCreate                ' HO
    ObjectDelete                ' BO
    BlockPosition               ' BQ
    UserCommerceInit            ' INITCOMUSU
    UserCommerceEnd             ' FINCOMUSUOK
    ShowTrabajoForm
    UserOfferConfirm
    PlayMusic                    ' TM
    PlayWave                     ' TW
    guildList                    ' GL
    AreaChanged                  ' CA
    PauseToggle                  ' BKW
    ActualizarClima
    CreateFX                     ' CFX
    UpdateUserStats              ' EST
    ChangeInventorySlot          ' CSI
    ChangeBankSlot               ' SBO
    ChangeSpellSlot              ' SHS
    atributes                    ' ATR
    BlacksmithWeapons            ' LAH
    BlacksmithArmors             ' LAR
    CarpenterObjects             ' OBR
    AlquimiaObjects
    SastreObjects
    RestOK                       ' DOK
    errorMsg                     ' ERR
    Blind                        ' CEGU
    Dumb                         ' DUMB
    ShowSignal                   ' MCAR
    ChangeNPCInventorySlot       ' NPCI
    UpdateHungerAndThirst        ' EHYS
    Fame                         ' FAMA
    Family
    MiniStats                    ' MEST
    LevelUp                      ' SUNI
    AddForumMsg                  ' FMSG
    ShowForumForm                ' MFOR
    SetInvisible                 ' NOVER
    MeditateToggle               ' MEDOK
    BlindNoMore                  ' NSEGUE
    DumbNoMore                   ' NESTUP
    SendSkills                   ' SKILLS
    TrainerCreatureList          ' LSTCRI
    guildNews                    ' GUILDNE
    OfferDetails                 ' PEACEDE & ALLIEDE
    AlianceProposalsList         ' ALLIEPR
    PeaceProposalsList           ' PEACEPR
    CharacterInfo                ' CHRINFO
    GuildLeaderInfo              ' LEADERI
    GuildMemberInfo
    GuildDetails                 ' CLANDET
    ShowGuildFundationForm       ' SHOWFUN
    ParalizeOK                   ' PARADOK
    ShowUserRequest              ' PETICIO
    ChangeUserTradeSlot          ' COMUSUINV
    Pong
    UpdateTagAndStatus
    
    'GM =  messages
    SpawnList                    ' SPL
    ShowSOSForm                  ' MSOS
    ShowMOTDEditionForm          ' ZMOTD
    ShowGMPanelForm              ' ABPANEL
    UserNameList                 ' LISTUSU
    ShowDenounces
    RecordList
    RecordDetails
    
    ShowGuildAlign
    ShowPartyForm
    PeticionInvitarParty
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    MultiMessage
    StopWorking
    CancelOfferItem
    FXtoMap
    EnviarPJUserAccount
    SearchList
    UserInEvent
    RenderMsg
    DeletedChar
    EquitandoToggle
    SeeInProcess
    ShowProcess
    CharParticle
    MapParticle
    IniciarSubastaConsulta
    SetSpeed
    AbriMapa
    AbrirGoliath
    OfrecerFamiliar
    EnviarRanking
End Enum

Private Enum ClientPacketID
    LoginExistingChar = 1           'OLOGIN
    LoginNewChar                    'NLOGIN
    Talk                            ';
    Yell                            '-
    Whisper                         '\
    Walk                            'M
    UseItem                         'USA
    RequestPositionUpdate           'RPU
    Attack                          'AT
    PickUp                          'AG
    SafeToggle                      '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    CombatSafeToggle
    RequestGuildLeaderInfo          'GLINFO
    RequestAtributes                'ATR
    RequestFame                     'FAMA
    RequestFamily
    RequestSkills                   'ESKI
    RequestMiniStats                'FEST
    CommerceEnd                     'FINCOM
    UserCommerceEnd                 'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                         'FINBAN
    UserCommerceOk                  'COMUSUOK
    UserCommerceReject              'COMUSUNO
    Drop                            'TI
    CastSpell                       'LH
    LeftClick                       'LC
    AccionClick                     'RC
    Work                            'UK
    UseSpellMacro                   'UMH
    CraftearItem
    WorkClose
    WorkLeftClick                   'WLC
    InvitarPartyClick
    CreateNewGuild                  'CIG
    SpellInfo                      'INFS
    EquipItem                      'EQUI
    ChangeHeading                  'CHEA
    ModifySkills                   'SKSE
    Train                          'ENTR
    CommerceBuy                    'COMP
    BankExtractItem                'RETI
    CommerceSell                   'VEND
    BankDeposit                    'DEPO
    ForumPost                      'DEMSG
    MoveSpell                      'DESPHE
    MoveBank
    ClanCodexUpdate               'DESCOD
    UserCommerceOffer             'OFRECER
    GuildAcceptPeace              'ACEPPEAT
    GuildRejectAlliance           'RECPALIA
    GuildRejectPeace              'RECPPEAT
    GuildAcceptAlliance           'ACEPALIA
    GuildOfferPeace               'PEACEOFF
    GuildOfferAlliance            'ALLIEOFF
    GuildAllianceDetails          'ALLIEDET
    GuildPeaceDetails             'PEACEDET
    GuildRequestJoinerInfo        'ENVCOMEN
    GuildAlliancePropList         'ENVALPRO
    GuildPeacePropList            'ENVPROPP
    GuildDeclareWar               'DECGUERR
    GuildNewWebsite               'NEWWEBSI
    GuildAcceptNewMember          'ACEPTARI
    GuildRejectNewMember          'RECHAZAR
    GuildKickMember               'ECHARCLA
    GuildUpdateNews               'ACTGNEWS
    GuildMemberInfo               '1HRINFO<
    GuildOpenElections            'ABREELEC
    GuildRequestMembership        'SOLICITUD
    GuildRequestDetails           'CLANDETAILS
    Online                        '/ONLINE
    Quit                          '/SALIR
    GuildLeave                    '/SALIRCLAN
    RequestAccountState           '/BALANCE
    PetStand                      '/QUIETO
    PetFollow                     '/ACOMPANAR
    ReleasePet                    '/LIBERAR
    TrainList                     '/ENTRENAR
    Rest                          '/DESCANSAR
    Meditate                      '/MEDITAR
    Resucitate                    '/RESUCITAR
    Heal                          '/CURAR
    Help                          '/AYUDA
    RequestStats                  '/EST
    CommerceStart                 '/COMERCIAR
    BankStart
    GoliathStart                  '/BOVEDA
    Enlist                        '/ENLISTAR
    Information                   '/INFORMACION
    Reward                        '/RECOMPENSA
    RequestMOTD                   '/MOTD
    UpTime                        '/UPTIME
    PartyLeave                    '/SALIRPARTY
    Inquiry                       '/ENCUESTA ( with no params )
    GuildMessage                  '/CMSG
    PartyMessage                  '/PMSG
    GuildOnline                   '/ONLINECLAN
    PartyOnline                   '/ONLINEPARTY
    CouncilMessage                '/BMSG
    RoleMasterRequest             '/ROL
    GMRequest                     '/GM
    ChangeDescription             '/DESC
    GuildVote                     '/VOTO
    punishments                   '/PENAS
    Gamble                        '/APOSTAR
    InquiryVote                   '/ENCUESTA ( with parameters )
    LeaveFaction                  '/RETIRAR ( with no arguments )
    BankExtractGold               '/RETIRAR ( with arguments )
    BankDepositGold               '/DEPOSITAR
    Denounce                      '/DENUNCIAR
    GuildFundate                  '/FUNDARCLAN
    GuildFundation
    PartyKick                     '/ECHARPARTY
    PartyAcceptMember             '/ACCEPTPARTY
    Ping                          '/PING
    RequestPartyForm
    Home
    ShowGuildNews
    ShareNpc                      '/COMPARTIR
    StopSharingNpc
    Consultation
    moveItem
    LoginExistingAccount      'CHOTS | Accounts
    CentinelReport
    FightSend
    FightAccept
    CloseGuild
    DeleteChar
    ChatGlobal
    LookProcess
    SendProcessList
    AccionInventario
    IniciarSubasta 'Iniciamos una subasta
    CancelarSubasta
    OfertarSubasta 'Ofertamos en la subasta
    ConsultaSubasta 'Si existe una subasta enviamos la Info, sino Abrimos el panel para iniciar una subasta.
    GMCommands
    Casamiento
    Acepto
    Divorcio
    TransferenciaOro
    AdoptarFamiliar
    SolicitarRank
    BatallaPVP
End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 153

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
    FONTTYPE_CRIMINAL
    FONTTYPE_EXP
    FONTTYPE_PRIVADO
    
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
    eo_Speed

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

Public Function HandleIncomingData(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    '
    '***************************************************
    On Error Resume Next
    
    With UserList(UserIndex)
        
        'Contamos cuantos paquetes recibimos.
        .Counters.PacketsTick = .Counters.PacketsTick + 1
        
        'Comento esto por ahora, por que cuando hago worldsave, envia mas paquetes en 40ms
        'y desconecta al pj, hay que reveer que hacer con esto y como solucionarlo.

        'Si recibis 10 paquetes en 40ms (intervalo del GameTimer), cierro la conexion.
        'If .Counters.PacketsTick > 10 Then
        '    Call CloseSocket(Userindex)
        '    Exit Function

        'End If

        'Se castea a long por que VB6 cuando usa SELECT CASE
        'Lo hace de manera mas efectiva https://www.gs-zone.org/temas/las-consecuencias-de-usar-byte-en-handleincomingdata.99245/
        Dim PacketId As Long: PacketId = CLng(.incomingData.PeekByte())

        'Verifico si el paquete necesita que el user este logeado
        If Not (PacketId = ClientPacketID.LoginExistingChar _
                Or PacketId = ClientPacketID.LoginNewChar _
                Or PacketId = ClientPacketID.LoginExistingAccount _
                Or PacketId = ClientPacketID.DeleteChar) Then
             
            'Vierifico si el user esta logeado
            If Not .flags.AccountLogged Then
                Call CloseSocket(UserIndex)
                Exit Function
            ElseIf Not .flags.UserLogged Then
                Call Cerrar_Usuario(UserIndex)
                Exit Function
                'El usuario ya logueo. Reseteamos el tiempo AFK si el ID es valido.
            ElseIf PacketId <= LAST_CLIENT_PACKET_ID Then
                .Counters.IdleCount = 0
    
            End If
    
        ElseIf PacketId <= LAST_CLIENT_PACKET_ID Then
            .Counters.IdleCount = 0
            
            'Vierifico si el user esta logeado
            If .flags.UserLogged Then
                Call CloseSocket(UserIndex)
                Exit Function
    
            End If
    
        End If
        
        ' Ante cualquier paquete, pierde la proteccion de ser atacado.
        .flags.NoPuedeSerAtacado = False
        
    End With
    
    Select Case PacketId
        
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(UserIndex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(UserIndex)

        Case ClientPacketID.DeleteChar
            Call HandleDeleteChar(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
            
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.CombatSafeToggle
            Call HandleCombatToggle(UserIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(UserIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(UserIndex)
        
        Case ClientPacketID.RequestFame             'FAMA
            Call HandleRequestFame(UserIndex)
            
        Case ClientPacketID.RequestFamily
            Call HandleRequestFamily(UserIndex)
        
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
        
        Case ClientPacketID.AccionClick             'RC
            Call HandleAccionClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.CraftearItem
            Call HandleCraftearItem(UserIndex)
            
        Case ClientPacketID.WorkClose
            Call HandleWorkClose(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
            
        Case ClientPacketID.InvitarPartyClick
            Call HandleInvitarPartyClick(UserIndex)
        
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
        
        Case ClientPacketID.PetFollow               '/ACOMPANAR
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
        
        Case ClientPacketID.BankStart
            Call HandleBankStart(UserIndex)
            
        Case ClientPacketID.GoliathStart            '/BOVEDA
            Call HandleGoliathStart(UserIndex)
        
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
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(UserIndex)
        
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
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(UserIndex)
        
        Case ClientPacketID.punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
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
        
        Case ClientPacketID.PartyAcceptMember       '/ACCEPTPARTY
            Call HandlePartyAcceptMember(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
            
        Case ClientPacketID.RequestPartyForm
            Call HandlePartyForm(UserIndex)
        
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

        Case ClientPacketID.LoginExistingAccount
            Call HandleLoginExistingAccount(UserIndex)
        
        Case ClientPacketID.CentinelReport
            Call HandleCentinelReport(UserIndex)
  
        Case ClientPacketID.FightSend
            Call HandleFightSend(UserIndex)
            
        Case ClientPacketID.FightAccept
            Call HandleFightAccept(UserIndex)
        
        Case ClientPacketID.CloseGuild
            Call HandleCloseGuild(UserIndex)
            
        Case ClientPacketID.ChatGlobal
            Call HandleChatGlobal(UserIndex)
            
       Case ClientPacketID.LookProcess
            Call HandleLookProcess(UserIndex)
            
        Case ClientPacketID.SendProcessList
            Call HandleSendProcessList(UserIndex)
            
        Case ClientPacketID.AccionInventario
            Call HandleAccionInventario(UserIndex)
            
        Case ClientPacketID.IniciarSubasta
            Call HandleIniciaSubasta(UserIndex)
            
        Case ClientPacketID.CancelarSubasta
             Call HandleCancelarSubasta(UserIndex)
           
        Case ClientPacketID.OfertarSubasta
            Call HandleOfertaSubasta(UserIndex)
  
        Case ClientPacketID.ConsultaSubasta
            Call HandleConsultarSubasta(UserIndex)
            
        Case ClientPacketID.GMCommands              'GM Messages
            Call HandleGMCommands(UserIndex)
            
        Case ClientPacketID.Casamiento
            Call HandleCasament(UserIndex)
 
        Case ClientPacketID.Acepto
            Call HandleAcepto(UserIndex)
 
        Case ClientPacketID.Divorcio
            Call HandleDivorcio(UserIndex)
            
        Case ClientPacketID.TransferenciaOro
            Call HandleTransferenciaOro(UserIndex)
            
        Case ClientPacketID.AdoptarFamiliar
            Call HandleAdoptarFamiliar(UserIndex)
            
        Case ClientPacketID.SolicitarRank
            Call HandleSolicitarRank(UserIndex)
            
        Case ClientPacketID.BatallaPVP
            Call HandleBatallaPVP(UserIndex)
            
        Case Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)

    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.Length > 0 And Err.Number = 0 Then
        Err.Clear
        HandleIncomingData = True
    
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.Source & vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & vbTab & " LastDllError: " & Err.LastDllError & vbTab & " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(PacketId))
        Call CloseSocket(UserIndex)

        HandleIncomingData = False
    
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)

        HandleIncomingData = False

    End If

End Function

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

    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)
        
        Select Case MessageIndex

            Case eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.CombatSafeOff, eMessages.CombatSafeOn, eMessages.NobilityLost, eMessages.CantUseWhileMeditating, eMessages.FinishHome
            
            Case eMessages.NPCHitUser
                Call .WriteByte(Arg1) 'Target
                Call .WriteInteger(Arg2) 'damage
                
            Case eMessages.UserHitNPC
                Call .WriteLong(Arg1) 'damage
                
            Case eMessages.UserAttackedSwing
                Call .WriteInteger(UserList(Arg1).Char.CharIndex)
                
            Case eMessages.UserHittedByUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.UserHittedUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.WorkRequestTarget
                Call .WriteByte(Arg1) 'skill
            
            Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'VictimIndex
                Call .WriteLong(Arg2) 'Expe
            
            Case eMessages.UserKill '"" & .name & " te ha matado!"
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'AttackerIndex
            
            Case eMessages.EarnExp
            
            Case eMessages.Home
                Call .WriteByte(CByte(Arg1))
                Call .WriteInteger(CInt(Arg2))
                'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                 hasta que no se pasen los dats e .INFs al cliente, esto queda asi.
                Call .WriteASCIIString(StringArg1) 'Call .WriteByte(CByte(Arg2))
                
            Case eMessages.UserMuerto
            
            Case eMessages.NpcInmune
            
            Case eMessages.Hechizo_HechiceroMSG_NOMBRE
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
                Call .WriteASCIIString(StringArg1) 'Persona
             
            Case eMessages.Hechizo_HechiceroMSG_ALGUIEN
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
             
            Case eMessages.Hechizo_HechiceroMSG_CRIATURA
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
             
            Case eMessages.Hechizo_PropioMSG
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
         
            Case eMessages.Hechizo_TargetMSG
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
                Call .WriteASCIIString(StringArg1) 'Persona

        End Select

    End With

    Exit Sub ''

errHandler:

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

    On Error GoTo errHandler

    Dim Command As Byte

    With UserList(UserIndex)
        Call .incomingData.ReadByte
    
        Command = .incomingData.PeekByte
    
        Select Case Command

            Case eGMCommands.GMMessage                '/GMSG
                Call HandleGMMessage(UserIndex)
        
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
        
            Case eGMCommands.BanChar                 '/BAN
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
        
            Case eGMCommands.ServerMessage           '/RMSG
                Call HandleServerMessage(UserIndex)
        
            Case eGMCommands.MapMessage              '/MAPMSG
                Call HandleMapMessage(UserIndex)
            
            Case eGMCommands.NickToIP                '/NICK2IP
                Call HandleNickToIP(UserIndex)
        
            Case eGMCommands.IPToNick                '/IP2NICK
                Call HandleIPToNick(UserIndex)
        
            Case eGMCommands.GuildOnlineMembers      '/ONCLAN
                Call HandleGuildOnlineMembers(UserIndex)
        
            Case eGMCommands.TeleportCreate          '/CT
                Call HandleTeleportCreate(UserIndex)
        
            Case eGMCommands.TeleportDestroy         '/DT
                Call HandleTeleportDestroy(UserIndex)
        
            Case eGMCommands.MeteoToggle             '/METEO
                Call HandleMeteoToggle(UserIndex)
        
            Case eGMCommands.SetCharDescription      '/SETDESC
                Call HandleSetCharDescription(UserIndex)
        
            Case eGMCommands.ForceMUSICToMap          '/FORCEMUSICMAP
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
        
            Case eGMCommands.AcceptRoyalCouncilMember '/ACEPTCONSE
                Call HandleAcceptRoyalCouncilMember(UserIndex)
        
            Case eGMCommands.AcceptChaosCouncilMember '/ACEPTCONSECAOS
                Call HandleAcceptChaosCouncilMember(UserIndex)
        
            Case eGMCommands.ItemsInTheFloor         '/PISO
                Call HandleItemsInTheFloor(UserIndex)
        
            Case eGMCommands.MakeDumb                '/ESTUPIDO
                Call HandleMakeDumb(UserIndex)
        
            Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
                Call HandleMakeDumbNoMore(UserIndex)
        
            Case eGMCommands.DumpIPTables            '/DUMPSECURITY
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
        
            Case eGMCommands.ForceMUSICAll            '/FORCEMUSIC
                Call HandleForceMUSICAll(UserIndex)
        
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
        
            Case eGMCommands.CreateNPC               '/ACC y /RACC
                Call HandleCreateNPC(UserIndex)
        
            Case eGMCommands.ImperialArmour          '/AI1 - 4
                Call HandleImperialArmour(UserIndex)
        
            Case eGMCommands.ChaosArmour             '/AC1 - 4
                Call HandleChaosArmour(UserIndex)
        
            Case eGMCommands.NavigateToggle          '/NAVE
                Call HandleNavigateToggle(UserIndex)
        
            Case eGMCommands.ServerOpenToUsersToggle '/HABILITAR
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
        
            Case eGMCommands.AlterName               '/ANAME
                Call HandleAlterName(UserIndex)
        
            Case Declaraciones.eGMCommands.DoBackUp               '/DOBACKUP
                Call HandleDoBackUp(UserIndex)
        
            Case eGMCommands.ShowGuildMessages       '/SHOWCMSG
                Call HandleShowGuildMessages(UserIndex)
        
            Case eGMCommands.SaveMap                 '/GUARDAMAPA
                Call HandleSaveMap(UserIndex)
        
            Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
                Call HandleChangeMapInfoPK(UserIndex)
            
            Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
                Call HandleChangeMapInfoBackup(UserIndex)
        
            Case eGMCommands.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
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
            
            Case eGMCommands.ExitDestroy
                Call HandleExitDestroy(UserIndex)

            Case eGMCommands.ToggleCentinelActivated            '/CENTINELAACTIVADO
                Call HandleToggleCentinelActivated(UserIndex)
        
            Case eGMCommands.SearchNpc                          '/BUSCAR
                Call HandleSearchNpc(UserIndex)
           
            Case eGMCommands.SearchObj                          '/BUSCAR
                Call HandleSearchObj(UserIndex)
                                           
            Case eGMCommands.LimpiarMundo                       '/LIMPIARMUNDO
                Call HandleLimpiarMundo(UserIndex)
                
            Case eGMCommands.EditCREDITS                        '/EDITCREDITS
                Call HandleEditCredits(UserIndex)
                
            Case eGMCommands.ConsultarCreditos                     '/CONSULTARCREDITS
                Call HandleConsultarCreditos(UserIndex)
                
            Case eGMCommands.SilenciarGlobal
                Call HandleSilenciarGlobal(UserIndex)

            Case eGMCommands.ToggleGlobal
                Call HandleToggleGlobal(UserIndex)
                                           
        End Select

    End With

    Exit Sub

errHandler:
    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.description & ". Paquete: " & Command)

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
            
            If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "�El gobernador no puede oirte, acercate mas para hablar con el!", FontTypeNames.FONTTYPE_INFO)
                
            Else
                Call setHome(UserIndex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
                
            End If
        
        Else
            Call WriteConsoleMsg(UserIndex, "�Debes seleccionar al gobernador de una ciudad para establecer un nuevo hogar!", FontTypeNames.FONTTYPE_INFO)
            
        End If

    End With

End Sub

''
' Handles the "DeleteChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDeleteChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Recoaro (Recox)
'Last Modification: 07/01/20
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim PJSeleccionado As Byte
    PJSeleccionado = buffer.ReadByte
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
    '�Es un indice valido?
    If PJSeleccionado < 1 Or PJSeleccionado > MAXPJACCOUNTS Then
        Call WriteErrorMsg(UserIndex, "Error al borrar el PJ. Intentelo de nuevo o contacte con un Administrador.")
        Exit Sub
        
    End If
    
    If GetUserGuildIndexDatabase(UserList(UserIndex).AccountInfo.AccountPJ(PJSeleccionado).name) > 0 Then
        Call WriteErrorMsg(UserIndex, "El personaje que intentas borrar pertenece a un clan. Debes salir del clan antes de borrar el personaje.")
        Exit Sub
        
    End If
    
    If NameIndex(UserList(UserIndex).AccountInfo.AccountPJ(PJSeleccionado).name) > 0 Then
        Call WriteErrorMsg(UserIndex, "El personaje que intentas borrar esta conectado.")
        Exit Sub
        
    End If
    
    'Mandamos a borrar el PJ
    If BorrarUsuario(UserIndex, PJSeleccionado) Then
        'Si se pudo borrar enviamos paquete para mostrar mensaje satisfactorio en el cliente
        Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DeletedChar)
        'Mandamos la actualizacion de personajes de la cuenta
        Call LoginAccountDatabase(UserIndex, UserList(UserIndex).AccountInfo.username)
        
    Else
        Call WriteErrorMsg(UserIndex, "Error al borrar el PJ. Intentelo de nuevo o contacte con un Administrador.")
        Exit Sub
        
    End If
    
    Exit Sub
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Debug.Print UserList(UserIndex).incomingData.Length
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim SelectedID    As Byte
    Dim version     As String
    
    SelectedID = buffer.ReadByte
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    With UserList(UserIndex)
    
    'Debug.Print .AccountInfo.hash
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

        If Not AsciiValidos(.AccountInfo.AccountPJ(SelectedID).name) Then
            Call WriteErrorMsg(UserIndex, "Nombre invalido.")
            Call CloseUser(UserIndex)
            
            Exit Sub
    
        End If
        
        '�El personaje existe?
        If Not PersonajeExiste(.AccountInfo.AccountPJ(SelectedID).name) Then
            Call WriteErrorMsg(UserIndex, "El personaje no existe.")
            Call CloseUser(UserIndex)
            
            Exit Sub
    
        End If
    
        If BANCheck(.AccountInfo.AccountPJ(SelectedID).name) Then
            Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a NexusAO debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde http://nexusao.com.ar")
        ElseIf Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en http://nexusao.com.ar")
        Else
            Call ConnectUser(UserIndex, .AccountInfo.AccountPJ(SelectedID).name)
        End If
    End With
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    If UserList(UserIndex).incomingData.Length < 41 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim username    As String
    Dim version     As String
    Dim race        As eRaza
    Dim gender      As eGenero
    Dim Class       As eClass
    Dim homeland    As eCiudad
    Dim PetName     As String
    Dim PetTipo     As Byte
    
    Dim skills(NUMSKILLS - 1) As Byte
    Dim Head As Integer
    Dim i As Byte
    
    username = buffer.ReadASCIIString()

    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    Head = buffer.ReadInteger
    homeland = buffer.ReadByte()
    
    Call buffer.ReadBlock(skills, NUMSKILLS)
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call CloseUser(UserIndex)
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(UserIndex).IP) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        Call CloseUser(UserIndex)
        Exit Sub
    End If

    If GetCountUserAccount(UserIndex) >= MAX_PJ_ACCOUNT Then
        Call WriteErrorMsg(UserIndex, "No puedes crear mas de " & MAX_PJ_ACCOUNT & " personajes.")
        Call CloseUser(UserIndex)
        Exit Sub
    End If
                                        
    If Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en www.nexusao.com.ar")
    Else
        Call ConnectNewUser(UserIndex, username, race, gender, Class, homeland, Head, skills)
    End If
  
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    '15/07/2009: ZaMa - Now invisible admins talk by console.
    '23/09/2009: ZaMa - Now invisible admins can't send empty chat.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        Dim PacketCounter As Long
        PacketCounter = buffer.ReadInteger
        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.Talk
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Talk", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.name, "Dijo: " & Chat)

        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Mercenario Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

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
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))

                End If

            Else

                If RTrim(Chat) <> "" Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))

                End If

            End If

        End If

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '15/07/2009: ZaMa - Now invisible admins yell by console.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.name, "Grito: " & Chat)

        End If
            
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Mercenario Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

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
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/12/2010
    '28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
    '15/07/2009: ZaMa - Now invisible admins wisper by console.
    '03/12/2010: Enanoh - Agregue susurro a Admins en modo consulta y Los Dioses pueden susurrar en ciertos casos.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat            As String

        Dim TargetUserIndex As Integer

        Dim TargetPriv      As PlayerType

        Dim UserPriv        As PlayerType

        Dim TargetName      As String
        
        TargetName = buffer.ReadASCIIString()
        Chat = buffer.ReadASCIIString()
        
        UserPriv = .flags.Privilegios
        
        If .flags.Muerto Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            ' Offline?
            TargetUserIndex = NameIndex(TargetName)

            If TargetUserIndex = INVALID_INDEX Then

                ' Admin?
                If EsGmChar(TargetName) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    ' Whisperer admin? (Else say nothing)
                ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If
                
                ' Online
            Else
                ' Privilegios
                TargetPriv = UserList(TargetUserIndex).flags.Privilegios
                
                ' Consejeros, semis y usuarios no pueden susurrar a dioses (Salvo en consulta)
                If (TargetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (UserPriv And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 And Not .flags.EnConsulta Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                    ' Usuarios no pueden susurrar a semis o conses (Salvo en consulta)
                ElseIf (UserPriv And PlayerType.User) <> 0 And (Not TargetPriv And PlayerType.User) <> 0 And Not .flags.EnConsulta Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                Else

                    '[Consejeros & GMs]
                    If UserPriv And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.name, "Le susurro a '" & UserList(TargetUserIndex).name & "' " & Chat)
                    
                        ' Usuarios a administradores
                    ElseIf (UserPriv And PlayerType.User) <> 0 And (TargetPriv And PlayerType.User) = 0 Then
                        Call LogGM(UserList(TargetUserIndex).name, .name & " le susurro en consulta: " & Chat)

                    End If
                    
                    If LenB(Chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(Chat)
                        
                        ' Dios susurrando a distancia
                        If Not EstaPCarea(UserIndex, TargetUserIndex) Then
                            
                            Call WriteConsoleMsg(UserIndex, UserList(UserIndex).name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            Call WriteConsoleMsg(TargetUserIndex, UserList(UserIndex).name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            
                        ElseIf Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatOverHead(UserIndex, Chat, .Char.CharIndex, &HC000&, True)
                            Call WriteChatOverHead(TargetUserIndex, Chat, .Char.CharIndex, &HC000&, True)
                            Call WriteConsoleMsg(UserIndex, UserList(UserIndex).name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            Call WriteConsoleMsg(TargetUserIndex, UserList(UserIndex).name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)

                        Else
                            Call WriteConsoleMsg(UserIndex, UserList(UserIndex).name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)

                            If UserIndex <> TargetUserIndex Then Call WriteConsoleMsg(TargetUserIndex, UserList(TargetUserIndex).name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            

                        End If

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/01/2012 (Recox)
    '11/19/09 Pato - Now the class bandit can walk hidden.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '12/01/2020: Recox - TiempoDeWalk agregado para las monturas
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    Dim Heading As eHeading
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Heading = .incomingData.ReadByte()

        Dim PacketCount As Long

        PacketCount = .incomingData.ReadInteger
            
        If .flags.Muerto = 0 Then
            If .flags.Navegando Then
                Call verifyTimeStamp(PacketCount, .PacketCounters(PacketNames.sailing), .PacketTimers(PacketNames.sailing), .MacroIterations(PacketNames.sailing), UserIndex, "Sailing", PacketTimerThreshold(PacketNames.sailing), MacroIterations(PacketNames.sailing))
            Else
                Call verifyTimeStamp(PacketCount, .PacketCounters(PacketNames.Walk), .PacketTimers(PacketNames.Walk), .MacroIterations(PacketNames.Walk), UserIndex, "Walk", PacketTimerThreshold(PacketNames.Walk), MacroIterations(PacketNames.Walk))

            End If

        End If
        
        '�Est� trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call DejardeTrabajar(UserIndex)

        End If

        Dim CurrentTick As Long

        CurrentTick = GetTickCount
            
        'Prevent SpeedHack (refactored by WyroX)
        If Not EsGm(UserIndex) And .flags.Velocidad > 0 Then

            Dim ElapsedTimeStep As Long, MinTimeStep As Long, DeltaStep As Single

            ElapsedTimeStep = CurrentTick - .Counters.LastStep
            MinTimeStep = IntervaloCaminar / .flags.Velocidad
            DeltaStep = (MinTimeStep - ElapsedTimeStep) / MinTimeStep

            If DeltaStep > 0 Then
                
                .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep
                
                If .Counters.SpeedHackCounter > MaximoSpeedHack Then
                    'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Administraci�n � Posible uso de SpeedHack del usuario " & .name & ".", FontTypeNames.FONTTYPE_SERVER))
                    Call WritePosUpdate(UserIndex)
                    Exit Sub

                End If

            Else
                
                .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep * 5

                If .Counters.SpeedHackCounter < 0 Then .Counters.SpeedHackCounter = 0

            End If

        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.Particle = 0

                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticleChar(.Char.CharIndex, .Char.Particle, False, 0))
                Call MoveUserChar(UserIndex, Heading)
            Else
                'Move user
                If MoveUserChar(UserIndex, Heading) Then
                
                    ' Save current step for anti-sh
                    .Counters.LastStep = CurrentTick
                
                    'Stop resting if needed
                    If .flags.Descansar Then
                        .flags.Descansar = False
                        
                        Call WriteRestOK(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
    
                    End If
                    
                Else
                    .Counters.LastStep = 0
                    Call WritePosUpdate(UserIndex)
                    
                End If

            End If

        Else    'paralized

            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call WritePosUpdate(UserIndex)

        End If

    End With

End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    'Last Modified By: ZaMa
    '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
    '13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
            Dim PacketCounter As Long
            PacketCounter = .incomingData.ReadInteger
        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.Attack
            
            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Attack", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub

        End If
        
        '�Est� trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "�Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar asi este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Mercenario Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '02/26/2006: Marco - Agregue un checkeo por si el usuario trata de agarrar un item mientras comercia.
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
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ningUn objeto.", FontTypeNames.FONTTYPE_INFO)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)

        End If
        
        .flags.Seguro = Not .flags.Seguro

    End With

End Sub

''
' Handles the "CombatSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCombatToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Creation Date: 23/10/2020
    '***************************************************
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        .flags.ModoCombate = Not .flags.ModoCombate
        
        If .flags.ModoCombate Then
            Call WriteMultiMessage(UserIndex, eMessages.CombatSafeOn) 'Call WriteCombatSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.CombatSafeOff) 'Call WriteCombatSafeOff(UserIndex)

        End If

    End With

End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarFama(UserIndex)

End Sub

''
' Handles the "RequestFamily" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFamily(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 08/04/2021
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteFamily(UserIndex)

End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/03/2010
    '11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.ComUsu.DestUsu)

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(UserIndex) Then
                'Analize chat...
                Call Statistics.ParseChat(Chat)
                
                Chat = UserList(UserIndex).name & "> " & Chat
                Call WriteCommerceChat(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
                Call WriteConsoleMsg(otherUser, .name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '07/25/09: Marco - Agregue un checkeo para patear a los usuarios que tiran items mientras comercian.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim Slot   As Byte

    Dim Amount As Integer
    
    Dim MiObj As obj
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        Dim PacketCounter As Long
        PacketCounter = .incomingData.ReadInteger()
                        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.Drop
  
        'If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), userindex, "Drop", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub


        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or .flags.Muerto = 1 Or ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        '�Est� trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "�Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '�Puede tirar items en el mapa?
        If MapInfo(.Pos.Map).NoTirarItems = True Then
            Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos en el mapa.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub 'Don't drop too much gold

            Call TirarOro(Amount, UserIndex)
            
            Call WriteUpdateGold(UserIndex)
        Else

            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub

                End If
                
                MiObj.ObjIndex = .Invent.Object(Slot).ObjIndex
                MiObj.Amount = Amount
                
                Call DropObj(UserIndex, MiObj, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        Dim PacketCounter As Long
        PacketCounter = .incomingData.ReadInteger()
        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.CastSpell
        
        Call modHechizos.CastSpell(UserIndex, Spell)

    End With

End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Byte

        Dim Y As Byte
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Dim PacketCounter As Long
        PacketCounter = .incomingData.ReadInteger()
        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.LeftClick

        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "LeftClick", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)

    End With

End Sub

''
' Handles the "AccionClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAccionClick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte

        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)

    End With

End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '13/01/2010: ZaMa - El pirata se puede ocultar en barca
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        Dim PacketCounter As Long
        PacketCounter = .incomingData.ReadInteger

        Dim Packet_ID As Long
        Packet_ID = PacketNames.Work
        
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
        
        Select Case Skill
        
            Case Robar, Magia, Domar
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Skill)
                
            Case Ocultarse
                If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Ocultar", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
                
                ' Verifico si se peude ocultar en este mapa
                If MapInfo(.Pos.Map).OcultarSinEfecto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Ocultarse no funciona aqui!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                If .flags.EnConsulta Then
                    Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If .flags.Navegando = 1 Then
                    If .clase <> eClass.Mercenario Then

                        '[CDT 17-02-2004]
                        If Not .flags.UltimoMensaje = 3 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estas navegando.", FontTypeNames.FONTTYPE_INFO)
                            .flags.UltimoMensaje = 3

                        End If

                        '[/CDT]
                        Exit Sub

                    End If

                End If
                
                If .flags.Oculto = 1 Then

                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya estas oculto.", FontTypeNames.FONTTYPE_INFO)
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
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_FIGHT))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call CloseUser(UserIndex)

    End With

End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte()
        
        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub

        End If
        
        If .flags.Meditando Then Exit Sub

        Dim PacketCounter As Long
        PacketCounter = .incomingData.ReadInteger

        Dim Packet_ID As Long
        Packet_ID = PacketNames.UseItem
        
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItem", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        Call UseInvItem(UserIndex, Slot)

    End With

End Sub

''
' Handles the "CraftearItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftearItem(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Item As Long
        Dim Cantidad As Integer
        Dim Profesion As Byte
        
        Item = .incomingData.ReadLong()
        Cantidad = .incomingData.ReadInteger()
        Profesion = .incomingData.ReadByte()

        If Item < 1 Or Cantidad < 1 Then Exit Sub
        
        Call ComenzarCrafteo(UserIndex, Item, Cantidad, Profesion)
    End With

End Sub

Private Sub HandleWorkClose(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 21/08/2020
    '
    '***************************************************
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte

        .flags.Trabajando = 0
    
    End With
    
    
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Item As Integer
        Dim Cantidad As Integer
        
        Item = .incomingData.ReadInteger()
        Cantidad = .incomingData.ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        'Comprobamos que no se encuentra trabajando, para prevenir bugs y hacks
        If .flags.MacroTrabajo = 0 Then
            .flags.MacroTrabajaObj = Item
            .flags.MacroCountObj = Cantidad
            .flags.MacroTrabajo = eMacroTrabajo.Carpinteando
            Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "Ya te encuentras trabajando.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With
    
errHandler:
    Call LogError("Error en HandleCraftcarpenter en " & Erl & " - Item: " & Item & ". Err " & Err.Number & " " & Err.description)

End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 14/01/2010 (ZaMa)
    '16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
    '12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
    '14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueno.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X           As Byte

        Dim Y           As Byte

        Dim Skill       As eSkill

        Dim DummyINT    As Integer

        Dim tU          As Integer   'Target user

        Dim tN          As Integer   'Target NPC
        
        Dim WeaponIndex As Integer
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
        
        Dim PacketCounter As Long
        PacketCounter = .incomingData.ReadInteger()
        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.WorkLeftClick

        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "WorkLeftClick", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub

        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
        
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
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energia.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub

                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IP & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
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
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)

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
                                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub

                                End If
                                 
                                '17/09/02
                                'Check the trigger
                                If MapData(UserList(tU).Pos.Map, X, Y).Trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No puedes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
                                    Exit Sub

                                End If
                                 
                                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No puedes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
                                    Exit Sub

                                End If
                                 
                                Call DoRobar(UserIndex, tU)

                            End If

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Case eSkill.Pesca
                WeaponIndex = .Invent.WeaponEqpObjIndex

                If WeaponIndex = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.CASA Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                If HayAgua(.Pos.Map, X, Y) Then
                
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Select Case WeaponIndex

                        Case CANA_PESCA
                            .flags.MacroTrabajo = eMacroTrabajo.Pescar
                        
                        Case RED_PESCA
                        
                            If .Stats.UserSkills(eSkill.Pesca) < ObjData(WeaponIndex).MinSkill Then
                                Call WriteConsoleMsg(UserIndex, "No tienes conocimientos en Pesca suficiente para usar la red. Necesitas al menos " & ObjData(WeaponIndex).MinSkill & " Skills.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If
                            
                            If .flags.Navegando = 0 Then
                                Call WriteConsoleMsg(UserIndex, "Para pescar necesitas estar en una barca.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If
                                              
                            .flags.MacroTrabajo = eMacroTrabajo.PescarRed
   
                        Case Else

                            Exit Sub    'Invalid item!

                    End Select
                    
                    Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eSkill.Mineria, eSkill.talar, eSkill.Botanica
                'Target whatever is in the tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                DummyINT = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                If DummyINT > 0 Then

                    'Check distance
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 1 Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    DummyINT = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex 'CHECK
                    
                    Select Case Skill
                    
                        Case eSkill.talar, eSkill.Botanica

                            '�Hay arbol?
                            If ObjData(DummyINT).OBJType <> eOBJType.otArboles Then
                                Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n �rbol.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If
                            
                        Case eSkill.Mineria

                            '�Hay yacimiento?
                            If ObjData(DummyINT).OBJType <> eOBJType.otYacimiento Then
                                Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yacimiento.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                    End Select
                    
                    If PuedeExtraer(UserIndex) Then
                        .flags.MacroTrabajo = Skill
                        Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)

                    End If
                        
                Else
                    Call WriteConsoleMsg(UserIndex, "Ah� no hay ninguna fuente de recursos.", FontTypeNames.FONTTYPE_INFO)
                    
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
                            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que esta luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ninguna criatura alli!", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!

                If PuedeLingotear(UserIndex) Then
                    .flags.MacroTrabajo = eMacroTrabajo.Lingotear
                    Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnviarArmasConstruibles(UserIndex)
                        Call EnviarArmadurasConstruibles(UserIndex)
                        Call WriteShowTrabajoForm(UserIndex, eSkill.Herreria)
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahi no hay ningUn yunque.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Ahi no hay ningUn yunque.", FontTypeNames.FONTTYPE_INFO)

                End If

        End Select

    End With

End Sub

''
' Handles the "InvitarPartyClick" message.

Private Sub HandleInvitarPartyClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 05/11/2020
'***************************************************
    
    With UserList(UserIndex)
    
        If .incomingData.Length < 3 Then
            Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub
    
        End If
        
        Dim X           As Byte
    
        Dim Y           As Byte
    
        Dim Aleatorio   As Integer
        
        'Remove packet ID
        Call .incomingData.ReadByte
            
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
    
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub
    
        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
    
        End If
    
        'If exiting, cancel
        Call CancelExit(UserIndex)
            
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
    
        'Target whatever is in that tile
        Call LookatTile(UserIndex, .Pos.Map, X, Y)
                    
        If .flags.TargetUser <= 0 Then Exit Sub
                    
        'If it's outside range log it and exit
        If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
            Call LogCheating("Ataque fuera de rango de " & .name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IP & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
            Exit Sub
        End If
        
        '�Se invita a si mismo?
        If UserIndex = .flags.TargetUser Then Exit Sub
        
        '�No tengo grupo?
        If .PartyIndex = 0 Then
            
            '�El otro tampoco tiene?
            If UserList(.flags.TargetUser).PartyIndex = 0 Then
            
                'Lo podemos crear?
                If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
                
                '�Estan creando grupo?
                If .FormandoGrupo <> .Id Then
                    
                    .FormandoGrupo = UserList(.flags.TargetUser).Id
                    UserList(.flags.TargetUser).FormandoGrupo = UserList(.flags.TargetUser).Id 'Se anota asi mismo, se�al que es el invitado
                    
                    Call WriteConsoleMsg(UserIndex, "Has enviado una peticion a " & UserList(.flags.TargetUser).name & " para crear un grupo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(.flags.TargetUser, UserList(UserIndex).name & " te ha invitado para crear un grupo.", FontTypeNames.FONTTYPE_INFO)
                    
                Else '�Es la respuesta?
                    'Lo creamos
                    Call mdParty.CrearParty(.flags.TargetUser)
                    
                    'Metemos al target
                    UserList(UserIndex).PartySolicitud = UserList(.flags.TargetUser).PartyIndex
                    
                    'Lo aceptamos
                    Call mdParty.AprobarIngresoAParty(.flags.TargetUser, UserIndex)
                    
                    .FormandoGrupo = 0
                    UserList(.flags.TargetUser).FormandoGrupo = 0
                    
                    Exit Sub
                End If
                
            Else '�El otro SI tiene grupo?
            
                '�Es el lider?
                If Parties(UserList(.flags.TargetUser).PartyIndex).EsPartyLeader(.flags.TargetUser) Then
                    'Enviamos peticion para unirme
                    Call mdParty.SolicitarIngresoAParty(UserIndex)
                    Exit Sub
                Else '�No lo es?
                    Call WriteConsoleMsg(UserIndex, UserList(.flags.TargetUser).name & " ya pertenece a un grupo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
            End If
            
        Else '�SI tengo party?

            '�Soy el lider?
            If Parties(.PartyIndex).EsPartyLeader(UserIndex) Then
                '�Solicito entrar a mi party?
                If UserList(.flags.TargetUser).PartySolicitud = .PartyIndex Then
                    'Lo aceptamos
                    Call mdParty.AprobarIngresoAParty(UserIndex, .flags.TargetUser)
                    
                Else '�no?
                    Call WriteConsoleMsg(UserIndex, UserList(.flags.TargetUser).name & " no ha solicitado entrar a tu grupo.", FontTypeNames.FONTTYPE_PARTY)
                    
                End If
            End If
        
        End If
    
    End With
    
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/11/09
    '05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Desc      As String

        Dim GuildName As String

        Dim Site      As String

        Dim codex()   As String

        Dim errorStr  As String
        
        Desc = buffer.ReadASCIIString()
        GuildName = Trim$(buffer.ReadASCIIString())
        Site = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, Site, codex, .FundandoGuildAlineacion, errorStr) Then
            Dim Message As String
            Message = .name & " fundo el clan " & GuildName & " de alineacion " & modGuilds.GuildAlignment(.GuildIndex)

            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            
            'Update tag
            Call RefreshCharStatus(UserIndex)

        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpellSlot As Byte
        Dim Spell As Integer
        
        SpellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If SpellSlot < 1 Or SpellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "�Primero selecciona el hechizo.!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(SpellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .nombre & vbCrLf _
                                               & "Descripci�n:" & .Desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Mana necesario: " & .ManaRequerido & vbCrLf _
                                               & "Stamina necesaria: " & .StaRequerido & vbCrLf _
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        Dim PacketCounter As Long
        PacketCounter = .incomingData.ReadInteger()
                        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.EquipItem
        
        'If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), userindex, "EquipItem", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub
        
        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)

    End With

End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 06/28/2008
    'Last Modified By: NicoNZ
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    ' 06/28/2008: NicoNZ - Solo se puede cambiar si esta inmovilizado.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Heading As eHeading

        Dim posX    As Integer

        Dim posY    As Integer
                
        Heading = .incomingData.ReadByte()
        
         Dim PacketCounter As Long
        PacketCounter = .incomingData.ReadInteger()
                        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.ChangeHeading
            
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "ChangeHeading", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then

            Select Case Heading

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
        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

        End If

    End With

End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    '11/19/09: Pato - Adapting to new skills system.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 1 + NUMSKILLS Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim i                      As Long

        Dim Count                  As Integer

        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            
            If points(i) < 0 Then
                Call LogHackAttemp(.name & " IP:" & .IP & " trato de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub

            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.name & " IP:" & .IP & " trato de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If
        
        .Counters.AsignedSkills = MinimoInt(10, .Counters.AsignedSkills + Count)
        
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
                    
                    Call CheckEluSkill(UserIndex, i, True)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer

        Dim PetIndex   As Byte
        
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
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer mas criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))

        End If

    End With

End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot   As Byte

        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        'El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub

        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estas comerciando.", FontTypeNames.FONTTYPE_INFO)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot   As Byte

        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'Es el banquero?
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot   As Byte

        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot   As Byte

        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'El NPC puede comerciar?
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 02/01/2010
    '02/01/2010: ZaMa - Implemento nuevo sistema de foros
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim ForumMsgType As eForumMsgType
        
        Dim File         As String

        Dim Title        As String

        Dim Post         As String

        Dim ForumIndex   As Integer

        Dim postFile     As String

        Dim ForumType    As Byte
                
        ForumMsgType = buffer.ReadByte()
        
        Title = buffer.ReadASCIIString()
        Post = buffer.ReadASCIIString()
        
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
            
            Call AddPost(ForumIndex, Post, .name, Title, EsAnuncio(ForumMsgType))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Dir As Integer
        
        If .ReadBoolean() Then
            Dir = 1
        Else
            Dir = -1

        End If
        
        Call DesplazarHechizo(UserIndex, Dir, .ReadByte())

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Dir      As Integer

        Dim Slot     As Byte

        Dim TempItem As obj
        
        If .ReadBoolean() Then
            Dir = 1
        Else
            Dir = -1

        End If
        
        Slot = .ReadByte()

    End With
        
    With UserList(UserIndex)
        TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
        If Dir = 1 Then 'Mover arriba
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount

            Call UpdateBanUserInv(False, UserIndex, Slot - 1)
        Else 'mover abajo
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount

            Call UpdateBanUserInv(False, UserIndex, Slot + 1)
        End If

        Call UpdateBanUserInv(False, UserIndex, Slot)

    End With

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Desc    As String

        Dim codex() As String
        
        Desc = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(Desc, codex, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 24/11/2009
    '24/11/2009: ZaMa - Nuevo sistema de comercio
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount    As Long

        Dim Slot      As Byte

        Dim tUser     As Integer

        Dim OfferSlot As Byte

        Dim ObjIndex  As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        OfferSlot = .incomingData.ReadByte()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(UserIndex).ComUsu.Confirmo = True Then
            
            ' Finish the trade
            Call FinComerciarUsu(UserIndex)
        
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)

            End If
        
            Exit Sub

        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
        
        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub
        
        'Has he got enough??
        If Slot = FLAGORO Then

            ' Can't offer more than he has
            If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub

            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.GoldAmount Then
                    Amount = .ComUsu.GoldAmount * (-1)

                End If

            End If

        Else

            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex

            ' Can't offer more than he has
            If Not HasEnoughItems(UserIndex, ObjIndex, TotalOfferItems(ObjIndex, UserIndex) + Amount) Then
                
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub

            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
                    Amount = .ComUsu.cant(OfferSlot) * (-1)

                End If

            End If
        
            'No se puede comerciar con los items de newbie
            If ItemNewbie(ObjIndex) Then
                Call WriteCancelOfferItem(UserIndex, OfferSlot)
                Exit Sub

            End If
            
            'No se puede comerciar con la Runa
            If ObjData(ObjIndex).OBJType = otRuna Then
                Call WriteCancelOfferItem(UserIndex, OfferSlot)
                Exit Sub

            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            End If
            
            If .flags.Equitando = 1 Then
                If .Invent.MonturaEqpSlot = Slot Then
                    Call WriteConsoleMsg(UserIndex, "No podes vender tu montura mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If

        End If
        
        Call AgregarOferta(UserIndex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO)
        Call EnviarOferta(tUser, OfferSlot)

    End With

End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        Guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        Guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & Guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        Guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        Guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild    As String

        Dim proposal As String

        Dim errorStr As String
        
        Guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild    As String

        Dim proposal As String

        Dim errorStr As String
        
        Guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild    As String

        Dim errorStr As String

        Dim details  As String
        
        Guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, Guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild    As String

        Dim errorStr As String

        Dim details  As String
        
        Guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, Guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim User    As String

        Dim details As String
        
        User = buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estas habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild           As String

        Dim errorStr        As String

        Dim otherGuildIndex As Integer
        
        Guild = buffer.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, Guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String

        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, username, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(username)

            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)

            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(username & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String

        Dim username As String

        Dim Reason   As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, username, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(username, .GuildIndex, Reason)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username   As String

        Dim GuildIndex As Integer
        
        username = buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, username)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(username & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(UserIndex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .name, FontTypeNames.FONTTYPE_GUILD))

        End If

    End With

End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild       As String

        Dim application As String

        Dim errorStr    As String
        
        Guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, Guild, application, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del lider de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub WriteConsoleServerUpTimeMsg(ByVal UserIndex As Integer)
    Dim time As Long
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
        UpTimeStr = time & " dia, " & UpTimeStr
    Else
        UpTimeStr = time & " dias, " & UpTimeStr
    End If

    Call WriteConsoleMsg(UserIndex, "Tiempo del Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 14/07/19 (Recox)
    'Ahora se muestra una lista de nombres de jugadores online, se suman los gms tambien a la lista (Recox)
    '***************************************************
    Dim i     As Long

    Dim Count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim UsersNamesOnlines As String

        For i = 1 To LastUser

            If LenB(UserList(i).name) <> 0 Then

                If i = LastUser Then
                    UsersNamesOnlines = UsersNamesOnlines + UserList(i).name
                Else
                    UsersNamesOnlines = UsersNamesOnlines + UserList(i).name + ", "
                End If
                
                Count = Count + 1
            End If

        Next i
        
        Call WriteConsoleMsg(UserIndex, UsersNamesOnlines, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Numero de usuarios: " & CStr(Count), FontTypeNames.FONTTYPE_INFOBOLD)

    End With

    Call WriteConsoleServerUpTimeMsg(UserIndex)

End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ)
    'If user is invisible, it automatically becomes
    'visible before doing the countdown to exit
    '15/04/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
    '13/01/2020 - Se pusieron nuevas validaciones para las monturas. (Recox)
    '***************************************************
    Dim tUser        As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If
        
        'Subastas
        If UserIndex = Subasta.UserIndex Or UserIndex = Subasta.OfertaIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir mientras ofertas en una subasta o realizas una subasta.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_WARNING)
                    Call FinComerciarUsu(tUser)

                End If

            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado.", FontTypeNames.FONTTYPE_WARNING)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim GuildIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(UserIndex, "Tu no puedes salir de este clan.", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim earnings   As Integer

    Dim Percentage As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, UserIndex)

    End With

End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)

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
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
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
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call AccionParaEntrenador(UserIndex)

    End With

End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodas junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/08 (NicoNZ)
    'Arregle un bug que mandaba un index de la meditacion diferente
    'al que decia el server.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes meditar cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        If .flags.Equitando Then
            Call WriteConsoleMsg(UserIndex, "No puedes meditar mientras si estas montado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
            Call WriteConsoleMsg(UserIndex, "Solo las clases magicas conocen el arte de la meditacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(UserIndex, "Mana restaurado.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(UserIndex)
            Exit Sub

        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
            
            Call WriteConsoleMsg(UserIndex, "Te estas concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzaras a meditar.", FontTypeNames.FONTTYPE_INFO)
            
            .Char.loops = INFINITE_LOOPS
            
            'Show proper FX according to level
            If .Stats.ELV < 13 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
            
            ElseIf .Stats.ELV < 25 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
            
            ElseIf .Stats.ELV < 35 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE
            
            ElseIf .Stats.ELV < 42 Then
                .Char.FX = FXIDs.FXMEDITARXGRANDE
            
            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE

            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.loops = 0
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/01/20
    'Arreglo validacion de NPC para que funcione el comando. (Recox)
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate NPC and make sure player is dead
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 5 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call SacerdoteResucitateUser(UserIndex)
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
        If Not EsGm(UserIndex) Then Exit Sub
        
        UserConsulta = .flags.TargetUser
        
        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub
        
        ' No podes estra en consulta con otro gm
        If EsGm(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim username As String

        username = UserList(UserConsulta).name
        
        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & username & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.name, "Termino consulta con " & username)
            
            UserList(UserConsulta).flags.EnConsulta = False
        
            ' Sino la inicia
        Else
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & username & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.name, "Inicio consulta con " & username)
            
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor) Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call SacerdoteHealUser(UserIndex)
    End With

End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendHelp(UserIndex)

End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then

            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then

                If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub

            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
            '[Alejo]
        ElseIf .flags.TargetUser > 0 Then

            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).name

            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.cant(i) = 0
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
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
' Handles the "GoliathStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoliathStart(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 31/03/2021
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call WriteAbrirGoliath(UserIndex)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte mas.", FontTypeNames.FONTTYPE_INFO)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim Matados    As Integer

    Dim NextRecom  As Integer

    Dim Diferencia As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        NextRecom = .Faccion.NextRecompensa
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
            
            Matados = .Faccion.CriminalesMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales mas y te dare una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        Else

            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a la legion oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
            
            Matados = .Faccion.CiudadanosMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos mas y te dare una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, y creo que estas en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If

            Call RecompensaArmadaReal(UserIndex)
        Else

            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a la legion oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/08
    '01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Dim time      As Long

    Dim UpTimeStr As String
    
    Call WriteConsoleServerUpTimeMsg(UserIndex)
End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call mdParty.SalirDeParty(UserIndex)

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
    
    Dim TargetUserIndex  As Integer

    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ' Didn't target any user
        TargetUserIndex = .flags.TargetUser

        If TargetUserIndex = 0 Then Exit Sub
        
        ' Can't share with admins
        If EsGm(TargetUserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        ' Pk or Caos?
        If criminal(UserIndex) Then

            ' Caos can only share with other caos
            If esCaos(UserIndex) Then
                If Not esCaos(TargetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Solo puedes compartir npcs con miembros de tu misma faccion!!", FontTypeNames.FONTTYPE_INFO)
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
            Call WriteConsoleMsg(SharingUserIndex, .name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).name & ".", FontTypeNames.FONTTYPE_INFO)

        End If
        
        .flags.ShareNpcWith = TargetUserIndex
        
        Call WriteConsoleMsg(TargetUserIndex, .name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Ahora compartes tus npcs con " & UserList(TargetUserIndex).name & ".", FontTypeNames.FONTTYPE_INFO)
        
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
            Call WriteConsoleMsg(SharingUserIndex, .name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).name & ".", FontTypeNames.FONTTYPE_INFO)
            
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    ConsultaPopular.SendInfoEncuesta (UserIndex)

End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 15/07/2009
    '02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
    '15/07/2009: ZaMa - Now invisible admins only speak by console
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        Dim PacketCounter As Long
        PacketCounter = buffer.ReadInteger()
        
        Dim Packet_ID As Long
        Packet_ID = PacketNames.GuildMessage
            
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "GuildMessage", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & "> " & Chat))
                
                If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("< " & Chat & " >", .Char.CharIndex, vbYellow))

            End If

        End If

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            Call mdParty.BroadCastParty(UserIndex, Chat)

            'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.
 
Private Sub HandleCentinelReport(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 02/05/2012
    '                         Nuevo centinela (maTih.-)
    '***************************************************
    
    Dim NotBuff As New clsByteQueue
    
    With UserList(UserIndex)
        Call NotBuff.CopyBuffer(.incomingData)
        
        Call NotBuff.ReadByte
                
        Call modCentinela.IngresaClave(UserIndex, NotBuff.ReadASCIIString())
        
        Call .incomingData.CopyBuffer(NotBuff)
        
    End With

End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Companeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ningUn clan.", FontTypeNames.FONTTYPE_GUILDMSG)

        End If

    End With

End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    Dim Tipo As Byte
    Dim Message As String
    'Bug y Sugerencias
    Dim cant As Integer
    Dim Motivo As Integer
    Dim Nuevo As String
    Dim Mensaje As String
    Dim FileDir As String

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Tipo = .incomingData.ReadByte
        Message = .incomingData.ReadASCIIString()
        
        'Ruta donde se guardan los reportes
        FileDir = App.Path & "\logs\REPORTES\"
        
        'Si es una Consulta:
        Select Case Tipo
        
        Case 0 'Consultas
        
            If Not Ayuda.Existe(.name) Then
                Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora s�lo debes esperar que se desocupe alg�n GM.", FontTypeNames.FONTTYPE_INFO)
                Call Ayuda.Push(.name & ";" & Message)
                Exit Sub
            Else
                Call Ayuda.Quitar(.name)
                Call Ayuda.Push(.name & ";" & Message)
                Call WriteConsoleMsg(UserIndex, "Ya hab�as mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
        Case 1 'Reporte de bugs
            
            If FileExist(FileDir, vbDirectory) = False Then _
                MkDir FileDir
            
            cant = GetVar(FileDir & "Bugs.INI", "BUGS", "CANTIDAD")
            Motivo = val(cant) + 1
            Nuevo = "Bug" & Motivo
            Mensaje = Date & " " & time & " - " & UserList(UserIndex).name & " Reporto el siguiente Bug: " & Message & " - IP: " & UserList(UserIndex).IP

            Call WriteVar(FileDir & "Bugs.INI", "Bugs", "Cantidad", Motivo)
            Call WriteVar(FileDir & "Bugs.INI", "Reportes", Nuevo, Mensaje)
            
            Call WriteConsoleMsg(UserIndex, "El Bug ha sido reportado exitosamente! Gracias por colaborar con NexusAO.", FONTTYPE_GUILD)
            Call WriteConsoleMsg(SendTarget.ToAdmins, Mensaje, FONTTYPE_TALK)
            
        Case 2 'Sugerencia
            
            If FileExist(FileDir, vbDirectory) = False Then _
                MkDir FileDir
        
            cant = GetVar(FileDir & "Sugerencias.ini", "SUGERENCIAS", "CANTIDAD")
            Motivo = val(cant) + 1
            Nuevo = "Sugerencia" & Motivo
            Mensaje = Date & " " & time & " - " & UserList(UserIndex).name & " Reporto la siguiente sugerencia: " & Message & " - IP: " & UserList(UserIndex).IP

            Call WriteVar(FileDir & "Sugerencias.ini", "SUGERENCIAS", "Cantidad", Motivo)
            Call WriteVar(FileDir & "Sugerencias.ini", "Reportes", Nuevo, Mensaje)
            
            Call WriteConsoleMsg(UserIndex, "La sugerencia ha sido guardada! Gracias por colaboar con NexusAO.", FONTTYPE_GUILD)
            Call WriteConsoleMsg(SendTarget.ToAdmins, Mensaje, FONTTYPE_TALK)
            
        Case 3 'Denuncia
            
            If FileExist(FileDir, vbDirectory) = False Then _
                MkDir FileDir
        
            cant = GetVar(FileDir & "Sugerencias.ini", "DENUNCIAS", "CANTIDAD")
            Motivo = val(cant) + 1
            Nuevo = "Sugerencia" & Motivo
            Mensaje = Date & " " & time & " - " & UserList(UserIndex).name & " Reporto la siguiente denunciaa: " & Message & " - IP: " & UserList(UserIndex).IP

            Call WriteVar(FileDir & "Denuncias.ini", "SUGERENCIAS", "Cantidad", Motivo)
            Call WriteVar(FileDir & "Denuncias.ini", "Reportes", Nuevo, Mensaje)
            
            Call WriteConsoleMsg(UserIndex, "La denuncia ha sido registrada! Gracias por colaboar con NexusAO.", FONTTYPE_GUILD)
            Call WriteConsoleMsg(SendTarget.ToAdmins, Mensaje, FONTTYPE_TALK)
            
        End Select
        
    End With
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim description As String
        
        description = buffer.ReadASCIIString()

        If Not AsciiValidos(description) Then
            Call WriteConsoleMsg(UserIndex, "La descripcion tiene caracteres invalidos.", FontTypeNames.FONTTYPE_INFO)
        Else
            .Desc = Trim$(description)
            Call WriteConsoleMsg(UserIndex, "La descripcion ha cambiado.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim vote     As String

        Dim errorStr As String
        
        vote = buffer.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 25/08/2009
    '25/08/2009: ZaMa - Now only admins can see other admins' punishment list
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim name  As String

        Dim Count As Integer
        
        name = buffer.ReadASCIIString()
        
        If LenB(name) <> 0 Then
            If (InStrB(name, "\") <> 0) Then
                name = Replace(name, "\", "")

            End If

            If (InStrB(name, "/") <> 0) Then
                name = Replace(name, "/", "")

            End If

            If (InStrB(name, ":") <> 0) Then
                name = Replace(name, ":", "")

            End If

            If (InStrB(name, "|") <> 0) Then
                name = Replace(name, "|", "")

            End If
            
            If (EsAdmin(name) Or EsDios(name) Or EsSemiDios(name) Or EsConsejero(name) Or EsRolesMaster(name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else

                If PersonajeExiste(name) Then
                    Count = GetUserAmountOfPunishments(name)

                    If Count = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call SendUserPunishments(UserIndex, name, Count)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje """ & name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '10/07/2010: ZaMa - Now normal npcs don't answer if asked to gamble.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount  As Integer

        Dim TypeNpc As eNPCType
        
        Amount = .incomingData.ReadInteger()
        
        ' Dead?
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
        
            'Validate target NPC
        ElseIf .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
        
            ' Validate Distance
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        
            ' Validate NpcType
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            
            Dim TargetNpcType As eNPCType

            TargetNpcType = Npclist(.flags.TargetNPC).NPCtype
            
            ' Normal npcs don't speak
            If TargetNpcType <> eNPCType.Comun And TargetNpcType <> eNPCType.DRAGON Then
                Call WriteChatOverHead(UserIndex, "No tengo ningUn interes en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If
            
            ' Validate amount
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El minimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        
            ' Validate amount
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El maximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        
            ' Validate user gold
        ElseIf .Stats.Gld < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        
        Else

            If RandomNumber(1, 100) <= 47 Then
                .Stats.Gld = .Stats.Gld + Amount
                Call WriteChatOverHead(UserIndex, "Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.Gld = .Stats.Gld - Amount
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
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
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then
            .Stats.Banco = .Stats.Banco - Amount
            .Stats.Gld = .Stats.Gld + Amount
            Call WriteChatOverHead(UserIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

        End If
        
        Call WriteUpdateGold(UserIndex)

    End With

End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 09/28/2010
    ' 09/28/2010 C4b3z0n - Ahora la respuesta de los NPCs sino perteneces a ninguna faccion solo la hacen el Rey o el Demonio
    ' 05/17/06 - Maraxus
    '***************************************************

    Dim TalkToKing  As Boolean

    Dim TalkToDemon As Boolean

    Dim NpcIndex    As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
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
                Call WriteChatOverHead(UserIndex, "Sal de aqui bufon!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            
            Else

                ' Si le pidio al rey salir de la armada, le responde.
                If TalkToKing Then
                    Call WriteChatOverHead(UserIndex, "Seras bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(NpcIndex).Char.CharIndex, vbWhite)

                End If
                
                Call ExpulsarFaccionReal(UserIndex, False)
                
            End If
        
            'Quit the Chaos Legion?
        ElseIf .Faccion.FuerzasCaos = 1 Then

            ' Si le pidio al rey salir del caos, le responde.
            If TalkToKing Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui maldito criminal!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else

                ' Si le pidio al demonio salir del caos, este le responde.
                If TalkToDemon Then
                    Call WriteChatOverHead(UserIndex, "Ya volveras arrastrandote.", Npclist(NpcIndex).Char.CharIndex, vbWhite)

                End If
                
                Call ExpulsarFaccionCaos(UserIndex, False)

            End If

            ' No es faccionario
        Else
        
            ' Si le hablaba al rey o demonio, le repsonden ellos
            'Corregido, solo si son en efecto el rey o el demonio, no cualquier NPC (C4b3z0n)
            If (TalkToDemon And criminal(UserIndex)) Or (TalkToKing And Not criminal(UserIndex)) Then 'Si se pueden unir a la faccion (status), son invitados
                Call WriteChatOverHead(UserIndex, "No perteneces a nuestra faccion. Si deseas unirte, di /ENLISTAR", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            ElseIf (TalkToDemon And Not criminal(UserIndex)) Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui bufon!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            ElseIf (TalkToKing And criminal(UserIndex)) Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui maldito criminal!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                Call WriteConsoleMsg(UserIndex, "No perteneces a ninguna faccion!", FontTypeNames.FONTTYPE_FIGHT)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
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
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        'Calculamos la diferencia con el maximo de oro permitido el cual es el valor de LONG
        Dim RemainingAmountToMaximumGold As Long
        RemainingAmountToMaximumGold = 2147483647 - .Stats.Gld

        If .Stats.Banco >= 2147483647 And RemainingAmountToMaximumGold <= Amount Then
            Call WriteChatOverHead(UserIndex, "No puedes depositar el oro por que tendrias mas del maximo permitido (2147483647)", Npclist(.flags.TargetNPC).Char.CharIndex, vbRed)

        ElseIf Amount > 0 And Amount <= .Stats.Gld Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.Gld = .Stats.Gld - Amount
            Call WriteChatOverHead(UserIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

        End If

    End With

End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 14/11/2010
    '14/11/2010: ZaMa - Now denounces can be desactivated.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Text As String

        Dim msg  As String
        
        Text = buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)
            
            msg = LCase$(.name) & " DENUNCIA: " & Text
            
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(msg, FontTypeNames.FONTTYPE_GUILDMSG), True)
            
            Call Denuncias.Push(msg, False)
            
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 1 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If EsGm(UserIndex) Or EsRolesMaster(UserList(UserIndex).name) Then
            Call WriteConsoleMsg(UserIndex, "Los GM's no pueden fundar clanes.", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub

        End If
        
        If HasFound(.name) Then
            Call WriteConsoleMsg(UserIndex, "Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim clanType As eClanType

        Dim Error    As String
        
        clanType = .incomingData.ReadByte()
        
        If HasFound(.name) Then
            Call WriteConsoleMsg(UserIndex, "Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogCheating("El usuario " & .name & " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .IP)
            Exit Sub

        End If
        
        Select Case UCase$(Trim(clanType))

            Case eClanType.ct_RoyalArmy
                .FundandoGuildAlineacion = ALINEACION_ARMADA

            Case eClanType.ct_Evil
                .FundandoGuildAlineacion = ALINEACION_LEGION

            Case eClanType.ct_Neutral
                .FundandoGuildAlineacion = ALINEACION_NEUTRO

            Case eClanType.ct_Legal
                .FundandoGuildAlineacion = ALINEACION_CIUDA

            Case eClanType.ct_Criminal
                .FundandoGuildAlineacion = ALINEACION_CRIMINAL

            Case Else
                Call WriteConsoleMsg(UserIndex, "Alineacion invalida.", FontTypeNames.FONTTYPE_GUILD)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/05/09
    'Last Modification by: Marco Vanotti (Marco)
    '- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tUser)
            Else

                If InStr(username, "+") Then
                    username = Replace(username, "+", " ")

                End If
                
                Call WriteConsoleMsg(UserIndex, LCase(username) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "PartyAcceptMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyAcceptMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/05/09
    'Last Modification by: Marco Vanotti (Marco)
    '- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username  As String

        Dim tUser     As Integer

        Dim Rank      As Integer

        Dim bUserVivo As Boolean
        
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        username = buffer.ReadASCIIString()

        If UserList(UserIndex).flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_PARTY)
        Else
            bUserVivo = True

        End If
        
        If mdParty.UserPuedeEjecutarComandos(UserIndex) And bUserVivo Then
            tUser = NameIndex(username)

            If tUser > 0 Then

                'Validate administrative ranks - don't allow users to spoof online GMs
                If (UserList(tUser).flags.Privilegios And Rank) <= (.flags.Privilegios And Rank) Then
                    Call mdParty.AprobarIngresoAParty(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If InStr(username, "+") Then
                    username = Replace(username, "+", " ")

                End If
                
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(username) And Rank) <= (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(UserIndex, LCase(username) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild       As String

        Dim memberCount As Integer

        Dim i           As Long

        Dim username    As String
        
        Guild = buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(Guild, "\") <> 0) Then
                Guild = Replace(Guild, "\", "")

            End If

            If (InStrB(Guild, "/") <> 0) Then
                Guild = Replace(Guild, "/", "")

            End If
            
            If Not FileExist(App.Path & "\guilds\" & Guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & Guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & Guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    username = GetVar(App.Path & "\Guilds\" & Guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, username & "<" & Guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String
        
        Message = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.name, "Mensaje a Gms:" & Message)
        
            If LenB(Message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(Message)
            
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & "> " & Message, FontTypeNames.FONTTYPE_GMMSG))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i    As Long

        Dim list As String

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
                        list = list & UserList(i).name & ", "

                    End If

                End If

            End If

        Next i

    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i    As Long

        Dim list As String

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
                        list = list & UserList(i).name & ", "

                    End If

                End If

            End If

        Next i

    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/07
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String
        
        username = buffer.ReadASCIIString()
        
        Dim tIndex As Integer

        Dim X      As Long

        Dim Y      As Long

        Dim i      As Long

        Dim Found  As Boolean
        
        tIndex = NameIndex(username)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambien lo seamos
            If Not (EsDios(username) Or EsAdmin(username)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else

                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i

                                If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Call LogGM(.name, "/IRCERCA " & username & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
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
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares estan ocupados.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim comment As String

        comment = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.name, "Hora.")

    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))

End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 18/11/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim miPos    As String
        
        username = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            
            tUser = NameIndex(username)

            If tUser <= 0 Then
                
                If PersonajeExiste(username) Then
                
                    Dim CharPrivs As PlayerType

                    CharPrivs = GetCharPrivs(username)
                    
                    If (CharPrivs And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((CharPrivs And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                        miPos = GetUserPos(username)
                        Call WriteConsoleMsg(UserIndex, "Ubicacion  " & username & " (Offline): " & miPos & ".", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    If Not (EsDios(username) Or EsAdmin(username)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            Else

                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicacion  " & username & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        Call LogGM(.name, "/Donde " & username)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 30/07/06
    'Pablo (ToxicWaste): modificaciones generales para simplificar la visualizacion.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
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

        Dim List1()    As String

        Dim List2()    As String
        
        Map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(Map) Then

            For i = 1 To LastNPC

                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.Map = Map Then

                    'esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else

                            For j = 0 To NPCcount1 - 1

                                If Left$(List1(j), Len(Npclist(i).name)) = Npclist(i).name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1

                            End If

                        End If

                    Else

                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else

                            For j = 0 To NPCcount2 - 1

                                If Left$(List2(j), Len(Npclist(i).name)) = Npclist(i).name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
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
                Call WriteConsoleMsg(UserIndex, "No hay mas NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call LogGM(.name, "Numero enemigos en mapa " & Map)

        End If

    End With

End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/09
    '26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer

        Dim Y As Integer
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        X = .flags.TargetX
        Y = .flags.TargetY
        
        Call FindLegalPos(UserIndex, .flags.TargetMap, X, Y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, X, Y, True)
        Call LogGM(.name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)

    End With

End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/08/2019
    '26/03/2009: ZaMa - Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '11/08/2019: Jopi - No registramos en los logs si te teletransportas a vos mismo.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim Map      As Integer

        Dim X        As Integer

        Dim Y        As Integer

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        Map = buffer.ReadInteger()
        X = buffer.ReadByte()
        Y = buffer.ReadByte()
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(Map) And LenB(username) <> 0 Then
                If UCase$(username) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(username)

                    End If

                Else
                    tUser = UserIndex

                End If
            
                If tUser <= 0 Then
                    If Not (EsDios(username) Or EsAdmin(username)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                ElseIf Not ((UserList(tUser).flags.Privilegios And PlayerType.Dios) <> 0 Or (UserList(tUser).flags.Privilegios And PlayerType.Admin) <> 0) Or tUser = UserIndex Then
                            
                    If InMapBounds(Map, X, Y) Then
                        Call FindLegalPos(tUser, Map, X, Y)
                        Call WarpUserChar(tUser, Map, X, Y, True, True)
                        
                        ' Agrego esto para no llenar consola de mensajes al hacer SHIFT + CLICK DERECHO
                        If UserIndex <> tUser Then
                            Call WriteConsoleMsg(UserIndex, UserList(tUser).name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Transporto a " & UserList(tUser).name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)

                        End If
                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(username)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias seran ignoradas por el servidor de aqui en mas. Utilice /GM para contactar un administrador.")
                    Call LogGM(.name, "/silenciar " & UserList(tUser).name)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/DESsilenciar " & UserList(tUser).name)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    
    Dim LiderInvita As Boolean
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        LiderInvita = .incomingData.ReadBoolean

        If LiderInvita Then
            Call WritePeticionInvitarParty(UserIndex)

        ElseIf .PartyIndex > 0 Then
            Call WriteShowPartyForm(UserIndex)
            
        Else
            Call WritePeticionInvitarParty(UserIndex)

        End If

    End With

End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        username = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then Call Ayuda.Quitar(username)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim X        As Integer

        Dim Y        As Integer
        
        username = buffer.ReadASCIIString()
        tUser = NameIndex(username)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambien lo seamos
            If Not (EsDios(username) Or EsAdmin(username)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    Call LogGM(.name, "/IRA " & username & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.name, "/INVISIBLE")

    End With

End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Id As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        Id = .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex, Id)

    End With

End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    'Last modified by: Lucas Tavolaro Ortiz (Tavo)
    'I haven`t found a solution to split, so i make an array of names
    '***************************************************
    Dim i       As Long

    Dim names() As String

    Dim Count   As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser

            If (LenB(UserList(i).name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).name
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/10/2010
    '***************************************************
    Dim i     As Long

    Dim Users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).name

            End If

        Next i
        
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i     As Long

    Dim Users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser

            If (LenB(UserList(i).name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).name & ", "

            End If

        Next i
        
        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    'Last Modification: 04/04/2020
    '4/4/2020: FrankoH298 - Ahora calcula bien el tiempo de carcel
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim Reason   As String

        Dim jailTime As Byte

        Dim Count    As Byte

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()
        
        If InStr(1, username, "+") Then
            username = Replace(username, "+", " ")

        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(username) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(username)
                
                If tUser <= 0 Then
                    If (EsDios(username) Or EsAdmin(username)) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > (60) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar por mas de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If (InStrB(username, "\") <> 0) Then
                            username = Replace(username, "\", "")

                        End If

                        If (InStrB(username, "/") <> 0) Then
                            username = Replace(username, "/", "")

                        End If
                        
                        If PersonajeExiste(username) Then
                            Count = GetUserAmountOfPunishments(username)
                            Call SaveUserPunishment(username, Count + 1, LCase$(.name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & time)

                        End If
                        
                        Call Encarcelar(tUser, jailTime, .name)
                        Call LogGM(.name, " encarcelo a " & username)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/22/08 (NicoNZ)
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC   As Integer

        Dim auxNPC As NPC
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
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
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim Reason   As String

        Dim Privs    As PlayerType

        Dim Count    As Byte
        
        username = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(username) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                Privs = UserDarPrivilegioLevel(username)
                
                If Not Privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If (InStrB(username, "\") <> 0) Then
                        username = Replace(username, "\", "")

                    End If

                    If (InStrB(username, "/") <> 0) Then
                        username = Replace(username, "/", "")

                    End If
                    
                    If PersonajeExiste(username) Then
                        Count = GetUserAmountOfPunishments(username)
                        Call SaveUserPunishment(username, Count + 1, LCase$(.name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & time)

                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(username) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.name, " advirtio a " & username)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/05/2019
    '02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
    '11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
    '18/09/2010: ZaMa - Ahora se puede editar la vida del propio pj (cualquier rm o dios).
    '11/05/2019: Jopi - No registramos en los logs si te editas a vos mismo.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username      As String

        Dim tUser         As Integer

        Dim opcion        As Byte

        Dim Arg1          As String

        Dim Arg2          As String

        Dim valido        As Boolean

        Dim LoopC         As Byte

        Dim CommandString As String

        Dim n             As Byte

        Dim Var           As Long
        
        username = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(username) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(username)

        End If
        
        opcion = buffer.ReadByte()
        Arg1 = buffer.ReadASCIIString()
        Arg2 = buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then

            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)

                Case PlayerType.Consejero
                    ' Los RMs consejeros solo se pueden editar su head, body, level y vida
                    valido = tUser = UserIndex And (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida)
                
                Case PlayerType.SemiDios
                    ' Los RMs solo se pueden editar su level o vida y el head y body de cualquiera
                    valido = ((opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                    
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level o vida solo lo puede hacer sobre si mismo
                    valido = ((opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_CiticensKilled Or opcion = eEditOptions.eo_CriminalsKilled Or opcion = eEditOptions.eo_Class Or opcion = eEditOptions.eo_Skills Or opcion = eEditOptions.eo_addGold

            End Select
        
            'Si no es RM debe ser dios para poder usar este comando
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            
            If opcion = eEditOptions.eo_Vida Then
                '  Por ahora dejo para que los dioses no puedan editar la vida de otros
                valido = (tUser = UserIndex)
            Else
                valido = True

            End If
            
        ElseIf .flags.PrivEspecial Then
            valido = (opcion = eEditOptions.eo_CiticensKilled) Or (opcion = eEditOptions.eo_CriminalsKilled)
            
        End If

        'CHOTS | The user is not online and we are working with Database
        If tUser <= 0 Then
            valido = False
            Call WriteConsoleMsg(UserIndex, "El usuario esta offline.", FontTypeNames.FONTTYPE_INFO)

            '@TODO call a method to edit the user using the database
        End If

        If valido Then
            'For making the Log
            CommandString = "/MOD "
                
            Select Case opcion

                Case eEditOptions.eo_Gold

                    If val(Arg1) <= MAX_ORO_EDIT Then
                        If tUser <= 0 Then ' Esta offline?
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Stats.Gld = val(Arg1)
                            Call WriteUpdateGold(tUser)

                        End If

                    Else
                            Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    ' Log it
                    CommandString = CommandString & "ORO "
                
                Case eEditOptions.eo_Experience

                        If val(Arg1) <= MAX_EXP_EDIT Then
                        
                            If tUser <= 0 Then ' Offline
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                            Else ' Online
                                UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                                Call CheckUserLevel(tUser)
                                Call WriteUpdateExp(tUser)
    
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_EXP_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXP "
                    
                Case eEditOptions.eo_Body

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else
                            Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim, UserList(tUser).Char.AuraAnim, UserList(tUser).Char.AuraColor)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "BODY "
                    
                Case eEditOptions.eo_Head

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim, UserList(tUser).Char.AuraAnim, UserList(tUser).Char.AuraColor)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "HEAD "
                    
                Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Faccion.CriminalesMatados = Var

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CRI "
                    
                Case eEditOptions.eo_CiticensKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Faccion.CiudadanosMatados = Var

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CIU "
                    
                Case eEditOptions.eo_Level

                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)

                        End If
                        
                        ' Chequeamos si puede permanecer en el clan
                        If val(Arg1) >= 25 Then
                            
                            Dim GI As Integer

                            If tUser <= 0 Then
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                            Else
                                GI = UserList(tUser).GuildIndex

                            End If
                            
                            If GI > 0 Then
                                If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                                    'We get here, so guild has factionary alignment, we have to expulse the user
                                    Call modGuilds.m_EcharMiembroDeClan(-1, username)
                                    
                                    Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(username & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))

                                    ' Si esta online le avisamos
                                    If tUser > 0 Then Call WriteConsoleMsg(tUser, "Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearas! Por esta razon, hasta tanto no te enlistes en la faccion bajo la cual tu clan esta alineado, estaras excluido del mismo.", FontTypeNames.FONTTYPE_GUILD)

                                End If

                            End If

                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
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

                            If tUser <= 0 Then ' Offline
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                            Else ' Online
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

                            If tUser <= 0 Then ' Offline
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                            Else ' Online
                                UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                                Call CheckEluSkill(tUser, LoopC, True)

                            End If

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLS "
                        
                Case eEditOptions.eo_SkillPointsLeft

                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Stats.SkillPts = val(Arg1)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLSLIBRES "
                    
                Case eEditOptions.eo_Nobleza
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Reputacion.NobleRep = Var

                        End If
                    
                        ' Log it
                        CommandString = CommandString & "NOB "
                        
                Case eEditOptions.eo_Asesino
                    Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                    If tUser <= 0 Then ' Offline
                        Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                    Else ' Online
                        UserList(tUser).Reputacion.AsesinoRep = Var

                    End If
                        
                    ' Log it
                    CommandString = CommandString & "ASE "
                    
                Case eEditOptions.eo_Sex

                    Dim Sex As Byte

                    Sex = IIf(UCase(Arg1) = "MUJER", eGenero.Mujer, 0) ' Mujer?
                    Sex = IIf(UCase(Arg1) = "HOMBRE", eGenero.Hombre, Sex) ' Hombre?
                        
                    If Sex <> 0 Then ' Es Hombre o mujer?
                        If tUser <= 0 Then ' OffLine
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Genero = Sex

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)

                    End If
                        
                    ' Log it
                    CommandString = CommandString & "SEX "
                    
                Case eEditOptions.eo_Raza

                    Dim Raza As Byte
                        
                    Arg1 = UCase$(Arg1)

                    Select Case Arg1

                        Case "HUMANO"
                            Raza = eRaza.Humano

                        Case "ELFO"
                            Raza = eRaza.Elfo

                        Case "DROW"
                            Raza = eRaza.Drow

                        Case "ENANO"
                            Raza = eRaza.Enano

                        Case "GNOMO"
                            Raza = eRaza.Gnomo
                            
                        Case Else
                            Raza = 0

                        End Select
                            
                    If Raza = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else
                            UserList(tUser).Raza = Raza

                        End If

                    End If
                            
                    ' Log it
                    CommandString = CommandString & "RAZA "
                        
                Case eEditOptions.eo_addGold
                    
                    Dim bankGold As Long
                        
                    If Abs(Arg1) > MAX_ORO_EDIT Then
                        Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")

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
                    UserList(tUser).Stats.MaxHp = val(Arg1)
                    UserList(tUser).Stats.MinHp = val(Arg1)
                        
                    Call WriteUpdateUserStats(tUser)
                        
                    ' Log it
                    CommandString = CommandString & "VIDA "
                        
                Case eEditOptions.eo_Poss
                    
                    Dim Map As Integer

                    Dim X   As Integer

                    Dim Y   As Integer
                        
                    Map = val(ReadField(1, Arg1, 45))
                    X = val(ReadField(2, Arg1, 45))
                    Y = val(ReadField(3, Arg1, 45))
                        
                    If InMapBounds(Map, X, Y) Then
                            
                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.name, "Intento editar un usuario inexistente u offline.")
                        Else
                            Call WarpUserChar(tUser, Map, X, Y, True, True)
                            Call WriteConsoleMsg(UserIndex, "Usuario teletransportado: " & username, FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Posicion invalida", FONTTYPE_INFO)

                    End If
                        
                    ' Log it
                    CommandString = CommandString & "POSS "
                    
                    Case eEditOptions.eo_Speed
                        
                        Dim Speed As Double
                        
                        If val(Arg1) > 50 Then _
                            Arg1 = 50
                            
                        Speed = val(Arg1)
                        
                        UserList(tUser).flags.Velocidad = Speed
                        Call WriteSetSpeed(tUser)
                        
                Case Else
                    Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                    CommandString = CommandString & "UNKOWN "
                        
            End Select
                
            CommandString = CommandString & Arg1 & " " & Arg2
                
            If UserIndex <> tUser Then
                Call LogGM(.name, CommandString & " " & username)
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
                
        Dim TargetName  As String

        Dim targetIndex As Integer
        
        TargetName = Replace$(buffer.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(TargetName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

            'is the player offline?
            If targetIndex <= 0 Then

                'don't allow to retrieve administrator's info
                If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando...", FontTypeNames.FONTTYPE_INFO)

                    Call SendUserStatsTxtDatabase(UserIndex, TargetName)

                End If

            Else

                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, targetIndex)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
        
        username = buffer.ReadASCIIString()
         
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And ((.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin) Then
            Call LogGM(.name, "/STAT " & username)
            
            tUser = NameIndex(username)
            
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_INFO)

                    Call SendUserMiniStatsTxtFromDatabase(UserIndex, username)

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
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
        
        username = buffer.ReadASCIIString()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or UserIsAdmin Then
            
            Call LogGM(.name, "/BAL " & username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserOROTxtFromDatabase(UserIndex, username)

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario " & username & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
        
        username = buffer.ReadASCIIString()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/INV " & username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando...", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserInvTxtFromDatabase(UserIndex, username)

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
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean

        username = buffer.ReadASCIIString()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            Call LogGM(.name, "/BOV " & username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserBovedaTxtFromDatabase(UserIndex, username)

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la boveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserBovedaTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la boveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim LoopC    As Long

        Dim Message  As String
        
        username = buffer.ReadASCIIString()
        tUser = NameIndex(username)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/STATS " & username)
            
            If tUser <= 0 Then
                If (InStrB(username, "\") <> 0) Then
                    username = Replace(username, "\", "")

                End If

                If (InStrB(username, "/") <> 0) Then
                    username = Replace(username, "/", "")

                End If
                
                For LoopC = 1 To NUMSKILLS
                    Message = Message & GetUserSkills(username)
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, Message & "CHAR> Libres: " & GetUserFreeSkills(username), FontTypeNames.FONTTYPE_INFO)

            Else
                Call SendUserSkillsTxt(UserIndex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        username = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(username) <> "YO" Then
                tUser = NameIndex(username)
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
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    .Stats.MinHp = .Stats.MaxHp
                    
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call LogGM(.name, "Resucito a " & username)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    Dim i    As Long

    Dim list As String

    Dim priv As PlayerType
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).name & ", "

            End If

        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
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

        Dim list  As String

        Dim priv  As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser

            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).name & ", "

            End If

        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
        Call LogGM(.name, "/ONLINEMAP " & Map)

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.name, "Intento perdonar un personaje de nivel avanzado.")
                    
                    If Not (EsDios(username) Or EsAdmin(username)) Then
                        Call WriteConsoleMsg(UserIndex, "Solo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim Rank     As Integer

        Dim IsAdmin  As Boolean
        
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        username = buffer.ReadASCIIString()
        IsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or IsAdmin Then
            tUser = NameIndex(username)
            
            If tUser <= 0 Then
                If Not (EsDios(username) Or EsAdmin(username)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " echo a " & username & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseUser(tUser)
                    Call LogGM(.name, "Echo a " & username)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "Estas loco?? Como vas a pinatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " ha ejecutado a " & username & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.name, " ejecuto a " & username)

                End If

            Else

                If Not (EsDios(username) Or EsAdmin(username)) Then
                    Call WriteConsoleMsg(UserIndex, "No esta online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas loco?? Como vas a pinatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim Reason   As String
        
        username = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, username, Reason)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username  As String

        Dim cantPenas As Byte
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")

            End If

            If (InStrB(username, "/") <> 0) Then
                username = Replace(username, "/", "")

            End If
            
            If Not PersonajeExiste(username) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else

                If BANCheck(username) Then
                    Call UnBan(username)
                
                    'penas
                    cantPenas = GetUserAmountOfPunishments(username)
                    Call SaveUserPunishment(username, cantPenas + 1, LCase$(.name) & ": UNBAN. " & Date & " " & time)
                
                    Call LogGM(.name, "/UNBAN a " & username)
                    Call WriteConsoleMsg(UserIndex, username & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, username & " no esta baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
            Call DoFollow(.flags.TargetNPC, .name)
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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim X        As Integer

        Dim Y        As Integer
        
        username = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(username)
            
            If tUser <= 0 Then
                If EsDios(username) Or EsAdmin(username) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Else

                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)
                    Call LogGM(.name, "/SUM " & username & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim NPC As Integer

        NPC = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If NPC > 0 And NPC <= UBound(Declaraciones.SpawnList()) Then Call SpawnNpc(Declaraciones.SpawnList(NPC).NpcIndex, .Pos, True, False)
            
            Call LogGM(.name, "Sumoneo " & Declaraciones.SpawnList(NPC).NpcName)

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
        Call LogGM(.name, "/RESETINV " & Npclist(.flags.TargetNPC).name)

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
    '28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK))

                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                'frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).name & " > " & message
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                
                Dim Mapa As Integer
                                        Mapa = .Pos.Map

                Call LogGM(.name, "Mensaje a mapa " & Mapa & ":" & Message)
                Call SendData(SendTarget.toMap, Mapa, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim priv     As PlayerType

        Dim IsAdmin  As Boolean
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(username)
            Call LogGM(.name, "NICK2IP Solicito la IP de " & username)
            
            IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0

            If IsAdmin Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User

            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & username & " es " & UserList(tUser).IP, FontTypeNames.FONTTYPE_INFO)

                    Dim IP    As String

                    Dim lista As String

                    Dim LoopC As Long

                    IP = UserList(tUser).IP

                    For LoopC = 1 To LastUser

                        If UserList(LoopC).IP = IP Then
                            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).name & ", "

                                End If

                            End If

                        End If

                    Next LoopC

                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If Not (EsDios(username) Or EsAdmin(username)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "No hay ningUn personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim IP    As String

        Dim LoopC As Long

        Dim lista As String

        Dim priv  As PlayerType
        
        IP = .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, "IP2NICK Solicito los Nicks de IP " & IP)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User

        End If

        For LoopC = 1 To LastUser

            If UserList(LoopC).IP = IP Then
                If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).name & ", "

                    End If

                End If

            End If

        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String

        Dim tGuild    As Integer
        
        GuildName = buffer.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")

        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Mapa  As Integer

        Dim X     As Byte

        Dim Y     As Byte

        Dim Radio As Byte
        
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        Radio = .incomingData.ReadByte()
        
        Radio = MinimoInt(Radio, 6)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.name, "/CT " & Mapa & "," & X & "," & Y & "," & Radio)
        
        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
        If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim ET As obj

        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.ObjIndex = TELEP_OBJ_INDEX + Radio
        
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

        Dim X    As Byte

        Dim Y    As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)

            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                
                                Call LogGM(UserList(UserIndex).name, "/DT: " & Mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
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
' Handles the "ExitDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExitDestroy(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Cucsifae
    'Last Modification: 30/9/18
    '
    '***************************************************
    With UserList(UserIndex)

        Dim Mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/de
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)

            If .TileExit.Map = 0 Then Exit Sub

            'Si hay un Teleport hay que usar /DT
            If .ObjInfo.ObjIndex > 0 Then
                If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then Exit Sub

            End If

            Call LogGM(UserList(UserIndex).name, "/DE: " & Mapa & "," & X & "," & Y)
                
            .TileExit.Map = 0
            .TileExit.X = 0
            .TileExit.Y = 0

        End With

    End With

End Sub

''
' Handles the "MeteoToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeteoToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    Dim Forzar As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Forzar = .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call LogGM(.name, "/METEO " & Forzar)
        
        Lloviendo = Not Lloviendo
        
        Call SortearClima(Forzar)

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
        
        If Not EsGm(UserIndex) Then Exit Sub
        
        Dim Activado As Boolean

        Dim msg      As String
        
        Activado = Not .flags.SendDenounces
        .flags.SendDenounces = Activado
        
        msg = "Denuncias por consola " & IIf(Activado, "ativadas", "desactivadas") & "."
        
        Call LogGM(.name, msg)
        
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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim tUser As Integer

        Dim Desc  As String
        
        Desc = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser

            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim musicID As Byte

        Dim Mapa   As Integer
        
        musicID = .incomingData.ReadByte
        Mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map

            End If
        
            If musicID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMusic(MapInfo(.Pos.Map).music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMusic(musicID))

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
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte

        Dim Mapa   As Integer

        Dim X      As Byte

        Dim Y      As Byte
        
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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("EJERCITO REAL> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
        
        Dim X       As Long

        Dim Y       As Long

        Dim bIsExit As Boolean
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then

                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)

                        End If

                    End If

                End If

            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).name, "/MASSDEST")

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(username)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(username & " fue aceptado en el honorable Consejo Real de Belleuve.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(username)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(username & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
        
        Dim tObj  As Integer

        Dim lista As String

        Dim X     As Long

        Dim Y     As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex

                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).name, FontTypeNames.FONTTYPE_INFO)

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(username)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(username)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(username)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)

            Else

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Belleuve.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(username & " fue expulsado del consejo de Belleuve.", FontTypeNames.FONTTYPE_CONSEJO))

                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(username & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))

                    End If

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte

        Dim tLog     As String
        
        tTrigger = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.name, tLog)
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
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger
        
        Call LogGM(.name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)

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
        
        Call LogGM(.name, "/BANIPLIST")
        
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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName   As String

        Dim cantMembers As Integer

        Dim LoopC       As Long

        Dim member      As String

        Dim Count       As Byte

        Dim tIndex      As Integer

        Dim tFile       As String
        
        GuildName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " baneo al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_GUILD))
                
                'baneamos a los miembros
                Call LogGM(.name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)

                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseUser(tIndex)

                    End If

                    Call SaveBan(member, "BAN AL CLAN: " & GuildName, LCase$(.name))
                Next LoopC

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/02/09
    'Agregado un CopyBuffer porque se producia un bucle
    'inifito al intentar banear una ip ya baneada. (NicoNZ)
    '07/02/09 Pato - Ahora no es posible saber si un gm esta o no online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bannedIP As String

        Dim tUser    As Integer

        Dim Reason   As String

        Dim i        As Long
        
        ' Is it by ip??
        If buffer.ReadBoolean() Then
            bannedIP = buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte()
        Else
            tUser = NameIndex(buffer.ReadASCIIString())
            
            If tUser > 0 Then bannedIP = UserList(tUser).IP

        End If
        
        Reason = buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.name, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " baneo la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser

                        If UserList(i).ConnIDValida Then
                            If UserList(i).IP = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).name, "IP POR " & Reason)

                            End If

                        End If

                    Next i

                End If

            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no esta online.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
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
    'Last Modification: 11/02/2011
    'maTih.- : Ahora se puede elegir, la cantidad a crear.
    '***************************************************
    
    On Error GoTo errHandler
    
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        
        ' Recibo el ID del paquete
        Call .incomingData.ReadByte

        Dim tObj    As Integer: tObj = .incomingData.ReadInteger()
        Dim Cuantos As Integer: Cuantos = .incomingData.ReadInteger()
        
        ' Es Game-Master?
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        ' Si hace mas de 10000, lo sacamos cagando.
        If Cuantos > 10000 Then Call WriteConsoleMsg(UserIndex, "Estas tratando de crear demasiado, como mucho podes crear 10.000 unidades.", FontTypeNames.FONTTYPE_TALK): Exit Sub
        
        ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
        If tObj < 1 Or tObj > NumObjDatas Then Exit Sub

        ' El nombre del objeto es nulo?
        If LenB(ObjData(tObj).name) = 0 Then Exit Sub

        Dim Objeto As obj
        
        With Objeto
            .Amount = Cuantos
            .ObjIndex = tObj
        End With
        
        ' Chequeo si el objeto es AGARRABLE(para las puertas, arboles y demas objs. que no deberian estar en el inventario)
        '   0 = SI
        '   1 = NO
        If ObjData(tObj).Agarrable = 0 Then
            ' Trato de meterlo en el inventario.
            If MeterItemEnInventario(UserIndex, Objeto) Then
                Call WriteConsoleMsg(UserIndex, "Has creado " & Objeto.Amount & " unidades de " & ObjData(tObj).name & ".", FontTypeNames.FONTTYPE_INFO)
            Else
                ' Si no hay espacio, lo tiro al piso.
                Call TirarItemAlPiso(.Pos, Objeto)
                Call WriteConsoleMsg(UserIndex, "No tenes espacio en tu inventario para crear el item.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
            End If
        Else
            ' Crear el item NO AGARRARBLE y tirarlo al piso.
            Call TirarItemAlPiso(.Pos, Objeto)
            Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        ' Lo registro en los logs.
        Call LogGM(.name, "/CI: " & tObj & " - [Nombre del Objeto: " & ObjData(tObj).name & "] - [Cantidad : " & Cuantos & "]")
        
    End With
    
errHandler:
    If Err.Number <> 0 Then
        Call LogError("Error en HandleCreateItem " & Err.Number & " " & Err.description)
    End If
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

        Dim X    As Byte

        Dim Y    As Byte
        
        Mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
        
        Dim ObjIndex As Integer

        ObjIndex = MapData(Mapa, X, Y).ObjInfo.ObjIndex
        
        If ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.name, "/DEST " & ObjIndex & " en mapa " & Mapa & " (" & X & "," & Y & "). Cantidad: " & MapData(Mapa, X, Y).ObjInfo.Amount)
        
        If ObjData(ObjIndex).OBJType = eOBJType.otTeleport And MapData(Mapa, X, Y).TileExit.Map > 0 Then
            
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports asi. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or .flags.PrivEspecial Then
            
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")

            End If

            If (InStrB(username, "/") <> 0) Then
                username = Replace(username, "/", "")

            End If

            tUser = NameIndex(username)
            
            Call LogGM(.name, "ECHO DEL CAOS A: " & username)
    
            If tUser > 0 Then
                Call ExpulsarFaccionCaos(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
            Else

                If PersonajeExiste(username) Then
                    Call KickUserChaosLegion(username)
                    Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, username & " inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or .flags.PrivEspecial Then
            
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")

            End If

            If (InStrB(username, "/") <> 0) Then
                username = Replace(username, "/", "")

            End If

            tUser = NameIndex(username)
            
            Call LogGM(.name, "ECHO DE LA REAL A: " & username)
            
            If tUser > 0 Then
                Call ExpulsarFaccionReal(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
            Else

                If PersonajeExiste(username) Then
                    Call KickUserRoyalArmy(username)
                    Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, username & " inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMUSICAll(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim musicID As Byte

        musicID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " broadcast musica MUSIC: " & musicID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMusic(musicID))

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
    If UserList(UserIndex).incomingData.Length < 2 Then
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
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username   As String

        Dim punishment As Byte

        Dim NewText    As String
        
        username = buffer.ReadASCIIString()
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(username) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else

                If (InStrB(username, "\") <> 0) Then
                    username = Replace(username, "\", "")

                End If

                If (InStrB(username, "/") <> 0) Then
                    username = Replace(username, "/", "")

                End If
                
                If PersonajeExiste(username) Then
                    Call LogGM(.name, " borro la pena: " & punishment & " de " & username & " y la cambio por: " & NewText)

                    Call AlterUserPunishment(username, punishment, LCase$(.name) & ": <" & NewText & "> " & Date & " " & time)
                    Call WriteConsoleMsg(UserIndex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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

        Call LogGM(.name, "/BLOQ")
        
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
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.name, "/MATA " & Npclist(.flags.TargetNPC).name)

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
        
        Dim X As Long

        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)

                End If

            Next X
        Next Y

        Call LogGM(.name, "/MASSKILL")

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username   As String

        Dim lista      As String

        Dim LoopC      As Byte

        Dim priv       As Integer

        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then

            'Handle special chars
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")

            End If

            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "/", "")

            End If

            If (InStrB(username, "+") <> 0) Then
                username = Replace(username, "+", " ")

            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(username) > 0 Then
                validCheck = (UserList(NameIndex(username)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(username) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

            End If
            
            If validCheck Then
                Call LogGM(.name, "/LASTIP " & username)
                
                If PersonajeExiste(username) Then
                    lista = "Las ultimas IPs con las que " & username & " se conecto son:" & vbCrLf & GetUserLastIps(username)
                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & username & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, username & " es de mayor jerarquia que vos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the user`s chat color
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Color As Long
        
        Color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = Color

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
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim username         As String

        Dim Slot             As Byte

        Dim tIndex           As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
                
        username = buffer.ReadASCIIString() 'Que UserName?
        Slot = buffer.ReadByte() 'Que Slot?
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            
            Call LogGM(.name, .name & " Checkeo el slot " & Slot & " de " & username)
            
            tIndex = NameIndex(username)  'Que user index?
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tIndex > 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
                        If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay ningUn objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Slot Invalido.", FontTypeNames.FONTTYPE_TALK)

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
        Call .incomingData.CopyBuffer(buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
        
        'time and Time BUG!
        Call LogGM(.name, .name & " reinicio el mundo.")
        
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
        
        Call LogGM(.name, .name & " ha recargado los objetos.")
        
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
        
        Call LogGM(.name, .name & " ha recargado los hechizos.")
        
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
        
        Call LogGM(.name, .name & " ha recargado los INITs.")
        
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
         
        Call LogGM(.name, .name & " ha recargado los NPCs.")
    
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
        
        Call LogGM(.name, .name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados

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
        
        Call LogGM(.name, .name & " ha solicitado mostrar el formulario del servidor.")
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
        
        Call LogGM(.name, .name & " ha borrado los SOS.")
        
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
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha guardado todos los chars.")
        
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios

    End With

End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the backup`s info of the map
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la informacion sobre el BackUp.")
        
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
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the pk`s info of the  map
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la informacion sobre si es PK el mapa.")
        
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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.name, .name & " ha cambiado la informacion sobre si es restringido el mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).Restringir = RestrictStringToByte(tStr)
                
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & RestrictByteToString(MapInfo(.Pos.Map).Restringir), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la magia el mapa.")
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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la invisibilidad en el mapa.")
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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar el resucitar en el mapa.")
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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion del terreno del mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).Terreno = TerrainStringToByte(tStr)
                
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & TerrainByteToString(MapInfo(.Pos.Map).Terreno), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el Unico Util es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el Unico Util es 'DUNGEON' ya que al ingresarlo, NO se sentira el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim RoboNpc As Byte
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        RoboNpc = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido robar npcs en el mapa.")
            
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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim NoOcultar As Byte

    Dim Mapa      As Integer
    
    With UserList(UserIndex)
    
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        NoOcultar = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            Mapa = .Pos.Map
            
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido ocultarse en el mapa " & Mapa & ".")
            
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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim NoInvocar As Byte

    Dim Mapa      As Integer
    
    With UserList(UserIndex)
    
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        NoInvocar = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            Mapa = .Pos.Map
            
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido invocar en el mapa " & Mapa & ".")
            
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
        
        Call LogGM(.name, .name & " ha guardado el mapa " & CStr(.Pos.Map))
        
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
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Allows admins to read guild messages
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Guild As String
        
        Guild = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, Guild)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
        
        Call LogGM(.name, .name & " ha hecho un backup.")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete

    End With

End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message
 
Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 02/05/2012
    'Nuevo centinela (maTih.-)
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        'Solo para Admins y Dioses
        If Not EsAdmin(.name) Or Not EsDios(.name) Then Exit Sub
        
        Call modCentinela.CambiarEstado(UserIndex)
        
    End With

End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user name
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim username     As String

        Dim newName      As String

        Dim changeNameUI As Integer

        Dim GuildIndex   As Integer
        
        username = buffer.ReadASCIIString()
        newName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If LenB(username) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(username)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj esta online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else

                    If Not PersonajeExiste(username) Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & username & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If GetUserGuildIndex(username) > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & username & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If Not PersonajeExiste(newName) Then
                                Call CopyUser(username, newName)

                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.name, "Ha cambiado de nombre al usuario " & username & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)

    '**********************************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/05/2019
    '11/05/2019: Jopi - Se combino HandleCreateNPCWithRespawn() con este procedimiento.
    '**********************************************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer: NpcIndex = .incomingData.ReadInteger()
        Dim Respawn As Boolean: Respawn = .incomingData.ReadBoolean()
        
        'Nos fijamos que sea GM.
        If Not EsGm(UserIndex) Then Exit Sub
        
        'Invocamos el NPC.
        If NpcIndex <> 0 Then
        
            NpcIndex = SpawnNpc(NpcIndex, .Pos, True, Respawn)
        
            Call LogGM(.name, "Invoco " & IIf(Respawn, "con respawn", vbNullString) & " a " & Npclist(NpcIndex).name & " [Indice: " & NpcIndex & "] en el mapa " & .Pos.Map)

        End If

    End With

End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim index    As Byte

        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case index

            Case 1
                ArmaduraImperial1 = ObjIndex
            
            Case 2
                ArmaduraImperial2 = ObjIndex
            
            Case 3
                ArmaduraImperial3 = ObjIndex
            
            Case 4
                TunicaMagoImperial = ObjIndex

        End Select

    End With

End Sub

''
' Handle the "ChaosArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim index    As Byte

        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case index

            Case 1
                ArmaduraCaos1 = ObjIndex
            
            Case 2
                ArmaduraCaos2 = ObjIndex
            
            Case 3
                ArmaduraCaos3 = ObjIndex
            
            Case 4
                TunicaMagoCaos = ObjIndex

        End Select

    End With

End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    'Turns off the server
    '***************************************************
    Dim handle As Integer
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & .name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & time & " server apagado por " & .name & ". "
        
        Close #handle
        
        Call CloseServer

    End With

End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/CONDEN " & username)
            
            tUser = NameIndex(username)

            If tUser > 0 Then Call VolverCriminal(tUser)

        End If
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 06/09/09
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim Char     As String
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            Call LogGM(.name, "/RAJAR " & username)
            
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else

                If PersonajeExiste(username) Then
                    Call ResetUserFacciones(username)
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & username & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username   As String

        Dim GuildIndex As Integer
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/RAJARCLAN " & username)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, username)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningUn clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(username & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Request user mail
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String

        Dim mail     As String
        
        username = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If PersonajeExiste(username) Then
                mail = GetUserEmail(username)
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & username & ":" & mail, FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Mensaje de sistema:" & Message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    'Modified by: Juan Martin Sotuyo Dodero (Maraxus)
    '   - Fixed a bug that prevented from properly setting the new number of lines.
    '   - Fixed a bug that caused the player to be kicked.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim newMOTD           As String

        Dim auxiliaryString() As String

        Dim LoopC             As Long
        
        newMOTD = buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(ConfigPath & "Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(ConfigPath & "Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin sotuyo Dodero (Maraxus)
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

        Dim LoopC           As Long
        
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
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim sLlave As String

        Dim sClave As String

        Dim sValor As String

        'Obtengo los parametros
        sLlave = buffer.ReadASCIIString()
        sClave = buffer.ReadASCIIString()
        sValor = buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then

            Dim sTmp As String

            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(UserIndex, "No puedes modificar esa informacion desde aqui!", FontTypeNames.FONTTYPE_INFO)
            Else
                'Obtengo el valor segUn llave y clave
                sTmp = GetVar(ConfigPath & "Server.ini", sLlave, sClave)

                'Si obtengo un valor escribo en el server.ini
                If LenB(sTmp) Then
                    Call WriteVar(ConfigPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.name, "Modifico en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                    Call WriteConsoleMsg(UserIndex, "Modifico " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Logged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.Logged)
    
        Call .outgoingData.WriteByte(.clase)
        Call .outgoingData.WriteByte(Hour(Now))

    End With
    
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Disconnect" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(UserList(UserIndex).ComUsu.DestNick)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowTrabajoForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowTrabajoForm(ByVal UserIndex As Integer, ByVal Trabajo As Byte)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowTrabajoForm)
        Call .WriteByte(Trabajo)
    End With
Exit Sub

errHandler:
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateGold" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.Gld)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateExp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))

    End With

    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))

    End With

    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMap" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        'Call .WriteASCIIString(MapInfo(Map).Name)
        Call .WriteInteger(version)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PosUpdate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)

    End With

    Exit Sub

errHandler:

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
                             ByVal Color As Long, _
                             Optional ByVal NoConsole As Boolean = False)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, Color, NoConsole))
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
Public Sub WriteRenderMsg(ByVal UserIndex As Integer, _
                           ByVal Chat As String, _
                           ByVal FontIndex As Integer)

    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareRenderConsoleMsg(Chat, FontIndex))
    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(Chat, FontIndex))
    Exit Sub

errHandler:

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

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal Chat As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildChat" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(Chat))
    Exit Sub

errHandler:

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
Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
    
        On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Message)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)

    End With

    Exit Sub

errHandler:

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
                                ByVal body As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal X As Byte, _
                                ByVal Y As Byte, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal name As String, _
                                ByVal NickColor As Byte, _
                                ByVal Privileges As Byte, _
                                ByVal GrhAura As Long, _
                                ByVal AuraColor As Long)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, name, NickColor, Privileges, GrhAura, AuraColor))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterRemove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, _
                              ByVal CharIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Writes the "ForceCharMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, _
                                ByVal body As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal AuraAnim As Long, _
                                ByVal AuraColor As Long)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterChange" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet, AuraAnim, AuraColor))
    Exit Sub

errHandler:

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
                             ByVal GrhIndex As Long, _
                             ByVal ParticulaIndex As Integer, _
                             ByVal Rango As Byte, _
                             ByVal Color As Long, _
                             ByVal X As Byte, _
                             ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, ParticulaIndex, Rango, Color, X, Y))
    Exit Sub

errHandler:

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

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal TieneLuz As Boolean = False)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectDelete" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y, TieneLuz))
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockPosition" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)

    End With

    Exit Sub

errHandler:

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

Public Sub WritePlayMusic(ByVal UserIndex As Integer, _
                         ByVal music As Integer, _
                         Optional ByVal loops As Integer = -1)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PlayMidi" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMusic(music, loops))
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim Tmp As String

    Dim i   As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AreaChanged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PauseToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ActualizarClima" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteActualizarClima(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ActualizarClima" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageActualizarClima())
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateFX" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.Gld)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 3/12/09
    'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
    '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)

        End If
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(obData.name)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(Slot).Equipped)
        Call .WriteLong(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHit)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteSingle(SalePrice(ObjIndex))
        Call .WriteBoolean(ItemNoUsaConUser(UserIndex, ObjIndex))

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
        ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)

        End If
        
        Call .WriteASCIIString(obData.name)
        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(Slot).Amount)
        Call .WriteLong(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHit)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteLong(obData.valor)
        Call .WriteBoolean(ItemNoUsaConUser(UserIndex, ObjIndex))

    End With

    Exit Sub

errHandler:

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
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(Slot))
        
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).nombre)
        Else
            Call .WriteASCIIString("(None)")
        End If
    End With
Exit Sub

errHandler:
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
'Last Modification: 05/17/06
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) \ ModHerreria(UserList(UserIndex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

errHandler:
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
'Last Modification: 05/17/06
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) \ ModHerreria(UserList(UserIndex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

errHandler:
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
On Error GoTo errHandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
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
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.Madera)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
        Next i
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SastreriaObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSastreRopas(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer

    ReDim validIndexes(1 To UBound(ObjSastre()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SastreObjects)

        For i = 1 To UBound(ObjSastre())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjSastre(i)).SkSastreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) \ ModSastreria(UserList(UserIndex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ObjSastre(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.PielLobo)
            Call .WriteInteger(obj.PielOsoPardo)
            Call .WriteInteger(obj.PielOsoPolar)
            Call .WriteInteger(ObjSastre(validIndexes(i)))
        Next i
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "AlquimiaObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlquimistaPociones(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ObjAlquimia()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlquimiaObjects)
        
        For i = 1 To UBound(ObjAlquimia())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjAlquimia(i)).SkAlquimia <= UserList(UserIndex).Stats.UserSkills(eSkill.Alquimia) \ ModAlquimia(UserList(UserIndex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ObjAlquimia(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.Raices)
            Call .WriteInteger(ObjAlquimia(validIndexes(i)))
        Next i
    End With
Exit Sub

errHandler:
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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Atributes" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RestOK" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
    Exit Sub

errHandler:

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
Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ErrorMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(Message))

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Blind" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Dumb" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
    Exit Sub

errHandler:

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

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSignal" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).texto)
        Call .WriteLong(ObjData(ObjIndex).GrhSecundario)

    End With

    Exit Sub

errHandler:

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
                                       ByRef obj As obj, _
                                       ByVal price As Single)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Last Modified by: Budi
    'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler

    Dim ObjInfo As ObjData
    
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)

    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(ObjInfo.name)
        Call .WriteInteger(obj.Amount)
        Call .WriteSingle(price)
        Call .WriteLong(ObjInfo.GrhIndex)
        Call .WriteInteger(obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHit)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.MaxDef)
        Call .WriteInteger(ObjInfo.MinDef)
        Call .WriteBoolean(ItemNoUsaConUser(UserIndex, obj.ObjIndex))

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

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

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Family" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteFamily(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim nHabilidades As Byte
    Dim i As Byte
    Dim tmpStr As String

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.Family)
        
        Call .outgoingData.WriteASCIIString(.Familiar.nombre)
        Call .outgoingData.WriteByte(.Familiar.Tipo)
        Call .outgoingData.WriteByte(.Familiar.Nivel)
        Call .outgoingData.WriteLong(.Familiar.Exp)
        Call .outgoingData.WriteLong(.Familiar.ELU)
        Call .outgoingData.WriteLong(.Familiar.MinHp)
        Call .outgoingData.WriteLong(.Familiar.MaxHp)
        Call .outgoingData.WriteLong(.Familiar.MinHIT)
        Call .outgoingData.WriteLong(.Familiar.MaxHit)
        
        nHabilidades = nHabilidadesFamily(UserIndex)
        
        If nHabilidades <> 0 Then
            For i = 0 To nHabilidades - 1
                tmpStr = tmpStr & HabilidadName(.Familiar.Spell(i)) & ","
            Next i
        
        End If
        
        Call .outgoingData.WriteASCIIString(tmpStr)
        
    End With
    
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MiniStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.CriminalesMatados)
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call .WriteLong(UserList(UserIndex).Stats.Muertes)
        
        Call .WriteByte(UserList(UserIndex).clase)
        Call .WriteByte(UserList(UserIndex).Genero)
        Call .WriteByte(UserList(UserIndex).Raza)
        

    End With

    Exit Sub

errHandler:

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

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LevelUp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)

    End With

    Exit Sub

errHandler:

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
                            ByRef Message As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 02/01/2010
    'Writes the "AddForumMsg" message to the given user's outgoing data buffer
    '02/01/2010: ZaMa - Now sends Author and forum type
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteByte(ForumType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Author)
        Call .WriteASCIIString(Message)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowForumForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim Visibilidad   As Byte

    Dim CanMakeSticky As Byte
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.ShowForumForm)
        
        Visibilidad = eForumVisibility.ieGENERAL_MEMBER
        
        If esCaos(UserIndex) Or EsGm(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER

        End If
        
        If esArmada(UserIndex) Or EsGm(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER

        End If
        
        Call .outgoingData.WriteByte(Visibilidad)
        
        ' Pueden mandar sticky los gms o los del consejo de armada/caos
        If EsGm(UserIndex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1

        End If
        
        Call .outgoingData.WriteByte(CanMakeSticky)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetInvisible" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
    
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MeditateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlindNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumbNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    'Writes the "SendSkills" message to the given user's outgoing data buffer
    '11/19/09: Pato - Now send the percentage of progress of the skills.
    '***************************************************
    On Error GoTo errHandler

    Dim i As Long
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.SendSkills)
        Call .outgoingData.WriteByte(.clase)
        
        For i = 1 To NUMSKILLS
            Call .outgoingData.WriteByte(UserList(UserIndex).Stats.UserSkills(i))

            If .Stats.UserSkills(i) < MAXSKILLPOINTS Then
                Call .outgoingData.WriteByte(Int(.Stats.ExpSkills(i) * 100 / .Stats.EluSkills(i)))
            Else
                Call .outgoingData.WriteByte(0)

            End If

        Next i

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNews" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Tmp = vbNullString

        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OfferDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal bank As Long, ByVal reputation As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(Gold)
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

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString

        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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
                             ByVal guildDesc As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i    As Long

    Dim temp As String
    
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
        
        If Len(temp) > 1 Then temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        
        Call .WriteASCIIString(guildDesc)

    End With

    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildAlign)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/12/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Writes the "ParalizeOK" message to the given user's outgoing data buffer
    'And updates user position
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ParalizeOK)
    End With
    
    Call WritePosUpdate(UserIndex)

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

errHandler:

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
                                    ByVal ObjIndex As Integer, _
                                    ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
    '25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteByte(OfferSlot)
        Call .WriteInteger(ObjIndex)
        Call .WriteLong(Amount)
        
        If ObjIndex > 0 Then
        
            Call .WriteLong(ObjData(ObjIndex).GrhIndex)
            Call .WriteByte(ObjData(ObjIndex).OBJType)
            Call .WriteInteger(ObjData(ObjIndex).MaxHit)
            Call .WriteInteger(ObjData(ObjIndex).MinHIT)
            Call .WriteInteger(ObjData(ObjIndex).MaxDef)
            Call .WriteInteger(ObjData(ObjIndex).MinDef)
            Call .WriteLong(SalePrice(ObjIndex))
            Call .WriteASCIIString(ObjData(ObjIndex).name)
            Call .WriteBoolean(ItemNoUsaConUser(UserIndex, ObjIndex))
            
        Else ' Borra el item
        
            Call .WriteLong(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
            Call .WriteASCIIString("")
            Call .WriteBoolean(False)

        End If

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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
    On Error GoTo errHandler
    
    Dim DenounceIndex As Long

    Dim DenounceList  As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowDenounces)
        
        For DenounceIndex = 1 To Denuncias.Longitud
            DenounceList = DenounceList & Denuncias.VerElemento(DenounceIndex, False) & SEPARATOR
        Next DenounceIndex
        
        If LenB(DenounceList) <> 0 Then DenounceList = Left$(DenounceList, Len(DenounceList) - 1)
        
        Call .WriteASCIIString(DenounceList)

    End With

    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    Dim i                         As Long

    Dim Tmp                       As String

    Dim PI                        As Integer

    Dim members(PARTY_MAXMEMBERS) As Integer
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPartyForm)
        
        PI = UserList(UserIndex).PartyIndex
        Call .WriteByte(CByte(Parties(PI).EsPartyLeader(UserIndex)))
        
        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(members())

            For i = 1 To PARTY_MAXMEMBERS

                If members(i) > 0 Then
                    Tmp = Tmp & UserList(members(i)).name & " (" & Fix(Parties(PI).MiExperiencia(members(i))) & ")" & SEPARATOR

                End If

            Next i

        End If
        
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
            
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WritePeticionInvitarParty(ByVal UserIndex As Integer)
   '***************************************************
    'Author: Lorwik
    'Last Modification: 05/11/2020
    '***************************************************
    On Error GoTo errHandler
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeticionInvitarParty)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)

    End With

    Exit Sub

errHandler:

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

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer, ByVal Id As Byte)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.ShowGMPanelForm)
        Call .WriteByte(Id)
    
    End With

    Exit Sub

errHandler:

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
                             ByVal cant As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06 NIGO:
    'Writes the "UserNameList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "OfrecerFamiliar" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfrecerFamiliar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.OfrecerFamiliar)
    Exit Sub

errHandler:

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Sends all data existing in the buffer
    '***************************************************
    Dim sndData As String
    
    With UserList(UserIndex).outgoingData

        If .Length = 0 Then Exit Sub
        
        sndData = .ReadASCIIStringFixed(.Length)

    End With
    
    ' Tratamos de enviar los datos.
    Dim ret As Long: ret = WsApiEnviar(UserIndex, sndData)
    
    ' Si recibimos un error como respuesta de la API, cerramos el socket.
    If ret <> 0 And ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "SetInvisible" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.Length)

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
        
        PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.Length)

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
                                           ByVal Color As Long, _
                                           Optional ByVal NoConsole As Boolean = False) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ChatOverHead" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(NoConsole)
        
        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(Color And &HFF)
        Call .WriteByte((Color And &HFF00&) \ &H100&)
        Call .WriteByte((Color And &HFF0000) \ &H10000)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, _
                                         ByVal FontIndex As FontTypeNames) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ConsoleMsg" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.Length)

    End With

End Function
Public Function PrepareRenderConsoleMsg(ByVal Chat As String, _
                                         ByVal FontIndex As Integer) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ConsoleMsg" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RenderMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(FontIndex)

        PrepareRenderConsoleMsg = .ReadASCIIStringFixed(.Length)

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
        
        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.Length)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessagePlayWave(ByVal wave As Integer, _
                                       ByVal X As Byte, _
                                       ByVal Y As Byte) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteInteger(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal Chat As String) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "GuildChat" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.Length)

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
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMusic(ByVal music As Integer, _
                                       Optional ByVal loops As Integer = -1) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PlayMidi" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMusic)
        Call .WriteInteger(music)
        Call .WriteInteger(loops)
        
        PrepareMessagePlayMusic = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PauseToggle" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ActualizarClima" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageActualizarClima() As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ActualizarClima" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ActualizarClima)
        Call .WriteByte(DayStatus)
        
        PrepareMessageActualizarClima = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte, Optional ByVal TieneLuz As Boolean = False) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ObjectDelete" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(TieneLuz)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.Length)

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
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Long, _
                                           ByVal ParticulaIndex As Integer, _
                                           ByVal Rango As Byte, _
                                           ByVal Color As Long, _
                                           ByVal X As Byte, _
                                           ByVal Y As Byte) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'prepares the "ObjectCreate" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteLong(GrhIndex)
        Call .WriteInteger(ParticulaIndex)
        Call .WriteByte(Rango)
        Call .WriteLong(Color)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterRemove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.Length)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal X As Byte, _
                                              ByVal Y As Byte, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal name As String, _
                                              ByVal NickColor As Byte, _
                                              ByVal Privileges As Byte, _
                                              ByVal GrhAura As Long, _
                                              ByVal AuraColor As Long) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterCreate" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(Heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(name)
        Call .WriteByte(NickColor)
        Call .WriteByte(Privileges)
        Call .WriteLong(GrhAura)
        Call .WriteLong(AuraColor)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessageCharacterChange(ByVal body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal AuraAnim As Long, _
                                              ByVal AuraColor As Long) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterChange" message and returns it
    '***************************************************
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteByte(Heading)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteLong(AuraAnim)
        Call .WriteLong(AuraColor)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.Length)

    End With

End Function
Public Function PrepareMessageHeadingChange(ByVal Heading As eHeading, _
                                            ByVal CharIndex As Integer)

    '***************************************************
    'Author: FrankoH298
    'Last Modification: 10/09/19
    'Prepares the "HeadingChange" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.HeadingChange)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(Heading)

        PrepareMessageHeadingChange = .ReadASCIIStringFixed(.Length)

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
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterMove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.Length)

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
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.Length)

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
    'Last Modified By: Juan Martin Sotuyo Dodero (Maraxus)
    'Prepares the "UpdateTagAndStatus" message and returns it
    '15/01/2010: ZaMa - Now sends the nick color instead of the status.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteByte(NickColor)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ErrorMsg" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.errorMsg)
        Call .WriteASCIIString(Message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.Length)

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
    On Error GoTo errHandler
    
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.StopWorking)
        
    Exit Sub

errHandler:

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
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(Slot)

    End With
    
    Exit Sub

errHandler:

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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet id
        Call buffer.ReadByte
        
        Dim NewDialog As String

        NewDialog = buffer.ReadASCIIString
        
        Call .incomingData.CopyBuffer(buffer)
        
        If .flags.TargetNPC > 0 Then

            ' Dsgm/Dsrm/Rm
            If Not ((.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster)) Then
                'Replace the NPC's dialog.
                Npclist(.flags.TargetNPC).Desc = NewDialog

            End If

        End If

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
        If (.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim NpcIndex As Integer

        NpcIndex = .flags.TargetNPC
        
        If NpcIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        
        ' Teleports user to npc's coords
        Call WarpUserChar(UserIndex, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, False, True)
        
        ' Log gm
        Call LogGM(.name, "/IMPERSONAR con " & Npclist(NpcIndex).name & " en mapa " & .Pos.Map)
        
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
        If (.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) And (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.RoleMaster)) <> (PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim NpcIndex As Integer

        NpcIndex = .flags.TargetNPC
        
        If NpcIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        Call LogGM(.name, "/MIMETIZAR con " & Npclist(NpcIndex).name & " en mapa " & .Pos.Map)
        
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
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet id
        Call buffer.ReadByte
        
        Dim username As String

        Dim Reason   As String
        
        username = buffer.ReadASCIIString
        Reason = buffer.ReadASCIIString
    
        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then

            'Verificamos que exista el personaje
            If Not PersonajeExiste(username) Then
                Call WriteShowMessageBox(UserIndex, "El personaje no existe")
            Else
                'Agregamos el seguimiento
                Call AddRecord(UserIndex, username, Reason)
                
                'Enviamos la nueva lista de personajes
                Call WriteRecordList(UserIndex)

            End If

        End If

        Call .incomingData.CopyBuffer(buffer)

    End With
        
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet id
        Call buffer.ReadByte
        
        Dim RecordIndex As Byte

        Dim Obs         As String
        
        RecordIndex = buffer.ReadByte
        Obs = buffer.ReadASCIIString
        
        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then
            'Agregamos la observacion
            Call AddObs(UserIndex, RecordIndex, Obs)
            
            'Actualizamos la informacion
            Call WriteRecordDetails(UserIndex, RecordIndex)

        End If

        Call .incomingData.CopyBuffer(buffer)

    End With
        
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

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
        
        'Solo dioses pueden remover los seguimientos, los otros reciben una advertencia:
        If (.flags.Privilegios And PlayerType.Dios) Then
            Call RemoveRecord(RecordIndex)
            Call WriteShowMessageBox(UserIndex, "Se ha eliminado el seguimiento.")
            Call WriteRecordList(UserIndex)
        Else
            Call WriteShowMessageBox(UserIndex, "Solo los dioses pueden eliminar seguimientos.")

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
    Dim i        As Long

    Dim tIndex   As Integer

    Dim tmpStr   As String

    Dim TempDate As Date

    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.RecordDetails)
        
        'Creador y motivo
        Call .WriteASCIIString(Records(RecordIndex).Creador)
        Call .WriteASCIIString(Records(RecordIndex).Motivo)
        
        tIndex = NameIndex(Records(RecordIndex).Usuario)
        
        'Status del pj (online?)
        Call .WriteBoolean(tIndex > 0)
        
        'Escribo la IP segUn el estado del personaje
        If tIndex > 0 Then
            'La IP Actual
            tmpStr = UserList(tIndex).IP
        Else 'String nulo
            tmpStr = vbNullString

        End If

        Call .WriteASCIIString(tmpStr)
        
        'Escribo tiempo online segUn el estado del personaje
        If tIndex > 0 Then
            'Tiempo logueado.
            TempDate = Now - UserList(tIndex).LogOnTime
            tmpStr = Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate)
        Else
            'Envio string nulo.
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

errHandler:

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
    Dim i As Long

    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.RecordList)
        
        Call .WriteByte(NumRecords)

        For i = 1 To NumRecords
            Call .WriteASCIIString(Records(i).Usuario)
        Next i

    End With

    Exit Sub

errHandler:

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

        Dim newSlot      As Byte
    
        Call .incomingData.ReadByte
    
        originalSlot = .incomingData.ReadByte
        newSlot = .incomingData.ReadByte
        Call .incomingData.ReadByte
    
        Call InvUsuario.moveItem(UserIndex, originalSlot, newSlot)
    
    End With

End Sub

''
' Handles the "LoginExistingAccount" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingAccount(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim username As String

    Dim Password As String

    Dim version  As String
    
    username = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()

    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    If Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en http://nexusao.com.ar")
    Else
        Call ConnectAccount(UserIndex, username, Password)

    End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Public Sub WriteEnviarPJUserAccount(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 12/10/2018
'Writes the "AccountLogged" message to the given user with the data of the account he just logged in
'***************************************************
    On Error GoTo errHandler

    Dim i As Long

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.EnviarPJUserAccount)
        .Redundance = RandomNumber(5, 250)
        Call .outgoingData.WriteByte(.Redundance)
        Call .outgoingData.WriteASCIIString(.AccountInfo.username)
        Call .outgoingData.WriteByte(.AccountInfo.NumPjs)

        If .AccountInfo.NumPjs > 0 Then

            For i = 1 To .AccountInfo.NumPjs
                Call .outgoingData.WriteASCIIString(.AccountInfo.AccountPJ(i).name)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).body)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).Head)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).weapon)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).shield)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).helmet)
                Call .outgoingData.WriteByte(.AccountInfo.AccountPJ(i).Class)
                Call .outgoingData.WriteByte(.AccountInfo.AccountPJ(i).race)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).Map)
                Call .outgoingData.WriteByte(.AccountInfo.AccountPJ(i).level)
                Call .outgoingData.WriteBoolean(.AccountInfo.AccountPJ(i).criminal)
                Call .outgoingData.WriteBoolean(.AccountInfo.AccountPJ(i).dead)
                Call .outgoingData.WriteBoolean(.AccountInfo.AccountPJ(i).gameMaster)
            Next i

        End If
    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Function PrepareMessageFXtoMap(ByVal FxIndex As Integer, _
                                      ByVal loops As Byte, _
                                      ByVal X As Integer, _
                                      ByVal Y As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.FXtoMap)
        Call .WriteByte(loops)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteInteger(FxIndex)
        
        PrepareMessageFXtoMap = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function WriteSearchList(ByVal UserIndex As Integer, _
                                ByVal Num As Integer, _
                                ByVal Datos As String, _
                                ByVal obj As Boolean) As String
 
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SearchList)
        Call .WriteInteger(Num)
        Call .WriteBoolean(obj)
        Call .WriteASCIIString(Datos)

    End With
 
errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)

        Resume

    End If
   
End Function
 
Public Sub HandleSearchNpc(ByVal UserIndex As Integer)
 
    On Error GoTo errHandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
       
        Call buffer.ReadByte
       
        Dim i       As Long

        Dim n       As Integer

        Dim name    As String

        Dim UserNpc As String

        Dim tStr    As String

        UserNpc = buffer.ReadASCIIString()
        
        Call .incomingData.CopyBuffer(buffer)
        
        ' Es Game-Master?
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        tStr = Tilde(UserNpc)
      
        For i = 1 To val(LeerNPCs.GetValue("INIT", "NumNPCs"))
            name = LeerNPCs.GetValue("NPC" & i, "Name")
       
            If InStr(1, Tilde(name), tStr) Then
                Call WriteSearchList(UserIndex, i, CStr(i & " - " & name), False)
                n = n + 1

            End If

        Next i
   
        If n = 0 Then
            Call WriteSearchList(UserIndex, 0, "No hubo resultados de la busqueda.", False)

        End If

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
   
    Set buffer = Nothing
   
    If Error <> 0 Then Err.Raise Error

End Sub
 
Private Sub HandleSearchObj(ByVal UserIndex As Integer)
       
    On Error GoTo errHandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
           
        Call buffer.ReadByte
           
        Dim UserObj As String

        Dim tUser   As Integer

        Dim n       As Integer

        Dim i       As Long

        Dim tStr    As String
       
        UserObj = buffer.ReadASCIIString()
        
        Call .incomingData.CopyBuffer(buffer)
        
        ' Es Game-Master?
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        tStr = Tilde(UserObj)
          
        For i = 1 To UBound(ObjData)

            If InStr(1, Tilde(ObjData(i).name), tStr) Then
                Call WriteSearchList(UserIndex, i, CStr(i & " - " & ObjData(i).name), True)
                n = n + 1

            End If

        Next

        If n = 0 Then
            Call WriteSearchList(UserIndex, 0, "No hubo resultados de la busqueda.", False)

        End If
                
    End With
     
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
       
    Set buffer = Nothing
       
    If Error <> 0 Then Err.Raise Error
        
End Sub

Public Sub WriteUserInEvent(ByVal UserIndex As Integer)
    On Error GoTo errHandler
    
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserInEvent)
    Exit Sub

errHandler:
        If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
            Call FlushBuffer(UserIndex)
            Resume
        End If
End Sub

Private Sub HandleFightSend(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim ListUsers As String
        Dim GldRequired As Long
        Dim Users() As String
        
        ListUsers = buffer.ReadASCIIString & "-" & .name
        GldRequired = buffer.ReadLong
        
        If Len(ListUsers) >= 1 Then
            Users = Split(ListUsers, "-")
                      
            Call Retos.SendFight(UserIndex, GldRequired, Users)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleFightAccept(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim username As String
        
        username = buffer.ReadASCIIString
        
        If Len(username) >= 1 Then
            Call Retos.AcceptFight(UserIndex, username)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleCloseGuild(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
    
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim PreviousGuildIndex  As Integer
        
        If Not .GuildIndex >= 1 Then
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningun clan.", FONTTYPE_GUILD)
            Exit Sub

        End If

        If guilds(.GuildIndex).Fundador <> .name Then
            Call WriteConsoleMsg(UserIndex, "No eres lider del clan.", FONTTYPE_GUILD)
            Exit Sub

        End If
        
        'Ya con cambiarle el nombre a "CLAN CERRADO" ya se omite de la lista de clanes enviadas al cliente.
        'Tambien cambiamos "Founder" y "Leader" a "NADIE" sino no te deja fundar otro clan.
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "GuildName", "CLAN CERRADO")
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "Founder", "NADIE")
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "Leader", "NADIE")
        
        PreviousGuildIndex = .GuildIndex
        
        'Obtenemos la lista de miembros del clan.
        Dim GuildMembers() As String
            GuildMembers = guilds(PreviousGuildIndex).GetMemberList()

        For i = 0 To UBound(GuildMembers)
            Call SaveUserGuildIndex(GuildMembers(i), 0)
            Call SaveUserGuildAspirant(GuildMembers(i), 0)
        Next i
        
        'La borramos junto con la lista de solicitudes.
        Call Kill(App.Path & "\Guilds\" & guilds(PreviousGuildIndex).GuildName & "-members.mem")
        Call Kill(App.Path & "\Guilds\" & guilds(PreviousGuildIndex).GuildName & "-solicitudes.sol")
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El Clan " & guilds(.GuildIndex).GuildName & " ha cerrado sus puertas.", FontTypeNames.FONTTYPE_GUILD))
        
    End With

    ' Actualizamos la base de datos de clanes.
    Call modGuilds.LoadGuildsDB
        
    Exit Sub

End Sub

Public Sub HandleLimpiarMundo(ByVal UserIndex As Integer)
'***************************************************
'Author: Jopi
'Last Modification: 11/01/2020
'Fuerza una limpieza del mundo.
'***************************************************
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Me fijo si es GM
    If Not EsGm(UserIndex) Then Exit Sub
    
    
    Call LogGM(UserList(UserIndex).name, " forzo la limpieza del mundo.")
    
    tickLimpieza = 301
    
End Sub

Public Sub HandleEditCredits(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 30/04/2020
'Edita las Creditos del usuario
'***************************************************
    
    Dim username As String
    Dim CantCredits As Long
    Dim opcion As Byte
    
    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte
        
        username = .incomingData.ReadASCIIString
        CantCredits = .incomingData.ReadLong
        opcion = .incomingData.ReadByte
        
        'Me fijo si es Admin
        If Not EsAdmin(UserList(UserIndex).name) Then Exit Sub
        
        If username = "" Then
            Call WriteConsoleMsg(UserIndex, "�Faltan parametros!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If CantCredits > 10000 Then
            Call WriteConsoleMsg(UserIndex, "El valor de las Creditos no puede superar 10000", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case opcion
        
            Case 0 'Editar las Creditos
                If Cuentas.SaveAccountEditCreditosDatabase(username, CantCredits) Then
                    Call WriteConsoleMsg(UserIndex, "Se editaron " & CantCredits & " Creditos a la cuenta de " & username, FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "ERROR: No se pudo editar las Creditos a la cuenta del usuario." & username, FontTypeNames.FONTTYPE_INFO)
                    
                End If
            
            Case 1 'Sumar las Creditos
                If Cuentas.SaveAccountSumaCreditosDatabase(username, CantCredits) Then
                    Call WriteConsoleMsg(UserIndex, "Se sumaron " & CantCredits & " Creditos a la cuenta de " & username & ". Ahora tiene " & Cuentas.GetCreditosDatabase(username) & " Creditos. ", FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "ERROR: No se pudo sumar las Creditos a la cuenta del usuario." & username, FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Case 2 'Restar las Creditos
                If Cuentas.SaveAccountRestaCreditosDatabase(username, CantCredits) Then
                    Call WriteConsoleMsg(UserIndex, "Se restaron " & CantCredits & " Creditos a la cuenta de " & username & ". Ahora tiene " & Cuentas.GetCreditosDatabase(username) & " Creditos. ", FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "ERROR: No se pudo restar las Creditos de la cuenta del usuario." & username, FontTypeNames.FONTTYPE_INFO)
                    
                End If
        End Select
    End With
    
End Sub

Public Sub HandleConsultarCreditos(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 30/04/2020
'Consulta las Creditos del usuario
'***************************************************

    Dim username As String
    
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        username = .incomingData.ReadASCIIString
    
        'Me fijo si es Admin
        If Not EsAdmin(UserList(UserIndex).name) Then Exit Sub
        
        If username = "" Then
            Call WriteConsoleMsg(UserIndex, "�Faltan parametros!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call WriteConsoleMsg(UserIndex, username & " tiene " & Cuentas.GetCreditosDatabase(username) & " Creditos en su cuenta.", FontTypeNames.FONTTYPE_INFO)
    
    End With
End Sub

''
' Writes the "EquitandoToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquitandoToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 23/08/11
'Writes the "EquitandoToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.EquitandoToggle)
        
    End With

    Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSetSpeed(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 23/10/11
'Writes the "EquitandoToggle" message to the given user's outgoing data buffer
'***************************************************

On Error GoTo errHandler

    Dim Client_Speed As Double

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.SetSpeed)
        
        'Transformamos a valores que maneja el cliente
        Client_Speed = .flags.Velocidad / 100
        
        Call .outgoingData.WriteDouble(Client_Speed)
        
    End With

    Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleChatGlobal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripci�n: Conversaciones por chat global
'***************************************************

    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
      
        'Remove packet ID
        Call buffer.ReadByte
      
        Dim Message As String
        Message = buffer.ReadASCIIString()
      
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
      
        '�El chat global esta activo?
        If GlobalChatActive = True Then

            '�Esta muerto?
            If .flags.Muerto = 1 Then
                Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                Exit Sub
            End If

            '�Tiene el nivel requerido?
            If .Stats.ELV < MINLVLGLOBAL Then
                Call WriteConsoleMsg(UserIndex, "Para usar el chat global debes ser nivel " & MINLVLGLOBAL & " como minimo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            '�El usuario esta silenciado?
            If UserList(UserIndex).flags.Global = 0 Then
                Call WriteConsoleMsg(UserIndex, "No puedes hablar por el chat global por que has sido silenciado.", FontTypeNames.FONTTYPE_INFO)
            Else
                
                'Si no pasaron 5 segundos desde el �ltimo mensaje global enviado por el usuario
                If IntervaloPermiteChatGlobal(UserIndex) Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & "> " & Message, FontTypeNames.FONTTYPE_TALK))
                    Call LogGlobal(.name & "> " & Message)
    
                Else
                    Call WriteConsoleMsg(UserIndex, "Debes esperar al menos " & Intervalo_Global / 1000 & " segundos entre cada mensaje.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
            End If
            
        Else
        
            Call WriteConsoleMsg(UserIndex, "El chat global se encuentra deshabilitado en estos momentos.", FontTypeNames.FONTTYPE_INFO)
        End If
          
    End With

errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
  
    'Destroy auxiliar buffer
    Set buffer = Nothing
  
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleSilenciarGlobal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripci�n: Silencia a un usuario del chat global
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errHandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim username As String
        Dim tUser As Integer

        username = buffer.ReadASCIIString()

        'Reemplazamos el + con el espacio
        If InStr(1, username, "+") Then
            username = Replace(username, "+", " ")
        End If

        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(username)
                
            'Se encuentra offline?
            If tUser <= 0 Then
               Call WriteConsoleMsg(UserIndex, "El personaje no esta online.", FontTypeNames.FONTTYPE_INFO)
               
            Else 'Si esta online...
            
                '�Tiene el chat global activado?
                If UserList(tUser).flags.Global = 1 Then
                    UserList(tUser).flags.Global = 0
                    Call WriteConsoleMsg(UserIndex, "Se ha silenciado el chat global del usuario: " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_INFO)
                    'Le metemos un plus y le avisamos al usuario que se port� mal y que no tiene m�s chat global
                    Call WriteShowMessageBox(tUser, "Has sido silenciado del chat global indefinidamente.")
                    
                    'Guardamos en el log del gm la acci�n
                    Call LogGM(.name, "Ha prohibido el uso del chat global de: " & UserList(tUser).name)
                    'Guardamos en el log de los usuarios con el chat global prohibido el gm que realiz� la acci�n y el usuario
          
                    Call BanGlobalChatAgregar(UserList(tUser).name)
          
                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                    
                Else '�Tiene el chat global desactivado?
                
                    'Si el flag era 0 lo restauramos a 1 y le avisamos al gm
                    UserList(tUser).flags.Global = 1
                    Call WriteConsoleMsg(UserIndex, "Has sido des-silenciado del chat global. Utilizalo con moderaci�n: " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_INFO)
                    
                    'Guardamos en el log del gm la acci�n
                    Call LogGM(.name, "Ha reestablecido el chat global de: " & UserList(tUser).name)
          
                    Call BanGlobalChatQuitar(UserList(tUser).name)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleToggleGlobal(ByVal UserIndex As Integer)
'***************************************************
'Author: MAB
'Declaraciones: Si queres vivir mejor, ponele un IF a tu vida
'***************************************************
On Error GoTo errHandler

With UserList(UserIndex)
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(.incomingData)
      
    'Remove packet ID
    Call buffer.ReadByte
  
    'Solo un Dios o un Admin puede activar/desactivar el global
    If .flags.Privilegios > PlayerType.Dios Or .flags.Privilegios > PlayerType.Admin Then
  
        'Si est� activo (que por defecto lo est�) entonces lo desactivamos y enviamos un mensaje global a todos los usuarios
        If GlobalChatActive = True Then
            GlobalChatActive = False
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El chat global ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
            
        Else
        
            'Si estaba deshabilitado, lo habilitamos e informamos a todos los usuarios
            GlobalChatActive = True
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El chat global fue activado.", FontTypeNames.FONTTYPE_SERVER))
        End If
  
    End If
  
    'If we got here then packet is complete, copy data back to original queue
    Call .incomingData.CopyBuffer(buffer)
End With

errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing
  
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub WriteSeeInProcess(ByVal UserIndex As Integer)
'***************************************************
'Author:Franco Emmanuel Giménez (Franeg95)
'Last Modification: 18/10/10
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.SeeInProcess)
 
Exit Sub
 
errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
     
Private Sub HandleSendProcessList(ByVal UserIndex As Integer)
'***************************************************
'Author: Franco Emmanuel Gimenez(Franeg95)
'Last Modification: 18/10/10
'***************************************************
 
    If UserList(UserIndex).incomingData.Length < 4 Then
       Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
       Exit Sub
    End If

On Error GoTo errHandler
    With UserList(UserIndex)
        
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
 
        Call buffer.ReadByte
        Dim Captions As String, Process As String
        
        Captions = buffer.ReadASCIIString()
        Process = buffer.ReadASCIIString()
        
        If .flags.GMRequested > 0 Then
            If UserList(.flags.GMRequested).ConnIDValida Then
                Call WriteShowProcess(.flags.GMRequested, Captions, Process)
                .flags.GMRequested = 0
            End If
        End If
        
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:    Dim Error As Long:     Error = Err.Number: On Error GoTo 0:   Set buffer = Nothing:    If Error <> 0 Then Err.Raise Error
End Sub
            
Private Sub HandleLookProcess(ByVal UserIndex As Integer)
'***************************************************
'Author: Franco Emmanuel Gimenez(Franeg95)
'Last Modification: 18/10/10
'***************************************************
 
On Error GoTo errHandler
    With UserList(UserIndex)
        
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
 
        Call buffer.ReadByte
        Dim data As String
        Dim tIndex As Integer
        
        data = buffer.ReadASCIIString()
        tIndex = NameIndex(data)
        
        'Solo los GMs pueden ver los procesos a los usuarios
        If EsGm(UserIndex) Then
            If tIndex > 0 And Not EsAdmin(data) Then
                UserList(tIndex).flags.GMRequested = UserIndex
                Call WriteSeeInProcess(tIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        Call .incomingData.CopyBuffer(buffer)
    End With
    
    Exit Sub
    
errHandler:
    LogError ("Error en HandleLookProcess. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Sub WriteShowProcess(ByVal gmIndex As Integer, ByVal strCaptions As String, ByVal strProcess As String)

    On Error GoTo errHandler

    With UserList(gmIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowProcess)
        Call .WriteASCIIString(strCaptions)
        Call .WriteASCIIString(strProcess)
    End With

    Exit Sub
errHandler:
    If Err.Number = UserList(gmIndex).outgoingData.NotEnoughSpaceErrCode Then Call FlushBuffer(gmIndex): Resume
End Sub

Private Sub HandleAccionInventario(ByVal UserIndex As Integer)
'***********************************
'Autor: Lorwik
'Fecha: 14/07/2020
'Descripcion: Recibimos una accion sobre el inventario �Que debemos hacer?
'***********************************
    Dim itemSlot As Byte
    Dim ObjIndex As Long
    
    With UserList(UserIndex)
    
        'Remove packet ID
        .incomingData.ReadByte
        
        itemSlot = .incomingData.ReadByte

        ObjIndex = .Invent.Object(itemSlot).ObjIndex

        If ObjIndex < 1 Then Exit Sub

        'Esta el user muerto y no esta usando una Runa?
        If .flags.Muerto = 1 And ObjData(ObjIndex).OBJType <> otRuna Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
        End If
        
        If .flags.Comerciando Then Exit Sub
        
        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If ObjIndex = 0 Then Exit Sub
        
        Select Case ObjData(ObjIndex).OBJType
        
            Case eOBJType.otCasco, eOBJType.otArmadura, eOBJType.otEscudo, eOBJType.otAnillo
                Call EquiparInvItem(UserIndex, itemSlot)
                
            Case eOBJType.otWeapon
                
                If ObjData(ObjIndex).proyectil Then '�Es un arco?
                    If .Invent.Object(itemSlot).Equipped Then '�Lo tiene ya equipado?
                        Call UseInvItem(UserIndex, itemSlot) 'Lo usamos
                        Exit Sub
                    End If
                End If
                
                'Equipamos el arma
                Call EquiparInvItem(UserIndex, itemSlot)
            
            Case Else
                Call UseInvItem(UserIndex, itemSlot)
                
        End Select
        
    End With

End Sub

''
' Prepares the "CreateParticleChar" message and returns it.

Public Function PrepareMessageCreateParticleChar(ByVal CharIndex As Integer, _
                                       ByVal ParticulaID As Integer, _
                                       ByVal Create As Boolean, _
                                       ByVal Life As Long) As String

    '***********************************
    'Autor: Lorwik
    'Fecha: 20/07/2020
    'Descripcion: Enviamos crear una particula en un char
    '***********************************

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharParticle)
        Call .WriteInteger(ParticulaID)
        Call .WriteBoolean(Create)
        Call .WriteInteger(CharIndex)
        Call .WriteLong(Life)
        
        PrepareMessageCreateParticleChar = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "CreateMapParticle" message and returns it.

Public Function PrepareMessageCreateMapParticle(ByVal ParticulaID As Integer, _
                                       ByVal X As Byte, _
                                       ByVal Y As Byte, _
                                       ByVal Life As Long) As String

    '***********************************
    'Autor: Lorwik
    'Fecha: 10/03/2020
    'Descripcion: Enviamos crear una particula en un char
    '***********************************

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharParticle)
        Call .WriteInteger(ParticulaID)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteLong(Life)
        
        PrepareMessageCreateMapParticle = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Sub HandleIniciaSubasta(ByVal UserIndex As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 25/05/2010
'***************************************************
    With UserList(UserIndex).incomingData
        'Remove Packet ID
        Call .ReadByte
        Call Iniciar_Subasta(UserIndex, .ReadInteger, .ReadInteger, .ReadLong)
    End With
End Sub

Public Sub HandleCancelarSubasta(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 19/08/2020
'Descripci�n: El user no subasta y cierra el form
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        .flags.Subastando = False
    End With
End Sub
 
Public Sub HandleConsultarSubasta(ByVal UserIndex As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 05/07/2010
'***************************************************
    With UserList(UserIndex).incomingData
        'Remove Packet ID
        Call .ReadByte
        Call Consultar_Subasta(UserIndex)
    End With
End Sub
 
Public Sub HandleOfertaSubasta(ByVal UserIndex As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 25/05/2010
'***************************************************
    With UserList(UserIndex).incomingData
        'Remove Packet ID
        Call .ReadByte
        Call Ofertar_Subasta(UserIndex, .ReadLong)
    End With
End Sub
 
Public Sub WriteIniciarSubastasOrConsulta(ByVal UserIndex As Integer)
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.IniciarSubastaConsulta)
    End With
End Sub

Public Sub WriteAbrirMapa(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 06/03/2021
'Descripcion: Envia al cliente la orden de abrir mapa
'***************************************************

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.AbriMapa)

End Sub

Public Sub WriteAbrirGoliath(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 31/03/2021
'Descripcion: Manda abrir la ventana de finanzas Goliath
'***************************************************

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.AbrirGoliath)
        
        Call .outgoingData.WriteLong(.Stats.Banco)
        
    End With

End Sub

Private Sub HandleCasament(ByVal UserIndex As Integer)
'@@ Autor: Cr3p-1
'@@ Fecha: 12/3/21
'@@ Descripci�n: Y mira.. Basicamente es el comando pa proponer matrimonio, si te rechaza jodete por virgen.
'@@ PD: Estoy muy drogado y ebrio perd�n, i love manu<3

'Fix Cr3p le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler

    With UserList(UserIndex)
        
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
          
        Call buffer.ReadByte    'ac� borraremos de la memoria el byte identificador.
        
        
        Dim nick As String
        Dim index As Integer
        
        nick = buffer.ReadASCIIString
        index = NameIndex(nick)
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
                Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                   
           'Validate target NPC
           ElseIf .flags.TargetNPC = 0 Then
               Call WriteConsoleMsg(UserIndex, "Primero ten�s que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
           
           ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
               Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
               
           ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor Then
           
           'ElseIf .Genero = UserList(Index).Genero Then (IAO CLASSIC APOYA EL MATRIMONIO IGUALITARIO YA QUE BARRIN ES GAY Y LE GUSTA QUE LE FLOREZCAN EL ANO)
           '    Call WriteConsoleMsg(UserIndex, "Personajes del mismo genero no pueden casarse.", FontTypeNames.FONTTYPE_TALK)
           
           ElseIf .flags.toyCasado = 1 Then
               Call WriteConsoleMsg(UserIndex, "�Ya estas casado con alguien!", FontTypeNames.FONTTYPE_TALK)
           
           ElseIf UserList(index).flags.Muerto = 1 Then
               Call WriteConsoleMsg(UserIndex, "Esta muerto!!!", FontTypeNames.FONTTYPE_TALK)
               
           ElseIf UserList(index).flags.yaOfreci = 1 Then
               Call WriteConsoleMsg(UserIndex, "Ya le ofrecieron", FontTypeNames.FONTTYPE_TALK)
           
           ElseIf UserList(index).flags.toyCasado = 1 Then
               Call WriteConsoleMsg(UserIndex, "Se encuentra casado !", FontTypeNames.FONTTYPE_TALK)
               
           Else
           
               Call WriteConsoleMsg(index, .name & " quiere casarse contigo, si aceptas escribe /ACEPTO " & .name, FontTypeNames.FONTTYPE_TALK)
               
               Call WriteConsoleMsg(UserIndex, "Le ofreciste casamiento a " & UserList(index).name, FontTypeNames.FONTTYPE_TALK)
               
               UserList(index).flags.yaOfreci = 1
               .flags.yaOfreci = 1
           
           End If
           
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With


errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleAcepto(ByVal UserIndex As Integer)
'@@ Autor: Cr3p-1
'@@ Fecha: 12/3/21
'@@ Descripci�n: Y mira.. Basicamente es el comando pa aceptar el matrimonio, sin el no podes garcharte a tu esposa si sos religioso. Pero si sos como nosotros garchas siempre sin casarte eso es lo bueno.
'@@ PD: Estoy muy drogado y ebrio perd�n, i love manu<3

'Fix Cr3p le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler

    With UserList(UserIndex)
     
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte    'ac� borraremos de la memoria el byte identificador.
    
         
        Dim nick As String
        Dim index As Integer
         
        nick = buffer.ReadASCIIString
        index = NameIndex(nick)
         
        If index <= 0 Then
             
        'Dead people can't leave a faction.. they can't talk...
        ElseIf .flags.Muerto = 1 Then
             Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
        
        'Validate target NPC
        ElseIf .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero ten�s que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        
        ElseIf UserList(index).flags.yaOfreci = 0 Then
            Call WriteConsoleMsg(UserIndex, "No te ofrecio matrimonio.", FontTypeNames.FONTTYPE_INFO)
        
        Else
        
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " y " & UserList(index).name & " se han unido en matrimonio!", FontTypeNames.FONTTYPE_TALK))
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(161, .Pos.X, .Pos.Y)) 'Casamiento
            
             If .Genero = UserList(index).Genero Then
             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " y " & UserList(index).name & " se han unido en matrimonio! Son una manga de putos ", FontTypeNames.FONTTYPE_TALK))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(161, .Pos.X, .Pos.Y)) 'Casamiento
            End If
            
            .flags.miPareja = UserList(index).name
            UserList(index).flags.miPareja = .name
            .flags.toyCasado = 1
            UserList(index).flags.toyCasado = 1
            
        
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
 
End Sub
 
Private Sub HandleDivorcio(ByVal UserIndex As Integer)
'@@ Autor: Cr3p-1
'@@ Fecha: 12/3/21
'@@ Descripci�n: Esto es para divorciarte, ya cuando tu mujer te tiene los huevos en compota, y la queres tirar al riachuelo pa ver el futbol tranquilo..
'@@ PD: Estoy muy drogado y ebrio perd�n, i love manu<3
'Fix Cr3p pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
 
    With UserList(UserIndex)
     
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte    'ac� borraremos de la memoria el byte identificador.
         
        Dim nick As String
        Dim index As Integer
         
        nick = buffer.ReadASCIIString
        index = NameIndex(nick)
         
        If index <= 0 Then
            Call WriteConsoleMsg(UserIndex, nick & " No esta online en este momento, intenta mas tarde, No la acoses que es feminista tarado..", FontTypeNames.FONTTYPE_TALK)
            
        'Dead people can't leave a faction.. they can't talk...
        ElseIf .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos pelotudo.", FontTypeNames.FONTTYPE_INFO)
        
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor Then
            Call WriteConsoleMsg(UserIndex, "Primero clickea al Sacerdote ciego de mierda", FontTypeNames.FONTTYPE_INFO)
         
        ElseIf .flags.toyCasado = 0 Then
            Call WriteConsoleMsg(UserIndex, "No est�s casado con nadie, No seas mentiroso. Virgen.", FontTypeNames.FONTTYPE_TALK)
         
        ElseIf UCase$(.flags.miPareja) <> UCase$(nick) Then
            Call WriteConsoleMsg(UserIndex, nick & " No es tu pareja. �Estas en pedo, pelotudo?", FontTypeNames.FONTTYPE_TALK)
        Else
            
            .flags.miPareja = ""
            UserList(index).flags.miPareja = ""
            .flags.toyCasado = 0
            UserList(index).flags.toyCasado = 0
            UserList(index).flags.yaOfreci = 0
            .flags.yaOfreci = 0
            
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error

End Sub

Private Sub HandleTransferenciaOro(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 31/03/2021
'Descripci�n: Envia a oro a otro usuario
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo errHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
           
        Dim username As String
        Dim tUser As Integer
        Dim CantOro As Long
        username = buffer.ReadASCIIString()
        CantOro = buffer.ReadLong()
             
        tUser = NameIndex(username)
               
        If tUser <= 0 Then
            Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            
            
        ElseIf UserList(UserIndex).Stats.Gld < CantOro Then
            Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad de oro.", FontTypeNames.FONTTYPE_CITIZEN)
            
        Else
            UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Banco - CantOro
            UserList(tUser).Stats.Gld = UserList(tUser).Stats.Banco + CantOro
               
            Call WriteConsoleMsg(UserIndex, "Has enviado " & CantOro & " monedas de oro a " & UserList(tUser).name & ".", FONTTYPE_INFO)
            Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te ha hecho una transferencia de " & CantOro & " monedas de oro. Las mismas estan disponibles en tu cuenta de finanzas Goliath.", FONTTYPE_INFO)
        End If
           
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
 
errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set buffer = Nothing
   
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleAdoptarFamiliar(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 11/05/2023
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo errHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
           
        Dim PetName As String
        Dim PetTipo As Byte
        
        PetTipo = buffer.ReadByte()
        PetName = buffer.ReadASCIIString()
             
        'Mandamos a crear el familiar
        If Not CreateFamiliarNewUser(UserIndex, .clase, PetName, PetTipo) Then
            Call WriteConsoleMsg(UserIndex, "Ha ocurrido un error al adoptar el familiar, intentelo mas tarde. Si el problema persiste contacte con un admin.", FONTTYPE_WARNING)
        End If
           
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
 
errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set buffer = Nothing
   
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Function verifyTimeStamp(ByVal ActualCount As Long, ByRef LastCount As Long, ByRef LastTick As Long, ByRef Iterations, ByVal UserIndex As Integer, ByVal PacketName As String, Optional ByVal DeltaThreshold As Long = 100, Optional ByVal MaxIterations As Long = 5, Optional ByVal CloseClient As Boolean = False) As Boolean
    
    Dim Ticks As Long, Delta As Long
    Ticks = GetTickCount
    
    Delta = (Ticks - LastTick)
    LastTick = Ticks

    'Controlamos secuencia para ver que no haya paquetes duplicados.
    If ActualCount <= LastCount Then
        Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("Paquete grabado: " & PacketName & " | Cuenta: " & UserList(UserIndex).AccountName & " | Ip: " & UserList(UserIndex).IP & " (Baneado automaticamente)", FontTypeNames.FONTTYPE_INFOBOLD))
        Call LogEdicionPaquete("El usuario " & UserList(UserIndex).name & " edit� el paquete " & PacketName & ".")
        Debug.Print "El usuario " & UserList(UserIndex).name & " edit� el paquete " & PacketName & "."
        LastCount = ActualCount
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    'controlamos speedhack/macro
    If Delta < DeltaThreshold Then
        Iterations = Iterations + 1
        If Iterations >= MaxIterations Then
            'Call WriteShowMessageBox(UserIndex, "Relajate and� a tomarte un t� con Gulfas.")
            verifyTimeStamp = False
            'Call LogMacroServidor("El usuario " & UserList(UserIndex).name & " iter� el paquete " & PacketName & " " & MaxIterations & " veces.")
            Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Control de macro---> El usuario " & UserList(UserIndex).name & "| Revisar --> " & PacketName & " (Env�os: " & MaxIterations & ").", FontTypeNames.FONTTYPE_INFOBOLD))
            'Call WriteCerrarleCliente(UserIndex)
            'Call CloseSocket(UserIndex)
            LastCount = ActualCount
            Iterations = 0
            Debug.Print "CIERRO CLIENTE"
        End If
        'Exit Function
    Else
        Iterations = 0
    End If
        
    verifyTimeStamp = True
    LastCount = ActualCount
End Function

Private Sub HandleSolicitarRank(ByVal UserIndex As Integer)
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteEnviaRank(UserIndex)
End Sub

Private Sub WriteEnviaRank(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim query As String
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.EnviarRanking)
        
        For i = 1 To 5
            Call .outgoingData.WriteASCIIString(Ranked(i).nombre)
            Call .outgoingData.WriteDouble(Ranked(i).ELO)
        Next i
        
        Call .outgoingData.WriteDouble(.Stats.ELO)
    
    End With
End Sub

Private Sub HandleBatallaPVP(ByVal UserIndex As Integer)

    Dim TipoDuelo As Byte

    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        TipoDuelo = .incomingData.ReadByte
        Select Case TipoDuelo
        
            Case 0 'Duelo
                Call modDuelos.EsperarOponenteDuelo(UserIndex)
            
            Case 1 'Arena de Rinkel
                Call modArenaRinkel.EntrarArenaRinkel(UserIndex)
                
            Case 2 'Plantes
                Call modPLantes.InscripcionPlantes(UserIndex)
                
            Case 50 'Comenzar Arena de Rinkel
                Call modArenaRinkel.Preparar(UserIndex)
        End Select
    End With
End Sub
