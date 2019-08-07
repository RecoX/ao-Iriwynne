Attribute VB_Name = "General"
Option Explicit

Public RestanteHappyHour As Long
Public AviseHappy As Boolean


Public RestanteViajero As Long
Public AviseViajero As Boolean

Global LeerNPCs As clsIniManager

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, _
                     Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************

    Dim CuerpoDesnudo As Integer

    With UserList(UserIndex)

        Select Case .Genero

        Case eGenero.Hombre

            Select Case .raza

            Case eRaza.Humano
                CuerpoDesnudo = 21

            Case eRaza.Drow
                CuerpoDesnudo = 32

            Case eRaza.Elfo
                CuerpoDesnudo = 210

            Case eRaza.Gnomo
                CuerpoDesnudo = 222

            Case eRaza.Enano
                CuerpoDesnudo = 53

            End Select

        Case eGenero.Mujer

            Select Case .raza

            Case eRaza.Humano
                CuerpoDesnudo = 39

            Case eRaza.Drow
                CuerpoDesnudo = 40

            Case eRaza.Elfo
                CuerpoDesnudo = 259

            Case eRaza.Gnomo
                CuerpoDesnudo = 260

            Case eRaza.Enano
                CuerpoDesnudo = 60

            End Select

        End Select

        If Mimetizado Then
            .CharMimetizado.Body = CuerpoDesnudo
        Else
            .Char.Body = CuerpoDesnudo

        End If

        .flags.Desnudo = 1

    End With

End Sub

Sub Bloquear(ByVal toMap As Boolean, _
             ByVal sndIndex As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal b As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'toMap = true -> Envia los datos a todo el mapa
'toMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
'***************************************************

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)

    End If

End Sub

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then

        With MapData(Map, X, Y)

            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
                (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
                (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
                .Graphic(2) = 0 Then
                HayAgua = True
            Else
                HayAgua = False

            End If

        End With

    Else
        HayAgua = False

    End If

End Function

Private Function HayLava(ByVal Map As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer) As Boolean

'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False

        End If

    Else
        HayLava = False

    End If

End Function

Sub LimpiarMundo()

'***************************************************
'Author: Unknown
'Last Modification: 05/09/2012 - ^[GS]^
'***************************************************
    On Error GoTo errhandler

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Preparando limpieza del mundo", FontTypeNames.FONTTYPE_INFOBOLD))

    If aLimpiarMundo.CantItems > 0 Then
        Call aLimpiarMundo.EraseAllItems

    End If

    Call SecurityIp.IpSecurityMantenimientoLista

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Se ha limpiado el mundo", FontTypeNames.FONTTYPE_INFOBOLD))

    Exit Sub

errhandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.description)

End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim k      As Long

    Dim npcNames() As String

    ReDim npcNames(1 To UBound(SpawnList)) As String

    For k = 1 To UBound(SpawnList)
        npcNames(k) = SpawnList(k).NPCNAME
    Next k

    Call WriteSpawnList(UserIndex, npcNames())

End Sub

Sub Main()
'***************************************************
'Author: Unknown
'Last Modification: 15/03/2011
'15/03/2011: ZaMa - Modularice todo, para que quede mas claro.
'***************************************************

    On Error Resume Next

    '  Dim dh, d1, d2, d3, d4 As String

    ' d2 = "06"
    ' d3 = "20"
    ' d1 = "15"
    ' d4 = "18"

    '  dh = d1 & "/" & d2 & "/" & d3 & d4

    ' If DateDiff("d", Date, dh) <= 0 Then End

    ChDir App.Path
    ChDrive App.Path

    Call LoadMotd
    Call BanIpCargar
    Call BanT0Cargar

    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision

    ' Start loading..
    frmCargando.Show

    ' Constants & vars
    frmCargando.Label1(2).Caption = "Cargando constantes..."
    Call LoadConstants
    Call CargarConfiguracionHAO
    Call CargarMsgs
    Call CargarSponsor
    Call CargarExperiencia
    Call CargarIntervalos

    DoEvents

    ' Arrays
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    Call LoadArrays

    ' Server.ini & Apuestas.dat
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    Call LoadSini
    Call CargaApuestas

    ' Npcs.dat
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat

    ' Obj.dat
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    Call LoadOBJData
    Call LoadCanjesData

    ' Hechizos.dat
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos

    ' Objetos de Herreria
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herrer�a"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria

    ' Objetos de Capinteria
    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpinter�a"
    Call LoadObjCarpintero

    ' Balance.dat
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance

    ' Armaduras faccionarias
    frmCargando.Label1(2).Caption = "Cargando ArmadurasFaccionarias.dat"
    Call LoadArmadurasFaccion

    ' Animaciones
    frmCargando.Label1(2).Caption = "Cargando Animaciones"
    Call LoadAnimations

    ' Pretorianos
    frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
    Call LoadPretorianData

    Torneo1vs1_CargarPos

    ' Mapas
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData

    End If

    ' Map Sounds
    Set SonidosMapas = New SoundMapInfo
    Call SonidosMapas.LoadSoundMapInfo

    ' Connections
    Call ResetUsersConnections

    ' Timers
    Call InitMainTimers

    ' Sockets
    Call SocketConfig

    centinelaActivado = 1

    ' End loading..
    Unload frmCargando

    'Log start time
    LogServerStartTime

    Call LoadRanking


1   Call m_JuegosDelHambre.Carga_JDH


    If FileExist(AntiFragPath) Then Call Kill(AntiFragPath)


    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)

    End If

    tInicioServer = GetTickCount() And &H7FFFFFFF

    LoadCastillos

    IniciarCastilloPretoriano
End Sub

Private Sub LoadConstants()

'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Loads all constants and general parameters.
'*****************************************************************
    On Error Resume Next

    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")

    ' Paths
    IniPath = App.Path & "\"
    DatPath = App.Path & "\Dat\"
    CharPath = App.Path & "\Charfile\"


    AntiFragPath = App.Path & "\Logs\ZZZ_AntiFrag.log"


    ' Skills by level
    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100

    ' Races
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ' Classes
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"

    ' Skills
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasi�n en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apu�alar) = "Apu�alar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    ' Attributes
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"

    ' Fishes
    ListaPeces(1) = PECES_POSIBLES.PESCADO1
    ListaPeces(2) = PECES_POSIBLES.PESCADO2
    ListaPeces(3) = PECES_POSIBLES.PESCADO3
    ListaPeces(4) = PECES_POSIBLES.PESCADO4

    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)

    Set Ayuda = New cCola
    Set Denuncias = New cCola
    Denuncias.MaxLenght = MAX_DENOUNCES

    MaxUsers = 0

    ' Initialize classes
    Protocol.InitAuxiliarBuffer

    Set aClon = New clsAntiMassClon
    Set aLimpiarMundo = New clsLimpiarMundo

End Sub

Private Sub LoadArrays()

'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Loads all arrays
'*****************************************************************
    On Error Resume Next

    ' Load Records
    Call LoadRecords

    ' Load guilds info
    Call LoadGuildsDB

    ' Load spawn list
    Call CargarSpawnList

    ' Load forbidden words
    Call CargarForbidenWords


    Call LoadQuests    ' GSZAO
    ' Retos 1 vs 1
    Call retos1vs1Load

    ' Retos 2 vs 2
    Call retos2vs2Load

End Sub

Private Sub ResetUsersConnections()

'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Resets Users Connections.
'*****************************************************************
    On Error Resume Next

    Dim LoopC  As Long

    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC

End Sub

Private Sub InitMainTimers()

'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Initializes Main Timers.
'*****************************************************************
    On Error Resume Next

    With frmMain
        .AutoSave.Enabled = True
        .tLluvia.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .tLluviaEvent.Enabled = True
        .FX.Enabled = True
        .Auditoria.Enabled = True
        .TIMER_AI.Enabled = True
        .npcataca.Enabled = True

    End With

End Sub

Private Sub SocketConfig()

'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Sets socket config.
'*****************************************************************
    On Error Resume Next

    Call SecurityIp.InitIpTables(1000)

    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

End Sub

Private Sub LogServerStartTime()

'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Logs Server Start Time.
'*****************************************************************
    Dim N      As Integer

    N = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #N
    Print #N, Date & " " & time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #N

End Sub

Function FileExist(ByVal file As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************

    FileExist = LenB(dir$(file, FileType)) <> 0

End Function

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String
'*****************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************

    Dim i      As Long

    Dim lastPos As Long

    Dim CurrentPos As Long

    Dim delimiter As String * 1

    delimiter = Chr$(SepASCII)

    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i

    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)

    End If

End Function

Function MapaValido(ByVal Map As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    MapaValido = Map >= 1 And Map <= NumMaps

End Function

Sub MostrarNumUsers()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim UsersMod As Integer

    UsersMod = NumUsers
    ' UsersMod = UsersMod * 2.1 ' kevin counter
    frmMain.txtNumUsers.Text = UsersMod

End Sub

Public Sub LogCriticEvent(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogIndex(ByVal Index As Integer, ByVal desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogError(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogStatic(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogTarea(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile(1)    ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogClanes(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & str
    Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & str
    Close #nfile

End Sub

Public Sub LogItemsEspeciales(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\ItemsEspeciales\ItemsEspeciales" & Month(Date$) & " - " & Year(Date$) & ".log" For Append Shared As #nfile
    Print #nfile, Date$ & " " & time$ & " " & str
    Close #nfile

End Sub

Public Sub LogDesarrollo(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & str
    Close #nfile

End Sub

Public Sub LogGM(Nombre As String, texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************�

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile

    Exit Sub

errhandler:

End Sub
Public Sub logCOSO()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************�

    On Error GoTo errhandler

    Dim nfile  As Integer

1   If Not FileExist(App.Path & "\logs\ServerUP.txt", vbArchive) Then
2       nfile = FreeFile    ' obtenemos un canal
3       'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
4       Open App.Path & "\logs\ServerUP.txt" For Append Shared As #nfile
5       Print #nfile, "[INIT]"
6       Print #nfile, "UP=1"
7       Close #nfile
8       Exit Sub
9   End If

10  WriteVar App.Path & "\logs\ServerUP.txt", "INIT", "UP", 1



11  Exit Sub

errhandler:

    LogError "ERROR EN LOGCOSO en linea " & Erl & " - " & Err.Number & " " & Err.description

End Sub

Public Sub LogAsesinato(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal

    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub logVentaCasa(ByVal texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal

    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim nfile  As Integer

    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, vbNullString
    Close #nfile

    Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Arg    As String

    Dim i      As Integer

    For i = 1 To 33

        Arg = ReadField(i, cad, 44)

        If LenB(Arg) = 0 Then Exit Function

    Next i

    ValidInputNP = True

End Function

Sub Restart()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

    Dim LoopC  As Long

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)

    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next

    'Initialize statistics!!
    Call Statistics.Initialize

    For LoopC = 1 To UBound(UserList())
        Set UserList(LoopC).incomingData = Nothing
        Set UserList(LoopC).outgoingData = Nothing
    Next LoopC

    ReDim UserList(1 To MaxUsers) As User

    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC

    LastUser = 0
    NumUsers = 0

    Call FreeNPCs
    Call FreeCharIndexes

    Call LoadSini

    Call ResetForums
    Call LoadOBJData

    Call LoadMapData

    Call CargarHechizos

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

    'Log it
    Dim N      As Integer

    N = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #N
    Print #N, Date & " " & time & " servidor reiniciado."
    Close #N

    'Ocultar

    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)

    End If

End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
'**************************************************************
'Author: Unknown
'Last Modify Date: 15/11/2009
'15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
'23/11/2009: ZaMa - Optimizacion de codigo.
'**************************************************************

    With UserList(UserIndex)

        If MapInfo(.Pos.Map).Zona <> "DUNGEON" Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 1 And _
               MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 2 And _
               MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 4 Then Intemperie = True
        Else
            Intemperie = False

        End If

    End With

    'En las arenas no te afecta la lluvia
    If IsArena(UserIndex) Then Intemperie = False

End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    If UserList(UserIndex).flags.UserLogged Then
        If Intemperie(UserIndex) Then
            If UserList(UserIndex).Pos.Map = 1 Then
                Dim modifi As Long
                modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 3)
                Call QuitarSta(UserIndex, modifi)
                Call FlushBuffer(UserIndex)
            End If
        End If
    End If
    Exit Sub
errhandler:
    LogError ("Error en EfectoLluvia")

End Sub

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i      As Integer

    For i = 1 To MAXMASCOTAS

        With UserList(UserIndex)

            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = _
                    Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia - 1

                    If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(.MascotasIndex(i), 0)

                End If

            End If

        End With

    Next i

End Sub
Public Sub EfectoPolo(ByVal UserIndex As Integer)

    Dim vida   As Integer

    With UserList(UserIndex)

        If .Invent.ArmourEqpObjIndex = 1086 Then Exit Sub    'Si tiene equipada la t�nica invernal no pasa nada.

        If .Invent.ArmourEqpObjIndex = 1087 Then Exit Sub    'Si tiene equipada la t�nica invernal no pasa nada.

        If .Counters.Congelado < IntervaloFrio Then    'Usamos el mismo intervalo que el fr�o.
            .Counters.Congelado = .Counters.Congelado + 1
        Else
            If .Pos.Map = 127 Or .Pos.Map = 128 Then    'Si est� en X mapa entonces:
                
                Call WriteConsoleMsg(UserIndex, "��Ponte tu vestimenta polar o morir�s congelado!!", FontTypeNames.FONTTYPE_INFO)
                vida = Porcentaje(.Stats.MaxHP, 15)
                .Stats.MinHp = .Stats.MinHp - vida

                If .Stats.MinHp < 1 Then    'Si la vida lleg� a 0 se muere
                    Call WriteConsoleMsg(UserIndex, "��Has muerto congelado!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)
                End If

                Call WriteUpdateHP(UserIndex)    'Actualizamos la vida.
            End If

            .Counters.Congelado = 0
        End If
        
    End With
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)

'***************************************************
'Autor: Unkonwn
'Last Modification: 23/11/2009
'If user is naked and it's in a cold map, take health points from him
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    Dim modifi As Integer

    With UserList(UserIndex)

        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
        Else

            If MapInfo(.Pos.Map).Terreno = eTerrain.terrain_nieve Then
                Call WriteConsoleMsg(UserIndex, "��Est�s muriendo de fr�o, abrigate o morir�s!!", FontTypeNames.FONTTYPE_INFO)
                modifi = Porcentaje(.Stats.MaxHP, 5)
                .Stats.MinHp = .Stats.MinHp - modifi

                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(UserIndex, "��Has muerto de fr�o!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)

                End If

                Call WriteUpdateHP(UserIndex)
            Else
                modifi = Porcentaje(.Stats.MaxSta, 5)
                Call QuitarSta(UserIndex, modifi)
                Call WriteUpdateSta(UserIndex)

            End If

            .Counters.Frio = 0

        End If

    End With

End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)

'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 23/11/2009
'If user is standing on lava, take health points from him
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    With UserList(UserIndex)

        If .Counters.Lava < IntervaloFrio Then    'Usamos el mismo intervalo que el del frio
            .Counters.Lava = .Counters.Lava + 1
        Else

            If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
                Call WriteConsoleMsg(UserIndex, "��Quitate de la lava, te est�s quemando!!", FontTypeNames.FONTTYPE_INFO)
                .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHP, 5)

                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(UserIndex, "��Has muerto quemado!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)

                End If

                Call WriteUpdateHP(UserIndex)

            End If

            .Counters.Lava = 0

        End If

    End With

End Sub

''
' Maneja  el efecto del estado atacable
'
' @param UserIndex  El index del usuario a ser afectado por el estado atacable
'

Public Sub EfectoEstadoAtacable(ByVal UserIndex As Integer)
'******************************************************
'Author: ZaMa
'Last Update: 18/09/2010 (ZaMa)
'18/09/2010: ZaMa - Ahora se activa el seguro cuando dejas de ser atacable.
'******************************************************

' Si ya paso el tiempo de penalizacion
    If Not IntervaloEstadoAtacable(UserIndex) Then
        ' Deja de poder ser atacado
        UserList(UserIndex).flags.AtacablePor = 0

        ' Activo el seguro si deja de estar atacable
        If Not UserList(UserIndex).flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn)

        End If

        ' Send nick normal
        Call RefreshCharStatus(UserIndex)

    End If

End Sub

''
' Maneja el tiempo de arrivo al hogar
'
' @param UserIndex  El index del usuario a ser afectado por el /hogar
'

Public Sub TravelingEffect(ByVal UserIndex As Integer)
'******************************************************
'Author: ZaMa
'Last Update: 01/06/2010 (ZaMa)
'******************************************************

' Si ya paso el tiempo de penalizacion
    If IntervaloGoHome(UserIndex) Then
        Call HomeArrival(UserIndex)

    End If

End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)

'******************************************************
'Author: Unknown
'Last Update: 16/09/2010 (ZaMa)
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'16/09/2010: ZaMa - Se recupera la apariencia de la barca correspondiente despues de terminado el mimetismo.
'******************************************************
    Dim Barco  As ObjData

    With UserList(UserIndex)

        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)

            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    Call ToggleBoatBody(UserIndex)
                Else
                    .Char.Body = iFragataFantasmal
                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                    .Char.CascoAnim = NingunCasco

                End If

            Else
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim

            End If

            With .Char
                Call ChangeUserChar(UserIndex, .Body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

            End With

            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Se fue el efecto del mimetismo, puede ser atacado por npcs
            .flags.Ignorado = False

        End If

    End With

End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/09/2010 (ZaMa)
'16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
'***************************************************

    With UserList(UserIndex)

        Debug.Print .Counters.Invisibilidad
        If .Counters.Invisibilidad < 310 Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + 1
        Else
            .Counters.Invisibilidad = -1    'RandomNumber(-100, 100) ' Invi variable :D
            .flags.invisible = 0
            WriteInviSegundos UserIndex, 0
            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)

                ' Si navega ya esta visible..
                If Not .flags.Navegando = 1 Then

                    'Si est� en un oscuro no lo hacemos visible
                    If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.zonaOscura Then
                        Call SetInvisible(UserIndex, .Char.CharIndex, False)

                    End If

                End If

            End If

        End If

    End With

End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)

        If .Contadores.Paralisis > 0 Then
            .Contadores.Paralisis = .Contadores.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0

        End If

    End With

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)

        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - 1
        Else

            If .flags.Ceguera = 1 Then
                .flags.Ceguera = 0
                Call WriteBlindNoMore(UserIndex)

            End If

            If .flags.Estupidez = 1 Then
                .flags.Estupidez = 0
                Call WriteDumbNoMore(UserIndex)

            End If

        End If

    End With

End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/12/2010
'02/12/2010: ZaMa - Now non-magic clases lose paralisis effect under certain circunstances.
'***************************************************

    With UserList(UserIndex)

        If .Counters.Paralisis > 0 Then

            Dim CasterIndex As Integer

            CasterIndex = .flags.ParalizedByIndex

            ' Only aplies to non-magic clases
            If .Stats.MaxMAN = 0 Then

                ' Paralized by user?
                If CasterIndex <> 0 Then

                    ' Close? => Remove Paralisis
                    If UserList(CasterIndex).Name <> .flags.ParalizedBy Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub

                        ' Caster dead? => Remove Paralisis
                    ElseIf UserList(CasterIndex).flags.Muerto = 1 Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub

                    ElseIf .Counters.Paralisis > IntervaloParalizadoReducido Then

                        ' Out of vision range? => Reduce paralisis counter
                        If Not InVisionRangeAndMap(UserIndex, UserList(CasterIndex).Pos) Then
                            ' Aprox. 1500 ms
                            .Counters.Paralisis = IntervaloParalizadoReducido
                            Exit Sub

                        End If

                    End If

                    ' Npc?
                Else
                    CasterIndex = .flags.ParalizedByNpcIndex

                    ' Paralized by npc?
                    If CasterIndex <> 0 Then

                        If .Counters.Paralisis > IntervaloParalizadoReducido Then

                            ' Out of vision range? => Reduce paralisis counter
                            If Not InVisionRangeAndMap(UserIndex, Npclist(CasterIndex).Pos) Then
                                ' Aprox. 1500 ms
                                .Counters.Paralisis = IntervaloParalizadoReducido
                                Exit Sub

                            End If

                        End If

                    End If

                End If

            End If

            .Counters.Paralisis = .Counters.Paralisis - 1

        Else
            Call RemoveParalisis(UserIndex)

        End If

    End With

End Sub

Public Sub RemoveParalisis(ByVal UserIndex As Integer)

'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Removes paralisis effect from user.
'***************************************************
    With UserList(UserIndex)
        .flags.Paralizado = 0
        .flags.Inmovilizado = 0
        .flags.ParalizedBy = vbNullString
        .flags.ParalizedByIndex = 0
        .flags.ParalizedByNpcIndex = 0
        .Counters.Paralisis = 0
        Call WriteParalizeOK(UserIndex)

    End With

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, _
                      ByRef EnviarStats As Boolean, _
                      ByVal Intervalo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 And _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 2 And _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub

        Dim massta As Integer

        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + 1
            Else
                EnviarStats = True
                .Counters.STACounter = 0

                If .flags.Desnudo Then Exit Sub    'Desnudo no sube energ�a. (ToxicWaste)

                massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))
                .Stats.MinSta = .Stats.MinSta + massta

                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta

                End If

            End If

        End If

    End With

End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N      As Integer

    With UserList(UserIndex)

        If .Counters.Veneno < IntervaloVeneno Then
            .Counters.Veneno = .Counters.Veneno + 1
        Else
            Call WriteConsoleMsg(UserIndex, "Est�s envenenado, si no te curas morir�s.", FontTypeNames.FONTTYPE_VENENO)
            .Counters.Veneno = 0
            N = RandomNumber(1, 5)
            .Stats.MinHp = .Stats.MinHp - N

            If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)

        End If

    End With

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)
'***************************************************
'Author: ??????
'Last Modification: 08/06/11 (CHOTS)
'Le agregu� que avise antes cuando se te est� por ir
'
'Cuando se pierde el efecto de la poci�n updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
'***************************************************

    Const SEGUNDOS_AVISO As Byte = 5

    'CHOTS | Los segundos antes que se te acabe que te avisa

    With UserList(UserIndex)

        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
            .flags.DuracionEfecto = .flags.DuracionEfecto - 1

            If ((.flags.DuracionEfecto / 25) <= SEGUNDOS_AVISO) And (.flags.UltimoMensaje <> 221) Then    'CHOTS | Lo divide por 25 por el intervalo del Timer (40x25=1000=1seg)
                Call WriteStrDextRunningOut(UserIndex)
                .flags.UltimoMensaje = 221

            End If

            If .flags.DuracionEfecto = 0 Then
                .flags.UltimoMensaje = 222
                .flags.TomoPocion = False
                .flags.TipoPocion = 0

                'volvemos los atributos al estado normal
                Dim LoopX As Integer

                For LoopX = 1 To NUMATRIBUTOS
                    .Stats.UserAtributos(LoopX) = .Stats.UserAtributosBackUP(LoopX)
                Next LoopX

                Call WriteUpdateStrenghtAndDexterity(UserIndex)

            End If

        End If

    End With

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)

        If Not .flags.Privilegios And PlayerType.User Then Exit Sub

        'Sed
        If .Stats.MinAGU > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + 1
            Else
                .Counters.AGUACounter = 0
                .Stats.MinAGU = .Stats.MinAGU - 10

                If .Stats.MinAGU <= 0 Then
                    .Stats.MinAGU = 0
                    .flags.Sed = 1

                End If

                fenviarAyS = True

            End If

        End If

        'hambre
        If .Stats.MinHam > 0 Then
            If .Counters.COMCounter < IntervaloHambre Then
                .Counters.COMCounter = .Counters.COMCounter + 1
            Else
                .Counters.COMCounter = 0
                .Stats.MinHam = .Stats.MinHam - 10

                If .Stats.MinHam <= 0 Then
                    .Stats.MinHam = 0
                    .flags.Hambre = 1

                End If

                fenviarAyS = True

            End If

        End If

    End With

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, _
                 ByRef EnviarStats As Boolean, _
                 ByVal Intervalo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 And _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 2 And _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub

        Dim mashit As Integer

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
        If .Stats.MinHp < .Stats.MaxHP Then
            If .Counters.HPCounter < Intervalo Then
                .Counters.HPCounter = .Counters.HPCounter + 1
            Else
                mashit = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))

                .Counters.HPCounter = 0
                .Stats.MinHp = .Stats.MinHp + mashit

                If .Stats.MinHp > .Stats.MaxHP Then .Stats.MinHp = .Stats.MaxHP
                Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
                EnviarStats = True

            End If

        End If

    End With

End Sub

Public Sub CargaNpcsDat()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim npcfile As String

    npcfile = DatPath & "NPCs.dat"
    Set LeerNPCs = New clsIniManager
    Call LeerNPCs.Initialize(npcfile)

End Sub

Sub PasarSegundo()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    Dim i      As Long


    Call Mod_Retos1vs1.reto_all_loop
    Call Mod_Retos2vs2.loop_reto

    If DeathMatch.Activo Then
        Call modDeath.Cuenta
    End If

    If JDH.Activo Then
        CuentaJDH
    End If

    'If Torneo1vs1.AutoCancelTime > 0 Then
    Call Mod_Torneo1vs1.Cuenta
    'End If



    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then

            Call Mod_Retos1vs1.loop_userReto(i)
            Call Mod_Retos2vs2.user_retoLoop(i)
            
            
            If UserList(i).flags.YaPediAyudaCount > 0 Then
                UserList(i).flags.YaPediAyudaCount = UserList(i).flags.YaPediAyudaCount - 1
                
                If UserList(i).flags.YaPediAyudaCount <= 0 Then
                    UserList(i).flags.YaPediAyuda = False
                    If UserList(i).GuildIndex > 0 Then
                        Call SendData(SendTarget.ToDiosesYclan, UserList(i).GuildIndex, PrepareMessageAyudaClan(i))
                    End If
                End If
            End If

            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1

                If UserList(i).Counters.Salir <= 0 Then
                    Call WriteConsoleMsg(i, "Gracias por jugar Iriwynne AO", FontTypeNames.FONTTYPE_INFO)
                    Call WriteDisconnect(i)
                    Call FlushBuffer(i)

                    Call CloseSocket(i)

                Else
                    ' @@ Miqueas El conteo de segundos siempre me gusto mucho ajaja
                    Call WriteConsoleMsg(i, "Cerrando juego en: " & CStr(UserList(i).Counters.Salir), FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'Cuenta regresiva
        If CountdownTime > 1 Then
            CountdownTime = CountdownTime - 1
            If CountdownMap > 0 Then

                If CountdownMap = Configuracion.Mapa1vs1 Or CountdownMap = Configuracion.Mapa2vs2 Then
                    Call SendData(SendTarget.toMap, CountdownMap, PrepareMessageConsoleMsg("Retos>" & CountdownTime, FontTypeNames.FONTTYPE_GUILD))
                Else
                    Call SendData(SendTarget.toMap, CountdownMap, PrepareMessageConsoleMsg("Conteo>" & CountdownTime, FontTypeNames.FONTTYPE_GUILD))
                End If
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Conteo>" & CountdownTime, FontTypeNames.FONTTYPE_GUILD))
            End If
        ElseIf CountdownTime = 1 Then
            If CountdownMap > 0 Then

                If CountdownMap = Configuracion.Mapa1vs1 Or CountdownMap = Configuracion.Mapa2vs2 Then
                    Call SendData(SendTarget.toMap, CountdownMap, PrepareMessageConsoleMsg("Retos> �YA!", FontTypeNames.FONTTYPE_FIGHT))
                Else
                    Call SendData(SendTarget.toMap, CountdownMap, PrepareMessageConsoleMsg("�YA!", FontTypeNames.FONTTYPE_FIGHT))
                End If

            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("�YA!", FontTypeNames.FONTTYPE_FIGHT))
            End If
            CountdownTime = 0
            CountdownMap = 0
        End If

    Next i

    ' HappyHour
    If (time > "08:00:00 p.m." And time < "08:15:00 p.m.") Or (time > "08:00:00" And time < "08:15:00") Then
        HappyHour = True
        If AviseHappy = False Then
            RestanteHappyHour = 15
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("HappyHour> Se ha activado la HappyHour durante 15 minutos!", FontTypeNames.FONTTYPE_AMARILLO))
            AviseHappy = True
        End If
    Else
        If AviseHappy Then
            RestanteHappyHour = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("HappyHour> Ha terminado!", FontTypeNames.FONTTYPE_AMARILLO))
        End If
        HappyHour = False
        AviseHappy = False
    End If


95  If (time = "02:55:00") Or (time = "02:55:00 p.m.") Then
96      Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Viajero>", "En 5 minutos llegara el npc viajero a polo norte mapa 128 con varios items de valor", FontTypeNames.FONTTYPE_AMARILLO))
97  End If

    ' Viajero
10451 If (time > "03:00:00" And time < "03:03:00") Or (time > "03:00:00 p.m." And time < "03:03:00 p.m.") Then
1026    ViajeroActivo = True
1503    If AviseViajero = False Then
1304        RestanteViajero = 3
1025        m_Viajero.MeterNPCViajero
1066        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Viajero>", "Ah llegado el viajero a estas tierras, se encuentra en el mapa Polonorte, en el muelle del mapa 128.", FontTypeNames.FONTTYPE_AMARILLO))
1078        AviseViajero = True
108     End If
109 Else
110     If AviseViajero Then
111         RestanteViajero = 0
112         m_Viajero.QuitarNPCViajero
113         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Viajero>", "El viajero se fue!", FontTypeNames.FONTTYPE_AMARILLO))
115     End If
114
        ViajeroActivo = False
        AviseViajero = False
    End If



98  If (time = "05:55:00") Or (time = "05:55:00 p.m.") Then
99      Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Viajero>", "En 5 minutos llegara el npc viajero a polo norte mapa 128 con varios items de valor", FontTypeNames.FONTTYPE_AMARILLO))
100 End If

    ' Viajero
    If (time > "06:00:00" And time < "06:03:00") Or (time > "06:00:00 p.m." And time < "06:03:00 p.m.") Then
        ViajeroActivo = True
        If AviseViajero = False Then
            RestanteViajero = 3
            m_Viajero.MeterNPCViajero
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Viajero>", "Ah llegado el viajero a estas tierras, se encuentra en el mapa Polonorte, en el muelle del mapa 128.", FontTypeNames.FONTTYPE_AMARILLO))
            AviseViajero = True
        End If
    Else
        If AviseViajero Then
            RestanteViajero = 0
            m_Viajero.QuitarNPCViajero
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Viajero>", "El viajero se fue!", FontTypeNames.FONTTYPE_AMARILLO))
        End If

        ViajeroActivo = False
        AviseViajero = False
    End If


434:

    Exit Sub

errhandler:

    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)

    Resume Next

End Sub

Public Function ReiniciarAutoUpdate() As Double
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function

Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'WorldSave
    Call ES.DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios

    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

Sub GuardarUsuarios()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

12  haciendoBK = True

11  Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
10  Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))

    Dim i      As Integer

9   For i = 1 To LastUser

8       If UserList(i).flags.UserLogged Then
7           Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr", False)

6       End If

5   Next i

    'se guardan los seguimientos
4   Call SaveRecords

3   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
2   Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

1   haciendoBK = False
    Exit Sub

errhandler:
    LogError "Error en GuardarUsuarios en linea " & Erl & " - " & Err.Number & " " & Err.description

End Sub

Public Sub FreeNPCs()

'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
    Dim LoopC  As Long

    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC

End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))

End Sub
