Attribute VB_Name = "m_JuegosDelHambre"
Option Explicit


Type JDHUser
    UserIndex As Integer       'UI del usuario.
    LastPosition As WorldPos      'Pos que estaba antes de entrar.
End Type

Private Type tCofres
    Objetos(1 To 7) As Obj
    X As Byte
    Y As Byte
    Abierto As Boolean
End Type

Type tJDH
    Cofres(1 To 12) As tCofres
    Mapa As Integer
    Cupos As Byte          'Cantidad de cupos.
    Ingresaron As Byte          'Cantidad que ingreso.
    UsUaRiOs() As JDHUser     'Tipo de usuarios
    Cuenta As Byte          'Cuenta regresiva.
    Activo As Boolean       'Hay deathmatch
    Ganador As JDHUser     'Datos del ganador.
    Inscripcion As Long
    Premio As Long
    EventStarted As Boolean
End Type

Private Const CUENTA_NUM As Byte = 5     'Segundos de cuenta.
Private Const ARENA_X As Byte = 50   'X de la arena(se suma por usuario)
Private Const ARENA_Y As Byte = 50   'Y de la arena.

'Private Const TIEMPO_AUTOCANCEL As Byte = 120     '2 Minutos antes del auto-cancel.

Private Const Cofre_Abierto As Byte = 10    'N�mero de cofre abierto.
Public Cofre_Cerrado As Integer    'N�mero de cofre cerrado.

Private Const TAG_EVENT As String = "Juegos del Hambre>"

Public JDH As tJDH

Public Sub Carga_JDH()

    On Error GoTo errhandleR

    Dim LoopX  As Long
    Dim LoopZ  As Long

    Dim DataCofre As Obj

    Dim Leer   As clsIniManager
    Set Leer = New clsIniManager
56  Call Leer.Initialize(App.Path & "\Dat\JuegosDelHambre.dat")



3   Cofre_Cerrado = CInt(Leer.GetValue("EVENTO", "COFREObj"))


2   DataCofre.Amount = 1
1   DataCofre.objIndex = Cofre_Cerrado


    With JDH

125     .Mapa = CInt(Leer.GetValue("EVENTO", "Mapa"))

173     For LoopX = 1 To UBound(.Cofres())
31          .Cofres(LoopX).X = CByte(Leer.GetValue("COFRE#" & LoopX, "X"))
12          .Cofres(LoopX).Y = CByte(Leer.GetValue("COFRE#" & LoopX, "Y"))

11          Call MakeObj(DataCofre, .Mapa, .Cofres(LoopX).X, .Cofres(LoopX).Y)

10          MapData(.Mapa, .Cofres(LoopX).X, .Cofres(LoopX).Y).Blocked = 1

9           MapData(.Mapa, .Cofres(LoopX).X, .Cofres(LoopX).Y).Cofre = LoopX

8           Call Bloquear(True, .Mapa, .Cofres(LoopX).X, .Cofres(LoopX).Y, True)

            For LoopZ = 1 To UBound(.Cofres(LoopX).Objetos())

7               .Cofres(LoopX).Objetos(LoopZ).objIndex = CInt(ReadField(1, (Leer.GetValue("COFRE#" & LoopX, "OBJETO#" & LoopZ)), 45))
6               .Cofres(LoopX).Objetos(LoopZ).Amount = CInt(ReadField(2, (Leer.GetValue("COFRE#" & LoopX, "OBJETO#" & LoopZ)), 45))
            Next LoopZ

5       Next LoopX

4   End With

    Exit Sub
errhandleR:
    LogError "error en cargajdh en " & Erl & ". err " & Err.Number & " " & Err.description & " - LoopX: " & LoopX & " - LoopZ: " & LoopZ
End Sub

Sub LimpiarJDH()

On Error GoTo errhandleR

' @ Limpia los datos anteriores.

    Dim DumpPos As WorldPos
    Dim LoopX As Long
    Dim LoopY As Long

1    With JDH
2        .Cuenta = 0
3        .Cupos = 0
4        .Ingresaron = 0
5        .Activo = False
6        .Inscripcion = 0
7        .Premio = 0
8        .EventStarted = False
        
        'Limpio el tipo de ganador.
9        With .Ganador
10            .UserIndex = 0
11            .LastPosition = DumpPos
12        End With

        'Recargamos los cofres
        'Call ReCargar_Cofres
13        Dim bIsExit As Boolean
        'Limpia los objetos que quedaron tira2.
14        For LoopX = 1 To 100
15            For LoopY = 1 To 100
16                With MapData(JDH.Mapa, LoopX, LoopY).ObjInfo
17                    'Hay objeto?
18                    If MapData(JDH.Mapa, LoopX, LoopY).ObjInfo.objIndex > 0 And MapData(JDH.Mapa, LoopX, LoopY).Blocked = 0 Then
19                        'No es del mapa.
20                        bIsExit = (MapData(JDH.Mapa, LoopX, LoopY).TileExit.Map > 0)

21                        If ItemNoEsDeMapa(MapData(JDH.Mapa, LoopX, LoopY).ObjInfo.objIndex, bIsExit) Then
22                            'Erase
23                            Call EraseObj(.Amount, JDH.Mapa, LoopX, LoopY)
24                        End If
                    End If

                    ' @@ Si es un cofre y esta bloqueado
25                    If MapData(JDH.Mapa, LoopX, LoopY).Cofre And MapData(JDH.Mapa, LoopX, LoopY).Blocked = 1 Then

26                        MapData(JDH.Mapa, LoopX, LoopY).Blocked = 0
27                        MapData(JDH.Mapa, LoopX, LoopY).Cofre = 0

28                        Call Bloquear(True, JDH.Mapa, LoopX, LoopY, False)
29                        Call EraseObj(.Amount, JDH.Mapa, LoopX, LoopY)

30                    End If

31                End With
32            Next LoopY
33        Next LoopX
    End With
    Exit Sub
    
errhandleR:
    LogError "error en LimpiarJDJ en " & Erl & ". Err " & Err.Number & " " & Err.description
End Sub

Sub CancelarJDH(Optional ByVal AutoCancel As Boolean = False)

' @ Cancela el jdh.

    Dim LoopX As Long
    Dim uIndex As Integer
    Dim UPos As WorldPos
    Dim Slot As Long
    Dim MiObj As Obj

    'Aviso.
    'If AutoCancel Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(TAG_EVENT, "El evento ha sido cancelado."))
    'End If

    'Llevo los usuarios que entraron a ulla.
    For LoopX = 1 To UBound(JDH.UsUaRiOs())
        uIndex = JDH.UsUaRiOs(LoopX).UserIndex
        'Hay usuario?
        If uIndex <> -1 Then

            'Est� logeado?
            If UserList(uIndex).ConnID <> -1 Then
                'Est� en jdh?
                If UserList(uIndex).EnJDH Then

                    'Telep to anterior posici�n.
                    Call AnteriorPos(uIndex, UPos)

                    'Desparalizamos los pts
                    If Not JDH.Ingresaron >= JDH.Cupos Then
                        Call WritePauseToggle(JDH.UsUaRiOs(LoopX).UserIndex)
                    End If

                    'Reset el flag.
                    UserList(uIndex).EnJDH = False
                    
                    'Lo mandamos a ulla.
                    Call FindLegalPos(uIndex, UPos.Map, UPos.X, UPos.Y)
                    Call WarpUserChar(uIndex, UPos.Map, UPos.X, UPos.Y, True)

                    'Devolvemos el oro de ingreso
                    If JDH.Inscripcion <> 0 Then

                        UserList(uIndex).Stats.GLD = UserList(uIndex).Stats.GLD + JDH.Inscripcion
                        Call WriteUpdateGold(uIndex)

                        Call WriteConsoleMsgNew(uIndex, TAG_EVENT, "El evento ha sido " & IIf(AutoCancel = True, "auto", "") & "cancelado, se te ha devuelto el costo de la inscripci�n.")

                    End If

                End If

            End If

        End If

    Next LoopX

    'Limpia el tipo
    LimpiarJDH

End Sub

Sub ActivarNuevoJDH(ByVal UserIndex As Integer, ByVal Cupos As Byte, ByVal Premio As Long, ByVal Inscripcion As Long)

On Error GoTo errhandleR

' @ Crea nuevo jdh.

    Dim LoopX As Long

    'Limpia el tipo.
1    LimpiarJDH

    'Llena los datos nuevos.
2    With JDH

3        If Cupos > 60 Then Cupos = 60
4        If Cupos < 2 Then Cupos = 2
7        If Premio < 0 Then Premio = 0
6        If Inscripcion < 0 Then Inscripcion = 0

8        .Cupos = Cupos
9        .Premio = Premio
10        .Inscripcion = Inscripcion
11        .Activo = True
12        .EventStarted = False

        'Redim array.
13        ReDim .UsUaRiOs(1 To Cupos) As JDHUser

        'Lleno el array con -1s
14        For LoopX = 1 To Cupos
15            .UsUaRiOs(LoopX).UserIndex = -1
16        Next LoopX

            'Avisa al mundo.
17          Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(TAG_EVENT, "Cantidad de cupos: " & .Cupos & ". Inscripci�n" & IIf(.Inscripcion > 0, " de: " & .Inscripcion & " Monedas de oro, ", " Gratis, ") & IIf(.Premio > 0, "Premio de: " & .Premio & " Monedas de oro.", " No hay premio.")))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(TAG_EVENT, "Para participar escribe /JDH, recuerda que debes tener el inventario vacio! Mucha suerte!"))

18        Call ReCargar_Cofres

    End With
    
    Exit Sub
    
errhandleR:
    LogError "Error en ActivarNuevoJDH en " & Erl & ". Err " & Err.Number

End Sub

Sub ActivarNuevoJDH2(ByVal Cupos As Byte)

' @ Crea nuevo jdh.

    Dim LoopX As Long

    'Limpia el tipo.
    LimpiarJDH

    'Llena los datos nuevos.
    With JDH

        If Cupos > 60 Then Cupos = 60
        If Cupos < 2 Then Cupos = 2

        .Cupos = Cupos
        .Premio = .Cupos * 50000
        .Inscripcion = 100000
        .Activo = True
        .EventStarted = False

        'Redim array.
        ReDim .UsUaRiOs(1 To Cupos) As JDHUser

        'Lleno el array con -1s
        For LoopX = 1 To Cupos
            .UsUaRiOs(LoopX).UserIndex = -1
        Next LoopX

        'Avisa al mundo.
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(TAG_EVENT, .Cupos & " Cupos, Inscripci�n" & IIf(.Inscripcion > 0, " de: " & .Inscripcion & " Monedas de oro, ", " Gratis, ") & IIf(.Premio > 0, "Premio de: " & .Premio & " Monedas de oro.", " No hay premio.") & vbNewLine & "Manden /JDH si desean participar."))

        Call ReCargar_Cofres

    End With

End Sub

Sub IngresarJDH(ByVal UserIndex As Integer)

' @ Usuario ingresa al death.

    Dim LibreSlot As Byte
    Dim SumarCount As Boolean
    Dim NuevaPos As WorldPos
    Dim FuturePos As WorldPos

    LibreSlot = ProximoSlot(SumarCount)

    ' @@ No hay slot.
    If Not LibreSlot <> 0 Then Exit Sub

    With JDH
        ' @@ Hay que sumar?
        If SumarCount Then .Ingresaron = .Ingresaron + 1

        ' @@ Lleno el usuario.
        .UsUaRiOs(LibreSlot).LastPosition = UserList(UserIndex).pos
        .UsUaRiOs(LibreSlot).UserIndex = UserIndex

        ' @@ Asignamos el flag
        UserList(UserIndex).EnJDH = True

        ' @@ Llevo a la arena.
        FuturePos.Map = JDH.Mapa
        FuturePos.X = ARENA_X
        FuturePos.Y = ARENA_Y

        Call ClosestStablePos(FuturePos, NuevaPos)
        Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

        Call WriteConsoleMsgNew(UserIndex, TAG_EVENT, "Has ingresado al evento" & IIf(.Inscripcion > 0, ", se te han descontado " & .Inscripcion & " Monedas de oro.", vbNullString))

        ' @@ Lo paralizamos
        Call WritePauseToggle(UserIndex)

        ' @@ Le descontamos el oro y lo updateamos
        If JDH.Inscripcion <> 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - .Inscripcion
            Call WriteUpdateGold(UserIndex)
        End If

        ' @@ Le quitamos la energia
        If UserList(UserIndex).Stats.MinSta <> 0 Then
            UserList(UserIndex).Stats.MinSta = 0
            Call WriteUpdateSta(UserIndex)
        End If
        
        ' @@ Lleno el cupo?
        If .Ingresaron >= .Cupos Then

            ' @@ Aviso que llen� el cupo
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(TAG_EVENT, "El cupo ha sido completado!"))

            ' @@ Doy inicio
            Iniciar

        End If

    End With

End Sub

Sub NoCancelAndInitialize()

    With JDH

        .Cupos = .Ingresaron

        ' @@ Aviso que llen� el cupo
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(TAG_EVENT, "El cupo ha sido completado!"))

        ' @@ Doy inicio
        Iniciar

    End With

End Sub

Sub CuentaJDH()

' @ Cuenta regresiva

    Dim LoopC As Long

    With JDH

        If .Cuenta > 0 Then
            'Resta el tiempo.
            .Cuenta = .Cuenta - 1

            If .Cuenta > 0 Then
                Call SendData(SendTarget.toMap, JDH.Mapa, PrepareMessageConsoleMsgNew(TAG_EVENT, .Cuenta, , FontTypeNames.FONTTYPE_GUILD))
            Else
                Call SendData(SendTarget.toMap, JDH.Mapa, PrepareMessageConsoleMsgNew(TAG_EVENT, "YA!", , FontTypeNames.FONTTYPE_FIGHT))

                For LoopC = 1 To .Ingresaron
                    Call WritePauseToggle(.UsUaRiOs(LoopC).UserIndex)
                Next LoopC

                .EventStarted = True
            End If
        End If

    End With

End Sub

Sub Iniciar()

' @ Inicia el evento.

    Dim LoopX As Long

    With JDH

        'Set la cuenta.
        .Cuenta = CUENTA_NUM

        'Aviso a los usuarios.
        For LoopX = 1 To UBound(.UsUaRiOs())
            'Hay usuario?
            If .UsUaRiOs(LoopX).UserIndex <> -1 Then
                'Est� logeado?
                If UserList(.UsUaRiOs(LoopX).UserIndex).ConnID <> -1 Then
                    Call WriteConsoleMsgNew(.UsUaRiOs(LoopX).UserIndex, TAG_EVENT, .Cuenta, , FontTypeNames.FONTTYPE_GUILD)
                Else    'No loged, limpio el tipo
                    .UsUaRiOs(LoopX).UserIndex = -1
                End If
            End If
        Next LoopX

    End With

End Sub

Sub MuereUserJDH(ByVal MuertoIndex As Integer, Optional ByVal sMessage As Boolean = True)

' @ Muere usuario en dm.

    Dim MuertoPos As WorldPos
    Dim QuedanEnJDH As Byte

    'Obtengo la anterior posici�n del usuario
    MuertoPos = Ullathorpe

    'Revivir usuario
    If UserList(MuertoIndex).flags.Muerto <> 0 Then
        Call RevivirUsuario(MuertoIndex)

        UserList(MuertoIndex).Stats.MinHp = UserList(MuertoIndex).Stats.MaxHP
        Call WriteUpdateHP(MuertoIndex)
    End If

    Call TirarTodosLosItems(MuertoIndex)

    'Reset el flag.
    UserList(MuertoIndex).EnJDH = False

    'Telep anterior pos.
    Call FindLegalPos(MuertoIndex, MuertoPos.Map, MuertoPos.X, MuertoPos.Y)
    Call WarpUserChar(MuertoIndex, MuertoPos.Map, MuertoPos.X, MuertoPos.Y, True)

    'Aviso al usuario
    If sMessage Then
        Call WriteConsoleMsgNew(MuertoIndex, TAG_EVENT, "Has caido en los Juegos del hambre, has sido revivido y llevado a tu posici�n anterior.")

        'Aviso al mapa.
        Call SendData(SendTarget.toMap, JDH.Mapa, PrepareMessageConsoleMsgNew(TAG_EVENT, UserList(MuertoIndex).Name & " ha sido derrotado."))
    End If

    ' @@ Si no comenzo no hacemos la wea de abajo.
    If Not JDH.EventStarted Then Exit Sub    ' @@ Arriba de esa linea es mejor porqe evitamos procesamiento al pedo.

    'Obtengo los usuarios que quedan..
    QuedanEnJDH = QuedanVivos()

    'Queda 1?
    If Not QuedanEnJDH <> 1 Then
        'Gan� ese usuario!
        Terminar
    End If

End Sub

Sub Terminar()

' @ Termina el death y gana un usuario.

    Dim WinnerIndex As Integer

    WinnerIndex = GanadorIndex

    'No hay ganador!! TRAGEDIAA XDD
    If Not WinnerIndex <> -1 Then
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("TRAGEDIA EN DEATHMATCHS!! WINNERINDEX = -1!!!!", FontTypeNames.FONTTYPE_GUILD))
        LimpiarJDH
        Exit Sub
    End If

    Call TirarTodosLosItems(WinnerIndex)
 
    'Hay ganador, le doi el premio..
    If JDH.Premio <> 0 Then
        UserList(WinnerIndex).Stats.GLD = UserList(WinnerIndex).Stats.GLD + JDH.Premio
        Call WriteUpdateGold(WinnerIndex)
    End If

    Dim Canjes As Integer

    Canjes = 8

    UserList(WinnerIndex).flags.PuntosShop = UserList(WinnerIndex).flags.PuntosShop + Canjes

    Call LogDesarrollo(UserList(WinnerIndex).Name & " gan� los Juegos del Hambre de " & JDH.Cupos & " cupos y gano " & 8 & " puntos de canje.")

    'Sacamos el .flags
    UserList(WinnerIndex).EnJDH = False

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew(TAG_EVENT, "Ganador del evento: " & UserList(WinnerIndex).Name & " se lleva una cantidad de " & JDH.Premio & " monedas de oro y 8 Canjes" & " Felicitaciones!"))

    'Ganador a su anterior posici�n..
    Dim ToPosition As WorldPos
    ToPosition = Ullathorpe

    'Warp.
    Call FindLegalPos(WinnerIndex, ToPosition.Map, ToPosition.X, ToPosition.Y)
    Call WarpUserChar(WinnerIndex, ToPosition.Map, ToPosition.X, ToPosition.Y, True)

    LimpiarJDH

End Sub

Sub DesconectaUserJDH(ByVal UserIndex As Integer)

    Dim i As Long

    For i = 1 To UBound(JDH.UsUaRiOs())
        If JDH.UsUaRiOs(i).UserIndex = UserIndex Then
            JDH.UsUaRiOs(i).UserIndex = -1
            Exit For
        End If
    Next i

    Call MuereUserJDH(UserIndex, False)
    Call SendData(SendTarget.toMap, JDH.Mapa, PrepareMessageConsoleMsgNew(TAG_EVENT, UserList(UserIndex).Name & " abandon� los Juegos del Hambre."))

End Sub

Sub AnteriorPos(ByVal UserIndex As Integer, ByRef MuertoPosition As WorldPos)

' @ Devuelve la posici�n anterior del usuario.

    Dim LoopX As Long

    For LoopX = 1 To UBound(JDH.UsUaRiOs())
        If JDH.UsUaRiOs(LoopX).UserIndex = UserIndex Then
            MuertoPosition = JDH.UsUaRiOs(LoopX).LastPosition
            Exit Sub
        End If
    Next LoopX

    'Posici�n de ulla u.u
    MuertoPosition = Ullathorpe

End Sub

Function AprobarIngresoJDH(ByVal ID As Integer, ByRef MensajeError As String) As Boolean

' @ Checks si puede ingresar al jdh.

    AprobarIngresoJDH = False

    ' @@ No hay death.
    If Not JDH.Activo Then
        MensajeError = "El evento no est� en curso."
        Exit Function
    End If

    ' @@ Inscripcion
    If UserList(ID).Stats.GLD < JDH.Inscripcion Then
        MensajeError = "Te faltan " & JDH.Ingresaron - UserList(ID).Stats.GLD & " monedas de oro"
        Exit Function
    End If

    ' @@ No hay cupos
    If JDH.Ingresaron >= JDH.Cupos Then
        MensajeError = "El evento ya no tiene cupos disponibles."
        Exit Function
    End If

    'Ya inscripto?
    If YaInscripto(ID) Then
        MensajeError = "Ya est�s en los Juegos del Hambre."
        Exit Function
    End If
        
    If UserList(ID).Invent.NroItems > 0 Then
        MensajeError = "Tenes que estar con inventario vacio para jugar."
        Exit Function
    End If
    
    
    AprobarIngresoJDH = True

End Function

Function ProximoSlot(ByRef Sumar As Boolean) As Byte

' @ Posici�n para un usuario.

    Dim LoopX As Long

    Sumar = False

    For LoopX = 1 To UBound(JDH.UsUaRiOs())

        ' @@ No hay usuario.
        If Not (JDH.UsUaRiOs(LoopX).UserIndex <> -1) Then

            ' @@ Slot encontrado.
            ProximoSlot = LoopX

            ' @@ Hay que sumar el contador?
            If JDH.Ingresaron < ProximoSlot Then
                Sumar = True
            End If

            Exit Function

        End If

    Next LoopX

    ProximoSlot = 0

End Function

Function QuedanVivos() As Byte

' @ Devuelve la cantidad de usuarios vivos que quedan.

    Dim LoopX As Long
    Dim Counter As Byte
    For LoopX = 1 To UBound(JDH.UsUaRiOs())

        With JDH.UsUaRiOs(LoopX)

            ' @@ Mientras halla usuario.
            If .UserIndex > 0 Then

                ' @@ Mientras est� logeado
                If (UserList(.UserIndex).ConnID <> -1) Then

                    ' @@ Mientras est� en jdh
                    If UserList(.UserIndex).EnJDH Then

                        ' @@ Sumo contador.
                        Counter = Counter + 1

                    End If

                End If

            End If

        End With

    Next LoopX

    QuedanVivos = Counter

End Function

Function GanadorIndex() As Integer

' @ Busca el ganador..

    Dim LoopX As Long
    For LoopX = 1 To UBound(JDH.UsUaRiOs())

        With JDH.UsUaRiOs(LoopX)

            If .UserIndex > 0 Then

                If UserList(.UserIndex).ConnID <> -1 Then

                    If UserList(.UserIndex).EnJDH Then
                        If Not (UserList(.UserIndex).flags.Muerto <> 0) Then
                            GanadorIndex = .UserIndex
                            Exit Function
                        End If
                    End If

                End If


            End If

        End With

    Next LoopX

    'No hay ganador! WTF!!!
    GanadorIndex = -1

End Function

Function YaInscripto(ByVal UserIndex As Integer) As Boolean

' @ Devuelve si ya est� inscripto.

    Dim LoopX As Long

    For LoopX = 1 To UBound(JDH.UsUaRiOs())
        If (JDH.UsUaRiOs(LoopX).UserIndex = UserIndex) Then
            YaInscripto = True
            Exit Function
        End If
    Next LoopX

    YaInscripto = False

End Function

Public Sub Clickea_Cofre(ByRef pos As WorldPos)

    Dim ID As Byte
    Dim DataCofre As Obj
    Dim LoopC As Long
    Dim n_Pos As WorldPos

    DataCofre.Amount = 1
    DataCofre.objIndex = Cofre_Abierto

    ID = MapData(pos.Map, pos.X, pos.Y).Cofre

    With JDH
        If .Activo = False Then Exit Sub
        If ID = 0 Then Exit Sub
        If JDH.Ingresaron < JDH.Cupos Then Exit Sub
        If .Cofres(ID).Abierto = True Then Exit Sub
        If .Cuenta > 0 Then Exit Sub

        .Cofres(ID).Abierto = True

        Call EraseObj(MapData(pos.Map, pos.X, pos.Y).ObjInfo.Amount, pos.Map, pos.X, pos.Y)
        Call MakeObj(DataCofre, .Mapa, pos.X, pos.Y)

        For LoopC = 1 To UBound(.Cofres(ID).Objetos())
            Call Tilelibre(pos, n_Pos, .Cofres(ID).Objetos(LoopC), False, True)
            Call MakeObj(.Cofres(ID).Objetos(LoopC), .Mapa, n_Pos.X, n_Pos.Y)
        Next LoopC

    End With

End Sub

Private Sub ReCargar_Cofres()

    Dim DataCofre As Obj
    Dim LoopC As Long
    Dim posX As Byte
    Dim posY As Byte

    DataCofre.Amount = 1
    DataCofre.objIndex = Cofre_Cerrado

    With JDH

        For LoopC = 1 To UBound(.Cofres())

            .Cofres(LoopC).Abierto = False

            posX = .Cofres(LoopC).X + RandomNumber(1, 10)
            posY = .Cofres(LoopC).Y + RandomNumber(1, 10)

            If posX > 100 Then posX = 90
            If posY > 100 Then posY = 90

            MapData(.Mapa, posX, posY).Blocked = 1
            MapData(.Mapa, posX, posY).Cofre = LoopC

            Call MakeObj(DataCofre, .Mapa, posX, posY)
            Call Bloquear(True, .Mapa, posX, posY, True)

        Next LoopC

    End With

End Sub

