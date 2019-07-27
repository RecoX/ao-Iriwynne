Attribute VB_Name = "Mod_Torneo1vs1"
Option Explicit

Private PuntosGanados As Integer

Private Type cClasesv
     Mago As Byte
     Clerigo As Byte
     Bardo As Byte
     Paladin As Byte
     Asesino As Byte
     Cazador As Byte
     Guerrero As Byte
     Druida As Byte
     Ladron As Byte
     Bandido As Byte
End Type

Private Type tMapPos
    MapaTorneo As Integer
    EsperaX As Byte
    EsperaY As Byte

    Esquina1x As Byte
    Esquina2x As Byte
    Esquina1y As Byte
    Esquina2y As Byte
    X1 As Byte
    Y1 As Byte
    X2 As Byte
    Y2 As Byte
End Type

Private Type tTorneo1vs1
    AutoCancelTime As Byte
    MinLevel As Byte
    MaxLevel As Byte
    ClasesValidas As cClasesv
    Cupos As Byte
    rondas As Integer
    Activo As Boolean
    Ingresaron As Boolean
    Premio As Long
    Inscripcion As Long
    Peleando(1 To 2) As String
    Cuenta As Byte
    Torneo_Luchadores() As Integer
End Type

Private Const CONST_CUENTA As Byte = 5

Private Const TAG_EVENT As String = "Torneo 1vs1>"

Private EventoMapPos As tMapPos

Public Torneo1vs1 As tTorneo1vs1

Public Sub Torneo1vs1_CargarPos()

    On Error GoTo errhandler

    With EventoMapPos

        ' @@ MAPA DONDE SE PELEA
        .MapaTorneo = 72
        
        .Esquina2x = 35
        .Esquina2y = 29
        .Esquina1x = 19
        .Esquina1y = 17

        .EsperaX = 44
        .EsperaY = 69

        .X1 = 35 'CInt(Leer.GetValue("INIT", "X1"))
        .X2 = 56 'CInt(Leer.GetValue("INIT", "X2"))
        .Y1 = 61 'CInt(Leer.GetValue("INIT", "Y1"))
        .Y2 = 82 'CInt(Leer.GetValue("INIT", "Y2"))

    End With
Exit Sub

errhandler:
LogError "ERROR AL CARGAR POS DE TORNEOI 1vs1"
Debug.Print "ERRORASO"

End Sub

Sub Rondas_Cancela(Optional ByVal AutoCancel As Boolean = False)

10  On Error GoTo Rondas_Cancela_Error

20  With Torneo1vs1
        .AutoCancelTime = 0
        
        If .Activo = False Then Exit Sub
        
'30      If (Not .Activo And Not .Ingresaron) Then
'            If AutoCancel = False Then Exit Sub
'        End If
        
40      .Activo = False
50      .Ingresaron = False
        
33        If Len(Torneo1vs1.Peleando(1)) > 0 Then
34            If (NameIndex(Torneo1vs1.Peleando(1))) > 0 Then
35                If Torneo1vs1.Cuenta > 0 Then
36                    WritePauseToggle (NameIndex(Torneo1vs1.Peleando(1)))
37                End If
83            End If
88        End If
        
89        If Len(Torneo1vs1.Peleando(2)) > 0 Then
77            If (NameIndex(Torneo1vs1.Peleando(2))) > 0 Then
76                If Torneo1vs1.Cuenta > 0 Then
75                    WritePauseToggle (NameIndex(Torneo1vs1.Peleando(2)))
74                End If
73            End If
        End If
        
70        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "El evento ha sido cancelado."))

        Dim i As Long
        
939        If UBound(.Torneo_Luchadores) <= 0 Then Exit Sub
        
90      For i = LBound(.Torneo_Luchadores) To UBound(.Torneo_Luchadores)

100         If (.Torneo_Luchadores(i) <> -1) Then

                Dim NuevaPos As WorldPos, FuturePos As WorldPos

110             FuturePos.Map = Ullathorpe.Map
120             FuturePos.X = Ullathorpe.X
130             FuturePos.Y = Ullathorpe.Y

140             Call ClosestLegalPos(FuturePos, NuevaPos)
                
                UserList(.Torneo_Luchadores(i)).Stats.GLD = UserList(.Torneo_Luchadores(i)).Stats.GLD + .Inscripcion
                WriteUpdateGold .Torneo_Luchadores(i)
                
                Dim N As WorldPos
                Call ClosestLegalPos(NuevaPos, N)
280             Call WarpUserChar(.Torneo_Luchadores(i), N.Map, N.X, N.Y, True)

290             UserList(.Torneo_Luchadores(i)).flags.Automatico = False

350         End If

360     Next i

370 End With
        
        
380 Exit Sub

Rondas_Cancela_Error:

390 Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Rondas_Cancela of Módulo m_Torneo1vs1" & Erl & ".")
    Resume Next

End Sub

Sub Rondas_UsuarioMuere(ByVal Userindex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)

10  On Error GoTo Rondas_UsuarioMuere_Error

20  With Torneo1vs1

        Dim i As Integer
        Dim pos As Integer
        Dim combate As Integer
        Dim LI1 As Integer, LI2 As Integer
        Dim UI1 As Integer, UI2 As Integer
        Dim N As WorldPos, k As WorldPos
30      If (Not Torneo1vs1.Activo) Then
40          Exit Sub
50      ElseIf (.Activo And .Ingresaron) Then
60          For i = LBound(.Torneo_Luchadores) To UBound(.Torneo_Luchadores)
70              If (.Torneo_Luchadores(i) = Userindex) Then
80                  .Torneo_Luchadores(i) = -1
                    Call ClosestLegalPos(Ullathorpe, N)
90                  Call WarpUserChar(Userindex, N.Map, N.X, N.Y, True)
100                 UserList(Userindex).flags.Automatico = False
110                 Exit Sub
120             End If
130         Next i
140         Exit Sub
150     End If

160     For pos = LBound(.Torneo_Luchadores) To UBound(.Torneo_Luchadores)
170         If (.Torneo_Luchadores(pos) = Userindex) Then Exit For
180     Next pos

        ' si no lo ha encontrado
190     If (.Torneo_Luchadores(pos) <> Userindex) Then Exit Sub

        '  Ojo con esta parte, aqui es donde verifica si el usuario esta en la posicion de espera del torneo, en estas cordenadas tienen que fijarse al crear su Mapa de torneos.

200     If UserList(Userindex).pos.X >= EventoMapPos.X1 And UserList(Userindex).pos.X <= EventoMapPos.X2 And UserList(Userindex).pos.Y >= EventoMapPos.Y1 And UserList(Userindex).pos.Y <= EventoMapPos.Y2 Then
210         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", UserList(Userindex).Name & " se fue mientras esperaba pelear!"))

            Call ClosestLegalPos(Ullathorpe, N)
            Call WarpUserChar(Userindex, N.Map, N.X, N.Y, True)
230         UserList(Userindex).flags.Automatico = False
240         .Torneo_Luchadores(pos) = -1
250         Exit Sub
260     End If

270     combate = 1 + (pos - 1) \ 2

        'ponemos li1 y li2 (luchador index) de los que combatian
280     LI1 = 2 * (combate - 1) + 1
290     LI2 = LI1 + 1

        'se informa a la gente
300     If (Real) Then
310         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", UserList(Userindex).Name & " pierde el combate!"))
320     Else
330         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", UserList(Userindex).Name & " se fue del combate!"))
340     End If

        'se le teleporta fuera si murio
460     If (Real) Then

            Dim NuevaPos As WorldPos, FuturePos As WorldPos

470         FuturePos.Map = Ullathorpe.Map
480         FuturePos.X = Ullathorpe.X
490         FuturePos.Y = Ullathorpe.Y

540         Call ClosestLegalPos(FuturePos, NuevaPos)
550         Call WarpUserChar(Userindex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

560         UserList(Userindex).flags.Automatico = False

570     ElseIf (Not CambioMapa) Then

            Call ClosestLegalPos(Ullathorpe, N)
620         Call WarpUserChar(Userindex, N.Map, N.X, N.Y, True)
630         UserList(Userindex).flags.Automatico = False

640     End If

        'se le borra de la lista y se mueve el segundo a li1
650     If (.Torneo_Luchadores(LI1) = Userindex) Then
660         .Torneo_Luchadores(LI1) = .Torneo_Luchadores(LI2)      'cambiamos slot
670         .Torneo_Luchadores(LI2) = -1
680     Else
690         .Torneo_Luchadores(LI2) = -1
700     End If

        'si es la ultima ronda
533     If (.rondas = 1) Then

            k = Ullathorpe
            Call ClosestLegalPos(k, N)
890         Call WarpUserChar(.Torneo_Luchadores(LI1), N.Map, N.X, N.Y, True)
900         'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "El ganador del torneo es " & UserList(.Torneo_Luchadores(LI1)).Name) & " por haber matado " & .Cupos & " usuarios.")
            
            With UserList(.Torneo_Luchadores(LI1))
                'Aviso al mundo.
                SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", .Name & " - " & ListaClases(.clase) & " " & ListaRazas(.raza) & " Nivel " & .Stats.ELV & " Ganó " & Format$(Torneo1vs1.Premio, "#,###") & " monedas de oro y " & Torneo1vs1.Cupos & " puntos de Canje por salir primero en el evento.")
            
            End With
            
            Dim Canjes As Integer

            Canjes = .Cupos
            
            UserList(.Torneo_Luchadores(LI1)).flags.PuntosShop = UserList(.Torneo_Luchadores(LI1)).flags.PuntosShop + Canjes

            'Oro de Premio
1080        If .Premio <> 0 Then

1090            UserList(.Torneo_Luchadores(LI1)).Stats.GLD = UserList(.Torneo_Luchadores(LI1)).Stats.GLD + .Premio
1100            Call WriteUpdateGold(.Torneo_Luchadores(LI1))

1110        End If

            'Sumamos un torneo ganado
'1120        UserList(.Torneo_Luchadores(LI1)).Stats.TorneosGanados = UserList(.Torneo_Luchadores(LI1)).Stats.TorneosGanados + 1

1130        UserList(.Torneo_Luchadores(LI1)).flags.Automatico = False

1140        .Activo = False

1150        Erase .Torneo_Luchadores()
            .Cuenta = 0
1170        .AutoCancelTime = 0

1190        Exit Sub

1200    Else

            'A su compañero se le teleporta dentro, condicional por seguridad
1340        Call WarpUserChar(.Torneo_Luchadores(LI1), EventoMapPos.MapaTorneo, EventoMapPos.EsperaX, EventoMapPos.EsperaY, True)

1350    End If

        'Si es el ultimo combate de la ronda
1360    If (2 ^ .rondas = 2 * combate) Then

1370        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Siguiente ronda"))
1380        .rondas = .rondas - 1

            'Antes de llamar a la proxima ronda hay q copiar a los tipos
1390        For i = 1 To 2 ^ .rondas
1400            UI1 = .Torneo_Luchadores(2 * (i - 1) + 1)
1410            UI2 = .Torneo_Luchadores(2 * i)
1420            If (UI1 = -1) Then UI1 = UI2
1430            .Torneo_Luchadores(i) = UI1
1440        Next i

1450        ReDim Preserve .Torneo_Luchadores(1 To 2 ^ .rondas) As Integer

1460        Call Rondas_Combate(1)

1470        Exit Sub

1480    End If

        'vamos al siguiente combate
1490    Call Rondas_Combate(combate + 1)

1500 End With

1510 Exit Sub

Rondas_UsuarioMuere_Error:

1520 Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Rondas_UsuarioMuere, line " & Erl & ".")

End Sub

Sub Rondas_UsuarioDesconecta(ByVal Userindex As Integer)

10  On Error GoTo Rondas_UsuarioDesconecta_Error

    ' @@ Por alguna razon explotaba todo al pingo y quedaba como userindex 0 xddx lo parcheamos para que no explote too.
20  If Userindex > 0 Then
        
        If (NameIndex(Torneo1vs1.Peleando(1))) > 0 Then
        If Torneo1vs1.Cuenta > 0 Then
            WritePauseToggle (NameIndex(Torneo1vs1.Peleando(1)))
        End If
        End If
        
        If (NameIndex(Torneo1vs1.Peleando(2))) > 0 Then
        If Torneo1vs1.Cuenta > 0 Then
            WritePauseToggle (NameIndex(Torneo1vs1.Peleando(2)))
        End If
        End If
        
        Torneo1vs1.Cuenta = 0
        
'30      Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", UserList(UserIndex).Name & " ha desconectado en Torneo1vs1."))

120     Call Rondas_UsuarioMuere(Userindex, False, False)

130 End If

140 Exit Sub

Rondas_UsuarioDesconecta_Error:

150 Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Rondas_UsuarioDesconecta, line " & Erl & ".")

End Sub

Sub Rondas_UsuarioCambiamapa(ByVal Userindex As Integer)

    Call Rondas_UsuarioMuere(Userindex, False, True)

End Sub

Sub Torneos_Inicia(ByVal Userindex As Integer, ByVal rondas As Integer, ByVal MinLevel As Byte, ByVal MaxLevel As Byte, ByVal cMago As Byte, ByVal cClerigo As Byte, ByVal cBardo As Byte, ByVal cPaladin As Byte, ByVal cAsesino As Byte, ByVal cCazador As Byte, ByVal cGuerrero As Byte, ByVal cDruida As Byte, ByVal cLadron As Byte, ByVal cBandido As Byte)
    With Torneo1vs1

        Dim i As Long

        ' @@ Ya esta en curso
        If (.Activo) Then WriteConsoleMsg Userindex, "Ya hay un torneo activo.", FontTypeNames.FONTTYPE_AMARILLO: Exit Sub   'Ya hay un torneo en curso
        
        .Cupos = val(2 ^ rondas)
        .rondas = rondas
        .AutoCancelTime = 0
        .Activo = True
        .Ingresaron = True
        
        .Inscripcion = 20000
        .Premio = 100000 ' * .Cupos

        .MaxLevel = MaxLevel
        .MinLevel = MinLevel
     
        Dim T As Boolean
        .AutoCancelTime = 180
        .ClasesValidas.Mago = cMago: If T = False Then T = (cMago > 0)
        .ClasesValidas.Clerigo = cClerigo: If T = False Then T = (cClerigo > 0)
        .ClasesValidas.Bardo = cBardo: If T = False Then T = (cBardo > 0)
        .ClasesValidas.Paladin = cPaladin: If T = False Then T = (cPaladin > 0)
        .ClasesValidas.Asesino = cAsesino: If T = False Then T = (cAsesino > 0)
        .ClasesValidas.Cazador = cCazador: If T = False Then T = (cCazador > 0)
        .ClasesValidas.Guerrero = cGuerrero: If T = False Then T = (cGuerrero > 0)
        .ClasesValidas.Druida = cDruida: If T = False Then T = (cDruida > 0)
        .ClasesValidas.Ladron = cLadron: If T = False Then T = (cLadron > 0)
        .ClasesValidas.Bandido = cBandido: If T = False Then T = (cBandido > 0)
           
        ReDim .Torneo_Luchadores(1 To 2 ^ rondas) As Integer
        
        For i = LBound(.Torneo_Luchadores) To UBound(.Torneo_Luchadores)
            .Torneo_Luchadores(i) = -1
        Next i

        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo Automático>", "Esta empezando un nuevo Torneo 1vs1 de " & .Cupos & " participantes!! para participar teclea /PARTICIPAR - (No cae inventario)"))
        
        PuntosGanados = .Cupos
           
        If T Then
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo Automático>", "Clases prohibidas: " & IIf(cMago > 0, " MAGO ", "") & IIf(cClerigo > 0, " CLERIGO ", "") & IIf(cBardo > 0, " BARDO ", "") & IIf(cPaladin > 0, " PALADIN ", "") & IIf(cAsesino > 0, " ASESINO ", "") & IIf(cCazador > 0, " CAZADOR ", "") & IIf(cGuerrero > 0, " GUERRERO ", "") & IIf(cDruida > 0, " DRUIDA ", "") & IIf(cLadron > 0, " LADRON ", "") & IIf(cBandido > 0, " BANDIDO ", ""), FontTypeNames.FONTTYPE_TORNEO)
        End If

    End With

End Sub


Sub Ingresar1vs1(ByVal Userindex As Integer)

    Dim i As Integer

10  On Error GoTo Torneos_Entra_Error

20  With UserList(Userindex)

        '@@ Tiene el oro de ingreso?
70      If .Stats.GLD < 20000 Then Call WriteConsoleMsgNew(Userindex, "Torneo 1vs1>", "No tienes el oro suficiente. Necesitas tener " & Torneo1vs1.Inscripcion & " monedas de oro."): Exit Sub

        '@@ Esta comerciando?
90      If .flags.Comerciando <> 0 Then Exit Sub

        If .pos.Map <> 1 Then Exit Sub

        '@@ No hay ningún torneo en curso
100     If (Not Torneo1vs1.Activo) Then Call WriteConsoleMsg(Userindex, "No hay ningún torneo en curso.", FontTypeNames.FONTTYPE_INFO): Exit Sub

        '@@ Cupos llenos.
110     If (Not Torneo1vs1.Ingresaron) Then Call WriteConsoleMsg(Userindex, "Cupos llenos.", FontTypeNames.FONTTYPE_INFO): Exit Sub
        
        If Torneo1vs1.MinLevel > UserList(Userindex).Stats.ELV Then
            Call WriteConsoleMsg(Userindex, "El mínimo nivel para entrar es de " & Torneo1vs1.MinLevel, FontTypeNames.FONTTYPE_TORNEO)
            Exit Sub
        End If
        
        If Torneo1vs1.MaxLevel < UserList(Userindex).Stats.ELV Then
            Call WriteConsoleMsg(Userindex, "El máximo nivel para entrar es de " & Torneo1vs1.MaxLevel, FontTypeNames.FONTTYPE_TORNEO)
            Exit Sub
        End If
        
        Dim MiClase As String
        MiClase = ListaClases(UserList(Userindex).clase)

On Error GoTo NoHayClases
    If MiClase = ListaClases(DeathMatch.ClasesValidas.Asesino) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Bandido) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Bardo) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Cazador) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Clerigo) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Druida) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Guerrero) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Ladron) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Mago) Or MiClase = ListaClases(DeathMatch.ClasesValidas.Paladin) Then

        Call WriteConsoleMsg(Userindex, "Tu clase no es válida.", FontTypeNames.FONTTYPE_TORNEO)
        Exit Sub

    End If

NoHayClases:

130 End With

    'Esta en torneo?
140 For i = LBound(Torneo1vs1.Torneo_Luchadores) To UBound(Torneo1vs1.Torneo_Luchadores)
150     If (Torneo1vs1.Torneo_Luchadores(i) = Userindex) Then
160         Call WriteConsoleMsg(Userindex, "Ya estás dentro del torneo", FontTypeNames.FONTTYPE_INFO)
170         Exit Sub
180     End If
190 Next i

200 For i = LBound(Torneo1vs1.Torneo_Luchadores) To UBound(Torneo1vs1.Torneo_Luchadores)

210     If (Torneo1vs1.Torneo_Luchadores(i) = -1) Then
220         Torneo1vs1.Torneo_Luchadores(i) = Userindex

            Dim NuevaPos As WorldPos, FuturePos As WorldPos

230         FuturePos.Map = EventoMapPos.MapaTorneo
240         FuturePos.X = EventoMapPos.EsperaX
250         FuturePos.Y = EventoMapPos.EsperaY

350         Call ClosestLegalPos(FuturePos, NuevaPos)

360         Call WarpUserChar(Torneo1vs1.Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, False)

370         UserList(Torneo1vs1.Torneo_Luchadores(i)).flags.Automatico = True

'380         Call WriteMensajes(UserIndex, eMensajes.mensaje484)          'Has ingresado al torneo

390         If Torneo1vs1.Inscripcion <> 0 Then
400             UserList(Torneo1vs1.Torneo_Luchadores(i)).Stats.GLD = UserList(Torneo1vs1.Torneo_Luchadores(i)).Stats.GLD - 20000
410             Call WriteUpdateGold(Torneo1vs1.Torneo_Luchadores(i))
420         End If

430         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "El personaje " & UserList(Userindex).Name & " ingresó al Torneo1vs1."))

440         If (i = UBound(Torneo1vs1.Torneo_Luchadores)) Then

450             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Comienza el torneo"))

460             Torneo1vs1.Ingresaron = False
                Torneo1vs1.AutoCancelTime = 0
480             Call Rondas_Combate(1)

490         End If

500         Exit Sub

510     End If

520 Next i

530 Exit Sub

Torneos_Entra_Error:

540 Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Torneos_Entra of Módulo m_Torneo1vs1" & Erl & ".")

End Sub

Sub Rondas_Combate(combate As Integer)

10  On Error GoTo Rondas_Combate_Error

20  With Torneo1vs1

        Dim UI1 As Integer
        Dim UI2 As Integer
        Dim k As WorldPos, N As WorldPos

30      UI1 = .Torneo_Luchadores(2 * (combate - 1) + 1)
40      UI2 = .Torneo_Luchadores(2 * combate)

50      If (UI2 = -1) Then
60          UI2 = .Torneo_Luchadores(2 * (combate - 1) + 1)
70          UI1 = .Torneo_Luchadores(2 * combate)
80      End If

90      If (UI1 = -1) Then

100         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Combate anulado por la desconexión de uno de los dos participantes."))

110         If (.rondas = 1) Then

120             If (UI2 <> -1) Then

130                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Torneo terminado, ganador del torneo por eliminación " & UserList(UI2).Name & "."))

                    Call ClosestLegalPos(Ullathorpe, N)
140                 Call WarpUserChar(UI2, N.Map, N.X, N.Y, True)

                    UserList(UI2).flags.PuntosShop = UserList(UI2).flags.PuntosShop + .Cupos
                    'Oro de Premio
310                 If .Premio <> 0 Then
320                     UserList(UI2).Stats.GLD = UserList(UI2).Stats.GLD + .Premio
330                     Call WriteUpdateGold(UI2)
340                 End If

360                 UserList(UI2).flags.Automatico = False

370                 .Activo = False

380                 Erase .Torneo_Luchadores()

400                 .AutoCancelTime = 0

420                 Exit Sub

430             End If

440             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "No hay ganador del evento por la desconexión de todos sus participantes."))

450             Exit Sub

460         End If

470         If (UI2 <> -1) Then Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "El usuario " & UserList(UI2).Name & " pasó a la siguiente ronda."))

480         If (2 ^ .rondas = 2 * combate) Then

490             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Siguiente ronda."))

500             .rondas = .rondas - 1

                'antes de llamar a la proxima ronda hay q copiar a los tipos
                Dim i As Integer

510             For i = 1 To 2 ^ .rondas
520                 UI1 = .Torneo_Luchadores(2 * (i - 1) + 1)
530                 UI2 = .Torneo_Luchadores(2 * i)
540                 If (UI1 = -1) Then UI1 = UI2
550                 .Torneo_Luchadores(i) = UI1
560             Next i

570             ReDim Preserve .Torneo_Luchadores(1 To 2 ^ .rondas) As Integer

580             Call Rondas_Combate(1)

590             Exit Sub

600         End If

610         Call Rondas_Combate(combate + 1)

620         Exit Sub

630     End If

640     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Torneo 1vs1>", UserList(UI1).Name & " vs. " & UserList(UI2).Name & ". Esquinas, Peleen"))

650     .Peleando(1) = UserList(UI1).Name
660     .Peleando(2) = UserList(UI2).Name
        WritePauseToggle UI1
        WritePauseToggle UI2
        .Cuenta = 5
        
        ' @@ Jugador 1
670     With UserList(UI1)

680         .Stats.MinHp = .Stats.MaxHP
690         .Stats.MinMAN = .Stats.MaxMAN
700         .Stats.MinSta = .Stats.MaxSta
134         .Stats.MinAGU = 100
720         .Stats.MinHam = 100

730         Call WriteUpdateUserStats(UI1)

740     End With

        ' @@ Jugador 2
750     With UserList(UI2)

760         .Stats.MinHp = .Stats.MaxHP
770         .Stats.MinMAN = .Stats.MaxMAN
780         .Stats.MinSta = .Stats.MaxSta
790         .Stats.MinAGU = 100
800         .Stats.MinHam = 100

810         Call WriteUpdateUserStats(UI2)

820     End With

840     Call WarpUserChar(UI1, EventoMapPos.MapaTorneo, EventoMapPos.Esquina1x, EventoMapPos.Esquina1y, True)
850     Call WarpUserChar(UI2, EventoMapPos.MapaTorneo, EventoMapPos.Esquina2x, EventoMapPos.Esquina2y, True)

860 End With

870 Exit Sub

Rondas_Combate_Error:

880 Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Rondas_Combate, line " & Erl & ".")

End Sub


Sub Cuenta()
 
 
 On Error Resume Next
 
' @ Cuenta regresiva y auto-cancel acá.
 
Dim PacketToSend    As String
Dim CanSendPackage  As Boolean
 
With Torneo1vs1
      
    'Hay cuenta?
    If .Cuenta <> 0 Then
        'Resta el tiempo.
        .Cuenta = .Cuenta - 1
       
        If .Cuenta > 1 Then
            SendData SendTarget.toMap, EventoMapPos.MapaTorneo, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "La ronda iniciará en " & .Cuenta & " segundos.")
        ElseIf .Cuenta = 1 Then
            SendData SendTarget.toMap, EventoMapPos.MapaTorneo, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "La ronda iniciará en 1 segundo!")
        ElseIf .Cuenta <= 0 Then
            WritePauseToggle (NameIndex(Torneo1vs1.Peleando(1)))
            WritePauseToggle (NameIndex(Torneo1vs1.Peleando(2)))
            SendData SendTarget.toMap, EventoMapPos.MapaTorneo, PrepareMessageConsoleMsgNew("Torneo 1vs1>", "¡PELEEN!")
        End If
    End If
    
    'Tiempo de auto-cancelamiento?
    If .AutoCancelTime <> 0 Then
       'Resto el contador
       If .AutoCancelTime <> 0 Then
           .AutoCancelTime = .AutoCancelTime - 1
       End If
             
       'Avisa cada 30 segundos.
       Select Case .AutoCancelTime
              Case 150      'Quedan 2:30.
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en 2:30 minutos")
              Case 120
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en 2 minutos")
              Case 90
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en 1:30 minutos")
              Case 60
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en 1 minuto")
              Case 30
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en 30 segundos")
              'Avisa a los 15
              Case 15
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en 15 segundos")
              'Avisa a los 10
              Case 10
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en 10 segundos")
              'Avisa a los 5
              Case 5
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en 5 segundos")
              'Avisa a los 3,2,1.
              Case 1, 2, 3
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Torneo 1vs1>", "Auto-Cancelará en " & .AutoCancelTime & " segundo/s")
              Case 0
                   CanSendPackage = False
                   Call Rondas_Cancela(1)
       End Select
       
       'Hay que enviar el mensaje?
       If CanSendPackage Then
          'Envia
          SendData SendTarget.ToAll, 0, PacketToSend
          'Reset el flag.
          CanSendPackage = False
       End If
       
    End If
   
End With
 
End Sub
 


