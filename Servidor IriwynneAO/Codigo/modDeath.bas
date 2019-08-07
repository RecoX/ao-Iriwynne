Attribute VB_Name = "modDeath"
Option Explicit
 
'Programado por maTih.-
 
Type DeathUser
     UserIndex      As Integer      'UI del usuario.
     LastPosition   As WorldPos     'Pos que estaba antes de entrar.
     Esperando      As Byte         'Tiempo de espera para volver.
End Type

Private Type cClasesx
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
 
Type tDeath
     Cupos          As Byte         'Cantidad de cupos.
     Ingresaron     As Byte         'Cantidad que ingreso.
     Usuarios()     As DeathUser    'Tipo de usuarios
     ClasesValidas As cClasesx
     Cuenta         As Byte         'Cuenta regresiva.
     Activo         As Boolean      'Hay deathmatch
     MaxLevel As Byte
     MinLevel As Byte
     AutoCancelTime As Byte         'Tiempo de auto-cancelamiento
     Ganador        As DeathUser    'Datos del ganador.
     EventStarted As Boolean
End Type
 
Private PuntosGanados As Integer
Const CUENTA_NUM    As Byte = 5     'Segundos de cuenta.
Const ARENA_MAP     As Integer = 72 'Mapa de la arena.
Const ARENA_X       As Byte = 74    'X de la arena(se suma por usuario)
Const ARENA_Y       As Byte = 53    'Y de la arena.
Const BANCO_X       As Byte = 74    'X donde aparece el banquero.
Const BANCO_Y       As Byte = 53    'Y Donde aparece el banquro.
 
Const PREMIO_POR_CABEZA As Long = 20000 'Premio en oro , el c�lculo es el de ac� abajo.
Const TIEMPO_AUTOCANCEL As Byte = 180     '3 Minutos antes del auto-cancel.

'C�lculo : PREMIO_POR_CABEZA * JUGADORES QUE PARTICIPARON
 
Public DeathMatch   As tDeath
 
 Sub DesconectaUser(ByVal UserIndex As Integer)
 
'
' @ maTih.-
On Error GoTo ErrHandler


    Dim i As Long

1    For i = 1 To UBound(DeathMatch.Usuarios())
2        If DeathMatch.Usuarios(i).UserIndex = UserIndex Then
3            DeathMatch.Usuarios(i).UserIndex = -1
4            Exit For
5        End If
6    Next i

7    Call MuereUser(UserIndex, False)
8    Call SendData(SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsgNew("Deathmatch>", UserList(UserIndex).Name & " abandon� el deathmatch."))

Exit Sub
ErrHandler:
LogError "Error DEath -Desconect: " & Err.Number & "  " & Err.description & " __ Linea: " & Erl


End Sub
Sub Limpiar()
 
' @ Limpia los datos anteriores.
 
Dim DumpPos     As WorldPos
Dim LoopX       As Long
Dim LoopY       As Long
Dim esSalida    As Boolean
 
With DeathMatch
     .Cuenta = 0
     .Cupos = 0
     .Ingresaron = 0
     .Activo = False
     .EventStarted = False
     
     'Limpio el tipo de ganador.
     With .Ganador
          .UserIndex = 0
          .LastPosition = DumpPos
          .Esperando = 0
     End With
     
     'Limpia los objetos que quedaron tira2.
     For LoopX = 1 To 100
         For LoopY = 1 To 100
             With MapData(ARENA_MAP, LoopX, LoopY)
                  'Hay objeto?
                  If .ObjInfo.objIndex <> 0 Then
                     'Flag por si hay salida.
                     esSalida = (.TileExit.Map <> 0)
                     'No es del mapa.
                     If Not ItemNoEsDeMapa(.ObjInfo.objIndex, esSalida) Then
                        'Erase :P
                        Call EraseObj(.ObjInfo.Amount, ARENA_MAP, LoopX, LoopY)
                     End If
                  End If
             End With
         Next LoopY
     Next LoopX
     
End With
 
End Sub
 
Sub Cancelar(ByRef CancelatedBy As String)
 
' @ Cancela el death.
 
Dim LoopX   As Long
Dim UIndex  As Integer
Dim UPos    As WorldPos
 
'Aviso.
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Deathmatch>", "Cancelado por : " & CancelatedBy & ".")
 
'Llevo los usuarios que entraron a ulla.
For LoopX = 1 To UBound(DeathMatch.Usuarios())
    UIndex = DeathMatch.Usuarios(LoopX).UserIndex
    'Hay usuario?
    If UIndex <> -1 Then
       'Est� logeado?
       If UserList(UIndex).ConnID <> -1 Then
          'Est� en death?
          If UserList(UIndex).Death Then
          'Reset el flag.
             UserList(UIndex).Death = False
             'Telep to anterior posici�n.
             Call AnteriorPos(UIndex, UPos)
             WarpUserChar UIndex, UPos.Map, UPos.X, UPos.Y, True
             
          End If
       End If
    End If
Next LoopX
 
'Limpia el tipo
Limpiar
 
End Sub
 
Sub ActivarNuevo(ByRef OrganizatedBy As String, ByVal Cupos As Byte, ByVal MinLevel As Byte, ByVal MaxLevel As Byte, ByVal cMago As Byte, ByVal cClerigo As Byte, ByVal cBardo As Byte, ByVal cPaladin As Byte, ByVal cAsesino As Byte, ByVal cCazador As Byte, ByVal cGuerrero As Byte, ByVal cDruida As Byte, ByVal cLadron As Byte, ByVal cBandido As Byte)

On Error GoTo ErrHandler
' @ Crea nuevo deathmatch.
 
Dim LoopX   As Long
 
'Limpia el tipo.
Limpiar

'Llena los datos nuevos.
With DeathMatch

     .Cupos = Cupos
     .Activo = True
     .MaxLevel = MaxLevel
     .MinLevel = MinLevel
     .EventStarted = False
     
     Dim T As Boolean
     
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
     
     
     
     'Redim array
     ReDim .Usuarios(1 To Cupos) As DeathUser
     
     'Lleno el array con -1s
     For LoopX = 1 To Cupos
         .Usuarios(LoopX).UserIndex = -1
     Next LoopX
     
     'Avisa al mundo.
     SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Deathmatch>", "Organizador: " & OrganizatedBy & " " & Cupos & " Cupos! para entrar /DEATH.")
     SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Deathmatch>", "Nivel minimo: " & DeathMatch.MinLevel & ", nivel m�ximo: " & .MaxLevel)
     
     If T Then
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Deathmatch>", "Clases prohibidas: " & IIf(cMago > 0, " MAGO ", "") & IIf(cClerigo > 0, " CLERIGO ", "") & IIf(cBardo > 0, " BARDO ", "") & IIf(cPaladin > 0, " PALADIN ", "") & IIf(cAsesino > 0, " ASESINO ", "") & IIf(cCazador > 0, " CAZADOR ", "") & IIf(cGuerrero > 0, " GUERRERO ", "") & IIf(cDruida > 0, " DRUIDA ", "") & IIf(cLadron > 0, " LADRON ", "") & IIf(cBandido > 0, " BANDIDO ", ""))
     End If
     
 '    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgnew("Deathmatch>","Quedan 3 minutos antes del auto-cancelamiento si no se llena el cupo.")
     
     'Set el tiempo de auto-cancelaci�n.
     .AutoCancelTime = TIEMPO_AUTOCANCEL
End With

Exit Sub
ErrHandler:
LogError "ActivarNuevo error. " & Err.Number & " " & Err.description
 
End Sub
 
Sub Ingresar(ByVal UserIndex As Integer)
 
' @ Usuario ingresa al death.
 
Dim LibreSlot   As Byte
Dim SumarCount  As Boolean
 
LibreSlot = ProximoSlot(SumarCount)
 
'No hay slot.
If Not LibreSlot <> 0 Then Exit Sub
 
With DeathMatch
     'Hay que sumar?
     If SumarCount Then .Ingresaron = .Ingresaron + 1
     
     'Lleno el usuario.
     .Usuarios(LibreSlot).LastPosition = UserList(UserIndex).Pos
     .Usuarios(LibreSlot).UserIndex = UserIndex
     
     'Llevo a la arena.
     Dim npos As WorldPos
     Dim Ppos As WorldPos
     Ppos.Map = ARENA_MAP
     Ppos.X = ARENA_X
     Ppos.Y = ARENA_Y
     UserList(UserIndex).Death = True
     Call ClosestLegalPos(Ppos, npos)
     WarpUserChar UserIndex, npos.Map, npos.X, npos.Y, True
     
     'Aviso..
     WriteConsoleMsgNew UserIndex, "Deathmatch>", "Has ingresado al deathmatch, eres el participante n�" & LibreSlot & "."
     
     'Lleno el cupo?
     If .Ingresaron >= .Cupos Then
         'Quito el tiempo de auto-cancelaci�n
         .AutoCancelTime = 0
         'Aviso que llen� el cupo
         SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Deathmatch>", "El cupo ha sido completado!")
         'Doy inicio
         Iniciar
     End If
     
End With
 
End Sub
 
Sub Cuenta()
 
' @ Cuenta regresiva y auto-cancel ac�.
 
Dim PacketToSend    As String
Dim CanSendPackage  As Boolean
 
With DeathMatch
     
    'Espera el ganador?
    If .Ganador.UserIndex <> 0 Then
       'Tiempo de espera
       If .Ganador.Esperando <> 0 Then
          'resta.
          .Ganador.Esperando = .Ganador.Esperando - 1
          'Llego al fin el tiempo.
          If Not .Ganador.Esperando <> 0 Then
             'Telep to anterior pos.
             WarpUserChar .Ganador.UserIndex, .Ganador.LastPosition.Map, .Ganador.LastPosition.X, .Ganador.LastPosition.Y, True
             'Aviso al usuario.
             WriteConsoleMsgNew .Ganador.UserIndex, "Deathmatch>", "El tiempo ha llegado a su fin, fuiste devuelto a tu posici�n anterior"
             'Limpiar.
             Limpiar
          End If
        End If
    End If
   
    'Hay cuenta?
    If .Cuenta <> 0 Then
        'Resta el tiempo.
        .Cuenta = .Cuenta - 1
       
        If .Cuenta > 1 Then
            SendData SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsgNew("Deathmatch>", "El death iniciar� en " & .Cuenta & " segundos.")
        ElseIf .Cuenta = 1 Then
            SendData SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsgNew("Deathmatch>", "El death iniciar� en 1 segundo!")
        ElseIf .Cuenta <= 0 Then
            SendData SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsgNew("Deathmatch>", "El death inici�! PELEEN!")
            MapInfo(ARENA_MAP).Pk = True
             .EventStarted = True
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
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en 2:30 minutos")
              Case 120
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en 2 minutos")
              Case 90
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en 1:30 minutos")
              Case 60
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en 1 minuto")
              Case 30
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en 30 segundos")
              'Avisa a los 15
              Case 15
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en 15 segundos")
              'Avisa a los 10
              Case 10
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en 10 segundos")
              'Avisa a los 5
              Case 5
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en 5 segundos")
              'Avisa a los 3,2,1.
              Case 1, 2, 3
                   CanSendPackage = True
                   PacketToSend = PrepareMessageConsoleMsgNew("Deathmatch>", "Auto-Cancelar� en " & .AutoCancelTime & " segundo/s")
              Case 0
                   CanSendPackage = False
                   Call Cancelar("Falta de participantes.")
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
 
Sub Iniciar()
 
' @ Inicia el evento.
 
Dim LoopX   As Long
 
With DeathMatch
     
     'Set la cuenta.
     .Cuenta = CUENTA_NUM
     
     'Aviso a los usuarios.
     For LoopX = 1 To UBound(.Usuarios())
         'Hay usuario?
         If .Usuarios(LoopX).UserIndex <> -1 Then
            'Est� logeado?
            If UserList(.Usuarios(LoopX).UserIndex).ConnID <> -1 Then
                WarpUserChar .Usuarios(LoopX).UserIndex, UserList(.Usuarios(LoopX).UserIndex).Pos.Map, UserList(.Usuarios(LoopX).UserIndex).Pos.X, UserList(.Usuarios(LoopX).UserIndex).Pos.Y, 0, 0
               WriteConsoleMsgNew .Usuarios(LoopX).UserIndex, "Deathmatch>", "Llen� el cupo! El deathmatch iniciar� en " & .Cuenta & " segundos!."
            Else    'No loged, limpio el tipo
               .Usuarios(LoopX).UserIndex = -1
            End If
         End If
     Next LoopX
   
    'Por default el mapa es seguro..
    MapInfo(ARENA_MAP).Pk = False
     
End With
 
End Sub
 
Sub MuereUser(ByVal MuertoIndex As Integer, Optional ByVal sMessage As Boolean = True)
 
 
On Error GoTo ErrHandler
' @ Muere usuario en dm.
 
Dim MuertoPos       As WorldPos
Dim QuedanEnDeath   As Byte
 
'Obtengo la anterior posici�n del usuario
Call AnteriorPos(MuertoIndex, MuertoPos)
 
 
'Revivir usuario
If UserList(MuertoIndex).flags.Muerto <> 0 Then RevivirUsuario MuertoIndex
 
'Llenar vida.
UserList(MuertoIndex).Stats.MinHp = UserList(MuertoIndex).Stats.MaxHP
 
'Actualizar hp.
WriteUpdateHP MuertoIndex
 
'Reset el flag.
UserList(MuertoIndex).Death = False
 
'Telep anterior pos.
WarpUserChar MuertoIndex, MuertoPos.Map, MuertoPos.X, MuertoPos.Y, True
 
 If sMessage Then
'Aviso al usuario
WriteConsoleMsgNew MuertoIndex, "Deathmatch>", "Has caido en el deathMatch, has sido revivido y llevado a tu posici�n anterior (Mapa : " & MapInfo(MuertoPos.Map).Name & ")"
 
'Aviso al mapa.
SendData SendTarget.toMap, ARENA_MAP, PrepareMessageConsoleMsgNew("Deathmatch>", UserList(MuertoIndex).Name & " Ha sido derrotado.")
 End If
 
 
 If Not DeathMatch.EventStarted Then Exit Sub
 
'Obtengo los usuarios que quedan..
QuedanEnDeath = QuedanVivos()
 
'Queda 1?
If Not QuedanEnDeath <> 1 Then
   'Gan� ese usuario!
   Terminar
End If


Exit Sub
ErrHandler:
LogError "ActivarNuevo error. " & Err.Number & " " & Err.description
   
End Sub
 
Sub Terminar()
 
' @ Termina el death y gana un usuario.
 On Error GoTo ErrHandler
 
Dim winnerIndex As Integer
Dim GoldPremio  As Long
 
winnerIndex = GanadorIndex
 
'No hay ganador!! TRAGEDIAA XDD
If Not winnerIndex <> -1 Then
   SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsgNew("Deathmatch>", "TRAGEDIA EN DEATHMATCHS!! WINNERINDEX = -1!!!!")
   Limpiar
   Exit Sub
End If
 
'Hay ganador, le doi el premio..
GoldPremio = (PREMIO_POR_CABEZA * DeathMatch.Cupos)
UserList(winnerIndex).Stats.GLD = UserList(winnerIndex).Stats.GLD + GoldPremio
 
 
UserList(winnerIndex).Death = False
 
UserList(winnerIndex).flags.PuntosShop = UserList(winnerIndex).flags.PuntosShop + UserList(winnerIndex).flags.tmpDeath

'Actualizo el oro
WriteUpdateGold winnerIndex
 
With UserList(winnerIndex)
    'Aviso al mundo.
    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Deathmatch>", .Name & " - " & ListaClases(.clase) & " " & ListaRazas(.raza) & " Nivel " & .Stats.ELV & " Gan� " & Format$(GoldPremio, "#,###") & " monedas de oro y " & UserList(winnerIndex).flags.tmpDeath & " puntos de Canje (Por haber matado a " & UserList(winnerIndex).flags.tmpDeath & " usuarios) por salir primero en el evento.")
End With
 
'Ganador a su anterior posici�n..
Dim ToPosition  As WorldPos
Call AnteriorPos(winnerIndex, ToPosition)
 
'Warp.
WarpUserChar winnerIndex, ToPosition.Map, ToPosition.X, ToPosition.Y, True
 
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Deathmatch>", "Finalizado")

Limpiar

Exit Sub
ErrHandler:
LogError "Error en Terminar de Death." & Err.Number & " - " & Err.description
Limpiar
 
End Sub
 
Sub AnteriorPos(ByVal UserIndex As Integer, ByRef MuertoPosition As WorldPos)
 
' @ Devuelve la posici�n anterior del usuario.
 
Dim LoopX   As Long
 
For LoopX = 1 To UBound(DeathMatch.Usuarios())
    If DeathMatch.Usuarios(LoopX).UserIndex = UserIndex Then
       MuertoPosition = DeathMatch.Usuarios(LoopX).LastPosition
       Exit Sub
    End If
Next LoopX
 
'Posici�n de ulla u.u
MuertoPosition = Ullathorpe
 
End Sub
 
Function AprobarIngreso(ByVal UserIndex As Integer, ByRef MensajeError As String) As Boolean
 
' @ Checks si puede ingresar al death.
 
Dim DumpBoolean As Boolean
 
AprobarIngreso = False
 
'No hay death.
If Not DeathMatch.Activo Then
   MensajeError = "No hay deathmatch en curso."
   Exit Function
End If
 
'No hay cupos.
If Not ProximoSlot(DumpBoolean) <> 0 Then
   MensajeError = "Hay un deathmatch, pero las inscripciones est�n cerradas"
   Exit Function
End If
 
'Ya inscripto?
If YaInscripto(UserIndex) Then
   MensajeError = "Ya est�s en el death!"
   Exit Function
End If
 
'Est� muerto
If UserList(UserIndex).flags.Muerto <> 0 Then
   MensajeError = "Muerto no puedes ingresar a un deathmatch, lo siento.."
   Exit Function
End If
 
'Est� preso
If UserList(UserIndex).Counters.Pena <> 0 Then
   MensajeError = "No puedes ingresar si est�s preso."
   Exit Function
End If

If UserList(UserIndex).Pos.Map <> 1 Then
   MensajeError = "No puedes ingresar si est�s fuera de Ullathorpe."
   Exit Function
End If

If DeathMatch.MinLevel > UserList(UserIndex).Stats.ELV Then
    MensajeError = "El m�nimo nivel para entrar es de " & DeathMatch.MinLevel
    Exit Function
End If

If DeathMatch.MaxLevel < UserList(UserIndex).Stats.ELV Then
    MensajeError = "El m�ximo nivel para entrar es de " & DeathMatch.MaxLevel
    Exit Function
End If
 
 Dim MiClase As String
 MiClase = ListaClases(UserList(UserIndex).clase)
 
 On Error GoTo NoHayClases
If MiClase = ListaClases(DeathMatch.ClasesValidas.Asesino) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Bandido) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Bardo) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Cazador) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Clerigo) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Druida) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Guerrero) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Ladron) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Mago) Or _
    MiClase = ListaClases(DeathMatch.ClasesValidas.Paladin) Then
    
        MensajeError = "Tu clase no es v�lida."
        Exit Function

End If
   
   AprobarIngreso = True
Exit Function

NoHayClases:
 
    AprobarIngreso = True

End Function
 
Function ProximoSlot(ByRef Sumar As Boolean) As Byte
 
' @ Posici�n para un usuario.
 
Dim LoopX   As Long
 
Sumar = False
 
For LoopX = 1 To UBound(DeathMatch.Usuarios())
    'No hay usuario.
    If Not DeathMatch.Usuarios(LoopX).UserIndex <> -1 Then
       'Slot encontrado.
       ProximoSlot = LoopX
       'Hay que sumar el contador?
       If DeathMatch.Ingresaron < ProximoSlot Then Sumar = True
       Exit Function
    End If
Next LoopX
 
ProximoSlot = 0
 
End Function
 
Function QuedanVivos() As Byte
 
' @ Devuelve la cantidad de usuarios vivos que quedan.
 
Dim LoopX   As Long
Dim Counter As Byte
 
For LoopX = 1 To UBound(DeathMatch.Usuarios())
    'Mientras halla usuario.
    If DeathMatch.Usuarios(LoopX).UserIndex <> -1 Then
       'Mientras est� logeado
       If UserList(DeathMatch.Usuarios(LoopX).UserIndex).ConnID <> -1 Then
          'Mientras est� en el mapa de death
          If Not UserList(DeathMatch.Usuarios(LoopX).UserIndex).Pos.Map <> ARENA_MAP Then
             'Sumo contador.
             Counter = Counter + 1
           End If
        End If
    End If
Next LoopX
 
QuedanVivos = Counter
 
End Function
 
Function GanadorIndex() As Integer
 
' @ Busca el ganador..
On Error GoTo ErrHandler
Dim LoopX   As Long
 
For LoopX = 1 To UBound(DeathMatch.Usuarios())
    If DeathMatch.Usuarios(LoopX).UserIndex <> -1 Then
   
       If UserList(DeathMatch.Usuarios(LoopX).UserIndex).ConnID <> -1 Then
          If Not UserList(DeathMatch.Usuarios(LoopX).UserIndex).Pos.Map <> ARENA_MAP Then
         
             If Not UserList(DeathMatch.Usuarios(LoopX).UserIndex).flags.Muerto <> 0 Then
                GanadorIndex = DeathMatch.Usuarios(LoopX).UserIndex
                Exit Function
             End If
             
           End If
        End If
       
    End If
Next LoopX
 
'No hay ganador! WTF!!!
GanadorIndex = -1

ErrHandler:
LogError "Err en GandorIndex Death. " & GanadorIndex & " - " & Err.Number & " " & Err.description
GanadorIndex = -1

 
End Function
 
Function YaInscripto(ByVal UserIndex As Integer) As Boolean
 
' @ Devuelve si ya est� inscripto.
 
Dim LoopX   As Long
 
For LoopX = 1 To UBound(DeathMatch.Usuarios())
    If DeathMatch.Usuarios(LoopX).UserIndex = UserIndex Then
       YaInscripto = True
       Exit Function
    End If
Next LoopX
 
YaInscripto = False
 
End Function
 
Function GetBanqueroPos() As WorldPos
 
' @ Devuelve una posici�n para el banquero.
 
'No hay objeto.
If Not MapData(ARENA_MAP, BANCO_X, BANCO_Y).ObjInfo.objIndex <> 0 Then
   'Si no hay usuario me quedo con esta pos.
   If Not MapData(ARENA_MAP, BANCO_X, BANCO_Y).UserIndex <> 0 Then
      GetBanqueroPos.Map = ARENA_MAP
      GetBanqueroPos.X = BANCO_X
      GetBanqueroPos.Y = BANCO_Y
      Exit Function
   End If
End If
 
'Si no estaba libre el anterior tile, busco uno en un radio de 5 tiles.
Dim LoopX   As Long
Dim LoopY   As Long
 
For LoopX = (BANCO_X - 5) To (BANCO_X + 5)
    For LoopY = (BANCO_Y - 5) To (BANCO_Y + 5)
        With MapData(ARENA_MAP, LoopX, LoopY)
             'No hay un objeto..
             If Not .ObjInfo.objIndex <> 0 Then
                'No hay usuario.
                If Not .UserIndex <> 0 Then
                   'Nos quedamos ac�.
                   GetBanqueroPos.Map = ARENA_MAP
                   GetBanqueroPos.X = LoopX
                   GetBanqueroPos.Y = LoopY
                   Exit Function
                End If
             End If
        End With
    Next LoopY
Next LoopX
 
'Poco probable, pero bueno, si no hay una posici�n libre
'Devolvemos la posici�n "ORIGINAL", lo peor que puede pasar
'Es pisar un objeto : P
GetBanqueroPos.Map = ARENA_MAP
GetBanqueroPos.X = BANCO_X
GetBanqueroPos.Y = BANCO_Y
 
End Function
