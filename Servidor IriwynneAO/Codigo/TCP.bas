Attribute VB_Name = "TCP"
Option Explicit

Public Function DarCabezaNueva(ByVal UserIndex As Integer, _
                               ByRef NewHead As Integer) As Boolean
        '*************************************************
        'Author: Miqueas
        'Last modified: 19/05/2015
        'otorga una nueva cabeza al usuario
        '*************************************************

        Dim UserRaza   As eRaza

        Dim UserGenero As eGenero
        
        UserRaza = UserList(UserIndex).raza
        UserGenero = UserList(UserIndex).Genero

        Select Case UserGenero

                Case eGenero.Hombre

                        Select Case UserRaza

                                Case eRaza.Humano
                                        NewHead = RandomNumber(1, 40)

                                Case eRaza.Elfo
                                        NewHead = RandomNumber(101, 122)

                                Case eRaza.Drow
                                        NewHead = RandomNumber(201, 221)

                                Case eRaza.Enano
                                        NewHead = RandomNumber(301, 319)

                                Case eRaza.Gnomo
                                        NewHead = RandomNumber(401, 416)

                        End Select

                Case eGenero.Mujer

                        Select Case UserRaza

                                Case eRaza.Humano
                                        NewHead = RandomNumber(70, 89)
                 
                                Case eRaza.Elfo
                                        NewHead = RandomNumber(170, 178)

                                Case eRaza.Drow
                                        NewHead = RandomNumber(270, 288)

                                Case eRaza.Gnomo
                                        NewHead = RandomNumber(370, 384)

                                Case eRaza.Enano
                                        NewHead = RandomNumber(470, 486)

                        End Select

        End Select

        DarCabezaNueva = True

        Exit Function
    
End Function

Sub DarCuerpo(ByVal UserIndex As Integer)

        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 14/03/2007
        'Elije una cabeza para el usuario y le da un body
        '*************************************************
        Dim NewBody    As Integer

        Dim UserRaza   As Byte

        Dim UserGenero As Byte

        UserGenero = UserList(UserIndex).Genero
        UserRaza = UserList(UserIndex).raza

        Select Case UserGenero

                Case eGenero.Hombre

                        Select Case UserRaza

                                Case eRaza.Humano
                                        NewBody = 1

                                Case eRaza.Elfo
                                        NewBody = 2

                                Case eRaza.Drow
                                        NewBody = 3

                                Case eRaza.Enano
                                        NewBody = 300

                                Case eRaza.Gnomo
                                        NewBody = 300

                        End Select

                Case eGenero.Mujer

                        Select Case UserRaza

                                Case eRaza.Humano
                                        NewBody = 1

                                Case eRaza.Elfo
                                        NewBody = 2

                                Case eRaza.Drow
                                        NewBody = 3

                                Case eRaza.Gnomo
                                        NewBody = 300

                                Case eRaza.Enano
                                        NewBody = 300

                        End Select

        End Select

        UserList(UserIndex).Char.Body = NewBody

End Sub

Private Function ValidarCabeza(ByVal UserRaza As Byte, _
                               ByVal UserGenero As Byte, _
                               ByVal Head As Integer) As Boolean

        Select Case UserGenero

                Case eGenero.Hombre

                        Select Case UserRaza

                                Case eRaza.Humano
                                        ValidarCabeza = (Head >= HUMANO_H_PRIMER_CABEZA And _
                                           Head <= HUMANO_H_ULTIMA_CABEZA)

                                Case eRaza.Elfo
                                        ValidarCabeza = (Head >= ELFO_H_PRIMER_CABEZA And _
                                           Head <= ELFO_H_ULTIMA_CABEZA)

                                Case eRaza.Drow
                                        ValidarCabeza = (Head >= DROW_H_PRIMER_CABEZA And _
                                           Head <= DROW_H_ULTIMA_CABEZA)

                                Case eRaza.Enano
                                        ValidarCabeza = (Head >= ENANO_H_PRIMER_CABEZA And _
                                           Head <= ENANO_H_ULTIMA_CABEZA)

                                Case eRaza.Gnomo
                                        ValidarCabeza = (Head >= GNOMO_H_PRIMER_CABEZA And _
                                           Head <= GNOMO_H_ULTIMA_CABEZA)

                        End Select
    
                Case eGenero.Mujer

                        Select Case UserRaza

                                Case eRaza.Humano
                                        ValidarCabeza = (Head >= HUMANO_M_PRIMER_CABEZA And _
                                           Head <= HUMANO_M_ULTIMA_CABEZA)

                                Case eRaza.Elfo
                                        ValidarCabeza = (Head >= ELFO_M_PRIMER_CABEZA And _
                                           Head <= ELFO_M_ULTIMA_CABEZA)

                                Case eRaza.Drow
                                        ValidarCabeza = (Head >= DROW_M_PRIMER_CABEZA And _
                                           Head <= DROW_M_ULTIMA_CABEZA)

                                Case eRaza.Enano
                                        ValidarCabeza = (Head >= ENANO_M_PRIMER_CABEZA And _
                                           Head <= ENANO_M_ULTIMA_CABEZA)

                                Case eRaza.Gnomo
                                        ValidarCabeza = (Head >= GNOMO_M_PRIMER_CABEZA And _
                                           Head <= GNOMO_M_ULTIMA_CABEZA)

                        End Select

        End Select
        
End Function

Function AsciiValidos(ByVal cad As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim car As Byte
        Dim i   As Integer

        cad = LCase$(cad)

        For i = 1 To Len(cad)
                car = Asc(mid$(cad, i, 1))
    
                If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
                        AsciiValidos = False

                        Exit Function

                End If
    
        Next i

        AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim car As Byte

        Dim i   As Integer

        cad = LCase$(cad)

        For i = 1 To Len(cad)
                car = Asc(mid$(cad, i, 1))
    
                If (car < 48 Or car > 57) Then
                        Numeric = False
                        Exit Function

                End If
    
        Next i

        Numeric = True

End Function

Function NombrePermitido(ByVal Nombre As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim i As Integer

        For i = 1 To UBound(ForbidenNames)

                If InStr(Nombre, ForbidenNames(i)) Then
                        NombrePermitido = False
                        Exit Function

                End If

        Next i

        NombrePermitido = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Integer

        For LoopC = 1 To NUMSKILLS

                If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
                        Exit Function

                        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100

                End If

        Next LoopC

        ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal UserIndex As Integer, _
                   ByRef Name As String, _
                   ByRef Password As String, _
                   ByVal UserRaza As eRaza, _
                   ByVal UserSexo As eGenero, _
                   ByVal Userclase As eClass, _
                   ByRef UserEmail As String, _
                   ByVal Hogar As eCiudad, _
                   ByVal Head As Integer, ByVal clave As String, ByVal discoDuro As String, ByVal CPU_ID As String)
        
        Dim i As Long

        With UserList(UserIndex)

                If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
                        Call WriteErrorMsg(UserIndex, "Nombre inválido.")
                        Exit Sub

                End If
    
                If UserList(UserIndex).flags.UserLogged Then
                        Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(UserIndex).Ip)
        
                        'Kick player ( and leave character inside :D )!
                        Call CloseSocketSL(UserIndex)
                        Call Cerrar_Usuario(UserIndex)
                        Exit Sub
                End If
    
                If (BuscarRegistroT0(CPU_ID) > 0) Then
                    Call WriteErrorMsg(UserIndex, "Usted estas baneado T0")
                    Call FlushBuffer(UserIndex)
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                '¿Existe el personaje?
                If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
                        Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
                        Exit Sub

                End If
    
                'Tiró los dados antes de llegar acá??
                'If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
                '        Call WriteErrorMsg(UserIndex, "Debe tirar los dados antes de poder crear un personaje.")
                '        Exit Sub

                'End If
    
                If Not ValidarCabeza(UserRaza, UserSexo, Head) Then
                        Call LogCheating("El usuario " & Name & " ha seleccionado la cabeza " & Head & " desde la IP " & .Ip)
        
                        Call WriteErrorMsg(UserIndex, "Cabeza inválida, elija una cabeza seleccionable.")
                        Exit Sub

                End If
    
                .flags.Muerto = 0
                .flags.Escondido = 0
                .flags.Vip = 0
                .discoDuro = discoDuro
                .CPU_ID = CPU_ID
                .Reputacion.AsesinoRep = 0
                .Reputacion.BandidoRep = 0
                .Reputacion.BurguesRep = 0
                .Reputacion.LadronesRep = 0
                .Reputacion.NobleRep = 1000
                .Reputacion.PlebeRep = 30
    
                .Reputacion.Promedio = 30 / 6
    
                .flags.Posteado = 0
                .flags.MeMando = 0
                .flags.Lemande = 0

                .Name = Name
                .clase = Userclase
                .raza = UserRaza
                .Genero = UserSexo
                .email = UserEmail
                .Hogar = Hogar
                .clave = clave
                
                '[Marcos Zeni (Miqueas150) 8/12/15]
                .Stats.UserAtributos(eAtributos.Fuerza) = Configuracion.Dados(0) + ModRaza(UserRaza).Fuerza
                .Stats.UserAtributos(eAtributos.Agilidad) = Configuracion.Dados(0) + ModRaza(UserRaza).Agilidad
                .Stats.UserAtributos(eAtributos.Inteligencia) = Configuracion.Dados(0) + ModRaza(UserRaza).Inteligencia
                .Stats.UserAtributos(eAtributos.Carisma) = Configuracion.Dados(0) + ModRaza(UserRaza).Carisma
                .Stats.UserAtributos(eAtributos.Constitucion) = Configuracion.Dados(0) + ModRaza(UserRaza).Constitucion
                '[/Marcos Zeni (Miqueas150)]
    
                For i = 1 To NUMSKILLS
                        .Stats.UserSkills(i) = 0
                Next i
    
                .Stats.SkillPts = 10
    
                .Char.heading = eHeading.SOUTH
    
                Call DarCuerpo(UserIndex)
                .Char.Head = Head
    
                .OrigChar = .Char
    
                Dim MiInt As Long

                'MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
                .Stats.MaxHP = 21 ' @@ 15 + MiInt
                .Stats.MinHp = 21 ' @@ 15 + MiInt
    
                MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)

                If MiInt = 1 Then MiInt = 2
    
                .Stats.MaxSta = 20 * MiInt
                .Stats.MinSta = 20 * MiInt
    
                .Stats.MaxAGU = 100
                .Stats.MinAGU = 100
    
                .Stats.MaxHam = 100
                .Stats.MinHam = 100
    
                '<-----------------MANA----------------------->
                If Userclase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
                        MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
                        .Stats.MaxMAN = MiInt
                        .Stats.MinMAN = MiInt
                ElseIf Userclase = eClass.Cleric Or Userclase = eClass.Druid _
                   Or Userclase = eClass.Bard Or Userclase = eClass.Assasin Then
                        .Stats.MaxMAN = 50
                        .Stats.MinMAN = 50
                ElseIf Userclase = eClass.Bandit Then 'Mana Inicial del Bandido (ToxicWaste)
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
                .Stats.Exp = 0
                .Stats.ELU = 300
                .Stats.ELV = 1
    
                '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
                Dim Slot      As Byte

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
    
                #If ConUpTime Then
                        .LogOnTime = Now
                        .UpTime = 0
                #End If

        End With

        'Valores Default de facciones al Activar nuevo usuario
        Call ResetFacciones(UserIndex)

        Call WriteVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password", Password) 'grabamos el password aqui afuera, para no mantenerlo cargado en memoria
        Call WriteVar(CharPath & UCase$(Name) & ".chr", "INIT", "VIP_DIAS", Date)
        
        Call SaveUser(UserIndex, CharPath & UCase$(Name) & ".chr")
  
        'Open User
        Call ConnectUser(UserIndex, Name, Password, discoDuro, CPU_ID)
  
End Sub

Sub CloseSocket(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error GoTo errhandler

        With UserList(UserIndex)
                Call SecurityIp.IpRestarConexion(GetLongIp(.Ip))
        
                If .ConnID <> -1 Then
                        Call CloseSocketSL(UserIndex)

                End If
                
                
        ' @@ Juegos del hambre
150     If .EnJDH Then
160         Call m_JuegosDelHambre.DesconectaUserJDH(UserIndex)
170     End If
        
                'Es el mismo user al que está revisando el centinela??
                'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
                ' y lo podemos loguear
                Dim CentinelaIndex As Byte

                CentinelaIndex = .flags.CentinelaIndex
        
                If CentinelaIndex <> 0 Then
                        Call modCentinela.CentinelaUserLogout(CentinelaIndex)

                End If
                
                If .flags.Automatico = True Then
                    Call Rondas_UsuarioDesconecta(UserIndex)
                End If
                
                If .Death = True Then
                    Call DesconectaUser(UserIndex)
                End If
                
                If .mReto.reto_Index <> 0 Then Call Mod_Retos1vs1.disconnectUser_reto(UserIndex)
                If .sReto.reto_used Then Call Mod_Retos2vs2.disconnect_Reto(UserIndex)
        
                If UserList(UserIndex).flags.List = 1 Then
                    UserList(UserIndex).flags.List = 0
                    Call ListEvent.Remove_User(UserIndex)
                End If
                
                'mato los comercios seguros
                If .ComUsu.DestUsu > 0 Then
                        If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                                        Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                                        Call FinComerciarUsu(.ComUsu.DestUsu)
                                        Call FlushBuffer(.ComUsu.DestUsu)

                                End If

                        End If

                End If
        
                'Empty buffer for reuse
                Call .incomingData.ReadASCIIStringFixed(.incomingData.length)
        
                If .flags.UserLogged Then
                        If NumUsers > 0 Then NumUsers = NumUsers - 1
                        Call CloseUser(UserIndex)
            
                Else
                        Call ResetUserSlot(UserIndex)

                End If
        
                Call FreeSlot(UserIndex)
                
                Call SendData(SendTarget.ToAll, 0, PrepareMessageMandarOnlines)
                
        End With

        Exit Sub

errhandler:
        Call ResetUserSlot(UserIndex)
        
        Call FreeSlot(UserIndex)
    
        Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & UserIndex)

End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
                Call WSApiCloseSocket(UserList(UserIndex).ConnID, UserIndex)
                UserList(UserIndex).ConnIDValida = False

        End If

End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, _
                                 ByRef Datos As String) As Long

        '***************************************************
        'Author: Unknown
        'Last Modification: 01/10/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
        '***************************************************
        On Error GoTo Err
    
        Dim ret As Long
    
        ret = WsApiEnviar(UserIndex, Datos)
    
        If ret <> 0 And ret <> WSAEWOULDBLOCK Then
                ' Close the socket avoiding any critical error
                Call CloseSocketSL(UserIndex)
                Call Cerrar_Usuario(UserIndex)

        End If

        Exit Function
    
Err:

End Function

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        Dim X As Integer, Y As Integer

        For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
                For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

                        If MapData(UserList(Index).Pos.Map, X, Y).UserIndex = Index2 Then
                                EstaPCarea = True
                                Exit Function

                        End If
        
                Next X
        Next Y

        EstaPCarea = False

End Function

Function HayPCarea(Pos As WorldPos) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim X As Integer, Y As Integer

        For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
                For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

                        If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                                        HayPCarea = True
                                        Exit Function

                                End If

                        End If

                Next X
        Next Y

        HayPCarea = False

End Function

Function HayOBJarea(Pos As WorldPos, objIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim X As Integer, Y As Integer

        For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
                For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

                        If MapData(Pos.Map, X, Y).ObjInfo.objIndex = objIndex Then
                                HayOBJarea = True
                                Exit Function

                        End If
        
                Next X
        Next Y

        HayOBJarea = False

End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        ValidateChr = UserList(UserIndex).Char.Head <> 0 _
           And UserList(UserIndex).Char.Body <> 0 _
           And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, _
                ByRef Name As String, _
                ByRef Password As String, ByVal discoDuro As String, ByVal CPU_ID As String)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 24/07/2010 (ZaMa)
        '26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
        '12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
        '14/09/2009: ZaMa - Ahora el usuario esta protegido del ataque de npcs al loguear
        '11/27/2009: Budi - Se envian los InvStats del personaje y su Fuerza y Agilidad
        '03/12/2009: Budi - Optimización del código
        '24/07/2010: ZaMa - La posicion de comienzo es namehuak, como se habia definido inicialmente.
        '***************************************************
        
        Dim N    As Integer

        Dim tStr As String

        With UserList(UserIndex)

                If .flags.UserLogged Then
                        Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .Ip)
                        'Kick player ( and leave character inside :D )!
                        Call CloseSocketSL(UserIndex)
                        Call Cerrar_Usuario(UserIndex)
                        Exit Sub

                End If
                
                .discoDuro = discoDuro
                
                'Reseteamos los FLAGS
                .flags.Escondido = 0
                .flags.TargetNPC = 0
                .flags.TargetNpcTipo = eNPCType.Comun
                .flags.TargetObj = 0
                .flags.TargetUser = 0
                .flags.LastNPCTalk = 0
                .Char.FX = 0
                .CPU_ID = CPU_ID
                .flags.MenuCliente = 255
                .flags.LastSlotClient = 255
    
                'Controlamos no pasar el maximo de usuarios
                If NumUsers >= MaxUsers Then
                        Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el máximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
                        Call FlushBuffer(UserIndex)
                        Call CloseSocket(UserIndex)
                        Exit Sub
                End If
                
                If (BuscarRegistroT0(CPU_ID) > 0) Then
                    Call WriteErrorMsg(UserIndex, "Usted esta baneado T0")
                    Call FlushBuffer(UserIndex)
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                '¿Este IP ya esta conectado?
                If AllowMultiLogins = 0 Then
                        If CheckForSameIP(UserIndex, .Ip) = True Then
                                Call WriteErrorMsg(UserIndex, "No es posible usar más de un personaje al mismo tiempo.")
                                Call FlushBuffer(UserIndex)
                                Call CloseSocket(UserIndex)
                                Exit Sub

                        End If

                End If
    
                '¿Existe el personaje?
                If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
                        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
                        Call FlushBuffer(UserIndex)
                        Call CloseSocket(UserIndex)
                        Exit Sub

                End If
    
                '¿Es el passwd valido?
                If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
                        Call WriteErrorMsg(UserIndex, "Password incorrecto.")
                        Call FlushBuffer(UserIndex)
                        Call CloseSocket(UserIndex)
                        Exit Sub

                End If
    
                '¿Ya esta conectado el personaje?
                If CheckForSameName(Name) Then
                        If UserList(NameIndex(Name)).Counters.Saliendo Then
                                Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
                        Else
                                Call WriteErrorMsg(UserIndex, "Perdón, un usuario con el mismo nombre se ha logueado.")

                        End If

                        Call FlushBuffer(UserIndex)
                        Call CloseSocket(UserIndex)
                        Exit Sub

                End If
    
                'Reseteamos los privilegios
                .flags.Privilegios = 0
    
                'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
                If EsAdmin(Name) Then
                        .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
                        Call LogGM(Name, "Se conecto con ip:" & .Ip)
                ElseIf EsDios(Name) Then
                        .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
                        Call LogGM(Name, "Se conecto con ip:" & .Ip)
                ElseIf EsSemiDios(Name) Then
                        .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        
                        .flags.PrivEspecial = EsGmEspecial(Name)
        
                        Call LogGM(Name, "Se conecto con ip:" & .Ip)
                ElseIf EsConsejero(Name) Then
                        .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
                        Call LogGM(Name, "Se conecto con ip:" & .Ip)
                Else
                        .flags.Privilegios = .flags.Privilegios Or PlayerType.User
                        .flags.AdminPerseguible = True

                End If
    
                'Add RM flag if needed
                If EsRolesMaster(Name) Then
                        .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster

                End If
    
                If ServerSoloGMs > 0 Then
                        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
                                Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                                Call FlushBuffer(UserIndex)
                                Call CloseSocket(UserIndex)
                                Exit Sub

                        End If

                End If
    
                'Cargamos el personaje
                Dim Leer As clsIniManager

                Set Leer = New clsIniManager
    
                Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")
    
                'Cargamos los datos del personaje
                Call LoadUserInit(UserIndex, Leer)
    
                Call LoadUserStats(UserIndex, Leer)
    
                'Cargamos los mensajes privados del usuario.
                Call CargarMensajes(UserIndex, Leer)
    
                If Not ValidateChr(UserIndex) Then
                        Call WriteErrorMsg(UserIndex, "Error en el personaje.")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                End If
    
                Call LoadUserReputacion(UserIndex, Leer)
                Call LoadQuestStats(UserIndex, Leer)    ' GSZAO
    
                Set Leer = Nothing
    
                If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
                If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
                If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma

                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS

                If (.flags.Muerto = 0) Then
                        .flags.SeguroResu = False
                        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff)
                Else
                        .flags.SeguroResu = True
                        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn)

                End If
    
                Call UpdateUserInv(True, UserIndex, 0)
                Call UpdateUserHechizos(True, UserIndex, 0)
                
                If .flags.Paralizado Then
                        Call WriteParalizeOK(UserIndex)

                End If
                
                ' @@ Miqueas : Los Gms loguean Siempre en el mapaGM
                If esGM(UserIndex) Then
                        .Pos.Map = Configuracion.MapaGm
                        .Pos.X = 50
                        .Pos.Y = 50

                End If

                Dim Mapa As Integer

                Mapa = .Pos.Map
    
                'Posicion de comienzo
                If Mapa = 0 Then
                        .Pos = Ullathorpe
                        Mapa = Ullathorpe.Map
                Else
    
                        If Not MapaValido(Mapa) Then
                                Call WriteErrorMsg(UserIndex, "El PJ se encuenta en un mapa inválido.")
                                Call CloseSocket(UserIndex)
                                Exit Sub

                        End If
        
                        ' If map has different initial coords, update it
                        Dim StartMap As Integer

                        StartMap = MapInfo(Mapa).StartPos.Map

                        If StartMap <> 0 Then
                                If MapaValido(StartMap) Then
                                        .Pos = MapInfo(Mapa).StartPos
                                        Mapa = StartMap

                                End If

                        End If
        
                End If
    
                'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
                'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
                If MapData(Mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(Mapa, .Pos.X, .Pos.Y).NpcIndex <> 0 Then

                        Dim FoundPlace As Boolean

                        Dim esAgua     As Boolean

                        Dim tX         As Long

                        Dim tY         As Long
        
                        FoundPlace = False
                        esAgua = HayAgua(Mapa, .Pos.X, .Pos.Y)
        
                        For tY = .Pos.Y - 1 To .Pos.Y + 1
                                For tX = .Pos.X - 1 To .Pos.X + 1

                                        If esAgua Then

                                                'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                                                If LegalPos(Mapa, tX, tY, True, False) Then
                                                        FoundPlace = True
                                                        Exit For

                                                End If

                                        Else

                                                'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                                                If LegalPos(Mapa, tX, tY, False, True) Then
                                                        FoundPlace = True
                                                        Exit For

                                                End If

                                        End If

                                Next tX
            
                                If FoundPlace Then _
                                   Exit For
                        Next tY
        
                        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                                .Pos.X = tX
                                .Pos.Y = tY
                        Else

                                'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
                                If MapData(Mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Then

                                        'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                                        If UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then

                                                'Le avisamos al que estaba comerciando que se tuvo que ir.
                                                If UserList(UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                                                        Call FinComerciarUsu(UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                                                        Call WriteConsoleMsg(UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                                                        Call FlushBuffer(UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)

                                                End If

                                                'Lo sacamos.
                                                If UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                                                        Call FinComerciarUsu(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex)
                                                        Call WriteErrorMsg(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                                                        Call FlushBuffer(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex)

                                                End If

                                        End If
                
                                        Call CloseSocket(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex)

                                End If

                        End If

                End If
    
                'Nombre de sistema
                .Name = Name
    
                .showName = True 'Por default los nombres son visibles
    
                'If in the water, and has a boat, equip it!
                If .Invent.BarcoObjIndex > 0 And _
                   (HayAgua(Mapa, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.Body)) Then

                        .Char.Head = 0

                        If .flags.Muerto = 0 Then
                                Call ToggleBoatBody(UserIndex)
                        Else
                                .Char.Body = iFragataFantasmal
                                .Char.ShieldAnim = NingunEscudo
                                .Char.WeaponAnim = NingunArma
                                .Char.CascoAnim = NingunCasco

                        End If
        
                        .flags.Navegando = 1

                End If
    
                'Info
                Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
                Call WriteChangeMap(UserIndex, .Pos.Map, MapInfo(.Pos.Map).MapVersion) 'Carga el mapa
                Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(.Pos.Map).Music, 45)))
    
                If .flags.Privilegios = PlayerType.Dios Then
                        .flags.ChatColor = RGB(250, 250, 150)
                ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
                        .flags.ChatColor = RGB(0, 255, 0)
                ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
                        .flags.ChatColor = RGB(0, 255, 255)
                ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
                        .flags.ChatColor = RGB(255, 128, 64)
                Else
                        .flags.ChatColor = vbWhite

                End If
    
                ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
                #If ConUpTime Then
                        .LogOnTime = Now
                #End If
    
                'Crea  el personaje del usuario
                Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
    
                If (.flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster)) = 0 Then
                        Call DoAdminInvisible(UserIndex)
                        .flags.SendDenounces = True
                Else

                        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))

                        End If

                End If
    
                Call WriteUserCharIndexInServer(UserIndex)
                
                ''[/el oso]
    
                Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
    
                Call CheckUserLevel(UserIndex)
                Call WriteUpdateUserStats(UserIndex)
    
                Call WriteUpdateHungerAndThirst(UserIndex)
                Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
                Call SendMOTD(UserIndex)
    
                If haciendoBK Then
                        Call WritePauseToggle(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER)

                End If
    
                If EnPausa Then
                        Call WritePauseToggle(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)

                End If
    
                If EnTesting And .Stats.ELV >= 18 Then
                        Call WriteErrorMsg(UserIndex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
                        Call FlushBuffer(UserIndex)
                        Call CloseSocket(UserIndex)
                        Exit Sub

                End If
    
                If TieneMensajesNuevos(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "¡Tienes mensajes privados sin leer!", FontTypeNames.FONTTYPE_FIGHT)

                End If
    
                'Actualiza el Num de usuarios
                'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
                NumUsers = NumUsers + 1
                .flags.UserLogged = True
                
                
                Call SendData(SendTarget.ToAll, 0, PrepareMessageMandarOnlines)
                
                
                'usado para borrar Pjs
                Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "1")

                MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
    
                If .Stats.SkillPts > 0 Then
                        Call WriteSendSkills(UserIndex)
                        Call WriteLevelUp(UserIndex, 1, .Stats.SkillPts)

                End If
          
                Dim VariableUsuarios As Integer

                VariableUsuarios = CInt(NumUsers) 'CInt(NumUsers * 2.1) 'kevin counter strike

                If VariableUsuarios > RECORDusuarios Then
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaneamente." & "Hay " & VariableUsuarios & " usuarios.", FontTypeNames.FONTTYPE_INFO))
                        RECORDusuarios = VariableUsuarios
                        Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str$(RECORDusuarios))
   
                End If
    
                If .NroMascotas > 0 And MapInfo(.Pos.Map).Pk Then

                        Dim i As Long

                        For i = 1 To MAXMASCOTAS

                                If .MascotasType(i) > 0 Then
                                        .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
                
                                        If .MascotasIndex(i) > 0 Then
                                                Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
                                                Call FollowAmo(.MascotasIndex(i))
                                        Else
                                                .MascotasIndex(i) = 0

                                        End If

                                End If

                        Next i

                End If
    
                If .flags.Navegando = 1 Then
                        Call WriteNavigateToggle(UserIndex)

                End If
    
                If criminal(UserIndex) Then
                        Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
                        .flags.Seguro = False
                Else
                        .flags.Seguro = True
                        Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)

                End If
    
                If .GuildIndex > 0 Then
            
                        'welcome to the show baby...
                        If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
                                Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)

                        End If

                End If
    
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
    
                Call WriteLoggedMessage(UserIndex)
    
                Call modGuilds.SendGuildNews(UserIndex)
    
                ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
                Call IntervaloPermiteSerAtacado(UserIndex, True)
    
                If Lloviendo Then
                       Call WriteRainToggle(UserIndex)

                End If
    
                tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)
    
                If LenB(tStr) <> 0 Then
                        Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)

                End If
                
                If RandomNumber(1, 2) = 2 Then EnviarInfoCastillo UserIndex
                
                With ControlMensajes

                        If .Activado = 1 Then
                 
                                Dim LoopC As Long
                      
                                For LoopC = 1 To UBound(.Mensajes())
                                        'Call WriteConsoleMsg(UserIndex, "Anuncio> " & .Mensajes(LoopC), FontTypeNames.FONTTYPE_DIOS)
                                        Call WriteConsoleMsgNew(UserIndex, "Anuncio>", .Mensajes(LoopC), FontTypeNames.FONTTYPE_VERDE, FontTypeNames.FONTTYPE_DIOS)
                                Next LoopC

                        End If
                    
                
                End With
                
                
                If .flags.Vip > 0 Then
                    'checkeamos.
                    Dim DiasVIP As String
                    DiasVIP = GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "VIP_DIAS")
                    Debug.Print DateDiff("d", Date, DiasVIP)
                    If DateDiff("d", Date, DiasVIP) < 0 Then
                        .flags.Vip = 0
                        WriteConsoleMsgNew UserIndex, "VIP>", "TU VIP HA VENCIDO !!!", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO
                        
                        WriteVar CharPath & UCase$(Name) & ".chr", "INIT", "VIP_DIAS", Date
                        
                        WarpUserChar UserIndex, .Pos.Map, .Pos.X, .Pos.Y, False, False
                    Else
                        If DateDiff("d", Date, DiasVIP) = 0 Then
                        WriteConsoleMsgNew UserIndex, "VIP>", "Es tu último dia de VIP.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO
                        
                        Else
                        WriteConsoleMsgNew UserIndex, "VIP>", "Te quedan " & DateDiff("d", Date, DiasVIP) & " días VIP.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO
                        End If
                    End If
                    
                End If
                
                
                If .GuildIndex > 0 Then
                
                    If Len(GetVar(App.Path & "\GUILDS\guildsinfo.inf", "GUILD" & .GuildIndex, "RecibioSolicitud")) > 0 Then
                    Dim tmps As Integer
                    tmps = GetVar(App.Path & "\GUILDS\guildsinfo.inf", "GUILD" & .GuildIndex, "RecibioSolicitud")
                        'If Len(tmps) > 0 Then
                            If tmps > 0 Then
                                WriteConsoleMsg UserIndex, "TIENES UNA SOLICITUD DE GUERRA DEL CLAN: " & GuildName(tmps), FontTypeNames.FONTTYPE_CENTINELA
                            End If
                        'End If
                    End If
                End If
                
                'Load the user statistics
                Call Statistics.UserConnected(UserIndex)
    
                Call MostrarNumUsers
    
                WriteChauClienteEditado UserIndex
                
                #If SeguridadAlkon Then
                        Call Security.UserConnected(UserIndex)
                #End If

                N = FreeFile
                Open App.Path & "\logs\numusers.log" For Output As N
                Print #N, NumUsers
                Close #N
    
                N = FreeFile
                'Log
                Open App.Path & "\logs\Connect.log" For Append Shared As #N
                Print #N, .Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & time & " " & Date
                Close #N
                 
                
                Call CheckRankingUser(UserIndex, topnivel)
                
        End With

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        'Dim j As Long
    
        'Call WriteGuildChat(UserIndex, "Mensajes de entrada:", True)

        'For j = 1 To MaxLines
        '        Call WriteGuildChat(UserIndex, MOTD(j).texto, True)
        'Next j

End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '*************************************************
        
        With UserList(UserIndex).Faccion
                .ArmadaReal = 0
                .CiudadanosMatados = 0
                .CriminalesMatados = 0
                .FuerzasCaos = 0
                .FechaIngreso = "No ingresó a ninguna Facción"
                .RecibioArmaduraCaos = 0
                .RecibioArmaduraReal = 0
                .RecibioExpInicialCaos = 0
                .RecibioExpInicialReal = 0
                .RecompensasCaos = 0
                .RecompensasReal = 0
                .Reenlistadas = 0
                .NivelIngreso = 0
                .MatadosIngreso = 0
                .NextRecompensa = 0

        End With

End Sub

Sub ResetContadores(ByVal UserIndex As Integer)

        '*************************************************
        'Author: Unknown
        'Last modified: 10/07/2010
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '05/20/2007 Integer - Agregue todas las variables que faltaban.
        '10/07/2010: ZaMa - Agrego los counters que faltaban.
        '*************************************************
        
        With UserList(UserIndex).Counters
            .HablaDeath = 0
                .AGUACounter = 0
                .LastCentinela = 0
                .bPuedeMeditar = True
                .Ceguera = 0
                .COMCounter = 0
                .Estupidez = 0
                .Frio = 0
                .goHome = 0
                .HPCounter = 0
                .IdleCount = 0
                .Invisibilidad = 0
                .Lava = 0
                .Congelado = 0
                .Mimetismo = 0
                .Ocultando = 0
                .Paralisis = 0
                .Pena = 0
                .PiqueteC = 0
                .Saliendo = False
                .Salir = 0
                .STACounter = 0
                .TiempoOculto = 0
                
                .TimerEstadoAtacable = 0
                .TimerGolpeMagia = 0
                .TimerGolpeUsar = 0
                .TimerLanzarSpell = 0
                .TimerMagiaGolpe = 0
                .TimerPerteneceNpc = 0
                .TimerPuedeAtacar = 0
                .TimerPuedeSerAtacado = 0
                .TimerPuedeTrabajar = 0
                .TimerPuedeUsarArco = 0
                .TimerUsar = 0
                .TimerUsarClick = 0
                .failedUsageAttempts = 0
                
                .IntervaloGolpe = 0
                .IntervaloHechizo = 0
                .LastPoteo = 0
                
                .Trabajando = 0
                .Veneno = 0

        End With

End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        With UserList(UserIndex).Char
                .Body = 0
                .CascoAnim = 0
                .CharIndex = 0
                .FX = 0
                .Head = 0
                .Loops = 0
                .heading = 0
                .Loops = 0
                .ShieldAnim = 0
                .WeaponAnim = 0

        End With

End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        With UserList(UserIndex)
                .Name = vbNullString
                .desc = vbNullString
                .DescRM = vbNullString
                .Pos.Map = 0
                .Pos.X = 0
                .Pos.Y = 0
                .Ip = vbNullString
                .clase = 0
                .email = vbNullString
                .Genero = 0
                .Hogar = 0
                .raza = 0
        
                .PartyIndex = 0
                .PartySolicitud = 0
        
                With .Stats
                        .Banco = 0
                        .ELV = 0
                        .ELU = 0
                        .Exp = 0
                        .def = 0
                        '.CriminalesMatados = 0
                        .NPCsMuertos = 0
                        .UsuariosMatados = 0
                        .SkillPts = 0
                        .GLD = 0
                        .UserAtributos(1) = 0
                        .UserAtributos(2) = 0
                        .UserAtributos(3) = 0
                        .UserAtributos(4) = 0
                        .UserAtributos(5) = 0
                        .UserAtributosBackUP(1) = 0
                        .UserAtributosBackUP(2) = 0
                        .UserAtributosBackUP(3) = 0
                        .UserAtributosBackUP(4) = 0
                        .UserAtributosBackUP(5) = 0

                End With
        
        End With

End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        With UserList(UserIndex).Reputacion
                .AsesinoRep = 0
                .BandidoRep = 0
                .BurguesRep = 0
                .LadronesRep = 0
                .NobleRep = 0
                .PlebeRep = 0
                .NobleRep = 0
                .Promedio = 0

        End With

End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If UserList(UserIndex).EscucheClan > 0 Then
                Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
                UserList(UserIndex).EscucheClan = 0

        End If

        If UserList(UserIndex).GuildIndex > 0 Then
                Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)

        End If

        UserList(UserIndex).GuildIndex = 0

End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)

        '*************************************************
        'Author: Unknown
        'Last modified: 06/28/2008
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '03/29/2006 Maraxus - Reseteo el CentinelaOK también.
        '06/28/2008 NicoNZ - Agrego el flag Inmovilizado
        '*************************************************
        With UserList(UserIndex).flags
        .RequiredKills = 0
            .YaPediAyuda = False
                .Automatico = False
                .PuntosShop = 0
                .tmpDeath = 0
                UserList(UserIndex).Death = False
                .Comerciando = False
                .Ban = 0
                .Escondido = 0
                .DuracionEfecto = 0
                .NpcInv = 0
                .StatsChanged = 0
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .TargetObj = 0
                .TargetObjMap = 0
                .TargetObjX = 0
                .TargetObjY = 0
                .TargetUser = 0
                .TipoPocion = 0
                .TomoPocion = False
                .Vip = 0
                .Descuento = vbNullString
                .Hambre = 0
                .Sed = 0
                .Descansar = False
                .Navegando = 0
                .Montando = 0
                .Oculto = 0
                .Envenenado = 0
                .invisible = 0
                .Paralizado = 0
                .Inmovilizado = 0
                .Maldicion = 0
                .Bendicion = 0
                .Meditando = 0
                .Privilegios = 0
                .PrivEspecial = False
                .PuedeMoverse = 0
                .OldBody = 0
                .OldHead = 0
                .AdminInvisible = 0
                .ValCoDe = 0
                .Hechizo = 0
                .TimesWalk = 0
                .StartWalk = 0
                .CountSH = 0
                .Silenciado = 0
                .CentinelaOK = False
                .CentinelaIndex = 0
                .AdminPerseguible = False
                .lastMap = 0
                .Traveling = 0
                .AtacablePor = 0
                .AtacadoPorNpc = 0
                .AtacadoPorUser = 0
                .NoPuedeSerAtacado = False
                .ShareNpcWith = 0
                .EnConsulta = False
                .Ignorado = False
                .SendDenounces = False
                .ParalizedBy = vbNullString
                .ParalizedByIndex = 0
                .ParalizedByNpcIndex = 0
        
                If .OwnedNpc <> 0 Then
                        Call PerdioNpc(UserIndex)

                End If
                
                .MenuCliente = 0
                .LastSlotClient = 0
        
        End With

End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Long

        For LoopC = 1 To MAXUSERHECHIZOS
                UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
        Next LoopC

End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Long
    
        UserList(UserIndex).NroMascotas = 0
        
        For LoopC = 1 To MAXMASCOTAS
                UserList(UserIndex).MascotasIndex(LoopC) = 0
                UserList(UserIndex).MascotasType(LoopC) = 0
        Next LoopC

End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Long
    
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
                UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
                UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
                UserList(UserIndex).BancoInvent.Object(LoopC).objIndex = 0
        Next LoopC
    
        UserList(UserIndex).BancoInvent.NroItems = 0

End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        With UserList(UserIndex).ComUsu

                If .DestUsu > 0 Then
                        Call FinComerciarUsu(.DestUsu)
                        Call FinComerciarUsu(UserIndex)

                End If

        End With

End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim i As Long

        UserList(UserIndex).ConnIDValida = False
        UserList(UserIndex).ConnID = -1

        Call LimpiarComercioSeguro(UserIndex)
        Call ResetFacciones(UserIndex)
        Call ResetContadores(UserIndex)
        Call ResetGuildInfo(UserIndex)
        Call ResetCharInfo(UserIndex)
        Call ResetBasicUserInfo(UserIndex)
        Call ResetReputacion(UserIndex)
        Call ResetUserFlags(UserIndex)
        Call LimpiarInventario(UserIndex)
        Call ResetUserSpells(UserIndex)
        Call ResetUserPets(UserIndex)
        Call ResetUserBanco(UserIndex)
        Call ResetQuestStats(UserIndex)    ' GSZAO
        Call LimpiarMensajes(UserIndex)

        With UserList(UserIndex).ComUsu
                .Acepto = False
    
                For i = 1 To MAX_OFFER_SLOTS
                        .Cant(i) = 0
                        .Objeto(i) = 0
                Next i
    
                .GoldAmount = 0
                .DestNick = vbNullString
                .DestUsu = 0

        End With
        
        With UserList(UserIndex).sReto
                .accept_count = 0
                .nick_sender = vbNullString
                .reto_Index = 0
                .reto_used = False
                .return_city = 0
                .acceptedOK = False
                .acceptLimit = 0
                .tmp_Time = 0

        End With
         
        With UserList(UserIndex).mReto
                .acceptLimitCount = 0
                .request_name = vbNullString
                .reto_Index = 0
                .return_home = 0
                .send_to_index = 0
                .temp_dropGamble = False
                .temp_goldGamble = 0

        End With
 
End Sub

Sub CloseUser(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error GoTo errhandler

        Dim N    As Integer

        Dim Map  As Integer

        Dim Name As String

        Dim i    As Integer

        Dim aN   As Integer

        With UserList(UserIndex)
        
                If UserIndex = GranPoder Then
                    GranPoder = 0
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Favor de los dioses>", .Name & " ha perdido el poder.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO))
                    'Call OtorgarFavordelosDioses(0)
                End If
                
                aN = .flags.AtacadoPorNpc

                If aN > 0 Then
                        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                        Npclist(aN).flags.AttackedBy = vbNullString

                End If
    
                aN = .flags.NPCAtacado

                If aN > 0 Then
                        If Npclist(aN).flags.AttackedFirstBy = .Name Then
                                Npclist(aN).flags.AttackedFirstBy = vbNullString

                        End If

                End If

                .flags.AtacadoPorNpc = 0
                .flags.NPCAtacado = 0
    
                Map = .Pos.Map
                Name = UCase$(.Name)
    
                .Char.FX = 0
                .Char.Loops = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
    
                .flags.UserLogged = False
                .Counters.Saliendo = False
    
                'Le devolvemos el body y head originales
                If .flags.AdminInvisible = 1 Then
                        .Char.Body = .flags.OldBody
                        .Char.Head = .flags.OldHead
                        .flags.AdminInvisible = 0

                End If
    
                'si esta en party le devolvemos la experiencia
                If .PartyIndex > 0 Then Call mdParty.SalirDeParty(UserIndex)
    
                'Save statistics
                Call Statistics.UserDisconnected(UserIndex)
    
                ' Grabamos el personaje del usuario
                Call SaveUser(UserIndex, CharPath & Name & ".chr")
    
                'usado para borrar Pjs
                Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "0")
    
                'Quitar el dialogo
                'If MapInfo(Map).NumUsers > 0 Then
                '    Call SendToUserArea(UserIndex, "QDL" & .Char.charindex)
                'End If
    
                If MapInfo(Map).NumUsers > 0 Then
                        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))

                End If
    
                'Borrar el personaje
                If .Char.CharIndex > 0 Then
                        Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)

                End If
    
                'Borrar mascotas
                For i = 1 To MAXMASCOTAS

                        If .MascotasIndex(i) > 0 Then
                                If Npclist(.MascotasIndex(i)).flags.NPCActive Then _
                                   Call QuitarNPC(.MascotasIndex(i))

                        End If

                Next i
    
                'Update Map Users
                MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
    
                If MapInfo(Map).NumUsers < 0 Then
                        MapInfo(Map).NumUsers = 0

                End If
    
                ' Si el usuario habia dejado un msg en la gm's queue lo borramos
                If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)
    
                Call ResetUserSlot(UserIndex)
    
                Call MostrarNumUsers
    
                N = FreeFile(1)
                Open App.Path & "\logs\Connect.log" For Append Shared As #N
                Print #N, Name & " ha dejado el juego. " & "User Index:" & UserIndex & " " & time & " " & Date
                Close #N

        End With

        Exit Sub

errhandler:
        Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.description)

End Sub

Sub ReloadSokcet()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error GoTo errhandler

        Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
        If NumUsers <= 0 Then
                Call WSApiReiniciarSockets
        Else

                '       Call apiclosesocket(SockListen)
                '       SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
        End If

        Exit Sub
errhandler:
        Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EnviarNoche(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Call WriteSendNight(UserIndex, IIf(DeNoche And (MapInfo(UserList(UserIndex).Pos.Map).Zona = eTerrain.terrain_campo Or MapInfo(UserList(UserIndex).Pos.Map).Zona = eTerrain.terrain_ciudad), True, False))
        Call WriteSendNight(UserIndex, IIf(DeNoche, True, False))

End Sub

Public Sub EcharPjsNoPrivilegiados()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Long
    
        For LoopC = 1 To LastUser

                If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
                        If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                                Call CloseSocket(LoopC)

                        End If

                End If

        Next LoopC

End Sub
