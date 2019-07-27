Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Const AumentoSTDef As Byte = 15
Public Const AumentoStBandido As Byte = AumentoSTDef + 23
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTTrabajador As Byte = AumentoSTDef + 25
Public Const AdicionalHPGuerrero As Byte = 2    'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1    'HP adicionales cuando sube de nivel


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Sub ActStats(ByVal victimIndex As Integer, ByVal attackerIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Ahora no te vuelve cirminal por matar un atacable
'***************************************************

    Dim DaExp  As Integer

    Dim EraCriminal As Boolean

    DaExp = CInt(UserList(victimIndex).Stats.ELV) * 2

    With UserList(attackerIndex)
    
        If .Death = True Then
            .flags.tmpDeath = .flags.tmpDeath + 1
        End If
        
        .Stats.Exp = .Stats.Exp + DaExp

        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP

        If TriggerZonaPelea(victimIndex, attackerIndex) <> TRIGGER6_PERMITE Then

            ' Es legal matarlo si estaba en atacable
            If UserList(victimIndex).flags.AtacablePor <> attackerIndex And .pos.Map <> Castillo.Mapa Then
                EraCriminal = criminal(attackerIndex)

                With .Reputacion

                    If Not criminal(victimIndex) Then
                        .AsesinoRep = .AsesinoRep + vlASESINO * 2

                        If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
                        .BurguesRep = 0
                        .NobleRep = 0
                        .PlebeRep = 0
                    Else
                        .NobleRep = .NobleRep + vlNoble

                        If .NobleRep > MAXREP Then .NobleRep = MAXREP

                    End If

                End With

                Dim esCriminal As Boolean

                esCriminal = criminal(attackerIndex)

                If EraCriminal <> esCriminal Then
                    Call RefreshCharStatus(attackerIndex)

                End If

            End If

        End If

        'Lo mata
        Call WriteMultiMessage(attackerIndex, eMessages.HaveKilledUser, victimIndex, DaExp)
        Call WriteMultiMessage(victimIndex, eMessages.UserKill, attackerIndex)

        Call FlushBuffer(victimIndex)

        'Log
        Call LogAsesinato(.Name & " asesino a " & UserList(victimIndex).Name)

        Dim tmpCantMatados As Integer
        ' @@ CHEQUEO DE SISTEMA DE GUERRAS BY CUICUI JAJAJA (no lo vean, da asco)
        If UserList(victimIndex).GuildIndex > 0 And UserList(attackerIndex).GuildIndex > 0 Then
            If guilds(UserList(victimIndex).GuildIndex).GetRelacion(UserList(attackerIndex).GuildIndex) = GUERRA Then
                ' @@ es la primera muerte?
                If Len(GetVar(App.Path & "\GUILDS\guildsinfo.inf", "GUILD" & UserList(attackerIndex).GuildIndex, "CantidadMatados")) <= 0 Then
                    Call WriteVar(App.Path & "\GUILDS\guildsinfo.inf", "GUILD" & UserList(attackerIndex).GuildIndex, "CantidadMatados", 1)
                Else
                    tmpCantMatados = GetVar(App.Path & "\GUILDS\guildsinfo.inf", "GUILD" & UserList(attackerIndex).GuildIndex, "CantidadMatados")
                    Call WriteVar(App.Path & "\GUILDS\guildsinfo.inf", "GUILD" & UserList(attackerIndex).GuildIndex, "CantidadMatados", tmpCantMatados + 1)
                End If
            End If
        End If
        
    End With

End Sub

Public Sub RevivirUsuario(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        With UserList(UserIndex)
                .flags.Muerto = 0
                .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion)
        
                If .Stats.MinHp > .Stats.MaxHP Then
                        .Stats.MinHp = .Stats.MaxHP

                End If
        
        If .flags.Montando = 1 Then
       Call WriteMontateToggle(UserIndex)
    End If
    
                If .flags.Navegando = 1 Then
                        Call ToggleBoatBody(UserIndex)
                Else
                        Call DarCuerpoDesnudo(UserIndex)
            
                        .Char.Head = .OrigChar.Head

                End If
        
                If .flags.Traveling Then
                        Call EndTravel(UserIndex, True)

                End If
        
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                Call WriteUpdateUserStats(UserIndex)

        End With

End Sub

Public Sub ToggleBoatBody(ByVal UserIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 25/07/2010
        'Gives boat body depending on user alignment.
        '25/07/2010: ZaMa - Now makes difference depending on faccion and atacable status.
        '***************************************************

        Dim Ropaje        As Integer

        Dim EsFaccionario As Boolean

        Dim NewBody       As Integer
    
        With UserList(UserIndex)
 
                .Char.Head = 0

                If .Invent.BarcoObjIndex = 0 Then Exit Sub
        
                Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
        
                ' Criminales y caos
                If criminal(UserIndex) Then
            
                        EsFaccionario = esCaos(UserIndex)
            
                        Select Case Ropaje

                                Case iBarca

                                        If EsFaccionario Then
                                                NewBody = iFragataCaos 'iBarcaCaos  @@ Miqueas : Por ahora desactivados
                                        Else
                                                NewBody = iBarcaPk

                                        End If
                
                                Case iGalera

                                        If EsFaccionario Then
                                                NewBody = iFragataCaos 'iGaleraCaos  @@ Miqueas : Por ahora desactivados
                                        Else
                                                NewBody = iGaleraPk

                                        End If
                    
                                Case iGaleon

                                        If EsFaccionario Then
                                                NewBody = iFragataCaos 'iGaleonCaos  @@ Miqueas : Por ahora desactivados
                                        Else
                                                NewBody = iGaleonPk

                                        End If

                        End Select
        
                        ' Ciudas y Armadas
                Else
            
                        EsFaccionario = esArmada(UserIndex)
            
                        Select Case Ropaje

                                Case iBarca

                                        If EsFaccionario Then
                                                NewBody = iFragataReal 'iBarcaReal  @@ Miqueas : Por ahora desactivados
                                        Else
                                                NewBody = iBarcaCiuda

                                        End If
                    
                                Case iGalera

                                        If EsFaccionario Then
                                                NewBody = iFragataReal   'iGaleraReal  @@ Miqueas : Por ahora desactivados
                                        Else
                                                NewBody = iGaleraCiuda

                                        End If
                        
                                Case iGaleon

                                        If EsFaccionario Then
                                                NewBody = iFragataReal  'iGaleonReal  @@ Miqueas : Por ahora desactivados
                                        Else
                                                NewBody = iGaleonCiuda

                                        End If

                        End Select
            
                End If
        
                .Char.Body = NewBody
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco

        End With

End Sub

Public Sub ChangeUserChar(ByVal UserIndex As Integer, _
                          ByVal Body As Integer, _
                          ByVal Head As Integer, _
                          ByVal heading As Byte, _
                          ByVal Arma As Integer, _
                          ByVal Escudo As Integer, _
                          ByVal casco As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        If UserList(UserIndex).Death Then
        Exit Sub
        End If
        
        With UserList(UserIndex).Char
                .Body = Body
                .Head = Head
                .heading = heading
                .WeaponAnim = Arma
                .ShieldAnim = Escudo
                .CascoAnim = casco
        
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(Body, Head, heading, .CharIndex, Arma, Escudo, .FX, .Loops, casco))

        End With

End Sub

Public Function GetWeaponAnim(ByVal UserIndex As Integer, _
                              ByVal objIndex As Integer) As Integer

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 03/29/10
        '
        '***************************************************
        Dim Tmp As Integer

        With UserList(UserIndex)
                Tmp = ObjData(objIndex).WeaponRazaEnanaAnim
            
                If Tmp > 0 Then
                        If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
                                GetWeaponAnim = Tmp
                                Exit Function

                        End If

                End If
        
                GetWeaponAnim = ObjData(objIndex).WeaponAnim

        End With

End Function

Public Sub EnviarFama(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim L As Long
    
        With UserList(UserIndex).Reputacion
                L = (-.AsesinoRep) + _
                   (-.BandidoRep) + _
                   .BurguesRep + _
                   (-.LadronesRep) + _
                   .NobleRep + _
                   .PlebeRep
                L = Round(L / 6)
        
                .Promedio = L

        End With
    
        Call WriteFame(UserIndex)

End Sub

Public Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean)
        '*************************************************
        'Author: Unknown
        'Last modified: 08/01/2009
        '08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
        '*************************************************

        On Error GoTo ErrorHandler
    
        With UserList(UserIndex)
                CharList(.Char.CharIndex) = 0
        
                If .Char.CharIndex = LastChar Then

                        Do Until CharList(LastChar) > 0
                                LastChar = LastChar - 1

                                If LastChar <= 1 Then Exit Do
                        Loop

                End If
        
                ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
             '   If IsAdminInvisible Then
               '         Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
              '  Else
                        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))

                'End If
        
                Call QuitarUser(UserIndex, .pos.Map)
        
                'MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0

                If InMapBounds(.pos.Map, .pos.X, .pos.Y) Then
                        MapData(.pos.Map, .pos.X, .pos.Y).UserIndex = 0

                End If
            
                .Char.CharIndex = 0

        End With
    
        NumChars = NumChars - 1
        Exit Sub
    
ErrorHandler:
    
        Dim UserName  As String

        Dim CharIndex As Integer
    
        If UserIndex > 0 Then
                UserName = UserList(UserIndex).Name
                CharIndex = UserList(UserIndex).Char.CharIndex

        End If

        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description & _
           ". User: " & UserName & "(UI: " & UserIndex & " - CI: " & CharIndex & ")")

End Sub

Public Sub RefreshCharStatus(ByVal UserIndex As Integer)

        '*************************************************
        'Author: Tararira
        'Last modified: 04/07/2009
        'Refreshes the status and tag of UserIndex.
        '04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
        '*************************************************
        Dim ClanTag   As String

        Dim NickColor As Byte
    
        With UserList(UserIndex)

                If .GuildIndex > 0 Then
                        ClanTag = modGuilds.GuildName(.GuildIndex)
                        ClanTag = " <" & ClanTag & ">"

                End If
        
                NickColor = GetNickColor(UserIndex)
        
                If .showName Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .Name & ClanTag))
                Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, vbNullString))

                End If
        
                'Si esta navengando, se cambia la barca.
                If .flags.Navegando Then
                        If .flags.Muerto = 1 Then
                                .Char.Body = iFragataFantasmal
                        Else
                                Call ToggleBoatBody(UserIndex)

                        End If
            
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                End If
                
                'ustedes se preguntaran que hace esto aca?
                'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
                'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
                'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable
                Dim NuevaA As Boolean
 
                Dim GI     As Integer
 
                Dim tStr   As String
 
                GI = .GuildIndex
 
                If GI > 0 Then
                        NuevaA = False
         
                        If Not modGuilds.m_ValidarPermanencia(UserIndex, True, NuevaA) Then
                                Call WriteConsoleMsg(UserIndex, "Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!", FontTypeNames.FONTTYPE_GUILD)

                        End If
 
                        If NuevaA Then
                                Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan ha pasado a tener alineación " & modGuilds.GuildAlignment(GI) & "!", FontTypeNames.FONTTYPE_GUILD))
                                tStr = modGuilds.GuildName(GI)
                                Call LogClanes("¡El clan " & tStr & " cambio de alineación!")

                        End If
           
                End If

        End With

End Sub

Public Function GetNickColor(ByVal UserIndex As Integer) As Byte
        '*************************************************
        'Author: ZaMa
        'Last modified: 15/01/2010
        '
        '*************************************************
    
        With UserList(UserIndex)
        
                If criminal(UserIndex) Then
                        GetNickColor = eNickColor.ieCriminal
                Else
                        GetNickColor = eNickColor.ieCiudadano

                End If
        
                If .flags.AtacablePor > 0 Then GetNickColor = GetNickColor Or eNickColor.ieAtacable

                If GranPoder = UserIndex Then GetNickColor = eNickColor.ieGranPoder
                    
                If .Death Then GetNickColor = eNickColor.ieCiudadano
                
        End With
    
End Function

Public Sub MakeUserChar(ByVal toMap As Boolean, _
                        ByVal sndIndex As Integer, _
                        ByVal UserIndex As Integer, _
                        ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        Optional ButIndex As Boolean = False)
'*************************************************
'Author: Unknown
'Last modified: 15/01/2010
'23/07/2009: Budi - Ahora se envía el nick
'15/01/2010: ZaMa - Ahora se envia el color del nick.
'*************************************************

    On Error GoTo errhandleR

    Dim CharIndex As Integer

    Dim ClanTag As String

    Dim NickColor As Byte

    Dim UserName As String

    Dim Privileges As Byte

    With UserList(UserIndex)

        If InMapBounds(Map, X, Y) Then

            'If needed make a new character in list
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = UserIndex

            End If

            'Place character on map if needed
            If toMap Then MapData(Map, X, Y).UserIndex = UserIndex

            'Send make character command to clients
            If Not toMap Then
                If .GuildIndex > 0 Then
                    ClanTag = modGuilds.GuildName(.GuildIndex)

                End If

                NickColor = GetNickColor(UserIndex)
                Privileges = .flags.Privilegios

                'Preparo el nick
                If .showName Then
                    UserName = .Name

                    If .flags.EnConsulta Then
                        UserName = UserName & " " & TAG_CONSULT_MODE
                    Else

                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                            If LenB(ClanTag) <> 0 Then _
                               UserName = UserName & " <" & ClanTag & ">"
                        Else

                            If (.flags.invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) Then
                                UserName = UserName & " " & TAG_USER_INVISIBLE
                            Else

                                If LenB(ClanTag) <> 0 Then _
                                   UserName = UserName & " <" & ClanTag & ">"

                            End If

                        End If

                    End If

                End If

                Dim esConquista As Boolean

                If Castillo.Dueño > 0 Then
                    If .GuildIndex > 0 Then
                        If .GuildIndex = Castillo.Dueño Then
                            esConquista = True
                        End If
                    End If
                End If

                If .Death Then
                
                    Call WriteCharacterCreate(sndIndex, 1, 1, .Char.heading, _
                                              .Char.CharIndex, X, Y, _
                                              1, 1, .Char.FX, 999, 1, _
                                              "Participante", NickColor, Privileges, 0, 0, .flags.Vip, 0)

                Else
                    Call WriteCharacterCreate(sndIndex, .Char.Body, .Char.Head, .Char.heading, _
                                              .Char.CharIndex, X, Y, _
                                              .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, _
                                              UserName, NickColor, Privileges, 0, 0, .flags.Vip, esConquista)
                End If


            Else
                'Hide the name and clan - set privs as normal user
                Call AgregarUser(UserIndex, .pos.Map, ButIndex)

            End If

        End If

    End With

    Exit Sub

errhandleR:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    'Resume Next
    Call CloseSocket(UserIndex)

End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal UserIndex As Integer)

'*************************************************
'Author: Unknown
'Last modified: 08/04/2011
'Chequea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
'07/08/2006 Integer - Modificacion de los valores
'01/10/2007 Tavo - Corregido el BUG de Configuracion.NivelMaximo
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
'13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
'09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
'12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y está en un clan, se lo expulsa para no sumar antifacción
'02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
'11/19/2009 Pato - Modifico la nueva fórmula de maná ganada para el bandido y se la limito a 499
'02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
'08/04/2011: Amraphen - Arreglada la distribución de probabilidades para la vida en el caso de promedio entero.
'*************************************************
10  On Error GoTo errhandleR

    Dim Pts    As Long
    Dim AumentoHIT As Integer
    Dim AumentoMANA As Integer
    Dim AumentoSTA As Integer
    Dim AumentoHP As Integer
    Dim WasNewbie As Boolean

    'Dim Promedio As Double
    'Dim aux As Integer
    'Dim DistVida(1 To 5) As Integer

    Dim GI     As Integer                'Guild Index

20  WasNewbie = EsNewbie(UserIndex)

30  With UserList(UserIndex)

40      Do While .Stats.Exp >= .Stats.ELU          ' Asi esta en la 13.3

            'Checkea si alcanzó el máximo nivel
50          If .Stats.ELV >= 47 Then
60              .Stats.Exp = 0
70              .Stats.ELU = 0
80              Call WriteUpdateUserStats(UserIndex)
90              Exit Sub
100         End If

            'Store it!
110         Call Statistics.UserLevelUp(UserIndex)

120         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .pos.X, .pos.Y))
130         Call WriteConsoleMsg(UserIndex, "Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)         ' ¡Has subido de nivel!

140         If .Stats.ELV = 1 Then
150             Pts = 10
160         Else
                'For multiple levels being rised at once
170             Pts = Pts + 5
180         End If

190         .Stats.ELV = .Stats.ELV + 1

200         .Stats.Exp = .Stats.Exp - .Stats.ELU

            .Stats.ELU = Get_ExpLvl(.Stats.ELV)

740         Select Case .clase

            Case eClass.Warrior
750             Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
760                 AumentoHP = RandomNumber(8, 12)
770             Case 20
780                 AumentoHP = RandomNumber(8, 12)
790             Case 19
800                 AumentoHP = RandomNumber(8, 11)
810             Case 18
820                 AumentoHP = RandomNumber(8, 11)
830             Case Else
840                 AumentoHP = RandomNumber(8, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
850             End Select

860             AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
870             AumentoSTA = AumentoSTDef

880         Case eClass.Hunter
890             Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
900                 AumentoHP = RandomNumber(9, 11)
910             Case 20
920                 AumentoHP = RandomNumber(8, 11)
930             Case 19
940                 AumentoHP = RandomNumber(8, 10)
950             Case 18
960                 AumentoHP = RandomNumber(7, 9)
970             Case Else
980                 AumentoHP = RandomNumber(7, 10) '.Stats.UserAtributos(eAtributos.Constitucion) \ 2)
990             End Select

1000            AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
1010            AumentoSTA = AumentoSTDef

1020        Case eClass.Pirat
1030            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
1040                AumentoHP = RandomNumber(9, 11)
1050            Case 20
1060                AumentoHP = RandomNumber(8, 11)
1070            Case 19
1080                AumentoHP = RandomNumber(7, 11)
1090            Case 18
1100                AumentoHP = RandomNumber(7, 11)
1110            Case Else
1120                AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
1130            End Select

1140            AumentoHIT = 2
1150            AumentoSTA = AumentoSTDef

1160        Case eClass.Paladin
1170            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
1180                AumentoHP = RandomNumber(8, 11)
1190            Case 20
1200                AumentoHP = RandomNumber(8, 10)
1210            Case 19
1220                AumentoHP = RandomNumber(8, 10)
1230            Case 18
1240                AumentoHP = RandomNumber(7, 9)
1250            Case Else
1260                AumentoHP = RandomNumber(7, 10)    ' .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPCazador
1270            End Select

1280            AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
1290            AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
1300            AumentoSTA = AumentoSTDef

1310        Case eClass.Thief
1320            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
1330                AumentoHP = RandomNumber(6, 9)
1340            Case 20
1350                AumentoHP = RandomNumber(6, 9)
1360            Case 19
1370                AumentoHP = RandomNumber(5, 9)
1380            Case 18
1390                AumentoHP = RandomNumber(5, 8)
1400            Case Else
1410                AumentoHP = RandomNumber(5, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
1420            End Select

1430            AumentoHIT = 2
1440            AumentoSTA = AumentoSTLadron

1450        Case eClass.Mage
1460            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
1470                AumentoHP = RandomNumber(7, 9)
1480            Case 20
1490                AumentoHP = RandomNumber(7, 8)
1500            Case 19
1510                AumentoHP = RandomNumber(6, 7)
1520            Case 18
1530                AumentoHP = RandomNumber(5, 7)
1540            Case Else
1550                AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
1560            End Select
1570            If AumentoHP < 1 Then AumentoHP = 4

1580            AumentoHIT = 1
                'AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
1590            AumentoSTA = AumentoSTMago

1600            If (.Stats.MaxMAN >= 2000) Then
1610                AumentoMANA = (3 * .Stats.UserAtributos(eAtributos.Inteligencia)) / 2
1620            Else
1630                AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
1640            End If

1650        Case eClass.Worker
1660            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
1670                AumentoHP = RandomNumber(9, 12)
1680            Case 20
1690                AumentoHP = RandomNumber(8, 12)
1700            Case 19
1710                AumentoHP = RandomNumber(7, 11)
1720            Case 18
1730                AumentoHP = RandomNumber(6, 11)
1740            Case Else
1750                AumentoHP = RandomNumber(5, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
1760            End Select

1770            AumentoHIT = 2
1780            AumentoSTA = AumentoSTTrabajador

1790        Case eClass.Cleric
1800            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
1810                AumentoHP = RandomNumber(7, 10)
1820            Case 20
1830                AumentoHP = RandomNumber(7, 9)
1840            Case 19
1850                AumentoHP = RandomNumber(7, 8)
1860            Case 18
1870                AumentoHP = RandomNumber(6, 9)
1880            Case Else
1890                AumentoHP = RandomNumber(7, 8)    '.Stats.UserAtributos(eAtributos.Constitucion) \ 2)
1900            End Select

1910            AumentoHIT = 2
1920            AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
1930            AumentoSTA = AumentoSTDef

1940        Case eClass.Druid
1950            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
1960                AumentoHP = RandomNumber(7, 9)
1970            Case 20
1980                AumentoHP = RandomNumber(7, 9)
1990            Case 19
2000                AumentoHP = RandomNumber(6, 9)
2010            Case 18
2020                AumentoHP = RandomNumber(6, 9)
2030            Case Else
2040                AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2050            End Select

2060            AumentoHIT = 2
2070            AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
2080            AumentoSTA = AumentoSTDef

2090        Case eClass.Assasin
2100            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
2110                AumentoHP = RandomNumber(7, 10)
2120            Case 20
2130                AumentoHP = RandomNumber(7, 9)
2140            Case 19
2150                AumentoHP = RandomNumber(7, 8)
2160            Case 18
2170                AumentoHP = RandomNumber(6, 8)
2180            Case Else
2190                AumentoHP = RandomNumber(7, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2200            End Select

2210            AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
2220            AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
2230            AumentoSTA = AumentoSTDef

2240        Case eClass.Bard
2250            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
2260                AumentoHP = RandomNumber(7, 10)
2270            Case 20
2280                AumentoHP = RandomNumber(7, 9)
2290            Case 19
2300                AumentoHP = RandomNumber(7, 8)
2310            Case 18
2320                AumentoHP = RandomNumber(6, 9)
2330            Case Else
2340                AumentoHP = RandomNumber(6, 8)    '.Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2350            End Select

2360            AumentoHIT = 2
2370            AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
2380            AumentoSTA = AumentoSTDef
            
            
            Case eClass.Bandit

                Select Case .Stats.UserAtributos(eAtributos.Constitucion)

                Case 21
                    AumentoHP = RandomNumber(9, 11)

                Case 20
                    AumentoHP = RandomNumber(8, 11)

                Case 19
                    AumentoHP = RandomNumber(8, 10)

                Case 18
                    AumentoHP = RandomNumber(7, 10)

                Case Else
                    AumentoHP = RandomNumber(7, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)

                End Select

                AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) / 3 * 2
                AumentoSTA = AumentoStBandido


2390        Case Else
2400            Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
2410                AumentoHP = RandomNumber(7, 9)
2420            Case 20
2430                AumentoHP = RandomNumber(7, 9)
2440            Case 19
2450                AumentoHP = RandomNumber(7, 8)
2460            Case 18
2470                AumentoHP = RandomNumber(6, 9)
2480            Case Else
2490                AumentoHP = RandomNumber(6, 9)    '.Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
2500            End Select

2510            AumentoHIT = 2
2520            AumentoSTA = AumentoSTDef
2530        End Select

            If AumentoHP < 8 And .clase = eClass.Paladin And .raza <> eRaza.Gnomo Then
                If RandomNumber(1, 2) > 1 Then
                    AumentoHP = AumentoHP + 1
                End If
            End If
            
            If .raza = eRaza.Humano Then
                
                If AumentoHP <= 9 Then
                    
                    If RandomNumber(1, 2) = 1 Then
                        AumentoHP = AumentoHP + 1
                    End If
                    
                End If
                
            End If
            
            'Actualizamos HitPoints
2540        .Stats.MaxHP = .Stats.MaxHP + AumentoHP
2550        If .Stats.MaxHP > STAT_MAXHP Then .Stats.MaxHP = STAT_MAXHP

            'Actualizamos Stamina
2560        .Stats.MaxSta = .Stats.MaxSta + AumentoSTA
2570        If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA

            'Actualizamos Mana
2580        .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
2590        If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN

            'Actualizamos Golpe Máximo
2600        .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT
2610        If .Stats.ELV < 36 Then
2620            If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
                   .Stats.MaxHIT = STAT_MAXHIT_UNDER36
2630        Else
2640            If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
                   .Stats.MaxHIT = STAT_MAXHIT_OVER36
2650        End If

            'Actualizamos Golpe Mínimo
2660        .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
2670        If .Stats.ELV < 36 Then
2680            If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
                   .Stats.MinHIT = STAT_MAXHIT_UNDER36
2690        Else
2700            If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
                   .Stats.MinHIT = STAT_MAXHIT_OVER36
2710        End If

            'Notificamos al user

2720        If AumentoHP > 0 Then
2730            Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
2740        End If

2750        If AumentoSTA > 0 Then
2760            Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de energía.", FontTypeNames.FONTTYPE_INFO)
2770        End If

2780        If AumentoMANA > 0 Then
2790            Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de manà.", FontTypeNames.FONTTYPE_INFO)
2800        End If

            ' @@ Le decimos las ups del personaje
            'Call CalcularPromedio(Userindex)

            'If AumentoHIT > 0 Then
            '    Call WriteConsoleMsg(UserIndex, "Tu golpe máximo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
            '    Call WriteConsoleMsg(UserIndex, "Tu golpe mínimo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
            'End If

2810        .Stats.MinHp = .Stats.MaxHP

            'If user is in a party, we modify the variable p_sumaniveleselevados
2820        Call mdParty.ActualizarSumaNivelesElevados(UserIndex)
            'If user reaches lvl 25 and he is in a guild, we check the guild's alignment and expulses the user if guild has factionary alignment

            If .Stats.ELV = 25 Then
                GI = .GuildIndex

                If GI > 0 Then
21                  If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                        'We get here, so guild has factionary alignment, we have to expulse the user
22                      Call modGuilds.m_EcharMiembroDeClan(-1, .Name)
23                      Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
24                      Call WriteConsoleMsg(UserIndex, "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo.", FontTypeNames.FONTTYPE_GUILD)

                    End If

                End If

            End If


            If EsNewbie(UserIndex) Then

                Dim tmpOro As Integer

                tmpOro = 500 * .Stats.ELV
                .Stats.GLD = .Stats.GLD + tmpOro

                Call WriteConsoleMsg(UserIndex, "Has ganado " & tmpOro & " monedas de oro, por subir de nivel mientras eres newbie", FontTypeNames.FONTTYPE_INFO)
            End If

2930    Loop

        'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
        If Not EsNewbie(UserIndex) And WasNewbie Then
16          Call QuitarNewbieObj(UserIndex)
17
18          If MapInfo(.pos.Map).Restringir = eRestrict.restrict_newbie Then
19              Call WarpUserChar(UserIndex, 1, 50, 50, True)
201             Call WriteConsoleMsg(UserIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        'Send all gained skill points at once (if any)
3030    If Pts > 0 Then


3040        Call WriteLevelUp(UserIndex, 1, Pts)
3050        .Stats.SkillPts = .Stats.SkillPts + Pts
3060        Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
3070    End If

3080 End With

3090 Call WriteUpdateUserStats(UserIndex)

    Call CheckRankingUser(UserIndex, topnivel)
3100 Exit Sub

errhandleR:

    Dim UserName As String
    Dim UserMap As Integer

3110 If UserIndex > 0 Then
3120    UserName = UserList(UserIndex).Name
3130    UserMap = UserList(UserIndex).pos.Map
3140 End If

3150 Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description & ". User: " & UserName & "(" & UserIndex & "). Map: " & UserMap)
End Sub

Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1

End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)

        '*************************************************
        'Author: Unknown
        'Last modified: 13/07/2009
        'Moves the char, sending the message to everyone in range.
        '30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
        '28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
        '13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
        '13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
        '*************************************************
        Dim npos               As WorldPos

        Dim sailing            As Boolean

        Dim CasperIndex        As Integer

        Dim CasperHeading      As eHeading

        Dim isAdminInvi        As Boolean

        Dim isZonaOscura       As Boolean

        Dim isZonaOscuraNewPos As Boolean
    
        sailing = PuedeAtravesarAgua(UserIndex)
        npos = UserList(UserIndex).pos
        isZonaOscura = (MapData(npos.Map, npos.X, npos.Y).trigger = eTrigger.zonaOscura)
    
        Call HeadtoPos(nHeading, npos)
    
        isZonaOscuraNewPos = (MapData(npos.Map, npos.X, npos.Y).trigger = eTrigger.zonaOscura)
    
        isAdminInvi = (UserList(UserIndex).flags.AdminInvisible = 1)
    
        If MoveToLegalPos(UserList(UserIndex).pos.Map, npos.X, npos.Y, sailing, Not sailing) Then

                'si no estoy solo en el mapa...
                If MapInfo(UserList(UserIndex).pos.Map).NumUsers > 1 Then
                        CasperIndex = MapData(UserList(UserIndex).pos.Map, npos.X, npos.Y).UserIndex
                        
                        'Si hay un usuario, y paso la validacion, entonces es un casper
                        If CasperIndex > 0 Then

                                ' Los admins invisibles no pueden patear caspers
                                If Not isAdminInvi Then
                    
                                        If TriggerZonaPelea(UserIndex, CasperIndex) = TRIGGER6_PROHIBE Then
                                                If UserList(CasperIndex).flags.SeguroResu = False Then
                                                        UserList(CasperIndex).flags.SeguroResu = True
                                                        Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)

                                                End If

                                        End If
    
                                        With UserList(CasperIndex)
                                                CasperHeading = InvertHeading(nHeading)
                                                Call HeadtoPos(CasperHeading, .pos)
                    
                                                ' Si es un admin invisible, no se avisa a los demas clientes
                                                If Not (.flags.AdminInvisible = 1) Then
                                                        Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .pos.X, .pos.Y))
                        
                                                        'Los valores de visible o invisible están invertidos porque estos flags son del UserIndex, por lo tanto si el UserIndex entra, el casper sale y viceversa :P
                                                        If isZonaOscura Then
                                                                If Not isZonaOscuraNewPos Then
                                                                        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, CasperIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))

                                                                End If

                                                        Else

                                                                If isZonaOscuraNewPos Then
                                                                        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, CasperIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

                                                                End If

                                                        End If

                                                End If
                        
                                                Call WriteForceCharMove(CasperIndex, CasperHeading)
                        
                                                'Update map and char
                                                .Char.heading = CasperHeading
                                                MapData(.pos.Map, .pos.X, .pos.Y).UserIndex = CasperIndex

                                        End With
                
                                        'Actualizamos las áreas de ser necesario
                                        Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)

                                End If

                        End If
            
                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not isAdminInvi Then _
                           Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, npos.X, npos.Y))
                        
                        'If isAdminInvi Then
                        'Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, npos.X, npos.Y))
                        'End If
                End If
        
                ' Los admins invisibles no pueden patear caspers
                If (Not isAdminInvi) Or (CasperIndex = 0) Then

                        With UserList(UserIndex)

                                ' Si no hay intercambio de pos con nadie
                                If CasperIndex = 0 Then
                                        MapData(.pos.Map, .pos.X, .pos.Y).UserIndex = 0

                                End If
                
                                .pos = npos
                                .Char.heading = nHeading
                                MapData(.pos.Map, .pos.X, .pos.Y).UserIndex = UserIndex
                                
                                If Extra.IsAreaResu(UserIndex) Then
                                        Call Extra.AutoCurar(UserIndex)

                                End If
                
                                If isZonaOscura Then
                                        If Not isZonaOscuraNewPos Then
                                                If (.flags.invisible Or .flags.Oculto) = 0 Then
                                                        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

                                                End If

                                        End If

                                Else

                                        If isZonaOscuraNewPos Then
                                                If (.flags.invisible Or .flags.Oculto) = 0 Then
                                                        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))

                                                End If

                                        End If

                                End If
                
                                Call DoTileEvents(UserIndex, .pos.Map, .pos.X, .pos.Y)

                        End With
            
                        'Actualizamos las áreas de ser necesario
                        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
                Else
                        Call WritePosUpdate(UserIndex)

                End If

        Else
                Call WritePosUpdate(UserIndex)

        End If
    
        If UserList(UserIndex).Counters.Trabajando Then _
           UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

        If UserList(UserIndex).Counters.Ocultando Then _
           UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1

End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading

        '*************************************************
        'Author: ZaMa
        'Last modified: 30/03/2009
        'Returns the heading opposite to the one passed by val.
        '*************************************************
        Select Case nHeading

                Case eHeading.EAST
                        InvertHeading = WEST

                Case eHeading.WEST
                        InvertHeading = EAST

                Case eHeading.SOUTH
                        InvertHeading = NORTH

                Case eHeading.NORTH
                        InvertHeading = SOUTH

        End Select

End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        UserList(UserIndex).Invent.Object(Slot) = Object
        Call WriteChangeInventorySlot(UserIndex, Slot)

End Sub

Function NextOpenCharIndex() As Integer
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Long
    
        For LoopC = 1 To MAXCHARS

                If CharList(LoopC) = 0 Then
                        NextOpenCharIndex = LoopC
                        NumChars = NumChars + 1
            
                        If LoopC > LastChar Then _
                           LastChar = LoopC
            
                        Exit Function

                End If

        Next LoopC

End Function

Function NextOpenUser() As Integer
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Long
    
        For LoopC = 1 To MaxUsers + 1

                If LoopC > MaxUsers Then Exit For
                If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
        Next LoopC
    
        NextOpenUser = LoopC

End Function

Public Sub FreeSlot(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 21/08/2015
        ' @@ By Miqueas : Fix Bug, LasUuser = 0,Nunca Salia del Loop
        '***************************************************
        UserList(UserIndex).ConnID = -1
        UserList(UserIndex).ConnIDValida = False

        If UserIndex = LastUser Then

                Do While (LastUser > 0) And (UserList(LastUser).ConnID = -1)
                        LastUser = LastUser - 1

                        If LastUser = 0 Then Exit Do ' @@ Miqueas
                Loop

        End If

End Sub

Public Function CanReceiveData(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 01/10/2012
        '
        '***************************************************
        CanReceiveData = UserList(UserIndex).ConnIDValida

End Function

Public Function isSlotFree(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 01/10/2012
        '
        '***************************************************
        isSlotFree = (UserList(UserIndex).ConnID = -1)

End Function

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 26/05/2011 (Amraphen)
        '26/05/2011: Amraphen - Ahora envía la defensa adicional de la armadura de segunda jerarquía
        '***************************************************

On Error GoTo errhandleR

        Dim GuildI             As Integer
        Dim j As Integer 'If UserList(UserIndex).clase = eClass.Mage Then asd
        Dim Name As String
        Dim Count As Integer
        Dim Miraza As String

        Dim ModificadorDefensa As Single 'Por las armaduras de segunda jerarquía.
    
1        With UserList(UserIndex)
2                Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & .Name, FontTypeNames.FONTTYPE_GUILDMSG)
3                Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_GUILDMSG)
4                Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHP & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_GUILDMSG)
5                Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(UserIndex).Stats.GLD & " Boveda: " & UserList(UserIndex).Stats.Banco & " Posicion: " & UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y & " en mapa " & UserList(UserIndex).pos.Map, FontTypeNames.FONTTYPE_GUILDMSG)
6                Call WriteConsoleMsg(sendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_GUILDMSG)

7        Select Case .raza

        Case eRaza.Drow
            Miraza = "Elfo Oscuro"
        Case eRaza.Elfo
            Miraza = "Elfo"
        Case eRaza.Enano
            Miraza = "Enano"
        Case eRaza.Gnomo
            Miraza = "Gnomo"
        Case eRaza.Humano
            Miraza = "Humano"
        End Select
        
8        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase) & " / Raza: " & Miraza & ".", FontTypeNames.FONTTYPE_GUILDMSG)
        
9        GuildI = UserList(UserIndex).GuildIndex
10    If GuildI > 0 Then
11        Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_GUILDMSG)
12        If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).Name) Then
13            Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_GUILDMSG)
14        End If
15    End If
    
    Call WriteConsoleMsg(sendIndex, SkillsNames(1) & " = " & UserList(UserIndex).Stats.UserSkills(1) & " | " _
    & SkillsNames(2) & " = " & UserList(UserIndex).Stats.UserSkills(2) & " | " _
   & SkillsNames(3) & " = " & UserList(UserIndex).Stats.UserSkills(3) & " | " _
    & SkillsNames(4) & " = " & UserList(UserIndex).Stats.UserSkills(4) & " | " _
    & SkillsNames(5) & " = " & UserList(UserIndex).Stats.UserSkills(5) & " | " _
    & SkillsNames(6) & " = " & UserList(UserIndex).Stats.UserSkills(6) & " | " _
    & SkillsNames(7) & " = " & UserList(UserIndex).Stats.UserSkills(7) & " | " _
    & SkillsNames(8) & " = " & UserList(UserIndex).Stats.UserSkills(8) & " | " _
    & SkillsNames(9) & " = " & UserList(UserIndex).Stats.UserSkills(9) & " | " _
    & SkillsNames(10) & " = " & UserList(UserIndex).Stats.UserSkills(10) & " | " _
    & SkillsNames(11) & " = " & UserList(UserIndex).Stats.UserSkills(11) & " | " _
    & SkillsNames(12) & " = " & UserList(UserIndex).Stats.UserSkills(12) & " | " _
    & SkillsNames(13) & " = " & UserList(UserIndex).Stats.UserSkills(13) & " | " _
    & SkillsNames(14) & " = " & UserList(UserIndex).Stats.UserSkills(14) & " | " _
    & SkillsNames(15) & " = " & UserList(UserIndex).Stats.UserSkills(15) & " | " _
    & SkillsNames(16) & " = " & UserList(UserIndex).Stats.UserSkills(16) & " | " _
    & SkillsNames(17) & " = " & UserList(UserIndex).Stats.UserSkills(17) & " | " _
    & SkillsNames(18) & " = " & UserList(UserIndex).Stats.UserSkills(18) & " | " _
    & SkillsNames(19) & " = " & UserList(UserIndex).Stats.UserSkills(19) & " | " _
    & SkillsNames(20) & " = " & UserList(UserIndex).Stats.UserSkills(20) & " | " _
    , FontTypeNames.FONTTYPE_GUILDMSG)
    
    
    'Items
22    Call WriteConsoleMsg(sendIndex, "Inventario: Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_GUILDMSG)
23    For j = 1 To MAX_NORMAL_INVENTORY_SLOTS
24        If .Invent.Object(j).objIndex > 0 Then
25            Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).objIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_GUILDMSG)
26         End If
27    Next j
    

    'boveda
28    Call WriteConsoleMsg(sendIndex, "Boveda: Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_GUILDMSG)
29    For j = 1 To 40
39        If UserList(UserIndex).BancoInvent.Object(j).objIndex > 0 Then
34            Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).objIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_GUILDMSG)
35        End If
36    Next
    
        End With

Exit Sub

errhandleR:

LogError "Error en sendUserStatsTxt en linea " & Erl & " - " & Err.Number & " " & Err.description

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Shows the users Stats when the user is online.
        '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
        '*************************************************
        With UserList(UserIndex)
                Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
        
                If .Faccion.ArmadaReal = 1 Then
                        Call WriteConsoleMsg(sendIndex, "Ejército real desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
                ElseIf .Faccion.FuerzasCaos = 1 Then
                        Call WriteConsoleMsg(sendIndex, "Legión oscura desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
                ElseIf .Faccion.RecibioExpInicialReal = 1 Then
                        Call WriteConsoleMsg(sendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
                ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
                        Call WriteConsoleMsg(sendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)

                End If
        
                Call WriteConsoleMsg(sendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
        
                If .GuildIndex > 0 Then
                        Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)

                End If

        End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Shows the users Stats when the user is offline.
        '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
        '*************************************************
        Dim CharFile      As String

        Dim Ban           As String

        Dim BanDetailPath As String
    
        BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
        CharFile = CharPath & charName & ".chr"
    
        If FileExist(CharFile) Then
                Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " usuarios matados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
                If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
                        Call WriteConsoleMsg(sendIndex, "Ejército real desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
                ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
                        Call WriteConsoleMsg(sendIndex, "Legión oscura desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
                ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
                        Call WriteConsoleMsg(sendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
                ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
                        Call WriteConsoleMsg(sendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)

                End If
        
                Call WriteConsoleMsg(sendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
        
                If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
                        Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)

                End If
        
                Ban = GetVar(CharFile, "FLAGS", "Ban")
                Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
                If Ban = "1" Then
                        Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)

                End If

        Else
                Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)

        End If

End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim j As Long
    
        With UserList(UserIndex)
                Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
                For j = 1 To .CurrentInventorySlots

                        If .Invent.Object(j).objIndex > 0 Then
                                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).objIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

                        End If

                Next j

        End With

End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim j        As Long

        Dim CharFile As String, Tmp As String

        Dim ObjInd   As Long, ObjCant As Long
    
        CharFile = CharPath & charName & ".chr"
    
        If FileExist(CharFile, vbNormal) Then
                Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
                For j = 1 To MAX_NORMAL_INVENTORY_SLOTS
                        Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
                        ObjInd = ReadField(1, Tmp, Asc("-"))
                        ObjCant = ReadField(2, Tmp, Asc("-"))

                        If ObjInd > 0 Then
                                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

                        End If

                Next j

        Else
                Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

        End If

End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim j As Integer
    
        Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
        For j = 1 To NUMSKILLS
                Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
        Next j
    
        Call WriteConsoleMsg(sendIndex, "SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)

End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, _
                                    ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If Npclist(NpcIndex).MaestroUser > 0 Then
                EsMascotaCiudadano = Not criminal(Npclist(NpcIndex).MaestroUser)

                If EsMascotaCiudadano Then
                        Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)

                End If

        End If

End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

        '**********************************************
        'Author: Unknown
        'Last Modification: 02/04/2010
        '24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
        '24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
        '06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
        '02/04/2010: ZaMa: Un ciuda no se vuelve mas criminal al atacar un npc no hostil.
        '**********************************************
        Dim EraCriminal As Boolean
    
        'Guardamos el usuario que ataco el npc.
        Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
    
        'Npc que estabas atacando.
        Dim LastNpcHit As Integer

        LastNpcHit = UserList(UserIndex).flags.NPCAtacado
        'Guarda el NPC que estas atacando ahora.
        UserList(UserIndex).flags.NPCAtacado = NpcIndex
    
        'Revisamos robo de npc.
        'Guarda el primer nick que lo ataca.
        If LenB(Npclist(NpcIndex).flags.AttackedFirstBy) = 0 Then

                'El que le pegabas antes ya no es tuyo
                If LastNpcHit <> 0 Then
                        If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

                        End If

                End If

                Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).Name
        ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then

                'Estas robando NPC
                'El que le pegabas antes ya no es tuyo
                If LastNpcHit <> 0 Then
                        If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

                        End If

                End If

        End If
    
        If Npclist(NpcIndex).MaestroUser > 0 Then
                If Npclist(NpcIndex).MaestroUser <> UserIndex Then
                        Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)

                End If

        End If
    
        If EsMascotaCiudadano(NpcIndex, UserIndex) Then
                Call VolverCriminal(UserIndex)
                Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
                Npclist(NpcIndex).Hostile = 1
        Else
                EraCriminal = criminal(UserIndex)
        
                'Reputacion
                If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                                Call VolverCriminal(UserIndex)

                        End If
        
                ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
                        UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2

                        If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
                           UserList(UserIndex).Reputacion.PlebeRep = MAXREP

                End If
        
                If Npclist(NpcIndex).MaestroUser <> UserIndex Then
                        'hacemos que el npc se defienda
                        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
                        Npclist(NpcIndex).Hostile = 1

                End If
        
                If EraCriminal And Not criminal(UserIndex) Then
                        Call VolverCiudadano(UserIndex)

                End If

        End If

        ' @ Flodeo del rey
        If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
            CheckFlodeoRey UserIndex, NpcIndex, guilds(UserList(UserIndex).GuildIndex).GuildName
        End If

End Sub

Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
    
        Dim WeaponIndex As Integer
     
        With UserList(UserIndex)
        
                WeaponIndex = .Invent.WeaponEqpObjIndex
        
                If WeaponIndex > 0 Then
                        If ObjData(WeaponIndex).Apuñala = 1 Then
                                PuedeApuñalar = .Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR _
                                   Or .clase = eClass.Assasin

                        End If

                End If
        
        End With
    
End Function

Public Function PuedeAcuchillar(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 25/01/2010 (ZaMa)
        '
        '***************************************************
    
        Dim WeaponIndex As Integer
    
        With UserList(UserIndex)

                If .clase = eClass.Pirat Then
        
                        WeaponIndex = .Invent.WeaponEqpObjIndex

                        If WeaponIndex > 0 Then
                                PuedeAcuchillar = (ObjData(WeaponIndex).Acuchilla = 1)

                        End If

                End If

        End With
    
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)

        '*************************************************
        'Author: Unknown
        'Last modified: 11/19/2009
        '*************************************************
        
        With UserList(UserIndex)

                If .flags.Hambre = 0 And .flags.Sed = 0 Then
                
                        With .Stats

                                If .UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
                
                                Dim Lvl As Integer

                                Lvl = .ELV
                
                                If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
                
                                If .UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
                           
                                .UserSkills(Skill) = .UserSkills(Skill) + 1
                                Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto! Ahora tienes " & .UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                    
                                .Exp = .Exp + 50

                                If .Exp > MAXEXP Then .Exp = MAXEXP
                    
                                Call WriteConsoleMsg(UserIndex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_NARANJA)
                    
                                Call WriteUpdateExp(UserIndex)
                                Call CheckUserLevel(UserIndex)

                        End With

                End If

        End With

End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Public Sub UserDie(ByVal UserIndex As Integer)

        '************************************************
        'Author: Uknown
        'Last Modified: 12/01/2010 (ZaMa)
        '04/15/2008: NicoNZ - Ahora se resetea el counter del invi
        '13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
        '27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
        '21/07/2009: Marco - Al morir se desactiva el comercio seguro.
        '16/11/2009: ZaMa - Al morir perdes la criatura que te pertenecia.
        '27/11/2009: Budi - Al morir envia los atributos originales.
        '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando mueren.
        '************************************************
        On Error GoTo ErrorHandler

        Dim i           As Long

        Dim aN          As Integer
    
        Dim iSoundDeath As Integer
    
        With UserList(UserIndex)

               
                
                'Sonido
                If .Genero = eGenero.Mujer Then
                        If HayAgua(.pos.Map, .pos.X, .pos.Y) Then
                                iSoundDeath = e_SoundIndex.MUERTE_MUJER_AGUA
                        Else
                                iSoundDeath = e_SoundIndex.MUERTE_MUJER

                        End If

                Else

                        If HayAgua(.pos.Map, .pos.X, .pos.Y) Then
                                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE_AGUA
                        Else
                                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE

                        End If

                End If
        
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, iSoundDeath)
        
                If UserIndex = GranPoder Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Favor de los dioses>", .Name & " ha perdido el poder.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO))
                    GranPoder = 0 'Call OtorgarFavordelosDioses(0)
                   
                
                    Call WarpUserChar(UserIndex, .pos.Map, .pos.X, .pos.Y, False, False)
                
                End If
                
                'Quitar el dialogo del user muerto
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
                .Stats.MinHp = 0
                .Stats.MinSta = 0
                .flags.AtacadoPorUser = 0
                .flags.Envenenado = 0
                .flags.Muerto = 1
        
                .Counters.Trabajando = 0
                .Counters.PescaBronce = 0
                .Counters.PescaOro = 0
                .Counters.PescaPlata = 0
        
                ' No se activa en arenas
                If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
                        .flags.SeguroResu = True
                        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
                Else
                        .flags.SeguroResu = False
                        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)

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
        
                Call PerdioNpc(UserIndex, False)
        
                '<<<< Atacable >>>>
                If .flags.AtacablePor > 0 Then
                        .flags.AtacablePor = 0
                        Call RefreshCharStatus(UserIndex)

                End If
        
                '<<<< Paralisis >>>>
                If .flags.Paralizado = 1 Then
                        .flags.Paralizado = 0
                        Call WriteParalizeOK(UserIndex)

                End If
        
                '<<< Estupidez >>>
                If .flags.Estupidez = 1 Then
                        .flags.Estupidez = 0
                        Call WriteDumbNoMore(UserIndex)

                End If
        
                '<<<< Descansando >>>>
                If .flags.Descansar Then
                        .flags.Descansar = False
                        Call WriteRestOK(UserIndex)
                End If
        
                '<<<< Meditando >>>>
                If .flags.Meditando Then
                        .flags.Meditando = False
                        Call WriteMeditateToggle(UserIndex)

                End If
        
                '<<<< Invisible >>>>
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                        .flags.Oculto = 0
                        .flags.invisible = 0
                        .Counters.TiempoOculto = 0
                        .Counters.Invisibilidad = 0
            
                        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)

                End If
        
                If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Then

                        ' << Si es newbie no pierde el inventario >>
                        If Not EsNewbie(UserIndex) Then
                                Call TirarTodo(UserIndex)
                        Else
                                Call TirarTodosLosItemsNoNewbies(UserIndex)

                        End If

                End If
        
                ' DESEQUIPA TODOS LOS OBJETOS
                'desequipar armadura
                If .Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

                End If
        
                'desequipar arma
                If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                End If
        
                'desequipar casco
                If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

                End If
        
                'desequipar herramienta
                If .Invent.AnilloEqpSlot > 0 Then
                        Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)

                End If
        
                'desequipar municiones
                If .Invent.MunicionEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)

                End If
        
                'desequipar escudo
                If .Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                End If
        
                ' << Reseteamos los posibles FX sobre el personaje >>
                If .Char.Loops = INFINITE_LOOPS Then
                        .Char.FX = 0
                        .Char.Loops = 0

                End If
        
                ' << Restauramos el mimetismo
                If .flags.Mimetizado = 1 Then
                        .Char.Body = .CharMimetizado.Body
                        .Char.Head = .CharMimetizado.Head
                        .Char.CascoAnim = .CharMimetizado.CascoAnim
                        .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                        .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                        .Counters.Mimetismo = 0
                        .flags.Mimetizado = 0
                        ' Puede ser atacado por npcs (cuando resucite)
                        .flags.Ignorado = False

                End If
        
                ' << Restauramos los atributos >>
                If .flags.TomoPocion = True Then

                        For i = 1 To 5
                                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
                        Next i

                End If
        
                '<< Cambiamos la apariencia del char >>
                If .flags.Navegando = 0 Then
                        .Char.Body = iCuerpoMuerto
                        .Char.Head = iCabezaMuerto
                        .Char.ShieldAnim = NingunEscudo
                        .Char.WeaponAnim = NingunArma
                        .Char.CascoAnim = NingunCasco
                Else
                        .Char.Body = iFragataFantasmal

                End If
        
                For i = 1 To MAXMASCOTAS

                        If .MascotasIndex(i) > 0 Then
                                Call MuereNpc(.MascotasIndex(i), 0)
                                ' Si estan en agua o zona segura
                        Else
                                .MascotasType(i) = 0

                        End If

                Next i
            
                
                .NroMascotas = 0
                    
                
        ' @@ Juegos del hambre
210     If .EnJDH Then Call m_JuegosDelHambre.MuereUserJDH(UserIndex)

                ' @@ Miqueas : Retos 1 vs 1
                If (.mReto.reto_Index <> 0) Then Call Mod_Retos1vs1.userdie_reto(UserIndex)
                
                ' @@ Miqueas : Retos 2 vs 2
                If (.sReto.reto_Index <> 0) Then
                        If (Mod_Retos2vs2.reto_List(.sReto.reto_Index).used_ring) Then
                                Call Mod_Retos2vs2.user_die_reto(UserIndex)

                        End If

                End If
        
                If .Death And DeathMatch.Activo Then
                   Call modDeath.MuereUser(UserIndex)
                End If
                
                If .flags.Automatico = True Then
                    Call Rondas_UsuarioMuere(UserIndex)
                End If
                
                '<< Actualizamos clientes >>
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                Call WriteUpdateUserStats(UserIndex)
                Call WriteUpdateStrenghtAndDexterity(UserIndex)

                '<<Castigos por party>>
                If .PartyIndex > 0 Then
                        Call mdParty.ObtenerExito(UserIndex, .Stats.ELV * -10 * mdParty.CantMiembros(UserIndex), .pos.Map, .pos.X, .pos.Y)

                End If
        
                '<<Cerramos comercio seguro>>
                Call LimpiarComercioSeguro(UserIndex)
        
                ' Hay que teletransportar?
                Dim Mapa As Integer

                Mapa = .pos.Map

                If Mapa > 0 Then

                        Dim MapaTelep As Integer

                        MapaTelep = MapInfo(Mapa).OnDeathGoTo.Map
        
                        If MapaTelep <> 0 Then
                                Call WriteConsoleMsg(UserIndex, "¡¡¡Tu estado no te permite permanecer en el mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
                                Call WarpUserChar(UserIndex, MapaTelep, MapInfo(Mapa).OnDeathGoTo.X, _
                                   MapInfo(Mapa).OnDeathGoTo.Y, True, True)

                        End If

                End If

        End With

        Exit Sub

ErrorHandler:
        Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)

End Sub

Public Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 13/07/2010
        '13/07/2010: ZaMa - Los matados en estado atacable ya no suman frag.
        '***************************************************

        If EsNewbie(Muerto) Then Exit Sub
        
        With UserList(Atacante)

                If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        
                If criminal(Muerto) Then
                        If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                                .flags.LastCrimMatado = UserList(Muerto).Name

                                If .Faccion.CriminalesMatados < MAXUSERMATADOS Then _
                                   .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1

                        End If
            
                        If .Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
                                .Faccion.Reenlistadas = 200  'jaja que trucho
                
                                'con esto evitamos que se vuelva a reenlistar
                        End If

                Else

                        If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                                .flags.LastCiudMatado = UserList(Muerto).Name

                                If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then _
                                   .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1

                        End If

                End If
        
                If .Stats.UsuariosMatados < MAXUSERMATADOS Then _
                   .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1

        End With

End Sub

Sub Tilelibre(ByRef pos As WorldPos, _
              ByRef npos As WorldPos, _
              ByRef Obj As Obj, _
              ByRef PuedeAgua As Boolean, _
              ByRef PuedeTierra As Boolean)

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 18/09/2010
        '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
        '18/09/2010: ZaMa - Aplico optimizacion de busqueda de tile libre en forma de rombo.
        '**************************************************************
        On Error GoTo errhandleR

        Dim Found As Boolean

        Dim LoopC As Integer

        Dim tX    As Long

        Dim tY    As Long
    
        npos = pos
        tX = pos.X
        tY = pos.Y
    
        LoopC = 1
    
        ' La primera posicion es valida?
        If LegalPos(pos.Map, npos.X, npos.Y, PuedeAgua, PuedeTierra, True) Then
        
                If Not HayObjeto(pos.Map, npos.X, npos.Y, Obj.objIndex, Obj.Amount) Then
                        Found = True

                End If
        
        End If
    
        ' Busca en las demas posiciones, en forma de "rombo"
        If Not Found Then

                While (Not Found) And LoopC <= 16

                        If RhombLegalTilePos(pos, tX, tY, LoopC, Obj.objIndex, Obj.Amount, PuedeAgua, PuedeTierra) Then
                                npos.X = tX
                                npos.Y = tY
                                Found = True

                        End If
        
                        LoopC = LoopC + 1
                Wend
        
        End If
    
        If Not Found Then
                npos.X = 0
                npos.Y = 0

        End If
    
        Exit Sub
    
errhandleR:
        Call LogError("Error en Tilelibre. Error: " & Err.Number & " - " & Err.description)

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal X As Integer, _
                 ByVal Y As Integer, _
                 ByVal FX As Boolean, _
                 Optional ByVal Teletransported As Boolean)

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 11/23/2010
        '15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
        '13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
        '16/09/2010 - ZaMa: No se pierde la visibilidad al cambiar de mapa al estar navegando invisible.
        '11/23/2010 - C4b3z0n: Ahora si no se permite Invi o Ocultar en el mapa al que cambias, te lo saca
        '**************************************************************
        Dim OldMap As Integer

        Dim OldX   As Integer

        Dim OldY   As Integer
    
        With UserList(UserIndex)
        
                'Quitar el dialogo
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
                OldMap = .pos.Map
                OldX = .pos.X
                OldY = .pos.Y

                Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
        
                
                If OldMap <> Map Then
                        
                        Call WriteChangeMap(UserIndex, Map, MapInfo(.pos.Map).MapVersion)
            
                        If MapInfo(.pos.Map).Pk = False And UserIndex = GranPoder Then
                        
                            GranPoder = 0
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Favor de los dioses>", .Name & " ha perdido el poder.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO))
                            
                        End If
                        
                        If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)

                                Dim AhoraVisible As Boolean 'Para enviar el mensaje de invi y hacer visible (C4b3z0n)

                                Dim WasInvi      As Boolean

                                'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
                                If MapInfo(Map).InviSinEfecto > 0 And .flags.invisible = 1 Then
                                        .flags.invisible = 0
                                        .Counters.Invisibilidad = 0
                                        AhoraVisible = True
                                        WasInvi = True 'si era invi, para el string

                                End If

                                'Chequeo de flags de mapa por ocultar (C4b3z0n)
                                If MapInfo(Map).OcultarSinEfecto > 0 And .flags.Oculto = 1 Then
                                        AhoraVisible = True
                                        .flags.Oculto = 0
                                        .Counters.TiempoOculto = 0

                                End If
                
                                If AhoraVisible Then 'Si no era visible y ahora es, le avisa. (C4b3z0n)
                                        Call SetInvisible(UserIndex, .Char.CharIndex, False)

                                        If WasInvi Then 'era invi
                                                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
                                        Else 'estaba oculto
                                                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", FontTypeNames.FONTTYPE_INFO)

                                        End If

                                End If

                        End If
            
                        Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(Map).Music, 45)))
            
                        'Update new Map Users
                        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
            
                        'Update old Map Users
                        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

                        If MapInfo(OldMap).NumUsers < 0 Then
                                MapInfo(OldMap).NumUsers = 0

                        End If
        
                        'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
                        Dim nextMap, previousMap As Boolean

                        'nextMap = IIf(distanceToCities(Map).distanceToCity(.Hogar) >= 0, True, False)
                        'previousMap = IIf(distanceToCities(.Pos.Map).distanceToCity(.Hogar) >= 0, True, False)

                        If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                                'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
                        ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
                                .flags.lastMap = .pos.Map
                        ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el último mapa es 0 ya que no esta en un dungeon)
                                .flags.lastMap = 0
                        ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
                                .flags.lastMap = .flags.lastMap

                        End If
            
                        Call WriteRemoveAllDialogs(UserIndex)

                End If
        
                .pos.X = X
                .pos.Y = Y
                .pos.Map = Map
                
                
                
                'Call SendData(SendTarget.ToAdmins, UserIndex, Protocol.PrepareMessageRemoveCharDialog(.Char.CharIndex))
                
                'Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
    
                Call MakeUserChar(True, Map, UserIndex, Map, X, Y)
                
                'Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, Protocol.PrepareMessageCharacterCreate(.Char.Body, .Char.Head, .Char.heading, .Char.CharIndex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.Loops, .Char.CascoAnim, .Name, GetNickColor(UserIndex), .flags.Privilegios, 0, 0))
                
                Call WriteUserCharIndexInServer(UserIndex)
        
                Call DoTileEvents(UserIndex, Map, X, Y)
        
                'Force a flush, so user index is in there before it's destroyed for teleporting
                Call FlushBuffer(UserIndex)
        
                'Seguis invisible al pasar de mapa
                If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            
                        ' No si estas navegando
                        If .flags.Navegando = 0 Then
                                Call SetInvisible(UserIndex, .Char.CharIndex, True)

                        End If

                End If
        
                If Teletransported Then
                        If .flags.Traveling = 1 Then
                                Call EndTravel(UserIndex, True)

                        End If

                End If
        
                If FX And .flags.AdminInvisible = 0 Then 'FX
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))

                End If
        
                If .NroMascotas Then Call WarpMascotas(UserIndex)
        
                ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
                Call IntervaloPermiteSerAtacado(UserIndex, True)
        
                ' Perdes el npc al cambiar de mapa
                Call PerdioNpc(UserIndex, False)
        
                ' Automatic toogle navigate
                If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
                        If HayAgua(.pos.Map, .pos.X, .pos.Y) Then
                                If .flags.Navegando = 0 Then
                                        .flags.Navegando = 1
                        
                                        'Tell the client that we are navigating.
                                        Call WriteNavigateToggle(UserIndex)

                                End If

                        Else

                                If .flags.Navegando = 1 Then
                                        .flags.Navegando = 0
                            
                                        'Tell the client that we are navigating.
                                        Call WriteNavigateToggle(UserIndex)

                                End If

                        End If

                End If
      
        End With

End Sub

Private Sub WarpMascotas(ByVal UserIndex As Integer)

        '************************************************
        'Author: Uknown
        'Last Modified: 26/10/2010
        '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
        '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
        '11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
        '26/10/2010: ZaMa - Ahora las mascotas rapswnean de forma aleatoria.
        '************************************************
        Dim i                As Integer

        Dim petType          As Integer

        Dim PetRespawn       As Boolean

        Dim PetTiempoDeVida  As Integer

        Dim NroPets          As Integer

        Dim InvocadosMatados As Integer

        Dim canWarp          As Boolean

        Dim Index            As Integer

        Dim iMinHP           As Integer
    
        NroPets = UserList(UserIndex).NroMascotas
        canWarp = (MapInfo(UserList(UserIndex).pos.Map).Pk = True) And (UserList(UserIndex).Death = 0 And UserList(UserIndex).sReto.reto_Index = 0 And UserList(UserIndex).mReto.reto_Index = 0 And UserList(UserIndex).flags.Automatico = False)
        
        For i = 1 To MAXMASCOTAS
                Index = UserList(UserIndex).MascotasIndex(i)
        
                If Index > 0 Then

                        ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
                        If Npclist(Index).Contadores.TiempoExistencia > 0 Then
                                Call QuitarNPC(Index)
                                UserList(UserIndex).MascotasIndex(i) = 0
                                InvocadosMatados = InvocadosMatados + 1
                                NroPets = NroPets - 1
                
                                petType = 0
                        Else
                                'Store data and remove NPC to recreate it after warp
                                'PetRespawn = Npclist(index).flags.Respawn = 0
                                petType = UserList(UserIndex).MascotasType(i)
                                'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                                ' Guardamos el hp, para restaurarlo uando se cree el npc
                                iMinHP = Npclist(Index).Stats.MinHp
                
                                Call QuitarNPC(Index)
                
                                ' Restauramos el valor de la variable
                                UserList(UserIndex).MascotasType(i) = petType

                        End If

                ElseIf UserList(UserIndex).MascotasType(i) > 0 Then
                        'Store data and remove NPC to recreate it after warp
                        PetRespawn = True
                        petType = UserList(UserIndex).MascotasType(i)
                        PetTiempoDeVida = 0
                Else
                        petType = 0

                End If
        
                If petType > 0 And canWarp Then
        
                        Dim SpawnPos As WorldPos
        
                        SpawnPos.Map = UserList(UserIndex).pos.Map
                        SpawnPos.X = UserList(UserIndex).pos.X + RandomNumber(-3, 3)
                        SpawnPos.Y = UserList(UserIndex).pos.Y + RandomNumber(-3, 3)
        
                        Index = SpawnNpc(petType, SpawnPos, False, PetRespawn)
            
                        'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                        ' Exception: Pets don't spawn in water if they can't swim
                        If Index = 0 Then
                                Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
                        Else
                                UserList(UserIndex).MascotasIndex(i) = Index

                                ' Nos aseguramos de que conserve el hp, si estaba dañado
                                Npclist(Index).Stats.MinHp = IIf(iMinHP = 0, Npclist(Index).Stats.MinHp, iMinHP)
            
                                Npclist(Index).MaestroUser = UserIndex
                                Npclist(Index).Contadores.TiempoExistencia = PetTiempoDeVida
                                Call FollowAmo(Index)

                        End If

                End If

        Next i
    
        If InvocadosMatados > 0 Then
                Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)

        End If
    
        If Not canWarp Then
                Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)

        End If
    
        UserList(UserIndex).NroMascotas = NroPets

End Sub

Public Sub WarpMascota(ByVal UserIndex As Integer, ByVal PetIndex As Integer)

        '************************************************
        'Author: ZaMa
        'Last Modified: 18/11/2009
        'Warps a pet without changing its stats
        '************************************************
        Dim petType   As Integer

        Dim NpcIndex  As Integer

        Dim iMinHP    As Integer

        Dim TargetPos As WorldPos
    
        With UserList(UserIndex)
        
                TargetPos.Map = .flags.TargetMap
                TargetPos.X = .flags.TargetX
                TargetPos.Y = .flags.TargetY
        
                NpcIndex = .MascotasIndex(PetIndex)
            
                'Store data and remove NPC to recreate it after warp
                petType = .MascotasType(PetIndex)
        
                ' Guardamos el hp, para restaurarlo cuando se cree el npc
                iMinHP = Npclist(NpcIndex).Stats.MinHp
        
                Call QuitarNPC(NpcIndex)
        
                ' Restauramos el valor de la variable
                .MascotasType(PetIndex) = petType
                .NroMascotas = .NroMascotas + 1
                NpcIndex = SpawnNpc(petType, TargetPos, False, False)
        
                'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                ' Exception: Pets don't spawn in water if they can't swim
                If NpcIndex = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
                Else
                        .MascotasIndex(PetIndex) = NpcIndex

                        With Npclist(NpcIndex)
                                ' Nos aseguramos de que conserve el hp, si estaba dañado
                                .Stats.MinHp = IIf(iMinHP = 0, .Stats.MinHp, iMinHP)
            
                                .MaestroUser = UserIndex
                                .Movement = TipoAI.SigueAmo
                                .Target = 0
                                .TargetNPC = 0

                        End With
            
                        Call FollowAmo(NpcIndex)

                End If

        End With

End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: 16/09/2010
        '16/09/2010 - ZaMa: Cuando se va el invi estando navegando, no se saca el invi (ya esta visible).
        '***************************************************
        Dim isNotVisible As Boolean

        Dim HiddenPirat  As Boolean
    
        With UserList(UserIndex)

                If .flags.UserLogged And Not .Counters.Saliendo Then
                        .Counters.Saliendo = True
                        .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.pos.Map).Pk, IntervaloCerrarConexion, 0)
            
                        isNotVisible = (.flags.Oculto Or .flags.invisible)

                        If isNotVisible Then
                                .flags.invisible = 0
                
                                If .flags.Oculto Then
                                        If .flags.Navegando = 1 Then
                                                If .clase = eClass.Pirat Then
                                                        ' Pierde la apariencia de fragata fantasmal
                                                        Call ToggleBoatBody(UserIndex)
                                                        Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                                                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                                           NingunEscudo, NingunCasco)
                                                        HiddenPirat = True

                                                End If

                                        End If

                                End If
                
                                .flags.Oculto = 0
                
                                ' Para no repetir mensajes
                                If Not HiddenPirat Then Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                                ' Si esta navegando ya esta visible
                                If .flags.Navegando = 0 Then
                                        Call SetInvisible(UserIndex, .Char.CharIndex, False)

                                End If

                        End If
            
                        If .flags.Traveling = 1 Then
                                Call EndTravel(UserIndex, True)

                        End If
            
                        Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)

                End If

        End With
        
        UserList(UserIndex).flags.MeMando = ""
        UserList(UserIndex).flags.Lemande = ""

End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/02/08
        '
        '***************************************************
        If UserList(UserIndex).Counters.Saliendo Then

                ' Is the user still connected?
                If UserList(UserIndex).ConnIDValida Then
                        UserList(UserIndex).Counters.Saliendo = False
                        UserList(UserIndex).Counters.Salir = 0
                        Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
                Else
                        'Simply reset
                        UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).pos.Map).Pk, IntervaloCerrarConexion, 0)

                End If

        End If

End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, _
                       ByVal UserIndexDestino As Integer, _
                       ByVal NuevoNick As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim ViejoNick       As String

        Dim ViejoCharBackup As String
    
        If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
        ViejoNick = UserList(UserIndexDestino).Name
    
        If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
                'hace un backup del char
                ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
                Name CharPath & ViejoNick & ".chr" As ViejoCharBackup

        End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
                Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Else
                Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Energía: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Maná: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        
                Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        
                Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
                #If ConUpTime Then

                        Dim TempSecs As Long

                        Dim TempStr  As String

                        TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
                        TempStr = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
                        Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
                #End If
    
                Call WriteConsoleMsg(sendIndex, "Dados: " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT1") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT2") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT3") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT4") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT5"), FontTypeNames.FONTTYPE_INFO)

        End If

End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim CharFile As String
    
        On Error Resume Next

        CharFile = CharPath & charName & ".chr"
    
        If FileExist(CharFile, vbNormal) Then
                Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
        Else
                Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

        End If

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 21/02/2010
        'Nacho: Actualiza el tag al cliente
        '21/02/2010: ZaMa - Ahora deja de ser atacable si se hace criminal.
        '**************************************************************
        With UserList(UserIndex)

                If MapData(.pos.Map, .pos.X, .pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
                If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                        .Reputacion.BurguesRep = 0
                        .Reputacion.NobleRep = 0
                        .Reputacion.PlebeRep = 0
                        .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO

                        If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
                        If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
            
                        If .flags.AtacablePor > 0 Then .flags.AtacablePor = 0

                End If

        End With
    
        Call RefreshCharStatus(UserIndex)

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 21/06/2006
        'Nacho: Actualiza el tag al cliente.
        '**************************************************************
        With UserList(UserIndex)

                If MapData(.pos.Map, .pos.X, .pos.Y).trigger = 6 Then Exit Sub
        
                .Reputacion.LadronesRep = 0
                .Reputacion.BandidoRep = 0
                .Reputacion.AsesinoRep = 0
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End With
    
        Call RefreshCharStatus(UserIndex)

End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal Body As Integer) As Boolean

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 10/07/2008
        'Checks if a given body index is a boat
        '**************************************************************
        'TODO : This should be checked somehow else. This is nasty....
        If Body = iFragataReal Or Body = iFragataCaos Or Body = iBarcaPk Or _
           Body = iGaleraPk Or Body = iGaleonPk Or Body = iBarcaCiuda Or _
           Body = iGaleraCiuda Or Body = iGaleonCiuda Or Body = iFragataFantasmal Then
                BodyIsBoat = True

        End If

End Function

Public Sub SetInvisible(ByVal UserIndex As Integer, _
                        ByVal userCharIndex As Integer, _
                        ByVal invisible As Boolean)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim sndNick As String

        With UserList(UserIndex)
                Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(userCharIndex, invisible))
    
                sndNick = .Name
    
                If invisible Then
                        sndNick = sndNick & " " & TAG_USER_INVISIBLE
                Else

                        If .GuildIndex > 0 Then
                                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"

                        End If

                End If
    
                Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, UserIndex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))

        End With

End Sub

Public Sub SetConsulatMode(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 05/06/10
        '
        '***************************************************

        Dim sndNick As String

        With UserList(UserIndex)
                sndNick = .Name
    
                If .flags.EnConsulta Then
                        sndNick = sndNick & " " & TAG_CONSULT_MODE
                Else

                        If .GuildIndex > 0 Then
                                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"

                        End If

                End If
    
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeNick(.Char.CharIndex, sndNick))

        End With

End Sub

Public Function IsArena(ByVal UserIndex As Integer) As Boolean
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 10/11/2009
        'Returns true if the user is in an Arena
        '**************************************************************
        IsArena = (TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE)

End Function

Public Sub PerdioNpc(ByVal UserIndex As Integer, _
                     Optional ByVal CheckPets As Boolean = True)
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 11/07/2010 (ZaMa)
        'The user loses his owned npc
        '18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdió.
        '11/07/2010: ZaMa - Coloco el indice correcto de las mascotas y ahora siguen al amo si existen.
        '13/07/2010: ZaMa - Ahora solo dejan de atacar las mascotas si estan atacando al npc que pierde su amo.
        '**************************************************************

        Dim PetCounter As Long

        Dim PetIndex   As Integer

        Dim NpcIndex   As Integer
    
        With UserList(UserIndex)
        
                NpcIndex = .flags.OwnedNpc

                If NpcIndex > 0 Then
            
                        If CheckPets Then

                                ' Dejan de atacar las mascotas
                                If .NroMascotas > 0 Then

                                        For PetCounter = 1 To MAXMASCOTAS
                    
                                                PetIndex = .MascotasIndex(PetCounter)
                        
                                                If PetIndex > 0 Then

                                                        ' Si esta atacando al npc deja de hacerlo
                                                        If Npclist(PetIndex).TargetNPC = NpcIndex Then
                                                                Call FollowAmo(PetIndex)

                                                        End If

                                                End If
                        
                                        Next PetCounter

                                End If

                        End If
            
                        ' Reset flags
                        Npclist(NpcIndex).Owner = 0
                        .flags.OwnedNpc = 0

                End If

        End With

End Sub

Public Sub ApropioNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 27/07/2010 (zaMa)
        'The user owns a new npc
        '18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
        '19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
        '27/07/2010: ZaMa - El sistema no aplica a mapas seguros.
        '**************************************************************

        With UserList(UserIndex)

                ' Los admins no se pueden apropiar de npcs
                If esGM(UserIndex) Then Exit Sub
        
                Dim Mapa As Integer

                Mapa = .pos.Map
        
                ' No aplica a triggers seguras
                If MapData(Mapa, .pos.X, .pos.Y).trigger = eTrigger.ZONASEGURA Then Exit Sub
        
                ' No se aplica a mapas seguros
                If MapInfo(Mapa).Pk = False Then Exit Sub
        
                ' No aplica a algunos mapas que permiten el robo de npcs
                If MapInfo(Mapa).RoboNpcsPermitido = 1 Then Exit Sub
        
                ' Pierde el npc anterior
                If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        
                ' Si tenia otro dueño, lo perdio aca
                Npclist(NpcIndex).Owner = UserIndex
                .flags.OwnedNpc = NpcIndex

        End With
    
        ' Inicializo o actualizo el timer de pertenencia
        Call IntervaloPerdioNpc(UserIndex, True)

End Sub

Public Function GetDireccion(ByVal UserIndex As Integer, _
                             ByVal OtherUserIndex As Integer) As String

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 17/11/2009
        'Devuelve la direccion hacia donde esta el usuario
        '**************************************************************
        Dim X As Integer

        Dim Y As Integer
    
        X = UserList(UserIndex).pos.X - UserList(OtherUserIndex).pos.X
        Y = UserList(UserIndex).pos.Y - UserList(OtherUserIndex).pos.Y
    
        If X = 0 And Y > 0 Then
                GetDireccion = "Sur"
        ElseIf X = 0 And Y < 0 Then
                GetDireccion = "Norte"
        ElseIf X > 0 And Y = 0 Then
                GetDireccion = "Este"
        ElseIf X < 0 And Y = 0 Then
                GetDireccion = "Oeste"
        ElseIf X > 0 And Y < 0 Then
                GetDireccion = "NorEste"
        ElseIf X < 0 And Y < 0 Then
                GetDireccion = "NorOeste"
        ElseIf X > 0 And Y > 0 Then
                GetDireccion = "SurEste"
        ElseIf X < 0 And Y > 0 Then
                GetDireccion = "SurOeste"

        End If

End Function

Public Function SameFaccion(ByVal UserIndex As Integer, _
                            ByVal OtherUserIndex As Integer) As Boolean
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 17/11/2009
        'Devuelve True si son de la misma faccion
        '**************************************************************
        SameFaccion = (esCaos(UserIndex) And esCaos(OtherUserIndex)) Or _
           (esArmada(UserIndex) And esArmada(OtherUserIndex))

End Function

Public Function FarthestPet(ByVal UserIndex As Integer) As Integer

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 18/11/2009
        'Devuelve el indice de la mascota mas lejana.
        '**************************************************************
        On Error GoTo errhandleR
    
        Dim PetIndex      As Integer

        Dim Distancia     As Integer

        Dim OtraDistancia As Integer
    
        With UserList(UserIndex)

                If .NroMascotas = 0 Then Exit Function
    
                For PetIndex = 1 To MAXMASCOTAS

                        ' Solo pos invocar criaturas que exitan!
                        If .MascotasIndex(PetIndex) > 0 Then

                                ' Solo aplica a mascota, nada de elementales..
                                If Npclist(.MascotasIndex(PetIndex)).Contadores.TiempoExistencia = 0 Then
                                        If FarthestPet = 0 Then
                                                ' Por si tiene 1 sola mascota
                                                FarthestPet = PetIndex
                                                Distancia = Abs(.pos.X - Npclist(.MascotasIndex(PetIndex)).pos.X) + _
                                                   Abs(.pos.Y - Npclist(.MascotasIndex(PetIndex)).pos.Y)
                                        Else
                                                ' La distancia de la proxima mascota
                                                OtraDistancia = Abs(.pos.X - Npclist(.MascotasIndex(PetIndex)).pos.X) + _
                                                   Abs(.pos.Y - Npclist(.MascotasIndex(PetIndex)).pos.Y)

                                                ' Esta mas lejos?
                                                If OtraDistancia > Distancia Then
                                                        Distancia = OtraDistancia
                                                        FarthestPet = PetIndex

                                                End If

                                        End If

                                End If

                        End If

                Next PetIndex

        End With

        Exit Function
    
errhandleR:
        Call LogError("Error en FarthestPet")

End Function

Public Function HasEnoughItems(ByVal UserIndex As Integer, _
                               ByVal objIndex As Integer, _
                               ByVal Amount As Long) As Boolean
        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 25/11/2009
        'Cheks Wether the user has the required amount of items in the inventory or not
        '**************************************************************

        Dim Slot          As Long

        Dim ItemInvAmount As Long
    
        With UserList(UserIndex)

                For Slot = 1 To .CurrentInventorySlots

                        ' Si es el item que busco
                        If .Invent.Object(Slot).objIndex = objIndex Then
                                ' Lo sumo a la cantidad total
                                ItemInvAmount = ItemInvAmount + .Invent.Object(Slot).Amount

                        End If

                Next Slot

        End With
    
        HasEnoughItems = Amount <= ItemInvAmount

End Function

Public Function TotalOfferItems(ByVal objIndex As Integer, _
                                ByVal UserIndex As Integer) As Long

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 25/11/2009
        'Cheks the amount of items the user has in offerSlots.
        '**************************************************************
        Dim Slot As Byte
    
        For Slot = 1 To MAX_OFFER_SLOTS

                ' Si es el item que busco
                If UserList(UserIndex).ComUsu.Objeto(Slot) = objIndex Then
                        ' Lo sumo a la cantidad total
                        TotalOfferItems = TotalOfferItems + UserList(UserIndex).ComUsu.Cant(Slot)

                End If

        Next Slot

End Function

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
   
        getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS

End Function

Public Sub goHome(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Budi
        'Last Modification: 01/06/2010
        '01/06/2010: ZaMa - Ahora usa otro tipo de intervalo (lo saque de tPiquetec)
        '***************************************************

        Dim Distance As Long

        Dim Tiempo   As Long
    
        With UserList(UserIndex)

                If .flags.Muerto = 1 Then
                        'If .flags.lastMap = 0 Then
                        '    Distance = distanceToCities(.Pos.Map).distanceToCity(.Hogar)
                        'Else
                        '    Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY
                        'End If
            
                        Tiempo = (Distance + 1) * 30 'seg
            
                        Call IntervaloGoHome(UserIndex, Tiempo * 1000, True)
                
                        If .flags.Navegando = 1 Then
                                .Char.FX = AnimHogarNavegando(.Char.heading)
                        Else
                                .Char.FX = AnimHogar(.Char.heading)

                        End If
                
                        .Char.Loops = INFINITE_LOOPS
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
                
                        'Call WriteMultiMessage(UserIndex, eMessages.Home, Distance, Tiempo, , MapInfo(Ciudades(.Hogar).Map).Name)
                Else
                        Call WriteConsoleMsg(UserIndex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)

                End If
        
        End With
    
End Sub

Public Function ToogleToAtackable(ByVal UserIndex As Integer, _
                                  ByVal OwnerIndex As Integer, _
                                  Optional ByVal StealingNpc As Boolean = True) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 15/01/2010
        'Change to Atackable mode.
        '***************************************************
    
        Dim AtacablePor As Integer
    
        With UserList(UserIndex)
        
                If MapInfo(.pos.Map).Pk = False Then
                        Call WriteConsoleMsg(UserIndex, "No puedes robar npcs en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                End If
        
                AtacablePor = .flags.AtacablePor
            
                If AtacablePor > 0 Then

                        ' Intenta robar un npc
                        If StealingNpc Then

                                ' Puede atacar el mismo npc que ya estaba robando, pero no una nuevo.
                                If AtacablePor <> OwnerIndex Then
                                        Call WriteConsoleMsg(UserIndex, "No puedes atacar otra criatura con dueño hasta que haya terminado tu castigo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Function

                                End If

                                ' Esta atacando a alguien en estado atacable => Se renueva el timer de atacable
                        Else
                                ' Renovar el timer
                                Call IntervaloEstadoAtacable(UserIndex, True)
                                ToogleToAtackable = True
                                Exit Function

                        End If

                End If
        
                .flags.AtacablePor = OwnerIndex
    
                ' Actualizar clientes
                Call RefreshCharStatus(UserIndex)
        
                ' Inicializar el timer
                Call IntervaloEstadoAtacable(UserIndex, True)
        
                ToogleToAtackable = True
        
        End With
    
End Function

Public Sub setHome(ByVal UserIndex As Integer, _
                   ByVal newHome As eCiudad, _
                   ByVal NpcIndex As Integer)

        '***************************************************
        'Author: Budi
        'Last Modification: 01/06/2010
        '30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
        '01/06/2010: ZaMa - Ahora te avisa si ya tenes ese hogar.
        '***************************************************
        If newHome < eCiudad.cUllathorpe Or newHome > eCiudad.cLastCity - 1 Then Exit Sub
    
        If UserList(UserIndex).Hogar <> newHome Then
                UserList(UserIndex).Hogar = newHome
    
                Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a nuestra humilde comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
        Else
                Call WriteChatOverHead(UserIndex, "¡¡¡Ya eres miembro de nuestra humilde comunidad!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)

        End If

End Sub

Public Function GetHomeArrivalTime(ByVal UserIndex As Integer) As Integer

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        'Calculates the time left to arrive home.
        '**************************************************************
        Dim TActual As Long
    
        TActual = GetTickCount() And &H7FFFFFFF
    
        With UserList(UserIndex)
                GetHomeArrivalTime = (.Counters.goHome - TActual) * 0.001

        End With

End Function

Public Sub HomeArrival(ByVal UserIndex As Integer)
        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        'Teleports user to its home.
        '**************************************************************
    
        Dim tX   As Integer

        Dim tY   As Integer

        Dim tMap As Integer

        With UserList(UserIndex)

                'Antes de que el pj llegue a la ciudad, lo hacemos dejar de navegar para que no se buguee.
                If .flags.Navegando = 1 Then
                        .Char.Body = iCuerpoMuerto
                        .Char.Head = iCabezaMuerto
                        .Char.ShieldAnim = NingunEscudo
                        .Char.WeaponAnim = NingunArma
                        .Char.CascoAnim = NingunCasco
            
                        .flags.Navegando = 0
            
                        Call WriteNavigateToggle(UserIndex)

                        'Le sacamos el navegando, pero no le mostramos a los demás porque va a ser sumoneado hasta ulla.
                End If
        
                'tX = Ciudades(.Hogar).X
                'tY = Ciudades(.Hogar).Y
                'tMap = Ciudades(.Hogar).Map
        
                Call FindLegalPos(UserIndex, tMap, tX, tY)
                Call WarpUserChar(UserIndex, tMap, tX, tY, True)
        
                Call WriteMultiMessage(UserIndex, eMessages.FinishHome)
        
                Call EndTravel(UserIndex, False)
        
        End With
    
End Sub

Public Sub EndTravel(ByVal UserIndex As Integer, ByVal Cancelado As Boolean)

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 11/06/2011
        'Ends travel.
        '**************************************************************
        With UserList(UserIndex)
                .Counters.goHome = 0
                .flags.Traveling = 0

                If Cancelado Then Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
                .Char.FX = 0
                .Char.Loops = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))

        End With

End Sub



Sub SendUserStatsTxtMercadoAO(ByVal sendIndex As Integer, ByVal UserIndex As String)
Dim j As Integer 'If UserList(UserIndex).clase = eClass.Mage Then asd
Dim GuildI As Integer
Dim Name As String
Dim Count As Integer
Dim Miraza As String

    Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(UserIndex).Name & ".", FontTypeNames.FONTTYPE_GUILDMSG)
    Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU, FontTypeNames.FONTTYPE_GUILDMSG)
    Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(UserIndex).Stats.MinHp & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta, FontTypeNames.FONTTYPE_GUILDMSG)
    
    Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_GUILDMSG)
    Call WriteConsoleMsg(sendIndex, "Usuarios matados: " & UserList(UserIndex).Stats.UsuariosMatados & " / Npcs matados: " & UserList(UserIndex).Stats.NPCsMuertos & ".", FontTypeNames.FONTTYPE_GUILDMSG)
    
    Select Case UserList(UserIndex).raza
        Case eRaza.Drow
            Miraza = "Elfo Oscuro"
        Case eRaza.Elfo
            Miraza = "Elfo"
        Case eRaza.Enano
            Miraza = "Enano"
        Case eRaza.Gnomo
            Miraza = "Gnomo"
        Case eRaza.Humano
            Miraza = "Humano"
    End Select
    Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(UserList(UserIndex).clase) & " / Raza: " & Miraza & ".", FontTypeNames.FONTTYPE_GUILDMSG)
        
    GuildI = UserList(UserIndex).GuildIndex
    If GuildI > 0 Then
        Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
        If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).Name) Then
            Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End If

    Call WriteConsoleMsg(sendIndex, SkillsNames(1) & " = " & UserList(UserIndex).Stats.UserSkills(1) & " | " _
    & SkillsNames(2) & " = " & UserList(UserIndex).Stats.UserSkills(2) & " | " _
    & SkillsNames(3) & " = " & UserList(UserIndex).Stats.UserSkills(3) & " | " _
    & SkillsNames(4) & " = " & UserList(UserIndex).Stats.UserSkills(4) & " | " _
    & SkillsNames(5) & " = " & UserList(UserIndex).Stats.UserSkills(5) & " | " _
    & SkillsNames(6) & " = " & UserList(UserIndex).Stats.UserSkills(6) & " | " _
    & SkillsNames(7) & " = " & UserList(UserIndex).Stats.UserSkills(7) & " | " _
    & SkillsNames(8) & " = " & UserList(UserIndex).Stats.UserSkills(8) & " | " _
    & SkillsNames(9) & " = " & UserList(UserIndex).Stats.UserSkills(9) & " | " _
    & SkillsNames(10) & " = " & UserList(UserIndex).Stats.UserSkills(10) & " | " _
    & SkillsNames(11) & " = " & UserList(UserIndex).Stats.UserSkills(11) & " | " _
    & SkillsNames(12) & " = " & UserList(UserIndex).Stats.UserSkills(12) & " | " _
    & SkillsNames(13) & " = " & UserList(UserIndex).Stats.UserSkills(13) & " | " _
    & SkillsNames(14) & " = " & UserList(UserIndex).Stats.UserSkills(14) & " | " _
    & SkillsNames(15) & " = " & UserList(UserIndex).Stats.UserSkills(15) & " | " _
    & SkillsNames(16) & " = " & UserList(UserIndex).Stats.UserSkills(16) & " | " _
    & SkillsNames(17) & " = " & UserList(UserIndex).Stats.UserSkills(17) & " | " _
    & SkillsNames(18) & " = " & UserList(UserIndex).Stats.UserSkills(18) & " | " _
    & SkillsNames(19) & " = " & UserList(UserIndex).Stats.UserSkills(19) & " | " _
    & SkillsNames(20) & " = " & UserList(UserIndex).Stats.UserSkills(20) & " | " _
    , FontTypeNames.FONTTYPE_GUILDMSG)
    
    'Items
    Call WriteConsoleMsg(sendIndex, "Inventario: Tiene " & UserList(UserIndex).Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_GUILDMSG)
    For j = 1 To 25
        If UserList(UserIndex).Invent.Object(j).objIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).objIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount, FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    Next j
    
    'boveda
    Call WriteConsoleMsg(sendIndex, "Boveda: Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_GUILDMSG)
    For j = 1 To 50
        If UserList(UserIndex).BancoInvent.Object(j).objIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).objIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    Next
    
 
End Sub


