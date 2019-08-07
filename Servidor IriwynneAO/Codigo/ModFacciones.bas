Attribute VB_Name = "ModFacciones"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public ArmaduraImperial1              As Integer

Public ArmaduraImperial2              As Integer

Public ArmaduraImperial3              As Integer

Public TunicaMagoImperial             As Integer

Public TunicaMagoImperialEnanos       As Integer

Public ArmaduraCaos1                  As Integer

Public ArmaduraCaos2                  As Integer

Public ArmaduraCaos3                  As Integer

Public TunicaMagoCaos                 As Integer

Public TunicaMagoCaosEnanos           As Integer

Public VestimentaImperialHumano       As Integer

Public VestimentaImperialEnano        As Integer

Public TunicaConspicuaHumano          As Integer

Public TunicaConspicuaEnano           As Integer

Public ArmaduraNobilisimaHumano       As Integer

Public ArmaduraNobilisimaEnano        As Integer

Public ArmaduraGranSacerdote          As Integer

Public VestimentaLegionHumano         As Integer

Public VestimentaLegionEnano          As Integer

Public TunicaLobregaHumano            As Integer

Public TunicaLobregaEnano             As Integer

Public TunicaEgregiaHumano            As Integer

Public TunicaEgregiaEnano             As Integer

Public SacerdoteDemoniaco             As Integer

Public Const NUM_RANGOS_FACCION       As Integer = 15

Private Const NUM_DEF_FACCION_ARMOURS As Byte = 3

Public Enum eTipoDefArmors

        ieBaja
        ieMedia
        ieAlta

End Enum

Public Type tFaccionArmaduras

        Armada(NUM_DEF_FACCION_ARMOURS - 1) As Integer
        Caos(NUM_DEF_FACCION_ARMOURS - 1) As Integer

End Type

' Matriz que contiene las armaduras faccionarias segun raza, clase, faccion y defensa de armadura
Public ArmadurasFaccion(1 To NUMCLASES, 1 To NUMRAZAS) As tFaccionArmaduras

' Contiene la cantidad de exp otorgada cada vez que aumenta el rango
Public RecompensaFacciones(NUM_RANGOS_FACCION)         As Long

Private Function GetArmourAmount(ByVal Rango As Integer, _
                                 ByVal TipoDef As eTipoDefArmors) As Integer
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 15/04/2010
        'Returns the amount of armours to give, depending on the specified rank
        '***************************************************

        Select Case TipoDef
        
                Case eTipoDefArmors.ieBaja
                        GetArmourAmount = 20 / (Rango + 1)
            
                Case eTipoDefArmors.ieMedia
                        GetArmourAmount = Rango * 2 / MaximoInt((Rango - 4), 1)
            
                Case eTipoDefArmors.ieAlta
                        GetArmourAmount = Rango * 1.35
            
        End Select
    
End Function

Private Sub GiveFactionArmours(ByVal UserIndex As Integer, ByVal IsCaos As Boolean)
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 15/04/2010
        'Gives faction armours to user
        '***************************************************
    
        Dim ObjArmour As Obj

        Dim Rango     As Integer
    
        With UserList(UserIndex)
    
                Rango = val(IIf(IsCaos, .Faccion.RecompensasCaos, .Faccion.RecompensasReal)) + 1
    
                ' Entrego armaduras de defensa baja
                ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieBaja)
        
                If IsCaos Then
                        ObjArmour.objIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieBaja)
                Else
                        ObjArmour.objIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieBaja)

                End If
        
                If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
                        Call TirarItemAlPiso(.Pos, ObjArmour)

                End If
        
                ' Entrego armaduras de defensa media
                ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieMedia)
        
                If IsCaos Then
                        ObjArmour.objIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieMedia)
                Else
                        ObjArmour.objIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieMedia)

                End If
        
                If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
                        Call TirarItemAlPiso(.Pos, ObjArmour)

                End If
    
                ' Entrego armaduras de defensa alta
                ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieAlta)
        
                If IsCaos Then
                        ObjArmour.objIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta)
                Else
                        ObjArmour.objIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta)

                End If
        
                If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
                        Call TirarItemAlPiso(.Pos, ObjArmour)

                End If

        End With

End Sub

Public Sub GiveExpReward(ByVal UserIndex As Integer, ByVal Rango As Long)
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 15/04/2010
        'Gives reward exp to user
        '***************************************************
    
        Dim GivenExp As Long
    
        With UserList(UserIndex)
        
                GivenExp = RecompensaFacciones(Rango)
        
                .Stats.Exp = .Stats.Exp + GivenExp
        
                If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
                Call WriteConsoleMsg(UserIndex, "Has sido recompensado con " & GivenExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_NARANJA)

                Call CheckUserLevel(UserIndex)
        
        End With
    
End Sub

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 15/04/2010
        'Handles the entrance of users to the "Armada Real"
        '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineaci�n neutral.
        '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
        '15/04/2010: ZaMa - Cambio en recompensas iniciales.
        '***************************************************

        With UserList(UserIndex)

                If .Faccion.ArmadaReal = 1 Then
                        Call WriteChatOverHead(UserIndex, "���Ya perteneces a las tropas reales!!! Ve a combatir criminales.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Faccion.FuerzasCaos = 1 Then
                        Call WriteChatOverHead(UserIndex, "���Maldito insolente!!! Vete de aqu� seguidor de las sombras.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If criminal(UserIndex) Then
                        Call WriteChatOverHead(UserIndex, "���No se permiten criminales en el ej�rcito real!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Faccion.CriminalesMatados < 30 Then
                        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, s�lo has matado " & .Faccion.CriminalesMatados & ".", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Stats.ELV < 32 Then
                        Call WriteChatOverHead(UserIndex, "���Para unirte a nuestras fuerzas debes ser al menos de nivel 32!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
     
                If .Faccion.CiudadanosMatados > 0 Then
                        Call WriteChatOverHead(UserIndex, "�Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Faccion.Reenlistadas > 4 Then
                        Call WriteChatOverHead(UserIndex, "�Has sido expulsado de las fuerzas reales demasiadas veces!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Reputacion.NobleRep < 50000 Then
                        Call WriteChatOverHead(UserIndex, "Necesitas ser a�n m�s noble para integrar el ej�rcito real, s�lo tienes " & .Reputacion.NobleRep & "/50.000 puntos de nobleza", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .GuildIndex > 0 Then
                        If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                                Call WriteChatOverHead(UserIndex, "���Perteneces a un clan neutro, sal de �l si quieres unirte a nuestras fuerzas!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                Exit Sub

                        End If

                End If
    
                .Faccion.ArmadaReal = 1
                .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
    
                
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Facci�n Real> " & UserList(UserIndex).Name & " se ha enlistado en nuestras fuerzas!", FontTypeNames.FONTTYPE_SERVER))
                Call WriteChatOverHead(UserIndex, "���Bienvenido al ej�rcito real!!! Aqu� tienes tus vestimentas. Cumple bien tu labor exterminando criminales y me encargar� de recompensarte.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                
                ' TODO: Dejo esta variable por ahora, pero con chequear las reenlistadas deberia ser suficiente :S
                If .Faccion.RecibioArmaduraReal = 0 Then
        
                        Call GiveFactionArmours(UserIndex, False)
                        Call GiveExpReward(UserIndex, 0)
        
                        .Faccion.RecibioArmaduraReal = 1
                        .Faccion.NivelIngreso = .Stats.ELV
                        .Faccion.FechaIngreso = Date
                        'Esto por ahora es in�til, siempre va a ser cero, pero bueno, despues va a servir.
                        .Faccion.MatadosIngreso = .Faccion.CiudadanosMatados
        
                        .Faccion.RecibioExpInicialReal = 1
                        .Faccion.RecompensasReal = 0
                        .Faccion.NextRecompensa = 30
        
                End If
    
                If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
    
                Call LogEjercitoReal(.Name & " ingres� el " & Date & " cuando era nivel " & .Stats.ELV)

        End With

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 15/04/2010
        'Handles the way of gaining new ranks in the "Armada Real"
        '15/04/2010: ZaMa - Agrego recompensas de oro y armaduras
        '***************************************************
        Dim Crimis    As Long

        Dim Lvl       As Byte

        Dim NextRecom As Long

        Dim Nobleza   As Long

        With UserList(UserIndex)
                Lvl = .Stats.ELV
                Crimis = .Faccion.CriminalesMatados
                NextRecom = .Faccion.NextRecompensa
                Nobleza = .Reputacion.NobleRep
    
                If Crimis < NextRecom Then
                        Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Crimis & " criminales m�s para recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                Select Case NextRecom

                        Case 30:
                                .Faccion.RecompensasReal = 1
                                .Faccion.NextRecompensa = 100
        
                        Case 100:
                                .Faccion.RecompensasReal = 2
                                .Faccion.NextRecompensa = 130
        
                        Case 130:
                                .Faccion.RecompensasReal = 3
                                .Faccion.NextRecompensa = 150
        
                        Case 150:
                                .Faccion.RecompensasReal = 4
                                .Faccion.NextRecompensa = 200
        
                        Case 200:
                                .Faccion.RecompensasReal = 5
                                .Faccion.NextRecompensa = 230
        
                        Case 230:

                                If Lvl < 27 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 27 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasReal = 6
                                .Faccion.NextRecompensa = 250
        
                        Case 250:
                                .Faccion.RecompensasReal = 7
                                .Faccion.NextRecompensa = 300
        
                        Case 300:
                                .Faccion.RecompensasReal = 8
                                .Faccion.NextRecompensa = 350
        
                        Case 350:

                                If Lvl < 30 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 30 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasReal = 9
                                .Faccion.NextRecompensa = 400
        
                        Case 400:

                                If Nobleza < 100000 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 2000000 - Nobleza & " puntos de nobleza para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasReal = 10
                                .Faccion.NextRecompensa = 500
        
                        Case 500:

                                If Nobleza < 300000 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 3000000 - Nobleza & " puntos de nobleza para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasReal = 11
                                .Faccion.NextRecompensa = 550
        
                        Case 550:

                                If Lvl < 35 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 35 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                If Nobleza < 400000 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 4000000 - Nobleza & " puntos de nobleza para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasReal = 12
                                .Faccion.NextRecompensa = 600
        
                        Case 600:

                                If Lvl < 36 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 36 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                If Nobleza < 500000 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 5000000 - Nobleza & " puntos de nobleza para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasReal = 13
                                .Faccion.NextRecompensa = 650
        
                        Case 650:

                                If Lvl < 37 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 37 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                If Nobleza < 600000 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 6000000 - Nobleza & " puntos de nobleza para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasReal = 14
                                .Faccion.NextRecompensa = 700
        
                        Case 700:
                                Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores soldados. Mataste " & Crimis & " criminales, sigue as�. Ya no tengo m�s recompensa para darte que mi agradecimiento. �Felicidades!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                Exit Sub
        
                        Case Else:
                                Exit Sub

                End Select
    
                Call WriteChatOverHead(UserIndex, "���Aqu� tienes tu recompensa " & TituloReal(UserIndex) & "!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

                ' Recompensas de armaduras y exp
                Call GiveFactionArmours(UserIndex, False)
                Call GiveExpReward(UserIndex, .Faccion.RecompensasReal)

        End With

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, _
                               Optional Expulsado As Boolean = True)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        ' 09/28/2010 C4b3z0n - Arreglado RT6 Overflow, el Desequipar() del escudo, pon�a de parametro el ObjIndex del escudo en vez del EqpSlot.
        '***************************************************

        With UserList(UserIndex)
                .Faccion.ArmadaReal = 0

                'Call PerderItemsFaccionarios(UserIndex)
                If Expulsado Then
                        Call WriteConsoleMsg(UserIndex, "���Has sido expulsado del ej�rcito real!!!", FontTypeNames.FONTTYPE_FIGHT)
                Else
                        Call WriteConsoleMsg(UserIndex, "���Te has retirado del ej�rcito real!!!", FontTypeNames.FONTTYPE_FIGHT)

                End If
    
                If .Invent.ArmourEqpObjIndex <> 0 Then

                        'Desequipamos la armadura real si est� equipada
                        If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

                End If
    
                If .Invent.EscudoEqpObjIndex <> 0 Then

                        'Desequipamos el escudo de caos si est� equipado
                        If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                End If
    
                If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

        End With

End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, _
                               Optional Expulsado As Boolean = True)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        ' 09/28/2010 C4b3z0n - Arreglado RT6 Overflow, el Desequipar() del escudo, pon�a de parametro el ObjIndex del escudo en vez del EqpSlot.
        '***************************************************

        With UserList(UserIndex)
                .Faccion.FuerzasCaos = 0

                'Call PerderItemsFaccionarios(UserIndex)
                If Expulsado Then
                        Call WriteConsoleMsg(UserIndex, "���Has sido expulsado de la Legi�n Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
                Else
                        Call WriteConsoleMsg(UserIndex, "���Te has retirado de la Legi�n Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)

                End If
    
                If .Invent.ArmourEqpObjIndex <> 0 Then

                        'Desequipamos la armadura de caos si est� equipada
                        If ObjData(.Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

                End If
    
                If .Invent.EscudoEqpObjIndex <> 0 Then

                        'Desequipamos el escudo de caos si est� equipado
                        If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                End If
    
                If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

        End With

End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
        '***************************************************
        'Autor: Unknown
        'Last Modification: 23/01/2007 Pablo (ToxicWaste)
        'Handles the titles of the members of the "Armada Real"
        '***************************************************

        Select Case UserList(UserIndex).Faccion.RecompensasReal
                        'Rango 1: Aprendiz (30 Criminales)
                        'Rango 2: Escudero (70 Criminales)
                        'Rango 3: Soldado (130 Criminales)
                        'Rango 4: Sargento (210 Criminales)
                        'Rango 5: Caballero (320 Criminales)
                        'Rango 6: Comandante (460 Criminales)
                        'Rango 7: Capit�n (640 Criminales + > lvl 27)
                        'Rango 8: Senescal (870 Criminales)
                        'Rango 9: Mariscal (1160 Criminales)
                        'Rango 10: Condestable (2000 Criminales + > lvl 30)
                        'Rangos de Honor de la Armada Real: (Consejo de Bander)
                        'Rango 11: Ejecutor Imperial (2500 Criminales + 2.000.000 Nobleza)
                        'Rango 12: Protector del Reino (3000 Criminales + 3.000.000 Nobleza)
                        'Rango 13: Avatar de la Justicia (3500 Criminales + 4.000.000 Nobleza + > lvl 35)
                        'Rango 14: Guardi�n del Bien (4000 Criminales + 5.000.000 Nobleza + > lvl 36)
                        'Rango 15: Campe�n de la Luz (5000 Criminales + 6.000.000 Nobleza + > lvl 37)
    
                Case 0
                        TituloReal = "Aprendiz"

                Case 1
                        TituloReal = "Escudero"

                Case 2
                        TituloReal = "Soldado"

                Case 3
                        TituloReal = "Sargento"

                Case 4
                        TituloReal = "Teniente"

                Case 5
                        TituloReal = "Comandante"

                Case 6
                        TituloReal = "Capit�n"

                Case 7
                        TituloReal = "Senescal"

                Case 8
                        TituloReal = "Mariscal"

                Case 9
                        TituloReal = "Condestable"

                Case 10
                        TituloReal = "Ejecutor Imperial"

                Case 11
                        TituloReal = "Protector del Reino"

                Case 12
                        TituloReal = "Avatar de la Justicia"

                Case 13
                        TituloReal = "Guardi�n del Bien"

                Case Else
                        TituloReal = "Campe�n de la Luz"

        End Select

End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 27/11/2009
        '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineaci�n neutral.
        '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
        'Handles the entrance of users to the "Legi�n Oscura"
        '***************************************************

        With UserList(UserIndex)

                If Not criminal(UserIndex) Then
                        Call WriteChatOverHead(UserIndex, "���L�rgate de aqu�, buf�n!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Faccion.FuerzasCaos = 1 Then
                        Call WriteChatOverHead(UserIndex, "���Ya perteneces a la legi�n oscura!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Faccion.ArmadaReal = 1 Then
                        Call WriteChatOverHead(UserIndex, "Las sombras reinar�n en Argentum. ���Fuera de aqu� insecto real!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                '[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
                If .Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ah�: �Recibio la experiencia para entrar?
                        Call WriteChatOverHead(UserIndex, "No permitir� que ning�n insecto real ingrese a mis tropas.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If

                '[/Barrin]
    
                If Not criminal(UserIndex) Then
                        Call WriteChatOverHead(UserIndex, "��Ja ja ja!! T� no eres bienvenido aqu� asqueroso ciudadano.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Faccion.CiudadanosMatados < 30 Then
                        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 30 ciudadanos, s�lo has matado " & .Faccion.CiudadanosMatados & ".", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .Stats.ELV < 32 Then
                        Call WriteChatOverHead(UserIndex, "���Para unirte a nuestras fuerzas debes ser al menos nivel 32!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                If .GuildIndex > 0 Then
                        If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                                Call WriteChatOverHead(UserIndex, "���Perteneces a un clan neutro, sal de �l si quieres unirte a nuestras fuerzas!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                Exit Sub

                        End If

                End If
    
                If .Faccion.Reenlistadas > 4 Then
                        If .Faccion.Reenlistadas = 200 Then
                                Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeld�a has atacado a mi ej�rcito. �Vete de aqu�!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Else
                                Call WriteChatOverHead(UserIndex, "�Has sido expulsado de las fuerzas oscuras demasiadas veces!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

                        End If

                        Exit Sub

                End If
    
                .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
                .Faccion.FuerzasCaos = 1
                
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Facci�n Caos> " & UserList(UserIndex).Name & " se ha enlistado en nuestras fuerzas!", FontTypeNames.FONTTYPE_SERVER))
                Call WriteChatOverHead(UserIndex, "���Bienvenido al lado oscuro!!! Aqu� tienes tus armaduras. Derrama sangre ciudadana y real, y ser�s recompensado, lo prometo.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
                If .Faccion.RecibioArmaduraCaos = 0 Then
                
                        Call GiveFactionArmours(UserIndex, True)
                        Call GiveExpReward(UserIndex, 0)
        
                        .Faccion.RecibioArmaduraCaos = 1
                        .Faccion.NivelIngreso = .Stats.ELV
                        .Faccion.FechaIngreso = Date
    
                        .Faccion.RecibioExpInicialCaos = 1
                        .Faccion.RecompensasCaos = 0
                        .Faccion.NextRecompensa = 30

                End If
    
                If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

                Call LogEjercitoCaos(.Name & " ingres� el " & Date & " cuando era nivel " & .Stats.ELV)

        End With

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 15/04/2010
        'Handles the way of gaining new ranks in the "Legi�n Oscura"
        '15/04/2010: ZaMa - Agrego recompensas de oro y armaduras
        '***************************************************
        Dim Ciudas    As Long

        Dim Lvl       As Byte

        Dim NextRecom As Long

        With UserList(UserIndex)
                Lvl = .Stats.ELV
                Ciudas = .Faccion.CiudadanosMatados
                NextRecom = .Faccion.NextRecompensa
    
                If Ciudas < NextRecom Then
                        Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Ciudas & " cuidadanos m�s para recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                        Exit Sub

                End If
    
                Select Case NextRecom

                        Case 30:
                                .Faccion.RecompensasCaos = 1
                                .Faccion.NextRecompensa = 130
        
                        Case 130:
                                .Faccion.RecompensasCaos = 2
                                .Faccion.NextRecompensa = 150
        
                        Case 150:
                                .Faccion.RecompensasCaos = 3
                                .Faccion.NextRecompensa = 200
        
                        Case 200:
                                .Faccion.RecompensasCaos = 4
                                .Faccion.NextRecompensa = 230
        
                        Case 230:
                                .Faccion.RecompensasCaos = 5
                                .Faccion.NextRecompensa = 250
        
                        Case 250:

                                If Lvl < 27 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 27 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasCaos = 6
                                .Faccion.NextRecompensa = 300
        
                        Case 300:
                                .Faccion.RecompensasCaos = 7
                                .Faccion.NextRecompensa = 350
        
                        Case 350:
                                .Faccion.RecompensasCaos = 8
                                .Faccion.NextRecompensa = 400
        
                        Case 400:

                                If Lvl < 30 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 30 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasCaos = 9
                                .Faccion.NextRecompensa = 500
        
                        Case 500:

                                If Lvl < 31 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 31 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasCaos = 10
                                .Faccion.NextRecompensa = 550
        
                        Case 550:

                                If Lvl < 33 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 33 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasCaos = 11
                                .Faccion.NextRecompensa = 600
        
                        Case 600:

                                If Lvl < 35 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 35 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasCaos = 12
                                .Faccion.NextRecompensa = 650
        
                        Case 650:

                                If Lvl < 36 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 36 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasCaos = 13
                                .Faccion.NextRecompensa = 700
        
                        Case 700:

                                If Lvl < 37 Then
                                        Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 37 - Lvl & " niveles para poder recibir la pr�xima recompensa.", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                        Exit Sub

                                End If

                                .Faccion.RecompensasCaos = 14
                                .Faccion.NextRecompensa = 850
        
                        Case 850:
                                Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores soldados. Mataste " & Ciudas & " ciudadanos . Tu �nica recompensa ser� la sangre derramada. ��Contin�a as�!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                Exit Sub
        
                        Case Else:
                                Exit Sub
            
                End Select
    
                Call WriteChatOverHead(UserIndex, "���Bien hecho " & TituloCaos(UserIndex) & ", aqu� tienes tu recompensa!!!", str$(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
                ' Recompensas de armaduras y exp
                Call GiveFactionArmours(UserIndex, True)
                Call GiveExpReward(UserIndex, .Faccion.RecompensasCaos)
    
        End With

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 23/01/2007 Pablo (ToxicWaste)
        'Handles the titles of the members of the "Legi�n Oscura"
        '***************************************************
        'Rango 1: Ac�lito (70)
        'Rango 2: Alma Corrupta (160)
        'Rango 3: Paria (300)
        'Rango 4: Condenado (490)
        'Rango 5: Esbirro (740)
        'Rango 6: Sanguinario (1100)
        'Rango 7: Corruptor (1500 + lvl 27)
        'Rango 8: Heraldo Impio (2010)
        'Rango 9: Caballero de la Oscuridad (2700)
        'Rango 10: Se�or del Miedo (4600 + lvl 30)
        'Rango 11: Ejecutor Infernal (5800 + lvl 31)
        'Rango 12: Protector del Averno (6990 + lvl 33)
        'Rango 13: Avatar de la Destrucci�n (8100 + lvl 35)
        'Rango 14: Guardi�n del Mal (9300 + lvl 36)
        'Rango 15: Campe�n de la Oscuridad (11500 + lvl 37)

        Select Case UserList(UserIndex).Faccion.RecompensasCaos

                Case 0
                        TituloCaos = "Ac�lito"

                Case 1
                        TituloCaos = "Alma Corrupta"

                Case 2
                        TituloCaos = "Paria"

                Case 3
                        TituloCaos = "Condenado"

                Case 4
                        TituloCaos = "Esbirro"

                Case 5
                        TituloCaos = "Sanguinario"

                Case 6
                        TituloCaos = "Corruptor"

                Case 7
                        TituloCaos = "Heraldo Imp�o"

                Case 8
                        TituloCaos = "Caballero de la Oscuridad"

                Case 9
                        TituloCaos = "Se�or del Miedo"

                Case 10
                        TituloCaos = "Ejecutor Infernal"

                Case 11
                        TituloCaos = "Protector del Averno"

                Case 12
                        TituloCaos = "Avatar de la Destrucci�n"

                Case 13
                        TituloCaos = "Guardi�n del Mal"

                Case Else
                        TituloCaos = "Campe�n de la Oscuridad"

        End Select

End Function

