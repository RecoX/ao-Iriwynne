Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
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
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget

        ToAll = 1
        toMap
        ToPCArea
        ToAllButIndex
        ToMapButIndex
        ToGM
        ToNPCArea
        ToGuildMembers
        ToAdmins
        ToPCAreaButIndex
        ToAdminsAreaButConsejeros
        ToDiosesYclan
        ToConsejo
        ToClanArea
        ToConsejoCaos
        ToRolesMasters
        ToDeadArea
        ToCiudadanos
        ToCriminales
        ToPartyArea
        ToReal
        ToCaos
        ToCiudadanosYRMs
        ToCriminalesYRMs
        ToRealYRMs
        ToCaosYRMs
        ToHigherAdmins
        ToGMsAreaButRmsOrCounselors
        ToUsersAreaButGMs
        ToUsersAndRmsAndCounselorsAreaButGMs

End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, _
                    ByVal sndIndex As Integer, _
                    ByVal sndData As String, _
                    Optional ByVal IsDenounce As Boolean = False)

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
        'Last Modify Date: 14/11/2010
        'Last modified by: ZaMa
        '14/11/2010: ZaMa - Now denounces can be desactivated.
        '**************************************************************
        On Error Resume Next

        Dim Loopc As Long
    
        Select Case sndRoute

                Case SendTarget.ToPCArea
                        Call SendToUserArea(sndIndex, sndData)
                        Exit Sub
        
                Case SendTarget.ToAdmins

                        For Loopc = 1 To LastUser

                                If UserList(Loopc).ConnID <> -1 Then
                                        If UserList(Loopc).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

                                                ' Denounces can be desactivated
                                                If IsDenounce Then
                                                        If UserList(Loopc).flags.SendDenounces Then
                                                                Call EnviarDatosASlot(Loopc, sndData)

                                                        End If

                                                Else
                                                        Call EnviarDatosASlot(Loopc, sndData)

                                                End If

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToAll

                        For Loopc = 1 To LastUser

                                If UserList(Loopc).ConnID <> -1 Then
                                        If UserList(Loopc).flags.UserLogged Then 'Esta logeado como usuario?
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToAllButIndex

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) And (Loopc <> sndIndex) Then
                                        If UserList(Loopc).flags.UserLogged Then 'Esta logeado como usuario?
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.toMap
                        Call SendToMap(sndIndex, sndData)
                        Exit Sub
          
                Case SendTarget.ToMapButIndex
                        Call SendToMapButIndex(sndIndex, sndData)
                        Exit Sub
        
                Case SendTarget.ToGuildMembers
                        Loopc = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

                        While Loopc > 0

                                If (UserList(Loopc).ConnID <> -1) Then
                                        Call EnviarDatosASlot(Loopc, sndData)

                                End If

                                Loopc = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                        Wend
                        Exit Sub
        
                Case SendTarget.ToDeadArea
                        Call SendToDeadUserArea(sndIndex, sndData)
                        Exit Sub
        
                Case SendTarget.ToPCAreaButIndex
                        Call SendToUserAreaButindex(sndIndex, sndData)
                        Exit Sub
        
                Case SendTarget.ToClanArea
                        Call SendToUserGuildArea(sndIndex, sndData)
                        Exit Sub
        
                Case SendTarget.ToPartyArea
                        Call SendToUserPartyArea(sndIndex, sndData)
                        Exit Sub
        
                Case SendTarget.ToAdminsAreaButConsejeros
                        Call SendToAdminsButConsejerosArea(sndIndex, sndData)
                        Exit Sub
        
                Case SendTarget.ToNPCArea
                        Call SendToNpcArea(sndIndex, sndData)
                        Exit Sub
        
                Case SendTarget.ToDiosesYclan
                        Loopc = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

                        While Loopc > 0

                                If (UserList(Loopc).ConnID <> -1) Then
                                        Call EnviarDatosASlot(Loopc, sndData)

                                End If

                                Loopc = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                        Wend
            
                        Loopc = modGuilds.Iterador_ProximoGM(sndIndex)

                        While Loopc > 0

                                If (UserList(Loopc).ConnID <> -1) Then
                                        Call EnviarDatosASlot(Loopc, sndData)

                                End If

                                Loopc = modGuilds.Iterador_ProximoGM(sndIndex)
                        Wend
            
                        Exit Sub
        
                Case SendTarget.ToConsejo

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If UserList(Loopc).flags.Privilegios And PlayerType.RoyalCouncil Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToConsejoCaos

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If UserList(Loopc).flags.Privilegios And PlayerType.ChaosCouncil Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToRolesMasters

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If UserList(Loopc).flags.Privilegios And PlayerType.RoleMaster Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToCiudadanos

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If Not criminal(Loopc) Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToCriminales

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If criminal(Loopc) Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToReal

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If UserList(Loopc).Faccion.ArmadaReal = 1 Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToCaos

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If UserList(Loopc).Faccion.FuerzasCaos = 1 Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToCiudadanosYRMs

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If Not criminal(Loopc) Or (UserList(Loopc).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToCriminalesYRMs

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If criminal(Loopc) Or (UserList(Loopc).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToRealYRMs

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If UserList(Loopc).Faccion.ArmadaReal = 1 Or (UserList(Loopc).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToCaosYRMs

                        For Loopc = 1 To LastUser

                                If (UserList(Loopc).ConnID <> -1) Then
                                        If UserList(Loopc).Faccion.FuerzasCaos = 1 Or (UserList(Loopc).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
        
                Case SendTarget.ToHigherAdmins

                        For Loopc = 1 To LastUser

                                If UserList(Loopc).ConnID <> -1 Then
                                        If UserList(Loopc).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                                                Call EnviarDatosASlot(Loopc, sndData)

                                        End If

                                End If

                        Next Loopc

                        Exit Sub
            
                Case SendTarget.ToGMsAreaButRmsOrCounselors
                        Call SendToGMsAreaButRmsOrCounselors(sndIndex, sndData)
                        Exit Sub
            
                Case SendTarget.ToUsersAreaButGMs
                        Call SendToUsersAreaButGMs(sndIndex, sndData)
                        Exit Sub

                Case SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs
                        Call SendToUsersAndRmsAndCounselorsAreaButGMs(sndIndex, sndData)
                        Exit Sub

        End Select

End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                        If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                                If UserList(tempIndex).ConnIDValida Then
                                        Call EnviarDatosASlot(tempIndex, sdData)

                                End If

                        End If

                End If

        Next Loopc

End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim TempInt   As Integer

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

        If Not MapaValido(Map) Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
            
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

                If TempInt Then  'Esta en el area?
                        TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

                        If TempInt Then
                                If tempIndex <> UserIndex Then
                                        If UserList(tempIndex).ConnIDValida Then
                                                Call EnviarDatosASlot(tempIndex, sdData)

                                        End If

                                End If

                        End If

                End If

        Next Loopc

End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                        If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

                                'Dead and admins read
                                If UserList(tempIndex).ConnIDValida And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
                                        Call EnviarDatosASlot(tempIndex, sdData)

                                End If

                        End If

                End If

        Next Loopc

End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        If UserList(UserIndex).GuildIndex = 0 Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                        If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                                If UserList(tempIndex).ConnIDValida And (UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex Or EsGm(tempIndex)) Then
                                        Call EnviarDatosASlot(tempIndex, sdData)

                                End If

                        End If

                End If

        Next Loopc

End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sdData As String)

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        If UserList(UserIndex).PartyIndex = 0 Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                        If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                                If UserList(tempIndex).ConnIDValida And UserList(tempIndex).PartyIndex = UserList(UserIndex).PartyIndex Then
                                        Call EnviarDatosASlot(tempIndex, sdData)

                                End If

                        End If

                End If

        Next Loopc

End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, _
                                          ByVal sdData As String)

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                        If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                                If UserList(tempIndex).ConnIDValida Then
                                        If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then _
                                           Call EnviarDatosASlot(tempIndex, sdData)

                                End If

                        End If

                End If

        Next Loopc

End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim TempInt   As Integer

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = Npclist(NpcIndex).Pos.Map
        AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
        AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

                If TempInt Then  'Esta en el area?
                        TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

                        If TempInt Then
                                If UserList(tempIndex).ConnIDValida Then
                                        Call EnviarDatosASlot(tempIndex, sdData)

                                End If

                        End If

                End If

        Next Loopc

End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, _
                           ByVal AreaX As Integer, _
                           ByVal AreaY As Integer, _
                           ByVal sdData As String)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim TempInt   As Integer

        Dim tempIndex As Integer
    
        AreaX = 2 ^ (AreaX \ 9)
        AreaY = 2 ^ (AreaY \ 9)
    
        If Not MapaValido(Map) Then Exit Sub

        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
            
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

                If TempInt Then  'Esta en el area?
                        TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

                        If TempInt Then
                                If UserList(tempIndex).ConnIDValida Then
                                        Call EnviarDatosASlot(tempIndex, sdData)

                                End If

                        End If

                End If

        Next Loopc

End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String)

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 5/24/2007
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        If Not MapaValido(Map) Then Exit Sub

        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(tempIndex, sdData)

                End If

        Next Loopc

End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 5/24/2007
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim Map       As Integer

        Dim tempIndex As Integer
    
        Map = UserList(UserIndex).Pos.Map
    
        If Not MapaValido(Map) Then Exit Sub

        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(tempIndex, sdData)

                End If

        Next Loopc

End Sub

Private Sub SendToGMsAreaButRmsOrCounselors(ByVal UserIndex As Integer, _
                                            ByVal sdData As String)

        '**************************************************************
        'Author: Torres Patricio(Pato)
        'Last Modify Date: 12/02/2010
        '12/02/2010: ZaMa - Restrinjo solo a dioses, admins y gms.
        '15/02/2010: ZaMa - Cambio el nombre de la funcion (viejo: ToGmsArea, nuevo: ToGmsAreaButRMsOrCounselors)
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                With UserList(tempIndex)

                        If .AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                                If .AreasInfo.AreaReciveY And AreaY Then
                                        If .ConnIDValida Then

                                                ' Exclusivo para dioses, admins y gms
                                                If (.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero _
                                                   And Not PlayerType.RoleMaster) = .flags.Privilegios Then
                                                        Call EnviarDatosASlot(tempIndex, sdData)

                                                End If

                                        End If

                                End If

                        End If

                End With

        Next Loopc

End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String)

        '**************************************************************
        'Author: Torres Patricio(Pato)
        'Last Modify Date: 10/17/2009
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                        If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                                If UserList(tempIndex).ConnIDValida Then
                                        If UserList(tempIndex).flags.Privilegios And PlayerType.User Then
                                                Call EnviarDatosASlot(tempIndex, sdData)

                                        End If

                                End If

                        End If

                End If

        Next Loopc

End Sub

Private Sub SendToUsersAndRmsAndCounselorsAreaButGMs(ByVal UserIndex As Integer, _
                                                     ByVal sdData As String)

        '**************************************************************
        'Author: Torres Patricio(Pato)
        'Last Modify Date: 10/17/2009
        '
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer
    
        Dim Map       As Integer

        Dim AreaX     As Integer

        Dim AreaY     As Integer
    
        Map = UserList(UserIndex).Pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
        If Not MapaValido(Map) Then Exit Sub
    
        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                        If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                                If UserList(tempIndex).ConnIDValida Then
                                        If UserList(tempIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                                                Call EnviarDatosASlot(tempIndex, sdData)

                                        End If

                                End If

                        End If

                End If

        Next Loopc

End Sub

Public Sub AlertarFaccionarios(ByVal UserIndex As Integer)

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 17/11/2009
        'Alerta a los faccionarios, dandoles una orientacion
        '**************************************************************
        Dim Loopc     As Long

        Dim tempIndex As Integer

        Dim Map       As Integer

        Dim Font      As FontTypeNames
    
        If esCaos(UserIndex) Then
                Font = FontTypeNames.FONTTYPE_CONSEJOCAOS
        Else
                Font = FontTypeNames.FONTTYPE_CONSEJO

        End If
    
        Map = UserList(UserIndex).Pos.Map
    
        If Not MapaValido(Map) Then Exit Sub

        For Loopc = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(Loopc)
        
                If UserList(tempIndex).ConnIDValida Then
                        If tempIndex <> UserIndex Then

                                ' Solo se envia a los de la misma faccion
                                If SameFaccion(UserIndex, tempIndex) Then
                                        Call EnviarDatosASlot(tempIndex, _
                                           PrepareMessageConsoleMsg("Escuchas el llamado de un compañero que proviene del " & _
                                           GetDireccion(UserIndex, tempIndex), Font))

                                End If

                        End If

                End If

        Next Loopc

End Sub
