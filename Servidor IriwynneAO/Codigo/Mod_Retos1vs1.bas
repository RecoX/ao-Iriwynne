Attribute VB_Name = "Mod_Retos1vs1"
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
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

'*****************************************************************

Option Explicit

Public Const MIN_GOLD As Long = 50000

Private Type player_Struct

        player_Index     As Integer
        round_Counter    As Byte

End Type

Public Type reto_Struct

        player_List(1)   As player_Struct
        count_Down       As Byte
        used_slot        As Boolean
        nextRoundCounter As Integer
       
        gold_gamble      As Long
        drop_gamble      As Boolean

End Type

Public Type userReto_Struct

        reto_Index       As Integer
        request_name     As String
        send_to_index    As String
       
        return_home      As Byte
        acceptLimitCount As Byte
       
        temp_goldGamble  As Long
        temp_dropGamble  As Boolean

End Type

Private Type retoPosStructs

        X As Integer
        Y As Integer

End Type

Private retoPoss()      As retoPosStructs

Public retoList(1 To 5) As reto_Struct

Private Function get_reto_slot() As Integer

        '
    
        Dim i As Long
    
        For i = 1 To UBound(retoList())

                If (retoList(i).used_slot = False) Then Exit For
        Next i
    
        If (i > UBound(retoList())) Then
                get_reto_slot = 0
        Else
                get_reto_slot = CInt(i)

        End If

End Function

Public Function can_sendReto(ByVal send_index As Integer, _
                             ByRef other_name As String, _
                             ByVal m_gold As Long, _
                             ByVal m_drop As Boolean, _
                             ByRef m_error As String) As Boolean

        '
    
        can_sendReto = False
    
        Dim other_index As Integer
    
        other_index = NameIndex(other_name)
    
        If (other_index = 0) Then
                m_error = other_name & " no est� online."

                Exit Function

        End If
    
        If m_gold < MIN_GOLD Then
                m_error = "La apuesta m�nima de oro es de " & CStr(MIN_GOLD) & " monedas de oro."

                Exit Function

        End If
        
        If send_index = other_index Then
            m_error = "No podes retarte a vos mismo."
            Exit Function
        End If
    
        can_sendReto = (check_player(send_index, m_gold, m_error) = True)
    
        If (can_sendReto) Then
                can_sendReto = (check_player(other_index, m_gold, m_error) = True)
        Else
                m_error = Replace$(m_error, UserList(send_index).Name & " ", vbNullString)
                m_error = Replace$(m_error, "est�", "est�s")
                m_error = Replace$(m_error, "tiene", "tienes")

        End If
    
End Function

Private Function check_player(ByVal player_Index As Integer, _
                              ByVal m_gold As Long, _
                              ByRef f_error As String) As Boolean

        '

        check_player = False
    
        With UserList(player_Index)

                If (.flags.Muerto <> 0) Then
                        f_error = .Name & " est� muerto."

                        Exit Function

                End If
         
                If (.Counters.Pena <> 0) Then
                        f_error = .Name & " est� en la c�rcel."

                        Exit Function

                End If
         
                If (.Stats.ELV < 35) Then
                        f_error = .Name & " tiene que ser mayor al nivel 35."

                        Exit Function

                End If
         
                If (.pos.Map <> eCiudad.cUllathorpe) Then
                        f_error = .Name & " est� fuera de su hogar."

                        Exit Function

                End If
         
                If (.mReto.reto_Index <> 0) Or (.sReto.reto_used = True) Then
                        f_error = .Name & " ya est� en reto."

                        Exit Function

                End If
         
                If (.Stats.GLD < m_gold) Then
                        f_error = .Name & " no tiene el oro suficiente."

                        Exit Function

                End If
         
                check_player = True

        End With

End Function

Public Sub send_Reto(ByVal send_index As Integer, _
                     ByVal other_index As Integer, _
                     ByVal GoldAmount As Long, _
                     ByVal dropItem As Boolean)

        '
    
        With UserList(send_index)

                Dim gamble_str As String
         
                gamble_str = "apostando " & Format$(GoldAmount, "#,###") & " monedas de oro"
         
                If (dropItem) Then
                        gamble_str = gamble_str & " y los items del inventario"

                End If
         
                .mReto.temp_dropGamble = dropItem
                .mReto.temp_goldGamble = GoldAmount
                .mReto.acceptLimitCount = 60
                .mReto.send_to_index = UserList(other_index).Name
         
                UserList(other_index).mReto.request_name = UCase$(.Name)
         
                Call Protocol.WriteConsoleMsg(send_index, "La solicitud ha sido enviada.", FontTypeNames.FONTTYPE_GUILD)
                Call Protocol.WriteConsoleMsg(other_index, "Solicitud de reto modalidad 1vs1 : " & .Name & " te desaf�a en un reto " & gamble_str & " si aceptas tipea /RETAR " & UCase$(.Name) & "." & vbNewLine & "Tienes 60 segundos para aceptar el reto, de lo contrario se auto-cancelar�.", FontTypeNames.FONTTYPE_GUILD)
                
                If .mReto.temp_dropGamble Then
                Call Protocol.WriteConsoleMsg(other_index, "Es por items!", FontTypeNames.FONTTYPE_FIGHT)
                End If
        End With
    
End Sub

Public Sub accept_Reto(ByVal user_Index As Integer, ByRef other_name As String)

'

    Dim send_index As Integer
    Dim tError As String
    If (Len(UserList(user_Index).mReto.request_name) = 0) Then Exit Sub

    If (UserList(user_Index).mReto.request_name <> other_name) Then
        Call Protocol.WriteConsoleMsg(user_Index, other_name & " no te est� retando.", FontTypeNames.FONTTYPE_GUILD)

        Exit Sub

    End If

    send_index = NameIndex(other_name)

    If (send_index <> 0) Then

        If can_AcceptReto(user_Index, send_index, tError) Then
            Call Protocol.WriteConsoleMsg(send_index, UserList(user_Index).Name & " acept� el reto.", FontTypeNames.FONTTYPE_GUILD)

            UserList(user_Index).mReto.acceptLimitCount = 0
            UserList(send_index).mReto.acceptLimitCount = 0

            Call init_reto(send_index, user_Index, UserList(send_index).mReto.temp_goldGamble, UserList(send_index).mReto.temp_dropGamble)
        Else
            Call WriteConsoleMsg(user_Index, tError, FontTypeNames.FONTTYPE_GUILD)
        End If
    Else
        Call Protocol.WriteConsoleMsg(user_Index, "El reto se ha cancelado porque " & other_name & " se ha desconectado.", FontTypeNames.FONTTYPE_GUILD)

    End If

End Sub

Private Function can_AcceptReto(ByVal userAccept As Integer, ByVal userSend As Integer, ByRef errorStr As String) As Boolean

    can_AcceptReto = False

    Dim gldAmount As Long

    gldAmount = UserList(userSend).mReto.temp_goldGamble

    With UserList(userAccept)

        If (.pos.Map <> 1) Then
            errorStr = "Debes estar en tu hogar para participar en un reto."
            Exit Function
        End If

        If (.flags.Muerto <> 0) Then
            errorStr = "No puedes retar en ese estado!"
            Exit Function
        End If

        If (.flags.Comerciando <> 0) Then
            errorStr = "Debes dejar de comerciar."
            Exit Function
        End If

        If (.Stats.GLD < gldAmount) Or (gldAmount < 0) Then
            errorStr = "No tienes el oro suficiente."
            Exit Function
        End If

        If .mReto.reto_Index <> 0 Then
            errorStr = "Ya est�s en reto!"
            Exit Function
        End If

        If .sReto.reto_used Then
            errorStr = "Ya est�s en reto!"
            Exit Function
        End If

    End With

    ' @@ El enviador esta en condiciones?

    With UserList(userSend)

        gldAmount = .mReto.temp_goldGamble

        If (.pos.Map <> 1) Then
            errorStr = "El oponente est� fuera de su hogar."
            Exit Function
        End If

        If (.flags.Muerto <> 0) Then
            errorStr = "Est� muerto."
            Exit Function
        End If

        If (.flags.Comerciando <> 0) Then
            errorStr = "El oponente est� comerciando."
            Exit Function
        End If

        If (.Stats.GLD < gldAmount) Or (gldAmount < 0) Then
            errorStr = "El oponente no tiene el oro suficiente."
            Exit Function
        End If

        If (.mReto.reto_Index <> 0) Then
            errorStr = "El oponente ya est� en reto!"
            Exit Function
        End If

        If (.sReto.reto_used) Then
            errorStr = "El oponente ya est� en reto!"
            Exit Function
        End If
        

    End With

    can_AcceptReto = True

End Function



Private Sub init_reto(ByVal send_index As Integer, _
                      ByVal other_index As Integer, _
                      ByVal gold As Long, _
                      ByVal Drop As Boolean)

        '
    
        Dim reto_Index As Integer
    
        reto_Index = get_reto_slot()
    
        If (reto_Index = 0) Then
                Call Protocol.WriteConsoleMsg(send_index, "El reto no ha podido iniciarse porque todas las salas est�n siendo ocupadas.", FontTypeNames.FONTTYPE_GUILD)
                Call Protocol.WriteConsoleMsg(other_index, "El reto no ha podido iniciarse porque todas las salas est�n siendo ocupadas.", FontTypeNames.FONTTYPE_GUILD)
        Else
        
                With retoList(reto_Index)
                        .count_Down = 6
                        .drop_gamble = Drop
                        .gold_gamble = gold
                        .player_List(0).player_Index = send_index
                        .player_List(0).round_Counter = 0
                        .player_List(1).player_Index = other_index
                        .player_List(1).round_Counter = 0
             
                        UserList(send_index).mReto.reto_Index = reto_Index
                        UserList(other_index).mReto.reto_Index = reto_Index
             
                        Call Protocol.WritePauseToggle(.player_List(0).player_Index)
                        Call Protocol.WritePauseToggle(.player_List(1).player_Index)
             
                        Call warp_Players(reto_Index)
             
                        .used_slot = True

                End With
        
        End If

End Sub

Public Sub userdie_reto(ByVal user_Index As Integer)

        '
    
    On Error GoTo errhandleR
    
        Dim other_user As Integer

        Dim reto_Index As Integer
    
12        reto_Index = UserList(user_Index).mReto.reto_Index
    
11        If (reto_Index = 0) Then Exit Sub
        
10       If (retoList(reto_Index).used_slot = False) Then logerror "Usedslot_False": Exit Sub
9
7        other_user = IIf(retoList(UserList(user_Index).mReto.reto_Index).player_List(0).player_Index = user_Index, 1, 0)
    
6        other_user = retoList(reto_Index).player_List(other_user).player_Index
    
5        If (other_user <> 0) Then
4                If (UserList(other_user).ConnID <> -1) Then
3                        Call winner_Reto(user_Index, other_user)

2                End If

1        End If
    
        Exit Sub
        
errhandleR:
  logerror "Error en userDie_Reto en " & Erl
    
End Sub

Private Sub winner_Reto(ByVal die_index As Integer, ByVal live_index As Integer)

        
On Error GoTo errhandleR

        '
    
        Dim reto_Index As Integer

        Dim winner_id  As Byte
    
11        reto_Index = UserList(die_index).mReto.reto_Index

12        With retoList(reto_Index)
    
10                winner_id = IIf(.player_List(0).player_Index = die_index, 1, 0)
9
8                .player_List(winner_id).round_Counter = (.player_List(winner_id).round_Counter + 1)
7
6                If (.player_List(winner_id).round_Counter = 2) Then
5                        Call end_reto(reto_Index, winner_id)
4                Else
3                        Call respawn_reto(reto_Index, winner_id)

2                End If
         
1        End With

Exit Sub

errhandleR:

logerror "Error en Wionner_reto en " & Erl

End Sub

Public Sub disconnectUser_reto(ByVal user_Index As Integer)

On Error GoTo errhandleR

        '
    
        Dim winnerID As Byte
    
1        winnerID = IIf(retoList(UserList(user_Index).mReto.reto_Index).player_List(0).player_Index = user_Index, 1, 0)
    
2        Call end_reto(UserList(user_Index).mReto.reto_Index, winnerID, True)
     
    Exit Sub
    
errhandleR:

logerror "Error en disconnectuser_reto en " & Erl

    
End Sub

Private Sub respawn_reto(ByVal reto_Index As Integer, ByVal winner_index As Byte)

        '
    
        Dim i As Long

        Dim T As String
    
        With retoList(reto_Index)
             
                T = UserList(.player_List(winner_index).player_Index).Name & " gana este round." & vbNewLine & "Resultado parcial : " & .player_List(0).round_Counter & "-" & .player_List(1).round_Counter & "!"
        
                For i = 0 To 1
                        Call Protocol.WriteConsoleMsg(.player_List(i).player_Index, T, FontTypeNames.FONTTYPE_GUILD)
                        Call Protocol.WriteConsoleMsg(.player_List(i).player_Index, "El siguiente round iniciar� en 5 segundos.", FontTypeNames.FONTTYPE_GUILD)
                Next i
         
                .nextRoundCounter = 6
         
                'Call warp_Players(reto_Index, True)
         
        End With
    
End Sub

Private Sub end_reto(ByVal reto_Index As Integer, _
                     ByVal winner As Byte, _
                     Optional ByVal disconnected As Boolean = False)

        '
    
        With retoList(reto_Index)
         
                Dim winner_index As Integer

                Dim looser_index As Integer

                Dim ullathorpe_p As WorldPos
                Dim k            As WorldPos

                'Dim rankIndex    As Integer
         
                winner_index = .player_List(winner).player_Index
                looser_index = .player_List(IIf(winner = 0, 1, 0)).player_Index
                
                Call SendData(SendTarget.ToAll, 0, Protocol.PrepareMessageConsoleMsg("Retos> " & UserList(winner_index).Name & " venci� en un reto a " & UserList(looser_index).Name & IIf(disconnected = True, " (Por desconexi�n)", "."), FontTypeNames.FONTTYPE_GUILD))
         
                If (.drop_gamble) Then
                        Call TirarTodosLosItems(looser_index)

                End If
         
                UserList(looser_index).Stats.GLD = (UserList(looser_index).Stats.GLD - .gold_gamble)
         
                'UserList(looser_index).Ranking.looser_counter = UserList(looser_index).Ranking.looser_counter + 1
                'UserList(looser_index).Ranking.gold_looser = UserList(looser_index).Ranking.gold_looser + .gold_gamble
         
                'rankIndex = modRankingRetos.in_Ranking(looser_index)
                  
                'If (rankIndex <> -1) Then
                '    ranking_List(rankIndex).user_Ranking = UserList(looser_index).Ranking
                'End If
         
                Call Protocol.WriteUpdateGold(looser_index)
                Call Protocol.WriteConsoleMsg(looser_index, "Has perdido el reto.", FontTypeNames.FONTTYPE_GUILD)
         
                ullathorpe_p.Map = 1
                ullathorpe_p.X = 50
                ullathorpe_p.Y = 50
         
                'Call FindLegalPos(looser_index, ullathorpe_p.Map, ullathorpe_p.X, ullathorpe_p.Y)
                
                Call ClosestStablePos(ullathorpe_p, k)
                
                Call WarpUserChar(looser_index, ullathorpe_p.Map, k.X, k.Y, True)
         
                If (.drop_gamble) Then
                        UserList(winner_index).mReto.return_home = 10
             
                        Call Protocol.WriteConsoleMsg(winner_index, "Has ganado el reto, en 10 segundos volver�s a la ciudad.", FontTypeNames.FONTTYPE_GUILD)
                Else
                        'ullathorpe_p = Ullathorpe
            
                        'Call FindLegalPos(winner_index, ullathorpe_p.Map, ullathorpe_p.X, ullathorpe_p.Y)
                        Call ClosestStablePos(ullathorpe_p, k)
                        Call WarpUserChar(winner_index, ullathorpe_p.Map, k.X, k.Y, True)
                        
                        Call reset_userReto(winner_index)

                End If
         
                UserList(winner_index).Stats.GLD = (UserList(winner_index).Stats.GLD + .gold_gamble)
         
                'UserList(winner_index).Ranking.winner_counter = UserList(winner_index).Ranking.winner_counter + 1
                'UserList(winner_index).Ranking.gold_winner = UserList(winner_index).Ranking.gold_winner + .gold_gamble
                  
                'rankIndex = modRankingRetos.ingress_Ranking(winner_index)
                  
                'If (rankIndex <> -1) Then
                '    Call modRankingRetos.set_Ranking(winner_index, rankIndex)
                'End If
         
                'rankIndex = modRankingRetos.in_Ranking(winner_index)
                  
                'If (rankIndex <> -1) Then
                '    ranking_List(rankIndex).user_Ranking = UserList(winner_index).Ranking
                'End If
         
                Call Protocol.WriteUpdateGold(winner_index)
                  
                Call reset_userReto(looser_index)
                
                Call erase_retoData(reto_Index)
                
        End With

End Sub

Private Sub erase_retoData(ByVal reto_Index As Integer)

        '
    
        With retoList(reto_Index)
    
                .count_Down = 0
                .drop_gamble = False
                .gold_gamble = 0
                .used_slot = False
         
                Dim i As Long
         
                For i = 0 To 1
                        .player_List(i).player_Index = 0
                        .player_List(i).round_Counter = 0
                Next i
    
        End With

End Sub

Private Sub warp_Players(ByVal reto_Index As Integer, _
                         Optional ByVal respawn As Boolean = False)

        '
    
        With retoList(reto_Index)
         
                Dim i As Long

                Dim N As Integer

                Dim p As WorldPos
         
                p.Map = Configuracion.Mapa1vs1
         
                For i = 0 To 1
                        N = .player_List(i).player_Index
             
                        If (N > 0) Then
                                If (UserList(N).ConnID <> -1) Then
                                        p.X = give_pos_X(reto_Index, i + 1)
                                        p.Y = give_pos_Y(reto_Index, i + 1)
                     
                                        Call WarpUserChar(N, p.Map, p.X, p.Y, True)
                     
                                        If (respawn) Then
                                                If UserList(N).flags.Muerto Then
                                                        Call RevivirUsuario(N)

                                                End If
                         
                                                UserList(N).Stats.MinHp = UserList(N).Stats.MaxHP
                                                UserList(N).Stats.MinMAN = UserList(N).Stats.MaxMAN
                                                UserList(N).Stats.MinSta = UserList(N).Stats.MaxSta
                                                UserList(N).Stats.MinAGU = 100
                                                UserList(N).Stats.MinHam = 100
                         
                                                Call Protocol.WriteUpdateUserStats(N)

                                        End If
                     
                                End If

                        End If

                Next i
    
        End With

End Sub

Public Sub reto_all_loop()

        '
    
        Dim i As Long
    
        For i = 1 To UBound(retoList())

                If (retoList(i).used_slot) Then Call reto_loop(CInt(i))
        Next i

End Sub

Private Sub reto_loop(ByVal reto_Index As Integer)

        '
    
        Dim T As String

        Dim i As Long

        Dim N As Integer

        Dim p As WorldPos
    
        With retoList(reto_Index)

                If (.nextRoundCounter <> 0) Then
                        .nextRoundCounter = (.nextRoundCounter - 1)
             
                        If (.nextRoundCounter = 0) Then

                                For i = 0 To 1
                                        N = .player_List(i).player_Index
                     
                                        If (N > 0) Then
                                                If UserList(N).ConnID <> -1 Then
                                                        p.Map = Configuracion.Mapa1vs1
                                                        p.X = give_pos_X(reto_Index, i + 1)
                                                        p.Y = give_pos_Y(reto_Index, i + 1)
                            
                                                        Call WarpUserChar(N, p.Map, p.X, p.Y, True)
                            
                                                        Call Protocol.WritePauseToggle(N)
                            
                                                        If UserList(N).flags.Muerto Then
                                                                Call RevivirUsuario(N)

                                                        End If
                         
                                                        UserList(N).Stats.MinHp = UserList(N).Stats.MaxHP
                                                        UserList(N).Stats.MinMAN = UserList(N).Stats.MaxMAN
                                                        UserList(N).Stats.MinSta = UserList(N).Stats.MaxSta
                                                        UserList(N).Stats.MinAGU = 100
                                                        UserList(N).Stats.MinHam = 100
                            
                                                        Call Protocol.WriteUpdateUserStats(N)
                            
                                                End If

                                        End If

                                Next i
                 
                                .count_Down = 6
                 
                        End If
             
                End If
         
                If (.count_Down <> 0) Then
                        .count_Down = .count_Down - 1
             
                        If (.count_Down = 0) Then
                                T = "�YA!"
                        Else
                                T = CStr(.count_Down) & "..."

                        End If
             
                        For i = 0 To 1
                                N = .player_List(i).player_Index
                 
                                If (N <> 0) Then
                                        If (UserList(N).ConnID <> -1) Then
                                                Call Protocol.WriteConsoleMsg(N, T, FontTypeNames.FONTTYPE_GUILD)

                                        End If

                                End If
                 
                                If (.count_Down = 0) Then Call Protocol.WritePauseToggle(N)
                        Next i
             
                End If

        End With

End Sub

Public Sub loop_userReto(ByVal user_Index As Integer)

        '
    
        With UserList(user_Index).mReto
        
                If (.acceptLimitCount <> 0) Then
                        .acceptLimitCount = .acceptLimitCount - 1
             
                        If (.acceptLimitCount = 0) Then

                                Dim temp_index As Integer
                 
                                temp_index = NameIndex(.send_to_index)
                 
                                If (temp_index <> 0) Then
                                        Call Protocol.WriteConsoleMsg(temp_index, "La solicitud de reto de " & UserList(user_Index).Name & " ha sido cancelada porque acab� el tiempo l�mite para aceptar.", FontTypeNames.FONTTYPE_GUILD)
                     
                                        UserList(temp_index).mReto.request_name = vbNullString

                                End If
                 
                                Call reset_userReto(user_Index)

                        End If
             
                End If
         
                If (.return_home <> 0) Then
                        .return_home = (.return_home - 1)
             
                        If (.return_home = 0) Then

                                Dim p As WorldPos
                                Dim k As WorldPos
                                
                                p.Map = 1
                                p.X = 50
                                p.Y = 50
        
                                'Call FindLegalPos(user_Index, p.Map, p.X, p.Y)
                                Call ClosestStablePos(p, k)
                                Call WarpUserChar(user_Index, p.Map, k.X, k.Y, True)
                 
                                Call Protocol.WriteConsoleMsg(user_Index, "Vuelves a la ciudad.", FontTypeNames.FONTTYPE_GUILD)
                                .request_name = vbNullString
                                .reto_Index = 0
                                Call reset_userReto(user_Index)

                        End If

                End If
    
        End With

End Sub

Public Sub reset_userReto(ByVal send_index As Integer)

        '
    
        UserList(send_index).mReto.request_name = vbNullString
        UserList(send_index).mReto.reto_Index = 0
                    
        With UserList(send_index).mReto
        
                .send_to_index = vbNullString
                .temp_dropGamble = False
                .temp_goldGamble = 0
                .request_name = vbNullString
                .return_home = 0
                .acceptLimitCount = 0
                .reto_Index = 0
                                
        End With
    
End Sub

Public Function give_pos_X(ByVal ring_Index As Integer, _
                           ByVal user_Index As Integer) As Integer

        '
        ' @ amishar.-
    
        Dim endPos As Integer
        
        endPos = retoPoss(ring_Index, user_Index).X
    
        give_pos_X = endPos
    
End Function

Public Function give_pos_Y(ByVal ring_Index As Integer, _
                           ByVal user_Index As Integer) As Integer

        '
        ' @ amishar.-
    
        Dim endPos As Integer
    
        endPos = retoPoss(ring_Index, user_Index).Y
    
        give_pos_Y = endPos
    
End Function

Public Sub retos1vs1Load()

        '
        ' @ amishar.-
    
        Dim nArenas As Integer

        Dim bReader As New clsIniManager
    
        bReader.Initialize DatPath & "Retos1vs1.ini"
    
        nArenas = val(bReader.GetValue("INIT", "Arenas"))
    
        If (nArenas = 0) Then Exit Sub
    
        ReDim retoPoss(1 To nArenas, 1 To 2) As retoPosStructs
    
        Dim i As Long

        Dim p As Long

        Dim s As String
    
        For i = 1 To nArenas
                For p = 1 To 2
                        s = bReader.GetValue("ARENA" & CStr(i), "Jugador" & CStr(p))
                
                        retoPoss(i, p).X = val(ReadField(2, s, Asc("-")))
                        retoPoss(i, p).Y = val(ReadField(3, s, Asc("-")))
                
                Next p
        Next i

End Sub

