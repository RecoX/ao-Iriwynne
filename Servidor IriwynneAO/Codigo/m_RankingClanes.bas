Attribute VB_Name = "m_RankingClanes"
Option Explicit

Public Const MAX_TOP As Byte = 10
Public Const MAX_RANKINGS As Byte = 3

Public Type tRanking
    Value(1 To MAX_TOP) As Long
    Nombre(1 To MAX_TOP) As String
End Type

Public Ranking(1 To MAX_RANKINGS) As tRanking

Public Enum eRanking
    TopClanes = 1
    TopHorasConquistadas = 2
    topnivel = 3
End Enum

Public Function RenameRanking(ByVal Ranking As eRanking) As String

'@ Devolvemos el nombre del TAG [] del archivo .DAT
    Select Case Ranking
        Case eRanking.TopHorasConquistadas
            RenameRanking = "Clanes_HorasConquistadas"
        Case eRanking.TopClanes
            RenameRanking = "Clanes_Level"
        Case eRanking.topnivel
            RenameRanking = "NIVEL"
        Case Else
            RenameRanking = vbNullString
    End Select
End Function

Public Function RenameValue(ByVal GuildExp As Integer, ByVal Ranking As eRanking) As Long
' @ Devolvemos a que hace referencia el ranking

        Select Case Ranking
            Case eRanking.TopHorasConquistadas
                RenameValue = guilds(GuildExp).GetGuildHorasConquistadas
            Case eRanking.TopClanes
                RenameValue = guilds(GuildExp).getGuildLevel
            Case eRanking.topnivel
                RenameValue = UserList(GuildExp).Stats.ELV ' Tengo que usar GuildExp jajaja
        End Select
        
End Function

Public Sub LoadRanking()

    Dim LoopI As Integer
    Dim LoopX As Integer
    Dim ln As String

    For LoopX = 1 To MAX_RANKINGS
        For LoopI = 1 To MAX_TOP
            ln = GetVar(App.Path & "\Dat\" & "Ranking.dat", RenameRanking(LoopX), "Top" & LoopI)
            Ranking(LoopX).Nombre(LoopI) = ReadField(1, ln, 45)
            Ranking(LoopX).Value(LoopI) = val(ReadField(2, ln, 45))
        Next LoopI
    Next LoopX

End Sub

Public Sub SaveRanking(ByVal Rank As eRanking)

    Dim LoopI As Integer

    For LoopI = 1 To MAX_TOP
        Call WriteVar(DatPath & "Ranking.Dat", RenameRanking(Rank), _
                      "Top" & LoopI, Ranking(Rank).Nombre(LoopI) & "-" & Ranking(Rank).Value(LoopI))
    Next LoopI
    
End Sub

Public Sub CheckRankingClan(ByVal GuildIndex As Integer, ByVal Rank As eRanking)

    Dim LoopX As Integer
    Dim LoopY As Integer
    Dim LoopZ As Integer
    Dim i As Integer
    Dim Value As Long
    Dim Actualizacion As Byte
    Dim Auxiliar As String
    Dim PosRanking As Byte

        Value = RenameValue(GuildIndex, Rank)

        ' @ Buscamos al personaje en el ranking
        For i = 1 To MAX_TOP
            If Ranking(Rank).Nombre(i) = UCase$(modGuilds.GuildName(GuildIndex)) Then
                PosRanking = i
                Exit For
            End If
        Next i

        ' @ Si el personaje esta en el ranking actualizamos los valores.
        If PosRanking <> 0 Then
            ' �Si est� actualizado pa que?
            If Value <> Ranking(Rank).Value(PosRanking) Then
                Call ActualizarPosRanking(PosRanking, Rank, Value)

                For LoopY = 1 To MAX_TOP
                    For LoopZ = 1 To MAX_TOP - LoopY

                        If Ranking(Rank).Value(LoopZ) < Ranking(Rank).Value(LoopZ + 1) Then

                            ' Actualizamos el valor
                            Auxiliar = Ranking(Rank).Value(LoopZ)
                            Ranking(Rank).Value(LoopZ) = Ranking(Rank).Value(LoopZ + 1)
                            Ranking(Rank).Value(LoopZ + 1) = Auxiliar

                            ' Actualizamos el nombre
                            Auxiliar = Ranking(Rank).Nombre(LoopZ)
                            Ranking(Rank).Nombre(LoopZ) = Ranking(Rank).Nombre(LoopZ + 1)
                            Ranking(Rank).Nombre(LoopZ + 1) = Auxiliar
                            Actualizacion = 1
                        End If
                    Next LoopZ
                Next LoopY
                'End If

                If Actualizacion <> 0 Then
                    Call SaveRanking(Rank)
                End If
            End If

            Exit Sub
        Else
            'Debug.Print "es menor"
        End If

        ' @ Nos fijamos si podemos ingresar al ranking
        For LoopX = 1 To MAX_TOP
            If Value <> Ranking(Rank).Value(LoopX) Then
                Call ActualizarRanking(LoopX, Rank, modGuilds.GuildName(GuildIndex), Value)
                Exit For
            End If
        Next LoopX
        
End Sub

Public Sub ActualizarPosRanking(ByVal Top As Byte, ByVal Rank As eRanking, ByVal Value As Long)
' @ Actualizamos la pos indicada en caso de que el personaje est� en el ranking

    With Ranking(Rank)
        .Value(Top) = Value
    End With

End Sub

Public Sub ActualizarRanking(ByVal Top As Byte, ByVal Rank As eRanking, ByVal UserName As String, ByVal Value As Long)

'@ Actualizamos la lista de ranking

    Dim LoopC As Integer

    Dim Valor(1 To MAX_TOP) As Long
    Dim Nombre(1 To MAX_TOP) As String

    ' @ Copia necesaria para evitar que se dupliquen repetidamente
    For LoopC = 1 To MAX_TOP
        Valor(LoopC) = Ranking(Rank).Value(LoopC)
        Nombre(LoopC) = Ranking(Rank).Nombre(LoopC)
    Next LoopC

    ' @ Corremos las pos, desde el "Top" que es la primera
    For LoopC = Top To MAX_TOP - 1
        Ranking(Rank).Value(LoopC + 1) = Valor(LoopC)
        Ranking(Rank).Nombre(LoopC + 1) = Nombre(LoopC)
    Next LoopC

    Ranking(Rank).Nombre(Top) = UCase$(UserName)
    Ranking(Rank).Value(Top) = Value
    Call SaveRanking(Rank)
    If Rank = TopClanes Or Rank = TopHorasConquistadas Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranking de clan " & IIf(RenameRanking(Rank) = "Clanes_Level", "de niveles de clan", "de clan") & "> El clan " & UserName & " ha subido al TOP " & Top & ".", FontTypeNames.FONTTYPE_GUILD))
    Else
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranking " & RenameRanking(Rank) & "> " & UserName & " ha subido al TOP " & Top & ".", FontTypeNames.FONTTYPE_GUILD))
    End If
End Sub


Public Sub CheckRankingUser(ByVal Userindex As Integer, ByVal Rank As eRanking)
' @ Desde aca nos hacemos la siguientes preguntas
' @ El personaje est� en el ranking?
' @ El personaje puede ingresar al ranking?

    Dim LoopX As Integer
    Dim LoopY As Integer
    Dim LoopZ As Integer
    Dim i As Integer
    Dim Value As Long
    Dim Actualizacion As Byte
    Dim Auxiliar As String
    Dim PosRanking As Byte

    With UserList(Userindex)

        ' @ Not gms
        If esGM(Userindex) Then Exit Sub

        Value = RenameValue(Userindex, Rank)

        ' @ Buscamos al personaje en el ranking
        For i = 1 To MAX_TOP
            If Ranking(Rank).Nombre(i) = UCase$(.Name) Then
                PosRanking = i
                Exit For
            End If
        Next i

        ' @ Si el personaje esta en el ranking actualizamos los valores.
        If PosRanking <> 0 Then
            ' �Si est� actualizado pa que?
            If Value <> Ranking(Rank).Value(PosRanking) Then
                Call ActualizarPosRanking(PosRanking, Rank, Value)

                ' �Es la pos 1? No hace falta ordenarlos
                'If Not PosRanking = 1 Then
                ' @ Chequeamos los datos para actualizar el ranking
                For LoopY = 1 To MAX_TOP
                    For LoopZ = 1 To MAX_TOP - LoopY

                        If Ranking(Rank).Value(LoopZ) < Ranking(Rank).Value(LoopZ + 1) Then

                            ' Actualizamos el valor
                            Auxiliar = Ranking(Rank).Value(LoopZ)
                            Ranking(Rank).Value(LoopZ) = Ranking(Rank).Value(LoopZ + 1)
                            Ranking(Rank).Value(LoopZ + 1) = Auxiliar

                            ' Actualizamos el nombre
                            Auxiliar = Ranking(Rank).Nombre(LoopZ)
                            Ranking(Rank).Nombre(LoopZ) = Ranking(Rank).Nombre(LoopZ + 1)
                            Ranking(Rank).Nombre(LoopZ + 1) = Auxiliar
                            Actualizacion = 1
                        End If
                    Next LoopZ
                Next LoopY
                'End If

                If Actualizacion <> 0 Then
                    Call SaveRanking(Rank)
                End If
            End If

            Exit Sub
        Else
            'Debug.Print "es menor"
        End If

        ' @ Nos fijamos si podemos ingresar al ranking
        For LoopX = 1 To MAX_TOP
            If Value > Ranking(Rank).Value(LoopX) Then
                Call ActualizarRanking(LoopX, Rank, .Name, Value)
                Exit For
            End If
        Next LoopX

    End With
End Sub

