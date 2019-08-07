Attribute VB_Name = "m_Castillo"
Option Explicit

Public Type tCastillo
    ReyCastillo As Integer
    Due�o As Integer
    Mapa As Integer
    Tiempo As Long
End Type

Public Castillo As tCastillo

Public Sub PasaMinutoCastillo()

' @@ Cuicui

    On Error GoTo errhandleR

    With Castillo
1        .Tiempo = .Tiempo + 1
        
2        If .Tiempo >= 60 Then
            
3            If .Due�o > 0 Then
            
4                guilds(.Due�o).Add_GuildExp (2)

5                Dim tmpHoras As Integer
6                tmpHoras = guilds(.Due�o).GetGuildHorasConquistadas
                
7                Call guilds(.Due�o).SetGuildHorasConquistadas(tmpHoras + 1)
                            
8                Call CheckRankingClan(.Due�o, TopHorasConquistadas)
                
9                Call SaveRanking(TopHorasConquistadas)
                
10            End If
            
11            .Tiempo = 0
12        End If
        
13    End With

Exit Sub

errhandleR:

LogError "Error en PasaMinutoCastillo en linea " & Erl & " - Error: " & Err.Number & " " & Err.description

End Sub

Public Sub IniciarCastilloPretoriano()
    
    ClanPretoriano(1).IsThiefActivated = val(GetVar(App.Path & "/Dat/Castillos.DAT", "MAIN", "SpawnearLadron"))
    ClanPretoriano(2).IsThiefActivated = val(GetVar(App.Path & "/Dat/Castillos.DAT", "MAIN", "SpawnearLadron"))
    
    ' @ Iniciamos los pretorianos
    If Not ClanPretoriano(1).SpawnClan(Castillo.Mapa, 39, 16, 1) Then
        
        Exit Sub
    End If

End Sub
    
Public Sub LoadCastillos()

On Error GoTo default
    ' @ Cargamos el castillo
    With Castillo
3        .ReyCastillo = GetVar(App.Path & "\Dat\Castillos.dat", "MAIN", "ReyCastillo")
        
        
1        .Mapa = GetVar(App.Path & "\Dat\Castillos.dat", "MAIN", "Mapa")
2        .Due�o = GetVar(App.Path & "\Dat\Castillos.dat", "MAIN", "Due�o")
    End With
    Exit Sub
    
default:
    
    LogError "ERROR EN LOADCASTILLOS " & Erl
    Castillo.Due�o = 0
    
End Sub

Public Sub EnviarInfoCastillo(ByVal Userindex As Integer)
    ' @ Enviamos la info del castillo
    With UserList(Userindex)
        If Castillo.Due�o > 0 Then
            Call WriteConsoleMsg(Userindex, "El castillo est� protegido por el clan " & modGuilds.GuildName(Castillo.Due�o) & ".", FontTypeNames.FONTTYPE_MARRON)
        Else
            Call WriteConsoleMsg(Userindex, "El castillo no est� siendo protegido por ning�n clan. �Aprovecha a conquistarlo con el tuyo!", FontTypeNames.FONTTYPE_MARRON)
        End If
        Call WriteConsoleMsg(Userindex, "Para obtener m�s informaci�n sobre el castillo tipea /CASTILLO.", FontTypeNames.FONTTYPE_MARRON)
    End With
End Sub

Public Sub ClanConquistaCastillo(ByVal GuildIndex As Integer)

    On Error GoTo errhandleR

    ' @ Se conquisto el castillo
17    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El castillo de Iriwynne ha caido. Ahora le pertenece al clan " & modGuilds.GuildName(GuildIndex) & ".", FontTypeNames.FONTTYPE_MARRON))

    Castillo.Due�o = GuildIndex
    Castillo.Tiempo = 0
    
16        Call WriteVar(App.Path & "\Dat\Castillos.dat", "MAIN", "Due�o", CStr(Castillo.Due�o))


    Dim i As Long, Integrantes() As String
    Dim tmpIndex As Integer

1    Call IniciarCastilloPretoriano

7    Integrantes = guilds(GuildIndex).GetMemberList
6    For i = LBound(Integrantes) To UBound(Integrantes)
         tmpIndex = NameIndex(Integrantes(i))
5        If tmpIndex > 0 Then
4            WarpUserChar tmpIndex, UserList(tmpIndex).pos.Map, UserList(tmpIndex).pos.X, UserList(tmpIndex).pos.Y, False, False
3        End If
2    Next i

15    If Castillo.Due�o > 0 Then
14        Integrantes = guilds(Castillo.Due�o).GetMemberList
13        For i = LBound(Integrantes) To UBound(Integrantes)
            
12       tmpIndex = NameIndex(Integrantes(i))
57        If tmpIndex > 0 Then
49            WarpUserChar tmpIndex, UserList(tmpIndex).pos.Map, UserList(tmpIndex).pos.X, UserList(tmpIndex).pos.Y, False, False
689        End If

9098        Next i
8767    End If
    
    Exit Sub

errhandleR:
    
    LogError "Error en ClanConquistaCastillo. " & Err.Number & " " & Err.description & " - Linea: " & Erl
    
    Resume Next

End Sub

Public Sub CheckFlodeoRey(ByVal Userindex As Integer, ByVal NpcIndex As Integer, ByVal GuildName As String)
    
    ' @ Avisamos por consola las distintas etapa de la conquista
    With Npclist(NpcIndex)
        If RandomNumber(1, 100) <= 15 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El castillo est� siendo atacado por el clan " & GuildName & ".", FontTypeNames.FONTTYPE_MARRON))
        End If
    End With
    
End Sub

Public Function PuedeAtacarCastillo(ByVal Userindex As Integer) As Boolean
    ' @ No podemos atacar nuestro castillo
    PuedeAtacarCastillo = False
        With UserList(Userindex)
            If Castillo.Mapa = .pos.Map Then
                If Castillo.Due�o = .GuildIndex Then
                PuedeAtacarCastillo = False
                    Exit Function
                End If
            End If
        End With
    PuedeAtacarCastillo = True
End Function


