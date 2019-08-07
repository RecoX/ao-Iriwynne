Attribute VB_Name = "Mod_Retos2vs2"
Option Explicit

Private Mapa_Arenas As Integer

Public Type ruleStruct

    drop_inv As Boolean
    gold_gamble As Long

End Type

Public Type teamStruct

    user_Index(1) As Integer
    round_count As Byte
    return_city As Byte

End Type

Public Type retoStruct

    team_array(1) As teamStruct
    general_rules As ruleStruct
    count_Down As Byte
    used_ring As Boolean
    haydrop As Boolean
    nextRoundCount As Integer

End Type

Public Type userStruct

    tempStruct As retoStruct
    accept_count As Byte
    reto_Index As Integer
    nick_sender As String
    reto_used As Boolean
    return_city As Byte
    tmp_Time As Byte
    acceptedOK As Boolean
    acceptLimit As Integer

End Type

Public Type retoPosStruct

    Map As Integer
    X As Integer
    Y As Integer

End Type

Public reto_List() As retoStruct
Public retoPos() As retoPosStruct

Public Sub loop_reto()

'
' @ amishar.-

    Dim LoopC As Long

1   On Error GoTo loop_reto_Error

2   For LoopC = LBound(reto_List()) To UBound(reto_List())

3       If (reto_List(LoopC).used_ring) Then
4           Call loop_reto_index(LoopC)

5       End If

6   Next LoopC

7   Exit Sub

loop_reto_Error:

8   Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure loop_reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Function can_Attack(ByVal attackerIndex As Integer, _
                           ByVal victimIndex As Integer) As Boolean

'
' @ amishar.-

    Dim retoIndex As Integer
    Dim teamIndex As Integer
    Dim tempIndex As Integer
    Dim teamLoop As Long

1   On Error GoTo can_Attack_Error

2   can_Attack = True

3   retoIndex = UserList(attackerIndex).sReto.reto_Index

4   teamIndex = -1

5   If reto_List(retoIndex).used_ring Then

6       For teamLoop = 0 To 1

7           If reto_List(retoIndex).team_array(teamLoop).user_Index(0) = attackerIndex Or reto_List(retoIndex).team_array(teamLoop).user_Index(1) = attackerIndex Then
8               teamIndex = teamLoop

9               Exit For

10          End If

11      Next teamLoop

12      If teamIndex <> -1 Then
13          tempIndex = IIf(reto_List(retoIndex).team_array(teamIndex).user_Index(0) = attackerIndex, 1, 0)

14          If reto_List(retoIndex).team_array(teamIndex).user_Index(tempIndex) = victimIndex Then
15              can_Attack = False

16          End If

17      End If

18  End If

19  Exit Function

can_Attack_Error:

20  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure can_Attack of M�dulo Mod_Retos2vs2" & Erl & ".")

End Function

Private Sub loop_reto_index(ByVal reto_Index As Integer)

'
' @ amishar.-

    Dim i As Long
    Dim j As Long
    Dim H As Integer
    Dim m As String

1   On Error GoTo loop_reto_index_Error

2   With reto_List(reto_Index)

3       If (.nextRoundCount <> 0) Then
4           .nextRoundCount = .nextRoundCount - 1

5           If (.nextRoundCount = 0) Then
6               Call warp_Teams(reto_Index, True)

7               .count_Down = 6

8           End If

9       End If

10      If (.count_Down <> 0) Then
11          .count_Down = (.count_Down - 1)

12          If (.count_Down > 0) Then
13              m = CStr(.count_Down) & "..."
14          Else
15              m = "�YA!"

16          End If

17          For i = 0 To 1
18              For j = 0 To 1
19                  H = .team_array(i).user_Index(j)

20                  If (H <> 0) Then
21                      If UserList(H).ConnID <> -1 Then
22                          Call Protocol.WriteConsoleMsg(H, m, FontTypeNames.FONTTYPE_GUILD)

23                          If (.count_Down = 0) Then Call Protocol.WritePauseToggle(H)

24                      End If

25                  End If

26              Next j
27          Next i

28      End If

29  End With

30  Exit Sub

loop_reto_index_Error:

31  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure loop_reto_index of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Function get_reto_index() As Integer

'
' @ amishar.-

    Dim LoopC As Long

1   On Error GoTo get_reto_index_Error

2   For LoopC = LBound(reto_List()) To UBound(reto_List())

3       If (reto_List(LoopC).used_ring = False) Then
4           get_reto_index = CInt(LoopC)

5           Exit Function

6       End If

7   Next LoopC

8   get_reto_index = -1

9   Exit Function

get_reto_index_Error:

10  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure get_reto_index of M�dulo Mod_Retos2vs2" & Erl & ".")

End Function

Public Sub set_reto_struct(ByVal user_Index As Integer, _
                           ByVal my_team As String, _
                           ByRef enemy_name As String, _
                           ByRef team_enemy As String, _
                           ByVal invDrop As Boolean, _
                           ByVal GoldAmount As Long)

'
' @ amishar.-

1   On Error GoTo set_reto_struct_Error

2   With UserList(user_Index).sReto
3       .accept_count = 0

4       With .tempStruct
5           .count_Down = 0
6           .used_ring = False

7           With .team_array(0)
8               .user_Index(0) = user_Index
9               .user_Index(1) = NameIndex(my_team)

10          End With

11          With .team_array(1)
12              .user_Index(0) = NameIndex(enemy_name)
13              .user_Index(1) = NameIndex(team_enemy)

14          End With

15          With .general_rules
16              .drop_inv = invDrop
17              .gold_gamble = GoldAmount

18          End With

19      End With

20  End With

21  Exit Sub

set_reto_struct_Error:

22  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure set_reto_struct of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Sub user_retoLoop(ByVal user_Index As Integer)

'
' @ amishar.-


1   On Error GoTo user_retoLoop_Error

2   With UserList(user_Index).sReto

3       If (.acceptLimit <> 0) Then
4           .acceptLimit = .acceptLimit - 1

5           If (.acceptLimit <= 0) Then
6               Call message_reto(.tempStruct, "El reto se ha autocancelado debido a que el tiempo para aceptar ha llegado a su l�mite.")

                Dim j As Long
                Dim i As Long
                Dim N As Integer
                Dim b As userStruct

7               For j = 0 To 1
8                   For i = 0 To 1
9                       N = .tempStruct.team_array(j).user_Index(i)

10                      If N > 0 Then
11                          If UCase$(UserList(N).sReto.nick_sender) = UCase$(UserList(user_Index).Name) Then
12                              UserList(N).sReto.nick_sender = vbNullString
13                              UserList(N).sReto.acceptedOK = False

14                          End If

15                      End If

16                  Next i
17              Next j

18              UserList(user_Index).sReto = b

19          End If

20      End If

21      If (.return_city <> 0) Then
22          .return_city = .return_city - 1

23          If (.return_city = 0) Then

                Dim p As WorldPos
                Dim k As WorldPos

24              p.Map = 1
25              p.X = 50
26              p.Y = 50

27              Call ClosestStablePos(p, k)
28              Call WarpUserChar(user_Index, p.Map, k.X, k.Y, True)

29              Call Protocol.WriteConsoleMsg(user_Index, "Regresas a la ciudad.", FontTypeNames.FONTTYPE_GUILD)

                Dim rIndex As Integer
30              rIndex = .reto_Index

31              .nick_sender = vbNullString
32              .reto_Index = 0

33              Call clear_data(rIndex)

34          End If

35      End If

36  End With

37  Exit Sub

user_retoLoop_Error:

38  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure user_retoLoop of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Sub erase_userData(ByVal user_Index As Integer)

'
' @ amishar.-


1   On Error GoTo erase_userData_Error

2   With UserList(user_Index).sReto

        Dim dumpStruct As retoStruct

3       .accept_count = 0
4       .nick_sender = vbNullString
5       .reto_Index = 0
6       .tmp_Time = 0
7       .reto_used = False
8       .tempStruct = dumpStruct
9       .return_city = 0
10      .acceptedOK = False

11  End With

12  Exit Sub

erase_userData_Error:

13  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure erase_userData of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Function can_send_reto(ByVal user_Index As Integer, _
                              ByRef fError As String) As Boolean

'
' @ amishar.-


1   On Error GoTo can_send_reto_Error

2   can_send_reto = False

3   With UserList(user_Index)

4       If (.flags.Muerto <> 0) Then
5           fError = "�Est�s muerto!"

6           Exit Function

7       End If

8       If (.Counters.Pena <> 0) Then
9           fError = "Est�s en la c�rcel"

10          Exit Function

11      End If

12      If (.Stats.GLD < .sReto.tempStruct.general_rules.gold_gamble) Then
13          fError = "No tienes el oro necesario"

14          Exit Function

15      End If

16      If (.Pos.Map <> eCiudad.cUllathorpe) Then
17          fError = .Name & " est� fuera de su hogar."

18          Exit Function

19      End If

20      If (.mReto.reto_Index <> 0) Or (.sReto.reto_used = True) Then
21          fError = .Name & " ya est� en reto."

22          Exit Function

23      End If

24      If (.Stats.ELV < 30) Then
25          fError = "Debes ser mayor a nivel 30!"

26          Exit Function

27      End If

28      With .sReto.tempStruct
29          can_send_reto = check_User(.team_array(0).user_Index(1), fError, .general_rules.gold_gamble)

30          If (can_send_reto) Then
31              can_send_reto = check_User(.team_array(1).user_Index(0), fError, .general_rules.gold_gamble)
32          Else

33              Exit Function

34          End If

35          If (can_send_reto) Then
36              can_send_reto = check_User(.team_array(1).user_Index(1), fError, .general_rules.gold_gamble)
37          Else

38              Exit Function

39          End If

40      End With

41  End With

42  Exit Function

can_send_reto_Error:

43  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure can_send_reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Function

Private Function check_User(ByVal user_Index As Integer, _
                            ByRef fError As String, _
                            ByVal goldGamble As Long) As Boolean

'
' @ amishar.-



1   On Error GoTo check_User_Error

2   check_User = False

3   If (user_Index = 0) Then
4       fError = "Alg�n usuario est� offline."

5       Exit Function

6   End If

7   With UserList(user_Index)

8       If (.flags.Muerto <> 0) Then
9           fError = .Name & " �Est� muerto!"

10          Exit Function

11      End If

12      If (.flags.Automatico <> 0) Then
13          fError = .Name & " �Esta en torneo!"
14          Exit Function
15      End If

16      If (.Counters.Pena <> 0) Then
17          fError = .Name & " Est� en la c�rcel"

18          Exit Function

19      End If

20      If (.Pos.Map <> eCiudad.cUllathorpe) Then
21          fError = .Name & " est� fuera de su hogar."

22          Exit Function

23      End If

24      If (.mReto.reto_Index <> 0) Or (.sReto.reto_used = True) Then
25          fError = .Name & " ya est� en reto."

26          Exit Function

27      End If

28      If (.Stats.GLD < goldGamble) Then
29          fError = .Name & " No tiene el oro necesario"

30          Exit Function

31      End If

32      If (.Stats.ELV < 30) Then
33          fError = .Name & " debe ser mayor a nivel 30!"

34          Exit Function

35      End If

36      check_User = True

37  End With

38  Exit Function

check_User_Error:

39  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure check_User of M�dulo Mod_Retos2vs2" & Erl & ".")

End Function

Public Sub send_Reto(ByVal user_Index As Integer)

'
' @ amishar.-



1   On Error GoTo send_Reto_Error

2   With UserList(user_Index).sReto

        Dim i As Long
        Dim j As Long

        Dim team_str As String
        Dim gamble_str As String

3       team_str = UserList(.tempStruct.team_array(0).user_Index(0)).Name & " y " & UserList(.tempStruct.team_array(0).user_Index(1)).Name & " vs " & UserList(.tempStruct.team_array(1).user_Index(0)).Name & " y " & UserList(.tempStruct.team_array(1).user_Index(1)).Name

4       gamble_str = " apostando " & Format$(.tempStruct.general_rules.gold_gamble, "#,###") & " monedas de oro"

5       If (.tempStruct.general_rules.drop_inv) Then
6           gamble_str = " y los items del inventario"

7       End If

8       For i = 0 To 1
9           For j = 0 To 1
10              UserList(.tempStruct.team_array(i).user_Index(j)).sReto.nick_sender = UCase$(UserList(user_Index).Name)

11              If (.tempStruct.team_array(i).user_Index(j) <> user_Index) Then
12                  Call Protocol.WriteConsoleMsg(.tempStruct.team_array(i).user_Index(j), "Solicitud de reto modalidad 2vs2 : " & team_str & " " & gamble_str & " para aceptar tipea /RETAR " & UCase$(UserList(user_Index).Name) & ".", FontTypeNames.FONTTYPE_GUILD)

13              End If

14          Next j
15      Next i

16      Call Protocol.WriteConsoleMsg(user_Index, "Se han enviado las solicitudes.", FontTypeNames.FONTTYPE_GUILD)
17      .acceptLimit = 60

18  End With

19  Exit Sub

send_Reto_Error:

20  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure send_Reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Sub disconnect_Reto(ByVal user_Index As Integer)

'
' @ amishar.-



    Dim team_Index As Integer
    Dim user_slot As Integer
    Dim team_winner As Byte
    Dim reto_Index As Integer

1   On Error GoTo disconnect_Reto_Error

2   reto_Index = UserList(user_Index).sReto.reto_Index

3   team_Index = find_Team(user_Index, reto_Index)

4   If (team_Index <> -1) Then
5       team_winner = IIf(team_Index = 1, 0, 1)
6       Call finish_reto(UserList(user_Index).sReto.reto_Index, team_winner)

7   End If

8   Exit Sub

disconnect_Reto_Error:

9   Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure disconnect_Reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Sub closeOtherReto(ByVal UserIndex As Integer)

'
' @ amishar.-


    Dim j As Long
    Dim i As Long
    Dim N As Integer
    Dim c As Boolean

1   On Error GoTo closeOtherReto_Error

2   N = NameIndex(UserList(UserIndex).sReto.nick_sender)

3   If (N > 0) Then

4       For i = 0 To 1
5           For j = 0 To 1

6               With UserList(N).sReto.tempStruct.team_array(i)

7                   If (.user_Index(j) = UserIndex) Then
8                       c = True

9                       Exit For

10                  End If

11              End With

12          Next j
13      Next i

14      If c = True Then

15          For i = 0 To 1
16              For j = 0 To 1

17                  With UserList(N).sReto.tempStruct.team_array(i)

18                      If (.user_Index(j) > 0) Then
19                          If UCase$(UserList(.user_Index(j)).sReto.nick_sender) = UCase$(UserList(N).Name) Then
20                              Call Protocol.WriteConsoleMsg(.user_Index(j), "El reto solicitado por " & UserList(N).Name & " ha sido cancelado debido a la desconexi�n de un participante.", FontTypeNames.FONTTYPE_GUILD)

21                          End If

22                      End If

23                  End With

24              Next j
25          Next i

26      End If

27  End If

28  Exit Sub

closeOtherReto_Error:

29  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure closeOtherReto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Sub accept_Reto(ByVal user_Index As Integer, ByVal requestName As String)

'
' @ amishar.-



    Dim sendIndex As Integer
    Dim i As Long
    Dim j As Long

1   On Error GoTo accept_Reto_Error

2   sendIndex = NameIndex(requestName)

3   If (sendIndex = 0) Or (UCase$(requestName) <> UserList(user_Index).sReto.nick_sender) Then
4       Call Protocol.WriteConsoleMsg(user_Index, requestName & " no te est� retando!!", FontTypeNames.FONTTYPE_GUILD)

5       Exit Sub

6   End If

7   If Not (UCase$(UserList(user_Index).Name) <> UserList(user_Index).sReto.nick_sender) Then
8       Call Protocol.WriteConsoleMsg(user_Index, "No te puedes aceptar a ti mismo", FontTypeNames.FONTTYPE_GUILD)

9       Exit Sub

10  End If

11  If (sendIndex = 0) Then Exit Sub

12  If UserList(user_Index).sReto.acceptedOK Then
13      Call Protocol.WriteConsoleMsg(user_Index, "�Ya has aceptado!", FontTypeNames.FONTTYPE_GUILD)

14      Exit Sub

15  End If

16  UserList(sendIndex).sReto.accept_count = (UserList(sendIndex).sReto.accept_count + 1)

17  Call message_reto(UserList(sendIndex).sReto.tempStruct, UserList(user_Index).Name & " acept� el reto.")

18  If (UserList(sendIndex).sReto.accept_count = 3) Then
19      Call init_reto(sendIndex)
20      Call message_reto(UserList(sendIndex).sReto.tempStruct, "Todos los participantes han aceptado el reto.")

21  End If

22  UserList(user_Index).sReto.acceptedOK = True

23  Exit Sub

accept_Reto_Error:

24  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure accept_Reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Private Sub init_reto(ByVal userSendIndex As Integer)

'
' @ amishar.-



    Dim reto_Index As Integer

1   On Error GoTo init_reto_Error

2   reto_Index = get_reto_index()

3   If (reto_Index = -1) Then
4       Call message_reto(UserList(userSendIndex).sReto.tempStruct, "Reto cancelado, todas las arenas est�n ocupadas.")

5       Exit Sub

6   End If

7   UserList(userSendIndex).sReto.acceptLimit = 0
8   reto_List(reto_Index) = UserList(userSendIndex).sReto.tempStruct
9   reto_List(reto_Index).used_ring = True
10  reto_List(reto_Index).count_Down = 6

11  Call warp_Teams(reto_Index)

12  Exit Sub

init_reto_Error:

13  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure init_reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Private Sub warp_Teams(ByVal reto_Index As Integer, _
                       Optional ByVal respawnUser As Boolean = False)

'
' @ amishar.-



1   On Error GoTo warp_Teams_Error

2   With reto_List(reto_Index)

        Dim LoopC As Long
        Dim mPosX As Byte
        Dim mPosY As Byte
        Dim nUser As Integer

3       .count_Down = 6

4       For LoopC = 0 To 1
5           nUser = .team_array(0).user_Index(LoopC)

6           If (nUser <> 0) Then
7               If (UserList(nUser).ConnID <> -1) Then
8                   mPosX = get_pos_x(reto_Index, 1, CInt(LoopC))
9                   mPosY = get_pos_y(reto_Index, 1, CInt(LoopC))

10                  UserList(nUser).sReto.reto_used = True
11                  UserList(nUser).sReto.reto_Index = reto_Index

12                  Call WarpUserChar(nUser, Mapa_Arenas, mPosX, mPosY, True)
13                  Call Protocol.WritePauseToggle(nUser)

14                  If (respawnUser) Then
15                      If (UserList(nUser).flags.Muerto) Then
16                          Call RevivirUsuario(nUser)

17                      End If

18                      UserList(nUser).Stats.MinHp = UserList(nUser).Stats.MaxHP
19                      UserList(nUser).Stats.MinMAN = UserList(nUser).Stats.MaxMAN
20                      UserList(nUser).Stats.MinHam = 100
21                      UserList(nUser).Stats.MinAGU = 100
22                      UserList(nUser).Stats.MinSta = UserList(nUser).Stats.MaxSta

23                      Call Protocol.WriteUpdateUserStats(nUser)

24                  End If

25              Else

26                  UserList(nUser).sReto.acceptedOK = False

27              End If

28          End If

29      Next LoopC

30      For LoopC = 0 To 1
31          nUser = .team_array(1).user_Index(LoopC)

32          If (nUser <> 0) Then
33              If (UserList(nUser).ConnID <> -1) Then
34                  mPosX = get_pos_x(reto_Index, 2, CInt(LoopC))
35                  mPosY = get_pos_y(reto_Index, 2, CInt(LoopC))

36                  UserList(nUser).sReto.reto_used = True
37                  UserList(nUser).sReto.reto_Index = reto_Index

38                  Call WarpUserChar(nUser, Mapa_Arenas, mPosX, mPosY, True)
39                  Call Protocol.WritePauseToggle(nUser)

40                  If (respawnUser) Then
41                      If (UserList(nUser).flags.Muerto) Then
42                          Call RevivirUsuario(nUser)

43                      End If

44                      UserList(nUser).Stats.MinHp = UserList(nUser).Stats.MaxHP
45                      UserList(nUser).Stats.MinMAN = UserList(nUser).Stats.MaxMAN
46                      UserList(nUser).Stats.MinHam = 100
47                      UserList(nUser).Stats.MinAGU = 100
48                      UserList(nUser).Stats.MinSta = UserList(nUser).Stats.MaxSta

49                      Call Protocol.WriteUpdateUserStats(nUser)

50                  End If

51              Else
52                  UserList(nUser).sReto.acceptedOK = False

53              End If

54          End If

55      Next LoopC

56  End With

57  Exit Sub

warp_Teams_Error:

58  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure warp_Teams of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Private Sub message_reto(ByRef retoStr As retoStruct, ByRef sMessage As String)

'
' @ amishar.-


1   On Error GoTo message_reto_Error

2   With retoStr


        Dim i As Long
        Dim j As Long
        Dim u As Integer

3       For i = 0 To 1
4           For j = 0 To 1
5               u = .team_array(i).user_Index(j)

6               If (u <> 0) Then
7                   If (UserList(u).ConnID <> -1) Then
8                       Call Protocol.WriteConsoleMsg(u, sMessage, FontTypeNames.FONTTYPE_GUILD)

9                   End If

10              End If

11          Next j
12      Next i

13  End With

14  Exit Sub

message_reto_Error:

15  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure message_reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Sub user_die_reto(ByVal user_Index As Integer)

'
' @ amishar.-


    Dim team_Index As Integer
    Dim user_slot As Integer
    Dim other_user As Integer
    Dim reto_Index As Integer

1   On Error GoTo user_die_reto_Error

2   reto_Index = UserList(user_Index).sReto.reto_Index

3   team_Index = find_Team(user_Index, reto_Index)

4   If (team_Index <> -1) Then
5       user_slot = find_user(team_Index, user_Index, reto_Index)
6   Else

7       Exit Sub

8   End If

9   If (user_slot = -1) Then Exit Sub

10  other_user = IIf(user_slot = 0, 1, 0)
11  other_user = reto_List(reto_Index).team_array(team_Index).user_Index(other_user)

    'is dead?

12  If (other_user) Then
13      If UserList(other_user).flags.Muerto Then
14          Call team_winner(reto_Index, IIf(team_Index = 0, 1, 0))

15      End If

16  Else
17      Call team_winner(reto_Index, IIf(team_Index = 0, 1, 0))

18  End If

19  Exit Sub

user_die_reto_Error:

20  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure user_die_reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Public Function find_Team(ByVal user_Index As Integer, _
                          ByVal reto_Index As Integer) As Integer

'
' @ amishar.-



    Dim i As Long
    Dim j As Long

1   On Error GoTo find_Team_Error

2   For i = 0 To 1
3       For j = 0 To 1

4           If reto_List(reto_Index).team_array(i).user_Index(j) = user_Index Then
5               find_Team = i

6               Exit Function

7           End If

8       Next j
9   Next i

10  find_Team = -1

11  Exit Function

find_Team_Error:

12  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure find_Team of M�dulo Mod_Retos2vs2" & Erl & ".")

End Function

Private Function find_user(ByVal team_Index As Integer, _
                           ByVal user_Index As Integer, _
                           ByVal reto_Index As Integer) As Integer

'
' @ amishar.-


    Dim i As Long

1   On Error GoTo find_user_Error

2   For i = 0 To 1

3       If reto_List(reto_Index).team_array(team_Index).user_Index(i) = user_Index Then
4           find_user = i

5           Exit Function

6       End If

7   Next i

8   find_user = -1

9   Exit Function

find_user_Error:

10  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure find_user of M�dulo Mod_Retos2vs2" & Erl & ".")

End Function

Private Sub team_winner(ByVal reto_Index As Integer, ByVal team_winner As Byte)

'
' @ amishar.-


1   On Error GoTo team_winner_Error

2   With reto_List(reto_Index)
3       .team_array(team_winner).round_count = (.team_array(team_winner).round_count + 1)

4       If (.team_array(team_winner).round_count = 2) Then
5           Call finish_reto(reto_Index, team_winner)
6       Else
7           Call respawn_reto(reto_Index, team_winner)

8       End If

9   End With

10  Exit Sub

team_winner_Error:

11  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure team_winner of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Private Sub respawn_reto(ByVal reto_Index As Integer, ByVal team_winner As Integer)

'
' @ amishar.-

'Call warp_Teams(reto_Index, True)



    Dim LoopX As Long
    Dim LoopC As Long
    Dim mStr As String
    Dim Index As Integer

1   On Error GoTo respawn_reto_Error

2   With reto_List(reto_Index)

3       mStr = "El equipo " & CStr(team_winner + 1) & " gana este duelo." & vbNewLine & "Resultado parcial : " & CStr(.team_array(0).round_count) & "-" & CStr(.team_array(1).round_count)

4       For LoopX = 0 To 1
5           For LoopC = 0 To 1
6               Index = .team_array(LoopX).user_Index(LoopC)

7               If (Index <> 0) Then
8                   If UserList(Index).ConnID <> -1 Then
9                       Call Protocol.WriteConsoleMsg(Index, mStr, FontTypeNames.FONTTYPE_GUILD)
10                      Call Protocol.WriteConsoleMsg(Index, "El siguiente round iniciar� en 3 segundos.", FontTypeNames.FONTTYPE_GUILD)

11                  End If

12              End If

13          Next LoopC
14      Next LoopX

15      .nextRoundCount = 3

16  End With

17  Exit Sub

respawn_reto_Error:

18  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure respawn_reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Private Sub finish_reto(ByVal reto_Index As Integer, ByVal team_winner As Byte)

'
' @ amishar.-



1   On Error GoTo finish_reto_Error

2   With reto_List(reto_Index)

        Dim retoMessage As String
        Dim team_looser As Byte
        Dim temp_index As Integer

3       retoMessage = get_reto_message(reto_Index)

4       retoMessage = retoMessage & vbNewLine & "Reto 2vs2> Ganador equipo " & CStr(team_winner + 1) & "."

5       Call SendData(SendTarget.ToAll, 0, Protocol.PrepareMessageConsoleMsg(retoMessage, FontTypeNames.FONTTYPE_GUILD))

6       team_looser = IIf(team_winner = 0, 1, 0)

        Dim LoopC As Long
        Dim bydrop As Boolean
        Dim byGold As Long

7       bydrop = (.general_rules.drop_inv = True)
8       byGold = .general_rules.gold_gamble

9       With .team_array(team_looser)

10          For LoopC = 0 To 1
11              temp_index = .user_Index(LoopC)

12              UserList(temp_index).sReto.reto_used = False
13              UserList(temp_index).sReto.acceptedOK = False

14              If (bydrop) Then
15                  Call TirarTodosLosItems(temp_index)

16              End If

17              Call WarpUserChar(temp_index, 1, 50 + LoopC, 50, True)

18              UserList(temp_index).Stats.GLD = (UserList(temp_index).Stats.GLD - byGold)

19              UserList(temp_index).sReto.nick_sender = vbNullString
20              UserList(temp_index).sReto.reto_Index = 0

21              Call Protocol.WriteUpdateGold(temp_index)

22          Next LoopC

23      End With

24      With .team_array(team_winner)

25          For LoopC = 0 To 1
26              temp_index = .user_Index(LoopC)

27              UserList(temp_index).sReto.reto_used = False
28              UserList(temp_index).sReto.acceptedOK = False

29              If (bydrop) Then
30                  UserList(temp_index).sReto.return_city = 15
31                  reto_List(reto_Index).haydrop = True

32                  Call Protocol.WriteConsoleMsg(temp_index, "Regresar�s a tu hogar en 15 segundos.", FontTypeNames.FONTTYPE_GUILD)
33              Else
34                  Call WarpUserChar(temp_index, 1, 50 + LoopC, 50, True)

35              End If

36              UserList(temp_index).Stats.GLD = (UserList(temp_index).Stats.GLD + byGold)

37              If reto_List(reto_Index).haydrop Then

38                  UserList(temp_index).sReto.nick_sender = vbNullString
39                  UserList(temp_index).sReto.reto_Index = 0

40              End If

41              Call Protocol.WriteUpdateGold(temp_index)

42          Next LoopC

43      End With

44      If .haydrop Then
45          Call clear_data(reto_Index)

46      End If

47  End With

48  Exit Sub

finish_reto_Error:

49  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure finish_reto of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Private Sub clear_data(ByVal reto_Index As Integer)

'
' @ amishar.-



1   On Error GoTo clear_data_Error

2   With reto_List(reto_Index)

3       .haydrop = False
4       .count_Down = 0

5       With .general_rules
6           .drop_inv = False
7           .gold_gamble = 0

8       End With

9       .used_ring = False

        Dim i As Long

10      For i = 0 To 1
11          .team_array(i).user_Index(0) = 0
12          .team_array(i).user_Index(1) = 0

13          .team_array(i).round_count = 0

14      Next i

15  End With

16  Exit Sub

clear_data_Error:

17  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure clear_data of M�dulo Mod_Retos2vs2" & Erl & ".")

End Sub

Private Function get_reto_message(ByVal reto_Index As Integer) As String

'
' @ amishar.-



    Dim TempStr As String
    Dim tempUser As Integer

1   On Error GoTo get_reto_message_Error

2   With reto_List(reto_Index)

3       TempStr = "Retos> "

4       With .team_array(0)
5           tempUser = .user_Index(0)

6           If (tempUser <> 0) Then
7               If UserList(tempUser).ConnID <> -1 Then
8                   TempStr = TempStr & UserList(tempUser).Name

9               End If

10          End If

11          tempUser = .user_Index(1)

12          If (tempUser <> 0) Then
13              If UserList(tempUser).ConnID <> -1 Then
14                  TempStr = TempStr & " y " & UserList(tempUser).Name

15              End If

16          End If

17      End With

18      With .team_array(1)
19          tempUser = .user_Index(0)

20          If (tempUser <> 0) Then
21              If UserList(tempUser).ConnID <> -1 Then
22                  TempStr = TempStr & " vs " & UserList(tempUser).Name

23              End If

24          End If

25          tempUser = .user_Index(1)

26          If (tempUser <> 0) Then
27              If UserList(tempUser).ConnID <> -1 Then
28                  TempStr = TempStr & " y " & UserList(tempUser).Name

29              End If

30          End If

31      End With

32      With .general_rules

33          TempStr = TempStr & " con apuesta de " & Format$(.gold_gamble, "#,###") & " monedas de oro"

34          If (.drop_inv) Then
35              TempStr = TempStr & " y los items del inventario"

36          End If

37      End With

38  End With

39  Exit Function

get_reto_message_Error:

40  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure get_reto_message of M�dulo Mod_Retos2vs2" & Erl & ".")

End Function

Private Function get_pos_x(ByVal ring_Index As Integer, _
                           ByVal team_Index As Integer, _
                           ByVal user_Index As Integer) As Integer

'
' @ amishar.-

    Dim endPos As Integer

    endPos = retoPos(ring_Index, team_Index, user_Index + 1).X

    get_pos_x = endPos

End Function

Private Function get_pos_y(ByVal ring_Index As Integer, _
                           ByVal team_Index As Integer, _
                           ByVal user_Index As Integer) As Integer

'
' @ amishar.-

    Dim endPos As Integer

    endPos = retoPos(ring_Index, team_Index, user_Index + 1).Y

    get_pos_y = endPos

End Function

Public Sub retos2vs2Load()

'
' @ amishar.-

    Dim nArenas As Integer

    Dim bReader As New clsIniManager

    bReader.Initialize DatPath & "Retos2vs2.ini"

    nArenas = val(bReader.GetValue("INIT", "Arenas"))

    If (nArenas = 0) Then Exit Sub

    ReDim Mod_Retos2vs2.retoPos(1 To nArenas, 1 To 2, 1 To 2) As Mod_Retos2vs2.retoPosStruct
    ReDim Mod_Retos2vs2.reto_List(1 To nArenas) As Mod_Retos2vs2.retoStruct

    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim s As String
    Dim Tmp As Integer

    Mapa_Arenas = Configuracion.Mapa2vs2

    Tmp = Asc("-")

    For i = 1 To nArenas
        For j = 1 To 2
            For p = 1 To 2
                s = bReader.GetValue("ARENA" & CStr(i), "Equipo" & CStr(j) & "Jugador" & CStr(p))

                Mod_Retos2vs2.retoPos(i, j, p).X = val(ReadField(2, s, Tmp))
                Mod_Retos2vs2.retoPos(i, j, p).Y = val(ReadField(3, s, Tmp))

            Next p
        Next j
    Next i

End Sub

Public Function eventAttack(ByVal attackerIndex As Integer, _
                            ByVal victimIndex As Integer) As Boolean

'
' @ amishar

    If UserList(attackerIndex).sReto.reto_used = True Then
        If Mod_Retos2vs2.can_Attack(attackerIndex, victimIndex) = False Then
            Call WriteConsoleMsg(attackerIndex, "No puedes atacar a tu compa�ero.", FontTypeNames.FONTTYPE_INFO)
            eventAttack = False

            Exit Function

        End If

    End If

    eventAttack = True

End Function



