Attribute VB_Name = "m_Viajero"

Option Explicit

' @@ Npc informador
Public iNpcViajero As Integer

Public Function MeterNPCViajero()

' @@ CuiCui
    Dim SpawnPos As WorldPos, newSpawnPos As WorldPos, UI As Integer, NpcIndex As Integer

    ' @@ Pos donde caiga.
1   On Error GoTo MeterNPCViajero_Error

2   SpawnPos.Map = 128: SpawnPos.X = 73: SpawnPos.Y = 50

3   UI = MapData(SpawnPos.Map, SpawnPos.X, SpawnPos.Y).Userindex

    ' @@ Si hay un user lo movemos
4   If UI > 0 Then

5       newSpawnPos = SpawnPos
6       newSpawnPos.Y = SpawnPos.Y + 1

7       Call ClosestStablePos(newSpawnPos, newSpawnPos)
8       Call WarpUserChar(UI, 1, newSpawnPos.X, newSpawnPos.Y, 0, 0)

9   End If

    ' @@ Verifico si no habia un npc en esa pos xd
10  NpcIndex = MapData(SpawnPos.Map, SpawnPos.X, SpawnPos.Y).NpcIndex

    ' @@ Si habia un npc lo quito xd
11  If NpcIndex > 0 Then Call QuitarNPC(NpcIndex)

    ' @@ Spawneo al NPC
12  iNpcViajero = SpawnNpc(745, SpawnPos, True, False)

29  Exit Function

MeterNPCViajero_Error:

30  Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure MeterNPCViajero of M�dulo m_viajero" & Erl & ".")

End Function

Public Function QuitarNPCViajero()

' @@ Cuicui

    If iNpcViajero > 0 Then

        On Error GoTo 10    ' CODIGO MISTICO DE CUICUI

        ' @@ Sacamos el NPC
        Call QuitarNPC(iNpcViajero)

10      iNpcViajero = 0

    End If
    
    If Err.Number > 0 Then LogError "Error en quitarNPC " & Err.description

End Function

