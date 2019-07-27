Attribute VB_Name = "Acciones"

Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, _
           ByVal Map As Integer, _
           ByVal X As Integer, _
           ByVal Y As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim tempIndex As Integer
    
        On Error Resume Next

        '¿Rango Visión? (ToxicWaste)
        If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
                Exit Sub

        End If
    
        '¿Posicion valida?
        If InMapBounds(Map, X, Y) Then

                With UserList(UserIndex)

                        If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                                tempIndex = MapData(Map, X, Y).NpcIndex
                
                                'Set the target NPC
                                .flags.TargetNPC = tempIndex
                
                                If Npclist(tempIndex).Comercia = 1 Then

                                        '¿Esta el user muerto? Si es asi no puede comerciar
                                        If .flags.Muerto = 1 Then
                                                Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If
                    
                                        'Is it already in commerce mode??
                                        If .flags.Comerciando Then
                                                Exit Sub

                                        End If
                    
                                        If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                                                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If
                    
                                        'Iniciamos la rutina pa' comerciar.
                                        Call IniciarComercioNPC(UserIndex)
                
                                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then

                                        '¿Esta el user muerto? Si es asi no puede comerciar
                                        If .flags.Muerto = 1 Then
                                                Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If
                    
                                        'Is it already in commerce mode??
                                        If .flags.Comerciando Then
                                                Exit Sub

                                        End If
                    
                                        If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                                                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If
                    
                                        'A depositar de una
                                        Call IniciarDeposito(UserIndex)
                
                                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then

                                        If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                                                Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If
                    
                                        'Revivimos si es necesario
                                        If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
                                                Call RevivirUsuario(UserIndex)

                                        End If
                    
                                        If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                                                'curamos totalmente
                                                .Stats.MinHp = .Stats.MaxHP
                                                Call WriteUpdateUserStats(UserIndex)

                                        End If

                                End If
                
                                '¿Es un obj?
                        ElseIf MapData(Map, X, Y).ObjInfo.objIndex > 0 Then
                                tempIndex = MapData(Map, X, Y).ObjInfo.objIndex
                
                                .flags.TargetObj = tempIndex
                
                                Select Case ObjData(tempIndex).OBJType

                                        Case eOBJType.otPuertas 'Es una puerta
                                                Call AccionParaPuerta(Map, X, Y, UserIndex)

                                        Case eOBJType.otCarteles 'Es un cartel
                                                Call AccionParaCartel(Map, X, Y, UserIndex)

                                        Case eOBJType.otForos 'Foro
                                                Call AccionParaForo(Map, X, Y, UserIndex)

                                        Case eOBJType.otLeña    'Leña

                                                If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                                                        Call AccionParaRamita(Map, X, Y, UserIndex)

                                                End If

                                End Select

                                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
                        ElseIf MapData(Map, X + 1, Y).ObjInfo.objIndex > 0 Then
                                tempIndex = MapData(Map, X + 1, Y).ObjInfo.objIndex
                                .flags.TargetObj = tempIndex
                
                                Select Case ObjData(tempIndex).OBJType
                    
                                        Case eOBJType.otPuertas 'Es una puerta
                                                Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
                    
                                End Select
            
                        ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.objIndex > 0 Then
                                tempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.objIndex
                                .flags.TargetObj = tempIndex
        
                                Select Case ObjData(tempIndex).OBJType

                                        Case eOBJType.otPuertas 'Es una puerta
                                                Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)

                                End Select
            
                        ElseIf MapData(Map, X, Y + 1).ObjInfo.objIndex > 0 Then
                                tempIndex = MapData(Map, X, Y + 1).ObjInfo.objIndex
                                .flags.TargetObj = tempIndex
                
                                Select Case ObjData(tempIndex).OBJType

                                        Case eOBJType.otPuertas 'Es una puerta
                                                Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

                                End Select

                        End If

                End With

        End If

End Sub

Public Sub AccionParaForo(ByVal Map As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer, _
                          ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 02/01/2010
        '02/01/2010: ZaMa - Agrego foros faccionarios
        '***************************************************

        On Error Resume Next

        Dim Pos As WorldPos
    
        Pos.Map = Map
        Pos.X = X
        Pos.Y = Y
    
        If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

        End If
    
        If SendPosts(UserIndex, ObjData(MapData(Map, X, Y).ObjInfo.objIndex).ForoID) Then
                Call WriteShowForumForm(UserIndex)

        End If
    
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
                If ObjData(MapData(Map, X, Y).ObjInfo.objIndex).Llave = 0 Then
                        If ObjData(MapData(Map, X, Y).ObjInfo.objIndex).Cerrada = 1 Then

                                'Abre la puerta
                                If ObjData(MapData(Map, X, Y).ObjInfo.objIndex).Llave = 0 Then
                    
                                        MapData(Map, X, Y).ObjInfo.objIndex = ObjData(MapData(Map, X, Y).ObjInfo.objIndex).IndexAbierta
                    
                                        Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.objIndex).GrhIndex, X, Y))
                    
                                        'Desbloquea
                                        MapData(Map, X, Y).Blocked = 0
                                        MapData(Map, X - 1, Y).Blocked = 0
                    
                                        'Bloquea todos los mapas
                                        Call Bloquear(True, Map, X, Y, 0)
                                        Call Bloquear(True, Map, X - 1, Y, 0)
                      
                                        'Sonido
                                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                                Else
                                        Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

                                End If

                        Else
                                'Cierra puerta
                                MapData(Map, X, Y).ObjInfo.objIndex = ObjData(MapData(Map, X, Y).ObjInfo.objIndex).IndexCerrada
                
                                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.objIndex).GrhIndex, X, Y))
                                
                                MapData(Map, X, Y).Blocked = 1
                                MapData(Map, X - 1, Y).Blocked = 1
                
                                Call Bloquear(True, Map, X - 1, Y, 1)
                                Call Bloquear(True, Map, X, Y, 1)
                
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))

                        End If
        
                        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.objIndex
                Else
                        Call WriteConsoleMsg(UserIndex, "La puerta está cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

                End If

        Else
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

        End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        If ObjData(MapData(Map, X, Y).ObjInfo.objIndex).OBJType = 8 Then
  
                If Len(ObjData(MapData(Map, X, Y).ObjInfo.objIndex).texto) > 0 Then
                        Call WriteShowSignal(UserIndex, MapData(Map, X, Y).ObjInfo.objIndex)

                End If
  
        End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim Suerte             As Byte

        Dim exito              As Byte

        Dim Obj                As Obj

        Dim SkillSupervivencia As Byte

        Dim Pos                As WorldPos

        Pos.Map = Map
        Pos.X = X
        Pos.Y = Y

        With UserList(UserIndex)

                If Distancia(Pos, .Pos) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                End If
    
                If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
                        Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                End If
    
                SkillSupervivencia = .Stats.UserSkills(eSkill.Supervivencia)
    
                If SkillSupervivencia < 6 Then
                        Suerte = 3
        
                ElseIf SkillSupervivencia <= 10 Then
                        Suerte = 2
        
                Else
                        Suerte = 1

                End If
    
                exito = RandomNumber(1, Suerte)
    
                If exito = 1 Then
                        If MapInfo(.Pos.Map).Zona <> eTerrain.terrain_ciudad Then
                                Obj.objIndex = FOGATA
                                Obj.Amount = 1
            
                                Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
            
                                Call MakeObj(Obj, Map, X, Y)
                                
                                'Las fogatas prendidas se deben eliminar
                                Call aLimpiarMundo.AddItem(Map, X, Y)
            
                                Call SubirSkill(UserIndex, eSkill.Supervivencia)
                        Else
                                Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                        End If

                Else
                        Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
                        Call SubirSkill(UserIndex, eSkill.Supervivencia)

                End If

        End With

End Sub
