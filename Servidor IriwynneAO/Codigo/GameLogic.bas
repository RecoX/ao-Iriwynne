Attribute VB_Name = "Extra"
Option Explicit

Function ItemShop(ByVal ItemIndex As Integer) As Boolean

        ' @@ Miqueas
        ' @@ 08/12/15
        ' @@ Identificamos los items SHOP del resto

        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
        ItemShop = ObjData(ItemIndex).Shop = 1

End Function

Public Sub DeleteAreaResuTheNpc(ByVal iNpc As Integer)
 
        ' @@ Miqueas
        ' @@ 17-10-2015
        ' @@ Delete Trigger in this NPC area
        Const Range = 4 ' @@ + 4 Tildes a la redonda del npc
 
        Dim X      As Long
 
        Dim Y      As Long
     
        Dim NpcPos As WorldPos
     
        NpcPos.Map = Npclist(iNpc).pos.Map
        NpcPos.X = Npclist(iNpc).pos.X
        NpcPos.Y = Npclist(iNpc).pos.Y

        For X = NpcPos.X - Range To NpcPos.X + Range
                For Y = NpcPos.Y - Range To NpcPos.Y + Range
 
                        If InMapBounds(NpcPos.Map, X, Y) Then
                                If (MapData(NpcPos.Map, X, Y).trigger = eTrigger.AutoResu) Then
                                        MapData(NpcPos.Map, X, Y).trigger = eTrigger.NADA

                                End If

                        End If
 
                Next Y
        Next X
 
End Sub

Public Sub SetAreaResuTheNpc(ByVal iNpc As Integer)
 
        ' @@ Miqueas
        ' @@ 17-10-2015
        ' @@ Set Trigger in this NPC area
        Const Range = 4 ' @@ + 4 Tildes a la redonda del npc
 
        Dim X      As Long
 
        Dim Y      As Long
     
        Dim NpcPos As WorldPos
     
        NpcPos.Map = Npclist(iNpc).pos.Map
        NpcPos.X = Npclist(iNpc).pos.X
        NpcPos.Y = Npclist(iNpc).pos.Y

        For X = NpcPos.X - Range To NpcPos.X + Range
                For Y = NpcPos.Y - Range To NpcPos.Y + Range
 
                        If InMapBounds(NpcPos.Map, X, Y) Then
                                If (MapData(NpcPos.Map, X, Y).trigger <> eTrigger.AutoResu) Or (MapData(NpcPos.Map, X, Y).trigger = eTrigger.NADA) Then
                                        MapData(NpcPos.Map, X, Y).trigger = eTrigger.AutoResu

                                End If

                        End If
 
                Next Y
        Next X
 
End Sub
 
Public Function IsAreaResu(ByVal Userindex As Integer) As Boolean
 
        ' @@ Miqueas
        ' @@ 17/10/2015
        ' @@ Validate Trigger Area
        With UserList(Userindex)
 
                If MapData(.pos.Map, .pos.X, .pos.Y).trigger = eTrigger.AutoResu Then
                        IsAreaResu = True
 
                        Exit Function
 
                End If
 
        End With
 
        IsAreaResu = False

End Function
 
Public Sub AutoCurar(ByVal Userindex As Integer)
 
        ' @@ Miqueas
        ' @@ 17-10-15
        ' @@ Zona de auto curacion
     
        With UserList(Userindex)
 
                If .flags.Muerto = 1 Then
                        Call RevivirUsuario(Userindex)
                        Call WriteConsoleMsg(Userindex, "El sacerdote te ha resucitado", FontTypeNames.FONTTYPE_INFO)
                        GoTo temp

                End If
 
                If .Stats.MinHp < .Stats.MaxHP Then
                        .Stats.MinHp = .Stats.MaxHP
                        Call WriteUpdateHP(Userindex)
                        Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado.", FontTypeNames.FONTTYPE_INFO)

                End If
 
temp:
 
                If .flags.Ceguera = 1 Then
                        .flags.Ceguera = 0

                End If

                If .flags.Envenenado = 1 Then
                        .flags.Envenenado = 0

                End If
 
        End With
 
End Sub
 
Public Function isNPCResucitador(ByVal iNpc As Integer) As Boolean
 
        ' @@ Miqueas
        ' @@ 17/10/2015
        ' @@ Validate NPC
        With Npclist(iNpc)
 
                If (.NPCtype = eNPCType.Revividor) Or (.NPCtype = eNPCType.ResucitadorNewbie) Then
                        isNPCResucitador = True
 
                        Exit Function
 
                End If
 
        End With
 
        isNPCResucitador = False

End Function

Public Function MapaLimpieza(ByVal Mapa As Integer) As Boolean

        Dim LoopC As Long

        For LoopC = 1 To Configuracion.MapNoLimpiezaCant

                If (Configuracion.MapNoLimpieza(LoopC) = Mapa) Then
                        MapaLimpieza = False
                        Exit Function

                End If

        Next LoopC

        MapaLimpieza = True
        
End Function

Public Function EsNoCreable(ByVal objIndex As Integer) As Boolean

        '***************************************************
        'Author: Miqueas
        'Last Modification: 18/11/2015
        '
        '***************************************************
        If objIndex < 1 Or objIndex > UBound(ObjData) Then Exit Function
    
        EsNoCreable = (ObjData(objIndex).ObjNoCreable = 1)

End Function

Public Function EsNewbie(ByVal Userindex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        EsNewbie = UserList(Userindex).Stats.ELV <= LimiteNewbie

End Function

Public Function esArmada(ByVal Userindex As Integer) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 23/01/2007
        '***************************************************

        esArmada = (UserList(Userindex).Faccion.ArmadaReal = 1)

End Function

Public Function esCaos(ByVal Userindex As Integer) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 23/01/2007
        '***************************************************

        esCaos = (UserList(Userindex).Faccion.FuerzasCaos = 1)

End Function

Public Function esGM(ByVal Userindex As Integer) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 23/01/2007
        '***************************************************

        esGM = (UserList(Userindex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))

End Function

Public Sub DoTileEvents(ByVal Userindex As Integer, _
                        ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 06/03/2010
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
' 06/03/2010 : Now we have 5 attemps to not fall into a map change or another teleport while going into a teleport. (Marco)
'***************************************************

    Dim npos   As WorldPos

    Dim FxFlag As Boolean

    Dim TelepRadio As Integer

    Dim DestPos As WorldPos

    On Error GoTo errhandler

    'Controla las salidas
    If InMapBounds(Map, X, Y) Then

        With MapData(Map, X, Y)

            If .ObjInfo.objIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.objIndex).OBJType = eOBJType.otTeleport
                TelepRadio = ObjData(.ObjInfo.objIndex).Radio

            End If

            If .TileExit.Map > 0 And .TileExit.Map <= NumMaps Then


                ' Es un teleport, entra en una posicion random, acorde al radio (si es 0, es pos fija)
                ' We have 5 attempts to not falling into another teleport or a map exit.. If we get to the fifth attemp,
                ' the teleport will act as if its radius = 0.
                If FxFlag And TelepRadio > 0 Then


                    Dim attemps As Long

                    Dim exitMap As Boolean

                    Do
                        DestPos.X = .TileExit.X + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.Y = .TileExit.Y + RandomNumber(TelepRadio * (-1), TelepRadio)

                        attemps = attemps + 1

                        exitMap = MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map > 0 And _
                                  MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map <= NumMaps
                    Loop Until (attemps >= 5 Or exitMap = False)

                    If attemps >= 5 Then
                        DestPos.X = .TileExit.X
                        DestPos.Y = .TileExit.Y

                    End If

                    ' Posicion fija
                Else
                    DestPos.X = .TileExit.X
                    DestPos.Y = .TileExit.Y

                End If

                DestPos.Map = .TileExit.Map

                If esGM(Userindex) Then
                    Call LogGM(UserList(Userindex).Name, "Utiliz� un teleport hacia el mapa " & _
                                                         DestPos.Map & " (" & DestPos.X & "," & DestPos.Y & ")")

                End If

                ' Si es un mapa que no admite muertos
                If MapInfo(DestPos.Map).OnDeathGoTo.Map <> 0 Then

                    ' Si esta muerto no puede entrar
                    If UserList(Userindex).flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "S�lo se permite entrar al mapa a los personajes vivos.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).pos, npos)

                        If npos.X <> 0 And npos.Y <> 0 Then
                            Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                        End If

                        Exit Sub

                    End If

                End If
                
                '�Es mapa de newbies?
                If MapInfo(DestPos.Map).Restringir = eRestrict.restrict_newbie Then

                    '�El usuario es un newbie?
                    If EsNewbie(Userindex) Or esGM(Userindex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, npos)

                            If npos.X <> 0 And npos.Y <> 0 Then
                                Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                            End If

                        End If

                    Else    'No es newbie
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).pos, npos)

                        If npos.X <> 0 And npos.Y <> 0 Then
                            Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, False)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_armada Then    '�Es mapa de Armadas?

                    '�El usuario es Armada?
                    If esArmada(Userindex) Or esGM(Userindex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, npos)

                            If npos.X <> 0 And npos.Y <> 0 Then
                                Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                            End If

                        End If

                    Else    'No es armada
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para miembros del ej�rcito real.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).pos, npos)

                        If npos.X <> 0 And npos.Y <> 0 Then
                            Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_caos Then    '�Es mapa de Caos?

                    '�El usuario es Caos?
                    If esCaos(Userindex) Or esGM(Userindex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, npos)

                            If npos.X <> 0 And npos.Y <> 0 Then
                                Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                            End If

                        End If

                    Else    'No es caos
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para miembros de la legi�n oscura.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).pos, npos)

                        If npos.X <> 0 And npos.Y <> 0 Then
                            Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_faccion Then    '�Es mapa de faccionarios?

                    '�El usuario es Armada o Caos?
                    If esArmada(Userindex) Or esCaos(Userindex) Or esGM(Userindex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, npos)

                            If npos.X <> 0 And npos.Y <> 0 Then
                                Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                            End If

                        End If

                    Else    'No es Faccionario
                        Call WriteConsoleMsg(Userindex, "Solo se permite entrar al mapa si eres miembro de alguna facci�n.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).pos, npos)

                        If npos.X <> 0 And npos.Y <> 0 Then
                            Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                        End If

                    End If

                Else    'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.

                    
                    
                If DestPos.Map = 122 And UserList(Userindex).Stats.ELV < 40 And Not esGM(Userindex) Then
                    Call WriteConsoleMsg(Userindex, "Debes ser nivel 40 o superior para ingresar a Magma.", FontTypeNames.FONTTYPE_INFO)
                    Call ClosestStablePos(UserList(Userindex).pos, npos)

                    If npos.X <> 0 And npos.Y <> 0 Then
                        Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, False)

                    End If
                    
                    Exit Sub
                    
                End If
                
                    If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then

                        If MapInfo(MapData(Map, X, Y).TileExit.Map).Pk = False And Userindex = GranPoder Then
                            GranPoder = 0
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Favor de los dioses>", UserList(Userindex).Name & " ha perdido el poder.", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO))
                            Call WarpUserChar(Userindex, UserList(Userindex).pos.Map, UserList(Userindex).pos.X, UserList(Userindex).pos.Y, False, False)

                        End If

                        Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                    Else

                        Call ClosestLegalPos(DestPos, npos)

                        If npos.X <> 0 And npos.Y <> 0 Then

                            Call WarpUserChar(Userindex, npos.Map, npos.X, npos.Y, FxFlag)

                            If MapInfo(MapData(Map, X, Y).TileExit.Map).Pk = False And Userindex = GranPoder Then
                                GranPoder = 0
                                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Favor de los dioses>", "Los dioses le otorgan el gran poder a " & UserList(Userindex).Name & " en el mapa " & UserList(Userindex).pos.Map & ".", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO))
                                'RefreshCharStatus GranPoder

                                Call WarpUserChar(Userindex, UserList(Userindex).pos.Map, UserList(Userindex).pos.X, UserList(Userindex).pos.Y, False, True)

                            End If


                        End If

                    End If

                End If

                'Te fusite del mapa. La criatura ya no es m�s tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer

                aN = UserList(Userindex).flags.AtacadoPorNpc

                If aN > 0 Then
                    Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                    Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                    Npclist(aN).flags.AttackedBy = vbNullString

                End If

                aN = UserList(Userindex).flags.NPCAtacado

                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(Userindex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString

                    End If

                End If

                UserList(Userindex).flags.AtacadoPorNpc = 0
                UserList(Userindex).flags.NPCAtacado = 0

            End If

        End With

    End If



    Exit Sub

errhandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.description)

End Sub

Function InRangoVision(ByVal Userindex As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If X > UserList(Userindex).pos.X - MinXBorder And X < UserList(Userindex).pos.X + MinXBorder Then
                If Y > UserList(Userindex).pos.Y - MinYBorder And Y < UserList(Userindex).pos.Y + MinYBorder Then
                        InRangoVision = True
                        Exit Function

                End If

        End If

        InRangoVision = False

End Function

Public Function InVisionRangeAndMap(ByVal Userindex As Integer, _
                                    ByRef OtherUserPos As WorldPos) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 20/11/2010
        '
        '***************************************************
    
        With UserList(Userindex)
        
                ' Same map?
                If .pos.Map <> OtherUserPos.Map Then Exit Function
    
                ' In x range?
                If OtherUserPos.X < .pos.X - MinXBorder Or OtherUserPos.X > .pos.X + MinXBorder Then Exit Function
        
                ' In y range?
                If OtherUserPos.Y < .pos.Y - MinYBorder And OtherUserPos.Y > .pos.Y + MinYBorder Then Exit Function

        End With

        InVisionRangeAndMap = True
    
End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, _
                          X As Integer, _
                          Y As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If X > Npclist(NpcIndex).pos.X - MinXBorder And X < Npclist(NpcIndex).pos.X + MinXBorder Then
                If Y > Npclist(NpcIndex).pos.Y - MinYBorder And Y < Npclist(NpcIndex).pos.Y + MinYBorder Then
                        InRangoVisionNPC = True
                        Exit Function

                End If

        End If

        InRangoVisionNPC = False

End Function

Function InMapBounds(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                InMapBounds = False
        Else
                InMapBounds = True

        End If
    
End Function

Private Function RhombLegalPos(ByRef pos As WorldPos, _
                               ByRef vX As Long, _
                               ByRef vY As Long, _
                               ByVal Distance As Long, _
                               Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False) As Boolean
        '***************************************************
        'Author: Marco Vanotti (Marco)
        'Last Modification: -
        ' walks all the perimeter of a rhomb of side  "distance + 1",
        ' which starts at Pos.x - Distance and Pos.y
        '***************************************************

        Dim i As Long
    
        vX = pos.X - Distance
        vY = pos.Y
    
        For i = 0 To Distance - 1

                If (LegalPos(pos.Map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
                        vX = vX + i
                        vY = vY - i
                        RhombLegalPos = True
                        Exit Function

                End If

        Next
    
        vX = pos.X
        vY = pos.Y - Distance
    
        For i = 0 To Distance - 1

                If (LegalPos(pos.Map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
                        vX = vX + i
                        vY = vY + i
                        RhombLegalPos = True
                        Exit Function

                End If

        Next
    
        vX = pos.X + Distance
        vY = pos.Y
    
        For i = 0 To Distance - 1

                If (LegalPos(pos.Map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
                        vX = vX - i
                        vY = vY + i
                        RhombLegalPos = True
                        Exit Function

                End If

        Next
    
        vX = pos.X
        vY = pos.Y + Distance
    
        For i = 0 To Distance - 1

                If (LegalPos(pos.Map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
                        vX = vX - i
                        vY = vY - i
                        RhombLegalPos = True
                        Exit Function

                End If

        Next
    
        RhombLegalPos = False
    
End Function

Public Function RhombLegalTilePos(ByRef pos As WorldPos, _
                                  ByRef vX As Long, _
                                  ByRef vY As Long, _
                                  ByVal Distance As Long, _
                                  ByVal objIndex As Integer, _
                                  ByVal ObjAmount As Long, _
                                  ByVal PuedeAgua As Boolean, _
                                  ByVal PuedeTierra As Boolean) As Boolean

        '***************************************************
        'Author: ZaMa
        'Last Modification: -
        ' walks all the perimeter of a rhomb of side  "distance + 1",
        ' which starts at Pos.x - Distance and Pos.y
        ' and searchs for a valid position to drop items
        '***************************************************
        On Error GoTo errhandler

        Dim i           As Long

        Dim HayObj      As Boolean
    
        Dim X           As Integer

        Dim Y           As Integer

        Dim MapObjIndex As Integer
    
        vX = pos.X - Distance
        vY = pos.Y
    
        For i = 0 To Distance - 1
        
                X = vX + i
                Y = vY - i
        
                If (LegalPos(pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
                        ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
                        If Not HayObjeto(pos.Map, X, Y, objIndex, ObjAmount) Then
                                vX = X
                                vY = Y
                
                                RhombLegalTilePos = True
                                Exit Function

                        End If
            
                End If

        Next
    
        vX = pos.X
        vY = pos.Y - Distance
    
        For i = 0 To Distance - 1
        
                X = vX + i
                Y = vY + i
        
                If (LegalPos(pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
                        ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
                        If Not HayObjeto(pos.Map, X, Y, objIndex, ObjAmount) Then
                                vX = X
                                vY = Y
                
                                RhombLegalTilePos = True
                                Exit Function

                        End If

                End If

        Next
    
        vX = pos.X + Distance
        vY = pos.Y
    
        For i = 0 To Distance - 1
        
                X = vX - i
                Y = vY + i
    
                If (LegalPos(pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
        
                        ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
                        If Not HayObjeto(pos.Map, X, Y, objIndex, ObjAmount) Then
                                vX = X
                                vY = Y
                
                                RhombLegalTilePos = True
                                Exit Function

                        End If

                End If

        Next
    
        vX = pos.X
        vY = pos.Y + Distance
    
        For i = 0 To Distance - 1
        
                X = vX - i
                Y = vY - i
    
                If (LegalPos(pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then

                        ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
                        If Not HayObjeto(pos.Map, X, Y, objIndex, ObjAmount) Then
                                vX = X
                                vY = Y
                
                                RhombLegalTilePos = True
                                Exit Function

                        End If

                End If

        Next
    
        RhombLegalTilePos = False
    
        Exit Function
    
errhandler:
        Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.description)

End Function

Public Function HayObjeto(ByVal Mapa As Integer, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal objIndex As Integer, _
                          ByVal ObjAmount As Long) As Boolean

        '***************************************************
        'Author: ZaMa
        'Last Modification: -
        'Checks if there's space in a tile to add an itemAmount
        '***************************************************
        Dim MapObjIndex As Integer

        MapObjIndex = MapData(Mapa, X, Y).ObjInfo.objIndex
            
        ' Hay un objeto tirado?
        If MapObjIndex <> 0 Then

                ' Es el mismo objeto?
                If MapObjIndex = objIndex Then
                        ' La suma es menor a 10k?
                        HayObjeto = (MapData(Mapa, X, Y).ObjInfo.Amount + ObjAmount > MAX_INVENTORY_OBJS)
                Else
                        HayObjeto = True

                End If

        Else
                HayObjeto = False

        End If

End Function

Sub ClosestLegalPos(pos As WorldPos, _
                    ByRef npos As WorldPos, _
                    Optional PuedeAgua As Boolean = False, _
                    Optional PuedeTierra As Boolean = True, _
                    Optional ByVal CheckExitTile As Boolean = False)
        '*****************************************************************
        'Author: Unknown (original version)
        'Last Modification: 09/14/2010 (Marco)
        'History:
        ' - 01/24/2007 (ToxicWaste)
        'Encuentra la posicion legal mas cercana y la guarda en nPos
        '*****************************************************************

        Dim Found As Boolean

        Dim LoopC As Integer

        Dim tX    As Long

        Dim tY    As Long
    
        npos = pos
        tX = pos.X
        tY = pos.Y
    
        LoopC = 1
    
        ' La primera posicion es valida?
        If LegalPos(pos.Map, npos.X, npos.Y, PuedeAgua, PuedeTierra, CheckExitTile) Then
                Found = True
    
                ' Busca en las demas posiciones, en forma de "rombo"
        Else

                While (Not Found) And LoopC <= 12

                        If RhombLegalPos(pos, tX, tY, LoopC, PuedeAgua, PuedeTierra, CheckExitTile) Then
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

End Sub

Public Sub ClosestStablePos(pos As WorldPos, ByRef npos As WorldPos)
        '***************************************************
        'Author: Unknown
        'Last Modification: 09/14/2010
        'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
        '*****************************************************************

        Call ClosestLegalPos(pos, npos, , , True)

End Sub

Function NameIndex(ByVal Name As String) As Integer
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim Userindex As Long
    
        '�Nombre valido?
        If LenB(Name) = 0 Then
                NameIndex = 0
                Exit Function

        End If
    
        If InStrB(Name, "+") <> 0 Then
                Name = UCase$(Replace(Name, "+", " "))

        End If
    
        Userindex = 1

        Do Until UCase$(UserList(Userindex).Name) = UCase$(Name)
        
                Userindex = Userindex + 1
        
                If Userindex > MaxUsers Then
                        NameIndex = 0
                        Exit Function

                End If

        Loop
     
        NameIndex = Userindex

End Function

Function CheckForSameIP(ByVal Userindex As Integer, ByVal UserIP As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Long
    
        For LoopC = 1 To LastUser

                If UserList(LoopC).flags.UserLogged = True Then
                        If UserList(LoopC).Ip = UserIP And Userindex <> LoopC Then
                                CheckForSameIP = True
                                Exit Function

                        End If

                End If

        Next LoopC
    
        CheckForSameIP = False

End Function

Function CheckForSameName(ByVal Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        'Controlo que no existan usuarios con el mismo nombre
        Dim LoopC As Long
    
        For LoopC = 1 To LastUser

                If UserList(LoopC).flags.UserLogged Then
            
                        'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
                        'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
                        'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
                        'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
                        'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
                        If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                                CheckForSameName = True
                                Exit Function

                        End If

                End If

        Next LoopC
    
        CheckForSameName = False

End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef pos As WorldPos)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        'Toma una posicion y se mueve hacia donde esta perfilado
        '*****************************************************************

        Select Case Head

                Case eHeading.NORTH
                        pos.Y = pos.Y - 1
        
                Case eHeading.SOUTH
                        pos.Y = pos.Y + 1
        
                Case eHeading.EAST
                        pos.X = pos.X + 1
        
                Case eHeading.WEST
                        pos.X = pos.X - 1

        End Select

End Sub

Function LegalPos(ByVal Map As Integer, _
                  ByVal X As Integer, _
                  ByVal Y As Integer, _
                  Optional ByVal PuedeAgua As Boolean = False, _
                  Optional ByVal PuedeTierra As Boolean = True, _
                  Optional ByVal CheckExitTile As Boolean = False) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Checks if the position is Legal.
        '***************************************************

        '�Es un mapa valido?
        If (Map <= 0 Or Map > NumMaps) Or _
           (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
                LegalPos = False
        Else

                With MapData(Map, X, Y)

                        If PuedeAgua And PuedeTierra Then
                                LegalPos = (.Blocked <> 1) And _
                                   (.Userindex = 0) And _
                                   (.NpcIndex = 0)
                        ElseIf PuedeTierra And Not PuedeAgua Then
                                LegalPos = (.Blocked <> 1) And _
                                   (.Userindex = 0) And _
                                   (.NpcIndex = 0) And _
                                   (Not HayAgua(Map, X, Y))
                        ElseIf PuedeAgua And Not PuedeTierra Then
                                LegalPos = (.Blocked <> 1) And _
                                   (.Userindex = 0) And _
                                   (.NpcIndex = 0) And _
                                   (HayAgua(Map, X, Y))
                        Else
                                LegalPos = False

                        End If

                End With
        
                If CheckExitTile Then
                        LegalPos = LegalPos And (MapData(Map, X, Y).TileExit.Map = 0)

                End If
        
        End If

End Function

Function MoveToLegalPos(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        Optional ByVal PuedeAgua As Boolean = False, _
                        Optional ByVal PuedeTierra As Boolean = True) As Boolean
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 13/07/2009
        'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
        '13/07/2009: ZaMa - Now it's also legal move where an invisible admin is.
        '***************************************************

        Dim Userindex        As Integer

        Dim IsDeadChar       As Boolean

        Dim IsAdminInvisible As Boolean

        '�Es un mapa valido?
        If (Map <= 0 Or Map > NumMaps) Or _
           (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
                MoveToLegalPos = False
        Else

                With MapData(Map, X, Y)
                        Userindex = .Userindex
        
                        If Userindex > 0 Then
                                IsDeadChar = (UserList(Userindex).flags.Muerto = 1)
                                IsAdminInvisible = (UserList(Userindex).flags.AdminInvisible = 1)
                        Else
                                IsDeadChar = False
                                IsAdminInvisible = False

                        End If
        
                        If PuedeAgua And PuedeTierra Then
                                MoveToLegalPos = (.Blocked <> 1) And _
                                   (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                                   (.NpcIndex = 0)
                        ElseIf PuedeTierra And Not PuedeAgua Then
                                MoveToLegalPos = (.Blocked <> 1) And _
                                   (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                                   (.NpcIndex = 0) And _
                                   (Not HayAgua(Map, X, Y))
                        ElseIf PuedeAgua And Not PuedeTierra Then
                                MoveToLegalPos = (.Blocked <> 1) And _
                                   (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                                   (.NpcIndex = 0) And _
                                   (HayAgua(Map, X, Y))
                        Else
                                MoveToLegalPos = False

                        End If

                End With

        End If

End Function

Public Sub FindLegalPos(ByVal Userindex As Integer, _
                        ByVal Map As Integer, _
                        ByRef X As Integer, _
                        ByRef Y As Integer)
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 26/03/2009
        'Search for a Legal pos for the user who is being teleported.
        '***************************************************

        If MapData(Map, X, Y).Userindex <> 0 Or _
           MapData(Map, X, Y).NpcIndex <> 0 Then
                    
                ' Se teletransporta a la misma pos a la que estaba
                If MapData(Map, X, Y).Userindex = Userindex Then Exit Sub
                            
                Dim FoundPlace     As Boolean

                Dim tX             As Long

                Dim tY             As Long

                Dim Rango          As Long

                Dim OtherUserIndex As Integer
    
                For Rango = 1 To 5
                        For tY = Y - Rango To Y + Rango
                                For tX = X - Rango To X + Rango

                                        'Reviso que no haya User ni NPC
                                        If MapData(Map, tX, tY).Userindex = 0 And _
                                           MapData(Map, tX, tY).NpcIndex = 0 Then
                        
                                                If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                                                Exit For

                                        End If

                                Next tX
        
                                If FoundPlace Then _
                                   Exit For
                        Next tY
            
                        If FoundPlace Then _
                           Exit For
                Next Rango
    
                If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                        X = tX
                        Y = tY
                Else
                        'Muy poco probable, pero..
                        'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
                        OtherUserIndex = MapData(Map, X, Y).Userindex

                        If OtherUserIndex <> 0 Then

                                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then

                                        'Le avisamos al que estaba comerciando que se tuvo que ir.
                                        If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                                                Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                                                Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                                                Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)

                                        End If

                                        'Lo sacamos.
                                        If UserList(OtherUserIndex).flags.UserLogged Then
                                                Call FinComerciarUsu(OtherUserIndex)
                                                Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor recon�ctate...")
                                                Call FlushBuffer(OtherUserIndex)

                                        End If

                                End If
            
                                Call CloseSocket(OtherUserIndex)

                        End If

                End If

        End If

End Sub

Function LegalPosNPC(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal AguaValida As Byte, _
                     Optional ByVal IsPet As Boolean = False) As Boolean

        '***************************************************
        'Autor: Unkwnown
        'Last Modification: 09/23/2009
        'Checks if it's a Legal pos for the npc to move to.
        '09/23/2009: Pato - If UserIndex is a AdminInvisible, then is a legal pos.
        '***************************************************
        
        On Error GoTo errhandler
        
        Dim IsDeadChar       As Boolean

        Dim Userindex        As Integer

        Dim IsAdminInvisible As Boolean
    
        If (Map <= 0 Or Map > NumMaps) Or _
           (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
                LegalPosNPC = False
                Exit Function

        End If

        With MapData(Map, X, Y)
                Userindex = .Userindex

                If Userindex > 0 Then
                        IsDeadChar = UserList(Userindex).flags.Muerto = 1
                        IsAdminInvisible = (UserList(Userindex).flags.AdminInvisible = 1)
                Else
                        IsDeadChar = False
                        IsAdminInvisible = False

                End If
    
                If AguaValida = 0 Then
                        LegalPosNPC = (.TileExit.Map <= 0) And (.Blocked <> 1) And _
                           (.Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                           (.NpcIndex = 0) And _
                           (.trigger <> eTrigger.POSINVALIDA Or IsPet) _
                           And Not HayAgua(Map, X, Y)
                Else
                        LegalPosNPC = (.TileExit.Map <= 0) And (.Blocked <> 1) And _
                           (.Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                           (.NpcIndex = 0) And _
                           (.trigger <> eTrigger.POSINVALIDA Or IsPet)

                End If

        End With

errhandler:

End Function

Sub SendHelp(ByVal Index As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim NumHelpLines As Integer

        Dim LoopC        As Integer

        NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

        For LoopC = 1 To NumHelpLines
                Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
        Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If Npclist(NpcIndex).NroExpresiones > 0 Then

                Dim randomi

                randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))

        End If

End Sub

Sub LookatTile(ByVal Userindex As Integer, _
               ByVal Map As Integer, _
               ByVal X As Integer, _
               ByVal Y As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/03/2009
'13/02/2009: ZaMa - El nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
'07/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'***************************************************

    On Error GoTo errhandler

    'Responde al click del usuario sobre el mapa
    Dim FoundChar As Byte

    Dim FoundSomething As Byte

    Dim TempCharIndex As Integer

    Dim Stat   As String

    Dim ft     As FontTypeNames

    With UserList(Userindex)

        '�Rango Visi�n? (ToxicWaste)
        If (Abs(.pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.pos.X - X) > RANGO_VISION_X) Then
            Exit Sub

        End If

        '�Posicion valida?
        If InMapBounds(Map, X, Y) Then

            With .flags
                .TargetMap = Map
                .TargetX = X
                .TargetY = Y

                '�Es un obj?
                If MapData(Map, X, Y).ObjInfo.objIndex > 0 Then
                    'Informa el nombre
                    .TargetObjMap = Map
                    .TargetObjX = X
                    .TargetObjY = Y
                    FoundSomething = 1
                ElseIf MapData(Map, X + 1, Y).ObjInfo.objIndex > 0 Then

                    'Informa el nombre
                    If ObjData(MapData(Map, X + 1, Y).ObjInfo.objIndex).OBJType = eOBJType.otPuertas Then
                        .TargetObjMap = Map
                        .TargetObjX = X + 1
                        .TargetObjY = Y
                        FoundSomething = 1

                    End If

                ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.objIndex > 0 Then

                    If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.objIndex).OBJType = eOBJType.otPuertas Then
                        'Informa el nombre
                        .TargetObjMap = Map
                        .TargetObjX = X + 1
                        .TargetObjY = Y + 1
                        FoundSomething = 1

                    End If

                ElseIf MapData(Map, X, Y + 1).ObjInfo.objIndex > 0 Then

                    If ObjData(MapData(Map, X, Y + 1).ObjInfo.objIndex).OBJType = eOBJType.otPuertas Then
                        'Informa el nombre
                        .TargetObjMap = Map
                        .TargetObjX = X
                        .TargetObjY = Y + 1
                        FoundSomething = 1

                    End If

                End If

                If FoundSomething = 1 Then
                    .TargetObj = MapData(Map, .TargetObjX, .TargetObjY).ObjInfo.objIndex

                    If (.TargetObj <> 0) Then    ' @@ Miqueas : Parchesuli

                        Dim TmpInt As Integer

                        TmpInt = MapData(Map, X, Y).ObjInfo.objIndex

                        If MostrarCantidad(.TargetObj) Then
                            Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name & " - " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount, FontTypeNames.FONTTYPE_INFO)

                        ElseIf (TmpInt <> 0) Then

                            If ObjData(TmpInt).OBJType = eOBJType.otTeleport Then

                                Dim DestinyMap As Integer

                                Dim MapName As String

                                DestinyMap = MapData(.TargetObjMap, .TargetObjX, .TargetObjY).TileExit.Map

                                If DestinyMap <> 0 Then
                                    MapName = MapInfo(DestinyMap).Name
                                Else
                                    MapName = vbNullString

                                End If

                                If Len(MapName) = 0 Then
                                    Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name & " a Mapa Desconocido", FontTypeNames.FONTTYPE_INFO)
                                Else
                                    Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name & " a " & MapName, FontTypeNames.FONTTYPE_INFO)

                                End If

                            End If

                        Else

                            Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name, FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End If

                '�Es un personaje?
                If Y + 1 <= YMaxMapSize Then
                    If MapData(Map, X, Y + 1).Userindex > 0 Then
                        TempCharIndex = MapData(Map, X, Y + 1).Userindex
                        FoundChar = 1

                    End If

                    If MapData(Map, X, Y + 1).NpcIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
                        FoundChar = 2

                    End If

                End If

                '�Es un personaje?
                If FoundChar = 0 Then
                    If MapData(Map, X, Y).Userindex > 0 Then
                        TempCharIndex = MapData(Map, X, Y).Userindex
                        FoundChar = 1

                    End If

                    If MapData(Map, X, Y).NpcIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y).NpcIndex
                        FoundChar = 2

                    End If

                End If

            End With

            'Reaccion al personaje
            If FoundChar = 1 Then    '  �Encontro un Usuario?
                If UserList(TempCharIndex).flags.AdminInvisible = 0 Or .flags.Privilegios And PlayerType.Dios Then

                    With UserList(TempCharIndex)

                        If LenB(.DescRM) = 0 And .showName Then    'No tiene descRM y quiere que se vea su nombre.
                            If EsNewbie(TempCharIndex) Then
                                Stat = " <NEWBIE>"

                            End If

                            If .Faccion.ArmadaReal = 1 Then
                                Stat = Stat & " <Ej�rcito Real> " & "<" & TituloReal(TempCharIndex) & ">"
                            ElseIf .Faccion.FuerzasCaos = 1 Then
                                Stat = Stat & " <Legi�n Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"

                            End If

                            If .GuildIndex > 0 Then
                                Stat = Stat & " <" & modGuilds.GuildName(.GuildIndex) & ">"

                            End If

                            If Len(.Desc) > 0 Then
                                Stat = "Ves a " & .Name & Stat & " - " & .Desc
                            Else
                                Stat = "Ves a " & .Name & Stat

                            End If

                            Stat = Stat & IIf(.flags.Vip > 0, " <VIP>", "")

                            Stat = Stat & " <Frags: " & .Faccion.CiudadanosMatados + .Faccion.CriminalesMatados & ">"

                            Stat = Stat & " <Puntos: " & .flags.PuntosShop & ">"

                            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                                Stat = Stat & " [CONSEJO DE BANDERBILL]"
                                ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                                Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
                                ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                            Else

                                If Not .flags.Privilegios And PlayerType.User Then

                                    ' @@ Miqueas 07/11/15 - Ordeno un poco el if de aca abajo

                                    'Stat = Stat & " <GAME MASTER>"
                                    ' Elijo el color segun el rango del GM:
                                    ' admin
                                    If .flags.Privilegios = PlayerType.Admin Then
                                        ft = FontTypeNames.FONTTYPE_VERDE
                                        Stat = Stat & " <Administrador>"
                                        ' Dios
                                    ElseIf .flags.Privilegios = PlayerType.Dios Then
                                        ft = FontTypeNames.FONTTYPE_DIOS
                                        Stat = Stat & " <Dios>"
                                        ' Gm
                                    ElseIf .flags.Privilegios = PlayerType.SemiDios Then
                                        ft = FontTypeNames.FONTTYPE_GM
                                        Stat = Stat & " <Semi Dios>"
                                        ' Conse
                                    ElseIf .flags.Privilegios = PlayerType.Consejero Then
                                        ft = FontTypeNames.FONTTYPE_CONSE
                                        Stat = Stat & " <Consejero>"
                                        ' Rm o Dsrm
                                    ElseIf .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Consejero) Or .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Dios) Then
                                        ft = FontTypeNames.FONTTYPE_EJECUCION
                                        Stat = Stat & " <Rol Master>"
                                    Else
                                        ft = FontTypeNames.FONTTYPE_WARNING
                                        Stat = Stat & " <Error, Informar a la administracion>"

                                    End If

                                ElseIf criminal(TempCharIndex) Then
                                    Stat = Stat & " <CRIMINAL>"
                                    ft = FontTypeNames.FONTTYPE_FIGHT
                                Else
                                    Stat = Stat & " <CIUDADANO>"
                                    ft = FontTypeNames.FONTTYPE_CITIZEN

                                End If

                            End If

                        Else  'Si tiene descRM la muestro siempre.
                            Stat = .DescRM
                            ft = FontTypeNames.FONTTYPE_INFOBOLD

                        End If

                        If UserList(TempCharIndex).flags.Vip > 0 Then
                            ft = FontTypeNames.FONTTYPE_VIP
                        End If
                        If TempCharIndex = GranPoder Then
                            Stat = Stat & " [Bendecido por los Dioses]"
                            ft = FontTypeNames.FONTTYPE_GRANPODER
                        End If
                        
                        If .Death Then
                            Stat = "Ves a un jugador"
                            ft = FontTypeNames.FONTTYPE_AMARILLO
                        End If
                    
                    End With

                    If LenB(Stat) > 0 Then
                        Call WriteConsoleMsg(Userindex, Stat, ft)
                    End If

                    FoundSomething = 1
                    .flags.TargetUser = TempCharIndex
                    .flags.TargetNPC = 0
                    .flags.TargetNpcTipo = eNPCType.Comun

                End If

            End If

            With .flags

                If FoundChar = 2 Then    '�Encontro un NPC?

                    Dim estatus As String

                    Dim MinHp As Long

                    Dim MaxHP As Long

                    'Dim SupervivenciaSkill As Byte

                    Dim sDesc As String

                    MinHp = Npclist(TempCharIndex).Stats.MinHp
                    MaxHP = Npclist(TempCharIndex).Stats.MaxHP
                    'SupervivenciaSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia)

                    If .Muerto = 0 Then    ' @@ Miqueas : Vida siempre visible
                        estatus = "(" & MinHp & "/" & MaxHP & ") "

                    End If

                    If Len(Npclist(TempCharIndex).Desc) > 1 Then
                        Stat = Npclist(TempCharIndex).Desc

                        '�Es el rey o el demonio?
                        If Npclist(TempCharIndex).NPCtype = eNPCType.Noble Then
                            If Npclist(TempCharIndex).flags.Faccion = 0 Then    'Es el Rey.

                                'Si es de la Legi�n Oscura mostramos el mensaje correspondiente y lo ejecutamos:
                                If UserList(Userindex).Faccion.FuerzasCaos = 1 Then
                                    Stat = MENSAJE_REY_CAOS

                                    If Not esGM(Userindex) Then    ' @@ Miqueas : 03/12/15 - No explota Gm's
                                        Call UserDie(Userindex)

                                    End If

                                ElseIf criminal(Userindex) Then

                                    'Nos fijamos si es criminal enlistable o no enlistable:
                                    If UserList(Userindex).Faccion.CiudadanosMatados > 0 Or _
                                       UserList(Userindex).Faccion.Reenlistadas > 4 Then    'Es criminal no enlistable.
                                        Stat = MENSAJE_REY_CRIMINAL_NOENLISTABLE
                                    Else    'Es criminal enlistable.
                                        Stat = MENSAJE_REY_CRIMINAL_ENLISTABLE

                                    End If

                                End If

                            Else    'Es el demonio

                                'Si es de la Armada Real mostramos el mensaje correspondiente y lo ejecutamos:
                                If UserList(Userindex).Faccion.ArmadaReal = 1 Then
                                    Stat = MENSAJE_DEMONIO_REAL

                                    If Not esGM(Userindex) Then    ' @@ Miqueas : 03/12/15 - No explota Gm's
                                        Call UserDie(Userindex)

                                    End If

                                ElseIf Not criminal(Userindex) Then

                                    'Nos fijamos si es ciudadano enlistable o no enlistable:
                                    If UserList(Userindex).Faccion.RecibioExpInicialReal = 1 Or _
                                       UserList(Userindex).Faccion.Reenlistadas > 4 Then    'Es ciudadano no enlistable.
                                        Stat = MENSAJE_DEMONIO_CIUDADANO_NOENLISTABLE
                                    Else    'Es ciudadano enlistable.
                                        Stat = MENSAJE_DEMONIO_CIUDADANO_ENLISTABLE

                                    End If

                                End If

                            End If

                        End If

                        'Enviamos el mensaje propiamente dicho:
                        Call WriteChatOverHead(Userindex, Stat, Npclist(TempCharIndex).Char.CharIndex, vbWhite)

                        If Not UserList(Userindex).flags.LastNPCTalk = 0 And Not UserList(Userindex).flags.LastNPCTalk = Npclist(TempCharIndex).Char.CharIndex Then
                            Call WriteChatOverHead(Userindex, vbNullString, UserList(Userindex).flags.LastNPCTalk, vbWhite)

                        End If

                        UserList(Userindex).flags.LastNPCTalk = Npclist(TempCharIndex).Char.CharIndex

                    Else

                        Dim CentinelaIndex As Integer

                        CentinelaIndex = EsCentinela(TempCharIndex)

                        If CentinelaIndex <> 0 Then
                            'Enviamos nuevamente el texto del centinela seg�n quien pregunta
                            Call modCentinela.CentinelaSendClave(Userindex, CentinelaIndex)
                        Else

                            If Npclist(TempCharIndex).MaestroUser > 0 Then
                                Call WriteConsoleMsg(Userindex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
                            Else
                                sDesc = estatus & Npclist(TempCharIndex).Name

                                If Npclist(TempCharIndex).Owner > 0 Then sDesc = sDesc & " le pertenece a " & UserList(Npclist(TempCharIndex).Owner).Name
                                sDesc = sDesc & "."

                                Call WriteConsoleMsg(Userindex, sDesc, FontTypeNames.FONTTYPE_INFO)

                                If .Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                                    Call WriteConsoleMsg(Userindex, "Le peg� primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)

                                End If

                            End If

                        End If

                    End If

                    FoundSomething = 1
                    .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
                    .TargetNPC = TempCharIndex
                    .TargetUser = 0
                    .TargetObj = 0

                End If

                If FoundChar = 0 Then
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0

                End If

                '*** NO ENCOTRO NADA ***
                If FoundSomething = 0 Then
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                    .TargetObj = 0
                    .TargetObjMap = 0
                    .TargetObjX = 0
                    .TargetObjY = 0
                    Call WriteMultiMessage(Userindex, eMessages.DontSeeAnything)

                End If

            End With

        Else

            If FoundSomething = 0 Then

                With .flags
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                    .TargetObj = 0
                    .TargetObjMap = 0
                    .TargetObjX = 0
                    .TargetObjY = 0

                End With

                Call WriteMultiMessage(Userindex, eMessages.DontSeeAnything)

            End If

        End If

    End With

    Exit Sub

errhandler:
    Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.description)

End Sub

Function FindDirection(pos As WorldPos, Target As WorldPos) As eHeading
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        'Devuelve la direccion en la cual el target se encuentra
        'desde pos, 0 si la direc es igual
        '*****************************************************************

        Dim X As Integer

        Dim Y As Integer
    
        X = pos.X - Target.X
        Y = pos.Y - Target.Y
    
        'NE
        If Sgn(X) = -1 And Sgn(Y) = 1 Then
                FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
                Exit Function

        End If
    
        'NW
        If Sgn(X) = 1 And Sgn(Y) = 1 Then
                FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
                Exit Function

        End If
    
        'SW
        If Sgn(X) = 1 And Sgn(Y) = -1 Then
                FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
                Exit Function

        End If
    
        'SE
        If Sgn(X) = -1 And Sgn(Y) = -1 Then
                FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
                Exit Function

        End If
    
        'Sur
        If Sgn(X) = 0 And Sgn(Y) = -1 Then
                FindDirection = eHeading.SOUTH
                Exit Function

        End If
    
        'norte
        If Sgn(X) = 0 And Sgn(Y) = 1 Then
                FindDirection = eHeading.NORTH
                Exit Function

        End If
    
        'oeste
        If Sgn(X) = 1 And Sgn(Y) = 0 Then
                FindDirection = eHeading.WEST
                Exit Function

        End If
    
        'este
        If Sgn(X) = -1 And Sgn(Y) = 0 Then
                FindDirection = eHeading.EAST
                Exit Function

        End If
    
        'misma
        If Sgn(X) = 0 And Sgn(Y) = 0 Then
                FindDirection = 0
                Exit Function

        End If

End Function

Public Function ItemNoEsDeMapa(ByVal Index As Integer, _
                               ByVal bIsExit As Boolean) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        With ObjData(Index)
                ItemNoEsDeMapa = .OBJType <> eOBJType.otPuertas And _
                   .OBJType <> eOBJType.otForos And _
                   .OBJType <> eOBJType.otCarteles And _
                   .OBJType <> eOBJType.otArboles And _
                   .OBJType <> eOBJType.otYacimiento And _
                   Not (.OBJType = eOBJType.otTeleport And bIsExit)
    
        End With

End Function

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        With ObjData(Index)
                MostrarCantidad = .OBJType <> eOBJType.otPuertas And _
                   .OBJType <> eOBJType.otForos And _
                   .OBJType <> eOBJType.otCarteles And _
                   .OBJType <> eOBJType.otArboles And _
                   .OBJType <> eOBJType.otYacimiento And _
                   .OBJType <> eOBJType.otTeleport

        End With

End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        EsObjetoFijo = OBJType = eOBJType.otForos Or _
           OBJType = eOBJType.otCarteles Or _
           OBJType = eOBJType.otArboles Or _
           OBJType = eOBJType.otYacimiento

End Function

Public Function RestrictStringToByte(ByRef restrict As String) As Byte
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 04/18/2011
        '
        '***************************************************
        restrict = UCase$(restrict)

        Select Case restrict

                Case "NEWBIE"
                        RestrictStringToByte = 1
        
                Case "ARMADA"
                        RestrictStringToByte = 2
        
                Case "CAOS"
                        RestrictStringToByte = 3
        
                Case "FACCION"
                        RestrictStringToByte = 4
        
                Case Else
                        RestrictStringToByte = 0

        End Select

End Function

Public Function RestrictByteToString(ByVal restrict As Byte) As String

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 04/18/2011
        '
        '***************************************************
        Select Case restrict

                Case 1
                        RestrictByteToString = "NEWBIE"
        
                Case 2
                        RestrictByteToString = "ARMADA"
        
                Case 3
                        RestrictByteToString = "CAOS"
        
                Case 4
                        RestrictByteToString = "FACCION"
        
                Case 0
                        RestrictByteToString = "NO"

        End Select

End Function

Public Function TerrainStringToByte(ByRef restrict As String) As Byte
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 04/18/2011
        '
        '***************************************************
        restrict = UCase$(restrict)

        Select Case restrict

                Case "NIEVE"
                        TerrainStringToByte = 1
        
                Case "DESIERTO"
                        TerrainStringToByte = 2
        
                Case "CIUDAD"
                        TerrainStringToByte = 3
        
                Case "CAMPO"
                        TerrainStringToByte = 4
        
                Case "DUNGEON"
                        TerrainStringToByte = 5
        
                Case Else
                        TerrainStringToByte = 0

        End Select

End Function

Public Function TerrainByteToString(ByVal restrict As Byte) As String

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 04/18/2011
        '
        '***************************************************
        Select Case restrict

                Case 1
                        TerrainByteToString = "NIEVE"
        
                Case 2
                        TerrainByteToString = "DESIERTO"
        
                Case 3
                        TerrainByteToString = "CIUDAD"
        
                Case 4
                        TerrainByteToString = "CAMPO"
        
                Case 5
                        TerrainByteToString = "DUNGEON"
        
                Case 0
                        TerrainByteToString = "BOSQUE"

        End Select

End Function
