Attribute VB_Name = "m_GranPoder"
Option Explicit

Public GranPoder As Integer


Public Sub OtorgarFavordelosDioses(ByVal UserIndex As Integer)
   
        Dim LoopC As Integer
        Dim EncontroIdeal As Byte
       
        If LastUser = 0 Then Exit Sub 'Debug.Print "FAVOR> LASTUSER=0": Exit Sub
       
        If GranPoder > 0 Then Exit Sub ' Ya hay
        
        If UserIndex = 0 Then
            
            Do While EncontroIdeal = 0 And LoopC < LastUser
                LoopC = LoopC + 1
                UserIndex = RandomNumber(1, LastUser)
                With UserList(UserIndex)
                
               ' Debug.Print "FAVOR> PRUEBA A " & .Name
                If .flags.UserLogged = True And .clase <> eClass.Worker And .Death = False And .flags.Automatico = False And .mReto.reto_Index = 0 And .sReto.reto_Index = 0 And .flags.EnConsulta = False And .flags.Muerto = 0 And .flags.Privilegios = User Then
                    If MapInfo(.Pos.Map).Pk = True And Not MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
                        EncontroIdeal = 1
                        GranPoder = UserIndex
                      '  Debug.Print "FAVOR> IDEAL= " & .Name
                    End If
                    
                End If
                End With
            Loop
            If Not EncontroIdeal = 1 Then
                UserIndex = 0
                GranPoder = 0
                Exit Sub
            End If
        End If
        
        If EncontroIdeal > 0 Then
         With UserList(UserIndex)
            If .flags.Muerto = 0 And .Death = False And .flags.Automatico = False And MapInfo(.Pos.Map).Pk = True Then
                'Call OtorgarFavordelosDioses(0)
                GranPoder = UserIndex
                
                Dim NombreMap As String

                If Len(MapInfo(UserList(UserIndex).Pos.Map).Name) > 0 Then
                    NombreMap = MapInfo(UserList(UserIndex).Pos.Map).Name
                Else
                    NombreMap = CStr(UserList(UserIndex).Pos.Map)
                End If
                
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Favor de los dioses>", "Los dioses le otorgan el gran poder a " & .Name & " en el mapa " & NombreMap & ".", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO))
                'RefreshCharStatus GranPoder
                
                Call WarpUserChar(UserIndex, .Pos.Map, .Pos.X, .Pos.Y, False, False)
                
            End If
    End With
        End If
End Sub


Public Sub DarPoder()
    Static Minutoss As Integer
    Minutoss = Minutoss + 1

    If Minutoss >= 1 Then
        If GranPoder = 0 Then
            OtorgarFavordelosDioses (0)
        ElseIf GranPoder > 0 Then

            If RandomNumber(1, 5) = 5 Then
                Dim NombreMap As String

                If Len(MapInfo(UserList(GranPoder).Pos.Map).Name) > 0 Then
                    NombreMap = MapInfo(UserList(GranPoder).Pos.Map).Name
                Else
                    NombreMap = CStr(UserList(GranPoder).Pos.Map)
                End If

                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Favor de los Dioses>", UserList(GranPoder).Name & " tiene el poder en el mapa " & NombreMap & ".", FontTypeNames.FONTTYPE_AMARILLO, FontTypeNames.FONTTYPE_BLANCO))
            End If

        End If
    End If

    Minutoss = 0


End Sub
