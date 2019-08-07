Attribute VB_Name = "Admin"

Option Explicit

Public Type tMotd
        texto As String
        Formato As String
End Type

Public MaxLines As Integer
Public MOTD()   As tMotd

Public Type tAPuestas
        Ganancias As Long
        Perdidas As Long
        Jugadas As Long
End Type

Public Apuestas                          As tAPuestas
Public tInicioServer                     As Long

'INTERVALOS
Public SanaIntervaloSinDescansar         As Integer
Public StaminaIntervaloSinDescansar      As Integer
Public SanaIntervaloDescansar            As Integer
Public StaminaIntervaloDescansar         As Integer
Public IntervaloSed                      As Integer
Public IntervaloHambre                   As Integer
Public IntervaloVeneno                   As Integer
Public IntervaloParalizado               As Integer

Public Const IntervaloParalizadoReducido As Integer = 37

Public IntervaloInvisible                As Integer

Public IntervaloFrio                     As Integer

Public IntervaloWavFx                    As Integer

Public IntervaloLanzaHechizo             As Integer

Public IntervaloNPCPuedeAtacar           As Integer

Public IntervaloNPCAI                    As Integer

Public IntervaloInvocacion               As Integer

Public IntervaloOculto                   As Integer '[Nacho]

Public IntervaloUserPuedeAtacar          As Long

Public IntervaloGolpeUsar                As Long

Public IntervaloMagiaGolpe               As Long

Public IntervaloGolpeMagia               As Long

Public IntervaloUserPuedeCastear         As Long

Public IntervaloUserPuedeTrabajar        As Long

Public IntervaloParaConexion             As Long

Public IntervaloCerrarConexion           As Long '[Gonzalo]

Public IntervaloUserPuedeUsar            As Long

Public IntervaloFlechasCazadores         As Long

Public IntervaloPuedeSerAtacado          As Long

Public IntervaloAtacable                 As Long

Public IntervaloOwnedNpc                 As Long

'BALANCE

Public PorcentajeRecuperoMana            As Integer

Public MinutosWs                         As Long

Public MinutosGuardarUsuarios            As Long

Public Puerto                            As Integer

Public BootDelBackUp                     As Byte

Public Lloviendo                         As Boolean

Public DeNoche                           As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        VersionOK = (Ver = ULTIMAVERSION)

End Function


Public Function BanHD_rem(ByVal HD As String) As Boolean
' GSZ-AO - Remueve un SerialHD como baneado

    On Error Resume Next

    Dim N As Long

    N = BanHD_find(HD)    ' buscar
    If N > 0 Then
        BanHDs.Remove N    ' quitar
        BanHD_save    ' guardar los cambios
        BanHD_rem = True
    Else
        BanHD_rem = False
    End If

End Function

Public Sub BanHD_add(ByVal HD As String)
' GSZ-AO - Agrega un nuevo SerialHD como baneado

    Dim N As Long

    N = BanHD_find(HD)    ' buscar

    If N > 0 Then
        ' ya estaba
    Else
        BanHDs.Add HD    ' agregar
        Call BanHD_save    ' guardar los cambios
    End If

End Sub

Public Function BanHD_find(ByVal HD As String) As Long
' GSZ-AO - Busca si un SerialHD est� baneado

    Dim Dale As Boolean
    Dim LoopC As Long

    Dale = True
    LoopC = 1
    On Error GoTo s
    Do While LoopC <= BanHDs.Count And Dale
        Dale = (BanHDs.Item(LoopC) <> HD)
        LoopC = LoopC + 1
    Loop
s:
    If Dale Then
        BanHD_find = 0
    Else
        BanHD_find = LoopC - 1
    End If

End Function

Public Sub BanHD_save()
' GSZ-AO - Guarda el listado de SerialHD's baneados
    On Error Resume Next
    Dim ArchivoBanHD As String
    Dim ArchN As Long
    Dim LoopC As Long

    ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"

    ArchN = FreeFile()
    Open ArchivoBanHD For Output As #ArchN

    For LoopC = 1 To BanHDs.Count
        Print #ArchN, BanHDs.Item(LoopC)
    Next LoopC

    Close #ArchN

End Sub

Public Sub BanHD_load()
' GSZ-AO - Carga el listado de SerialHD's baneados
    On Error Resume Next

    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanHD As String

    ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"

    Set BanHDs = New Collection

    ArchN = FreeFile()
    Open ArchivoBanHD For Input As #ArchN

    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanHDs.Add Tmp
    Loop

    Close #ArchN
End Sub
Sub ReSpawnOrigPosNpcs()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim i     As Integer

        Dim MiNPC As npc
       
        For i = 1 To LastNPC

                'OJO
                If Npclist(i).flags.NPCActive Then
            
                        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                                MiNPC = Npclist(i)
                                Call QuitarNPC(i)
                                Call ReSpawnNpc(MiNPC)

                        End If
            
                        'tildada por sugerencia de yind
                        'If Npclist(i).Contadores.TiempoExistencia > 0 Then
                        '        Call MuereNpc(i, 0)
                        'End If
                End If
       
        Next i
    
End Sub

Sub WorldSave()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim LoopX As Integer

        Dim hFile As Integer
    
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))
    
        #If SeguridadAlkon Then
                Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
        #End If
    
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    
        Dim j As Integer, k As Integer
    
        For j = 1 To NumMaps

                If MapInfo(j).BackUp = 1 Then k = k + 1
        Next j
    
        FrmStat.ProgressBar1.min = 0
        FrmStat.ProgressBar1.max = k
        FrmStat.ProgressBar1.Value = 0
    
        For LoopX = 1 To NumMaps
                'DoEvents
        
                If MapInfo(LoopX).BackUp = 1 Then
                        Call GrabarMapa(LoopX, App.Path & "\WorldBackUp\Mapa" & LoopX)
                        FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1

                End If
    
        Next LoopX
    
        FrmStat.Visible = False
    
        If FileExist(DatPath & "\bkNpcs.dat") Then Kill (DatPath & "bkNpcs.dat")
    
        hFile = FreeFile()
    
        Open DatPath & "\bkNpcs.dat" For Output As hFile
    
        For LoopX = 1 To LastNPC

                If Npclist(LoopX).flags.BackUp = 1 Then
                        Call BackUPnPc(LoopX, hFile)

                End If

        Next LoopX
        
        Close hFile
    
        Call SaveForums
    
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha conclu�do.", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub PurgarPenas()

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        ' @@ Todo
        Dim i As Long
        
        
        
        For i = 1 To LastUser

                If UserList(i).flags.UserLogged Then
                        If UserList(i).Counters.Pena > 0 Then
                                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                                If UserList(i).Counters.Pena < 1 Then
                                        UserList(i).Counters.Pena = 0
                                        ' @@ Miqueas : 07/11/15
                                        Call WarpUserChar(i, Configuracion.Libertad.Map, Configuracion.Libertad.X, Configuracion.Libertad.Y, True)
                                        Call WriteConsoleMsg(i, "�Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                                        Call FlushBuffer(i)

                                End If

                        End If

                End If

        Next i

End Sub

Public Sub Encarcelar(ByVal Userindex As Integer, _
                      ByVal Minutos As Long, _
                      Optional ByVal GmName As String = vbNullString)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        UserList(Userindex).Counters.Pena = Minutos
    
        ' @@ Miqueas : 07/11/15
        Call WarpUserChar(Userindex, Configuracion.Prision.Map, Configuracion.Prision.X, Configuracion.Prision.Y, True)
    
        If LenB(GmName) = 0 Then
                Call WriteConsoleMsg(Userindex, "Has sido encarcelado, deber�s permanecer en la c�rcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
        Else
                Call WriteConsoleMsg(Userindex, GmName & " te ha encarcelado, deber�s permanecer en la c�rcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)

        End If

        If UserList(Userindex).flags.Traveling = 1 Then
                Call EndTravel(Userindex, True)

        End If

End Sub

Public Sub BorrarUsuario(ByVal UserName As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
                Kill CharPath & UCase$(UserName) & ".chr"

        End If

End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        'Unban the character
        Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
    
        'Remove it from the banned people database
        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")

End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim i As Integer
    
        If MD5ClientesActivado = 1 Then

                For i = 0 To UBound(MD5s)

                        If (md5formateado = MD5s(i)) Then
                                MD5ok = True
                                Exit Function

                        End If

                Next i

                MD5ok = False
        Else
                MD5ok = True

        End If

End Function

Public Sub MD5sCarga()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim LoopC As Integer
    
        MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))
    
        If MD5ClientesActivado = 1 Then
                ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))

                For LoopC = 0 To UBound(MD5s)
                        MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
                        MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
                Next LoopC

        End If

End Sub

Public Sub BanIpAgrega(ByVal Ip As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        BanIps.Add Ip
    
        Call BanIpGuardar

End Sub

Public Function BanIpBuscar(ByVal Ip As String) As Long
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim Dale  As Boolean

        Dim LoopC As Long
    
        Dale = True
        LoopC = 1

        Do While LoopC <= BanIps.Count And Dale
                Dale = (BanIps.Item(LoopC) <> Ip)
                LoopC = LoopC + 1
        Loop
    
        If Dale Then
                BanIpBuscar = 0
        Else
                BanIpBuscar = LoopC - 1

        End If

End Function

Public Function BanIpQuita(ByVal Ip As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim N As Long
    
        N = BanIpBuscar(Ip)

        If N > 0 Then
                BanIps.Remove N
                BanIpGuardar
                BanIpQuita = True
        Else
                BanIpQuita = False

        End If

End Function

Public Sub BanIpGuardar()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim ArchivoBanIp As String

        Dim ArchN        As Long

        Dim LoopC        As Long
    
        ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
        ArchN = FreeFile()
        Open ArchivoBanIp For Output As #ArchN
    
        For LoopC = 1 To BanIps.Count
                Print #ArchN, BanIps.Item(LoopC)
        Next LoopC
    
        Close #ArchN

End Sub

Public Sub BanIpCargar()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim ArchN        As Long

        Dim Tmp          As String

        Dim ArchivoBanIp As String
    
        ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
        Set BanIps = New Collection
    
        ArchN = FreeFile()
        Open ArchivoBanIp For Input As #ArchN
    
        Do While Not EOF(ArchN)
                Line Input #ArchN, Tmp
                BanIps.Add Tmp
        Loop
    
        Close #ArchN

End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
        '***************************************************
        'Author: Unknown
        'Last Modification: 03/02/07
        'Last Modified By: Juan Mart�n Sotuyo Dodero (Maraxus)
        '***************************************************

        If EsAdmin(Name) Then
                UserDarPrivilegioLevel = PlayerType.Admin
        ElseIf EsDios(Name) Then
                UserDarPrivilegioLevel = PlayerType.Dios
        ElseIf EsSemiDios(Name) Then
                UserDarPrivilegioLevel = PlayerType.SemiDios
        ElseIf EsConsejero(Name) Then
                UserDarPrivilegioLevel = PlayerType.Consejero
        Else
                UserDarPrivilegioLevel = PlayerType.User

        End If

End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, _
                        ByVal UserName As String, _
                        ByVal Reason As String)
        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 03/02/07
        '22/05/2010: Ya no se peude banear admins de mayor rango si estan online.
        '***************************************************

        Dim tUser     As Integer

        Dim userPriv  As Byte

        Dim cantPenas As Byte

        Dim Rank      As Integer
    
        If InStrB(UserName, "+") Then
                UserName = Replace(UserName, "+", " ")

        End If
    
        tUser = NameIndex(UserName)
    
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
        With UserList(bannerUserIndex)

                If tUser <= 0 Then
                        Call WriteConsoleMsg(bannerUserIndex, "El usuario no est� online.", FontTypeNames.FONTTYPE_TALK)
            
                        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                                userPriv = UserDarPrivilegioLevel(UserName)
                
                                If (userPriv And Rank) > (.flags.Privilegios And Rank) Then
                                        Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarqu�a.", FontTypeNames.FONTTYPE_INFO)
                                Else

                                        If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                                                Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                                        Else
                                                Call LogBanFromName(UserName, bannerUserIndex, Reason)
                                                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                                                'ponemos el flag de ban a 1
                                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                                                'ponemos la pena
                                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                        
                                                If (userPriv And Rank) = (.flags.Privilegios And Rank) Then
                                                        .flags.Ban = 1
                                                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                                                        Call CloseSocket(bannerUserIndex)

                                                End If
                        
                                                Call LogGM(.Name, "BAN a " & UserName)

                                        End If

                                End If

                        Else
                                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                        End If

                Else

                        If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarqu�a.", FontTypeNames.FONTTYPE_INFO)
                        Else
            
                                Call LogBan(tUser, bannerUserIndex, Reason)
                                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
                
                                'Ponemos el flag de ban a 1
                                UserList(tUser).flags.Ban = 1
                
                                If (UserList(tUser).flags.Privilegios And Rank) = (.flags.Privilegios And Rank) Then
                                        .flags.Ban = 1
                                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                                        Call CloseSocket(bannerUserIndex)

                                End If
                
                                Call LogGM(.Name, "BAN a " & UserName)
                
                                'ponemos el flag de ban a 1
                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                                'ponemos la pena
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                
                                Call CloseSocket(tUser)

                        End If

                End If

        End With

End Sub


Public Function RemoverRegistroT0(ByVal CPU_ID As String) As Boolean    '//Disco.

    On Error Resume Next

    Dim N As Long

    N = BuscarRegistroT0(CPU_ID)

    If N > 0 Then
        BanT0s.Remove N
        RegistroBanT0
        RemoverRegistroT0 = True
    Else
        RemoverRegistroT0 = False
    End If

End Function

Public Sub AgregarRegistroT0(ByVal CPU_ID As String)

    BanT0s.Add CPU_ID
    Call RegistroBanT0
    
End Sub

Public Function BuscarRegistroT0(ByVal CPU_ID As String) As Long    '//Disco.

    Dim Dale As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1
    
    Do While LoopC <= BanT0s.Count And Dale
        Dale = (BanT0s.Item(LoopC) <> CPU_ID)
        LoopC = LoopC + 1
    Loop

    If Dale Then
        BuscarRegistroT0 = 0
    Else
        BuscarRegistroT0 = LoopC - 1
    End If

End Function

Public Sub RegistroBanT0()    '//Disco.

    Dim ArchivoBanT0 As String
    Dim ArchN As Long
    Dim LoopC As Long

    ArchivoBanT0 = App.Path & "\Dat\BanT0s.dat"

    ArchN = FreeFile()
    Open ArchivoBanT0 For Output As #ArchN

    For LoopC = 1 To BanT0s.Count
        Print #ArchN, BanT0s.Item(LoopC)
    Next LoopC

    Close #ArchN

End Sub

Public Sub BanT0Cargar()    '//Disco.

On Error Resume Next
    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanT0 As String
    
    ArchivoBanT0 = App.Path & "\Dat\BanT0s.dat"

    Do While BanT0s.Count > 0
        BanT0s.Remove 1
    Loop

    ArchN = FreeFile()
    Open ArchivoBanT0 For Input As #ArchN

    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanT0s.Add Tmp
    Loop

    Close #ArchN

End Sub




