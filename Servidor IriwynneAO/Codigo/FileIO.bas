Attribute VB_Name = "ES"
Option Explicit

Private Type ConfigHAO

        NivelMaximo As Byte
        Dados(0 To 1) As Byte
        
        ' @@ [Mapas]
        MapaGm As Integer
        Mapa1vs1 As Integer
        Mapa2vs2  As Integer
        Prision As WorldPos
        Libertad As WorldPos
        MapaFaccionReal As Integer
        MapaFaccionCaos As Integer
        
        MAPA_PRETORIANO As Integer
        
        MAPA_PRETORIANOLEFT(0 To 1) As Byte
        MAPA_PRETORIANORIGHT(0 To 1) As Byte
        
        ' @@ [Limpieza]
        MapNoLimpiezaCant As Integer
        MapNoLimpieza() As Integer

        ' @@ [Rates]
        Experiencia As Byte
        Oro As Byte
        
        ' @@ [Trabajador]
        
        PescaTrabajador(0 To 1) As Byte
        PescaOtras(0 To 1) As Byte

        TalaTrabajador(0 To 1) As Byte
        TalaOtras(0 To 1) As Byte

        MinaTrabajador(0 To 1) As Byte
        MinaOtras(0 To 1) As Byte

End Type

Public Configuracion As ConfigHAO

Private Type nMess

        Activado       As Byte
        Mensajes()     As String
        Tiempo         As Byte
        Conteo         As Byte

End Type

Public ControlMensajes As nMess
Public ControlMensajes2 As nMess

Public ExpForLvl()     As Long

Public Function Get_ExpLvl(ByVal UserLvl As Byte) As Long

    On Error GoTo r

    Get_ExpLvl = ExpForLvl(UserLvl)
    Exit Function
    
r:
    
    Err.Clear
    On Error GoTo yafue
    Get_ExpLvl = 610000000
    Exit Function
    
yafue:
    Get_ExpLvl = 99999999
End Function

Public Sub CargarExperiencia()

        Dim rPath As String

        Dim Cant  As Byte

        Dim LoopX As Long

        rPath = App.Path & "\Configuracion.ini"

        Cant = val(GetVar(rPath, "Experiencia", "NumLvl"))
      
        ReDim ExpForLvl(1 To Cant) As Long
    
        For LoopX = 1 To Cant
                ExpForLvl(LoopX) = val(GetVar(rPath, "Experiencia", LoopX))

                If ExpForLvl(LoopX) = 0 Then ExpForLvl(LoopX) = 10000000 ' @@ Parche rarito, no creo, NI TENDRIA QUE PASAR
        Next LoopX

End Sub

Sub CargarMsgs()

        Dim rPath   As String

        Dim cantMss As Byte

        Dim LoopX   As Long

        rPath = App.Path & "\Configuracion.ini"

        cantMss = val(GetVar(rPath, "MENSAJES", "Cantidad"))

        With ControlMensajes

                ReDim .Mensajes(1 To cantMss) As String
    
                .Activado = val(GetVar(rPath, "MENSAJES", "Activado"))
    
                .Tiempo = val(GetVar(rPath, "MENSAJES", "Tiempo"))
    
                .Conteo = 0
    
                For LoopX = 1 To cantMss
                        .Mensajes(LoopX) = GetVar(rPath, "MENSAJES", "M" & LoopX)
                Next LoopX
    
        End With

End Sub

Sub CargarSponsor()

        Dim rPath   As String

        Dim cantMss As Byte

        Dim LoopX   As Long

        rPath = App.Path & "\SPONSOR.ini"

        cantMss = val(GetVar(rPath, "MENSAJES", "Cantidad"))

        With ControlMensajes2

                ReDim .Mensajes(1 To cantMss) As String
    
                .Activado = val(GetVar(rPath, "MENSAJES", "Activado"))
    
                .Tiempo = val(GetVar(rPath, "MENSAJES", "Tiempo"))
    
                .Conteo = 0
    
                For LoopX = 1 To cantMss
                        .Mensajes(LoopX) = GetVar(rPath, "MENSAJES", "M" & LoopX)
                Next LoopX
    
        End With

End Sub
Public Sub CargarConfiguracionHAO()
        '***************************************************
        'Author: Miqueas - (Marcos Zeni)
        'Last Modification: - 04/09/15
        '
        '***************************************************

        Dim Leer As clsIniManager

        Set Leer = New clsIniManager
        
        Dim TempStr As String

        Dim Ansci   As Byte
        
        Leer.Initialize "Configuracion.ini"
        
        ' @@ Rates del Servidor
        Configuracion.Experiencia = Leer.GetValue("Rates", "Experiencia")
        Configuracion.Oro = Leer.GetValue("Rates", "Oro")
        
        
        Ansci = Asc("-")
        
        TempStr = Leer.GetValue("Servidor", "Dados")
        Configuracion.Dados(0) = General.ReadField(1, TempStr, Ansci)
        Configuracion.Dados(1) = General.ReadField(2, TempStr, Ansci)
        Configuracion.NivelMaximo = Leer.GetValue("Servidor", "NivelMaximo")
        
        ' @@ Mapas
        Configuracion.MapaGm = Leer.GetValue("Mapas", "MapaGm")
        Configuracion.Mapa1vs1 = Leer.GetValue("Mapas", "Mapa1vs1")
        Configuracion.Mapa2vs2 = Leer.GetValue("Mapas", "Mapa2vs2")
        
        ' @@ Prision
        TempStr = vbNullString
        TempStr = Leer.GetValue("Mapas", "Prision")
        Configuracion.Prision.Map = General.ReadField(1, TempStr, Ansci)
        Configuracion.Prision.X = General.ReadField(2, TempStr, Ansci)
        Configuracion.Prision.Y = General.ReadField(3, TempStr, Ansci)
        
        ' @@ Libertad
        TempStr = vbNullString
        TempStr = Leer.GetValue("Mapas", "Libertad")
        Configuracion.Libertad.Map = General.ReadField(1, TempStr, Ansci)
        Configuracion.Libertad.X = General.ReadField(2, TempStr, Ansci)
        Configuracion.Libertad.Y = General.ReadField(3, TempStr, Ansci)
        
        TempStr = vbNullString
        TempStr = Leer.GetValue("Mapas", "MapaPretoriano")
        Configuracion.MAPA_PRETORIANO = General.ReadField(1, TempStr, Ansci)
        
        Configuracion.MAPA_PRETORIANOLEFT(0) = General.ReadField(2, TempStr, Ansci)
        Configuracion.MAPA_PRETORIANOLEFT(1) = General.ReadField(3, TempStr, Ansci)
        
        Configuracion.MAPA_PRETORIANORIGHT(0) = General.ReadField(4, TempStr, Ansci)
        Configuracion.MAPA_PRETORIANORIGHT(1) = General.ReadField(5, TempStr, Ansci)
        
        Dim LoopC As Long
     
        Configuracion.MapNoLimpiezaCant = val(Leer.GetValue("Limpieza", "Cantidad"))
        
        ReDim Configuracion.MapNoLimpieza(1 To Configuracion.MapNoLimpiezaCant) As Integer
        
        For LoopC = 1 To Configuracion.MapNoLimpiezaCant
                Configuracion.MapNoLimpieza(LoopC) = Leer.GetValue("Limpieza", "Mapa" & LoopC)
        Next LoopC

        ' @@ Faccionarios
        Configuracion.MapaFaccionReal = Leer.GetValue("Mapas", "MapReal")
        Configuracion.MapaFaccionCaos = Leer.GetValue("Mapas", "MapCaos")
        
        
        ' @@ Trabajo
        TempStr = vbNullString
        TempStr = Leer.GetValue("Trabajo", "PescaTrabajador")
        Configuracion.PescaTrabajador(0) = val(General.ReadField(1, TempStr, Ansci))
        Configuracion.PescaTrabajador(1) = val(General.ReadField(2, TempStr, Ansci))
        
        TempStr = vbNullString
        TempStr = Leer.GetValue("Trabajo", "PescaOtras")
        Configuracion.PescaOtras(0) = val(General.ReadField(1, TempStr, Ansci))
        Configuracion.PescaOtras(1) = val(General.ReadField(2, TempStr, Ansci))
        
        TempStr = vbNullString
        TempStr = Leer.GetValue("Trabajo", "TalaTrabajador")
        Configuracion.TalaTrabajador(0) = val(General.ReadField(1, TempStr, Ansci))
        Configuracion.TalaTrabajador(1) = val(General.ReadField(2, TempStr, Ansci))
        
        TempStr = vbNullString
        TempStr = Leer.GetValue("Trabajo", "TalaOtras")
        Configuracion.TalaOtras(0) = val(General.ReadField(1, TempStr, Ansci))
        Configuracion.TalaOtras(1) = val(General.ReadField(2, TempStr, Ansci))
        
        TempStr = vbNullString
        TempStr = Leer.GetValue("Trabajo", "MinaTrabajador")
        Configuracion.MinaTrabajador(0) = val(General.ReadField(1, TempStr, Ansci))
        Configuracion.MinaTrabajador(1) = val(General.ReadField(2, TempStr, Ansci))
        
        TempStr = vbNullString
        TempStr = Leer.GetValue("Trabajo", "MinaOtras")
        Configuracion.MinaOtras(0) = val(General.ReadField(1, TempStr, Ansci))
        Configuracion.MinaOtras(1) = val(General.ReadField(2, TempStr, Ansci))
                
        Set Leer = Nothing

End Sub

Public Sub CargarSpawnList()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim N As Integer, LoopC As Integer

        N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
        ReDim SpawnList(N) As tCriaturasEntrenador

        For LoopC = 1 To N
                SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
                SpawnList(LoopC).NPCNAME = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
        Next LoopC
    
End Sub

Function EsAdmin(ByRef Name As String) As Boolean

        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)

End Function

Function EsDios(ByRef Name As String) As Boolean

        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)

End Function

Function EsSemiDios(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)

End Function

Function EsGmEspecial(ByRef Name As String) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        EsGmEspecial = (val(Administradores.GetValue("Especial", Name)) = 1)

End Function

Function EsConsejero(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        EsConsejero = (val(Administradores.GetValue("Consejero", Name)) = 1)

End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        EsRolesMaster = (val(Administradores.GetValue("RM", Name)) = 1)

End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 27/03/2011
        'Returns true if char is administrative user.
        '***************************************************
    
        Dim esGM As Boolean
    
        ' Admin?
        esGM = EsAdmin(Name)

        ' Dios?
        If Not esGM Then esGM = EsDios(Name)

        ' Semidios?
        If Not esGM Then esGM = EsSemiDios(Name)

        ' Consejero?
        If Not esGM Then esGM = EsConsejero(Name)

        EsGmChar = esGM

End Function

Public Sub loadAdministrativeUsers()
        'Admines     => Admin
        'Dioses      => Dios
        'SemiDioses  => SemiDios
        'Especiales  => Especial
        'Consejeros  => Consejero
        'RoleMasters => RM

        'Si esta mierda tuviese array asociativos el c�digo ser�a tan lindo.
        Dim buf  As Integer

        Dim i    As Long

        Dim Name As String
       
        ' Public container
        Set Administradores = New clsIniManager
    
        ' Server ini info file
        Dim ServerIni As clsIniManager

        Set ServerIni = New clsIniManager
    
        Call ServerIni.Initialize(IniPath & "Server.ini")
       
        ' Admines
        buf = val(ServerIni.GetValue("INIT", "Admines"))
    
        For i = 1 To buf
                Name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
                If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
                ' Add key
                Call Administradores.ChangeValue("Admin", Name, "1")

        Next i
    
        ' Dioses
        buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
        For i = 1 To buf
                Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
                If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
                ' Add key
                Call Administradores.ChangeValue("Dios", Name, "1")
        
        Next i
    
        ' Especiales
        buf = val(ServerIni.GetValue("INIT", "Especiales"))
    
        For i = 1 To buf
                Name = UCase$(ServerIni.GetValue("Especiales", "Especial" & i))
        
                If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
                ' Add key
                Call Administradores.ChangeValue("Especial", Name, "1")
        
        Next i
    
        ' SemiDioses
        buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
        For i = 1 To buf
                Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
                If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
                ' Add key
                Call Administradores.ChangeValue("SemiDios", Name, "1")
        
        Next i
    
        ' Consejeros
        buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
        For i = 1 To buf
                Name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
                If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
                ' Add key
                Call Administradores.ChangeValue("Consejero", Name, "1")
        
        Next i
    
        ' RolesMasters
        buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
        For i = 1 To buf
                Name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
                If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
                ' Add key
                Call Administradores.ChangeValue("RM", Name, "1")
        Next i
    
        Set ServerIni = Nothing
    
End Sub

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
        '****************************************************
        'Author: ZaMa
        'Last Modification: 18/11/2010
        'Reads the user's charfile and retrieves its privs.
        '***************************************************

        Dim Privs As PlayerType

        If EsAdmin(UserName) Then
                Privs = PlayerType.Admin
        
        ElseIf EsDios(UserName) Then
                Privs = PlayerType.Dios

        ElseIf EsSemiDios(UserName) Then
                Privs = PlayerType.SemiDios
        
        ElseIf EsConsejero(UserName) Then
                Privs = PlayerType.Consejero
    
        Else
                Privs = PlayerType.User

        End If

        GetCharPrivs = Privs

End Function

Public Function TxtDimension(ByVal Name As String) As Long
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim N As Integer, cad As String, Tam As Long

        N = FreeFile(1)
        Open Name For Input As #N
        Tam = 0

        Do While Not EOF(N)
                Tam = Tam + 1
                Line Input #N, cad
        Loop
        Close N
        TxtDimension = Tam

End Function

Public Sub CargarForbidenWords()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))

        Dim N As Integer, i As Integer

        N = FreeFile(1)
        Open DatPath & "NombresInvalidos.txt" For Input As #N
    
        For i = 1 To UBound(ForbidenNames)
                Line Input #N, ForbidenNames(i)
        Next i
    
        Close N

End Sub

Public Sub CargarHechizos()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        '###################################################
        '#               ATENCION PELIGRO                  #
        '###################################################
        '
        '  ���� NO USAR GetVar PARA LEER Hechizos.dat !!!!
        '
        'El que ose desafiar esta LEY, se las tendr� que ver
        'con migo. Para leer Hechizos.dat se deber� usar
        'la nueva clase clsLeerInis.
        '
        'Alejo
        '
        '###################################################

        On Error GoTo errhandler

        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
    
        Dim Hechizo As Integer

        Dim Leer    As clsIniManager

        Set Leer = New clsIniManager
    
        Call Leer.Initialize(DatPath & "Hechizos.dat")
    
        'obtiene el numero de hechizos
        NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    
        ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
        frmCargando.cargar.min = 0
        frmCargando.cargar.max = NumeroHechizos
        frmCargando.cargar.Value = 0
    
        'Llena la lista
        For Hechizo = 1 To NumeroHechizos

                With Hechizos(Hechizo)
                        .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
                        .desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
                        .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
                        .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
                        .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
                        .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
                        .Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
                        .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
                        .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
                        .Loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
                        '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
                        .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
                        .MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
                        .MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
                        .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
                        .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
                        .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
                        .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
                        .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
                        .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
                        .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
                        .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
                        .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
                        .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
                        .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
                        .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
                        .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
                        .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
                        .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
                        .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
                        .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
                        .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
                        .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
                        .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
                        .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
                        .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
                        .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
                        .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
                        .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
                        .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
                        .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
                        .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
                        .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
                        .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
                        .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
                        .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
                        .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
                        .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
                        .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
                        .Warp = val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
                        .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
                        .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
                        .Cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
                        .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
                        '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
                        '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
                        .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
                        .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
                        'Barrin 30/9/03
                        .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
                        .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
                        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            
                        .NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
                        .StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))

                End With

        Next Hechizo
    
        Set Leer = Nothing
    
        Exit Sub

errhandler:
        MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim i As Integer
    
        MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
    
        'ReDim MOTD(1 To MaxLines)

        'For i = 1 To MaxLines
        '        MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
        '        MOTD(i).Formato = vbNullString
        'Next i

End Sub

Public Sub DoBackUp()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        haciendoBK = True
    
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
        'Call LimpiarMundo
        Call WorldSave
        Call modGuilds.v_RutinaElecciones
        'Call ResetCentinelaInfo     'Reseteamos al centinela 'Lo saco porque ahora el reset lo maneja el modCentinela [C4b3z0n]
    
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
        haciendoBK = False
    
        'Log
        On Error Resume Next

        Dim nfile As Integer

        nfile = FreeFile ' obtenemos un canal
        Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
        Print #nfile, Date & " " & time
        Close #nfile

End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MAPFILE As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/2011
        '10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
        '28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
        '12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados y Centinela)
        '***************************************************

        'On Error Resume Next

        Dim FreeFileMap As Long

        Dim FreeFileInf As Long

        Dim Y           As Long

        Dim X           As Long

        Dim ByFlags     As Byte

        Dim LoopC       As Long

        Dim MapWriter   As clsByteBuffer

        Dim InfWriter   As clsByteBuffer

        Dim IniManager  As clsIniManager

        Dim NpcInvalido As Boolean
    
        Set MapWriter = New clsByteBuffer
        Set InfWriter = New clsByteBuffer
        Set IniManager = New clsIniManager
    
        If FileExist(MAPFILE & ".map", vbNormal) Then
                Kill MAPFILE & ".map"

        End If
    
        If FileExist(MAPFILE & ".inf", vbNormal) Then
                Kill MAPFILE & ".inf"

        End If
    
        'Open .map file
        FreeFileMap = FreeFile
        Open MAPFILE & ".Map" For Binary As FreeFileMap
    
        Call MapWriter.initializeWriter(FreeFileMap)
    
        'Open .inf file
        FreeFileInf = FreeFile
        Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
        Call InfWriter.initializeWriter(FreeFileInf)
    
        'map Header
        Call MapWriter.putInteger(MapInfo(Map).MapVersion)
        
        Call MapWriter.putString(MiCabecera.desc, False)
        Call MapWriter.putLong(MiCabecera.crc)
        Call MapWriter.putLong(MiCabecera.MagicWord)
    
        Call MapWriter.putDouble(0)
    
        'inf Header
        Call InfWriter.putDouble(0)
        Call InfWriter.putInteger(0)
    
        'Write .map file
        For Y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize

                        With MapData(Map, X, Y)
                                ByFlags = 0
                
                                If .Blocked Then ByFlags = ByFlags Or 1
                                If .Graphic(2) Then ByFlags = ByFlags Or 2
                                If .Graphic(3) Then ByFlags = ByFlags Or 4
                                If .Graphic(4) Then ByFlags = ByFlags Or 8
                                If .trigger Then ByFlags = ByFlags Or 16
                
                                Call MapWriter.putByte(ByFlags)
                
                                Call MapWriter.putInteger(.Graphic(1))
                
                                For LoopC = 2 To 4

                                        If .Graphic(LoopC) Then _
                                           Call MapWriter.putInteger(.Graphic(LoopC))
                                Next LoopC
                
                                If .trigger Then _
                                   Call MapWriter.putInteger(CInt(.trigger))
                
                                '.inf file
                                ByFlags = 0
                
                                If .ObjInfo.objIndex > 0 Then
                                        If ObjData(.ObjInfo.objIndex).OBJType = eOBJType.otFogata Then
                                                .ObjInfo.objIndex = 0
                                                .ObjInfo.Amount = 0

                                        End If

                                End If
    
                                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                                ' No hacer backup de los NPCs inv�lidos (Pretorianos, Mascotas, Invocados y Centinela)
                                If .NpcIndex Then
                                        NpcInvalido = (Npclist(.NpcIndex).NPCtype = eNPCType.Pretoriano) Or (Npclist(.NpcIndex).MaestroUser > 0) Or EsCentinela(.NpcIndex)
                    
                                        If Not NpcInvalido Then ByFlags = ByFlags Or 2

                                End If
                
                                If .ObjInfo.objIndex Then ByFlags = ByFlags Or 4
                
                                Call InfWriter.putByte(ByFlags)
                
                                If .TileExit.Map Then
                                        Call InfWriter.putInteger(.TileExit.Map)
                                        Call InfWriter.putInteger(.TileExit.X)
                                        Call InfWriter.putInteger(.TileExit.Y)

                                End If
                
                                If .NpcIndex And Not NpcInvalido Then _
                                   Call InfWriter.putInteger(Npclist(.NpcIndex).Numero)
                
                                If .ObjInfo.objIndex Then
                                        Call InfWriter.putInteger(.ObjInfo.objIndex)
                                        Call InfWriter.putInteger(.ObjInfo.Amount)

                                End If
                
                                NpcInvalido = False

                        End With

                Next X
        Next Y
    
        Call MapWriter.saveBuffer
        Call InfWriter.saveBuffer
    
        'Close .map file
        Close FreeFileMap

        'Close .inf file
        Close FreeFileInf
    
        Set MapWriter = Nothing
        Set InfWriter = Nothing

        With MapInfo(Map)
                'write .dat file
                Call IniManager.ChangeValue("Mapa" & Map, "Name", .Name)
                Call IniManager.ChangeValue("Mapa" & Map, "MusicNum", .Music)
                Call IniManager.ChangeValue("Mapa" & Map, "MagiaSinefecto", .MagiaSinEfecto)
                Call IniManager.ChangeValue("Mapa" & Map, "InviSinEfecto", .InviSinEfecto)
                Call IniManager.ChangeValue("Mapa" & Map, "ResuSinEfecto", .ResuSinEfecto)
                Call IniManager.ChangeValue("Mapa" & Map, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
                Call IniManager.ChangeValue("Mapa" & Map, "OnDeathGoTo", .OnDeathGoTo.Map & "-" & .OnDeathGoTo.X & "-" & .OnDeathGoTo.Y)
    
                Call IniManager.ChangeValue("Mapa" & Map, "Terreno", TerrainByteToString(.Terreno))
                Call IniManager.ChangeValue("Mapa" & Map, "Zona", .Zona)
                Call IniManager.ChangeValue("Mapa" & Map, "Restringir", RestrictByteToString(.Restringir))
                Call IniManager.ChangeValue("Mapa" & Map, "BackUp", str$(.BackUp))
    
                If .Pk Then
                        Call IniManager.ChangeValue("Mapa" & Map, "Pk", "0")
                Else
                        Call IniManager.ChangeValue("Mapa" & Map, "Pk", "1")

                End If
        
                Call IniManager.ChangeValue("Mapa" & Map, "OcultarSinEfecto", .OcultarSinEfecto)
                Call IniManager.ChangeValue("Mapa" & Map, "InvocarSinEfecto", .InvocarSinEfecto)
                Call IniManager.ChangeValue("Mapa" & Map, "NoEncriptarMP", .NoEncriptarMP)
                Call IniManager.ChangeValue("Mapa" & Map, "RoboNpcsPermitido", .RoboNpcsPermitido)
    
                Call IniManager.DumpFile(MAPFILE & ".dat")

        End With
    
        Set IniManager = Nothing

End Sub

Sub LoadArmasHerreria()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim N As Integer, lc As Integer
    
        N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
        ReDim Preserve ArmasHerrero(1 To N) As Integer
    
        For lc = 1 To N
                ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
        Next lc

End Sub

Sub LoadArmadurasHerreria()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim N As Integer, lc As Integer
    
        N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
        ReDim Preserve ArmadurasHerrero(1 To N) As Integer
    
        For lc = 1 To N
                ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
        Next lc

End Sub

Sub LoadBalance()
        '***************************************************
        'Author: Unknown
        'Last Modification: 15/04/2010
        '15/04/2010: ZaMa - Agrego recompensas faccionarias.
        '***************************************************

        Dim i As Long
    
        'Modificadores de Clase
        For i = 1 To NUMCLASES

                With ModClase(i)
                        .Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
                        .AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
                        .AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
                        .AtaqueWrestling = val(GetVar(DatPath & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
                        .Da�oArmas = val(GetVar(DatPath & "Balance.dat", "MODDA�OARMAS", ListaClases(i)))
                        .Da�oProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDA�OPROYECTILES", ListaClases(i)))
                        .Da�oWrestling = val(GetVar(DatPath & "Balance.dat", "MODDA�OWRESTLING", ListaClases(i)))
                        .Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))

                End With

        Next i
    
        'Modificadores de Raza
        For i = 1 To NUMRAZAS

                With ModRaza(i)
                        .Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
                        .Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
                        .Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
                        .Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
                        .Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))

                End With

        Next i
    
        'Modificadores de Vida
        For i = 1 To NUMCLASES
                ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
        Next i
    
        'Distribuci�n de Vida
        For i = 1 To 5
                DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
        Next i

        For i = 1 To 4
                DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
        Next i
    
        'Extra
        PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

        'Party
        ExponenteNivelParty = val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
    
        ' Recompensas faccionarias
        For i = 1 To NUM_RANGOS_FACCION
                RecompensaFacciones(i - 1) = val(GetVar(DatPath & "Balance.dat", "RECOMPENSAFACCION", "Rango" & i))
        Next i
    
End Sub

Sub LoadObjCarpintero()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim N As Integer, lc As Integer
    
        N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
        ReDim Preserve ObjCarpintero(1 To N) As Integer
    
        For lc = 1 To N
                ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
        Next lc

End Sub

Sub LoadOBJData()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        '###################################################
        '#               ATENCION PELIGRO                  #
        '###################################################
        '
        '���� NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
        '
        'El que ose desafiar esta LEY, se las tendr� que ver
        'con migo. Para leer desde el OBJ.DAT se deber� usar
        'la nueva clase clsLeerInis.
        '
        'Alejo
        '
        '###################################################

        'Call LogTarea("Sub LoadOBJData")

        On Error GoTo errhandler

        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
    
        '*****************************************************************
        'Carga la lista de objetos
        '*****************************************************************
        Dim Object As Integer

        Dim Leer   As clsIniManager

        Set Leer = New clsIniManager
    
        Call Leer.Initialize(DatPath & "Obj.dat")
    
        'obtiene el numero de obj
        NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
        frmCargando.cargar.min = 0
        frmCargando.cargar.max = NumObjDatas
        frmCargando.cargar.Value = 0
    
        ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
        'Llena la lista
        For Object = 1 To NumObjDatas

                With ObjData(Object)
48                        .Name = Leer.GetValue("OBJ" & Object, "Name")
            
                        'Pablo (ToxicWaste) Log de Objetos.
                        .Log = val(Leer.GetValue("OBJ" & Object, "Log"))
                        .NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
                        '07/09/07
            
                        .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))

                        If .GrhIndex = 0 Then
                                .GrhIndex = .GrhIndex

                        End If
            
45                        .OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
            
46                        .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
            
                        Select Case .OBJType

                                Case eOBJType.otArmadura
                                        .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                                        .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                                        .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                                        .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                                        .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                                        .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                
                                Case eOBJType.otESCUDO
                                        .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                                        .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                                        .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                                        .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                                        .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                                        .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                                        .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                                Case eOBJType.otCASCO
                                        .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                                        .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                                        .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                                        .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                                        .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                                        .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                                        .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                                Case eOBJType.otWeapon
                                        .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                                        .Apu�ala = val(Leer.GetValue("OBJ" & Object, "Apu�ala"))
                                        .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                                        .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                                        .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                                        .proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                                        .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                                        .StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
                                        .StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                                        .Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                    
                                        .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                                        .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                                        .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                                        .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                                        .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                                        .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                                        .WeaponRazaEnanaAnim = val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                
                                Case eOBJType.otInstrumentos
                                        .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                                        .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                                        .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                                        'Pablo (ToxicWaste)
                                        .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                                        .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                                Case eOBJType.otMinerales
                                        .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                
                                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                                        .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                                        .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                                        .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                                Case otPociones
                                        .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                                        .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                                        .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                                        .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                                Case eOBJType.otBarcos
                                        .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                                        .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                                        .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                
                                Case eOBJType.otFlechas
                                        .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                                        .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                                        .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                                        .Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                                        
                                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                                        .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                                        .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                                        .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                                        .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                                        .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                                        .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    
                                Case eOBJType.otTeleport
                                        .Radio = val(Leer.GetValue("OBJ" & Object, "Radio"))
                
                                Case eOBJType.otForos
                                        Call AddForum(Leer.GetValue("OBJ" & Object, "ID"))
                                        
                                Case eOBJType.otCofresDonador
33                                        Call Mod_Cofres.LoadItemDrop(Object, Leer)
                   
                        End Select
            
                        .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
                        .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
                        .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
                        .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
                        .MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
                        .MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
                        .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
                        .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
                        .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
                        .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
                        .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
                        .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
                        .def = (.MinDef + .MaxDef) / 2
            
                        .RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
                        .RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
                        .RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
                        .RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
                        .RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
            
1                        .Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
3                        .Shop = val(Leer.GetValue("OBJ" & Object, "Shop"))
2                        .copas = val(Leer.GetValue("OBJ" & Object, "Copas"))
5                        .Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
                        
                        .SeLimpia = val(Leer.GetValue("OBJ" & Object, "SeLimpia"))
            
                        .Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))

                        If .Cerrada = 1 Then
                                .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
                                .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))

                        End If
            
                        'Puertas y llaves
                        .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            
21                        .texto = Leer.GetValue("OBJ" & Object, "Texto")
22                        .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
                        .Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
                        .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
                        .Acuchilla = val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
            
                        .Guante = val(Leer.GetValue("OBJ" & Object, "Guante"))
            
                        'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
                        Dim i As Integer

                        Dim N As Integer

                        Dim s As String

                        For i = 1 To NUMCLASES
                                s = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
                                N = 1

                                Do While LenB(s) > 0 And UCase$(ListaClases(N)) <> s
                                        N = N + 1
                                Loop
                                .ClaseProhibida(i) = IIf(LenB(s) > 0, N, 0)
                        Next i
            
                        .DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
                        .DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            
                        .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
                        If .SkCarpinteria > 0 Then _
                           .Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
                        .MaderaElfica = val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
            
                        'Bebidas
55                        .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
            
54                        .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
453                        .Upgrade = val(Leer.GetValue("OBJ" & Object, "Upgrade"))
                        
                        .ObjNoCreable = val(Leer.GetValue("OBJ" & Object, "NoCreable"))
            
                        frmCargando.cargar.Value = frmCargando.cargar.Value + 1

                End With

        Next Object
    
        Set Leer = Nothing
    
        ' Inicializo los foros faccionarios
        Call AddForum(FORO_CAOS_ID)
        Call AddForum(FORO_REAL_ID)
    
        Exit Sub

errhandler:
        MsgBox "error cargando objetos " & Err.Number & ": " & Err.description & " - " & Erl

End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

        '*************************************************
        'Author: Unknown
        'Last modified: 11/19/2009
        '11/19/2009: Pato - Load the EluSkills and ExpSkills
        '*************************************************
        Dim LoopC As Long

        With UserList(UserIndex)
                With .Stats

                        For LoopC = 1 To NUMATRIBUTOS
                                .UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
                                .UserAtributosBackUP(LoopC) = .UserAtributos(LoopC)
                        Next LoopC
        
                        For LoopC = 1 To NUMSKILLS
                                .UserSkills(LoopC) = val(UserFile.GetValue("SKILLS", "SK" & LoopC))
                        Next LoopC
        
                        For LoopC = 1 To MAXUSERHECHIZOS
                                .UserHechizos(LoopC) = val(UserFile.GetValue("Hechizos", "H" & LoopC))
                        Next LoopC
        
                        .GLD = CLng(UserFile.GetValue("STATS", "GLD"))
                        .Banco = CLng(UserFile.GetValue("STATS", "BANCO"))
        
                        .MaxHP = CInt(UserFile.GetValue("STATS", "MaxHP"))
                        .MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))
        
                        .MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
                        .MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
        
                        .MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
                        .MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))
        
                        .MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
                        .MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))
        
                        .MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
                        .MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))
        
                        .MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
                        .MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))
        
                        .SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))
        
                        .Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
                        .ELU = CLng(UserFile.GetValue("STATS", "ELU"))
                        .ELV = CByte(UserFile.GetValue("STATS", "ELV"))
        
                        .UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
                        .NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

                End With
    
                With .flags

                        If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then _
                           .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
        
                        If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then _
                           .Privilegios = .Privilegios Or PlayerType.ChaosCouncil

                End With

        End With

End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        With UserList(UserIndex).Reputacion
                .AsesinoRep = val(UserFile.GetValue("REP", "Asesino"))
                .BandidoRep = val(UserFile.GetValue("REP", "Bandido"))
                .BurguesRep = val(UserFile.GetValue("REP", "Burguesia"))
                .LadronesRep = val(UserFile.GetValue("REP", "Ladrones"))
                .NobleRep = val(UserFile.GetValue("REP", "Nobles"))
                .PlebeRep = val(UserFile.GetValue("REP", "Plebe"))
                .Promedio = val(UserFile.GetValue("REP", "Promedio"))

        End With
    
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

        '*************************************************
        'Author: Unknown
        'Last modified: 19/11/2006
        'Loads the Users RECORDs
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
        '*************************************************
        Dim LoopC As Long

        Dim ln    As String
    
        With UserList(UserIndex)
                
                If Len(UserFile.GetValue("DISCODURO", "VAL")) <> 0 Then
                .discoDuro = UserFile.GetValue("DISCODURO", "VAL")
                End If
                
                With .Faccion
                        .ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
                        .FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
                        .CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
                        .CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
                        .RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
                        .RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
                        .RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
                        .RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
                        .RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
                        .RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
                        .Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
                        .NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
                        .FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
                        .MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
                        .NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))

                End With
        
                With .flags
                                                
                        .Vip = CByte(UserFile.GetValue("FLAGS", "VIP"))
        
                        .Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
                        .Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
            
                        .Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
                        .Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
                        .Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
                        .Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
                        .Montando = CByte(UserFile.GetValue("FLAGS", "Montando"))
                        .Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
                        .Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
            
                        .PuntosShop = CLng(UserFile.GetValue("FLAGS", "PuntosShop"))
            
                        'Matrix
                        .lastMap = val(UserFile.GetValue("FLAGS", "LastMap"))

                End With
        
                If .flags.Paralizado = 1 Then
                        .Counters.Paralisis = IntervaloParalizado

                End If
        
                .Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
 
                .email = UserFile.GetValue("CONTACTO", "Email")
                .clave = UserFile.GetValue("CONTACTO", "Pin")
                
                .Genero = UserFile.GetValue("INIT", "Genero")
                .clase = UserFile.GetValue("INIT", "Clase")
                .raza = UserFile.GetValue("INIT", "Raza")
                .Hogar = UserFile.GetValue("INIT", "Hogar")
                .Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))
        
                With .OrigChar
                        .Head = CInt(UserFile.GetValue("INIT", "Head"))
                        .Body = CInt(UserFile.GetValue("INIT", "Body"))
                        .WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
                        .ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
                        .CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
            
                        .heading = eHeading.SOUTH

                End With
        
                #If ConUpTime Then
                        .UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
                #End If
        
                If .flags.Muerto = 0 Then
                        .Char = .OrigChar
                Else
                        .Char.Body = iCuerpoMuerto
                        .Char.Head = iCabezaMuerto
                        .Char.WeaponAnim = NingunArma
                        .Char.ShieldAnim = NingunEscudo
                        .Char.CascoAnim = NingunCasco

                End If
        
                .desc = UserFile.GetValue("INIT", "Desc")
        
                .Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
                .Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
                .Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))
        
                .Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))
        
                '[KEVIN]--------------------------------------------------------------------
                '***********************************************************************************
                .BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))

                'Lista de objetos del banco
                For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
                        ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
                        .BancoInvent.Object(LoopC).objIndex = CInt(ReadField(1, ln, 45))
                        .BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
                Next LoopC

                '------------------------------------------------------------------------------------
                '[/KEVIN]*****************************************************************************
        
                'Lista de objetos
                For LoopC = 1 To MAX_NORMAL_INVENTORY_SLOTS
                        ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
                        .Invent.Object(LoopC).objIndex = val(ReadField(1, ln, 45))
                        .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
                        .Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
                Next LoopC
        
                'Obtiene el indice-objeto del arma
                .Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))

                If .Invent.WeaponEqpSlot > 0 Then
                        .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).objIndex

                End If
        
                'Obtiene el indice-objeto del armadura
                .Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))

                If .Invent.ArmourEqpSlot > 0 Then
                        .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).objIndex
                        .flags.Desnudo = 0
                Else
                        .flags.Desnudo = 1

                End If
        
                'Obtiene el indice-objeto del escudo
                .Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))

                If .Invent.EscudoEqpSlot > 0 Then
                        .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).objIndex

                End If
        
                'Obtiene el indice-objeto del casco
                .Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))

                If .Invent.CascoEqpSlot > 0 Then
                        .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).objIndex

                End If
        
                'Obtiene el indice-objeto barco
                .Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
                If .Invent.BarcoSlot > 0 Then
                        .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).objIndex

                End If
        
                'Obtiene el indice-objeto municion
                .Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))

                If .Invent.MunicionEqpSlot > 0 Then
                        .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).objIndex

                End If
        
                '[Alejo]
                'Obtiene el indice-objeto anilo
                .Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))

                If .Invent.AnilloEqpSlot > 0 Then
                        .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).objIndex

                End If
        
                .NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))

                For LoopC = 1 To MAXMASCOTAS
                        .MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
                Next LoopC
        
                ln = UserFile.GetValue("Guild", "GUILDINDEX")
                
                If IsNumeric(ln) Then
                        .GuildIndex = CInt(ln)
                Else
                        .GuildIndex = 0

                End If
                

                .flags.Posteado = UserFile.GetValue("MERCADO", "POSTEADO")
                .flags.MeMando = UserFile.GetValue("MERCADO", "MEMANDO")
                .flags.Lemande = UserFile.GetValue("MERCADO", "LEMANDE")

        End With

End Sub

Function GetVar(ByVal file As String, _
                ByVal Main As String, _
                ByVal Var As String, _
                Optional EmptySpaces As Long = 1024) As String
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim sSpaces  As String ' This will hold the input that the program will retrieve

        Dim szReturn As String ' This will be the defaul value if the string is not found
      
        szReturn = vbNullString
      
        sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
        GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
      
        GetVar = RTrim$(sSpaces)
        GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."
    
        Dim Map       As Integer

        Dim tFileName As String
    
        On Error GoTo man
        
        NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        Call InitAreas
        
        frmCargando.cargar.min = 0
        frmCargando.cargar.max = NumMaps
        frmCargando.cargar.Value = 0
        
        MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        ReDim MapInfo(1 To NumMaps) As MapInfo
        
        For Map = 1 To NumMaps

                If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
                        tFileName = App.Path & "\WorldBackUp\Mapa" & Map
                
                        If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                                tFileName = App.Path & MapPath & "Mapa" & Map

                        End If

                Else
                        tFileName = App.Path & MapPath & "Mapa" & Map

                End If
            
                Call CargarMapa(Map, tFileName)
            
                frmCargando.cargar.Value = frmCargando.cargar.Value + 1
                DoEvents
        Next Map
    
        Exit Sub

man:
        MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
        Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."
    
        Dim Map       As Integer

        Dim tFileName As String
    
        On Error GoTo man
        
        NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        Call InitAreas
        
        frmCargando.cargar.min = 0
        frmCargando.cargar.max = NumMaps
        frmCargando.cargar.Value = 0
        
        MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        ReDim MapInfo(1 To NumMaps) As MapInfo
          
        For Map = 1 To NumMaps
            
                tFileName = App.Path & MapPath & "Mapa" & Map
                Call CargarMapa(Map, tFileName)
            
                frmCargando.cargar.Value = frmCargando.cargar.Value + 1
                DoEvents
        Next Map
    
        Exit Sub

man:
        MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
        Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByRef MAPFl As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: 10/08/2010
        '10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
        '***************************************************

        On Error GoTo errh

        Dim hFile     As Integer

        Dim X         As Long

        Dim Y         As Long

        Dim ByFlags   As Byte

        Dim npcfile   As String

        Dim Leer      As clsIniManager

        Dim MapReader As clsByteBuffer

        Dim InfReader As clsByteBuffer

        Dim Buff()    As Byte
    
        Set MapReader = New clsByteBuffer
        Set InfReader = New clsByteBuffer
        Set Leer = New clsIniManager
    
        npcfile = DatPath & "NPCs.dat"
    
1        hFile = FreeFile

2        Open MAPFl & ".map" For Binary As #hFile
3        Seek hFile, 1

4        ReDim Buff(LOF(hFile) - 1) As Byte
    
5        Get #hFile, , Buff
6        Close hFile
    
7        Call MapReader.initializeReader(Buff)

        'inf
8        Open MAPFl & ".inf" For Binary As #hFile
9        Seek hFile, 1

10        ReDim Buff(LOF(hFile) - 1) As Byte
    
11        Get #hFile, , Buff
12        Close hFile
    
13        Call InfReader.initializeReader(Buff)
    
        'map Header
14        MapInfo(Map).MapVersion = MapReader.getInteger
    
15        MiCabecera.desc = MapReader.getString(Len(MiCabecera.desc))
16        MiCabecera.crc = MapReader.getLong
17        MiCabecera.MagicWord = MapReader.getLong
    
18        Call MapReader.getDouble

        'inf Header
19        Call InfReader.getDouble
20        Call InfReader.getInteger

21        For Y = YMinMapSize To YMaxMapSize
22                For X = XMinMapSize To XMaxMapSize

23                        With MapData(Map, X, Y)
24                                '.map file
25                                ByFlags = MapReader.getByte
26
27                                If ByFlags And 1 Then .Blocked = 1
28
29                                .Graphic(1) = MapReader.getInteger
30
31                                'Layer 2 used?
32                                If ByFlags And 2 Then .Graphic(2) = MapReader.getInteger
33
34                                'Layer 3 used?
35                                If ByFlags And 4 Then .Graphic(3) = MapReader.getInteger
36
37                                'Layer 4 used?
38                                If ByFlags And 8 Then .Graphic(4) = MapReader.getInteger
39
40                                'Trigger used?
41                                If ByFlags And 16 Then .trigger = MapReader.getInteger
42
43                                '.inf file
44                                ByFlags = InfReader.getByte
45
46                                If ByFlags And 1 Then
47                                        .TileExit.Map = InfReader.getInteger
48                                        .TileExit.X = InfReader.getInteger
49                                        .TileExit.Y = InfReader.getInteger
50
51                                End If
52
53                                If ByFlags And 2 Then
54                                        'Get and make NPC
55                                        .NpcIndex = InfReader.getInteger
56
57                                        If .NpcIndex > 0 And .NpcIndex < 32000 Then
58
59                                                'Si el npc debe hacer respawn en la pos
60                                                'original la guardamos
61                                                If val(GetVar(npcfile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
62                                                        .NpcIndex = OpenNPC(.NpcIndex)
63                                                        Npclist(.NpcIndex).Orig.Map = Map
64                                                        Npclist(.NpcIndex).Orig.X = X
65                                                        Npclist(.NpcIndex).Orig.Y = Y
66                                                Else
67                                                        .NpcIndex = OpenNPC(.NpcIndex)
68
69                                                End If
70
If .NpcIndex < 32000 Then
71                                                Npclist(.NpcIndex).Pos.Map = Map
72                                                Npclist(.NpcIndex).Pos.X = X
73                                                Npclist(.NpcIndex).Pos.Y = Y
74
75                                                Call MakeNPCChar(True, 0, .NpcIndex, Map, X, Y)
76
End If
77                                        End If
78
79                                End If
80
81                                If ByFlags And 4 Then
82                                        'Get and make Object
83                                        .ObjInfo.objIndex = InfReader.getInteger
84                                        .ObjInfo.Amount = InfReader.getInteger
85
86                                End If
87
88                        End With
89
90                Next X
91        Next Y
92
93        Call Leer.Initialize(MAPFl & ".dat")
94
95        With MapInfo(Map)
96                .Name = Leer.GetValue("Mapa" & Map, "Name")
97                .Music = Leer.GetValue("Mapa" & Map, "MusicNum")
98                .StartPos.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
99                .StartPos.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
100                .StartPos.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        
101                .OnDeathGoTo.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
102                .OnDeathGoTo.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
103                .OnDeathGoTo.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        
104                .MagiaSinEfecto = val(Leer.GetValue("Mapa" & Map, "MagiaSinEfecto"))
105                .InviSinEfecto = val(Leer.GetValue("Mapa" & Map, "InviSinEfecto"))
106                .ResuSinEfecto = val(Leer.GetValue("Mapa" & Map, "ResuSinEfecto"))
107                .OcultarSinEfecto = val(Leer.GetValue("Mapa" & Map, "OcultarSinEfecto"))
108                .InvocarSinEfecto = val(Leer.GetValue("Mapa" & Map, "InvocarSinEfecto"))
        
109                .NoEncriptarMP = val(Leer.GetValue("Mapa" & Map, "NoEncriptarMP"))

110                .RoboNpcsPermitido = val(Leer.GetValue("Mapa" & Map, "RoboNpcsPermitido"))
        
111                If val(Leer.GetValue("Mapa" & Map, "Pk")) = 0 Then
112                        .Pk = True
113                Else
114                        .Pk = False

                End If
        
115                .Terreno = TerrainStringToByte(Leer.GetValue("Mapa" & Map, "Terreno"))
116                .Zona = Leer.GetValue("Mapa" & Map, "Zona")
117                .Restringir = RestrictStringToByte(Leer.GetValue("Mapa" & Map, "Restringir"))
118                .BackUp = val(Leer.GetValue("Mapa" & Map, "BACKUP"))

        End With
    
998        Set MapReader = Nothing
999        Set InfReader = Nothing
997        Set Leer = Nothing
    
996        Erase Buff
        Exit Sub

errh:
        
        'Call LogError("Error cargando mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.description & " - Linea: " & Erl)

        Set MapReader = Nothing
        Set InfReader = Nothing
        Set Leer = Nothing

End Sub

Sub LoadSini()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

On Error GoTo errhandler

        Dim Temporal As Long
    
        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    
        BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

        Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
        HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
        AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
        IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
        'Lee la version correcta del cliente
        ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
    
        PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
        ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))
    
        ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
        ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
        ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
        TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
        TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))
        ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
        ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
        ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
        TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
        TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))
    
        
        VestimentaImperialHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialHumano"))
        VestimentaImperialEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialEnano"))
        TunicaConspicuaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaHumano"))
        TunicaConspicuaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaEnano"))
        ArmaduraNobilisimaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaHumano"))
        ArmaduraNobilisimaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaEnano"))
        ArmaduraGranSacerdote = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraGranSacerdote"))
    
        VestimentaLegionHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionHumano"))
        VestimentaLegionEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionEnano"))
        TunicaLobregaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaHumano"))
        TunicaLobregaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaEnano"))
        TunicaEgregiaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaHumano"))
        TunicaEgregiaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaEnano"))
        SacerdoteDemoniaco = val(GetVar(IniPath & "Server.ini", "INIT", "SacerdoteDemoniaco"))

        EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))
    
        'Intervalos
        SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
        FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
    
        StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
        FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
    
        SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
        FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
    
        StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
        FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
    
        IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
        FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
        IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
        FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
    
        IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
        FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
    
        IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
        FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
    
        IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
        FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
    
        IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
        FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
    
        IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
        FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
    
        IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
        FrmInterv.txtInvocacion.Text = IntervaloInvocacion
    
        IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
        FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion
    
        '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
        IntervaloPuedeSerAtacado = 5000 ' Cargar desde balance.dat
        IntervaloAtacable = 60000 ' Cargar desde balance.dat
        IntervaloOwnedNpc = 18000 ' Cargar desde balance.dat
    
        IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
        FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
        frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
        FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval
    
        frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
        FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval
    
        IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
        FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
    
        IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
        FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
        'TODO : Agregar estos intervalos al form!!!
        IntervaloMagiaGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
        IntervaloGolpeMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))
        IntervaloGolpeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar"))
    
        frmMain.tLluvia.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
        FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
    
        MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))

        If MinutosWs < 60 Then MinutosWs = 180
    
        MinutosGuardarUsuarios = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGuardarUsuarios"))
    
        IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
        IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
        IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))
    
114         IntervaloOculto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))
    
        '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
      
113        RECORDusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "RECORD"))
      
        'Max users
112        Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))

        If MaxUsers = 0 Then
119                MaxUsers = Temporal
117                ReDim UserList(1 To MaxUsers) As User

        End If
    
        '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
        'Se agreg� en LoadBalance y en el Balance.dat
        'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))
    
        ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
668        Call Statistics.Initialize
    
667        Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
665        Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
656        Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    
55        Call MD5sCarga
    
57        Set ConsultaPopular = New ConsultasPopulares
56        Call ConsultaPopular.LoadData

        #If SeguridadAlkon Then
509                Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
        #End If
    
        ' Admins
5        Call loadAdministrativeUsers
4        CargarSponsor
        
        '@ Cui
3        Call CargarMalasPalabras
        
2        Call BanHD_load
        
        
        Exit Sub
errhandler:
        LogError "error en loadsini " & Erl
        Resume Next
End Sub

Sub WriteVar(ByVal file As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        'Escribe VAR en un archivo
        '***************************************************

        writeprivateprofilestring Main, Var, Value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, _
             ByVal UserFile As String, _
             Optional ByVal SaveTimeOnline As Boolean = True)
        '*************************************************
        'Author: Unknown
        'Last modified: 10/10/2010 (Pato)
        'Saves the Users RECORDs
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '11/19/2009: Pato - Save the EluSkills and ExpSkills
        '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
        '10/10/2010: Pato - Saco el WriteVar e implemento la clase clsIniManager
        '*************************************************

        On Error GoTo errhandler

        Dim Manager As clsIniManager

        Dim Existe  As Boolean

999        With UserList(UserIndex)

                'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
                'clase=0 es el error, porq el enum empieza de 1!!
998                If .clase = 0 Or .Stats.ELV = 0 Then
997                        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
996                        Exit Sub

995                End If
    
994                Set Manager = New clsIniManager
    
993                If FileExist(UserFile) Then
992                        Call Manager.Initialize(UserFile)
        
991                        If FileExist(UserFile & ".bk") Then Call Kill(UserFile & ".bk")
990                        Name UserFile As UserFile & ".bk"
        
989                        Existe = True

988                End If
    
987                If .flags.Mimetizado = 1 Then
986                        .Char.Body = .CharMimetizado.Body
                        .Char.Head = .CharMimetizado.Head
                        .Char.CascoAnim = .CharMimetizado.CascoAnim
                        .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                        .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                        .Counters.Mimetismo = 0
                        .flags.Mimetizado = 0
                        ' Se fue el efecto del mimetismo, puede ser atacado por npcs
                        .flags.Ignorado = False

                End If
    
                Dim LoopC As Integer
            
985
                Call Manager.ChangeValue("FLAGS", "Muerto", CStr(.flags.Muerto))
                Call Manager.ChangeValue("FLAGS", "Escondido", CStr(.flags.Escondido))
                Call Manager.ChangeValue("FLAGS", "Hambre", CStr(.flags.Hambre))
                Call Manager.ChangeValue("FLAGS", "Sed", CStr(.flags.Sed))
                Call Manager.ChangeValue("FLAGS", "Desnudo", CStr(.flags.Desnudo))
                Call Manager.ChangeValue("FLAGS", "Ban", CStr(.flags.Ban))
                Call Manager.ChangeValue("FLAGS", "Navegando", CStr(.flags.Navegando))
                Call Manager.ChangeValue("FLAGS", "Montando", CStr(.flags.Montando))
                Call Manager.ChangeValue("FLAGS", "Envenenado", CStr(.flags.Envenenado))
                Call Manager.ChangeValue("FLAGS", "Paralizado", CStr(.flags.Paralizado))
                Call Manager.ChangeValue("FLAGS", "PuntosShop", CStr(.flags.PuntosShop))
                
                'Matrix
                Call Manager.ChangeValue("FLAGS", "LastMap", CStr(.flags.lastMap))
    
                Call Manager.ChangeValue("CONSEJO", "PERTENECE", IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
                Call Manager.ChangeValue("CONSEJO", "PERTENECECAOS", IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))
    
                Call Manager.ChangeValue("COUNTERS", "Pena", CStr(.Counters.Pena))
   
                Call Manager.ChangeValue("FACCIONES", "EjercitoReal", CStr(.Faccion.ArmadaReal))
                Call Manager.ChangeValue("FACCIONES", "EjercitoCaos", CStr(.Faccion.FuerzasCaos))
                Call Manager.ChangeValue("FACCIONES", "CiudMatados", CStr(.Faccion.CiudadanosMatados))
                Call Manager.ChangeValue("FACCIONES", "CrimMatados", CStr(.Faccion.CriminalesMatados))
                Call Manager.ChangeValue("FACCIONES", "rArCaos", CStr(.Faccion.RecibioArmaduraCaos))
                Call Manager.ChangeValue("FACCIONES", "rArReal", CStr(.Faccion.RecibioArmaduraReal))
                Call Manager.ChangeValue("FACCIONES", "rExCaos", CStr(.Faccion.RecibioExpInicialCaos))
                Call Manager.ChangeValue("FACCIONES", "rExReal", CStr(.Faccion.RecibioExpInicialReal))
                Call Manager.ChangeValue("FACCIONES", "recCaos", CStr(.Faccion.RecompensasCaos))
                Call Manager.ChangeValue("FACCIONES", "recReal", CStr(.Faccion.RecompensasReal))
                Call Manager.ChangeValue("FACCIONES", "Reenlistadas", CStr(.Faccion.Reenlistadas))
                Call Manager.ChangeValue("FACCIONES", "NivelIngreso", CStr(.Faccion.NivelIngreso))
                Call Manager.ChangeValue("FACCIONES", "FechaIngreso", .Faccion.FechaIngreso)
                Call Manager.ChangeValue("FACCIONES", "MatadosIngreso", CStr(.Faccion.MatadosIngreso))
                Call Manager.ChangeValue("FACCIONES", "NextRecompensa", CStr(.Faccion.NextRecompensa))
    
                '�Fueron modificados los atributos del usuario?
                If Not .flags.TomoPocion Then

113                        For LoopC = 1 To UBound(.Stats.UserAtributos)
112                                Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributos(LoopC)))
111                        Next LoopC

110                Else

109                        For LoopC = 1 To UBound(.Stats.UserAtributos)
                                '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
108                                Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributosBackUP(LoopC)))
107                        Next LoopC

106                End If
    
105                For LoopC = 1 To UBound(.Stats.UserSkills)
104                        Call Manager.ChangeValue("SKILLS", "SK" & LoopC, CStr(.Stats.UserSkills(LoopC)))
103                Next LoopC
    
102                Call Manager.ChangeValue("CONTACTO", "Email", .email)
101                Call Manager.ChangeValue("CONTACTO", "Pin", .clave)
                
                
100                Call Manager.ChangeValue("INIT", "Genero", .Genero)
99                Call Manager.ChangeValue("INIT", "Raza", .raza)
98                Call Manager.ChangeValue("INIT", "Hogar", .Hogar)
97                Call Manager.ChangeValue("INIT", "Clase", .clase)
96                Call Manager.ChangeValue("INIT", "Desc", .desc)
    
95                Call Manager.ChangeValue("INIT", "Heading", CStr(.Char.heading))
94                Call Manager.ChangeValue("INIT", "Head", CStr(.OrigChar.Head))
    
93                If .flags.Muerto = 0 Then
92                        If .Char.Body <> 0 Then
91                                Call Manager.ChangeValue("INIT", "Body", CStr(.Char.Body))

90                        End If

89                End If
    
88                Call Manager.ChangeValue("INIT", "Arma", CStr(.Char.WeaponAnim))
87                Call Manager.ChangeValue("INIT", "Escudo", CStr(.Char.ShieldAnim))
86                Call Manager.ChangeValue("INIT", "Casco", CStr(.Char.CascoAnim))
    
                #If ConUpTime Then
    
                        If SaveTimeOnline Then

                                Dim TempDate As Date

                                TempDate = Now - .LogOnTime
                                .LogOnTime = Now
                                .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
                                .UpTime = .UpTime
                                Call Manager.ChangeValue("INIT", "UpTime", .UpTime)

                        End If

                #End If
    
                'First time around?
85                If LenB(Manager.GetValue("INIT", "LastIP1")) = 0 Then
84                        Call Manager.ChangeValue("INIT", "LastIP1", .Ip & " - " & Date & ":" & time)
                        'Is it a different ip from last time?
83                ElseIf .Ip <> Left$(Manager.GetValue("INIT", "LastIP1"), InStr(1, Manager.GetValue("INIT", "LastIP1"), " ") - 1) Then

                        Dim i As Integer

82                        For i = 5 To 2 Step -1
81                                Call Manager.ChangeValue("INIT", "LastIP" & i, Manager.GetValue("INIT", "LastIP" & CStr(i - 1)))
80                        Next i

79                        Call Manager.ChangeValue("INIT", "LastIP1", .Ip & " - " & Date & ":" & time)
                        'Same ip, just update the date
                Else
78                        Call Manager.ChangeValue("INIT", "LastIP1", .Ip & " - " & Date & ":" & time)

                End If
    
77                Call Manager.ChangeValue("INIT", "Position", .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y)
    
76                Call Manager.ChangeValue("STATS", "GLD", CStr(.Stats.GLD))
75                Call Manager.ChangeValue("STATS", "BANCO", CStr(.Stats.Banco))
    
74                Call Manager.ChangeValue("STATS", "MaxHP", CStr(.Stats.MaxHP))
73                Call Manager.ChangeValue("STATS", "MinHP", CStr(.Stats.MinHp))
    
72                Call Manager.ChangeValue("STATS", "MaxSTA", CStr(.Stats.MaxSta))
71                Call Manager.ChangeValue("STATS", "MinSTA", CStr(.Stats.MinSta))
    
70                Call Manager.ChangeValue("STATS", "MaxMAN", CStr(.Stats.MaxMAN))
69                Call Manager.ChangeValue("STATS", "MinMAN", CStr(.Stats.MinMAN))
    
68                Call Manager.ChangeValue("STATS", "MaxHIT", CStr(.Stats.MaxHIT))
67                Call Manager.ChangeValue("STATS", "MinHIT", CStr(.Stats.MinHIT))
    
66                Call Manager.ChangeValue("STATS", "MaxAGU", CStr(.Stats.MaxAGU))
65                Call Manager.ChangeValue("STATS", "MinAGU", CStr(.Stats.MinAGU))
    
64                Call Manager.ChangeValue("STATS", "MaxHAM", CStr(.Stats.MaxHam))
63                Call Manager.ChangeValue("STATS", "MinHAM", CStr(.Stats.MinHam))
    
62                Call Manager.ChangeValue("STATS", "SkillPtsLibres", CStr(.Stats.SkillPts))
      
61                Call Manager.ChangeValue("STATS", "EXP", CStr(.Stats.Exp))
60                Call Manager.ChangeValue("STATS", "ELV", CStr(.Stats.ELV))
    
59                Call Manager.ChangeValue("STATS", "ELU", CStr(.Stats.ELU))
58                Call Manager.ChangeValue("MUERTES", "UserMuertes", CStr(.Stats.UsuariosMatados))
                'Call Manager.ChangeValue( "MUERTES", "CrimMuertes", CStr(.Stats.CriminalesMatados))
57                Call Manager.ChangeValue("MUERTES", "NpcsMuertes", CStr(.Stats.NPCsMuertos))
      
                '[KEVIN]----------------------------------------------------------------------------
                '*******************************************************************************************
56                Call Manager.ChangeValue("BancoInventory", "CantidadItems", val(.BancoInvent.NroItems))

                Dim loopd As Integer

55                For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
54                        Call Manager.ChangeValue("BancoInventory", "Obj" & loopd, .BancoInvent.Object(loopd).objIndex & "-" & .BancoInvent.Object(loopd).Amount)
53                Next loopd

                '*******************************************************************************************
                '[/KEVIN]-----------
      
                'Save Inv
52                Call Manager.ChangeValue("Inventory", "CantidadItems", val(.Invent.NroItems))
    
51                For LoopC = 1 To MAX_NORMAL_INVENTORY_SLOTS
50                        Call Manager.ChangeValue("Inventory", "Obj" & LoopC, .Invent.Object(LoopC).objIndex & "-" & .Invent.Object(LoopC).Amount & "-" & .Invent.Object(LoopC).Equipped)
49                Next LoopC
    
48                Call Manager.ChangeValue("Inventory", "WeaponEqpSlot", CStr(.Invent.WeaponEqpSlot))
47                Call Manager.ChangeValue("Inventory", "ArmourEqpSlot", CStr(.Invent.ArmourEqpSlot))
46                Call Manager.ChangeValue("Inventory", "CascoEqpSlot", CStr(.Invent.CascoEqpSlot))
45                Call Manager.ChangeValue("Inventory", "EscudoEqpSlot", CStr(.Invent.EscudoEqpSlot))
44                Call Manager.ChangeValue("Inventory", "BarcoSlot", CStr(.Invent.BarcoSlot))
43                Call Manager.ChangeValue("Inventory", "MunicionSlot", CStr(.Invent.MunicionEqpSlot))
                '/Nacho
    
42                Call Manager.ChangeValue("Inventory", "AnilloSlot", CStr(.Invent.AnilloEqpSlot))
    
                'Reputacion
41                Call Manager.ChangeValue("REP", "Asesino", CStr(.Reputacion.AsesinoRep))
40                Call Manager.ChangeValue("REP", "Bandido", CStr(.Reputacion.BandidoRep))
39                Call Manager.ChangeValue("REP", "Burguesia", CStr(.Reputacion.BurguesRep))
38                Call Manager.ChangeValue("REP", "Ladrones", CStr(.Reputacion.LadronesRep))
37                Call Manager.ChangeValue("REP", "Nobles", CStr(.Reputacion.NobleRep))
36                Call Manager.ChangeValue("REP", "Plebe", CStr(.Reputacion.PlebeRep))
    
                Dim L As Long

35                L = (-.Reputacion.AsesinoRep) + _
                   (-.Reputacion.BandidoRep) + _
                   .Reputacion.BurguesRep + _
                   (-.Reputacion.LadronesRep) + _
                   .Reputacion.NobleRep + _
                   .Reputacion.PlebeRep
34                L = L / 6
33                Call Manager.ChangeValue("REP", "Promedio", CStr(L))
    
    
        Call SaveQuestStats(UserIndex, Manager)    ' GSZAO

                Dim cad As String
    
32                For LoopC = 1 To MAXUSERHECHIZOS
31                        cad = .Stats.UserHechizos(LoopC)
30                        Call Manager.ChangeValue("HECHIZOS", "H" & LoopC, cad)
29                Next
    
                Dim NroMascotas As Long

28                NroMascotas = .NroMascotas
    
27                For LoopC = 1 To MAXMASCOTAS

                        ' Mascota valida?
26                        If .MascotasIndex(LoopC) > 0 Then

                                ' Nos aseguramos que la criatura no fue invocada
25                                If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
24                                        cad = .MascotasType(LoopC)
23                                Else 'Si fue invocada no la guardamos
22                                        cad = "0"
21                                        NroMascotas = NroMascotas - 1

20                                End If

19                                Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)
18                        Else
17                                cad = .MascotasType(LoopC)
16                                Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)

                        End If
15
14                Next
    
13                Call Manager.ChangeValue("MASCOTAS", "NroMascotas", CStr(NroMascotas))
    
                'Guarda los mensajes privados del usuario.
12                Call GuardarMensajes(UserIndex, Manager)
    
                'Devuelve el head de muerto
11                If .flags.Muerto = 1 Then
10                        .Char.Head = iCabezaMuerto

9                End If
                
8                Call Manager.ChangeValue("MERCADO", "POSTEADO", .flags.Posteado)
7                Call Manager.ChangeValue("MERCADO", "MEMANDO", .flags.MeMando)
6                Call Manager.ChangeValue("MERCADO", "LEMANDE", .flags.Lemande)
                
5                Call Manager.ChangeValue("DISCODURO", "VAL", .discoDuro)
                
                 Call Manager.ChangeValue("FLAGS", "Vip", CStr(.flags.Vip))
4        End With

3        Call Manager.DumpFile(UserFile)

2        Set Manager = Nothing

1        If Existe Then Call Kill(UserFile & ".bk")
        
        Exit Sub

errhandler:
        Call LogError("Error en SaveUser en linea " & Erl & " - " & Err.Number & " " & Err.description)
        Set Manager = Nothing

End Sub

Function criminal(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim L As Long
    
        With UserList(UserIndex).Reputacion
                L = (-.AsesinoRep) + _
                   (-.BandidoRep) + _
                   .BurguesRep + _
                   (-.LadronesRep) + _
                   .NobleRep + _
                   .PlebeRep
                L = L / 6
                criminal = (L < 0)

        End With

End Function

Sub BackUPnPc(ByVal NpcIndex As Integer, ByVal hFile As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 10/09/2010
        '10/09/2010 - Pato: Optimice el BackUp de NPCs
        '***************************************************

        Dim LoopC As Integer
    
        Print #hFile, "[NPC" & Npclist(NpcIndex).Numero & "]"
    
        With Npclist(NpcIndex)
        
                'General
                Print #hFile, "Name=" & .Name
                Print #hFile, "Desc=" & .desc
                Print #hFile, "Head=" & val(.Char.Head)
                Print #hFile, "Body=" & val(.Char.Body)
                Print #hFile, "Heading=" & val(.Char.heading)
                Print #hFile, "Casco=" & val(.Char.CascoAnim)
                Print #hFile, "Escudo=" & val(.Char.ShieldAnim)
                Print #hFile, "Arma=" & val(.Char.WeaponAnim)
                Print #hFile, "Movement=" & val(.Movement)
                Print #hFile, "Attackable=" & val(.Attackable)
                Print #hFile, "Comercia=" & val(.Comercia)
                Print #hFile, "TipoItems=" & val(.TipoItems)
                Print #hFile, "Hostil=" & val(.Hostile)
                Print #hFile, "GiveEXP=" & val(.GiveEXP)
                Print #hFile, "GiveGLD=" & val(.GiveGLD)
                Print #hFile, "InvReSpawn=" & val(.InvReSpawn)
                Print #hFile, "NpcType=" & val(.NPCtype)
        
                'Stats
                Print #hFile, "Alineacion=" & val(.Stats.Alineacion)
                Print #hFile, "DEF=" & val(.Stats.def)
                Print #hFile, "MaxHit=" & val(.Stats.MaxHIT)
                Print #hFile, "MaxHp=" & val(.Stats.MaxHP)
                Print #hFile, "MinHit=" & val(.Stats.MinHIT)
                Print #hFile, "MinHp=" & val(.Stats.MinHp)
        
                'Flags
                Print #hFile, "ReSpawn=" & val(.flags.respawn)
                Print #hFile, "BackUp=" & val(.flags.BackUp)
                Print #hFile, "Domable=" & val(.flags.Domable)
        
                'Inventario
                Print #hFile, "NroItems=" & val(.Invent.NroItems)

                If .Invent.NroItems > 0 Then

                        For LoopC = 1 To .Invent.NroItems
                                Print #hFile, "Obj" & LoopC & "=" & .Invent.Object(LoopC).objIndex & "-" & .Invent.Object(LoopC).Amount
                        Next LoopC

                End If
        
                Print #hFile, vbNullString

        End With

End Sub

Sub CargarNpcBackUp(ByVal NpcIndex As Integer, ByVal NpcNumber As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        'Status
        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"
    
        Dim npcfile As String
    
        'If NpcNumber > 499 Then
        '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
        'Else
        npcfile = DatPath & "bkNPCs.dat"
        'End If
    
        With Npclist(NpcIndex)
    
                .Numero = NpcNumber
                .Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
                .desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
                .Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
                .NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
        
                .Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
                .Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
                .Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
                
                .Char.CascoAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "Casco"))
                .Char.ShieldAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "Escudo"))
                .Char.WeaponAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "Arma"))
        
                .Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
                .Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
                .Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
                .GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))
        
                .GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
        
                .InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
        
                .Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
                .Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
                .Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
                .Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
                .Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
                .Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
        
                Dim LoopC As Integer

                Dim ln    As String

                .Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))

                If .Invent.NroItems > 0 Then

                        For LoopC = 1 To MAX_NORMAL_INVENTORY_SLOTS
                                ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
                                .Invent.Object(LoopC).objIndex = val(ReadField(1, ln, 45))
                                .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
               
                        Next LoopC

                Else

                        For LoopC = 1 To MAX_NORMAL_INVENTORY_SLOTS
                                .Invent.Object(LoopC).objIndex = 0
                                .Invent.Object(LoopC).Amount = 0
                        Next LoopC

                End If
        
                For LoopC = 1 To MAX_NPC_DROPS
                        ln = GetVar(npcfile, "NPC" & NpcNumber, "Drop" & LoopC)
                        .Drop(LoopC).objIndex = val(ReadField(1, ln, 45))
                        .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
                Next LoopC
        
                .flags.NPCActive = True
                .flags.respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
                .flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
                .flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
                .flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
        
                'Tipo de items con los que comercia
                .TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

        End With

End Sub

Sub LogBan(ByVal BannedIndex As Integer, _
           ByVal UserIndex As Integer, _
           ByVal Motivo As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
        Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)
    
        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

        mifile = FreeFile
        Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
        Print #mifile, UserList(BannedIndex).Name
        Close #mifile

End Sub

Sub LogBanFromName(ByVal BannedName As String, _
                   ByVal UserIndex As Integer, _
                   ByVal Motivo As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

        mifile = FreeFile
        Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
        Print #mifile, BannedName
        Close #mifile

End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

        mifile = FreeFile
        Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
        Print #mifile, BannedName
        Close #mifile

End Sub

Public Sub CargaApuestas()
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
        Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
        Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub

Public Sub LoadArmadurasFaccion()

        '***************************************************
        'Author: ZaMa
        'Last Modification: 15/04/2010
        '
        '***************************************************
        Dim ClassIndex    As Long
    
        Dim ArmaduraIndex As Integer
    
        For ClassIndex = 1 To NUMCLASES
    
                ' Defensa minima para armadas altos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyAlto"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
                ' Defensa minima para armadas bajos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyBajo"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
                ' Defensa minima para caos altos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosAlto"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
                ' Defensa minima para caos bajos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosBajo"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
    
                ' Defensa media para armadas altos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyAlto"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
                ' Defensa media para armadas bajos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyBajo"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
                ' Defensa media para caos altos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosAlto"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
                ' Defensa media para caos bajos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosBajo"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
    
                ' Defensa alta para armadas altos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyAlto"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
                ' Defensa alta para armadas bajos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyBajo"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
                ' Defensa alta para caos altos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosAlto"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
                ' Defensa alta para caos bajos
                ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosBajo"))
        
                ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
                ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
    
        Next ClassIndex
    
End Sub

Public Sub LoadAnimations()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 11/06/2011
        '
        '***************************************************
        AnimHogar(eHeading.NORTH) = 40
        AnimHogar(eHeading.EAST) = 42
        AnimHogar(eHeading.SOUTH) = 39
        AnimHogar(eHeading.WEST) = 41
    
        AnimHogarNavegando(eHeading.NORTH) = 44
        AnimHogarNavegando(eHeading.EAST) = 46
        AnimHogarNavegando(eHeading.SOUTH) = 43
        AnimHogarNavegando(eHeading.WEST) = 45

End Sub

Public Sub CargarMalasPalabras()

'Copied from CargarForbidenWords xd alta paja

        ReDim BadWords(1 To TxtDimension(DatPath & "malaspalabras.txt"))

        Dim N As Integer, i As Integer

        N = FreeFile(1)
        Open DatPath & "malaspalabras.txt" For Input As #N
    
        For i = 1 To UBound(BadWords)
                Line Input #N, BadWords(i)
        Next i
    
        Close N

End Sub
