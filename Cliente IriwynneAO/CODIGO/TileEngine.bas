Attribute VB_Name = "Mod_TileEngine"

Option Explicit

'Map sizes in tiles
Public Const XMaxMapSize     As Byte = 100
Public Const XMinMapSize     As Byte = 1
Public Const YMaxMapSize     As Byte = 100
Public Const YMinMapSize     As Byte = 1

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

Private Const GrhFogata      As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Encabezado bmp
Type BITMAPFILEHEADER

        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long

End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER

        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long

End Type

'Posicion en un mapa
Public Type Position

        X As Long
        Y As Long

End Type

'Posicion en el Mundo
Public Type WorldPos

        Map As Integer
        X As Integer
        Y As Integer

End Type

'Contiene info acerca de donde se puede encontrar un grh tama�o y animacion
Public Type GrhData

        sX As Integer
        sY As Integer
    
        FileNum As Long
    
        pixelWidth As Integer
        pixelHeight As Integer
    
        TileWidth As Single
        TileHeight As Single
    
        NumFrames As Integer
        Frames() As Long
    
        Speed As Single

End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh

        GrhIndex As Integer
        FrameCounter As Single
        Speed As Single
        Started As Byte
        Loops As Integer
        

End Type

'Lista de cuerpos
Public Type BodyData

        Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
        HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

        Head(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

        WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

        ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Apariencia del personaje
Public Type Char
        EsDonador As Byte
        Min_HP As Long
        Max_Hp As Long
        esNPC As Byte
        Active As Byte
        Heading As E_Heading
        Pos As Position
    
        iHead As Integer
        iBody As Integer
        Body As BodyData
        Head As HeadData
        Casco As HeadData
        Arma As WeaponAnimData
        Escudo As ShieldAnimData
        UsandoArma As Boolean
    
        fX As Grh
        FxIndex As Integer
    
        Criminal As Byte
        Atacable As Boolean
        EsGranPoder As Boolean
        
        Nombre As String
    
        scrollDirectionX As Integer
        scrollDirectionY As Integer
    
        Moving As Byte
        MoveOffsetX As Single
        MoveOffsetY As Single
    
        pie As Boolean
        muerto As Boolean
        invisible As Boolean
        priv As Byte
        Conquisto As Boolean

End Type

'Info de un objeto
Public Type Obj

        ObjIndex As Integer
        Amount As Integer

End Type

'Tipo de las celdas del mapa
Public Type MapBlock

        Graphic(1 To 4) As Grh
        CharIndex As Integer
        ObjGrh As Grh
            
        NpcIndex As Integer
        OBJInfo As Obj
        TileExit As WorldPos
        Blocked As Byte
    
        Trigger As Integer
        Damage As DList

End Type

'Info de cada mapa
Public Type MapInfo

        Music As String
        Name As String
        StartPos As WorldPos
        MapVersion As Integer

End Type

'DX7 Objects
Public DirectX                 As DirectX7
Public DirectDraw              As DirectDraw7

Public PrimarySurface         As DirectDrawSurface7

Public PrimaryClipper         As DirectDrawClipper

Public BackBufferSurface      As DirectDrawSurface7

'Bordes del mapa
Public MinXBorder              As Byte
Public MaxXBorder              As Byte
Public MinYBorder              As Byte
Public MaxYBorder              As Byte

'Status del user
Public CurMap                  As Integer 'Mapa actual

Public UserIndex               As Integer

Public UserMoving              As Byte

Public UserBody                As Integer

Public UserHead                As Integer

Public UserPos                 As Position 'Posicion

Public AddtoUserPos            As Position 'Si se mueve

Public UserCharIndex           As Integer

Public EngineRun               As Boolean

Public FPS                     As Long

Public FramesPerSecCounter     As Long

Private fpsLastCheck           As Long

'Tama�o del la vista en Tiles
Private WindowTileWidth        As Integer

Private WindowTileHeight       As Integer

Private HalfWindowTileWidth    As Integer

Private HalfWindowTileHeight   As Integer

'Offset del desde 0,0 del main view
Private MainViewTop            As Integer

Private MainViewLeft           As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tama�o muy grande puede
'volver el engine muy lento
Public TileBufferSize          As Integer

Private TileBufferPixelOffsetX As Integer

Private TileBufferPixelOffsetY As Integer

'Tama�o de los tiles en pixels
Public TilePixelHeight         As Integer

Public TilePixelWidth          As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX   As Integer

Public ScrollPixelsPerFrameY   As Integer

Dim timerElapsedTime           As Single

Dim timerTicksPerFrame         As Single

Dim engineBaseSpeed            As Single

Public NumBodies               As Integer

Public Numheads                As Integer

Public NumFxs                  As Integer

Public NumChars                As Integer

Public LastChar                As Integer

Public NumWeaponAnims          As Integer

Public NumShieldAnims          As Integer

Public MainDestRect           As RECT

Public MainViewRect           As RECT

Public BackBufferRect         As RECT

Public MainViewWidth          As Integer

Public MainViewHeight         As Integer

'�?�?�?�?�?�?�?�?�?�Graficos�?�?�?�?�?�?�?�?�?�?�?
Public GrhData()               As GrhData 'Guarda todos los grh

Public BodyData()              As BodyData

Public HeadData()              As HeadData

Public FxData()                As tIndiceFx

Public WeaponAnimData()        As WeaponAnimData

Public ShieldAnimData()        As ShieldAnimData

Public CascoAnimData()         As HeadData
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�Mapa?�?�?�?�?�?�?�?�?�?�?�?
Public MapData()               As MapBlock ' Mapa

Public MapInfo                 As MapInfo ' Info acerca del mapa en uso
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

Public bRain                   As Boolean 'est� raineando?

Public bTecho                  As Boolean 'hay techo?

Public brstTick                As Long

Private RLluvia(7)             As RECT  'RECT de la lluvia

Private iFrameIndex            As Byte  'Frame actual de la LL

Private llTick                 As Long  'Contador

Private LTLluvia(4)            As Integer

Public charlist(1 To 10000)    As Char

' Used by GetTextExtentPoint32
Private Type size

        cx As Long
        cy As Long

End Type

'[CODE 001]:MatuX
Public Enum PlayLoop

        plNone = 0
        plLluviain = 1
        plLluviaout = 2

End Enum

'[END]'
'
'       [END]
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 _
                Lib "gdi32" _
                Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                               ByVal lpsz As String, _
                                               ByVal cbString As Long, _
                                               lpSize As size) As Long

Private Declare Function SetPixel _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long

'Text width computation. Needed to center text.kidral

Private wcol         As Integer

Private method       As Boolean

Private scol(0 To 3) As Integer

Private subidor      As Boolean

Private aldus        As Integer

Private aldstep      As Integer

Private porcual      As Integer

Private Function alduath() As String

        Dim elnombre As String

        Dim templine As String

        elnombre = "kidral"

        If aldstep >= 1 Then
                templine = Left$(elnombre, aldstep)

        End If

        'Aldstep es cuantas letras ya hay en su lugar, por ende me dice la letra nueva
        '(7-aldstep) es la cantidad de espacios
        templine = templine & Space$(7 - aldstep - porcual) & mid$(elnombre, aldstep + 1, 1) & Space$(porcual)
        aldus = aldus + 1

        If aldus >= 10 Then
                If aldstep = 7 Then
                        If aldus = 100 Then
                                aldstep = 0

                        End If

                Else
                        porcual = porcual + 1
                        aldus = 0

                        If (7 - aldstep - porcual) = 0 Then
                                aldstep = aldstep + 1
                                porcual = 0

                        End If

                End If

        End If

        alduath = templine

End Function

Private Function MABColor() As Long

        On Error Resume Next

        MABColor = RGB(255, RandomNumber(150, 255), 0)

End Function

Private Function betaniaColor() As Long

        On Error Resume Next

        betaniaColor = RGB(252, 150, 177)

End Function

Private Function AfroditaColor() As Long

        On Error Resume Next
            
        AfroditaColor = RGB(247, 15, 206)

End Function
Private Function ACEColor() As Long

        On Error Resume Next
            
        ACEColor = RGB(68, 221, 178)

End Function
Private Function kidralColor() As Long

        On Error Resume Next

        method = Not method
   
        If method Then
   
                If subidor Then
                        scol(wcol) = scol(wcol) + 1

                        If scol(wcol) = 255 Then wcol = wcol + 1
                Else
                        scol(wcol) = scol(wcol) - 1

                        If scol(wcol) = 0 Then wcol = wcol + 1

                End If
    
                If wcol = 3 Then
                        subidor = Not subidor
                        wcol = 0

                End If
            
                '    kidralColor = RGB(scol(0), scol(1), scol(2))
        Else
                kidralColor = RandomNumber(0, &HFFFFFF)

        End If

End Function

Sub CargarCabezas()

        Dim N            As Integer

        Dim i            As Long

        Dim Numheads     As Integer

        Dim Miscabezas() As tIndiceCabeza
    
        N = FreeFile()
        Open DirInit & "Cabezas.ind" For Binary Access Read As #N
    
        'cabecera
        Get #N, , MiCabecera
    
        'num de cabezas
        Get #N, , Numheads
    
        'Resize array
        ReDim HeadData(0 To Numheads) As HeadData
        ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
        For i = 1 To Numheads
                Get #N, , Miscabezas(i)
        
                If Miscabezas(i).Head(1) Then
                        Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
                        Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
                        Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
                        Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)

                End If

        Next i
    
        Close #N

End Sub

Sub CargarCascos()

        Dim N            As Integer

        Dim i            As Long

        Dim NumCascos    As Integer

        Dim Miscabezas() As tIndiceCabeza
    
        N = FreeFile()
        Open DirInit & "Cascos.ind" For Binary Access Read As #N
    
        'cabecera
        Get #N, , MiCabecera
    
        'num de cabezas
        Get #N, , NumCascos
    
        'Resize array
        ReDim CascoAnimData(0 To NumCascos) As HeadData
        ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
        For i = 1 To NumCascos
                Get #N, , Miscabezas(i)
        
                If Miscabezas(i).Head(1) Then
                        Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
                        Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
                        Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
                        Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)

                End If

        Next i
    
        Close #N

End Sub

Sub CargarCuerpos()

        Dim N            As Integer

        Dim i            As Long

        Dim NumCuerpos   As Integer

        Dim MisCuerpos() As tIndiceCuerpo
    
        N = FreeFile()
        Open DirInit & "Personajes.ind" For Binary Access Read As #N
    
        'cabecera
        Get #N, , MiCabecera
    
        'num de cabezas
        Get #N, , NumCuerpos
    
        'Resize array
        ReDim BodyData(0 To NumCuerpos) As BodyData
        ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
        For i = 1 To NumCuerpos
                Get #N, , MisCuerpos(i)
        
                If MisCuerpos(i).Body(1) Then
                        InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
                        InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
                        InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
                        InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
                        BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
                        BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY

                End If

        Next i
    
        Close #N

End Sub

Sub CargarFxs()

        Dim N      As Integer

        Dim i      As Long

        Dim NumFxs As Integer
    
        N = FreeFile()
        Open DirInit & "Fxs.ind" For Binary Access Read As #N
    
        'cabecera
        Get #N, , MiCabecera
    
        'num de cabezas
        Get #N, , NumFxs
    
        'Resize array
        ReDim FxData(1 To NumFxs) As tIndiceFx
    
        For i = 1 To NumFxs
                Get #N, , FxData(i)
        Next i
    
        Close #N

End Sub

Sub CargarArrayLluvia()

        Dim N  As Integer

        Dim i  As Long

        Dim Nu As Integer
    
        N = FreeFile()
        Open DirInit & "fk.ind" For Binary Access Read As #N
    
        'cabecera
        Get #N, , MiCabecera
    
        'num de cabezas
        Get #N, , Nu
    
        'Resize array
        ReDim bLluvia(1 To Nu) As Byte
    
        For i = 1 To Nu
                Get #N, , bLluvia(i)
        Next i
    
        Close #N

End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, _
                  ByVal viewPortY As Integer, _
                  ByRef tx As Byte, _
                  ByRef TY As Byte)
                  
        '******************************************
        'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
        '******************************************
        
        ' @@ Miqueas : Fix bug, a veces UserPos = 0, y explotaba todo a la mierda :P - 07/11/15

        If InMapBounds(UserPos.X, UserPos.Y) Then
                tx = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
                TY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2

                Exit Sub

        End If

        ' @@ Parche a un posible error de mierda
        'LogError "InMapBounds(CurrentUser.UserPos.X, CurrentUser.UserPos.Y) = false", "ConvertCPtoTP"
End Sub

Sub MakeChar(ByVal CharIndex As Integer, _
             ByVal Body As Integer, _
             ByVal Head As Integer, _
             ByVal Heading As Byte, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal Arma As Integer, _
             ByVal Escudo As Integer, _
             ByVal Casco As Integer)

        On Error Resume Next

        'Apuntamos al ultimo Char
        If CharIndex > LastChar Then LastChar = CharIndex
    
        With charlist(CharIndex)

                'If the char wasn't allready active (we are rewritting it) don't increase char count
                If .Active = 0 Then _
                   NumChars = NumChars + 1
        
                If Arma = 0 Then Arma = 2
                If Escudo = 0 Then Escudo = 2
                If Casco = 0 Then Casco = 2
        
                .iHead = Head
                .iBody = Body
                .Head = HeadData(Head)
                .Body = BodyData(Body)
                .Arma = WeaponAnimData(Arma)
        
                .Escudo = ShieldAnimData(Escudo)
                .Casco = CascoAnimData(Casco)
               
                .Heading = Heading
        
                'Reset moving stats
                .Moving = 0
                .MoveOffsetX = 0
                .MoveOffsetY = 0
        
                'Update position
                .Pos.X = X
                .Pos.Y = Y
        
                'Make active
                .Active = 1

        End With
    
        'Plot on map
        MapData(X, Y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

        With charlist(CharIndex)
                .Active = 0
                .Criminal = 0
                .Atacable = False
                .FxIndex = 0
                .invisible = False
                .Moving = 0
                .Conquisto = False
                .muerto = False
                
                ' @@fixeos?
                .EsDonador = 0
                .EsGranPoder = False
                .esNPC = 0
                
                '.Nombre = vbNullString
                .pie = False
                .Pos.X = 0
                .Pos.Y = 0
                .UsandoArma = False

        End With

End Sub

Sub EraseChar(ByVal CharIndex As Integer)

        '*****************************************************************
        'Erases a character from CharList and map
        '*****************************************************************
        'On Error Resume Next ' @@ Miqueas : 07/11 - No hace falta

        charlist(CharIndex).Active = 0
    
        'Update lastchar
        If CharIndex = LastChar Then

                Do Until charlist(LastChar).Active = 1
                        LastChar = LastChar - 1

                        If LastChar = 0 Then Exit Do
                Loop

        End If

        ' @@ Miqueas 07/11/15 - Soluciono Error, y posible bug de clones
        Dim X As Long

        Dim Y As Long

        For X = 1 To 100
                For Y = 1 To 100

                        If MapData(X, Y).CharIndex = CharIndex Then
                                MapData(X, Y).CharIndex = 0

                                Exit For

                        End If

                Next Y
        Next X

        'If InMapBounds(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y) Then
        '        MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).CharIndex = 0
        'Else
        '        LogError "UserChar Pos x - y = 0", "EraseChar"
        'End If
    
        'Remove char's dialog
        Call Dialogos.RemoveDialog(CharIndex)
    
        Call ResetCharInfo(CharIndex)
    
        'Update NumChars
        NumChars = NumChars - 1

End Sub

Public Sub InitGrh(ByRef Grh As Grh, _
                   ByVal GrhIndex As Integer, _
                   Optional ByVal Started As Byte = 2)

        '*****************************************************************
        'Sets up a grh. MUST be done before rendering
        '*****************************************************************
    
        
        If Not GrhIndex <> 0 Then Exit Sub
        
        Grh.GrhIndex = GrhIndex
    
        If Started = 2 Then
                If GrhData(Grh.GrhIndex).NumFrames > 1 Then
                        Grh.Started = 1
                Else
                        Grh.Started = 0

                End If

        Else

                'Make sure the graphic can be started
                If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
                Grh.Started = Started

        End If
    
        If Grh.Started Then
                Grh.Loops = INFINITE_LOOPS
        Else
                Grh.Loops = 0

        End If
    
        Grh.FrameCounter = 1
        Grh.Speed = GrhData(Grh.GrhIndex).Speed

End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)

        '*****************************************************************
        'Starts the movement of a character in nHeading direction
        '*****************************************************************
        Dim addX As Integer

        Dim addY As Integer

        Dim X    As Integer

        Dim Y    As Integer

        Dim nX   As Integer

        Dim nY   As Integer
    
        With charlist(CharIndex)
                X = .Pos.X
                Y = .Pos.Y
        
                'Figure out which way to move
                Select Case nHeading

                        Case E_Heading.NORTH
                                addY = -1
        
                        Case E_Heading.EAST
                                addX = 1
        
                        Case E_Heading.SOUTH
                                addY = 1
            
                        Case E_Heading.WEST
                                addX = -1

                End Select
        
                nX = X + addX
                nY = Y + addY
        
                MapData(nX, nY).CharIndex = CharIndex
                .Pos.X = nX
                .Pos.Y = nY
                MapData(X, Y).CharIndex = 0
        
                .MoveOffsetX = -1 * (TilePixelWidth * addX)
                .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
                .Moving = 1
                .Heading = nHeading
        
                .scrollDirectionX = addX
                .scrollDirectionY = addY

        End With
    
        If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
        'areas viejos
        If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
                If CharIndex <> UserCharIndex Then
                        Call EraseChar(CharIndex)

                End If

        End If

End Sub

Public Sub DoFogataFx()

        Dim location As Position
    
        If bFogata Then
                bFogata = HayFogata(location)

                If Not bFogata Then
                        Call Audio.StopWave(FogataBufferIndex)
                        FogataBufferIndex = 0

                End If

        Else
                bFogata = HayFogata(location)

                If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)

        End If

End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean

        '***************************************************
        'Author: Unknown
        'Last Modification: 09/21/2010
        ' 09/21/2010: C4b3z0n - Changed from Private Funtion tu Public Function.
        '***************************************************
        With charlist(CharIndex).Pos
                EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder

        End With

End Function

Sub DoPasosFx(ByVal CharIndex As Integer)

        If Not UserNavegando Then

                With charlist(CharIndex)

                        If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                                .pie = Not .pie
                
                                If .pie Then
                                        Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                                Else
                                        Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)

                                End If

                        End If

                End With

        Else
                ' TODO : Actually we would have to check if the CharIndex char is in the water or not....
                Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)

        End If

End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

        'On Error Resume Next ' @@ Miqueas 07/11/15 - Ya no hace falta

        Dim X        As Integer

        Dim Y        As Integer

        Dim addX     As Integer

        Dim addY     As Integer

        Dim nHeading As E_Heading
    
        With charlist(CharIndex)
                X = .Pos.X
                Y = .Pos.Y
        
                If InMapBounds(X, Y) Then ' @@ Miqueas 07/11/15 Fix bug
                        MapData(X, Y).CharIndex = 0

                End If
        
                addX = nX - X
                addY = nY - Y
        
                If Sgn(addX) = 1 Then
                        nHeading = E_Heading.EAST
                ElseIf Sgn(addX) = -1 Then
                        nHeading = E_Heading.WEST
                ElseIf Sgn(addY) = -1 Then
                        nHeading = E_Heading.NORTH
                ElseIf Sgn(addY) = 1 Then
                        nHeading = E_Heading.SOUTH

                End If
        
                MapData(nX, nY).CharIndex = CharIndex
        
                .Pos.X = nX
                .Pos.Y = nY
        
                .MoveOffsetX = -1 * (TilePixelWidth * addX)
                .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
                .Moving = 1
                .Heading = nHeading
        
                .scrollDirectionX = Sgn(addX)
                .scrollDirectionY = Sgn(addY)
        
                'parche para que no medite cuando camina
                If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
                        .FxIndex = 0

                End If

        End With
    
        If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
        If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
                Call EraseChar(CharIndex)

        End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)

        '******************************************
        'Starts the screen moving in a direction
        '******************************************
        Dim X  As Integer

        Dim Y  As Integer

        Dim tx As Integer

        Dim TY As Integer
    
        'Figure out which way to move
        Select Case nHeading

                Case E_Heading.NORTH
                        Y = -1
        
                Case E_Heading.EAST
                        X = 1
        
                Case E_Heading.SOUTH
                        Y = 1
        
                Case E_Heading.WEST
                        X = -1

        End Select
    
        'Fill temp pos
        tx = UserPos.X + X
        TY = UserPos.Y + Y
    
        'Check to see if its out of bounds
        If tx < MinXBorder Or tx > MaxXBorder Or TY < MinYBorder Or TY > MaxYBorder Then
                Exit Sub
        Else
                'Start moving... MainLoop does the rest
                AddtoUserPos.X = X
                UserPos.X = tx
                AddtoUserPos.Y = Y
                UserPos.Y = TY
                UserMoving = 1
        
                bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                   MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                   MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

        End If

End Sub

Private Function HayFogata(ByRef location As Position) As Boolean

        Dim j As Long

        Dim K As Long
    
        For j = UserPos.X - 8 To UserPos.X + 8
                For K = UserPos.Y - 6 To UserPos.Y + 6

                        If InMapBounds(j, K) Then
                                If MapData(j, K).ObjGrh.GrhIndex = GrhFogata Then
                                        location.X = j
                                        location.Y = K
                    
                                        HayFogata = True
                                        Exit Function

                                End If

                        End If

                Next K
        Next j

End Function

Function NextOpenChar() As Integer

        '*****************************************************************
        'Finds next open char slot in CharList
        '*****************************************************************
        Dim loopc As Long

        Dim Dale  As Boolean
    
        loopc = 1

        Do While charlist(loopc).Active And Dale
                loopc = loopc + 1
                Dale = (loopc <= UBound(charlist))
        Loop
    
        NextOpenChar = loopc

End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean

        On Error GoTo errorhandler

        Dim Grh         As Long

        Dim Frame       As Long

        Dim GrhCount    As Long

        Dim Handle      As Integer

        Dim FileVersion As Long
    
        'Open files
        Handle = FreeFile()
    
        Open DirInit & "Graficos.ind" For Binary Access Read As Handle
        Seek #1, 1
    
        'Get file version
        Get Handle, , FileVersion
    
        'Get number of grhs
        Get Handle, , GrhCount
    
        'Resize arrays
        ReDim GrhData(1 To GrhCount) As GrhData
    
        While Not EOF(Handle)

                Get Handle, , Grh
        
                If Grh <> 0 Then

                        With GrhData(Grh)
                                'Get number of frames
                                Get Handle, , .NumFrames

                                If .NumFrames <= 0 Then GoTo errorhandler
                
                                ReDim .Frames(1 To GrhData(Grh).NumFrames)
                
                                If .NumFrames > 1 Then

                                        'Read a animation GRH set
                                        For Frame = 1 To .NumFrames
                                                Get Handle, , .Frames(Frame)

                                                If .Frames(Frame) <= 0 Or .Frames(Frame) > GrhCount Then
                                                        GoTo errorhandler

                                                End If

                                        Next Frame
                    
                                        Get Handle, , .Speed
                    
                                        If .Speed <= 0 Then GoTo errorhandler
                    
                                        'Compute width and height
                                        .pixelHeight = GrhData(.Frames(1)).pixelHeight

                                        If .pixelHeight <= 0 Then GoTo errorhandler
                    
                                        .pixelWidth = GrhData(.Frames(1)).pixelWidth

                                        If .pixelWidth <= 0 Then GoTo errorhandler
                    
                                        .TileWidth = GrhData(.Frames(1)).TileWidth

                                        If .TileWidth <= 0 Then GoTo errorhandler
                    
                                        .TileHeight = GrhData(.Frames(1)).TileHeight

                                        If .TileHeight <= 0 Then GoTo errorhandler
                                Else
                                        'Read in normal GRH data
                                        Get Handle, , .FileNum

                                        If .FileNum <= 0 Then GoTo errorhandler
                    
                                        Get Handle, , GrhData(Grh).sX

                                        If .sX < 0 Then GoTo errorhandler
                    
                                        Get Handle, , .sY

                                        If .sY < 0 Then GoTo errorhandler
                    
                                        Get Handle, , .pixelWidth

                                        If .pixelWidth <= 0 Then GoTo errorhandler
                    
                                        Get Handle, , .pixelHeight

                                        If .pixelHeight <= 0 Then GoTo errorhandler
                    
                                        'Compute width and height
                                        .TileWidth = .pixelWidth / TilePixelHeight
                                        .TileHeight = .pixelHeight / TilePixelWidth
                    
                                        .Frames(1) = Grh

                                End If

                        End With

                End If

        Wend
    
        Close Handle
    
        LoadGrhData = True
        Exit Function

errorhandler:
        LoadGrhData = False

End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

        '*****************************************************************
        'Checks to see if a tile position is legal
        '*****************************************************************
        'Limites del mapa
        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                Exit Function

        End If
    
        'Tile Bloqueado?
        If MapData(X, Y).Blocked = 1 Then
                Exit Function

        End If
    
        '�Hay un personaje?
        If MapData(X, Y).CharIndex > 0 Then
                Exit Function

        End If
   
        If UserNavegando <> HayAgua(X, Y) Then
                Exit Function

        End If
    
        LegalPos = True

End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

        '*****************************************************************
        'Author: ZaMa
        'Last Modify Date: 01/08/2009
        'Checks to see if a tile position is legal, including if there is a casper in the tile
        '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
        '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
        '*****************************************************************
        Dim CharIndex As Integer
    
        'Limites del mapa
        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                Exit Function

        End If
    
        'Tile Bloqueado?
        If MapData(X, Y).Blocked = 1 Then
                Exit Function

        End If
    
        CharIndex = MapData(X, Y).CharIndex

        '�Hay un personaje?
        If CharIndex > 0 Then
    
                If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
                        Exit Function

                End If
        
                With charlist(CharIndex)

                        ' Si no es casper, no puede pasar
                        If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                                Exit Function
                        Else

                                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                                If HayAgua(UserPos.X, UserPos.Y) Then
                                        If Not HayAgua(X, Y) Then Exit Function
                                Else

                                        ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                                        If HayAgua(X, Y) Then Exit Function

                                End If
                
                                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                                        If charlist(UserCharIndex).invisible = True Then Exit Function

                                End If

                        End If

                End With

        End If
   
        If UserNavegando <> HayAgua(X, Y) Then
                Exit Function

        End If
    
        MoveToLegalPos = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean

        '*****************************************************************
        'Checks to see if a tile position is in the maps bounds
        '*****************************************************************
        If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
                Exit Function

        End If
    
        InMapBounds = True

End Function

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, _
                              ByVal X As Integer, _
                              ByVal Y As Integer, _
                              ByVal center As Byte, _
                              ByVal Animate As Byte)

        Dim CurrentGrhIndex As Integer

        Dim SourceRect      As RECT

        On Error GoTo error
        
        If Animate Then
                If Grh.Started = 1 Then
                        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)

                        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                                If Grh.Loops <> INFINITE_LOOPS Then
                                        If Grh.Loops > 0 Then
                                                Grh.Loops = Grh.Loops - 1
                                        Else
                                                Grh.Started = 0
                                                Exit Sub

                                        End If

                                End If

                        End If

                End If

        End If
    
        'Figure out what frame to draw (always 1 if not animated)
        CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
        With GrhData(CurrentGrhIndex)

                'Center Grh over X,Y pos
                If center Then
                        If .TileWidth <> 1 Then
                                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

                        End If
            
                        If .TileHeight <> 1 Then
                                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

                        End If

                End If
        
                SourceRect.Left = .sX
                SourceRect.Top = .sY
                SourceRect.Right = SourceRect.Left + .pixelWidth
                SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
                'Draw
                Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_WAIT)

        End With

        Exit Sub

error:

        If Err.number = 9 And Grh.FrameCounter < 1 Then
                Grh.FrameCounter = 1
                Resume
        Else
                MsgBox Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
                End

        End If

End Sub

Public Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, _
                                       ByVal X As Integer, _
                                       ByVal Y As Integer, _
                                       ByVal center As Byte)

        Dim SourceRect As RECT
    
        With GrhData(GrhIndex)

                'Center Grh over X,Y pos
                If center Then
                        If .TileWidth <> 1 Then
                                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

                        End If
            
                        If .TileHeight <> 1 Then
                                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

                        End If

                End If
        
                SourceRect.Left = .sX
                SourceRect.Top = .sY
                SourceRect.Right = SourceRect.Left + .pixelWidth
                SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
                'Draw
                Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)

        End With

End Sub

Sub DDrawTransGrhtoSurfaceAlpha(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT

    If Animate Then

        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1

                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If

    End If

    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
        If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    With SourceRect
        .Left = GrhData(CurrentGrhIndex).sX
        .Top = GrhData(CurrentGrhIndex).sY
        .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
    End With

    Dim Src    As DirectDrawSurface7
    Dim rDest  As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo   As Long

    Set Src = SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum)

    Src.GetSurfaceDesc ddsdSrc
    BackBufferSurface.GetSurfaceDesc ddsdDest

    With rDest
        .Left = X
        .Top = Y
        .Right = X + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = Y + GrhData(CurrentGrhIndex).pixelHeight

        If .Right > ddsdDest.lWidth Then
            .Right = ddsdDest.lWidth
        End If
        If .Bottom > ddsdDest.lHeight Then
            .Bottom = ddsdDest.lHeight
        End If
    End With

    ' 0 -> 16 bits 555
    ' 1 -> 16 bits 565
    ' 2 -> 16 bits raro (Sin implementar)
    ' 3 -> 24 bits
    ' 4 -> 32 bits

    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0& And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0& Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0& And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0& Then
        Modo = 1
        'TODO : Revisar las m�scaras de 24!! Quiz�s mirando el campo lRGBBitCount para diferenciar 24 de 32...
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0& And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0& Then
        Modo = 3
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &HFF00& And ddsdSrc.ddpfPixelFormat.lGBitMask = &HFF00& Then
        Modo = 4
    Else
        'Modo = 2 '16 bits raro ?
        Call BackBufferSurface.BltFast(X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        Exit Sub
    End If

    Dim SrcLock As Boolean
    Dim DstLock As Boolean

    SrcLock = False
    DstLock = False

    On Local Error GoTo HayErrorAlpha

    Call Src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    SrcLock = True
    Call BackBufferSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
    DstLock = True

    Call BackBufferSurface.GetLockedArray(dArray())
    Call Src.GetLockedArray(sArray())

    Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)

    BackBufferSurface.Unlock rDest
    DstLock = False
    Src.Unlock SourceRect
    SrcLock = False
    Exit Sub

HayErrorAlpha:
    If SrcLock Then Src.Unlock SourceRect
    If DstLock Then BackBufferSurface.Unlock rDest
End Sub
Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, _
                           ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal center As Byte, _
                           ByVal Animate As Byte, _
                           Optional ByVal killAtEnd As Byte = 1)

        '*****************************************************************
        'Draws a GRH transparently to a X and Y position
        '*****************************************************************
        Dim CurrentGrhIndex As Integer

        Dim SourceRect      As RECT
    
        On Error GoTo error
    
        If Animate Then
                If Grh.Started = 1 Then
                        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
                        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                                If Grh.Loops <> INFINITE_LOOPS Then
                                        If Grh.Loops > 0 Then
                                                Grh.Loops = Grh.Loops - 1
                                        Else
                                                Grh.Started = 0

                                                If killAtEnd Then Exit Sub

                                        End If

                                End If

                        End If

                End If

        End If
    
        'Figure out what frame to draw (always 1 if not animated)
        CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
        With GrhData(CurrentGrhIndex)

                'Center Grh over X,Y pos
                If center Then
                        If .TileWidth <> 1 Then
                                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

                        End If
            
                        If .TileHeight <> 1 Then
                                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

                        End If

                End If
                
                SourceRect.Left = .sX
                SourceRect.Top = .sY
                SourceRect.Right = SourceRect.Left + .pixelWidth
                SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
                If X < BackBufferRect.Left Then
                        SourceRect.Left = SourceRect.Left - X
                        X = 0

                End If
        
                If Y < BackBufferRect.Top Then
                        SourceRect.Top = SourceRect.Top - Y
                        Y = 0

                End If
        
                'Draw
                Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)

        End With

        Exit Sub

error:

        If Err.number = 9 And Grh.FrameCounter < 1 Then
                Grh.FrameCounter = 1
                Resume
        Else
                MsgBox "Ocurri� un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripci�n del error: " & _
                   vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
                End

        End If

End Sub

Sub DrawGrhtoHdc(ByVal hdc As Long, _
                 ByVal GrhIndex As Integer, _
                 ByRef SourceRect As RECT, _
                 ByRef destRect As RECT)
        '*****************************************************************
        'Draws a Grh's portion to the given area of any Device Context
        '*****************************************************************
        Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hdc, SourceRect, destRect)

End Sub

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, _
                                   ByVal dstX As Long, _
                                   ByVal dstY As Long, _
                                   ByVal GrhIndex As Integer, _
                                   ByRef SourceRect As RECT, _
                                   ByVal TransparentColor As Long)

        '**************************************************************
        'Author: Torres Patricio (Pato)
        'Last Modify Date: 12/22/2009
        'This method is SLOW... Don't use in a loop if you care about
        'speed!
        '*************************************************************
        Dim color   As Long

        Dim X       As Long

        Dim Y       As Long

        Dim srchdc  As Long

        Dim Surface As DirectDrawSurface7
    
        Set Surface = SurfaceDB.Surface(GrhData(GrhIndex).FileNum)
    
        srchdc = Surface.GetDC
    
        For X = SourceRect.Left To SourceRect.Right - 1
                For Y = SourceRect.Top To SourceRect.Bottom - 1
                        color = GetPixel(srchdc, X, Y)
            
                        If color <> TransparentColor Then
                                Call SetPixel(dsthdc, dstX + (X - SourceRect.Left), dstY + (Y - SourceRect.Top), color)

                        End If

                Next Y
        Next X
    
        Call Surface.ReleaseDC(srchdc)

End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, _
                              ByRef Picture As StdPicture, _
                              ByVal x1 As Single, _
                              ByVal Y1 As Single, _
                              Optional Width1, _
                              Optional Height1, _
                              Optional x2, _
                              Optional Y2, _
                              Optional Width2, _
                              Optional Height2)
        '**************************************************************
        'Author: Torres Patricio (Pato)
        'Last Modify Date: 12/28/2009
        'Draw Picture in the PictureBox
        '*************************************************************

        Call PictureBox.PaintPicture(Picture, x1, Y1, Width1, Height1, x2, Y2, Width2, Height2)

End Sub

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)

'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y      As Long     'Keeps track of where on map we are

    Dim X      As Long     'Keeps track of where on map we are

    Dim screenminY As Integer  'Start Y pos on current screen

    Dim screenmaxY As Integer  'End Y pos on current screen

    Dim screenminX As Integer  'Start X pos on current screen

    Dim screenmaxX As Integer  'End X pos on current screen

    Dim minY   As Integer  'Start Y pos on current map

    Dim maxY   As Integer  'End Y pos on current map

    Dim minX   As Integer  'Start X pos on current map

    Dim maxX   As Integer  'End X pos on current map

    Dim ScreenX As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset As Integer

    Dim minYOffset As Integer

    Dim PixelOffsetXTemp As Integer    'For centering grhs

    Dim PixelOffsetYTemp As Integer    'For centering grhs

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth

    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize

    'Make sure mins and maxs are allways in map bounds
    If minY < YMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If

    If maxY > YMaxMapSize Then maxY = YMaxMapSize

    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If

    If maxX > XMaxMapSize Then maxX = XMaxMapSize

    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If

    If screenmaxY < YMaxMapSize Then
        screenmaxY = screenmaxY + 1
    ElseIf screenmaxY > YMaxMapSize Then
        screenmaxY = YMaxMapSize

    End If

    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If

    If screenmaxX < XMaxMapSize Then
        screenmaxX = screenmaxX + 1
    ElseIf screenmaxX > XMaxMapSize Then
        screenmaxX = XMaxMapSize

    End If

    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX

            'Layer 1 **********************************
            Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), _
                                   (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                                   (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                                   0, 1)
            '******************************************

            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y

    'Draw floor layer 2
    ScreenY = minYOffset

    For Y = minY To maxY
        ScreenX = minXOffset

        For X = minX To maxX

            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), _
                                            (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                                            (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                                            1, 1)

            End If

            '******************************************

            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    'Draw Transparent Layers
    ScreenY = minYOffset

    For Y = minY To maxY
        ScreenX = minXOffset

        For X = minX To maxX
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY

            With MapData(X, Y)

                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.ObjGrh, _
                                                PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)

                End If

                '***********************************************

                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)

                End If

                '*************************************************

                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                       PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If

                '************************************************


'If .Damage.Activated Then
'               Mod_rDamage.Draw X, Y, PixelOffsetXTemp + 5, PixelOffsetYTemp - 5
'            End If
0                If .Damage.Activated Then
                   Mod_rDamage.Draw X, Y, PixelOffsetXTemp + 20, PixelOffsetYTemp - 30

                End If

            End With

            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    If Not bTecho Then
        'Draw blocked tiles and grid
        ScreenY = minYOffset

        For Y = minY To maxY
            ScreenX = minXOffset

            For X = minX To maxX

                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    'Draw
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                                                (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                                                (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                                                1, 1)

                End If

                '**********************************

                ScreenX = ScreenX + 1
            Next X

            ScreenY = ScreenY + 1
        Next Y

    End If

    'TODO : Check this!!
    If UserMap = 127 Or UserMap = 128 Then

        If bRain Then

4994:            'Figure out what frame to draw
            If llTick < DirectX.TickCount - 50 Then
                iFrameIndex = iFrameIndex + 1
                If iFrameIndex > 7 Then iFrameIndex = 0
                llTick = DirectX.TickCount

            End If

            For Y = 0 To 4
                For X = 0 To 4
                    Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(X), SurfaceDB.Surface(15168), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
                Next X
            Next Y

        End If

    End If
    
    If SegundosInvisible > 0 Then
        Mod_TileEngine.RenderText 620, 260, "Segundos de invi restantes: " & SegundosInvisible, RGB(255, 255, 50), frmMain.font
    End If
    
    
    If AyudaClanActive Then
        Dim Cantt As Byte
        For Cantt = 1 To AyudaClanMax
            If AyudaClan(Cantt).Active Then
                Mod_TileEngine.RenderText 700, 550 + (Cantt * 10), AyudaClan(Cantt).Linea, RGB(255, 255, 50), frmMSG.font
            End If
        Next Cantt
    End If
    
    If DialogosClanes.QuedanDialogos = False Then
        If ZonaSeguraM = False Then
            Mod_TileEngine.RenderText 260, 260, "Zona segura", RGB(55, 255, 50), frmMain.font
        Else
            Mod_TileEngine.RenderText 260, 260, "Zona insegura", RGB(255, 50, 50), frmMain.font
        End If
    End If
    
End Sub

Public Function RenderSounds()

        '**************************************************************
        'Author: Juan Mart�n Sotuyo Dodero
        'Last Modify Date: 3/30/2008
        'Actualiza todos los sonidos del mapa.
        '**************************************************************
        If bLluvia(UserMap) = 1 Then
        
                If bRain Then
                If UserMap <> 1 Then Exit Function
                        If bTecho Then
                                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                                        If RainBufferIndex Then _
                                           Call Audio.StopWave(RainBufferIndex)
                                        RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                                        frmMain.IsPlaying = PlayLoop.plLluviain

                                End If

                        Else

                                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                                        If RainBufferIndex Then _
                                           Call Audio.StopWave(RainBufferIndex)
                                        RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                                        frmMain.IsPlaying = PlayLoop.plLluviaout

                                End If

                        End If

                End If

        End If
    
        DoFogataFx

End Function

Function HayUserAbajo(ByVal X As Integer, _
                      ByVal Y As Integer, _
                      ByVal GrhIndex As Integer) As Boolean

        If GrhIndex > 0 Then
                HayUserAbajo = _
                   charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                   And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                   And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                   And charlist(UserCharIndex).Pos.Y <= Y

        End If

End Function

Sub LoadGraphics()
        '**************************************************************
        'Author: Juan Mart�n Sotuyo Dodero - complete rewrite
        'Last Modify Date: 11/03/2006
        'Initializes the SurfaceDB and sets up the rain rects
        '**************************************************************
        'New surface manager :D
        Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
    
        'Set up te rain rects
        RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
        RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
        RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
        RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
        RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
        RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
        RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
        RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256

End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, _
                               ByVal setMainViewTop As Integer, _
                               ByVal setMainViewLeft As Integer, _
                               ByVal setTilePixelHeight As Integer, _
                               ByVal setTilePixelWidth As Integer, _
                               ByVal setWindowTileHeight As Integer, _
                               ByVal setWindowTileWidth As Integer, _
                               ByVal setTileBufferSize As Integer, _
                               ByVal pixelsToScrollPerFrameX As Integer, _
                               pixelsToScrollPerFrameY As Integer, _
                               ByVal engineSpeed As Single) As Boolean

        '***************************************************
        'Author: Aaron Perkins
        'Last Modification: 08/14/07
        'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Creates all DX objects and configures the engine to start running.
        '***************************************************
        Dim SurfaceDesc As DDSURFACEDESC2

        Dim ddck        As DDCOLORKEY
        
        'Fill startup variables
        MainViewTop = setMainViewTop
        MainViewLeft = setMainViewLeft
        TilePixelWidth = setTilePixelWidth
        TilePixelHeight = setTilePixelHeight
        WindowTileHeight = setWindowTileHeight
        WindowTileWidth = setWindowTileWidth
        TileBufferSize = setTileBufferSize
    
        HalfWindowTileHeight = setWindowTileHeight \ 2
        HalfWindowTileWidth = setWindowTileWidth \ 2
    
        'Compute offset in pixels when rendering tile buffer.
        'We diminish by one to get the top-left corner of the tile for rendering.
        TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
        TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
        engineBaseSpeed = engineSpeed
    
        'Set FPS value to 60 for startup
        FPS = 60
        FramesPerSecCounter = 60
    
        MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
        MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
        MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
        MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
        MainViewWidth = TilePixelWidth * WindowTileWidth
        MainViewHeight = TilePixelHeight * WindowTileHeight
    
        'Resize mapdata array
        ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
        'Set intial user position
        UserPos.X = MinXBorder
        UserPos.Y = MinYBorder
    
        'Set scroll pixels per frame
        ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
        ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
        'Set the view rect
        With MainViewRect
                .Left = MainViewLeft
                .Top = MainViewTop
                .Right = .Left + MainViewWidth
                .Bottom = .Top + MainViewHeight

        End With
    
        'Set the dest rect
        With MainDestRect
                .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
                .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
                .Right = .Left + MainViewWidth
                .Bottom = .Top + MainViewHeight

        End With
    
        On Error Resume Next

        Set DirectX = New DirectX7
    
        If Err Then
                MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
                Exit Function

        End If
    
        '****** INIT DirectDraw ******
        ' Create the root DirectDraw object
        Set DirectDraw = DirectX.DirectDrawCreate(vbNullString)
    
        If Err Then
                MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
                Exit Function

        End If
    
        On Error GoTo 0

        Call DirectDraw.SetCooperativeLevel(setDisplayFormhWnd, DDSCL_NORMAL)
    
        'Primary Surface
        ' Fill the surface description structure
        With SurfaceDesc
                .lFlags = DDSD_CAPS
                .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE

        End With

        ' Create the surface
        Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)
    
        'Create Primary Clipper
        Set PrimaryClipper = DirectDraw.CreateClipper(0)
        Call PrimaryClipper.SetHWnd(frmMain.hwnd)
        Call PrimarySurface.SetClipper(PrimaryClipper)
    
        With BackBufferRect
                .Left = 0
                .Top = 0
                .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
                .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)

        End With
    
        With SurfaceDesc
                .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH

                If ClientSetup.bUseVideo Then
                        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
                Else
                        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY

                End If

                .lHeight = BackBufferRect.Bottom
                .lWidth = BackBufferRect.Right

        End With
    
        ' Create surface
        Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)
    
        'Set color key
        ddck.low = 0
        ddck.high = 0
        Call BackBufferSurface.SetColorKey(DDCKEY_SRCBLT, ddck)
    
        'Set font transparency
        Call BackBufferSurface.SetFontTransparency(D_TRUE)
    
        Call LoadGrhData
        Call CargarCuerpos
        Call CargarCabezas
        Call CargarCascos
        Call CargarFxs
    
        LTLluvia(0) = 224
        LTLluvia(1) = 352
        LTLluvia(2) = 480
        LTLluvia(3) = 608
        LTLluvia(4) = 736
    
        Call LoadGraphics
    
        InitTileEngine = True

End Function

Public Sub DeinitTileEngine()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modification: 08/14/07
        'Destroys all DX objects
        '***************************************************
        On Error Resume Next

        Set PrimarySurface = Nothing
        Set PrimaryClipper = Nothing
        Set BackBufferSurface = Nothing
    
        Set DirectDraw = Nothing
    
        Set DirectX = Nothing

End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, _
   ByVal DisplayFormLeft As Integer)

        '***************************************************
        'Author: Arron Perkins
        'Last Modification: 08/14/07
        'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Updates the game's model and renders everything.
        '***************************************************
        Static OffsetCounterX As Single

        Static OffsetCounterY As Single
    
        '****** Set main view rectangle ******
        MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
        MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
        MainViewRect.Right = MainViewRect.Left + MainViewWidth
        MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
        If EngineRun Then
                If UserMoving Then

                        '****** Move screen Left and Right if needed ******
                        If AddtoUserPos.X <> 0 Then
                                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame

                                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                                        OffsetCounterX = 0
                                        AddtoUserPos.X = 0
                                        UserMoving = False

                                End If

                        End If
            
                        '****** Move screen Up and Down if needed ******
                        If AddtoUserPos.Y <> 0 Then
                                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame

                                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                                        OffsetCounterY = 0
                                        AddtoUserPos.Y = 0
                                        UserMoving = False

                                End If

                        End If

                End If
                
                '****** Update screen ******
                If UserCiego Then
                        Call CleanViewPort
                Else
                        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

                End If
        
                Call Dialogos.Render
                Call DibujarCartel
        
                Call DialogosClanes.Draw
        
                'Display front-buffer!
                Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT)
        
                'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
                While (GetTickCount - fpsLastCheck) \ 10 < FramesPerSecCounter

                        Sleep 5
                Wend
        
                'Si est� activado el FragShooter y est� esperando para sacar una foto, lo hacemos:
                If ClientSetup.bActive Then
                        If FragShooterCapturePending Then
                                DoEvents
                                Call ScreenCapture(True)
                                FragShooterCapturePending = False

                        End If

                End If
        
                'FPS update
                If fpsLastCheck + 1000 < GetTickCount Then
                        FPS = FramesPerSecCounter
                        FramesPerSecCounter = 1
                        fpsLastCheck = GetTickCount
                Else
                        FramesPerSecCounter = FramesPerSecCounter + 1

                End If
    
                'Get timing info
                timerElapsedTime = GetElapsedTime()
                timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

        End If

End Sub

Public Sub RenderText(ByVal lngXPos As Integer, _
                      ByVal lngYPos As Integer, _
                      ByRef strText As String, _
                      ByVal lngColor As Long, _
                      ByRef font As StdFont)

        If Len(strText) <> 0 Then
                Call BackBufferSurface.SetForeColor(vbBlack)
                Call BackBufferSurface.SetFont(font)
                Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, False)
        
                Call BackBufferSurface.SetForeColor(lngColor)
                Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)

        End If

End Sub

Public Sub RenderTextCentered(ByVal lngXPos As Integer, _
                              ByVal lngYPos As Integer, _
                              ByRef strText As String, _
                              ByVal lngColor As Long, _
                              ByRef font As StdFont)

        Dim hdc As Long

        Dim Ret As size
    
        If Len(strText) <> 0 Then
                Call BackBufferSurface.SetFont(font)
        
                'Get width of text once rendered
                hdc = BackBufferSurface.GetDC()
                Call GetTextExtentPoint32(hdc, strText, Len(strText), Ret)
                Call BackBufferSurface.ReleaseDC(hdc)
        
                lngXPos = lngXPos - Ret.cx \ 2
        
                Call BackBufferSurface.SetForeColor(vbBlack)
                Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, False)
        
                Call BackBufferSurface.SetForeColor(lngColor)
                Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)

        End If

End Sub

Private Function GetElapsedTime() As Single

        '**************************************************************
        'Author: Aaron Perkins
        'Last Modify Date: 10/07/2002
        'Gets the time that past since the last call
        '**************************************************************
        Dim start_time    As Currency

        Static end_time   As Currency

        Static timer_freq As Currency

        'Get the timer frequency
        If timer_freq = 0 Then
                QueryPerformanceFrequency timer_freq

        End If
    
        'Get current time
        Call QueryPerformanceCounter(start_time)
    
        'Calculate elapsed time
        GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
        'Get next end time
        Call QueryPerformanceCounter(end_time)

End Function

Private Sub CharRender(ByVal CharIndex As Long, _
                       ByVal PixelOffsetX As Integer, _
                       ByVal PixelOffsetY As Integer)

'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 25/05/2011 (Amraphen)
'Draw char's to screen without offcentering them
'16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
'25/05/2011: Amraphen - Agregado movimiento de armas al golpear.
'***************************************************
    Dim moved  As Boolean

    Dim attacked As Boolean

    Dim Pos    As Integer

    Dim line   As String

    Dim color  As Long

    With charlist(CharIndex)

        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame

                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                   .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1

                'Char moved
                moved = True

                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                   (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If

            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame

                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                   .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1

                'Char moved
                moved = True

                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                   (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If

        If .Heading = 0 Then
            .Heading = E_Heading.SOUTH

        End If

        If .UsandoArma And .Arma.WeaponWalk(.Heading).Started Then _
           attacked = True

        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1

            If Not attacked Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
                .UsandoArma = False

            End If

            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1

            .Moving = False

        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY


        If .Min_HP > 0 Then
            'If (Val(UserPos.X + frmMain.MouseX \ 32 - PixelOffsetX \ 64)) = .Pos.X And _
             Val((UserPos.Y + frmMain.MouseY / 32 - PixelOffsetY \ 64)) = .Pos.Y Then
            'If Not .Body.HeadOffset Then

            Dim TY As Byte, tx As Byte
            Call ConvertCPtoTP(frmMain.MouseX, frmMain.MouseY, tx, TY)



            If (Abs(tx - charlist(CharIndex).Pos.X) <= 1) And (Abs(TY - charlist(CharIndex).Pos.Y) <= 1) Then
                Call RenderBoxNPC(CharIndex, PixelOffsetX, PixelOffsetY)
            End If
            'End If
            'End If
        End If
        ' If .Min_HP > 0 Then Call RenderBoxNPC(CharIndex, PixelOffsetX, PixelOffsetY)

        If Not .invisible Then

            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
               Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, 0)

            'Draw Head
            If .Head.Head(.Heading).GrhIndex Then
                Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)

                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                   Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 0)

                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                   Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, 0)

                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                   Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, 0)

            End If

            'Draw name over head
            If LenB(.Nombre) > 0 Then
                If Nombres Then
                    Pos = getTagPosition(.Nombre)

                    If .priv = 0 Then
                        If .Atacable Then
                            color = ColoresPJ(48)
                        Else
                            If .Criminal Then
                                color = ColoresPJ(50)
                            Else
                                color = ColoresPJ(49)
                            End If
                        End If
                    Else
                        color = ColoresPJ(.priv)
                    End If

                    If .EsGranPoder Then
                        color = RGB(192, 192, 192)
                    End If
                    
                    Dim tmpB As Boolean
                    If .Nombre = "SHOP" Then
                        color = RGB(RandomNumber(100, 255), 255, 100): tmpB = True
                    ElseIf InStr(1, UCase$(.Nombre), "SUGAR") Then
                        color = RGB(RandomNumber(100, 255), 255, 100): tmpB = True
                    End If
                    
                    ' @@ DURO
                    If .Nombre = "Participante" Then
                        color = RGB(192, 192, 192)
                    End If

                    'Nick
                    line = Left$(.Nombre, Pos - 2)

                    If .priv = 0 Then
                        If .Heading = E_Heading.SOUTH And .Max_Hp <= 0 Then
                            If .esNPC > 0 Then

                                color = RGB(1, 210, 1)

                            End If
                        End If
                    End If

                    If .priv > 0 And .priv < 4 Then
                        Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, color, frmMain.font)

                    ElseIf tmpB Then
                        Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, color, frmList.font)
                    Else
                        Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, color, frmMain.font)
                    End If

                    ' If .EsDonador <> 0 Then

                    '     Dim Estrella As Grh
                    '     Estrella.GrhIndex = 24935
                    '
                    '     If Len(line) > 15 Then
                    '        Call DDrawTransGrhtoSurface(Estrella, (PixelOffsetX - Round(TilePixelWidth \ 2) - Len(line) * 2) - Len(line) - 20, PixelOffsetY + 10, 1, 1)
                    '     Else
                    '         Call DDrawTransGrhtoSurface(Estrella, (PixelOffsetX - Round(TilePixelWidth \ 2) - Len(line)) - Len(line), PixelOffsetY + 10, 1, 1)
                    '     End If
                    ' End If

                    'Clan
44                  line = mid$(.Nombre, Pos)

                    Dim TAG_USER_INVISIBLE As String

                    TAG_USER_INVISIBLE = "[INVISIBLE]"

                    If line = TAG_USER_INVISIBLE Then
                        color = RGB(128, 128, 128)

                    End If

                    If .Conquisto Then
                        Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 8, PixelOffsetY + 45, line, RGB(255, 255, 0), frmScreenshots.font)
                    Else
                        Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 8, PixelOffsetY + 45, line, RGB(131, 131, 131), frmScreenshots.font)
                    End If
                End If

            End If

        Else

            Call modConsole.DrawInvisibleChar(CharIndex, PixelOffsetX, PixelOffsetY)

        End If

        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex)    '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo

        'Draw FX
        If .FxIndex <> 0 Then
            Call DDrawTransGrhtoSurface(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1)

            'Check if animation is over
            If .fX.Started = 0 Then _
               .FxIndex = 0

        End If

    End With

End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, _
                          ByVal fX As Integer, _
                          ByVal Loops As Integer)

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modify Date: 12/03/04
        'Sets an FX to the character.
        '***************************************************
        With charlist(CharIndex)
                .FxIndex = fX
        
                If .FxIndex > 0 Then
                        Call InitGrh(.fX, FxData(fX).Animacion)
        
                        .fX.Loops = Loops

                End If

        End With

End Sub

Private Sub CleanViewPort()

        '***************************************************
        'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
        'Last Modify Date: 12/03/04
        'Fills the viewport with black.
        '***************************************************
        Dim r As RECT

        Call BackBufferSurface.BltColorFill(r, vbBlack)

End Sub

Private Sub RenderBoxNPC(ByVal CharIndex As Long, _
                         ByVal PixelOffsetX As Integer, _
                         ByVal PixelOffsetY As Integer)

        ' / kevin LOL
    
        With charlist(CharIndex)

                If .Min_HP > 0 Then
                
                        BackBufferSurface.SetForeColor vbBlack
                        BackBufferSurface.SetFillColor vbRed
                    
                        BackBufferSurface.DrawBox PixelOffsetX - 20, PixelOffsetY + 52, (((.Min_HP / 100) / (.Max_Hp / 100)) * 55) + PixelOffsetX, PixelOffsetY + 44
            
                End If

        End With

End Sub

Public Function ClanTag(ByVal CharIndex As Long) As Boolean
        '---------------------------------------------------------------------------------------
        ' Procedimiento : ClanTag
        ' Autor         : Lagalot
        ' Fecha         : 20/06/12
        ' Proposito     : Revisa si determinado charindex tiene el mismo clan que el usuario actual.
        '---------------------------------------------------------------------------------------
        '

        Dim sclan    As String

        Dim UserClan As String

        Dim Pos      As Integer
    
        Pos = InStr(charlist(CharIndex).Nombre, "<")

        If Pos > 0 Then
                UserClan = mid$(charlist(CharIndex).Nombre, InStr(charlist(CharIndex).Nombre, "<"))

                If InStr(charlist(UserCharIndex).Nombre, "<") > 0 Then
                        sclan = mid$(charlist(UserCharIndex).Nombre, InStr(charlist(UserCharIndex).Nombre, "<"))
                Else
                        sclan = "0"

                End If

        Else
                sclan = "0"
                UserClan = "1"

        End If
        
        If UserClan = sclan Or UCase$(charlist(CharIndex).Nombre) = UCase$(frmMain.lblName.Caption) Then
                ClanTag = True

                Exit Function

        End If

        ClanTag = False

End Function
