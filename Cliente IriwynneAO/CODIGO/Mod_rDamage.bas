Attribute VB_Name = "Mod_rDamage"
Option Explicit
 
Const DAMAGE_TIME   As Integer = 50

Const DAMAGE_FONT_S As Byte = 20
 
Enum EDType

        edPuñal = 1                'Apuñalo.
        edNormal = 2               'Hechizo o golpe común.

End Enum
 
Private DNormalFont As New StdFont
 
Type DList

        DamageVal      As Integer     'Cantidad de daño.
        ColorRGB       As Long       'Color.
        DamageType     As EDType     'Tipo, se usa para saber si es apu o no.
        DamageFont     As New StdFont      'Efecto del apu.
        TimeRendered   As Integer    'Tiempo transcurrido.
        Downloading    As Byte       'Contador para la posicion Y.
        Activated      As Boolean    'Si está activado..
        lastUpdated    As Long

End Type
 
Sub Initialize()
 
        ' @ Inicializa la font.
 
        With DNormalFont
     
                .size = 15
                .Italic = False
                .Bold = True
                .Name = "Tahoma"
     
        End With
 
End Sub
 
Sub Create(ByVal X As Byte, _
           ByVal Y As Byte, _
           ByVal ColorRGB As Long, _
           ByVal DamageValue As Integer, _
           ByVal edMode As Byte)
 
        ' @ Agrega un nuevo daño.
        If InMapBounds(X, Y) Then ' @@ Miqueas : Por si las dudas, nunca sabemos cuando va a explotar todo

                With MapData(X, Y).Damage
     
                        .Activated = True
                        .ColorRGB = ColorRGB
                        .DamageType = edMode
                        .DamageVal = DamageValue
                        .TimeRendered = 0
                        .Downloading = 0
     
                        If .DamageType = EDType.edPuñal Then

                                With .DamageFont
                                        .size = DAMAGE_FONT_S
                                        .Name = "Tahoma"
                                        .Bold = True
                                        Exit Sub

                                End With

                        End If
     
                        .DamageFont = DNormalFont
                        .DamageFont.size = 15
                        .DamageFont.Bold = True
                
                        .lastUpdated = GetTickCount()

                        If edMode = 10 Then .ColorRGB = RGB(255, 255, 0)
     
                End With

        End If

End Sub
 
Sub Draw(ByVal X As Byte, _
         ByVal Y As Byte, _
         ByVal PixelX As Integer, _
         ByVal PixelY As Integer)
 
        ' @ Dibuja un daño
 
        With MapData(X, Y).Damage
     
                If (Not .Activated) Or (Not .DamageVal <> 0) Then
                        Exit Sub

                End If

                If .TimeRendered < DAMAGE_TIME Then
           
                        'Sumo el contador del tiempo.
                        If GetTickCount() - .lastUpdated > 15 Then
                                .TimeRendered = .TimeRendered + 1
                                .lastUpdated = GetTickCount()

                        End If
           
                        If .DamageType = 10 Then
                        'If (.TimeRendered) > 0 Then
                                    .Downloading = (.TimeRendered * 2.5)
    
                           ' End If
                        Else
                            If (.TimeRendered * 0.5) > 0 Then
                                    .Downloading = (.TimeRendered * 0.5)
    
                            End If
                        End If
           
                        .ColorRGB = ModifyColour(.TimeRendered, .DamageType)
           
                        'Efectito para el apu :P
                        'If .DamageType = EDType.edPuñal Then
                        '        .DamageFont.size = NewSize(.TimeRendered)
                        'End If
                        
                        'Dibujo ; D
                        Dim signe As String
                  
                        If .DamageType = 10 Then
                                .ColorRGB = vbYellow
                                
                                .DamageFont.size = 12
                                .DamageFont.Bold = True
                                .DamageFont.Italic = False
                                signe = "+"
                        Else
                                signe = "-"

                        End If
               
                        'Dibujo ; D
                        RenderTextCentered PixelX - 5, PixelY + 8 - .Downloading, signe & .DamageVal, .ColorRGB, .DamageFont
              
                        'Si llego al tiempo lo limpio
                        If .DamageType = 10 Then
                        
                        If .TimeRendered >= 20 Then
                                Clear X, Y

                        End If
                        Else
                        If .TimeRendered >= DAMAGE_TIME Then
                                Clear X, Y

                        End If
                        End If
           
                End If
       
        End With
 
End Sub
 
Sub Clear(ByVal X As Byte, ByVal Y As Byte)
 
        ' @ Limpia todo.
 
        With MapData(X, Y).Damage
        
                .Activated = False
                .ColorRGB = 0
                .DamageVal = 0
                .TimeRendered = 0
                .lastUpdated = 0
                .Downloading = 0

        End With
 
End Sub
 
Function ModifyColour(ByVal TimeNowRendered As Byte, ByVal DamageType As Byte) As Long
 
        ' @ Se usa para el "efecto" de desvanecimiento.
 
        Select Case DamageType
                   
                Case EDType.edPuñal
                        ModifyColour = RGB(200, 200, 11)
                   
                Case EDType.edNormal
                        ModifyColour = RGB(255 - (TimeNowRendered * 3), 0, 0)
       
        End Select
 
End Function
 
