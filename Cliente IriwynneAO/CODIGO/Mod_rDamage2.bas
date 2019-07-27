Attribute VB_Name = "Mod_rDamage"
Option Explicit
 
Const DAMAGE_TIME   As Integer = 50
Const DAMAGE_FONT_S As Byte = 14
 
Enum EDType
     edPuñal = 1                'Apuñalo.
     edNormal = 2               'Hechizo o golpe común.
End Enum
 
Private DNormalFont    As New StdFont
 
Type DList
     DamageVal      As Integer  'Cantidad de daño.
     ColorRGB       As Long     'Color.
     DamageType     As EDType   'Tipo, se usa para saber si es apu o no.
     DamageFont     As New StdFont  'Efecto del apu.
     TimeRendered   As Integer  'Tiempo transcurrido.
     Downloading    As Byte     'Contador para la posicion Y.
     Activated      As Boolean  'Si está activado..
End Type
 
Sub Initialize()
 
' @ Inicializa la font.
 
With DNormalFont
     
     .size = 12
     .Italic = False
     .Bold = True
     .Name = "Tahoma"
     
End With
 
 
End Sub
 
Sub Create(ByVal X As Byte, ByVal Y As Byte, ByVal ColorRGB As Long, ByVal DamageValue As Integer, ByVal edMode As Byte)
 
' @ Agrega un nuevo daño.
 
With MapData(X, Y).Damage
     
     .Activated = True
     .ColorRGB = ColorRGB
     .DamageType = edMode
     .DamageVal = DamageValue
     .TimeRendered = 0
     .Downloading = 0
     
     If .DamageType = EDType.edPuñal Then
        With .DamageFont
             .size = Val(DAMAGE_FONT_S)
             .Name = "Tahoma"
             .Bold = True
             Exit Sub
        End With
     End If
     
     .DamageFont = DNormalFont
     .DamageFont.size = 14
     
End With
 
End Sub
 
Sub Draw(ByVal X As Byte, ByVal Y As Byte, ByVal PixelX As Integer, ByVal PixelY As Integer)
 
' @ Dibuja un daño
 
With MapData(X, Y).Damage
     
     If (Not .Activated) Or (Not .DamageVal <> 0) Then Exit Sub
        If .TimeRendered < DAMAGE_TIME Then
           
           'Sumo el contador del tiempo.
           .TimeRendered = .TimeRendered + 1
           
           If (.TimeRendered / 2) > 0 Then
               .Downloading = (.TimeRendered / 2)
           End If
           
           .ColorRGB = ModifyColour(.TimeRendered, .DamageType)
           
           'Efectito para el apu :P
           If .DamageType = EDType.edPuñal Then
              .DamageFont.size = NewSize(.TimeRendered)
           End If
               
           'Dibujo ; D
            If .DamageType <> EDType.edNormal And .DamageType <> EDType.edPuñal Then
           RenderTextCentered PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB, .DamageFont
            
            Else
           RenderTextCentered PixelX, PixelY - .Downloading, "" & .DamageVal, .ColorRGB, .DamageFont
           End If
           'Si llego al tiempo lo limpio
           If .TimeRendered >= DAMAGE_TIME Then
              Clear X, Y
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
End With
 
End Sub
 
Function ModifyColour(ByVal TimeNowRendered As Byte, ByVal DamageType As Byte) As Long
 
' @ Se usa para el "efecto" de desvanecimiento.
 
Select Case DamageType
                   
       Case EDType.edPuñal
            ModifyColour = RGB(200, 200, 100)
           'ModifyColour = GetPuñalNewColour()
                   
       Case EDType.edNormal
            ModifyColour = RGB(255 - (TimeNowRendered * 3), 0, 0)
            
        Case Else
        ModifyColour = RGB(255, 255, 10)
        
       
End Select
 
End Function
 
Function NewSize(ByVal timeRenderedNow As Byte) As Byte
 
' @ Se usa para el "efecto" del apu.
 
Dim tResult As Byte
 
Select Case timeRenderedNow
 
       Case 1 To 10
            NewSize = 14
       
       Case 11 To 20
            NewSize = 13
           
       Case 21 To 30
            NewSize = 12
           
       Case 31 To 40
            NewSize = 11
       
       Case 41 To 50
            NewSize = 10
       
       
End Select
 
End Function
