Attribute VB_Name = "Mod_rDamage"
Option Explicit
 
Const DAMAGE_TIME   As Integer = 50
Const DAMAGE_FONT_S As Byte = 14
 
Enum EDType
     edPu�al = 1                'Apu�alo.
     edNormal = 2               'Hechizo o golpe com�n.
End Enum
 
Private DNormalFont    As New StdFont
 
Type DList
     DamageVal      As Integer  'Cantidad de da�o.
     ColorRGB       As Long     'Color.
     DamageType     As EDType   'Tipo, se usa para saber si es apu o no.
     DamageFont     As New StdFont  'Efecto del apu.
     TimeRendered   As Integer  'Tiempo transcurrido.
     Downloading    As Byte     'Contador para la posicion Y.
     Activated      As Boolean  'Si est� activado..
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
 
' @ Agrega un nuevo da�o.
 
With MapData(X, Y).Damage
     
     .Activated = True
     .ColorRGB = ColorRGB
     .DamageType = edMode
     .DamageVal = DamageValue
     .TimeRendered = 0
     .Downloading = 0
     
     If .DamageType = EDType.edPu�al Then
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
 
' @ Dibuja un da�o
 
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
           If .DamageType = EDType.edPu�al Then
              .DamageFont.size = NewSize(.TimeRendered)
           End If
               
           'Dibujo ; D
            If .DamageType <> EDType.edNormal And .DamageType <> EDType.edPu�al Then
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
                   
       Case EDType.edPu�al
            ModifyColour = RGB(200, 200, 100)
           'ModifyColour = GetPu�alNewColour()
                   
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
