VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Canjear Items Shop"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtCantMensajes 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "5"
      Top             =   2505
      Width           =   285
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   10
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Image imgChkActivarFragShooter 
      Height          =   255
      Left            =   480
      Top             =   3510
      Width           =   255
   End
   Begin VB.Image imgChkDesactivarFragShooter 
      Height          =   255
      Left            =   2490
      Top             =   3510
      Width           =   255
   End
   Begin VB.Image imgChkPantalla 
      Height          =   255
      Left            =   2010
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image imgChkConsola 
      Height          =   255
      Left            =   480
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image imgChkEfectosSonido 
      Height          =   255
      Left            =   480
      Top             =   1635
      Width           =   255
   End
   Begin VB.Image imgChkSonidos 
      Height          =   255
      Left            =   480
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image imgChkMusica 
      Height          =   255
      Left            =   480
      Top             =   990
      Width           =   255
   End
   Begin VB.Image imgCambiarPasswd 
      Height          =   285
      Left            =   1440
      Top             =   4920
      Width           =   2010
   End
   Begin VB.Image imgMsgPersonalizado 
      Height          =   285
      Left            =   2520
      Top             =   4320
      Width           =   2010
   End
   Begin VB.Image imgConfigTeclas 
      Height          =   285
      Left            =   360
      Top             =   4320
      Width           =   2010
   End
   Begin VB.Image imgSalir 
      Height          =   285
      Left            =   1440
      Top             =   5520
      Width           =   2010
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private clsFormulario          As clsFormMovementManager

Private cBotonConfigTeclas     As clsGraphicalButton

Private cBotonMsgPersonalizado As clsGraphicalButton

Private cBotonCambiarPasswd    As clsGraphicalButton

Private cBotonSalir            As clsGraphicalButton

Public LastButtonPressed       As clsGraphicalButton

Private picCheckBox            As Picture

Private bMusicActivated        As Boolean

Private bSoundActivated        As Boolean

Private bSoundEffectsActivated As Boolean

Private loading                As Boolean

Private Sub Command1_Click()
    frmCanjes.Show
        frmCanjes.List1.Clear
        
        Call WriteCanje
        
        Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgCambiarPasswd_Click()
        Call frmNewPassword.Show(vbModal, Me)

End Sub

Private Sub imgChkActivarFragShooter_Click()

        ClientSetup.bActive = True
        
        Set imgChkDesactivarFragShooter.Picture = Nothing
     
        imgChkActivarFragShooter.Picture = picCheckBox

End Sub

Private Sub imgChkDesactivarFragShooter_Click()

        ClientSetup.bActive = False
       
        Set imgChkActivarFragShooter.Picture = Nothing
       
        imgChkDesactivarFragShooter.Picture = picCheckBox

End Sub

Private Sub txtCantMensajes_Change()
        txtCantMensajes.Text = Val(txtCantMensajes.Text)
    
        If txtCantMensajes.Text > 0 Then
                DialogosClanes.CantidadDialogos = txtCantMensajes.Text
        Else
                txtCantMensajes.Text = 5

        End If

End Sub

Private Sub txtLevel_Change()

        ClientSetup.byMurderedLevel = 25

End Sub

Private Sub imgChkConsola_Click()
        DialogosClanes.Activo = False
    
        imgChkConsola.Picture = picCheckBox
        Set imgChkPantalla.Picture = Nothing

End Sub

Private Sub imgChkEfectosSonido_Click()

        If loading Then Exit Sub
    
        Call Audio.PlayWave(SND_CLICK)
        
        bSoundEffectsActivated = Not bSoundEffectsActivated
    
        Audio.SoundEffectsActivated = bSoundEffectsActivated
    
        If bSoundEffectsActivated Then
                imgChkEfectosSonido.Picture = picCheckBox
        Else
                Set imgChkEfectosSonido.Picture = Nothing

        End If
            
End Sub

Private Sub imgChkMusica_Click()

        If loading Then Exit Sub
    
        Call Audio.PlayWave(SND_CLICK)
    
        bMusicActivated = Not bMusicActivated
            
        If Not bMusicActivated Then
                Audio.MusicActivated = False
                Slider1(0).Enabled = False
                Set imgChkMusica.Picture = Nothing
        Else

                If Not Audio.MusicActivated Then  'Prevent the music from reloading
                        Audio.MusicActivated = True
                        Slider1(0).Enabled = True
                        Slider1(0).Value = Audio.MusicVolume

                End If
        
                imgChkMusica.Picture = picCheckBox

        End If

End Sub

Private Sub imgChkPantalla_Click()
        DialogosClanes.Activo = True
    
        imgChkPantalla.Picture = picCheckBox
        Set imgChkConsola.Picture = Nothing

End Sub

Private Sub imgChkSonidos_Click()

        If loading Then Exit Sub
    
        Call Audio.PlayWave(SND_CLICK)
    
        bSoundActivated = Not bSoundActivated
    
        If Not bSoundActivated Then
                Audio.SoundActivated = False
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
                Slider1(1).Enabled = False
        
                Set imgChkSonidos.Picture = Nothing
        Else
                Audio.SoundActivated = True
                Slider1(1).Enabled = True
                Slider1(1).Value = Audio.SoundVolume
        
                imgChkSonidos.Picture = picCheckBox

        End If

End Sub

Private Sub imgConfigTeclas_Click()

        If Not loading Then _
           Call Audio.PlayWave(SND_CLICK)
        Call frmCustomKeys.Show(vbModal, Me)

End Sub

Private Sub imgManual_Click()

        If Not loading Then _
           Call Audio.PlayWave(SND_CLICK)
        Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/manual/", "", App.path, SW_SHOWNORMAL)

End Sub

Private Sub imgMapa_Click()
        Call frmMapa.Show(vbModal, Me)

End Sub

Private Sub imgMsgPersonalizado_Click()
        Call frmMessageTxt.Show(vbModeless, Me)

End Sub

Private Sub imgSalir_Click()
        Unload Me
        frmMain.SetFocus

End Sub

Private Sub Form_Load()

        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    
        Me.Picture = LoadPicture(DirInterfaces & "VentanaOpciones.jpg")
        LoadButtons
    
        loading = True      'Prevent sounds when setting check's values
        LoadUserConfig
        loading = False     'Enable sounds when setting check's values

End Sub

Private Sub LoadButtons()

        Dim GrhPath As String
    
        GrhPath = DirInterfaces & "Buttons/"

        Set cBotonConfigTeclas = New clsGraphicalButton
        Set cBotonMsgPersonalizado = New clsGraphicalButton
        Set cBotonCambiarPasswd = New clsGraphicalButton
        Set cBotonSalir = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonConfigTeclas.Initialize(imgConfigTeclas, GrhPath & "boton Configurar teclas .jpg", _
           GrhPath & "boton Configurar teclas 1.jpg", _
           GrhPath & "boton Configurar teclas 1.jpg", Me)
                                    
        Call cBotonMsgPersonalizado.Initialize(imgMsgPersonalizado, GrhPath & "boton Mensajes .jpg", _
           GrhPath & "boton Mensajes 1.jpg", _
           GrhPath & "boton Mensajes 1.jpg", Me)
                                    
        Call cBotonCambiarPasswd.Initialize(imgCambiarPasswd, GrhPath & "boton Cambiar contraseña.jpg", _
           GrhPath & "boton Cambiar contraseña 1.jpg", _
           GrhPath & "boton Cambiar contraseña 1.jpg", Me)
                                                                        
        Call cBotonSalir.Initialize(imgSalir, GrhPath & "Boton Salir.bmp", _
           GrhPath & "Boton Salir 1.bmp", _
           GrhPath & "Boton Salir 1.bmp", Me)
                                    
        Set picCheckBox = LoadPicture(DirInterfaces & "CheckBoxOpciones.jpg")

End Sub

Private Sub LoadUserConfig()

        ' Load music config
        bMusicActivated = Audio.MusicActivated
        Slider1(0).Enabled = bMusicActivated
    
        If bMusicActivated Then
                imgChkMusica.Picture = picCheckBox
        
                Slider1(0).Value = Audio.MusicVolume

        End If
    
        ' Load Sound config
        bSoundActivated = Audio.SoundActivated
        Slider1(1).Enabled = bSoundActivated
    
        If bSoundActivated Then
                imgChkSonidos.Picture = picCheckBox
        
                Slider1(1).Value = Audio.SoundVolume

        End If
    
        ' Load Sound Effects config
        bSoundEffectsActivated = Audio.SoundEffectsActivated

        If bSoundEffectsActivated Then imgChkEfectosSonido.Picture = picCheckBox
    
        txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)
    
        If DialogosClanes.Activo Then
                imgChkPantalla.Picture = picCheckBox
        Else
                imgChkConsola.Picture = picCheckBox

        End If
        
        If ClientSetup.bKill Then imgChkActivarFragShooter.Picture = picCheckBox
        If Not ClientSetup.bActive Then imgChkDesactivarFragShooter.Picture = picCheckBox

End Sub

Private Sub Slider1_Change(Index As Integer)

        Select Case Index

                Case 0
                        Audio.MusicVolume = Slider1(0).Value

                Case 1
                        Audio.SoundVolume = Slider1(1).Value

        End Select

End Sub

Private Sub Slider1_Scroll(Index As Integer)

        Select Case Index

                Case 0
                        Audio.MusicVolume = Slider1(0).Value

                Case 1
                        Audio.SoundVolume = Slider1(1).Value

        End Select

End Sub
