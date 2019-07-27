VERSION 5.00
Begin VB.Form frmEntrenador 
   BorderStyle     =   0  'None
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2175
      Left            =   870
      TabIndex        =   0
      Top             =   675
      Width           =   2355
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   2160
      Tag             =   "1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Image imgLuchar 
      Height          =   375
      Left            =   600
      Tag             =   "1"
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "frmEntrenador"
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

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^,
'   le puse borde a la ventana y le cambié la letra a
'   una más linda :)
'
'[END]'

Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonLuchar     As clsGraphicalButton

Private cBotonSalir      As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
        
        Me.Picture = LoadPicture(DirInterfaces & "VentanaEntrenador.jpg")
    
        Call LoadButtons
    
End Sub

Private Sub LoadButtons()

        Dim grhpath As String
    
        grhpath = DirInterfaces

        Set cBotonLuchar = New clsGraphicalButton
        Set cBotonSalir = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonLuchar.Initialize(imgLuchar, grhpath & "Buttons/Boton Siguiente.jpg", _
         grhpath & "buttons/Boton Siguiente 1.jpg", _
         grhpath & "Buttons/Boton Siguiente 1.jpg", Me)

        Call cBotonSalir.Initialize(imgSalir, grhpath & "Buttons/boton salir.bmp", _
         grhpath & "Buttons/boton salir 1.bmp", _
         grhpath & "Buttons/boton salir 1.bmp", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgLuchar_Click()
        LastNpcSummoned = lstCriaturas.listIndex + 1
        Call WriteTrain(LastNpcSummoned)
        Unload Me

End Sub

Private Sub imgSalir_Click()
        Unload Me

End Sub

Private Sub lstCriaturas_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub
