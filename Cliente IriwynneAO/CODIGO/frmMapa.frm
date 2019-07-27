VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMapa.frx":0000
   ScaleHeight     =   7035
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblMapaInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6600
      Width           =   7455
   End
   Begin VB.Label lblMapa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pasa por el mouse para obtener información acerca del mapa."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6300
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   93
      Left            =   5400
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   46
      Left            =   1800
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   45
      Left            =   1800
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   44
      Left            =   3240
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   43
      Left            =   3960
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   42
      Left            =   6120
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   41
      Left            =   6120
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   40
      Left            =   6120
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   39
      Left            =   5400
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   38
      Left            =   4680
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   37
      Left            =   3960
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   36
      Left            =   3240
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   35
      Left            =   2520
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   34
      Left            =   360
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   33
      Left            =   1080
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   32
      Left            =   1800
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   31
      Left            =   3240
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   30
      Left            =   2160
      Top             =   -120
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   29
      Left            =   3960
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   28
      Left            =   1800
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   27
      Left            =   2520
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   26
      Left            =   1800
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   25
      Left            =   2520
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   24
      Left            =   1080
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   23
      Left            =   1800
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   22
      Left            =   2520
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   21
      Left            =   1920
      Top             =   -120
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   20
      Left            =   1440
      Top             =   -240
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   19
      Left            =   4680
      Top             =   3960
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   18
      Left            =   3960
      Top             =   3960
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   17
      Left            =   6120
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   16
      Left            =   6120
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   15
      Left            =   5400
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   14
      Left            =   5400
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   13
      Left            =   4680
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   12
      Left            =   4680
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   11
      Left            =   3960
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   10
      Left            =   1800
      Top             =   3960
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   9
      Left            =   3960
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   8
      Left            =   3240
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   7
      Left            =   2520
      Top             =   3960
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   6
      Left            =   5400
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   5
      Left            =   3240
      Top             =   3960
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   4
      Left            =   5400
      Top             =   3960
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   3
      Left            =   3240
      Top             =   360
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   2
      Left            =   3240
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   735
      Index           =   1
      Left            =   3240
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image imgMap 
      Height          =   375
      Index           =   0
      Left            =   1080
      Top             =   0
      Width           =   855
   End
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   8040
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
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
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private Enum eMaps

        ieGeneral
        ieDungeon

End Enum


Private picMaps(1) As Picture

Private CurrentMap As eMaps
Private Const lblMapaCaption As String = "Nombre del mapa: "
Private Const lblMapaInfoCaption As String = "Información del mapa: "


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

''
' Toggle which image is visible.
'
Private Sub ToggleImgMaps()
        '*************************************************
        'Author: Marco Vanotti (MarKoxX)
        'Last modified: 24/07/08
        '
        '*************************************************

        '    imgToogleMap(CurrentMap).Visible = False
    
        '   If CurrentMap = eMaps.ieGeneral Then
        '           ImgCerrar.Visible = False
        '         CurrentMap = eMaps.ieDungeon
        '          ImgCerrar.Visible = True
        '          CurrentMap = eMaps.ieGeneral
        

        '  End If
    
        ' imgToogleMap(CurrentMap).Visible = True
        '  Me.Picture = picMaps(CurrentMap)

End Sub

''
' Load the images. Resizes the form, adjusts image's left and top and set lblTexto's Top and Left.
'
Private Sub Form_Load()
        '*************************************************
        'Author: Marco Vanotti (MarKoxX)
        'Last modified: 24/07/08
        '
        '*************************************************

        On Error GoTo error
    
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
        
        'Cargamos las imagenes de los mapas
        'Set picMaps(eMaps.ieGeneral) =
        'Set picMaps(eMaps.ieDungeon) = LoadPicture(DirInterfaces & "mapa1.jpg")
    
        ' Imagen de fondo
        'CurrentMap = eMaps.ieGeneral
        Me.Picture = LoadPicture(DirInterfaces & "mapa1.jpg")
    
        ImgCerrar.MouseIcon = picMouseIcon
        ' imgToogleMap(0).MouseIcon = picMouseIcon
        ' imgToogleMap(1).MouseIcon = picMouseIcon
    
        Exit Sub
error:
        MsgBox Err.Description, vbInformation, "Error: " & Err.number
        Unload Me

End Sub

Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub imgToogleMap_Click(Index As Integer)
        'ToggleImgMaps

End Sub

Private Sub imgMap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


Select Case Index

Case 1
lblMapaInfo.Caption = lblMapaInfoCaption & "Ciudad principal. Ningún bando ríge sobre éstas tierras."
lblMapa.Caption = lblMapaCaption & "Ullathorpe"
Case 2
lblMapaInfo.Caption = lblMapaInfoCaption & "Ciudad segura regída por la Armada Real"
lblMapa.Caption = lblMapaCaption & "Nix"
Case 3
lblMapaInfo.Caption = lblMapaInfoCaption & "Ciudad segura regída por la Armada Real"
lblMapa.Caption = lblMapaCaption & "Banderbille"
Case 4
lblMapaInfo.Caption = lblMapaInfoCaption & "Ciudad insegura regída por los Criminales"
lblMapa.Caption = lblMapaCaption & "Arkhein"
Case 6
lblMapaInfo.Caption = lblMapaInfoCaption & "Ciudad segura. Regída por la Armada Real"
lblMapa.Caption = lblMapaCaption & "Lindos"
Case 9
lblMapaInfo.Caption = lblMapaInfoCaption & "Mapa inseguro, hay yacimientos de: Hierro"
lblMapa.Caption = lblMapaCaption & "Minas de Hierro"
Case 10
lblMapaInfo.Caption = lblMapaInfoCaption & "Mapa inseguro, hay yacimientos de: Hierro, Plata"
lblMapa.Caption = lblMapaCaption & "Minas de Plata"
Case 16
lblMapaInfo.Caption = lblMapaInfoCaption & "Dungeon inseguro, hay: Dragon Rojo"
lblMapa.Caption = lblMapaCaption & "Dungeon Dragon"
Case 32
lblMapaInfo.Caption = lblMapaInfoCaption & "Dungeon inseguro, hay: Liche, Zombie, Orco, Araña Gigante, Demonio"
lblMapa.Caption = lblMapaCaption & "Dungeon Marabel"
Case 34
lblMapaInfo.Caption = lblMapaInfoCaption & "Dungeon inseguro, hay: NPCs acuáticos"
lblMapa.Caption = lblMapaCaption & "Dungeon Aqua"
Case 24
lblMapaInfo.Caption = lblMapaInfoCaption & "Mapa inseguro, hay yacimientos de: Oro"
lblMapa.Caption = lblMapaCaption & "Minas de Oro"
Case 39
lblMapaInfo.Caption = lblMapaInfoCaption & "Dungeon inseguro, hay: NPCs de fuego"
lblMapa.Caption = lblMapaCaption & "Dungeon Infierno"
Case 43
lblMapaInfo.Caption = lblMapaInfoCaption & "Mapa inseguro, ideal para clases no mágicas."
lblMapa.Caption = lblMapaCaption & "Montaña enana"
Case 46
lblMapaInfo.Caption = lblMapaInfoCaption & "Mapa inseguro, hay árboles élficos."
lblMapa.Caption = lblMapaCaption & "Bosque élfico"
Case 42
lblMapaInfo.Caption = lblMapaInfoCaption & "Dungeon inseguro, hay NPCs que dan mucha experiencia. (Medusas)"
lblMapa.Caption = lblMapaCaption & "Dungeon Veriil"
Case Else
lblMapaInfo.Caption = lblMapaInfoCaption & "Zona insegura"
lblMapa.Caption = lblMapaCaption & "Mapa inseguro"

End Select

End Sub

Private Sub Label1_Click()
Unload Me
frmMain.SetFocus

End Sub
