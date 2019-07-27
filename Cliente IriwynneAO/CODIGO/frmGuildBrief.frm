VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7620
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildBrief.frx":0000
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "&H8000000A&"
   Begin VB.TextBox Desc 
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
      Height          =   915
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   6090
      Width           =   6930
   End
   Begin VB.Label GuildLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel: "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Image imgSolicitarIngreso 
      Height          =   375
      Left            =   5850
      Tag             =   "1"
      Top             =   7095
      Width           =   1335
   End
   Begin VB.Image imgDeclararGuerra 
      Height          =   375
      Left            =   4560
      Tag             =   "1"
      Top             =   7125
      Width           =   1335
   End
   Begin VB.Image imgOfrecerAlianza 
      Height          =   375
      Left            =   3120
      Tag             =   "1"
      Top             =   7125
      Width           =   1335
   End
   Begin VB.Image imgOfrecerPaz 
      Height          =   375
      Left            =   1680
      Tag             =   "1"
      Top             =   7095
      Width           =   1335
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   120
      Tag             =   "1"
      Top             =   7095
      Width           =   1455
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   375
      TabIndex        =   18
      Top             =   3600
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   375
      TabIndex        =   17
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   375
      TabIndex        =   16
      Top             =   4080
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Index           =   3
      Left            =   375
      TabIndex        =   15
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Index           =   4
      Left            =   375
      TabIndex        =   14
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Index           =   5
      Left            =   375
      TabIndex        =   13
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Index           =   6
      Left            =   375
      TabIndex        =   12
      Top             =   5040
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Index           =   7
      Left            =   375
      TabIndex        =   11
      Top             =   5280
      Width           =   6735
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
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
      Left            =   1485
      TabIndex        =   10
      Top             =   555
      Width           =   4695
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
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
      Left            =   1605
      TabIndex        =   9
      Top             =   885
      Width           =   2775
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
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
      Left            =   6000
      TabIndex        =   8
      Top             =   885
      Width           =   1455
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
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
      Left            =   1155
      TabIndex        =   7
      Top             =   1215
      Width           =   3135
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
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
      Left            =   1575
      TabIndex        =   6
      Top             =   1545
      Width           =   2655
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
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
      Left            =   5115
      TabIndex        =   5
      Top             =   1545
      Width           =   1935
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
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
      Left            =   5160
      TabIndex        =   4
      Top             =   1215
      Width           =   1815
   End
   Begin VB.Label lblAlineacion 
      BackStyle       =   0  'Transparent
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
      Left            =   5130
      TabIndex        =   3
      Top             =   1875
      Width           =   1815
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
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
      Left            =   2310
      TabIndex        =   2
      Top             =   1875
      Width           =   2175
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
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
      Left            =   2160
      TabIndex        =   1
      Top             =   2205
      Width           =   1575
   End
   Begin VB.Label antifaccion 
      BackStyle       =   0  'Transparent
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
      Left            =   2910
      TabIndex        =   0
      Top             =   2535
      Width           =   2415
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario          As clsFormMovementManager

Private cBotonGuerra           As clsGraphicalButton

Private cBotonAlianza          As clsGraphicalButton

Private cBotonPaz              As clsGraphicalButton

Private cBotonSolicitarIngreso As clsGraphicalButton

Private cBotonCerrar           As clsGraphicalButton

Public LastButtonPressed       As clsGraphicalButton

Public EsLeader                As Boolean

Private Sub Desc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub Form_Load()

        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    
        Me.Picture = LoadPicture(DirInterfaces & "VentanaDetallesClan.jpg")
    
        Call LoadButtons
    
End Sub

Private Sub LoadButtons()

        Dim grhpath As String
    
        grhpath = DirInterfaces & "Buttons/"

        Set cBotonGuerra = New clsGraphicalButton
        Set cBotonAlianza = New clsGraphicalButton
        Set cBotonPaz = New clsGraphicalButton
        Set cBotonSolicitarIngreso = New clsGraphicalButton
        Set cBotonCerrar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonGuerra.Initialize(imgDeclararGuerra, grhpath & "boton Declarar guerra.jpg", _
           grhpath & "boton Declarar guerra 1.jpg", _
           grhpath & "boton Declarar guerra 1.jpg", Me)

        Call cBotonAlianza.Initialize(imgOfrecerAlianza, grhpath & "boton Ofrecer Alianza.jpg", _
           grhpath & "boton Ofrecer Alianza 1.jpg", _
           grhpath & "boton Ofrecer Alianza 1.jpg", Me)

        Call cBotonPaz.Initialize(imgOfrecerPaz, grhpath & "boton Ofrecer Paz.jpg", _
           grhpath & "boton Ofrecer Paz 1.jpg", _
           grhpath & "boton Ofrecer Paz 1.jpg", Me)

        Call cBotonSolicitarIngreso.Initialize(imgSolicitarIngreso, grhpath & "boton Solicitar ingreso.jpg", _
           grhpath & "boton Solicitar ingreso 1.jpg", _
           grhpath & "boton Solicitar ingreso 1.jpg", Me)

        Call cBotonCerrar.Initialize(ImgCerrar, grhpath & "boton Cerrar.bmp", _
           grhpath & "boton Cerrar 1.bmp", _
           grhpath & "boton Cerrar 1.bmp", Me)

        If Not EsLeader Then
                imgDeclararGuerra.Visible = False
                imgOfrecerAlianza.Visible = False
                imgOfrecerPaz.Visible = False

        End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub imgDeclararGuerra_Click()
        Call WriteGuildDeclareWar(Nombre.Caption)
        Unload Me

End Sub

Private Sub imgOfrecerAlianza_Click()
        frmCommet.Nombre = Nombre.Caption
        frmCommet.T = TIPO.ALIANZA
        Call frmCommet.Show(vbModal, frmGuildBrief)

End Sub

Private Sub imgOfrecerPaz_Click()
        frmCommet.Nombre = Nombre.Caption
        frmCommet.T = TIPO.PAZ
        Call frmCommet.Show(vbModal, frmGuildBrief)

End Sub

Private Sub imgSolicitarIngreso_Click()
        Call frmGuildSol.RecieveSolicitud(Nombre.Caption)
        Call frmGuildSol.Show(vbModal, frmGuildBrief)

End Sub

