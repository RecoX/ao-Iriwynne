VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblHelp 
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
      Height          =   1215
      Left            =   600
      TabIndex        =   21
      Top             =   4710
      Width           =   7815
   End
   Begin VB.Image imgCancelar 
      Height          =   360
      Left            =   510
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3495
      TabIndex        =   20
      Top             =   840
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3495
      TabIndex        =   19
      Top             =   1215
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3495
      TabIndex        =   18
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3495
      TabIndex        =   17
      Top             =   1950
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3495
      TabIndex        =   16
      Top             =   2325
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3495
      TabIndex        =   15
      Top             =   2700
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3495
      TabIndex        =   14
      Top             =   3075
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   3495
      TabIndex        =   13
      Top             =   3450
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3495
      TabIndex        =   12
      Top             =   3825
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   3495
      TabIndex        =   11
      Top             =   4200
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7635
      TabIndex        =   10
      Top             =   840
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   7635
      TabIndex        =   9
      Top             =   1215
      Width           =   405
   End
   Begin VB.Image imgMas1 
      Height          =   300
      Left            =   3960
      Top             =   780
      Width           =   300
   End
   Begin VB.Image imgMas2 
      Height          =   300
      Left            =   3960
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image imgMenos2 
      Height          =   300
      Left            =   3120
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image imgMas3 
      Height          =   300
      Left            =   3960
      Top             =   1515
      Width           =   300
   End
   Begin VB.Image imgMenos3 
      Height          =   300
      Left            =   3120
      Top             =   1515
      Width           =   300
   End
   Begin VB.Image imgMas4 
      Height          =   300
      Left            =   3960
      Top             =   1890
      Width           =   300
   End
   Begin VB.Image imgMenos4 
      Height          =   300
      Left            =   3120
      Top             =   1890
      Width           =   300
   End
   Begin VB.Image imgMas5 
      Height          =   300
      Left            =   3960
      Top             =   2265
      Width           =   300
   End
   Begin VB.Image imgMenos5 
      Height          =   300
      Left            =   3120
      Top             =   2265
      Width           =   300
   End
   Begin VB.Image imgMas6 
      Height          =   300
      Left            =   3960
      Top             =   2640
      Width           =   300
   End
   Begin VB.Image imgMenos6 
      Height          =   300
      Left            =   3120
      Top             =   2640
      Width           =   300
   End
   Begin VB.Image imgMas7 
      Height          =   300
      Left            =   3960
      Top             =   3015
      Width           =   300
   End
   Begin VB.Image imgMenos7 
      Height          =   300
      Left            =   3120
      Top             =   3015
      Width           =   300
   End
   Begin VB.Image imgMas8 
      Height          =   300
      Left            =   3960
      Top             =   3390
      Width           =   300
   End
   Begin VB.Image imgMenos8 
      Height          =   300
      Left            =   3120
      Top             =   3390
      Width           =   300
   End
   Begin VB.Image imgMas9 
      Height          =   300
      Left            =   3960
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image imgMenos9 
      Height          =   300
      Left            =   3120
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image imgMas10 
      Height          =   300
      Left            =   3960
      Top             =   4140
      Width           =   300
   End
   Begin VB.Image imgMenos10 
      Height          =   300
      Left            =   3120
      Top             =   4140
      Width           =   300
   End
   Begin VB.Image imgMas11 
      Height          =   285
      Left            =   8100
      Top             =   780
      Width           =   345
   End
   Begin VB.Image imgMenos11 
      Height          =   285
      Left            =   7260
      Top             =   780
      Width           =   345
   End
   Begin VB.Image imgMas12 
      Height          =   285
      Left            =   8100
      Top             =   1155
      Width           =   345
   End
   Begin VB.Image imgMenos12 
      Height          =   285
      Left            =   7260
      Top             =   1155
      Width           =   345
   End
   Begin VB.Image imgMas13 
      Height          =   285
      Left            =   8100
      Top             =   1515
      Width           =   345
   End
   Begin VB.Image imgMenos13 
      Height          =   285
      Left            =   7260
      Top             =   1515
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   7635
      TabIndex        =   8
      Top             =   1575
      Width           =   405
   End
   Begin VB.Image imgMas14 
      Height          =   285
      Left            =   8100
      Top             =   1890
      Width           =   345
   End
   Begin VB.Image imgMenos14 
      Height          =   285
      Left            =   7260
      Top             =   1890
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   7635
      TabIndex        =   7
      Top             =   1950
      Width           =   405
   End
   Begin VB.Image imgMas15 
      Height          =   285
      Left            =   8100
      Top             =   2265
      Width           =   345
   End
   Begin VB.Image imgMenos15 
      Height          =   285
      Left            =   7260
      Top             =   2265
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7635
      TabIndex        =   6
      Top             =   2325
      Width           =   405
   End
   Begin VB.Image imgMas16 
      Height          =   285
      Left            =   8100
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image imgMenos16 
      Height          =   285
      Left            =   7260
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7635
      TabIndex        =   5
      Top             =   2700
      Width           =   405
   End
   Begin VB.Image imgMas17 
      Height          =   285
      Left            =   8100
      Top             =   3015
      Width           =   345
   End
   Begin VB.Image imgMenos17 
      Height          =   285
      Left            =   7260
      Top             =   3015
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7635
      TabIndex        =   4
      Top             =   3075
      Width           =   405
   End
   Begin VB.Image imgMas18 
      Height          =   285
      Left            =   8100
      Top             =   3390
      Width           =   345
   End
   Begin VB.Image imgMenos18 
      Height          =   285
      Left            =   7260
      Top             =   3390
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7635
      TabIndex        =   3
      Top             =   3450
      Width           =   405
   End
   Begin VB.Image imgMenos1 
      Height          =   300
      Left            =   3120
      Top             =   780
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   7635
      TabIndex        =   2
      Top             =   3825
      Width           =   405
   End
   Begin VB.Image imgMas19 
      Height          =   285
      Left            =   8100
      Top             =   3765
      Width           =   345
   End
   Begin VB.Image imgMenos19 
      Height          =   285
      Left            =   7260
      Top             =   3765
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   7635
      TabIndex        =   1
      Top             =   4200
      Width           =   405
   End
   Begin VB.Image imgMas20 
      Height          =   285
      Left            =   8100
      Top             =   4140
      Width           =   345
   End
   Begin VB.Image imgMenos20 
      Height          =   285
      Left            =   7260
      Top             =   4140
      Width           =   345
   End
   Begin VB.Image imgAceptar 
      Height          =   360
      Left            =   6990
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   90
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario               As clsFormMovementManager

Private cBotonMas(1 To NUMSKILLS)   As clsGraphicalButton

Private cBotonMenos(1 To NUMSKILLS) As clsGraphicalButton

Private cBtonAceptar                As clsGraphicalButton

Private cBotonCancelar              As clsGraphicalButton

Public LastButtonPressed            As clsGraphicalButton

Private bPuedeMagia                 As Boolean

Private bPuedeMeditar               As Boolean

Private bPuedeEscudo                As Boolean

Private bPuedeCombateDistancia      As Boolean

Private vsHelp(1 To NUMSKILLS)      As String

Private Sub Form_Load()
    
        MirandoAsignarSkills = True
    
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    
        'Flags para saber que skills se modificaron
        ReDim flags(1 To NUMSKILLS)
    
        Call ValidarSkills
    
        Me.Picture = LoadPicture(DirInterfaces & "VentanaSkills.jpg")
        Call LoadButtons
    
        Call LoadHelp

End Sub

Private Sub LoadButtons()

        Dim grhpath As String

        Dim i       As Long
    
        grhpath = DirInterfaces

        For i = 1 To NUMSKILLS
                Set cBotonMas(i) = New clsGraphicalButton
                Set cBotonMenos(i) = New clsGraphicalButton
        Next i
    
        Set cBtonAceptar = New clsGraphicalButton
        Set cBotonCancelar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBtonAceptar.Initialize(imgAceptar, grhpath & "buttons/Boton Aceptar.bmp", _
           grhpath & "buttons/Boton Aceptar 1.bmp", _
           grhpath & "buttons/Boton Aceptar 1.bmp", Me)

        Call cBotonCancelar.Initialize(imgCancelar, grhpath & "buttons/Boton Cancelar.bmp", _
           grhpath & "buttons/Boton Cancelar 1.bmp", _
           grhpath & "buttons/Boton Cancelar 1.bmp", Me)

        Call cBotonMas(1).Initialize(imgMas1, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me, _
           grhpath & "BotonMasSkills.jpg", Not bPuedeMagia)

        Call cBotonMas(2).Initialize(imgMas2, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(3).Initialize(imgMas3, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(4).Initialize(imgMas4, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)
    
        Call cBotonMas(5).Initialize(imgMas5, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me, _
           grhpath & "BotonMasSkills.jpg", Not bPuedeMeditar)

        Call cBotonMas(6).Initialize(imgMas6, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(7).Initialize(imgMas7, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(8).Initialize(imgMas8, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)
    
        Call cBotonMas(9).Initialize(imgMas9, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(10).Initialize(imgMas10, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(11).Initialize(imgMas11, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me, _
           grhpath & "BotonMasSkills.jpg", Not bPuedeEscudo)

        Call cBotonMas(12).Initialize(imgMas12, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)
    
        Call cBotonMas(13).Initialize(imgMas13, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(14).Initialize(imgMas14, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(15).Initialize(imgMas15, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(16).Initialize(imgMas16, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)
    
        Call cBotonMas(17).Initialize(imgMas17, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(18).Initialize(imgMas18, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me, _
           grhpath & "BotonMasSkills.jpg", Not bPuedeCombateDistancia)

        Call cBotonMas(19).Initialize(imgMas19, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)

        Call cBotonMas(20).Initialize(imgMas20, grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasSkills.jpg", _
           grhpath & "BotonMasClickSkills.jpg", Me)
    
        Call cBotonMenos(1).Initialize(imgMenos1, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me, _
           grhpath & "BotonMenosSkills.jpg", Not bPuedeMagia)

        Call cBotonMenos(2).Initialize(imgMenos2, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(3).Initialize(imgMenos3, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(4).Initialize(imgMenos4, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)
    
        Call cBotonMenos(5).Initialize(imgMenos5, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me, _
           grhpath & "BotonMenosSkills.jpg", Not bPuedeMeditar)

        Call cBotonMenos(6).Initialize(imgMenos6, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(7).Initialize(imgMenos7, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(8).Initialize(imgMenos8, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)
    
        Call cBotonMenos(9).Initialize(imgMenos9, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(10).Initialize(imgMenos10, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(11).Initialize(imgMenos11, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me, _
           grhpath & "BotonMenosSkills.jpg", Not bPuedeEscudo)

        Call cBotonMenos(12).Initialize(imgMenos12, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)
    
        Call cBotonMenos(13).Initialize(imgMenos13, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(14).Initialize(imgMenos14, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(15).Initialize(imgMenos15, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(16).Initialize(imgMenos16, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)
    
        Call cBotonMenos(17).Initialize(imgMenos17, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(18).Initialize(imgMenos18, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me, _
           grhpath & "BotonMenosSkills.jpg", Not bPuedeCombateDistancia)

        Call cBotonMenos(19).Initialize(imgMenos19, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

        Call cBotonMenos(20).Initialize(imgMenos20, grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosSkills.jpg", _
           grhpath & "BotonMenosClickSkills.jpg", Me)

End Sub

Private Sub SumarSkillPoint(ByVal SkillIndex As Integer)

        If Alocados > 0 Then

                If Val(Text1(SkillIndex).Caption) < MAXSKILLPOINTS Then
                        Text1(SkillIndex).Caption = Val(Text1(SkillIndex).Caption) + 1
                        flags(SkillIndex) = flags(SkillIndex) + 1
                        Alocados = Alocados - 1

                End If
            
        End If
    
        puntos.Caption = Alocados

End Sub

Private Sub RestarSkillPoint(ByVal SkillIndex As Integer)

        If Alocados < SkillPoints Then
        
                If Val(Text1(SkillIndex).Caption) > 0 And flags(SkillIndex) > 0 Then
                        Text1(SkillIndex).Caption = Val(Text1(SkillIndex).Caption) - 1
                        flags(SkillIndex) = flags(SkillIndex) - 1
                        Alocados = Alocados + 1

                End If

        End If
    
        puntos.Caption = Alocados

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal
        lblHelp.Caption = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
        MirandoAsignarSkills = False

End Sub

Private Sub imgAceptar_Click()

        Dim skillChanges(NUMSKILLS) As Byte

        Dim i                       As Long

        For i = 1 To NUMSKILLS
                skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
                'Actualizamos nuestros datos locales
                UserSkills(i) = Val(Text1(i).Caption)
        Next i
    
        Call WriteModifySkills(skillChanges())
    
        SkillPoints = Alocados
    
        Unload Me

End Sub

Private Sub imgApunialar_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
                                   
        Call ShowHelp(eSkill.Apu�alar)

End Sub

Private Sub imgCancelar_Click()
        Unload Me

End Sub

Private Sub imgCarpinteria_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
        Call ShowHelp(eSkill.Carpinteria)

End Sub

Private Sub imgCombateArmas_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
        Call ShowHelp(eSkill.Armas)

End Sub

Private Sub imgCombateDistancia_MouseMove(Button As Integer, _
                                          Shift As Integer, _
                                          X As Single, _
                                          Y As Single)
        Call ShowHelp(eSkill.Proyectiles)

End Sub

Private Sub imgCombateSinArmas_MouseMove(Button As Integer, _
                                         Shift As Integer, _
                                         X As Single, _
                                         Y As Single)
        Call ShowHelp(eSkill.Wrestling)

End Sub

Private Sub imgComercio_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
        Call ShowHelp(eSkill.Comerciar)

End Sub

Private Sub imgDomar_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
        Call ShowHelp(eSkill.Domar)

End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
        Call ShowHelp(eSkill.Defensa)

End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
        Call ShowHelp(eSkill.Tacticas)

End Sub

Private Sub imgHerreria_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
        Call ShowHelp(eSkill.Herreria)

End Sub

Private Sub imgLiderazgo_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
        Call ShowHelp(eSkill.Liderazgo)

End Sub

Private Sub imgMagia_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
        Call ShowHelp(eSkill.Magia)

End Sub

Private Sub imgMas1_Click()
        Call SumarSkillPoint(1)

End Sub

Private Sub imgMas10_Click()
        Call SumarSkillPoint(10)

End Sub

Private Sub imgMas11_Click()
        Call SumarSkillPoint(11)

End Sub

Private Sub imgMas12_Click()
        Call SumarSkillPoint(12)

End Sub

Private Sub imgMas13_Click()
        Call SumarSkillPoint(13)

End Sub

Private Sub imgMas14_Click()
        Call SumarSkillPoint(14)

End Sub

Private Sub imgMas15_Click()
        Call SumarSkillPoint(15)

End Sub

Private Sub imgMas16_Click()
        Call SumarSkillPoint(16)

End Sub

Private Sub imgMas17_Click()
        Call SumarSkillPoint(17)

End Sub

Private Sub imgMas18_Click()
        Call SumarSkillPoint(18)

End Sub

Private Sub imgMas19_Click()
        Call SumarSkillPoint(19)

End Sub

Private Sub imgMas2_Click()
        Call SumarSkillPoint(2)

End Sub

Private Sub imgMas20_Click()
        Call SumarSkillPoint(20)

End Sub

Private Sub imgMas3_Click()
        Call SumarSkillPoint(3)

End Sub

Private Sub imgMas4_Click()
        Call SumarSkillPoint(4)

End Sub

Private Sub imgMas5_Click()
        Call SumarSkillPoint(5)

End Sub

Private Sub imgMas6_Click()
        Call SumarSkillPoint(6)

End Sub

Private Sub imgMas7_Click()
        Call SumarSkillPoint(7)

End Sub

Private Sub imgMas8_Click()
        Call SumarSkillPoint(8)

End Sub

Private Sub imgMas9_Click()
        Call SumarSkillPoint(9)

End Sub

Private Sub imgMeditar_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
        Call ShowHelp(eSkill.Meditar)

End Sub

Private Sub imgMenos1_Click()
        Call RestarSkillPoint(1)

End Sub

Private Sub imgMenos10_Click()
        Call RestarSkillPoint(10)

End Sub

Private Sub imgMenos11_Click()
        Call RestarSkillPoint(11)

End Sub

Private Sub imgMenos12_Click()
        Call RestarSkillPoint(12)

End Sub

Private Sub imgMenos13_Click()
        Call RestarSkillPoint(13)

End Sub

Private Sub imgMenos14_Click()
        Call RestarSkillPoint(14)

End Sub

Private Sub imgMenos15_Click()
        Call RestarSkillPoint(15)

End Sub

Private Sub imgMenos16_Click()
        Call RestarSkillPoint(16)

End Sub

Private Sub imgMenos17_Click()
        Call RestarSkillPoint(17)

End Sub

Private Sub imgMenos18_Click()
        Call RestarSkillPoint(18)

End Sub

Private Sub imgMenos19_Click()
        Call RestarSkillPoint(19)

End Sub

Private Sub imgMenos2_Click()
        Call RestarSkillPoint(2)

End Sub

Private Sub imgMenos20_Click()
        Call RestarSkillPoint(20)

End Sub

Private Sub imgMenos3_Click()
        Call RestarSkillPoint(3)

End Sub

Private Sub imgMenos4_Click()
        Call RestarSkillPoint(4)

End Sub

Private Sub imgMenos5_Click()
        Call RestarSkillPoint(5)

End Sub

Private Sub imgMenos6_Click()
        Call RestarSkillPoint(6)

End Sub

Private Sub imgMenos7_Click()
        Call RestarSkillPoint(7)

End Sub

Private Sub imgMenos8_Click()
        Call RestarSkillPoint(8)

End Sub

Private Sub imgMenos9_Click()
        Call RestarSkillPoint(9)

End Sub

Private Sub LoadHelp()
    
        vsHelp(eSkill.Magia) = "Magia:" & vbCrLf & _
           "- Representa la habilidad de un personaje de las �reas m�gica." & vbCrLf & _
           "- Indica la variedad de hechizos que es capaz de dominar el personaje."

        If Not bPuedeMagia Then
                vsHelp(eSkill.Magia) = vsHelp(eSkill.Magia) & vbCrLf & _
                   "* Habilidad inhabilitada para tu clase."

        End If
    
        vsHelp(eSkill.Robar) = "Robar:" & vbCrLf & _
           "- Habilidades de hurto. Nunca por medio de la violencia." & vbCrLf & _
           "- Indica la probabilidad de �xito del personaje al intentar apoderarse de oro de otro, en caso de ser Ladr�n, tambien podr� apoderarse de items."
    
        vsHelp(eSkill.Tacticas) = "Evasi�n en Combate:" & vbCrLf & _
           "- Representa la habilidad general para moverse en combate entre golpes enemigos sin morir o tropezar en el intento." & vbCrLf & _
           "- Indica la posibilidad de evadir un golpe f�sico del personaje."
    
        vsHelp(eSkill.Armas) = "Combate con Armas:" & vbCrLf & _
           "- Representa la habilidad del personaje para manejar armas de combate cuerpo a cuerpo." & vbCrLf & _
           "- Indica la probabilidad de impactar al oponente con armas cuerpo a cuerpo."
    
        vsHelp(eSkill.Meditar) = "Meditar:" & vbCrLf & _
           "- Representa la capacidad del personaje de concentrarse para abstrarse dentro de su mente, y as� revitalizar su fuerza espiritual." & vbCrLf & _
           "- Indica la velocidad a la que el personaje recupera man� (Clases m�gicas)."
    
        If Not bPuedeMeditar Then
                vsHelp(eSkill.Meditar) = vsHelp(eSkill.Meditar) & vbCrLf & _
                   "* Habilidad inhabilitada para tu clase."

        End If

        vsHelp(eSkill.Apu�alar) = "Apu�alar:" & vbCrLf & _
           "- Representa la destreza para inflingir da�o grave con armas cortas." & vbCrLf & _
           "- Indica la posibilidad de apu�alar al enemigo en un ataque. El Asesino es la �nica clase que no necesitar� 10 skills para comenzar a entrenar esta habilidad."

        vsHelp(eSkill.Ocultarse) = "Ocultarse:" & vbCrLf & _
           "- La habilidad propia de un personaje para mimetizarse con el medio y evitar se perciba su presencia." & vbCrLf & _
           "- Indica la facilidad con la que uno puede desaparecer de la vista de los dem�s y por cuanto tiempo."
    
        vsHelp(eSkill.Supervivencia) = "Superivencia:" & vbCrLf & _
           "- Es el conjunto de habilidades necesarias para sobrevivir fuera de una ciudad en base a lo que la naturaleza ofrece." & vbCrLf & _
           "- Permite conocer la salud de las criaturas gui�ndose exclusivamente por su aspecto, as� como encender fogatas junto a las que descansar."
    
        vsHelp(eSkill.Talar) = "Talar:" & vbCrLf & _
           "- Es la habilidad en el uso del hacha para evitar desperdiciar le�a y maximizar la efectividad de cada golpe dado." & vbCrLf & _
           "- Indica la probabilidad de obtener le�a por golpe."
    
        vsHelp(eSkill.Comerciar) = "Comercio:" & vbCrLf & _
           "- Es la habilidad para regatear los precios exigidos en la compra y evitar ser regateado al vender." & vbCrLf & _
           "- Indica que tan caro se compra en el comercio con NPCs."
    
        vsHelp(eSkill.Defensa) = "Defensa con Escudos:" & vbCrLf & _
           "- Es la habilidad de interponer correctamente el escudo ante cada embate enemigo para evitar ser impactado sin perder el equilibrio y poder responder r�pidamente con la otra mano." & vbCrLf & _
           "- Indica las probabilidades de bloquear un impacto con el escudo."
    
        If Not bPuedeEscudo Then
                vsHelp(eSkill.Defensa) = vsHelp(eSkill.Defensa) & vbCrLf & _
                   "* Habilidad inhabilitada para tu clase."

        End If

        vsHelp(eSkill.Pesca) = "Pesca:" & vbCrLf & _
           "- Es el conjunto de conocimientos b�sicos para poder armar un se�uelo, poner la carnada en el anzuelo y saber d�nde buscar peces." & vbCrLf & _
           "- Indica la probabilidad de tener �xito en cada intento de pescar."
    
        vsHelp(eSkill.Mineria) = "Miner�a:" & vbCrLf & _
           "- Es el conjunto de conocimientos sobre los distintos minerales, el d�nde se obtienen, c�mo deben ser extra�dos y trabajados." & vbCrLf & _
           "- Indica la probabilidad de tener �xito en cada intento de minar y la capacidad, o no de convertir estos minerales en lingotes."
    
        vsHelp(eSkill.Carpinteria) = "Carpinter�a:" & vbCrLf & _
           "- Es el conjunto de conocimientos para saber serruchar, lijar, encolar y clavar madera con un buen nivel de terminaci�n." & vbCrLf & _
           "- Indica la habilidad en el manejo de estas herramientas, el que tan bueno se es en el oficio de carpintero."
    
        vsHelp(eSkill.Herreria) = "Herrer�a:" & vbCrLf & _
           "- Es el conjunto de conocimientos para saber procesar cada tipo de mineral para fundirlo, forjarlo y crear aleaciones." & vbCrLf & _
           "- Indica la habilidad en el manejo de estas t�cnicas, el que tan bueno se es en el oficio de herrero."
    
        vsHelp(eSkill.Liderazgo) = "Liderazgo:" & vbCrLf & _
           "- Es la habilidad propia del personaje para convencer a otros a seguirlo en batalla." & vbCrLf & _
           "- Permite crear clanes y partys"
    
        vsHelp(eSkill.Domar) = "Domar Animales:" & vbCrLf & _
           "- Es la habilidad en el trato con animales para que estos te sigan y ayuden en combate." & vbCrLf & _
           "- Indica la posibilidad de lograr domar a una criatura y qu� clases de criaturas se puede domar."
    
        vsHelp(eSkill.Proyectiles) = "Combate a distancia:" & vbCrLf & _
           "- Es el manejo de las armas de largo alcance." & vbCrLf & _
           "- Indica la probabilidad de �xito para impactar a un enemigo con este tipo de armas."
    
        If Not bPuedeCombateDistancia Then
                vsHelp(eSkill.Proyectiles) = vsHelp(eSkill.Proyectiles) & vbCrLf & _
                   "* Habilidad inhabilitada para tu clase."

        End If

        vsHelp(eSkill.Wrestling) = "Combate sin armas:" & vbCrLf & _
           "- Es la habilidad del personaje para entrar en combate sin arma alguna salvo sus propios brazos." & vbCrLf & _
           "- Indica la probabilidad de �xito para impactar a un enemigo estando desarmado. El Bandido y Ladr�n tienen habilidades extras asociadas a esta habilidad."
    
        vsHelp(eSkill.Navegacion) = "Navegaci�n:" & vbCrLf & _
           "- Es la habilidad para controlar barcos en el mar sin naufragar." & vbCrLf & _
           "- Indica que clase de barcos se pueden utilizar."
    
End Sub

Private Sub imgMineria_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
        Call ShowHelp(eSkill.Mineria)

End Sub

Private Sub imgNavegacion_MouseMove(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)
        Call ShowHelp(eSkill.Navegacion)

End Sub

Private Sub imgOcultarse_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
        Call ShowHelp(eSkill.Ocultarse)

End Sub

Private Sub imgPesca_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
        Call ShowHelp(eSkill.Pesca)

End Sub

Private Sub imgRobar_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
        Call ShowHelp(eSkill.Robar)

End Sub

Private Sub imgSupervivencia_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
        Call ShowHelp(eSkill.Supervivencia)

End Sub

Private Sub imgTalar_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
        Call ShowHelp(eSkill.Talar)

End Sub

Private Sub lblHelp_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub ShowHelp(ByVal eeSkill As eSkill)
        lblHelp.Caption = vsHelp(eeSkill)

End Sub

Private Sub ValidarSkills()

        bPuedeMagia = True
        bPuedeMeditar = True
        bPuedeEscudo = True
        bPuedeCombateDistancia = True

        Select Case UserClase

                Case eClass.Warrior, eClass.Hunter, eClass.Worker, eClass.Thief
                        bPuedeMagia = False
                        bPuedeMeditar = False
        
                Case eClass.Pirat
                        bPuedeMagia = False
                        bPuedeMeditar = False
                        bPuedeEscudo = False
        
                Case eClass.Mage, eClass.Druid
                        bPuedeEscudo = False
                        bPuedeCombateDistancia = False
            
        End Select
    
        ' Magia
        imgMas1.Enabled = bPuedeMagia
        imgMenos1.Enabled = bPuedeMagia

        ' Meditar
        imgMas5.Enabled = bPuedeMeditar
        imgMenos5.Enabled = bPuedeMeditar

        ' Escudos
        imgMas11.Enabled = bPuedeEscudo
        imgMenos11.Enabled = bPuedeEscudo

        ' Proyectiles
        imgMas18.Enabled = bPuedeCombateDistancia
        imgMenos18.Enabled = bPuedeCombateDistancia

End Sub
