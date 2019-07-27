VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   0  'None
   Caption         =   "Invocar"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2175
      Left            =   420
      TabIndex        =   0
      Top             =   495
      Width           =   2355
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   120
      Top             =   2910
      Width           =   855
   End
   Begin VB.Image imgInvocar 
      Height          =   375
      Left            =   1800
      Top             =   2910
      Width           =   1335
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonInvocar    As clsGraphicalButton

Private cBotonSalir      As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    
        Me.Picture = LoadPicture(DirInterfaces & "VentanaMsj.jpg")
    
        Call LoadButtons

End Sub

Private Sub LoadButtons()

        Dim grhpath As String
    
        grhpath = DirInterfaces & "Buttons/"

        Set cBotonInvocar = New clsGraphicalButton
        Set cBotonSalir = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonInvocar.Initialize(imgInvocar, grhpath & "boton tirar.bmp", _
         grhpath & "boton Tirar 1.bmp", _
         grhpath & "boton Tirar 1.bmp", Me)

        Call cBotonSalir.Initialize(imgSalir, grhpath & "Boton Salir.bmp", _
           grhpath & "Boton Salir 1.bmp", _
           grhpath & "Boton Salir 1.bmp", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgInvocar_Click()
        Call WriteSpawnCreature(lstCriaturas.listIndex + 1)

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
