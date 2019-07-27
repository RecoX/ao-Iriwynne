VERSION 5.00
Begin VB.Form frmRanking 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "rank"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "raking"
   Picture         =   "frmRanking.frx":0000
   ScaleHeight     =   3315
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   540
      Left            =   2640
      Picture         =   "frmRanking.frx":DDC7
      Top             =   2640
      Width           =   1680
   End
   Begin VB.Image ImgClanes 
      Height          =   540
      Left            =   960
      Picture         =   "frmRanking.frx":11D0B
      Top             =   2040
      Width           =   2640
   End
   Begin VB.Image ImgClanesHoras 
      Height          =   540
      Left            =   960
      Picture         =   "frmRanking.frx":1A112
      Top             =   1440
      Width           =   2640
   End
   Begin VB.Image ImgLevel 
      Height          =   540
      Left            =   960
      Picture         =   "frmRanking.frx":227F2
      Top             =   840
      Width           =   2640
   End
End
Attribute VB_Name = "frmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager


Private cBotonSalir            As clsGraphicalButton

Private cBotonHorasC             As clsGraphicalButton

Private cBotonNiveles            As clsGraphicalButton

Private cBotonNivelClan            As clsGraphicalButton

Public LastButtonPressed       As clsGraphicalButton

Private Sub Form_Load()

' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

LoadButtons
    
End Sub


Private Sub LoadButtons()
    
    
    Dim grhpath As String
    
    grhpath = DirInterfaces & "buttons/"
    Set LastButtonPressed = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Call cBotonSalir.Initialize(Image1, grhpath & "boton salir.bmp", _
    grhpath & "boton salir 1.bmp", _
    grhpath & "boton salir 1.bmp", Me)
    
    Set cBotonNivelClan = New clsGraphicalButton
    
    Call cBotonNivelClan.Initialize(ImgClanes, grhpath & "boton nivel clan.jpg", _
    grhpath & "boton nivel clan 1.jpg", _
    grhpath & "boton nivel clan 1.jpg", Me)
    
    Set cBotonNiveles = New clsGraphicalButton
    
    Call cBotonNiveles.Initialize(ImgLevel, grhpath & "Boton niveles.jpg", _
    grhpath & "Boton niveles 1.jpg", _
    grhpath & "Boton niveles 1.jpg", Me)
    
    Set cBotonHorasC = New clsGraphicalButton
    
    Call cBotonHorasC.Initialize(ImgClanesHoras, grhpath & "boton horasconquistadas.jpg", _
    grhpath & "boton horasconquistadas 1.jpg", _
    grhpath & "boton horasconquistadas 1.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LastButtonPressed.ToggleToNormal
End Sub

Private Sub Image1_Click()
frmMain.SetFocus
Unload Me
End Sub

Private Sub ImgLevel_Click()
'1 nivel - horas - nivel
CualPedi = 1
Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopLevel)
End Sub

           
Private Sub ImgClanesHoras_Click()
'1 nivel - horas - nivel
    CualPedi = 2
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopHorasClanes)
    
End Sub

Private Sub ImgClanes_Click()
'1 nivel - horas - nivel
CualPedi = 3
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopClanes)
End Sub

Private Sub Label1_Click()
    Unload Me
    frmMain.SetFocus
End Sub
