VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creaci�n de un Clan"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClanName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   360
      TabIndex        =   0
      Top             =   1815
      Width           =   3345
   End
   Begin VB.TextBox txtWeb 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   3345
   End
   Begin VB.Image imgSiguiente 
      Height          =   375
      Left            =   2400
      Tag             =   "1"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image imgCancelar 
      Height          =   375
      Left            =   240
      Tag             =   "1"
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonSiguiente  As clsGraphicalButton

Private cBotonCancelar   As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Deactivate()
        Me.SetFocus

End Sub

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me

        Me.Picture = LoadPicture(DirInterfaces & "VentanaNombreClan.jpg")
        
        Call LoadButtons

End Sub

Private Sub LoadButtons()

        Dim GrhPath As String
    
        GrhPath = DirInterfaces & "Buttons/"

        Set cBotonSiguiente = New clsGraphicalButton
        Set cBotonCancelar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonSiguiente.Initialize(imgSiguiente, GrhPath & "Boton Siguiente.jpg", _
           GrhPath & "Boton Siguiente 1.jpg", _
           GrhPath & "Boton Siguiente 1.jpg", Me)

        Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "Boton Cancelar.bmp", _
           GrhPath & "Boton Cancelar 1.bmp", _
           GrhPath & "Boton Cancelar 1.bmp", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgCancelar_Click()
        Unload Me

End Sub

Private Sub imgSiguiente_Click()

        If Len(txtClanName.Text) <= 30 Then
                If Not AsciiValidos(txtClanName.Text) Then
                        MsgBox "Nombre invalido."
                        Exit Sub

                End If

        Else
                MsgBox "Nombre demasiado extenso."
                Exit Sub

        End If
    
        'ClanName = complexNameToSimple(txtClanName.Text, True) 'string com�n
        'ClanName = Trim$(ClanName)
        'secClanName = txtClanName.Text 'string con caracteres especiales
        'Site = txtWeb.Text
        'Unload Me
        'frmGuildDetails.Show , frmMain
        
        ClanName = txtClanName.Text
        Site = txtWeb.Text
        Unload Me
        frmGuildDetails.Show , frmMain
    
End Sub

Private Sub txtWeb_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

