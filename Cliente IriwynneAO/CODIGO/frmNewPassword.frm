VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contrase�a"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2265
      Width           =   4005
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1545
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   825
      Width           =   4005
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2640
      Tag             =   "1"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image imgAceptar 
      Height          =   375
      Left            =   600
      Tag             =   "1"
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonAceptar    As clsGraphicalButton
Private cBotonSalir    As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    
        Me.Picture = LoadPicture(DirInterfaces & "VentanaCambiarcontrasenia.jpg")
    
        Call LoadButtons

End Sub

Private Sub LoadButtons()

        Dim grhpath As String
    
        grhpath = DirInterfaces & "buttons/"

        Set cBotonAceptar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonAceptar.Initialize(imgAceptar, grhpath & "boton Cambiar contrase�a.jpg", _
           grhpath & "boton Cambiar contrase�a 1.jpg", _
           grhpath & "boton Cambiar contrase�a 1.jpg", Me)
           
        Set cBotonSalir = New clsGraphicalButton
    
        Call cBotonSalir.Initialize(Image1, grhpath & "boton cancelar.bmp", _
           grhpath & "boton cancelar 1.bmp", _
           grhpath & "boton cancelar 1.bmp", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub Image1_Click()
Unload Me
frmMain.SetFocus

End Sub

Private Sub imgAceptar_Click()

        If Text2.Text <> Text3.Text Then
                Call MsgBox("Las contrase�as no coinciden", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contrase�a")
                Exit Sub

        End If
    
        Call WriteChangePassword(Text1.Text, Text2.Text)
        Unload Me

End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub
