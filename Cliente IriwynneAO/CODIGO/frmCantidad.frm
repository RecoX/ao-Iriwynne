VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1470
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
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
      Height          =   315
      Left            =   480
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   2190
   End
   Begin VB.Image imgTirarTodo 
      Height          =   375
      Left            =   1680
      Tag             =   "1"
      Top             =   975
      Width           =   1335
   End
   Begin VB.Image imgTirar 
      Height          =   375
      Left            =   210
      Tag             =   "1"
      Top             =   975
      Width           =   1335
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonTirar      As clsGraphicalButton

Private cBotonTirarTodo  As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    
        Me.Picture = LoadPicture(DirInterfaces & "VentanaTirarOro.jpg")
    
        Call LoadButtons

End Sub

Private Sub LoadButtons()

        Dim grhpath As String
    
        grhpath = DirInterfaces & "Buttons/"
    
        Set cBotonTirar = New clsGraphicalButton
        Set cBotonTirarTodo = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton

        Call cBotonTirar.Initialize(imgTirar, grhpath & "Boton Tirar.bmp", grhpath & "Boton Tirar 1.bmp", grhpath & "Boton Tirar 1.bmp", Me)
           
        Call cBotonTirarTodo.Initialize(imgTirarTodo, grhpath & "Boton Tirar Todo.bmp", grhpath & "Boton Tirar Todo 1.bmp", grhpath & "Boton Tirar Todo 1.bmp", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgTirar_Click()

        If LenB(txtCantidad.Text) > 0 Then
                If Not IsNumeric(txtCantidad.Text) Then Exit Sub  'Should never happen
        
                Call WriteDrop(Inventario.SelectedItem, frmCantidad.txtCantidad.Text)
                frmCantidad.txtCantidad.Text = ""

        End If
    
        Unload Me

End Sub

Private Sub imgTirarTodo_Click()

        If Inventario.SelectedItem = 0 Then Exit Sub
    
        If Inventario.SelectedItem <> FLAGORO Then
                Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem))
                Unload Me
        Else

                If UserGLD > 10000 Then
                        Call WriteDrop(Inventario.SelectedItem, 10000)
                        Unload Me
                Else
                        Call WriteDrop(Inventario.SelectedItem, UserGLD)
                        Unload Me

                End If

        End If

        frmCantidad.txtCantidad.Text = ""

End Sub

Private Sub txtCantidad_Change()

        On Error GoTo errhandler

        If Val(txtCantidad.Text) < 0 Then
                txtCantidad.Text = "1"

        End If
    
        If Val(txtCantidad.Text) > MAX_INVENTORY_OBJS Then
                txtCantidad.Text = "10000"

        End If
    
        Exit Sub
    
errhandler:
        'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
        txtCantidad.Text = "1"

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

        If (KeyAscii <> 8) Then
                If (KeyAscii < 48 Or KeyAscii > 57) Then
                        KeyAscii = 0

                End If

        End If

End Sub

Private Sub txtCantidad_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub
