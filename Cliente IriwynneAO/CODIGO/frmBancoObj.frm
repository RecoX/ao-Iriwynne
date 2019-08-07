VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "frmBancoObj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBancoObj.frx":1CCA
   ScaleHeight     =   530
   ScaleMode       =   0  'User
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox CantidadOro 
      Alignment       =   2  'Center
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
      Height          =   270
      Left            =   3525
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "1"
      Top             =   1410
      Width           =   1035
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
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
      Height          =   270
      Left            =   3195
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "1"
      Top             =   3930
      Width           =   615
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   465
      ScaleHeight     =   3990
      ScaleWidth      =   2565
      TabIndex        =   4
      Top             =   2400
      Width           =   2595
   End
   Begin VB.PictureBox PicInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   4095
      ScaleHeight     =   17.61
      ScaleMode       =   0  'User
      ScaleWidth      =   882.58
      TabIndex        =   3
      Top             =   2400
      Width           =   2595
   End
   Begin VB.Label lblUserGld 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3600
      TabIndex        =   5
      Top             =   945
      Width           =   135
   End
   Begin VB.Image imgDepositarOro 
      Height          =   690
      Left            =   1200
      Tag             =   "0"
      Top             =   960
      Width           =   690
   End
   Begin VB.Image imgRetirarOro 
      Height          =   765
      Left            =   4695
      Tag             =   "0"
      Top             =   945
      Width           =   945
   End
   Begin VB.Image imgCerrar 
      Height          =   450
      Left            =   6480
      Picture         =   "frmBancoObj.frx":1BB29
      Tag             =   "0"
      Top             =   120
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   3360
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   3360
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   6990
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   1
      Top             =   7245
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   6750
      Width           =   750
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Private clsFormulario    As clsFormMovementManager

'Private cBotonRetirarOro   As clsGraphicalButton

'Private cBotonDepositarOro As clsGraphicalButton

Private cBotonCerrar     As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Public LasActionBuy      As Boolean

Public LastIndex1        As Integer

Public LastIndex2        As Integer

Public NoPuedeMover      As Boolean

Private Sub cantidad_Change()

        If Val(cantidad.Text) < 1 Then
                cantidad.Text = 1

        End If
    
        If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
                cantidad.Text = MAX_INVENTORY_OBJS

        End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

        If (KeyAscii <> 8) Then
                If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
                        KeyAscii = 0

                End If

        End If

End Sub

Private Sub CantidadOro_Change()

        If Val(CantidadOro.Text) < 1 Then
                cantidad.Text = 1

        End If

End Sub

Private Sub CantidadOro_KeyPress(KeyAscii As Integer)

        If (KeyAscii <> 8) Then
                If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
                        KeyAscii = 0

                End If

        End If

End Sub

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me

        'Cargamos la interfase
        Me.Picture = LoadPicture(DirInterfaces & "VentanaBoveda.jpg")
    
        Call LoadButtons
    
End Sub

Private Sub LoadButtons()

        Dim GrhPath As String
    
        GrhPath = DirInterfaces
        'CmdMoverBov(1).Picture = LoadPicture(DirInterfaces & "FlechaSubirObjeto.jpg")
        'CmdMoverBov(0).Picture = LoadPicture(DirInterfaces & "FlechaBajarObjeto.jpg")
    
        'Set cBotonRetirarOro = New clsGraphicalButton
        'Set cBotonDepositarOro = New clsGraphicalButton
        Set cBotonCerrar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton

        'Call cBotonDepositarOro.Initialize(imgDepositarOro, vbNullString, GrhPath & "BotonDepositaOroApretado.jpg", GrhPath & "BotonDepositaOroApretado.jpg", Me)
        'Call cBotonRetirarOro.Initialize(imgRetirarOro, vbNullString, GrhPath & "BotonRetirarOroApretado.jpg", GrhPath & "BotonRetirarOroApretado.jpg", Me)
        'Call cBotonCerrar.Initialize(imgCerrar, vbNullString, GrhPath & "xPrendida.bmp", GrhPath & "xPrendida.bmp", Me)
    
        Image1(0).MouseIcon = picMouseIcon
        Image1(1).MouseIcon = picMouseIcon
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call LastButtonPressed.ToggleToNormal

End Sub

Private Sub Image1_Click(Index As Integer)
    
        Call Audio.PlayWave(SND_CLICK)
    
        If InvBanco(Index).SelectedItem = 0 Then Exit Sub
    
        If Not IsNumeric(cantidad.Text) Then Exit Sub
    
        Select Case Index

                Case 0
                        LastIndex1 = InvBanco(0).SelectedItem
                        LasActionBuy = True
                        Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
            
                Case 1
                        LastIndex2 = InvBanco(1).SelectedItem
                        LasActionBuy = False
                        Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)

        End Select

End Sub

Private Sub imgDepositarOro_Click()
        Call WriteBankDepositGold(Val(CantidadOro.Text))

End Sub

Private Sub imgRetirarOro_Click()
        Call WriteBankExtractGold(Val(CantidadOro.Text))

End Sub

Private Sub Label2_Click()
        Call WriteBankEnd
        NoPuedeMover = False

End Sub

Private Sub PicBancoInv_Click()

        If InvBanco(0).SelectedItem <> 0 Then

                With UserBancoInventory(InvBanco(0).SelectedItem)
                        Label1(0).Caption = .Name
            
                        Select Case .ObjType

                                Case 2, 32
                                        Label1(1).Caption = "M�x Golpe:" & .MaxHit
                                        Label1(2).Caption = "M�n Golpe:" & .MinHit
                                        Label1(1).Visible = True
                                        Label1(2).Visible = True
                    
                                Case 3, 16, 17
                                        Label1(1).Caption = "M�x Defensa:" & .MaxDef
                                        Label1(2).Caption = "M�n Defensa:" & .MinDef
                                        Label1(1).Visible = True
                                        Label1(2).Visible = True
                    
                                Case Else
                                        Label1(1).Visible = False
                                        Label1(2).Visible = False
                    
                        End Select
            
                End With
        
        Else
                Label1(0).Caption = ""
                Label1(1).Visible = False
                Label1(2).Visible = False

        End If

End Sub

Private Sub PicBancoInv_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
        Call LastButtonPressed.ToggleToNormal

End Sub

Private Sub PicInv_Click()
    
        If InvBanco(1).SelectedItem <> 0 Then

                With Inventario
                        Label1(0).Caption = .ItemName(InvBanco(1).SelectedItem)
            
                        Select Case .ObjType(InvBanco(1).SelectedItem)

                                Case eObjType.otWeapon, eObjType.otFlechas
                                        Label1(1).Caption = "M�x Golpe:" & .MaxHit(InvBanco(1).SelectedItem)
                                        Label1(2).Caption = "M�n Golpe:" & .MinHit(InvBanco(1).SelectedItem)
                                        Label1(1).Visible = True
                                        Label1(2).Visible = True
                    
                                Case eObjType.otcasco, eObjType.otArmadura, eObjType.otescudo ' 3, 16, 17
                                        Label1(1).Caption = "M�x Defensa:" & .MaxDef(InvBanco(1).SelectedItem)
                                        Label1(2).Caption = "M�n Defensa:" & .MinDef(InvBanco(1).SelectedItem)
                                        Label1(1).Visible = True
                                        Label1(2).Visible = True
                    
                                Case Else
                                        Label1(1).Visible = False
                                        Label1(2).Visible = False
                    
                        End Select
            
                End With

        Else
                Label1(0).Caption = ""
                Label1(1).Visible = False
                Label1(2).Visible = False

        End If

End Sub

Private Sub PicInv_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
        Call LastButtonPressed.ToggleToNormal

End Sub

Private Sub ImgCerrar_Click()
        Call WriteBankEnd
        NoPuedeMover = False

End Sub
