VERSION 5.00
Begin VB.Form frmRetos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "              Sistema de Retos"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   2970
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
   Picture         =   "frmRetos.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox bItems 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   1740
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      Top             =   3470
      Width           =   200
   End
   Begin VB.TextBox bGold 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Left            =   400
      TabIndex        =   5
      Top             =   2980
      Width           =   2175
   End
   Begin VB.OptionButton dRetos 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Index           =   1
      Left            =   970
      TabIndex        =   4
      Top             =   3480
      Width           =   200
   End
   Begin VB.OptionButton dRetos 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   3470
      Value           =   -1  'True
      Width           =   200
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   2
      Left            =   400
      TabIndex        =   2
      Top             =   2335
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   1
      Left            =   400
      TabIndex        =   1
      Top             =   1720
      Width           =   2160
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   0
      Left            =   400
      TabIndex        =   0
      Top             =   1075
      Width           =   2175
   End
   Begin VB.Image ImgMandar 
      Height          =   540
      Left            =   600
      Picture         =   "frmRetos.frx":11FDB
      Top             =   3840
      Width           =   1680
   End
   Begin VB.Image ImgCerrar 
      Height          =   210
      Left            =   2640
      Picture         =   "frmRetos.frx":15F1F
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonRetar   As clsGraphicalButton

Public LastButtonPressed      As clsGraphicalButton


Private Function CheckDatos() As Boolean
        ' @@ Chequeamos los Datos para no precesar mierda al pedo

        If dRetos(0).Value = True Then
      
                If Not Len(bName(1).Text) <> 0 Then
                        Call ShowConsoleMsg("Introduce el nombre de tu Enemigo.")
                        CheckDatos = False

                        Exit Function

                End If

                If Not IsNumeric(bGold.Text) Then
                        Call ShowConsoleMsg("Introduce el Oro en numeros.")
                        CheckDatos = False

                        Exit Function

                End If

        ElseIf dRetos(1).Value = True Then

                If Not Len(bName(0).Text) <> 0 Then
                        Call ShowConsoleMsg("Introduce el nombre de tu Compa�ero.")
                        CheckDatos = False

                        Exit Function

                End If

                If Not Len(bName(1).Text) <> 0 Then
                        Call ShowConsoleMsg("Introduce el nombre de tu Enemigo.")
                        CheckDatos = False

                        Exit Function

                End If

                If Not Len(bName(2).Text) <> 0 Then
                        Call ShowConsoleMsg("Introduce el nombre del compa�ero de tu Enemigo.")
                        CheckDatos = False

                        Exit Function

                End If

                If Not IsNumeric(bGold.Text) Then
                        Call ShowConsoleMsg("Introduce el Oro en numeros.")
                        CheckDatos = False

                        Exit Function

                End If
            
        End If

        CheckDatos = True

End Function

Private Sub dRetos_Click(Index As Integer)

        Select Case Index

                Case 0
                        dRetos(1).Value = False
                        dRetos(0).Value = True
                        
                        bName(0).Enabled = False
                        bName(1).Enabled = True
                        bName(2).Enabled = False
                  
                        'bName(0).BackColor = &H80000000
                        'bName(2).BackColor = &H80000000
                        'dRetos(1).Enabled = False

                Case 1
                        dRetos(1).Value = True
                        dRetos(0).Value = False
                        bName(0).Enabled = True
                        bName(2).Enabled = True
                  
                        'bName(0).BackColor = &HFFFFFF
                        'bName(2).BackColor = &HFFFFFF
                        'dRetos(0).Enabled = False
                  
        End Select

End Sub

Private Sub Form_Load()

        Me.Picture = LoadPicture(DirInterfaces & "VentanaRetos.jpg")

        ' @@ Empezamos con la opcion de 1 vs 1, por default
        Call dRetos_Click(0)
    Call LoadButtons
    
End Sub

Private Sub LoadButtons()

        Dim grhpath As String
    
        grhpath = DirInterfaces & "Buttons/"

        Set cBotonRetar = New clsGraphicalButton
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonRetar.Initialize(ImgMandar, grhpath & "boton enviar.bmp", _
           grhpath & "boton enviar 1.bmp", _
           grhpath & "boton enviar 1.bmp", Me)


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub


Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub ImgMandar_Click()

        If Not CheckDatos Then Exit Sub
      
        ' @@ Chequear el uso de Rtrim$(), puede llegar a ser mucho mejor usar Trim$()

        If dRetos(0).Value = True Then
                Call WriteOtherSendReto(RTrim$(bName(1).Text), Val(bGold.Text), (bItems.Value <> 0))
        ElseIf dRetos(1).Value = True Then
                Call WriteSendReto(RTrim$(bName(0).Text), RTrim$(bName(1).Text), RTrim$(bName(2).Text), Val(bGold.Text), (bItems.Value <> 0))
        End If

        Unload Me
        
End Sub

Private Sub Label1_Click()
ImgMandar_Click
End Sub

Private Sub Label2_Click()
Unload Me
frmMain.SetFocus

End Sub
