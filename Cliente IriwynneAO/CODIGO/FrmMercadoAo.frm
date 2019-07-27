VERSION 5.00
Begin VB.Form FrmMercadoAo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMercadoAo.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstUsers 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2730
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Despublicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   1800
      Picture         =   "FrmMercadoAo.frx":9C1F
      Top             =   3840
      Width           =   1125
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   120
      Picture         =   "FrmMercadoAo.frx":CD4A
      Top             =   3840
      Width           =   1680
   End
   Begin VB.Image Image3 
      Height          =   540
      Left            =   600
      Picture         =   "FrmMercadoAo.frx":10C8E
      Top             =   2760
      Width           =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   1000
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Publicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   -120
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "FrmMercadoAo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
'Call ParseUserCommand("/VEO " & Text1.Text)
'Call ParseUserCommand("/PENAS " & Text1.Text)


    
End Sub

Private Sub Command4_Click()
'Call WriteRetos(Text1.Text, "2")

End Sub

Private Sub Command5_Click()
Call ParseUserCommand("/POSTEADOS")
End Sub


Private Sub Command7_Click()
'Call WriteRetos(Text1.Text, "1")


    
End Sub

Private Sub Form_Load()
Label2.Caption = lstUsers.Text
Call lstUsers_Click
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_Click()
'Call ParseUserCommand("/POSTE")


End Sub

Private Sub Image3_Click()
    Dim tStr As String
    Dim Nick As String
    Nick = lstUsers.Text
    
    If LenB(Nick) <> 0 Then
     '   tStr = InputBox("Confirma la visualizacion con tu PIN.", "Ingresa tu PIN")
        Call WriteRetos(Nick, "0", tStr)
    End If
    
    Call ParseUserCommand("/PENAS " & Nick)
End Sub

Private Sub Image4_Click()
    Dim tStr As String
    Dim Nick As String
    Nick = lstUsers.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Confirma el envio de solicitud con tu PIN.", "Ingresa tu PIN")
        Call WriteRetos(Nick, "1", tStr)
    End If
End Sub

Private Sub Image5_Click()

    Dim tStr As String
    Dim Nick As String
    Nick = lstUsers.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Confirma con tu PIN el intercambio.", "Ingresa tu PIN")
        Call WriteRetos(Nick, "2", tStr)
    End If
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Call writeRegresar(250)
End Sub

Private Sub Label4_Click()
Call writeRegresar(251)
End Sub

Private Sub lstUsers_Click()
Label2.Caption = lstUsers.Text
End Sub
