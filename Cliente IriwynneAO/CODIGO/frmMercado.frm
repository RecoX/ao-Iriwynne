VERSION 5.00
Begin VB.Form frmMercado 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMercado.frx":0000
   ScaleHeight     =   5355
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   8400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmMercado.frx":7BD1
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver Informacion"
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enviar Solicitud de Cambio"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox cboListaUsus 
      Height          =   3570
      Left            =   8400
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6480
      TabIndex        =   14
      Top             =   200
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   150
      TabIndex        =   13
      Top             =   3530
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   3840
      TabIndex        =   12
      Top             =   3525
      Width           =   3495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   150
      TabIndex        =   11
      Top             =   4650
      Width           =   3495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   3860
      TabIndex        =   10
      Top             =   4650
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   3860
      TabIndex        =   9
      Top             =   2340
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   150
      TabIndex        =   8
      Top             =   2340
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   3860
      TabIndex        =   7
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   150
      TabIndex        =   6
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Miembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmMercado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager

Private Sub cboListaUsus_Click()
    Dim nick As String
    nick = cboListaUsus.Text
    
Call WritePublicarMAO(nick)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim nick As String
nick = cboListaUsus.Text
frmCharInfo.frmType = CharInfoFrmType.frmMemberMercado
Call WriteGuildMemberInfo(nick)
End Sub

Private Sub Form_Load()
Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub Label9_Click()
Unload Me
End Sub
