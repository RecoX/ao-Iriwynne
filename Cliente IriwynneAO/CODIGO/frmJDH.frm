VERSION 5.00
Begin VB.Form frmJDH 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "JUEGOS DEL HAMBRE CUI"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fJDH 
      BackColor       =   &H00000000&
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtPremioORO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtInscripcionORO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtCupos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblCrearJuegos 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crear Juegos del hambre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label lblEventoCaptura 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evento JDH"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   525
         Index           =   0
         Left            =   720
         TabIndex        =   7
         Top             =   120
         Width           =   2400
      End
      Begin VB.Label lblInscripcionCTF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cupos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblPremiocTF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Premio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblInscripcionCTF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   1680
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmJDH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblCrearJuegos_Click()

    If Not IsNumeric(txtCupos.Text) Then 'Or Not IsNumeric(txtPremioORO.Text) Or Not IsNumeric(txtInscripcionORO.Text) Then
        Call MsgBox("El valor tiene que ser valor num�rico", vbCritical, "Error")
        Exit Sub
    End If

    'Call WriteEventCreate(e_Events.eJuegosDelHambre, CByte(txtCupos.Text), CLng(txtPremioORO.Text), CLng(txtInscripcionORO.Text))
    Call WriteCrearJDH(Val(CByte(txtCupos.Text)))
End Sub

