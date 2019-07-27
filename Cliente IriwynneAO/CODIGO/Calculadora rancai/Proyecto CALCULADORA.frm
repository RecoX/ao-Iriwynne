VERSION 5.00
Begin VB.Form frmCalcu 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Proyecto CALCULADORA.frx":0000
   ScaleHeight     =   3660
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cboCLASE 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Proyecto CALCULADORA.frx":ED8E
      Left            =   3480
      List            =   "Proyecto CALCULADORA.frx":EDB6
      TabIndex        =   4
      Text            =   "Clerigo"
      Top             =   2000
      Width           =   2055
   End
   Begin VB.ComboBox cboRAZA 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Proyecto CALCULADORA.frx":EE22
      Left            =   3480
      List            =   "Proyecto CALCULADORA.frx":EE35
      TabIndex        =   3
      Text            =   "HUMANO"
      Top             =   1390
      Width           =   2055
   End
   Begin VB.TextBox txtELV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtVidaInic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "20"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtVida 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1420
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   240
      Picture         =   "Proyecto CALCULADORA.frx":EE62
      Top             =   2880
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   4920
      Picture         =   "Proyecto CALCULADORA.frx":16A09
      Top             =   2880
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su nivel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Elija su raza:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   8
      Top             =   1120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Elija su clase:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valor fijo"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su vida actual"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   1120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCalcu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Clase As String
Dim Raza As String
Dim intVida As Integer
Dim intELV As Integer

Private Sub Image1_Click()

    On Error GoTo salto
    Raza = cboRAZA
    Clase = cboCLASE
    intVida = txtVida.Text
    intELV = txtELV.Text
    
    GoTo 1
    
salto:
        
    MsgBox "Debes llenar los casilleros correctamente."
    Exit Sub

1    If Not IsNumeric(txtELV.Text) Then
        MsgBox "Debes ingresar un nivel entre 1 y 47."
        Exit Sub
    End If

    If Not IsNumeric(txtVida.Text) Then
        MsgBox "Debes ingresar una vida válida."
        Exit Sub
    End If

    If Val(txtVida.Text) < 21 Then
        MsgBox "Debes ingresar una vida válida."
        Exit Sub
    End If

    If intELV <= 0 Or intELV > 47 Then
        MsgBox ("El nivel es incorrecto.")
        Exit Sub
    End If

    Dim M As Double 'Promedio segun qué pija es.
    
    If UCase(Raza) = "HUMANO" Then

        Select Case UCase$(Clase)
        Case "MAGO", "TRABAJADOR"
            M = 6.5
        Case "PALADIN", "CAZADOR", "PIRATA", "BANDIDO"
            M = 9.5
        Case "CLERIGO", "DRUIDA", "BARDO", "ASESINO"
            M = 8
        Case "GUERRERO"
            M = 10
        Case "LADRON"
            M = 7
        End Select

    End If

    If UCase(Raza) = "ELFO" Or UCase(Raza) = "ELFO OSCURO" Then

        Select Case UCase$(Clase)
        Case "MAGO", "TRABAJADOR"
            M = 6
        Case "PALADIN", "CAZADOR", "PIRATA", "BANDIDO"
            M = 9
        Case "CLERIGO", "DRUIDA", "BARDO", "ASESINO"
            M = 7.5
        Case "GUERRERO"
            M = 9.5
        Case "LADRON"
            M = 6.5
        End Select

    End If

    If UCase(Raza) = "ENANO" Then

        Select Case UCase$(Clase)
        Case "MAGO", "TRABAJADOR"
            M = 7
        Case "PALADIN", "CAZADOR", "PIRATA", "BANDIDO"
            M = 10
        Case "CLERIGO", "DRUIDA", "BARDO", "ASESINO"
            M = 8.8
        Case "GUERRERO"
            M = 10.5
        Case "LADRON"
            M = 7.5
        End Select

    End If

    If UCase(Raza) = "GNOMO" Then

        Select Case UCase$(Clase)
        Case "MAGO", "TRABAJADOR"
            M = 6
        Case "PALADIN", "CAZADOR", "PIRATA", "BANDIDO"
            M = 9
        Case "CLERIGO", "DRUIDA", "BARDO", "ASESINO"
            M = 7.5
        Case "GUERRERO"
            M = 9
        Case "LADRON"
            M = 6
        End Select

    End If


Dim Promedio As Double
Dim VidaPromedio As Integer
Dim Diferencia As Integer

Promedio = (intVida - 20) / (intELV - 1)
VidaPromedio = (M * (intELV - 1)) + 20
Diferencia = intVida - VidaPromedio + 5

MsgBox "Tu personaje con " & intVida & " de vida tiene un promedio de " & Round(Promedio, 2) & ". El promedio es de " & M ' deberias tener " & VidaPromedio & " de vida para estar en promedio justo"

Exit Sub

If Diferencia > 0 Then
MsgBox "FELICIDADES!!. Tu vida está " & Diferencia & " por encima de lo normal"
ElseIf Diferencia < 0 Then
MsgBox "NO TE DESANIMES!!. Tu vida está " & Diferencia & " por debajo de lo normal"
Else
MsgBox "SAFASTE!!. Tu vida está en promedio justo."
End If

 
End Sub

Private Sub Image2_Click()

End

End Sub
