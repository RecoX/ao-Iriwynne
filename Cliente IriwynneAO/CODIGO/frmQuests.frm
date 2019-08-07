VERSION 5.00
Begin VB.Form frmQuests 
   BorderStyle     =   0  'None
   Caption         =   "Misiones"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10050
   Icon            =   "frmQuests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuests.frx":000C
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cAbandonar 
      Caption         =   "Abandonar Quest"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   6600
      Width           =   2895
   End
   Begin VB.CommandButton cVolver 
      Caption         =   "Volver"
      Height          =   735
      Left            =   7320
      TabIndex        =   3
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5895
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   6135
   End
   Begin VB.ListBox lstQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5880
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quests"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   810
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub cAbandonar_Click()
    Call Audio.PlayWave(SND_CLICK)

    'Chequeamos si hay items.
    If lstQuests.ListCount = 0 Then
        MsgBox "�No tienes ninguna misi�n!", vbOKOnly + vbExclamation
        Exit Sub
    End If

    'Chequeamos si tiene algun item seleccionado.
    If lstQuests.listIndex < 0 Then
        MsgBox "�Primero debes seleccionar una misi�n!", vbOKOnly + vbExclamation
        Exit Sub
    End If

    Select Case MsgBox("�Est�s seguro que deseas abandonar la misi�n?", vbYesNo + vbExclamation)
        Case vbYes  'Bot�n S�.
            'Enviamos el paquete para abandonar la quest
            Call WriteQuestAbandon(lstQuests.listIndex + 1)

        Case vbNo   'Bot�n NO.
            'Como seleccion� que no, no hace nada.
            Exit Sub
    End Select

End Sub

Private Sub cVolver_Click()

    Call Audio.PlayWave(SND_CLICK)
    Unload Me

End Sub

Private Sub Form_Load()
' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

'    Me.Picture = LoadPicture(DirGUI & "frmCargando.jpg")    ' TODO: Falta una ventana para esto


End Sub

Private Sub lstQuests_Click()

    If lstQuests.listIndex < 0 Then Exit Sub

    Call WriteQuestDetailsRequest(lstQuests.listIndex + 1)
End Sub
