VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas de paz"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1785
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   240
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   4620
   End
   Begin VB.Image imgRechazar 
      Height          =   480
      Left            =   3840
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgAceptar 
      Height          =   480
      Left            =   2640
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgDetalle 
      Height          =   480
      Left            =   1440
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgCerrar 
      Height          =   480
      Left            =   240
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonAceptar    As clsGraphicalButton

Private cBotonCerrar     As clsGraphicalButton

Private cBotonDetalles   As clsGraphicalButton

Private cBotonRechazar   As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private TipoProp         As TIPO_PROPUESTA

Public Enum TIPO_PROPUESTA

        ALIANZA = 1
        PAZ = 2

End Enum

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    
        Call LoadBackGround
        Call LoadButtons

End Sub

Private Sub LoadButtons()

        Dim GrhPath As String
    
        GrhPath = DirInterfaces

        Set cBotonAceptar = New clsGraphicalButton
        Set cBotonCerrar = New clsGraphicalButton
        Set cBotonDetalles = New clsGraphicalButton
        Set cBotonRechazar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarOferta.jpg", _
           GrhPath & "BotonAceptarOferta.jpg", _
           GrhPath & "BotonAceptarClickOferta.jpg", Me)

        Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarOferta.jpg", _
           GrhPath & "BotonCerrarOferta.jpg", _
           GrhPath & "BotonCerrarClickOferta.jpg", Me)

        Call cBotonDetalles.Initialize(imgDetalle, GrhPath & "BotonDetallesOferta.jpg", _
           GrhPath & "BotonDetallesOferta.jpg", _
           GrhPath & "BotonDetallesClickOferta.jpg", Me)

        Call cBotonRechazar.Initialize(imgRechazar, GrhPath & "BotonRechazarOferta.jpg", _
           GrhPath & "BotonRechazarOferta.jpg", _
           GrhPath & "BotonRechazarClickOferta.jpg", Me)

End Sub

Private Sub LoadBackGround()

        If TipoProp = TIPO_PROPUESTA.ALIANZA Then
                Me.Picture = LoadPicture(DirInterfaces & "VentanaOfertaAlianza.jpg")
        Else
                Me.Picture = LoadPicture(DirInterfaces & "VentanaOfertaPaz.jpg")

        End If

End Sub

Public Property Let ProposalType(ByVal nValue As TIPO_PROPUESTA)
        TipoProp = nValue

End Property

Private Sub imgAceptar_Click()

        If TipoProp = PAZ Then
                Call WriteGuildAcceptPeace(lista.List(lista.listIndex))
        Else
                Call WriteGuildAcceptAlliance(lista.List(lista.listIndex))

        End If
    
        Me.Hide
    
        Unload Me

End Sub

Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub imgDetalle_Click()

        If TipoProp = PAZ Then
                Call WriteGuildPeaceDetails(lista.List(lista.listIndex))
        Else
                Call WriteGuildAllianceDetails(lista.List(lista.listIndex))

        End If

End Sub

Private Sub imgRechazar_Click()

        If TipoProp = PAZ Then
                Call WriteGuildRejectPeace(lista.List(lista.listIndex))
        Else
                Call WriteGuildRejectAlliance(lista.List(lista.listIndex))

        End If
    
        Me.Hide
    
        Unload Me

End Sub
