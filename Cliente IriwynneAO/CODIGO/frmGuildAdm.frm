VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4065
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
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBuscar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   495
      TabIndex        =   1
      Top             =   4650
      Width           =   3105
   End
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3540
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   495
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   0
      Top             =   570
      Width           =   3075
   End
   Begin VB.Image imgDetalles 
      Height          =   375
      Left            =   2280
      Tag             =   "1"
      Top             =   4995
      Width           =   1335
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   480
      Tag             =   "1"
      Top             =   4995
      Width           =   855
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonCerrar     As clsGraphicalButton

Private cBotonDetalles   As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
        
        Me.Picture = LoadPicture(DirInterfaces & "VentanaListaClanes.jpg")
    
        Call LoadButtons
    
End Sub

Private Sub LoadButtons()

        Dim grhpath As String
    
        grhpath = DirInterfaces & "buttons/"

        Set cBotonCerrar = New clsGraphicalButton
        Set cBotonDetalles = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonCerrar.Initialize(ImgCerrar, grhpath & "boton Cerrar.bmp", _
           grhpath & "boton Cerrar 1.bmp", _
           grhpath & "boton Cerrar 1.bmp", Me)

        Call cBotonDetalles.Initialize(imgDetalles, grhpath & "boton Detalles.bmp", _
           grhpath & "boton Detalles 1.bmp", _
           grhpath & "boton Detalles 1.bmp", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub guildslist_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub ImgCerrar_Click()
        Unload Me
        frmMain.SetFocus

End Sub

Private Sub imgDetalles_Click()
        frmGuildBrief.EsLeader = False
        Dim guild As String
        guild = guildslist.List(guildslist.listIndex)
        
        ' If it's too long to fit, split it
        Do While InStr(1, guild, "[", vbTextCompare) > 0
                'tmp = InStrRev(str, " ")
        
                'If tmp = 0 Then Exit Do
        
                guild = Left$(guild, Len(guild) - 1)
        Loop
        
        guild = Left$(guild, Len(guild) - 1)
        Call WriteGuildRequestDetails(guild)

End Sub

Private Sub txtBuscar_Change()
        Call FiltrarListaClanes(txtBuscar.Text)

End Sub

Private Sub txtBuscar_GotFocus()

        With txtBuscar
                .SelStart = 0
                .SelLength = Len(.Text)

        End With

End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

        Dim lIndex As Long
    
        If UBound(GuildNames) <> 0 Then

                With guildslist
                        'Limpio la lista
                        .Clear
            
                        .Visible = False
            
                        ' Recorro los arrays
                        For lIndex = 0 To UBound(GuildNames)

                                ' Si coincide con los patrones
                                If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                                        ' Lo agrego a la lista
                                        .AddItem GuildNames(lIndex)

                                End If

                        Next lIndex
            
                        .Visible = True

                End With

        End If

End Sub
