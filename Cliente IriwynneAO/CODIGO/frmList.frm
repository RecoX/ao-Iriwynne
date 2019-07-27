VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   0  'None
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Segoe Print"
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
   Picture         =   "frmList.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLadron 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   22
      Top             =   4850
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkClerigo 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   21
      Top             =   2600
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkBardo 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   20
      Top             =   2910
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkPaladin 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   19
      Top             =   3150
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkAsesino 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   18
      Top             =   3440
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkCazador 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   17
      Top             =   3710
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkGuerrero 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   16
      Top             =   3960
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkDruida 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   15
      Top             =   4270
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkBandido 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   14
      Top             =   4560
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkMago 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   13
      Top             =   2330
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   4800
      Width           =   1695
   End
   Begin VB.ComboBox cMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmList.frx":DFBC
      Left            =   1440
      List            =   "frmList.frx":E070
      TabIndex        =   10
      Text            =   "2"
      Top             =   1440
      Width           =   855
   End
   Begin VB.ComboBox cMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmList.frx":E124
      Left            =   3480
      List            =   "frmList.frx":E1D9
      TabIndex        =   9
      Text            =   "47"
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   480
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt_value 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Text            =   "ACA PONE NUMERO DOM"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmd_start 
      Caption         =   "Iniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cerrar lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmd_sumall 
      Caption         =   "Sumonear todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmd_sum 
      Caption         =   "Sumonear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.ListBox list_participant 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox cCupos1vs1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmList.frx":E28D
      Left            =   3120
      List            =   "frmList.frx":E2A2
      TabIndex        =   24
      Text            =   "2"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cCupos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmList.frx":E2B7
      Left            =   3120
      List            =   "frmList.frx":E36B
      TabIndex        =   11
      Text            =   "2"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "        SELECCIONA LAS CLASES VALIDAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   600
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label lbl_participant 
      Caption         =   "No hay una lista actualmente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
 

Private Sub cmd_start_Click()
 
    If MsgBox("Si le das a 'SÍ', se borrará la otra lista", vbYesNo) = vbYes Then
   
        List.AmountList = Val(txt_value.Text)
        
        If List.AmountList = 0 Then MsgBox "Debes elegir un numero de usuarios para la lista.": Exit Sub
        
        list_participant.Clear
        
        For i = 1 To List.Indexs
            List.Names(i) = ""
        Next i
        
        List.Indexs = 0
        
        If Val(frmList.txt_value) > 0 Then
        Call WriteStartList(List.AmountList) '@@ enviamos la cantidad de participantes dentro de la lista
        End If
    End If
 
End Sub
 
Private Sub cmd_sumall_Click()
 
    Dim loopc As Long
   
    If List.AmountList = 0 Then Exit Sub
  
    For loopc = 0 To List.AmountList
    If Len(list_participant.List(loopc)) > 0 Then
        Call WriteSummonChar(list_participant.List(loopc))
        End If
    Next loopc
 
End Sub
 
Private Sub cmd_cancel_Click()
 
    Call WriteCancelList
   
End Sub
 
Private Sub cmd_sum_Click()
 
    If List.AmountList = 0 Then Exit Sub
    If Not list_participant.listIndex <> -1 Then Exit Sub
    
    Call WriteSummonChar(frmList.list_participant.List(list_participant.listIndex))
   
End Sub
 
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
   
    '@@ para cuando abra el form, por si le cierra o algo.
   
    Dim loopX As Long
   
    With List
   
        If .Indexs <> 0 Then
       
            For loopX = 1 To .Indexs
           
                list_participant.AddItem .Names(loopX)
                lbl_participant.Caption = .Indexs & " / " & .AmountList
               
            Next loopX
   
        End If
       
    End With
    
    cCupos1vs1.listIndex = 0
    cCupos.listIndex = 0
    cMax.listIndex = 45
    
    cMin.listIndex = 0
    
End Sub


Private Sub Image1_Click()


    Dim TipoEvento As String
    If Option1(0).Value = True Then
        TipoEvento = "1vs1"
    Else
        TipoEvento = "Deathmatch"
    End If


    If Val(cMin.List(cMin.listIndex)) > Val(cMax.List(cMin.listIndex)) Then
        MsgBox "No podes elegir un número minimo mayor al numero máximo.", vbOKOnly
    End If


    ' @@ CLASES
    Dim cMago  As Byte, _
        cClerigo As Byte, _
        cBardo As Byte, _
        cPaladin As Byte, _
        cAsesino As Byte, _
        cCazador As Byte, _
        cGuerrero As Byte, _
        cDruida As Byte, _
        cLadron As Byte, _
        cBandido As Byte

    cMago = Val(chkMago.Value)
    cClerigo = Val(chkClerigo.Value)
    cBardo = Val(chkBardo.Value)
    cPaladin = Val(chkPaladin.Value)
    cAsesino = Val(chkAsesino.Value)
    cCazador = Val(chkCazador.Value)
    cGuerrero = Val(chkGuerrero.Value)
    cDruida = Val(chkDruida.Value)
    cLadron = Val(chkLadron.Value)
    cBandido = Val(chkBandido.Value)

    If cMago = 0 Then
        cMago = eClass.Mage
    Else
        cMago = 0
    End If
    If cClerigo = 0 Then
        cClerigo = eClass.Cleric
    Else
        cClerigo = 0
    End If
    If cBardo = 0 Then
        cBardo = eClass.Bard
    Else
        cBardo = 0
    End If
    If cPaladin = 0 Then
        cPaladin = eClass.Paladin
    Else
        cPaladin = 0
    End If
    If cAsesino = 0 Then
        cAsesino = eClass.Assasin
    Else
        cAsesino = 0
    End If
    If cCazador = 0 Then
        cCazador = eClass.Hunter
    Else
        cCazador = 0
    End If
    If cGuerrero = 0 Then
        cGuerrero = eClass.Warrior
    Else
        cGuerrero = 0
    End If
    If cLadron = 0 Then
        cLadron = eClass.Thief
    Else
        cLadron = 0
    End If
    If cBandido = 0 Then
        cBandido = eClass.Bandit
    Else
        cBandido = 0
    End If
    If cDruida = 0 Then
        cDruida = eClass.Druid
    Else
        cDruida = 0
    End If

    If cCupos.listIndex = -1 Then
        MsgBox "Seleccione los cupos."
        Exit Sub
    End If

    If cMin.listIndex = -1 Then
        MsgBox "Seleccione el nivel MINIMO"
        Exit Sub
    End If

    If cMax.listIndex = -1 Then
        MsgBox "Seleccione el nivel MAXIMO"
        Exit Sub
    End If

    If Val(cMin.List(cMin.listIndex)) > Val(cMax.List(cMax.listIndex)) Then
        MsgBox "El nivel minimo tiene que ser menor al nivel maximo permitido"
        Exit Sub
    End If

    If Val(cMin.List(cMin.listIndex) > 47) Or Val(cMin.List(cMin.listIndex) < 1) Then Exit Sub

    If Val(cMax.Text) > 47 Or Val(cMax.Text) < 1 Then Exit Sub
    If Val(cMin.Text) > 47 Or Val(cMin.Text) < 1 Then Exit Sub

    If TipoEvento = "1vs1" Then
        If cCupos1vs1.listIndex = -1 Then
            MsgBox "Elegi los cupos de 1vs1."
            '1-2
            '2-4
            '3-8
            '4-16
            Exit Sub
        End If

        Dim cupos As Byte
        Select Case cCupos1vs1.listIndex
        
        Case 0
        cupos = 1
        Case 1
        cupos = 2
        Case 2
        cupos = 3
        Case 3
        cupos = 4
        Case 4
        cupos = 5
        Case Else
           cupos = 6
        
        End Select
        Call WriteHacerT(cupos, Val(cMin.List(cMin.listIndex)), Val(cMax.List(cMax.listIndex)), cMago, cClerigo, cBardo, cPaladin, cAsesino, cCazador, cGuerrero, cDruida, cLadron, cBandido)
    Else
        WriteActivarDeath cCupos.List(cCupos.listIndex), Val(cMin.List(cMin.listIndex)), Val(cMax.List(cMax.listIndex)), cMago, cClerigo, cBardo, cPaladin, cAsesino, cCazador, cGuerrero, cDruida, cLadron, cBandido
    End If

End Sub


Private Sub Option1_Click(Index As Integer)

If Index = 0 Then
    Option1(1).Value = 0
    cCupos.Visible = False
    cCupos1vs1.Visible = True
ElseIf Index = 1 Then
    Option1(0).Value = 0
    cCupos.Visible = True
    cCupos1vs1.Visible = False
End If

End Sub

Private Sub txt_value_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If

End Sub
