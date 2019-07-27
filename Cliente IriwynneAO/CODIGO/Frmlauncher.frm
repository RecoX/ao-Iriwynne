VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form Frmlauncher 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frmlauncher.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   8160
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2985
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   3720
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   5265
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Frmlauncher.frx":1381F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   1215
      Left            =   3720
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   3600
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   7800
      Picture         =   "Frmlauncher.frx":1389D
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   3600
      Top             =   3000
      Width           =   2655
   End
End
Attribute VB_Name = "Frmlauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub Command1_Click()
Dim tx As String
tx = GetVar(App.path & "\INIT\Update.ini", "INIT", "X")
Label1.Caption = tx
End Sub

Private Sub Command2_Click()
Dim ix As String
ix = Inet1.OpenURL("http://iriwynneupdater.ucoz.es/VEREXE.txt")
Label1.Caption = ix
End Sub

Private Sub Command3_Click()
Analizar
End Sub

Private Sub Image1_Click()

If EnProceso Then Exit Sub
Call addConsole("Conectando...", 255, 255, 0, True, True) '>> Informacion
'EnProceso = True
'Analizar 'Iniciamos la función Analizar =).

Call Main

End Sub

Private Function LeerInt(ByVal Ruta As String) As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    f = FreeFile
    Open Ruta For Output As f
    Print #f, data
    Close #f
End Sub
Private Sub Image2_Click()
End
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

Function Analizar()


    Dim ix     As Integer
    Dim tx     As Integer
    Dim DifX   As Integer
    Dim strsX  As String
    
    Call addConsole("Analizando cliente..", 255, 255, 255, True, False)
    tx = LeerInt(App.path & "\INIT\Update.ini")    'tx = GetVar(App.path & "\INIT\Update.ini", "INIT", "X")
    Call addConsole("Listo.", 1, 255, 1, True, False)
    On Error GoTo noanda
    
    Call addConsole("Accediendo al AutoUpdater.", 255, 255, 255, True, False)
    'LINK1            'Variable que contiene el numero de actualización correcto del servidor
    ix = Inet1.OpenURL("http://iriwynneupdater.ucoz.es/VEREXE.txt")
    Call addConsole("Listo.", 1, 255, 1, True, False)
    
    If ix < tx Then
        Call GuardarInt(App.path & "/INIT/Update.ini", "0")
        Call addConsole("Se resetearon las actualizaciones, se repetirá el proceso!", 255, 1, 1, True, False)
        ix = 0
        tx = 0
        Analizar
        Exit Function
    End If
    
    ' @@ Llego a conseguir iX.
    GoTo 99
    
noanda:
    MiError = Err.number & " " & Err.Description & " - " & Err.Source
    Call addConsole("El autoupdater tiró error : " & MiError, 255, 255, 255, True, False)
    Err.Clear
    ix = tx
GoTo 99

99    Debug.Print "No hubo error."
  On Error GoTo avisar

    Label1.Caption = ix
    'Variable que contiene el numero de actualización del cliente
    'Variable con la diferencia de actualizaciones servidor-cliente
    DifX = ix - tx

    If Not (DifX = 0) Then    'Si la diferencia no es nula,
        Call ShellExecute(Me.hWnd, "open", App.path & "/Autoupdate.exe", "", "", 1)
        End
    Else
        Call addConsole("No hay actualizaciones pendientes", 10, 255, 10, True, True)    '>> Informacion
    End If

    EnProceso = False

    Call addConsole("El cliente ya está listo para jugar", 255, 255, 255, True, True)  '>> Informacion


    
    Call Main
    Unload Me    'Me.Visible = False

    Exit Function

avisar:
    MiError = MiError & "--" & "Hay un error al actualizar, intente nuevamente."
    MsgBox "Hay un error al actualizar, intente nuevamente."
    EnProceso = False

End Function
       
        Public Sub AutoDownload(Numero As Integer)
            On Error Resume Next
           
            sRGY.Picture = SR.Picture
           
            Inet1.AccessType = icUseDefault
            Dim B() As Byte
           
           
            B() = Inet1.OpenURL(strURL, icByteArray)
           
            'Descargamos y guardamos el archivo
            Open Darchivo For Binary Access _
            Write As #1
            Put #1, , B()
            Close #1
           
            'Informacion
            Call addConsole("   Instalando actualización.", 0, 100, 255, False, False)    '>> Informacion
           
            sRGY.Picture = sY.Picture
           
            'Unzipeamos
            UnZip Darchivo, App.path & "\"
           
            'Borramos el zip
            Kill Darchivo
        End Sub


Private Sub Image3_Click()
    ' @ Cui
    Call ShellExecute(0, "Open", "http://www.argentumonline.org/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub Image4_Click()
    Call ShellExecute(0, "Open", "http://www.argentumonline.org/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub Label2_Click()

End Sub

