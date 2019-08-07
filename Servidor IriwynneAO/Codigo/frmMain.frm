VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   6390
   ClientLeft      =   1950
   ClientTop       =   1515
   ClientWidth     =   5250
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6390
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CheckBox chkServerHabilitado 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Server Habilitado Solo Gms"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtNumUsers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdSystray 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Systray"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrarServer 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cerrar Servidor"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   3495
   End
   Begin VB.CommandButton cmdConfiguracion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Configuraci�n General"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   4935
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3000
      Top             =   2580
   End
   Begin VB.CommandButton cmdDump 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Crear Log Cr�tico de Usuarios"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   4935
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3960
      Top             =   2580
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   3060
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3960
      Top             =   2100
   End
   Begin VB.Timer tLluviaEvent 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3480
      Top             =   3060
   End
   Begin VB.Timer tLluvia 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   2580
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3000
      Top             =   3060
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   4440
      Top             =   3060
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4455
      Top             =   2580
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mensajea todos los clientes (Solo testeo)"
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin VB.TextBox txtChat 
         BackColor       =   &H00C0FFFF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1320
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enviar por Consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enviar por Pop-Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4800
      Top             =   0
      Width           =   495
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "N�mero de usuarios jugando:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2460
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA

        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64

End Type
   
Const NIM_ADD = 0

Const NIM_DELETE = 2

Const NIF_MESSAGE = 1

Const NIF_ICON = 2

Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200

Const WM_LBUTTONDBLCLK = &H203

Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal hWnd As Long, _
                              lpdwProcessId As Long) As Long

Private Declare Function Shell_NotifyIconA _
                Lib "SHELL32" (ByVal dwMessage As Long, _
                               lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, _
                                   ID As Long, _
                                   flags As Long, _
                                   CallbackMessage As Long, _
                                   Icon As Long, _
                                   Tip As String) As NOTIFYICONDATA

        Dim nidTemp As NOTIFYICONDATA

        nidTemp.cbSize = Len(nidTemp)
        nidTemp.hWnd = hWnd
        nidTemp.uID = ID
        nidTemp.uFlags = flags
        nidTemp.uCallbackMessage = CallbackMessage
        nidTemp.hIcon = Icon
        nidTemp.szTip = Tip & Chr$(0)

        setNOTIFYICONDATA = nidTemp

End Function

Sub CheckIdleUser()

        Dim iUserIndex As Long
    
        For iUserIndex = 1 To LastUser

                With UserList(iUserIndex)

                        'Conexion activa? y es un usuario loggeado?
                        If .ConnID <> -1 And .flags.UserLogged And Not esGM(iUserIndex) Then

                                'Actualiza el contador de inactividad
                                If .flags.Traveling = 0 Then
                                        .Counters.IdleCount = .Counters.IdleCount + 1

                                End If
                
                                If .Counters.IdleCount >= IdleLimit Then
                                        Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado.")

                                        'mato los comercios seguros
                                        If .ComUsu.DestUsu > 0 Then
                                                If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                                                        If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                                                                Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                                                                Call FinComerciarUsu(.ComUsu.DestUsu)
                                                                Call FlushBuffer(.ComUsu.DestUsu) 'flush the buffer to send the message right away

                                                        End If

                                                End If

                                                Call FinComerciarUsu(iUserIndex)

                                        End If

                                        Call Cerrar_Usuario(iUserIndex)

                                End If

                        End If

                End With

        Next iUserIndex

End Sub

Private Sub Auditoria_Timer()

        On Error GoTo errhand

        Static centinelSecs As Byte

        centinelSecs = centinelSecs + 1

        If centinelSecs = 5 Then
                'Every 5 seconds, we try to call the player's attention so it will report the code.
                Call modCentinela.CallUserAttention
    
                centinelSecs = 0

        End If

        Call PasarSegundo 'sistema de desconexion de 10 segs

        Exit Sub

errhand:

        Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)

        Resume Next

End Sub

Private Sub AutoSave_Timer()

        On Error GoTo errhandleR

        ' @@ Miqueas : 29/08/15
        
        'fired every minute
        Static Minutos          As Integer

        Static MinutosLatsClean As Integer

        Static MinsPjesSave     As Integer

        Static Current          As Integer
        
        Static Current2          As Integer
        
        Current = Current + 1
        Current2 = Current2 + 1
        Minutos = Minutos + 1
        MinsPjesSave = MinsPjesSave + 1
        MinutosLatsClean = MinutosLatsClean + 1

        '�?�?�?�?�?�?�?�?�?�?�
1        Call ModAreas.AreasOptimizacion
        '�?�?�?�?�?�?�?�?�?�?�

        'Actualizamos el centinela
2        Call modCentinela.PasarMinutoCentinela

9        Select Case Minutos

                Case MinutosWs - 1
3                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))

7                Case Is >= MinutosWs
4                        Call ES.DoBackUp
5                        Call aClon.VaciarColeccion
6                        Minutos = 0
 
        End Select

        Select Case MinsPjesSave
        
                Case MinutosGuardarUsuarios - 1
11                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("CharSave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))

12                Case Is >= MinutosGuardarUsuarios
13                        Call mdParty.ActualizaExperiencias
14                        Call GuardarUsuarios
15                        MinsPjesSave = 0

16        End Select

17        Select Case MinutosLatsClean
        
                Case 14
19                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza De Mundo en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
        
20                Case Is >= 15
21                        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
22                        Call LimpiarMundo
23                        MinutosLatsClean = 0

24        End Select
        
        
28                                Dim LoopC As Long

25        With ControlMensajes ' @@ Miqueas : 07/11/15

26                If .Activado = 1 Then
27                        If Current >= .Tiempo Then


29                                Current = 0

30                                For LoopC = 1 To UBound(.Mensajes())
31                                       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("Anuncio>", .Mensajes(LoopC), FontTypeNames.FONTTYPE_VERDE, FontTypeNames.FONTTYPE_DIOS))
32                                Next LoopC

33                        End If

34                End If

35        End With

251        With ControlMensajes2 ' @@ Miqueas : 07/11/15

262                If .Activado = 1 Then
273                        If Current2 >= .Tiempo Then


295                                Current2 = 0

306                                For LoopC = 1 To UBound(.Mensajes())
317                                       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsgNew("SPONSOR>", .Mensajes(LoopC), FontTypeNames.FONTTYPE_VERDE, FontTypeNames.FONTTYPE_VIOLETA))
328                                Next LoopC

339                        End If

3410                End If

351        End With

36        Call PurgarPenas
37        Call CheckIdleUser
        
38        If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
39        If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
        If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
        If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
        If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
        If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
                If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"

        End If

        
        '<<<<<-------- Log the number of users online ------>>>
40        Call DarPoder
        
        Dim N As Integer

        N = FreeFile()
        Open App.Path & "\logs\numusers.log" For Output Shared As N
        Print #N, NumUsers
        Close #N
        '<<<<<-------- Log the number of users online ------>>>

        Exit Sub

errhandleR:
        Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.description & " - en linea: " & Erl)

        Resume Next

End Sub

Private Sub chkServerHabilitado_Click()
        ServerSoloGMs = chkServerHabilitado.Value

End Sub

Private Sub cmdCerrarServer_Click()

        If MsgBox("��Atencion!! Si cierra el servidor puede provocar la perdida de datos. " & _
           "�Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
        
                Dim f

                For Each f In Forms

                        Unload f
                Next
            End
        End If

End Sub

Private Sub cmdConfiguracion_Click()
        frmServidor.Visible = True

End Sub

Private Sub CMDDUMP_Click()

        On Error Resume Next

        Dim i As Integer

        For i = 1 To MaxUsers
                Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & _
                   ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & _
                   " UserLogged: " & UserList(i).flags.UserLogged)
        Next i
    
        Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub cmdSystray_Click()
        SetSystray

End Sub

Private Sub Command1_Click()
        Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
        ''''''''''''''''SOLO PARA EL TESTEO'''''''
        ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
        txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text

End Sub

Public Sub InitMain(ByVal f As Byte)

        If f = 1 Then
                Call SetSystray
        Else
                frmMain.Show

        End If

End Sub

Private Sub Command2_Click()
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
        ''''''''''''''''SOLO PARA EL TESTEO'''''''
        ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
        txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text

End Sub

Private Sub Command3_Click()



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

        On Error Resume Next
   
        If Not Visible Then

                Select Case X \ Screen.TwipsPerPixelX
                
                        Case WM_LBUTTONDBLCLK
                                WindowState = vbNormal
                                Visible = True

                                Dim hProcess As Long

                                GetWindowThreadProcessId hWnd, hProcess
                                AppActivate hProcess

                        Case WM_RBUTTONUP
                                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                                PopupMenu mnuPopUp

                                If hHook Then UnhookWindowsHookEx hHook: hHook = 0

                End Select

        End If
   
End Sub

Private Sub QuitarIconoSystray()

        On Error Resume Next

        'Borramos el icono del systray
        Dim i   As Integer

        Dim nid As NOTIFYICONDATA

        nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, vbNullString)

        i = Shell_NotifyIconA(NIM_DELETE, nid)

End Sub

Private Sub Form_Unload(Cancel As Integer)

        On Error Resume Next

        'Save stats!!!
        Call Statistics.DumpStatistics

        Call QuitarIconoSystray

        Call LimpiaWsApi

        Dim LoopC As Integer

        For LoopC = 1 To MaxUsers

                If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
        Next

        'Log
        Dim N As Integer

        N = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #N
        Print #N, Date & " " & time & " server cerrado."
        Close #N

        End

        Set SonidosMapas = Nothing

End Sub

Private Sub FX_Timer()

        On Error GoTo hayerror

        Call SonidosMapas.ReproducirSonidosDeMapas

        Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()

        '********************************************************
        'Author: Unknown
        'Last Modify Date: -
        '********************************************************
        Dim iUserIndex   As Long

        Dim bEnviarStats As Boolean

        Dim bEnviarAyS   As Boolean
    
        On Error GoTo hayerror
    
        '<<<<<< Procesa eventos de los usuarios >>>>>>
        For iUserIndex = 1 To LastUser

                With UserList(iUserIndex)

                        'Conexion activa?
                        If .ConnID <> -1 Then
                                '�User valido?
                
                                If .ConnIDValida And .flags.UserLogged Then
                    
                                        '[Alejo-18-5]
                                        bEnviarStats = False
                                        bEnviarAyS = False
                               
                                        If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                                        If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                                        
                                        If .Counters.IntervaloGolpe > 0 Then .Counters.IntervaloGolpe = .Counters.IntervaloGolpe - 1
                                        If .Counters.IntervaloHechizo > 0 Then .Counters.IntervaloHechizo = .Counters.IntervaloHechizo - 1
                                        
                                        If .flags.Muerto = 0 Then
                        
                                                '[Consejeros]
                                                If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)
                                        
                                                If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)
                                                If (.flags.Privilegios And PlayerType.User) Then Call EfectoPolo(iUserIndex)
                                                
                                                If Not esGM(iUserIndex) Then
                                                    If .flags.Desnudo <> 0 Then
                                                        Call EfectoFrio(iUserIndex)
                                                    End If
                                                End If
                                                
                                                
                                                If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        
                                                If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                        
                                                If .flags.AdminInvisible <> 1 Then
                                                        If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                                                        If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)

                                                End If
                        
                                                If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        
                                                If .flags.AtacablePor <> 0 Then Call EfectoEstadoAtacable(iUserIndex)
                        
                                                Call DuracionPociones(iUserIndex)
                        
                                                Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                                                If .flags.Hambre = 0 And .flags.Sed = 0 Then
                                                        If Lloviendo And .Pos.Map = 127 Or Lloviendo And .Pos.Map = 128 And 1 = 2 Then
                                                                If Not Intemperie(iUserIndex) Then
                                                                        If Not .flags.Descansar Then
                                                                                'No esta descansando
                                                                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                                                                If bEnviarStats Then
                                                                                        Call WriteUpdateHP(iUserIndex)
                                                                                        bEnviarStats = False

                                                                                End If

                                                                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                                                                If bEnviarStats Then
                                                                                        Call WriteUpdateSta(iUserIndex)
                                                                                        bEnviarStats = False

                                                                                End If

                                                                        Else
                                                                                'esta descansando
                                                                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                                                                If bEnviarStats Then
                                                                                        Call WriteUpdateHP(iUserIndex)
                                                                                        bEnviarStats = False

                                                                                End If

                                                                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                                                                If bEnviarStats Then
                                                                                        Call WriteUpdateSta(iUserIndex)
                                                                                        bEnviarStats = False

                                                                                End If

                                                                                'termina de descansar automaticamente
                                                                                If .Stats.MaxHP = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                                                                        Call WriteRestOK(iUserIndex)
                                                                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                                                                        .flags.Descansar = False

                                                                                End If
                                        
                                                                        End If

                                                                End If

                                                        Else

                                                                If Not .flags.Descansar Then
                                                                        'No esta descansando
                                    
                                                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                                                        If bEnviarStats Then
                                                                                Call WriteUpdateHP(iUserIndex)
                                                                                bEnviarStats = False

                                                                        End If

                                                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                                                        If bEnviarStats Then
                                                                                Call WriteUpdateSta(iUserIndex)
                                                                                bEnviarStats = False

                                                                        End If
                                    
                                                                Else
                                                                        'esta descansando
                                    
                                                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                                                        If bEnviarStats Then
                                                                                Call WriteUpdateHP(iUserIndex)
                                                                                bEnviarStats = False

                                                                        End If

                                                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                                                        If bEnviarStats Then
                                                                                Call WriteUpdateSta(iUserIndex)
                                                                                bEnviarStats = False

                                                                        End If

                                                                        'termina de descansar automaticamente
                                                                        If .Stats.MaxHP = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                                                                Call WriteRestOK(iUserIndex)
                                                                                Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                                                                .flags.Descansar = False

                                                                        End If
                                    
                                                                End If

                                                        End If

                                                End If
                        
                                                If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                                                If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                                        Else

                                                If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)
                                        End If 'Muerto

                                Else 'no esta logeado?
                                        'Inactive players will be removed!
                                        .Counters.IdleCount = .Counters.IdleCount + 1

                                        If .Counters.IdleCount > IntervaloParaConexion Then
                                                .Counters.IdleCount = 0
                                                Call CloseSocket(iUserIndex)

                                        End If

                                End If 'UserLogged
                
                                'If there is anything to be sent, we send it
                                Call FlushBuffer(iUserIndex)

                        End If

                End With

        Next iUserIndex

        Exit Sub

hayerror:
        LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)

End Sub

Private Sub Image1_Click()

    Dim i      As Long

    On Error GoTo errhandleR

    For i = 1 To LastUser
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida And UserList(i).flags.UserLogged Then
            Call WriteReCheckCpuID(i)
        End If
    Next i

    Exit Sub

errhandleR:
    LogError "error en ReCheckCpuID"

End Sub

Private Sub Label1_Click()

Lloviendo = Not Lloviendo

End Sub

Private Sub mnusalir_Click()
        Call cmdCerrarServer_Click

End Sub

Public Sub mnuMostrar_Click()

        On Error Resume Next

        WindowState = vbNormal
        Form_MouseMove 0, 0, 7725, 0

End Sub

Private Sub SetSystray()

        Dim i   As Integer

        Dim s   As String

        Dim nid As NOTIFYICONDATA
    
        s = "ARGENTUM-ONLINE"
        nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
        i = Shell_NotifyIconA(NIM_ADD, nid)
        
        If WindowState <> vbMinimized Then WindowState = vbMinimized
        Visible = False

End Sub

Private Sub npcataca_Timer()

        On Error Resume Next

        Dim npc As Long
    
        For npc = 1 To LastNPC
                Npclist(npc).CanAttack = 1
        Next npc

End Sub

Private Sub TIMER_AI_Timer()

    On Error GoTo ErrorHandler

    Dim NpcIndex As Long

    Dim Mapa   As Integer

    Dim e_p    As Integer

    'Barrin 29/9/03
1   If Not haciendoBK And Not EnPausa Then

        'Update NPCs
2       For NpcIndex = 1 To LastNPC

3           With Npclist(NpcIndex)

4               If .flags.NPCActive Then    'Nos aseguramos que sea INTELIGENTE!

                    ' Chequea si contiua teniendo due�o
5                   If .Owner > 0 Then Call ValidarPermanenciaNpc(NpcIndex)

6                   If .flags.Paralizado = 1 Then
7                       Call EfectoParalisisNpc(NpcIndex)
8                   Else

                        ' Preto? Tienen ai especial
9                       If .NPCtype = eNPCType.Pretoriano And Not .Name = "Ladr�n Pretoriano" Then
10                          Call ClanPretoriano(.ClanIndex).PerformPretorianAI(NpcIndex)
11                      Else

                            'Usamos AI si hay algun user en el mapa
12                          If .flags.Inmovilizado = 1 Then
13                              Call EfectoParalisisNpc(NpcIndex)

14                          End If

15                          Mapa = .Pos.Map

16                          If Mapa > 0 Then
17                              If MapInfo(Mapa).NumUsers > 0 Then
18                                    If .Movement <> TipoAI.ESTATICO Then
19                                        Call NPCAI(NpcIndex)

20                                    End If

21                                End If

22                            End If

23                        End If

24                    End If

25                End If

            End With

26        Next NpcIndex

    End If

    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & _
                  Npclist(NpcIndex).Pos.Map & " - Linea: " & Erl)
    Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub tLluvia_Timer()

        On Error GoTo errhandleR

        Dim iCount As Long

        If Lloviendo Then

                For iCount = 1 To LastUser
                        Call EfectoLluvia(iCount)
                Next iCount

        End If

        Exit Sub
errhandleR:
        Call LogError("tLluvia " & Err.Number & ": " & Err.description)

End Sub

Private Sub tLluviaEvent_Timer()

    On Error GoTo ErrorHandler
    
    Static MinutosLloviendo As Long

    Static MinutosSinLluvia As Long

        MinutosSinLluvia = MinutosSinLluvia + 1
        
        frmMain.Label1.Caption = MinutosSinLluvia & " - " & Lloviendo
        
        If MinutosSinLluvia = 20 Then
            Lloviendo = True
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        ElseIf MinutosSinLluvia = 30 Then
            Lloviendo = False
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        ElseIf MinutosSinLluvia = 40 Then
            Lloviendo = True
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        ElseIf MinutosSinLluvia > 60 Then
            MinutosSinLluvia = 0
            Lloviendo = False
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        End If
        
Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")

End Sub

Private Sub tPiqueteC_Timer()
    
        Dim NuevaA As Boolean

        ' Dim NuevoL As Boolean
        Dim GI     As Integer

        Dim i      As Long
    
        On Error GoTo errhandleR

        For i = 1 To LastUser

                With UserList(i)

                        ' @@ Miqueas - 07/11/15: Gms excluidos del piquete
                        If .flags.UserLogged Then
                        
                                If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then ' @@ Parche - 17/11/2015
                                        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE And .flags.Privilegios = User Then
                                                If .flags.Muerto = 0 Then
                                                        .Counters.PiqueteC = .Counters.PiqueteC + 1
                                                        Call WriteConsoleMsg(i, "���Est�s obstruyendo la v�a p�blica, mu�vete o ser�s encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                        
                                                        If .Counters.PiqueteC > 23 Then
                                                                .Counters.PiqueteC = 0
                                                                Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)

                                                        End If

                                                Else
                                                        .Counters.PiqueteC = 0

                                                End If

                                        Else
                                                .Counters.PiqueteC = 0

                                        End If

                                End If

                                Call FlushBuffer(i)

                        End If

                End With

        Next i

        Exit Sub

errhandleR:
        Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.description)

End Sub
