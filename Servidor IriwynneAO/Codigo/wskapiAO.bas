Attribute VB_Name = "wskapiAO"
'**************************************************************
' wskapiAO.bas
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

''
' Modulo para manejar Winsock
'

'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).
#Const WSAPI_CREAR_LABEL = True

Private Const SD_BOTH As Long = &H2

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx _
                Lib "user32" _
                Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                         ByVal lpClassName As String, _
                                         ByVal lpWindowName As String, _
                                         ByVal dwStyle As Long, _
                                         ByVal X As Long, _
                                         ByVal Y As Long, _
                                         ByVal nWidth As Long, _
                                         ByVal nHeight As Long, _
                                         ByVal hwndParent As Long, _
                                         ByVal hMenu As Long, _
                                         ByVal hInstance As Long, _
                                         lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WS_CHILD = &H40000000

Private Const GWL_WNDPROC = (-4)

Private Const SIZE_RCVBUF As Long = 8192

Private Const SIZE_SNDBUF As Long = 8192

Private Const WM_USER     As Long = &H400

Public Const WM_WINSOCK   As Long = WM_USER + 1

' ====================================================================================
' ====================================================================================

Private OldWProc          As Long

Private ActualWProc       As Long

Public hWndMsg            As Long

' ====================================================================================
' ====================================================================================

Public SockListen         As Long

' ====================================================================================
' ====================================================================================

Public Sub IniciaWsApi(ByVal hwndParent As Long)
        Call LogApiSock("IniciaWsApi")
        Debug.Print "IniciaWsApi"

        #If WSAPI_CREAR_LABEL Then
                hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
        #Else
                hWndMsg = hwndParent
        #End If 'WSAPI_CREAR_LABEL

        OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
        ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

        Dim Desc As String

        Call StartWinsock(Desc)

End Sub

Public Sub LimpiaWsApi()
        Call LogApiSock("LimpiaWsApi")

        If WSAStartedUp Then
                Call EndWinsock

        End If

        If OldWProc <> 0 Then
                SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
                OldWProc = 0

        End If

        #If WSAPI_CREAR_LABEL Then

                If hWndMsg <> 0 Then
                        DestroyWindow hWndMsg

                End If

        #End If

End Sub

Public Function WndProc(ByVal hWnd As Long, _
                        ByVal msg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

        On Error Resume Next

        Dim ret      As Long

        Dim Tmp()    As Byte

        Dim s        As Long

        Dim E        As Long

        Dim n        As Integer

        Dim UltError As Long
    
        If msg = WM_WINSOCK Then
                s = wParam
                E = WSAGetSelectEvent(lParam)
        
                If E = FD_ACCEPT Then
                        If s = SockListen Then
                                Call EventoSockAccept(s)

                        End If

                End If

        ElseIf (msg > WM_WINSOCK) And (msg <= (WM_WINSOCK + MaxUsers)) Then
                s = wParam
                E = WSAGetSelectEvent(lParam)

                n = msg - WM_WINSOCK

                Select Case E

                        Case FD_READ
                                'create appropiate sized buffer
                                ReDim Preserve Tmp(SIZE_RCVBUF - 1) As Byte
                
                                ret = recv(s, Tmp(0), SIZE_RCVBUF, 0)

                                ' Comparo por = 0 ya que esto es cuando se cierra
                                ' "gracefully". (mas abajo)
                                If ret < 0 Then
                                        UltError = Err.LastDllError

                                        If UltError = WSAEMSGSIZE Then
                                                Debug.Print "WSAEMSGSIZE"
                                                ret = SIZE_RCVBUF
                                        Else
                                                Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                                                Call LogApiSock("Error en Recv: N=" & n & " S=" & s & " Str=" & GetWSAErrorString(UltError))
                        
                                                'no hay q llamar a CloseSocket() directamente,
                                                'ya q pueden abusar de algun error para
                                                'desconectarse sin los 10segs. CREEME.
                                                Call CloseSocketSL(n)
                                                Call Cerrar_Usuario(n)
                                                Exit Function

                                        End If

                                ElseIf ret = 0 Then
                                        Call CloseSocketSL(n)
                                        Call Cerrar_Usuario(n)

                                End If
                
                                ReDim Preserve Tmp(ret - 1) As Byte
                
                                Call EventoSockRead(n, Tmp)
                
                        Case FD_WRITE
                                Call FlushBuffer(n)
            
                        Case FD_CLOSE
                                Call apiclosesocket(s)
                
                                If UserList(n).ConnID <> -1 Then 'Si se cerró bien el socket en esta instancia ConnID tendría que ser -1
                                        UserList(n).ConnID = -1
                                        UserList(n).ConnIDValida = False
                                        Call EventoSockClose(n)

                                End If

                End Select

        Else
                WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)

        End If

End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal Slot As Integer, ByRef str As String) As Long

        Dim ret     As String

        Dim Retorno As Long

        Dim data()  As Byte

        Dim length  As Long
    
        ReDim Preserve data(Len(str) - 1) As Byte

        data = StrConv(str, vbFromUnicode)
    
        #If SeguridadAlkon Then
                Call Security.DataSent(Slot, data)
        #End If
    
        length = UBound(data) + 1 'No hago con len(str) porque tengo las esperanzas de cambiar el parametro string por un array de bytes
    
        If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
                ret = send(ByVal UserList(Slot).ConnID, data(0), ByVal length, ByVal 0)

                If ret < 0 Then
                        ret = Err.LastDllError

                        If ret = WSAEWOULDBLOCK Then
                
                                #If SeguridadAlkon Then
                                        Call Security.DataStored(Slot)
                                #End If
                
                                ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
                                Call UserList(Slot).outgoingData.WriteASCIIStringFixed(str)

                        End If

                ElseIf ret < length Then

                        Dim Buff() As Byte
            
                        ReDim Buff(ret - 1) As Byte
            
                        data = StrConv(str, vbFromUnicode)
            
                        Call CopyMemory(Buff(0), data(0), ret)
            
                        #If SeguridadAlkon Then
                                Call Security.DataStored(Slot)
                                Call Security.DataSent(Slot, Buff())
                        #End If
            
                        ReDim Buff((length - ret) - 1) As Byte
            
                        Call CopyMemory(Buff(0), data(ret), length - ret)
            
                        Call UserList(Slot).outgoingData.WriteBlock(Buff())

                End If

        ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then

                If Not UserList(Slot).Counters.Saliendo Then
                        Retorno = -1

                End If

        End If
    
        WsApiEnviar = Retorno

End Function

Public Sub LogApiSock(ByVal str As String)

        On Error GoTo Errhandler

        Dim nfile As Integer

        nfile = FreeFile ' obtenemos un canal
        Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
        Print #nfile, Date & " " & time & " " & str
        Close #nfile

        Exit Sub

Errhandler:

End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
        '==========================================================
        'USO DE LA API DE WINSOCK
        '========================
    
        Dim NewIndex  As Integer

        Dim ret       As Long

        Dim Tam       As Long, sa As sockaddr

        Dim NuevoSock As Long

        Dim i         As Long

        Dim str       As String

        Dim data()    As Byte
    
        Tam = sockaddr_size
    
        '=============================================
        'SockID es en este caso es el socket de escucha,
        'a diferencia de socketwrench que es el nuevo
        'socket de la nueva conn
    
        'Modificado por Maraxus
        'ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
        ret = accept(SockID, sa, Tam)

        If ret = INVALID_SOCKET Then
                i = Err.LastDllError
                Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
                Exit Sub

        End If

        'If ret = INVALID_SOCKET Then
        '    If Err.LastDllError = 11002 Then
        '        ' We couldn't decide if to accept or reject the connection
        '        'Force reject so we can get it out of the queue
        '        ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 1)
        '        Call LogCriticEvent("Error en WSAAccept() API 11002: No se pudo decidir si aceptar o rechazar la conexión.")
        '    Else
        '        i = Err.LastDllError
        '        Call LogCriticEvent("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
        '        Exit Sub
        '    End If
        'End If

        NuevoSock = ret
    
        If setsockopt(NuevoSock, SOL_SOCKET, SO_LINGER, 0, 4) <> 0 Then
                i = Err.LastDllError
                Call LogCriticEvent("Error al setear lingers." & i & ": " & GetWSAErrorString(i))

        End If
    
        If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
                Call WSApiCloseSocket(NuevoSock, 0)
                Exit Sub

        End If
    
        If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
                str = Protocol.PrepareMessageErrorMsg("Limite de conexiones para su IP alcanzado.")
        
                ReDim Preserve data(Len(str) - 1) As Byte
        
                data = StrConv(str, vbFromUnicode)
        
                #If SeguridadAlkon Then
                        Call Security.DataSent(Security.NO_SLOT, data)
                #End If
        
                Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
                Call WSApiCloseSocket(NuevoSock, 0)
                Exit Sub

        End If
    
        'Seteamos el tamaño del buffer de entrada
        If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
                i = Err.LastDllError
                Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))

        End If

        'Seteamos el tamaño del buffer de salida
        If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
                i = Err.LastDllError
                Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))

        End If
    
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   BIENVENIDO AL SERVIDOR!!!!!!!!
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
        NewIndex = NextOpenUser ' Nuevo indice
    
        If NewIndex <= MaxUsers Then
        
                'Make sure both outgoing and incoming data buffers are clean
                Call UserList(NewIndex).incomingData.ReadASCIIStringFixed(UserList(NewIndex).incomingData.length)
                Call UserList(NewIndex).outgoingData.ReadASCIIStringFixed(UserList(NewIndex).outgoingData.length)

                #If SeguridadAlkon Then
                        Call Security.NewConnection(NewIndex)
                #End If
        
                UserList(NewIndex).Ip = GetAscIP(sa.sin_addr)

                'Busca si esta banneada la ip
                For i = 1 To BanIps.Count

                        If BanIps.Item(i) = UserList(NewIndex).Ip Then
                                'Call apiclosesocket(NuevoSock)
                                Call WriteErrorMsg(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
                                Call FlushBuffer(NewIndex)
                                Call SecurityIp.IpRestarConexion(sa.sin_addr)
                                Call WSApiCloseSocket(NuevoSock, 0)
                                Exit Sub

                        End If

                Next i
        
                If NewIndex > LastUser Then LastUser = NewIndex
        
                UserList(NewIndex).ConnID = NuevoSock
                UserList(NewIndex).ConnIDValida = True
        
                If WSAAsyncSelect(NuevoSock, hWndMsg, ByVal WM_WINSOCK + NewIndex, ByVal (FD_READ Or FD_WRITE Or FD_CLOSE)) Then
                        Call WSApiCloseSocket(NuevoSock, 0)

                End If

        Else
                str = Protocol.PrepareMessageErrorMsg("El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
        
                ReDim Preserve data(Len(str) - 1) As Byte
        
                data = StrConv(str, vbFromUnicode)
        
                #If SeguridadAlkon Then
                        Call Security.DataSent(Security.NO_SLOT, data)
                #End If
        
                Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
                Call WSApiCloseSocket(NuevoSock, 0)

        End If

End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos() As Byte)

        With UserList(Slot)
    
                #If SeguridadAlkon Then
                        Call Security.DataReceived(Slot, Datos)
                #End If
    
                Call .incomingData.WriteBlock(Datos)
    
                If .ConnID <> -1 Then
                        Call HandleIncomingData(Slot)
                Else
                        Exit Sub

                End If

        End With

End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)

        'Es el mismo user al que está revisando el centinela??
        'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
        Dim CentinelaIndex As Byte

        CentinelaIndex = UserList(Slot).flags.CentinelaIndex
        
        If CentinelaIndex <> 0 Then
                Call modCentinela.CentinelaUserLogout(CentinelaIndex)

        End If
    
        #If SeguridadAlkon Then
                Call Security.UserDisconnected(Slot)
        #End If
    
        If UserList(Slot).flags.UserLogged Then
                Call CloseSocketSL(Slot)
                Call Cerrar_Usuario(Slot)
        Else
                Call CloseSocket(Slot)

        End If

End Sub

Public Sub WSApiReiniciarSockets()

        Dim i As Long

        'Cierra el socket de escucha
        If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
        'Cierra todas las conexiones
        For i = 1 To MaxUsers

                If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
                        Call CloseSocket(i)

                End If
        
                'Call ResetUserSlot(i)
        Next i
    
        For i = 1 To MaxUsers
                Set UserList(i).incomingData = Nothing
                Set UserList(i).outgoingData = Nothing
        Next i
    
        ' No 'ta el PRESERVE :p
        ReDim UserList(1 To MaxUsers)

        For i = 1 To MaxUsers
                UserList(i).ConnID = -1
                UserList(i).ConnIDValida = False
        
                Set UserList(i).incomingData = New clsByteQueue
                Set UserList(i).outgoingData = New clsByteQueue
        Next i
    
        LastUser = 1
        NumUsers = 0
    
        Call LimpiaWsApi
        Call Sleep(100)
        Call IniciaWsApi(frmMain.hWnd)
        SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)

End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long, ByVal UserIndex As Long)
        Call WSAAsyncSelect(Socket, hWndMsg, ByVal WM_WINSOCK + UserIndex, ByVal FD_CLOSE)
        Call ShutDown(Socket, SD_BOTH)

End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, _
                                ByRef lpCallerData As WSABUF, _
                                ByRef lpSQOS As FLOWSPEC, _
                                ByVal Reserved As Long, _
                                ByRef lpCalleeId As WSABUF, _
                                ByRef lpCalleeData As WSABUF, _
                                ByRef Group As Long, _
                                ByVal dwCallbackData As Long) As Long

        Dim sa As sockaddr
    
        'Check if we were requested to force reject

        If dwCallbackData = 1 Then
                CondicionSocket = CF_REJECT
                Exit Function

        End If
    
        'Get the address

        CopyMemory sa, ByVal lpCallerId.lpBuffer, lpCallerId.dwBufferLen
    
        If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
                CondicionSocket = CF_REJECT
                Exit Function

        End If

        CondicionSocket = CF_ACCEPT 'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero así es más claro....

End Function
