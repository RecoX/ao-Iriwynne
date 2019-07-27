Attribute VB_Name = "modNuevoTimer"
Option Explicit

Public Const Tolerancia_FailIntervalo As Byte = 7

Public Enum eIntervalos

        ' @@ Acciones
        iPuedeAtacar = 1
        iPuedeAtacarConFlechas = 2
        iPuedeAtacarConHechizos = 3
        iPuedeUsarItem = 4
        iPuedeUsarItemDblClick = 5
        iPuedeUsarPocion = 6
        
        ' @@ Combos
        iComboMagiaGolpe = 7
        iComboGolpeMagia = 8
        iComboGolpeUsar = 9
    
End Enum

Private Const MaxIntervalos           As Byte = 9

Public Intervalos(1 To MaxIntervalos) As Integer

Public Sub CargarIntervalos()

        Dim s_File As String

        s_File = App.Path & "\Configuracion.ini"
    
        If FileExist(s_File, vbArchive) Then
        
                ' @@ Acciones
                Intervalos(eIntervalos.iPuedeAtacar) = val(GetVar(s_File, "INTERVALOS", "PuedeAtacar"))
                Intervalos(eIntervalos.iPuedeAtacarConFlechas) = val(GetVar(s_File, "INTERVALOS", "PuedeAtacarConFlechas"))
                Intervalos(eIntervalos.iPuedeAtacarConHechizos) = val(GetVar(s_File, "INTERVALOS", "PuedeAtacarConHechizos"))
                Intervalos(eIntervalos.iPuedeUsarItem) = val(GetVar(s_File, "INTERVALOS", "PuedeUsarItem"))
                Intervalos(eIntervalos.iPuedeUsarItemDblClick) = val(GetVar(s_File, "INTERVALOS", "PuedeUsarItemDblClick"))
                Intervalos(eIntervalos.iPuedeUsarPocion) = val(GetVar(s_File, "INTERVALOS", "PuedeUsarPocion"))
                
                ' @@ Combos
                Intervalos(eIntervalos.iComboMagiaGolpe) = val(GetVar(s_File, "INTERVALOS", "ComboMagiaGolpe"))
                Intervalos(eIntervalos.iComboGolpeMagia) = val(GetVar(s_File, "INTERVALOS", "ComboGolpeMagia"))
                Intervalos(eIntervalos.iComboGolpeUsar) = val(GetVar(s_File, "INTERVALOS", "ComboGolpeUsar"))

        End If
        
        ' @@ Acciones
        If Intervalos(eIntervalos.iPuedeAtacar) < 0 Then Intervalos(eIntervalos.iPuedeAtacar) = 1500
        If Intervalos(eIntervalos.iPuedeAtacarConFlechas) < 0 Then Intervalos(eIntervalos.iPuedeAtacarConFlechas) = 1400
        If Intervalos(eIntervalos.iPuedeAtacarConHechizos) < 0 Then Intervalos(eIntervalos.iPuedeAtacarConHechizos) = 1400
        If Intervalos(eIntervalos.iPuedeUsarItem) < 0 Then Intervalos(eIntervalos.iPuedeUsarItem) = 450
        If Intervalos(eIntervalos.iPuedeUsarItemDblClick) < 0 Then Intervalos(eIntervalos.iPuedeUsarItemDblClick) = 125
        If Intervalos(eIntervalos.iPuedeUsarPocion) < 0 Then Intervalos(eIntervalos.iPuedeUsarPocion) = 300
        
        ' @@ Combos
        If Intervalos(eIntervalos.iComboMagiaGolpe) < 0 Then Intervalos(eIntervalos.iComboMagiaGolpe) = 1000
        If Intervalos(eIntervalos.iComboGolpeMagia) < 0 Then Intervalos(eIntervalos.iComboGolpeMagia) = 1000
        If Intervalos(eIntervalos.iComboGolpeUsar) < 0 Then Intervalos(eIntervalos.iComboGolpeUsar) = 1000

End Sub

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, _
                                            Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim TActual As Long

        TActual = GetTickCount() And &H7FFFFFFF

        With UserList(UserIndex)

                If getInterval(TActual, .Counters.TimerLanzarSpell) >= Intervalos(eIntervalos.iPuedeAtacarConHechizos) Then
                        If Actualizar Then
                                .Counters.TimerLanzarSpell = TActual

                        End If

                        IntervaloPermiteLanzarSpell = True
                Else
                        IntervaloPermiteLanzarSpell = False

                End If

        End With

End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, _
                                       Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim TActual As Long

        TActual = GetTickCount() And &H7FFFFFFF

        With UserList(UserIndex)

                If getInterval(TActual, .Counters.TimerPuedeAtacar) >= Intervalos(eIntervalos.iPuedeAtacar) Then
                        If Actualizar Then
                                .Counters.TimerPuedeAtacar = TActual
                                .Counters.TimerGolpeUsar = TActual

                        End If

                        IntervaloPermiteAtacar = True
                Else
                        IntervaloPermiteAtacar = False

                End If

        End With

End Function

Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: ZaMa
        'Checks if the time that passed from the last hit is enough for the user to use a potion.
        'Last Modification: 06/04/2009
        '***************************************************

        Dim TActual As Long

        TActual = GetTickCount() And &H7FFFFFFF

        With UserList(UserIndex)

                If getInterval(TActual, .Counters.TimerGolpeUsar) >= Intervalos(eIntervalos.iComboGolpeUsar) Then
                        If Actualizar Then
                                .Counters.TimerGolpeUsar = TActual

                        End If

                        IntervaloPermiteGolpeUsar = True
                Else
                        IntervaloPermiteGolpeUsar = False

                End If

        End With

End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = True) As Boolean

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        Dim TActual As Long
    
        With UserList(UserIndex)

                If .Counters.TimerMagiaGolpe > .Counters.TimerLanzarSpell Then
                        Exit Function

                End If

                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Function
        
                TActual = GetTickCount() And &H7FFFFFFF
       
                If getInterval(TActual, .Counters.TimerLanzarSpell) >= Intervalos(eIntervalos.iComboMagiaGolpe) Then
                        If Actualizar Then
                                .Counters.TimerMagiaGolpe = TActual
                                .Counters.TimerPuedeAtacar = TActual
                                .Counters.TimerGolpeUsar = TActual

                        End If

                        IntervaloPermiteMagiaGolpe = True
                Else
                        IntervaloPermiteMagiaGolpe = False

                End If

        End With

End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim TActual As Long
    
        If UserList(UserIndex).Counters.TimerGolpeMagia > UserList(UserIndex).Counters.TimerPuedeAtacar Then
                Exit Function

        End If
    
        TActual = GetTickCount() And &H7FFFFFFF
    
        If getInterval(TActual, UserList(UserIndex).Counters.TimerPuedeAtacar) >= Intervalos(eIntervalos.iComboGolpeMagia) Then
                If Actualizar Then
                        UserList(UserIndex).Counters.TimerGolpeMagia = TActual
                        UserList(UserIndex).Counters.TimerLanzarSpell = TActual

                End If

                IntervaloPermiteGolpeMagia = True
        Else
                IntervaloPermiteGolpeMagia = False

        End If

End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, _
                                         Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim TActual As Long
    
        TActual = GetTickCount() And &H7FFFFFFF

        With UserList(UserIndex)

                If getInterval(TActual, .Counters.TimerPuedeTrabajar) >= IntervaloUserPuedeTrabajar Then
                
                        If Actualizar Then
                                .Counters.TimerPuedeTrabajar = TActual

                        End If

                        IntervaloPermiteTrabajar = True
                Else
                        IntervaloPermiteTrabajar = False

                End If

        End With

End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, _
                                     Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 25/01/2010 (ZaMa)
        '25/01/2010: ZaMa - General adjustments.
        '***************************************************

        Dim TActual As Long
    
        TActual = GetTickCount() And &H7FFFFFFF

        With UserList(UserIndex)

                If getInterval(TActual, .Counters.TimerUsar) >= 150 Then
                        If Actualizar Then
                                .Counters.TimerUsar = TActual
                                '.Counters.TimerUsarClick = TActual
                                .Counters.failedUsageAttempts = 0

                        End If

                        IntervaloPermiteUsar = True
                Else
                        IntervaloPermiteUsar = False
        
                        .Counters.failedUsageAttempts = .Counters.failedUsageAttempts + 1
        
                        'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
                        If .Counters.failedUsageAttempts = Tolerancia_FailIntervalo Then
                              '  Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > posible modificación de intervalos por parte de " & .Name & " Hora: " & time$, FontTypeNames.FONTTYPE_EJECUCION))
                                .Counters.failedUsageAttempts = 0
                                'Call CloseSocket(UserIndex)

                        End If

                End If

        End With

End Function

Public Function IntervaloPermiteUsarClick(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean

        Dim TActual As Long

        TActual = GetTickCount() And &H7FFFFFFF

        With UserList(UserIndex)

                If getInterval(TActual, .Counters.TimerUsarClick) >= Intervalos(eIntervalos.iPuedeUsarItemDblClick) Then
                
                        If Actualizar Then
                                .Counters.TimerUsar = TActual
                                .Counters.TimerUsarClick = TActual
                                .Counters.failedUsageAttempts = 0

                        End If

                        IntervaloPermiteUsarClick = True
                Else
                        IntervaloPermiteUsarClick = False
    
                        .Counters.failedUsageAttempts = .Counters.failedUsageAttempts + 1
    
                        'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
                        If .Counters.failedUsageAttempts = Tolerancia_FailIntervalo Then
                               ' Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("ANTICHEAT > posible modificación de intervalos por parte de " & .Name & " Hora: " & time$, FontTypeNames.FONTTYPE_EJECUCION))
                                .Counters.failedUsageAttempts = 0
                                'Call CloseSocket(UserIndex)

                        End If

                End If

        End With

End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim TActual As Long
    
        TActual = GetTickCount() And &H7FFFFFFF

        With UserList(UserIndex)

                If getInterval(TActual, .Counters.TimerPuedeUsarArco) >= Intervalos(eIntervalos.iPuedeAtacarConFlechas) Then
                        If Actualizar Then
                                .Counters.TimerPuedeUsarArco = TActual

                        End If

                        IntervaloPermiteUsarArcos = True
                Else
                        IntervaloPermiteUsarArcos = False

                End If

        End With

End Function

Public Function IntervaloPermiteSerAtacado(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = False) As Boolean

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 13/11/2009
        '13/11/2009: ZaMa - Add the Timer which determines wether the user can be atacked by a NPc or not
        '**************************************************************
        Dim TActual As Long
    
        TActual = GetTickCount() And &H7FFFFFFF
    
        With UserList(UserIndex)

                ' Inicializa el timer
                If Actualizar Then
                        .Counters.TimerPuedeSerAtacado = TActual
                        .flags.NoPuedeSerAtacado = True
                        IntervaloPermiteSerAtacado = False
                Else

                        If getInterval(TActual, .Counters.TimerPuedeSerAtacado) >= IntervaloPuedeSerAtacado Then
                                .flags.NoPuedeSerAtacado = False
                                IntervaloPermiteSerAtacado = True
                        Else
                                IntervaloPermiteSerAtacado = False

                        End If

                End If

        End With

End Function

Public Function IntervaloPerdioNpc(ByVal UserIndex As Integer, _
                                   Optional ByVal Actualizar As Boolean = False) As Boolean

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 13/11/2009
        '13/11/2009: ZaMa - Add the Timer which determines wether the user still owns a Npc or not
        '**************************************************************
        Dim TActual As Long
    
        TActual = GetTickCount() And &H7FFFFFFF
    
        With UserList(UserIndex)

                ' Inicializa el timer
                If Actualizar Then
                        .Counters.TimerPerteneceNpc = TActual
                        IntervaloPerdioNpc = False
                Else

                        If getInterval(TActual, .Counters.TimerPerteneceNpc) >= IntervaloOwnedNpc Then
                                IntervaloPerdioNpc = True
                        Else
                                IntervaloPerdioNpc = False

                        End If

                End If

        End With

End Function

Public Function IntervaloEstadoAtacable(ByVal UserIndex As Integer, _
                                        Optional ByVal Actualizar As Boolean = False) As Boolean

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 13/01/2010
        '13/01/2010: ZaMa - Add the Timer which determines wether the user can be atacked by an user or not
        '**************************************************************
        Dim TActual As Long
    
        TActual = GetTickCount() And &H7FFFFFFF
    
        With UserList(UserIndex)

                ' Inicializa el timer
                If Actualizar Then
                        .Counters.TimerEstadoAtacable = TActual
                        IntervaloEstadoAtacable = True
                Else

                        If getInterval(TActual, .Counters.TimerEstadoAtacable) >= IntervaloAtacable Then
                                IntervaloEstadoAtacable = False
                        Else
                                IntervaloEstadoAtacable = True

                        End If

                End If

        End With

End Function

Public Function IntervaloGoHome(ByVal UserIndex As Integer, _
                                Optional ByVal TimeInterval As Long, _
                                Optional ByVal Actualizar As Boolean = False) As Boolean

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        '01/06/2010: ZaMa - Add the Timer which determines wether the user can be teleported to its home or not
        '**************************************************************
        Dim TActual As Long
    
        TActual = GetTickCount() And &H7FFFFFFF
    
        With UserList(UserIndex)

                ' Inicializa el timer
                If Actualizar Then
                        .flags.Traveling = 1
                        .Counters.goHome = TActual + TimeInterval
                Else

                        If TActual >= .Counters.goHome Then
                                IntervaloGoHome = True

                        End If

                End If

        End With

End Function

Public Function checkInterval(ByRef StartTime As Long, _
                              ByVal TimeNow As Long, _
                              ByVal Interval As Long) As Boolean

        Dim lInterval As Long

        If TimeNow < StartTime Then
                lInterval = &H7FFFFFFF - StartTime + TimeNow + 1
        Else
                lInterval = TimeNow - StartTime

        End If

        If lInterval >= Interval Then
                StartTime = TimeNow
                checkInterval = True
        Else
                checkInterval = False

        End If

End Function

Public Function getInterval(ByVal TimeNow As Long, _
                            ByVal StartTime As Long) As Long ' 0.13.5

        If TimeNow < StartTime Then
                getInterval = &H7FFFFFFF - StartTime + TimeNow + 1
        Else
                getInterval = TimeNow - StartTime

        End If

End Function

