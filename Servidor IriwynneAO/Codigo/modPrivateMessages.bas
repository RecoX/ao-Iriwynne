Attribute VB_Name = "modPrivateMessages"
Option Explicit

Public Sub AgregarMensaje(ByVal UserIndex As Integer, _
                          ByRef Autor As String, _
                          ByRef Mensaje As String)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Agrega un nuevo mensaje privado a un usuario online.
        '***************************************************
        Dim Loopc As Long

        With UserList(UserIndex)

                If .UltimoMensaje < MAX_PRIVATE_MESSAGES Then
                        .UltimoMensaje = .UltimoMensaje + 1
                Else

                        For Loopc = 1 To MAX_PRIVATE_MESSAGES - 1
                                .Mensajes(Loopc) = .Mensajes(Loopc + 1)
                        Next

                End If
        
                With .Mensajes(.UltimoMensaje)
                        .Contenido = UCase$(Autor) & ": " & Mensaje & " (" & Now & ")"
                        .Nuevo = True

                End With
        
                Call WriteConsoleMsg(UserIndex, "¡Has recibido un mensaje privado de un Game Master!", FontTypeNames.FONTTYPE_GM)

        End With

End Sub

Public Sub AgregarMensajeOFF(ByRef Destinatario As String, _
                             ByRef Autor As String, _
                             ByRef Mensaje As String)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Agrega un nuevo mensaje privado a un usuario offline.
        '***************************************************
        Dim UltimoMensaje As Byte

        Dim CharFile      As String

        Dim Contenido     As String

        Dim Loopc         As Long

        CharFile = CharPath & Destinatario & ".chr"
        UltimoMensaje = CByte(GetVar(CharFile, "MENSAJES", "UltimoMensaje"))
        Contenido = UCase$(Autor) & ": " & Mensaje & " (" & Now & ")"

        If UltimoMensaje < MAX_PRIVATE_MESSAGES Then
                UltimoMensaje = UltimoMensaje + 1
        Else

                For Loopc = 1 To MAX_PRIVATE_MESSAGES - 1
                        Call WriteVar(CharFile, "MENSAJES", "MSJ" & Loopc, GetVar(CharFile, "MENSAJES", "MSJ" & Loopc + 1))
                        Call WriteVar(CharFile, "MENSAJES", "MSJ" & Loopc & "_NUEVO", GetVar(CharFile, "MENSAJES", "MSJ" & Loopc + 1 & "_NUEVO"))
                Next Loopc

        End If
        
        Call WriteVar(CharFile, "MENSAJES", "MSJ" & UltimoMensaje, Contenido)
        Call WriteVar(CharFile, "MENSAJES", "MSJ" & UltimoMensaje & "_NUEVO", 1)
        Call WriteVar(CharFile, "MENSAJES", "UltimoMensaje", UltimoMensaje)

End Sub

Public Function TieneMensajesNuevos(ByVal UserIndex As Integer) As Boolean

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Determina si el usuario tiene mensajes nuevos.
        '***************************************************
        Dim Loopc As Long

        For Loopc = 1 To MAX_PRIVATE_MESSAGES

                If UserList(UserIndex).Mensajes(Loopc).Nuevo Then
                        TieneMensajesNuevos = True
                        Exit Function

                End If

        Next Loopc
    
        TieneMensajesNuevos = False

End Function

Public Sub GuardarMensajes(ByVal UserIndex As Integer, ByRef Manager As clsIniManager)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Guarda los mensajes del usuario.
        '***************************************************
        Dim Loopc As Long
    
        With UserList(UserIndex)
        
                Call Manager.ChangeValue("MENSAJES", "UltimoMensaje", CStr(.UltimoMensaje))
        
                For Loopc = 1 To MAX_PRIVATE_MESSAGES
                        Call Manager.ChangeValue("MENSAJES", "MSJ" & Loopc, .Mensajes(Loopc).Contenido)

                        If .Mensajes(Loopc).Nuevo Then
                                Call Manager.ChangeValue("MENSAJES", "MSJ" & Loopc & "_NUEVO", 1)
                        Else
                                Call Manager.ChangeValue("MENSAJES", "MSJ" & Loopc & "_NUEVO", 0)

                        End If

                Next Loopc

        End With

End Sub

Public Sub CargarMensajes(ByVal UserIndex As Integer, ByRef Manager As clsIniManager)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Carga los mensajes del usuario.
        '***************************************************
        Dim Loopc As Long

        With UserList(UserIndex)
                .UltimoMensaje = val(Manager.GetValue("MENSAJES", "UltimoMensaje"))
        
                For Loopc = 1 To MAX_PRIVATE_MESSAGES

                        With .Mensajes(Loopc)
                                .Nuevo = val(Manager.GetValue("MENSAJES", "MSJ" & Loopc & "_NUEVO"))
                                .Contenido = CStr(Manager.GetValue("MENSAJES", "MSJ" & Loopc))

                        End With

                Next Loopc

        End With

End Sub

Private Sub LimpiarMensajeSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Limpia el un mensaje de un usuario online.
        '***************************************************
        
        With UserList(UserIndex).Mensajes(Slot)
                .Contenido = vbNullString
                .Nuevo = False

        End With

End Sub

Public Sub LimpiarMensajes(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Limpia los mensajes del slot.
        '***************************************************
        Dim Loopc As Long

        With UserList(UserIndex)
                .UltimoMensaje = 0
        
                For Loopc = 1 To MAX_PRIVATE_MESSAGES
                        Call LimpiarMensajeSlot(UserIndex, Loopc)
                Next Loopc

        End With

End Sub

Public Sub BorrarMensaje(ByVal UserIndex As Integer, ByVal Slot As Byte)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Borra un mensaje de un usuario.
        '***************************************************
        Dim Loopc As Long

        With UserList(UserIndex)

                If Slot > .UltimoMensaje Or Slot < 1 Then Exit Sub

                If Slot = .UltimoMensaje Then
                        Call LimpiarMensajeSlot(UserIndex, Slot)
                Else

                        For Loopc = Slot To MAX_PRIVATE_MESSAGES - 1
                                .Mensajes(Loopc) = .Mensajes(Loopc + 1)
                        Next Loopc

                        Call LimpiarMensajeSlot(UserIndex, .UltimoMensaje)

                End If
        
                .UltimoMensaje = .UltimoMensaje - 1

        End With

End Sub

Public Sub BorrarMensajeOFF(ByVal UserName As String, ByVal Slot As Byte)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Borra un mensaje de un usuario.
        '***************************************************
        Dim CharFile      As String

        Dim UltimoMensaje As Byte

        Dim Loopc         As Long

        CharFile = CharPath & UserName & ".chr"
    
        UltimoMensaje = GetVar(CharFile, "MENSAJES", "UltimoMensaje")
    
        If Slot > UltimoMensaje Or Slot < 1 Then Exit Sub
    
        If Slot = UltimoMensaje Then
                Call WriteVar(CharFile, "MENSAJES", "MSJ" & Slot, vbNullString)
                Call WriteVar(CharFile, "MENSAJES", "MSJ" & Slot & "_Nuevo", vbNullString)
        Else

                For Loopc = Slot To UltimoMensaje - 1
                        Call WriteVar(CharFile, "MENSAJES", "MSJ" & Loopc, GetVar(CharFile, "MENSAJES", "MSJ" & Loopc + 1))
                        Call WriteVar(CharFile, "MENSAJES", "MSJ" & Loopc & "_NUEVO", GetVar(CharFile, "MENSAJES", "MSJ" & Loopc + 1 & "_NUEVO"))
                Next Loopc

                Call WriteVar(CharFile, "MENSAJES", "MSJ" & UltimoMensaje, vbNullString)
                Call WriteVar(CharFile, "MENSAJES", "MSJ" & UltimoMensaje & "_Nuevo", vbNullString)

        End If
    
        Call WriteVar(CharFile, "MENSAJES", "UltimoMensaje", UltimoMensaje - 1)

End Sub

Public Sub LimpiarMensajesOFF(ByVal UserName As String)

        '***************************************************
        'Author: Amraphen
        'Last Modification: 18/08/2011
        'Borra los mensajes de un usuario offline.
        '***************************************************
        Dim CharFile      As String

        Dim UltimoMensaje As Byte

        Dim Loopc         As Long

        CharFile = CharPath & UserName & ".chr"
    
        UltimoMensaje = GetVar(CharFile, "MENSAJES", "UltimoMensaje")
    
        If UltimoMensaje > 0 Then

                For Loopc = 1 To UltimoMensaje
                        Call WriteVar(CharFile, "MENSAJES", "MSJ" & Loopc, vbNullString)
                        Call WriteVar(CharFile, "MENSAJES", "MSJ" & Loopc & "_NUEVO", vbNullString)
                Next Loopc
        
                Call WriteVar(CharFile, "MENSAJES", "UltimoMensaje", 0)

        End If

End Sub
