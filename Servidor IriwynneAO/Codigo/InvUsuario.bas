Attribute VB_Name = "InvUsuario"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        ' 22/05/2010: Los items newbies ya no son robables.
        '***************************************************

        '17/09/02
        'Agregue que la función se asegure que el objeto no es un barco

        On Error GoTo errhandleR

        Dim i        As Integer

        Dim objIndex As Integer
    
        For i = 1 To UserList(UserIndex).CurrentInventorySlots
                objIndex = UserList(UserIndex).Invent.Object(i).objIndex

                If objIndex > 0 Then
                        If (ObjData(objIndex).OBJType <> eOBJType.otLlaves And _
                        ObjData(objIndex).OBJType <> eOBJType.otMonturas And _
                           ObjData(objIndex).OBJType <> eOBJType.otBarcos And _
                           Not ItemNewbie(objIndex)) Then
                                TieneObjetosRobables = True
                                Exit Function

                        End If

                End If

        Next i
    
        Exit Function

errhandleR:
        Call LogError("Error en TieneObjetosRobables. Error: " & Err.Number & " - " & Err.description)

End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, _
                            ByVal objIndex As Integer, _
                            Optional ByRef sMotivo As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 14/01/2010 (ZaMa)
        '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
        '***************************************************

        On Error GoTo manejador
    
        'Admins can use ANYTHING!
        If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
                If ObjData(objIndex).ClaseProhibida(1) <> 0 Then

                        Dim i As Integer

                        For i = 1 To NUMCLASES

                                If ObjData(objIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
                                        ClasePuedeUsarItem = False
                                        sMotivo = "Tu clase no puede usar este objeto."
                                        Exit Function

                                End If

                        Next i

                End If

        End If
    
        ClasePuedeUsarItem = True

        Exit Function

manejador:
        LogError ("Error en ClasePuedeUsarItem")

End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim j As Integer

        With UserList(UserIndex)

                For j = 1 To UserList(UserIndex).CurrentInventorySlots

                        If .Invent.Object(j).objIndex > 0 Then
             
                                If ObjData(.Invent.Object(j).objIndex).Newbie = 1 Then _
                                   Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                                Call UpdateUserInv(False, UserIndex, j)
        
                        End If

                Next j
    
                '[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
                'es transportado a su hogar de origen ;)
                If MapInfo(.pos.Map).Restringir = eRestrict.restrict_newbie Then
        
                        Dim DeDonde As WorldPos
        
                        Select Case .Hogar
    
                                Case eCiudad.cUllathorpe
                                        DeDonde = Ullathorpe

                        End Select
        
                        Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
    
                End If

                '[/Barrin]
        End With

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim j As Integer

        With UserList(UserIndex)

                For j = 1 To .CurrentInventorySlots
                        .Invent.Object(j).objIndex = 0
                        .Invent.Object(j).Amount = 0
                        .Invent.Object(j).Equipped = 0
                Next j
    
                .Invent.NroItems = 0
    
                .Invent.ArmourEqpObjIndex = 0
                .Invent.ArmourEqpSlot = 0
    
                .Invent.WeaponEqpObjIndex = 0
                .Invent.WeaponEqpSlot = 0
    
                .Invent.CascoEqpObjIndex = 0
                .Invent.CascoEqpSlot = 0
    
                .Invent.EscudoEqpObjIndex = 0
                .Invent.EscudoEqpSlot = 0
    
                .Invent.AnilloEqpObjIndex = 0
                .Invent.AnilloEqpSlot = 0
    
                .Invent.MunicionEqpObjIndex = 0
                .Invent.MunicionEqpSlot = 0
    
                .Invent.BarcoObjIndex = 0
                .Invent.BarcoSlot = 0
                
                .Invent.MonturaObjIndex = 0
                .Invent.MonturaSlot = 0
    
    
        End With

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 23/01/2007
        '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
        '***************************************************
        On Error GoTo errhandleR

        'If Cantidad > 100000 Then Exit Sub

        With UserList(UserIndex)

                'SI EL Pjta TIENE ORO LO TIRAMOS
                If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then

                        Dim MiObj    As Obj
                    
                        Dim Extra    As Long

                        Dim TeniaOro As Long

                        'info debug
                        Dim Loops    As Integer
            
                        'Seguridad Alkon (guardo el oro tirado si supera los 50k)
                        If Cantidad > 100000 Then

                                Dim j        As Integer

                                Dim k        As Integer

                                Dim m        As Integer

                                Dim Cercanos As String

                                m = .pos.Map

                                For j = .pos.X - 10 To .pos.X + 10
                                        For k = .pos.Y - 10 To .pos.Y + 10

                                                If InMapBounds(m, j, k) Then
                                                        If MapData(m, j, k).UserIndex > 0 Then
                                                                Cercanos = Cercanos & UserList(MapData(m, j, k).UserIndex).Name & ","

                                                        End If

                                                End If

                                        Next k
                                Next j

                                Call LogDesarrollo(.Name & " tira oro. Cercanos: " & Cercanos)

                        End If

                        '/Seguridad
                        
                        TeniaOro = .Stats.GLD

                        If Cantidad > 500000 Then 'Para evitar explotar demasiado
                                Extra = Cantidad - 500000
                                Cantidad = 500000

                        End If
            
                        Do While (Cantidad > 0)
                
                                If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
                                        MiObj.Amount = MAX_INVENTORY_OBJS
                                        Cantidad = Cantidad - MiObj.Amount
                                Else
                                        MiObj.Amount = Cantidad
                                        Cantidad = Cantidad - MiObj.Amount

                                End If
    
                                MiObj.objIndex = iORO
                
                                If esGM(UserIndex) Then Call LogGM(.Name, "Tiró cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.objIndex).Name)

                                Dim AuxPos As WorldPos
                
                                If .clase = eClass.Pirat And .Invent.BarcoObjIndex = 476 Then
                                        AuxPos = TirarItemAlPiso(.pos, MiObj, False)

                                        If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                                                .Stats.GLD = .Stats.GLD - MiObj.Amount

                                        End If

                                Else
                                        AuxPos = TirarItemAlPiso(.pos, MiObj, True)

                                        If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                                                .Stats.GLD = .Stats.GLD - MiObj.Amount

                                        End If

                                End If
                
                                'info debug
                                Loops = Loops + 1

                                If Loops > 100 Then
                                        LogError ("Error en tiraroro")
                                        Exit Sub

                                End If
                
                        Loop

                        If TeniaOro = .Stats.GLD Then Extra = 0
                        If Extra > 0 Then
                                .Stats.GLD = .Stats.GLD - Extra

                        End If
        
                End If

        End With

        Exit Sub

errhandleR:
        Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Cantidad As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error GoTo errhandleR

        If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
        With UserList(UserIndex).Invent.Object(Slot)

                If .Amount <= Cantidad And .Equipped = 1 Then
                        Call Desequipar(UserIndex, Slot)

                End If
        
                'Quita un objeto
                .Amount = .Amount - Cantidad

                '¿Quedan mas?
                If .Amount <= 0 Then
                        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                        .objIndex = 0
                        .Amount = 0

                End If

        End With

        Exit Sub

errhandleR:
        Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, _
                  ByVal UserIndex As Integer, _
                  ByVal Slot As Byte)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error GoTo errhandleR

        Dim NullObj As UserOBJ

        Dim LoopC   As Long

        With UserList(UserIndex)

                'Actualiza un solo slot
                If Not UpdateAll Then
    
                        'Actualiza el inventario
                        If .Invent.Object(Slot).objIndex > 0 Then
                                Call ChangeUserInv(UserIndex, Slot, .Invent.Object(Slot))
                        Else
                                Call ChangeUserInv(UserIndex, Slot, NullObj)

                        End If
    
                Else
    
                        'Actualiza todos los slots
                        For LoopC = 1 To .CurrentInventorySlots

                                'Actualiza el inventario
                                If .Invent.Object(LoopC).objIndex > 0 Then
                                        Call ChangeUserInv(UserIndex, LoopC, .Invent.Object(LoopC))
                                Else
                                        Call ChangeUserInv(UserIndex, LoopC, NullObj)

                                End If

                        Next LoopC

                End If
    
                Exit Sub

        End With

errhandleR:
        Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByVal Slot As Byte, _
            ByVal num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 11/5/2010
        '11/5/2010 - ZaMa: Arreglo bug que permitia apilar mas de 10k de items.
        '***************************************************

        Dim DropObj As Obj

        Dim MapObj  As Obj

        With UserList(UserIndex)

                If num > 0 Then
        
                        DropObj.objIndex = .Invent.Object(Slot).objIndex
        
                        If (ItemNewbie(DropObj.objIndex) And (.flags.Privilegios And PlayerType.User)) Then
                                Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                        End If
                
                        If num < 0 Then Exit Sub
                        If .flags.Comerciando Then Exit Sub
                        
                        DropObj.Amount = MinimoInt(num, .Invent.Object(Slot).Amount)

                        'Check objeto en el suelo
                        MapObj.objIndex = MapData(.pos.Map, X, Y).ObjInfo.objIndex
                        MapObj.Amount = MapData(.pos.Map, X, Y).ObjInfo.Amount
        
                        If MapObj.objIndex = 0 Or MapObj.objIndex = DropObj.objIndex Then
        
                                If MapObj.Amount = MAX_INVENTORY_OBJS Then
                                        Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub
                                End If
            
                                If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
                                        DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount
                                End If
            
                                Call MakeObj(DropObj, Map, X, Y)
                                Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
                                Call UpdateUserInv(False, UserIndex, Slot)
                                
                                If MapaLimpieza(.pos.Map) Then
                                
                                        If ObjData(DropObj.objIndex).SeLimpia = 0 Then ' se puede eliminar?
                                                If ObjData(DropObj.objIndex).OBJType <> eOBJType.otGuita Or ObjData(DropObj.objIndex).OBJType <> eOBJType.otTeleport Then ' las monedas no se borran
                                                        If MapData(Map, X, Y).Blocked <> 1 Then
                                                                Call aLimpiarMundo.AddItem(Map, X, Y)

                                                        End If

                                                End If

                                        End If

                                End If

If ObjData(DropObj.objIndex).OBJType = eOBJType.otMonturas Then
                Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU MONTURA!", FontTypeNames.FONTTYPE_TALK)
            End If
            
                                If ObjData(DropObj.objIndex).OBJType = eOBJType.otBarcos Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)

                                End If
            
                                If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Tiró cantidad:" & num & " Objeto:" & ObjData(DropObj.objIndex).Name)
            
                                'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
                                'Es un Objeto que tenemos que loguear?
                                If ObjData(DropObj.objIndex).Log = 1 Then
                                        Call LogItemsEspeciales(.Name & " tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.objIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                                ElseIf DropObj.Amount > 5000 Then 'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)

                                        'Si no es de los prohibidos de loguear, lo logueamos.
                                        If ObjData(DropObj.objIndex).NoLog <> 1 Then
                                                Call LogItemsEspeciales(.Name & " tiró al piso " & DropObj.Amount & " " & ObjData(DropObj.objIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)

                                        End If

                                End If

                        Else
                                Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)

                        End If

                End If

        End With

End Sub

Sub EraseObj(ByVal num As Integer, _
             ByVal Map As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        With MapData(Map, X, Y)
                .ObjInfo.Amount = .ObjInfo.Amount - num
    
                If .ObjInfo.Amount <= 0 Then
                        .ObjInfo.objIndex = 0
                        .ObjInfo.Amount = 0
        
                        Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))

                End If

        End With

End Sub

Sub MakeObj(ByRef Obj As Obj, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
    
        If Obj.objIndex > 0 And Obj.objIndex <= UBound(ObjData) Then
    
                With MapData(Map, X, Y)

                        If .ObjInfo.objIndex = Obj.objIndex Then
                                .ObjInfo.Amount = .ObjInfo.Amount + Obj.Amount
                        Else
                                .ObjInfo = Obj
                
                                If .trigger <> eTrigger.zonaOscura Then
                                        Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.objIndex).GrhIndex, X, Y))

                                End If

                        End If

                End With

        End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error GoTo errhandleR

        Dim Slot As Byte
    
        With UserList(UserIndex)
                '¿el user ya tiene un objeto del mismo tipo?
                Slot = 1
        
                Do Until .Invent.Object(Slot).objIndex = MiObj.objIndex And _
                   .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
                        Slot = Slot + 1

                        If Slot > .CurrentInventorySlots Then
                                Exit Do

                        End If

                Loop
            
                'Sino busca un slot vacio
                If Slot > .CurrentInventorySlots Then
                        Slot = 1

                        Do Until .Invent.Object(Slot).objIndex = 0
                                Slot = Slot + 1

                                If Slot > .CurrentInventorySlots Then
                                        Call WriteConsoleMsg(UserIndex, "No puedes cargar más objetos.", FontTypeNames.FONTTYPE_FIGHT)
                                        MeterItemEnInventario = False
                                        Exit Function

                                End If

                        Loop
                        .Invent.NroItems = .Invent.NroItems + 1

                End If

                'Mete el objeto
                If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
                        'Menor que MAX_INV_OBJS
                        .Invent.Object(Slot).objIndex = MiObj.objIndex
                        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
                Else
                        .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS

                End If

        End With
    
        MeterItemEnInventario = True
           
        Call UpdateUserInv(False, UserIndex, Slot)
    
        Exit Function
errhandleR:
        Call LogError("Error en MeterItemEnInventario. Error " & Err.Number & " : " & Err.description)

End Function

Sub GetObj(ByVal UserIndex As Integer)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 18/12/2009
        '18/12/2009: ZaMa - Oro directo a la billetera.
        '***************************************************

        Dim Obj    As ObjData

        Dim MiObj  As Obj

        Dim ObjPos As String
    
        With UserList(UserIndex)

                '¿Hay algun obj?
                If MapData(.pos.Map, .pos.X, .pos.Y).ObjInfo.objIndex > 0 Then

                        '¿Esta permitido agarrar este obj?
                        If ObjData(MapData(.pos.Map, .pos.X, .pos.Y).ObjInfo.objIndex).Agarrable <> 1 Then

                                Dim X As Integer

                                Dim Y As Integer
                
                                X = .pos.X
                                Y = .pos.Y
                
                                Obj = ObjData(MapData(.pos.Map, .pos.X, .pos.Y).ObjInfo.objIndex)
                                MiObj.Amount = MapData(.pos.Map, X, Y).ObjInfo.Amount
                                MiObj.objIndex = MapData(.pos.Map, X, Y).ObjInfo.objIndex
                
                                ' Oro directo a la billetera!
                                If Obj.OBJType = otGuita Then
                                        .Stats.GLD = .Stats.GLD + MiObj.Amount
                                        
                                        'Quitamos el objeto
                                        Call EraseObj(MapData(.pos.Map, X, Y).ObjInfo.Amount, .pos.Map, .pos.X, .pos.Y)
                                        Call SendData(SendTarget.ToPCArea, UserIndex, Protocol.PrepareMessageCreateDamage(.pos.X, .pos.Y, MiObj.Amount, 10))
                        
                                        Call WriteUpdateGold(UserIndex)
                                Else

                                        If MeterItemEnInventario(UserIndex, MiObj) Then
                    
                                                'Quitamos el objeto
                                                Call EraseObj(MapData(.pos.Map, X, Y).ObjInfo.Amount, .pos.Map, .pos.X, .pos.Y)

                                                If ObjData(MiObj.objIndex).SeLimpia = 0 Then ' juntamos algo que puede estar en la lista de limpieza
                                                        Call aLimpiarMundo.RemoveItem(.pos.Map, .pos.X, .pos.Y)

                                                End If

                                                If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.objIndex).Name)
        
                                                'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                                                'Es un Objeto que tenemos que loguear?
                                                If ObjData(MiObj.objIndex).Log = 1 Then
                                                        ObjPos = " Mapa: " & .pos.Map & " X: " & .pos.X & " Y: " & .pos.Y
                                                        Call LogItemsEspeciales(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.objIndex).Name & ObjPos)
                                                ElseIf MiObj.Amount > 5000 Then 'Es mucha cantidad?

                                                        'Si no es de los prohibidos de loguear, lo logueamos.
                                                        If ObjData(MiObj.objIndex).NoLog <> 1 Then
                                                                ObjPos = " Mapa: " & .pos.Map & " X: " & .pos.X & " Y: " & .pos.Y
                                                                Call LogItemsEspeciales(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.objIndex).Name & ObjPos)

                                                        End If

                                                End If

                                        End If

                                End If

                        End If

                Else
                        Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)

                End If

        End With

End Sub

Public Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '***************************************************
        'Author: Unknown
        'Last Modification: 26/05/2011
        '26/05/2011: Amraphen - Agregadas armaduras faccionarias de segunda jerarquía.
        '***************************************************

        On Error GoTo errhandleR

        'Desequipa el item slot del inventario
        Dim Obj As ObjData
    
        With UserList(UserIndex)
                With .Invent

                        If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
                                Exit Sub
                        ElseIf .Object(Slot).objIndex = 0 Then
                                Exit Sub

                        End If
            
                        Obj = ObjData(.Object(Slot).objIndex)

                End With
        
                Select Case Obj.OBJType

                        Case eOBJType.otWeapon

                                With .Invent
                                        .Object(Slot).Equipped = 0
                                        .WeaponEqpObjIndex = 0
                                        .WeaponEqpSlot = 0

                                End With
                
                                If Not .flags.Mimetizado = 1 Then

                                        With .Char
                                                .WeaponAnim = NingunArma
                                                Call ChangeUserChar(UserIndex, .Body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                                        End With

                                End If
            
                        Case eOBJType.otFlechas

                                With .Invent
                                        .Object(Slot).Equipped = 0
                                        .MunicionEqpObjIndex = 0
                                        .MunicionEqpSlot = 0

                                End With
            
                        Case eOBJType.otAnillo

                                With .Invent
                                        .Object(Slot).Equipped = 0
                                        .AnilloEqpObjIndex = 0
                                        .AnilloEqpSlot = 0

                                End With
            
                        Case eOBJType.otArmadura
                        
                        If .flags.Montando = 1 Then Exit Sub

                                'Nos fijamos si es armadura de segunda jerarquía
                                If Obj.Real = 2 Or Obj.Caos = 2 Then

                                        With .Invent
                                                .Object(Slot).Equipped = 0
                                                .FactionArmourEqpObjIndex = 0
                                                .FactionArmourEqpSlot = 0

                                        End With
                    
                                        'Cambiamos el body si tiene una armadura faccionaria de defensa alta, sino no pasa nada
                                        If .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta) Or _
                                           .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta) Then
                        
                                                If .flags.Mimetizado = 1 Then
                                                        .CharMimetizado.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                                                Else
                                                        .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje

                                                        With .Char
                                                                Call ChangeUserChar(UserIndex, .Body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                                                        End With

                                                End If

                                        End If
                    
                                        'Actualizo AMBOS slots de armaduras (el normal y el de segunda jerarquía)
                                        Call UpdateUserInv(False, UserIndex, Slot)
                                        Call UpdateUserInv(False, UserIndex, .Invent.ArmourEqpSlot)
                    
                                        Exit Sub
                    
                                Else

                                        With .Invent

                                                'Si tiene armadura faccionaria de segunda jerarquía equipada la sacamos:
                                                If .FactionArmourEqpObjIndex Then
                                                        Call Desequipar(UserIndex, .FactionArmourEqpSlot)

                                                End If
                    
                                                .Object(Slot).Equipped = 0
                                                .ArmourEqpObjIndex = 0
                                                .ArmourEqpSlot = 0

                                        End With
                    
                                        Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)

                                        With .Char
                                                Call ChangeUserChar(UserIndex, .Body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                                        End With

                                End If
                 
                        Case eOBJType.otCASCO

                                With .Invent
                                        .Object(Slot).Equipped = 0
                                        .CascoEqpObjIndex = 0
                                        .CascoEqpSlot = 0

                                End With
                
                                If Not .flags.Mimetizado = 1 Then

                                        With .Char
                                                .CascoAnim = NingunCasco
                                                Call ChangeUserChar(UserIndex, .Body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                                        End With

                                End If
            
                        Case eOBJType.otESCUDO

                                With .Invent
                                        .Object(Slot).Equipped = 0
                                        .EscudoEqpObjIndex = 0
                                        .EscudoEqpSlot = 0

                                End With
                
                                If Not .flags.Mimetizado = 1 Then

                                        With .Char
                                                .ShieldAnim = NingunEscudo
                                                Call ChangeUserChar(UserIndex, .Body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                                        End With

                                End If
            
                End Select

        End With
    
        Call WriteUpdateUserStats(UserIndex)
        Call UpdateUserInv(False, UserIndex, Slot)
    
        Exit Sub

errhandleR:
        Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.description)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, _
                           ByVal objIndex As Integer, _
                           Optional ByRef sMotivo As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 14/01/2010 (ZaMa)
        '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
        '***************************************************

        On Error GoTo errhandleR
    
        If ObjData(objIndex).Mujer = 1 Then
                SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
        ElseIf ObjData(objIndex).Hombre = 1 Then
                SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
        Else
                SexoPuedeUsarItem = True

        End If
    
        If Not SexoPuedeUsarItem Then sMotivo = "Tu género no puede usar este objeto."
    
        Exit Function
errhandleR:
        Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, _
                              ByVal objIndex As Integer, _
                              Optional ByRef sMotivo As String) As Boolean

        '***************************************************
        'Author: Unknown
        'Last Modification: 26/05/2011 (Amraphen)
        '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
        '26/05/2011: Amraphen - Agrego validación para armaduras faccionarias de segunda jerarquía.
        '***************************************************
        Dim ArmourIndex           As Integer

        Dim FaltaPrimeraJerarquia As Boolean
        If esGM(UserIndex) Then FaccionPuedeUsarItem = True: Exit Function
        
        If ObjData(objIndex).Real Then
                If Not criminal(UserIndex) And esArmada(UserIndex) Then
                        If ObjData(objIndex).Real = 2 Then
                                ArmourIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
                
                                If ArmourIndex > 0 And ObjData(ArmourIndex).Real = 1 Then
                                        FaccionPuedeUsarItem = True
                                Else
                                        FaccionPuedeUsarItem = False
                                        FaltaPrimeraJerarquia = True

                                End If

                        Else 'Es item faccionario común
                                FaccionPuedeUsarItem = True

                        End If

                Else
                        FaccionPuedeUsarItem = False

                End If

        ElseIf ObjData(objIndex).Caos Then

                If criminal(UserIndex) And esCaos(UserIndex) Then
                        If ObjData(objIndex).Caos = 2 Then
                                ArmourIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
                
                                If ArmourIndex > 0 And ObjData(ArmourIndex).Caos = 1 Then
                                        FaccionPuedeUsarItem = True
                                Else
                                        FaccionPuedeUsarItem = False
                                        FaltaPrimeraJerarquia = True

                                End If

                        Else 'Es item faccionario común
                                FaccionPuedeUsarItem = True

                        End If

                Else
                        FaccionPuedeUsarItem = False

                End If

        Else
                FaccionPuedeUsarItem = True

        End If
    
        If Not FaccionPuedeUsarItem Then
                If FaltaPrimeraJerarquia Then
                        sMotivo = "Debes tener equipada una armadura faccionaria."
                Else
                        sMotivo = "Tu alinación no puede usar este objeto."

                End If

        End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '*************************************************
        'Author: Unknown
        'Last modified: 26/05/2011 (Amraphen)
        '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
        '14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
        '26/05/2011: Amraphen - Agregadas armaduras faccionarias de segunda jerarquía.
        '*************************************************

        On Error GoTo errhandleR

        'Equipa un item del inventario
        Dim Obj      As ObjData

        Dim objIndex As Integer

        Dim sMotivo  As String
    
        With UserList(UserIndex)
                objIndex = .Invent.Object(Slot).objIndex
                Obj = ObjData(objIndex)
        
                If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                End If
                
                Select Case Obj.OBJType

                        Case eOBJType.otWeapon

                                If ClasePuedeUsarItem(UserIndex, objIndex, sMotivo) And _
                                   FaccionPuedeUsarItem(UserIndex, objIndex, sMotivo) Then

                                        'Si esta equipado lo quita
                                        If .Invent.Object(Slot).Equipped Then
                                                'Quitamos del inv el item
                                                Call Desequipar(UserIndex, Slot)

                                                'Animacion por defecto
                                                If .flags.Mimetizado = 1 Then
                                                        .CharMimetizado.WeaponAnim = NingunArma
                                                Else
                                                        .Char.WeaponAnim = NingunArma
                                                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                                End If

                                                Exit Sub

                                        End If
                    
                                        'Quitamos el elemento anterior
                                        If .Invent.WeaponEqpObjIndex > 0 Then
                                                Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                                        End If
                    
                                        .Invent.Object(Slot).Equipped = 1
                                        .Invent.WeaponEqpObjIndex = objIndex
                                        .Invent.WeaponEqpSlot = Slot
                    
                                        'El sonido solo se envia si no lo produce un admin invisible
                                        If Not (.flags.AdminInvisible = 1) Then _
                                           Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .pos.X, .pos.Y))
                    
                                        If .flags.Mimetizado = 1 Then
                                                .CharMimetizado.WeaponAnim = GetWeaponAnim(UserIndex, objIndex)
                                        Else
                                                .Char.WeaponAnim = GetWeaponAnim(UserIndex, objIndex)
                                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                        End If

                                Else
                                        Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                                End If
            
                        Case eOBJType.otAnillo

                                If ClasePuedeUsarItem(UserIndex, objIndex, sMotivo) And _
                                   FaccionPuedeUsarItem(UserIndex, objIndex, sMotivo) Then

                                        'Si esta equipado lo quita
                                        If .Invent.Object(Slot).Equipped Then
                                                'Quitamos del inv el item
                                                Call Desequipar(UserIndex, Slot)
                                                Exit Sub

                                        End If
                        
                                        'Quitamos el elemento anterior
                                        If .Invent.AnilloEqpObjIndex > 0 Then
                                                Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)

                                        End If
                
                                        .Invent.Object(Slot).Equipped = 1
                                        .Invent.AnilloEqpObjIndex = objIndex
                                        .Invent.AnilloEqpSlot = Slot
                        
                                Else
                                        Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                                End If
            
                        Case eOBJType.otFlechas

                                If ClasePuedeUsarItem(UserIndex, objIndex, sMotivo) And _
                                   FaccionPuedeUsarItem(UserIndex, objIndex, sMotivo) Then
                        
                                        'Si esta equipado lo quita
                                        If .Invent.Object(Slot).Equipped Then
                                                'Quitamos del inv el item
                                                Call Desequipar(UserIndex, Slot)
                                                Exit Sub

                                        End If
                        
                                        'Quitamos el elemento anterior
                                        If .Invent.MunicionEqpObjIndex > 0 Then
                                                Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)

                                        End If
                
                                        .Invent.Object(Slot).Equipped = 1
                                        .Invent.MunicionEqpObjIndex = objIndex
                                        .Invent.MunicionEqpSlot = Slot
                        
                                Else
                                        Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                                End If
            
                        Case eOBJType.otArmadura

                                If .flags.Navegando = 1 Then Exit Sub
                
                                'Nos aseguramos que puede usarla
                                If ClasePuedeUsarItem(UserIndex, objIndex, sMotivo) And _
                                   SexoPuedeUsarItem(UserIndex, objIndex, sMotivo) And _
                                   CheckRazaUsaRopa(UserIndex, objIndex, sMotivo) And _
                                   FaccionPuedeUsarItem(UserIndex, objIndex, sMotivo) Then
                    
                                        'Nos fijamos si es armadura de segunda jerarquia
                                        If Obj.Real = 2 Or Obj.Caos = 2 Then

                                                'Si esta equipado lo quita
                                                If .Invent.Object(Slot).Equipped Then
                                                        Call Desequipar(UserIndex, Slot)
                            
                                                        If Not .flags.Mimetizado = 1 Then
                                                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                                        End If
                            
                                                        Exit Sub

                                                End If
                        
                                                'Quita el anterior
                                                If .Invent.FactionArmourEqpObjIndex > 0 Then
                                                        Call Desequipar(UserIndex, .Invent.FactionArmourEqpSlot)

                                                End If
                        
                                                'Lo equipa
                                                .Invent.Object(Slot).Equipped = 1
                                                .Invent.FactionArmourEqpObjIndex = objIndex
                                                .Invent.FactionArmourEqpSlot = Slot
                        
                                                If .flags.Mimetizado = 1 Then
                                                        .CharMimetizado.Body = Obj.Ropaje
                                                Else

                                                        If .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta) Or _
                                                           .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta) Then
                                                                .Char.Body = Obj.Ropaje
                                                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                                        End If

                                                End If

                                        Else

                                                'Si esta equipado lo quita
                                                If .Invent.Object(Slot).Equipped Then
                                                        Call Desequipar(UserIndex, Slot)
                            
                                                        'Esto está de más:
                                                        'Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                                                        If Not .flags.Mimetizado = 1 Then
                                                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                                        End If
                            
                                                        Exit Sub

                                                End If
                
                                                'Quita el anterior
                                                If .Invent.ArmourEqpObjIndex > 0 Then
                                                        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

                                                End If
                
                                                'Lo equipa
                                                .Invent.Object(Slot).Equipped = 1
                                                .Invent.ArmourEqpObjIndex = objIndex
                                                .Invent.ArmourEqpSlot = Slot
                            
                                                If .flags.Mimetizado = 1 Then
                                                        .CharMimetizado.Body = Obj.Ropaje
                                                Else
                                                        .Char.Body = Obj.Ropaje
                                                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                                End If

                                                .flags.Desnudo = 0

                                        End If

                                Else
                                        Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                                End If
            
                        Case eOBJType.otCASCO

                                If .flags.Navegando = 1 Then Exit Sub
                                If ClasePuedeUsarItem(UserIndex, objIndex, sMotivo) Then

                                        'Si esta equipado lo quita
                                        If .Invent.Object(Slot).Equipped Then
                                                Call Desequipar(UserIndex, Slot)

                                                If .flags.Mimetizado = 1 Then
                                                        .CharMimetizado.CascoAnim = NingunCasco
                                                Else
                                                        .Char.CascoAnim = NingunCasco
                                                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                                End If

                                                Exit Sub

                                        End If
            
                                        'Quita el anterior
                                        If .Invent.CascoEqpObjIndex > 0 Then
                                                Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

                                        End If
            
                                        'Lo equipa
                    
                                        .Invent.Object(Slot).Equipped = 1
                                        .Invent.CascoEqpObjIndex = objIndex
                                        .Invent.CascoEqpSlot = Slot

                                        If .flags.Mimetizado = 1 Then
                                                .CharMimetizado.CascoAnim = Obj.CascoAnim
                                        Else
                                                .Char.CascoAnim = Obj.CascoAnim
                                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                        End If

                                Else
                                        Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                                End If
            
                        Case eOBJType.otESCUDO

                                If .flags.Navegando = 1 Then Exit Sub
                
                                If ClasePuedeUsarItem(UserIndex, objIndex, sMotivo) And _
                                   FaccionPuedeUsarItem(UserIndex, objIndex, sMotivo) Then
        
                                        'Si esta equipado lo quita
                                        If .Invent.Object(Slot).Equipped Then
                                                Call Desequipar(UserIndex, Slot)

                                                If .flags.Mimetizado = 1 Then
                                                        .CharMimetizado.ShieldAnim = NingunEscudo
                                                Else
                                                        .Char.ShieldAnim = NingunEscudo
                                                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                                End If

                                                Exit Sub

                                        End If
             
                                        'Quita el anterior
                                        If .Invent.EscudoEqpObjIndex > 0 Then
                                                Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                                        End If
             
                                        'Lo equipa
                     
                                        .Invent.Object(Slot).Equipped = 1
                                        .Invent.EscudoEqpObjIndex = objIndex
                                        .Invent.EscudoEqpSlot = Slot
                     
                                        If .flags.Mimetizado = 1 Then
                                                .CharMimetizado.ShieldAnim = Obj.ShieldAnim
                                        Else
                                                .Char.ShieldAnim = Obj.ShieldAnim
                         
                                                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                                        End If

                                Else
                                        Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                                End If
            
                End Select

        End With
    
        'Actualiza
        Call UpdateUserInv(False, UserIndex, Slot)
    
        Exit Sub
    
errhandleR:
        Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)

End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, _
                                  ByVal ItemIndex As Integer, _
                                  Optional ByRef sMotivo As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 14/01/2010 (ZaMa)
        '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
        '***************************************************

        On Error GoTo errhandleR
        If esGM(UserIndex) Then CheckRazaUsaRopa = True: Exit Function
        
        With UserList(UserIndex)

                'Verifica si la raza puede usar la ropa
                If .raza = eRaza.Humano Or _
                   .raza = eRaza.Elfo Or _
                   .raza = eRaza.Drow Then
                        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
                Else
                        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)

                End If
        
                'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
                If (.raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
                        CheckRazaUsaRopa = False

                End If

        End With
    
        If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    
        Exit Function
    
errhandleR:
        Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal byClick As Byte)

        '*************************************************
        'Author: Unknown
        'Last modified: 10/12/2009
        'Handels the usage of items from inventory box.
        '24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
        '24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
        '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
        '17/11/2009: ZaMa - Ahora se envia una orientacion de la posicion hacia donde esta el que uso el cuerno.
        '27/11/2009: Budi - Se envia indivualmente cuando se modifica a la Agilidad o la Fuerza del personaje.
        '08/12/2009: ZaMa - Agrego el uso de hacha de madera elfica.
        '10/12/2009: ZaMa - Arreglos y validaciones en todos las herramientas de trabajo.
        '*************************************************

        Dim Obj      As ObjData

        Dim objIndex As Integer

        Dim TargObj  As ObjData

        Dim MiObj    As Obj
    
        With UserList(UserIndex)
    
                If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
                Obj = ObjData(.Invent.Object(Slot).objIndex)
        
                If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                End If
        
                If Obj.OBJType = eOBJType.otWeapon Then
                        If Obj.proyectil = 1 Then
                
                                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                                If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
                        Else

                                'dagas
                                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

                        End If

                Else

                        If byClick = 1 Then
                                
                                If Not IntervaloPermiteUsarClick(UserIndex) Then
                                Exit Sub
                                End If
                        Else

                                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

                        End If

                End If
        
                objIndex = .Invent.Object(Slot).objIndex
                .flags.TargetObjInvIndex = objIndex
                .flags.TargetObjInvSlot = Slot
        
                Select Case Obj.OBJType

                        Case eOBJType.otUseOnce
        
                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If
        
                                'Usa el item
                                .Stats.MinHam = .Stats.MinHam + Obj.MinHam

                                If .Stats.MinHam > .Stats.MaxHam Then _
                                   .Stats.MinHam = .Stats.MaxHam
                                .flags.Hambre = 0
                                Call WriteUpdateHungerAndThirst(UserIndex)
                                'Sonido
                
                                If objIndex = e_ObjetosCriticos.Manzana Or objIndex = e_ObjetosCriticos.Manzana2 Or objIndex = e_ObjetosCriticos.ManzanaNewbie Then
                                        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
                                Else
                                        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)

                                End If
                
                                'Quitamos del inv el item
                                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                                Call UpdateUserInv(False, UserIndex, Slot)
        
                        Case eOBJType.otGuita

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).Amount
                                .Invent.Object(Slot).Amount = 0
                                .Invent.Object(Slot).objIndex = 0
                                .Invent.NroItems = .Invent.NroItems - 1
                
                                Call UpdateUserInv(False, UserIndex, Slot)
                                Call WriteUpdateGold(UserIndex)
                
                        Case eOBJType.otWeapon

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                If Not .Stats.MinSta > 0 Then
                                        Call WriteConsoleMsg(UserIndex, "Estás muy cansad" & _
                                           IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If
                
                                If ObjData(objIndex).proyectil = 1 Then
                                        If .Invent.Object(Slot).Equipped = 0 Then
                                                Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If

                                        Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                                ElseIf .flags.TargetObj = Leña Then

                                        If .Invent.Object(Slot).objIndex = DAGA Then
                                                If .Invent.Object(Slot).Equipped = 0 Then
                                                        Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
                                                        Exit Sub

                                                End If
                            
                                                Call TratarDeHacerFogata(.flags.TargetObjMap, _
                                                   .flags.TargetObjX, .flags.TargetObjY, UserIndex)

                                        End If

                                Else
                    
                                        Select Case objIndex
                    
                                                Case CAÑA_PESCA, RED_PESCA, CAÑA_PESCA_NEWBIE
                            
                                                        ' Lo tiene equipado?
                                                        If .Invent.WeaponEqpObjIndex = objIndex Then
                                                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Pesca)  'Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
                                                        Else
                                                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                                                        End If
                            
                                                Case HACHA_LEÑADOR, HACHA_LEÑA_ELFICA, HACHA_LEÑADOR_NEWBIE
                            
                                                        ' Lo tiene equipado?
                                                        If .Invent.WeaponEqpObjIndex = objIndex Then
                                                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Talar)
                                                        Else
                                                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                                                        End If
                            
                                                Case PIQUETE_MINERO, PIQUETE_MINERO_NEWBIE
                        
                                                        ' Lo tiene equipado?
                                                        If .Invent.WeaponEqpObjIndex = objIndex Then
                                                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
                                                        Else
                                                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                                                        End If
                            
                                                Case MARTILLO_HERRERO, MARTILLO_HERRERO_NEWBIE
                        
                                                        ' Lo tiene equipado?
                                                        If .Invent.WeaponEqpObjIndex = objIndex Then
                                                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Herreria)
                                                        Else
                                                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                                                        End If
                            
                                                Case SERRUCHO_CARPINTERO, SERRUCHO_CARPINTERO_NEWBIE
                            
                                                        ' Lo tiene equipado?
                                                        If .Invent.WeaponEqpObjIndex = objIndex Then
                                                                Call EnivarObjConstruibles(UserIndex)
                                                                Call WriteShowCarpenterForm(UserIndex)
                                                        Else
                                                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                                                        End If
                            
                                                Case Else ' Las herramientas no se pueden fundir

                                                        If ObjData(objIndex).SkHerreria > 0 Then
                                                                ' Solo objetos que pueda hacer el herrero
                                                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)

                                                        End If

                                        End Select

                                End If
            
                        Case eOBJType.otPociones
                
                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If
                                
                                If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If
                                        
                                Dim TActual As Long
    
                                TActual = GetTickCount() And &H7FFFFFFF
    
                                If getInterval(TActual, .Counters.LastPoteo) >= Intervalos(eIntervalos.iPuedeUsarItemDblClick) Then
                
                                        .Counters.LastPoteo = TActual
                                        .Counters.failedUsageAttempts = 0
                             
                                        .flags.TomoPocion = True
                                        .flags.TipoPocion = Obj.TipoPocion
                        
                                        Select Case .flags.TipoPocion
                
                                                Case 1 'Modif la agilidad
                                                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                                                        'Usa el item
                                                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                                                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                                                           .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS

                                                        If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
                        
                                                        'Quitamos del inv el item
                                                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                                                        ' Los admin invisibles solo producen sonidos a si mismos
                                                        If .flags.AdminInvisible = 1 Then
                                                                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))
                                                        Else
                                                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))

                                                        End If

                                                        Call WriteUpdateDexterity(UserIndex)
                        
                                                Case 2 'Modif la fuerza
                                                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                                                        'Usa el item
                                                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                                                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                                                           .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS

                                                        If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                        
                                                        'Quitamos del inv el item
                                                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                                                        ' Los admin invisibles solo producen sonidos a si mismos
                                                        If .flags.AdminInvisible = 1 Then
                                                                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))
                                                        Else
                                                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))

                                                        End If

                                                        Call WriteUpdateStrenght(UserIndex)
                        
                                                Case 3 'Pocion roja, restaura HP
                                                        'Usa el item
                                                        .Stats.MinHp = .Stats.MinHp + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                                                        If .Stats.MinHp > .Stats.MaxHP Then _
                                                           .Stats.MinHp = .Stats.MaxHP
                        
                                                        'Quitamos del inv el item
                                                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                                                        ' Los admin invisibles solo producen sonidos a si mismos
                                                        If .flags.AdminInvisible = 1 Then
                                                                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))
                                                        Else
                                                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))

                                                        End If
                    
                                                Case 4 'Pocion azul, restaura MANA
                                                        'Usa el item
                                                        'nuevo calculo para recargar mana
                                                        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV

                                                        If .Stats.MinMAN > .Stats.MaxMAN Then _
                                                           .Stats.MinMAN = .Stats.MaxMAN
                        
                                                        'Quitamos del inv el item
                                                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                                                        ' Los admin invisibles solo producen sonidos a si mismos
                                                        If .flags.AdminInvisible = 1 Then
                                                                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))
                                                        Else
                                                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))

                                                        End If
                        
                                                Case 5 ' Pocion violeta

                                                        If .flags.Envenenado = 1 Then
                                                                .flags.Envenenado = 0
                                                                Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)

                                                        End If

                                                        'Quitamos del inv el item
                                                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                                                        ' Los admin invisibles solo producen sonidos a si mismos
                                                        If .flags.AdminInvisible = 1 Then
                                                                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))
                                                        Else
                                                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))

                                                        End If
                        
                                                Case 6  ' Pocion Negra

                                                        If .flags.Privilegios And PlayerType.User Then
                                                                Call QuitarUserInvItem(UserIndex, Slot, 1)
                                                                Call UserDie(UserIndex)
                                                                Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)

                                                        End If

                                        End Select

                                        Call WriteUpdateUserStats(UserIndex)
                                        Call UpdateUserInv(False, UserIndex, Slot)
                                Else
                                        .Counters.failedUsageAttempts = .Counters.failedUsageAttempts + 1
                    
                                        If .Counters.failedUsageAttempts = Tolerancia_FailIntervalo Then ' @@ Avisamos por consola posible chitero
                                              '  Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Posible alteracion de intervalos por parte de : " & .Name & " Hora: " & time$, FontTypeNames.FONTTYPE_EJECUCION))
                                                .Counters.failedUsageAttempts = 0
                                                Exit Sub

                                        End If

                                End If
        
                        Case eOBJType.otBebidas

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed

                                If .Stats.MinAGU > .Stats.MaxAGU Then _
                                   .Stats.MinAGU = .Stats.MaxAGU
                                .flags.Sed = 0
                                Call WriteUpdateHungerAndThirst(UserIndex)
                
                                'Quitamos del inv el item
                                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                                ' Los admin invisibles solo producen sonidos a si mismos
                                If .flags.AdminInvisible = 1 Then
                                        Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))
                                Else
                                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.X, .pos.Y))

                                End If
                
                                Call UpdateUserInv(False, UserIndex, Slot)
            
                        Case eOBJType.otLlaves

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                If .flags.TargetObj = 0 Then Exit Sub
                                TargObj = ObjData(.flags.TargetObj)

                                '¿El objeto clickeado es una puerta?
                                If TargObj.OBJType = eOBJType.otPuertas Then

                                        '¿Esta cerrada?
                                        If TargObj.Cerrada = 1 Then

                                                '¿Cerrada con llave?
                                                If TargObj.Llave > 0 Then
                                                        If TargObj.clave = Obj.clave Then
                 
                                                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex _
                                                                   = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex).IndexCerrada
                                                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex
                                                                Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                                                                Exit Sub
                                                        Else
                                                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                                                Exit Sub

                                                        End If

                                                Else

                                                        If TargObj.clave = Obj.clave Then
                                                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex _
                                                                   = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex).IndexCerradaLlave
                                                                Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex
                                                                Exit Sub
                                                        Else
                                                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                                                Exit Sub

                                                        End If

                                                End If

                                        Else
                                                Call WriteConsoleMsg(UserIndex, "No está cerrada.", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If

                                End If
            
                        Case eOBJType.otBotellaVacia

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                If Not HayAgua(.pos.Map, .flags.TargetX, .flags.TargetY) Then
                                        Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                MiObj.Amount = 1
                                MiObj.objIndex = ObjData(.Invent.Object(Slot).objIndex).IndexAbierta
                                Call QuitarUserInvItem(UserIndex, Slot, 1)

                                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                                        Call TirarItemAlPiso(.pos, MiObj)

                                End If
                
                                Call UpdateUserInv(False, UserIndex, Slot)
            
                        Case eOBJType.otBotellaLlena

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed

                                If .Stats.MinAGU > .Stats.MaxAGU Then _
                                   .Stats.MinAGU = .Stats.MaxAGU
                                .flags.Sed = 0
                                Call WriteUpdateHungerAndThirst(UserIndex)
                                MiObj.Amount = 1
                                MiObj.objIndex = ObjData(.Invent.Object(Slot).objIndex).IndexCerrada
                                Call QuitarUserInvItem(UserIndex, Slot, 1)

                                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                                        Call TirarItemAlPiso(.pos, MiObj)

                                End If
                
                                Call UpdateUserInv(False, UserIndex, Slot)
            
                        Case eOBJType.otPergaminos

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                If .Stats.MaxMAN > 0 Then
                                        If .flags.Hambre = 0 And _
                                           .flags.Sed = 0 Then
                                                Call AgregarHechizo(UserIndex, Slot)
                                                Call UpdateUserInv(False, UserIndex, Slot)
                                        Else
                                                Call WriteConsoleMsg(UserIndex, "Estás demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)

                                        End If

                                Else
                                        Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)

                                End If

                        Case eOBJType.otMinerales

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
               
                        Case eOBJType.otInstrumentos

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                If Obj.Real Then '¿Es el Cuerno Real?
                                        If FaccionPuedeUsarItem(UserIndex, objIndex) Then
                                                If MapInfo(.pos.Map).Pk = False Then
                                                        Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                                                        Exit Sub

                                                End If
                        
                                                ' Los admin invisibles solo producen sonidos a si mismos
                                                If .flags.AdminInvisible = 1 Then
                                                        Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .pos.X, .pos.Y))
                                                Else
                                                        Call AlertarFaccionarios(UserIndex)
                                                        Call SendData(SendTarget.toMap, .pos.Map, PrepareMessagePlayWave(Obj.Snd1, .pos.X, .pos.Y))

                                                End If
                        
                                                Exit Sub
                                        Else
                                                Call WriteConsoleMsg(UserIndex, "Sólo miembros del ejército real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If

                                ElseIf Obj.Caos Then '¿Es el Cuerno Legión?

                                        If FaccionPuedeUsarItem(UserIndex, objIndex) Then
                                                If MapInfo(.pos.Map).Pk = False Then
                                                        Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                                                        Exit Sub

                                                End If
                        
                                                ' Los admin invisibles solo producen sonidos a si mismos
                                                If .flags.AdminInvisible = 1 Then
                                                        Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .pos.X, .pos.Y))
                                                Else
                                                        Call AlertarFaccionarios(UserIndex)
                                                        Call SendData(SendTarget.toMap, .pos.Map, PrepareMessagePlayWave(Obj.Snd1, .pos.X, .pos.Y))

                                                End If
                        
                                                Exit Sub
                                        Else
                                                Call WriteConsoleMsg(UserIndex, "Sólo miembros de la legión oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub

                                        End If

                                End If

                                'Si llega aca es porque es o Laud o Tambor o Flauta
                                ' Los admin invisibles solo producen sonidos a si mismos
                                If .flags.AdminInvisible = 1 Then
                                        Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .pos.X, .pos.Y))
                                Else
                                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .pos.X, .pos.Y))

                                End If
               
                        Case eOBJType.otBarcos

                                'Verifica si esta aproximado al agua antes de permitirle navegar
                                If .Stats.ELV < 25 Then

                                        ' Solo pirata y trabajador pueden navegar antes
                                        If .clase <> eClass.Worker And .clase <> eClass.Pirat Then
                                                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                                                Exit Sub
                                        Else

                                                ' Pero a partir de 20
                                                If .Stats.ELV < 20 Then
                            
                                                        If .clase = eClass.Worker And .Stats.UserSkills(eSkill.Pesca) <> 100 Then
                                                                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 y además tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                                                        Else
                                                                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)

                                                        End If
                            
                                                        Exit Sub
                                                Else

                                                        ' Esta entre 20 y 25, si es trabajador necesita tener 100 en pesca
                                                        If .clase = eClass.Worker Then
                                                                If .Stats.UserSkills(eSkill.Pesca) <> 100 Then
                                                                        Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior y además tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                                                                        Exit Sub

                                                                End If

                                                        End If

                                                End If

                                        End If

                                End If
                
                                If ((LegalPos(.pos.Map, .pos.X - 1, .pos.Y, True, False) _
                                   Or LegalPos(.pos.Map, .pos.X, .pos.Y - 1, True, False) _
                                   Or LegalPos(.pos.Map, .pos.X + 1, .pos.Y, True, False) _
                                   Or LegalPos(.pos.Map, .pos.X, .pos.Y + 1, True, False)) _
                                   And .flags.Navegando = 0) _
                                   Or .flags.Navegando = 1 Then
                                        Call DoNavega(UserIndex, Obj, Slot)
                                Else
                                        Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)

                                End If
                                
                                Case eOBJType.otMonturas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "|¡¡Estas muerto!! Los muertos no dominan los animales. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                End If
                If ((LegalPos(.pos.Map, .pos.X, .pos.Y, True, False) _
                        Or LegalPos(.pos.Map, .pos.X, .pos.Y, True, False) _
                        Or LegalPos(.pos.Map, .pos.X, .pos.Y, True, False) _
                        Or LegalPos(.pos.Map, .pos.X, .pos.Y, True, False)) _
                        And .flags.Navegando = 0) _
                        Or .flags.Navegando = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes montar en el agua!", FontTypeNames.FONTTYPE_INFO)
                Else
                Call DoEquita(UserIndex, Obj, Slot)
                End If
                
                        Case eOBJType.otCofresDonador

                                If .flags.Muerto = 1 Then
                                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub

                                End If

                                If .InvSuma = 0 Then
                                        If Not .flags.UltimoMensaje = 100 Then
                                                .flags.UltimoMensaje = 100
                        
                                                Call WriteConsoleMsg(UserIndex, "Recuerda tener 5 slots libres en tu inventario, el staff no se hace responsable de perdidas de objetos.", FontTypeNames.FONTTYPE_INFO)

                                        End If

                                        .InvSuma = .InvSuma + 1
                                        Exit Sub

                                End If

                                If .InvSuma = 1 Then
                                        Call ItemDrop_Shop(UserIndex, objIndex, Slot)
                                        .InvSuma = 0
                                        Exit Sub

                                End If

                End Select
    
        End With

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error GoTo errhandleR

        With UserList(UserIndex)

                If MapData(.pos.Map, .pos.X, .pos.Y).trigger = 6 Then Exit Sub
        
                Call TirarTodosLosItems(UserIndex)
        
                'Dim Cantidad As Long

                'Cantidad = .Stats.GLD - CLng(.Stats.ELV) * 10000
        
                'If Cantidad > 0 Then _
                 Call TirarOro(Cantidad, UserIndex)

        End With

        Exit Sub

errhandleR:
        Call LogError("Error en TirarTodo. Error: " & Err.Number & " - " & Err.description)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        With ObjData(Index)
                ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And _
                   (.Caos <> 1 Or .NoSeCae = 0) And _
                   .OBJType <> eOBJType.otLlaves And _
                   .OBJType <> eOBJType.otBarcos And _
                   .OBJType <> eOBJType.otMonturas And _
                   .NoSeCae = 0

        End With

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/2010 (ZaMa)
        '12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
        '***************************************************
        On Error GoTo errhandleR

        Dim i         As Byte

        Dim NuevaPos  As WorldPos

        Dim MiObj     As Obj

        Dim ItemIndex As Integer

        Dim DropAgua  As Boolean
    
        With UserList(UserIndex)

                For i = 1 To .CurrentInventorySlots
                        ItemIndex = .Invent.Object(i).objIndex

                        If ItemIndex > 0 Then
                                If ItemSeCae(ItemIndex) Then
                                        NuevaPos.X = 0
                                        NuevaPos.Y = 0
                    
                                        'Creo el Obj
                                        MiObj.Amount = .Invent.Object(i).Amount
                                        MiObj.objIndex = ItemIndex

                                        DropAgua = True

                                        ' Es pirata?
                                        If .clase = eClass.Pirat Then

                                                ' Si tiene galeon equipado
                                                If .Invent.BarcoObjIndex = 476 Then

                                                        ' Limitación por nivel, después dropea normalmente
                                                        If .Stats.ELV = 20 Then
                                                                ' No dropea en agua
                                                                DropAgua = False

                                                        End If

                                                End If

                                        End If
                    
                                        Call Tilelibre(.pos, NuevaPos, MiObj, DropAgua, True)
                    
                                        If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                                                Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                                        End If

                                End If

                        End If

                Next i

        End With
    
        Exit Sub
    
errhandleR:
        Call LogError("Error en TirarTodosLosItems. Error: " & Err.Number & " - " & Err.description)

End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
        ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: 23/11/2009
        '07/11/09: Pato - Fix bug #2819911
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
        Dim i         As Byte

        Dim NuevaPos  As WorldPos

        Dim MiObj     As Obj

        Dim ItemIndex As Integer
    
        With UserList(UserIndex)

                If MapData(.pos.Map, .pos.X, .pos.Y).trigger = 6 Then Exit Sub
        
                For i = 1 To UserList(UserIndex).CurrentInventorySlots
                        ItemIndex = .Invent.Object(i).objIndex

                        If ItemIndex > 0 Then
                                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                                        NuevaPos.X = 0
                                        NuevaPos.Y = 0
                    
                                        'Creo MiObj
                                        MiObj.Amount = .Invent.Object(i).Amount
                                        MiObj.objIndex = ItemIndex
                                        'Pablo (ToxicWaste) 24/01/2007
                                        'Tira los Items no newbies en todos lados.
                                        Tilelibre .pos, NuevaPos, MiObj, True, True

                                        If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                                                Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                                        End If

                                End If

                        End If

                Next i

        End With

End Sub

Public Function getObjType(ByVal objIndex As Integer) As eOBJType
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        If objIndex > 0 Then
                getObjType = ObjData(objIndex).OBJType

        End If
    
End Function
