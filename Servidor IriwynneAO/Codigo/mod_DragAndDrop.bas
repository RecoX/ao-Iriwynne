Attribute VB_Name = "mod_DragAndDrop"
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
 
Public Sub DragToUser(ByVal Userindex As Integer, _
                      ByVal tIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Amount As Integer)
   
        ' @@ Drag un slot a un usuario.
On Error GoTo errhandleR

        Dim tObj       As Obj

        Dim tString    As String
        
        Dim errorFound As String

        Dim Espacio    As Boolean
        
        ' Puede dragear ?
        If Not CanDragObj(UserList(Userindex).pos.Map, UserList(Userindex).flags.Navegando, errorFound) Then
                WriteConsoleMsg Userindex, errorFound, FontTypeNames.FONTTYPE_INFOBOLD

                Exit Sub
End If
                      If UserList(tIndex).flags.Muerto = 1 Then 'cambiado userindex por targetindex mankeada de HEKAMIAH
                                Call WriteConsoleMsg(Userindex, "�El usuario est� muerto!", FontTypeNames.FONTTYPE_INFO)
Exit Sub
        End If
        
        
        If UserList(tIndex).Death Then Exit Sub
        If UserList(tIndex).EnJDH Then Exit Sub
        
        
        'Preparo el objeto.
        tObj.Amount = Amount
        tObj.objIndex = UserList(Userindex).Invent.Object(Slot).objIndex
        
        If Not EsAdmin(UserList(Userindex).Name) Then ' @@ Solo los admins Pueden crear los items con NoCreable = 1 -18/11/2015
                
                If Not EsAdmin(UserList(Userindex).Name) Then
                        If ItemShop(UserList(Userindex).Invent.Object(Slot).objIndex) = True Then
                                Call WriteConsoleMsg(Userindex, "No puedes tirar items shop.", FontTypeNames.FONTTYPE_INFO) ' IvanLisz
                                Exit Sub

                        End If

                End If

        End If
        
         If (Amount <= 0 Or Amount > UserList(Userindex).Invent.Object(Slot).Amount) Then
                Call WriteConsoleMsg(Userindex, "No Tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

        End If
        
        'TmpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        Espacio = MeterItemEnInventario(tIndex, tObj)
 
        'No tiene espacio.

        If Not Espacio Then
        
                WriteConsoleMsg Userindex, "El usuario no tiene espacio en su inventario.", FontTypeNames.FONTTYPE_INFOBOLD



                Exit Sub

        End If
 
        'Quito el objeto.
        QuitarUserInvItem Userindex, Slot, Amount
 
        'Hago un update de su inventario.
        UpdateUserInv False, Userindex, Slot
 
        'Preparo el mensaje para userINdex (quien dragea)
 
        tString = "Le has arrojado"
 
        If tObj.Amount <> 1 Then
                tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
        Else
                tString = tString & " Tu " & ObjData(tObj.objIndex).Name

        End If
 
        tString = tString & " a " & UserList(tIndex).Name
 
        'Envio el mensaje
        WriteConsoleMsg Userindex, tString, FontTypeNames.FONTTYPE_INFOBOLD
 
        'Preparo el mensaje para el otro usuario (quien recibe)
        tString = UserList(Userindex).Name & " Te ha arrojado"
 
        If tObj.Amount <> 1 Then
                tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
        Else
                tString = tString & " su " & ObjData(tObj.objIndex).Name

        End If
 
        'Envio el mensaje al otro usuario
        WriteConsoleMsg tIndex, tString & ".", FontTypeNames.FONTTYPE_INFOBOLD
 
Exit Sub

errhandleR:
LogError "Error en DragToUser: " & Err.Number & " " & Err.description

End Sub
 
Public Sub DragToNPC(ByVal Userindex As Integer, _
                     ByVal tNpc As Integer, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)

        ' @@ Drag un slot a un npc.

        On Error GoTo errhandleR
 
        Dim TeniaOro   As Long

        Dim TeniaObj   As Integer

        Dim tmpIndex   As Integer
        
        Dim errorFound As String
 
1        tmpIndex = UserList(Userindex).Invent.Object(Slot).objIndex
2        TeniaOro = UserList(Userindex).Stats.GLD
3        TeniaObj = UserList(Userindex).Invent.Object(Slot).Amount
        
        ' Puede dragear ?
4        If Not CanDragObj(UserList(Userindex).pos.Map, UserList(Userindex).flags.Navegando, errorFound) Then
5                WriteConsoleMsg Userindex, errorFound, FontTypeNames.FONTTYPE_INFOBOLD

6                Exit Sub

        End If

7        If Not EsAdmin(UserList(Userindex).Name) Then
8                If ItemShop(tmpIndex) = True Then
9                        Call WriteConsoleMsg(Userindex, "No puedes tirar items shop.", FontTypeNames.FONTTYPE_INFO) ' IvanLisz
                        Exit Sub

                End If

        End If

        'End If
              
10         If (Amount <= 0 Or Amount > UserList(Userindex).Invent.Object(Slot).Amount) Then
11                Call WriteConsoleMsg(Userindex, "No Tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
12                Exit Sub

        End If
 
        'Es un banquero?

13        If Npclist(tNpc).NPCtype = eNPCType.Banquero Then
14                Call UserDejaObj(Userindex, Slot, Amount)
                
                'No tiene m�s el mismo amount que antes? entonces deposit�.

15                If TeniaObj <> UserList(Userindex).Invent.Object(Slot).Amount Then
16                        WriteConsoleMsg Userindex, "Has depositado " & Amount & " - " & ObjData(tmpIndex).Name & ".", FontTypeNames.FONTTYPE_INFOBOLD
17                        UpdateUserInv False, Userindex, Slot

18                End If

                'Es un npc comerciante?
19        ElseIf Npclist(tNpc).Comercia = 1 Then
                'El npc compra cualquier tipo de items?

20                If Not Npclist(tNpc).TipoItems <> eOBJType.otCualquiera Or Npclist(tNpc).TipoItems = ObjData(UserList(Userindex).Invent.Object(Slot).objIndex).OBJType Then
                        
21                        Call Comercio(eModoComercio.Venta, Userindex, tNpc, Slot, Amount)
                        
                        'Gan� oro? si es as� es porque lo vendi�.

22                        If TeniaOro <> UserList(Userindex).Stats.GLD Then
23                                WriteConsoleMsg Userindex, "Le has vendido al " & Npclist(tNpc).Name & " " & Amount & " - " & ObjData(tmpIndex).Name & ".", FontTypeNames.FONTTYPE_INFOBOLD

24                        End If

25                Else
26                        WriteConsoleMsg Userindex, "El npc no est� interesado en comprar este tipo de objetos.", FontTypeNames.FONTTYPE_INFOBOLD

                End If

        End If
 
        Exit Sub
 
errhandleR:
LogError "Error en DragToNPC en la linea " & Erl & ": " & Err.Number & " " & Err.description

End Sub
 
Public Sub DragToPos(ByVal Userindex As Integer, _
                     ByVal X As Byte, _
                     ByVal Y As Byte, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)

        '            Drag un slot a una posici�n.
 
 On Error GoTo errhandleR
 
        Dim errorFound As String

        Dim tObj       As Obj

        Dim tString    As String
        
        Dim tmpIndex   As Integer
        
        'No puede dragear en esa pos?
        
        
        If Not CanDragToPos(UserList(Userindex).pos.Map, X, Y, errorFound) Then
                WriteConsoleMsg Userindex, errorFound, FontTypeNames.FONTTYPE_INFOBOLD

                Exit Sub

        End If
     
        ' Puede dragear ?
        If Not CanDragObj(UserList(Userindex).pos.Map, UserList(Userindex).flags.Navegando, errorFound) Then
                WriteConsoleMsg Userindex, errorFound, FontTypeNames.FONTTYPE_INFOBOLD

                Exit Sub

        End If

        'Creo el objeto.
        tObj.objIndex = UserList(Userindex).Invent.Object(Slot).objIndex
        tObj.Amount = Amount
       
        If Not EsAdmin(UserList(Userindex).Name) Then
                If ItemShop(tObj.objIndex) = True Then
                        Call WriteConsoleMsg(Userindex, "No puedes tirar items shop.", FontTypeNames.FONTTYPE_INFO) ' IvanLisz
                        Exit Sub

                End If

        End If
        
         If (Amount <= 0 Or Amount > UserList(Userindex).Invent.Object(Slot).Amount) Then
                Call WriteConsoleMsg(Userindex, "No Tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

        End If
 
        'Agrego el objeto a la posici�n.
        MakeObj tObj, UserList(Userindex).pos.Map, CInt(X), CInt(Y)
 
        'Quito el objeto.
        QuitarUserInvItem Userindex, Slot, Amount
 
        'Actualizo el inventario
        UpdateUserInv False, Userindex, Slot
 
        'Preparo el mensaje.
        tString = "Has arrojado "
 
        If tObj.Amount <> 1 Then
                tString = tString & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
        Else
                tString = tString & "tu " & ObjData(tObj.objIndex).Name 'faltaba el tstring &

        End If
 
        'ENvio.
        WriteConsoleMsg Userindex, tString & ".", FontTypeNames.FONTTYPE_INFOBOLD
 Exit Sub
 
errhandleR:
LogError "Error en DragToPos: " & Err.Number & " " & Err.description
End Sub
 
Private Function CanDragToPos(ByVal Map As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByRef Error As String) As Boolean
 
        '            Devuelve si se puede dragear un item a x posici�n.
 
        CanDragToPos = False
 
        'Zona segura?

        If Not MapInfo(Map).Pk Then
                Error = "No est� permitido arrojar objetos al suelo en zonas seguras."

                Exit Function

        End If
 
        'Ya hay objeto?

        If Not MapData(Map, X, Y).ObjInfo.objIndex = 0 Then
                Error = "Hay un objeto en esa posici�n!"

                Exit Function

        End If
 
        'Tile bloqueado?
        
        

        If Not MapData(Map, X, Y).Blocked = 0 Then
                Error = "No puedes arrojar objetos en esa posici�n"

                Exit Function

        End If
        
        If HayAgua(Map, X, Y) Then
                Error = "No puedes arrojar objetos al agua"
                
                Exit Function

        End If

        CanDragToPos = True
 
End Function
 
Private Function CanDragObj(ByVal objIndex As Integer, _
                            ByVal Navegando As Boolean, _
                            ByRef Error As String) As Boolean
 
        '            Devuelve si un objeto es drageable.
        
        CanDragObj = False
 
        If objIndex < 1 Or objIndex > UBound(ObjData()) Then Exit Function
 
        'Objeto newbie?

        If ObjData(objIndex).Newbie <> 0 Then
                Error = "No puedes arrojar objetos newbies!"

                Exit Function

        End If
        
        
    
        If ObjData(objIndex).OBJType = eOBJType.otGuita Then
    
                Error = "No puedes arrojar oro!"
                Exit Function
        End If
        
        'Est� navgeando?

        If Navegando Then
                Error = "No puedes arrojar un barco si est�s navegando!"

                Exit Function

        End If
 
        CanDragObj = True
 
End Function

Public Sub moveItem(ByVal Userindex As Integer, _
                    ByVal originalSlot As Integer, _
                    ByVal newSlot As Integer)

        Dim tmpObj As UserOBJ

        If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub

        With UserList(Userindex)

                If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
    
                tmpObj = .Invent.Object(originalSlot)
                .Invent.Object(originalSlot) = .Invent.Object(newSlot)
                .Invent.Object(newSlot) = tmpObj
    
                'Viva VB6 y sus putas deficiencias.
                If .Invent.AnilloEqpSlot = originalSlot Then
                        .Invent.AnilloEqpSlot = newSlot
                ElseIf .Invent.AnilloEqpSlot = newSlot Then
                        .Invent.AnilloEqpSlot = originalSlot

                End If
    
                If .Invent.ArmourEqpSlot = originalSlot Then
                        .Invent.ArmourEqpSlot = newSlot
                ElseIf .Invent.ArmourEqpSlot = newSlot Then
                        .Invent.ArmourEqpSlot = originalSlot

                End If
    
                If .Invent.BarcoSlot = originalSlot Then
                        .Invent.BarcoSlot = newSlot
                ElseIf .Invent.BarcoSlot = newSlot Then
                        .Invent.BarcoSlot = originalSlot

                End If
    
                If .Invent.CascoEqpSlot = originalSlot Then
                        .Invent.CascoEqpSlot = newSlot
                ElseIf .Invent.CascoEqpSlot = newSlot Then
                        .Invent.CascoEqpSlot = originalSlot

                End If
    
                If .Invent.EscudoEqpSlot = originalSlot Then
                        .Invent.EscudoEqpSlot = newSlot
                ElseIf .Invent.EscudoEqpSlot = newSlot Then
                        .Invent.EscudoEqpSlot = originalSlot

                End If
    
                If .Invent.MunicionEqpSlot = originalSlot Then
                        .Invent.MunicionEqpSlot = newSlot
                ElseIf .Invent.MunicionEqpSlot = newSlot Then
                        .Invent.MunicionEqpSlot = originalSlot

                End If
    
                If .Invent.WeaponEqpSlot = originalSlot Then
                        .Invent.WeaponEqpSlot = newSlot
                ElseIf .Invent.WeaponEqpSlot = newSlot Then
                        .Invent.WeaponEqpSlot = originalSlot

                End If

                Call UpdateUserInv(False, Userindex, originalSlot)
                Call UpdateUserInv(False, Userindex, newSlot)

        End With

End Sub

