Attribute VB_Name = "InvNpc"
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

'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'                        Modulo Inv & Obj
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Modulo para controlar los objetos y los inventarios.
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
Public Function TirarItemAlPiso(Pos As WorldPos, _
                                Obj As Obj, _
                                Optional NotPirata As Boolean = True) As WorldPos
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error GoTo ErrHandler

        Dim NuevaPos As WorldPos

        NuevaPos.X = 0
        NuevaPos.Y = 0
    
        Tilelibre Pos, NuevaPos, Obj, NotPirata, True

        If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then

                ' @@ Los items dropeados por NPCs, tambien suman a la limpieza de mundo
                If MapaLimpieza(Pos.Map) Then
                                
                        If ObjData(Obj.objIndex).SeLimpia = 0 Then ' se puede eliminar?
                                If ObjData(Obj.objIndex).OBJType <> eOBJType.otGuita Or ObjData(Obj.objIndex).OBJType <> eOBJType.otTeleport Then ' las monedas no se borran
                                        If MapData(Pos.Map, NuevaPos.X, NuevaPos.Y).Blocked <> 1 Then
                                                Call aLimpiarMundo.AddItem(Pos.Map, NuevaPos.X, NuevaPos.Y)

                                        End If

                                End If

                        End If

                End If
        
                Call MakeObj(Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)

        End If

        TirarItemAlPiso = NuevaPos

        Exit Function
ErrHandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc, _
                           ByVal IsPretoriano As Boolean, _
                           ByVal UserIndex As Integer)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 28/11/2009
        'Give away npc's items.
        '28/11/2009: ZaMa - Implementado drops complejos
        '02/04/2010: ZaMa - Los pretos vuelven a tirar oro.
        '10/04/2011: ZaMa - Logueo los objetos logueables dropeados.
        '***************************************************
        On Error Resume Next

        With npc
        
                Dim i        As Byte

                Dim MiObj    As Obj

                Dim NroDrop  As Integer

                Dim Random   As Integer

                Dim objIndex As Integer
        
                ' Tira todo el inventario
                If IsPretoriano Then

                        For i = 1 To MAX_NORMAL_INVENTORY_SLOTS

                                If .Invent.Object(i).objIndex > 0 Then
                                        MiObj.Amount = .Invent.Object(i).Amount
                                        MiObj.objIndex = .Invent.Object(i).objIndex
                                        Call TirarItemAlPiso(.Pos, MiObj)

                                End If

                        Next i
            
                        ' Dropea oro?
                        If .GiveGLD > 0 Then _
                           Call TirarOroNpc(IIf(HappyHour, .GiveGLD * 2, .GiveGLD), .Pos, UserIndex)
                
                        Exit Sub

                End If
        
                Random = RandomNumber(1, 100)
        
                ' Tiene 10% de prob de no tirar nada
                If Random <= 90 Then
                        NroDrop = 1
            
                        If Random <= 10 Then
                                NroDrop = NroDrop + 1
                 
                                For i = 1 To 3

                                        ' 10% de ir pasando de etapas
                                        If RandomNumber(1, 100) <= 10 Then
                                                NroDrop = NroDrop + 1
                                        Else
                                                Exit For

                                        End If

                                Next i
                
                        End If

                        objIndex = .Drop(NroDrop).objIndex

                        If objIndex > 0 Then
            
                                If objIndex = iORO Then
                                        Call TirarOroNpc(.Drop(NroDrop).Amount, npc.Pos, UserIndex)
                                Else
                                        MiObj.Amount = .Drop(NroDrop).Amount
                                        MiObj.objIndex = objIndex
                    
                                        Call TirarItemAlPiso(.Pos, MiObj)
                    
                                        If ObjData(objIndex).Log = 1 Then
                                                Call LogItemsEspeciales(npc.Name & " drope� " & MiObj.Amount & " " & _
                                                   ObjData(objIndex).Name & "[" & objIndex & "]")

                                        End If
                    
                                End If

                        End If

                End If

        End With

End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal objIndex As Integer) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim i As Integer

        If Npclist(NpcIndex).Invent.NroItems > 0 Then

                For i = 1 To MAX_NORMAL_INVENTORY_SLOTS

                        If Npclist(NpcIndex).Invent.Object(i).objIndex = objIndex Then
                                QuedanItems = True
                                Exit Function

                        End If

                Next

        End If

        QuedanItems = False

End Function

''
' Gets the amount of a certain item that an npc has.
'
' @param npcIndex Specifies reference to npcmerchant
' @param ObjIndex Specifies reference to object
' @return   The amount of the item that the npc has
' @remarks This function reads the Npc.dat file
Function EncontrarCant(ByVal NpcIndex As Integer, ByVal objIndex As Integer) As Integer

        '***************************************************
        'Author: Unknown
        'Last Modification: 03/09/08
        'Last Modification By: Marco Vanotti (Marco)
        ' - 03/09/08 EncontrarCant now returns 0 if the npc doesn't have it (Marco)
        '***************************************************
        On Error Resume Next

        'Devuelve la cantidad original del obj de un npc

        Dim ln As String, npcfile As String

        Dim i  As Integer
    
        npcfile = DatPath & "NPCs.dat"
     
        For i = 1 To MAX_NORMAL_INVENTORY_SLOTS
                ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)

                If objIndex = val(ReadField(1, ln, 45)) Then
                        EncontrarCant = val(ReadField(2, ln, 45))
                        Exit Function

                End If

        Next
                       
        EncontrarCant = 0

End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        On Error Resume Next

        Dim i As Integer
    
        With Npclist(NpcIndex)
                .Invent.NroItems = 0
        
                For i = 1 To MAX_NORMAL_INVENTORY_SLOTS
                        .Invent.Object(i).objIndex = 0
                        .Invent.Object(i).Amount = 0
                Next i
        
                .InvReSpawn = 0

        End With

End Sub

''
' Removes a certain amount of items from a slot of an npc's inventory
'
' @param npcIndex Specifies reference to npcmerchant
' @param Slot Specifies reference to npc's inventory's slot
' @param antidad Specifies amount of items that will be removed
Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, _
                     ByVal Slot As Byte, _
                     ByVal Cantidad As Integer)

        '***************************************************
        'Author: Unknown
        'Last Modification: 23/11/2009
        'Last Modification By: Marco Vanotti (Marco)
        ' - 03/09/08 Now this sub checks that te npc has an item before respawning it (Marco)
        '23/11/2009: ZaMa - Optimizacion de codigo.
        '***************************************************
        Dim objIndex As Integer

        Dim iCant    As Integer
    
        With Npclist(NpcIndex)
                objIndex = .Invent.Object(Slot).objIndex
    
                'Quita un Obj
                If ObjData(.Invent.Object(Slot).objIndex).Crucial = 0 Then
                        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            
                        If .Invent.Object(Slot).Amount <= 0 Then
                                .Invent.NroItems = .Invent.NroItems - 1
                                .Invent.Object(Slot).objIndex = 0
                                .Invent.Object(Slot).Amount = 0

                                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                                        Call CargarInvent(NpcIndex) 'Reponemos el inventario

                                End If

                        End If

                Else
                        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            
                        If .Invent.Object(Slot).Amount <= 0 Then
                                .Invent.NroItems = .Invent.NroItems - 1
                                .Invent.Object(Slot).objIndex = 0
                                .Invent.Object(Slot).Amount = 0
                
                                If Not QuedanItems(NpcIndex, objIndex) Then
                                        'Check if the item is in the npc's dat.
                                        iCant = EncontrarCant(NpcIndex, objIndex)

                                        If iCant Then
                                                .Invent.Object(Slot).objIndex = objIndex
                                                .Invent.Object(Slot).Amount = iCant
                                                .Invent.NroItems = .Invent.NroItems + 1

                                        End If

                                End If
                
                                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                                        Call CargarInvent(NpcIndex) 'Reponemos el inventario

                                End If

                        End If

                End If

        End With

End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        'Vuelve a cargar el inventario del npc NpcIndex
        Dim LoopC   As Integer

        Dim ln      As String

        Dim npcfile As String
    
        npcfile = DatPath & "NPCs.dat"
    
        With Npclist(NpcIndex)
                .Invent.NroItems = val(GetVar(npcfile, "NPC" & .Numero, "NROITEMS"))
        
                For LoopC = 1 To .Invent.NroItems
                        ln = GetVar(npcfile, "NPC" & .Numero, "Obj" & LoopC)
                        .Invent.Object(LoopC).objIndex = val(ReadField(1, ln, 45))
                        .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            
                Next LoopC

        End With

End Sub

Public Sub TirarOroNpc(ByVal Cantidad As Long, _
                       ByRef Pos As WorldPos, _
                       ByVal UserIndex As Integer)
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 13/02/2010
        '***************************************************

        On Error GoTo ErrHandler
      
        With UserList(UserIndex)
        
                .Stats.GLD = (.Stats.GLD + Cantidad)

                If (.Stats.GLD > MAXORO) Then
                        .Stats.GLD = MAXORO

                End If

                Call SendData(SendTarget.ToPCArea, UserIndex, Protocol.PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Cantidad, 10))
                        
                Call WriteUpdateGold(UserIndex)

        End With

        Exit Sub

ErrHandler:
        Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)

End Sub
