Attribute VB_Name = "modSistemaComercio"
'*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************
'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************
Option Explicit

Enum eModoComercio

        Compra = 1
        Venta = 2

End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, _
                    ByVal userIndex As Integer, _
                    ByVal NpcIndex As Integer, _
                    ByVal Slot As Integer, _
                    ByVal Cantidad As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 07/06/2010
        '27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
        '  - 06/13/08 (NicoNZ)
        '07/06/2010: ZaMa - Los objetos se loguean si superan la cantidad de 1k (antes era solo si eran 1k).
        '*************************************************
        Dim Precio As Long
        Dim ValorCopas As Integer
        Dim Objeto As Obj
        
        On Error GoTo errhandleR

1        If Cantidad < 1 Or Slot < 1 Then Exit Sub
        
2        If Cantidad > 10000 Then Cantidad = 10000
        
3        If Modo = eModoComercio.Compra Then
4                If Slot > MAX_NORMAL_INVENTORY_SLOTS Then
5                        Exit Sub
6                ElseIf Cantidad > MAX_INVENTORY_OBJS Then
7                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(userIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
8                        Call Ban(UserList(userIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados �tems:" & Cantidad)
9                        UserList(userIndex).flags.Ban = 1
10                        Call WriteErrorMsg(userIndex, "Has sido baneado por el Sistema AntiCheat.")
11                        Call FlushBuffer(userIndex)
12                        Call CloseSocket(userIndex)
13                        Exit Sub
14                ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).Amount > 0 Then
                        Exit Sub

                End If

15                If Cantidad > Npclist(NpcIndex).Invent.Object(Slot).Amount Then Cantidad = Npclist(UserList(userIndex).flags.TargetNPC).Invent.Object(Slot).Amount
16                Objeto.Amount = Cantidad
17                Objeto.objIndex = Npclist(NpcIndex).Invent.Object(Slot).objIndex
                'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
                'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
18                Precio = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).objIndex).Valor / Descuento(userIndex) * Cantidad) + 0.5)
19
                
                ValorCopas = ObjData(Npclist(NpcIndex).Invent.Object(Slot).objIndex).copas
                
20                 If ValorCopas > 0 Then
21                    Cantidad = 1
22                    Objeto.Amount = 1
23                End If
                
24                If UserList(userIndex).Stats.GLD < Precio Then
25                        Call WriteConsoleMsg(userIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
26                        Exit Sub
                End If
                
27                If ValorCopas > 0 Then
28                    Cantidad = 1
29                    Objeto.Amount = 1
                End If
                
30                If Not TieneObjetos(710, ValorCopas * Cantidad, userIndex) Then
31                    Call WriteConsoleMsg(userIndex, "No tienes copas suficientes", FontTypeNames.FONTTYPE_GM)
32                    Exit Sub
                End If
                
33                If MeterItemEnInventario(userIndex, Objeto) = False Then
34                        'Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
35                        Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
36                        Call WriteTradeOK(userIndex)
37                        Exit Sub

                End If

38                UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD - Precio
               
                
39                If ValorCopas > 0 Then
40                    Call QuitarObjetos(710, ValorCopas * Cantidad, userIndex)
41                    LogDesarrollo UserList(userIndex).Name & " compr� por Copas: " & ValorCopas * Cantidad
42                    Call UpdateUserInv(True, userIndex, 0)
43                End If
                
44                Call QuitarNpcInvItem(UserList(userIndex).flags.TargetNPC, CByte(Slot), Cantidad)

                'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
                'Es un Objeto que tenemos que loguear?
46                If ObjData(Objeto.objIndex).Log = 1 Then
45                        Call LogItemsEspeciales(UserList(userIndex).Name & " compr� del NPC " & Objeto.Amount & " " & ObjData(Objeto.objIndex).Name)
              '  ElseIf Objeto.Amount >= 1000 Then 'Es mucha cantidad?

                        'Si no es de los prohibidos de loguear, lo logueamos.
               '         If ObjData(Objeto.objIndex).NoLog <> 1 Then
                '                Call LogItemsEspeciales(UserList(UserIndex).Name & " compr� del NPC " & Objeto.Amount & " " & ObjData(Objeto.objIndex).Name)

                 '       End If

                End If

                'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
47                If ObjData(Objeto.objIndex).OBJType = otLlaves Then
48                        Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & Slot, Objeto.objIndex & "-0")
49                        Call logVentaCasa(UserList(userIndex).Name & " compr� " & ObjData(Objeto.objIndex).Name)

                End If

        ElseIf Modo = eModoComercio.Venta Then
                    
                If Cantidad < 0 Then Exit Sub
                
51                If Cantidad > UserList(userIndex).Invent.Object(Slot).Amount Then Cantidad = UserList(userIndex).Invent.Object(Slot).Amount
50                Objeto.Amount = Cantidad
52                Objeto.objIndex = UserList(userIndex).Invent.Object(Slot).objIndex

53                If Objeto.objIndex = 0 Then
                        Exit Sub
57                ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.objIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.objIndex = iORO Then
54                        Call WriteConsoleMsg(userIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
55                        Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
56                        Call WriteTradeOK(userIndex)
                        Exit Sub
58                ElseIf ObjData(Objeto.objIndex).Real = 1 Then

59                        If Npclist(NpcIndex).Name <> "SR" Then
                                Call WriteConsoleMsg(userIndex, "Las armaduras del ej�rcito real s�lo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                                Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
                                Call WriteTradeOK(userIndex)
                                Exit Sub

                        End If

                ElseIf ObjData(Objeto.objIndex).Caos = 1 Then

                        If Npclist(NpcIndex).Name <> "SC" Then
                                Call WriteConsoleMsg(userIndex, "Las armaduras de la legi�n oscura s�lo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                                Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
                                Call WriteTradeOK(userIndex)
                                Exit Sub

                        End If

                ElseIf ItemShop(Objeto.objIndex) = True Then
                        Call WriteConsoleMsg(userIndex, "Este item es una reliquia, no deberias venderlo!", FontTypeNames.FONTTYPE_INFO)
                        Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
                        Call WriteTradeOK(userIndex)

                        Exit Sub
                ElseIf UserList(userIndex).Invent.Object(Slot).Amount < 0 Or Cantidad = 0 Then
                        Exit Sub
                ElseIf Slot < LBound(UserList(userIndex).Invent.Object()) Or Slot > UBound(UserList(userIndex).Invent.Object()) Then
                        Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
                        Exit Sub
                ElseIf UserList(userIndex).flags.Privilegios And PlayerType.Consejero Then
                        Call WriteConsoleMsg(userIndex, "No puedes vender �tems.", FontTypeNames.FONTTYPE_WARNING)
                        Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
                        Call WriteTradeOK(userIndex)
                        Exit Sub

                End If

66                Call QuitarUserInvItem(userIndex, Slot, Cantidad)
                'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
67                Precio = Fix(SalePrice(Objeto.objIndex) * Cantidad)
68                UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD + Precio

69                If UserList(userIndex).Stats.GLD > MAXORO Then _
                   UserList(userIndex).Stats.GLD = MAXORO
                Dim NpcSlot As Integer
101                NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.objIndex, Objeto.Amount)

102                If NpcSlot <= MAX_NORMAL_INVENTORY_SLOTS Then 'Slot valido
                        'Mete el obj en el slot
103                        Npclist(NpcIndex).Invent.Object(NpcSlot).objIndex = Objeto.objIndex
104                        Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount

105                        If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
106                                Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS

                        End If

                End If

                'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
                'Es un Objeto que tenemos que loguear?
107                If ObjData(Objeto.objIndex).Log = 1 Then
108                        Call LogItemsEspeciales(UserList(userIndex).Name & " vendi� al NPC " & Objeto.Amount & " " & ObjData(Objeto.objIndex).Name)
109                ElseIf Objeto.Amount >= 1000 Then 'Es mucha cantidad?

                        'Si no es de los prohibidos de loguear, lo logueamos.
1000                        If ObjData(Objeto.objIndex).NoLog <> 1 Then
1001                                Call LogItemsEspeciales(UserList(userIndex).Name & " vendi� al NPC " & Objeto.Amount & " " & ObjData(Objeto.objIndex).Name)

                        End If

                End If

        End If

1002        Call UpdateUserInv(True, userIndex, 0)
1003        Call WriteUpdateUserStats(userIndex)
1004        Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
1005        Call WriteTradeOK(userIndex)
1006        Call SubirSkill(userIndex, eSkill.Comerciar)

Exit Sub
errhandleR:
LogError "Error en Comercio en linea " & Erl & ". Err " & Err.Number & " " & Err.description


End Sub

Public Sub IniciarComercioNPC(ByVal userIndex As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        Call EnviarNpcInv(userIndex, UserList(userIndex).flags.TargetNPC)
        UserList(userIndex).flags.Comerciando = True
        Call WriteCommerceInit(userIndex)

End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, _
                              ByVal Objeto As Integer, _
                              ByVal Cantidad As Integer) As Integer
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        SlotEnNPCInv = 1

        Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).objIndex = Objeto _
           And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS
                SlotEnNPCInv = SlotEnNPCInv + 1

                If SlotEnNPCInv > MAX_NORMAL_INVENTORY_SLOTS Then Exit Do
        Loop

        If SlotEnNPCInv > MAX_NORMAL_INVENTORY_SLOTS Then
                SlotEnNPCInv = 1

                Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).objIndex = 0
                        SlotEnNPCInv = SlotEnNPCInv + 1

                        If SlotEnNPCInv > MAX_NORMAL_INVENTORY_SLOTS Then Exit Do
                Loop

                If SlotEnNPCInv <= MAX_NORMAL_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1

        End If

End Function

Private Function Descuento(ByVal userIndex As Integer) As Single
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        Descuento = 1 + UserList(userIndex).Stats.UserSkills(eSkill.Comerciar) / 100

End Function

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC
Private Sub EnviarNpcInv(ByVal userIndex As Integer, ByVal NpcIndex As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last Modified: 06/14/08
        'Last Modified By: Nicol�s Ezequiel Bouhid (NicoNZ)
        '*************************************************


    Dim UserName As String
    Dim NPCNAME As String
    On Error Resume Next
31  UserName = UserList(userIndex).Name
30  NPCNAME = Npclist(NpcIndex).Name

    If Err.Number > 0 Then Err.Clear
    
    If NpcIndex = 0 Then LogError "NpcIndex=0 en EnviarNpcInv": Exit Sub
    
On Error GoTo errhandleR

        Dim Slot As Byte
        Dim val  As Single

1        For Slot = 1 To MAX_NORMAL_INVENTORY_SLOTS
                
2                If Npclist(NpcIndex).Invent.Object(Slot).objIndex > 0 Then
3                        Dim thisObj As Obj
4                        thisObj.objIndex = Npclist(NpcIndex).Invent.Object(Slot).objIndex
5                        thisObj.Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount
6                        val = (ObjData(thisObj.objIndex).Valor) / Descuento(userIndex)
7                        Call WriteChangeNPCInventorySlot(userIndex, Slot, thisObj, val)
8                Else
9                        Dim DummyObj As Obj
10                        Call WriteChangeNPCInventorySlot(userIndex, Slot, DummyObj, 0)

                End If

11        Next Slot
  
  Exit Sub
  
errhandleR:

LogError "error en EnviarNpcInv en " & Erl & ". Err " & Err.Number & " " & Err.description & "... Nombre:" & UserName & " - Npc: " & NPCNAME
Resume Next


End Sub

''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El n�mero de objeto al cual le calculamos el precio de venta
Public Function SalePrice(ByVal objIndex As Integer) As Single

        '*************************************************
        'Author: Nicol�s (NicoNZ)
        '
        '*************************************************
        If objIndex < 1 Or objIndex > UBound(ObjData) Then Exit Function
        If ItemNewbie(objIndex) Then Exit Function
        SalePrice = ObjData(objIndex).Valor / REDUCTOR_PRECIOVENTA

End Function
