VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
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
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private MouseDownSelectedItem      As Long

Private InventoryOffset            As Long             'Number of lines we scrolled down from topmost

Private InvSelectedItem            As Long             'Currently selected item

Private ownTilePixelWidth          As Integer

Private ownTilePixelHeight         As Integer

Private startdX                    As Integer

Private startdY                    As Integer

Private ShowValue                  As Boolean

Private ValueFormated              As Boolean

Private ImgContainer               As Boolean

Private InvSurface                 As DirectDrawSurface7            'DD Surface used to render everything

Private bufferInvSurface           As DirectDrawSurface7

Private Inventory()                As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Private Sub Class_Initialize()

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        '
        '***************************************************
End Sub

Private Sub Class_Terminate()
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        '
        '***************************************************
        Set InvSurface = Nothing
        Set bufferInvSurface = Nothing

End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 14/01/2010
        'Scrolls the graphical inventory up / down
        '14/01/2010: ZaMa - Now redraws inventory only if needed
        '***************************************************

        'Only move if there is something to show up /down
        If CanScroll(Up) Then
                InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
                'Redraw the inventory if needed
                DrawInventory

        End If

End Sub

Public Function CanScroll(ByVal Up As Boolean)

        '***************************************************
        'Author: ZaMa
        'Last Modify Date: 14/01/2010
        'Determines wether inventory can be scrolled up / down
        '***************************************************
        If Up Then
                CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * _
                   (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
        Else
                CanScroll = InventoryOffset > 0

        End If
    
End Function

Public Sub SelectGold()

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Sets the gold as the currently selected item
        '***************************************************
        Dim prevSelectedItem As Long
    
        'Store preivously selected item
        prevSelectedItem = InvSelectedItem
    
        'Select the gold
        InvSelectedItem = FLAGORO
    
        'Redraw old item to deselect it
        If prevSelectedItem <> FLAGORO Then
                Call DrawInventorySlot(prevSelectedItem)

        End If

End Sub

Public Sub DeselectItem()

        '***************************************************
        'Author: ZaMa
        'Last Modify Date: 30/11/2009
        'Deselect the currently selected item
        '***************************************************
        If InvSelectedItem = 0 Then Exit Sub
    
        Dim ClearSlot As Byte
    
        ClearSlot = InvSelectedItem
    
        'Select nothing
        InvSelectedItem = 0
    
        'Redraw old item to deselect it
        Call DrawInventorySlot(ClearSlot)

End Sub

Public Sub ChangeSlotItemAmount(ByVal Slot As Byte, ByVal Amount As Long)
        '***************************************************
        'Author: ZaMa
        'Created Date: 30/11/2009
        'Last Modify Date: 09/12/2009 (Budi)
        'Change the amount of the required item, and update the slot.
        '***************************************************
    
        Inventory(Slot).Amount = Amount
    
        'Update Amount
        Call DrawInventorySlot(Slot)

End Sub

Public Property Get SelectedItem() As Long
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the selected item index
        '***************************************************
        SelectedItem = InvSelectedItem

End Property

Public Property Get MaxHit(ByVal Slot As Byte) As Integer
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the max hit of the selected item
        '***************************************************
        MaxHit = Inventory(Slot).MaxHit

End Property

Public Property Get MinHit(ByVal Slot As Byte) As Integer
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the min hit of the selected item
        '***************************************************
        MinHit = Inventory(Slot).MinHit

End Property

Public Property Get MaxDef(ByVal Slot As Byte) As Integer
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the defense of the selected item
        '***************************************************
        MaxDef = Inventory(Slot).MaxDef

End Property

Public Property Get MinDef(ByVal Slot As Byte) As Integer
        '***************************************************
        'Author: Budi
        'Last Modify Date: 02/1209
        'Retrieves the defense of the selected item
        '***************************************************
        MinDef = Inventory(Slot).MinDef

End Property

Public Property Get GrhIndex(ByVal Slot As Byte) As Integer
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the grh index of the selected item
        '***************************************************
        GrhIndex = Inventory(Slot).GrhIndex

End Property

Public Property Get Valor(ByVal Slot As Byte) As Single
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the value of the selected item
        '***************************************************
        Valor = Inventory(Slot).Valor

End Property

Public Property Get Amount(ByVal Slot As Byte) As Long

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the selected item's amount
        '***************************************************
        If Slot = FLAGORO Then
                Amount = UserGLD
        ElseIf Slot >= LBound(Inventory) And Slot <= UBound(Inventory) Then
                Amount = Inventory(Slot).Amount

        End If

End Property

Public Property Get ObjIndex(ByVal Slot As Byte) As Integer
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the selected item's object index
        '***************************************************
        ObjIndex = Inventory(Slot).ObjIndex

End Property

Public Property Get ObjType(ByVal Slot As Byte) As Integer
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the selected item's object type
        '***************************************************
        ObjType = Inventory(Slot).ObjType

End Property

Public Property Get ItemName(ByVal Slot As Byte) As String
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves the selected item's name
        '***************************************************
        ItemName = Inventory(Slot).Name

End Property

Public Property Get Equipped(ByVal Slot As Byte) As Boolean
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Retrieves True if the item at the given pos is eqiupped
        '***************************************************
        Equipped = Inventory(Slot).Equipped

End Property

Public Property Get MaxObjs() As Byte
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modify Date: 09/16/09
        'Retrieves the capacity of the Inventory
        '***************************************************
        MaxObjs = UBound(Inventory)

End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)

        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modify Date: 09/16/09
        'Set the capacity of the Inventary
        '***************************************************
        If UBound(Inventory) = MaxObjs Then Exit Sub
    
        ReDim Preserve Inventory(1 To MaxObjs) As Inventory

End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
        '***************************************************
        'Author: ZaMa
        'Last Modify Date: 11/12
        'Indicates if the amount of the items in the inventory is shown
        '***************************************************
        ShowValue = bShow
    
        Dim Slot As Integer

        ' Update Inventory
        For Slot = 1 To Me.MaxObjs
                DrawInventorySlot (Slot)
        Next Slot

End Sub

Public Sub ClearAllSlots()
        '***************************************************
        'Author: ZaMa
        'Last Modify Date: 15/03/2011
        'Clear All Slots.
        '***************************************************

        Dim Slot As Long

        For Slot = 1 To Me.MaxObjs
                Call SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)
        Next Slot

End Sub

Public Sub SetItem(ByVal Slot As Byte, _
                   ByVal eOBJIndex As Integer, _
                   ByVal eAmount As Long, _
                   ByVal eEquipped As Byte, _
                   ByVal eGrhIndex As Integer, _
                   ByVal eObjType As Integer, _
                   ByVal eMaxHit As Integer, _
                   ByVal eMinHit As Integer, _
                   ByVal MaxDef As Integer, _
                   ByVal MinDef As Integer, _
                   ByVal eValor As Single, _
                   ByVal eName As String)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 12/04/06
        'Sets all data for a given inventory slot
        '***************************************************
        If Slot < 1 Or Slot > UBound(Inventory) Then Exit Sub
    
        With Inventory(Slot)
                .Amount = eAmount
                .MaxDef = MaxDef
                .MinDef = MinDef
                .Equipped = eEquipped
                .GrhIndex = eGrhIndex
                .MaxHit = eMaxHit
                .MinHit = eMinHit
                .Name = eName
                .ObjIndex = eOBJIndex
                .ObjType = eObjType
                .Valor = eValor

        End With
    
        'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
        'Render inventory slot (we don't need to render the whole inventory)
        Call DrawInventorySlot(Slot)

End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Selects the item clicked if it's valid and return's it's index
        '***************************************************
        Dim TempItem As Long

        Dim temp_x   As Long

        Dim temp_y   As Long
    
        temp_x = X \ ownTilePixelWidth
        temp_y = Y \ ownTilePixelHeight
    
        TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
        'Make sure it's within limits
        If TempItem <= UBound(Inventory) Then
                ClickItem = TempItem

        End If

End Function

Private Sub DrawInventory()

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Renders the inventory to the given PictureBox
        '***************************************************
        Dim loopc    As Long

        Dim SrcRect  As RECT

        Dim TempRect As RECT

        Dim destRect As RECT
    
        With TempRect
                .Bottom = InventoryWindow.Height
                .Right = InventoryWindow.Width

        End With
    
        With destRect
                .Bottom = ownTilePixelHeight
                .Right = ownTilePixelWidth

        End With
    
        'Clear the inventory window
        InventoryWindow.Cls
    
        'Clear the slot area
        Call InvSurface.BltColorFill(TempRect, InventoryWindow.BackColor)
    
        For loopc = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
                'For loopc = 1 To UBound(Inventory)
    
                If Inventory(loopc).GrhIndex Then

                        'Get source rect
                        With SrcRect
                                .Left = GrhData(Inventory(loopc).GrhIndex).sX
                                .Top = GrhData(Inventory(loopc).GrhIndex).sY
                                .Right = .Left + ownTilePixelWidth
                                .Bottom = .Top + ownTilePixelHeight

                        End With
            
                        'Render the item grh and the amount
                        InvSurface.SetForeColor vbWhite
                        InvSurface.BltFast destRect.Left + startdX, destRect.Top + startdY, SurfaceDB.Surface(GrhData(Inventory(loopc).GrhIndex).FileNum), SrcRect, DDBLTFAST_WAIT
                        InvSurface.DrawText destRect.Left + startdX / 2, destRect.Top + startdY / 2, Inventory(loopc).Amount, False
            
                        'If equipped we render "E"
                        If Inventory(loopc).Equipped Then
                                InvSurface.SetForeColor vbCyan
                                InvSurface.DrawRoundedBox 1, 1, 32, 32, RGB(255, 255, 255), RGB(255, 255, 255)
                                'InvSurface.SetForeColor vbYellow
                                InvSurface.DrawText destRect.Left + 20, destRect.Top + 20, "E", False
                                
                        End If
            
                        'Render a box around the selected item
                        If InvSelectedItem = loopc Then
                                If ImgContainer Then
                                        InvSurface.SetForeColor vbGreen
                                        InvSurface.setDrawStyle DrawStyleConstants.vbDot
                                        InvSurface.DrawBox destRect.Left, destRect.Top, destRect.Left + ownTilePixelWidth, destRect.Top + ownTilePixelHeight

                                End If

                        End If

                End If
        
                'Render the item to the Inventory Window
                With destRect

                        'Check if we have reached the right end of the inventory picturebox
                        If .Left = (InventoryWindow.ScaleWidth - (InventoryWindow.ScaleWidth Mod ownTilePixelWidth) - ownTilePixelWidth) Then
                                .Left = 0
                                .Right = ownTilePixelWidth
                                .Bottom = .Bottom + ownTilePixelHeight
                                .Top = .Top + ownTilePixelHeight
                        Else
                                .Left = .Left + ownTilePixelWidth
                                .Right = .Right + ownTilePixelWidth

                        End If

                End With

        Next loopc

        Call bufferInvSurface.BltFast(0, 0, InvSurface, TempRect, DDBLTFAST_WAIT)
        Call InvSurface.BltToDC(InventoryWindow.hdc, TempRect, TempRect)
    
        Call InventoryWindow.Refresh

End Sub

Private Sub DrawInventorySlot(ByVal Slot As Integer)

        Dim destRect      As RECT
        Dim SrcRectSlot   As RECT
        Dim upperLeftSlot As Integer

        upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
        'If not in renderable area we exit
        If Slot < upperLeftSlot Then Exit Sub
    
        With destRect
                .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth \ ownTilePixelWidth)) * ownTilePixelHeight
                .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth \ ownTilePixelWidth)) * ownTilePixelWidth
                .Bottom = .Top + ownTilePixelHeight
                .Right = .Left + ownTilePixelWidth
        End With
    
        'Clear the slot area
        Call InvSurface.BltColorFill(destRect, InventoryWindow.BackColor)
    
        If Inventory(Slot).GrhIndex Then

                'Get source rect
                With SrcRectSlot
                        .Left = GrhData(Inventory(Slot).GrhIndex).sX
                        .Top = GrhData(Inventory(Slot).GrhIndex).sY
                        .Right = .Left + TilePixelWidth
                        .Bottom = .Top + TilePixelHeight

                End With
        
                'Render the item grh and the amount
                InvSurface.SetForeColor vbWhite
                InvSurface.BltFast destRect.Left + startdX, destRect.Top + startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(Slot).GrhIndex).FileNum)), SrcRectSlot, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
        
                If ShowValue Then

                        Dim sValue As String

                        sValue = IIf(ValueFormated, Format(Inventory(Slot).Amount, "#,###,###,###"), Inventory(Slot).Amount)
                        '     sValue = Inventory(slot).Amount
                        InvSurface.DrawText destRect.Left + startdX / 2, destRect.Top + startdY / 2, sValue, False

                End If
        
                'If equipped we render "E"
                If Inventory(Slot).Equipped Then
                        InvSurface.SetForeColor vbYellow
                        InvSurface.DrawText destRect.Left + 20, destRect.Top + 20, "E", False

                End If
        
                'Render a box around the selected item
                If InvSelectedItem = Slot Then
                        If ImgContainer Then
                                InvSurface.SetForeColor vbGreen
                                InvSurface.setDrawStyle DrawStyleConstants.vbDot
                                InvSurface.DrawBox destRect.Left, destRect.Top, destRect.Left + ownTilePixelWidth, destRect.Top + ownTilePixelHeight

                        End If

                End If

        End If
    
        'Render the item to the Inventory Window
        InvSurface.BltToDC InventoryWindow.hdc, destRect, destRect
    
        Call InventoryWindow.Refresh

End Sub

Public Sub Initialize(ByRef DirectDraw As DirectDraw7, _
                      ByRef InvPic As PictureBox, _
                      ByVal MaxObjs As Byte, _
                      Optional ByVal FontSize As Integer = 7, _
                      Optional ByVal TileWidth As Integer = 32, _
                      Optional ByVal TileHeight As Integer = 32, _
                      Optional ByVal startX As Integer = 0, _
                      Optional ByVal startY As Integer = 0, _
                      Optional ByVal bImgContainer As Boolean = True, _
                      Optional ByVal bShowText As Boolean = True, _
                      Optional ByVal FormatValue As Boolean = False)
    
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 03/12/09
        'Sets the reference to the picture box on which to render the inventory
        '03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
        '***************************************************

        Dim ddck As DDCOLORKEY

        Set InventoryWindow = InvPic
        ValueFormated = FormatValue
    
        'Make sure auto-redraw is set to true
        InventoryWindow.AutoRedraw = True
    
        'Set apropiate scale (pixel)
        InventoryWindow.ScaleMode = 3
    
        ReDim Inventory(1 To MaxObjs) As Inventory
    
        'initialize DX stuff
        Dim SurfaceDesc As DDSURFACEDESC2
    
        'Make sure DirectDraw was correctly initialized
        If DirectDraw Is Nothing Then Exit Sub
            
        'Set surface's description
        With SurfaceDesc
                .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
                .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
                .lHeight = InventoryWindow.Height
                .lWidth = InventoryWindow.Width

        End With
    
        'Create inventory surface
        Set InvSurface = DirectDraw.CreateSurface(SurfaceDesc)
        Set bufferInvSurface = DirectDraw.CreateSurface(SurfaceDesc)
    
        ownTilePixelHeight = TileHeight
        ownTilePixelWidth = TileWidth
        startdX = startX
        startdY = startY
    
        ImgContainer = bImgContainer
        ShowValue = bShowText
   
        'Prepare surfaces for text rendering
        InvSurface.SetFontTransparency True

        'TODO : Fonts should be in a separate class / collection
        Dim font As StdFont

        Set font = New StdFont

        Dim Ifnt As IFont
    
        font.Name = "Verdana"
        font.Bold = True
        font.Italic = False
        font.size = FontSize
        font.Underline = False
        font.Strikethrough = False
    
        Set Ifnt = font
    
        InvSurface.SetFont Ifnt
        bufferInvSurface.SetFont Ifnt
        '\TODO
    
        If ImgContainer Then _
           InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Implements the mouse move event of the inventory picture box
        'Displays a ToolTip of the item under the mouse
        '***************************************************
        Dim temp_x   As Integer

        Dim temp_y   As Integer

        Dim TempItem As Integer
    
        'Exit if it got outside the control's areanv
        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
           Exit Sub
  
        temp_x = X \ ownTilePixelWidth
        temp_y = Y \ ownTilePixelHeight
        
        TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
        If TempItem <= UBound(Inventory) Then
                InventoryWindow.ToolTipText = Inventory(TempItem).Name

        End If
    
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)

        Dim TempItem As Long
        
        'Exit if it got outside the control's area
        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
                
        If (Button = vbRightButton) Then
                TempItem = ClickItem(X, Y)

                If (Inventory(TempItem).GrhIndex) Then
                        MouseDownSelectedItem = TempItem

                        Exit Sub

                End If

        End If
   
        MouseDownSelectedItem = 0
        
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 27/07/04
        'Implements the mouse up event of the inventory picture box
        'Check outs which item was clicked
        '***************************************************
        'Store previously selected item
        Dim prevSelItem As Long

        Dim TempItem    As Long

        Dim temp_x      As Long

        Dim temp_y      As Long
    
        'Exit if it got outside the control's area
        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
           Exit Sub
    
        prevSelItem = InvSelectedItem
    
        'Get the currently clickced item

        If Not (frmMain.UsabaDrag) Then
                InvSelectedItem = ClickItem(CInt(X), CInt(Y))

        End If

        If (frmMain.UsandoDrag) And (InvSelectedItem <= MAX_INVENTORY_SLOTS) And (InvSelectedItem >= 0) And (Not frmMain.UsabaDrag) Then
                If (InvSelectedItem <> MouseDownSelectedItem) And (MouseDownSelectedItem <> 0) And (InvSelectedItem <> 0) Then
                        
                        If UserEstado = 0 Then
                        Call WriteMoveItem(InvSelectedItem, MouseDownSelectedItem, eMoveType.Inventory)
                        End If
                        
                        frmMain.UsandoDrag = False
                Else
                        frmMain.UsandoDrag = False

                End If

        End If
   
        'Update needed inventory slots
        If prevSelItem <> InvSelectedItem Then
                If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call DrawInventorySlot(prevSelItem)
                If InvSelectedItem Then Call DrawInventorySlot(InvSelectedItem)

        End If

End Sub

