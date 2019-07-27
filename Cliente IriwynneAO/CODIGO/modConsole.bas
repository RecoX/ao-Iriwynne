Attribute VB_Name = "modConsole"
Option Explicit

Private Type NMHDR

        hWndFrom As Long
        idFrom As Long
        code As Long

End Type

Private Type CHARRANGE

        cpMin As Long
        cpMax As Long

End Type

Private Type ENLINK

        hdr As NMHDR
        msg As Long
        wParam As Long
        lParam As Long
        chrg As CHARRANGE

End Type

Private Type TEXTRANGE

        chrg As CHARRANGE
        lpstrText As String

End Type

Private Declare Function SetWindowLong _
                Lib "USER32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc _
                Lib "USER32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hwnd As Long, _
                                         ByVal msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Declare Function SendMessage _
                Lib "USER32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Private Declare Function ShellExecute _
                Lib "shell32" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Private Const WM_NOTIFY = &H4E

Private Const EM_SETEVENTMASK = &H445

Private Const EM_GETEVENTMASK = &H43B

Private Const EM_GETTEXTRANGE = &H44B

Private Const EM_AUTOURLDETECT = &H45B

Private Const EN_LINK = &H70B

Private Const WM_LBUTTONDOWN = &H201

Private Const ENM_LINK = &H4000000

Private Const GWL_WNDPROC = (-4)

Private Const SW_SHOW = 5

Private lOldProc   As Long

Private hWndRTB    As Long

Private hWndParent As Long

Public Sub DrawInvisibleChar(ByVal CharIndex As Integer, _
                             ByVal PixelOffsetX As Long, _
                             ByVal PixelOffsetY As Long)

        Dim line  As String

        Dim color As Long

        Dim Pos   As Integer

        With charlist(CharIndex)

                If ClanTag(CharIndex) And (esGM(CharIndex) = False Or esGM(UserCharIndex) = False) Then
           
                        'Draw Body
                        If .Body.Walk(.Heading).GrhIndex Then _
                           Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, 0)
        
                        'Draw Head
                        If .Head.Head(.Heading).GrhIndex Then
                                Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                
                                'Draw Helmet
                                If .Casco.Head(.Heading).GrhIndex Then _
                                   Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 0)
                
                                'Draw Weapon
                                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                                   Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, 0)
                
                                'Draw Shield
                                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                                   Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, 0)

                        End If

                        'Draw name over head
                        If LenB(.Nombre) > 0 Then
                                        
                                If Nombres Then
                                                                        
                                        Pos = getTagPosition(.Nombre)
                                        line = Left$(.Nombre, Pos - 2)
                                        color = RGB(255, 255, 255)
                                                        
                                        Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, color, frmMain.font)
                                
                                        'Clan
                                        line = mid$(.Nombre, Pos)
                                    
                                        If .Conquisto Then
                                            Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 7, PixelOffsetY + 45, line, RGB(255, 255, 0), frmScreenshots.font)
                                        Else
                                            Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 7, PixelOffsetY + 45, line, color, frmScreenshots.font)
                                        End If
                                End If

                        End If
    
                End If

        End With

End Sub

Public Sub EnableURLDetect(ByVal hWndRichTextbox As Long, ByVal hWndOwner As Long)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 13/12/2012
        'Enables url detection in richtexbox.
        '***************************************************
        SendMessage hWndRichTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndRichTextbox, EM_GETEVENTMASK, 0, 0)
        SendMessage hWndRichTextbox, EM_AUTOURLDETECT, 1, ByVal 0
    
        hWndParent = hWndOwner
        hWndRTB = hWndRichTextbox

End Sub

Public Sub DisableURLDetect()

        '***************************************************
        'Author: ZaMa
        'Last Modification: 13/12/2012
        'Disables url detection in richtexbox.
        '***************************************************
 
        SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0

        StopCheckingLinks

End Sub

Public Sub StartCheckingLinks()

        '***************************************************
        'Author: ZaMa
        'Last Modification: 18/11/2010
        'Starts checking links (in console range)
        '***************************************************
        If lOldProc = 0 Then
                lOldProc = SetWindowLong(hWndParent, GWL_WNDPROC, AddressOf WndProc)

        End If

End Sub

Public Sub StopCheckingLinks()

        '***************************************************
        'Author: ZaMa
        'Last Modification: 18/11/2010
        'Stops checking links (out of console range)
        '***************************************************
        If lOldProc Then
                SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
                lOldProc = 0

        End If

End Sub

Public Function WndProc(ByVal hwnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

        '***************************************************
        'Author: ZaMa
        'Last Modification: 13/02/2012
        'Get "Click" event on link and open browser.
        '***************************************************
        Dim uHead As NMHDR

        Dim eLink As ENLINK

        Dim eText As TEXTRANGE

        Dim sText As String

        Dim lLen  As Long
    
        If uMsg = WM_NOTIFY Then
                CopyMemory uHead, ByVal lParam, Len(uHead)

                If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then
                    
                        CopyMemory eLink, ByVal lParam, Len(eLink)
            
                        Select Case eLink.msg

                                Case WM_LBUTTONDOWN
                                        eText.chrg.cpMin = eLink.chrg.cpMin
                                        eText.chrg.cpMax = eLink.chrg.cpMax
                                        eText.lpstrText = Space$(1024)
                    
                                        lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)

                                        sText = Left$(eText.lpstrText, lLen)
                                        ShellExecute hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW

                        End Select

                End If

        End If
    
        WndProc = CallWindowProc(lOldProc, hwnd, uMsg, wParam, lParam)

End Function

