#include-once
#AutoIt3Wrapper_Au3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

; #INDEX# =======================================================================================================================
; Title .........: GUIDarkMenu UDF Library for AutoIt3
; AutoIt Version : 3.3.18.0
; Language ......: English
; Description ...: UDF library for applying dark theme to menubar
; Author(s) .....: WildByDesign (including previous code from ahmet, argumentum, UEZ)
; Version .......: 0.9.2
; ===============================================================================================================================

; Windows messages used by GUIDarkMenu:
; $WM_DRAWITEM
; $WM_MEASUREITEM
; $WM_WINDOWPOSCHANGED
; $WM_ACTIVATE

#include <GuiMenu.au3>
#include <GUIConstantsEx.au3>
#include <APIGdiConstants.au3>
#include <WindowsNotifsConstants.au3>
#include <Array.au3>

#include "GUIDarkInternal.au3"

; Menu Info
Const $ODT_MENU = 1
Const $ODS_SELECTED = 0x0001
Const $ODS_DISABLED = 0x0004
Const $ODS_HOTLIGHT = 0x0040

; DPI
Global $iDPIpct = 100

; Dark Mode Colors (RGB)
Global $COLOR_BG_DARK = 0x121212
Global $COLOR_TEXT_LIGHT = 0xE0E0E0
Global $COLOR_CONTROL_BG = 0x202020
Global $COLOR_BORDER = 0x3F3F3F
Global $COLOR_MENU_BG = __WinAPI_ColorAdjustLuma($COLOR_BG_DARK, 5)
Global $COLOR_MENU_HOT = __WinAPI_ColorAdjustLuma($COLOR_MENU_BG, 20)
Global $COLOR_MENU_SEL = __WinAPI_ColorAdjustLuma($COLOR_MENU_BG, 10)
Global $COLOR_MENU_TEXT = $COLOR_TEXT_LIGHT

; Store handle for GUI that called menu functions
Global $hGUI

;GUIRegisterMsg($WM_DRAWITEM, "WM_DRAWITEM")
GUIRegisterMsg($WM_MEASUREITEM, "WM_MEASUREITEM")

;********************************************************************
; WM_MEASURE procedure
;********************************************************************
Func WM_MEASUREITEM($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam
    Local Const $DEFAULT_GUI_FONT = 17
    Local $tagMEASUREITEM = "uint CtlType;uint CtlID;uint itemID;uint itemWidth;uint itemHeight;ulong_ptr itemData"
    Local $t = DllStructCreate($tagMEASUREITEM, $lParam)
    If Not IsDllStruct($t) Then Return $GUI_RUNDEFMSG

    If $t.CtlType <> $ODT_MENU Then Return $GUI_RUNDEFMSG

    Local $itemID = $t.itemID

    Local $sText = _GUICtrlMenu_GetItemText(_GUICtrlMenu_GetMenu($hWnd), $itemID, False)
	Local $arSize = _GetTextDimension($hWnd, $sText)

    ; Set dimensions with padding (with high DPI)
	$t.itemWidth = _CalcMenuItemWidth($iDPIpct, $arSize[0])
    $t.itemHeight = $arSize[1] + 1

    Return 1
EndFunc   ;==>WM_MEASUREITEM_Handler

Func _GetTextDimension($hWnd, $sText)
    ; Calculate text dimensions
    Local $hDC = __WinAPI_GetDC($hWnd)
    Local $hFont = __SendMessage($hWnd, $WM_GETFONT, 0, 0)
    If Not $hFont Then $hFont = __WinAPI_GetStockObject($DEFAULT_GUI_FONT)
    Local $hOldFont = __WinAPI_SelectObject($hDC, $hFont)

    Local $tSize = __WinAPI_GetTextExtentPoint32($hDC, $sText)
    Local $iTextWidth = $tSize.X
    Local $iTextHeight = $tSize.Y

    __WinAPI_SelectObject($hDC, $hOldFont)
    __WinAPI_ReleaseDC($hWnd, $hDC)
	Local $arSize = [$iTextWidth, $iTextHeight]
	Return $arSize
EndFunc

Func _CalcMenuItemWidth($iDPIpct, $iTextWidth)
    If $iDPIpct < 100 Or $iDPIpct > 400 Then
        Return $iTextWidth - 5
    EndIf

    Local Const $iSteps = Int(($iDPIpct - 100) / 25)
    Return $iTextWidth - (4 * $iSteps)
EndFunc   ;==>_CalcMenuItemWidth

;********************************************************************
; WM_DRAWITEM procedure
;********************************************************************
Func WM_DRAWITEM($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam
    Local Const $SM_CXDLGFRAME = 7
    Local Const $DEFAULT_GUI_FONT = 17
    Local $tagDRAWITEM = "uint CtlType;uint CtlID;uint itemID;uint itemAction;uint itemState;ptr hwndItem;handle hDC;" & _
            "long left;long top;long right;long bottom;ulong_ptr itemData"
    Local $t = DllStructCreate($tagDRAWITEM, $lParam)
    If Not IsDllStruct($t) Then Return $GUI_RUNDEFMSG

    If $t.CtlType <> $ODT_MENU Then Return $GUI_RUNDEFMSG


    Local $hDC = $t.hDC
    Local $left = $t.left
    Local $top = $t.top
    Local $right = $t.right
    Local $bottom = $t.bottom
    Local $state = $t.itemState
    Local $itemID = $t.itemID

	Local $hMenu = _GUICtrlMenu_GetMenu($hWnd)
    ; convert itemID to position
    Local $iPos = -1
    For $i = 0 To _GUICtrlMenu_GetItemCount($hMenu) - 1
        If $itemID = _GUICtrlMenu_GetItemID($hMenu, $i) Then
            $iPos = $i
            ExitLoop
        EndIf
    Next
    If $iPos < 0 Then Return 1 ; something must have gone seriously wrong

    Local $sText = _GUICtrlMenu_GetItemText($hMenu, $iPos)
    $sText = StringReplace($sText, "&", "")

    ; Colors
    Local $clrBG = _ColorToCOLORREF($COLOR_MENU_BG)
    Local $clrSel = _ColorToCOLORREF($COLOR_MENU_SEL)
    Local $clrText = _ColorToCOLORREF($COLOR_MENU_TEXT)

    ;Static $iDrawCount = 0
    ; Static $bFullBarDrawn = False

    ; Count how many items were drawn in this "draw cycle"
    ;$iDrawCount += 1

    ; argumentum ; pre-declare all the "Local" in those IF-THEN that could be needed
    Local $tClient, $iFullWidth, $tFullMenuBar, $hFullBrush
    Local $tEmptyArea, $hEmptyBrush

    ; If we are at the first item AND the bar has not yet been drawn
    If $iPos = 0 Then ; And Not $bFullBarDrawn Then
        ; Get the full window width
        $tClient = __WinAPI_GetClientRect($hWnd)
        $iFullWidth = $tClient.right

        ; Fill the entire menu bar
        $tFullMenuBar = DllStructCreate($tagRECT)
        With $tFullMenuBar
            .left = 0
            .top = $top - 1
            .right = $iFullWidth + 3
            .bottom = $bottom
        EndWith

        $hFullBrush = __WinAPI_CreateSolidBrush($clrBG)
        __WinAPI_FillRect($hDC, $tFullMenuBar, $hFullBrush)
        __WinAPI_DeleteObject($hFullBrush)
    EndIf

    ; After drawing all items, mark as "drawn"
    ;If $iDrawCount >= UBound($g_aMenuText) Then
        ; $bFullBarDrawn = True
        ;$iDrawCount = 0
    ;EndIf

    ; Draw background for the area AFTER the last menu item
    If $iPos = (_GUICtrlMenu_GetItemCount($hMenu) - 1) Then ; Last menu
        $tClient = __WinAPI_GetClientRect($hWnd)
        $iFullWidth = $tClient.right

        ; Fill only the area to the RIGHT of the last menu item
        If $right < $iFullWidth Then
            $tEmptyArea = DllStructCreate($tagRECT)
            With $tEmptyArea
                .left = $right
                .top = $top ;        argumentum ; replace magic numbers with it's parameter name when possible
                .right = $iFullWidth + __WinAPI_GetSystemMetrics($SM_CXDLGFRAME) ; 7 = $SM_CXDLGFRAME
                .bottom = $bottom
            EndWith

            $hEmptyBrush = __WinAPI_CreateSolidBrush($clrBG)
            __WinAPI_FillRect($hDC, $tEmptyArea, $hEmptyBrush)
            __WinAPI_DeleteObject($hEmptyBrush)
        EndIf
    EndIf

    ; Draw item background (selected = lighter)
    Local $bSelected = BitAND($state, $ODS_SELECTED)
    Local $bHot = BitAND($state, $ODS_HOTLIGHT)
    Local $hBrush

    If $bSelected Then
        $hBrush = __WinAPI_CreateSolidBrush($clrSel)
    ElseIf $bHot Then
        $hBrush = __WinAPI_CreateSolidBrush($COLOR_MENU_HOT)
    Else
        $hBrush = __WinAPI_CreateSolidBrush($clrBG)
    EndIf

    Local $tItemRect = DllStructCreate($tagRECT)
    With $tItemRect
        .left = $left
        .top = $top
        .right = $right
        .bottom = $bottom
    EndWith

    __WinAPI_FillRect($hDC, $tItemRect, $hBrush)
    __WinAPI_DeleteObject($hBrush)

    ; Setup font
    Local $hFont = __SendMessage($hWnd, $WM_GETFONT, 0, 0)
    If Not $hFont Then $hFont = __WinAPI_GetStockObject($DEFAULT_GUI_FONT)
    Local $hOldFont = __WinAPI_SelectObject($hDC, $hFont)

    __WinAPI_SetBkMode($hDC, $TRANSPARENT)
    __WinAPI_SetTextColor($hDC, $clrText)

    ; Draw text
    Local $tTextRect = DllStructCreate($tagRECT)
    With $tTextRect
        .left = $left + 10
        .top = $top + 4
        .right = $right - 10
        .bottom = $bottom - 4
    EndWith

    DllCall($hUser32Dll, "int", "DrawTextW", "handle", $hDC, "wstr", $sText, "int", -1, "ptr", _
            DllStructGetPtr($tTextRect), "uint", BitOR($DT_SINGLELINE, $DT_VCENTER, $DT_LEFT))

    If $hOldFont Then __WinAPI_SelectObject($hDC, $hOldFont)

    Return 1
EndFunc   ;==>WM_DRAWITEM

Func _GUITopMenuTheme($hWnd)
	$hGUI = $hWnd

    ; get top menu handle
    Local $hMenu = _GUICtrlMenu_GetMenu($hWnd)
    If Not $hMenu Then Return False
	GUIRegisterMsg($WM_WINDOWPOSCHANGED, "WM_WINDOWPOSCHANGED_Handler")
    GUIRegisterMsg($WM_ACTIVATE, "WM_ACTIVATE_Handler")

	; get window DPI for measurement adjustments
	$iDPIpct = Round(__WinAPI_GetDpiForWindow($hGUI) / 96, 2) * 100
	If @error Then $iDPIpct = 100

	For $i = 0 To _GUICtrlMenu_GetItemCount($hMenu) - 1
		_GUICtrlMenu_SetItemType($hMenu, $i, $MFT_OWNERDRAW, True)
	Next
    MenuBarBKColor($hMenu, $COLOR_MENU_BG)
EndFunc   ;==>_GUITopMenuTheme

Func _SetMenuColors($hWnd, $MenuBG, $MenuHot, $MenuSel, $MenuText)
    Local $hMenu = _GUICtrlMenu_GetMenu($hWnd)
    $COLOR_MENU_BG = $MenuBG
    $COLOR_MENU_HOT = $MenuHot
    $COLOR_MENU_SEL = $MenuSel
    $COLOR_MENU_TEXT = $MenuText
    ; redraw menubar background area
    MenuBarBKColor($hMenu, $COLOR_MENU_BG)
    ; redraw menubar and force refresh
    _GUICtrlMenu_DrawMenuBar($hWnd)
    __WinAPI_RedrawWindow($hWnd, 0, 0, BitOR($RDW_INVALIDATE, $RDW_UPDATENOW))
EndFunc   ;==>_SetMenuColors

Func WM_WINDOWPOSCHANGED_Handler($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam, $lParam
	If $hWnd <> $hGUI Then Return $GUI_RUNDEFMSG
	_drawUAHMenuNCBottomLine($hWnd)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_WINDOWPOSCHANGED_Handler

Func _drawUAHMenuNCBottomLine($hWnd)
    Local $rcClient = __WinAPI_GetClientRect($hWnd)

    DllCall($hUser32Dll, "int", "MapWindowPoints", _
        "hwnd", $hWnd, _ ; hWndFrom
        "hwnd", 0, _     ; hWndTo
        "ptr", DllStructGetPtr($rcClient), _
        "uint", 2)       ;number of points - 2 for RECT structure

    If @error Then
        ;MsgBox($MB_ICONERROR, "Error", @error)
        Exit
    EndIf

    Local $rcWindow = __WinAPI_GetWindowRect($hWnd)

    __WinAPI_OffsetRect($rcClient, -$rcWindow.left, -$rcWindow.top)

    Local $rcAnnoyingLine = DllStructCreate($tagRECT)
    $rcAnnoyingLine.left = $rcClient.left
    $rcAnnoyingLine.top = $rcClient.top
    $rcAnnoyingLine.right = $rcClient.right
    $rcAnnoyingLine.bottom = $rcClient.bottom

    $rcAnnoyingLine.bottom = $rcAnnoyingLine.top
    $rcAnnoyingLine.top = $rcAnnoyingLine.top - 1

    Local $hRgn = __WinAPI_CreateRectRgn(0,0,8000,8000)

    Local $hDC = __WinAPI_GetDCEx($hWnd,$hRgn, BitOR($DCX_WINDOW,$DCX_INTERSECTRGN))
    Local $hFullBrush = __WinAPI_CreateSolidBrush($COLOR_MENU_BG)
    __WinAPI_FillRect($hDC, $rcAnnoyingLine, $hFullBrush)
    __WinAPI_ReleaseDC($hWnd, $hDC)
    __WinAPI_DeleteObject($hFullBrush)

EndFunc   ;==>_drawUAHMenuNCBottomLine

Func WM_ACTIVATE_Handler($hWnd, $MsgID, $wParam, $lParam)
    #forceref $MsgID, $wParam, $lParam
    If $hWnd <> $hGUI Then Return $GUI_RUNDEFMSG
    _drawUAHMenuNCBottomLine($hWnd)

    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_ACTIVATE_Handler

Func MenuBarBKColor($hMenu, $nColor)
	Local $tInfo,$aResult
	Local $hBrush = DllCall($hGdi32Dll, 'hwnd', 'CreateSolidBrush', 'int', $nColor)
	If @error Then Return
	;$tInfo = DllStructCreate("int Size;int Mask;int Style;int YMax;int hBack;int ContextHelpID;ptr MenuData")
	$tInfo = DllStructCreate("int Size;int Mask;int Style;int YMax;handle hBack;int ContextHelpID;ptr MenuData")
	DllStructSetData($tInfo, "Mask", 2)
	DllStructSetData($tInfo, "hBack", $hBrush[0])
	DllStructSetData($tInfo, "Size", DllStructGetSize($tInfo))
	$aResult = DllCall($hUser32Dll, "int", "SetMenuInfo", "hwnd", $hMenu, "ptr", DllStructGetPtr($tInfo))
	Return $aResult[0] <> 0
EndFunc   ;==>_GUICtrlMenu_SetMenuBackground

Func _ColorToCOLORREF($iColor) ;RGB to BGR
    Local $iR = BitAND(BitShift($iColor, 16), 0xFF)
    Local $iG = BitAND(BitShift($iColor, 8), 0xFF)
    Local $iB = BitAND($iColor, 0xFF)
    Return BitOR(BitShift($iB, -16), BitShift($iG, -8), $iR)
EndFunc   ;==>_ColorToCOLORREF
