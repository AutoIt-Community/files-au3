#include-once

; #INDEX# =======================================================================================================================
; Title .........: GUIDarkMenu UDF Library for AutoIt3
; AutoIt Version : 3.3.18.0
; Language ......: English
; Description ...: UDF library for applying dark theme to menubar
; Author(s) .....: WildByDesign, Kanashius (including previous code from ahmet, argumentum, UEZ)
; Version .......: 0.9.2
; ===============================================================================================================================

#include <WinAPISysWin.au3>
#include <GuiMenu.au3>
#include <GUIConstantsEx.au3>
#include <APIGdiConstants.au3>
#include <WindowsNotifsConstants.au3>
#include <WinAPIGdiDC.au3>
#include <StructureConstants.au3>
#include <AutoItConstants.au3>
#include <WinAPIGdi.au3>
#include <WinAPIHObj.au3>

; Menu Info
Const $ODT_MENU = 1
Const $ODS_SELECTED = 0x0001
Const $ODS_DISABLED = 0x0004
Const $ODS_HOTLIGHT = 0x0040

Global $__GUIDarkMenu_mData[]
Global Const $__GUIDarkMenu_iThemeLight = 1, $__GUIDarkMenu_iThemeDark = 2
Global Const $__GUIDarkMenu_vColorBG = "iColorBG", $__GUIDarkMenu_vColor = "iColor", $__GUIDarkMenu_vColorCtrlBG = "iColorCtrlBG", $__GUIDarkMenu_vColorBorder = "iColorBorder"
Global Const $__GUIDarkMenu_vColorMenuBG = "iColorMenuBG", $__GUIDarkMenu_vColorMenuHot = "iColorMenuHot", $__GUIDarkMenu_vColorMenuSel = "iColorMenuSel", $__GUIDarkMenu_vColorMenuText = "iColorMenuText"

Func __GUIDarkMenu_StartUp()
	$__GUIDarkMenu_mData.hDllGDI = DllOpen("gdi32.dll")
	$__GUIDarkMenu_mData.hDllUser = DllOpen("user32.dll")
	Local $mGuis[]
	$__GUIDarkMenu_mData.mGuis = $mGuis
	$__GUIDarkMenu_mData.hProc = DllCallbackRegister('__GUIDarkMenu_WinProc', 'ptr', 'hwnd;uint;wparam;lparam')
	Local $mThemes[]
	Local $mDarkTheme[]
	$mDarkTheme[$__GUIDarkMenu_vColorBG] = 0x121212
	$mDarkTheme[$__GUIDarkMenu_vColor] = 0xE0E0E0
	$mDarkTheme[$__GUIDarkMenu_vColorCtrlBG] = 0x202020
	$mDarkTheme[$__GUIDarkMenu_vColorBorder] = 0x3F3F3F
	$mDarkTheme[$__GUIDarkMenu_vColorMenuBG] = _WinAPI_ColorAdjustLuma($mDarkTheme[$__GUIDarkMenu_vColorBG], 5)
	$mDarkTheme[$__GUIDarkMenu_vColorMenuHot] = _WinAPI_ColorAdjustLuma($mDarkTheme[$__GUIDarkMenu_vColorMenuBG], 20)
	$mDarkTheme[$__GUIDarkMenu_vColorMenuSel] = _WinAPI_ColorAdjustLuma($mDarkTheme[$__GUIDarkMenu_vColorMenuBG], 10)
	$mDarkTheme[$__GUIDarkMenu_vColorMenuText] = $mDarkTheme[$__GUIDarkMenu_vColor]
	$mThemes[$__GUIDarkMenu_iThemeDark] = $mDarkTheme
	Local $mLightTheme[]
	$mLightTheme[$__GUIDarkMenu_vColorBG] = 0xFFFFFF
	$mLightTheme[$__GUIDarkMenu_vColor] = 0x000000
	$mLightTheme[$__GUIDarkMenu_vColorCtrlBG] = 0xDDDDDD
	$mLightTheme[$__GUIDarkMenu_vColorBorder] = 0xCCCCCC
	$mLightTheme[$__GUIDarkMenu_vColorMenuBG] = _WinAPI_ColorAdjustLuma($mLightTheme[$__GUIDarkMenu_vColorBG], 95)
	$mLightTheme[$__GUIDarkMenu_vColorMenuHot] = _WinAPI_ColorAdjustLuma($mLightTheme[$__GUIDarkMenu_vColorMenuBG], 80)
	$mLightTheme[$__GUIDarkMenu_vColorMenuSel] = _WinAPI_ColorAdjustLuma($mLightTheme[$__GUIDarkMenu_vColorMenuBG], 70)
	$mLightTheme[$__GUIDarkMenu_vColorMenuText] = $mLightTheme[$__GUIDarkMenu_vColor]
	$mThemes[$__GUIDarkMenu_iThemeLight] = $mLightTheme
	$__GUIDarkMenu_mData.mThemes = $mThemes
EndFunc

Func __GUIDarkMenu_Shutdown()
	If Not __GUIDarkMenu__IsInitialized() Then Return SetError(2, 0, False)
	For $hGui In MapKeys($__GUIDarkMenu_mData.mGuis)
		__GUIDarkMenu_GuiRemove($hGui)
	Next
	DllCallbackFree($__GUIDarkMenu_mData.hProc)
	DllClose($__GUIDarkMenu_mData.hDllGDI)
	DllClose($__GUIDarkMenu_mData.hDllUser)
	Local $mNewMap[]
	$__GUIDarkMenu_mData = $mNewMap
	Return True
EndFunc

Func __GUIDarkMenu_GuiAdd($hGui)
	If Not __GUIDarkMenu__IsInitialized() Then Return SetError(2, 0, False)
	If MapExists($__GUIDarkMenu_mData.mGuis, $hGui) Then Return True
	Local $mGui[]
	Local $iDpiPct = Round(__WinAPI_GetDpiForWindow($hGUI) / 96, 2) * 100
	$mGui.iDpi = @error?100:$iDpiPct
	Local $hProc = _WinAPI_SetWindowLong($hGui, -4, DllCallbackGetPtr($__GUIDarkMenu_mData.hProc))
	If @error Then Return SetError(2, 0, False)
	$mGui.hPrevProc = $hProc
	$mGui.iTheme = $__GUIDarkMenu_iThemeLight
	$mGui.iTextSpaceHori = 20
	$mGui.iTextSpaceVert = 8
	$mGui.iFontSize = 9
	$__GUIDarkMenu_mData["mGuis"][$hGui] = $mGui
	Return True
EndFunc

Func __GUIDarkMenu_GuiRemove($hGui)
	If Not __GUIDarkMenu__IsInitialized() Then Return SetError(2, 0, False)
	If Not MapExists($__GUIDarkMenu_mData.mGuis, $hGui) Then Return SetError(1, 1, False)
	_WinAPI_SetWindowLong($hGui, -4, $__GUIDarkMenu_mData.mGuis[$hGui].hPrevProc)
	MapRemove($__GUIDarkMenu_mData.mGuis, $hGui)
	Return True
EndFunc

Func __GUIDarkMenu__IsInitialized()
	Return UBound($__GUIDarkMenu_mData)>0
EndFunc

Func __GUIDarkMenu_WinProc($hWnd, $iMsg, $iwParam, $ilParam)
	Local $sContinue = $GUI_RUNDEFMSG
    Switch $iMsg
        Case $WM_WINDOWPOSCHANGED
            $sContinue = __GUIDarkMode__WM_WINDOWPOSCHANGED($hWnd, $iMsg, $iwParam, $ilParam)
        Case $WM_ACTIVATE
            $sContinue = __GUIDarkMode__WM_ACTIVATE($hWnd, $iMsg, $iwParam, $ilParam)
        Case $WM_MEASUREITEM
            $sContinue = __GUIDarkMode__WM_MEASUREITEM($hWnd, $iMsg, $iwParam, $ilParam)
        Case $WM_DRAWITEM
            $sContinue = __GUIDarkMode__WM_DRAWITEM($hWnd, $iMsg, $iwParam, $ilParam)
	EndSwitch
	If $sContinue=$GUI_RUNDEFMSG And MapExists($__GUIDarkMenu_mData.mGuis, $hWnd) Then Return _WinAPI_CallWindowProc($__GUIDarkMenu_mData.mGuis[$hWnd].hPrevProc, $hWnd, $iMsg, $iwParam, $ilParam)
EndFunc

Func __GUIDarkMenu_SetTheme($hGui, $iTheme)
	If Not __GUIDarkMenu__IsInitialized() Then Return SetError(2, 0, False)
	__GUIDarkMenu_GuiAdd($hGui)
	If @error Then Return SetError(1, 1, False)
	If Not MapExists($__GUIDarkMenu_mData.mThemes, $iTheme) Then Return SetError(1, 2, False)
	$__GUIDarkMenu_mData["mGuis"][$hGui]["iTheme"] = $iTheme
    Local $hMenu = _GUICtrlMenu_GetMenu($hGui)
	If Not $hMenu Then Return False
	For $i = 0 To _GUICtrlMenu_GetItemCount($hMenu) - 1
		_GUICtrlMenu_SetItemType($hMenu, $i, $MFT_OWNERDRAW, True)
	Next
	__GUIDarkMenu__MenuBarSetBKColor($hMenu, __GUIDarkMenu__GetColor($hGui, $__GUIDarkMenu_vColorMenuBG))
    _GUICtrlMenu_DrawMenuBar($hGui)
    _WinAPI_RedrawWindow($hGui, 0, 0, BitOR($RDW_INVALIDATE, $RDW_UPDATENOW))
	Return True
EndFunc

Func __GUIDarkMenu__GetColor($hGui, $sColor)
	If Not __GUIDarkMenu__IsInitialized() Then Return SetError(2, 0, False)
	If Not MapExists($__GUIDarkMenu_mData.mGuis, $hGui) Then Return SetError(1, 1, 0)
	If Not MapExists($__GUIDarkMenu_mData.mThemes[$__GUIDarkMenu_mData.mGuis[$hGui].iTheme], $sColor) Then Return SetError(1, 2, 0)
	Return $__GUIDarkMenu_mData.mThemes[$__GUIDarkMenu_mData.mGuis[$hGui].iTheme][$sColor]
EndFunc

Func __GUIDarkMenu__MenuBarSetBKColor($hMenu, $iColor)
	Local $tInfo,$aResult
	Local $hBrush = DllCall($__GUIDarkMenu_mData.hDllGDI, 'hwnd', 'CreateSolidBrush', 'int', $iColor)
	If @error Then Return
	;$tInfo = DllStructCreate("int Size;int Mask;int Style;int YMax;int hBack;int ContextHelpID;ptr MenuData")
	$tInfo = DllStructCreate("int Size;int Mask;int Style;int YMax;handle hBack;int ContextHelpID;ptr MenuData")
	DllStructSetData($tInfo, "Mask", 2)
	DllStructSetData($tInfo, "hBack", $hBrush[0])
	DllStructSetData($tInfo, "Size", DllStructGetSize($tInfo))
	$aResult = DllCall($__GUIDarkMenu_mData.hDllUser, "int", "SetMenuInfo", "hwnd", $hMenu, "ptr", DllStructGetPtr($tInfo))
	Return $aResult[0] <> 0
EndFunc   ;==>_GUICtrlMenu_SetMenuBackground

Func __WinAPI_GetDpiForWindow($hWnd)
    Local $aResult = DllCall($__GUIDarkMenu_mData.hDllUser, "uint", "GetDpiForWindow", "hwnd", $hWnd) ;requires Win10 v1607+ / no server support
    If Not IsArray($aResult) Or @error Then Return SetError(1, @extended, 0)
    If Not $aResult[0] Then Return SetError(2, @extended, 0)
    Return $aResult[0]
EndFunc   ;==>__WinAPI_GetDpiForWindow

Func __GUIDarkMenu__GetTextDimension($hWnd, $sText)
    ; Calculate text dimensions
    Local $hDC = _WinAPI_GetDC($hWnd)
    Local $hFont = _SendMessage($hWnd, $WM_GETFONT, 0, 0)
    If Not $hFont Then $hFont = _WinAPI_GetStockObject($DEFAULT_GUI_FONT)
    Local $hOldFont = _WinAPI_SelectObject($hDC, $hFont)

    Local $tSize = _WinAPI_GetTextExtentPoint32($hDC, $sText)
    Local $iTextWidth = $tSize.X + 6
    Local $iTextHeight = $tSize.Y

    _WinAPI_SelectObject($hDC, $hOldFont)
    _WinAPI_ReleaseDC($hWnd, $hDC)
	Local $arSize = [$iTextWidth, $iTextHeight]
	Return $arSize
EndFunc

Func __GUIDarkMode__WM_MEASUREITEM($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam
	If Not MapExists($__GUIDarkMenu_mData.mGuis, $hWnd) Then Return $GUI_RUNDEFMSG
    Local $tagMEASUREITEM = "uint CtlType;uint CtlID;uint itemID;uint itemWidth;uint itemHeight;ulong_ptr itemData"
    Local $t = DllStructCreate($tagMEASUREITEM, $lParam)
    If Not IsDllStruct($t) Then Return $GUI_RUNDEFMSG

    If $t.CtlType <> $ODT_MENU Then Return $GUI_RUNDEFMSG

    Local $itemID = $t.itemID

    Local $sText = _GUICtrlMenu_GetItemText(_GUICtrlMenu_GetMenu($hWnd), $itemID, False)
	Local $arSize = __GUIDarkMenu__GetTextDimension($hWnd, $sText)

    ; Set dimensions with padding (with high DPI)
	$t.itemWidth = $arSize[0] + $__GUIDarkMenu_mData.mGuis[$hWnd].iTextSpaceHori/2
    $t.itemHeight = $arSize[1] + $__GUIDarkMenu_mData.mGuis[$hWnd].iTextSpaceVert
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_MEASUREITEM_Handler

Func __GUIDarkMode__WM_DRAWITEM($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam
    Local $tagDRAWITEM = "uint CtlType;uint CtlID;uint itemID;uint itemAction;uint itemState;ptr hwndItem;handle hDC;" & _
            "long left;long top;long right;long bottom;ulong_ptr itemData"
    Local $tDrawItem = DllStructCreate($tagDRAWITEM, $lParam)
    If Not IsDllStruct($tDrawItem) Then Return $GUI_RUNDEFMSG
    If $tDrawItem.CtlType <> $ODT_MENU Then Return $GUI_RUNDEFMSG

    Local $hDC = $tDrawItem.hDC
    Local $iLeft = $tDrawItem.left
    Local $iTop = $tDrawItem.top
    Local $iRight = $tDrawItem.right
    Local $iBottom = $tDrawItem.bottom
    Local $iState = $tDrawItem.itemState
    Local $iItemID = $tDrawItem.itemID

	Local $hMenu = _GUICtrlMenu_GetMenu($hWnd)
    ; convert itemID to position
    Local $iPos = -1
    For $i = 0 To _GUICtrlMenu_GetItemCount($hMenu) - 1
        If $iItemID = _GUICtrlMenu_GetItemID($hMenu, $i) Then
            $iPos = $i
            ExitLoop
        EndIf
    Next
    If $iPos < 0 Then Return $GUI_RUNDEFMSG ; something must have gone seriously wrong

    Local $sText = _GUICtrlMenu_GetItemText($hMenu, $iPos)
    $sText = StringReplace($sText, "&", "")

    ; Draw item background (selected = lighter)
    Local $bSelected = BitAND($iState, $ODS_SELECTED)
    Local $bHot = BitAND($iState, $ODS_HOTLIGHT)
    Local $hBrush

	; __GUIDarkMenu__ColorRGBToBGR()
    If $bSelected Then
        $hBrush = _WinAPI_CreateSolidBrush(__GUIDarkMenu__GetColor($hWnd, $__GUIDarkMenu_vColorMenuSel))
    ElseIf $bHot Then
        $hBrush = _WinAPI_CreateSolidBrush(__GUIDarkMenu__GetColor($hWnd, $__GUIDarkMenu_vColorMenuHot))
    Else
        $hBrush = _WinAPI_CreateSolidBrush(__GUIDarkMenu__GetColor($hWnd, $__GUIDarkMenu_vColorMenuBG))
    EndIf

    Local $tItemRect = DllStructCreate($tagRECT)
    With $tItemRect
        .left = $iLeft
        .top = $iTop
        .right = $iRight
        .bottom = $iBottom
    EndWith

    _WinAPI_FillRect($hDC, $tItemRect, $hBrush)
    _WinAPI_DeleteObject($hBrush)

    ; Setup font
    Local $hFont = __GUIDarkMenu__CreateMenuFontByName("Segoe UI", $__GUIDarkMenu_mData.mGuis[$hWnd].iFontSize)
    If Not $hFont Then $hFont = _WinAPI_GetStockObject($DEFAULT_GUI_FONT)
    Local $hOldFont = _WinAPI_SelectObject($hDC, $hFont)

    _WinAPI_SetBkMode($hDC, $TRANSPARENT)
    _WinAPI_SetTextColor($hDC, __GUIDarkMenu__GetColor($hWnd, $__GUIDarkMenu_vColorMenuText))

    ; Draw text
    Local $tTextRect = DllStructCreate($tagRECT)
    With $tTextRect
        .left = $iLeft + $__GUIDarkMenu_mData.mGuis[$hWnd].iTextSpaceHori/2
        .top = $iTop + $__GUIDarkMenu_mData.mGuis[$hWnd].iTextSpaceVert/2
        .right = $iRight - $__GUIDarkMenu_mData.mGuis[$hWnd].iTextSpaceHori/2
        .bottom = $iBottom - $__GUIDarkMenu_mData.mGuis[$hWnd].iTextSpaceVert/2
    EndWith

    DllCall($__GUIDarkMenu_mData.hDllUser, "int", "DrawTextW", "handle", $hDC, "wstr", $sText, "int", -1, "ptr", _
            DllStructGetPtr($tTextRect), "uint", BitOR($DT_SINGLELINE, $DT_VCENTER, $DT_LEFT))

    If $hOldFont Then _WinAPI_SelectObject($hDC, $hOldFont)

    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_DRAWITEM

Func __GUIDarkMode__WM_WINDOWPOSCHANGED($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam, $lParam
	If Not MapExists($__GUIDarkMenu_mData.mGuis, $hWnd) Then Return $GUI_RUNDEFMSG
	__GUIDarkMode__DrawUAHMenuNCBottomLine($hWnd)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_WINDOWPOSCHANGED_Handler

Func __GUIDarkMode__DrawUAHMenuNCBottomLine($hWnd)
    Local $rcClient = _WinAPI_GetClientRect($hWnd)
    DllCall($__GUIDarkMenu_mData.hDllUser, "int", "MapWindowPoints", _
        "hwnd", $hWnd, _ ; hWndFrom
        "hwnd", 0, _     ; hWndTo
        "ptr", DllStructGetPtr($rcClient), _
        "uint", 2)       ;number of points - 2 for RECT structure

    Local $rcWindow = _WinAPI_GetWindowRect($hWnd)
    _WinAPI_OffsetRect($rcClient, -$rcWindow.left, -$rcWindow.top)

    Local $rcAnnoyingLine = DllStructCreate($tagRECT)
    $rcAnnoyingLine.left = $rcClient.left
    $rcAnnoyingLine.top = $rcClient.top
    $rcAnnoyingLine.right = $rcClient.right
    $rcAnnoyingLine.bottom = $rcClient.bottom

    $rcAnnoyingLine.bottom = $rcAnnoyingLine.top
    $rcAnnoyingLine.top = $rcAnnoyingLine.top - 1

    Local $hRgn = _WinAPI_CreateRectRgn(0,0,8000,8000)

    Local $hDC = _WinAPI_GetDCEx($hWnd,$hRgn, BitOR($DCX_WINDOW,$DCX_INTERSECTRGN))
    Local $hFullBrush = _WinAPI_CreateSolidBrush(__GUIDarkMenu__GetColor($hWnd, $__GUIDarkMenu_vColorMenuBG))
    _WinAPI_FillRect($hDC, $rcAnnoyingLine, $hFullBrush)
    _WinAPI_ReleaseDC($hWnd, $hDC)
    _WinAPI_DeleteObject($hFullBrush)

EndFunc   ;==>__GUIDarkMode__DrawUAHMenuNCBottomLine

Func __GUIDarkMode__WM_ACTIVATE($hWnd, $MsgID, $wParam, $lParam)
    #forceref $MsgID, $wParam, $lParam
    If Not MapExists($__GUIDarkMenu_mData.mGuis, $hWnd) Then Return $GUI_RUNDEFMSG
    __GUIDarkMode__DrawUAHMenuNCBottomLine($hWnd)

    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_ACTIVATE_Handler

Func __GUIDarkMenu__ColorRGBToBGR($iColor) ;RGB to BGR
    Local $iR = BitAND(BitShift($iColor, 16), 0xFF)
    Local $iG = BitAND(BitShift($iColor, 8), 0xFF)
    Local $iB = BitAND($iColor, 0xFF)
    Return BitOR(BitShift($iB, -16), BitShift($iG, -8), $iR)
EndFunc   ;==>__GUIDarkMenu__ColorRGBToBGR

Func __GUIDarkMenu__CreateFont($nHeight, $nWidth, $nEscape, $nOrientn, $fnWeight, $bItalic, $bUnderline, $bStrikeout, $nCharset, $nOutputPrec, $nClipPrec, $nQuality, $nPitch, $ptrFontName)
	Local $hFont = DllCall($__GUIDarkMenu_mData.hDllGDI , "hwnd", "CreateFont", _
												"int", $nHeight, _
												"int", $nWidth, _
												"int", $nEscape, _
												"int", $nOrientn, _
												"int", $fnWeight, _
												"long", $bItalic, _
												"long", $bUnderline, _
												"long", $bStrikeout, _
												"long", $nCharset, _
												"long", $nOutputPrec, _
												"long", $nClipPrec, _
												"long", $nQuality, _
												"long", $nPitch, _
												"ptr", $ptrFontName)
	Return $hFont[0]
EndFunc

Func __GUIDarkMenu__CreateMenuFontByName($sFontName, $nHeight = 9, $nWidth = 400)
	Local $stFontName = DllStructCreate("char[260]")
	DllStructSetData($stFontName, 1, $sFontName)
    Local $hDC		= _WinAPI_GetDC(0)
    Local $nPixel	= _WinAPI_GetDeviceCaps($hDC, 90)
    $nHeight	= 0 - _WinAPI_MulDiv($nHeight, $nPixel, 72)
    _WinAPI_ReleaseDC(0, $hDC)
	Local $hFont = __GUIDarkMenu__CreateFont($nHeight, 0, 0, 0, $nWidth, 0, 0, 0, 0, 0, 0, 0, 0, DllStructGetPtr($stFontName))
	$stFontName = 0
	Return $hFont
EndFunc