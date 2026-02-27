#include-once

#include <WinAPITheme.au3>
#include <Array.au3>
#include <GuiTreeView.au3>

Global $hKernel32, $hGdi32, $hUser32, $hShlwapi, $hShell32

Func __Timer_QueryPerformanceFrequency_mod()
	Local $aCall = DllCall($hKernel32, "bool", "QueryPerformanceFrequency", "int64*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	Return SetExtended($aCall[0], $aCall[1])
EndFunc   ;==>__Timer_QueryPerformanceFrequency_mod

Func __Timer_QueryPerformanceCounter_mod()
	Local $aCall = DllCall($hKernel32, "bool", "QueryPerformanceCounter", "int64*", 0)
	If @error Then Return SetError(@error, @extended, -1)
	Return SetExtended($aCall[0], $aCall[1])
EndFunc   ;==>__Timer_QueryPerformanceCounter_mod

Func _Timer_Diff_mod($iTimeStamp)
	Return 1000 * (__Timer_QueryPerformanceCounter_mod() - $iTimeStamp) / __Timer_QueryPerformanceFrequency_mod()
EndFunc   ;==>_Timer_Diff_mod

Func _Timer_Init_mod()
	Return __Timer_QueryPerformanceCounter_mod()
EndFunc   ;==>_Timer_Init_mod

Func _WinAPI_ReleaseDC_mod($hWnd, $hDC)
	Local $aCall = DllCall($hUser32, "int", "ReleaseDC", "hwnd", $hWnd, "handle", $hDC)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_ReleaseDC_mod

Func _WinAPI_GetDCEx_mod($hWnd, $hRgn, $iFlags)
	Local $aCall = DllCall($hUser32, 'handle', 'GetDCEx', 'hwnd', $hWnd, 'handle', $hRgn, 'dword', $iFlags)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_GetDCEx_mod

Func _WinAPI_CreateRectRgn_mod($iLeftRect, $iTopRect, $iRightRect, $iBottomRect)
	Local $aCall = DllCall($hGdi32, "handle", "CreateRectRgn", "int", $iLeftRect, "int", $iTopRect, "int", $iRightRect, _
			"int", $iBottomRect)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_CreateRectRgn_mod

Func _WinAPI_OffsetRect_mod(ByRef $tRECT, $iDX, $iDY)
	Local $aCall = DllCall($hUser32, 'bool', 'OffsetRect', 'struct*', $tRECT, 'int', $iDX, 'int', $iDY)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_OffsetRect_mod

Func _WinAPI_GetWindowRect_mod($hWnd)
	Local $tRECT = DllStructCreate($tagRECT)
	Local $aCall = DllCall($hUser32, "bool", "GetWindowRect", "hwnd", $hWnd, "struct*", $tRECT)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tRECT
EndFunc   ;==>_WinAPI_GetWindowRect_mod

Func _WinAPI_ShellGetFileInfo_mod($sFilePath, $iFlags, $iAttributes, ByRef $tSHFILEINFO)
	Local $aCall = DllCall($hShell32, 'dword_ptr', 'SHGetFileInfoW', 'wstr', $sFilePath, 'dword', $iAttributes, _
			'struct*', $tSHFILEINFO, 'uint', DllStructGetSize($tSHFILEINFO), 'uint', $iFlags)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_ShellGetFileInfo_mod

Func _WinAPI_GetClientRect_mod($hWnd)
	Local $tRECT = DllStructCreate($tagRECT)
	Local $aCall = DllCall($hUser32, "bool", "GetClientRect", "hwnd", $hWnd, "struct*", $tRECT)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tRECT
EndFunc   ;==>_WinAPI_GetClientRect_mod

Func _WinAPI_GetWindowLong_mod($hWnd, $iIndex)
	Local $sFuncName = "GetWindowLongW"
	If @AutoItX64 Then $sFuncName = "GetWindowLongPtrW"
	Local $aCall = DllCall($hUser32, "long_ptr", $sFuncName, "hwnd", $hWnd, "int", $iIndex)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_GetWindowLong_mod

Func _WinAPI_DefWindowProc_mod($hWnd, $iMsg, $wParam, $lParam)
	Local $aCall = DllCall($hUser32, "lresult", "DefWindowProc", "hwnd", $hWnd, "uint", $iMsg, "wparam", $wParam, _
			"lparam", $lParam)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_DefWindowProc_mod

Func _WinAPI_SetLastError_mod($iErrorCode, Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
	DllCall($hKernel32, "none", "SetLastError", "dword", $iErrorCode)
	Return SetError($_iCallerError, $_iCallerExtended, Null)
EndFunc   ;==>_WinAPI_SetLastError_mod

Func _WinAPI_SetWindowLong_mod($hWnd, $iIndex, $iValue)
	_WinAPI_SetLastError_mod(0) ; as suggested in MSDN
	Local $sFuncName = "SetWindowLongW"
	If @AutoItX64 Then $sFuncName = "SetWindowLongPtrW"
	Local $aCall = DllCall($hUser32, "long_ptr", $sFuncName, "hwnd", $hWnd, "int", $iIndex, "long_ptr", $iValue)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_SetWindowLong_mod

Func _WinAPI_SetWindowPos_mod($hWnd, $hAfter, $iX, $iY, $iCX, $iCY, $iFlags)
	Local $aCall = DllCall($hUser32, "bool", "SetWindowPos", "hwnd", $hWnd, "hwnd", $hAfter, "int", $iX, "int", $iY, _
			"int", $iCX, "int", $iCY, "uint", $iFlags)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_SetWindowPos_mod

Func _WinAPI_DrawText_mod($hDC, $sText, ByRef $tRECT, $iFlags)
	Local $aCall = DllCall($hUser32, "int", "DrawTextW", "handle", $hDC, "wstr", $sText, "int", -1, "struct*", $tRECT, _
			"uint", $iFlags)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_DrawText_mod

Func _WinAPI_SetBkColor_mod($hDC, $iColor)
	Local $aCall = DllCall($hGdi32, "INT", "SetBkColor", "handle", $hDC, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_SetBkColor_mod

Func _WinAPI_DeleteObject_mod($hObject)
	Local $aCall = DllCall($hGdi32, "bool", "DeleteObject", "handle", $hObject)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_DeleteObject_mod

Func _WinAPI_InflateRect_mod(ByRef $tRECT, $iDX, $iDY)
	Local $aCall = DllCall($hUser32, 'bool', 'InflateRect', 'struct*', $tRECT, 'int', $iDX, 'int', $iDY)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_InflateRect_mod

Func _WinAPI_FillRect_mod($hDC, $tRECT, $hBrush)
	Local $aCall
	If IsPtr($hBrush) Then
		$aCall = DllCall($hUser32, "int", "FillRect", "handle", $hDC, "struct*", $tRECT, "handle", $hBrush)
	Else
		$aCall = DllCall($hUser32, "int", "FillRect", "handle", $hDC, "struct*", $tRECT, "dword_ptr", $hBrush)
	EndIf
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_FillRect_mod

Func _WinAPI_CreateSolidBrush_mod($iColor)
	Local $aCall = DllCall($hGdi32, "handle", "CreateSolidBrush", "INT", $iColor)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_CreateSolidBrush_mod

Func _WinAPI_GetClassName_mod($hWnd)
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $aCall = DllCall($hUser32, "int", "GetClassNameW", "hwnd", $hWnd, "wstr", "", "int", 4096)
	If @error Or Not $aCall[0] Then Return SetError(@error, @extended, '')

	Return SetExtended($aCall[0], $aCall[2])
EndFunc   ;==>_WinAPI_GetClassName_mod

Func _WinAPI_SetTextColor_mod($hDC, $iColor)
	Local $aCall = DllCall($hGdi32, "INT", "SetTextColor", "handle", $hDC, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_SetTextColor_mod

Func _WinAPI_PathIsRoot_mod($sFilePath)
	Local $aCall = DllCall($hShlwapi, 'bool', 'PathIsRootW', 'wstr', $sFilePath & "\")
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_PathIsRoot_mod

Func _WinAPI_ScreenToClient_mod($hWnd, ByRef $tPoint)
	Local $aCall = DllCall($hUser32, "bool", "ScreenToClient", "hwnd", $hWnd, "struct*", $tPoint)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_ScreenToClient_mod

Func TreeItemToPath($hTree, $hItem, $bArray = False)
	Local $sPath = StringReplace(_GUICtrlTreeView_GetTree($hTree, $hItem), "|", "\")
	$sPath = StringTrimLeft($sPath, StringInStr($sPath, "\"))     ; remove this pc at the beginning
	If StringInStr(FileGetAttrib($sPath), "D") Then $sPath &= "\"   ; let folders end with \
	If $bArray Then
		Local $aPath = _ArrayFromString($sPath)
		_ArrayInsert($aPath, 0, 1)
		Return $aPath
	EndIf
	Return $sPath
EndFunc   ;==>TreeItemToPath

Func _PathSplit_mod($sFilePath)
	Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	Local $aArray = StringRegExp($sFilePath, "^\h*((?:\\\\\?\\)*(\\\\[^\?\/\\]+|[A-Za-z]:)?(.*[\/\\]\h*)?((?:[^\.\/\\]|(?(?=\.[^\/\\]*\.)\.))*)?([^\/\\]*))$", $STR_REGEXPARRAYMATCH)
	If @error Then ; This error should never happen.
		ReDim $aArray[5]
		$aArray[$PATH_ORIGINAL] = $sFilePath
	EndIf
	$sDrive = $aArray[$PATH_DRIVE]
	If StringLeft($aArray[$PATH_DIRECTORY], 1) == "/" Then
		$sDir = StringRegExpReplace($aArray[$PATH_DIRECTORY], "\h*[\/\\]+\h*", "\/")
	Else
		$sDir = StringRegExpReplace($aArray[$PATH_DIRECTORY], "\h*[\/\\]+\h*", "\\")
	EndIf
	$aArray[$PATH_DIRECTORY] = $sDir
	$sFileName = $aArray[$PATH_FILENAME]
	$sExtension = $aArray[$PATH_EXTENSION]

	Return $aArray
EndFunc   ;==>_PathSplit_mod

Func _WinAPI_FindWindowEx($hParent, $hAfter, $sClass, $sTitle = "")
    Local $ret = DllCall($hUser32, "hwnd", "FindWindowExW", "hwnd", $hParent, "hwnd", $hAfter, "wstr", $sClass, "wstr", $sTitle)
    If @error Or Not IsArray($ret) Then Return 0
    Return $ret[0]
EndFunc
