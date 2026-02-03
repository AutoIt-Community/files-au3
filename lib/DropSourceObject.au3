#include-once
#include "InternalObject.au3"
#include "IUnknown.au3"

Global $__g_iDropSourceCount

Global $__g_hMthd_QueryContinueDrag, $__g_hMthd_GiveFeedback
Global $__g_hMthd_DragEnterTarget, $__g_hMthd_DragLeaveTarget

Global $tagSourceObjIntData = "hwnd hTarget;"


Func CreateDropSource()
	$__g_iDropSourceCount += 1

	Local $iObjectId = PrepareInternalObject(3)
	Local $tObject = $__g_aObjects[$iObjectId][1]
	Local $tSupportedIIDs = $__g_aObjects[$iObjectId][2]

	If Not $__g_hMthd_QueryContinueDrag Then
		$__g_hMthd_QueryContinueDrag = DllCallbackRegister("__Mthd_QueryContinueDrag", "long", "ptr;bool;dword")
		$__g_hMthd_GiveFeedback = DllCallbackRegister("__Mthd_GiveFeedback", "long", "ptr;dword")
		$__g_hMthd_DragEnterTarget = DllCallbackRegister("__Mthd_DragEnterTarget", "long", "ptr;hwnd")
		$__g_hMthd_DragLeaveTarget = DllCallbackRegister("__Mthd_DragLeaveTarget", "long", "ptr")
	EndIf

	Local $tIDropSrcVTab = DllStructCreate("ptr pFunc[5]")
	$tIDropSrcVTab.pFunc(1) = DllCallbackGetPtr($__g_hMthd_QueryInterfaceThunk)
	$tIDropSrcVTab.pFunc(2) = DllCallbackGetPtr($__g_hMthd_AddRefThunk)
	$tIDropSrcVTab.pFunc(3) = DllCallbackGetPtr($__g_hMthd_ReleaseThunk)
	$tIDropSrcVTab.pFunc(4) = DllCallbackGetPtr($__g_hMthd_QueryContinueDrag)
	$tIDropSrcVTab.pFunc(5) = DllCallbackGetPtr($__g_hMthd_GiveFeedback)

	Local $tIDropSrcNotifyVTab = DllStructCreate("ptr pFunc[5]")
	$tIDropSrcNotifyVTab.pFunc(1) = DllCallbackGetPtr($__g_hMthd_QueryInterfaceThunk)
	$tIDropSrcNotifyVTab.pFunc(2) = DllCallbackGetPtr($__g_hMthd_AddRefThunk)
	$tIDropSrcNotifyVTab.pFunc(3) = DllCallbackGetPtr($__g_hMthd_ReleaseThunk)
	$tIDropSrcNotifyVTab.pFunc(4) = DllCallbackGetPtr($__g_hMthd_DragEnterTarget)
	$tIDropSrcNotifyVTab.pFunc(5) = DllCallbackGetPtr($__g_hMthd_DragLeaveTarget)

	Local $tInternalData = DllStructCreate($tagSourceObjIntData)

	$tObject.pVTab(2) = DllStructGetPtr($tIDropSrcVTab)
	$tObject.pVTab(3) = DllStructGetPtr($tIDropSrcNotifyVTab)
	$tObject.pData = DllStructGetPtr($tInternalData)
	_WinAPI_GUIDFromStringEx($sIID_IDropSource, DllStructGetPtr($tSupportedIIDs, 2))
	_WinAPI_GUIDFromStringEx($sIID_IDropSourceNotify, DllStructGetPtr($tSupportedIIDs, 3))

	$__g_aObjects[$iObjectId][4] = $tIDropSrcVTab
	$__g_aObjects[$iObjectId][5] = $tIDropSrcNotifyVTab
	$__g_aObjects[$iObjectId][6] = $tInternalData

;~ 	ConsoleWrite("IUnknown Location: " & DllStructGetPtr($tObject, "pVTab") & @CRLF)
;~ 	ConsoleWrite("IDropSource Location: " & DllStructGetPtr($tObject, "pVTab") + $PTR_LEN & @CRLF)
;~ 	ConsoleWrite("IDropNotify Location: " & DllStructGetPtr($tObject, "pVTab") + 2*$PTR_LEN & @CRLF)

	$__g_aObjects[$iObjectId][0] = DllStructGetPtr($tObject, "pVTab") + $PTR_LEN
	Return $__g_aObjects[$iObjectId][0]
EndFunc   ;==>CreateDropSource

Func DestroyDropSource($pObject)
	If (Not $pObject) Or (Not IsPtr($pObject)) Then Return SetError($ERROR_INVALID_PARAMETER, 0, False)

	DestroyInternalObject($pObject)
	If Not @error Then
		$__g_iDropSourceCount -= 1
		If Not $__g_iDropSourceCount Then
			DllCallbackFree($__g_hMthd_QueryContinueDrag)
			DllCallbackFree($__g_hMthd_GiveFeedback)
			DllCallbackFree($__g_hMthd_DragEnterTarget)
			DllCallbackFree($__g_hMthd_DragLeaveTarget)

			$__g_hMthd_QueryContinueDrag = 0
			$__g_hMthd_GiveFeedback = 0
			$__g_hMthd_DragEnterTarget = 0
			$__g_hMthd_DragLeaveTarget = 0
		EndIf
	EndIf
EndFunc   ;==>DestroyDropSource


Func __Mthd_QueryContinueDrag($pThis, $bEscapePressed, $iKeyState)
	#forceref $pThis, $bEscapePressed, $iKeyState

	Local $iReturn = $S_OK
	If $bEscapePressed Then
		$iReturn = $DRAGDROP_S_CANCEL
	Else
		If Not BitAND($iKeyState, BitOR($MK_LBUTTON, $MK_RBUTTON)) Then $iReturn = $DRAGDROP_S_DROP
	EndIf

	Return $iReturn
EndFunc   ;==>__Mthd_QueryContinueDrag


Func __Mthd_GiveFeedback($pThis, $iEffect)
	#forceref $pThis, $iEffect

;~ 	Local Const $iDataOffset = $PTR_LEN * 3 + 4
;~ 	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $iDataOffset)), 1)
;~ 	Local $tData = DllStructCreate($tagSourceObjIntData, $pData)

	Return $DRAGDROP_S_USEDEFAULTCURSORS
EndFunc   ;==>__Mthd_GiveFeedback

Func __Mthd_DragEnterTarget($pThis, $hTarget)
	Local Const $iDataOffset = $PTR_LEN * 2 + 4
	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $iDataOffset)), 1)
	Local $tData = DllStructCreate($tagSourceObjIntData, $pData)
	DllStructSetData($tData, "hTarget", $hTarget)

	Return $S_OK
EndFunc   ;==>__Mthd_DragEnterTarget

Func __Mthd_DragLeaveTarget($pThis)
	#forceref $pThis
	Local Const $iDataOffset = $PTR_LEN * 2 + 4
	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $iDataOffset)), 1)
	Local $tData = DllStructCreate($tagSourceObjIntData, $pData)
	DllStructSetData($tData, "hTarget", 0)

	Return $S_OK
EndFunc   ;==>__Mthd_DragLeaveTarget
