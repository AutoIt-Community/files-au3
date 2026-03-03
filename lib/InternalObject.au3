#include-once
#include "ProjectConstants.au3"
#include "IUnknown.au3"
#include "IFileOperation.au3"

Global $__g_aObjects[1][20]
Global $__g_hMthd_QueryInterface, $__g_hMthd_AddRef, $__g_hMthd_Release
Global $__g_hMthd_QueryInterfaceThunk, $__g_hMthd_AddRefThunk, $__g_hMthd_ReleaseThunk

Func PrepareInternalObject($iImplIfaces)
	If Not $__g_hMthd_QueryInterface Then
		$__g_hMthd_QueryInterface = DllCallbackRegister("__Mthd_QueryInterface", "long", "ptr;ptr;ptr")
		$__g_hMthd_AddRef = DllCallbackRegister("__Mthd_AddRef", "long", "ptr")
		$__g_hMthd_Release = DllCallbackRegister("__Mthd_Release", "long", "ptr")
		$__g_hMthd_QueryInterfaceThunk = DllCallbackRegister("__Mthd_QueryInterfaceThunk", "long", "ptr;ptr;ptr")
		$__g_hMthd_AddRefThunk = DllCallbackRegister("__Mthd_AddRefThunk", "long", "ptr")
		$__g_hMthd_ReleaseThunk = DllCallbackRegister("__Mthd_ReleaseThunk", "long", "ptr")
	EndIf

	Local $iObjectId = UBound($__g_aObjects)
	ReDim $__g_aObjects[$iObjectId + 1][UBound($__g_aObjects, 2)]
	$__g_aObjects[0][0] += 1

	Local $tIUnknownVTab = DllStructCreate("ptr pFunc[3]")
	$tIUnknownVTab.pFunc(1) = DllCallbackGetPtr($__g_hMthd_QueryInterface)
	$tIUnknownVTab.pFunc(2) = DllCallbackGetPtr($__g_hMthd_AddRef)
	$tIUnknownVTab.pFunc(3) = DllCallbackGetPtr($__g_hMthd_Release)

	Local $tagSupportedIIDs
	For $i = 1 To $iImplIfaces
		$tagSupportedIIDs &= "byte[16];"
	Next
	Local $tSupportedIIDs = DllStructCreate($tagSupportedIIDs)
	_WinAPI_GUIDFromStringEx($sIID_IUnknown, DllStructGetPtr($tSupportedIIDs, 1))

	Local $tagObject = StringFormat("align 4;int iImplIfaces;ptr pVTab[%d];int iRefCnt;" & _
			"ptr pSupportedIIDs;ptr pData", $iImplIfaces)

	Local $tObject = DllStructCreate($tagObject)
	$tObject.iRefCnt = 1
	$tObject.iImplIfaces = $iImplIfaces
	$tObject.pVTab(1) = DllStructGetPtr($tIUnknownVTab)
	$tObject.pSupportedIIDs = DllStructGetPtr($tSupportedIIDs)

	Local $pObject = DllStructGetPtr($tObject, "pVTab")

	$__g_aObjects[$iObjectId][0] = $pObject
	$__g_aObjects[$iObjectId][1] = $tObject
	$__g_aObjects[$iObjectId][2] = $tSupportedIIDs
	$__g_aObjects[$iObjectId][3] = $tIUnknownVTab

	Return $iObjectId
EndFunc

Func DestroyInternalObject($pObject)
	If (Not $pObject) Or (Not IsPtr($pObject)) Then Return SetError($ERROR_INVALID_PARAMETER, 0, False)

	For $i = 0 To UBound($__g_aObjects) - 1
		If $__g_aObjects[$i][0] = $pObject Then ExitLoop
	Next
	If $i = UBound($__g_aObjects) Then Return SetError($ERROR_INVALID_PARAMETER, 0, False)

	For $j = 0 To UBound($__g_aObjects, 2) - 1
		$__g_aObjects[$i][$j] = 0
	Next
	$__g_aObjects[0][0] -= 1

	If Not $__g_aObjects[0][0] Then
		DllCallbackFree($__g_hMthd_QueryInterface)
		DllCallbackFree($__g_hMthd_AddRef)
		DllCallbackFree($__g_hMthd_Release)
		DllCallbackFree($__g_hMthd_QueryInterfaceThunk)
		DllCallbackFree($__g_hMthd_AddRefThunk)
		DllCallbackFree($__g_hMthd_ReleaseThunk)

		$__g_hMthd_QueryInterface = 0
		$__g_hMthd_AddRef = 0
		$__g_hMthd_Release = 0
		$__g_hMthd_QueryInterfaceThunk = 0
		$__g_hMthd_AddRefThunk = 0
		$__g_hMthd_ReleaseThunk = 0
	EndIf
EndFunc   ;==>DestroyDropSource


#Region Internal IUnknown Methods

Func __Mthd_QueryInterface($pThis, $pIID, $ppObj)
;~ 	ConsoleWrite("QI: " & _WinAPI_StringFromGUID($pIID) & @CRLF)

	Local $hResult = $S_OK

	Local $iIIDCnt = DllStructGetData(DllStructCreate("int", Ptr($pThis - 4)), 1)
	Local $tThis = DllStructCreate(StringFormat("align 4;ptr pVTab[%d];int iRefCnt;ptr pSupportedIIDs", $iIIDCnt), $pThis)
	Local $pTestIID = $tThis.pSupportedIIDs

	If Not $ppObj Then
		$hResult = $E_POINTER
	Else
		For $i = 0 To $iIIDCnt - 1
			If _WinAPI_StringFromGUID($pIID) = _WinAPI_StringFromGUID(Ptr($pTestIID)) Then
				DllStructSetData(DllStructCreate("ptr", $ppObj), 1, Ptr($pThis + $i * $PTR_LEN))
;~ 				ConsoleWrite("FOUND: " & $pThis + $i * $PTR_LEN & @CRLF)

				__Mthd_AddRef($pThis)
				ExitLoop
			EndIf
			$pTestIID += 16
		Next
		If $i = $iIIDCnt Then $hResult = $E_NOINTERFACE

	EndIf

	Return $hResult
EndFunc   ;==>__Mthd_QueryInterface

Func __Mthd_AddRef($pThis)

	Local $iImplIfaces = DllStructGetData(DllStructCreate("int", Ptr($pThis - 4)), 1)
	Local $tThis = DllStructCreate(StringFormat("align 4;ptr pVTab[%d];int iRefCnt", $iImplIfaces), $pThis)
	$tThis.iRefCnt += 1

	Return $tThis.iRefCnt
EndFunc   ;==>__Mthd_AddRef

Func __Mthd_Release($pThis)
	Local $iImplIfaces = DllStructGetData(DllStructCreate("int", Ptr($pThis - 4)), 1)
	Local $tThis = DllStructCreate(StringFormat("align 4;ptr pVTab[%d];int iRefCnt", $iImplIfaces), $pThis)
	$tThis.iRefCnt -= 1
	Return $tThis.iRefCnt
EndFunc   ;==>__Mthd_Release

Func __Mthd_QueryInterfaceThunk($pThis, $pIID, $ppObj)
	Local $hResult = $S_OK
	Local $tIID = DllStructCreate($tagGUID, $pIID)
	If _WinAPI_StringFromGUID($tIID) = $sIID_IUnknown Then
		DllStructSetData(DllStructCreate("ptr", $ppObj), 1, $pThis)
	Else
		$pThis = Ptr($pThis - $PTR_LEN)
		Local $pVTab = DllStructGetData(DllStructCreate("ptr", $pThis), 1)
		Local $pFunc = DllStructGetData(DllStructCreate("ptr", $pVTab), 1)
		Local $aCall = DllCallAddress("long", $pFunc, "ptr", $pThis, "ptr", $pIID, "ptr", $ppObj)
		$hResult = $aCall[0]
	EndIf
	Return $hResult
EndFunc   ;==>__Mthd_QueryInterfaceThunk

Func __Mthd_AddRefThunk($pThis)
	$pThis = Ptr($pThis - $PTR_LEN)
	Return _AddRef($pThis)
EndFunc   ;==>__Mthd_AddRefThunk

Func __Mthd_ReleaseThunk($pThis)
	$pThis = Ptr($pThis - $PTR_LEN)
	Return _Release($pThis)
EndFunc   ;==>__Mthd_ReleaseThunk
#EndRegion Internal IUnknown Methods
