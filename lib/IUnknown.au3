#include-once
#include "ProjectConstants.au3"
#include <WinAPIConv.au3>

;This allows us to call the IUnknown methods directly from an object pointer.
;(You could also use ObjCreateInterface to create an object type from the ptr).

Func _QueryInterface($pThis, $sIID)
	If (Not $pThis) Or (Not IsPtr($pThis)) Then Return SetError($ERROR_INVALID_PARAMETER)
	Local $pVTab = DllStructGetData(DllStructCreate("ptr", $pThis), 1)
	Local $pFunc = DllStructGetData(DllStructCreate("ptr", $pVTab), 1)
	Local $tIID = _WinAPI_GUIDFromString($sIID)
	Local $aCall = DllCallAddress("long", $pFunc, "ptr", $pThis, "struct*", $tIID, "ptr*", 0)
	Return SetError($aCall[0], 0, $aCall[3])
EndFunc   ;==>_QueryInterface

Func _AddRef($pThis)
	If (Not $pThis) Or (Not IsPtr($pThis)) Then Return SetError($ERROR_INVALID_PARAMETER)
	Local Const $PTR_LEN = @AutoItX64 ? 8 : 4
	Local $pVTab = DllStructGetData(DllStructCreate("ptr", $pThis), 1)
	Local $pFunc = DllStructGetData(DllStructCreate("ptr", $pVTab + $PTR_LEN), 1)
	Local $aCall = DllCallAddress("uint", $pFunc, "ptr", $pThis)
	Return $aCall[0]
EndFunc   ;==>_AddRef
