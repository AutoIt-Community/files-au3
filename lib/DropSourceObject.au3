#include-once
#include "ProjectConstants.au3"
#include "IUnknown.au3"

#include <WinAPISysWin.au3>
#include <WinAPIMisc.au3>
#include <GuiTreeView.au3>

Global $g_hListview, $g_hTreeView, $sTargetCtrl, $hTLESystem

; In this file we're probably just interested in these methods.  The rest is just stuff to make our internal object work!
; __Mthd_QueryContinueDrag, __Mthd_GiveFeedback
; __Mthd_DragEnterTarget, __Mthd_DragLeaveTarget

; __Mthd_QueryContinueDrag is called in a loop - it asks us if we want to continue the drag/drop.
; And we say:
; "drop!" if the left and right buttons are up.
; "cancel!" if the esc key is pressed.
; otherwise "Continue looping!"

; __Mthd_GiveFeedback provides us with the oportunity to update the mouse cursor.
; We're delegating this to the system by returning DRAGDROP_S_USEDEFAULTCURSORS

; __Mthd_DragEnterTarget & __Mthd_DragLeaveTarget lets us know what we're currently hovering over. (as a window handle)
; Just in case you want to do anything with that info!.

;--------------------------------------------------------------------------------------------------------------
; Our DropSource Object in memory...
;--------------------------------------------------------------------------------------------------------------
;
; The left most column aren't part of the object - they're interface pointers.
; All of them are valid "object pointers" for our object. Regardless, some functions require us to provide a ptr to a specific interface.
;
; The "Thunk" methods are stubs, and redirect to the methods on pIUnkVtab.
; IIDs in aSupportedIIDs[] must be in the same order as the vtables in our layout.
;
;
;                   +- iIfaceCount
;                   |
;  pIUnknown----->  +- pIUnkVtab -------->  +- pQueryInterface
;                   |                       +- pAddRef
;                   |                       +- pRelease
;                   |
;  pIDropSource ->  += pIDrpSrcVTab ----->  +- pQueryInterfaceThunk
;                   |                       +- pAddRefThunk
;                   |                       +- pReleaseThunk
;                   |                       +- pQueryContinueDrag
;                   |                       +- pGiveFeedback
;                   |
;  pIDrpSrcNtfy ->  += pIDrpSrcNtfyVTab-->  +- pQueryInterfaceThunk
;                   |                       +- pAddRefThunk
;                   |                       +- pReleaseThunk
;                   |                       +- pDragEnterTarget
;                   |                       +- pDragLeaveTarget
;                   +- iRefCount
;                   |
;                   +- aSupportedIIDs[IID_IUnknown, IID_IDropSource, IID_IDropSourceNotify]
;                   |                   |
;                   +- hTarget
;
;--------------------------------------------------------------------------------------------------------------



Global $__g_aObjects[1][20]
Global $__g_hMthd_QueryInterface, $__g_hMthd_AddRef, $__g_hMthd_Release
Global $__g_hMthd_QueryInterfaceThunk, $__g_hMthd_AddRefThunk, $__g_hMthd_ReleaseThunk
Global $__g_hMthd_QueryContinueDrag, $__g_hMthd_GiveFeedback
Global $__g_hMthd_DragEnterTarget, $__g_hMthd_DragLeaveTarget

Func CreateDropSource()
	If Not $__g_hMthd_QueryInterface Then
		$__g_hMthd_QueryInterface = DllCallbackRegister("__Mthd_QueryInterface", "long", "ptr;ptr;ptr")
		$__g_hMthd_AddRef = DllCallbackRegister("__Mthd_AddRef", "long", "ptr")
		$__g_hMthd_Release = DllCallbackRegister("__Mthd_Release", "long", "ptr")

		$__g_hMthd_QueryInterfaceThunk = DllCallbackRegister("__Mthd_QueryInterfaceThunk", "long", "ptr;ptr;ptr")
		$__g_hMthd_AddRefThunk = DllCallbackRegister("__Mthd_AddRefThunk", "long", "ptr")
		$__g_hMthd_ReleaseThunk = DllCallbackRegister("__Mthd_ReleaseThunk", "long", "ptr")
	EndIf

	If Not $__g_hMthd_QueryContinueDrag Then
		$__g_hMthd_QueryContinueDrag = DllCallbackRegister("__Mthd_QueryContinueDrag", "long", "ptr;bool;dword")
		$__g_hMthd_GiveFeedback = DllCallbackRegister("__Mthd_GiveFeedback", "long", "ptr;dword")
	EndIf

	If Not $__g_hMthd_DragEnterTarget Then
		$__g_hMthd_DragEnterTarget = DllCallbackRegister("__Mthd_DragEnterTarget", "long", "ptr;hwnd")
		$__g_hMthd_DragLeaveTarget = DllCallbackRegister("__Mthd_DragLeaveTarget", "long", "ptr")
	EndIf

	Local $iObjectId = UBound($__g_aObjects)
	ReDim $__g_aObjects[$iObjectId + 1][UBound($__g_aObjects, 2)]
	$__g_aObjects[0][0] += 1

	Local $tUnknownVTab = DllStructCreate("ptr pFunc[3]")
	$tUnknownVTab.pFunc(1) = DllCallbackGetPtr($__g_hMthd_QueryInterface)
	$tUnknownVTab.pFunc(2) = DllCallbackGetPtr($__g_hMthd_AddRef)
	$tUnknownVTab.pFunc(3) = DllCallbackGetPtr($__g_hMthd_Release)

	Local $tDropSrcVTab = DllStructCreate("ptr pFunc[5]")
	$tDropSrcVTab.pFunc(1) = DllCallbackGetPtr($__g_hMthd_QueryInterfaceThunk)
	$tDropSrcVTab.pFunc(2) = DllCallbackGetPtr($__g_hMthd_AddRefThunk)
	$tDropSrcVTab.pFunc(3) = DllCallbackGetPtr($__g_hMthd_ReleaseThunk)
	$tDropSrcVTab.pFunc(4) = DllCallbackGetPtr($__g_hMthd_QueryContinueDrag)
	$tDropSrcVTab.pFunc(5) = DllCallbackGetPtr($__g_hMthd_GiveFeedback)

	Local $tDropSrcNotifyVTab = DllStructCreate("ptr pFunc[5]")
	$tDropSrcNotifyVTab.pFunc(1) = DllCallbackGetPtr($__g_hMthd_QueryInterfaceThunk)
	$tDropSrcNotifyVTab.pFunc(2) = DllCallbackGetPtr($__g_hMthd_AddRefThunk)
	$tDropSrcNotifyVTab.pFunc(3) = DllCallbackGetPtr($__g_hMthd_ReleaseThunk)
	$tDropSrcNotifyVTab.pFunc(4) = DllCallbackGetPtr($__g_hMthd_DragEnterTarget)
	$tDropSrcNotifyVTab.pFunc(5) = DllCallbackGetPtr($__g_hMthd_DragLeaveTarget)

	Local $tagObject = "align 4;int iImplIfaces;ptr pVTab[3];int iRefCnt;" & _
			"byte IID_IUnknown[16];" & _
			"byte IID_IDropSource[16];" & _
			"byte IID_IDropSourceNotify[16];" & _
			"hwnd hTarget"

	Local $tObject = DllStructCreate($tagObject)
	$tObject.pVTab(1) = DllStructGetPtr($tUnknownVTab)
	$tObject.pVTab(2) = DllStructGetPtr($tDropSrcVTab)
	$tObject.pVTab(3) = DllStructGetPtr($tDropSrcNotifyVTab)

	$tObject.iRefCnt = 1
	$tObject.iImplIfaces = 3
	_WinAPI_GUIDFromStringEx($sIID_IUnknown, DllStructGetPtr($tObject, "IID_IUnknown"))
	_WinAPI_GUIDFromStringEx($sIID_IDropSource, DllStructGetPtr($tObject, "IID_IDropSource"))
	_WinAPI_GUIDFromStringEx($sIID_IDropSourceNotify, DllStructGetPtr($tObject, "IID_IDropSourceNotify"))

	Local $pObject = DllStructGetPtr($tObject, "pVTab")

	$__g_aObjects[$iObjectId][0] = $pObject
	$__g_aObjects[$iObjectId][1] = $tObject
	$__g_aObjects[$iObjectId][2] = $tUnknownVTab
	$__g_aObjects[$iObjectId][3] = $tDropSrcVTab
	$__g_aObjects[$iObjectId][4] = $tDropSrcNotifyVTab

	Return $pObject
EndFunc   ;==>CreateDropSource

Func DestroyDropSource($pObject)
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
		DllCallbackFree($__g_hMthd_QueryContinueDrag)
		DllCallbackFree($__g_hMthd_GiveFeedback)
		DllCallbackFree($__g_hMthd_DragEnterTarget)
		DllCallbackFree($__g_hMthd_DragLeaveTarget)

		$__g_hMthd_QueryInterface = 0
		$__g_hMthd_AddRef = 0
		$__g_hMthd_Release = 0
		$__g_hMthd_QueryInterfaceThunk = 0
		$__g_hMthd_AddRefThunk = 0
		$__g_hMthd_ReleaseThunk = 0
		$__g_hMthd_QueryContinueDrag = 0
		$__g_hMthd_GiveFeedback = 0
		$__g_hMthd_DragEnterTarget = 0
		$__g_hMthd_DragLeaveTarget = 0
	EndIf
EndFunc   ;==>DestroyDropSource


#Region Internal IUnknown Methods

Func __Mthd_QueryInterface($pThis, $pIID, $ppObj)
	Local $hResult = $S_OK

	Local $iIIDCnt = DllStructGetData(DllStructCreate("int", Ptr($pThis - 4)), 1)
	Local $tThis = DllStructCreate(StringFormat("align 4;ptr pVTab[%d];int iRefCnt", $iIIDCnt), $pThis)

	Local $pTestIID = DllStructGetPtr($tThis, "iRefCnt") + 4
	If Not $ppObj Then
		$hResult = $E_POINTER
	Else
		For $i = 0 To $iIIDCnt - 1
			If _WinAPI_StringFromGUID($pIID) = _WinAPI_StringFromGUID(Ptr($pTestIID)) Then
				DllStructSetData(DllStructCreate("ptr", $ppObj), 1, Ptr($pThis + $i * $PTR_LEN))
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

#Region Internal IDataSource Methods

Func __Mthd_QueryContinueDrag($pThis, $bEscapePressed, $iKeyState)
	#forceref $pThis, $bEscapePressed, $iKeyState

;~  Key State values may be combined. Test using BitAND!
;~ 	If BitAND($MK_RBUTTON, $iKeyState) Then ...

;~  Key Constants:
;~  $MK_RBUTTON
;~  $MK_SHIFT
;~  $MK_CONTROL
;~  $MK_MBUTTON
;~  $MK_XBUTTON1
;~  $MK_XBUTTON2

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

;~ 	Drop effect values may be combined. Test using BitAND!
;~ 	If BitAND($DROPEFFECT_COPY, $iEffect) Then ...

;~ 	Effect Constants...
;~ 	$DROPEFFECT_NONE
;~ 	$DROPEFFECT_COPY
;~ 	$DROPEFFECT_MOVE
;~ 	$DROPEFFECT_LINK
;~ 	$DROPEFFECT_SCROLL

	Select
		Case $sTargetCtrl = 'Tree'
			; use __TreeListExplorer_GetPath($hTLESystem) to obtain initial tree path
			; reset selection back to original after
			Local $hItemHover = TreeItemFromPoint2($g_hTreeView)
			_WinAPI_SetFocus($g_hTreeView)
			_GUICtrlTreeView_SelectItem($g_hTreeView, $hItemHover)
			_GUICtrlTreeView_SetState($g_hTreeView, __TreeListExplorer_GetPath($hTLESystem), $TVIS_SELECTED, True)
		Case $sTargetCtrl = 'List'
			;
		Case Else
			; clear
	EndSelect

	Return $DRAGDROP_S_USEDEFAULTCURSORS
EndFunc   ;==>__Mthd_GiveFeedback

#EndRegion Internal IDataSource Methods

#Region Internal IDataSourceNotify Methods
Func __Mthd_DragEnterTarget($pThis, $hTarget)
	Local Const $iDataOffset = 52 + $PTR_LEN
	Local $tData = DllStructCreate("align 4; hwnd hTarget", Ptr($pThis + $iDataOffset))
	DllStructSetData($tData, "hTarget", $hTarget)
	Local $hTargetCtrl = DllStructSetData($tData, "hTarget", $hTarget)
	If $hTargetCtrl = $g_hTreeView Then
		ConsoleWrite("drag target is treeview" & @CRLF)
		$sTargetCtrl = 'Tree'
	EndIf
	If $hTargetCtrl = $g_hListview Then
		ConsoleWrite("drag target is listview" & @CRLF)
		$sTargetCtrl = 'List'
	EndIf
	;ConsoleWrite("target: " & DllStructSetData($tData, "hTarget", $hTarget) & @CRLF)
	;ConsoleWrite("parent: " & _WinAPI_GetAncestor(DllStructSetData($tData, "hTarget", $hTarget), $GA_PARENT) & @CRLF)

	Return $S_OK
EndFunc   ;==>__Mthd_DragEnterTarget

Func __Mthd_DragLeaveTarget($pThis)
	Local Const $iDataOffset = 52 + $PTR_LEN
	Local $tData = DllStructCreate("align 4; hwnd hTarget", Ptr($pThis + $iDataOffset))
	DllStructSetData($tData, "hTarget", 0)
	$sTargetCtrl = ''

	Return $S_OK
EndFunc   ;==>__Mthd_DragLeaveTarget
#EndRegion Internal IDataSourceNotify Methods

Func TreeItemFromPoint2($hWnd)
	Local $tMPos = _WinAPI_GetMousePos(True, $hWnd)
	Return _GUICtrlTreeView_HitTestItem($hWnd, DllStructGetData($tMPos, 1), DllStructGetData($tMPos, 2))
EndFunc   ;==>TreeItemFromPoint2
