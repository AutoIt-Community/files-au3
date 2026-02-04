#include-once
#include <WinAPISysWin.au3>
#include <SendMessage.au3>
#include "InternalObject.au3"
#include "IUnknown.au3"
#include <GuiListView.au3>
#include <GuiTreeView.au3>
#include <StructureConstants.au3>
#include <memory.au3>
#include <clipboard.au3>

#include <String.au3>
#include "TreeListExplorer.au3"

Global $__g_iDropTargetCount
Global $__g_hMthd_DragEnter, $__g_hMthd_DragOver, $__g_hMthd_DragLeave, $__g_hMthd_Drop
Global $tagTargetObjIntData = "hwnd hTarget;bool bAcceptDrop;ptr pDataObject;ptr pIDropTgtHelper"
Global Const $__g_iTargetObjDataOffset = $PTR_LEN * 2 + 4

Global $hTreeOrig
Global $iPreviousHot

Func CreateDropTarget($hTarget = 0)
	$__g_iDropTargetCount += 1

	Local $iObjectId = PrepareInternalObject(2)
	Local $tObject = $__g_aObjects[$iObjectId][1]
	Local $tSupportedIIDs = $__g_aObjects[$iObjectId][2]

	If Not $__g_hMthd_DragEnter Then
		$__g_hMthd_DragEnter = DllCallbackRegister("__Mthd_DragEnter", "long", "ptr;ptr;dword;uint64;ptr")
		$__g_hMthd_DragOver = DllCallbackRegister("__Mthd_DragOver", "long", "ptr;dword;uint64;ptr")
		$__g_hMthd_DragLeave = DllCallbackRegister("__Mthd_DragLeave", "long", "ptr")
		$__g_hMthd_Drop = DllCallbackRegister("__Mthd_Drop", "long", "ptr;ptr;dword;uint64;ptr")
	EndIf

	Local $tIDropTgtVTab = DllStructCreate("ptr pFunc[7]")
	$tIDropTgtVTab.pFunc(1) = DllCallbackGetPtr($__g_hMthd_QueryInterfaceThunk)
	$tIDropTgtVTab.pFunc(2) = DllCallbackGetPtr($__g_hMthd_AddRefThunk)
	$tIDropTgtVTab.pFunc(3) = DllCallbackGetPtr($__g_hMthd_ReleaseThunk)
	$tIDropTgtVTab.pFunc(4) = DllCallbackGetPtr($__g_hMthd_DragEnter)
	$tIDropTgtVTab.pFunc(5) = DllCallbackGetPtr($__g_hMthd_DragOver)
	$tIDropTgtVTab.pFunc(6) = DllCallbackGetPtr($__g_hMthd_DragLeave)
	$tIDropTgtVTab.pFunc(7) = DllCallbackGetPtr($__g_hMthd_Drop)

	Local $tInternalData = DllStructCreate($tagTargetObjIntData)
	$tInternalData.hTarget = $hTarget

	$tObject.pVTab(2) = DllStructGetPtr($tIDropTgtVTab)
	$tObject.pData = DllStructGetPtr($tInternalData)
	_WinAPI_GUIDFromStringEx($sIID_IDropTarget, DllStructGetPtr($tSupportedIIDs, 2))

	$__g_aObjects[$iObjectId][4] = $tIDropTgtVTab
	$__g_aObjects[$iObjectId][5] = $tInternalData

	Local $oDropTgtHelper = ObjCreateInterface($sCLSID_DragDropHelper, $sIID_IDropTargetHelper, $tagIDropTargetHelper)

	Local $pIDropTgtHelper
	$oDropTgtHelper.QueryInterface($sIID_IDropTargetHelper, $pIDropTgtHelper)
	$tInternalData.pIDropTgtHelper = $pIDropTgtHelper

;~ 	ConsoleWrite("IUnknown Location: " & DllStructGetPtr($tObject, "pVTab") & @CRLF)
;~ 	ConsoleWrite("IDropTarget Location: " & DllStructGetPtr($tObject, "pVTab") + $PTR_LEN & @CRLF)

	$__g_aObjects[$iObjectId][0] = DllStructGetPtr($tObject, "pVTab") + $PTR_LEN
	Return $__g_aObjects[$iObjectId][0]
EndFunc   ;==>CreateDropTarget

Func DestroyDropTarget($pObject)
	If (Not $pObject) Or (Not IsPtr($pObject)) Then Return SetError($ERROR_INVALID_PARAMETER, 0, False)

	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pObject + $__g_iTargetObjDataOffset)), 1)
	Local $tData = DllStructCreate($tagTargetObjIntData, $pData)
	_Release($tData.pIDropTgtHelper)

	DestroyInternalObject($pObject)
	If Not @error Then
		$__g_iDropTargetCount -= 1
		If Not $__g_iDropTargetCount Then
			DllCallbackFree($__g_hMthd_DragEnter)
			DllCallbackFree($__g_hMthd_DragOver)
			DllCallbackFree($__g_hMthd_DragLeave)
			DllCallbackFree($__g_hMthd_Drop)

			$__g_hMthd_DragEnter = 0
			$__g_hMthd_DragOver = 0
			$__g_hMthd_DragLeave = 0
			$__g_hMthd_Drop = 0
		EndIf
	EndIf
EndFunc   ;==>DestroyDropTarget

Func __Mthd_DragEnter($pThis, $pDataOject, $iKeyState, $iPoint, $piEffect)

	#forceref $pThis, $pDataOject, $iKeyState, $iPoint, $piEffect

	Local $sDirText = ""

	Local $tPoint = DllStructCreate($tagPoint)
	$tPoint.X = _WinAPI_LoDWord($iPoint)
	$tPoint.Y = _WinAPI_HiDWord($iPoint)

	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $__g_iTargetObjDataOffset)), 1)
	Local $tData = DllStructCreate($tagTargetObjIntData, $pData)
	$tData.bAcceptDrop = False
	$tData.pDataObject = $pDataOject

	Local $oDataObject = ObjCreateInterface($pDataOject, $sIID_IDataObject, $tagIDataObject)
	;Addref to counteract the automatic release() when $oDataObject falls out of scope.
	$oDataObject.AddRef()

	;Accept only if dataobject contains a HDrop.
	Local $pIEnumFmtEtc
	$oDataObject.EnumFormatEtc($DATADIR_GET, $pIEnumFmtEtc)
	Local $oIEnumFmtEtc = ObjCreateInterface($pIEnumFmtEtc, $sIID_IEnumFORMATETC, $tagIEnumFORMATETC)
	Local $tFormatEtc = DllStructCreate($tagFORMATETC)
	Local $pFmtEtc = DllStructGetPtr($tFormatEtc), $iFetched
	$oIEnumFmtEtc.Reset()
	While $oIEnumFmtEtc.Next(1, $pFmtEtc, $iFetched) = $S_OK
;~ 		ConsoleWrite(Hex($tFormatEtc.cfFormat) & " " & _ClipBoard_GetFormatName($tFormatEtc.cfFormat) &  @CRLF)
		If $tFormatEtc.cfFormat = $CF_HDROP Then
			$tData.bAcceptDrop = True
			ExitLoop
		EndIf
	WEnd

	Switch _WinAPI_GetClassName($tData.hTarget)
		Case $WC_LISTVIEW
			;
		Case $WC_TREEVIEW
			$hTreeOrig = _GUICtrlTreeView_GetSelection($tData.hTarget)
		Case Else
			$tData.bAcceptDrop = False
	EndSwitch

	__DoDropResponse($tData, $iKeyState, $tPoint, $piEffect, $sDirText)

	Local $oIDropTgtHelper = ObjCreateInterface($tData.pIDropTgtHelper, $sIID_IDropTargetHelper, $tagIDropTargetHelper)
	$oIDropTgtHelper.AddRef()

	$oIDropTgtHelper.DragEnter($tData.hTarget, $pDataOject, $tPoint, $piEffect)

	Return $S_OK
EndFunc   ;==>__Mthd_DragEnter

Func __Mthd_DragOver($pThis, $iKeyState, $iPoint, $piEffect)

	#forceref $pThis, $iKeyState, $iPoint, $piEffect

	Local $sDirText = ""
	Local $bIsFolder

	Local $tPoint = DllStructCreate($tagPoint)
	$tPoint.X = _WinAPI_LoDWord($iPoint)
	$tPoint.Y = _WinAPI_HiDWord($iPoint)

	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $__g_iTargetObjDataOffset)), 1)
	Local $tData = DllStructCreate($tagTargetObjIntData, $pData)

	_WinAPI_ScreenToClient($tData.hTarget, $tPoint)

	Switch _WinAPI_GetClassName($tData.hTarget)
		Case $WC_LISTVIEW
			Local $aListItem = _GUICtrlListView_HitTest($tData.hTarget, $tPoint.X, $tPoint.Y)
			Local $sItemText = _GUICtrlListView_GetItemText($tData.hTarget, $aListItem[0])
			Local $sFullPath = __TreeListExplorer_GetPath(1) & $sItemText
			If StringInStr(FileGetAttrib($sFullPath), "D") Then
				$sDirText = $sItemText
				$bIsFolder = True
			Else
				; get the currently selected path
				$sDirPath = __TreeListExplorer_GetPath(1)
				; obtain folder name only for drag tooltip
				Local $aPath = _StringBetween($sDirPath, "\", "\")
				$sDirText = $aPath[UBound($aPath) - 1]
			EndIf
			; handle case when cursor is over ListView but not on item; need to show current directory as drop dir
			If $sDirText = "" Then
				; get the currently selected path
				$sDirPath = __TreeListExplorer_GetPath(1)
				; obtain folder name only for drag tooltip
				Local $aPath = _StringBetween($sDirPath, "\", "\")
				$sDirText = $aPath[UBound($aPath) - 1]
			EndIf

			; clear previously DROPHILITED listview item
			_GUICtrlListView_SetItemState($tData.hTarget, $iPreviousHot, 0, $LVIS_DROPHILITED)

			If $aListItem[0] >= 0 Then
				$iPreviousHot = $aListItem[0]
				; bring focus to listview to show hot item (needed for listview to listview drag)
				_WinAPI_SetFocus($tData.hTarget)
				;_GUICtrlListView_SetHotItem($tData.hTarget, $aListItem[0])
				If $bIsFolder Then
					_GUICtrlListView_SetItemState($tData.hTarget, $aListItem[0], $LVIS_DROPHILITED, $LVIS_DROPHILITED)
				EndIf
			Else
				; clear previously DROPHILITED listview item
				_GUICtrlListView_SetItemState($tData.hTarget, $iPreviousHot, 0, $LVIS_DROPHILITED)
			EndIf
		Case $WC_TREEVIEW
			;Local $tMPos = _WinAPI_GetMousePos(True, $tData.hTarget)
			;Local $tMPos = _WinAPI_GetMousePos()
			Local $hTreeItem = _GUICtrlTreeView_HitTestItem($tData.hTarget, $tPoint.X, $tPoint.Y)
			$sDirText = _GUICtrlTreeView_GetText($tData.hTarget, $hTreeItem)
			If $hTreeItem <> 0 Then
				; bring focus to treeview to properly show DROPHILITE
				_WinAPI_SetFocus($tData.hTarget)
				_GUICtrlTreeView_SelectItem($tData.hTarget, $hTreeItem)
				_GUICtrlTreeView_SetState($tData.hTarget, $hTreeOrig, $TVIS_SELECTED, True)
			EndIf
		Case Else
			$sDirText = ""
	EndSwitch

	__DoDropResponse($tData, $iKeyState, $tPoint, $piEffect, $sDirText)

	Local $oIDropTgtHelper = ObjCreateInterface($tData.pIDropTgtHelper, $sIID_IDropTargetHelper, $tagIDropTargetHelper)
	$oIDropTgtHelper.AddRef()
	$oIDropTgtHelper.DragOver($tPoint, $piEffect)

	Return $S_OK
EndFunc   ;==>__Mthd_DragOver

Func __Mthd_DragLeave($pThis)
	#forceref $pThis
	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $__g_iTargetObjDataOffset)), 1)
	Local $tData = DllStructCreate($tagTargetObjIntData, $pData)

	__SetDropDescription($tData.pDataObject, $DROPIMAGE_INVALID)

	Local $oIDropTgtHelper = ObjCreateInterface($tData.pIDropTgtHelper, $sIID_IDropTargetHelper, $tagIDropTargetHelper)
	$oIDropTgtHelper.AddRef()
	$oIDropTgtHelper.DragLeave()

	Switch _WinAPI_GetClassName($tData.hTarget)
		Case $WC_LISTVIEW
			; clear previously DROPHILITED listview item
			_GUICtrlListView_SetItemState($tData.hTarget, $iPreviousHot, 0, $LVIS_DROPHILITED)
		Case $WC_TREEVIEW
			; restore original treeview selection if cursor leaves treeview
			_GUICtrlTreeView_SelectItem($tData.hTarget, $hTreeOrig)
	EndSwitch

	$tData.bAcceptDrop = False
	$tData.pDataObject = 0

	Return $S_OK
EndFunc   ;==>__Mthd_DragLeave

Func __Mthd_Drop($pThis, $pDataOject, $iKeyState, $iPoint, $piEffect)
	#forceref $pThis, $iKeyState, $iPoint, $piEffect

	Local $sDirText = ""

	Local $tPoint = DllStructCreate($tagPoint)
	$tPoint.X = _WinAPI_LoDWord($iPoint)
	$tPoint.Y = _WinAPI_HiDWord($iPoint)

	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $__g_iTargetObjDataOffset)), 1)
	Local $tData = DllStructCreate($tagTargetObjIntData, $pData)

	Local $vItem = __DoDropResponse($tData, $iKeyState, $tPoint, $piEffect, $sDirText)
	If @extended Then $vItem += 1

	Local $oIDropTgtHelper = ObjCreateInterface($tData.pIDropTgtHelper, $sIID_IDropTargetHelper, $tagIDropTargetHelper)
	$oIDropTgtHelper.AddRef()
	$oIDropTgtHelper.Drop($pDataOject, $tPoint, $piEffect)

	Local $tEffect = DllStructCreate("dword iEffect", $piEffect)
	If $tEffect.iEffect <> $DROPEFFECT_NONE Then

		Local $tFormatEtc = DllStructCreate($tagFORMATETC)
		$tFormatEtc.cfFormat = $CF_HDROP
		$tFormatEtc.iIndex = -1
		$tFormatEtc.tymed = $TYMED_HGLOBAL

		Local $oDataObject = ObjCreateInterface($pDataOject, $sIID_IDataObject, $tagIDataObject)
		Local $tStgMedium = DllStructCreate($tagSTGMEDIUM)
		$oDataObject.AddRef()
		$oDataObject.GetData($tFormatEtc, $tStgMedium)

		Local $asFilenames = _WinAPI_DragQueryFileEx($tStgMedium.handle)

		Switch _WinAPI_GetClassName($tData.hTarget)
			Case $WC_LISTVIEW
				; clear previously DROPHILITED listview item
				_GUICtrlListView_SetItemState($tData.hTarget, $iPreviousHot, 0, $LVIS_DROPHILITED)
				For $i = 1 To $asFilenames[0]
					;_GUICtrlListView_InsertItem($tData.hTarget, $asFilenames[$i], $vItem)
					$vItem += 1
				Next
				;_GUICtrlListView_SetInsertMark($tData.hTarget, -1)

			Case $WC_TREEVIEW
				Local $hInsAfter = _GUICtrlTreeView_GetPrev($tData.hTarget, $vItem)
				If Not $hInsAfter Then $hInsAfter = $TVI_FIRST
				For $i = 1 To $asFilenames[0]
					;$hInsAfter = _GUICtrlTreeView_InsertItem($tData.hTarget, $asFilenames[$i], 0, $hInsAfter)
				Next
				;_GUICtrlTreeView_SetInsertMark($tData.hTarget, 0)
		EndSwitch
	EndIf

	;__SetPerformedDropEffect($pDataOject, $piEffect)

	$tData.bAcceptDrop = False
	$tData.pDataObject = 0

	Return $S_OK
EndFunc   ;==>__Mthd_Drop

Func __DoInsertMark($hTarget, $tPoint)
	Return
EndFunc

Func __DoInsertMark_orig($hTarget, $tPoint)
	Local $vItem = -1, $bAfter = False
	_WinAPI_ScreenToClient($hTarget, $tPoint)

	Switch _WinAPI_GetClassName($hTarget)
		Case $WC_LISTVIEW
			Local $tLVINSERTMARK = DllStructCreate($tagLVINSERTMARK)
			$tLVINSERTMARK.Size = DllStructGetSize($tLVINSERTMARK)
			If _SendMessage($hTarget, $LVM_INSERTMARKHITTEST, $tPoint, $tLVINSERTMARK, 0, "struct*", "struct*") Then
				If _SendMessage($hTarget, $LVM_SETINSERTMARK, 0, $tLVINSERTMARK, 0, "wparam", "struct*") Then $vItem = $tLVINSERTMARK.Item
				$bAfter = BitAND($tLVINSERTMARK.Flags, $LVIM_AFTER)
			EndIf

		Case $WC_TREEVIEW
			Local $tTVHITTESTINFO = DllStructCreate($tagTVHITTESTINFO)
			$tTVHITTESTINFO.X = $tPoint.X
			$tTVHITTESTINFO.Y = $tPoint.Y
			If _SendMessage($hTarget, $TVM_HITTEST, $tPoint, $tTVHITTESTINFO, 0, "struct*", "struct*") Then
				If _SendMessage($hTarget, $TVM_SETINSERTMARK, 0, $tTVHITTESTINFO.Item, 0, "wparam", "handle") Then $vItem = $tTVHITTESTINFO.Item
			EndIf

	EndSwitch

	Return SetExtended($bAfter, $vItem)
EndFunc

Func __SetDropDescription($pDataObject, $iType, $sMessage = "", $sInsert = "")
	Local $tFormatEtc = DllStructCreate($tagFORMATETC)
	$tFormatEtc.cfFormat = _ClipBoard_RegisterFormat($CFSTR_DROPDESCRIPTION)
	$tFormatEtc.ptd = 0
	$tFormatEtc.aspect = $DVASPECT_CONTENT
	$tFormatEtc.index = -1
	$tFormatEtc.tymed = $TYMED_HGLOBAL

	Local $tDropDesc = DllStructCreate($tagDROPDESCRIPTION)
	Local $hGblMem = _MemGlobalAlloc(DllStructGetSize($tDropDesc), $GPTR)
	Local $pDropDesc = _MemGlobalLock($hGblMem)
	$tDropDesc = DllStructCreate($tagDROPDESCRIPTION, $pDropDesc)
	$tDropDesc.iType = $iType
	$tDropDesc.sMessage = $sMessage
	$tDropDesc.sInsert = $sInsert
	_MemGlobalUnlock($hGblMem)

	Local $tStgMedium = DllStructCreate($tagSTGMEDIUM)
	$tStgMedium.tymed = $TYMED_HGLOBAL
	$tStgMedium.handle = $hGblMem
	$tStgMedium.pUnkForRelease = 0

	Local $oDataObj = ObjCreateInterface($pDataObject, $sIID_IDataObject, $tagIDataObject)
	$oDataObj.AddRef()
	$oDataObj.SetData($tFormatEtc, $tStgMedium, 1)
EndFunc

Func __DoDropResponse($tData, $iKeyState, $tPoint, $piEffect, $sDirText)

	#forceref $tData, $iKeyState, $tPoint, $piEffect

	Local $vItem, $bAfter
	Local $iRetEffect = $DROPEFFECT_NONE
	Local $tEffect = DllStructCreate("dword iEffect", $piEffect)

	;Only Accept copy ops (will expand this at some point!)

	If $tData.bAcceptDrop Then
		If BitAND($tEffect.iEffect, $DROPEFFECT_MOVE) Then $iRetEffect = $DROPEFFECT_MOVE
		If BitAND($tEffect.iEffect, $DROPEFFECT_COPY) Then
			If BitAND($iKeyState, $MK_CONTROL) Or $iRetEffect = $DROPEFFECT_NONE Then
				$iRetEffect = $DROPEFFECT_COPY
			EndIf
		EndIf
		If BitAND($tEffect.iEffect, $DROPEFFECT_LINK) Then
			If BitAND($iKeyState, $MK_ALT) Or $iRetEffect = $DROPEFFECT_NONE Then
				$iRetEffect = $DROPEFFECT_LINK
			EndIf
		EndIf
	EndIf

	$tEffect.iEffect = $iRetEffect
	If $tEffect.iEffect <> $DROPEFFECT_NONE Then
		;$vItem = __DoInsertMark($tData.hTarget, $tPoint)
		$bAfter = @extended
	Else
		Switch _WinAPI_GetClassName($tData.hTarget)
			Case $WC_LISTVIEW
				;_GUICtrlListView_SetInsertMark($tData.hTarget, -1)
				$vItem = -1

			Case $WC_TREEVIEW
				;_GUICtrlTreeView_SetInsertMark($tData.hTarget, 0)
				$vItem = 0

		EndSwitch
	EndIf

	Switch $tEffect.iEffect
		Case $DROPEFFECT_NONE
			__SetDropDescription($tData.pDataObject, $DROPIMAGE_NONE, "No Op %1", $sDirText)

		Case $DROPEFFECT_LINK
			__SetDropDescription($tData.pDataObject, $DROPIMAGE_LINK, "Link to %1", $sDirText)

		Case $DROPEFFECT_COPY
			__SetDropDescription($tData.pDataObject, $DROPIMAGE_COPY, "Copy to %1", $sDirText)

		Case $DROPEFFECT_MOVE
			__SetDropDescription($tData.pDataObject, $DROPIMAGE_MOVE, "Move to %1", $sDirText)

	EndSwitch

	Return SetExtended($bAfter, $vItem)
EndFunc

Func __SetPerformedDropEffect($pDataObject, $iDropEffect)
    Local $tFormatEtc = DllStructCreate($tagFORMATETC)
    $tFormatEtc.cfFormat = _ClipBoard_RegisterFormat($CFSTR_PERFORMEDDROPEFFECT)
    $tFormatEtc.ptd = 0
    $tFormatEtc.aspect = $DVASPECT_CONTENT
    $tFormatEtc.index = -1
    $tFormatEtc.tymed = $TYMED_HGLOBAL

    Local $hGblMem = _MemGlobalAlloc(4, $GPTR)
    Local $pDropEffect = _MemGlobalLock($hGblMem)
    Local $tDropEffect = DllStructCreate("dword iEffect", $pDropEffect)
    $tDropEffect.iEffect = $iDropEffect
    _MemGlobalUnlock($hGblMem)

    Local $tStgMedium = DllStructCreate($tagSTGMEDIUM)
    $tStgMedium.tymed = $TYMED_HGLOBAL
    $tStgMedium.handle = $hGblMem
    $tStgMedium.pUnkForRelease = 0

    Local $oDataObj = ObjCreateInterface($pDataObject, $sIID_IDataObject, $tagIDataObject)
    $oDataObj.AddRef()
    $oDataObj.SetData($tFormatEtc, $tStgMedium, 1)
EndFunc
