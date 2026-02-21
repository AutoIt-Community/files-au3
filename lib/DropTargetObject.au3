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
#include "IFileOperation.au3"
#include "SharedFunctions.au3"
#include "../lib/SubFileOperations.au3"

Global $__g_iDropTargetCount
Global $__g_hMthd_DragEnter, $__g_hMthd_DragOver, $__g_hMthd_DragLeave, $__g_hMthd_Drop
Global $tagTargetObjIntData = "hwnd hTarget;bool bAcceptDrop;ptr pDataObject;ptr pIDropTgtHelper"
Global Const $__g_iTargetObjDataOffset = $PTR_LEN * 2 + 4

Global $hTreeOrig
Global $iPreviousHot, $sDropTV, $iFinalEffect
Global $sSourceDrive, $sSourcePath, $bIsSameDrive, $bIsSameFolder

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

Func __Mthd_DragEnter($pThis, $pDataObject, $iKeyState, $iPoint, $piEffect)

	#forceref $pThis, $pDataObject, $iKeyState, $iPoint, $piEffect

	Local $sDirText = ""

	Local $tPoint = DllStructCreate($tagPoint)
	$tPoint.X = _WinAPI_LoDWord($iPoint)
	$tPoint.Y = _WinAPI_HiDWord($iPoint)

	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $__g_iTargetObjDataOffset)), 1)
	Local $tData = DllStructCreate($tagTargetObjIntData, $pData)
	$tData.bAcceptDrop = False
	$tData.pDataObject = $pDataObject

	Local $oDataObject = ObjCreateInterface($pDataObject, $sIID_IDataObject, $tagIDataObject)
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

	; test getting filename to compare
	Local $tFormatEtc = DllStructCreate($tagFORMATETC)
	$tFormatEtc.cfFormat = $CF_HDROP
	$tFormatEtc.iIndex = -1
	$tFormatEtc.tymed = $TYMED_HGLOBAL

	Local $oDataObject = ObjCreateInterface($pDataObject, $sIID_IDataObject, $tagIDataObject)
	Local $tStgMedium = DllStructCreate($tagSTGMEDIUM)
	$oDataObject.AddRef()
	$oDataObject.GetData($tFormatEtc, $tStgMedium)

	Local $asFilenames = _WinAPI_DragQueryFileEx($tStgMedium.handle)
	Local $sPathName = $asFilenames[1]
	If StringInStr(FileGetAttrib($sPathName), "D") Then
		$sPathName = $sPathName & "\"
	EndIf
	Local $aPath = _PathSplit_mod($sPathName)
	$sSourceDrive = $aPath[$PATH_DRIVE]
	$sSourcePath = $aPath[$PATH_DRIVE] & $aPath[$PATH_DIRECTORY]
	; test

	Switch _WinAPI_GetClassName_mod($tData.hTarget)
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

	$oIDropTgtHelper.DragEnter($tData.hTarget, $pDataObject, $tPoint, $piEffect)

	Return $S_OK
EndFunc   ;==>__Mthd_DragEnter

Func __Mthd_DragOver($pThis, $iKeyState, $iPoint, $piEffect)

	#forceref $pThis, $iKeyState, $iPoint, $piEffect

	Local $sDirText = ""
	Local $bIsFolder
	Local $sDestDrive, $sDestPath

	Local $tPoint = DllStructCreate($tagPoint)
	$tPoint.X = _WinAPI_LoDWord($iPoint)
	$tPoint.Y = _WinAPI_HiDWord($iPoint)

	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $__g_iTargetObjDataOffset)), 1)
	Local $tData = DllStructCreate($tagTargetObjIntData, $pData)

	_WinAPI_ScreenToClient_mod($tData.hTarget, $tPoint)

	Switch _WinAPI_GetClassName_mod($tData.hTarget)
		Case $WC_LISTVIEW
			Local $aListItem = _GUICtrlListView_HitTest($tData.hTarget, $tPoint.X, $tPoint.Y)
			Local $sItemText = _GUICtrlListView_GetItemText($tData.hTarget, $aListItem[0])
			Local $sFullPath = __TreeListExplorer_GetPath(1) & $sItemText
			If Not $sItemText Then
				; get the currently selected path
				$sDirPath = __TreeListExplorer_GetPath(1)
				; obtain folder name only for drag tooltip
				Local $aPath = _StringBetween($sDirPath, "\", "\")
				$sDirText = $aPath[UBound($aPath) - 1]
				; get full path for comparison
				$sDestPath = $sDirPath
			ElseIf StringInStr(FileGetAttrib($sFullPath), "D") Then
				$sDirText = $sItemText
				$bIsFolder = True
				; get full path for comparison
				$sDestPath = $sFullPath & "\"
			Else
				; get the currently selected path
				$sDirPath = __TreeListExplorer_GetPath(1)
				; obtain folder name only for drag tooltip
				Local $aPath = _StringBetween($sDirPath, "\", "\")
				$sDirText = $aPath[UBound($aPath) - 1]
				; get full path for comparison
				$sDestPath = $sDirPath
			EndIf

			; clear previously DROPHILITED listview item
			_GUICtrlListView_SetItemState($tData.hTarget, $iPreviousHot, 0, $LVIS_DROPHILITED)

			If $aListItem[0] >= 0 Then
				$iPreviousHot = $aListItem[0]
				; bring focus to listview to show hot item (needed for listview to listview drag)
				_WinAPI_SetFocus($tData.hTarget)
				If $bIsFolder Then
					_GUICtrlListView_SetItemState($tData.hTarget, $aListItem[0], $LVIS_DROPHILITED, $LVIS_DROPHILITED)
				EndIf
			Else
				; clear previously DROPHILITED listview item
				_GUICtrlListView_SetItemState($tData.hTarget, $iPreviousHot, 0, $LVIS_DROPHILITED)
			EndIf
		Case $WC_TREEVIEW
			Local $hTreeItem = _GUICtrlTreeView_HitTestItem($tData.hTarget, $tPoint.X, $tPoint.Y)
			$sDirText = _GUICtrlTreeView_GetText($tData.hTarget, $hTreeItem)
			If $hTreeItem <> 0 Then
				; bring focus to treeview to properly show DROPHILITE
				_WinAPI_SetFocus($tData.hTarget)
				_GUICtrlTreeView_SelectItem($tData.hTarget, $hTreeItem)
				_GUICtrlTreeView_SetState($tData.hTarget, $hTreeOrig, $TVIS_SELECTED, True)
				; get full path for comparison
				$sDestPath = TreeItemToPath($tData.hTarget, $hTreeItem)
				$sDropTV = $sDestPath
			EndIf
		Case Else
			$sDirText = ""
	EndSwitch

	; compare source and target to determine if they are on the same drive
	If StringInStr($sDestPath, $sSourceDrive) Then
		$bIsSameDrive = True
	Else
		$bIsSameDrive = False
	EndIf

	; compare source and target to determine if they are the same
	If $sDestPath = $sSourcePath Then
		$bIsSameFolder = True
	Else
		$bIsSameFolder = False
	EndIf

	__DoDropResponse($tData, $iKeyState, $tPoint, $piEffect, $sDirText, $bIsSameDrive, $bIsSameFolder)

	Local $tEffect = DllStructCreate("dword iEffect", $piEffect)
	$iFinalEffect = $tEffect.iEffect

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

	Switch _WinAPI_GetClassName_mod($tData.hTarget)
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

Func __Mthd_Drop($pThis, $pDataObject, $iKeyState, $iPoint, $piEffect)
	#forceref $pThis, $iKeyState, $iPoint, $piEffect

	Local $sDirText = ""
	Local $bIsFolder, $iFlags, $sAction

	Local $tPoint = DllStructCreate($tagPoint)
	$tPoint.X = _WinAPI_LoDWord($iPoint)
	$tPoint.Y = _WinAPI_HiDWord($iPoint)

	Local $pData = DllStructGetData(DllStructCreate("ptr", Ptr($pThis + $__g_iTargetObjDataOffset)), 1)
	Local $tData = DllStructCreate($tagTargetObjIntData, $pData)

	_WinAPI_ScreenToClient_mod($tData.hTarget, $tPoint)

	__DoDropResponse($tData, $iKeyState, $tPoint, $piEffect, $sDirText)

	Local $oIDropTgtHelper = ObjCreateInterface($tData.pIDropTgtHelper, $sIID_IDropTargetHelper, $tagIDropTargetHelper)
	$oIDropTgtHelper.AddRef()
	$oIDropTgtHelper.Drop($pDataObject, $tPoint, $piEffect)

	Local $tEffect = DllStructCreate("dword iEffect", $piEffect)
	If $tEffect.iEffect <> $DROPEFFECT_NONE Then

		Local $tFormatEtc = DllStructCreate($tagFORMATETC)
		$tFormatEtc.cfFormat = $CF_HDROP
		$tFormatEtc.iIndex = -1
		$tFormatEtc.tymed = $TYMED_HGLOBAL

		Local $oDataObject = ObjCreateInterface($pDataObject, $sIID_IDataObject, $tagIDataObject)
		Local $tStgMedium = DllStructCreate($tagSTGMEDIUM)
		$oDataObject.AddRef()
		$oDataObject.GetData($tFormatEtc, $tStgMedium)

		Local $asFilenames = _WinAPI_DragQueryFileEx($tStgMedium.handle)
		; remove amount of items at [0]
		Local $arPaths[UBound($asFilenames)-1]
		For $i=0 To UBound($arPaths)-1
			$arPaths[$i] = $asFilenames[$i+1]
		Next

		Local $sTargetPathAbs = Default

		Switch _WinAPI_GetClassName_mod($tData.hTarget)
			Case $WC_LISTVIEW
				; clear previously DROPHILITED listview item
				_GUICtrlListView_SetItemState($tData.hTarget, $iPreviousHot, 0, $LVIS_DROPHILITED)

				Local $aListItem = _GUICtrlListView_HitTest($tData.hTarget, $tPoint.X, $tPoint.Y)
				Local $sItemText = _GUICtrlListView_GetItemText($tData.hTarget, $aListItem[0])
				$sTargetPathAbs = __TreeListExplorer_GetPath(1) & $sItemText
				If StringInStr(FileGetAttrib($sTargetPathAbs), "D") Then
					$sDirText = $sItemText
					$sTargetPathAbs = $sTargetPathAbs
					$bIsFolder = True
				Else
					; get the currently selected path
					$sDirPath = __TreeListExplorer_GetPath(1)
					$sTargetPathAbs = $sDirPath
					; obtain folder name only for drag tooltip
					Local $aPath = _StringBetween($sDirPath, "\", "\")
					$sDirText = $aPath[UBound($aPath) - 1]
				EndIf
				; handle case when cursor is over ListView but not on item; current directory becomes drop dir
				If $sDirText = "" Then
					; get the currently selected path
					$sDirPath = __TreeListExplorer_GetPath(1)
					$sFullPath = $sDirPath
					; obtain folder name only for drag tooltip
					Local $aPath = _StringBetween($sDirPath, "\", "\")
					$sDirText = $aPath[UBound($aPath) - 1]
				EndIf
			Case $WC_TREEVIEW
				; restore original treeview selection
				_GUICtrlTreeView_SelectItem($tData.hTarget, $hTreeOrig)

				; full treeview drop path is most recently DROPHILITED item
				$sTargetPathAbs = $sDropTV
		EndSwitch
		If $sTargetPathAbs<>Default Then
			; determine if IFileOperation needs to copy or move files
			Switch $iFinalEffect
				Case $DROPEFFECT_COPY
					__FilesOperation_DoInSub($__FileOperation_Copy, $sTargetPathAbs, $arPaths)
				Case $DROPEFFECT_MOVE
					__FilesOperation_DoInSub($__FileOperation_Move, $sTargetPathAbs, $arPaths)
					$tEffect.iEffect = $DROPEFFECT_NONE
					__SetPerformedDropEffect($pDataObject, $DROPEFFECT_NONE)
			EndSwitch
		EndIf
	EndIf

	$tData.bAcceptDrop = False
	$tData.pDataObject = 0

	Return $S_OK
EndFunc   ;==>__Mthd_Drop

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
EndFunc   ;==>__SetDropDescription

Func __DoDropResponse($tData, $iKeyState, $tPoint, $piEffect, $sDirText, $bIsSameDrive = False, $bIsSameFolder = False)

	#forceref $tData, $iKeyState, $tPoint, $piEffect

	Local $iRetEffect = $DROPEFFECT_NONE, $iReqOp
    Local $tEffect = DllStructCreate("dword iEffect", $piEffect)

    ;See what the user is asking for based on key modifiers.
    If $tData.bAcceptDrop Then
        Switch BitAND($iKeyState, BitOR($MK_CONTROL, $MK_ALT, $MK_SHIFT))
            Case BitOr($MK_CONTROL, $MK_ALT), 0
                $iReqOp = $DROPEFFECT_MOVE
            Case $MK_CONTROL
                $iReqOp = $DROPEFFECT_COPY
            Case BitOR($MK_CONTROL, $MK_SHIFT), $MK_ALT
                $iReqOp = $DROPEFFECT_LINK
        EndSwitch

        ;If move is legally an option
        If BitAND($tEffect.iEffect, $DROPEFFECT_MOVE) Then
            If $iReqOp = $DROPEFFECT_MOVE And $bIsSameDrive Then $iRetEffect = $DROPEFFECT_MOVE
        EndIf
		;If copy is legally an option
        If BitAND($tEffect.iEffect, $DROPEFFECT_COPY) Then
            If $iReqOp = $DROPEFFECT_COPY Or $iRetEffect = $DROPEFFECT_NONE Then $iRetEffect = $DROPEFFECT_COPY
        EndIf
        ;If link is legally an option
        If BitAND($tEffect.iEffect, $DROPEFFECT_LINK) Then
            If $iReqOp = $DROPEFFECT_LINK Or $iRetEffect = $DROPEFFECT_NONE Then $iRetEffect = $DROPEFFECT_LINK
        EndIf
        If $bIsSameFolder Then $iRetEffect = $DROPEFFECT_NONE
    EndIf
    $tEffect.iEffect = $iRetEffect

	Switch $tEffect.iEffect
		Case $DROPEFFECT_NONE
			__SetDropDescription($tData.pDataObject, $DROPIMAGE_NOIMAGE, "", "")

		Case $DROPEFFECT_LINK
			__SetDropDescription($tData.pDataObject, $DROPIMAGE_LINK, "Link to %1", $sDirText)

		Case $DROPEFFECT_COPY
			__SetDropDescription($tData.pDataObject, $DROPIMAGE_COPY, "Copy to %1", $sDirText)

		Case $DROPEFFECT_MOVE
			__SetDropDescription($tData.pDataObject, $DROPIMAGE_MOVE, "Move to %1", $sDirText)
	EndSwitch

	Return
EndFunc   ;==>__DoDropResponse

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
EndFunc   ;==>__SetPerformedDropEffect
