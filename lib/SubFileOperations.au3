#cs ----------------------------------------------------------------------------

	 AutoIt Version: 3.3.18.0
	 Author:         Kanashius

	 Script Function:
		Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include-once
#include "../lib/IPC.au3"
#include "../lib/TreeListExplorer.au3"
#include "../lib/IFileOperation.au3"
#include "../lib/IUnknown.au3"

Global $__FileOperation_Copy = 1, $__FileOperation_Move = 2, $__FileOperation_ProcessReady = 3, $__FileOperation_Successful = 4, $__FileOperation_Failed = 5
Global $__FileOperation_mData[]

; ====================== sub process ======================

Func __FileOperations_Sub($hSubProcess)
	__IPC_SubSendCmd($__FileOperation_ProcessReady)
	While Sleep(100)
	WEnd
EndFunc

Func __FileOperations_SubExit()
	; consider handling the abortion of any running copy/move
	__IPC_Shutdown()
	Exit
EndFunc

Func __FileOperations_SubReceive($iCmd, $arData)
	Switch $iCmd
		Case $__FileOperation_Copy, $__FileOperation_Move
			If UBound($arData)<3 Then Return __IPC_Log($__IPC_LOG_ERROR, "$__FileOperation_Copy/$__FileOperation_Move failed: Missing parameters.")
			Local $bResult = __FileOperations_SubCopyOrMove(($iCmd=$__FileOperation_Copy)?"CopyItems":"MoveItems", $arData[0], $arData[1], $arData[2])
			If @error=1 Then Return __IPC_Log($__IPC_LOG_ERROR, "$__FileOperation_Copy/$__FileOperation_Move failed: Invalid parameter ("&@extended&").")
			__IPC_SubSendCmd($bResult?$__FileOperation_Successful:$__FileOperation_Failed)
			__FileOperations_SubExit()
	EndSwitch
EndFunc

Func __FileOperations_SubCopyOrMove($sAction, $sTargetPath, ByRef $arPaths, $iFlags)
	If Not IsString($sTargetPath) Then Return SetError(1, 1, False)
	If Not IsArray($arPaths) Then Return SetError(1, 2, False)
	If Not IsInt($iFlags) Then Return SetError(1, 3, False)
	#cs
	Local $sPath = ""
	For $i = 0 To UBound($arPaths)-1
		If $i>0 Then $sPath &= ","
		$sPath &= $arPaths[$i]
	Next
	__IPC_Log($__IPC_LOG_DEBUG, $sAction&" to: "&$sTargetPath&" >> "&$sPath)
	#ce
	Local $pDataObj = GetDataObjectOfFile_B($arPaths)
	Local $iResult = _IFileOperationFile($pDataObj, $sTargetPath, $sAction, $iFlags)
	_Release($pDataObj)
	Return $iResult?(True):(False)
EndFunc

; ====================== main process ======================

Func __FilesOperation_DoInSub($iOperation, $sTarget, ByRef $arPaths)
	Local $mOperationElem[]
	$mOperationElem.iType = $iOperation
	$mOperationElem.sTarget = $sTarget
	$mOperationElem.arPaths = $arPaths
	$mOperationElem.iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
	Local $hSubProcess = __IPC_StartProcess("__FileOperations_Receive", Default, "__FilesOperation_RemoveSub")
	If @error Then __IPC_Log($__IPC_LOG_ERROR, "Error starting subprocess: "&@error&":"&@extended)
	; should be fine here, but maybe cause problems, when $__FileOperation_ProcessReady is received, before this was done...
	; look out for => Error: Unknown operation process ready
	$__FileOperation_mData[$hSubProcess] = $mOperationElem
EndFunc

Func __FilesOperation_RemoveSub($hSubProcess)
	__IPC_ProcessStop($hSubProcess)
	MapRemove($__FileOperation_mData, $hSubProcess)
EndFunc

Func __FileOperations_Receive($hSubProcess, $iCmd, $arData)
	Switch $iCmd
		Case $__FileOperation_ProcessReady
			If Not MapExists($__FileOperation_mData, $hSubProcess) Then
				__IPC_Log($__IPC_LOG_ERROR, "Unknown operation process ready")
				__IPC_ProcessStop($hSubProcess)
				Return
			EndIf
			Local $mOperationElem = $__FileOperation_mData[$hSubProcess]
			__IPC_MainSendCmd($hSubProcess, $mOperationElem.iType, $mOperationElem.sTarget, $mOperationElem.arPaths, $mOperationElem.iFlags)
		Case $__FileOperation_Successful
			__IPC_Log($__IPC_LOG_INFO, "Operation successfull")
			__TreeListExplorer_Reload(1)
		Case $__FileOperation_Failed
			__IPC_Log($__IPC_LOG_ERROR, "Operation failed")
	EndSwitch
EndFunc