#cs ----------------------------------------------------------------------------

	 AutoIt Version: 3.3.18.0
	 Author:         Kanashius

	 Script Function:
		History management

#ce ----------------------------------------------------------------------------
#include-once
Global $__History__Data[]

; #FUNCTION# ====================================================================================================================
; Name ..........: __History_Create
; Description ...: Create a History object.
; Syntax ........:__History_Create($sCallbackUnReDo, $iMaxSteps = 100, $sCallbackChange = Default)
; Parameters ....: $sCallbackUnReDo     - callback function as string.
;                  $iMaxSteps           - [optional] the amount of history elements, which should be stored (max possible undo calls)
;                                         must be a number >0
;                  $sCallbackChange     - [optional] callback function as string. Using Default will not call any function.
; Return values .: None
; Author ........: Kanashius
; Modified ......:
; Remarks .......: The $sCallbackUnReDo calls the provided function, which must have 3 parameters ($hHistory, $bRedo, $data) and
;                  is called to execute the undo/redo operation. If $bRedo is true, the operation is to redo, otherwise it is undo.
;                  $data contains the data added by the __History_Add function.
;                  The $sCallbackChange calls the provided function, which must have 1 parameter ($hHistory) and is called, when
;                  the history changes to enable monitoring of the available amount of undo/redo operations (__History_UndoCount/
;                  __History_RedoCount)
;
;                  Errors:
;                  1 - Parameter is invalid (@extended 1 - $sCallbackUnReDo, 2 - $iMaxSteps, 3 - $sCallbackChange)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History_Create($sCallbackUnReDo, $iMaxSteps = 100, $sCallbackChange = Default)
	If Not IsFunc(Execute($sCallbackUnReDo)) Then Return SetError(1, 1, -1)
	If Not IsInt($iMaxSteps) Or $iMaxSteps<1 Then Return SetError(1, 2, -1)
	If $sCallbackChange <> Default And Not IsFunc(Execute($sCallbackChange)) Then Return SetError(1, 3, -1)
	Local $mHistory[], $mHistoryIntern[]
	$mHistory.sCallbackUnReDo = $sCallbackUnReDo
	$mHistory.sCallbackChange = $sCallbackChange
	$mHistory.mHistory = $mHistoryIntern
	$mHistory.iMaxSteps = $iMaxSteps
	$mHistory.startIndex = 0
	$mHistory.currentIndex = 0
	$mHistory.stopIndex = 0
	Return MapAppend($__History__Data, $mHistory)+1
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __History_Delete
; Description ...: Deletes the history and cleans up the resources
; Syntax ........: __History_Delete($hHistory)
; Parameters ....: $hHistory             - the history handle.
; Return values .: None
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - $hHistory is not a valid history handle
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History_Delete($hHistory)
	Local $iHistoryId = __History__GetIdFromHandle($hHistory)
	If @error Then Return SetError(1, 0, False)
	MapRemove($__History__Data, $iHistoryId)
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __History_Clear
; Description ...: Clears the history
; Syntax ........: __History_Clear($hHistory)
; Parameters ....: $hHistory             - the history handle.
; Return values .: True on success, False otherwise
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - $hHistory is not a valid history handle
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History_Clear($hHistory)
	Local $iHistoryId = __History__GetIdFromHandle($hHistory)
	If @error Then Return SetError(1, 0, False)
	Local $mHistory[]
	$__History__Data[$iHistoryId]["mHistory"] = $mHistory
	$__History__Data[$iHistoryId]["startIndex"] = 0
	$__History__Data[$iHistoryId]["currentIndex"] = 0
	$__History__Data[$iHistoryId]["stopIndex"] = 0
	If $__History__Data[$iHistoryId].sCallbackChange<>Default Then Call($__History__Data[$iHistoryId].sCallbackChange, $hHistory)
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __History_Add
; Description ...: Add an element to the history
; Syntax ........: __History_Add($hHistory, $data)
; Parameters ....: $hHistory             - the history handle.
;                  $data                 - the data object to add to the history
; Return values .: True on success, False otherwise
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - $hHistory is not a valid history handle
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History_Add($hHistory, $data)
	Local $iHistoryId = __History__GetIdFromHandle($hHistory)
	If @error Then Return SetError(1, 0, False)
	$__History__Data[$iHistoryId]["mHistory"][$__History__Data[$iHistoryId].currentIndex] = $data
	$__History__Data[$iHistoryId]["currentIndex"] += 1
	If $__History__Data[$iHistoryId].currentIndex>=$__History__Data[$iHistoryId].iMaxSteps Then $__History__Data[$iHistoryId]["currentIndex"]=0
	If $__History__Data[$iHistoryId].startIndex=$__History__Data[$iHistoryId].currentIndex Then $__History__Data[$iHistoryId]["startIndex"]+=1
	If $__History__Data[$iHistoryId].startIndex>=$__History__Data[$iHistoryId].iMaxSteps Then $__History__Data[$iHistoryId]["startIndex"]=0
	$__History__Data[$iHistoryId].stopIndex = $__History__Data[$iHistoryId].currentIndex
	If $__History__Data[$iHistoryId].sCallbackChange<>Default Then Call($__History__Data[$iHistoryId].sCallbackChange, $hHistory)
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __History_Undo
; Description ...: Execute an undo operation
; Syntax ........: __History_Undo($hHistory, $iSteps = 1)
; Parameters ....: $hHistory             - the history handle.
;                  $iSteps               - how many times to undo
; Return values .: The amount of undo operations executed
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - $hHistory is not a valid history handle
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History_Undo($hHistory, $iSteps = 1)
	Local $iHistoryId = __History__GetIdFromHandle($hHistory)
	If @error Then Return SetError(1, 0, 0)
	Local $iDoneSteps = 0
	For $i=0 to $iSteps-1
		If __History_UndoCount($hHistory)>0 Then
			$iDoneSteps+=1
			$__History__Data[$iHistoryId]["currentIndex"] -= 1
			If $__History__Data[$iHistoryId].currentIndex<0 Then $__History__Data[$iHistoryId]["currentIndex"]=$__History__Data[$iHistoryId].iMaxSteps-1
			Call($__History__Data[$iHistoryId].sCallbackUnReDo, $hHistory, False, $__History__Data[$iHistoryId]["mHistory"][$__History__Data[$iHistoryId].currentIndex])
		EndIf
	Next
	If $__History__Data[$iHistoryId].sCallbackChange<>Default And $iDoneSteps>0 Then Call($__History__Data[$iHistoryId].sCallbackChange, $hHistory)
	return $iDoneSteps
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __History_Redo
; Description ...: Execute a redo operation
; Syntax ........: __History_Redo($hHistory, $iSteps = 1)
; Parameters ....: $hHistory             - the history handle.
;                  $iSteps               - how many times to redo
; Return values .: The amount of redo operations executed
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - $hHistory is not a valid history handle
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History_Redo($hHistory, $iSteps = 1)
	Local $iHistoryId = __History__GetIdFromHandle($hHistory)
	If @error Then Return SetError(1, 0, 0)
	Local $iDoneSteps = 0
	For $i=0 to $iSteps-1
		If __History_RedoCount($hHistory)>0 Then
			$iDoneSteps+=1
			Call($__History__Data[$iHistoryId].sCallbackUnReDo, $hHistory, True, $__History__Data[$iHistoryId]["mHistory"][$__History__Data[$iHistoryId].currentIndex])
			$__History__Data[$iHistoryId]["currentIndex"] += 1
			If $__History__Data[$iHistoryId].currentIndex>=$__History__Data[$iHistoryId].iMaxSteps Then $__History__Data[$iHistoryId]["currentIndex"]=0
		EndIf
	Next
	If $__History__Data[$iHistoryId].sCallbackChange<>Default And $iDoneSteps>0 Then Call($__History__Data[$iHistoryId].sCallbackChange, $hHistory)
	return $iDoneSteps
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __History_UndoCount
; Description ...: How many undo operations can be executed
; Syntax ........: __History_UndoCount($hHistory)
; Parameters ....: $hHistory             - the history handle.
; Return values .: The amount of possible undo operations
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - $hHistory is not a valid history handle
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History_UndoCount($hHistory)
	Local $iHistoryId = __History__GetIdFromHandle($hHistory)
	If @error Then Return SetError(1, 0, 0)
	If $__History__Data[$iHistoryId].startIndex = $__History__Data[$iHistoryId].currentIndex Then
		Return 0
	ElseIf $__History__Data[$iHistoryId].startIndex<$__History__Data[$iHistoryId].currentIndex Then
		Return $__History__Data[$iHistoryId].currentIndex-$__History__Data[$iHistoryId].startIndex
	Else
		Return $__History__Data[$iHistoryId].iMaxSteps-$__History__Data[$iHistoryId].startIndex+$__History__Data[$iHistoryId].currentIndex
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __History_RedoCount
; Description ...: How many redo operations can be executed
; Syntax ........: __History_RedoCount($hHistory)
; Parameters ....: $hHistory             - the history handle.
; Return values .: The amount of possible redo operations
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - $hHistory is not a valid history handle
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History_RedoCount($hHistory)
	Local $iHistoryId = __History__GetIdFromHandle($hHistory)
	If @error Then Return SetError(1, 0, 0)
	If $__History__Data[$iHistoryId].stopIndex = $__History__Data[$iHistoryId].currentIndex Then
		Return 0
	ElseIf $__History__Data[$iHistoryId].stopIndex>$__History__Data[$iHistoryId].currentIndex Then
		Return $__History__Data[$iHistoryId].stopIndex-$__History__Data[$iHistoryId].currentIndex
	Else
		Return $__History__Data[$iHistoryId].iMaxSteps-$__History__Data[$iHistoryId].currentIndex+$__History__Data[$iHistoryId].stopIndex
	EndIf
EndFunc
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __History__GetIdFromHandle
; Description ...: Convert a history handle to a history ID
; Syntax ........: __History__GetIdFromHandle($hHistory)
; Parameters ....: $hHistory             - the history handle.
; Return values .: The history ID.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - No history with the handle $hHistory exists
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History__GetIdFromHandle($hHistory)
	If MapExists($__History__Data, $hHistory-1) Then Return $hHistory-1
	Return SetError(1, 0, 0)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __History__GetHandleFromId
; Description ...: Convert a history ID to a history handle
; Syntax ........: __History__GetHandleFromId($iHistoryId)
; Parameters ....: $iHistoryId             - the history ID.
; Return values .: The history handle.
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __History__GetHandleFromId($iHistoryId)
	Return $iHistoryId+1
EndFunc