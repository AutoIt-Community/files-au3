#include-once
#include <StaticConstants.au3>
#include <WinAPISysWin.au3>
#include <FileConstants.au3>

; #INDEX# =======================================================================================================================
; Title .........: Lang (AutoIt Language UDF)
; AutoIt Version : 3.3.18.1
; Language ......: English
; Description ...: UDF to help managing multiple languages.
; Author(s) .....: Kanashius
; Version .......: 1.0.1
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; __Lang_CallbackLabel
; __Lang_CallbackGuiCtrlSetData
; __Lang_GetLanguages
; __Lang_GetCurrentLanguage
; __Lang_Load
; __Lang_Get
; __Lang_CreateCallback
; __Lang_SetCallback
; __Lang_DeleteCallback
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __Lang__HandleCallback
; __Lang__Get
; __Lang__LoadIni
; __Lang__GUICtrlGetStyle
; ===============================================================================================================================

; #INTERNAL_USE_ONLY GLOBAL VARIABLES # =========================================================================================
Global $__Lang__mLang[]
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_CallbackLabel
; Description ...: Callback to register a label control for language changes, including the reading direction.
;                  To be used to register with __Lang_CreateCallback, e.g. __Lang_CreateCallback("__Lang_CallbackLabel", $idLabel).
; Syntax ........: __Lang_CallbackLabel($sKey, $sVal, $bRTL, $idCtrl)
; Parameters ....: $sKey	         - The data key. Provided by the callback.
;                  $sVal	         - The data value for that key. Provided by the callback.
;                  $bRTL  			 - True, for reading direction right-to-left. Provided by the callback.
;                  $idCtrl           - The label control id, provided as parameter when registering the callback
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_CallbackLabel($sKey, $sVal, $bRTL, $idCtrl)
	GUICtrlSetData($idCtrl, $sVal)
	Local $iStyle = __Lang__GUICtrlGetStyle($idCtrl)
	If $bRTL And Not BitAND($iStyle, $SS_RIGHT) Then
		GUICtrlSetStyle($idCtrl, BitOR(BitAND($iStyle, BitNOT($SS_LEFT)), $SS_RIGHT))
	ElseIf Not $bRTL And Not BitAND($iStyle, $SS_LEFT) Then
		GUICtrlSetStyle($idCtrl, BitOR(BitAND($iStyle, BitNOT($SS_RIGHT)), $SS_LEFT))
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_CallbackGuiCtrlSetData
; Description ...: Callback to register any control for language changes, if that can be done by calling GuiCtrlSetData.
;                  To be used to register with __Lang_CreateCallback, e.g. __Lang_CreateCallback("__Lang_CallbackGuiCtrlSetData", $idCtrl).
; Syntax ........: __Lang_CallbackGuiCtrlSetData($sKey, $sVal, $bRTL, $idCtrl)
; Parameters ....: $sKey	         - The data key. Provided by the callback.
;                  $sVal	         - The data value for that key. Provided by the callback.
;                  $bRTL  			 - True, for reading direction right-to-left. Provided by the callback.
;                  $idCtrl           - The control id, provided as parameter when registering the callback
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_CallbackGuiCtrlSetData($sKey, $sVal, $bRTL, $idCtrl)
	GUICtrlSetData($idCtrl, $sVal)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_GetLanguages
; Description ...: Get all languages provided by the language config.
; Syntax ........: __Lang_GetLanguages([$sLangIni = "lang.ini"])
; Parameters ....: $sLangIni         - [optional] Default: "lang.ini". The path to the language config file.
; Return values .: A 2D-Array with the languages.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: See __Lang_Load for more information about the language config file.
;                  Returns all found languages in a 2D-Array with the section names (e.g. en/de/...) and the language name
;                  (if the name is missing, the section name will be inserted).
;                  [0][0] = 1st section name (usually the ISO 639 language code)
;                  [0][1] = 1st language name (usually in the original language itself)
;                  ...
;                  [n][0] = n-th section name
;                  [n][1] = n-th language name
;
;                  Errors:
;                  1 - Parameter not valid (@extended: 1 - $sLangIni) => File could not be properly read.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_GetLanguages($sLangIni = "lang.ini")
	Local $hFile = FileOpen($sLangIni, BitOR($FO_READ, $FO_FULLFILE_DETECT))
	If @error Then Return SetError(2, @error, False)
	Local $arSections[0][2]
	While True
		Local $sLine = FileReadLine($hFile)
		If @error=-1 Then ExitLoop
		If @error Then Return SetError(2, @error, False)
		Local $arData = StringRegExp($sLine, "^(?:\[(.*?)\]|([^=]*)=(.*))$", 1)
		If UBound($arData)=1 Then ; is a section line or file end
			ReDim $arSections[UBound($arSections)+1][2]
			$arSections[UBound($arSections)-1][0] = $arData[0]
		ElseIf UBound($arSections)>0 And UBound($arData)=3 And $arData[1] = "Name" Then
			$arSections[UBound($arSections)-1][1] = $arData[2]
		EndIf
	WEnd
	FileClose($hFile)
	Return $arSections
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_GetCurrentLanguage
; Description ...: Get the currently selected language section name (language code)
; Syntax ........: __Lang_GetCurrentLanguage()
; Parameters ....:
; Return values .: The section name on success, False otherwise
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  2 - No language selected, call __Lang_Load first
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_GetCurrentLanguage()
	If Not MapExists($__Lang__mLang, "sLang") Then Return SetError(2, 0, False)
	Return $__Lang__mLang.sLang
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_Load
; Description ...: Set the current language for the application.
; Syntax ........: __Lang_Load([$sLang=Default, [$sLangIni = "lang.ini"]])
; Parameters ....: $sLang	         - [optional] The section name of the language to load (usually the ISO 639 language code).
;                                                 Default: The first section name in the config file.
;                  $sLangIni	     - [optional] Default: "lang.ini". The path to the language config file.
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: The $sLangIni language config file must be an ini file with the following format:
;                  - It supports different string encodings (automatically detected, see FileOpen with $FO_FULLFILE_DETECT)
;                  - Sections must be a string surrounded by []. Leading/Trailing spaces are removed.
;                  - Key/Value-Pairs must be seperated by "=". Everything before the first "=" is the key, everything after that a value.
;                      Leading/Trailing spaces are removed.
;                  - Newlines must be \r\n (@CRLF).
;                  - Every section will be interpreted as a language and should be a ISO 639 language code
;                  - Keys should start with a lower-case character
;                  - Keys should represent their value in the default language (if a key is not found, the key will be returned as value for that string)
;                  - The keys "Name", "RightToLeft" and "File" are reserved for specific purposes:
;                      - "Name" defines the name of the language for that section (Must be specified in the main config file)
;                      - "RightToLeft" can be set to "true" to mark a language with a reading direction of right-to-left
;                      - "File" can be used to specify a different file, where the language should be loaded from (if the main file gets to large)
;                        => These files follow the same rules, but they are not allowed to have "File" keys themselves.
;                           They can contain multiple languages, but all languages must be defined in the main file.
;                           All keys will be loaded and added to the existing keys from the original file.
;                  - The values need to follow the rules for StringFormat. When retrieving a value with a key, the value will be put into StringFormat with the additional parameters (See __Lang_Get).
;
;                  @extended - 1: The language was not found and the default (first) language was used instead
;
;                  Errors:
;                  1 - Parameter not valid (@extended: 1 - $sLang (no language found in config),
;                                                      2 - $sLangIni (File could not be read/parsed))
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_Load($sLang=Default, $sLangIni = "lang.ini")
	If Not MapExists($__Lang__mLang, "mCallbacks") Then
		Local $mCallbacks[]
		$__Lang__mLang.mCallbacks = $mCallbacks
	EndIf
	If Not MapExists($__Lang__mLang, "sLangDefault") Then $__Lang__mLang.sLangDefault = Default
	__Lang__LoadIni($sLangIni, $sLang)
	If @error Then Return SetError(1, 2, False)
	Local $bFallbackToDefault = False
	If $sLang = Default Or Not MapExists($__Lang__mLang.mIni, $sLang) Then
		If $sLang<>Default Then $bFallbackToDefault = True
		$sLang = $__Lang__mLang.sLangDefault
	EndIf
	If $sLang = Default Then Return SetError(1, 1, False) ; no language found in config
	$__Lang__mLang.sLang = $sLang
	For $iKey In MapKeys($__Lang__mLang.mCallbacks)
		__Lang__HandleCallback($iKey+1)
	Next
	Return SetExtended($bFallbackToDefault?1:0, True)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_Get
; Description ...: Get the text value for a specific key.
; Syntax ........: __Lang_Get($sKey, [$vParam1 = Default, [$vParam2 = Default, ... , [$vParam32 = Default]]])
; Parameters ....: $sKey         - The key to retrieve the text value for.
;                  $vParam1      - [optional] Default: None. Up to 32 variables that will be output. See Remarks.
;                  ...
;                  $vParam32     - [optional] Default: None. Up to 32 variables that will be output. See Remarks.
; Return values .: The requested formatted text on success. "<<"&$sKey&">>" on failure.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: The value retrieved for $sKey is first processed with StringFormat, before it is returned.
;                  The parameters $vParam1 to $vParam32 are passed to StringFormat.
;                  The result of the StringFormat call is returned by this method.
;                  e.g. __Lang_Get($sKey, 1, 2, 3) will result in: StringFormat(<<GetValueForKey>>($sKey), 1, 2, 3)
;
;                  If $sKey is not found for the current language, the value for the default (first) language is returned.
;                  If $sKey is also missing for the default language, then "<<"&$sKey&">>" will be returned.
;
;                  Errors:
;                  1 - Parameter not valid (@extended: 1 - $sKey)
;                  2 - No language loaded, call __Lang_Load first
;                  3 - StringFormat of the value failed
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_Get($sKey, $vParam1 = Default, $vParam2 = Default, $vParam3 = Default, $vParam4 = Default, $vParam5 = Default, _
						$vParam6 = Default, $vParam7 = Default, $vParam8 = Default, $vParam9 = Default, $vParam10 = Default, _
						$vParam11 = Default, $vParam12 = Default, $vParam13 = Default, $vParam14 = Default, $vParam15 = Default, _
						$vParam16 = Default, $vParam17 = Default, $vParam18 = Default, $vParam19 = Default, $vParam20 = Default, _
						$vParam21 = Default, $vParam22 = Default, $vParam23 = Default, $vParam24 = Default, $vParam25 = Default, _
						$vParam26 = Default, $vParam27 = Default, $vParam28 = Default, $vParam29 = Default, $vParam30 = Default, _
						$vParam31 = Default, $vParam32 = Default)
	Local $arParams[@NumParams-1]
	For $i=0 To UBound($arParams)-1
		$arParams[$i] = Eval("vParam"&($i+1))
	Next
	Local $sVal = __Lang__Get($sKey, $arParams)
	If @error Then Return SetError(@error, @extended, $sVal)
	Return SetExtended(@extended, $sVal)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_CreateCallback
; Description ...: Create a callback, which is called when the language changes.
; Syntax ........: __Lang_CreateCallback($sFunc, [$vParam1 = Default, ..., [$vParam10 = Default]])
; Parameters ....: $sFunc         - The name of the function to call
;                  $vParam1      - [optional] Default: None. Up to 10 variables that will be passed to the callback.
;                  ...
;                  $vParam32     - [optional] Default: None. Up to 10 variables that will be passed to the callback.
; Return values .: The callback id or 0 on failure.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: The callback is created and added to the system, but it will not be called, unless __Lang_SetCallback was called first.
;                  Depending on __Lang_SetCallback, different parameters are required for the provided callback function.
;                  If __Lang_SetCallback is called with a $sKey, the callback requires ($sKey, $sVal, $bRTL,...).
;                  If __Lang_SetCallback is called with $sKey = Default, the callback requires ($bRTL,...).
;                  Additional parameters are required at the callback function for every $vParamN passed to __Lang_CreateCallback.
;                  These parameters are stored and added to every callback call.
;                  e.g. - Local $iCallback = __Lang_CreateCallback("_callbackFunc", 1, 2, 3) with __Lang_SetCallback($iCallback)
;                         requires Func _callbackFunc($bRTL, $iParam1, $iParam2, $iParam3)
;                       - Local $iCallback = __Lang_CreateCallback("_callbackFunc", 1, 2, 3) with __Lang_SetCallback($iCallback, "someText")
;                         requires Func _callbackFunc($sKey, $sVal, $bRTL, $iParam1, $iParam2, $iParam3)
;                  If $sKey is not found for the current language, the value for the default (first) language is provided for $sVal.
;                  If $sKey is also missing for the default language, then "<<"&$sKey&">>" will be provided for $sVal.
;
;                  Errors:
;                  1 - Parameter not valid (@extended: 1 - $sFunc)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_CreateCallback($sFunc, $vParam1 = Default, $vParam2 = Default, $vParam3 = Default, $vParam4 = Default, $vParam5 = Default, _
									$vParam6 = Default, $vParam7 = Default, $vParam8 = Default, $vParam9 = Default, $vParam10 = Default)
	If Not IsFunc(Execute($sFunc)) Then Return SetError(1, 1, 0)
	Local $arCallbackParams[@NumParams-1]
	For $i=0 To UBound($arCallbackParams)-1
		$arCallbackParams[$i] = Eval("vParam"&($i+1))
	Next
	Local $mCallback[]
	$mCallback.sCallback = $sFunc
	$mCallback.arCallbackParams = $arCallbackParams
	$mCallback.sKey = Default
	$mCallback.arParams = Default
	Return MapAppend($__Lang__mLang.mCallbacks, $mCallback) + 1
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_SetCallback
; Description ...: Activates the callback and optionally sets/updates the key with its parameters.
; Syntax ........: __Lang_SetCallback($iCallback, [$sKey = Default, [$vParam1 = Default, ..., [$vParam32 = Default]]])
; Parameters ....: $iCallback         - The callback id
;                  $sKey              - [optional] Default: None. The key to use when calling the callback.
;                  $vParam1           - [optional] Default: None. The parameters for formatting. (See __Lang_Get)
;                  ...
;                  $vParam32          - [optional] Default: None. The parameters for formatting. (See __Lang_Get)
; Return values .: True on success
; Author ........: Kanashius
; Modified ......:
; Remarks .......: After __Lang_SetCallback was called, the callback gets activated, so the callback function will be called at language changes.
;                  When __Lang_SetCallback is called, the first callback call is triggered to set/update the callback.
;                  This function can used to update/change the key for a callback.
;
;                  Errors:
;                  1 - Parameter not valid (@extended: 1 - $iCallback, 3 - $vParamN)
;                  4 - Error calling the callback function, check the required parameters (__Lang_CreateCallback)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_SetCallback($iCallback, $sKey = Default, _
						$vParam1 = Default, $vParam2 = Default, $vParam3 = Default, $vParam4 = Default, $vParam5 = Default, _
						$vParam6 = Default, $vParam7 = Default, $vParam8 = Default, $vParam9 = Default, $vParam10 = Default, _
						$vParam11 = Default, $vParam12 = Default, $vParam13 = Default, $vParam14 = Default, $vParam15 = Default, _
						$vParam16 = Default, $vParam17 = Default, $vParam18 = Default, $vParam19 = Default, $vParam20 = Default, _
						$vParam21 = Default, $vParam22 = Default, $vParam23 = Default, $vParam24 = Default, $vParam25 = Default, _
						$vParam26 = Default, $vParam27 = Default, $vParam28 = Default, $vParam29 = Default, $vParam30 = Default, _
						$vParam31 = Default, $vParam32 = Default)
	If Not MapExists($__Lang__mLang.mCallbacks, $iCallback-1) Then Return SetError(1, 1, False)
	Local $arParams[@NumParams-2]
	For $i=0 To UBound($arParams)-1
		$arParams[$i] = Eval("vParam"&($i+1))
	Next
	$__Lang__mLang["mCallbacks"][$iCallback-1]["sKey"] = $sKey
	$__Lang__mLang["mCallbacks"][$iCallback-1]["arParams"] = $arParams
	__Lang__HandleCallback($iCallback)
	If @error Then Return SetError(@error, @extended, False)
	Return True
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: __Lang_DeleteCallback
; Description ...: Delete the specified callback.
; Syntax ........: __Lang_DeleteCallback($iCallback)
; Parameters ....: $iCallback         - The callback id
; Return values .: True on success
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Should be called for every callback, when it is no longer needed.
;                  This removes the callback and the corresponding values and frees their memory.
;                  It is not required to be called when the program exits.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang_DeleteCallback($iCallback)
	If $iCallback>0 And MapExists($__Lang__mLang.mCallbacks, $iCallback-1) Then
		MapRemove($__Lang__mLang.mCallbacks, $iCallback-1)
		Return True
	EndIf
	Return False
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Lang__HandleCallback
; Description ...: Handle the specified callback and trigger a call.
; Syntax ........: __Lang__HandleCallback($iCallback)
; Parameters ....: $iCallback        - the callback id
; Return values .: True on success
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter not valid (@extended: 1 - $iCallback, 3 - callback parameter not specified => call __Lang_SetCallback first)
;                  4 - Error calling the callback function, check the required parameters (__Lang_CreateCallback)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang__HandleCallback($iCallback)
	If Not MapExists($__Lang__mLang.mCallbacks, $iCallback-1) Then Return SetError(1, 1, False)
	Local $mCallback = $__Lang__mLang.mCallbacks[$iCallback-1]
	If $mCallback.arParams = Default Then Return SetError(1, 3, False)
	Local $sVal = Default, $iExt = 0
	If $mCallback.sKey<>Default Then
		$sVal = __Lang__Get($mCallback.sKey, $mCallback.arParams)
		; Errors will be ignored to provide the key on error
		; If @error And @error=1 Then Return SetError(@error, @extended+1, False)
		; If @error Then Return SetError(@error, @extended, False)
		$iExt = @error?0:@extended
	Else
		$iExt = $__Lang__mLang.mIni[$__Lang__mLang.sLang].bRTL
	EndIf
	Local $iAdditionalArgs = 4
	If $mCallback.sKey=Default Then $iAdditionalArgs = 2
	Local $arArgs[UBound($mCallback.arCallbackParams)+$iAdditionalArgs]
	$arArgs[0] = "CallArgArray"
	If $iAdditionalArgs = 4 Then
		$arArgs[1] = $mCallback.sKey
		$arArgs[2] = $sVal
	EndIf
	$arArgs[$iAdditionalArgs-1] = $iExt?(True):(False)
	For $i=0 To UBound($mCallback.arCallbackParams)-1
		$arArgs[$i+$iAdditionalArgs] = $mCallback.arCallbackParams[$i]
	Next
	Call($mCallback.sCallback, $arArgs)
	If @error = 0xDEAD And @extended = 0xBEEF Then Return SetError(4, 0, False)
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Lang__Get
; Description ...: Get the value for a key. Returns for the current language if available, otherwise for the default language.
; Syntax ........: __Lang__Get($sKey, ByRef $arParams)
; Parameters ....: $sKey        - the key
;                  $arParams    - the array with parameters to pass to StringFormat.
; Return values .: True on success
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Reading direction is returned with @extended (0 - ltr, 1 - rtl)
;
;                  Errors:
;                  1 - Parameter not valid (@extended: 1 - $sKey)
;                  2 - Language not loaded
;                  3 - StringFormat error
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang__Get($sKey, ByRef $arParams)
	If Not MapExists($__Lang__mLang, "sLang") Then Return SetError(2, 0, "<<"&$sKey&">>")
	Local $sLang = $__Lang__mLang.sLang
	If Not MapExists($__Lang__mLang.mIni[$sLang].mEntries, $sKey) Then
		$sLang = $__Lang__mLang.sLangDefault
		If $sLang=Default Or Not MapExists($__Lang__mLang.mIni[$sLang].mEntries, $sKey) Then Return SetError(1, 1, "<<"&$sKey&">>")
	EndIf
	Local $arArgs[UBound($arParams)+2]
	$arArgs[0] = "CallArgArray"
	$arArgs[1] = $__Lang__mLang.mIni[$sLang].mEntries[$sKey]
	Local $bRTL = $__Lang__mLang.mIni[$sLang].bRTL
	For $i=0 To UBound($arParams)-1
		$arArgs[$i+2] = $arParams[$i]
	Next
	Local $sVal = Call("StringFormat", $arArgs)
	If @error Then Return SetError(3, @error, "<<"&$sKey&">>")
	Return SetExtended($bRTL?1:0, $sVal)
EndFunc


; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Lang__LoadIni
; Description ...: Load a language config file
; Syntax ........: __Lang__LoadIni($sFile[, $bRec = False])
; Parameters ....: $sFile       - the language config file
;                  $sLang       - [optional] Default: none. The language to load additionally to the default on.
;                  $bRec        - [optional] Default: False. True for recursive calls.
; Return values .: True on success, the sub ini for recursive calls
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  2 - File could not be opened
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang__LoadIni($sFile, $sLang = Default, $bRec = False)
	Local $mIni[]
	Local $hFile = FileOpen($sFile, BitOR($FO_READ, $FO_FULLFILE_DETECT))
	If @error Then Return SetError(2, @error, False)
	If Not $bRec Then $__Lang__mLang.sLangDefault = Default
	; iterate all lines and handle them
	Local $bAdd = False, $mEntries[], $iCount = 0, $mFileEntries[], $sSection = Default
	While True
		Local $sLine = FileReadLine($hFile)
		If @error=-1 Then ExitLoop
		If @error Then Return SetError(2, @error, False)
		Local $arData = StringRegExp($sLine, "^(?:\[\s*(.*?)\s*\]|\s*([^=]*?)\s*=\s*(.*?)\s*)$", 1)
		If UBound($arData)=1 Then ; is a section line
			If $iCount = 0 Or $arData[0] = $sLang Then
				If $sSection<>Default Then $mIni[$sSection]["mEntries"] = $mEntries
				$sSection = $arData[0]
				If Not $bRec And $iCount = 0 Then $__Lang__mLang.sLangDefault = $sSection
				$bAdd = True
				Local $mSection[], $mNewEntries[]
				$mEntries = $mNewEntries
				$mSection.sName = $sSection
				$mSection.iPos = $iCount
				$iCount += 1
				$mSection.bRTL = False
				$mIni[$sSection] = $mSection
			Else
				$bAdd = False
			EndIf
		ElseIf $bAdd And UBound($arData)=3 Then ; is a line with key/value pair
			$mEntries[$arData[1]] = $arData[2]
			If $arData[1]="File" And Not $bRec And Not MapExists($mFileEntries, $arData[2]) Then
				Local $sFolder = ""
				Local $arFolder = StringRegExp($sFile, "^(.*[\/\\]).*$", 1)
				If UBound($arFolder)>0 Then $sFolder = $arFolder[0]
				Local $mSubIni = __Lang__LoadIni($sFolder&$arData[2], $sLang, True)
				If Not @error Then $mFileEntries[$arData[2]] = $mSubIni
			EndIf
		EndIf
	WEnd
	If $sSection<>Default Then $mIni[$sSection]["mEntries"] = $mEntries
	FileClose($hFile)
	; Add entries from additional files
	For $iKey In MapKeys($mFileEntries)
		Local $mSubIni = $mFileEntries[$iKey]
		For $sSection In MapKeys($mSubIni)
			If MapExists($mIni, $sSection) Then
				For $sKey In MapKeys($mSubIni[$sSection].mEntries)
					$mIni[$sSection]["mEntries"][$sKey] = $mSubIni[$sSection].mEntries[$sKey]
					If $sKey="RightToLeft" And $mIni[$sSection]["mEntries"][$sKey]="true" Then $mIni[$sSection]["bRTL"] = True
				Next
			EndIf
		Next
	Next
	For $sKey In MapKeys($mIni)
		If MapExists($mIni[$sKey].mEntries, "RightToLeft") And StringLower($mIni[$sKey].mEntries["RightToLeft"]) = "true" Then $mIni[$sKey].bRTL = True
	Next
	If $bRec Then Return $mIni
	$__Lang__mLang.mIni = $mIni
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Lang__GUICtrlGetStyle
; Description ...: Get the control $iStyle
; Syntax ........: __Lang__GUICtrlGetStyle($hWnd)
; Parameters ....: $hWnd       - the control id/handle
; Return values .: $iStyle on success, 0 otherwise
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  2 - $iStyle could not be retrieved
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang__GUICtrlGetStyle($hWnd)
    If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $iStyle = _WinAPI_GetWindowLong($hWnd, $GWL_STYLE)
	If @error Then Return SetError(2, @error, 0)
    Return $iStyle
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Lang__GUICtrlGetStyleEx
; Description ...: Get the controls $iExStyle
; Syntax ........: __Lang__GUICtrlGetStyleEx($hWnd)
; Parameters ....: $hWnd       - the control id/handle
; Return values .: $iExStyle on success, 0 otherwise
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  2 - $iExStyle could not be retrieved
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Lang__GUICtrlGetStyleEx($hWnd)
    If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $iExStyle = _WinAPI_GetWindowLong($hWnd, $GWL_EXSTYLE)
	If @error Then Return SetError(2, @error, 0)
    Return $iExStyle
EndFunc