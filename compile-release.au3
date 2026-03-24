#cs ----------------------------------------------------------------------------

	 AutoIt Version: 3.3.18.0
	 Author:         Kanashius

	 Script Function:
		Compile release zip files for the project.

#ce ----------------------------------------------------------------------------

If @Compiled Then ; requires @AutoItExe to contain the AutoIt-Installation path to find the Au2Exe
	ConsoleWrite("This application should not be compiled."&@crlf)
	Exit
EndIf
; get Aut2Exe path
Global $sAu2Exe = _fileGetFolder(@AutoItExe)&"Aut2Exe\Aut2exe.exe"

; Define output folder
Global $sReleaseFolder = "release\"

; Define files/folders
Global $sMainSrc = "src\main.au3"
Global $sAppIcon = "resources\images\app.ico"
Global $arFiles = ["resources", "CHANGELOG.md", "LICENSE.md", "README.md"]
Global $arSourceFiles = ["src", "lib", "compile-release.au3"]

; parse version
Local $arVersion = StringRegExp(FileRead($sMainSrc), '(?m)^\s*Global\s*\$sVersion\s*=\s*"(\d+\.\d+\.\d+)\s*-\s*(\d+-\d+-\d+)\s*"$', 1)
If UBound($arVersion)<2 Then
	ConsoleWrite("Error: Version could not be parsed."&@crlf)
	Exit
EndIf
Global $sVersion = $arVersion[0]
Global $sDate = $arVersion[1] ; maybe use current date as release date instead

ConsoleWrite("============ TARGET x86 ============"&@crlf)
Global $sExeX86 = $sReleaseFolder&"FilesAu3-x86.exe"
Global $arReleaseFilesX86 = $arFiles
ReDim $arReleaseFilesX86[UBound($arReleaseFilesX86)+1]
$arReleaseFilesX86[UBound($arReleaseFilesX86)-1] = $sExeX86
If Not _au3Compile($sAu2Exe, $sMainSrc, False, $sExeX86, $sAppIcon) Then
	ConsoleWrite("Error: During compilation.")
	Exit
EndIf
If Not _zipFiles($sReleaseFolder&"FilesAu3-portable-"&$sVersion&"-x86", $arReleaseFilesX86) Then
	ConsoleWrite("Error: Failed to create the x86 release"&@crlf)
EndIf
FileDelete($sExeX86)

ConsoleWrite("============ TARGET x64 ============"&@crlf)
Global $sExeX64 = $sReleaseFolder&"FilesAu3-x64.exe"
Global $arReleaseFilesX64 = $arFiles
ReDim $arReleaseFilesX64[UBound($arReleaseFilesX64)+1]
$arReleaseFilesX64[UBound($arReleaseFilesX64)-1] = $sExeX64
If Not _au3Compile($sAu2Exe, $sMainSrc, True, $sExeX64, $sAppIcon) Then
	ConsoleWrite("Error: During compilation.")
	Exit
EndIf
If Not _zipFiles($sReleaseFolder&"FilesAu3-portable-"&$sVersion&"-x64", $arReleaseFilesX64) Then
	ConsoleWrite("Error: Failed to create the x64 release"&@crlf)
EndIf
FileDelete($sExeX64)

ConsoleWrite("============ TARGET src ============"&@crlf)
Global $arReleaseFilesSrc = _arraysCombine($arFiles, $arSourceFiles)
If Not _zipFiles($sReleaseFolder&"FilesAu3-"&$sVersion&"-src", $arReleaseFilesSrc) Then
	ConsoleWrite("Error: Failed to create the src release"&@crlf)
EndIf

ConsoleWrite("Finished successfully"&@crlf)
Exit

; $iComp values 0-4 (Default=2)
Func _au3Compile($sAu2Exe, $sSrc, $bX64 = False, $sExe = Default, $sIco = Default, $iComp = Default, $bPack = True)
	Local $sCommand = '"'&$sAu2Exe&'"'&' /in "'&$sSrc&'"'
	If $sExe<>Default Then $sCommand &=' /out "'&$sExe&'"'
	If $sIco<>Default Then $sCommand &=' /icon "'&$sIco&'"'
	If $iComp<>Default Then $sCommand &=' /comp "'&$iComp&'"'
	If Not $bPack Then $sCommand &=' /nopack "'
	$sCommand &= $bX64?' /x64':' /x86'
	Local $sToTarget = ""
	If $sExe<>Default Then $sToTarget = " to "&$sExe
	ConsoleWrite("Start compiling "&$sSrc&$sToTarget&@crlf)
	Local $iExit = RunWait($sCommand)
	If $iExit<>0 Then Return SetError(2, 0, False)
	ConsoleWrite("Done compiling "&$sSrc&$sToTarget&@crlf)
	Return True
EndFunc

Func _zipFiles($sTarget, ByRef $arFiles)
	Local $sFiles = ""
	For $i=0 To UBound($arFiles)-1
		If $i>0 Then $sFiles&=", "
		$sFiles&="'"&$arFiles[$i]&"'"
	Next
	Local $sTargetFile = $sTarget&".zip"
	If FileExists($sTargetFile) Then FileDelete($sTargetFile)
	Local $sCommand = '"PowerShell.exe" -Command "Compress-Archive -Path '&$sFiles&' -DestinationPath '&"'"&$sTargetFile&"'"&'"'
	ConsoleWrite("Start creating "&$sTargetFile&@crlf)
	Local $iExit = RunWait($sCommand, "", @SW_HIDE)
	If $iExit<>0 Then Return SetError(2, 0, False)
	ConsoleWrite("Done creating "&$sTargetFile&@crlf)
	Return True
EndFunc

Func _fileGetFolder($sFile)
	Local $arPath = StringRegExp($sFile, "^(.*[\/\\]).*$", 1)
	If UBound($arPath)<1 Then Return SetError(1, 1, False)
	Return $arPath[0]
EndFunc

Func _arraysCombine(ByRef $arArray1, ByRef $arArray2)
	Local $arResult[UBound($arArray1)+UBound($arArray2)]
	For $i=0 To UBound($arArray1)-1
		$arResult[$i] = $arArray1[$i]
	Next
	Local $iStart = UBound($arArray1)
	For $i=0 To UBound($arArray2)-1
		$arResult[$iStart+$i] = $arArray2[$i]
	Next
	Return $arResult
EndFunc