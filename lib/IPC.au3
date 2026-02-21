#include-once
#include <AutoItConstants.au3>
#include <WinAPIProc.au3>

; #INDEX# =======================================================================================================================
; Title .........: IPC (InterProcessCommunication)
; AutoIt Version : 3.3.18.1
; Language ......: English
; Description ...: UDF for inter process communication between the main and child processes using TCP.
;                  Strings can be send as well as integer commands.
; Author(s) .....: Kanashius
; Version .......: 1.1.0
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; __IPC_StartUp
; __IPC_Shutdown
; __IPC_GetScriptExecutable
; __IPC_StartProcess
; __IPC_SubGetPID
; __IPC_SubCheck
; __IPC_SubConnect
; __IPC_SubSend
; __IPC_SubSendCmd
; __IPC_MainSend
; __IPC_MainSendCmd
; __IPC_ProcessStop
; __IPC_SubProcessing
; __IPC_MainProcessing
; __IPC_Log
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __IPC__SendMsg
; __IPC__ToBinary
; __IPC__BinaryConsumeVar
; __IPC__BinaryConsumeBytes
; __IPC__SubDisconnect
; __IPC__ServerStart
; __IPC__ServerIsRunning
; __IPC__ServerStop
; __IPC__ServerProcessStdOut
; __IPC__ServerProcessLogStd
; __IPC__ServerAccept
; __IPC__AddSocket
; __IPC__ProcessMessages
; __IPC__ProcessMessagesAtSocket
; __IPC__SocketReadBytes
; __IPC__SocketDisconnect
; __IPC__ServerProcessRemove
; __IPC__ProcessIdToHandle
; __IPC__ProcessHandleToId
; ===============================================================================================================================

; #GLOBAL CONSTANTS# ============================================================================================================
Global Const $__IPC_LOG_FATAL = 1, $__IPC_LOG_ERROR = 2, $__IPC_LOG_WARN = 3, $__IPC_LOG_INFO = 4, $__IPC_LOG_DEBUG = 5
Global Const $__IPC_LOG_TRACE = 6
; ===============================================================================================================================

; #INTERNAL_USE_ONLY GLOBAL VARIABLES # =========================================================================================
Global Const $__IPC_CONN_TO_MAIN = 1, $__IPC_CONN_TO_SUB = 2
Global Const $__IPC_MSG_CONNECT = Int(1, 1), $__IPC_MSG_DISCONNECT = Int(2, 1), $__IPC_MSG_ACK = Int(3, 1)
Global Const $__IPC_MSG_DATA = Int(4, 1), $__IPC_MSG_CMD = Int(5, 1), $__IPC_MSG_CMD_DATA = Int(6, 1)
Global Const $__IPC_Port = 40001, $__IPC_MainPullRate = 100, $__IPC_MaxByteRecv = 1024, $__IPC_SubPullRate = 100
Global Const $__IPC_PARAM_CONNECT = "--IPC-CONNECT"
Global Const $__IPC_DataType_Binary = Int(1, 1), $__IPC_DataType_Bool = Int(2, 1), $__IPC_DataType_Func = Int(3, 1)
Global Const $__IPC_DataType_Double = Int(4, 1), $__IPC_DataType_Ptr = Int(5, 1), $__IPC_DataType_Hwnd = Int(6, 1)
Global Const $__IPC_DataType_Int32 = Int(7, 1), $__IPC_DataType_Int64 = Int(8, 1), $__IPC_DataType_Object = Int(9, 1)
Global Const $__IPC_DataType_String = Int(10, 1), $__IPC_DataType_Keyword = Int(11, 1), $__IPC_DataType_DLLStruct = Int(12, 1)
Global Const $__IPC_DataType_Map = Int(13, 1), $__IPC_DataType_Array = Int(14, 1)
Global $__IPC__Data[]
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_StartUp
; Description ...: StartUp of the ICP UDF initializing required variables. Must be called before using other UDF functions.
; Syntax ........: __IPC_StartUp([$iLogLevel = $__IPC_LOG_INFO[, $iMainPullRate = Default[, $iMaxReceiveCount = Default[, $iMainPort = Default]]]])
; Parameters ....: $iLogLevel             - [optional] Default: $__IPC_LOG_INFO. All logging equal or lower to the level will be shown.
;                  $iMainPullRate         - [optional] Default: 100 ms. How often the main process looks for new data.
;                  $iMaxReceiveCount      - [optional] Default: Default. Unlimited, however much data can be received.
;                  $iMainPort             - [optional] Default: 40001. The port to start looking for an open port for the TCPServer.
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
;                 Possible log levels: $__IPC_LOG_FATAL, $__IPC_LOG_ERROR, $__IPC_LOG_WARN, $__IPC_LOG_INFO, $__IPC_LOG_DEBUG and $__IPC_LOG_TRACE = 6
;
;                 The $iMainPullRate defines how often the main process checks TCP/STDOUT/STDERR streams for new data or connections.
;                 This should not be set very low, because that may cause blocking of the AutoIt-Script (Freezing).
;                 If a lot of data is sent, consider setting the $iMainPullRate higher to avoid application freezes (see AdlibRegister blocking).
;                 Setting the $iMainPullRate to 0 will disable the automatic handling of data (AdlibRegister). Then the __IPC_MainProcessing
;                 function must be called manually.
;
;                 $iMainPort can be set, but it will not garantee the usage of that port. The Script will start with that number
;                 and try to listen at that port. If it is already bound, the port will increase, until a free port is found.
;                 If none is found, __IPC_StartProcess will return an error (when the TCPServer is started).
;
;                 $iMaxReceiveCount can be set to limit the number of calls to TCPRecv, receiving $__IPC_MaxByteRecv bytes of data.
;                 This can be useful, if large amount of data are sent in a short time, to avoid the main process from freezing.
;                 Together with manual handling of data (see $iMainPullRate), this can be used to ensure that the main process does not freeze.
;                 $iMaxReceiveCount resets for every sub process => every sub process gets pulled $iMaxReceiveCount times.
;
;                 Errors:
;                 1 - Parameter not valid (@extended: 1 - $iLogLevel, 2 - $iMainPullRate, 3 - $iMainPort)
;                 2 - __IPC_StartUp was already called, call __IPC_Shutdown first
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_StartUp($iLogLevel = $__IPC_LOG_INFO, $iMainPullRate = Default, $iMaxReceiveCount = Default, $iMainPort = Default)
	; handle parameters
	If Not UBound(MapKeys($__IPC__Data))=0 Then Return SetError(2, 0, False)
	If $iMainPullRate = Default Then $iMainPullRate = $__IPC_MainPullRate
	If $iMainPort = Default Then $iMainPort = $__IPC_Port
	If Not IsInt($iLogLevel) Or $iLogLevel<0 Or $iLogLevel>$__IPC_LOG_TRACE Then Return SetError(1, 1, False)
	If Not IsInt($iMainPullRate) Or $iMainPullRate<0 Then Return SetError(1, 2, False)
	If $iMaxReceiveCount <> Default And (Not IsInt($iMaxReceiveCount) Or $iMaxReceiveCount < 0) Then Return SetError(1, 3, False)
	If Not IsInt($iMainPort) Or $iMainPort<1024 Or $iMainPort>65535 Then Return SetError(1, 4, False)
	If Not MapExists($__IPC__Data, "iLogLevel") Then $__IPC__Data.iLogLevel = $iLogLevel
	; init tcp
	$__IPC__Data.iStartUp = TCPStartup()
	$__IPC__Data.iMaxReceiveCount = $iMaxReceiveCount ; how often tcprecv is called during processing, if more data is available
	; init socket map
	Local $mConnects[]
	$__IPC__Data.mConnects = $mConnects
	; init server data map
	Local $mServer[], $mProcesses[]
	$mServer.mProcesses = $mProcesses ; all started processes are stored here
	$mServer.iMainPullRate = $iMainPullRate
	$mServer.iMainStartPort = $iMainPort
	$mServer.iListen = Default ; the tcplisten handle
	$mServer.iPort = Default ; the actually used port
	$mServer.iOpenProcesses = 0 ; the number of processes started, but not yet successfully identified at the socket
                                ; ($__IPC_MSG_CONNECT with $hProcess to assign the socket to the process)
	$__IPC__Data.mServer = $mServer
	; init client data map
	Local $mClient[]
	$mClient.iPullRate = $__IPC_SubPullRate
	$mClient.iSocket = Default ; the subprocess tcp socket
	$mClient.sCallback = Default ; the callback of the sub process for message handling
	$mClient.sExitCallback = Default ; the callback called when the connection to the main process is closed/lost
	$__IPC__Data.mClient = $mClient
	__IPC_Log($__IPC_LOG_INFO, "IPC started")
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_Shutdown
; Description ...: Shutdown of the ICP UDF. Should be called to end the usage of the IPC UDF (or on exit).
; Syntax ........: __IPC_Shutdown()
; Parameters ....:
; Return values .: True on success, False if __IPC_StartUp was never called.
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_Shutdown()
	If UBound(MapKeys($__IPC__Data))=0 Then Return False ; startup was not called or shutdown was already called
	; disconnect from all processes and send them the message to terminate
	Local $arProcesses = MapKeys($__IPC__Data.mServer.mProcesses)
	For $i=0 To UBound($arProcesses)-1
		__IPC__ServerProcessRemove($arProcesses[$i], True)
	Next
	; disconnect if it is a subprocess (main process happens nothing)
	__IPC__SubDisconnect()
	; shutdown tcp if it was started successfully
	If MapExists($__IPC__Data, "iStartUp") And $__IPC__Data.iStartUp=1 Then TCPShutdown()
	; reset data map
	Local $mData[]
	$__IPC__Data = $mData
	__IPC_Log($__IPC_LOG_INFO, "IPC shutdown")
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_GetScriptExecutable
; Description ...: Returns the $sExecutable parameter for __IPC_StartProcess, when an .au3 or .exe with the name exists.
; Syntax ........: __IPC_GetScriptExecutable($sScriptPath)
; Parameters ....: $sScriptPath         -
; Return values .: The $sExecutable for the __IPC_StartProcess function.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Provides the AutoIt executable to the path of any .au3 file.
;                 Automatically completes the filename with the .au3 or .exe extension, if such a file exists.
;                 The .au3 will be preferred.
;                 e.g.: "IPC-Example-SubSeperated" will automatically start the corresponding .au3 file, if present.
;                 If not, but a .exe is present, that will be executed.
;                 If "IPC-Example-SubSeperated.au3" is given, the result will be an AutoIt executable with the script as parameter.
;                 If the main process is a compiled AutoIt executable, it includes the AutoIt Interpreter.
;                 So calling it with a Script as parameter will execute that Script.
;
;                 Errors:
;                 1 - Parameter $sScriptPath not valid (@extended: 1 - no possible .au3/.exe file exists,
;                     2 - the file is not a .au3 or .exe)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_GetScriptExecutable($sScriptPath)
	; handle $sScriptPath, if it is provided without the extension => autocomplete
	If Not FileExists($sScriptPath) Then
		If FileExists($sScriptPath&".au3") Then
			$sScriptPath = $sScriptPath&".au3"
		ElseIf FileExists($sScriptPath&".exe") Then
			$sScriptPath = $sScriptPath&".exe"
		Else
			Return SetError(1, 1, False)
		EndIf
	EndIf
	; check the extension and return the desired executable (with arguments)
	Local $sExt = StringRight($sScriptPath, 4)
	If $sExt = ".au3" Then Return '"'&@AutoItExe&'" /AutoIt3ExecuteScript "'&$sScriptPath&'"'
	If $sExt = ".exe" Then Return $sScriptPath
	Return SetError(1, 2, False)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_StartProcess
; Description ...: Starts a sub process from the main process.
; Syntax ........: __IPC_StartProcess([$sCallback = Default[, $arguments = Default[, $sDoneCallback = Default[, $sExecutable = Default[, $sWorkingDir = ""[, $show = @SW_HIDE[, $opt_flag = BitOR($STDOUT_CHILD, $STDERR_CHILD)]]]]]]])
; Parameters ....: $sCallback         - [optional] Default: no callback. The callback function to call for incoming data from the subprocess.
;                  $arguments         - [optional] Default: no arguments. The arguments to call the subprocess with. Can be a String or a 1D-Array.
;                  $sDoneCallback     - [optional] Default: no callback. This function is called with the $hProcess, when the sub process is closed.
;                  $sExecutable       - [optional] Default: the script itself. The executable to be executed.
;                  $sWorkingDir       - [optional] Default: the script directory. The working dir for the subprocess.
;                  $show              - [optional] Default: @SW_HIDE. See function: "Run".
;                  $opt_flag          - [optional] Default: $STDOUT_CHILD+$STDERR_CHILD. See function: "Run".
; Return values .: The $hSubProcess of the started subprocess. 0 on failure.
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
;                 $sCallback must be a function with 3 parameters ($hSubProcess, $iCmd, $arData).
;                 $iCmd is the integer command and $arData is a 1D-Array with all values sent.
;                 $iCmd or $arData can be Default, if they were not sent.
;
;                 If $arguments is a 1D-Array, all not string values are converted to string with String().
;
;                 $sDoneCallback must be a function with 1 parameter ($hProcess). This function will be called, when the sub process does not exist anymore (ProcessExist).
;
;                 If $sExecutable should be different then the script of the main process, it can be provided here. This can be anything Run() can execute with parameters.
;                 If it should be another script, __IPC_GetScriptExecutable can be used to get the path to the .au3 or .exe (whichever is present).
;
;                 To disable console output of the subprocess, $opt_flag can be set to 0 (Or any other $opt_flag, see Run() ).
;
;                 Errors:
;                 1 - Parameter invalid (@extended: 1 - $sCallback, 2 - $arguments, 3 - $sDoneCallback, 4 - $sExecutable, 5 - $sWorkingDir, 6 - $show, 7 - $opt_flag)
;                 100 - Calling __IPC__ServerStart failed. @extended contains the @error of__IPC__ServerStart
;                 200 - Error creating process handle
;                 ? - Look at possible @error/@extended from Run()
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_StartProcess($sCallback = Default, $arguments = Default, $sDoneCallback = Default, $sExecutable = Default, $sWorkingDir = "", $show = @SW_HIDE, $opt_flag = BitOR($STDOUT_CHILD, $STDERR_CHILD))
	; start the main process server, if it is not running already
	If Not __IPC__ServerIsRunning() Then
		__IPC__ServerStart()
		If @error Then Return SetError(100, @error, 0)
	EndIf
	; handle all parameters
	If $sExecutable=Default Then $sExecutable = __IPC_GetScriptExecutable(@ScriptFullPath)
	If $sCallback<>Default And Not IsFunc(Execute($sCallback)) Then Return SetError(1, 1, -1)
	If $arguments<>Default And Not IsString($arguments) And Not IsArray($arguments) Then Return SetError(1, 2, 0)
	If $sDoneCallback<>Default And Not IsFunc(Execute($sDoneCallback)) Then Return SetError(1, 3, -1)
	If $sExecutable<>Default And Not IsString($sExecutable) Then Return SetError(1, 4, 0)
	If Not IsString($sWorkingDir) Then Return SetError(1, 5, 0)
	If $show<>@SW_SHOW And $show<>@SW_HIDE And $show<>@SW_MINIMIZE And $show<>@SW_MAXIMIZE Then Return SetError(1, 6, 0)
	If Not IsInt($opt_flag) Then Return SetError(1, 7, 0)
	; create process data
	Local $mProcess[]
	$mProcess.bStdErr = (BitAND($opt_flag, $STDERR_CHILD)?(True):(False))
	$mProcess.bStdOut = ((BitAND($opt_flag, $STDOUT_CHILD) Or BitAND($opt_flag, $STDERR_MERGED))?(True):(False))
	$mProcess.iSocket = Default
	$mProcess.bWaitForSocket = True
	$mProcess.sCallback = $sCallback
	$mProcess.sExitCallback = $sDoneCallback
	; create process handle
	Local $iProcess = MapAppend($__IPC__Data.mServer.mProcesses, $mProcess)
	If @error Then Return SetError(200, 0, 0)
	Local $hProcess = __IPC__ProcessIdToHandle($iProcess)
	; create arguments
	Local $sArguments = $__IPC_PARAM_CONNECT&" "&$__IPC__Data.mServer.iPort&" "&$hProcess
	If IsString($arguments) Then
		$sArguments &= " "&$arguments
	ElseIf UBound($arguments)>0 Then
		For $i=0 to UBound($arguments)-1 Step 1
			$sArguments&=' "'&StringReplace($arguments[$i], '"', '""')&'"'
		Next
	EndIf
	__IPC_Log($__IPC_LOG_INFO, "Start process: "&$sExecutable)
	__IPC_Log($__IPC_LOG_DEBUG, @TAB&" with arguments: "&$sArguments)
	; start the sub process
	Local $sCmd = $sExecutable&" "&$sArguments
	Local $iPID = Run($sCmd, $sWorkingDir, $show, $opt_flag)
	Local $iError = @error, $iExtended = @extended
	; save additional process information
	$__IPC__Data["mServer"]["mProcesses"][$iProcess]["iPID"] = $iPID
	$__IPC__Data["mServer"]["mProcesses"][$iProcess]["hProcess"] = $hProcess
	; add one process to the number of processes started, but not yet successfully connected ($__IPC_MSG_CONNECT with $hProcess)
	$__IPC__Data["mServer"]["iOpenProcesses"] += 1
	; remove all process data, if the startup failed
	If $iError Then
		__IPC__ServerProcessRemove($iProcess, True)
		__IPC_Log($__IPC_LOG_ERROR, "Failed process start: "&$sCmd)
		Return SetError($iError, $iExtended, 0)
	EndIf
	Return SetExtended($iExtended, $hProcess)
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_SubGetPID
; Description ...: Get the PID of a running sub process.
; Syntax ........: __IPC_SubGetPID($hProcess)
; Parameters ....: $hProcess         - the sub process handle
; Return values .: The PID. 0 on failure.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $hProcess)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_SubGetPID($hProcess)
	Local $iProcess = __IPC__ProcessHandleToId($hProcess)
	If @error Then Return SetError(1, 1, 0)
	Return $__IPC__Data.mServer.mProcesses[$iProcess].iPID
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_SubCheck
; Description ...: Check if the script is running as sub process and call the corresponding sub process function.
; Syntax ........: __IPC_SubCheck($sFunctionSub[, $sFunctionMain = Default[, $sCallback = Default[, $sExitCallback = Default[, $iLogLevel = $__IPC_LOG_INFO[, $iPullRate = Default]]]]])
; Parameters ....: $sFunctionSub         - the function to call for the sub process execution.
;                  $sFunctionMain        - [optional] Default: none. The function to call for the main process execution.
;                  $sCallback            - [optional] Default: no callback. The callback for the sub process messages.
;                  $sExitCallback        - [optional] Default: no callback. The callback for the disconnect/close command from the main process.
;                  $iLogLevel            - [optional] Default: $__IPC_LOG_INFO. All logging equal or lower to the level will be shown.
;                  $iPullRate            - [optional] Default: 100 ms. How often the sub process looks for new data.
; Return values .: The sub process handle on success, 0 otherwise.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: $sFunctionSub must be a function with 1 parameter ($hSubProcess). It is called, when the script is executed as sub process.
;
;                  $sFunctionMain must be a function without parameters. It is called, when the script is executed as main process.
;
;                  $sCallback must be a function with 2 parameters ($iCmd, $arData).
;                  $iCmd contains the integer command and $arData is a 1D-Array with all values sent.
;                  $iCmd or $arData can be Default, if they were not sent.
;
;                  $sExitCallback must be a function without parameters. This function will be called, when the main process disconnects from the sub process.
;
;                  Possible log levels: $__IPC_LOG_FATAL, $__IPC_LOG_ERROR, $__IPC_LOG_WARN, $__IPC_LOG_INFO, $__IPC_LOG_DEBUG and $__IPC_LOG_TRACE = 6
;
;                  The $iPullRate defines how often the sub process checks TCP streams for new data.
;                  This should not be set very low, because that may cause blocking of the AutoIt-Script (freezing).
;                  If a lot of data is sent, consider setting the $iPullRate higher to avoid application freezes (see AdlibRegister blocking).
;                  Setting the $iPullRate to 0 will disable the automatic handling of data (AdlibRegister). Then the __IPC_SubProcessing
;                  function must be called manually.
;
;                  Errors:
;                  1 - Parameter invalid (@extended: 1 - $sFunctionSub, 2 - $sFunctionMain, 3 - $sCallback, 4 - $sExitCallback, 5 - $iLogLevel, 6 - $iPullRate, 7 - $iMainPullRate)
;                  2 - __IPC_StartUp failed. @extended contains the error from __IPC_StartUp.
;                  3 - Connect to main process failed (TCPConnect)
;                  4 - Failed to sent the connect command to the main process with tcp
;                  5 - Calling $sFunctionSub failed.
;                  6 - Calling $sFunctionMain failed.
;                  7 - Parameter in commandline invalid (@extended: 1 - iPort, 2 - hProcess)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_SubCheck($sFunctionSub, $sFunctionMain = Default, $sCallback = Default, $sExitCallback = Default, $iLogLevel = $__IPC_LOG_INFO, $iPullRate = Default)
	; handle parameters
	If Not IsFunc(Execute($sFunctionSub)) Then Return SetError(1, 1, 0)
	If $sFunctionMain<>Default And Not IsFunc(Execute($sFunctionMain)) Then Return SetError(1, 2, 0)
	; other parameters will be checked/handled in __IPC_SubConnect

	; check if the script was started with commandline argument $__IPC_PARAM_CONNECT and the corresponding port and process number
	Local $iPort = Default, $hProcess = Default
	If UBound($CmdLine)>=4 Then
		If $CmdLine[1]=$__IPC_PARAM_CONNECT Then
			$iPort = Int($CmdLine[2])
			$hProcess = Int($CmdLine[3])
		EndIf
		Local $arCmdLine[UBound($CmdLine)-3]
		$arCmdLine[0] = UBound($arCmdLine)-1
		For $i=4 to UBound($CmdLine)-1
			$arCmdLine[$i-3] = $CmdLine[$i]
		Next
		$CmdLine = $arCmdLine
	EndIf
	; no subprocess detected
	If $iPort=Default Or $hProcess=Default Then
		If $sFunctionMain<>Default Then
			__IPC_StartUp($iLogLevel)
			If @error And @error<>2 Then Return SetError(2, @error, False) ; ignore the "startup already called" message, otherwise return the startup error
			Call($sFunctionMain)
			If @error = 0xDEAD And @extended = 0xBEEF Then Return SetError(6, 0, 0)
		EndIf
		Return 0
	EndIf
	Local $hSubProcess = __IPC_SubConnect($iPort, $hProcess, $sCallback, $sExitCallback, $iLogLevel, $iPullRate)
	If @error=1 And @extended>2 Then Return SetError(1, @extended, 0)
	If @error=1 Then Return SetError(7, @extended, 0)
	If @error Then Return SetError(@error, @extended, 0)
	Call($sFunctionSub, $hSubProcess)
	If @error = 0xDEAD And @extended = 0xBEEF Then Return SetError(5, 0, 0)
	Return $hSubProcess
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_SubConnect
; Description ...: Connect to the main process.
; Syntax ........: __IPC_SubConnect($iPort, $hProcess)
; Parameters ....: $iPort         - the port the main process is listening at.
;                  $hProcess      - the hSubProcessHandle of this script.
;                  $sCallback            - [optional] Default: no callback. The callback for the sub process messages.
;                  $sExitCallback        - [optional] Default: no callback. The callback for the disconnect/close command from the main process.
;                  $iLogLevel            - [optional] Default: $__IPC_LOG_INFO. All logging equal or lower to the level will be shown.
;                  $iPullRate            - [optional] Default: 100 ms. How often the sub process looks for new data.
; Return values .: The sub process handle on success, 0 otherwise.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: $iPort must be and integer >=1024 and <65535.
;                  $hProcess must be an integer >=1 and must correspond to the process handle at the main process.
;
;                  For more information on $sCallback, $sExitCallback, $iLogLevel and $iPullRate see __IPC_SubCheck.
;
;                  Errors:
;                  1 - Parameter invalid (@extended: 1 - $iPort, 2 - $hProcess)
;                  2 - Connect to main process failed (TCPConnect)
;                  3 - __IPC_StartUp failed. @extended contains the error from __IPC_StartUp.
;                  4 - Failed to sent the connect command to the main process with tcp
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_SubConnect($iPort, $hProcess, $sCallback = Default, $sExitCallback = Default, $iLogLevel = $__IPC_LOG_INFO, $iPullRate = Default)
	If Not IsInt($iPort) Or $iPort<1024 Or $iPort>65535 Then Return SetError(1, 1, 0)
	If Not IsInt($hProcess) Or $hProcess<1 Then Return SetError(1, 2, 0)
	If $sCallback<>Default And Not IsFunc(Execute($sCallback)) Then Return SetError(1, 3, 0)
	If $sExitCallback<>Default And Not IsFunc(Execute($sExitCallback)) Then Return SetError(1, 4, 0)
	If $iPullRate = Default Then $iPullRate = $__IPC_SubPullRate
	If Not IsInt($iLogLevel) Or $iLogLevel<0 Or $iLogLevel>$__IPC_LOG_TRACE Then Return SetError(1, 5, False)
	If Not IsInt($iPullRate) Or $iPullRate<0 Then Return SetError(1, 6, False)
	; handle detected sub process
	__IPC_StartUp($iLogLevel)
	If @error And @error<>2 Then Return SetError(2, @error, 0) ; ignore the "startup already called" message, otherwise return the startup error
	; connect to main process
	__IPC_Log($__IPC_LOG_INFO, "Connect: "&$iPort&" >> "&$hProcess)
	Local $iSocket = TCPConnect("127.0.0.1", $iPort)
	If @error Then
		__IPC_Log($__IPC_LOG_ERROR, "Could not connect to main process: "&$iPort&" > "&$hProcess)
		Return SetError(3, 0, 0)
	EndIf
	; connect to the main process
	TCPSend($iSocket, Binary($__IPC_MSG_CONNECT)&Binary($hProcess))
	IF @error Then  Return SetError(4, @error, 0)
	; save data for the connection
	$__IPC__Data.mClient.iSocket = $iSocket
	$__IPC__Data.mClient.sCallback = $sCallback
	$__IPC__Data.mClient.sExitCallback = $sExitCallback
	$__IPC__Data.mClient.iPullRate = $iPullRate
	If $iPullRate>0 And $__IPC__Data.mClient.sCallback<>Default Then AdlibRegister("__IPC_SubProcessing", $__IPC__Data.mClient.iPullRate)
	__IPC__AddSocket($iSocket, $__IPC_CONN_TO_MAIN)
	Return $hProcess
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_SubSend
; Description ...: Send data to the main process
; Syntax ........: __IPC_SubSend($data1, [$data2 = Default, [$data3 = Default, [$data4 = Default, [$data5 = Default,
;                                [$data6 = Default, [$data7 = Default, [$data8 = Default, [$data9 = Default, [$data10 = Default]]]]]]]]])
; Parameters ....: $hProcess           - the sub process handle
;                  $data1              - the data
;                  $data2              - [optional] the data (if provided, it will be send, even if default)
;                  ...                 - ...
;                  $data10             - [optional] the data (if provided, it will be send, even if default)
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid
;                  2 - TCP Send failed
;                  3 - Not connected to sub process
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_SubSend($data1, $data2 = Default, $data3 = Default, $data4 = Default, $data5 = Default, $data6 = Default, _
	                $data7 = Default, $data8 = Default, $data9 = Default, $data10 = Default)
	If $__IPC__Data.mClient.iSocket=Default Then Return SetError(3, 0, False)

	Local $arArgs[@NumParams+3]
    $arArgs[0] = "CallArgArray"
	$arArgs[1] = $__IPC__Data.mClient.iSocket
	$arArgs[2] = Default
	For $i=1 To @NumParams
		$arArgs[$i+2] = Eval("data"&$i)
	Next

	Local $bResult = Call("__IPC__SendMsg", $arArgs)
	If @error Then Return SetError(@error, @extended-2, $bResult)
	Return $bResult
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_SubSendCmd
; Description ...: Send an integer command (with data) to the main process
; Syntax ........: __IPC_SubSendCmd($hProcess, $iCmd, [$data1 = Default, [$data2 = Default, [$data3 = Default, [$data4 = Default,
;                                   [$data5 = Default, [$data6 = Default, [$data7 = Default, [$data8 = Default, [$data9 = Default,
;                                   [$data10 = Default]]]]]]]]]])
; Parameters ....: $iCmd               - integer as command
;                  $data1              - [optional] the data (if provided, it will be send, even if default)
;                  ...                 - ...
;                  $data10             - [optional] the data (if provided, it will be send, even if default)
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (1 - $iCmd)
;                  2 - TCP Send failed
;                  3 - Not connected to sub process
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_SubSendCmd($iCmd, $data1 = Default, $data2 = Default, $data3 = Default, $data4 = Default, $data5 = Default, $data6 = Default, _
	                $data7 = Default, $data8 = Default, $data9 = Default, $data10 = Default)
	If $__IPC__Data.mClient.iSocket=Default Then Return SetError(3, 0, False)
	If Not IsInt($iCmd) Then Return SetError(1, 2, False)

	Local $arArgs[@NumParams+2]
    $arArgs[0] = "CallArgArray"
	$arArgs[1] = $__IPC__Data.mClient.iSocket
	$arArgs[2] = $iCmd
	For $i=1 To @NumParams-1
		$arArgs[$i+2] = Eval("data"&$i)
	Next

	Local $bResult = Call("__IPC__SendMsg", $arArgs)
	If @error Then Return SetError(@error, @extended-1, $bResult)
	Return $bResult
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_MainSend
; Description ...: Send data to a sub process
; Syntax ........: __IPC_MainSend($hProcess, $data1, [$data2 = Default, [$data3 = Default, [$data4 = Default, [$data5 = Default,
;                                 [$data6 = Default, [$data7 = Default, [$data8 = Default, [$data9 = Default,
;                                 [$data10 = Default]]]]]]]]])
; Parameters ....: $hProcess           - the sub process handle
;                  $data1              - the data
;                  $data2              - [optional] the data (if provided, it will be send, even if default)
;                  ...                 - ...
;                  $data10             - [optional] the data (if provided, it will be send, even if default)
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $hProcess)
;                  2 - TCP Send failed
;                  3 - Not connected to sub process
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_MainSend($hProcess, $data1, $data2 = Default, $data3 = Default, $data4 = Default, $data5 = Default, $data6 = Default, _
	                $data7 = Default, $data8 = Default, $data9 = Default, $data10 = Default)
	Local $iProcess = __IPC__ProcessHandleToId($hProcess)
	If @error Then Return SetError(1, 1, False)

	Local $iSocket = $__IPC__Data.mServer.mProcesses[$iProcess].iSocket
	If $iSocket=Default Then Return SetError(3, 0, False)

	Local $arArgs[@NumParams+2]
    $arArgs[0] = "CallArgArray"
	$arArgs[1] = $iSocket
	$arArgs[2] = Default
	For $i=1 To @NumParams
		$arArgs[$i+2] = Eval("data"&$i)
	Next

	Local $bResult = Call("__IPC__SendMsg", $arArgs)
	If @error=1 Then Return SetError(@error, @extended, $bResult)
	If @error Then Return SetError(@error, @extended, $bResult)
	Return $bResult
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_MainSendCmd
; Description ...: Send an integer command (with data) to a sub process
; Syntax ........: __IPC_MainSendCmd($hProcess, $iCmd, [$data1 = Default, [$data2 = Default, [$data3 = Default, [$data4 = Default,
;                                    [$data5 = Default, [$data6 = Default, [$data7 = Default, [$data8 = Default,
;                                    [$data9 = Default, [$data10 = Default]]]]]]]]]])
; Parameters ....: $hProcess           - the sub process handle
;                  $iCmd               - integer as command
;                  $data1              - [optional] the data (if provided, it will be send, even if default)
;                  ...                 - ...
;                  $data10             - [optional] the data (if provided, it will be send, even if default)
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $hProcess, 2 - $iCmd)
;                  2 - TCP Send failed
;                  3 - Not connected to sub process
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_MainSendCmd($hProcess, $iCmd, $data1 = Default, $data2 = Default, $data3 = Default, $data4 = Default, $data5 = Default, _
					   $data6 = Default, $data7 = Default, $data8 = Default, $data9 = Default, $data10 = Default)
	If Not IsInt($iCmd) Then Return SetError(1, 2, False)
	Local $iProcess = __IPC__ProcessHandleToId($hProcess)
	If @error Then Return SetError(1, 1, False)

	Local $iSocket = $__IPC__Data.mServer.mProcesses[$iProcess].iSocket
	If $iSocket=Default Then Return SetError(3, 0, False)

	Local $arArgs[@NumParams+2]
    $arArgs[0] = "CallArgArray"
	$arArgs[1] = $iSocket
	$arArgs[2] = $iCmd
	For $i=2 To @NumParams-1
		$arArgs[$i+1] = Eval("data"&($i-1))
	Next

	Local $bResult = Call("__IPC__SendMsg", $arArgs)
	If @error=1 Then Return SetError(@error, @extended, $bResult)
	If @error Then Return SetError(@error, @extended, $bResult)
	Return $bResult
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_ProcessStop
; Description ...: Disconnect from a sub process. This triggers a call to the sub process $sExitCallback (if provided).
; Syntax ........: __IPC_ProcessStop($hProcess)
; Parameters ....: $hProcess           - the sub process handle
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $hProcess)
;                  2 - __IPC__SocketDisconnect failed. (@extended: @error from __IPC__SocketDisconnect)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_ProcessStop($hProcess)
	Local $iProcess = __IPC__ProcessHandleToId($hProcess)
	If @error Then Return SetError(1, 1, False)

	__IPC__SocketDisconnect($__IPC__Data.mServer.mProcesses[$iProcess].iSocket)
	If @error Then Return SetError(2, @error, False)
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_SubProcessing
; Description ...: Process messages for the sub process. Only call if the sub process pullrate $iPullRate was set to 0.
; Syntax ........: __IPC_SubProcessing()
; Parameters ....:
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_SubProcessing()
	__IPC__ProcessMessages()
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_MainProcessing
; Description ...: Process messages for the main process. Only call if the main process pullrate $iMainPullRate was set to 0.
; Syntax ........: __IPC_MainProcessing()
; Parameters ....:
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_MainProcessing()
	If $__IPC__Data.mServer.iOpenProcesses>0 Then __IPC__ServerAccept()
	__IPC__ProcessMessages()
	__IPC__ServerProcessStdOut()
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: __IPC_Log
; Description ...: Log a message to the console with the provided log level.
; Syntax ........: __IPC_Log($iLevel, $sMsg[, $iError=Default[, $iExtended = Default[, $hSubProcess = Default[, $iLine = @ScriptLineNumber]]]])
; Parameters ....: $iLevel           - the sub process handle
;                  $sMsg             - the string to print
;                  $iError           - [optional] Default: None. An error if some should be printed.
;                  $iExtended        - [optional] Default: None. Extended data if some should be printed.
;                  $hSubProcess      - [optional] Default: None. The sub process, where the log belongs to.
;                  $iLine            - [optional] Default: @ScriptLineNumber. The Scriptline to print
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $iLevel, 6 - $iLine)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC_Log($iLevel, $sMsg, $iError=Default, $iExtended = Default, $hSubProcess = Default, $iLine = @ScriptLineNumber)
	; handle parameters
	If Not IsInt($iLevel) Or $iLevel<0 Or $iLevel>$__IPC_LOG_TRACE Then Return SetError(1, 1, False)
	If Not IsInt($iLine) Or $iLevel<0 Then Return SetError(1, 6, False)
	If $iLevel>$__IPC__Data.iLogLevel Then Return False
	; handle $iError and $iExtended
	Local $sError = ""
	If $iError<>Default And $iExtended<>Default Then
		$sError = " [Error: "&$iError&", Extended: "&$iExtended&"]"
	ElseIf $iError<>Default Then
		$sError = " [Error: "&$iError&"]"
	ElseIf $iExtended<>Default Then
		$sError = " [Extended: "&$iExtended&"]"
	EndIf
	; create log prefix depending on log level
	Local $sLevel
	Switch $iLevel
		Case $__IPC_LOG_FATAL
			$sLevel = "FATAL"
		Case $__IPC_LOG_ERROR
			$sLevel = "ERROR"
		Case $__IPC_LOG_WARN
			$sLevel = "WARN "
		Case $__IPC_LOG_INFO
			$sLevel = "INFO "
		Case $__IPC_LOG_DEBUG
			$sLevel = "DEBUG"
		Case $__IPC_LOG_TRACE
			$sLevel = "TRACE"
	EndSwitch
	Local $arLines = StringSplit($sMsg, @CRLF, 3)
	For $i=0 to UBound($arLines)-1 Step 1
		If $hSubProcess<>Default Then
			ConsoleWrite(">"&"Sub["&$hSubProcess&"] "&$sLevel&" ["&@YEAR&"/"&@MON&"/"&@MDAY&" "&@HOUR&":"&@MIN&":"&@SEC&"."&@MSEC&"] "&$arLines[$i])
		Else
			ConsoleWrite(">"&$sLevel&" ["&@YEAR&"/"&@MON&"/"&@MDAY&" "&@HOUR&":"&@MIN&":"&@SEC&"."&@MSEC&"] "&$arLines[$i])
		EndIf
		If $i=UBound($arLines)-1 Then ConsoleWrite($sError)
		ConsoleWrite(@crlf)
	Next
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__SendMsg
; Description ...: Load or update the content of a treeview item to fill it with files/folders/drives
; Syntax ........: __IPC__SendMsg($iSocket[, $iCmdOrData[, $data = Default]])
; Parameters ....: $iSocket       - the tcp socket
;                  $iCmd          - [optional] the integer command or default for no command
;                  $data1         - [optional] the data (if provided, it will be send, even if default)
;                  ...            - ...
;                  $data10        - [optional] the data (if provided, it will be send, even if default)
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $iCmdOrData)
;                  2 - TCPSend failed
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__SendMsg($iSocket, $iCmd = Default, $data1 = Default, $data2 = Default, $data3 = Default, $data4 = Default, _
					$data5 = Default, $data6 = Default, $data7 = Default, $data8 = Default, $data9 = Default, $data10 = Default)
	; check if data is a string
	Local $iDataParamCount = @NumParams-2
	Local $dData = __IPC__ToBinary($iDataParamCount)
	For $i=0 To $iDataParamCount-1
		$dData &= __IPC__ToBinary(Eval("data"&$i+1))
		If @error Then Return SetError(1, $i+2, False)
	Next
	Local $iMsg = $__IPC_MSG_DATA
	; handle the cmd parameter
	If $iCmd<>Default Then
		; check cmd parameter if it is an integer
		If Not IsInt($iCmd) Then Return SetError(1, 1, False)
		$iMsg = $__IPC_MSG_CMD
		If $iDataParamCount>0 Then $iMsg = $__IPC_MSG_CMD_DATA
		; prepend the command to the data
		$dData = __IPC__ToBinary($iCmd) & $dData
	EndIf
	; check the data length
	Local $iLen = BinaryLen($dData)
	; send the message binary data length as integer first
	TCPSend($iSocket, Binary($iMsg)&Binary(Int($iLen, 2)))
	If @error Then Return SetError(2, __IPC__SubDisconnect(True), False)
	; send the data in $__IPC_MaxByteRecv sized packages
	For $i=1 To $iLen Step $__IPC_MaxByteRecv
		Local $dSend = BinaryMid($dData, $i, $__IPC_MaxByteRecv)
		TCPSend($iSocket, $dSend)
		If @error Then Return SetError(2, __IPC__SubDisconnect(True), False)
	Next
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ToBinary
; Description ...: Convert variable to binary data
; Syntax ........: __IPC__ToBinary($data, $bPrefixed = True)
; Parameters ....: $data          - the variable to convert to binary
; Return values .: The binary data on success
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $iCmdOrData)
;                  2 - Unknown datatype
;                  3 - Could not be converted to binary
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ToBinary($data)
	Local $bAddSize = False, $iDataType = 0, $dData = Default
	Switch VarGetType($data)
		Case "Binary"
			$iDataType = $__IPC_DataType_Binary
			$dData = $data
			$bAddSize = True
		Case "Bool"
			$iDataType = $__IPC_DataType_Bool
			$dData = Binary($data)
		Case "UserFunction", "Function"
			$iDataType = $__IPC_DataType_Func
			$dData = StringToBinary(FuncName($data), 4)
			$bAddSize = True
		Case "Double"
			$iDataType = $__IPC_DataType_Double
			$dData = Binary($data)
		Case "Ptr" ; => also handles
			$iDataType = IsHWnd($data)?$__IPC_DataType_Hwnd:$__IPC_DataType_Ptr
			$dData = Binary($data)
			$bAddSize = True
		Case "Int32"
			$iDataType = $__IPC_DataType_Int32
			$dData = Binary($data)
		Case "Int64"
			$iDataType = $__IPC_DataType_Int64
			$dData = Binary($data)
		Case "Object"
			#cs
			Local $iPid = _WinAPI_GetCurrentProcessID()
			__IPC__ToBinary($iPid) ; add to binary
			_WinAPI_DuplicateHandle()

			Local $sCLSID = ObjName($data, $OBJ_PROGID); $OBJ_CLSID)
			Local $sIID = ObjName($data, $OBJ_IID)
			ConsoleWrite($sCLSID&@crlf)
			ConsoleWrite($sIID&@crlf)
			Local $oObj = ObjGet(ObjName($data, $OBJ_MODULE)) ; ObjCreateInterface($sCLSID, $sIID)
			$oObj.add(3, "drei")
			ConsoleWrite("---------1--------"&@crlf)
			For $vKey In $data
			   ConsoleWrite($vKey&" > "&$data.Item($vKey) & " , ")
			Next
			ConsoleWrite(@crlf)
			ConsoleWrite("---------2--------"&@crlf)
			For $vKey In $oObj
			   ConsoleWrite($vKey&" > "&$oObj.Item($vKey) & " , ")
			Next
			ConsoleWrite(@crlf)
			ConsoleWrite("-----------------"&@crlf)
			#ce
		Case "String"
			$iDataType = $__IPC_DataType_String
			$dData = StringToBinary($data, 2)
			$bAddSize = True
		Case "Keyword"
			$iDataType = $__IPC_DataType_Keyword
			$dData = StringToBinary($data, 4)
			$bAddSize = True
		Case "DLLStruct"
			#cs
			; ??? _WinAPI_DuplicateHandle()
			; maybe see objcreate to share memory across processes
			Local $iIndex = 1
			While True
				Local $elem = DllStructGetData($data, $iIndex)
				If IsString($elem) Then
					; iterate DllStructGetData($data, $iIndex, 1...) and retrieve every index seperately to get the size
				EndIf
				ConsoleWrite(VarGetType($elem)&" >> "&$elem&@crlf)
				$iIndex += 1
				If @error Then ExitLoop
			WEnd
			#ce
		Case "Map"
			$iDataType = $__IPC_DataType_Map
			Local $arKeys = MapKeys($data)
			Local $dData = __IPC__ToBinary(UBound($arKeys))
			For $i=0 to UBound($arKeys)-1
				$dData &= __IPC__ToBinary($arKeys[$i])
				$dData &= __IPC__ToBinary($data[$arKeys[$i]])
			Next
		Case "Array"
			$iDataType = $__IPC_DataType_Array
			Local $iDim = UBound($data, 0)
			$dData = __IPC__ToBinary($iDim)
			Switch $iDim
				Case 1
					$dData &= __IPC__ToBinary(UBound($data))
					For $i=0 to UBound($data)-1
						$dData &= __IPC__ToBinary($data[$i])
					Next
				Case 2
					$dData &= __IPC__ToBinary(UBound($data))
					$dData &= __IPC__ToBinary(UBound($data, 2))
					For $i=0 to UBound($data)-1
						For $j=0 to UBound($data, 2)-1
							$dData &= __IPC__ToBinary($data[$i][$j])
						Next
					Next
				Case 3
					$dData &= __IPC__ToBinary(UBound($data))
					$dData &= __IPC__ToBinary(UBound($data, 2))
					$dData &= __IPC__ToBinary(UBound($data, 3))
					For $i=0 to UBound($data)-1
						For $j=0 to UBound($data, 2)-1
							For $k=0 to UBound($data, 3)-1
								$dData &= __IPC__ToBinary($data[$i][$j][$k])
							Next
						Next
					Next
			EndSwitch
		Case Else
			Return SetError(2, 0, 0)
	EndSwitch
	If $dData=Default Then Return SetError(3, 0, 0)
	If $bAddSize Then
		Return Binary($iDataType)&__IPC__ToBinary(BinaryLen($dData))&$dData
	Else
		Return Binary($iDataType)&$dData
	EndIf
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__BinaryConsumeVar
; Description ...: Read a variable from binary data (and consume it from the binary data)
; Syntax ........: __IPC__BinaryConsumeVar(ByRef $dData)
; Parameters ....: $dData          - the binary data
; Return values .: The variable with the correct type on success
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  ?
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__BinaryConsumeVar(ByRef $dData)
	; todo error handling
	Local $iDataType = Int(__IPC__BinaryConsumeBytes($dData, 4))
	Switch $iDataType
		Case $__IPC_DataType_Binary
			Local $iSize = __IPC__BinaryConsumeVar($dData)
			Return __IPC__BinaryConsumeBytes($dData, $iSize)
		Case $__IPC_DataType_Bool
			Return Int(__IPC__BinaryConsumeBytes($dData, 1))?(True):(False)
		Case $__IPC_DataType_Func, $__IPC_DataType_Keyword
			Local $iSize = __IPC__BinaryConsumeVar($dData)
			Local $dBinaryVal = __IPC__BinaryConsumeBytes($dData, $iSize)
			Local $sName = BinaryToString($dBinaryVal, 4)
			Return Execute($sName)
		Case $__IPC_DataType_Double
			Local $dDouble = __IPC__BinaryConsumeBytes($dData, 8)
			Local $tDoubleStruct = DllStructCreate("byte[8]")
			DllStructSetData($tDoubleStruct, 1, $dDouble)
			Return DllStructGetData(DllStructCreate("double", DllStructGetPtr($tDoubleStruct)), 1)
		Case $__IPC_DataType_Ptr, $__IPC_DataType_Hwnd
			Local $iSize = __IPC__BinaryConsumeVar($dData)
			Local $dPtr = Ptr(Int(__IPC__BinaryConsumeBytes($dData, $iSize)))
			If $iDataType=$__IPC_DataType_Hwnd Then Return HWnd($dPtr)
			Return $dPtr
		Case $__IPC_DataType_Int32
			Return Int(__IPC__BinaryConsumeBytes($dData, 4), 1)
		Case $__IPC_DataType_Int64
			Return Int(__IPC__BinaryConsumeBytes($dData, 8), 2)
		Case $__IPC_DataType_Object
			; todo
			;Local $iPid = __IPC__BinaryConsumeVar($dData)
			;_WinAPI_GetProcessHandleCount
			;_WinAPI_DuplicateHandle(?, ?, ?, 0, True, 2) ; 2 => DUPLICATE_SAME_ACCESS
		Case $__IPC_DataType_String
			Local $iSize = __IPC__BinaryConsumeVar($dData)
			Return BinaryToString(__IPC__BinaryConsumeBytes($dData, $iSize), 2)
		Case $__IPC_DataType_DLLStruct
			; todo
		Case $__IPC_DataType_Map
			Local $iPairs = __IPC__BinaryConsumeVar($dData)
			Local $mMap[]
			For $i=0 to $iPairs-1
				Local $key = __IPC__BinaryConsumeVar($dData)
				Local $val = __IPC__BinaryConsumeVar($dData)
				$mMap[$key] = $val
			Next
			Return $mMap
		Case $__IPC_DataType_Array
			Local $iDim = __IPC__BinaryConsumeVar($dData)
			Switch $iDim
				Case 1
					Local $iSize = __IPC__BinaryConsumeVar($dData)
					Local $arData[$iSize]
					For $i=0 to UBound($arData)-1
						$arData[$i] = __IPC__BinaryConsumeVar($dData)
					Next
					Return $arData
				Case 2
					Local $iSize = __IPC__BinaryConsumeVar($dData)
					Local $iSize2 = __IPC__BinaryConsumeVar($dData)
					Local $arData[$iSize][$iSize2]
					For $i=0 to UBound($arData)-1
						For $j=0 to UBound($arData, 2)-1
							$arData[$i][$j] = __IPC__BinaryConsumeVar($dData)
						Next
					Next
					Return $arData
				Case 3
					Local $iSize = __IPC__BinaryConsumeVar($dData)
					Local $iSize2 = __IPC__BinaryConsumeVar($dData)
					Local $iSize3 = __IPC__BinaryConsumeVar($dData)
					Local $arData[$iSize][$iSize2][$iSize3]
					For $i=0 to UBound($arData)-1
						For $j=0 to UBound($arData, 2)-1
							For $k=0 to UBound($arData, 3)-1
								$arData[$i][$j][$k] = __IPC__BinaryConsumeVar($dData)
							Next
						Next
					Next
					Return $arData
			EndSwitch
		Case Else
			; todo error
	EndSwitch
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__BinaryConsumeBytes
; Description ...: Consume an amount of bytes from the binary variable and return them. The returned bytes are removed from $dData.
; Syntax ........: __IPC__BinaryConsumeBytes(ByRef $dData)
; Parameters ....: $dData          - the binary data
; Return values .: The binary data with $iBytes number of bytes
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $dData)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__BinaryConsumeBytes(ByRef $dData, $iBytes)
	If BinaryLen($dData)<$iBytes Then Return SetError(1, 1, False)
	Local $bResultData = BinaryMid($dData, 1, $iBytes)
	$dData = BinaryMid($dData, 1+$iBytes)
	Return $bResultData
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__SubDisconnect
; Description ...: Disconnect sub process
; Syntax ........: __IPC__SubDisconnect([$bErr = False])
; Parameters ....: $bErr       - [optional] Default False. True if disconnect is because of an error
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__SubDisconnect($bErr = False)
	If $__IPC__Data.mClient.iSocket=Default Then Return False ; not a sub process or already disconnected
	If $bErr Then __IPC_Log($__IPC_LOG_ERROR, "Connection to main process lost.")
	__IPC_Log($__IPC_LOG_DEBUG, "Close Client connection.")
	; tell main process that sub process is disconnecting (ignore error => connection already lost)
	Local $iSocket = $__IPC__Data.mClient.iSocket
	TCPSend($iSocket, Binary($__IPC_MSG_DISCONNECT))
	; wait for the connection to be closed => provides the main process with the opportunity to read all TCP/STD data
	Local $bTimeOut = True
	For $i=0 to 1000 ; timeout after 10 seconds
		TCPSend($iSocket, Binary($__IPC_MSG_ACK))
		If @error Then
			$bTimeOut = False
			ExitLoop
		EndIf
		Sleep(10)
	Next
	If $bTimeOut Then __IPC_Log($__IPC_LOG_ERROR, "Connection to main process did not close properly.") ; disconnect took too long
	TCPCloseSocket($__IPC__Data.mClient.iSocket)
	; reset sub process
	$__IPC__Data.mClient.iPullRate = $__IPC_SubPullRate
	$__IPC__Data.mClient.iSocket = Default
	; remove tcp message processing
	If $__IPC__Data.mClient.sCallback<>Default Then
		If $__IPC__Data.mClient.iPullRate>0 Then AdlibUnRegister("__IPC_SubProcessing")
		$__IPC__Data.mClient.sCallback = Default
	EndIf
	; call the exit/close callback of the sub process, if desired
	If $__IPC__Data.mClient.sExitCallback<>Default Then
		Call($__IPC__Data.mClient.sExitCallback)
		$__IPC__Data.mClient.sExitCallback = Default
	EndIf
	__IPC_Log($__IPC_LOG_INFO, "Client connection closed.")
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ServerStart
; Description ...: Start the main process server
; Syntax ........: __IPC__ServerStart()
; Parameters ....:
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ServerStart()
	If $__IPC__Data.mServer.iListen<>Default Then Return SetError(2, 0, False) ; server already open
	; look for the next open port to listen to, starting with iMainStartPort
	Local $iPort = $__IPC__Data.mServer.iMainStartPort
	Local $iListen = -1
	While True
		$iListen = TCPListen("127.0.0.1", $iPort)
		If @error=1 Then Return SetError(2, 0, False)
		If $iListen>=0 Then ExitLoop
		$iPort+=1
		If $iPort>65535 Then $iPort = 1024
		If $iPort=$__IPC__Data.mServer.iMainStartPort-1 Then Return SetError(3, 0, False)
	WEnd
	If $iListen<=0 Then Return False ; failed to open tcp connection
	; save connection data and start listening for tcp/std data
	If $__IPC__Data.mServer.iMainPullRate>0 Then AdlibRegister("__IPC_MainProcessing", $__IPC__Data.mServer.iMainPullRate)
	$__IPC__Data.mServer.iListen = $iListen
	$__IPC__Data.mServer.iPort = $iPort
	__IPC_Log($__IPC_LOG_INFO, "IPC Server listening at 127.0.0.1:"&$iPort)
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ServerIsRunning
; Description ...: Check if the main process server is running
; Syntax ........: __IPC__ServerIsRunning()
; Parameters ....:
; Return values .: True if the server is running.
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ServerIsRunning()
	Return $__IPC__Data.mServer.iListen<>Default
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ServerStop
; Description ...: Stop the main process server.
; Syntax ........: __IPC__ServerStop()
; Parameters ....:
; Return values .: True on success.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  2 - Server is not running
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ServerStop()
	If $__IPC__Data.mServer.iListen=Default Then Return SetError(2, 0, False) ; server not running
	TCPCloseSocket($__IPC__Data.mServer.iListen)
	$__IPC__Data.mServer.iListen = Default
	If $__IPC__Data.mServer.iMainPullRate>0 Then AdlibUnRegister("__IPC_MainProcessing") ; remove automatic message processing
	__IPC_Log($__IPC_LOG_INFO, "IPC Server stopped")
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ServerProcessStdOut
; Description ...: Iterate all sub processes and check for stdout/stderr data and if they are still connected.
; Syntax ........: __IPC__ServerProcessStdOut()
; Parameters ....:
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ServerProcessStdOut()
	__IPC_Log($__IPC_LOG_TRACE, "Run __IPC__ServerProcessStdOut")
	Local $arProcesses = MapKeys($__IPC__Data.mServer.mProcesses)
	For $i=0 to UBound($arProcesses)-1 Step 1
		__IPC_Log($__IPC_LOG_TRACE, "Check std of $iProcess: "&$arProcesses[$i])
		Local $mProcess = $__IPC__Data.mServer.mProcesses[$arProcesses[$i]]
		If $mProcess.bStdOut Then __IPC__ServerProcessLogStd($arProcesses[$i])
		If $mProcess.bStdErr Then __IPC__ServerProcessLogStd($arProcesses[$i], True)
		If Not ProcessExists($mProcess.iPID) Then __IPC__ServerProcessRemove($arProcesses[$i])
	Next
	__IPC_Log($__IPC_LOG_TRACE, "Done __IPC__ServerProcessStdOut")
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ServerProcessStdOut
; Description ...: Process the stdout/stderr data from a sub processes.
; Syntax ........: __IPC__ServerProcessStdOut($iProcess[, $bErr = False])
; Parameters ....: $iProcess         - the process id
;                  $bErr             - [optional] Default: False. True if stderr should be read. False for stdout.
; Return values .: False on error.
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ServerProcessLogStd($iProcess, $bErr = False)
	If Not MapExists($__IPC__Data.mServer.mProcesses, $iProcess) Then Return SetError(1, 1, True) ; process does not exists
	; read process stdout/stderr
	Local $mProcess = $__IPC__Data.mServer.mProcesses[$iProcess]
	Local $sData = $bErr?StderrRead($mProcess.iPID):StdoutRead($mProcess.iPID)
	If @error Then Return SetError(2, 0, True)
	If @extended<=0 Then Return True
	; handle output by printing to console
	Local $arData = StringSplit($sData, @CRLF, 3)
	For $i=0 To UBound($arData)-1
		If $i=UBound($arData)-1 And StringLen($arData[$i])=0 Then ContinueLoop
		If $bErr Then
			__IPC_Log($__IPC_LOG_ERROR, $arData[$i], Default, Default, $mProcess.hProcess)
		Else
			__IPC_Log($__IPC_LOG_INFO, $arData[$i], Default, Default, $mProcess.hProcess)
		EndIf
	Next
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ServerAccept
; Description ...: Accept new connections of sub processes to the server
; Syntax ........: __IPC__ServerAccept()
; Parameters ....:
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ServerAccept()
	__IPC_Log($__IPC_LOG_TRACE, "Run __IPC__ServerAccept")
	While True
		Local $iTimeout = Opt("TCPTimeout", 0) ; disable 100ms timeout for TCPAccept
		Local $iSocket = TCPAccept($__IPC__Data.mServer.iListen)
		Opt("TCPTimeout", $iTimeout) ; reset to old timeout for users
		If $iSocket=-1 Then ExitLoop
		__IPC_Log($__IPC_LOG_INFO, "Client connect: "&$iSocket)
		__IPC__AddSocket($iSocket, $__IPC_CONN_TO_SUB)
		$__IPC__Data["mServer"]["iOpenProcesses"] -= 1
	WEnd
	__IPC_Log($__IPC_LOG_TRACE, "Done __IPC__ServerAccept")
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__AddSocket
; Description ...: Add a new socket to the open connections
; Syntax ........: __IPC__AddSocket($iSocket, $iType)
; Parameters ....: $iSocket       - the socket
;                  $iType         - the type ($__IPC_CONN_TO_SUB or $__IPC_CONN_TO_MAIN)
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__AddSocket($iSocket, $iType)
	Local $mConn[]
	$mConn.iType = $iType ; is the socket from main->sub process or from sub->main process
	$mConn.iSocket = $iSocket
	$mConn.iLastCommand = Default ; contains the last command
	$mConn.dataBuffer = Default ; collects received data
	$mConn.iDataBufferSize = Default ; size of $mConn.dataBuffer
	$mConn.hProcess = Default ; only set for $__IPC_CONN_TO_SUB
	$mConn.iCommandBytes = Default ; amount of bytes to receive before executing command $mConn.iLastCommand (only for commands with custom data with unknown size)
	$__IPC__Data["mConnects"][$iSocket] = $mConn
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ProcessMessages
; Description ...: Read all data from all open tcp connections. After that they all get processed.
; Syntax ........: __IPC__ProcessMessages()
; Parameters ....:
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ProcessMessages()
	__IPC_Log($__IPC_LOG_TRACE, "Run __IPC__ProcessMessages")
	Local $arSockets = MapKeys($__IPC__Data.mConnects)
	__IPC_Log($__IPC_LOG_TRACE, "Run __IPC__ProcessMessages with "&UBound($arSockets)&" sockets")
	; read new data from socket
	Local $iLoopCount = 0
	For $i=0 to UBound($arSockets)-1 Step 1
		Local $iSocket = $arSockets[$i]
		Local $iTimeout = Opt("TCPTimeout", 0) ; disable 100ms timeout for TCPAccept
		Local $bData = TCPRecv($iSocket, $__IPC_MaxByteRecv, 1)
		Opt("TCPTimeout", $iTimeout) ; reset to old timeout for users
		If @error Then ; connection lost
			__IPC_Log($__IPC_LOG_ERROR, "Connection lost: "&$iSocket)
			__IPC__SocketDisconnect($iSocket, True)
			ContinueLoop
		ElseIf BinaryLen($bData)=0 Then ; no data received
			__IPC_Log($__IPC_LOG_TRACE, "Nothing received: "&$iSocket)
			ContinueLoop
		EndIf
		; put received data into the buffer
		Local $iBinLen = BinaryLen($bData)
		__IPC_Log($__IPC_LOG_TRACE, "Received: "&$iSocket&":"&$iBinLen&" >> "&$bData)
		If $__IPC__Data["mConnects"][$iSocket]["dataBuffer"]=Default Then
			$__IPC__Data["mConnects"][$iSocket]["dataBuffer"] = $bData
			$__IPC__Data["mConnects"][$iSocket]["iDataBufferSize"] = $iBinLen
		Else
			$__IPC__Data["mConnects"][$iSocket]["dataBuffer"] &= $bData
			$__IPC__Data["mConnects"][$iSocket]["iDataBufferSize"] += $iBinLen
		EndIf
		If $iBinLen>=$__IPC_MaxByteRecv And ($__IPC__Data.iMaxReceiveCount=Default Or $iLoopCount<$__IPC__Data.iMaxReceiveCount) Then
			__IPC_Log($__IPC_LOG_TRACE, "Receive more data: "&$iSocket&" >> "&$iBinLen&" >> "&$iLoopCount)
			$i-=1 ; more data may be available, maybe limit this to 10 tries
			$iLoopCount += 1
		Else
			$iLoopCount = 0
		EndIf
	Next
	__IPC_Log($__IPC_LOG_TRACE, "Do process __IPC__ProcessMessages")
	; process data from socket
	Local $arSockets = MapKeys($__IPC__Data.mConnects)
	__IPC_Log($__IPC_LOG_TRACE, "Do process __IPC__ProcessMessages with "&UBound($arSockets)&" sockets")
	For $i=0 to UBound($arSockets)-1 Step 1
		__IPC__ProcessMessagesAtSocket($arSockets[$i])
	Next
	__IPC_Log($__IPC_LOG_TRACE, "Done __IPC__ProcessMessages")
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ProcessMessagesAtSocket
; Description ...: Process all incoming tcp data.
; Syntax ........: __IPC__ProcessMessagesAtSocket($iSocket)
; Parameters ....: $iSocket        - the socket
; Return values .:
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ProcessMessagesAtSocket($iSocket)
	While MapExists($__IPC__Data.mConnects, $iSocket) And $__IPC__Data.mConnects[$iSocket]["iDataBufferSize"]>0
		If $__IPC__Data.mConnects[$iSocket]["iLastCommand"]=Default Then
			Local $bData = __IPC__SocketReadBytes($iSocket, 4) ; int32
			If @error Then ExitLoop ; wait for more data to be received
			If Not @error Then $__IPC__Data["mConnects"][$iSocket]["iLastCommand"] = $bData
		EndIf
		Local $iCmd = $__IPC__Data.mConnects[$iSocket]["iLastCommand"]
		Switch $iCmd
			Case $__IPC_MSG_DATA, $__IPC_MSG_CMD, $__IPC_MSG_CMD_DATA
				; check if the amount of message data bytes was already read, otherwise read them first
				If $__IPC__Data["mConnects"][$iSocket]["iCommandBytes"] = Default Then
					Local $bData = __IPC__SocketReadBytes($iSocket, 8) ; int64
					If @error Then ExitLoop ; wait for more data to be received
					If Not @error Then $__IPC__Data["mConnects"][$iSocket]["iCommandBytes"] = Int($bData)
				EndIf
				; read the specified amount of message data bytes
				Local $dData = __IPC__SocketReadBytes($iSocket, $__IPC__Data.mConnects[$iSocket]["iCommandBytes"])
				If @error Then ExitLoop ; wait for more data to be received
				; reset the command and message bytes
				$__IPC__Data["mConnects"][$iSocket]["iLastCommand"] = Default
				$__IPC__Data["mConnects"][$iSocket]["iCommandBytes"] = Default
				; Handle the first byte specifying the command if applicable
				Local $iDataCommand = Default
				If $iCmd = $__IPC_MSG_CMD Or $iCmd = $__IPC_MSG_CMD_DATA Then $iDataCommand = __IPC__BinaryConsumeVar($dData)
				Local $iDataCount = __IPC__BinaryConsumeVar($dData)
				Local $arData[$iDataCount]
				Local $sLogData = "arData["&$iDataCount&"]"
				If $iDataCount>0 Then
					$sLogData = "arData["
					For $i=0 To $iDataCount-1
						If $i<>0 Then $sLogData &= ", "
						$arData[$i] = __IPC__BinaryConsumeVar($dData)
						If IsFunc("_ToStringC") Then
							$sLogData &= Execute('Call("_ToStringC", $arData[$i], False, "", "", " ")')
						Else
							$sLogData &= String($arData[$i])
						EndIf
					Next
					$sLogData &= "]"
				EndIf
				If UBound($arData)<=0 Then $arData = Default
				; get the callback (if present) and the hProcess (if on main process and the message came from a sub process)
				Local $sCallback = Default
				Local $hProcess = Default
				If $__IPC__Data.mConnects[$iSocket].iType = $__IPC_CONN_TO_MAIN Then
					$sCallback = $__IPC__Data.mClient.sCallback
				ElseIf $__IPC__Data.mConnects[$iSocket].iType = $__IPC_CONN_TO_SUB Then
					$hProcess = $__IPC__Data.mConnects[$iSocket].hProcess
					Local $iProcess = __IPC__ProcessHandleToId($hProcess)
					If @error Then ContinueLoop ; should not happen, as long as the connection is there
					$sCallback = $__IPC__Data.mServer.mProcesses[$iProcess].sCallback
				EndIf
				If $sCallback<>Default Then
					If $iDataCommand<>Default And $arData<>Default Then
						__IPC_Log($__IPC_LOG_DEBUG, "__IPC_MSG_CMD_DATA Received command: "&$iSocket&" >> "&$iDataCommand&" >> "&$sLogData)
						If $hProcess<>Default Then
							Call($sCallback, $hProcess, $iDataCommand, $arData)
							If @error Then __IPC_Log($__IPC_LOG_ERROR, "Error calling: "&$sCallback&" with 3 parameters")
						Else
							Call($sCallback, $iDataCommand, $arData)
							If @error Then __IPC_Log($__IPC_LOG_ERROR, "Error calling: "&$sCallback&" with 2 parameters")
						EndIf
					ElseIf $iDataCommand<>Default Then
						__IPC_Log($__IPC_LOG_DEBUG, "__IPC_MSG_CMD Received command: "&$iSocket&" >> "&$iDataCommand)
						If $hProcess<>Default Then
							Call($sCallback, $hProcess, $iDataCommand, Default)
							If @error Then __IPC_Log($__IPC_LOG_ERROR, "Error calling: "&$sCallback&" with 2 parameters")
						Else
							Call($sCallback, $iDataCommand, Default)
							If @error Then __IPC_Log($__IPC_LOG_ERROR, "Error calling: "&$sCallback&" with 1 parameter")
						EndIf
					ElseIf $arData<>Default Then
						__IPC_Log($__IPC_LOG_DEBUG, "__IPC_MSG_DATA Received data: "&$iSocket&" >> "&$sLogData)
						If $hProcess<>Default Then
							Call($sCallback, $hProcess, Default, $arData)
							If @error Then __IPC_Log($__IPC_LOG_ERROR, "Error calling: "&$sCallback&" with 2 parameters")
						Else
							Call($sCallback, Default, $arData)
							If @error Then __IPC_Log($__IPC_LOG_ERROR, "Error calling: "&$sCallback&" with 1 parameters")
						EndIf
					EndIf
				EndIf
			Case $__IPC_MSG_CONNECT ; server side only
				; read the provided hProcess to assign to the socket
				Local $hProcess = Int(__IPC__SocketReadBytes($iSocket, 4))
				If @error Then ExitLoop ; wait for more data to be received
				__IPC_Log($__IPC_LOG_INFO, "Subprocess connected: "&$iSocket&" Process: "&$hProcess)
				$__IPC__Data["mConnects"][$iSocket]["iLastCommand"] = Default
				Local $iProcess = __IPC__ProcessHandleToId($hProcess)
				If @error Then ContinueLoop ; tcp sent a not existing hProcess => protocol error, skip the command
				$__IPC__Data["mConnects"][$iSocket]["hProcess"] = $hProcess
				$__IPC__Data["mServer"]["mProcesses"][$iProcess]["iSocket"] = $iSocket
				$__IPC__Data["mServer"]["mProcesses"][$iProcess]["bWaitForSocket"] = False
			Case $__IPC_MSG_DISCONNECT
				__IPC_Log($__IPC_LOG_INFO, "Received MSG_DISCONNECT: "&$iSocket)
				$__IPC__Data["mConnects"][$iSocket]["iLastCommand"] = Default
				__IPC__SocketDisconnect($iSocket)
				ExitLoop
			Case $__IPC_MSG_ACK
				__IPC_Log($__IPC_LOG_TRACE, "Received $__IPC_MSG_ACK: "&$iSocket)
				$__IPC__Data["mConnects"][$iSocket]["iLastCommand"] = Default
			Case Default
				__IPC_Log($__IPC_LOG_TRACE, "Received unknown command: "&$iCmd&" at socket "&$iSocket)
				$__IPC__Data["mConnects"][$iSocket]["iLastCommand"] = Default
		EndSwitch
	WEnd
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__SocketReadBytes
; Description ...: Read a specified amount of bytes from a socket. If $iBytes are in the buffer, they are consumed and returned.
;                  Otherwise an error is returned.
; Syntax ........: __IPC__SocketReadBytes($iSocket, $iBytes)
; Parameters ....: $iSocket        - the socket
;                  $iBytes         - the number of bytes
; Return values .: the requested binary data or 0 on failure
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $iSocket, 2 - $iBytes)
;                  2 - $iBytes are not in the data buffer
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__SocketReadBytes($iSocket, $iBytes)
	; check parameters
	If Not IsInt($iSocket) Or Not MapExists($__IPC__Data.mConnects, $iSocket) Then Return SetError(1, 1, 0)
	If Not IsInt($iBytes) Or $iBytes<=0 Then Return SetError(1, 2, 0)
	 ; check amount of bytes available
	If $__IPC__Data["mConnects"][$iSocket]["iDataBufferSize"] < $iBytes Then Return SetError(2, 0, 0)
	; consume specified amount of bytes and return them
	Local $bData = BinaryMid($__IPC__Data["mConnects"][$iSocket]["dataBuffer"], 1, $iBytes)
	$__IPC__Data["mConnects"][$iSocket]["dataBuffer"] = BinaryMid($__IPC__Data["mConnects"][$iSocket]["dataBuffer"], $iBytes+1)
	$__IPC__Data["mConnects"][$iSocket]["iDataBufferSize"] = BinaryLen($__IPC__Data.mConnects[$iSocket]["dataBuffer"])
	Return $bData
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__SocketDisconnect
; Description ...: Disconnect a socket from the open connections
; Syntax ........: __IPC__SocketDisconnect($iSocket[, $bError = False])
; Parameters ....: $iSocket        - the socket
;                  $bError         - [optional] Default: False. True, if the cause was an error
; Return values .: True on sucess
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $iSocket)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__SocketDisconnect($iSocket, $bError = False)
	; make sure socket exists
	If Not IsInt($iSocket) Or Not MapExists($__IPC__Data.mConnects, $iSocket) Then Return SetError(1, 1, False)
	; check if it is the sub process receiving data from the main process
	If $__IPC__Data.mConnects[$iSocket].iType = $__IPC_CONN_TO_MAIN Then
		__IPC_Log($__IPC_LOG_DEBUG, "Disconnect from main: "&$iSocket)
		MapRemove($__IPC__Data.mConnects, $iSocket)
		__IPC__SubDisconnect($bError)
	; check if it is the main process receiving data from the sub process
	ElseIf $__IPC__Data.mConnects[$iSocket].iType = $__IPC_CONN_TO_SUB Then
		Local $hProcess = $__IPC__Data["mConnects"][$iSocket]["hProcess"]
		If $hProcess=Default Then ; socket never received $__IPC_MSG_CONNECT with $hProcess to identify the socket
			__IPC_Log($__IPC_LOG_DEBUG, "Disconnect unidentified sub process with socket "&$iSocket)
			$__IPC__Data["mServer"]["iOpenProcesses"] += 1
		Else ; remove the socket for the process and send the terminate signal to the sub process
			__IPC_Log($__IPC_LOG_DEBUG, "Disconnect sub process: "&$hProcess&" with socket "&$iSocket)
			TCPSend($iSocket, Binary($__IPC_MSG_DISCONNECT))
			__IPC__ProcessMessagesAtSocket($iSocket) ; process last data
			Local $iProcess = __IPC__ProcessHandleToId($hProcess)
			If Not @error Then $__IPC__Data["mServer"]["mProcesses"][$iProcess]["iSocket"] = Default
		EndIf
		MapRemove($__IPC__Data.mConnects, $iSocket)
		TCPCloseSocket($iSocket)
		If $hProcess=Default Then
			__IPC_Log($__IPC_LOG_INFO, "Disconnected unidentified sub process with socket "&$iSocket)
		Else
			__IPC_Log($__IPC_LOG_INFO, "Disconnected sub process: "&$hProcess&" with socket "&$iSocket)
		EndIf
		If UBound(MapKeys($__IPC__Data.mServer.mProcesses))=0 Then __IPC__ServerStop()
	EndIf
	Return True
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ServerProcessRemove
; Description ...: Remove a process from the sub process list
; Syntax ........: __IPC__ServerProcessRemove($iProcess[, $bError = False])
; Parameters ....: $iProcess       - the process id
;                  $bError         - [optional] Default: False. True, if the cause was an error
; Return values .: True on sucess
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $iProcess)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ServerProcessRemove($iProcess, $bError = False)
	; make sure process exists
	If Not IsInt($iProcess) Or $iProcess < 0 Or Not MapExists($__IPC__Data.mServer.mProcesses, $iProcess) Then Return SetError(1, 0, False)
	Local $mProcess = $__IPC__Data.mServer.mProcesses[$iProcess]
	__IPC__ServerProcessLogStd($iProcess) ; handle output a last time
	; disconnect socket if still connected => send terminate signal to sub processes
	If $mProcess.iSocket<>Default Then __IPC__SocketDisconnect($mProcess.iSocket, $bError)
	If $__IPC__Data["mServer"]["mProcesses"][$iProcess]["bWaitForSocket"] Then $__IPC__Data["mServer"]["iOpenProcesses"] -= 1
	If $__IPC__Data.mServer.mProcesses[$iProcess].sExitCallback<>Default Then Call($__IPC__Data.mServer.mProcesses[$iProcess].sExitCallback, $mProcess.hProcess)
	MapRemove($__IPC__Data.mServer.mProcesses, $iProcess)
	__IPC_Log($__IPC_LOG_INFO, "Process removed: "&$mProcess.hProcess)
	; todo if $bError Then ; todo callback for undesired error
	; call $mServer.mProcesses[$iProcess]["sCallback"]
	If UBound(MapKeys($__IPC__Data.mServer.mProcesses))=0 Then __IPC__ServerStop()
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ProcessIdToHandle
; Description ...: Get the process handle for a process id
; Syntax ........: __IPC__ProcessIdToHandle($iId)
; Parameters ....: $iId       - the process id
; Return values .: $hProcess
; Author ........: Kanashius
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ProcessIdToHandle($iId)
	Return $iId+1
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __IPC__ProcessHandleToId
; Description ...: Get the process id for a process handle
; Syntax ........: __IPC__ProcessHandleToId($hProcess)
; Parameters ....: $hProcess       - the process handle
; Return values .: $iProcess on success
; Author ........: Kanashius
; Modified ......:
; Remarks .......: Errors:
;                  1 - Parameter invalid (@extended: 1 - $hProcess)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __IPC__ProcessHandleToId($hProcess)
	If Not IsInt($hProcess) Or $hProcess-1<0 Or Not MapExists($__IPC__Data.mServer.mProcesses, $hProcess-1) Then Return SetError(1, 1, 0)
	Return $hProcess-1
EndFunc
