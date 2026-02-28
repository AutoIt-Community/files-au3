#include-once

#include <WinAPIShellEx.au3>
#include <APIErrorsConstants.au3>

; #FUNCTIONS# ===================================================================================================================
; _IFileOperationNewItem
; _IFileOperationRenameItem
; _IFileOperationRenameItems (not working properly)
; _IFileOperationDeleteItem
; _IFileOperationDeleteItems
; _IFileOperationCopyItem
; _IFileOperationCopyItems
; _IFileOperationMoveItem
; _IFileOperationMoveItems
; ===============================================================================================================================

Global Const $FOFX_ADDUNDORECORD = 0x20000000
Global Const $FOFX_NOSKIPJUNCTIONS = 0x00010000
Global Const $FOFX_PREFERHARDLINK = 0x00020000
Global Const $FOFX_SHOWELEVATIONPROMPT = 0x00040000
Global Const $FOFX_EARLYFAILURE = 0x00100000
Global Const $FOFX_PRESERVEFILEEXTENSIONS = 0x00200000
Global Const $FOFX_KEEPNEWERFILE = 0x00400000
Global Const $FOFX_NOCOPYHOOKS = 0x00800000
Global Const $FOFX_NOMINIMIZEBOX = 0x01000000
Global Const $FOFX_MOVEACLSACROSSVOLUMES = 0x02000000
Global Const $FOFX_DONTDISPLAYSOURCEPATH = 0x04000000
Global Const $FOFX_DONTDISPLAYDESTPATH = 0x08000000
Global Const $FOFX_RECYCLEONDELETE = 0x00080000
Global Const $FOFX_REQUIREELEVATION = 0x10000000
Global Const $FOFX_COPYASDOWNLOAD = 0x40000000
Global Const $FOFX_DONTDISPLAYLOCATIONS = 0x80000000

Global Const $IID_IShellItem = "{43826d1e-e718-42ee-bc55-a1e261c37bfe}"
Global Const $tIIDIShellItem = _WinAPI_GUIDFromString($IID_IShellItem)
Global Const $dtag_IShellItem = _
        "BindToHandler hresult(ptr;clsid;clsid;ptr*);" & _
        "GetParent hresult(ptr*);" & _
        "GetDisplayName hresult(int;ptr*);" & _
        "GetAttributes hresult(int;int*);" & _
        "Compare hresult(ptr;int;int*);"

Global Const $IID_IShellItemArray = "{b63ea76d-1f85-456f-a19c-48159efa858b}"
Global Const $dtagIShellItemArray = "BindToHandler hresult();GetPropertyStore hresult();" & _
        "GetPropertyDescriptionList hresult();GetAttributes hresult();GetCount hresult(dword*);" & _
        "GetItemAt hresult();EnumItems hresult();"

Global Const $BHID_EnumItems = "{94F60519-2850-4924-AA5A-D15E84868039}"
Global Const $IID_IEnumShellItems = "{70629033-e363-4a28-a567-0db78006e6d7}"
Global Const $dtagIEnumShellItems = "Next hresult(ulong;ptr*;ulong*);Skip hresult();Reset hresult();Clone hresult();"

Global Const $CLSID_IFileOperation = "{3AD05575-8857-4850-9277-11B85BDB8E09}"
Global Const $IID_IFileOperation = "{947AAB5F-0A5C-4C13-B4D6-4BF7836FC9F8}"
Global Const $tagIFileOperation = _
    "Advise hresult(ptr;dword*);" & _
    "Unadvise hresult(dword);" & _
    "SetOperationFlags hresult(dword);" & _
    "SetProgressMessage hresult(wstr);" & _
    "SetProgressDialog hresult(ptr);" & _
    "SetProperties hresult(ptr);" & _
    "SetOwnerWindow hresult(hwnd);" & _
    "ApplyPropertiesToItem hresult(ptr);" & _
    "ApplyPropertiesToItems hresult(ptr);" & _
    "RenameItem hresult(ptr;wstr;ptr);" & _
    "RenameItems hresult(ptr;wstr);" & _
    "MoveItem hresult(ptr;ptr;wstr;ptr);" & _
    "MoveItems hresult(ptr;ptr);" & _
    "CopyItem hresult(ptr;ptr;wstr;ptr);" & _
    "CopyItems hresult(ptr;ptr);" & _
    "DeleteItem hresult(ptr;ptr);" & _
    "DeleteItems hresult(ptr);" & _
    "NewItem hresult(ptr;dword;wstr;wstr;ptr);" & _
    "PerformOperations hresult();" & _
    "GetAnyOperationsAborted hresult(ptr*);"

Global Const $PTR_LEN = @AutoItX64 ? 8 : 4
Global Const $sIID_IDataObject = "{0000010e-0000-0000-C000-000000000046}"
Global Const $sIID_IShellFolder = "{000214E6-0000-0000-C000-000000000046}"

Global Const $tagIShellFolder = _
		"ParseDisplayName hresult(hwnd; ptr; wstr; ulong*; ptr*; ulong*);" & _
		"EnumObjects hresult(hwnd; ulong; ptr*);" & _
		"BindToObject hresult(ptr; ptr*; struct*; ptr*);" & _
		"BindToStorage hresult(ptr; ptr*; struct*; ptr*);" & _
		"CompareIDs hresult(lparam; ptr; ptr);" & _
		"CreateViewObject hresult(hwnd; struct*; ptr*);" & _
		"GetAttributesOf hresult(uint; ptr; ulong*);" & _
		"GetUIObjectOf hresult(hwnd; uint; ptr; struct*; uint*; ptr*);" & _
		"GetDisplayNameOf hresult(ptr; ulong; ptr*);" & _
		"SetNameOf hresult(hwnd; ptr; wstr; ulong; ptr*);"

Global $hShell32 = DllOpen('shell32.dll')
OnAutoItExitRegister("_Cleanup")

Func _IFileOperationNewItem($sPath, $sName, $bFolder = False)
    Local $Attrib = 0
    If $bFolder Then $Attrib = $FILE_ATTRIBUTE_DIRECTORY
    If Not FileExists($sPath) Then DirCreate($sPath)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    Local $pIShellItem = SHCreateItemFromParsingName($sPath, 0, DllStructGetPtr($tIIDIShellItem))
    $oIFileOperation.NewItem($pIShellItem, $Attrib, $sName, "", 0)
    $oIFileOperation.PerformOperations()
EndFunc   ;==>_IFileOperationNewItem

Func _IFileOperationRenameItem($sPath, $sNewName)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    Local $pIShellItem = SHCreateItemFromParsingName($sPath, 0, DllStructGetPtr($tIIDIShellItem))
    $oIFileOperation.RenameItem($pIShellItem, $sNewName, 0)
    $oIFileOperation.PerformOperations()
EndFunc   ;==>_IFileOperationRenameItem

Func _IFileOperationRenameItems($pDataObj, $sNewName)
    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $tIIDIShellItemArray = CLSIDFromString($IID_IShellItemArray)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)
    $oIFileOperation.RenameItems($pDataObj, $sNewName)
    Return $oIFileOperation.PerformOperations() = 0
EndFunc   ;==>_IFileOperationRenameItems

Func _IFileOperationDeleteItem($sPathFrom, $bPermanent = False, $iFlags = 0)
    $iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
    If $bPermanent Then $iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_NOCOPYHOOKS)
    If Not FileExists($sPathFrom) Then Return SetError(1, 0, False)
    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)
    Local $pIShellItemFrom = SHCreateItemFromParsingName($sPathFrom, 0, DllStructGetPtr($tIIDIShellItem))
    If Not $pIShellItemFrom Then Return SetError(3, 0, False)
    $oIFileOperation.SetOperationFlags($iFlags)
    $oIFileOperation.DeleteItem($pIShellItemFrom, 0)
    Return $oIFileOperation.PerformOperations() = 0
EndFunc   ;==>_IFileOperationDeleteItem

Func _IFileOperationDeleteItems($pDataObj, $bPermanent = False, $iFlags = 0)
    $iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
    If $bPermanent Then $iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_NOCOPYHOOKS)
    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $tIIDIShellItemArray = CLSIDFromString($IID_IShellItemArray)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)
    $oIFileOperation.SetOperationFlags($iFlags)
    $oIFileOperation.DeleteItems($pDataObj)
    Return $oIFileOperation.PerformOperations() = 0
EndFunc   ;==>_IFileOperationDeleteItems

Func _IFileOperationCopyItem($sPathFrom, $sPathTo, $iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS))
    If Not FileExists($sPathFrom) Then Return SetError(1, 0, False)
    If Not FileExists($sPathTo) Then DirCreate($sPathTo)
    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)
    Local $pIShellItemFrom = SHCreateItemFromParsingName($sPathFrom, 0, DllStructGetPtr($tIIDIShellItem))
    Local $pIShellItemTo = SHCreateItemFromParsingName($sPathTo, 0, DllStructGetPtr($tIIDIShellItem))
    If Not $pIShellItemFrom Or Not $pIShellItemTo Then Return SetError(3, 0, False)
    $oIFileOperation.SetOperationFlags($iFlags)
    $oIFileOperation.CopyItem($pIShellItemFrom, $pIShellItemTo, Null, 0)
    Return $oIFileOperation.PerformOperations() = 0
EndFunc   ;==>_IFileOperationCopyItem

Func _IFileOperationCopyItems($pDataObj, $sPathTo, $iFlags = 0)
    If Not FileExists($sPathTo) Then DirCreate($sPathTo)
    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $tIIDIShellItemArray = CLSIDFromString($IID_IShellItemArray)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)
    Local $pIShellItemTo = SHCreateItemFromParsingName($sPathTo, 0, DllStructGetPtr($tIIDIShellItem))
    If Not $pIShellItemTo Then Return SetError(3, 0, False)
    $oIFileOperation.SetOperationFlags($iFlags)
    $oIFileOperation.CopyItems($pDataObj, $pIShellItemTo)
    Return $oIFileOperation.PerformOperations() = 0
EndFunc   ;==>_IFileOperationCopyItems

Func _IFileOperationMoveItem($sPathFrom, $sPathTo, $iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS))
    If Not FileExists($sPathFrom) Then Return SetError(1, 0, False)
    If Not FileExists($sPathTo) Then DirCreate($sPathTo)
    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)
    Local $pIShellItemFrom = SHCreateItemFromParsingName($sPathFrom, 0, DllStructGetPtr($tIIDIShellItem))
    Local $pIShellItemTo = SHCreateItemFromParsingName($sPathTo, 0, DllStructGetPtr($tIIDIShellItem))
    If Not $pIShellItemFrom Or Not $pIShellItemTo Then Return SetError(3, 0, False)
    $oIFileOperation.SetOperationFlags($iFlags)
    $oIFileOperation.MoveItem($pIShellItemFrom, $pIShellItemTo, Null, 0)
    Return $oIFileOperation.PerformOperations() = 0
EndFunc   ;==>_IFileOperationMoveItem

Func _IFileOperationMoveItems($pDataObj, $sPathTo, $iFlags = 0)
    If Not FileExists($sPathTo) Then DirCreate($sPathTo)
    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $tIIDIShellItemArray = CLSIDFromString($IID_IShellItemArray)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)
    Local $pIShellItemTo = SHCreateItemFromParsingName($sPathTo, 0, DllStructGetPtr($tIIDIShellItem))
    If Not $pIShellItemTo Then Return SetError(3, 0, False)
    $oIFileOperation.SetOperationFlags($iFlags)
    $oIFileOperation.MoveItems($pDataObj, $pIShellItemTo)
    Return $oIFileOperation.PerformOperations() = 0
EndFunc   ;==>_IFileOperationMoveItems

; Original function name, still needed for Files Au3
; TO DO: would like to switch to using _IFileOperationMoveItems and _IFileOperationCopyItems
Func _IFileOperationFile($pDataObj, $sPathTo, $sAction, $iFlags = 0)
    If Not FileExists($sPathTo) Then DirCreate($sPathTo)
    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $tIIDIShellItemArray = CLSIDFromString($IID_IShellItemArray)
    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $tagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)
    Local $pIShellItemTo = SHCreateItemFromParsingName($sPathTo, 0, DllStructGetPtr($tIIDIShellItem))
    If Not $pIShellItemTo Then Return SetError(3, 0, False)
    $oIFileOperation.SetOperationFlags($iFlags)
    Switch $sAction
        Case "CopyItems"
            $oIFileOperation.CopyItems($pDataObj, $pIShellItemTo)
        Case "MoveItems"
            $oIFileOperation.MoveItems($pDataObj, $pIShellItemTo)
    EndSwitch
    Return $oIFileOperation.PerformOperations() = 0
EndFunc   ;==>_IFileOperationFile

Func SHCreateItemFromParsingName($sPath, $pPbc, $pRIID)
  Local $aRes = DllCall($hShell32, "long", "SHCreateItemFromParsingName", "wstr", $sPath, "ptr", $pPbc, "ptr", $pRIID, "ptr*", 0)
  If @error Or $aRes[0] Then Return SetError(1)
  Return $aRes[4]
EndFunc   ;==>SHCreateItemFromParsingName

Func CLSIDFromString($sString)
    Local $tCLSID = DllStructCreate("dword;word;word;byte[8]")
    Local $aRet = DllCall("Ole32.dll", "long", "CLSIDFromString", "wstr", $sString, "ptr", DllStructGetPtr($tCLSID))
    If @error Then Return SetError(1, 0, @error)
    If $aRet[0] <> 0 Then Return SetError(2, $aRet[0], 0)
    Return $tCLSID
EndFunc   ;==>CLSIDFromString

Func _Cleanup()
    DllClose($hShell32)
EndFunc   ;==>_Cleanup

Func _Release($pThis)
	If (Not $pThis) Or (Not IsPtr($pThis)) Then Return SetError($ERROR_INVALID_PARAMETER)
	Local $pVTab = DllStructGetData(DllStructCreate("ptr", $pThis), 1)
	Local $pFunc = DllStructGetData(DllStructCreate("ptr", $pVTab + 2 * $PTR_LEN), 1)
	Local $aCall = DllCallAddress("uint", $pFunc, "ptr", $pThis)
	Return $aCall[0]
EndFunc   ;==>_Release

Func GetDataObject(ByRef $sPath) ; code by jugador
    Local $iCount = UBound($sPath)
    If $iCount = 0 Then Return 0
    Local $sParentPath = StringLeft($sPath[0], StringInStr($sPath[0], "\", 0, -1) - 1)
    Local $pParentPidl = _WinAPI_ShellILCreateFromPath($sParentPath)
    Local $tPidls = DllStructCreate("ptr[" & $iCount & "]")
    Local $pFullPidl, $pRelativePidl, $last_SHITEMID
    For $i = 0 To $iCount - 1
        $pFullPidl = _WinAPI_ShellILCreateFromPath($sPath[$i])
        $last_SHITEMID = DllCall($hShell32, "ptr", "ILFindLastID", "ptr", $pFullPidl)[0]
        $pRelativePidl = DllCall($hShell32, "ptr", "ILClone", "ptr", $last_SHITEMID)[0]
        DllStructSetData($tPidls, 1, $pRelativePidl, $i + 1)
        DllCall($hShell32, "none", "ILFree", "ptr", $pFullPidl)
    Next
    Local $tIID_IDataObject = _WinAPI_GUIDFromString($sIID_IDataObject)
    Local $pIDataObject = __SHCreateDataObject($tIID_IDataObject, $pParentPidl, $iCount, DllStructGetPtr($tPidls), 0)
    DllCall($hShell32, "none", "ILFree", "ptr", $pParentPidl)
    For $i = 1 To $iCount
        DllCall($hShell32, "none", "ILFree", "ptr", DllStructGetData($tPidls, 1, $i))
    Next
    If Not $pIDataObject Then Return 0
    Return $pIDataObject
EndFunc   ;==>GetDataObject

Func __SHCreateDataObject($tIID_IDataObject, $ppidlFolder = 0, $cidl = 0, $papidl = 0, $pdtInner = 0)
    Local $aRes = DllCall($hShell32, "long", "SHCreateDataObject", _
                                         "ptr", $ppidlFolder, _          
                                         "uint", $cidl, _
                                         "ptr", $papidl, _ 
                                         "ptr", $pdtInner, _
                                         "struct*", $tIID_IDataObject, _
                                         "ptr*", 0)
    If @error Then Return SetError(1, 0, $aRes[0])
    Return $aRes[6]
EndFunc   ;==>__SHCreateDataObject
