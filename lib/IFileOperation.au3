#include-once

#include <WinAPIShellEx.au3>

; originally from Danyfirex

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
Global Const $dtagIFileOperation = "Advise hresult(ptr;dword*);" & _
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

; Local $iFlags = BitOR($FOF_NOERRORUI, $FOFX_KEEPNEWERFILE, $FOFX_NOCOPYHOOKS, $FOF_NOCONFIRMATION)
; FOFX_ADDUNDORECORD (preferred)
; FOFX_RECYCLEONDELETE
; FOFX_NOCOPYHOOKS

;Local $sAction = "CopyItems"
;Local $sAction = "MoveItems"

;_IFileOperationFile($pDataObj, $sPathTo, $sAction, $iFlags)

Func _IFileOperationFile($pDataObj, $sPathTo, $sAction, $iFlags = 0)

    If Not FileExists($sPathTo) Then
        DirCreate($sPathTo)
    EndIf


    Local $tIIDIShellItem = CLSIDFromString($IID_IShellItem)
    Local $tIIDIShellItemArray = CLSIDFromString($IID_IShellItemArray)


    Local $oIFileOperation = ObjCreateInterface($CLSID_IFileOperation, $IID_IFileOperation, $dtagIFileOperation)
    If Not IsObj($oIFileOperation) Then Return SetError(2, 0, False)

    Local $pIShellItemTo = 0

    _SHCreateItemFromParsingName($sPathTo, 0, DllStructGetPtr($tIIDIShellItem), $pIShellItemTo)


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


Func _SHCreateItemFromParsingName($szPath, $pbc, $riid, ByRef $pv)
    Local $aRes = DllCall("shell32.dll", "long", "SHCreateItemFromParsingName", "wstr", $szPath, "ptr", $pbc, "ptr", $riid, "ptr*", 0)
    If @error Then Return SetError(1, 0, @error)
    $pv = $aRes[4]
    Return $aRes[0]
EndFunc   ;==>_SHCreateItemFromParsingName


Func CLSIDFromString($sString)
    Local $tCLSID = DllStructCreate("dword;word;word;byte[8]")
    Local $aRet = DllCall("Ole32.dll", "long", "CLSIDFromString", "wstr", $sString, "ptr", DllStructGetPtr($tCLSID))
    If @error Then Return SetError(1, 0, @error)
    If $aRet[0] <> 0 Then Return SetError(2, $aRet[0], 0)
    Return $tCLSID
EndFunc   ;==>CLSIDFromString
