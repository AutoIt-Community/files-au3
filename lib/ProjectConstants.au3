#include-once

#include <GUIConstants.au3>
#include <WindowsConstants.au3>
#include <StructureConstants.au3>
#include <ApiErrorsConstants.au3>
#include <WinAPIConstants.au3>

Global Const $PTR_LEN = @AutoItX64 ? 8 : 4

Global Const $sIID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Global Const $sIID_IShellFolder = "{000214E6-0000-0000-C000-000000000046}"
Global Const $sIID_IDataObject = "{0000010e-0000-0000-C000-000000000046}"
Global Const $sIID_IDropSource = "{00000121-0000-0000-C000-000000000046}"
Global Const $sIID_IDropSourceNotify = "{0000012B-0000-0000-C000-000000000046}"


Global Const $DRAGDROP_S_DROP = 0x00040100
Global Const $DRAGDROP_S_CANCEL = 0x00040101
Global Const $DRAGDROP_S_USEDEFAULTCURSORS = 0x00040102
Global Const $DRAGDROP_E_NOTREGISTERED = 0x80040100
Global Const $DRAGDROP_E_ALREADYREGISTERED = 0x80040101

Global Const $DROPEFFECT_NONE = 0
Global Const $DROPEFFECT_COPY = 1
Global Const $DROPEFFECT_MOVE = 2
Global Const $DROPEFFECT_LINK = 4
Global Const $DROPEFFECT_SCROLL = 0x80000000

Global Const $MK_LBUTTON = 0x0001
Global Const $MK_RBUTTON = 0x0002
Global Const $MK_SHIFT = 0x0004
Global Const $MK_CONTROL = 0x0008
Global Const $MK_MBUTTON = 0x0010
Global Const $MK_XBUTTON1 = 0x0020
Global Const $MK_XBUTTON2 = 0x0040

Global Const $tagIShellFolder = _
		"ParseDisplayName hresult(hwnd; ptr*; wstr; ulong*; ptr*; ulong*);" & _
		"EnumObjects hresult(hwnd; ulong; ptr*);" & _
		"BindToObject hresult(ptr; ptr*; struct*; ptr*);" & _
		"BindToStorage hresult(ptr; ptr*; struct*; ptr*);" & _
		"CompareIDs hresult(lparam; ptr; ptr);" & _
		"CreateViewObject hresult(hwnd; struct*; ptr*);" & _
		"GetAttributesOf hresult(uint; ptr; ulong*);" & _
		"GetUIObjectOf hresult(hwnd; uint; ptr*; struct*; uint*; ptr*);" & _
		"GetDisplayNameOf hresult(ptr; ulong; ptr*);" & _
		"SetNameOf hresult(hwnd; ptr; wstr; ulong; ptr*);"