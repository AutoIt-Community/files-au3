#include-once

#include <GUIConstants.au3>
#include <WindowsConstants.au3>
#include <StructureConstants.au3>
#include <ApiErrorsConstants.au3>
#include <WinAPIConstants.au3>
#include <Clipboard.au3>

Global Const $S_FALSE = 1

Global Const $sCLSID_DragDropHelper = "{4657278A-411B-11D2-839A-00C04FD918D0}"
Global Const $sIID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Global Const $sIID_IDropSource = "{00000121-0000-0000-C000-000000000046}"
Global Const $sIID_IDropSourceNotify = "{0000012B-0000-0000-C000-000000000046}"
Global Const $sIID_IEnumFORMATETC = "{00000103-0000-0000-C000-000000000046}"
Global Const $sIID_IDropTarget = "{00000122-0000-0000-C000-000000000046}"
Global Const $sIID_IDragSourceHelper = "{DE5BF786-477A-11D2-839D-00C04FD918D0}"
Global Const $sIID_IDragSourceHelper2 = "{83E07D0D-0C5F-4163-BF1A-60B274051E40}"
Global Const $sIID_IDropTargetHelper = "{4657278B-411B-11D2-839A-00C04FD918D0}"

Global Const $CLSCTX_INPROC_SERVER = 0x01
Global Const $CLSCTX_INPROC_HANDLER = 0x2
Global Const $CLSCTX_LOCAL_SERVER = 0x4
Global Const $CLSCTX_REMOTE_SERVER = 0x10
Global Const $CLSCTX_NO_CODE_DOWNLOAD = 0x400
Global Const $CLSCTX_NO_CUSTOM_MARSHAL = 0x1000
Global Const $CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000
Global Const $CLSCTX_NO_FAILURE_LOG = 0x4000
Global Const $CLSCTX_DISABLE_AAA = 0x8000
Global Const $CLSCTX_ENABLE_AAA = 0x10000
Global Const $CLSCTX_FROM_DEFAULT_CONTEXT = 0x20000
Global Const $CLSCTX_ACTIVATE_X86_SERVER = 0x40000
Global Const $CLSCTX_ACTIVATE_32_BIT_SERVER = $CLSCTX_ACTIVATE_X86_SERVER
Global Const $CLSCTX_ACTIVATE_64_BIT_SERVER = 0x80000
Global Const $CLSCTX_ENABLE_CLOAKING = 0x100000
Global Const $CLSCTX_APPCONTAINER = 0x400000
Global Const $CLSCTX_ACTIVATE_AAA_AS_IU = 0x800000
Global Const $CLSCTX_ACTIVATE_ARM32_SERVER = 0x2000000
Global Const $CLSCTX_ALLOW_LOWER_TRUST_REGISTRATION = 0x4000000
Global Const $CLSCTX_ALL = BitOR($CLSCTX_INPROC_SERVER, $CLSCTX_INPROC_HANDLER, $CLSCTX_LOCAL_SERVER, $CLSCTX_REMOTE_SERVER)

Global Const $CFSTR_PERFORMEDDROPEFFECT = "Performed DropEffect"

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
Global Const $MK_ALT = 0x0020
Global Const $MK_XBUTTON1 = 0x0020
Global Const $MK_XBUTTON2 = 0x0040

Global Const $TYMED_HGLOBAL = 1
Global Const $TYMED_FILE = 2
Global Const $TYMED_ISTREAM = 4
Global Const $TYMED_ISTORAGE = 8
Global Const $TYMED_GDI = 16
Global Const $TYMED_MFPICT = 32
Global Const $TYMED_ENHMF = 64
Global Const $TYMED_NULL = 0

Global Const $DVASPECT_CONTENT = 1
Global Const $DVASPECT_THUMBNAIL = 2
Global Const $DVASPECT_ICON = 4
Global Const $DVASPECT_DOCPRINT = 8

Global Const $DV_E_FORMATETC = 0x80040064
Global Const $DV_E_DVTARGETDEVICE = 0x80040065
Global Const $DV_E_STGMEDIUM = 0x80040066
Global Const $DV_E_STATDATA = 0x80040067
Global Const $DV_E_LINDEX = 0x80040068
Global Const $DV_E_TYMED = 0x80040069
Global Const $DV_E_CLIPFORMAT = 0x8004006A
Global Const $DV_E_DVASPECT = 0x8004006B
Global Const $DV_E_DVTARGETDEVICE_SIZE = 0x8004006C
Global Const $DV_E_NOIVIEWOBJECT = 0x8004006D

Global Const $DATADIR_GET = 1
Global Const $DATADIR_SET = 2

Global Const $CFSTR_DROPDESCRIPTION = "DropDescription"

Global Const $DROPIMAGE_INVALID = -1
Global Const $DROPIMAGE_NONE = 0
Global Const $DROPIMAGE_COPY = 1
Global Const $DROPIMAGE_MOVE = 2
Global Const $DROPIMAGE_LINK = 4
Global Const $DROPIMAGE_LABEL = 6
Global Const $DROPIMAGE_WARNING = 7
Global Const $DROPIMAGE_NOIMAGE = 8

Global Const $DSH_ALLOWDROPDESCRIPTIONTEXT = 0x0001

Global Const $tagFORMATETC = "struct;word cfFormat;ptr ptd;dword aspect;long index;dword tymed;endstruct"
Global Const $tagSTGMEDIUM = "struct;dword tymed;ptr handle;ptr pUnkForRelease;endstruct"
Global Const $tagDROPFILES = "struct;dword pFiles;long pt[2];bool fNC;bool fWide;endstruct"
Global Const $tagSHDRAGIMAGE = "struct;long aiDragImageSize[2];long aiOffset[2];handle hDragImage;dword iColorKey;endstruct"
Global Const $tagDROPDESCRIPTION = "struct;int iType;wchar sMessage[260];wchar sInsert[260];endstruct"
Global Const $tagSHFILEINFOW = "struct;handle hIcon;int iIcon;dword iAttributes;wchar sDisplayName[260];wchar sTypeName[80];endstruct;"

;~ 		"ParseDisplayName hresult(hwnd; ptr; wstr; ulong*; ptr*; ulong*);" & _

Global Const $tagIDataObject = _
		"GetData hresult(struct*; struct*);" & _
		"GetDataHere hresult(struct*; struct*);" & _
		"QueryGetData hresult(struct*);" & _
		"GetCanonicalFormatEtc hresult(struct*; struct*);" & _
		"SetData hresult(struct*; struct*; bool);" & _
		"EnumFormatEtc hresult(dword; ptr*);" & _
		"DAdvise hresult(struct*; dword; ptr*; dword*);" & _
		"DUnadvise hresult(dword);" & _
		"EnumDAdvise hresult(ptr*);"

Global Const $tagIEnumFORMATETC = _
		"Next hresult(ulong; ptr; ulong*);" & _
		"Skip hresult(ulong);" & _
		"Reset hresult();" & _
		"Clone hresult(ptr*);"

Global Const $tagIDragSourceHelper = _
		"InitializeFromBitmap hresult(struct*; ptr);" & _
		"InitializeFromWindow hresult(hwnd; ptr; ptr);"

Global Const $tagIDragSourceHelper2 = $tagIDragSourceHelper & _
		"SetFlags hresult(dword);"

Global Const $tagIDropTargetHelper = _
		"DragEnter hresult(hwnd; ptr; struct*; dword);" & _
		"DragLeave hresult();" & _
		"DragOver hresult(struct*; dword);" & _
		"Drop hresult(ptr; struct*; dword);" & _
		"Show hresult(bool);"
