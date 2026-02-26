#NoTrayIcon
#AutoIt3Wrapper_UseX64=Y
;#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include <GUIConstantsEx.au3>
#include <GuiToolTip.au3>
#include <GuiTreeView.au3>
#include <GuiListView.au3>
#include <String.au3>
#include <WinAPITheme.au3>
#include <WindowsConstants.au3>
#include <WindowsNotifsConstants.au3>
#include <WindowsStylesConstants.au3>

Global $hKernel32 = DllOpen('kernel32.dll')
Global $hGdi32 = DllOpen('gdi32.dll')
Global $hUser32 = DllOpen('user32.dll')
Global $hShlwapi = DllOpen('shlwapi.dll')
Global $hShell32 = DllOpen('shell32.dll')

#include "../lib/SharedFunctions.au3"
#include "../lib/GUIFrame_WBD_Mod.au3"
#include "../lib/History.au3"
#include "../lib/TreeListExplorer.au3"
#include "../lib/ProjectConstants.au3"
#include "../lib/DropSourceObject.au3"
#include "../lib/DropTargetObject.au3"

; CREDITS:
; Kanashius     TreeListExplorer UDF
; SOLVE-SMART	Code review, organization and refactoring
; pixelsearch   Detached Header and ListView synchronization
; ioa747        Detached Header subclassing for dark mode
; Nine          Custom Draw for Buttons
; argumentum    Dark Mode functions
; NoNameCode    Dark Mode functions
; Melba23       GUIFrame UDF
; ahmet         Non-client painting of white menubar line in dark mode
; UEZ           Lots and lots and lots
; DonChunior    Code review, bug fixes and refactoring
; MattyD		Drag and drop code
; jugador		ListView multiple item drag and drop
; Danyfirex		IFileOperation code

Global $sVersion = "0.4.0 - 2026-01-22"

; set base DPI scale value and apply DPI
Global $DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = -2
Global $iDPI = 1
$iDPI = ApplyDPI()

; DPI must be set before ownerdrawn menu
#include "../lib/ModernMenuRaw.au3"

Opt("GUIOnEventMode", 1)
Opt("GUICloseOnESC", 0)

Global $hTLESystem, $iFrame_A, $hSeparatorFrame, $aWinSize2, $idInputPath, $g_hInputPath, $g_hStatus, $idTreeView
Global $g_hGUI, $g_hChild, $g_hHeader, $g_hListview, $idListview, $iHeaderHeight, $hParentFrame, $g_iIconWidth, $g_hTreeView
Global $g_hSizebox, $g_hOldProc, $g_iHeight, $g_hDots
Global $idPropertiesItem, $idPropertiesLV, $sCurrentPath
Global $hListImgList, $iListDragIndex, $sTargetCtrl, $hTreeItemOrig, $hIcon
Global $sBack, $sForward, $sUpLevel, $sRefresh
Global $sTreeDragItem, $sListDragItems, $bDragToolActive = False
Global $pLVDropTarget, $pTVDropTarget
Global $bLoadStatus = False, $bCursorOverride = False
Global $idExitItem, $idAboutItem, $idDeleteItem, $idRenameItem, $idCopyItem, $idPasteItem, $idUndoItem, $idHiddenItem, $idSystemItem
Global $bHideHidden = False, $bHideSystem = False
Global $hCursor, $hProc
Global $sSelectedItems, $g_aText[4], $gText
Global $idSeparator, $idThemeItem, $hToolTip1, $hToolTip2, $bTooltipActive
Global $isDarkMode = _WinAPI_ShouldAppsUseDarkMode()
Global $hFolderHistory = __History_Create("_doUnReDo", 100, "_historyChange"), $bFolderHistoryChanging = False
Global $hSolidBrush = _WinAPI_CreateBrushIndirect($BS_SOLID, 0x000000)
Global $iTopSpacer = Round(12 * $iDPI)
Global $aPosTip, $iOldaPos0, $iOldaPos1
Global $sRenameFrom, $sControlFocus, $bFocusChanged = False, $bSelectChanged = False, $bSaveEdit = False
Global $bCopy = False, $pCopyObj
; force light mode
;$isDarkMode = False

; get Windows build
Global $iOSBuild = @OSBuild
Global $iRevision = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "UBR")

; button colors
If $isDarkMode Then
	Global $iBackColorDef = 0x202020
	Global $iBackColorDis = 0x202020
	If $iOSBuild >= 22621 Then
		$iBackColorDef = 0x000000
		$iBackColorDis = 0x000000
	EndIf
	Global $iBackColorHot = _WinAPI_ColorAdjustLuma($iBackColorDef, 30)
	Global $iBackColorSel = _WinAPI_ColorAdjustLuma($iBackColorDef, 10)
	Global $iTextColorDef = 0xFFFFFF
	Global $iTextColorDis = _WinAPI_ColorAdjustLuma($iTextColorDef, -50)
Else
	Global $iBackColorDef = _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE))
	Global $iBackColorHot = _WinAPI_ColorAdjustLuma($iBackColorDef, -10)
	Global $iBackColorSel = _WinAPI_ColorAdjustLuma($iBackColorDef, -5)
	Global $iBackColorDis = _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE))
	Global $iTextColorDef = _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT))
	Global $iTextColorDis = _WinAPI_ColorAdjustLuma(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT)), 70)
EndIf

; Structure Definitions (using $tagNMHDR and $tagRECT which are defined in includes)
Global Const $tagNMCUSTOMDRAW = $tagNMHDR & ";dword DrawStage;handle hdc;" & $tagRECT & ";dword_ptr ItemSpec;uint ItemState;lparam lItemParam;"

If $isDarkMode Then
	_SetMenuBkColor($iBackColorDef)
	_SetMenuSelectBkColor(_WinAPI_ColorAdjustLuma($iBackColorDef, 30))
	_SetMenuSelectRectColor(_WinAPI_ColorAdjustLuma($iBackColorDef, 30))
	_SetMenuSelectTextColor(0xffffff)
	_SetMenuTextColor(0xffffff)
Else
	_SetMenuBkColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)))
	_SetMenuSelectBkColor(_WinAPI_ColorAdjustLuma(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)), -6))
	_SetMenuSelectRectColor(_WinAPI_ColorAdjustLuma(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)), -6))
	_SetMenuSelectTextColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT)))
	_SetMenuTextColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT)))
EndIf

OnAutoItExitRegister(_CleanExit)

_FilesAu3()

Func _FilesAu3()
	_GDIPlus_Startup()
	Local $sButtonSpacing = 20
	; calculate icon size for TreeListExplorer
	Local $iTreeListIconSize = 16 * $iDPI

	; check font availability
	Local $sButtonFont
	If _WinAPI_GetFontName("Segoe Fluent Icons") Then
		; Segoe Fluent Icons are available, use for buttons (Windows 11)
		$sButtonFont = "Segoe Fluent Icons"
	Else
		; Segoe Fluent Icons are not available, fall back to Segoe MDL2 Assets (Windows 10)
		$sButtonFont = "Segoe MDL2 Assets"
	EndIf

	; Startup of the TreeListExplorer
	__TreeListExplorer_StartUp($__TreeListExplorer_Lang_EN)

	; Create GUI and register events
	$g_hGUI = GUICreate("Files Au3", @DesktopWidth - 600, @DesktopHeight - 400, -1, -1, $WS_OVERLAPPEDWINDOW)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_EventsGUI")
	GUISetOnEvent($GUI_EVENT_MAXIMIZE, "_EventsGUI")
	GUISetOnEvent($GUI_EVENT_RESIZED, "_EventsGUI")
	GUISetOnEvent($GUI_EVENT_DROPPED, "_EventsGUI")

	; used to determine separator position
	$FrameWidth1 = (@DesktopWidth - 600) / 3

	_InitDarkSizebox()

	; statusbar create
	$g_hStatus = _GUICtrlStatusBar_Create($g_hGUI, -1, "", $WS_CLIPSIBLINGS)

	$aClientSize = WinGetClientSize($g_hGUI)

	; set tooltips theming
	$hToolTip1 = _GUIToolTip_Create(0)
	_GUIToolTip_SetMaxTipWidth($hToolTip1, 400)
	$hToolTip2 = _GUIToolTip_Create(0)
	_GUIToolTip_SetMaxTipWidth($hToolTip2, 400)

	GUISetFont(10, $FW_NORMAL, $GUI_FONTNORMAL, $sButtonFont)

	$sBack = GUICtrlCreateButton(ChrW(0xE64E), $sButtonSpacing, 10, -1, -1)
	GUICtrlSetOnEvent(-1, "_ButtonFunctions")
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKWIDTH)
	$aPos = ControlGetPos($g_hGUI, "", $sBack)
	$sBackPosV = $aPos[1] + $aPos[3]
	$sBackPosH = $aPos[0] + $aPos[2]
	$iButtonHeight = $aPos[3]
	GUICtrlSetState($sBack, $GUI_DISABLE)
	_GUIToolTip_AddTool($hToolTip2, $g_hGUI, "Back", GUICtrlGetHandle($sBack))

	$sForward = GUICtrlCreateButton(ChrW(0xE64D), $sBackPosH + $sButtonSpacing, 10, -1, -1)
	GUICtrlSetOnEvent(-1, "_ButtonFunctions")
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKWIDTH)
	$aPos = ControlGetPos($g_hGUI, "", $sForward)
	$sForwardPosV = $aPos[1] + $aPos[3]
	$sForwardPosH = $aPos[0] + $aPos[2]
	GUICtrlSetState($sForward, $GUI_DISABLE)
	_GUIToolTip_AddTool($hToolTip2, $g_hGUI, "Forward", GUICtrlGetHandle($sForward))

	$sUpLevel = GUICtrlCreateButton(ChrW(0xE64C), $sForwardPosH + $sButtonSpacing, 10, -1, -1)
	GUICtrlSetOnEvent(-1, "_ButtonFunctions")
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKWIDTH)
	$aPos = ControlGetPos($g_hGUI, "", $sUpLevel)
	$sUpLevelPosV = $aPos[1] + $aPos[3]
	$sUpLevelPosH = $aPos[0] + $aPos[2]
	_GUIToolTip_AddTool($hToolTip2, $g_hGUI, "Up", GUICtrlGetHandle($sUpLevel))

	$sRefresh = GUICtrlCreateButton(ChrW(0xE72C), $sUpLevelPosH + $sButtonSpacing, 10, -1, -1)
	GUICtrlSetOnEvent(-1, "_ButtonFunctions")
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKWIDTH)
	$aPos = ControlGetPos($g_hGUI, "", $sRefresh)
	$sRefreshPosV = $aPos[1] + $aPos[3]
	$sRefreshPosH = $aPos[0] + $aPos[2]
	_GUIToolTip_AddTool($hToolTip2, $g_hGUI, "Refresh", GUICtrlGetHandle($sRefresh))

	; reset GUI font
	GUISetFont(10, $FW_NORMAL, $GUI_FONTNORMAL, "Segoe UI")

	; Menubar
	Local $idFileMenu = _GUICtrlCreateODTopMenu("& File", $g_hGUI)
	Local $idEditMenu = _GUICtrlCreateODTopMenu("& Edit", $g_hGUI)
	Local $idViewMenu = _GUICtrlCreateODTopMenu("& View", $g_hGUI)
	Local $idHelpMenu = _GUICtrlCreateODTopMenu("& Help", $g_hGUI)

	; File menu
	$idDeleteItem = GUICtrlCreateMenuItem("&Delete" & @TAB & "Delete", $idFileMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlSetState($idDeleteItem, $GUI_DISABLE)
	$idRenameItem = GUICtrlCreateMenuItem("&Rename", $idFileMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlSetState($idRenameItem, $GUI_DISABLE)
	$idPropertiesItem = GUICtrlCreateMenuItem("&Properties" & @TAB & "Shift+P", $idFileMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlSetState($idPropertiesItem, $GUI_DISABLE)
	GUICtrlCreateMenuItem("", $idFileMenu)
	$idExitItem = GUICtrlCreateMenuItem("&Exit", $idFileMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	; Edit menu
	$idUndoItem = GUICtrlCreateMenuItem("&Undo" & @TAB & "Ctrl+Z", $idEditMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlSetState($idUndoItem, $GUI_DISABLE)
	GUICtrlCreateMenuItem("", $idEditMenu)
	$idCopyItem = GUICtrlCreateMenuItem("&Copy" & @TAB & "Ctrl+C", $idEditMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlSetState($idCopyItem, $GUI_DISABLE)
	$idPasteItem = GUICtrlCreateMenuItem("&Paste" & @TAB & "Ctrl+V", $idEditMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlSetState($idPasteItem, $GUI_DISABLE)
	; View menu
	$idThemeItem = GUICtrlCreateMenuItem("&Dark Mode", $idViewMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlCreateMenuItem("", $idViewMenu)
	$idHiddenItem = GUICtrlCreateMenuItem("&Show Hidden Files", $idViewMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlSetState($idHiddenItem, $GUI_CHECKED)
	$bHideHidden = False
	$idSystemItem = GUICtrlCreateMenuItem("&Hide Protected System Files", $idViewMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")
	GUICtrlSetState($idSystemItem, $GUI_CHECKED)
	$bHideSystem = True
	; Help menu
	$idAboutItem = GUICtrlCreateMenuItem("&About", $idHelpMenu)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")

	If $isDarkMode Then GUICtrlSetState($idThemeItem, $GUI_CHECKED)

	If $isDarkMode Then
		If $iOSBuild >= 26100 And $iRevision >= 6899 Then
			$idInputPath = GUICtrlCreateInput("", $sRefreshPosH + ($sButtonSpacing * 1.5), 10, $aClientSize[0] - $sRefreshPosH - ($sButtonSpacing * 3), $iButtonHeight, -1, -1)
		Else
			$idInputPath = GUICtrlCreateInput("", $sRefreshPosH + ($sButtonSpacing * 1.5), 10, $aClientSize[0] - $sRefreshPosH - ($sButtonSpacing * 3), $iButtonHeight, -1, $WS_EX_STATICEDGE)
		EndIf
	Else
		$idInputPath = GUICtrlCreateInput("", $sRefreshPosH + ($sButtonSpacing * 1.5), 10, $aClientSize[0] - $sRefreshPosH - ($sButtonSpacing * 3), $iButtonHeight)
	EndIf
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKRIGHT)
	$g_hInputPath = GUICtrlGetHandle($idInputPath)

	Local $iStatusHeight = _WinAPI_GetWindowHeight($g_hStatus)
	$aClientSize2 = WinGetClientSize($g_hGUI)
	Local $iFrameHeight = $aClientSize2[1] - $iStatusHeight - $sRefreshPosV - $iTopSpacer
	$iFrame_A = _GUIFrame_Create($g_hGUI, 0, $FrameWidth1, 9, 0, $sRefreshPosV + $iTopSpacer)

	; Set minimum sizes for the frames
	_GUIFrame_SetMin($iFrame_A, 200, 600, True)

	; Create treeview frame
	_GUIFrame_Switch($iFrame_A, 1)
	$aWinSize1 = WinGetClientSize(_GUIFrame_GetHandle($iFrame_A, 1))

	; create treeview
	Local $iStyle = BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_SHOWSELALWAYS, $TVS_TRACKSELECT, $TVS_EDITLABELS)
	$idTreeView = GUICtrlCreateTreeView(0, 0, $aWinSize1[0], $iFrameHeight, $iStyle)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
	$g_hTreeView = GUICtrlGetHandle($idTreeView)

	; Create TLE system
	$hTLESystem = __TreeListExplorer_CreateSystem($g_hGUI, "", "_folderCallback")

	; Add Views to TLE system
	__TreeListExplorer_AddView($hTLESystem, $idInputPath)
	__TreeListExplorer_AddView($hTLESystem, $idTreeView)
	__TreeListExplorer_SetViewIconSize($idTreeView, $iTreeListIconSize)

	; set callback to allow filtering of hidden and/or protected system files
	__TreeListExplorer_SetCallback($idTreeView, $__TreeListExplorer_Callback_Filter, "_filterCallback")

	; Create listview frame
	_GUIFrame_Switch($iFrame_A, 2)

	$aWinSize2 = WinGetClientSize(_GUIFrame_GetHandle($iFrame_A, 2))

	; create header control
	Local $hChildLV = _GUIFrame_GetHandle($iFrame_A, 2)
	$g_hHeader = _GUICtrlHeader_Create($hChildLV, BitOR($HDS_BUTTONS, $HDS_DRAGDROP, $HDS_FULLDRAG))
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
	_GUICtrlHeader_AddItem($g_hHeader, "Name", 300)
	_GUICtrlHeader_AddItem($g_hHeader, "Size", 100)
	_GUICtrlHeader_AddItem($g_hHeader, "Date Modified", 150)
	_GUICtrlHeader_AddItem($g_hHeader, "Type", 150)

	; Set Size column alignment
	_GUICtrlHeader_SetItemAlign($g_hHeader, 1, 1)

	; Set sort arrow
	;_GUICtrlHeader_SetItemFormat($g_hHeader, 0, $HDF_SORTUP)

	; get header height
	$iHeaderHeight = _WinAPI_GetWindowHeight($g_hHeader)

	; create listview control
	Local $iExStyles = BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_DOUBLEBUFFER, $LVS_EX_TRACKSELECT)
	$idListview = GUICtrlCreateListView("Name|Size|Date Modified|Type", 0, $iHeaderHeight, $aWinSize2[0], $iFrameHeight - $iHeaderHeight, BitOR($LVS_SHOWSELALWAYS, $LVS_NOCOLUMNHEADER, $LVS_EDITLABELS), $iExStyles)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)

	$g_hListview = GUICtrlGetHandle($idListview)

	;Create Target for our GUI & register.
	$pLVDropTarget = CreateDropTarget($g_hListview)
	RegisterDragDrop($g_hListview, $pLVDropTarget)

	$pTVDropTarget = CreateDropTarget($g_hTreeView)
	RegisterDragDrop($g_hTreeView, $pTVDropTarget)

	_GUIToolTip_AddTool($hToolTip1, $g_hGUI, "", $g_hListview)

	; right align Size column
	_GUICtrlListView_JustifyColumn($idListview, 1, 1)

	; listview context menu
	Local $idContextLV = GUICtrlCreateContextMenu($idListview)
	$idPropertiesLV = GUICtrlCreateMenuItem("Properties", $idContextLV)
	GUICtrlSetOnEvent(-1, "_MenuFunctions")

	; add listview and callbacks to TLE system
	__TreeListExplorer_AddView($hTLESystem, $idListview, True, True, True, False, False)
	__TreeListExplorer_SetViewIconSize($idListview, $iTreeListIconSize)
	__TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_Loading, "_loadingCallback")
	__TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_DoubleClick, "_doubleClickCallback")
	__TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_ListViewPaths, "_handleListViewData")
	__TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_ListViewItemCreated, "_handleListViewItemCreated")

	; set callback to allow filtering of hidden and/or protected system files
	__TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_Filter, "_filterCallback")

	; Set resizing flag for all created frames
	_GUIFrame_ResizeSet(0)

	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND2")
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY2")

	; get rid of the focus rectangle dots
	GUICtrlSendMsg($idTreeView, $WM_CHANGEUISTATE, 65537, 0)

	; resize listview columns to match header widths + update global variable $g_iIconWidth
	_resizeLVCols()

	; set ownerdraw parts for statusbar
	Local $iParts = $aClientSize[0] / 4
	Local $aParts[4] = [$iParts, $iParts * 2, $iParts * 3, -1]
	Dim $g_aText[UBound($aParts)] = [" ", " ", " ", " "]
	_GUICtrlStatusBar_SetParts($g_hStatus, $aParts)
	_GUICtrlStatusBar_SetText($g_hStatus, $g_aText[0], 0, $SBT_OWNERDRAW)
	_GUICtrlStatusBar_SetText($g_hStatus, $g_aText[1], 1, $SBT_OWNERDRAW)
	_GUICtrlStatusBar_SetText($g_hStatus, $g_aText[2], 2, $SBT_OWNERDRAW)
	_GUICtrlStatusBar_SetText($g_hStatus, $g_aText[3], 3, $SBT_OWNERDRAW)

	_setThemeColors()

	GUIRegisterMsg($WM_MOVE, "WM_MOVE")
	GUIRegisterMsg($WM_SIZE, "WM_SIZE")
	GUIRegisterMsg($WM_DRAWITEM, "WM_DRAWITEM2")
	GUIRegisterMsg($WM_ACTIVATE, "WM_ACTIVATE_Handler")
	GUIRegisterMsg($WM_WINDOWPOSCHANGED, "WM_WINDOWPOSCHANGED_Handler")

	_GUICtrl_SetFont($g_hHeader, 16 * $iDPI, 400, 0, "Segoe UI")

	; update variable for header height
	$iHeaderHeight = _WinAPI_GetWindowHeight($g_hHeader)

	If $isDarkMode Then
		_WinAPI_DwmSetWindowAttribute_unr($g_hGUI, 38, 2)
		_WinAPI_DwmExtendFrameIntoClientArea($g_hGUI, _WinAPI_CreateMargins(-1, -1, -1, -1))
	EndIf

	; enumerate all visible and non-visible child windows
	Local $aWinList = _WinAPI_EnumWindows(False, _GetHwndFromPID(@AutoItPID))
	Local $aWindows[$aWinList[0][0]][3]
	For $i = 1 To $aWinList[0][0]
		$aWindows[$i - 1][0] = $aWinList[$i][0]
		$aWindows[$i - 1][1] = $aWinList[$i][1]
		$aWindows[$i - 1][2] = WinGetTitle($aWinList[$i][0])
	Next

	; get handles for parent frame and separator frame
	For $i = 0 To UBound($aWindows) - 1
		If $aWindows[$i][2] = "FrameParent" Then $hParentFrame = $aWindows[$i][0]
		If $aWindows[$i][2] = "SeparatorFrame" Then $hSeparatorFrame = $aWindows[$i][0]
	Next

	; set background color for separator frame and parent frame
	If $isDarkMode Then
		If @OSBuild >= 22621 Then
			GUISetBkColor(0x000000, $hSeparatorFrame)
			GUISetBkColor(0x000000, $hParentFrame)
		Else
			GUISetBkColor($iBackColorDef, $hSeparatorFrame)
			GUISetBkColor($iBackColorDef, $hParentFrame)
		EndIf
	Else
		GUISetBkColor($iBackColorDef, $hSeparatorFrame)
		GUISetBkColor($iBackColorDef, $hParentFrame)
	EndIf

	; get imagelist handles for treeview and listview
	$hListImgList = _GUICtrlListView_GetImageList($idListview, 1)

	GUISetState(@SW_SHOW, $g_hGUI)
	_drawUAHMenuNCBottomLine($g_hGUI)

	; set GUI icon
	_WinSetIcon($g_hGUI, @ScriptDir & "\app.ico")

	_removeExStyles()

	_GUICtrlListView_SetHoverTime($idListview, 500)

	; apply theme to tooltips
	_themeTooltips()

	; add composited to treeview frame
	Local $i_ExStyle_Old = _WinAPI_GetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 1), $GWL_EXSTYLE)
	_WinAPI_SetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 1), $GWL_EXSTYLE, BitOR($i_ExStyle_Old, $WS_EX_COMPOSITED))

	; add drop files support to treeview frame and listview frame
	$i_ExStyle_Old = _WinAPI_GetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 1), $GWL_EXSTYLE)
	_WinAPI_SetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 1), $GWL_EXSTYLE, BitOR($i_ExStyle_Old, $WS_EX_ACCEPTFILES))

	; add drop files support to treeview frame and listview frame
	$i_ExStyle_Old = _WinAPI_GetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 2), $GWL_EXSTYLE)
	_WinAPI_SetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 2), $GWL_EXSTYLE, BitOR($i_ExStyle_Old, $WS_EX_ACCEPTFILES))

	Local $sMsg = "Attention: The file operation code is new and caution is advised." & @CRLF & @CRLF
	$sMsg &= "Please consider testing any file operations in less important areas of your file system. Creating an area "
	$sMsg &= "on your file system with test folders and test files would be a good idea for testing purposes." & @CRLF & @CRLF
	$sMsg &= "At the moment, Files Au3 allows you to Undo the most recent drag and drop, copy, move, delete, rename, etc. "
	$sMsg &= "by pressing Ctrl+Z or Undo from the Edit menu. Future versions will expand to allow more than just the most "
	$sMsg &= "recent Undo operation."
	MsgBox($MB_ICONWARNING, "Files Au3", $sMsg)

	; TreeView has initial focus
	$sControlFocus = 'Tree'
	$bFocusChanged = True

	While True
		If $bTooltipActive Then
			; check if cursor is still over listview
			Local $aCursor = GUIGetCursorInfo($g_hGUI)
			If $aCursor[4] <> $idListview Then
				; cancel tooltip when not over listview
				_GUIToolTip_TrackActivate($hToolTip1, False, $g_hGUI, $g_hListview)
				; reset the value stored in the tooltip
				$gText = ""
				_GUIToolTip_UpdateTipText($hToolTip1, $g_hGUI, $g_hListview, $gText)
				$bTooltipActive = False
			EndIf
		EndIf

		If $bFocusChanged Or $bSelectChanged Then
			; keep track of which control currently has focus to determine which menu items to enable/disable
			$bFocusChanged = False
			$bSelectChanged = False
			Select
				Case $sControlFocus = 'List'
					; ListView currently has focus
					Local $aSelectedLV = _GUICtrlListView_GetSelectedIndices($idListview, True)
					If $aSelectedLV[0] = 0 Then
						; no selections in ListView currently
						GUICtrlSetState($idRenameItem, $GUI_DISABLE)
						GUICtrlSetState($idCopyItem, $GUI_DISABLE)
						HotKeySet("^c")
						GUICtrlSetState($idDeleteItem, $GUI_DISABLE)
						HotKeySet("{DELETE}")
						GUICtrlSetState($idPasteItem, $bCopy ? $GUI_ENABLE : $GUI_DISABLE)
						HotKeySet("^v", $bCopy ? "_PasteItems" : "")
						GUICtrlSetState($idPropertiesItem, $GUI_ENABLE)
						HotKeySet("+p", "_Properties")
						GUICtrlSetState($idPropertiesLV, $GUI_ENABLE)
					ElseIf $aSelectedLV[0] = 1 Then
						; 1 item selection in ListView
						GUICtrlSetState($idRenameItem, $GUI_ENABLE)
						GUICtrlSetState($idCopyItem, $GUI_ENABLE)
						HotKeySet("^c", "_CopyItems")
						GUICtrlSetState($idDeleteItem, $GUI_ENABLE)
						HotKeySet("{DELETE}", "_DeleteItems")
						Local $sSelectedItem = _GUICtrlListView_GetItemText($idListview, $aSelectedLV[1], 0)
						Local $sSelectedLV = __TreeListExplorer_GetPath($hTLESystem) & $sSelectedItem
						; is selected path a folder
						If StringInStr(FileGetAttrib($sSelectedLV), "D") Then
							GUICtrlSetState($idPasteItem, $bCopy ? $GUI_ENABLE : $GUI_DISABLE)
							HotKeySet("^v", $bCopy ? "_PasteItems" : "")
						Else
							GUICtrlSetState($idPasteItem, $GUI_DISABLE)
							HotKeySet("^v")
						EndIf
						GUICtrlSetState($idPropertiesItem, $GUI_ENABLE)
						HotKeySet("+p", "_Properties")
						GUICtrlSetState($idPropertiesLV, $GUI_ENABLE)
					Else
						; multiple items selected in ListView
						GUICtrlSetState($idRenameItem, $GUI_DISABLE) ; not supporting multiple file Rename right now
						GUICtrlSetState($idCopyItem, $GUI_ENABLE)
						HotKeySet("^c", "_CopyItems")
						GUICtrlSetState($idDeleteItem, $GUI_ENABLE)
						HotKeySet("{DELETE}", "_DeleteItems")
						GUICtrlSetState($idPasteItem, $GUI_DISABLE)
						HotKeySet("^v")
						GUICtrlSetState($idPropertiesItem, $GUI_ENABLE)
						HotKeySet("+p", "_Properties")
						GUICtrlSetState($idPropertiesLV, $GUI_ENABLE)
					EndIf
				Case $sControlFocus = 'Tree'
					; TreeView currently has focus
					; treeview always has a selection
					GUICtrlSetState($idRenameItem, $GUI_ENABLE)
					GUICtrlSetState($idCopyItem, $GUI_ENABLE)
					HotKeySet("^c", "_CopyItems")
					GUICtrlSetState($idDeleteItem, $GUI_ENABLE)
					HotKeySet("{DELETE}", "_DeleteItems")
					GUICtrlSetState($idPasteItem, $bCopy ? $GUI_ENABLE : $GUI_DISABLE)
					HotKeySet("^v", $bCopy ? "_PasteItems" : "")
					GUICtrlSetState($idPropertiesItem, $GUI_ENABLE)
					HotKeySet("+p", "_Properties")
					GUICtrlSetState($idPropertiesLV, $GUI_DISABLE)
				Case Not $sControlFocus
					; Neither the ListView or TreeView has focus right now
					; in this case likely disable menu options
					GUICtrlSetState($idRenameItem, $GUI_DISABLE)
					GUICtrlSetState($idCopyItem, $GUI_DISABLE)
					HotKeySet("^c")
					GUICtrlSetState($idDeleteItem, $GUI_DISABLE)
					HotKeySet("{DELETE}")
					GUICtrlSetState($idPasteItem, $GUI_DISABLE)
					HotKeySet("^v")
					GUICtrlSetState($idPropertiesItem, $GUI_DISABLE)
					HotKeySet("+p")
					GUICtrlSetState($idPropertiesLV, $GUI_DISABLE)
			EndSelect
		EndIf

		Sleep(200)
	WEnd
EndFunc   ;==>_FilesAu3

Func _themeTooltips()
	Local $aData = _WinAPI_EnumProcessWindows(0, False)

	For $i = 1 To $aData[0][0]
		If $aData[$i][1] = "tooltips_class32" Then
			If $isDarkMode Then
				_WinAPI_SetWindowTheme($aData[$i][0], 'DarkMode_Explorer', 'ToolTip')
			Else
				_WinAPI_SetWindowTheme($aData[$i][0], 'Explorer', 'ToolTip')
			EndIf
		EndIf
	Next
EndFunc   ;==>_themeTooltips

Func _switchTheme()
	_addExStyles()
	If $isDarkMode Then
		; switch to light mode
		$isDarkMode = False
		_WinAPI_DwmSetWindowAttribute_unr($g_hGUI, 38, 0)
		_WinAPI_DwmExtendFrameIntoClientArea($g_hGUI, _WinAPI_CreateMargins(0, 0, 0, 0))
		_GUISetDarkTheme($g_hGUI, False)
		GUICtrlSetBkColor($idSeparator, 0x909090)
		GUISetBkColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE)), $hSeparatorFrame)
		$iBackColorDef = _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE))
		$iBackColorHot = _WinAPI_ColorAdjustLuma($iBackColorDef, -10)
		$iBackColorSel = _WinAPI_ColorAdjustLuma($iBackColorDef, -5)
		$iBackColorDis = _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE))
		$iTextColorDef = _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT))
		$iTextColorDis = _WinAPI_ColorAdjustLuma(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT)), 70)

		GUISetBkColor($iBackColorDef, $hSeparatorFrame)
		GUISetBkColor($iBackColorDef, $hParentFrame)

		_RefreshButtons()

		_SetMenuBkColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)))
		_SetMenuSelectBkColor(_WinAPI_ColorAdjustLuma(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)), -6))
		_SetMenuSelectRectColor(_WinAPI_ColorAdjustLuma(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)), -6))
		_SetMenuSelectTextColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT)))
		_SetMenuTextColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT)))

		_GUIMenuBarSetBkColor($g_hGUI, _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)))

		GUICtrlSetState($idThemeItem, $GUI_UNCHECKED)
		_themeTooltips()
		_setThemeColors()
		_WinAPI_RedrawWindow($g_hStatus)
	Else
		; switch to dark mode
		$isDarkMode = True
		_WinAPI_DwmSetWindowAttribute_unr($g_hGUI, 38, 2)
		_WinAPI_DwmExtendFrameIntoClientArea($g_hGUI, _WinAPI_CreateMargins(-1, -1, -1, -1))
		_GUISetDarkTheme($g_hGUI)
		$iBackColorDef = 0x202020
		$iBackColorDis = 0x202020
		If $iOSBuild >= 22621 Then
			$iBackColorDef = 0x000000
			$iBackColorDis = 0x000000
		EndIf
		$iBackColorHot = _WinAPI_ColorAdjustLuma($iBackColorDef, 30)
		$iBackColorSel = _WinAPI_ColorAdjustLuma($iBackColorDef, 10)
		$iTextColorDef = 0xFFFFFF
		$iTextColorDis = _WinAPI_ColorAdjustLuma($iTextColorDef, -50)

		GUICtrlSetBkColor($idSeparator, 0x505050)
		If @OSBuild >= 22621 Then
			GUISetBkColor(0x000000, $hSeparatorFrame)
			GUISetBkColor(0x000000, $hParentFrame)
		Else
			GUISetBkColor($iBackColorDef, $hSeparatorFrame)
			GUISetBkColor($iBackColorDef, $hParentFrame)
		EndIf

		_RefreshButtons()

		_SetMenuBkColor($iBackColorDef)
		_SetMenuSelectBkColor(_WinAPI_ColorAdjustLuma($iBackColorDef, 30))
		_SetMenuSelectRectColor(_WinAPI_ColorAdjustLuma($iBackColorDef, 30))
		_SetMenuSelectTextColor(0xffffff)
		_SetMenuTextColor(0xffffff)

		_GUIMenuBarSetBkColor($g_hGUI, $iBackColorDef)

		GUICtrlSetState($idThemeItem, $GUI_CHECKED)
		_themeTooltips()
		_setThemeColors()
		_WinAPI_RedrawWindow($g_hStatus)
	EndIf
EndFunc   ;==>_switchTheme

Func _RefreshButtons()
	; Refresh navigation buttons to apply theme changes
	; This forces a redraw by hiding and showing the controls
	Local $aButtons = [$sBack, $sForward, $sUpLevel, $sRefresh]
	For $i = 0 To UBound($aButtons) - 1
		GUICtrlSetState($aButtons[$i], $GUI_HIDE)
		GUICtrlSetState($aButtons[$i], $GUI_SHOW)
	Next
EndFunc   ;==>_RefreshButtons

Func _selectionChangedLV()
	Local $sSelectedLV, $iDirCount = 0, $iFileCount = 0, $iFileSizes = 0
	; get selections from listview
	Local $aSelectedLV = _GUICtrlListView_GetSelectedIndices($idListview, True)
	Local $sSelectedItem = ""
	$sSelectedItems = ""

	For $i = 1 To $aSelectedLV[0]
		$sSelectedItem = _GUICtrlListView_GetItemText($idListview, $aSelectedLV[$i], 0)
		$sSelectedLV = __TreeListExplorer_GetPath($hTLESystem) & $sSelectedItem
		; is selected path a folder
		If StringInStr(FileGetAttrib($sSelectedLV), "D") Then
			; increase folder count
			$iDirCount += 1
			$sSelectedItems &= $sSelectedItem & "|"
		Else
			; increase file count
			$iFileCount += 1
			$sSelectedItems &= $sSelectedItem & "|"
			; count file sizes
			$iFileSizes += FileGetSize($sSelectedLV)
		EndIf
	Next

	Local $iItemCount = $iFileCount + $iDirCount

	If $iItemCount > 1 Then
		$g_aText[1] = "  " & $iItemCount & " items selected"
	ElseIf $iItemCount = 1 Then
		$g_aText[1] = "  1 item selected"
	Else
		$g_aText[1] = " "
	EndIf

	If $iFileSizes = 0 Then
		; clear size status if size equals zero
		$g_aText[2] = " "
	Else
		$g_aText[2] = "  " & __TreeListExplorer__GetSizeString($iFileSizes)
	EndIf

	; update number of items (files and folders) in statusbar
	Local $iLVItemCount = _GUICtrlListView_GetItemCount($idListview)
	$g_aText[0] = "  " & $iLVItemCount & " item"
	If $iLVItemCount > 1 Then
		$g_aText[0] &= "s"
	EndIf

	_WinAPI_RedrawWindow($g_hStatus)
EndFunc   ;==>_selectionChangedLV

Func _handleListViewData($hSystem, $hView, $sPath, ByRef $arPaths)
	ReDim $arPaths[UBound($arPaths)][_GUICtrlListView_GetColumnCount($idListview)] ; resize the array (and return it at the end)
	For $i = 0 To UBound($arPaths) - 1
		Local $sFilePath = $sPath & $arPaths[$i][0]
		If __TreeListExplorer__PathIsFolder($sFilePath) Then
			$arPaths[$i][2] = __TreeListExplorer__GetTimeString(FileGetTime($sFilePath, 0))             ; add time modified
			$arPaths[$i][3] = _getType($sFilePath, True)
		Else
			$arPaths[$i][1] = FileGetSize($sFilePath)             ; Put size as integer numbers here to enable the default sorting
			$arPaths[$i][2] = __TreeListExplorer__GetTimeString(FileGetTime($sFilePath, 0))             ; add time modified
			$arPaths[$i][3] = _getType($sFilePath)
		EndIf
	Next
	; custom sorting could be done here as well, setting the parameter $bEnableSorting to False when adding the ListView. Sorting can then be handled by the user
	Return $arPaths
EndFunc   ;==>_handleListViewData

Func _handleListViewItemCreated($hSystem, $hView, $sPath, $sFilename, $iIndex, $bFolder)
	If Not $bFolder Then _GUICtrlListView_SetItemText($hView, $iIndex, __TreeListExplorer__GetSizeString(_GUICtrlListView_GetItemText($hView, $iIndex, 1)), 1) ; convert size in bytes to the short text form, after sorting
EndFunc   ;==>_handleListViewItemCreated

Func _getType($sPath, $bFolder = False)
	Local $tSHFILEINFO = DllStructCreate($tagSHFILEINFO)
	Local $iAttr = ($bFolder ? $FILE_ATTRIBUTE_DIRECTORY : $FILE_ATTRIBUTE_NORMAL)
	_WinAPI_ShellGetFileInfo($sPath, BitOR($SHGFI_TYPENAME, $SHGFI_USEFILEATTRIBUTES), $iAttr, $tSHFILEINFO)
	Return DllStructGetData($tSHFILEINFO, 5)
EndFunc   ;==>_getType

Func _doubleClickCallback($hSystem, $hView, $sRoot, $sFolder, $sSelected, $item)
	; get filename for currently selected ListView item
	Local $Array = _GUICtrlListView_GetItemTextArray($idListview, -1)
	; ensure that the array contains information and that filename is not blank
	If $Array[0] <> 0 And $Array[1] <> "" Then
		; open file in ListView when double-clicking (uses Windows defaults per extension)
		ShellExecute($sRoot & $sFolder & $Array[1])
	EndIf
EndFunc   ;==>_doubleClickCallback

Func _loadingCallback($hSystem, $hView, $sRoot, $sFolder, $sSelected, $sPath, $bLoading)
	$bLoadStatus = $bLoading
	If $bLoading Then
		; add delay before changing cursor and clearing status item count
		AdlibRegister("_ListViewLoadWait", 250)
		Return
	EndIf

	If $bCursorOverride Then
		; reset GUI cursor if it has been overridden
		GUISetCursor($MCID_ARROW, 0, $g_hGUI)
		GUISetCursor($MCID_ARROW, 0, _GUIFrame_GetHandle($iFrame_A, 2))
		GUICtrlSetState($idListview, $GUI_SHOW)
		$bCursorOverride = False
	EndIf

	; update statusbar item count
	_PathInputChanged()
EndFunc   ;==>_loadingCallback

Func _ListViewLoadWait()
	If $bLoadStatus Then
		; override GUI with loading/waiting cursor on directories that are slower to load
		$bCursorOverride = True
		GUISetCursor($MCID_WAIT, 1, $g_hGUI)
		GUISetCursor($MCID_WAIT, 1, _GUIFrame_GetHandle($iFrame_A, 2))
		GUICtrlSetState($idListview, $GUI_HIDE)
		; clear statusbar item count
		$g_aText[0] = ""
		_WinAPI_RedrawWindow($g_hStatus)
	EndIf
	AdlibUnRegister("_ListViewLoadWait")
EndFunc   ;==>_ListViewLoadWait

Func _folderCallback($hSystem, $sRoot, $sFolder, $sSelected)
	Local Static $sFolderPrev = ""
	If $sFolder <> $sFolderPrev Then
		Local $arData = [$hSystem, $sFolderPrev, $sFolder]
		If Not $bFolderHistoryChanging Then __History_Add($hFolderHistory, $arData)
		$sFolderPrev = $sFolder
	EndIf
	GUICtrlSetState($sUpLevel, $sFolder <> "" ? $GUI_ENABLE : $GUI_DISABLE)
EndFunc   ;==>_folderCallback

Func _doUnReDo($hHistory, $bRedo, $arData)
	$bFolderHistoryChanging = True
	If $bRedo Then
		__TreeListExplorer_OpenPath($arData[0], $arData[2])
	Else
		__TreeListExplorer_OpenPath($arData[0], $arData[1])
	EndIf
	$bFolderHistoryChanging = False
EndFunc   ;==>_doUnReDo

Func _historyChange($hHistory)
	GUICtrlSetState($sBack, __History_UndoCount($hHistory) > 0 ? $GUI_ENABLE : $GUI_DISABLE)
	GUICtrlSetState($sForward, __History_RedoCount($hHistory) > 0 ? $GUI_ENABLE : $GUI_DISABLE)
EndFunc   ;==>_historyChange

Func WM_NOTIFY2($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam
	__TreeListExplorer__WinProc($hWnd, $iMsg, $wParam, $lParam)

	; used for dark mode header text color and header and listview combined functionality
	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)

	; keep previous listview item row (related to tooltips)
	Local Static $iItemPrev
	Local $iItemRow

	Local $tText

	; header and listview combined functionality
	Local $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	Local $iCode = DllStructGetData($tNMHDR, "Code")
	Switch $hWndFrom
		Case $g_hHeader
			Switch $iCode
				Case $NM_CUSTOMDRAW
					; === HEADER LOGIC ===
					If $isDarkMode Then
						Local $tNMCustomDraw = DllStructCreate($tagNMCUSTOMDRAW, $lParam)

						; STEP 1 (CDDS_PREPAINT): Request item notifications.
						If $tNMCustomDraw.DrawStage = $CDDS_PREPAINT Then
							Return $CDRF_NOTIFYITEMDRAW
						EndIf

						; STEP 2 (CDDS_ITEMPREPAINT): Set text color.
						If BitAND($tNMCustomDraw.DrawStage, $CDDS_ITEMPREPAINT) Then
							; Set the text color to white
							_WinAPI_SetTextColor_mod($tNMCustomDraw.hdc, $CLR_WHITE)

							; Tell the system to use the new font settings and continue.
							Return $CDRF_NEWFONT
						EndIf
					EndIf

				Case $HDN_ITEMCHANGED, $HDN_ITEMCHANGEDW
					Local $tNMHEADER = DllStructCreate($tagNMHEADER, $lParam)
					Local $iHeaderItem = $tNMHEADER.Item
					Local $tHDITEM = DllStructCreate($tagHDITEM, $tNMHEADER.pItem)
					Local $iHeaderItemWidth = $tHDITEM.XY
					_GUICtrlListView_SetColumnWidth($g_hListview, $iHeaderItem, $iHeaderItemWidth)
					_resizeLVCols2()
					Return False                     ; to continue tracking the divider

				Case $HDN_ENDDRAG
					AdlibRegister("_EndDrag", 10)
					Return False                     ; to allow the control to automatically place and reorder the item

				Case $HDN_DIVIDERDBLCLICK, $HDN_DIVIDERDBLCLICKW
					; "Notifies a header control's parent window that the user double-clicked the divider area of the control." (msdn)
					; Let's use it to auto-size the corresponding listview column.
					Local $tNMHEADER = DllStructCreate($tagNMHEADER, $lParam)
					Local $iHeaderItem = $tNMHEADER.Item
					_GUICtrlListView_SetColumnWidth($g_hListview, $iHeaderItem, $LVSCW_AUTOSIZE)
					_GUICtrlHeader_SetItemWidth($g_hHeader, $iHeaderItem, _GUICtrlListView_GetColumnWidth($g_hListview, $iHeaderItem))
					_resizeLVCols2()
					_WinAPI_RedrawWindow($g_hHeader)

				Case $HDN_ITEMCLICK, $HDN_ITEMCLICKW
					; "Notifies a header control's parent window that the user clicked the control." (msdn)
					; Let's use it to sort the corresponding listview column.
					Local $tNMHEADER = DllStructCreate($tagNMHEADER, $lParam)
					Local $iHeaderItem = $tNMHEADER.Item
					Local $iCol = $iHeaderItem
					Local $hHeader = $g_hHeader
					Local $hView = $g_hListview
					For $i = 0 To _GUICtrlHeader_GetItemCount($hHeader) - 1
						If $i = $iCol Then ContinueLoop
						_GUICtrlHeader_SetItemFormat($hHeader, $i, BitAND(_GUICtrlHeader_GetItemFormat($hHeader, $i), BitNOT(BitOR($HDF_SORTDOWN, $HDF_SORTUP))))
					Next
					Local $iFormat = _GUICtrlHeader_GetItemFormat($hHeader, $iCol)
					$__TreeListExplorer__Data["mViews"][$hView]["mSorting"]["iCol"] = $iCol
					If BitAND($iFormat, $HDF_SORTUP) Then ; ascending
						_GUICtrlHeader_SetItemFormat($hHeader, $iCol, BitOR(BitXOR($iFormat, $HDF_SORTUP), $HDF_SORTDOWN))
						$__TreeListExplorer__Data["mViews"][$hView]["mSorting"]["iDir"] = 1
					Else ; descending
						_GUICtrlHeader_SetItemFormat($hHeader, $iCol, BitOR(BitXOR($iFormat, $HDF_SORTDOWN), $HDF_SORTUP))
						$__TreeListExplorer__Data["mViews"][$hView]["mSorting"]["iDir"] = 0
					EndIf
					__TreeListExplorer__UpdateView($hView, True)
			EndSwitch

		Case $g_hListview
			Switch $iCode
				Case $LVN_ENDSCROLL
					Local Static $tagNMLVSCROLL = $tagNMHDR & ";int dx;int dy"
					Local $tNMLVSCROLL = DllStructCreate($tagNMLVSCROLL, $lParam)
					If $tNMLVSCROLL.dy = 0 Then                     ; ListView horizontal scrolling
						_resizeLVCols2()
					EndIf
					If $tNMLVSCROLL.dx = 0 Then
						; ensure that tooltip doen't show during vertical scrolling
						_GUIToolTip_TrackActivate($hToolTip1, False, $g_hGUI, $g_hListview)
						; reset the value stored in the tooltip
						$gText = ""
						_GUIToolTip_UpdateTipText($hToolTip1, $g_hGUI, $g_hListview, $gText)
					EndIf
				Case $LVN_ITEMCHANGED
					; follow up in main While loop
					$bSelectChanged = True
					; item selection(s) have changed
					_selectionChangedLV()
				Case $LVN_BEGINDRAG, $LVN_BEGINRDRAG
					Local $tNMListView = DllStructCreate($tagNMLISTVIEW, $lParam)
					$hTreeItemOrig = _GUICtrlTreeView_GetSelection($g_hTreeView)

					; create array with list of selected listview items
					Local $aItems = _GUICtrlListView_GetSelectedIndices($tNMHDR.hwndFrom, True)
					For $i = 1 To $aItems[0]
						$aItems[$i] = __TreeListExplorer_GetPath($hTLESystem) & _GUICtrlListView_GetItemText($tNMHDR.hwndFrom, $aItems[$i])
					Next

					Local $pDataObj, $pDropSource
					;$pDataObj = GetDataObjectOfFiles($hWnd, $aItems) ; MattyD function

					_ArrayDelete($aItems, 0) ; only needed for GetDataObjectOfFile_B
					$pDataObj = GetDataObjectOfFile_B($aItems) ; jugador function

					;Create an IDropSource to handle our end of the drag/drop operation.
					$pDropSource = CreateDropSource()

					Local $iResult = _SHDoDragDrop($pDataObj, $pDropSource, BitOR($DROPEFFECT_MOVE, $DROPEFFECT_COPY, $DROPEFFECT_LINK))
					;__TreeListExplorer_Reload($hTLESystem)

					DestroyDropSource($pDropSource)
					_Release($pDataObj)

					; allow Undo if drop returns successful
					If $iResult = $DRAGDROP_S_DROP Then _AllowUndo()

					; there is not supposed to be a Return value on LVN_BEGINDRAG
					; however it fixes an issue with built-in drag-drop mechanism
					Return 0
				Case $LVN_HOTTRACK                ; Sent by a list-view control When the user moves the mouse over an item
					Local $tInfo2 = DllStructCreate($tagNMLISTVIEW, $lParam)
					$gText = _GUICtrlListView_GetItemText($hWndFrom, DllStructGetData($tInfo2, "Item"), 0)
					$iItemRow = DllStructGetData($tInfo2, "Item")
					; clear tooltip if cursor not over column 0 or different item
					If DllStructGetData($tInfo2, "SubItem") <> 0 Or $iItemRow <> $iItemPrev Then
						; ensure that tooltip only shows when over column 0
						_GUIToolTip_TrackActivate($hToolTip1, False, $g_hGUI, $g_hListview)
						; reset the value stored in the tooltip
						$gText = ""
						_GUIToolTip_UpdateTipText($hToolTip1, $g_hGUI, $g_hListview, $gText)
					EndIf
					$iItemPrev = $iItemRow
					Return 0                    ; Allow the ListView to perform its normal track select processing.
				Case $NM_HOVER                  ; Sent by a list-view control when the mouse hovers over an item
					; need to determine if file or folder to get more details
					If $gText <> "" Then
						$gText = __TreeListExplorer_GetPath($hTLESystem) & $gText & @CRLF
						Local $gText2 = __TreeListExplorer_GetPath($hTLESystem) & _GUICtrlListView_GetItemText($g_hListview, _GUICtrlListView_GetHotItem($g_hListview), 0)
						; is selected path a folder
						If StringInStr(FileGetAttrib($gText2), "D") Then
							If _WinAPI_PathIsRoot_mod($gText2) Then
								; update drive space information
								Local $iDriveFree = Round(DriveSpaceFree($gText2) / 1024, 1)
								Local $iDriveTotal = Round(DriveSpaceTotal($gText2) / 1024, 1)
								Local $iPercentFree = Round(($iDriveFree / $iDriveTotal) * 100)
								Local $sDriveInfo = $iDriveFree & " GB free" & " (" & $iPercentFree & "%)"
								$gText = $sDriveInfo
							Else
								$gText &= "Size: " & __TreeListExplorer__GetSizeString(DirGetSize($gText2)) & @CRLF
								$gText &= "Date Modified: " & _GUICtrlListView_GetItemText($g_hListview, _GUICtrlListView_GetHotItem($g_hListview), 2)
							EndIf
						Else
							$gText &= "Size: " & _GUICtrlListView_GetItemText($g_hListview, _GUICtrlListView_GetHotItem($g_hListview), 1) & @CRLF
							$gText &= "Date Modified: " & _GUICtrlListView_GetItemText($g_hListview, _GUICtrlListView_GetHotItem($g_hListview), 2)
						EndIf
					EndIf
					$aPosTip = MouseGetPos()
					If $aPosTip[0] <> $iOldaPos0 Or $aPosTip[1] <> $iOldaPos1 Then
						_GUIToolTip_TrackPosition($hToolTip1, $aPosTip[0] + 10, $aPosTip[1] + 20)
						_GUIToolTip_UpdateTipText($hToolTip1, $g_hGUI, $g_hListview, $gText)
						$iOldaPos0 = $aPosTip[0]
						$iOldaPos1 = $aPosTip[1]
					EndIf
					If $gText = "" Then
						_GUIToolTip_TrackActivate($hToolTip1, False, $g_hGUI, $g_hListview)
					Else
						_GUIToolTip_TrackActivate($hToolTip1, True, $g_hGUI, $g_hListview)
						; check if cursor is over listview (if not, clear tooltip)
						$bTooltipActive = True
					EndIf
					Return 1                     ; prevent the hover from being processed
				Case $LVN_KEYDOWN
					Local $tLVKeyDown = DllStructCreate($tagNMLVKEYDOWN, $lParam)
					Local $iVKey = DllStructGetData($tLVKeyDown, "VKey")
					If $iVKey = 46 Then
						; create array with list of selected listview items
						Local $aItems = _GUICtrlListView_GetSelectedIndices($tNMHDR.hwndFrom, True)
						For $i = 1 To $aItems[0]
							$aItems[$i] = __TreeListExplorer_GetPath($hTLESystem) & _GUICtrlListView_GetItemText($tNMHDR.hwndFrom, $aItems[$i])
						Next

						;$pDataObj = GetDataObjectOfFiles($hWnd, $aItems) ; MattyD function

						_ArrayDelete($aItems, 0) ; only needed for GetDataObjectOfFile_B
						Local $pDataObj = GetDataObjectOfFile_B($aItems) ; jugador function

						Local $iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
						_IFileOperationDelete($pDataObj, $iFlags)

						__TreeListExplorer_Reload($hTLESystem)

						_Release($pDataObj)
						_AllowUndo()
					EndIf
				Case $LVN_BEGINLABELEDITA, $LVN_BEGINLABELEDITW
					Local $aSelectedLV = _GUICtrlListView_GetSelectedIndices($idListview, True)
					; there should only be one selected item during a rename
					Local $iItemLV = $aSelectedLV[1]
					Local $sRenameItem = _GUICtrlListView_GetItemText($idListview, $iItemLV, 0)
					$sRenameFrom = __TreeListExplorer_GetPath($hTLESystem) & $sRenameItem
					; set hotkeys to ensure that file name cannot contain illegal characters
					; \ / : * ? " < > |
					HotKeySet ('{\}', "_RenameCheckLV")
					HotKeySet ('{/}', "_RenameCheckLV")
					HotKeySet ('{:}', "_RenameCheckLV")
					HotKeySet ('{*}', "_RenameCheckLV")
					HotKeySet ('{?}', "_RenameCheckLV")
					HotKeySet ('{"}', "_RenameCheckLV")
					HotKeySet ('{<}', "_RenameCheckLV")
					HotKeySet ('{>}', "_RenameCheckLV")
					HotKeySet ('{|}', "_RenameCheckLV")
					Return False
				Case $LVN_ENDLABELEDITA, $LVN_ENDLABELEDITW
					; unset hotkeys that block illegal characters from being set
					HotKeySet ('{\}')
					HotKeySet ('{/}')
					HotKeySet ('{:}')
					HotKeySet ('{*}')
					HotKeySet ('{?}')
					HotKeySet ('{"}')
					HotKeySet ('{<}')
					HotKeySet ('{>}')
					HotKeySet ('{|}')
					Local $sRenameTo
                    $tText = DllStructCreate($tagNMLVDISPINFO, $lParam)
                    Local $tBuffer = DllStructCreate("wchar Text[" & DllStructGetData($tText, "TextMax") & "]", DllStructGetData($tText, "Text"))
					Local $sTextRet = DllStructGetData($tBuffer, "Text")
                    Local $sIllegal = "A file name can't contain any of the following characters:" & @CRLF & @CRLF
                    $sIllegal &= '\ / : * ? " < > |'
                    ;   A file name can't contain any of the following characters:
                    ;   \/:*?"<>|
                    Select
                        Case StringInStr($sTextRet, '\', 2)
                            Return False
                        Case StringInStr($sTextRet, '/', 2)
                            Return False
                        Case StringInStr($sTextRet, ':', 2)
                            Return False
                        Case StringInStr($sTextRet, '*', 2)
                            Return False
                        Case StringInStr($sTextRet, '?', 2)
                            Return False
                        Case StringInStr($sTextRet, '"', 2)
                            Return False
                        Case StringInStr($sTextRet, '<', 2)
                            Return False
                        Case StringInStr($sTextRet, '>', 2)
                            Return False
                        Case StringInStr($sTextRet, '|', 2)
                            Return False
                        Case Not $sTextRet
                            Return False
                        Case Else
							$sRenameTo = __TreeListExplorer_GetPath($hTLESystem) & $sTextRet
							_WinAPI_ShellFileOperation($sRenameFrom, $sRenameTo, $FO_RENAME, BitOR($FOF_ALLOWUNDO, $FOF_NO_UI))
							; refresh TLE system to pick up any folder changes, file type changes, etc.
							__TreeListExplorer_Reload($hTLESystem)
							_AllowUndo()
                            Return True     ; allow rename to occur
                    EndSelect
				Case $NM_SETFOCUS
					$sControlFocus = 'List'
					$bFocusChanged = True
				Case $NM_KILLFOCUS
					$sControlFocus = ''
					$bFocusChanged = True
			EndSwitch
		Case $g_hTreeView
			Switch $iCode
				Case $TVN_BEGINDRAGW, $TVN_BEGINRDRAGW
					Local $tTree = DllStructCreate($tagNMTREEVIEW, $lParam)
					Local $hDragItem = DllStructGetData($tTree, "NewhItem")
					$hTreeItemOrig = _GUICtrlTreeView_GetSelection($g_hTreeView)

					Local $sItemText = TreeItemToPath($g_hTreeView, $hDragItem)

					Local $pDataObj, $pDropSource

					;Get an IDataObject representing the file to copy
					$pDataObj = GetDataObjectOfFile($hWnd, $sItemText)
					If Not @error Then

						;Create an IDropSource to handle our end of the drag/drop operation.
						$pDropSource = CreateDropSource()

						If Not @error Then
							Local $iResult = _SHDoDragDrop($pDataObj, $pDropSource,  BitOR($DROPEFFECT_MOVE, $DROPEFFECT_COPY, $DROPEFFECT_LINK))

							;Operation done, destroy our drop source. (Can't just IUnknown_Release() this one!)
							DestroyDropSource($pDropSource)
						EndIf

						;Relase the data object so the system can destroy it (prevent memory leaks)
						_Release($pDataObj)

						; allow Undo if drop returns successful
						If $iResult = $DRAGDROP_S_DROP Then _AllowUndo()
					EndIf
				Case $TVN_KEYDOWN
					Local $tTVKeyDown = DllStructCreate($tagNMTVKEYDOWN, $lParam)
					Local $iVKey = DllStructGetData($tTVKeyDown, "VKey")
					If $iVKey = 46 Then
						Local $hTreeItemSel = _GUICtrlTreeView_GetSelection($g_hTreeView)
						Local $sItemText = TreeItemToPath($g_hTreeView, $hTreeItemSel)

						Local $pDataObj, $pDropSource

						;Get an IDataObject representing the file to copy
						$pDataObj = GetDataObjectOfFile($hWnd, $sItemText)
						$iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
						_IFileOperationDelete($pDataObj, $iFlags)

						__TreeListExplorer_Reload($hTLESystem)

							;Relase the data object so the system can destroy it (prevent memory leaks)
						_Release($pDataObj)
						_AllowUndo()
					EndIf
				Case $TVN_BEGINLABELEDITA, $TVN_BEGINLABELEDITW
					HotKeySet("{Enter}", "_SaveEditTV")
					HotKeySet("{Esc}", "_CancelEditTV")
					$hTreeItemOrig = _GUICtrlTreeView_GetSelection($g_hTreeView)
					$sRenameFrom = TreeItemToPath($g_hTreeView, $hTreeItemOrig)
					; set hotkeys to ensure that file name cannot contain illegal characters
					; \ / : * ? " < > |
					HotKeySet ('{\}', "_RenameCheckTV")
					HotKeySet ('{/}', "_RenameCheckTV")
					HotKeySet ('{:}', "_RenameCheckTV")
					HotKeySet ('{*}', "_RenameCheckTV")
					HotKeySet ('{?}', "_RenameCheckTV")
					HotKeySet ('{"}', "_RenameCheckTV")
					HotKeySet ('{<}', "_RenameCheckTV")
					HotKeySet ('{>}', "_RenameCheckTV")
					HotKeySet ('{|}', "_RenameCheckTV")
					Return False
				Case $TVN_ENDLABELEDITA, $TVN_ENDLABELEDITW
					Local $sRenameTo
					HotKeySet("{Enter}")
					HotKeySet("{Esc}")
					; unset hotkeys that block illegal characters from being set
					HotKeySet ('{\}')
					HotKeySet ('{/}')
					HotKeySet ('{:}')
					HotKeySet ('{*}')
					HotKeySet ('{?}')
					HotKeySet ('{"}')
					HotKeySet ('{<}')
					HotKeySet ('{>}')
					HotKeySet ('{|}')
					If $bSaveEdit Then
						$bSaveEdit = False
						$tText = DllStructCreate($tagNMTVDISPINFO, $lParam)
						Local $tBuffer = DllStructCreate("wchar Text[" & DllStructGetData($tText, "TextMax") & "]", DllStructGetData($tText, "Text"))
						Local $sTextRet = DllStructGetData($tBuffer, "Text")
						;   A file name can't contain any of the following characters:
						;   \/:*?"<>|
						Select
							Case StringInStr($sTextRet, '\', 2)
								Return False
							Case StringInStr($sTextRet, '/', 2)
								Return False
							Case StringInStr($sTextRet, ':', 2)
								Return False
							Case StringInStr($sTextRet, '*', 2)
								Return False
							Case StringInStr($sTextRet, '?', 2)
								Return False
							Case StringInStr($sTextRet, '"', 2)
								Return False
							Case StringInStr($sTextRet, '<', 2)
								Return False
							Case StringInStr($sTextRet, '>', 2)
								Return False
							Case StringInStr($sTextRet, '|', 2)
								Return False
							Case Not $sTextRet
								Return False
							Case Else
								Local $aPath = _StringBetween($sRenameFrom, "\", "\")
								Local $sRenameItem = $aPath[UBound($aPath) - 1]
								$sRenameTo = StringReplace($sRenameFrom, $sRenameItem, $sTextRet)
								_WinAPI_ShellFileOperation($sRenameFrom, $sRenameTo, $FO_RENAME, BitOR($FOF_ALLOWUNDO, $FOF_NO_UI))
								;__TreeListExplorer_Reload($hTLESystem)
								_AllowUndo()
								Return True     ; allow rename to occur
                    	EndSelect
					EndIf
				Case $NM_SETFOCUS
					$sControlFocus = 'Tree'
					$bFocusChanged = True
				Case $NM_KILLFOCUS
					$sControlFocus = ''
					$bFocusChanged = True
				Case $TVN_SELCHANGINGA, $TVN_SELCHANGINGW
					; follow up in main While loop
					$bSelectChanged = True
			EndSwitch
	EndSwitch

	; custom draw buttons
	Local $tInfo = DllStructCreate($tagNMCUSTOMDRAW, $lParam)
	If _WinAPI_GetClassName_mod($tInfo.hWndFrom) = "Button" And IsString(GUICtrlRead($tInfo.IDFrom)) And $tInfo.Code = $NM_CUSTOMDRAW Then
		Local $tRECT = DllStructCreate($tagRECT, DllStructGetPtr($tInfo, "left"))
		Switch $tInfo.DrawStage
			Case $CDDS_PREPAINT
				If BitAND($tInfo.ItemState, $CDIS_HOT) Then
					; set hot track back color
					$hBrush = _WinAPI_CreateSolidBrush_mod($iBackColorHot)
				EndIf
				If BitAND($tInfo.ItemState, $CDIS_SELECTED) Then
					; set selected back color
					$hBrush = _WinAPI_CreateSolidBrush_mod($iBackColorSel)
				EndIf
				If BitAND($tInfo.ItemState, $CDIS_DISABLED) Then
					; set disabled back color
					$hBrush = _WinAPI_CreateSolidBrush_mod($iBackColorDis)
				EndIf
				If Not BitAND($tInfo.ItemState, $CDIS_HOT) And Not BitAND($tInfo.ItemState, $CDIS_SELECTED) And Not BitAND($tInfo.ItemState, $CDIS_DISABLED) Then
					$hBrush = _WinAPI_CreateSolidBrush_mod($iBackColorDef)
				EndIf
				_WinAPI_FillRect_mod($tInfo.hDC, $tRECT, $hBrush)
				_WinAPI_DeleteObject_mod($hBrush)
				Return $CDRF_NOTIFYPOSTPAINT
			Case $CDDS_POSTPAINT
				_WinAPI_InflateRect_mod($tRECT, -4, -6)
				If BitAND($tInfo.ItemState, $CDIS_HOT) Then
					; set hot track back color
					_WinAPI_SetBkColor_mod($tInfo.hDC, $iBackColorHot)
					; set default text color
					_WinAPI_SetTextColor_mod($tInfo.hDC, $iTextColorDef)
				EndIf
				If BitAND($tInfo.ItemState, $CDIS_SELECTED) Then
					; set selected back color
					_WinAPI_SetBkColor_mod($tInfo.hDC, $iBackColorSel)
					; set default text color
					_WinAPI_SetTextColor_mod($tInfo.hDC, $iTextColorDef)
				EndIf
				If BitAND($tInfo.ItemState, $CDIS_DISABLED) Then
					; set disabled back color
					_WinAPI_SetBkColor_mod($tInfo.hDC, $iBackColorDis)
					; set disabled text color
					_WinAPI_SetTextColor_mod($tInfo.hDC, $iTextColorDis)
				EndIf
				If Not BitAND($tInfo.ItemState, $CDIS_HOT) And Not BitAND($tInfo.ItemState, $CDIS_SELECTED) And Not BitAND($tInfo.ItemState, $CDIS_DISABLED) Then
					; set default back color
					_WinAPI_SetBkColor_mod($tInfo.hDC, $iBackColorDef)
					; set default text color
					_WinAPI_SetTextColor_mod($tInfo.hDC, $iTextColorDef)
				EndIf

				_WinAPI_DrawText_mod($tInfo.hDC, GUICtrlRead($tInfo.IDFrom), $tRECT, BitOR($DT_CENTER, $DT_VCENTER))
		EndSwitch
	EndIf

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY2

Func _RenameCheckLV()
	Local $sIllegal = "A file name can't contain any of the following characters:" & @CRLF & @CRLF
    $sIllegal &= '\ / : * ? " < > |'
	Local $hEdit = _GUICtrlListView_GetEditControl($g_hListview)
	_GUICtrlEdit_ShowBalloonTip($hEdit, '', $sIllegal, $TTI_INFO)
EndFunc

Func _RenameCheckTV()
	Local $sIllegal = "A file name can't contain any of the following characters:" & @CRLF & @CRLF
    $sIllegal &= '\ / : * ? " < > |'
	Local $hEdit = _GUICtrlTreeView_GetEditControl($g_hTreeView)
	_GUICtrlEdit_ShowBalloonTip($hEdit, '', $sIllegal, $TTI_INFO)
EndFunc

Func _SaveEditTV()
	$bSaveEdit = True
	_GUICtrlTreeView_EndEdit($g_hTreeView)
EndFunc

Func _CancelEditTV()
	$bSaveEdit = False
	_GUICtrlTreeView_EndEdit($g_hTreeView)
EndFunc

Func _removeExStyles()
	; remove WS_EX_COMPOSITED from GUI
	Local $i_ExStyle_Old = _WinAPI_GetWindowLong_mod($hParentFrame, $GWL_EXSTYLE)
	_WinAPI_SetWindowLong_mod($hParentFrame, $GWL_EXSTYLE, BitXOR($i_ExStyle_Old, $WS_EX_COMPOSITED))

	; add LVS_EX_DOUBLEBUFFER to ListView
	_GUICtrlListView_SetExtendedListViewStyle($g_hListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_DOUBLEBUFFER, $LVS_EX_TRACKSELECT))
EndFunc   ;==>_removeExStyles

Func _addExStyles()
	Local Static $hStyleTimer
	; Return if already triggered and timer is under 750ms
	If $hStyleTimer And _Timer_Diff_mod($hStyleTimer) < 750 Then Return

	; add WS_EX_COMPOSITED to GUI
	Local $i_ExStyle_Old = _WinAPI_GetWindowLong_mod($hParentFrame, $GWL_EXSTYLE)
	_WinAPI_SetWindowLong_mod($hParentFrame, $GWL_EXSTYLE, BitOR($i_ExStyle_Old, $WS_EX_COMPOSITED))

	; remove LVS_EX_DOUBLEBUFFER from ListView
	_GUICtrlListView_SetExtendedListViewStyle($g_hListview, $LVS_EX_FULLROWSELECT, $LVS_EX_TRACKSELECT)

	; Set trigger status and timer
	$hStyleTimer = _Timer_Init_mod()

	; set adlib to reset those values
	AdlibRegister("_resetExStylesAdlib", 1000)
EndFunc   ;==>_addExStyles

Func _resetExStylesAdlib()
	; remove WS_EX_COMPOSITED from GUI
	Local $i_ExStyle_Old = _WinAPI_GetWindowLong_mod($hParentFrame, $GWL_EXSTYLE)
	_WinAPI_SetWindowLong_mod($hParentFrame, $GWL_EXSTYLE, BitXOR($i_ExStyle_Old, $WS_EX_COMPOSITED))

	; add LVS_EX_DOUBLEBUFFER to ListView
	_GUICtrlListView_SetExtendedListViewStyle($g_hListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_DOUBLEBUFFER, $LVS_EX_TRACKSELECT))

	; add composited to treeview frame
	Local $i_ExStyle_Old = _WinAPI_GetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 1), $GWL_EXSTYLE)
	_WinAPI_SetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 1), $GWL_EXSTYLE, BitOR($i_ExStyle_Old, $WS_EX_COMPOSITED))

	; unregister adlib
	AdlibUnRegister("_resetExStylesAdlib")
EndFunc   ;==>_resetExStylesAdlib

;==============================================
Func _resizeLVCols() ; resize listview columns to match header widths (1st display only, before any horizontal scrolling)
	For $i = 0 To _GUICtrlHeader_GetItemCount($g_hHeader) - 1
		_GUICtrlListView_SetColumnWidth($g_hListview, $i, _GUICtrlHeader_GetItemWidth($g_hHeader, $i))
	Next

	; In case column 0 got an icon, retrieve the width of the icon
	Local $aRectLV = _GUICtrlListView_GetItemRect($g_hListview, 0, $LVIR_ICON)     ; bounding rectangle of the icon (if any)
	$g_iIconWidth = $aRectLV[2] - $aRectLV[0]     ; without icon : 4 - 4 => 0 (tested, the famous "4" !)
	; with icon of 20 pixels : 24 - 4 = 20
EndFunc   ;==>_resizeLVCols

;==============================================
Func _resizeLVCols2() ; called while a header item is tracked or a divider is double-clicked. Also called while the listview is scrolled horizontally.
	Local $iCol, $aRectLV
	Local $aOrder = _GUICtrlHeader_GetOrderArray($g_hHeader)
	$iCol = $aOrder[1]     ; left column (may not be column 0, if column 0 was dragged/dropped elsewhere)
	If $iCol > 0 Then     ; LV subitem
		$aRectLV = _GUICtrlListView_GetSubItemRect($g_hListview, 0, $iCol)
	Else     ; column 0 needs _GUICtrlListView_GetItemRect()
		$aRectLV = _GUICtrlListView_GetItemRect($g_hListview, 0, $LVIR_LABEL)         ; bounding rectangle of the item text
		$aRectLV[0] -= (4 + $g_iIconWidth)         ; adjust LV col 0 left coord (+++)
	EndIf
	If $aRectLV[0] < 0 Then     ; horizontal scrollbar is NOT at left => move and resize the detached header (mimic a normal listview)
		WinMove($g_hHeader, "", $aRectLV[0], 0, WinGetPos($g_hChild)[2] - $aRectLV[0], Default)
	Else     ; horizontal scrollbar is at left => move and resize the detached header to its initial coords & size
		WinMove($g_hHeader, "", 0, 0, WinGetPos($g_hChild)[2], Default)
	EndIf
EndFunc   ;==>_resizeLVCols2

;==============================================
Func _reorderLVCols()
	; remove LVS_NOCOLUMNHEADER from listview
	Local $i_Style_Old = _WinAPI_GetWindowLong_mod($g_hListview, $GWL_STYLE)
	_WinAPI_SetWindowLong_mod($g_hListview, $GWL_STYLE, BitXOR($i_Style_Old, $LVS_NOCOLUMNHEADER))

	; reorder listview columns order to match header items order
	Local $aOrder = _GUICtrlHeader_GetOrderArray($g_hHeader)
	_GUICtrlListView_SetColumnOrderArray($g_hListview, $aOrder)

	; add LVS_NOCOLUMNHEADER back to listview
	_WinAPI_SetWindowLong_mod($g_hListview, $GWL_STYLE, $i_Style_Old)
EndFunc   ;==>_reorderLVCols

; Function for getting HWND from PID
Func _GetHwndFromPID($PID)
	$hWnd = 0
	$winlist = WinList()
	Do
		For $i = 1 To $winlist[0][0]
			If $winlist[$i][0] <> "" Then
				$iPID2 = WinGetProcess($winlist[$i][1])
				If $iPID2 = $PID Then
					$hWnd = $winlist[$i][1]
					ExitLoop
				EndIf
			EndIf
		Next
	Until $hWnd <> 0
	Return $hWnd
EndFunc   ;==>_GetHwndFromPID

Func WM_COMMAND2($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg

	Local $iCode = BitShift($wParam, 16)
	Switch $lParam
		Case GUICtrlGetHandle($idInputPath)
			Switch $iCode
				Case $EN_SETFOCUS
					; select all text in path input box
					AdlibRegister("_PathSelectAll", 10)
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND2

Func _About()
	Local $sMsg
	$sMsg = "Program Version: " & @TAB & $sVersion & @CRLF & @CRLF
	$sMsg &= "TreeListExplorer: " & @TAB & _VersionToString(_UDFGetVersion("../lib/TreeListExplorer.au3")) & @CRLF & @CRLF
	$sMsg &= "Made by: " & @TAB & "AutoIt Community"
	MsgBox(0, "Files Au3", $sMsg)
EndFunc   ;==>_About

Func _CleanExit()
	__TreeListExplorer_Shutdown()
	_GUICtrlHeader_Destroy($g_hHeader)
	_GUIToolTip_Destroy($hToolTip1)
	_GUIToolTip_Destroy($hToolTip2)
	RevokeDragDrop($g_hListview)
	DestroyDropTarget($pLVDropTarget)
	RevokeDragDrop($g_hTreeView)
	DestroyDropTarget($pTVDropTarget)
	GUIDelete($g_hGUI)
	_ClearDarkSizebox()

	DllClose($hKernel32)
	DllClose($hGdi32)
	DllClose($hShlwapi)
EndFunc   ;==>_CleanExit

Func _ClearDarkSizebox()
	_GDIPlus_BitmapDispose($g_hDots)
	_WinAPI_DestroyCursor($hCursor)
	_WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, $g_hOldProc)
	DllCallbackFree($hProc)
	_GDIPlus_Shutdown()
EndFunc   ;==>_ClearDarkSizebox

Func _InitDarkSizebox()
	Local Const $SBS_SIZEBOX = 0x08

	; Create a sizebox window (Scrollbar class) BEFORE creating the StatusBar control
	$g_hSizebox = _WinAPI_CreateWindowEx(0, "Scrollbar", "", $WS_CHILD + $WS_VISIBLE + $SBS_SIZEBOX, _
			0, 0, 0, 0, $g_hGUI)

	; Subclass the sizebox (by changing the window procedure associated with the Scrollbar class)
	$hProc = DllCallbackRegister('ScrollbarProc', 'lresult', 'hwnd;uint;wparam;lparam')
	$g_hOldProc = _WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, DllCallbackGetPtr($hProc))

	$hCursor = _WinAPI_LoadCursor(0, $OCR_SIZENWSE)
	_WinAPI_SetClassLongEx($g_hSizebox, -12, $hCursor)     ; $GCL_HCURSOR = -12

	; Sizebox height with DPI
	$g_iHeight = (16 * $iDPI) + 2
	$g_hDots = CreateDots($g_iHeight, $g_iHeight, 0x00000000 + $iBackColorDef, 0xFF000000 + 0xBFBFBF)
EndFunc   ;==>_InitDarkSizebox

; Resize the status bar when GUI size changes
Func WM_SIZE($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam, $lParam
	; GUIFrame resizing
	_GUIFrame_SIZE_Handler($hWnd, $iMsg, $wParam, $lParam)
	; resize statusbar parts
	Local $aClientSize = WinGetClientSize($g_hGUI)
	Local Static $bIsSizeBoxShown = True
	Local $iParts = $aClientSize[0] / 4
	Local $aParts[4] = [$iParts, $iParts * 2, $iParts * 3, -1]
	_GUICtrlStatusBar_SetParts($g_hStatus, $aParts)
	_GUICtrlStatusBar_Resize($g_hStatus)

	If BitAND(WinGetState($g_hGUI), $WIN_STATE_MAXIMIZED) Then
		_WinAPI_ShowWindow($g_hSizebox, @SW_HIDE)
		$bIsSizeBoxShown = False
	Else
		WinMove($g_hSizebox, "", $aClientSize[0] - $g_iHeight, $aClientSize[1] - $g_iHeight, $g_iHeight, $g_iHeight)
		If Not $bIsSizeBoxShown Then
			_WinAPI_ShowWindow($g_hSizebox, @SW_SHOW)
			$bIsSizeBoxShown = True
		EndIf
	EndIf

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_SIZE

Func _Properties()
	Select
		Case $sControlFocus = 'List'
			Local $aSelectedLV = _GUICtrlListView_GetSelectedIndices($idListview, True)
			If $aSelectedLV[0] = 1 Then
				Local $sSelectedItem = _GUICtrlListView_GetItemText($idListview, $aSelectedLV[1], 0)
				Local $sSelectedLV = __TreeListExplorer_GetPath($hTLESystem) & $sSelectedItem
				_WinAPI_ShellObjectProperties($sSelectedLV)
			ElseIf $aSelectedLV[0] = 0 Then
				_WinAPI_ShellObjectProperties(__TreeListExplorer_GetPath($hTLESystem))
			Else
				;$sSelectedItems
				Local $aFiles = StringSplit($sSelectedItems, "|")
				_ArrayDelete($aFiles, $aFiles[0])
				_ArrayDelete($aFiles, 0)
				_WinAPI_SHMultiFileProperties(__TreeListExplorer_GetPath($hTLESystem), $aFiles)
			EndIf
		Case $sControlFocus = 'Tree'
			Local $hTreeItem = _GUICtrlTreeView_GetSelection($g_hTreeView)
			Local $sItemText = TreeItemToPath($g_hTreeView, $hTreeItem)
			_WinAPI_ShellObjectProperties($sItemText)
	EndSelect
EndFunc   ;==>_Properties

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_SHMultiFileProperties
; Description....: Displays a merged property sheet for a set of files.
; Syntax.........: _WinAPI_SHMultiFileProperties ($sPath, $aNames)
; Parameters.....: $sPath - Sting of the path containing the files.
;                  $aNames - Array of filenames.
; Return values..: Success - 1.
;                  Failure - 0 and sets the @error:
;                  1 - failed to show properties.
;                  2 - failed to create PIDL from path.
;                  3 - Name in $aFiles not found in $sPath.
;                  4 - failed to create IDataObject Interface.
; Author.........: line333
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================

Func _WinAPI_SHMultiFileProperties($sPath, $aNames)
	Local $sNames, $err = 0, $iCount = UBound($aNames), $PIDLChild
	Local $aPIDLAbsolute[$iCount]

	If StringRight($sPath, 1) = '\' Then $sPath = StringTrimRight($sPath, 1)
	$PIDL = _WinAPI_ShellILCreateFromPath($sPath)
	If @error Then Return SetError(2, 0, 0)

	For $i = 0 To $iCount - 1
		$sNames &= 'ptr;'
	Next

	$aPIDL = DllStructCreate($sNames)
	For $i = 0 To $iCount - 1
		$aPIDLAbsolute[$i] = _WinAPI_ShellILCreateFromPath($sPath & '\' & $aNames[$i])
		If @error Then
			$err = 2
			ExitLoop
		Else
			$PIDLChild = _WinAPI_ILFindChild($PIDL, $aPIDLAbsolute[$i])
			If @error Then
				$err = 3
				ExitLoop
			Else
				DllStructSetData($aPIDL, $i + 1, $PIDLChild)
			EndIf
		EndIf
	Next

	If $err = 0 Then
		$pDataObject = _WinAPI_CIDLData_CreateFromIDArray($PIDL, $iCount, $aPIDL)
		If @error Then
			$err = 4
		Else
			$Ret = DllCall($hShell32, 'uint', 'SHMultiFileProperties', 'ptr', DllStructGetData($pDataObject, 1), 'dword', 0)
			If @error Then $err = 1
		EndIf
	EndIf

	_WinAPI_CoTaskMemFree($PIDL)
	For $i = 0 To $iCount - 1
		$PIDL = DllStructGetData($aNames, $i + 1)
		If $PIDL Then _WinAPI_CoTaskMemFree(DllStructGetData($aNames, $i + 1))
		If $aPIDLAbsolute[$i] Then _WinAPI_CoTaskMemFree($aPIDLAbsolute[$i])
	Next

	If $err <> 0 Then Return SetError($err, 0, 0)
	Return 1
EndFunc   ;==>_WinAPI_SHMultiFileProperties

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_ILFindChild
; Description....: Determines whether a specified ITEMIDLIST structure is the child of another ITEMIDLIST structure and returns a pointer to the child's simple ITEMIDLIST.
; Syntax.........: _WinAPI_ILFindChild ($PIDLParent, $PIDLChild)
; Parameters.....: $PIDLParent - A pointer to the parent ITEMIDLIST structure.
;                  $PIDLChild - A pointer to the child ITEMIDLIST structure.
; Return values..: Success - ITEMIDLIST relative to the ITEMIDLIST or the parent.
;                  Failure - 0 and sets the @error flag to non-zero.
; Author.........: line333
; Modified.......:
; Remarks .......: To free the returned PIDL, call the _WinAPI_CoTaskMemFree() function.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================

Func _WinAPI_ILFindChild($PIDLParent, $PIDLChild)
	Local $err = 0
	$Ret = DllCall($hShell32, 'ptr', 'ILFindChild', 'ptr', $PIDLParent, 'ptr', $PIDLChild)
	If @error Or $Ret[0] = 0 Then Return SetError(1, 0, 0)
	Return $Ret[0]
EndFunc   ;==>_WinAPI_ILFindChild

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_CIDLData_CreateFromIDArray
; Description....: Creates a data object with the default vtable pointer.
; Syntax.........: _WinAPI_CIDLData_CreateFromIDArray ($PIDL, $iItems, $aPIDL)
; Parameters.....: $PIDL - A fully qualified IDLIST for the root of the items specified in apidl.
;                  $iItems - Number of items in $aPIDL
;                  $aPIDL - The array of item IDs relative to $PIDL.
; Return values..: Success - IDataObject Interface.
;                  Failure - 0 and sets the @error flag to non-zero.
; Author.........: line333
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ IDataObject
; Example .......: Yes
; ===============================================================================================================================

Func _WinAPI_CIDLData_CreateFromIDArray($PIDL, $iItems, $aPIDL)
	Local $err = 0, $pDataObject = DllStructCreate('ptr')
	DllCall($hShell32, 'uint', 'CIDLData_CreateFromIDArray', 'ptr', $PIDL, 'uint', $iItems, 'ptr', DllStructGetPtr($aPIDL), 'ptr', DllStructGetPtr($pDataObject))
	If @error Then Return SetError(1, 0, 0)
	Return $pDataObject
EndFunc   ;==>_WinAPI_CIDLData_CreateFromIDArray

Func _GUICtrl_SetFont($hWnd, $iHeight = 15, $iWeight = 400, $iFontAtrributes = 0, $sFontName = "Arial")
	; Author: Rasim
	$hFont = _WinAPI_CreateFont($iHeight, 0, 0, 0, $iWeight, BitAND($iFontAtrributes, 2), BitAND($iFontAtrributes, 4), _
			BitAND($iFontAtrributes, 8), $DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, _
			$DEFAULT_QUALITY, 0, $sFontName)

	_SendMessage($hWnd, $WM_SETFONT, $hFont, 1)
EndFunc   ;==>_GUICtrl_SetFont

Func WM_DRAWITEM2($hWnd, $Msg, $wParam, $lParam)
	#forceref $Msg, $wParam, $lParam

	; modernmenuraw
	WM_DRAWITEM($hWnd, $Msg, $wParam, $lParam)

	Local $tDRAWITEMSTRUCT = DllStructCreate("uint CtlType;uint CtlID;uint itemID;uint itemAction;uint itemState;HWND hwndItem;HANDLE hDC;long rcItem[4];ULONG_PTR itemData", $lParam)

	If DllStructGetData($tDRAWITEMSTRUCT, "hwndItem") <> $g_hStatus Then Return $GUI_RUNDEFMSG     ; Only process the statusbar

	Local $itemID = DllStructGetData($tDRAWITEMSTRUCT, "itemID")     ; part number
	Local $hDC = DllStructGetData($tDRAWITEMSTRUCT, "hDC")
	Local $tRECT = DllStructCreate("long left;long top;long right; long bottom", DllStructGetPtr($tDRAWITEMSTRUCT, "rcItem"))
	Local $iTop = DllStructGetData($tRECT, "top")
	Local $iLeft = DllStructGetData($tRECT, "left")
	Local $hBrush

	$hBrush = _WinAPI_CreateSolidBrush($iBackColorDef)     ; Background Color
	_WinAPI_FillRect($hDC, DllStructGetPtr($tRECT), $hBrush)
	_WinAPI_SetTextColor($hDC, $iTextColorDef)     ; Font Color
	_WinAPI_SetBkMode($hDC, $TRANSPARENT)
	DllStructSetData($tRECT, "top", $iTop + 1)
	DllStructSetData($tRECT, "left", $iLeft + 1)
	_WinAPI_DrawText($hDC, $g_aText[$itemID], $tRECT, $DT_LEFT)
	_WinAPI_DeleteObject($hBrush)

	$tDRAWITEMSTRUCT = 0

	_WinAPI_RedrawWindow($g_hSizebox)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_DRAWITEM2

;==============================================
Func ScrollbarProc($hWnd, $iMsg, $wParam, $lParam) ; Andreik

	If $hWnd = $g_hSizebox And $iMsg = $WM_PAINT Then
		Local $tPAINTSTRUCT
		Local $hDC = _WinAPI_BeginPaint($hWnd, $tPAINTSTRUCT)
		Local $iWidth = DllStructGetData($tPAINTSTRUCT, 'rPaint', 3) - DllStructGetData($tPAINTSTRUCT, 'rPaint', 1)
		Local $iHeight = DllStructGetData($tPAINTSTRUCT, 'rPaint', 4) - DllStructGetData($tPAINTSTRUCT, 'rPaint', 2)
		Local $hGraphics = _GDIPlus_GraphicsCreateFromHDC($hDC)
		_GDIPlus_GraphicsDrawImageRect($hGraphics, $g_hDots, 0, 0, $iWidth, $iHeight)
		_GDIPlus_GraphicsDispose($hGraphics)
		_WinAPI_EndPaint($hWnd, $tPAINTSTRUCT)
		Return 0
	EndIf
	Return _WinAPI_CallWindowProc($g_hOldProc, $hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>ScrollbarProc

;==============================================
Func CreateDots($iWidth, $iHeight, $iBackgroundColor, $iDotsColor) ; Andreik
	Local $iDotSize, $iDotSpace, $iDotFrame
	Local $iDPIpct = $iDPI * 100
	Switch $iDPIpct
		;       Dot Size    Spacing     From Right  From Bottom Confirmed (* from aero.msstyles resources)
		; 100%      2           1           2           2           yes
		; 125%      3           1           2           2           yes
		; 150%      3           1           2           2           yes
		; 175%      3           1           2           2           yes
		; 200%      4           2           4           4           yes
		; 250%      5           2           5           5           yes
		; 300%      6           3           6           6           yes
		; 400%      8           4           8           8           yes
		Case 100
			$iDotSize = 2
			$iDotSpace = $iDotSize - 0.5                ; gives some control over the spacing between dots
			$iDotFrame = 0.5                            ; gives some control over the spacing from frame
		Case 125 To 175
			$iDotSize = 3
			$iDotSpace = $iDotSize - 1
			$iDotFrame = 1
		Case 200 To 225
			$iDotSize = 4
			$iDotSpace = $iDotSize - 1
			$iDotFrame = 1
		Case 250 To 275
			$iDotSize = 5
			$iDotSpace = $iDotSize - 1.5
			$iDotFrame = 0
		Case 300 To 375
			$iDotSize = 6
			$iDotSpace = $iDotSize - 1.5
			$iDotFrame = 1
		Case 400 To 500
			$iDotSize = 8
			$iDotSpace = $iDotSize - 2
			$iDotFrame = 2
		Case Else
			$iDotSize = 3
			$iDotSpace = $iDotSize - 1
			$iDotFrame = 1
	EndSwitch

	Local $a[6][2] = [[3, 7], [3, 5], [3, 3], [5, 5], [5, 3], [7, 3]]
	Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iWidth, $iHeight)
	Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
	Local $hBrush = _GDIPlus_BrushCreateSolid($iDotsColor)
	_GDIPlus_GraphicsClear($hGraphics, $iBackgroundColor)
	For $i = 0 To UBound($a) - 1
		_GDIPlus_GraphicsFillRect($hGraphics, $iWidth - ($iDotSpace * $a[$i][0]) + $iDotFrame, $iHeight - ($iDotSpace * $a[$i][1]) + $iDotFrame, $iDotSize, $iDotSize, $hBrush)
	Next
	_GDIPlus_BrushDispose($hBrush)
	_GDIPlus_GraphicsDispose($hGraphics)
	Return $hBitmap
EndFunc   ;==>CreateDots

;==============================================
Func WM_MOVE($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam, $lParam
	Local Static $bSizeboxOffScreen = False

	If $hWnd = $g_hGUI Then
		Local $aPos = WinGetPos($g_hSizebox)
		If $aPos[0] + $aPos[2] > @DesktopWidth Or $aPos[1] + $aPos[3] > @DesktopHeight Then
			$bSizeboxOffScreen = True
		Else
			If $bSizeboxOffScreen Then
				; sizebox was off-screen but is back in range now, redraw GUI
				_WinAPI_RedrawWindow($g_hGUI)
				$bSizeboxOffScreen = False
			EndIf
		EndIf
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_MOVE

Func ApplyDPI()
	; apply System DPI awareness and calculate factor
	; Returns DPI scaling factor (1.0 = 100%), defaults to 1.0 on error
	_WinAPI_SetProcessDpiAwarenessContext($DPI_AWARENESS_CONTEXT_SYSTEM_AWARE)
	If @error Then Return 1

	Local $iDPI2 = Round(_WinAPI_GetDpiForSystem() / 96, 2)
	If @error Then Return 1

	Return $iDPI2
EndFunc   ;==>ApplyDPI

Func _WinAPI_SetProcessDpiAwarenessContext($DPI_AWARENESS_CONTEXT_value) ; UEZ
    Local $aResult = DllCall("user32.dll", "bool", "SetProcessDpiAwarenessContext", @AutoItX64 ? "int64" : "int", $DPI_AWARENESS_CONTEXT_value) ;requires Win10 v1703+ / Windows Server 2016+
    If Not IsArray($aResult) Or @error Then Return SetError(1, @extended, 0)
    If Not $aResult[0] Then Return SetError(2, @extended, 0)
    Return $aResult[0]
EndFunc   ;==>_WinAPI_SetProcessDpiAwarenessContext

Func _WinAPI_GetDpiForSystem() ; UEZ
	Local $aResult = DllCall('user32.dll', "uint", "GetDpiForSystem")     ; requires Win10 v1607+ / no server support
	If Not IsArray($aResult) Or @error Then Return SetError(1, @extended, 0)
	If Not $aResult[0] Then Return SetError(2, @extended, 0)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_GetDpiForSystem

Func _setThemeColors()
	; set theme
	If $isDarkMode = True Then
		_GUISetDarkTheme($g_hGUI)
		_GUISetDarkTheme(_GUIFrame_GetHandle($iFrame_A, 1))
		_GUISetDarkTheme(_GUIFrame_GetHandle($iFrame_A, 2))
		GUICtrlSetBkColor($idListview, $iBackColorDef)
		GUICtrlSetBkColor($idTreeView, $iBackColorDef)
		GUICtrlSetColor($idListview, $iTextColorDef)
		GUICtrlSetColor($idTreeView, $iTextColorDef)
		_WinAPI_SetWindowTheme($g_hListview, 'DarkMode_Explorer')
		_WinAPI_SetWindowTheme($g_hTreeView, 'DarkMode_Explorer')
		_WinAPI_SetWindowTheme($g_hStatus, 'DarkMode', 'ExplorerStatusBar')
		_WinAPI_SetWindowTheme($g_hHeader, 'DarkMode_ItemsView', 'Header')
		If $iOSBuild >= 22621 Then
			GUICtrlSetBkColor($idInputPath, 0x101010)
			If $iOSBuild >= 26100 And $iRevision >= 6899 Then
				$sThemeName = 'DarkMode_DarkTheme'
				_WinAPI_SetWindowTheme($g_hInputPath, 'DarkMode_DarkTheme')
			Else
				$sThemeName = 'DarkMode_Explorer'
				_WinAPI_SetWindowTheme($g_hInputPath, 'DarkMode_Explorer')
			EndIf
		Else
			GUICtrlSetBkColor($idInputPath, 0x303030)
		EndIf
		GUICtrlSetColor($idInputPath, $iTextColorDef)
		GUICtrlSetBkColor($idSeparator, 0x505050)
		$hSolidBrush = _WinAPI_CreateBrushIndirect($BS_SOLID, $iBackColorDef)
		_drawUAHMenuNCBottomLine($g_hGUI)
	Else
		_GUISetDarkTheme($g_hGUI, False)
		_GUISetDarkTheme(_GUIFrame_GetHandle($iFrame_A, 1), False)
		_GUISetDarkTheme(_GUIFrame_GetHandle($iFrame_A, 2), False)
		GUICtrlSetBkColor($idListview, $iBackColorDef)
		GUICtrlSetBkColor($idTreeView, $iBackColorDef)
		GUICtrlSetColor($idListview, $iTextColorDef)
		GUICtrlSetColor($idTreeView, $iTextColorDef)
		_WinAPI_SetWindowTheme($g_hListview, 'Explorer')
		_WinAPI_SetWindowTheme($g_hTreeView, 'Explorer')
		_WinAPI_SetWindowTheme($g_hStatus, 'Explorer')
		_WinAPI_SetWindowTheme($g_hHeader, 'ItemsView', 'Header')
		_WinAPI_SetWindowTheme($g_hInputPath, 'Explorer')
		GUICtrlSetBkColor($idInputPath, 0xFFFFFF)
		GUICtrlSetColor($idInputPath, $iTextColorDef)
		GUICtrlSetBkColor($idSeparator, 0x909090)
		$hSolidBrush = _WinAPI_CreateBrushIndirect($BS_SOLID, $iBackColorDef)
		_drawUAHMenuNCBottomLine($g_hGUI)
	EndIf
EndFunc   ;==>_setThemeColors

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_ShouldAppsUseDarkMode
; Description ...: Checks if apps should use the dark mode.
; Syntax ........: _WinAPI_ShouldAppsUseDarkMode()
; Parameters ....: None
; Return values .: Success: Returns True if apps should use dark mode.
;                  Failure: Returns False and sets @error:
;                           -1: Operating system version is earlier than Windows 10 (version 1809, build 17763).
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: Requires Windows 10 (version 1809, build 17763) or later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_ShouldAppsUseDarkMode()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnShouldAppsUseDarkMode = 132
	Local $aResult = DllCall('uxtheme.dll', 'bool', $fnShouldAppsUseDarkMode)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_ShouldAppsUseDarkMode

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_DwmSetWindowAttribute_unr
; Description ...: Dose the same as _WinAPI_DwmSetWindowAttribute; But has no Restrictions
; Syntax ........: _WinAPI_DwmSetWindowAttribute_unr($hWnd, $iAttribute, $iData)
; Parameters ....: $hWnd                - a handle value.
;                  $iAttribute          - an integer value.
;                  $iData               - an integer value.
; Return values .: Success: 1 Failure: @error, @extended & False
; Author ........: argumentum
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/211475-winapithemeex-darkmode-for-autoits-win32guis/?do=findComment&comment=1530103
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_DwmSetWindowAttribute_unr($hWnd, $iAttribute, $iData) ; #include <WinAPIGdi.au3> ; unthoughtful unrestricting mod.
	Local $aCall = DllCall('dwmapi.dll', 'long', 'DwmSetWindowAttribute', 'hwnd', $hWnd, 'dword', $iAttribute, _
			'dword*', $iData, 'dword', 4)
	If @error Then Return SetError(@error, @extended, 0)
	If $aCall[0] Then Return SetError(10, $aCall[0], 0)
	Return 1
EndFunc   ;==>_WinAPI_DwmSetWindowAttribute_unr

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUISetDarkTheme
; Description ...: Sets the theme for a specified window to either dark or light mode on Windows 10.
; Syntax ........: _GUISetDarkTheme($hwnd, $dark_theme = True)
; Parameters ....: $hwnd          - The handle to the window.
;                  $dark_theme    - If True, sets the dark theme; if False, sets the light theme.
;                                   (Default is True for dark theme.)
; Return values .: None
; Author ........: DK12000, NoNameCode
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/211196-gui-title-bar-dark-theme-an-elegant-solution-using-dwmapi/
; Example .......: No
; ===============================================================================================================================
Func _GUISetDarkTheme($hWnd, $bEnableDarkTheme = True)
	Local Const $APPMODE_FORCEDARK = 2, $APPMODE_FORCELIGHT = 3
	Local $iPreferredAppMode = ($bEnableDarkTheme == True) ? $APPMODE_FORCEDARK : $APPMODE_FORCELIGHT
	Local $iGUI_BkColor = $iBackColorDef
	_WinAPI_SetPreferredAppMode($iPreferredAppMode)
	_WinAPI_RefreshImmersiveColorPolicyState()
	_WinAPI_FlushMenuThemes()
	GUISetBkColor($iGUI_BkColor, $hWnd)
	_WinAPI_DwmSetWindowAttribute_unr($hWnd, $DWMWA_USE_IMMERSIVE_DARK_MODE, $bEnableDarkTheme)
EndFunc   ;==>_GUISetDarkTheme

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_SetPreferredAppMode
; Description ...: Sets the preferred application mode for Windows 10 (version 1903, build 18362) and later.
; Syntax ........: _WinAPI_SetPreferredAppMode($PREFERREDAPPMODE)
; Parameters ....: $PREFERREDAPPMODE - The preferred application mode. See enum PreferredAppMode for possible values.
;                    $APPMODE_DEFAULT (0)
;                    $APPMODE_ALLOWDARK (1)
;                    $APPMODE_FORCEDARK (2)
;                    $APPMODE_FORCELIGHT (3)
;                    $APPMODE_MAX (4)
; Return values .: Success: The PreferredAppMode retuned by the DllCall
;                  Failure: '' and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 18362)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function is applicable for Windows 10 (version 18362) and later.
; Related .......: _WinAPI_AllowDarkModeForApp
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_SetPreferredAppMode($PREFERREDAPPMODE)
	If @OSBuild < 18362 Then Return SetError(-1, 0, False)
	Local $fnSetPreferredAppMode = 135
	Local $aResult = DllCall('uxtheme.dll', 'int', $fnSetPreferredAppMode, 'int', $PREFERREDAPPMODE)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, '')
	Return $aResult[0]
EndFunc   ;==>_WinAPI_SetPreferredAppMode

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_RefreshImmersiveColorPolicyState
; Description ...: Refreshes the system's immersive color policy state, allowing changes to take effect.
; Syntax ........: _WinAPI_RefreshImmersiveColorPolicyState()
; Parameters ....: None
; Return values .: Success: True
;                  Failure: False and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 17763)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function is applicable for Windows 10 (version 17763) and later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_RefreshImmersiveColorPolicyState()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnRefreshImmersiveColorPolicyState = 104
	Local $aResult = DllCall('uxtheme.dll', 'none', $fnRefreshImmersiveColorPolicyState)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, False)
	Return True
EndFunc   ;==>_WinAPI_RefreshImmersiveColorPolicyState

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_FlushMenuThemes
; Description ...: Refreshes the system's immersive color policy state, allowing changes to take effect.
; Syntax ........: _WinAPI_FlushMenuThemes()
; Parameters ....: None
; Return values .: Success: True
;                  Failure: False and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 17763)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function is applicable for Windows 10 (version 17763) and later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_FlushMenuThemes()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnFlushMenuThemes = 136
	Local $aResult = DllCall('uxtheme.dll', 'none', $fnFlushMenuThemes)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, False)
	Return True
EndFunc   ;==>_WinAPI_FlushMenuThemes

Func _EventsGUI()
	Local Static $hTreeItemOrig = 0
	Local Static $bTreeOrigStored = False
	Local $hIcon, $iImgIndex
	Switch @GUI_CtrlId
		Case $GUI_EVENT_CLOSE
			Exit
		Case $GUI_EVENT_MAXIMIZE
			_resizeLVCols2()
		Case $GUI_EVENT_RESIZED
			_resizeLVCols2()
	EndSwitch
EndFunc   ;==>_EventsGUI

Func TreeItemFromPoint($hWnd)
	Local $tMPos = _WinAPI_GetMousePos(True, $hWnd)
	Return _GUICtrlTreeView_HitTestItem($hWnd, DllStructGetData($tMPos, 1), DllStructGetData($tMPos, 2))
EndFunc   ;==>TreeItemFromPoint

Func ListItemFromPoint($hWnd)
	Local $aListItem = _GUICtrlListView_HitTest($hWnd, -1, -1)
	Return $aListItem[0]
EndFunc   ;==>ListItemFromPoint

Func _ButtonFunctions()
	Switch @GUI_CtrlId
		Case $sBack
			__History_Undo($hFolderHistory)
		Case $sForward
			__History_Redo($hFolderHistory)
		Case $sUpLevel
			__TreeListExplorer_OpenPath($hTLESystem, __TreeListExplorer__GetPathAndLast(__TreeListExplorer_GetPath($hTLESystem))[0])
		Case $sRefresh
			__TreeListExplorer_Reload($hTLESystem)
	EndSwitch
EndFunc   ;==>_ButtonFunctions

Func _MenuFunctions()
	Switch @GUI_CtrlId
		Case $idThemeItem
			_switchTheme()
		Case $idExitItem
			Exit
		Case $idPropertiesItem
			_Properties()
		Case $idPropertiesLV
			_Properties()
		Case $idAboutItem
			_About()
		Case $idDeleteItem
			_DeleteItems()
			_AllowUndo()
		Case $idRenameItem
			_RenameItem()
			_AllowUndo()
		Case $idCopyItem
			_CopyItems()
		Case $idPasteItem
			_PasteItems()
			_AllowUndo()
		Case $idUndoItem
			_UndoOp()
		Case $idHiddenItem
			If BitAND(GUICtrlRead($idHiddenItem), $GUI_CHECKED) = $GUI_CHECKED Then
				GUICtrlSetState($idHiddenItem, $GUI_UNCHECKED)
				$bHideHidden = True
			Else
				GUICtrlSetState($idHiddenItem, $GUI_CHECKED)
				$bHideHidden = False
			EndIf

			__TreeListExplorer_ReloadView($idListView, True)
			__TreeListExplorer_ReloadView($idTreeView, True)
		Case $idSystemItem
			If BitAND(GUICtrlRead($idSystemItem), $GUI_CHECKED) = $GUI_CHECKED Then
				GUICtrlSetState($idSystemItem, $GUI_UNCHECKED)
				$bHideSystem = False
			Else
				GUICtrlSetState($idSystemItem, $GUI_CHECKED)
				$bHideSystem = True
			EndIf

			__TreeListExplorer_ReloadView($idListView, True)
			__TreeListExplorer_ReloadView($idTreeView, True)
	EndSwitch
EndFunc   ;==>_MenuFunctions

Func _UndoOp()
	; perform Undo by sending Ctrl+Z to the Desktop (Progman class, SysListView32 class)
    Local Const $hProgman = WinGetHandle("[CLASS:Progman]")
    Local Const $hCurrent = WinGetHandle("[ACTIVE]")

    Local Const $hSHELLDLL_DefView = _WinAPI_FindWindowEx($hProgman, 0, "SHELLDLL_DefView", "")
    Local Const $hSysListView32    = _WinAPI_FindWindowEx($hSHELLDLL_DefView, 0, "SysListView32", "FolderView")

    _WinAPI_SetForegroundWindow($hSysListView32)
    _WinAPI_SetFocus($hSysListView32)

    ControlSend($hSysListView32, "", "", "^z")
    WinActivate($hCurrent)

	__TreeListExplorer_Reload($hTLESystem)

	GUICtrlSetState($idUndoItem, $GUI_DISABLE)
	HotKeySet("^z")
EndFunc

Func _AllowUndo()
	GUICtrlSetState($idUndoItem, $GUI_ENABLE)
	HotKeySet("^z", "_UndoOp")
EndFunc

Func _PasteItems()
	Local $sFullPath
	Select
		Case $sControlFocus = 'List'
			Local $aSelectedLV = _GUICtrlListView_GetSelectedIndices($idListview, True)
			If $aSelectedLV[0] = 0 Then
				; no selections in ListView currently, current path is Paste directory
				Local $sSelectedLV = __TreeListExplorer_GetPath($hTLESystem)
				$sFullPath = $sSelectedLV
			ElseIf $aSelectedLV[0] = 1 Then
				; 1 item selection in ListView
				Local $sSelectedItem = _GUICtrlListView_GetItemText($idListview, $aSelectedLV[1], 0)
				Local $sSelectedLV = __TreeListExplorer_GetPath($hTLESystem) & $sSelectedItem
				; is selected path a folder
				If StringInStr(FileGetAttrib($sSelectedLV), "D") Then
					$sSelectedLV = $sSelectedLV
				Else
					Return
				EndIf
			EndIf
			$iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
			$sAction = "CopyItems"
			_IFileOperationFile($pCopyObj, $sFullPath, $sAction, $iFlags)
			__TreeListExplorer_Reload($hTLESystem)
		Case $sControlFocus = 'Tree'
			Local $hTreeItem = _GUICtrlTreeView_GetSelection($g_hTreeView)
			Local $sItemText = TreeItemToPath($g_hTreeView, $hTreeItem)
			$sFullPath = $sItemText
			$iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
			$sAction = "CopyItems"
			_IFileOperationFile($pCopyObj, $sFullPath, $sAction, $iFlags)
			__TreeListExplorer_Reload($hTLESystem)
	EndSelect
EndFunc

Func _CopyItems()
	Select
		Case $sControlFocus = 'List'
			; if previous copy object exists and user initiates new copy, release previous object
			If $bCopy Then _Release($pCopyObj)

			; create array with list of selected listview items
			Local $aItems = _GUICtrlListView_GetSelectedIndices($g_hListView, True)
			For $i = 1 To $aItems[0]
				$aItems[$i] = __TreeListExplorer_GetPath($hTLESystem) & _GUICtrlListView_GetItemText($g_hListView, $aItems[$i])
			Next

			;Local $pDataObj = GetDataObjectOfFiles($hWnd, $aItems) ; MattyD function

			_ArrayDelete($aItems, 0) ; only needed for GetDataObjectOfFile_B
			$pCopyObj = GetDataObjectOfFile_B($aItems) ; jugador function

			; we don't want to release this until after Paste
			;_Release($pCopyObj)

			; keep track of copy status for menu
			$bCopy = True
		Case $sControlFocus = 'Tree'
			; if previous copy object exists and user initiates new copy, release previous object
			If $bCopy Then _Release($pCopyObj)

			Local $hTreeItem = _GUICtrlTreeView_GetSelection($g_hTreeView)
			Local $sItemText = TreeItemToPath($g_hTreeView, $hTreeItem)

			;Get an IDataObject representing the file to copy
			$pCopyObj = GetDataObjectOfFile(_GUIFrame_GetHandle($iFrame_A, 1), $sItemText)

			; we don't want to release this until after Paste
			;_Release($pCopyObj)

			; keep track of copy status for menu
			$bCopy = True
	EndSelect
EndFunc

Func _RenameItem()
	Select
		Case $sControlFocus = 'List'
			Local $aSelectedLV = _GUICtrlListView_GetSelectedIndices($idListview, True)
			; there should only be one selected item during a rename
			Local $iItemLV = $aSelectedLV[1]
			Local $hEditLabel = _GUICtrlListView_EditLabel($g_hListView, $iItemLV)
		Case $sControlFocus = 'Tree'
			Local $hTreeItem = _GUICtrlTreeView_GetSelection($g_hTreeView)
			_GUICtrlTreeView_EditText($g_hTreeView, $hTreeItem)
	EndSelect
EndFunc

Func _DeleteItems()
	Select
		Case $sControlFocus = 'List'
			; create array with list of selected listview items
			Local $aItems = _GUICtrlListView_GetSelectedIndices($g_hListview, True)
			For $i = 1 To $aItems[0]
				$aItems[$i] = __TreeListExplorer_GetPath($hTLESystem) & _GUICtrlListView_GetItemText($g_hListview, $aItems[$i])
			Next

			;$pDataObj = GetDataObjectOfFiles($hWnd, $aItems) ; MattyD function

			_ArrayDelete($aItems, 0) ; only needed for GetDataObjectOfFile_B
			Local $pDataObj = GetDataObjectOfFile_B($aItems) ; jugador function

			Local $iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
			_IFileOperationDelete($pDataObj, $iFlags)

			__TreeListExplorer_Reload($hTLESystem)

			_Release($pDataObj)
		Case $sControlFocus = 'Tree'
			Local $hTreeItemSel = _GUICtrlTreeView_GetSelection($g_hTreeView)
			Local $sItemText = TreeItemToPath($g_hTreeView, $hTreeItemSel)

			Local $pDataObj, $pDropSource

			;Get an IDataObject representing the file to copy
			$pDataObj = GetDataObjectOfFile(_GUIFrame_GetHandle($iFrame_A, 1), $sItemText)
			$iFlags = BitOR($FOFX_ADDUNDORECORD, $FOFX_RECYCLEONDELETE, $FOFX_NOCOPYHOOKS)
			_IFileOperationDelete($pDataObj, $iFlags)

			__TreeListExplorer_Reload($hTLESystem)

				;Relase the data object so the system can destroy it (prevent memory leaks)
			_Release($pDataObj)
	EndSelect
EndFunc

Func _EndDrag()
	; reorder columns after header drag and drop
	_reorderLVCols()
	AdlibUnRegister("_EndDrag")
EndFunc   ;==>_EndDrag

Func _PathSelectAll()
	; select all text in path input box
	ControlFocus($g_hGUI, "", $idInputPath)
	_GUICtrlEdit_SetSel($g_hInputPath, 0, -1)
	AdlibUnRegister("_PathSelectAll")
EndFunc   ;==>_PathSelectAll

Func _PathInputChanged()
	Local Const $SB_LEFT = 6

	; reset position of header and listview
	GUICtrlSendMsg($idListview, $WM_HSCROLL, $SB_LEFT, 0)
	WinMove($g_hHeader, "", 0, 0, WinGetPos($g_hChild)[2], Default)

	; update number of items (files and folders) in statusbar
	Local $iLVItemCount = _GUICtrlListView_GetItemCount($idListview)
	$g_aText[0] = "  " & $iLVItemCount & " item"
	If $iLVItemCount > 1 Then
		$g_aText[0] &= "s"
	EndIf

	; update drive space information
	Local $iDriveFree = Round(DriveSpaceFree(__TreeListExplorer_GetPath($hTLESystem)) / 1024, 1)
	Local $iDriveTotal = Round(DriveSpaceTotal(__TreeListExplorer_GetPath($hTLESystem)) / 1024, 1)
	Local $iPercentFree = Round(($iDriveFree / $iDriveTotal) * 100)
	$g_aText[3] = "  " & $iDriveFree & " GB free" & " (" & $iPercentFree & "%)"
	_WinAPI_RedrawWindow($g_hStatus)
EndFunc   ;==>_PathInputChanged

Func _drawUAHMenuNCBottomLine($hWnd) ; ahmet
	$rcClient = _WinAPI_GetClientRect($hWnd)

	Local $aCall = DllCall($hUser32, "int", "MapWindowPoints", _
			"hwnd", $hWnd, _         ; hWndFrom
			"hwnd", 0, _             ; hWndTo
			"ptr", DllStructGetPtr($rcClient), _
			"uint", 2)               ; number of points - 2 for RECT structure

	$rcWindow = _WinAPI_GetWindowRect($hWnd)

	_WinAPI_OffsetRect($rcClient, -$rcWindow.left, -$rcWindow.top)

	$rcAnnoyingLine = DllStructCreate($tagRECT)
	$rcAnnoyingLine.left = $rcClient.left
	$rcAnnoyingLine.top = $rcClient.top
	$rcAnnoyingLine.right = $rcClient.right
	$rcAnnoyingLine.bottom = $rcClient.bottom

	$rcAnnoyingLine.bottom = $rcAnnoyingLine.top
	$rcAnnoyingLine.top = $rcAnnoyingLine.top - 1

	$hRgn = _WinAPI_CreateRectRgn(0, 0, 8000, 8000)

	$hDC = _WinAPI_GetDCEx($hWnd, $hRgn, BitOR($DCX_WINDOW, $DCX_INTERSECTRGN))
	_WinAPI_FillRect($hDC, $rcAnnoyingLine, $hSolidBrush)
	_WinAPI_ReleaseDC($hWnd, $hDC)
EndFunc   ;==>_drawUAHMenuNCBottomLine

Func WM_ACTIVATE_Handler($hWnd, $MsgID, $wParam, $lParam) ; ioa747
	_drawUAHMenuNCBottomLine($g_hGUI)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_ACTIVATE_Handler

Func WM_WINDOWPOSCHANGED_Handler($hWnd, $iMsg, $wParam, $lParam)
	If $hWnd <> $g_hGUI Then Return $GUI_RUNDEFMSG
	_drawUAHMenuNCBottomLine($hWnd)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_WINDOWPOSCHANGED_Handler

;Convert @error codes from DllCall into win32 codes.
Func TranslateDllError($iError = @error)
	Switch $iError
		Case 0
			$iError = $ERROR_SUCCESS
		Case 1
			$iError = $ERROR_DLL_INIT_FAILED
		Case Else
			$iError = $ERROR_INVALID_PARAMETER
	EndSwitch
EndFunc

Func CoTaskMemFree($pMemBlock)
	DllCall("Ole32.dll", "none", "CoTaskMemFree", "ptr", $pMemBlock)
EndFunc   ;==>CoTaskMemFree

Func _SHDoDragDrop($pDataObj, $pDropSource, $iOKEffects)
    ;We must pass a IID_IDropSource ptr.
    $pDropSource = _QueryInterface($pDropSource, $sIID_IDropSource)
    _Release($pDropSource)

	;Local $aCall = DllCall($hShell32, "long", "SHDoDragDrop", "hwnd", $g_hGUI, "ptr", $pDataObj, "ptr", $pDropSource, "dword", $iOKEffects, "ptr*", 0)
	Local $aCall = DllCall($hShell32, "long", "SHDoDragDrop", "hwnd", Null, "ptr", $pDataObj, "ptr", Null, "dword", $iOKEffects, "ptr*", 0)

    If @error Then Return SetError(@error, @extended, $aCall)
    Return '0x' & Hex($aCall[0])
EndFunc   ;==>_SHDoDragDrop

Func GetDataObjectOfFile($hWnd, $sPath)
	;Get the path as an idList. This is allocated memory that we should free later on.
	Local $aCall = DllCall("Shell32.dll", "long", "SHParseDisplayName", "wstr", $sPath, "ptr", 0, "ptr*", 0, "ulong", 0, "ulong*", 0)
	Local $iError = @error ? TranslateDllError() : $aCall[0]
	If $iError Then Return SetError($iError, 0, False)
	Local $pIdl = $aCall[3]

	;From the idList, get two things: IShellFolder for the parent folder & the item's IDL relative to the parent folder.
	Local $tIID_IShellFolder = _WinAPI_GUIDFromString($sIID_IShellFolder)
	$aCall = DllCall("Shell32.dll", "long", "SHBindToParent", "ptr", $pIdl, "struct*", $tIID_IShellFolder, "ptr*", 0, "ptr*", 0)
	$iError = @error ? TranslateDllError() : $aCall[0]
	If $iError Then Return SetError($iError, 0, False)
	Local $pShellFolder = $aCall[3]
	;SHBindToParent does not allocate a new PID, so we're not responsible for freeing $pIdlChild.
	Local $tpIdlChild = DllStructCreate("ptr")
	DllStructSetData($tpIdlChild, 1, $aCall[4])
	Local $ppIdlChild = DllStructGetPtr($tpIdlChild)

	;We have an interface tag for IShellFolder, so we can use ObjCreateInterface to "convert" it into an object datatype.
	;$oShellFolder will automatically release when it goes out of scope, so we don't need to manually _Release($pShellFolder).
	Local $oShellFolder = ObjCreateInterface($pShellFolder, $sIID_IShellFolder, $tagIShellFolder)
	Local $pDataObject, $tIID_IDataObject = _WinAPI_GUIDFromString($sIID_IDataObject)
	$iError = $oShellFolder.GetUIObjectOf($hWnd, 1, $ppIdlChild, $tIID_IDataObject, 0, $pDataObject)

	;Free the IDL we initially created
	CoTaskMemFree($pIdl)

	Return SetError($iError, 0, Ptr($pDataObject))
EndFunc   ;==>GetDataObjectOfFile

Func GetDataObjectOfFiles($hWnd, $asPaths)
	;If we use the DesktopFolder object as the parent, children can all be defined by normal file paths.
	;So we don't need to worry about sibling folders to the root etc...

	Local $aCall = DllCall("Shell32.dll", "long", "SHGetDesktopFolder", "ptr*", 0)
	Local $iError = @error ? TranslateDllError() : $aCall[0]
	If $iError Then Return SetError($iError, 0, False)

	Local $pShellFolder = $aCall[1]
	Local $tChildren = DllStructCreate(StringFormat("ptr pIdls[%d]", UBound($asPaths)))

	Local $iEaten, $pChildIDL, $iAttributes
	Local $oShellFolder = ObjCreateInterface($pShellFolder, $sIID_IShellFolder, $tagIShellFolder)

	For $i = 1 To $asPaths[0]
		$oShellFolder.ParseDisplayName($hWnd, 0, $asPaths[$i], $iEaten, $pChildIDL, $iAttributes)
		$tChildren.pIdls(($i)) = $pChildIDL
	Next

	;We have an interface tag for IShellFolder, so we can use ObjCreateInterface to "convert" it into an object datatype.
	;$oShellFolder will automatically release when it goes out of scope, so we don't need to manually _Release($pShellFolder).
	Local $pDataObject, $tIID_IDataObject = _WinAPI_GUIDFromString($sIID_IDataObject)
	$iError = $oShellFolder.GetUIObjectOf($hWnd,  $asPaths[0], DllStructGetPtr($tChildren), $tIID_IDataObject, 0, $pDataObject)

	;Free the IDLs now we have a data object.
	For $i = 1 To $asPaths[0]
		CoTaskMemFree($tChildren.pIdls(($i)) & @CRLF)
	Next

	Return SetError($iError, 0, Ptr($pDataObject))
EndFunc   ;==>GetDataObjectOfFiles

Func RegisterDragDrop($hWnd, $pDropTarget)
	Local $aCall = DllCall("ole32.dll", "long", "RegisterDragDrop", "hwnd", $hWnd, "ptr", $pDropTarget)
	If @error Then Return SetError(TranslateDllError(), 0, False)
	Return SetError($aCall[0], 0, $aCall[0] = $S_OK)
EndFunc   ;==>RegisterDragDrop

Func RevokeDragDrop($hWnd)
	Local $aCall = DllCall("ole32.dll", "long", "RevokeDragDrop", "hwnd", $hWnd)
	If @error Then Return SetError(TranslateDllError(), 0, False)
	Return SetError($aCall[0], 0, $aCall[0] = $S_OK)
EndFunc   ;==>RevokeDragDrop


; jugador code

Func GetDataObjectOfFile_B(ByRef $sPath)
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
EndFunc

Func GetDataObjectOfFile_C(ByRef $sPath)
    If UBound($sPath) = 0 Then Return 0

    Local $tIID_IDataObject = _WinAPI_GUIDFromString($sIID_IDataObject)
    Local $pIDataObject = __SHCreateDataObject($tIID_IDataObject, 0, 0, 0, 0)
    If Not $pIDataObject Then Return 0

    Local Const $tag_IDataObject = _
                        "GetData hresult(ptr;ptr*);" & _
                        "GetDataHere hresult(ptr;ptr*);" & _
                        "QueryGetData hresult(ptr);" & _
                        "GetCanonicalFormatEtc hresult(ptr;ptr*);" & _
                        "SetData hresult(ptr;ptr;bool);" & _
                        "EnumFormatEtc hresult(dword;ptr*);" & _
                        "DAdvise hresult(ptr;dword;ptr;dword*);" & _
                        "DUnadvise hresult(dword);" & _
                        "EnumDAdvise hresult(ptr*);"
    Local $oIDataObject = ObjCreateInterface($pIDataObject, $sIID_IDataObject, $tag_IDataObject)
    If Not IsObj($oIDataObject) Then
        _Release($pIDataObject)
        Return 0
    Endif

    Local $tFORMATETC, $tSTGMEDIUM
    __Fill_tag_FORMATETC($tFORMATETC)
    __Fill_tag_STGMEDIUM($tSTGMEDIUM, $sPath)

    $oIDataObject.SetData(DllStructGetPtr($tFORMATETC), DllStructGetPtr($tSTGMEDIUM), 1)
    _AddRef($pIDataObject)

    Return $pIDataObject
EndFunc

Func __Fill_tag_FORMATETC(Byref $tFORMATETC)
    Local Const $CF_HDROP = 15
    Local Const $TYMED_HGLOBAL = 1

    $tFORMATETC = DllStructCreate("ushort cfFormat; ptr ptd; uint dwAspect; int lindex; uint tymed")
    DllStructSetData($tFORMATETC, "cfFormat", $CF_HDROP)
    DllStructSetData($tFORMATETC, "dwAspect", 1)
    DllStructSetData($tFORMATETC, "lindex", -1)
    DllStructSetData($tFORMATETC, "tymed", $TYMED_HGLOBAL)
EndFunc

Func __Fill_tag_STGMEDIUM(Byref $tSTGMEDIUM, Byref $aFiles)
    Local Const $CF_HDROP = 15
    Local Const $TYMED_HGLOBAL = 1

    Local $sFileList = ""
    For $i = 0 To UBound($aFiles) - 1
        $sFileList &= $aFiles[$i] & Chr(0)
    Next
    $sFileList &= Chr(0)

    Local $iSize = 20 + (StringLen($sFileList) * 2)

    Local $hGlobal = DllCall($hKernel32, "ptr", "GlobalAlloc", "uint", 0x2042, "ulong_ptr", $iSize)[0]
    Local $pLock = DllCall($hKernel32, "ptr", "GlobalLock", "ptr", $hGlobal)[0]

    Local $tDROPFILES = DllStructCreate("dword pFiles; int x; int y; bool fNC; bool fWide", $pLock)
    DllStructSetData($tDROPFILES, "pFiles", 20) 
    DllStructSetData($tDROPFILES, "fWide", True)

    Local $tPaths = DllStructCreate("wchar[" & StringLen($sFileList) & "]", $pLock + 20)
    DllStructSetData($tPaths, 1, $sFileList)

    DllCall($hKernel32, "bool", "GlobalUnlock", "ptr", $hGlobal)

    $tSTGMEDIUM = DllStructCreate("uint tymed; ptr hGlobal; ptr pUnkForRelease")
    DllStructSetData($tSTGMEDIUM, "tymed", $TYMED_HGLOBAL)
    DllStructSetData($tSTGMEDIUM, "hGlobal", $hGlobal)
    DllStructSetData($tSTGMEDIUM, "pUnkForRelease", 0)
EndFunc

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
EndFunc

; jugador code above

Func _VersionToString($arVersion, $sSep = " ")
    If Not IsArray($arVersion) Or UBound($arVersion, 0)<2 Or UBound($arVersion, 1)<2 Then Return SetError(1, 1, "Version not parsed.")
    Local $sVersion = ""
    If $arVersion[1][0]>0 Then
        $sVersion &= $arVersion[1][1]
    EndIf
    If $sVersion = "" Then Return "Version unknown"
    Return $sVersion
EndFunc

Func _UDFGetVersion($sFile)
    Local $sCode = FileRead($sFile)
    If @error Then Return SetError(@error, @extended, 0)
    Local $arVersion = _GetVersion($sCode)
    If @error Then Return SetError(@error, @extended, -1)
    Return $arVersion
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetVersion
; Description ...: Get the AutoIt Version as well as the UDF Version.
; Syntax ........: _GetVersion($sUdfCode)
; Parameters ....: $sUdfCode               - the sourcecode of the udf
; Return values .: Array with version information.
; Author ........: Kanashius
; Modified ......:
; Remarks .......: @extended (1 - Only AutoIt Version found, 2 - Only UDF Version found, 3 - Both found)
;                  Resurns a 2D-Array with:
;                  [0][0] being the amount of version parts found for the AutoIt Version + 1
;                  [0][1] If [0][0]>0 then this is the full autoit version as string
;                  [0][2] The first part of the autoit version (index 2-5 is a number)
;                  ...
;                  [0][6] The last part of the autoit version (last part is a/b/rc)
;                  [1][0] being the amount of version parts found for the UDF Version + 1
;                  [1][1] If [1][0]>0 then this is the full udf version as string
;                  [1][2] The first part of the udf version (index 2-5 is a number)
;                  ...
;                  [1][6] The last part of the udf version (last part is a/b/rc)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetVersion($sUdfCode)
    Local $iExtended = 0
    Local $arAutoItVersion = StringRegExp($sUdfCode, "(?m)(?s)^;\s*#INDEX#\s*=*.*?;\s*AutoIt\s*Version\s*\.*:\s((\d+)(?:\.(\d+)(?:\.(\d+)(?:\.(\d+))?)?)?(a|b|rc)?)\s*$", 1)
    If Not @error Then $iExtended = 1
    Local $arUDFVersion =  StringRegExp($sUdfCode, "(?m)(?s)^;\s*#INDEX#\s*=*.*?;\s*Version\s*\.*:\s((\d+)(?:\.(\d+)(?:\.(\d+))?)?(a|b|rc)?)\s*$", 1)
    If Not @error Then $iExtended += 2
    Local $iVerNumbers = UBound($arAutoItVersion)
    If UBound($arUDFVersion)>$iVerNumbers Then $iVerNumbers = UBound($arUDFVersion)
    Local $arResult[2][$iVerNumbers+1]
    $arResult[0][0] = UBound($arAutoItVersion)
    For $i=0 to UBound($arAutoItVersion)-1 Step 1
        $arResult[0][$i+1] = $arAutoItVersion[$i]
    Next
    $arResult[1][0] = UBound($arUDFVersion)
    For $i=0 to UBound($arUDFVersion)-1 Step 1
        $arResult[1][$i+1] = $arUDFVersion[$i]
    Next
    Return SetExtended($iExtended, $arResult)
EndFunc

Func _filterCallback($hSystem, $hView, $bIsFolder, $sPath, $sName, $sExt)
	#forceref $hSystem, $hView, $bIsFolder

	; nothing to filter
	If Not $bHideHidden And Not $bHideSystem Then Return True

	; always show root drive letters
	Local $sFullPath = $sPath & $sName & $sExt
	If _WinAPI_PathIsRoot_mod($sFullPath) Then Return True

	Switch $sName
		Case "$RECYCLE.BIN"
			Return False
		Case "System Volume Information"
			Return False
	EndSwitch

	; fetch attributes and apply both filters (System and Hidden) independently
	Local $sAttrib = FileGetAttrib($sFullPath)
	If $bHideSystem And StringInStr($sAttrib, "S", 2) > 0 Then Return False
	If $bHideHidden And StringInStr($sAttrib, "H", 2) > 0 Then Return False

	Return True
EndFunc

Func _WinSetIcon($hWnd, $sFile, $iIndex = 0, $bSmall = False) ; https://www.autoitscript.com/forum/topic/168698-changing-a-windows-icon/#findComment-1461109
  Local $WM_SETICON = 128, $ICON_SMALL = 0, $ICON_BIG = 1, $hIcon = _WinAPI_ExtractIcon($sFile, $iIndex, $bSmall)
  If Not $hIcon Then Return SetError(1, 0, 1) ; https://learn.microsoft.com/en-us/windows/win32/winmsg/wm-seticon
  _SendMessage($hWnd, $WM_SETICON, Int(Not $bSmall), $hIcon)
  _WinAPI_DestroyIcon($hIcon)
EndFunc   ;==>_WinSetIcon
