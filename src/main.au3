#AutoIt3Wrapper_UseX64=Y
;#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include <GUIConstantsEx.au3>
#include <WindowsNotifsConstants.au3>
#include <WindowsStylesConstants.au3>
#include <WindowsConstants.au3>
#include <GuiTreeView.au3>
#include <WinAPITheme.au3>
#include <GuiToolTip.au3>

#include "../lib/TreeListExplorer.au3"
#include "../lib/GUIFrame_WBD_Mod.au3"
#include "../lib/History.au3"

; CREDITS:
; Kanashius     TreeListExplorer UDF
; pixelsearch   Detached Header and ListView synchronization
; ioa747        Detached Header subclassing for dark mode
; Nine          Custom Draw for Buttons
; argumentum    Dark Mode functions
; NoNameCode    Dark Mode functions
; Melba23       GUIFrame UDF
; UEZ           Lots and lots and lots

Global $sVersion = "2026-01-12"

; set base DPI scale value and apply DPI
Global $DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = -2
Global $iDPI = 1
$iDPI = ApplyDPI()

; DPI must be set before ownerdrawn menu
#include "../lib/ModernMenuRaw.au3"

Global $hTLESystem, $iFrame_A, $hSecondFrame, $g_bHeaderEndDrag, $hSeparatorFrame, $aWinSize2, $idInputPath, $g_hInputPath, $g_hStatus
Global $g_hGUI, $g_hChild, $g_hHeader, $g_hListview, $idListview, $iHeaderHeight, $hChildLV, $hParentFrame, $g_iIconWidth, $g_hTreeView
Global $g_hSizebox, $g_hOldProc, $g_iHeight, $g_aTextStatus, $g_aRatioW, $g_hDots
Global $idPropertiesItem, $idPropertiesLV, $sCurrentPath
Global $sBack, $sForward, $sUpLevel, $sRefresh
Global $hCursor, $hProc, $g_hBrush
Global $iTimeCalled, $iTimeDiff
Global $bPathSelectAll, $bPathInputChanged, $sSelectedItems, $g_aText
Global $idSeparator, $idThemeItem, $hToolTip2, $hToolTip1
Global $isDarkMode = _WinAPI_ShouldAppsUseDarkMode()
Global $hFolderHistory=__History_Create("_doUnReDo", 100, "_historyChange"), $bFolderHistoryChanging = False
Global $sBack, $sForward, $sUpLevel
Global Const $SBS_SIZEBOX = 0x08, $SBS_SIZEGRIP = 0x10
Global Const $CLR_TEXT = 0xFFFFFF ; The text color
Global Const $SB_LEFT = 6
Global Const $APPMODE_FORCEDARK = 2
Global Const $APPMODE_FORCELIGHT = 3
Global $iTopSpacer = 14
Global $aPosTip, $iOldaPos0, $iOldaPos1
; force light mode
;$isDarkMode = False

Global $tInfo, $gText

Global $hKernel32 = DllOpen('kernel32.dll')
Global $hGdi32 = DllOpen('gdi32.dll')
Global $hUser32 = DllOpen('user32.dll')
Global $hShlwapi = DllOpen('shlwapi.dll')
Global $hShell32 = DllOpen('shell32.dll')

Global $global_StatusBar_Text = "  Part 1"

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

; menu line color
Global $iColorNew = $iBackColorDef
Global $iColorOld = _WinAPI_GetSysColor($COLOR_MENU)

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
    ; dpi
    Local $iTreeListIconSize
    If $iDPI = 1 Then
        $iTreeListIconSize = 16
    ElseIf $iDPI = 1.25 Then
        $iTreeListIconSize = 20
    ElseIf $iDPI = 1.5 Then
        $iTreeListIconSize = 24
    ElseIf $iDPI = 2 Then
        $iTreeListIconSize = 32
    Else
        $iTreeListIconSize = 16
    EndIf

    ; check font availability
    Local $sButtonFont
    If _WinAPI_GetFontName("Segoe Fluent Icons") Then
        ; Segoe Fluent Icons are available, use for buttons (Windows 11)
        $sButtonFont = "Segoe Fluent Icons"
    Else
        ; Segoe Fluent Icons are not available, fall back to Segoe MDL2 Assets (Windows 10)
        $sButtonFont = "Segoe MDL2 Assets"
    EndIf

    ; StartUp of the TreeListExplorer UDF (required)
    __TreeListExplorer_StartUp($__TreeListExplorer_Lang_EN, $iTreeListIconSize)
    If @error Then ConsoleWrite("__TreeListExplorer_StartUp failed: "&@error&":"&@extended&@crlf)

    ;Create GUI
    $g_hGUI = GUICreate("Files Au3", @DesktopWidth - 600, @DesktopHeight - 400, -1, -1, $WS_OVERLAPPEDWINDOW)
    $FrameWidth1 = @DesktopWidth / 4

    _InitDarkSizebox()

    ; statusbar create
    $g_hStatus = _GUICtrlStatusBar_Create($g_hGUI, -1, "", $WS_CLIPSIBLINGS)

    $aClientSize = WinGetClientSize($g_hGUI)

    ; set tooltips theming
    $hToolTip2 = _GUIToolTip_Create(0)
    _GUIToolTip_SetMaxTipWidth($hToolTip2, 400)
    $hToolTip1 = _GUIToolTip_Create(0)
    _GUIToolTip_SetMaxTipWidth($hToolTip1, 400)

    GUISetFont(10, $FW_NORMAL, $GUI_FONTNORMAL, $sButtonFont)

    $sBack = GUICtrlCreateButton(ChrW(0xE64E), $sButtonSpacing, 10, -1, -1)
    GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKWIDTH)
    $aPos = ControlGetPos($g_hGUI, "", $sBack)
    $sBackPosV = $aPos[1] + $aPos[3]
    $sBackPosH = $aPos[0] + $aPos[2]
    $iButtonHeight = $aPos[3]
	GUICtrlSetState($sBack, $GUI_DISABLE)
    _GUIToolTip_AddTool($hToolTip2, $g_hGUI, "Back", GUICtrlGetHandle($sBack))

    $sForward = GUICtrlCreateButton(ChrW(0xE64D), $sBackPosH + $sButtonSpacing, 10, -1, -1)
    GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKWIDTH)
    $aPos = ControlGetPos($g_hGUI, "", $sForward)
    $sForwardPosV = $aPos[1] + $aPos[3]
    $sForwardPosH = $aPos[0] + $aPos[2]
	GUICtrlSetState($sForward, $GUI_DISABLE)
    _GUIToolTip_AddTool($hToolTip2, $g_hGUI, "Forward", GUICtrlGetHandle($sForward))

    $sUpLevel = GUICtrlCreateButton(ChrW(0xE64C), $sForwardPosH + $sButtonSpacing, 10, -1, -1)
    GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKWIDTH)
    $aPos = ControlGetPos($g_hGUI, "", $sUpLevel)
    $sUpLevelPosV = $aPos[1] + $aPos[3]
    $sUpLevelPosH = $aPos[0] + $aPos[2]
    _GUIToolTip_AddTool($hToolTip2, $g_hGUI, "Up", GUICtrlGetHandle($sUpLevel))

    $sRefresh = GUICtrlCreateButton(ChrW(0xE72C), $sUpLevelPosH + $sButtonSpacing, 10, -1, -1)
    GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR + $GUI_DOCKLEFT + $GUI_DOCKWIDTH)
    $aPos = ControlGetPos($g_hGUI, "", $sRefresh)
    $sRefreshPosV = $aPos[1] + $aPos[3]
    $sRefreshPosH = $aPos[0] + $aPos[2]
    _GUIToolTip_AddTool($hToolTip2, $g_hGUI, "Refresh", GUICtrlGetHandle($sRefresh))

    ; reset GUI font
    GUISetFont(10, $FW_NORMAL, $GUI_FONTNORMAL, "Segoe UI")

    ; Menu
    If $isDarkMode Then
        Local $idFileMenu = _GUICtrlCreateODTopMenu("& File", $g_hGUI)
        Local $idViewMenu = _GUICtrlCreateODTopMenu("& View", $g_hGUI)
        Local $idHelpMenu = _GUICtrlCreateODTopMenu("& Help", $g_hGUI)
    Else
        Local $idFileMenu = _GUICtrlCreateODTopMenu("& File", $g_hGUI)
        Local $idViewMenu = _GUICtrlCreateODTopMenu("& View", $g_hGUI)
        Local $idHelpMenu = _GUICtrlCreateODTopMenu("& Help", $g_hGUI)
    EndIf

    $idPropertiesItem = GUICtrlCreateMenuItem("&Properties", $idFileMenu)
    GUICtrlSetState($idPropertiesItem, $GUI_DISABLE)
    GUICtrlCreateMenuItem("", $idFileMenu)
    Local $idExitItem = GUICtrlCreateMenuItem("&Exit", $idFileMenu)
    $idThemeItem = GUICtrlCreateMenuItem("&Dark Mode", $idViewMenu)
    Local $idAboutItem = GUICtrlCreateMenuItem("&About", $idHelpMenu)

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

    ;Create Frames
    Local $iSpacer
    If $iDPI = 1 Then
        $iSpacer = 4
    ElseIf $iDPI = 1.25 Then
        $iSpacer = 9
    ElseIf $iDPI = 1.5 Then
        $iSpacer = 13
    ElseIf $iDPI = 2 Then
        $iSpacer = 18
    Else
        $iSpacer = 6 * $iDPI
    EndIf

    ;$iFrame_A = _GUIFrame_Create($g_hGUI, 0, $FrameWidth1, 5, 0, $sRefreshPosV + 4)
    ;$iFrame_A = _GUIFrame_Create($g_hGUI, 0, $FrameWidth1, 5, 0, $sRefreshPosV + 4, 0, $aClientSize[1] - _GUICtrlStatusBar_GetHeight($g_hStatus) - $sRefreshPosV - 9)
    $iFrame_A = _GUIFrame_Create($g_hGUI, 0, $FrameWidth1, 9, 0, $sRefreshPosV + $iTopSpacer, 0, $aClientSize[1] - _GUICtrlStatusBar_GetHeight($g_hStatus) - $sRefreshPosV - $iSpacer)

    ;Set min sizes for the frames
    _GUIFrame_SetMin($iFrame_A, 200, 600)

    ;Create Explorer Listviews
    _GUIFrame_Switch($iFrame_A, 1)
    $aWinSize1 = WinGetClientSize(_GUIFrame_GetHandle($iFrame_A, 1))

    Global $idTreeView = GUICtrlCreateTreeView(0, 0, $aWinSize1[0], $aWinSize1[1] - $sRefreshPosV - $iTopSpacer, BitOR($GUI_SS_DEFAULT_TREEVIEW, $TVS_TRACKSELECT))
    GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
    $g_hTreeView = GUICtrlGetHandle($idTreeView)

    ; Create TLE system
    Local $hTLESystem = __TreeListExplorer_CreateSystem($g_hGUI, "", "_folderCallback")
    If @error Then ConsoleWrite("__TreeListExplorer_CreateSystem left failed: "&@error&":"&@extended&@crlf)
    ; Add Views to TLE system
    __TreeListExplorer_AddView($hTLESystem, $idInputPath)
    If @error Then ConsoleWrite("__TreeListExplorer_AddView $idInputPath failed: "&@error&":"&@extended&@crlf)
    __TreeListExplorer_AddView($hTLESystem, $idTreeView)
    If @error Then ConsoleWrite("__TreeListExplorer_AddView $idTreeView failed: "&@error&":"&@extended&@crlf)
    ;__TreeListExplorer_SetCallback($idTreeView, $__TreeListExplorer_Callback_Click, "_clickCallback")
    ;__TreeListExplorer_SetCallback($idTreeView, $__TreeListExplorer_Callback_DoubleClick, "_doubleClickCallback")

    _GUIFrame_Switch($iFrame_A, 2)

    $aWinSize2 = WinGetClientSize(_GUIFrame_GetHandle($iFrame_A, 2))

    $hChildLV = _GUIFrame_GetHandle($iFrame_A, 2)
    $g_hHeader = _GUICtrlHeader_Create($hChildLV, BitOR($HDS_BUTTONS, $HDS_DRAGDROP, $HDS_FULLDRAG))
    ;$g_hHeader = _GUICtrlHeader_Create($hChildLV, BitOR($HDS_BUTTONS, $HDS_HOTTRACK, $HDS_DRAGDROP))
    GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
    _GUICtrlHeader_AddItem($g_hHeader, "Name", 300)
    _GUICtrlHeader_AddItem($g_hHeader, "Size", 100)
    _GUICtrlHeader_AddItem($g_hHeader, "Date Modified", 150)
    _GUICtrlHeader_AddItem($g_hHeader, "Type", 150)

    ; Set Size column alignment
    _GUICtrlHeader_SetItemAlign($g_hHeader, 1, 1)

    ; Set sort arrow
    _GUICtrlHeader_SetItemFormat($g_hHeader, 0, $HDF_SORTUP)

    ; get header height
    $iHeaderHeight = _WinAPI_GetWindowHeight($g_hHeader)

    ;_WinAPI_SetWindowPos_mod($g_hHeader, 0, 10, 0, $aWinSize2[0] - 11, $iHeaderHeight, BitOR($SWP_NOZORDER, $SWP_NOACTIVATE))

    $idListview = GUICtrlCreateListView("Name|Size|Date Modified|Type", 0, $iHeaderHeight, $aWinSize2[0], $aWinSize2[1] - $iHeaderHeight - $sRefreshPosV - $iTopSpacer, BitOR($LVS_SHOWSELALWAYS, $LVS_NOCOLUMNHEADER), BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_DOUBLEBUFFER, $LVS_EX_TRACKSELECT))
    GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)

    $g_hListview = GUICtrlGetHandle($idListview)

    _GUIToolTip_AddTool($hToolTip1, $g_hGUI, "", $g_hListview)

    ; right align Size column
    _GUICtrlListView_JustifyColumn($idListview, 1, 1)

    ; listview context menu
    Local $idContextLV = GUICtrlCreateContextMenu($idListview)
    Local $idPropertiesLV = GUICtrlCreateMenuItem("Properties", $idContextLV)

    __TreeListExplorer_AddView($hTLESystem, $idListview, True, True, True, False, False)
    If @error Then ConsoleWrite("__TreeListExplorer_AddView $idListview failed: "&@error&":"&@extended&@crlf)
    ;__TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_Click, "_clickCallback")
    __TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_DoubleClick, "_doubleClickCallback")
    __TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_ListViewPaths, "_handleListViewData")
    __TreeListExplorer_SetCallback($idListview, $__TreeListExplorer_Callback_ListViewItemCreated, "_handleListViewItemCreated")

    ; Set resizing flag for all created frames
    _GUIFrame_ResizeSet(0)

    ; Register the $WM_SIZE handler to permit resizing
    ;_GUIFrame_ResizeReg()

    GUIRegisterMsg($WM_COMMAND, "WM_COMMAND2")
    GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY2")

    ; get rid of the focus rectangle dots
    GUICtrlSendMsg($idTreeView, $WM_CHANGEUISTATE, 65537, 0)

    ; add listview column for Type (for some reason this only works when added later)
    ;_GUICtrlListView_AddColumn($g_hListview, "Type", 500)
    ;_GUICtrlListView_SetColumnWidth($g_hListview, 3, 500)

    ; resize listview columns to match header widths + update global variable $g_iIconWidth
    _resizeLVCols()

    ; statusbar
    Local $iParts = $aClientSize[0] / 4
    Local $aParts[4] = [$iParts, $iParts * 2, $iParts * 3, -1]
    ;Local $g_aText[4] = [" ", " ", " ", " "]
    Dim $g_aText[Ubound($aParts)] = [" ", " ", " ", " "]
    _GUICtrlStatusBar_SetParts($g_hStatus, $aParts)
    _GUICtrlStatusBar_SetText($g_hStatus, $g_aText[0], 0, $SBT_OWNERDRAW)
    _GUICtrlStatusBar_SetText($g_hStatus, $g_aText[1], 1, $SBT_OWNERDRAW)
    _GUICtrlStatusBar_SetText($g_hStatus, $g_aText[2], 2, $SBT_OWNERDRAW)
    _GUICtrlStatusBar_SetText($g_hStatus, $g_aText[3], 3, $SBT_OWNERDRAW)

    _setThemeColors()

    If $isDarkMode Then
        ; set COLOR_MENU system color
        _WinAPI_SetSysColors($COLOR_MENU, $iColorNew)
    EndIf

    GUIRegisterMsg($WM_MOVE, "WM_MOVE")
    GUIRegisterMsg($WM_SIZE, "WM_SIZE")
    GUIRegisterMsg($WM_DRAWITEM, "WM_DRAWITEM2")

    _GUICtrl_SetFont($g_hHeader, 16 * $iDPI, 400, 0, "Segoe UI")

    ; update variable for header height
    $iHeaderHeight = _WinAPI_GetWindowHeight($g_hHeader)

    If $isDarkMode Then
        _WinAPI_DwmSetWindowAttribute_unr($g_hGUI, 38, 2)
        _WinAPI_DwmExtendFrameIntoClientArea($g_hGUI, _WinAPI_CreateMargins(-1, -1, -1, -1))
    EndIf

    GUISetState(@SW_SHOW, $g_hGUI)

    ; get parent frame handle
    Local $aData = _WinAPI_EnumChildWindows(_GetHwndFromPID(@AutoItPID))
    For $i = 1 to $aData[0][0]
        $aData[$i][1] = _WinAPI_GetWindowText($aData[$i][0])
        If $aData[$i][1] = "FrameParent" Then $hParentFrame = $aData[$i][0]
    Next

    ; get separator frame handle
    Local $aData = _WinAPI_EnumChildWindows(_GetHwndFromPID(@AutoItPID))
    For $i = 1 to $aData[0][0]
        $aData[$i][1] = _WinAPI_GetWindowText($aData[$i][0])
        If $aData[$i][1] = "SeparatorFrame" Then $hSeparatorFrame = $aData[$i][0]
    Next

    ; set background color for separator frame and parent frame
    If $isDarkMode = True Then
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

    _removeExStyles()

    ;_GUICtrlListView_RegisterSortCallBack($g_hListview, 2) ; 2 = natural sort
    ; add sort arrow to header
    ;$__g_aListViewSortInfo[1][10] = $g_hHeader

    _GUICtrlListView_SetHoverTime($idListview, 500)

    ; apply theme separately to tooltips
    _themeTooltips()

    ; add composited to treeview frame
    Local $i_ExStyle_Old = _WinAPI_GetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 1), $GWL_EXSTYLE)
    _WinAPI_SetWindowLong_mod(_GUIFrame_GetHandle($iFrame_A, 1), $GWL_EXSTYLE, BitOR($i_ExStyle_Old, $WS_EX_COMPOSITED))

    While True
        Local $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $idThemeItem
                _switchTheme()
			Case $sBack
				__History_Undo($hFolderHistory)
            Case $sForward
				__History_Redo($hFolderHistory)
            Case $sUpLevel
				__TreeListExplorer_OpenPath($hTLESystem, __TreeListExplorer__GetPathAndLast(__TreeListExplorer_GetPath($hTLESystem))[0])
            Case $sRefresh
				__TreeListExplorer_Reload($hTLESystem)
            Case $GUI_EVENT_CLOSE
                ExitLoop
            Case $GUI_EVENT_RESIZED, $GUI_EVENT_MAXIMIZE
                _resizeLVCols2()
            Case $idExitItem
                ExitLoop
            Case $idPropertiesItem, $idPropertiesLV
                _Properties()
            Case $idAboutItem
                _About()
        EndSwitch

        If $g_bHeaderEndDrag Then
            $g_bHeaderEndDrag = False
            _reorderLVCols()
        EndIf

        If $bPathSelectAll Then
            ; select all text in path input box
            ControlFocus($g_hGUI, "", $idInputPath)
            _GUICtrlEdit_SetSel($g_hInputPath, 0, -1)
            ; reset variable
            $bPathSelectAll = False
        EndIf

        If $bPathInputChanged Then
            ; reset position of header and listview
            GUICtrlSendMsg($idListview, $WM_HSCROLL, $SB_LEFT, 0)
            WinMove($g_hHeader, "", 0, 0, WinGetPos($g_hChild)[2], Default)

            ; update number of items (files and folders) in statusbar
            $g_aText[0] = "  " & _GUICtrlListView_GetItemCount($idListview) & " items"
            ; update drive space information
            Local $iDriveFree = Round(DriveSpaceFree(GUICtrlRead($idInputPath)) / 1024, 1)
            Local $iDriveTotal = Round(DriveSpaceTotal(GUICtrlRead($idInputPath)) / 1024, 1)
            Local $iPercentFree = Round(($iDriveFree / $iDriveTotal) * 100)
            $g_aText[3] = $iDriveFree & " GB free" & " (" & $iPercentFree & "%)"
            _WinAPI_RedrawWindow($g_hStatus)
            ; reset variable
            $bPathInputChanged = False
        EndIf

        Local $aCursor = GUIGetCursorInfo($g_hGUI)
        If $aCursor[4] <> $idListview Then
            ; cancel tooltip when not over listview
            _GUIToolTip_TrackActivate($hToolTip1, False, $g_hGUI, $g_hListview)
            ; reset the value stored in the tooltip
            $gText = ""
            _GUIToolTip_UpdateTipText($hToolTip1, $g_hGUI, $g_hListview, $gText)
        EndIf
        Sleep(10)
    WEnd
EndFunc

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
EndFunc

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

        GUICtrlSetState($sBack, $GUI_HIDE)
        GUICtrlSetState($sBack, $GUI_SHOW)
        GUICtrlSetState($sForward, $GUI_HIDE)
        GUICtrlSetState($sForward, $GUI_SHOW)
        GUICtrlSetState($sUpLevel, $GUI_HIDE)
        GUICtrlSetState($sUpLevel, $GUI_SHOW)
        GUICtrlSetState($sRefresh, $GUI_HIDE)
        GUICtrlSetState($sRefresh, $GUI_SHOW)

        _SetMenuBkColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)))
        _SetMenuSelectBkColor(_WinAPI_ColorAdjustLuma(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)), -6))
        _SetMenuSelectRectColor(_WinAPI_ColorAdjustLuma(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)), -6))
        _SetMenuSelectTextColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT)))
        _SetMenuTextColor(_WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT)))

        _GUIMenuBarSetBkColor($g_hGUI, _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOW)))

        GUICtrlSetState($idThemeItem, $GUI_UNCHECKED)
        _themeTooltips()
        $iColorNew = $iBackColorDef
        _WinAPI_SetSysColors($COLOR_MENU, $iColorNew)
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

        ; separator
        ;Local $i_ExStyle_Old = _WinAPI_GetWindowLong_mod(GUICtrlGetHandle($idSeparator), $GWL_EXSTYLE)
        ;_WinAPI_SetWindowLong_mod(GUICtrlGetHandle($idSeparator), $GWL_EXSTYLE, BitXOR($i_ExStyle_Old, $WS_EX_DLGMODALFRAME))
        GUICtrlSetBkColor($idSeparator, 0x505050)
        If @OSBuild >= 22621 Then
            GUISetBkColor(0x000000, $hSeparatorFrame)
            GUISetBkColor(0x000000, $hParentFrame)
        Else
            GUISetBkColor($iBackColorDef, $hSeparatorFrame)
            GUISetBkColor($iBackColorDef, $hParentFrame)
        EndIf

        GUICtrlSetState($sBack, $GUI_HIDE)
        GUICtrlSetState($sBack, $GUI_SHOW)
        GUICtrlSetState($sForward, $GUI_HIDE)
        GUICtrlSetState($sForward, $GUI_SHOW)
        GUICtrlSetState($sUpLevel, $GUI_HIDE)
        GUICtrlSetState($sUpLevel, $GUI_SHOW)
        GUICtrlSetState($sRefresh, $GUI_HIDE)
        GUICtrlSetState($sRefresh, $GUI_SHOW)

        _SetMenuBkColor($iBackColorDef)
        _SetMenuSelectBkColor(_WinAPI_ColorAdjustLuma($iBackColorDef, 30))
        _SetMenuSelectRectColor(_WinAPI_ColorAdjustLuma($iBackColorDef, 30))
        _SetMenuSelectTextColor(0xffffff)
        _SetMenuTextColor(0xffffff)

        _GUIMenuBarSetBkColor($g_hGUI, $iBackColorDef)

        GUICtrlSetState($idThemeItem, $GUI_CHECKED)
        _themeTooltips()
        $iColorNew = $iBackColorDef
        _WinAPI_SetSysColors($COLOR_MENU, $iColorNew)
        _setThemeColors()
        _WinAPI_RedrawWindow($g_hStatus)
    EndIf
EndFunc

Func _selectionChangedLV()
    Local $sSelectedLV, $sSelectedInfo, $iDirCount = 0, $iFileCount = 0, $iFileSizes = 0
    ; get selections from listview
    Local $aSelectedLV = _GUICtrlListView_GetSelectedIndices($idListview, True)
    Local $sSelectedItem = ""
    $sSelectedItems = ""

    For $i = 1 To $aSelectedLV[0]
        $sSelectedItem = _GUICtrlListView_GetItemText($idListview, $aSelectedLV[$i], 0)
        $sSelectedLV = GUICtrlRead($idInputPath) & $sSelectedItem
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

    If $iDirCount <> 0 Or $iDirCount = 0 Then
        $sSelectedInfo = "(" & $iDirCount & " folders selected" & ")"
    ElseIf $iDirCount = 1 Then
        $sSelectedInfo &= "(" & $iDirCount & " folder selected" & ")"
    EndIf

    If $iFileCount <> 0 Or $iFileCount = 0 Then
        $sSelectedInfo &= @TAB & "(" & $iFileCount & " files selected" & ")"
    ElseIf $iFileCount = 1 Then
        $sSelectedInfo &= @TAB & "(" & $iFileCount & " file selected" & ")"
    EndIf

    Local $iItemCount = $iFileCount + $iDirCount

    ; Properties dialog
    If $iItemCount = 1 Then
        $sCurrentPath = $sSelectedLV
        GUICtrlSetState($idPropertiesItem, $GUI_ENABLE)
        GUICtrlSetState($idPropertiesLV, $GUI_ENABLE)
    ElseIf $iItemCount <> 1 And $iItemCount <> 0 Then
        ; multi-properties
        ; need number of selected items to declare array
        GUICtrlSetState($idPropertiesItem, $GUI_ENABLE)
        GUICtrlSetState($idPropertiesLV, $GUI_ENABLE)
    Else
        GUICtrlSetState($idPropertiesItem, $GUI_DISABLE)
        GUICtrlSetState($idPropertiesLV, $GUI_DISABLE)
    EndIf

    If $iItemCount <> 0 And $iItemCount <> 1 Then
        $g_aText[1] = $iItemCount & " items selected"
        _WinAPI_RedrawWindow($g_hStatus)
    ElseIf $iItemCount = 1 Then
        $g_aText[1] = $iItemCount & " item selected"
        _WinAPI_RedrawWindow($g_hStatus)
    Else
        $g_aText[1] = " "
        _WinAPI_RedrawWindow($g_hStatus)
    EndIf

    If $iFileSizes = 0 Then
        ; clear size status if size equals zero
        $g_aText[2] = " "
        _WinAPI_RedrawWindow($g_hStatus)
    Else
        $g_aText[2] = __TreeListExplorer__GetSizeString($iFileSizes)
        _WinAPI_RedrawWindow($g_hStatus)
    EndIf
EndFunc

Func _handleListViewData($hSystem, $hView, $sPath, ByRef $arPaths)
	ReDim $arPaths[UBound($arPaths)][_GUICtrlListView_GetColumnCount($idListview)] ; resize the array (and return it at the end)
	For $i=0 To UBound($arPaths)-1
		Local $sFilePath = $sPath & $arPaths[$i][0]
        If __TreeListExplorer__PathIsFolder($sFilePath) Then
            $arPaths[$i][2] = __TreeListExplorer__GetTimeString(FileGetTime($sFilePath, 0)) ; add time modified
            $arPaths[$i][3] = _getType($sFilePath, True)
        Else
            $arPaths[$i][1] = FileGetSize($sFilePath) ; Put size as integer numbers here to enable the default sorting
            $arPaths[$i][2] = __TreeListExplorer__GetTimeString(FileGetTime($sFilePath, 0)) ; add time modified
            $arPaths[$i][3] = _getType($sFilePath)
        EndIf
	Next
	; custom sorting could be done here as well, setting the parameter $bEnableSorting to False when adding the ListView. Sorting can then be handled by the user
	Return $arPaths
EndFunc

Func _handleListViewItemCreated($hSystem, $hView, $sPath, $sFilename, $iIndex, $bFolder)
	If Not $bFolder Then _GUICtrlListView_SetItemText($hView, $iIndex, __TreeListExplorer__GetSizeString(_GUICtrlListView_GetItemText($hView, $iIndex, 1)), 1) ; convert size in bytes to the short text form, after sorting
EndFunc

Func _getType($sPath, $bFolder = False)
    Local $tSHFILEINFO = DllStructCreate($tagSHFILEINFO)
    Local $iAttr = ($bFolder?$FILE_ATTRIBUTE_DIRECTORY:$FILE_ATTRIBUTE_NORMAL)
    _WinAPI_ShellGetFileInfo($sPath, BitOR($SHGFI_TYPENAME, $SHGFI_USEFILEATTRIBUTES), $iAttr, $tSHFILEINFO)
    Return DllStructGetData($tSHFILEINFO, 5)
EndFunc

Func _selectCallback($hSystem, $sRoot, $sFolder, $sSelected)
    ;__TreeListExplorer__FileGetIconIndex($sRoot&$sFolder&$sSelected)
	ConsoleWrite("Select "&$hSystem&": "&$sRoot&$sFolder&"["&$sSelected&"]"&@CRLF)
EndFunc

Func _clickCallback($hSystem, $hView, $sRoot, $sFolder, $sSelected, $item)
    ;ConsoleWrite("Click at "&$hView&": "&$sRoot&$sFolder&"["&$sSelected&"] :"&$item&@CRLF)

    ConsoleWrite("Click at "&$hView&": "&$sRoot&$sFolder&"["&$sSelected&"] :"&$item&@CRLF)
	If $hView=GUICtrlGetHandle($idListview) Then
		Local $sSel = _GUICtrlListView_GetSelectedIndices($hView)
		If StringInStr($sSel, "|") Then ConsoleWrite("Multiple selected items: "&$sSel&@CRLF)
	EndIf

    #cs
    ; get filename for currently selected ListView item
    Local $Array = _GUICtrlListView_GetItemTextArray($idListview, -1)
    ; ensure that the array contains information and that file size is not blank
    If $Array[0] <> 0 And $Array[2] <> "" Then
        ; update statusbar part with file size
        _GUICtrlStatusBar_SetText($g_hStatus, $Array[2] & " (selected)", 1, $SBT_NOBORDERS)
    Else
        ; selected file has been cleared, so clear status text
        _GUICtrlStatusBar_SetText($g_hStatus, " ", 1, $SBT_NOBORDERS)
    EndIf
    #ce
EndFunc

Func _doubleClickCallback($hSystem, $hView, $sRoot, $sFolder, $sSelected, $item)
    ; get filename for currently selected ListView item
    Local $Array = _GUICtrlListView_GetItemTextArray($idListview, -1)
    ; ensure that the array contains information and that filename is not blank
    If $Array[0] <> 0 And $Array[1] <> "" Then
        ; open file in ListView when double-clicking (uses Windows defaults per extension)
        ShellExecute($sRoot & $sFolder & $Array[1])
    EndIf
EndFunc

Func _filterCallback($hSystem, $hView, $bIsFolder, $sPath, $sName, $sExt)
	; ConsoleWrite("Filter: "&$hSystem&" > "&$hView&" -- Folder: "&$bIsFolder&" Path: "&$sPath&" Filename: "&$sName&" Ext: "&$sExt&@crlf)
	Return $bIsFolder Or $sExt=".au3"
EndFunc

Func _loadingCallback($hSystem, $hView, $sRoot, $sFolder, $sSelected, $sPath, $bLoading)
	; ConsoleWrite("Loading "&$hSystem&": Status: "&$bLoading&" View: "&$hView&" >> "&$sRoot&$sFolder&"["&$sSelected&"] >> "&$sPath&@CRLF)
EndFunc

Func _folderCallback($hSystem, $sRoot, $sFolder, $sSelected)
	Local Static $sFolderPrev=""
	If $sFolder<>$sFolderPrev Then
		Local $arData = [$hSystem, $sFolderPrev, $sFolder]
		If Not $bFolderHistoryChanging Then __History_Add($hFolderHistory, $arData)
		$sFolderPrev = $sFolder
	EndIf
	GUICtrlSetState($sUpLevel, $sFolder<>""?$GUI_ENABLE:$GUI_DISABLE)
EndFunc

Func _doUnReDo($hHistory, $bRedo, $arData)
	$bFolderHistoryChanging = True
	If $bRedo Then
		__TreeListExplorer_OpenPath($arData[0], $arData[2])
	Else
		__TreeListExplorer_OpenPath($arData[0], $arData[1])
	EndIf
	$bFolderHistoryChanging = False
EndFunc

Func _historyChange($hHistory)
	GUICtrlSetState($sBack, __History_UndoCount($hHistory)>0?$GUI_ENABLE:$GUI_DISABLE)
	GUICtrlSetState($sForward, __History_RedoCount($hHistory)>0?$GUI_ENABLE:$GUI_DISABLE)
EndFunc

Func WM_NOTIFY2($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam
	__TreeListExplorer__WinProc($hWnd, $iMsg, $wParam, $lParam)

    ; used for dark mode header text color and header and listview combined functionality
    Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)

    ; keep previous listview item row (related to tooltips)
    Local Static $iItemPrev
    Local $iItemRow

    ;Local $tInfo, $gText

    ; header and listview combined functionality
    ;Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
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
                    ;Local $tRECT, $iHeaderColWidth
                    ;$tRECT = _GUICtrlHeader_GetItemRectEx($g_hHeader, $iHeaderItem)
                    ;$iHeaderColWidth = DllStructGetData($tRECT, "Right") - DllStructGetData($tRECT, "Left")
                    ;_GUICtrlListView_SetColumnWidth($g_hListview, $iHeaderItem, $iHeaderColWidth)
                    _GUICtrlListView_SetColumnWidth($g_hListview, $iHeaderItem, $iHeaderItemWidth)
                    _resizeLVCols2()
                    Return False ; to continue tracking the divider

                Case $HDN_ENDDRAG
                    $g_bHeaderEndDrag = True
                    Return False ; to allow the control to automatically place and reorder the item

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
                    ;_GUICtrlListView_SortItems($g_hListview, $iHeaderItem)
                    ;Local $tNMListView = DllStructCreate($tagNMLISTVIEW, $ilParam)
					Local $iCol = $iHeaderItem
					Local $hHeader = $g_hHeader
                    Local $hView = $g_hListView
					For $i = 0 To _GUICtrlHeader_GetItemCount($hHeader) - 1
						If $i=$iCol Then ContinueLoop
						_GUICtrlHeader_SetItemFormat($hHeader, $i, BitAND(_GUICtrlHeader_GetItemFormat($hHeader, $i), BitNOT(BitOR($HDF_SORTDOWN, $HDF_SORTUP))))
					Next
					$iFormat = _GUICtrlHeader_GetItemFormat($hHeader, $iCol)
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

        Case $g_hListView
            Switch $iCode
                Case $LVN_ENDSCROLL
                    Local Static $tagNMLVSCROLL = $tagNMHDR & ";int dx;int dy"
                    Local $tNMLVSCROLL = DllStructCreate($tagNMLVSCROLL, $lParam)
                    If $tNMLVSCROLL.dy = 0 Then ; ListView horizontal scrolling
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
                    ; item selection(s) have changed
                    _selectionChangedLV()
                Case $LVN_HOTTRACK; Sent by a list-view control When the user moves the mouse over an item
                    $tInfo = DllStructCreate($tagNMLISTVIEW, $lParam)
                    $gText = _GUICtrlListView_GetItemText($hWndFrom, DllStructGetData($tInfo, "Item"), 0)
                    $iItemRow = DllStructGetData($tInfo, "Item")
                    ; clear tooltip if cursor not over column 0 or different item
                    If DllStructGetData($tInfo, "SubItem") <> 0 Or $iItemRow <> $iItemPrev Then
                        ; ensure that tooltip only shows when over column 0
                        _GUIToolTip_TrackActivate($hToolTip1, False, $g_hGUI, $g_hListview)
                        ; reset the value stored in the tooltip
                        $gText = ""
                        _GUIToolTip_UpdateTipText($hToolTip1, $g_hGUI, $g_hListview, $gText)
                    EndIf
                    $iItemPrev = $iItemRow
                    Return 0; Allow the ListView to perform its normal track select processing.
                Case $NM_HOVER ; Sent by a list-view control when the mouse hovers over an item
                    ; need to determine if file or folder to get more details
                    If $gText <> "" Then
                        $gText = GUICtrlRead($idInputPath) & $gText & @CRLF
                        Local $gText2 = GUICtrlRead($idInputPath) & _GUICtrlListView_GetItemText($g_hListview, _GUICtrlListView_GetHotItem($g_hListview), 0)
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
                    EndIf
                    Return 1 ; prevent the hover from being processed
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
			;_WinAPI_InflateRect_mod($tRECT, -4, -6)
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
            Local $tDTTOPTS = DllStructCreate($tagDTTOPTS)
            DllStructSetData($tDTTOPTS, 'Size', DllStructGetSize($tDTTOPTS))
            DllStructSetData($tDTTOPTS, 'Flags', BitOR($DTT_TEXTCOLOR, $DTT_GLOWSIZE, $DTT_COMPOSITED))
            DllStructSetData($tDTTOPTS, 'clrText', $iTextColorDef)
            DllStructSetData($tDTTOPTS, 'GlowSize', 12)
            ;Local $hTheme = _WinAPI_OpenThemeData($g_hGUI, 'Globals')
			_WinAPI_DrawText_mod($tInfo.hDC, GUICtrlRead($tInfo.IDFrom), $tRECT, BitOR($DT_CENTER, $DT_VCENTER))
            ;_WinAPI_DrawThemeTextEx($hTheme, 0, 0, $tInfo.hDC, GUICtrlRead($tInfo.IDFrom), $tRECT, BitOR($DT_CENTER, $DT_SINGLELINE, $DT_VCENTER), $tDTTOPTS)
            ;_WinAPI_CloseThemeData($hTheme)
		EndSwitch
	EndIf

    Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_NOTIFY2

Func _removeExStyles()
    ; remove WS_EX_COMPOSITED from GUI
    Local $i_ExStyle_Old = _WinAPI_GetWindowLong_mod($hParentFrame, $GWL_EXSTYLE)
    _WinAPI_SetWindowLong_mod($hParentFrame, $GWL_EXSTYLE, BitXOR($i_ExStyle_Old, $WS_EX_COMPOSITED))

    ; add LVS_EX_DOUBLEBUFFER to ListView
    _GUICtrlListView_SetExtendedListViewStyle($g_hListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_DOUBLEBUFFER, $LVS_EX_TRACKSELECT))
EndFunc

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
EndFunc

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
EndFunc

;==============================================
Func _resizeLVCols() ; resize listview columns to match header widths (1st display only, before any horizontal scrolling)
    For $i = 0 To _GUICtrlHeader_GetItemCount($g_hHeader) - 1
        _GUICtrlListView_SetColumnWidth($g_hListview, $i, _GUICtrlHeader_GetItemWidth($g_hHeader, $i))
    Next

    ; In case column 0 got an icon, retrieve the width of the icon
    Local $aRectLV = _GUICtrlListView_GetItemRect($g_hListView, 0, $LVIR_ICON) ; bounding rectangle of the icon (if any)
    $g_iIconWidth = $aRectLV[2] - $aRectLV[0] ; without icon : 4 - 4 => 0 (tested, the famous "4" !)
                                              ; with icon of 20 pixels : 24 - 4 = 20
EndFunc   ;==>_resizeLVCols

;==============================================
Func _resizeLVCols2() ; called while a header item is tracked or a divider is double-clicked. Also called while the listview is scrolled horizontally.
    Local $iCol, $aRectLV
    Local $aOrder = _GUICtrlHeader_GetOrderArray($g_hHeader)
    $iCol = $aOrder[1] ; left column (may not be column 0, if column 0 was dragged/dropped elsewhere)
    If $iCol > 0 Then ; LV subitem
        $aRectLV = _GUICtrlListView_GetSubItemRect($g_hListView, 0, $iCol)
    Else ; column 0 needs _GUICtrlListView_GetItemRect()
        $aRectLV = _GUICtrlListView_GetItemRect($g_hListView, 0, $LVIR_LABEL) ; bounding rectangle of the item text
        $aRectLV[0] -= (4 + $g_iIconWidth) ; adjust LV col 0 left coord (+++)
    EndIf
    If $aRectLV[0] < 0 Then ; horizontal scrollbar is NOT at left => move and resize the detached header (mimic a normal listview)
        ;WinMove($g_hHeader, "", $aRectLV[0], 0, WinGetClientSize($g_hChild)[0] - $aRectLV[0], Default)
        WinMove($g_hHeader, "", $aRectLV[0], 0, WinGetPos($g_hChild)[2] - $aRectLV[0], Default)
    Else ; horizontal scrollbar is at left => move and resize the detached header to its initial coords & size
        ;WinMove($g_hHeader, "", 0, 0, WinGetClientSize($g_hChild)[0], Default)
        WinMove($g_hHeader, "", 0, 0, WinGetPos($g_hChild)[2], Default)
    EndIf
EndFunc   ;==>_resizeLVCols2

;==============================================
Func _reorderLVCols()
    ; remove LVS_NOCOLUMNHEADER from listview
    Local $i_Style_Old = _WinAPI_GetWindowLong_mod($g_hListView, $GWL_STYLE)
    _WinAPI_SetWindowLong_mod($g_hListView, $GWL_STYLE, BitXOR($i_Style_Old, $LVS_NOCOLUMNHEADER))

    ; reorder listview columns order to match header items order
    Local $aOrder = _GUICtrlHeader_GetOrderArray($g_hHeader)
    _GUICtrlListView_SetColumnOrderArray($g_hListview, $aOrder)

    ; add LVS_NOCOLUMNHEADER back to listview
    _WinAPI_SetWindowLong_mod($g_hListView, $GWL_STYLE, $i_Style_Old)

EndFunc   ;==>_reorderLVCols

;Function for getting HWND from PID
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
EndFunc;==>_GetHwndFromPID

Func WM_COMMAND2($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg

    Local $iCode = BitShift($wParam, 16)
    Switch $lParam
        Case GUICtrlGetHandle($idInputPath)
            Switch $iCode
                Case $EN_SETFOCUS
                    ; select all text in path input box
                    $bPathSelectAll = True
                Case $EN_CHANGE
                    $bPathInputChanged = True
            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc  ;==>WM_COMMAND2

Func _About()
    Local $sMsg
    $sMsg = "Version: " & @TAB & @TAB & $sVersion & @CRLF & @CRLF
    $sMsg &= "Created by: " & @TAB & "AutoIt Community"
	MsgBox(0, "Files Au3", $sMsg)
EndFunc

Func _CleanExit()
    __TreeListExplorer_Shutdown()
    ;_GUICtrlListView_UnRegisterSortCallBack($g_hListview)
    _GUICtrlHeader_Destroy($g_hHeader)
    GUIDelete($g_hGUI)
    _ClearDarkSizebox()

    ; reset COLOR_MENU system color
    _WinAPI_SetSysColors($COLOR_MENU, $iColorOld)

    DllClose($hKernel32)
    DllClose($hGdi32)
    DllClose($hShlwapi)
    ;DllClose($hUser32)
EndFunc

Func _ClearDarkSizebox()
    _GDIPlus_BitmapDispose($g_hDots)
    _WinAPI_DestroyCursor($hCursor)
    _WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, $g_hOldProc)
    DllCallbackFree($hProc)
    _GDIPlus_Shutdown()
EndFunc

Func _InitDarkSizebox()
        ;-----------------
    ; Create a sizebox window (Scrollbar class) BEFORE creating the StatusBar control
    $g_hSizebox = _WinAPI_CreateWindowEx(0, "Scrollbar", "", $WS_CHILD + $WS_VISIBLE + $SBS_SIZEBOX, _
    0, 0, 0, 0, $g_hGUI) ; $SBS_SIZEBOX or $SBS_SIZEGRIP

    ; Subclass the sizebox (by changing the window procedure associated with the Scrollbar class)
    $hProc = DllCallbackRegister('ScrollbarProc', 'lresult', 'hwnd;uint;wparam;lparam')
    $g_hOldProc = _WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, DllCallbackGetPtr($hProc))

    $hCursor = _WinAPI_LoadCursor(0, $OCR_SIZENWSE)
    _WinAPI_SetClassLongEx($g_hSizebox, -12, $hCursor) ; $GCL_HCURSOR = -12
    $g_hBrush = _WinAPI_CreateSolidBrush($iBackColorDef)

    ; Sizebox height with DPI
    $g_iHeight = (16 * $iDPI) + 2
    $g_hDots = CreateDots($g_iHeight, $g_iHeight, 0x00000000 + $iBackColorDef, 0xFF000000 + 0xBFBFBF)

EndFunc

Func __Timer_QueryPerformanceFrequency_mod()
	Local $aCall = DllCall($hKernel32, "bool", "QueryPerformanceFrequency", "int64*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	Return SetExtended($aCall[0], $aCall[1])
EndFunc   ;==>__Timer_QueryPerformanceFrequency_mod

Func __Timer_QueryPerformanceCounter_mod()
	Local $aCall = DllCall($hKernel32, "bool", "QueryPerformanceCounter", "int64*", 0)
	If @error Then Return SetError(@error, @extended, -1)
	Return SetExtended($aCall[0], $aCall[1])
EndFunc   ;==>__Timer_QueryPerformanceCounter_mod

Func _Timer_Diff_mod($iTimeStamp)
	Return 1000 * (__Timer_QueryPerformanceCounter_mod() - $iTimeStamp) / __Timer_QueryPerformanceFrequency_mod()
EndFunc   ;==>_Timer_Diff_mod

Func _Timer_Init_mod()
	Return __Timer_QueryPerformanceCounter_mod()
EndFunc   ;==>_Timer_Init_mod

Func _WinAPI_ReleaseDC_mod($hWnd, $hDC)
	Local $aCall = DllCall($hUser32, "int", "ReleaseDC", "hwnd", $hWnd, "handle", $hDC)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_ReleaseDC_mod

Func _WinAPI_GetDCEx_mod($hWnd, $hRgn, $iFlags)
	Local $aCall = DllCall($hUser32, 'handle', 'GetDCEx', 'hwnd', $hWnd, 'handle', $hRgn, 'dword', $iFlags)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_GetDCEx_mod

Func _WinAPI_CreateRectRgn_mod($iLeftRect, $iTopRect, $iRightRect, $iBottomRect)
	Local $aCall = DllCall($hGdi32, "handle", "CreateRectRgn", "int", $iLeftRect, "int", $iTopRect, "int", $iRightRect, _
			"int", $iBottomRect)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_CreateRectRgn_mod

Func _WinAPI_OffsetRect_mod(ByRef $tRECT, $iDX, $iDY)
	Local $aCall = DllCall($hUser32, 'bool', 'OffsetRect', 'struct*', $tRECT, 'int', $iDX, 'int', $iDY)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_OffsetRect_mod

Func _WinAPI_GetWindowRect_mod($hWnd)
	Local $tRECT = DllStructCreate($tagRECT)
	Local $aCall = DllCall($hUser32, "bool", "GetWindowRect", "hwnd", $hWnd, "struct*", $tRECT)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tRECT
EndFunc   ;==>_WinAPI_GetWindowRect_mod

Func _WinAPI_ShellGetFileInfo_mod($sFilePath, $iFlags, $iAttributes, ByRef $tSHFILEINFO)
	Local $aCall = DllCall($hShell32, 'dword_ptr', 'SHGetFileInfoW', 'wstr', $sFilePath, 'dword', $iAttributes, _
			'struct*', $tSHFILEINFO, 'uint', DllStructGetSize($tSHFILEINFO), 'uint', $iFlags)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_ShellGetFileInfo_mod

Func _WinAPI_GetClientRect_mod($hWnd)
	Local $tRECT = DllStructCreate($tagRECT)
	Local $aCall = DllCall($hUser32, "bool", "GetClientRect", "hwnd", $hWnd, "struct*", $tRECT)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tRECT
EndFunc   ;==>_WinAPI_GetClientRect_mod

Func _WinAPI_GetWindowLong_mod($hWnd, $iIndex)
	Local $sFuncName = "GetWindowLongW"
	If @AutoItX64 Then $sFuncName = "GetWindowLongPtrW"
	Local $aCall = DllCall($hUser32, "long_ptr", $sFuncName, "hwnd", $hWnd, "int", $iIndex)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_GetWindowLong_mod

Func _WinAPI_DefWindowProc_mod($hWnd, $iMsg, $wParam, $lParam)
	Local $aCall = DllCall($hUser32, "lresult", "DefWindowProc", "hwnd", $hWnd, "uint", $iMsg, "wparam", $wParam, _
			"lparam", $lParam)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_DefWindowProc_mod

Func _WinAPI_SetLastError_mod($iErrorCode, Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
	DllCall($hKernel32, "none", "SetLastError", "dword", $iErrorCode)
	Return SetError($_iCallerError, $_iCallerExtended, Null)
EndFunc   ;==>_WinAPI_SetLastError_mod

Func _WinAPI_SetWindowLong_mod($hWnd, $iIndex, $iValue)
	_WinAPI_SetLastError_mod(0) ; as suggested in MSDN
	Local $sFuncName = "SetWindowLongW"
	If @AutoItX64 Then $sFuncName = "SetWindowLongPtrW"
	Local $aCall = DllCall($hUser32, "long_ptr", $sFuncName, "hwnd", $hWnd, "int", $iIndex, "long_ptr", $iValue)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_SetWindowLong_mod

Func _WinAPI_SetWindowPos_mod($hWnd, $hAfter, $iX, $iY, $iCX, $iCY, $iFlags)
	Local $aCall = DllCall($hUser32, "bool", "SetWindowPos", "hwnd", $hWnd, "hwnd", $hAfter, "int", $iX, "int", $iY, _
			"int", $iCX, "int", $iCY, "uint", $iFlags)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_SetWindowPos_mod

Func _WinAPI_DrawText_mod($hDC, $sText, ByRef $tRECT, $iFlags)
	Local $aCall = DllCall($hUser32, "int", "DrawTextW", "handle", $hDC, "wstr", $sText, "int", -1, "struct*", $tRECT, _
			"uint", $iFlags)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_DrawText_mod

Func _WinAPI_SetBkColor_mod($hDC, $iColor)
	Local $aCall = DllCall($hGdi32, "INT", "SetBkColor", "handle", $hDC, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_SetBkColor_mod

Func _WinAPI_DeleteObject_mod($hObject)
	Local $aCall = DllCall($hGdi32, "bool", "DeleteObject", "handle", $hObject)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_DeleteObject_mod

Func _WinAPI_InflateRect_mod(ByRef $tRECT, $iDX, $iDY)
	Local $aCall = DllCall($hUser32, 'bool', 'InflateRect', 'struct*', $tRECT, 'int', $iDX, 'int', $iDY)
	If @error Then Return SetError(@error, @extended, False)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_InflateRect_mod

Func _WinAPI_FillRect_mod($hDC, $tRECT, $hBrush)
	Local $aCall
	If IsPtr($hBrush) Then
		$aCall = DllCall($hUser32, "int", "FillRect", "handle", $hDC, "struct*", $tRECT, "handle", $hBrush)
	Else
		$aCall = DllCall($hUser32, "int", "FillRect", "handle", $hDC, "struct*", $tRECT, "dword_ptr", $hBrush)
	EndIf
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_FillRect_mod

Func _WinAPI_CreateSolidBrush_mod($iColor)
	Local $aCall = DllCall($hGdi32, "handle", "CreateSolidBrush", "INT", $iColor)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_CreateSolidBrush_mod

Func _WinAPI_GetClassName_mod($hWnd)
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $aCall = DllCall($hUser32, "int", "GetClassNameW", "hwnd", $hWnd, "wstr", "", "int", 4096)
	If @error Or Not $aCall[0] Then Return SetError(@error, @extended, '')

	Return SetExtended($aCall[0], $aCall[2])
EndFunc   ;==>_WinAPI_GetClassName_mod

Func _WinAPI_SetTextColor_mod($hDC, $iColor)
	Local $aCall = DllCall($hGdi32, "INT", "SetTextColor", "handle", $hDC, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_SetTextColor_mod

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
    Local $aSelectedLV = _GUICtrlListView_GetSelectedIndices($idListview, True)
    If $aSelectedLV[0] = 1 Then
        _WinAPI_ShellObjectProperties($sCurrentPath)
    ElseIf $aSelectedLV[0] = 0 Then
        _WinAPI_ShellObjectProperties(GUICtrlRead($idInputPath))
    Else
        ;$sSelectedItems
        Local $aFiles = StringSplit($sSelectedItems, "|")
        _ArrayDelete($aFiles, $aFiles[0])
        _ArrayDelete($aFiles, 0)
        _WinAPI_SHMultiFileProperties(GUICtrlRead($idInputPath), $aFiles)
    EndIf
EndFunc

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
                DLLStructSetData($aPIDL, $i+1, $PIDLChild)
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
        $PIDL = DllStructGetData($aNames, $i+1)
        If $PIDL Then _WinAPI_CoTaskMemFree(DllStructGetData($aNames, $i+1))
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
    ;Author: Rasim
        $hFont = _WinAPI_CreateFont($iHeight, 0, 0, 0, $iWeight, BitAND($iFontAtrributes, 2), BitAND($iFontAtrributes, 4), _
                                    BitAND($iFontAtrributes, 8), $DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, _
                                    $DEFAULT_QUALITY, 0, $sFontName)

        _SendMessage($hWnd, $WM_SETFONT, $hFont, 1)
EndFunc ;==>_GUICtrl_SetFont

Func WM_DRAWITEM2($hWnd, $Msg, $wParam, $lParam)
    #forceref $Msg, $wParam, $lParam

    ; modernmenuraw
    WM_DRAWITEM($hWnd, $Msg, $wParam, $lParam)

    Local $tDRAWITEMSTRUCT = DllStructCreate("uint CtlType;uint CtlID;uint itemID;uint itemAction;uint itemState;HWND hwndItem;HANDLE hDC;long rcItem[4];ULONG_PTR itemData", $lParam)

    If DllStructGetData($tDRAWITEMSTRUCT, "hwndItem") <> $g_hStatus Then Return $GUI_RUNDEFMSG ; Only process the statusbar

    Local $itemID = DllStructGetData($tDRAWITEMSTRUCT, "itemID") ;part number
    Local $hDC = DllStructGetData($tDRAWITEMSTRUCT, "hDC")
    Local $tRect = DllStructCreate("long left;long top;long right; long bottom", DllStructGetPtr($tDRAWITEMSTRUCT, "rcItem"))
    Local $iTop = DllStructGetData($tRect, "top")
    Local $iLeft = DllStructGetData($tRect, "left")
    Local $hBrush

    $hBrush = _WinAPI_CreateSolidBrush($iBackColorDef) ; Background Color
    _WinAPI_FillRect($hDC, DllStructGetPtr($tRect), $hBrush)
    _WinAPI_SetTextColor($hDC, $iTextColorDef) ; Font Color
    _WinAPI_SetBkMode($hDC, $TRANSPARENT)
    DllStructSetData($tRect, "top", $iTop + 1)
    DllStructSetData($tRect, "left", $iLeft + 1)
    _WinAPI_DrawText($hDC, $g_aText[$itemID], $tRect, $DT_LEFT)
    _WinAPI_DeleteObject($hBrush)

    $tDRAWITEMSTRUCT = 0

    _WinAPI_RedrawWindow($g_hSizebox)

    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_DRAWITEM2

Func _WinAPI_PathIsRoot_mod($sFilePath)
	Local $aCall = DllCall($hShlwapi, 'bool', 'PathIsRootW', 'wstr', $sFilePath & "\")
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>_WinAPI_PathIsRoot_mod

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
            $iDotSpace = $iDotSize - 0.5    ; gives some control over the spacing between dots
            $iDotFrame = 0.5                ; gives some control over the spacing from frame
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

    Local $a[6][2] = [[3,7], [3,5], [3,3], [5,5], [5,3], [7,3]]
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
                _WinAPI_RedrawWindow($g_hGui)
                $bSizeboxOffScreen = False
            EndIf
        EndIf
        ;_WinAPI_RedrawWindow($g_hSizebox)
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_MOVE

Func ApplyDPI()
    ; apply System DPI awareness and calculate factor
    _WinAPI_SetThreadDpiAwarenessContext($DPI_AWARENESS_CONTEXT_SYSTEM_AWARE)
    If Not @error Then
        $iDPI2 = Round(_WinAPI_GetDpiForSystem() / 96, 2)
        Return $iDPI2
        If @error Then Return $iDPI2 = 1
    Else
        Return $iDPI2 = 1
    EndIf
EndFunc

Func _WinAPI_SetThreadDpiAwarenessContext($DPI_AWARENESS_CONTEXT_value) ; UEZ
    Local $aResult = DllCall("user32.dll", "uint", "SetThreadDpiAwarenessContext", @AutoItX64 ? "int64" : "int", $DPI_AWARENESS_CONTEXT_value) ;requires Win10 v1703+ / Windows Server 2016+
    If Not IsArray($aResult) Or @error Then Return SetError(1, @extended, 0)
    If Not $aResult[0] Then Return SetError(2, @extended, 0)
    Return $aResult[0]
EndFunc   ;==>_WinAPI_SetThreadDpiAwarenessContext

Func _WinAPI_GetDpiForSystem() ; UEZ
    Local $aResult = DllCall("user32.dll", "uint", "GetDpiForSystem") ;requires Win10 v1607+ / no server support
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
        _WinAPI_SetWindowTheme($g_hListView, 'DarkMode_Explorer')
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
    Else
        _GUISetDarkTheme($g_hGUI, False)
        _GUISetDarkTheme(_GUIFrame_GetHandle($iFrame_A, 1), False)
        _GUISetDarkTheme(_GUIFrame_GetHandle($iFrame_A, 2), False)
        GUICtrlSetBkColor($idListview, $iBackColorDef)
        GUICtrlSetBkColor($idTreeView, $iBackColorDef)
        GUICtrlSetColor($idListview, $iTextColorDef)
        GUICtrlSetColor($idTreeView, $iTextColorDef)
        _WinAPI_SetWindowTheme($g_hListView, 'Explorer')
        _WinAPI_SetWindowTheme($g_hTreeView, 'Explorer')
        _WinAPI_SetWindowTheme($g_hStatus, 'Explorer')
        _WinAPI_SetWindowTheme($g_hHeader, 'ItemsView', 'Header')
        _WinAPI_SetWindowTheme($g_hInputPath, 'Explorer')
        GUICtrlSetBkColor($idInputPath, 0xFFFFFF)
        GUICtrlSetColor($idInputPath, $iTextColorDef)
        GUICtrlSetBkColor($idSeparator, 0x909090)
    EndIf
EndFunc

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
	Local $iPreferredAppMode = ($bEnableDarkTheme == True) ? $APPMODE_FORCEDARK : $APPMODE_FORCELIGHT
	Local $iGUI_BkColor = $iBackColorDef
	_WinAPI_SetPreferredAppMode($iPreferredAppMode)
	_WinAPI_RefreshImmersiveColorPolicyState()
	_WinAPI_FlushMenuThemes()
	GUISetBkColor($iGUI_BkColor, $hWnd)
	;_GUICtrlSetDarkTheme($hWnd, $bEnableDarkTheme)            ;To Color the GUI's own Scrollbar
;~ 	DllCall('dwmapi.dll', 'long', 'DwmSetWindowAttribute', 'hwnd', $hWnd, 'dword', $DWMWA_USE_IMMERSIVE_DARK_MODE, 'dword*', Int($bEnableDarkTheme), 'dword', 4)
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
