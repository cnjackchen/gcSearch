#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=gcSearch.ico
#AutoIt3Wrapper_Res_Comment=ChemStation文件搜索工具
#AutoIt3Wrapper_Res_Description=ChemStation文件搜索工具
#AutoIt3Wrapper_Res_Fileversion=2.6.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Jack Chen <jack.chen@iff.com>
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-q
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs
	AutoIt Version: 3.3.14.2
	作者:        Jack Chen
	网站:        http://hi.baidu.com/jdchenjian
	脚本说明：   本脚本可从 Agilent ChemStation 生成的.ch文件中提取样品名、方法、日期等信息，
	加上文件路径信息生成index文件，可通过样品名、方法等关键词在索引中搜索色谱数据信息。

	不同版本 Agilent ChemStation 数据文件格式:
	HP / Agilent ChemStation data file (*.ch, *.ms, *.uv)
	Read only format. Chromatographic signal from Agilent / HP ChemStation. Each run, each detector - one separate file.
	Contains raw instrument data and dataset information. The method information and integration parameters are in separate files.
	Versions supported:
	2 - GC/MS Data;
	Read the TIC and associated mass-spectra.
	30 - ADC Data, LC Data
	31 - UV Spectrum data;
	8, 81 - GC Data (A.xx Chemstation)
	179 - GC Data (B.03 Chemstation)
	180, 181 - GC Data (B.04 Chemstation)
	130 - ADC Data, LC Data (B.04 Chemstation)
	131 - UV-Spectrum Data (B.04 Chemstation)
	ChemStation Data file(*.ch)格式:
	;~ 索引文件
	; 版本  文件开头             SampleName       Operator     DateTime       Method
	; A.xx  8/81/30/31           0025-1           0149-1       0179-1         0229-1
	; B.xx  179/180/181/130/131  0859-1           1881-1       2392-1         2575-1
#ce

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <GuiStatusBar.au3>
#include <Date.au3>
#include <GuiDateTimePicker.au3>
#include <ComboConstants.au3>
#include <SQLite.au3>

AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("TrayAutoPause", 0) ; Script will not pause when click on tray icon.
AutoItSetOption("TrayOnEventMode", 1) ; Enable OnEvent functions notifications for the tray
AutoItSetOption("TrayMenuMode", 1 + 2) ; Default tray menu items (Script Paused/Exit) will not be shown.
AutoItSetOption("WinTitleMatchMode", 4)
AutoItSetOption("GUIOnEventMode", 1)
AutoItSetOption("ExpandVarStrings", 1) ; use variables and macros inside strings, e.g., "The value of var1 is $var1$".
AutoItSetOption("GUIResizeMode", $GUI_DOCKALL)

Global Const $AppName = "gcSearch"
Global Const $AppVer = "2.6" ; 程序版本
Global Const $STDChromDir = @ScriptDir & "\STDChrom" ; 标准谱图文件夹
Global Const $SettingsFile = @ScriptDir & "\" & $AppName & ".ini" ; 配置文件
Global Const $IndexCacheFile = @ScriptDir & "\Cache.db"
Global $IndexFile, $IndexCached, $Data_MethDirs
Global $hSetAsSTD, $hCompareWithSTD, $hCompare, $IsRebuilding = False
Global $hCompareWithRef, $hSetAsRef, $RefDataFile, $RefMethod, $RefInfo
Global $__hListView_Editable ; 可编辑的列表ID

Global $MenuID_OtherMeth ; 右键菜单其它积分方法(MenuID - 方法路径)
Global $NewCHFiles ; 暂存新.ch文件路径

Global $aSearchResult ; 搜索结果
Global $B_DESCENDING[7] ; 列表排序用数组
Global $aKeyWords ; _Search() 关键词
Global $hIndexDirList, $hIndexSettingsRebuild, $RecentDays = 2
Global $IndexDirs, $AutoIndex, $AnalMethDir, $RunOnStart, $DataMethFirst, $IniFile, $DataPaths, $InstNames, $MethPaths, $DirMap

Global $Instruments, $hMainGUI, $RunOnStartItem, $hDataMethFirst
Global $MainContextMenu, $a_Date[7], $hDate1, $hDate2, $hSampleInfo, $hSampleName, $hMethod, $hDataFile, $hSearch, $hMainList
Global $hLoadMenu, $hDataFileMenu, $hStatus, $hLoadWithMethMenu
;Global $CMDCtrlEdit = "Edit1"
Global $CMDCtrlEdit = "WindowsForms10.EDIT.app.0.378734a1"


; 避免软件的多个实例同时运行
Local $list, $i
$list = WinList("[REGEXPTITLE:(?i)^" & $AppName & "; REGEXPCLASS:(?i)AutoIt *v3 *GUI]")
For $i = 1 To $list[0][0]
	If Not BitAND(WinGetState($list[$i][1], ""), 2) Then
		WinSetState($list[$i][1], "", @SW_SHOW)
	EndIf
	WinActivate($list[$i][1])
	Exit
Next

FileChangeDir(@ScriptDir)
FileInstall("sqlite3.dll", "sqlite3.dll")
_SQLite_Startup()
If Not FileExists($SettingsFile) Then
	Local $file = FileOpen($SettingsFile, 2 + 32) ; Unicode UTF16 Little Endian格式，解决部分英文系统中文乱码问题
	FileClose($file)
	IniWriteSection($SettingsFile, "Settings", _
			"RunOnStart=1" & @LF & _
			"Dirs=" & @LF & _
			"IndexFile=" & @LF & _
			"AutoIndex=1" & @LF & _
			"AnalMethDir=" & @LF & _
			"DataMethFirst=0" & @LF & _
			"Data_MethDirs=" & @LF & _
			"DirMap=")
EndIf
$RunOnStart = IniRead($SettingsFile, "Settings", "RunOnStart", 1) * 1 ; 开机启动
$IndexDirs = IniRead($SettingsFile, "Settings", "Dirs", "") ; 目录,用于建index/监控
$IndexFile = IniRead($SettingsFile, "Settings", "IndexFile", "") ; 索引文件
$AutoIndex = IniRead($SettingsFile, "Settings", "AutoIndex", 1) * 1 ; 自动索引新文件
$AnalMethDir = IniRead($SettingsFile, "Settings", "AnalMethDir", "") ; 用于数据分析的方法文件夹
$DataMethFirst = IniRead($SettingsFile, "Settings", "DataMethFirst", 0) * 1 ; DA.M 优先
$Data_MethDirs = IniRead($SettingsFile, "Settings", "Data_MethDirs", "") ; 为数据文件目录指定默认的方法目录
$DirMap = IniRead($SettingsFile, "Settings", "DirMap", "") ; 路径转换
If $IndexFile = "" Then
	$IndexFile = @ScriptDir & "\index.db"
EndIf
If DriveGetType($IndexFile) = "Network" Then ; 索引文件不在本机，则在本地建立缓存
	CacheIndexDb()
	AdlibRegister("CacheIndexDb", 30000) ; 定时更新索引缓存
Else
	FileDelete($IndexCacheFile)
	$IndexCached = False
EndIf
If Not FileExists($STDChromDir) Then
	DirCreate($STDChromDir) ; 建立标准谱图文件夹
EndIf

; 从 Windows 目录下 ini文件中读取 _DataPath$，InstNames等ChemStation配置信息
; Rev A.xx 系列配置信息在win.ini文件中
; Rev B.xx 系列配置信息在chemstation.ini文件中
$Instruments = IniRead(@WindowsDir & "\CHEMSTATION.ini", "PCS", "Instruments", "") ; Instruments=1,2,3...
If $Instruments <> "" Then
	$IniFile = @WindowsDir & "\CHEMSTATION.ini"
Else
	$Instruments = IniRead(@WindowsDir & "\win.ini", "PCS", "Instruments", "")
	If $Instruments <> "" Then
		$IniFile = @WindowsDir & "\win.ini"
	EndIf
EndIf
Local $NewDirs
If $IniFile <> "" Then
	GetChemIni($IniFile, $DataPaths, $InstNames, $MethPaths) ; 读取ChemStation配置文件信息（_DATAPATH$-->DataPath，InstName-->InstNames）
	$NewDirs = $IndexDirs & "|" & $DataPaths
	$NewDirs = CheckDirs($NewDirs) ; 去掉重叠或不存在的目录
	If $IndexDirs <> $NewDirs Then
		$IndexDirs = $NewDirs
		IniWrite($SettingsFile, "Settings", "Dirs", $IndexDirs)
	EndIf
EndIf

;~ GUI主界面
$hMainGUI = GUICreate($AppName & " v" & $AppVer & " - Agilent ChemStation 数据文件搜索工具", 900, 580, -1, -1, $WS_OVERLAPPEDWINDOW)
GUISetOnEvent($GUI_EVENT_CLOSE, "OnEvent_Close", $hMainGUI)
GUISetOnEvent($GUI_EVENT_MINIMIZE, "OnEvent_Min", $hMainGUI)

;~ 菜单
Local $FileMenu = GUICtrlCreateMenu("文件")
GUICtrlCreateMenuItem("调用标准谱图文件", $FileMenu)
GUICtrlSetOnEvent(-1, "_LoadSTD")
GUICtrlCreateMenuItem("打开标准谱图文件夹", $FileMenu)
GUICtrlSetOnEvent(-1, "_OpenSTDChromDir")
GUICtrlCreateMenuItem("退出...", $FileMenu)
GUICtrlSetOnEvent(-1, "_ShutApp")
Local $SettingsMenu = GUICtrlCreateMenu("设置")
$RunOnStartItem = GUICtrlCreateMenuItem("开机自动运行", $SettingsMenu)
GUICtrlSetOnEvent(-1, "_RunOnStart")
GUICtrlCreateMenuItem("设置...", $SettingsMenu)
GUICtrlSetOnEvent(-1, "_ShowSettings")
Local $HelpMenu = GUICtrlCreateMenu("帮助")
GUICtrlCreateMenuItem("帮助", $HelpMenu)
GUICtrlSetOnEvent(-1, "_Help")
GUICtrlCreateMenuItem("主页", $HelpMenu)
GUICtrlSetOnEvent(-1, "_Website")
GUICtrlCreateMenuItem("关于", $HelpMenu)
GUICtrlSetOnEvent(-1, "_About")
;~ 托盘菜单
TraySetClick(16) ; Releasing secondary mouse button
TrayCreateItem($AppName)
TrayItemSetOnEvent(-1, "_ShowMainGUI")
TrayItemSetState(-1, 512) ; $TRAY_DEFAULT = -1
TrayCreateItem("关于")
TrayItemSetOnEvent(-1, "_About")
TrayCreateItem("退出...")
TrayItemSetOnEvent(-1, "_ShutApp")
TraySetToolTip($AppName & " for ChemStation")
TraySetOnEvent(-7, "_ShowMainGUI") ; $TRAY_EVENT_PRIMARYDOWN = -7
;~ 主界面
GUICtrlCreateLabel("开始日期：", 10, 25, 70, 20)
$hDate1 = GUICtrlGetHandle(GUICtrlCreateDate("", 80, 20, 125, 20, BitOR($DTS_SHOWNONE, $WS_TABSTOP)))
_GUICtrlDTP_SetFormat($hDate1, "yyyy年MM月dd日")
GUICtrlSetTip(-1, "只搜索此日期之后的数据")
GUICtrlCreateLabel("结束日期：", 215, 25, 70, 20, BitOR($GUI_SS_DEFAULT_LABEL, $SS_CENTER))
$hDate2 = GUICtrlGetHandle(GUICtrlCreateDate("", 285, 20, 125, 20, BitOR($DTS_SHOWNONE, $WS_TABSTOP)))
_GUICtrlDTP_SetFormat($hDate2, "yyyy年MM月dd日")
GUICtrlSetTip(-1, "只搜索此日期之前的数据")
$a_Date[0] = True
_GUICtrlDTP_SetSystemTime($hDate1, $a_Date)
_GUICtrlDTP_SetSystemTime($hDate2, $a_Date)
GUICtrlCreateCombo("", 420, 20, 60, 20, $CBS_DROPDOWNLIST)
GUICtrlSetData(-1, "最近...|今天|1 天|2 天|3 天|4 天|7 天|30天|其它...", "最近...")
GUICtrlSetOnEvent(-1, "_SetDateRecent")
GUICtrlCreateLabel("样品信息：", 490, 25, 70, 20)
$hSampleInfo = GUICtrlCreateInput("", 560, 20, 110, 20)
GUICtrlSetTip(-1, "样品信息（备注）")
GUICtrlCreateButton("清除", 690, 20, 50, 20)
GUICtrlSetTip(-1, "清除输入信息和搜索结果")
GUICtrlSetOnEvent(-1, "_OnEventClear")
GUICtrlCreateLabel("样品名称：", 10, 60, 70, 20)
$hSampleName = GUICtrlCreateInput("", 80, 55, 125, 20)
GUICtrlSetTip(-1, "样品名称")
GUICtrlCreateLabel("分析方法：", 215, 60, 70, 20)
$hMethod = GUICtrlCreateInput("", 285, 55, 125, 20)
GUICtrlSetTip(-1, "分析方法")
GUICtrlCreateLabel("数据文件：", 420, 60, 70, 20)
$hDataFile = GUICtrlCreateInput("", 490, 55, 125, 20)
GUICtrlSetTip(-1, "数据文件路径")
$hSearch = GUICtrlCreateButton("立即搜索", 650, 52, 90, 25, $BS_DEFPUSHBUTTON)
GUICtrlSetTip(-1, "立即搜索")
GUICtrlSetOnEvent(-1, "_Search")
Global $tText = DllStructCreate("wchar Text[1024]") ; 建个结构，用来放listview列数据
$hMainList = GUICtrlCreateListView("样品名称|分析方法|分析日期|分析时间|数据文件|操作者|样品信息", _
		10, 95, 880, 430, BitOR($LVS_SHOWSELALWAYS, $LVS_SINGLESEL, $LVS_OWNERDATA), _
		BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_HEADERDRAGDROP, $LVS_EX_DOUBLEBUFFER))
GUICtrlSendMsg($hMainList, $LVM_SETITEMCOUNT, 0, 0) ; 分配列表内存。 为什么要这样做？因为虚拟列表必须要知道数据总量
GUICtrlSendMsg($hMainList, $LVM_SETCOLUMNWIDTH, 0, 100)
GUICtrlSendMsg($hMainList, $LVM_SETCOLUMNWIDTH, 1, 100)
GUICtrlSendMsg($hMainList, $LVM_SETCOLUMNWIDTH, 2, 90)
GUICtrlSendMsg($hMainList, $LVM_SETCOLUMNWIDTH, 3, 90)
GUICtrlSendMsg($hMainList, $LVM_SETCOLUMNWIDTH, 4, 280)
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
GUICtrlSetTip(-1, "双击在ChemStation中打开" & @CRLF & "右键单击显示选项")
GUICtrlSetOnEvent(-1, "_SortListView")

;~ 关联菜单
$MainContextMenu = GUICtrlCreateContextMenu($hMainList)
$hLoadMenu = GUICtrlCreateMenuItem("在ChemStation中打开", $MainContextMenu)
GUICtrlSetOnEvent(-1, "_LoadInChemStation")
GUICtrlSetState(-1, $GUI_DEFBUTTON)

;~ 用于积分的方法目录
If $AnalMethDir = "" Then
	Global $aMethDir = StringSplit($MethPaths, "|")
	For $i = 1 To $aMethDir[0]
		$aMethDir[$i] = StringRegExpReplace($aMethDir[$i], "\\$", "") ; 去掉结尾的"\"
		If FileExists($aMethDir[$i]) Then
			$AnalMethDir = StringRegExpReplace($aMethDir[$i], ".*\\", "") & "|" & $aMethDir[$i]
			IniWrite($SettingsFile, "Settings", "AnalMethDir", $AnalMethDir)
			ExitLoop
		EndIf
	Next
EndIf
_OtherMethMenu($AnalMethDir)

GUICtrlCreateMenuItem("", $MainContextMenu)

$hCompare = GUICtrlCreateMenuItem("比较谱图", $MainContextMenu)
GUICtrlSetOnEvent(-1, "_Compare")

GUICtrlCreateMenuItem("", $MainContextMenu)

$hCompareWithRef = GUICtrlCreateMenuItem('和“' & $RefInfo & '”比较', $MainContextMenu)
GUICtrlSetOnEvent(-1, "_CompareWithRef")
$hSetAsRef = GUICtrlCreateMenuItem("作为待比较谱图", $MainContextMenu)
GUICtrlSetOnEvent(-1, "_SetAsRef")

GUICtrlCreateMenuItem("", $MainContextMenu)

$hCompareWithSTD = GUICtrlCreateMenuItem("和标准谱图比较...", $MainContextMenu)
GUICtrlSetOnEvent(-1, "_CompareWithSTD")
$hSetAsSTD = GUICtrlCreateMenuItem("设为标准谱图...", $MainContextMenu)
GUICtrlSetOnEvent(-1, "_SetAsSTD")

GUICtrlCreateMenuItem("", $MainContextMenu)

$hDataFileMenu = GUICtrlCreateMenu("数据文件", $MainContextMenu)
GUICtrlCreateMenuItem("浏览数据文件", $hDataFileMenu)
GUICtrlSetOnEvent(-1, "_OpenFolder")
GUICtrlCreateMenuItem("修改文件信息...", $hDataFileMenu)
GUICtrlSetOnEvent(-1, "_EditFileInfo")


;~ 	状态栏
Global $aParts[3] = [130, 200, -1]
$hStatus = _GUICtrlStatusBar_Create($hMainGUI, -1, "", BitOR($SBARS_SIZEGRIP, $SBARS_TOOLTIPS))
_GUICtrlStatusBar_SetParts($hStatus, $aParts)
_GUICtrlStatusBar_SetTipText($hStatus, 0, "索引更新时间")
_GUICtrlStatusBar_SetTipText($hStatus, 1, "已索引文件数")
If Not FileExists($IndexFile) Then
	_GUICtrlStatusBar_SetText($hStatus, "没有找到索引文件，请先建立索引。", 2)
	GUICtrlSetState($hSearch, $GUI_DISABLE) ; 无index.db文件，使Search Now按钮无效
Else
	_ShowIndexStatus($IndexFile, $hStatus)
EndIf

_GUICtrlStatusBar_SetText($hStatus, "搜索安捷伦化学工作站色谱、电泳谱数据文件(*.ch)", 2)
GUIRegisterMsg($WM_SIZE, "WM_SIZE")

$NewDirs = CheckDirs($IndexDirs) ; 检查目录
If $IndexDirs <> $NewDirs Then
	$IndexDirs = $NewDirs
	IniWrite($SettingsFile, "Settings", "Dirs", $IndexDirs)
EndIf
Global $gaDirs
If $IndexDirs <> "" And $AutoIndex = 1 Then
	$gaDirs = StringSplit($IndexDirs, "|", 2)
	MonitorDirectory($gaDirs)
	AdlibRegister("IndexNewFile", 10000)
EndIf
_SetRunOnStart() ; 设置注册表-开机自动运行
If $CmdLine[0] = 0 Or $CmdLine[1] <> "-Hide" Then ; 开机自动运行快捷方式带－Hide参数，不显示主界面
	GUISetState(@SW_SHOW, $hMainGUI)
EndIf

GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
OnAutoItExitRegister("_OnExit")

If $IniFile <> "" Then
	AdlibRegister("_WatchIniFile", 2000)
EndIf

If Not FileExists($IndexFile) Then
	MsgBox(16, $AppName, '索引文件 "' & $IndexFile & '" 不存在！' & @CRLF & @CRLF & _
			'请重新设置索引文件路径，或者重建索引。', 0, $hMainGUI)
	GUI_Settings() ; 显示设置界面
EndIf

While 1
	Sleep(100)
WEnd
;~ ===========================================以上是主程序==========================================


;~ 缓存索引数据库到本地
Func CacheIndexDb()
	Local $t1, $t2
	$IndexCached = False
	If Not FileExists($IndexFile) Then Return

	$t1 = FileGetTime($IndexFile, 0, 1)
	$t2 = FileGetTime($IndexCacheFile, 0, 1)
	If $t1 = $t2 Or FileCopy($IndexFile, $IndexCacheFile, 1) Then
		$IndexCached = True
	EndIf
EndFunc   ;==>CacheIndexDb


;~ 显示、修改数据文件信息
Func _EditFileInfo()
	Global $FileInfoSelectedItem, $hFileInfoGUI, $hFileInfoDataFile, $hFileInfoSampleName, $hFileInfoMethod
	Global $hFileInfoDateTime, $hFileInfoOperator, $hFileInfoSampleInfo, $hReserveFileTime

	$FileInfoSelectedItem = _GUICtrlListView_GetNextItem($hMainList, -1, 0, 8)
	Local $item = _GUICtrlListView_GetItemTextArray($hMainList, $FileInfoSelectedItem) ; 读取选择的条目
	If $item[0] < 7 Then Return
	Local $filepath = $item[5]
	Local $search = FileFindFirstFile($filepath & "\*.ch")
	If $search = -1 Then Return
	Local $var, $FileInfoDataFile
	While 1
		$var = FileFindNextFile($search)
		If @error Then ExitLoop
		If @extended Then ContinueLoop
		If $FileInfoDataFile = "" Then
			$FileInfoDataFile = $filepath & "\" & $var
		Else
			$FileInfoDataFile &= "|" & $filepath & "\" & $var
		EndIf
	WEnd
	FileClose($search)
	If $FileInfoDataFile = "" Then Return
	If StringInStr($FileInfoDataFile, "|") Then ; 多个文件，让用户选择一个
		$FileInfoDataFile = FileOpenDialog("选择数据文件（*.ch）", $filepath, "数据文件 (*.ch)", 3, "", $hMainGUI)
		If @error Then Return
	EndIf

	Local $FileInfo = _ReadFileInfo($FileInfoDataFile)
	$hFileInfoGUI = GUICreate("数据文件信息", 500, 360, -1, -1, -1, -1, $hMainGUI)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_FileInfoClose")
	GUICtrlCreateLabel("文件路径", 10, 20, 60, 20)
	$hFileInfoDataFile = GUICtrlCreateInput($FileInfoDataFile, 70, 16, 420, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
	GUICtrlCreateLabel("样品名称", 10, 50, 60, 20)
	$hFileInfoSampleName = GUICtrlCreateInput($FileInfo[0], 70, 46, 420, 20, $ES_AUTOHSCROLL)
	GUICtrlCreateLabel("分析方法", 10, 80, 60, 20)
	$hFileInfoMethod = GUICtrlCreateInput($FileInfo[3], 70, 76, 420, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
	GUICtrlCreateLabel("分析时间", 10, 110, 60, 20)
	$hFileInfoDateTime = GUICtrlCreateInput($FileInfo[2], 70, 106, 420, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
	GUICtrlCreateLabel("操作者", 10, 140, 60, 20)
	$hFileInfoOperator = GUICtrlCreateInput($FileInfo[1], 70, 136, 420, 20)
	GUICtrlCreateLabel("样品信息", 10, 170, 60, 20)
	$hFileInfoSampleInfo = GUICtrlCreateInput($FileInfo[4], 70, 166, 420, 60, BitOR($ES_LEFT, $ES_AUTOHSCROLL, $ES_AUTOVSCROLL, $ES_MULTILINE, $ES_WANTRETURN))
	$hReserveFileTime = GUICtrlCreateCheckbox("保留原始文件的最后修改时间", 10, 240)
	GUICtrlSetTip(-1, "修改文件但不更新文件修改时间")
	GUICtrlCreateLabel("注：除样品信息外，其它信息长度不得超过255个字符。", 10, 280)

	GUICtrlCreateButton("恢复", 250, 320, 60, 20)
	GUICtrlSetTip(-1, "撤消所有后期修改，" & @CRLF & "恢复原始文件信息。")
	If Not FileExists($FileInfoDataFile & ".bak") Then GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlSetOnEvent(-1, "_FileInfoRestore")
	GUICtrlCreateButton("修改", 340, 320, 60, 20)
	GUICtrlSetTip(-1, "将修改后的信息写入数据文件")
	GUICtrlSetOnEvent(-1, "_FileInfoSave")
	GUICtrlCreateButton("退出", 430, 320, 60, 20)
	GUICtrlSetOnEvent(-1, "_FileInfoClose")
	GUISetState(@SW_SHOW)
EndFunc   ;==>_EditFileInfo

Func _FileInfoClose()
	GUIDelete(@GUI_WinHandle)
EndFunc   ;==>_FileInfoClose

Func _FileInfoSave()
	; 备份
	Local $FileInfoDataFile = GUICtrlRead($hFileInfoDataFile)
	If Not FileExists($FileInfoDataFile & ".bak") Then
		FileCopy($FileInfoDataFile, $FileInfoDataFile & ".bak") ; 备份 .ch 文件
	EndIf
	Local $MACFile = StringLeft($FileInfoDataFile, StringInStr($FileInfoDataFile, "\", 0, -1)) & "SAMPLE.MAC"
	If FileExists($MACFile) And Not FileExists($MACFile & ".bak") Then
		FileCopy($MACFile, $MACFile & ".bak") ; 备份 SAMPLE.MAC
	EndIf

	; 读取信息
	Local $FileInfoSampleName = GUICtrlRead($hFileInfoSampleName)
	If StringLen($FileInfoSampleName) > 255 Then
		$FileInfoSampleName = StringLeft($FileInfoSampleName, 255)
	EndIf
	Local $FileInfoMethod = GUICtrlRead($hFileInfoMethod)
	If StringLen($FileInfoMethod) > 255 Then
		$FileInfoMethod = StringLeft($FileInfoMethod, 255)
	EndIf
	Local $FileInfoDateTime = GUICtrlRead($hFileInfoDateTime)
	If StringLen($FileInfoDateTime) > 255 Then
		$FileInfoDateTime = StringLeft($FileInfoDateTime, 255)
	EndIf
	Local $FileInfoOperator = GUICtrlRead($hFileInfoOperator)
	If StringLen($FileInfoOperator) > 255 Then
		$FileInfoOperator = StringLeft($FileInfoOperator, 255)
	EndIf
	Local $FileInfoSampleInfo = GUICtrlRead($hFileInfoSampleInfo)
	Local $ReserveFileTime = GUICtrlRead($hReserveFileTime)
	Local $DataFileTime = FileGetTime($FileInfoDataFile, 0, 1)

	; 写入文件
	Local $file, $header
	$file = FileOpen($FileInfoDataFile, 1 + 16) ; Force binary mode
	$header = _ReadChars($file, 0, 1)
	; 按.ch文件的开头区分不同版本。ChemStation A.xx系列：
	If StringInStr('|8|81|30|31|', '|' & $header & '|') Then ; ANSI
		_WriteChars($file, $FileInfoSampleName, 24, 1) ; SampleName
		_WriteChars($file, $FileInfoOperator, 148, 1) ; Operator
		_WriteChars($file, $FileInfoDateTime, 178, 1) ; DateTime
		_WriteChars($file, $FileInfoMethod, 228, 1) ; Method
		; ChemStation B.xx, C.01.XX系列：
	Else ; If StringInStr('|179|180|181|130|131|', '|' & $header & '|') Then ; 每个字符占2个字节，UTF16 LE
		_WriteChars($file, $FileInfoSampleName, 858, 2) ; SampleName
		_WriteChars($file, $FileInfoOperator, 1880, 2) ; Operator
		_WriteChars($file, $FileInfoDateTime, 2391, 2) ; DateTime
		_WriteChars($file, $FileInfoMethod, 2574, 2) ; Method
	EndIf
	FileClose($file)
	If $ReserveFileTime = $GUI_CHECKED Then
		FileSetTime($FileInfoDataFile, $DataFileTime)
	EndIf

	; 样品信息写入SAMPLE.MAC
	Local $SpInfo, $aInfo, $i, $NewSpInfo, $MACFileTime
	If FileExists($MACFile) Then
		$MACFileTime = FileGetTime($MACFile, 0, 1)
		$SpInfo = FileRead($MACFile)
		$SpInfo = StringRegExpReplace($SpInfo, '(?i)RPTSAMPLEINFO ".*"' & @CRLF, '')
		$FileInfoSampleInfo = StringRegExpReplace($FileInfoSampleInfo, "[" & @CR & @LF & "]+", @CRLF) ; 删除空行
		If $FileInfoSampleInfo <> "" Then
			$aInfo = StringSplit($FileInfoSampleInfo, @CRLF, 1)
			For $i = 1 To $aInfo[0]
				$NewSpInfo &= 'RPTSAMPLEINFO "' & $aInfo[$i] & '"' & @CRLF
			Next
			$NewSpInfo = $NewSpInfo & 'REMOVE SAMPLEINFO'
			$SpInfo = StringReplace($SpInfo, 'REMOVE SAMPLEINFO', $NewSpInfo)
		EndIf
		FileDelete($MACFile)
		Local $file = FileOpen($MACFile, 2 + 32) ; Unicode UTF16 Little Endian
		FileWrite($file, $SpInfo)
		FileClose($file)
		If $ReserveFileTime = $GUI_CHECKED Then
			FileSetTime($MACFile, $MACFileTime)
		EndIf
	EndIf

	Local $aIndex = _IndexFile($FileInfoDataFile)
	;$aIndex[7] = [$SampleName, $Method, $Date, $Time, $ChFile, $Operator, $SampleInfo]
	Local $hDb = _SQLite_Open($IndexFile)
	_SQLite_Exec($hDb, "REPLACE INTO Index1 VALUES ('" & $aIndex[0] & "','" & $aIndex[1] & "','" & _
			$aIndex[2] & "','" & $aIndex[3] & "','" & $aIndex[4] & "','" & $aIndex[5] & "','" & $aIndex[6] & "');")
	_SQLite_Close($hDb)
	For $i = 0 To UBound($aIndex) - 1
		$aSearchResult[$FileInfoSelectedItem + 1][$i] = $aIndex[$i]
	Next
	GUICtrlSendMsg($hMainList, $LVM_REDRAWITEMS, 0, 20)
	GUIDelete(@GUI_WinHandle)
EndFunc   ;==>_FileInfoSave

Func _FileInfoRestore()
	Local $FileInfoDataFile = GUICtrlRead($hFileInfoDataFile)
	If FileExists($FileInfoDataFile & ".bak") Then
		FileMove($FileInfoDataFile & ".bak", $FileInfoDataFile, 1)
	EndIf
	Local $MACFile = StringLeft($FileInfoDataFile, StringInStr($FileInfoDataFile, "\", 0, -1)) & "SAMPLE.MAC"
	If FileExists($MACFile & ".bak") Then
		FileMove($MACFile & ".bak", $MACFile, 1)
	EndIf

	Local $hDb, $aIndex
	$aIndex = _IndexFile($FileInfoDataFile)
	; $aIndex[7] = [$SampleName, $Method, $Date, $Time, $ChFile, $Operator, $SampleInfo]
	$hDb = _SQLite_Open($IndexFile)
	_SQLite_Exec($hDb, "REPLACE INTO Index1 VALUES ('" & $aIndex[0] & "','" & $aIndex[1] & "','" & _
			$aIndex[2] & "','" & $aIndex[3] & "','" & $aIndex[4] & "','" & $aIndex[5] & "','" & $aIndex[6] & "');")
	_SQLite_Close($hDb)
	For $i = 0 To UBound($aIndex) - 1
		$aSearchResult[$FileInfoSelectedItem + 1][$i] = $aIndex[$i]
	Next
	GUICtrlSendMsg($hMainList, $LVM_REDRAWITEMS, 0, 20)
	GUIDelete(@GUI_WinHandle)
	_EditFileInfo()
EndFunc   ;==>_FileInfoRestore

; 将信息写入文件
;~ $file - handle， $Pos - offset, $flag - StringToBinary() 转换标识 1- ANSI, 2 - UTF16 LE, 3 - UTF16 - BE, 4 - UTF8
Func _WriteChars($file, $str, $Pos, $flag)
	Local $len, $bytes, $data
	If $flag = 1 Or $flag = 4 Then
		$bytes = 1 ; 每个字符的字节数
	Else
		$bytes = 2
	EndIf
	FileSetPos($file, $Pos, 0)
	$len = FileRead($file, 1)
	$len = Number($len) * $bytes
	For $i = 1 To $len
		$data &= "00"
	Next
	FileWrite($file, Binary("0x" & $data)) ; 删除原有信息
	$len = StringLen($str)
	FileSetPos($file, $Pos, 0)
	FileWrite($file, BinaryMid($len, 1, 1))
	FileWrite($file, StringToBinary($str, $flag))
EndFunc   ;==>_WriteChars

; 版本  文件开头             SampleName       Operator     DateTime       Method
; A.xx  8/81/30/31           0025-1           0149-1       0179-1         0229-1
; B.03  179/180/181/130/131  0859-1           1881-1       2392-1         2575-1
Func _ReadFileInfo($ChFile)
	Local $file, $header, $MACFile, $SpInfo, $i, $match
	Local $Info[5] ; SampleName, Operator, DateTime, Method, SampleInfo
	$file = FileOpen($ChFile, 16) ; Force binary mode
	$header = _ReadChars($file, 0, 1)
	; 按.ch文件的开头区分不同版本。ChemStation A.xx系列：
	If StringInStr('|8|81|30|31|', '|' & $header & '|') Then ; ANSI
		$Info[0] = _ReadChars($file, 24, 1) ; SampleName
		$Info[1] = _ReadChars($file, 148, 1) ; Operator
		$Info[2] = _ReadChars($file, 178, 1) ; DateTime
		$Info[3] = _ReadChars($file, 228, 1) ; Method
	; ChemStation B.xx, C.01.XX系列：
	Else ; If StringInStr('|179|180|181|130|131|', '|' & $header & '|') Then ; 每个字符占2个字节，UTF16 LE
		$Info[0] = _ReadChars($file, 858, 2) ; SampleName
		$Info[1] = _ReadChars($file, 1880, 2) ; Operator
		$Info[2] = _ReadChars($file, 2391, 2) ; DateTime
		$Info[3] = _ReadChars($file, 2574, 2) ; Method
	EndIf
	FileClose($file)
	; 读取SAMPLE.MAC中的样品信息
	$MACFile = StringLeft($ChFile, StringInStr($ChFile, "\", 0, -1)) & "SAMPLE.MAC"
	$SpInfo = FileRead($MACFile)
	$match = StringRegExp($SpInfo, '(?m)^(?i)RPTSAMPLEINFO "(.*?)"', 3)
	For $i = 0 To UBound($match) - 1
		$Info[4] = $Info[4] & $match[$i] & @CRLF
	Next
	$Info[4] = StringRegExpReplace($Info[4], @CRLF & "$", "")
	Return $Info
EndFunc   ;==>_ReadFileInfo


Func _ShowIndexStatus($IdxFile, $hStatus)
	Local $t[6], $hDb, $hQuery, $sOut
	Local $IndexTime, $IndexCount = 0
	If FileExists($IdxFile) Then
		$t = FileGetTime($IdxFile) ; 索引文件更新日期
		$IndexTime = $t[0] & "-" & $t[1] & "-" & $t[2] & " " & $t[3] & ":" & $t[4] & ":" & $t[5]
		$hDb = _SQLite_Open($IdxFile, $SQLITE_OPEN_READONLY) ; 只读方式打开
		_SQLite_Query($hDb, "SELECT COUNT(*) FROM Index1;", $hQuery)
		_SQLite_FetchData($hQuery, $sOut)
		_SQLite_QueryFinalize($hQuery)
		_SQLite_Close($hDb)
		$IndexCount = Number($sOut[0])
	EndIf
	_GUICtrlStatusBar_SetText($hStatus, $IndexTime & "                    ", 0)
	_GUICtrlStatusBar_SetText($hStatus, $IndexCount & "                    ", 1)
EndFunc   ;==>_ShowIndexStatus


;~ 其它积分方法菜单
Func _OtherMethMenu($mDir)
	Local $fname, $MenuItemID, $n, $i, $j, $SubMenu, $SubMenuID

	GUICtrlDelete($hLoadWithMethMenu)
	$hLoadWithMethMenu = GUICtrlCreateMenu("用其它方法积分", $MainContextMenu, 1)
	$MenuID_OtherMeth = ""

	Local $arr, $brr
	If $mDir <> "" Then
		$arr = StringSplit($mDir, "||", 1)
		$n = $arr[0]
		Dim $aDirs[$n][2]
		For $i = 1 To $arr[0]
			$brr = StringSplit($arr[$i], "|", 1)
			If $brr[0] < 2 Then
				; gcSearch 2.5之前版本如 AnalMethDir=C:\Chem32\1\METHODS 没有"|"分隔，会出错
				$aDirs[$i - 1][0] = "方法"
				$aDirs[$i - 1][1] = $brr[1]
			Else
				$aDirs[$i - 1][0] = $brr[1]
				$aDirs[$i - 1][1] = $brr[2]
			EndIf
		Next
	EndIf

	If $n > 0 Then
		For $i = 0 To $n - 1
			If $n > 1 Then
				$SubMenu = $aDirs[$i][0]
				$SubMenuID = GUICtrlCreateMenu($SubMenu, $hLoadWithMethMenu)
			Else
				$SubMenuID = $hLoadWithMethMenu
			EndIf
			If StringRight($aDirs[$i][1], 1) = "\" Then $aDirs[$i][1] = StringTrimRight($aDirs[$i][1], 1)
			$fname = _FileListToArray($aDirs[$i][1], "*.M", 2)
			If @error Then ContinueLoop
			_ArraySort($fname, 0, 1)
			; _ArrayDisplay($fname)
			For $j = 1 To $fname[0]
				$MenuItemID = GUICtrlCreateMenuItem($fname[$j], $SubMenuID)
				GUICtrlSetOnEvent(-1, "_LoadWithMeth")
				$MenuID_OtherMeth &= $MenuItemID & "|" & $aDirs[$i][1] & @CRLF
			Next
		Next
	EndIf
	$MenuID_OtherMeth = StringTrimRight($MenuID_OtherMeth, 2) ; 去掉最后的换行符
	GUICtrlCreateMenuItem("设定方法目录...", $hLoadWithMethMenu)
	GUICtrlSetOnEvent(-1, "_MethSettings")
EndFunc   ;==>_OtherMethMenu

Func _MethSettings()
	GUI_Settings(1)
EndFunc   ;==>_MethSettings

;~ 用其它方法打开
Func _LoadWithMeth()
	Local $file, $method, $ID
	Local $item = _GUICtrlListView_GetItemTextArray($hMainList, -1) ; 读取选择的条目
	If $item[0] < 7 Then Return
	$method = GUICtrlRead(@GUI_CtrlId, 1)
	$file = $item[5]
	If Not FileExists($file) Then
		_GUICtrlStatusBar_SetText($hStatus, "数据文件 " & $file & " 不存在！", 2)
		Return
	EndIf

	Local $arr = StringSplit($MenuID_OtherMeth, @CRLF, 1)
	Local $idx = _ArraySearch($arr, @GUI_CtrlId & "|", 1, 0, 0, 1)
	If @error Then Return
	$arr = StringSplit($arr[$idx], "|")
	If $arr[0] < 2 Then Return

	$ID = _GetChemID() ;~ 取得 ChemStation (Offline) 窗口 ID
	If Not $ID Then
		_MsgChemStationNotRunning()
		Return
	EndIf

	WinActivate($ID)
	; 加载方法
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadMethod_DAOnly "' & $arr[2] & '\","' & $method & '"')
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")

	Sleep(20)
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadFile "' & $file & '"') ; 调用谱图
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
EndFunc   ;==>_LoadWithMeth

;~ 打开标准谱图数据文件夹
Func _OpenSTDChromDir()
	ShellExecute($STDChromDir)
EndFunc   ;==>_OpenSTDChromDir


; Resize the status bar when GUI size changes
Func WM_SIZE($hWnd, $Msg, $wParam, $lParam)
	_GUICtrlStatusBar_Resize($hStatus)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_SIZE

; 关闭窗口，隐藏
Func OnEvent_Close()
	WinSetState($hMainGUI, "", @SW_HIDE)
	ReduceMemory()
EndFunc   ;==>OnEvent_Close

;~ 窗口最小化时整理内存
Func OnEvent_Min()
	ReduceMemory()
EndFunc   ;==>OnEvent_Min


; 退出程序
Func _ShutApp()
	Local $a = MsgBox(4 + 32 + 256, $AppName, "退出后新的 ChemStation 数据文件将不会被索引。" & @CRLF & @CRLF & _
			"确定要退出程序吗？", 60, $hMainGUI)
	If $a = 6 Then Exit
EndFunc   ;==>_ShutApp

;~ 退出前执行
Func _OnExit()
	GUIRegisterMsg($WM_NOTIFY, "")
	GUIRegisterMsg($WM_SIZE, "")
	MonitorDirectory()
	FileDelete($IndexCacheFile)
	_SQLite_Shutdown()
EndFunc   ;==>_OnExit

;~ 帮助
Func _Help()
	If FileExists(@ScriptDir & "\gcSearch.pdf") Then ; 打开pdf帮助文件
		ShellExecute(@ScriptDir & "\gcSearch.pdf", "", "", "open")
		If Not @error Then Return
	EndIf
	Local $string = $AppName & ' 用于搜索 HP/Agilent ChemStation、OpenLab Chemstation Edition 的数据文件，'
	$string &= '可在 Windows 2000以上操作系统中运行。'
	$string &= @CRLF & @CRLF & '软件只有一个文件（' & $AppName & '.exe），可Copy到电脑中任意位置运行。'
	$string &= '第一次运行时，会出现索引设置界面。请将保存数据文件的文件夹添加到列表中，'
	$string &= '选中“重建索引”，按“确定”后开始扫描、索引数据文件。仅第一次运行或修改索引目录后需重建索引，平时软件在后台运行，'
	$string &= '能自动索引新的数据文件。建立索引后就可以搜索了。在搜索界面中输入日期、方法、样品名称、样品信息等，'
	$string &= '按回车或“立即搜索”，瞬间即可得到搜索结果。在搜索结果列表中双击相应的条目，'
	$string &= '可在ChemStation（Offline）中调用该色谱文件。右键单击相应条目，通过右键菜单可实现在化学工作站中显示、'
	$string &= '比较色谱图等功能。关闭主界面后，软件最小化到通知区域在后台运行，实时地监视、索引新的数据文件。'
	$string &= @CRLF & @CRLF & '在后台运行时，占用的系统资源非常小，不会影响电脑的使用。单击通知区域的放大镜图标可调出主界面；'
	$string &= '右键单击图标会弹出右键菜单，选“退出”可完全退出程序。注意：退出程序后新的数据文件将不会被索引，'
	$string &= '因此也就不能被搜索到了。'
	$string &= @CRLF & @CRLF & '要卸载软件，先将菜单“设置”->“开机自动运行”前的勾去掉，然后直接删除文件即可。'
	; 通常MsgBox会让整个程序暂停运行。启动另一进程显示MsgBox可以避免这个问题
	Run(@AutoItExe & ' /AutoIt3ExecuteLine  "MsgBox(64, ''' & $AppName & ' 帮助' & _ ; <- Msgbox 的 TITLE
			''', ''' & $string & ''')"') ; Msgbox 的 Text
EndFunc   ;==>_Help

;~  软件主页
Func _Website()
	ShellExecute("https://github.com/cnjackchen/gcSearch")
EndFunc   ;==>_Website

;~ 关于
Func _About()
	Run(@AutoItExe & ' /AutoIt3ExecuteLine  "MsgBox(64, ''' & _
			'关于 ' & $AppName & ''', ''' & _ ; 以下为 Msgbox 的 Text
			'版本：' & $AppVer & @CRLF & @CRLF & _
			'作者：Jack Chen <jack.chen@iff.com>' & ''')"')
EndFunc   ;==>_About

; 显示搜索窗口
Func _ShowMainGUI()
	WinSetState($hMainGUI, "", @SW_SHOW)
	WinActivate($hMainGUI)
EndFunc   ;==>_ShowMainGUI

; 重建索引
Func _Rebuild()
	Local $ans = MsgBox(4 + 64, $AppName, "重建索引可能需要几分钟时间。" & @CRLF & @CRLF & "确定要重建索引吗？", 0, $hMainGUI)
	If $ans <> 6 Then Return ; not YES
	AdlibRegister("_DoRebuild", 10) ; 尽快返回，以Timer来启动重建索引。否则在索引期间其它按键不起作用
EndFunc   ;==>_Rebuild

; 重建索引
Func _DoRebuild()
	Local $t, $hIndexDb, $i
	AdlibUnRegister("_DoRebuild")
	$t = TimerInit()
	_GUICtrlStatusBar_SetText($hStatus, "正在重建索引，请稍候...", 2)
	$IsRebuilding = True ; 标志
	GUICtrlSetState($hSearch, $GUI_DISABLE) ; 暂时禁止搜索
	FileDelete($IndexFile)
	$hIndexDb = _SQLite_Open($IndexFile)
	_SQLite_Exec($hIndexDb, "CREATE TABLE Index1 (样品名称 TEXT,分析方法 TEXT,分析日期 TEXT," & _
			"分析时间 TEXT,数据文件 TEXT UNIQUE,操作者 TEXT,样品信息 TEXT);")
	_SQLite_Exec($hIndexDb, "BEGIN TRANSACTION;")
	Local $IndexCount = 0
	Local $aDirs = StringSplit($IndexDirs, "|")
	For $i = 1 To $aDirs[0]
		If FileExists($aDirs[$i]) Then
			_IndexDir($aDirs[$i], $hIndexDb, $IndexCount)
		EndIf
	Next
	_SQLite_Exec($hIndexDb, "COMMIT;")
	_SQLite_Close($hIndexDb)

	$t = Round(TimerDiff($t) / 60000, 2)
	_ShowIndexStatus($IndexFile, $hStatus)
	_GUICtrlStatusBar_SetText($hStatus, "索引重建完毕，用时 " & $t & " 分钟。", 2)
	$IsRebuilding = False
	GUICtrlSetState($hSearch, $GUI_ENABLE)
	ReduceMemory()
EndFunc   ;==>_DoRebuild

Func _IndexDir($sDir, $hIndexDb, ByRef $IndexCount)
	Local $iCnt = 0, $search, $fname, $aIndex
	Local $aDir = StringSplit($sDir, "|", 2)
	Do
		If StringRight($aDir[$iCnt], 1) = "\" Then $aDir[$iCnt] = StringTrimRight($aDir[$iCnt], 1)
		$search = FileFindFirstFile($aDir[$iCnt] & "\*")
		If $search <> -1 Then
			While 1
				$fname = FileFindNextFile($search)
				If @error = 1 Then ExitLoop
				If @extended = 1 And $fname <> "." And $fname <> ".." Then ; 如果是文件夹则加入列表
					$sDir &= "|" & $aDir[$iCnt] & "\" & $fname
				EndIf
			WEnd
			FileClose($search)
		EndIf
		$search = FileFindFirstFile($aDir[$iCnt] & "\*.ch")
		If $search <> -1 Then
			While 1
				$fname = FileFindNextFile($search)
				If @error = 1 Then ExitLoop
				If @extended <> 1 Then
					$aIndex = _IndexFile($aDir[$iCnt] & "\" & $fname)
					_SQLite_Exec($hIndexDb, "INSERT INTO Index1 VALUES ('" & $aIndex[0] & "','" & $aIndex[1] & "','" & _
							$aIndex[2] & "','" & $aIndex[3] & "','" & $aIndex[4] & "','" & $aIndex[5] & "','" & $aIndex[6] & "');")
					$IndexCount += 1
					If Mod($IndexCount, 10) = 0 Then
						_GUICtrlStatusBar_SetText($hStatus, "@YEAR@-@MON@-@MDAY@ @HOUR@:@MIN@:@SEC@" & "                    ", 0)
						_GUICtrlStatusBar_SetText($hStatus, $IndexCount & "                    ", 1)
						Sleep(1)
					EndIf
				EndIf
			WEnd
			FileClose($search)
		EndIf
		$iCnt += 1
		If UBound($aDir) <= $iCnt Then $aDir = StringSplit($sDir, "|", 2)
	Until UBound($aDir) <= $iCnt
EndFunc   ;==>_IndexDir

;~ 索引文件
; 版本  文件开头             SampleName       Operator     DateTime       Method
; A.xx  8/81/30/31           0025-1           0149-1       0179-1         0229-1
; B.03  179/180/181/130/131  0859-1           1881-1       2392-1         2575-1
Func _IndexFile($ChFile)
	Local $file, $header, $MACFile, $SpInfo, $i
	Local $SampleName, $Operator, $DateTime, $SampleInfo, $method, $match, $Date, $Time
	$file = FileOpen($ChFile, 16) ; Force binary mode
	$header = _ReadChars($file, 0, 1)
	; 按.ch文件的开头区分不同版本。ChemStation A.xx系列：
	If StringInStr('|8|81|30|31|', '|' & $header & '|') Then
		$SampleName = _ReadChars($file, 24, 1) ; ANSI
		$Operator = _ReadChars($file, 148, 1)
		$DateTime = _ReadChars($file, 178, 1)
		$method = _ReadChars($file, 228, 1)
	; ChemStation B.xx, C.01.xx系列：
	Else ; If StringInStr('|179|180|181|130|131|', '|' & $header & '|') Then
		$SampleName = _ReadChars($file, 858, 2) ; 每个字符占2个字节，UTF16 LE
		$Operator = _ReadChars($file, 1880, 2)
		$DateTime = _ReadChars($file, 2391, 2)
		$method = _ReadChars($file, 2574, 2)
	EndIf
	FileClose($file)
	; 分解日期、时间
	$match = StringRegExp($DateTime, '([^,]*?),? +(\d{1,2}:\d{1,2}.*)', 2)
	If Not @error Then
		$Date = $match[1]
		$Time = $match[2]
		$Date = _DateParse($Date)
	EndIf
	; 读取SAMPLE.MAC中的样品信息
	$MACFile = StringLeft($ChFile, StringInStr($ChFile, "\", 0, -1)) & "SAMPLE.MAC"
	$SpInfo = FileRead($MACFile)
	$match = StringRegExp($SpInfo, '(?m)^(?i)RPTSAMPLEINFO "(.*?)"', 3)
	For $i = 0 To UBound($match) - 1
		$SampleInfo = $SampleInfo & $match[$i] & ";"
	Next
	If StringRight($SampleInfo, 1) = ";" Then
		$SampleInfo = StringTrimRight($SampleInfo, 1)
	EndIf
	; $Index = $SampleName & "|" & $Method & "|" & $Date & "|" & $Time & "|" & $ChFile & "|" & $Operator & "|" & $SampleInfo
	Local $aIndex[7] = [$SampleName, $method, $Date, $Time, $ChFile, $Operator, $SampleInfo]
	Return $aIndex
EndFunc   ;==>_IndexFile

; 读取某一偏移量的字符串
;~ $file - handle， $Pos - offset, $flag - BinaryToString() 转换标识 1- ANSI, 2 - UTF16 LE, 3 - UTF16 - BE, 4 - UTF8
Func _ReadChars($file, $Pos = 0, $flag = 2)
	Local $len, $str, $bytes
	If $flag = 1 Or $flag = 4 Then
		$bytes = 1 ; 每个字符的字节数
	Else
		$bytes = 2
	EndIf
	FileSetPos($file, $Pos, 0)
	$len = FileRead($file, 1)
	$len = Number($len) * $bytes
	$str = FileRead($file, $len)
	Return BinaryToString($str, $flag)
EndFunc   ;==>_ReadChars

;~ 函数。将日期转换为 yyyy-MM-dd 标准格式
#cs
	类型1: 月份用字母表示
	dd-MMM-yy  26-Oct-09  -->  2009-10-26
	dd MMM yy  17 Mar 94  -->  1994-05-17
	类型2：用“-”分隔
	yy-M-d     10-5-25  -->  2010-5-25
	类型3：用“.”分隔
	dd.MM.yy   23.03.06  -->  2006-03-23
	类型4：用“/”分隔
	M/d/yyyy    5/8/2010   -->   2010-05-08
	M/d/yy  8/28/95  --> 1995-08-28
#ce
;~ 例：
;~ MsgBox(0, "", "26-Oct-09 --> " & _DateParse("26-Oct-09") & @CRLF & "17 Mar 94 --> " & _
;~ 		_DateParse("17 Mar 94") & @CRLF & "10-5-25 --> " & _DateParse("10-5-25") & @CRLF & "23.03.06 --> " & _
;~ 		_DateParse("23.03.06") & @CRLF & "8/28/95 --> " & _DateParse("8/28/95") & @CRLF & "5/8/2010 --> " & _DateParse("5/8/2010"))
Func _DateParse($str)
	Local $Year, $Month, $Day, $YMD, $match
	$str = StringStripWS($str, 3) ; 去掉开头、结尾的空字符
	; 类型1
	If StringRegExp($str, "^(\d{1,2})(?:-| )(\D{3,})(?:-| )(\d{2,4})$", 0) Then
		$match = StringRegExp($str, "^(\d{1,2})(?:-| )(\D{3,})(?:-| )(\d{2,4})$", 2)
		$Year = $match[3]
		$Month = $match[2]
		$Day = $match[1]
		; 类型2
	ElseIf StringRegExp($str, "^(\d{2,4})-(\d{1,2})-(\d{1,2})$", 0) Then
		$match = StringRegExp($str, "^(\d{2,4})-(\d{1,2})-(\d{1,2})$", 2)
		$Year = $match[1]
		$Month = $match[2]
		$Day = $match[3]
		; 类型3
	ElseIf StringRegExp($str, "^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$", 0) Then
		$match = StringRegExp($str, "^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$", 2)
		$Year = $match[3]
		$Month = $match[2]
		$Day = $match[1]
		; 类型4
	ElseIf StringRegExp($str, "^(\d{1,2})/(\d{1,2})/(\d{2,4})$", 0) Then
		$match = StringRegExp($str, "^(\d{1,2})/(\d{1,2})/(\d{2,4})$", 2)
		$Year = $match[3]
		$Month = $match[1]
		$Day = $match[2]
	EndIf
	; 将字母式月份转换为数字。
	Local $MonthStr = "Jan01Feb02Mar03Apr04May05Jun06Jul07Aug08Sep09Oct10Nov11Dec12"
	If StringRegExp($Month, "^\D\D\D", 0) Then
		$match = StringRegExp($Month, "^\D\D\D", 2)
		$Month = StringMid($MonthStr, StringInStr($MonthStr, $match[0]) + 3, 2)
	EndIf
	; 若年份不足4位数，转换为4位
	If StringLen($Year) < 4 Then
		If $Year > 70 Then
			$Year = "19" & $Year
		Else
			$Year = "20" & $Year
		EndIf
	EndIf
	; 若月份不足2位，月份前加0
	If StringLen($Month) < 2 Then
		$Month = "0" & $Month
	EndIf
	; 若日不足2位数，日前加0
	If StringLen($Day) < 2 Then
		$Day = "0" & $Day
	EndIf
	$YMD = $Year & "-" & $Month & "-" & $Day
	Return $YMD
EndFunc   ;==>_DateParse


; 搜索
Func _Search()

	Local $a_Date[6]
	Local $SampleInfo, $SampleName, $method, $DataFile
	Local $sDate1, $sDate2
	GUICtrlSetState($hSearch, $GUI_DISABLE)
	$SampleName = GUICtrlRead($hSampleName)
	$SampleInfo = GUICtrlRead($hSampleInfo)
	$method = GUICtrlRead($hMethod)
	$DataFile = GUICtrlRead($hDataFile)
	$a_Date = _GUICtrlDTP_GetSystemTime($hDate1)
	If @error = 0 And $a_Date[0] <> 0 Then
		$sDate1 = StringFormat("%04d-%02d-%02d", $a_Date[0], $a_Date[1], $a_Date[2])
	Else
		$sDate1 = "1970-01-01"
	EndIf
	$a_Date = _GUICtrlDTP_GetSystemTime($hDate2)
	If @error = 0 And $a_Date[0] <> 0 Then
		$sDate2 = StringFormat("%04d-%02d-%02d", $a_Date[0], $a_Date[1], $a_Date[2])
	Else
		$sDate2 = @YEAR & "-" & @MON & "-" & @MDAY
	EndIf
	If $sDate1 > $sDate2 Then
		_GUICtrlStatusBar_SetText($hStatus, "日期错误！结束日期不能早于开始日期。", 2)
		GUICtrlSetState($hSearch, $GUI_ENABLE)
		Return
	EndIf

	$SampleName = _KeyWords($SampleName)
	$SampleInfo = _KeyWords($SampleInfo)
	$method = _KeyWords($method)
	$DataFile = _KeyWords($DataFile)

	_GUICtrlStatusBar_SetText($hStatus, "正在搜索，请稍候 ...", 2)

	#cs ====内存数据库搜索测试====
		$hDb = _SQLite_Open()
		_SQLite_Exec(-1, "ATTACH DATABASE '" & $IndexFile & "' AS TempDB")
		_SQLite_Exec(-1, "CREATE TABLE Index1 AS SELECT * FROM TempDB.Index1")
		_SQLite_Exec(-1, "DETACH DATABASE TempDB")
	#ce

	Local $file
	If $IndexCached Then ; 搜索缓存
		$file = $IndexCacheFile
	Else ; 直接搜索数据库
		$file = $IndexFile
	EndIf
	_ShowIndexStatus($file, $hStatus)

	Local $t = TimerInit()
	Local $hDb = _SQLite_Open($file, $SQLITE_OPEN_READONLY)
	Local $iRows, $iColumns, $iITEM_COUNT = 0
	Local $iRval = _SQLite_GetTable2d($hDb, "SELECT * FROM Index1 WHERE" & _
			" 样品名称 LIKE '" & $SampleName & "'" & _
			" AND 分析方法 LIKE '" & $method & "'" & _
			" AND 分析日期 BETWEEN '" & $sDate1 & "' AND '" & $sDate2 & "'" & _
			" AND 数据文件 LIKE '" & $DataFile & "'" & _
			" AND 样品信息 LIKE '" & $SampleInfo & "'" & _
			" ORDER BY 分析日期,分析时间 LIMIT 1000", _ ; 按时间排序，限制搜索结果500条
			$aSearchResult, $iRows, $iColumns)
	If $iRval = $SQLITE_OK Then
		$iITEM_COUNT = UBound($aSearchResult) - 1
		GUICtrlSendMsg($hMainList, $LVM_SETITEMCOUNT, $iITEM_COUNT, 0)
		GUICtrlSendMsg($hMainList, $LVM_ENSUREVISIBLE, 0, 0)
		GUICtrlSendMsg($hMainList, $LVM_REDRAWITEMS, 0, 20)
	EndIf
	_SQLite_Close($hDb)
	$t = Round(TimerDiff($t)) / 1000

	If $iITEM_COUNT < 1000 Then
		_GUICtrlStatusBar_SetText($hStatus, "找到符合条件的结果 " & $iITEM_COUNT & " 条，" & "用时 " & $t & " 秒。", 2)
	Else
		_GUICtrlStatusBar_SetText($hStatus, "符合条件的结果 > 1000 条（已显示前 1000 条），" & "用时 " & $t & _
				" 秒。部分结果未显示，请修改关键词后重新搜索。", 2)
	EndIf

	GUICtrlSetState($hSearch, $GUI_ENABLE)
	Local $KeyWords = $SampleName & "|" & $method & "|" & $sDate1 & "|" & $sDate2 & "|" & $DataFile & "|" & $SampleInfo
	$KeyWords = StringReplace($KeyWords, "%", ".*")
	$aKeyWords = StringSplit($KeyWords, "|", 2)
	Dim $B_DESCENDING[7] ; 排序数组初始化
	ReduceMemory()
EndFunc   ;==>_Search

; 路径转换
Func _DirMap($dir)
	If $DirMap = "" Then Return $dir
	Local $i, $arr1, $arr2
	$arr1 = StringSplit($DirMap, "||", 1)
	For $i = 1 To $arr1[0]
		$arr2 = StringSplit($arr1[$i], "|", 1)
		If StringInStr($dir, $arr2[1]) = 1 Then
			$dir = StringReplace($dir, $arr2[1], $arr2[2], 1)
			ExitLoop
		EndIf
	Next
	Return $dir
EndFunc   ;==>_DirMap

;~ 列表排序
Func _SortListView()
	Local $iCol = GUICtrlGetState(@GUI_CtrlId)
	_ArraySort($aSearchResult, $B_DESCENDING[$iCol], 1, 0, $iCol)
	$B_DESCENDING[$iCol] = Not $B_DESCENDING[$iCol]
	GUICtrlSendMsg(@GUI_CtrlId, $LVM_ENSUREVISIBLE, 0, 0)
	GUICtrlSendMsg(@GUI_CtrlId, $LVM_REDRAWITEMS, 0, 20)
EndFunc   ;==>_SortListView

Func _LVEditableOff($hListView = "")
	If $hListView = "" Then
		$__hListView_Editable = ""
	Else
		$__hListView_Editable = StringReplace($__hListView_Editable, "|" & $hListView & "|", "")
	EndIf
;~ 	If $__hListView_Editable = "" Then
;~ 		GUIRegisterMsg($WM_NOTIFY, "")
;~ 	EndIf
EndFunc   ;==>_LVEditableOff

Func _LVEditableOn($hListView)
	$__hListView_Editable &= "|" & $hListView & "|"
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
EndFunc   ;==>_LVEditableOn

Func WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
	#forceref $hWnd, $Msg, $wParam
	Local $tNMHDR = DllStructCreate($tagNMITEMACTIVATE, $lParam)
	Local $IDFrom = DllStructGetData($tNMHDR, 'IDFrom')
	Local $Code = DllStructGetData($tNMHDR, 'Code')
	Local $Index = DllStructGetData($tNMHDR, 'Index')
	If $IDFrom = $hMainList Then ; ===============搜索结果列表=========================
		Switch $Code
			Case $NM_DBLCLK ; 双击
				If $Index <> -1 Then
					_LoadinChemStation()
				EndIf
			Case $NM_RCLICK ; 右键单击
				If $Index = -1 Then
					GUICtrlSetState($hLoadMenu, $GUI_DISABLE)
					GUICtrlSetState($hLoadWithMethMenu, $GUI_DISABLE)
					GUICtrlSetState($hCompare, $GUI_DISABLE)
					GUICtrlSetState($hCompareWithRef, $GUI_DISABLE)
					GUICtrlSetState($hSetAsRef, $GUI_DISABLE)
					GUICtrlSetState($hCompareWithSTD, $GUI_DISABLE)
					GUICtrlSetState($hSetAsSTD, $GUI_DISABLE)
					GUICtrlSetState($hDataFileMenu, $GUI_DISABLE)
				Else
					GUICtrlSetState($hLoadMenu, $GUI_ENABLE)
					_OtherMethMenu($AnalMethDir)
					GUICtrlSetState($hCompare, $GUI_ENABLE)
					GUICtrlSetState($hCompareWithRef, $GUI_ENABLE)
					GUICtrlSetState($hSetAsRef, $GUI_ENABLE)
					GUICtrlSetState($hCompareWithSTD, $GUI_ENABLE)
					GUICtrlSetState($hSetAsSTD, $GUI_ENABLE)
					GUICtrlSetState($hDataFileMenu, $GUI_ENABLE)
					GUICtrlSetState($hLoadMenu, $GUI_DEFBUTTON)
				EndIf
			Case -150, -177 ; $LVN_GETDISPINFOA = -150, $LVN_GETDISPINFOW = -177 -----更新虚拟列表-----
				If Not IsArray($aSearchResult) Then ContinueCase
				Local $tInfo = DllStructCreate($tagNMLVDISPINFO, $lParam)
				Local $iIndex = Int(DllStructGetData($tInfo, "Item"))
				Local $iSub = Int(DllStructGetData($tInfo, "SubItem"))
				Local $s = $aSearchResult[$iIndex + 1][$iSub]
				If $iSub = 4 Then
					$s = StringLeft($s, StringInStr($s, "\", 0, -1) - 1)
					If $DirMap <> "" Then
						$s = _DirMap($s) ; 路径转换
					EndIf
				EndIf
				DllStructSetData($tText, "Text", $s);列数据放入$tText结构
				DllStructSetData($tInfo, "Text", DllStructGetPtr($tText));用$tText结构的指针来设置列数据
				DllStructSetData($tInfo, "TextMax", StringLen($s));设置列数据长度
		EndSwitch
	ElseIf StringInStr($__hListView_Editable, "|" & $IDFrom & "|") Then ; ================修改Listview的内容===================
		Switch $Code
			Case $NM_DBLCLK
				If $Index = -1 Then
					Return
				EndIf
				Local $SubItem = DllStructGetData($tNMHDR, 'SubItem')
				Local $sText = _GUICtrlListView_GetItemText($IDFrom, $Index, $SubItem)
				$sText = InputBox("修改条目", "请输入新的内容", $sText, "", 280, 140, Default, Default, Default, $hWnd)
				If Not @error Then
					_GUICtrlListView_SetItemText($IDFrom, $Index, $sText, $SubItem)
				EndIf
		EndSwitch
	EndIf
	Return $GUI_RUNDEFMSG ; allow the default processing
EndFunc   ;==>WM_NOTIFY

;~ 取得 ChemStation (Offline) 窗口 ID
Func _GetChemID()
	Local $ChID = WinGetHandle("[REGEXPCLASS:(?i)WindowsForms10.Window.8.app.0.378734a; REGEXPTITLE:(?i)(?:Offline|脱机)]")
	Return $ChID
EndFunc   ;==>_GetChemID

Func _MsgChemStationNotRunning()
	_GUICtrlStatusBar_SetText($hStatus, "离线化学工作站未启动！请先启动 ChemStation (Offline)。", 2)
	Run(@AutoItExe & ' /AutoIt3ExecuteLine  "MsgBox(64, ''' & _
			$AppName & ''', ''' & _ ; 以下为 Msgbox 的 Text
			'数据文件需在离线化学工作站中打开，请先启动 ChemStation (Offline)！' & ''')"')
EndFunc   ;==>_MsgChemStationNotRunning

; 在ChemStation中打开
Func _LoadinChemStation()
	Local $idx, $file, $method, $ID
	Local $Index = _GUICtrlListView_GetNextItem($hMainList)
	$method = $aSearchResult[$Index + 1][1]
	$file = $aSearchResult[$Index + 1][4]
	$file = StringLeft($file, StringInStr($file, "\", 0, -1) - 1)
	; MsgBox(0, $Index, $method & @crlf & $file)

	If Not FileExists($file) Then
		_GUICtrlStatusBar_SetText($hStatus, "数据文件 " & $file & " 不存在！", 2)
		Return
	EndIf
	$ID = _GetChemID() ;~ 取得 ChemStation (Offline) 窗口 ID
	If Not $ID Then
		_MsgChemStationNotRunning()
		Return
	EndIf
	WinActivate($ID)
	_LoadMeth($ID, $file, $method, $DataMethFirst) ; 加载方法
	Sleep(20)
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadFile "' & $file & '"') ; 调用谱图
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
EndFunc   ;==>_LoadinChemStation

;~ 调用方法
Func _LoadMeth($ID, $DataFilePath, $method, $DAMFirst)
	Local $mpath = '_CONFIGMETPATH$' ; ChemStation 默认方法路径
	If $DAMFirst = 1 And FileExists($DataFilePath & "\DA.M") Then
		$mpath = '"' & $DataFilePath & '\"'
		$method = "DA.M" ; 调用数据文件中自带的方法
	ElseIf $Data_MethDirs <> "" Then
		Local $i, $arr1, $arr2
		$arr1 = StringSplit($Data_MethDirs, "||", 1)
		For $i = 1 To $arr1[0]
			$arr2 = StringSplit($arr1[$i], "|", 1)
			If StringInStr($DataFilePath, $arr2[1]) = 1 Then
				$mpath = '"' & $arr2[2] & '\"'
				ExitLoop
			EndIf
		Next
	EndIf

	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadMethod_DAOnly ' & $mpath & ',"' & $method & '"') ; 修改方法
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
EndFunc   ;==>_LoadMeth

Func _CompareWithRef()
	Local $file, $ID
	Local $item = _GUICtrlListView_GetItemTextArray($hMainList, -1) ; 读取选择的条目
	If $item[0] < 7 Then Return
	$file = $item[5]
	If Not FileExists($file) Then
		_GUICtrlStatusBar_SetText($hStatus, "数据文件 " & $file & " 不存在！", 2)
		Return
	EndIf
	$ID = _GetChemID() ;~ 取得 ChemStation (Offline) 窗口 ID
	If Not $ID Then
		_MsgChemStationNotRunning()
		Return
	EndIf
	WinActivate($ID)
	_LoadMeth($ID, $RefDataFile, $RefMethod, $DataMethFirst) ; 加载方法
	Sleep(20)
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadFile "' & $RefDataFile & '"') ; 调用参比谱图
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
	Sleep(20)
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadFile "' & $file & '", 1') ; 合并谱图
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
EndFunc   ;==>_CompareWithRef


;~ 作为参比谱图
Func _SetAsRef()
	Local $RefSampleName
	Local $item = _GUICtrlListView_GetItemTextArray($hMainList, -1) ; 读取选择的条目
	If $item[0] < 7 Then Return
	$RefSampleName = $item[1]
	$RefMethod = $item[2]
	$RefDataFile = $item[5]
	If Not FileExists($RefDataFile) Then
		_GUICtrlStatusBar_SetText($hStatus, "数据文件 " & $RefDataFile & " 不存在！", 2)
		Return
	EndIf
	If $RefSampleName <> "" Then
		$RefInfo = $RefSampleName
	Else
		$RefInfo = StringRegExpReplace($RefDataFile, ".*\\", "")
	EndIf
	GUICtrlSetData($hCompareWithRef, '和“' & $RefInfo & '”比较')
EndFunc   ;==>_SetAsRef


;~ 调用标准谱图
Func _LoadSTD()
	Local $method, $ID, $STDDataFile, $search, $ChFile, $aIndex
	$STDDataFile = FileSelectFolder("选择一个标准谱图数据文件（*.D）", $STDChromDir, 2, "", $hMainGUI)
	If @error Or StringRight($STDDataFile, 2) <> ".D" Or Not FileExists($STDDataFile) Then Return
	$search = FileFindFirstFile($STDDataFile & "\*.ch")
	If $search = -1 Then
		FileClose($search)
		Return
	EndIf
	$ChFile = FileFindNextFile($search)
	$aIndex = _IndexFile($STDDataFile & "\" & $ChFile)
	$method = $aIndex[1]
	FileClose($search)
	$ID = _GetChemID() ;~ 取得 ChemStation (Offline) 窗口 ID
	If Not $ID Then
		_MsgChemStationNotRunning()
		Return
	EndIf
	WinActivate($ID)
	_LoadMeth($ID, $STDDataFile, $method, $DataMethFirst)
	Sleep(20)
;~ LoadFile [file$], [merge], [LoadAndInteg], [LoadAndReport], [FreezeLayout], [bUseSigDetailsOverride] {44318 bytes}
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadFile "' & $STDDataFile & '"') ; 调用谱图
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
EndFunc   ;==>_LoadSTD


; 比较谱图
Func _Compare()
	Local $file, $ID
	Local $item = _GUICtrlListView_GetItemTextArray($hMainList, -1) ; 读取选择的条目
	If $item[0] < 7 Then Return
	$file = $item[5]
	If Not FileExists($file) Then
		_GUICtrlStatusBar_SetText($hStatus, "数据文件 " & $file & " 不存在！", 2)
		Return
	EndIf
	$ID = _GetChemID() ;~ 取得 ChemStation (Offline) 窗口 ID
	If Not $ID Then
		_MsgChemStationNotRunning()
		Return
	EndIf
	WinActivate($ID)
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadFile "' & $file & '", 1') ; 合并谱图
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
EndFunc   ;==>_Compare

;~ 和标准谱图比较
Func _CompareWithSTD()
	Local $file, $method, $ID, $STDDataFile
	Local $item = _GUICtrlListView_GetItemTextArray($hMainList, -1) ; 读取选择的条目
	If $item[0] < 7 Then Return
	$method = $item[2]
	$file = $item[5]
	If Not FileExists($file) Then
		_GUICtrlStatusBar_SetText($hStatus, "数据文件 " & $file & " 不存在！", 2)
		Return
	EndIf
	$STDDataFile = FileSelectFolder("选择一个标准谱图数据文件（*.D）", $STDChromDir, 2, "", $hMainGUI)
	If @error Or StringRight($STDDataFile, 2) <> ".D" Or Not FileExists($STDDataFile) Then Return
	$ID = _GetChemID() ;~ 取得 ChemStation (Offline) 窗口 ID
	If Not $ID Then
		_MsgChemStationNotRunning()
		Return
	EndIf
	WinActivate($ID)
	_LoadMeth($ID, $file, $method, $DataMethFirst) ; 加载方法
	Sleep(20)
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadFile "' & $STDDataFile & '"') ; 调用谱图
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
	Sleep(20)
	ControlSetText($ID, "", $CMDCtrlEdit, 'LoadFile "' & $file & '", 1') ; 合并谱图
	ControlSend($ID, "", $CMDCtrlEdit, "{Enter}")
EndFunc   ;==>_CompareWithSTD


;~ 作为标准谱图
Func _SetAsSTD()
	Local $method, $file, $DataFileName, $Msg, $aMethPath, $i
	Local $item = _GUICtrlListView_GetItemTextArray($hMainList, -1) ; 读取选择的条目
	If $item[0] < 7 Then Return
	$method = $item[2]
	$file = $item[5]
	If Not FileExists($file) Then
		_GUICtrlStatusBar_SetText($hStatus, "数据文件 " & $file & " 不存在！", 2)
		Return
	EndIf
;~ 	根据安捷伦化学工作站文件命名规则：文件或目录名称中不得包含以下字符： < > : " / \ | @ % * ? ' 以及空格。
;~ 数据文件名不得超过40个字符
	While 1
		$DataFileName = InputBox("标准谱图", '输入谱图说明（作为数据文件名），其中不能包含以下字符：\/:*?"<>|@%‘', _
				"", " M40", -1, -1, @DesktopWidth / 2 - 40, @DesktopHeight / 2 - 40, 60, $hMainGUI)
		If @error Then Return
		If StringRegExp($DataFileName, '[\\/:*?"<>|@%]') Then
			_GUICtrlStatusBar_SetText($hStatus, '文件名中不能包含以下字符：\/:*?"<>|@%‘。请重新输入！', 2)
			ContinueLoop ; 重新输入
		EndIf
		$DataFileName = StringReplace($DataFileName, " ", "_") & ".D" ; 去掉空格+.D
		If FileExists($STDChromDir & "\" & $DataFileName) Then
			$Msg = MsgBox(4 + 32 + 256, "标准谱图", "标准谱图数据文件 " & $DataFileName & " 已存在，是否覆盖？", 60, $hMainGUI)
			If $Msg <> 6 Then ContinueLoop ; 重新输入
		EndIf
		ExitLoop
	WEnd
	DirRemove($STDChromDir & "\" & $DataFileName, 1)
	DirCopy($file, $STDChromDir & "\" & $DataFileName, 1)
	If Not FileExists($STDChromDir & "\" & $DataFileName & "\DA.M") Then
;~ 		复制方法文件到标准谱图数据文件夹
		$aMethPath = StringSplit($MethPaths, "|")
		For $i = 1 To $aMethPath[0]
			If FileExists($aMethPath[$i] & $method) Then ; $aMethPath[$i] 路径中带有“\”
				DirCopy($aMethPath[$i] & $method, $STDChromDir & "\" & $DataFileName & "\DA.M", 1)
				FileSetAttrib($STDChromDir & "\" & $DataFileName & "\DA.M", "-R", 1)
			EndIf
		Next
	EndIf
	_GUICtrlStatusBar_SetText($hStatus, "已创建标准谱图文件： " & $DataFileName, 2)
EndFunc   ;==>_SetAsSTD

; 在资源管理器中打开Data文件夹
Func _OpenFolder()
	Local $file
	Local $item = _GUICtrlListView_GetItemTextArray($hMainList, -1) ; 读取选择的条目
	If $item[0] < 7 Then Return
	$file = $item[5]
	If Not FileExists($file) Then
		_GUICtrlStatusBar_SetText($hStatus, "数据文件 " & $file & " 不存在！", 2)
		Return
	Else
		ShellExecute($file)
	EndIf
EndFunc   ;==>_OpenFolder

#Region ; ================================= 设置 ====================================
Func _ShowSettings()
	GUI_Settings()
EndFunc   ;==>_ShowSettings

Func GUI_Settings($TabNo = 0)
	Local $arr, $i, $Tab[3]
	GUICreate("设置", 600, 460, -1, -1, _
			BitOR($WS_CAPTION, $DS_MODALFRAME, $WS_SYSMENU), -1, $hMainGUI)
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_SettingsClose")
	GUICtrlCreateTab(5, 5, 590, 420)

	$Tab[0] = GUICtrlCreateTabItem("索引设置") ; ===================索引设置======================

	Global $hAutoIndex = GUICtrlCreateCheckbox('后台监视并自动索引以下目录中新增的 ChemStation 数据文件', 15, 35)
	GUICtrlSetTip(-1, "选中此项并保持软件在后台运行，" & @CRLF & _
			"这样才能实现自动索引新增的数据文件。")
	If $AutoIndex = 1 Then
		GUICtrlSetState(-1, $GUI_CHECKED)
	EndIf
	GUICtrlSetOnEvent(-1, "_MsgDisbleAutoIndex")
	Global $hIndexDirList = GUICtrlCreateListView("索引目录", _
			15, 60, 500, 140, -1, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_SetColumnWidth(-1, 0, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSetTip(-1, "ChemStation 数据文件夹" & @CRLF & "双击可修改")
	If $IndexDirs <> "" Then
		$arr = StringSplit($IndexDirs, "|", 1)
		For $i = 1 To $arr[0]
			GUICtrlCreateListViewItem($arr[$i], $hIndexDirList)
		Next
	EndIf
	_LVEditableOn($hIndexDirList)
	GUICtrlCreateButton("帮助", 525, 60, 60, 20)
	GUICtrlSetOnEvent(-1, "_IndexSettingsHelp")
	GUICtrlCreateButton("添加", 525, 90, 60, 20)
	GUICtrlSetOnEvent(-1, "_IndexSettingsAddDir")
	GUICtrlSetTip(-1, "添加目录")
	GUICtrlCreateButton("移除", 525, 120, 60, 20)
	GUICtrlSetTip(-1, "移除选定的目录")
	GUICtrlSetOnEvent(-1, "_IndexSettingsRemoveDir")
	Global $hIndexSettingsCustom = GUICtrlCreateLabel("索引文件：", 15, 225, 85, 20)
	Global $hIndexFile = GUICtrlCreateEdit($IndexFile, 100, 220, 345, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
	GUICtrlSetTip(-1, "索引数据库文件路径")
	Global $hIndexSettingsSetFile = GUICtrlCreateButton("浏览", 455, 220, 60, 20)
	GUICtrlSetTip(-1, "选择索引数据库文件")
	GUICtrlSetOnEvent(-1, "_IndexSettingsSetFile")
	GUICtrlCreateButton("默认", 525, 220, 60, 20)
	GUICtrlSetTip(-1, "将索引文件设为" & @CRLF & "软件目录\index.db")
	GUICtrlSetOnEvent(-1, "_IndexSettingsDefaultFile")
	Global $hIndexSettingsRebuild = GUICtrlCreateCheckbox("重建索引", 15, 255, 120, 20)
	GUICtrlSetTip(-1, "重建索引数据库")
	If $IndexDirs = "" Or $IsRebuilding Then
		GUICtrlSetState($hIndexSettingsRebuild, $GUI_DISABLE)
	Else
		If Not FileExists($IndexFile) Then
			GUICtrlSetState($hIndexSettingsRebuild, $GUI_CHECKED)
		EndIf
	EndIf


	$Tab[1] = GUICtrlCreateTabItem("积分方法") ; ===================积分方法======================

	GUICtrlCreateLabel("以下目录中的方法将出现在搜索结果右键""用其它方法积分""菜单中：", 15, 40)
	Global $hMethDirList = GUICtrlCreateListView("菜单标识|方法目录", _
			15, 60, 500, 140, -1, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_SetColumnWidth(-1, 1, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSetTip(-1, "方法目录" & @CRLF & "双击可修改")
	If $AnalMethDir <> "" Then
		$arr = StringSplit($AnalMethDir, "||", 1)
		For $i = 1 To $arr[0]
			GUICtrlCreateListViewItem($arr[$i], $hMethDirList)
		Next
	EndIf
	_LVEditableOn($hMethDirList) ; 允许修改
	GUICtrlCreateButton("帮助", 525, 60, 60, 20)
	GUICtrlSetOnEvent(-1, "GUI_SettingsHelp1")
	GUICtrlCreateButton("添加", 525, 90, 60, 20)
	GUICtrlSetOnEvent(-1, "_MethDirAdd")
	GUICtrlSetTip(-1, "添加目录")
	GUICtrlCreateButton("移除", 525, 120, 60, 20)
	GUICtrlSetTip(-1, "移除选定的目录")
	GUICtrlSetOnEvent(-1, "_MethDirRemove")
	GUICtrlCreateLabel("为 ChemStation 数据文件指定默认的积分方法目录：", 15, 220, 400)
	Global $hData_MethDirsList = GUICtrlCreateListView("数据文件目录|默认方法目录", _
			15, 240, 500, 140, -1, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_SetColumnWidth(-1, 0, 240)
	_GUICtrlListView_SetColumnWidth(-1, 1, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSetTip(-1, "数据文件目录-方法目录对应表" & @CRLF & "双击可修改")
	If $Data_MethDirs <> "" Then
		$arr = StringSplit($Data_MethDirs, "||", 1)
		For $i = 1 To $arr[0]
			GUICtrlCreateListViewItem($arr[$i], $hData_MethDirsList)
		Next
	EndIf
	_LVEditableOn($hData_MethDirsList) ; 允许修改
	GUICtrlCreateButton("帮助", 525, 240, 60, 20)
	GUICtrlSetOnEvent(-1, "GUI_SettingsHelp2")
	GUICtrlCreateButton("添加", 525, 270, 60, 20)
	GUICtrlSetOnEvent(-1, "_Data_MethAdd")
	GUICtrlSetTip(-1, "添加条目")
	GUICtrlCreateButton("移除", 525, 300, 60, 20)
	GUICtrlSetTip(-1, "移除条目")
	GUICtrlSetOnEvent(-1, "_Data_MethRemove")
	$hDataMethFirst = GUICtrlCreateCheckbox("DA.M 方法优先", 15, 395, 200, 20)
	GUICtrlSetTip(-1, "优先调用数据文件中的DA.M方法")
	If $DataMethFirst = 1 Then GUICtrlSetState(-1, $GUI_CHECKED)

	$Tab[2] = GUICtrlCreateTabItem("路径转换") ; ===================路径转换======================

	GUICtrlCreateLabel("列表第一列中的路径将被转换成相应第二列中的路径：", 15, 40)
	Global $hDirMap = GUICtrlCreateListView("原始路径|替换路径", _
			15, 60, 500, 140, -1, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_SetColumnWidth(-1, 0, 240)
	_GUICtrlListView_SetColumnWidth(-1, 1, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSetTip(-1, "路径转换" & @CRLF & "双击可修改")
	If $DirMap <> "" Then
		$arr = StringSplit($DirMap, "||", 1)
		For $i = 1 To $arr[0]
			GUICtrlCreateListViewItem($arr[$i], $hDirMap)
		Next
	EndIf
	_LVEditableOn($hDirMap) ; 允许修改
	GUICtrlCreateButton("帮助", 525, 60, 60, 20)
	GUICtrlSetOnEvent(-1, "GUI_SettingsHelp3")
	GUICtrlCreateButton("添加", 525, 90, 60, 20)
	GUICtrlSetOnEvent(-1, "_DirMapAdd")
	GUICtrlSetTip(-1, "添加规则")
	GUICtrlCreateButton("移除", 525, 120, 60, 20)
	GUICtrlSetTip(-1, "移除规则")
	GUICtrlSetOnEvent(-1, "_DirMapRemove")

	GUICtrlCreateTabItem("")

	GUICtrlCreateButton("确定", 440, 430, 60, 20)
	GUICtrlSetOnEvent(-1, "GUI_SettingsOK")
	GUICtrlSetTip(-1, "确定")
	GUICtrlCreateButton("取消", 525, 430, 60, 20)
	GUICtrlSetOnEvent(-1, "GUI_SettingsClose")
	GUICtrlSetTip(-1, "取消")

	GUICtrlSetState($Tab[$TabNo], $GUI_SHOW)
	GUISetState()
EndFunc   ;==>GUI_Settings

Func GUI_SettingsHelp1() ; 通常MsgBox会让整个程序暂停运行。启动另一进程显示MsgBox可以避免这个问题
	Run(@AutoItExe & ' /AutoIt3ExecuteLine  "MsgBox(64, ''' & _
			$AppName & ' 帮助' & _ ; <- Msgbox 的 TITLE
			''', ''' & _ ; 以下为 Msgbox 的 Text
			'在搜索结果右键菜单中，有""用其它方法积分""的选项。' & _
			'该功能允许您方便地在 ChemStation 离线工作站中调用其它方法，对数据文件重新进行积分。' & @CRLF & _
			'将包含方法文件的目录加入列表，这些方法将出现在右键菜单中。' & _
			''')"') ; 以上为 Msgbox 的 Text
EndFunc   ;==>GUI_SettingsHelp1

Func GUI_SettingsHelp2() ; 通常MsgBox会让整个程序暂停运行。启动另一进程显示MsgBox可以避免这个问题
	Run(@AutoItExe & ' /AutoIt3ExecuteLine  "MsgBox(64, ''' & _
			$AppName & ' 帮助' & _ ; <- Msgbox 的 TITLE
			''', ''' & _ ; 以下为 Msgbox 的 Text
			'双击搜索结果中的数据文件，' & _
			'软件会在 ChemStation 的方法目录（通常是"C:\Chem32\1\METHODS"）中查找相应的方法，' & _
			'用于数据分析（积分）。' & @CRLF & _
			'要改变默认的方法目录，让某个目录中的数据文件在指定的方法目录中的查找方法，' & _
			'可以将数据文件目录、对应的方法目录加入这个列表中。' & _
			'如：要让""D:\3#\DATA""中的数据文件默认用""D:\3#\METHODS""中的方法积分，可点击""添加""，' & _
			'然后分别选择数据文件目录、积分方法目录。' & _
			''')"') ; 以上为 Msgbox 的 Text
EndFunc   ;==>GUI_SettingsHelp2

Func GUI_SettingsHelp3() ; 通常MsgBox会让整个程序暂停运行。启动另一进程显示MsgBox可以避免这个问题
	Run(@AutoItExe & ' /AutoIt3ExecuteLine  "MsgBox(64, ''' & _
			$AppName & ' 帮助' & _ ; <- Msgbox 的 TITLE
			''', ''' & _ ; 以下为 Msgbox 的 Text
			'在客户机上，我们可以将远程主机的 Chemstation 数据文件夹映射到本地，' & _
			'将索引文件设为主机的索引文件(index.db)，' & _
			'然后使用本软件完成数据文件的远程搜索、调用。' & _
			'搜索结果中显示的数据文件路径，与客户端访问数据文件的实际路径可能不相同，' & _
			'但是我们可以用路径转换功能得到正确的路径。' & @CRLF & _
			'例如：' & @CRLF & _
			'远程搜索结果中的一个数据文件路径为：""C:\Chem32\1\DATA\SIG123456.D""，' & _
			'这个文件位于主机的""C:\""中，而不在本地（客户机），' & _
			'实际从客户机访问这个文件的路径可能是""X:\Chem32\1\DATA\SIG123456.D""或""Z:\DATA\SIG123456.D""' & _
			'（取决于路径映射的设置）。' & _
			'要正确调用这个数据文件，我们需要将""C:\Chem32\1\""转换为""X:\Chem32\1\""或者""Z:\""。' & _
			''')"') ; 以上为 Msgbox 的 Text
EndFunc   ;==>GUI_SettingsHelp3


Func _MsgDisbleAutoIndex()
	If GUICtrlRead(@GUI_CtrlId) = $GUI_UNCHECKED Then
		Run(@AutoItExe & ' /AutoIt3ExecuteLine  "MsgBox(64, ''' & _
				$AppName & _ ; <- Msgbox 的 TITLE
				''', ''' & _ ; 以下为 Msgbox 的 Text
				'如果取消此选项，索引目录中新增的数据文件将不会被自动索引，' & _
				'您也将无法搜索到新的ChemStation数据文件。' & @CRLF & _
				''')"') ; 以上为 Msgbox 的 Text
	EndIf
EndFunc   ;==>_MsgDisbleAutoIndex


Func _DirMapAdd()
	Local $dir1 = InputBox("路径转换", "请输入需转换的路径", "", "", 280, 140, Default, Default, Default, @GUI_WinHandle)
	If @error Then Return
	Local $dir2 = InputBox("路径转换", '将路径："' & $dir1 & '"' & @CRLF & @CRLF & _
			'转换成：', '', '', 280, 160, Default, Default, Default, @GUI_WinHandle)
	If @error Then Return
	GUICtrlCreateListViewItem($dir1 & "|" & $dir2, $hDirMap)
EndFunc   ;==>_DirMapAdd

Func _DirMapRemove()
	_GUICtrlListView_DeleteItemsSelected($hDirMap)
EndFunc   ;==>_DirMapRemove

Func _Data_MethAdd()
	Local $dir1 = FileSelectFolder("选择 ChemStation 数据文件目录", "", 2, "", @GUI_WinHandle)
	If Not FileExists($dir1) Then Return
	Local $dir2 = FileSelectFolder("选择与 " & $dir1 & " 对应的方法目录", "", 2, "", @GUI_WinHandle)
	If Not FileExists($dir2) Then Return
	GUICtrlCreateListViewItem($dir1 & "|" & $dir2, $hData_MethDirsList)
EndFunc   ;==>_Data_MethAdd

Func _Data_MethRemove()
	_GUICtrlListView_DeleteItemsSelected($hData_MethDirsList)
EndFunc   ;==>_Data_MethRemove

; 添加目录
Func _MethDirAdd()
	Local $dir = FileSelectFolder("选择用于数据分析的方法目录", "", 2, "", @GUI_WinHandle)
	If Not FileExists($dir) Then Return
	GUICtrlCreateListViewItem(StringRegExpReplace($dir, ".*\\", "") & "|" & $dir, $hMethDirList)
EndFunc   ;==>_MethDirAdd

; 移除目录
Func _MethDirRemove()
	_GUICtrlListView_DeleteItemsSelected($hMethDirList)
EndFunc   ;==>_MethDirRemove

Func GUI_SettingsClose()
	_LVEditableOff()
	GUIDelete(@GUI_WinHandle)
	WinActivate($hMainGUI)
EndFunc   ;==>GUI_SettingsClose

Func GUI_SettingsOK()
	$AnalMethDir = _ListviewRead($hMethDirList, "||")
	IniWrite($SettingsFile, "Settings", "AnalMethDir", $AnalMethDir)

	$Data_MethDirs = _ListviewRead($hData_MethDirsList, "||")
	IniWrite($SettingsFile, "Settings", "Data_MethDirs", $Data_MethDirs)

	If GUICtrlRead($hDataMethFirst) = $GUI_CHECKED Then
		$DataMethFirst = 1
	Else
		$DataMethFirst = 0
	EndIf
	IniWrite($SettingsFile, "Settings", "DataMethFirst", $DataMethFirst)

	$DirMap = _ListviewRead($hDirMap, "||")
	IniWrite($SettingsFile, "Settings", "DirMap", $DirMap)

	$IndexFile = GUICtrlRead($hIndexFile)
	If DriveGetType($IndexFile) = "Network" Then
		CacheIndexDb()
		AdlibRegister("CacheIndexDb", 30000)
	Else
		AdlibUnRegister("CacheIndexDb")
		FileDelete($IndexCacheFile)
		$IndexCached = False
	EndIf
	Local $file = StringReplace($IndexFile, @ScriptDir & "\index.db", "")
	IniWrite($SettingsFile, "Settings", "IndexFile", $file)

	If GUICtrlRead($hAutoIndex) = $GUI_CHECKED Then
		$AutoIndex = 1
	Else
		$AutoIndex = 0
	EndIf
	IniWrite($SettingsFile, "Settings", "AutoIndex ", $AutoIndex)

	Local $NewDirs = _ListviewRead($hIndexDirList, "|")
	If $IndexDirs <> $NewDirs Then
		$IndexDirs = $NewDirs
		IniWrite($SettingsFile, "Settings", "Dirs", $IndexDirs)
	EndIf

	MonitorDirectory()
	AdlibUnRegister("IndexNewFile")

	If $IndexDirs <> "" And $AutoIndex = 1 Then
		Local $aDirs = StringSplit($IndexDirs, "|", 2)
		MonitorDirectory($aDirs)
		AdlibRegister("IndexNewFile", 10000)
	EndIf

	Local $Rebuild = GUICtrlRead($hIndexSettingsRebuild)
	GUI_SettingsClose()
	If $Rebuild = $GUI_CHECKED Then
		_Rebuild()
	Else
		_ShowIndexStatus($IndexFile, $hStatus)
		If FileExists($IndexFile) Then
			GUICtrlSetState($hSearch, $GUI_ENABLE)
		EndIf
	EndIf
EndFunc   ;==>GUI_SettingsOK

Func _ListviewRead($hListView, $d = "||") ; Item之间的分隔符
	Local $i, $items
	Local $cnt = _GUICtrlListView_GetItemCount($hListView)
	If $cnt > 0 Then
		For $i = 0 To $cnt - 1
			$items &= $d & _GUICtrlListView_GetItemTextString($hListView, $i)
		Next
		$items = StringTrimLeft($items, StringLen($d))
	EndIf
	Return $items
EndFunc   ;==>_ListviewRead

Func _IndexSettingsHelp() ; 通常MsgBox会让整个程序暂停运行。启动另一进程显示MsgBox可以避免这个问题
	Run(@AutoItExe & ' /AutoIt3ExecuteLine  "MsgBox(64, ''' & _
			$AppName & ' 帮助' & _ ; <- Msgbox 的 TITLE
			''', ''' & _ ; 以下为 Msgbox 的 Text
			'为加快搜索速度，本软件在指定的目录下查找ChemStation数据文件，' & _
			'提取文件相关信息并建立索引数据库。' & _
			'执行搜索时，直接在索引数据库中查找。' & @CRLF & _
			'请将包含ChemStation数据文件的目录加入列表，然后选择“重建索引”。' & _
			'只有在修改目录后需要重建索引，平时软件在后台运行，' & _
			'会自动索引新的数据文件。' & _
			''')"') ; 以上为 Msgbox 的 Text
EndFunc   ;==>_IndexSettingsHelp

Func _IndexSettingsSetFile()
	Local $path = FileOpenDialog("选择索引数据库文件", StringRegExpReplace($IndexFile, "[^\\]*$", ""), "SQLite3 files (*.db)|All files (*.*)", 0, "index.db", @GUI_WinHandle)
	If Not @error Then
		GUICtrlSetData($hIndexFile, $path)
		If FileExists($path) Then
			GUICtrlSetState($hIndexSettingsRebuild, $GUI_UNCHECKED)
		Else
			GUICtrlSetState($hIndexSettingsRebuild, $GUI_CHECKED)
		EndIf
	EndIf
EndFunc   ;==>_IndexSettingsSetFile

Func _IndexSettingsDefaultFile()
	GUICtrlSetData($hIndexFile, @ScriptDir & "\index.db")
EndFunc   ;==>_IndexSettingsDefaultFile

; 添加目录
Func _IndexSettingsAddDir()
	Local $NewDir = FileSelectFolder("选择含 ChemStation 数据文件的文件夹", "", 2, "", @GUI_WinHandle)
	If Not FileExists($NewDir) Then Return
	Local $sDirs = _ListviewRead($hIndexDirList, "|")
	Local $NewDirs = $sDirs & "|" & $NewDir
	$NewDirs = CheckDirs($NewDirs)
	If $sDirs = $NewDirs Then
		_GUICtrlStatusBar_SetText($hStatus, '“' & $NewDir & '”已包含在列表中', 2)
		Return
	EndIf

	_GUICtrlListView_DeleteAllItems($hIndexDirList)
	If $NewDirs <> "" Then
		Local $arr = StringSplit($NewDirs, "|", 1)
		For $i = 1 To $arr[0]
			GUICtrlCreateListViewItem($arr[$i], $hIndexDirList)
		Next
	EndIf
	GUICtrlSetState($hIndexSettingsRebuild, $GUI_ENABLE)
	If Not FileExists($IndexFile) Then
		GUICtrlSetState($hIndexSettingsRebuild, $GUI_CHECKED)
	EndIf
EndFunc   ;==>_IndexSettingsAddDir

; 移除目录
Func _IndexSettingsRemoveDir()
	Local $rDir = _GUICtrlListView_GetItemTextString($hIndexDirList, -1)
	If $rDir = "" Then Return
	If StringInStr('|' & $DataPaths & '|', '|' & $rDir & '|') Then
		_GUICtrlStatusBar_SetText($hStatus, '“' & $rDir & '”是ChemStation当前数据目录，不能移除', 2)
		Return
	EndIf
	_GUICtrlListView_DeleteItemsSelected($hIndexDirList)
	If _GUICtrlListView_GetItemCount($hIndexDirList) < 1 Then
		GUICtrlSetState($hIndexSettingsRebuild, $GUI_DISABLE)
	EndIf
EndFunc   ;==>_IndexSettingsRemoveDir
#EndRegion ; ================================= 设置 ====================================

; 清除输入框及搜索结果
Func _OnEventClear()
	Local $Date_Disable[7] = [True]
	Local $Date_Enable[7] = [False, @YEAR, @MON, @MDAY]
	_GUICtrlDTP_SetSystemTime($hDate1, $Date_Enable)
	_GUICtrlDTP_SetSystemTime($hDate1, $Date_Disable)
	_GUICtrlDTP_SetSystemTime($hDate2, $Date_Enable)
	_GUICtrlDTP_SetSystemTime($hDate2, $Date_Disable)
	GUICtrlSetData($hSampleInfo, "")
	GUICtrlSetData($hSampleName, "")
	GUICtrlSetData($hMethod, "")
	GUICtrlSetData($hDataFile, "")
	$aKeyWords = ""
	GUICtrlSendMsg($hMainList, $LVM_SETITEMCOUNT, 0, 0)
	GUICtrlSendMsg($hMainList, $LVM_REDRAWITEMS, 0, 20)
	_GUICtrlStatusBar_SetText($hStatus, "搜索安捷伦化学工作站色谱、电泳谱数据文件(*.ch)", 2)
EndFunc   ;==>_OnEventClear

; 日期设为最近
Func _SetDateRecent()
	Local $rdays = GUICtrlRead(@GUI_CtrlId)
	GUICtrlSetData(@GUI_CtrlId, "最近...")
	If $rdays = "最近..." Then Return
	If $rdays = "其它..." Then
		$rdays = InputBox("日期", "将开始日期设为最近几天", $RecentDays, "", _
				250, 140, Default, Default, Default, @GUI_WinHandle)
		If $rdays = "" Or Not Int($rdays) Then Return
		$RecentDays = Int($rdays)
	EndIf
	$rdays = Int($rdays)
	Local $DateFrom = _DateAdd('D', -$rdays, @YEAR & "/" & @MON & "/" & @MDAY)
	Local $a = StringSplit($DateFrom, "/")
	Local $aDate[7] = [False, $a[1], $a[2], $a[3]]
	_GUICtrlDTP_SetSystemTime($hDate1, $aDate)
	Local $aDate[7] = [False, @YEAR, @MON, @MDAY]
	_GUICtrlDTP_SetSystemTime($hDate2, $aDate)
	$aDate[0] = True
	_GUICtrlDTP_SetSystemTime($hDate2, $aDate)
	GUICtrlSetState($hSearch, $GUI_FOCUS)
EndFunc   ;==>_SetDateRecent

; 在Windows启动时运行
Func _RunOnStart()
	If BitAND(GUICtrlRead($RunOnStartItem), $GUI_CHECKED) = $GUI_CHECKED Then
		$RunOnStart = 0
	Else
		$RunOnStart = 1
	EndIf
	_SetRunOnStart()
	IniWrite($SettingsFile, "Settings", "RunOnStart", $RunOnStart)
EndFunc   ;==>_RunOnStart

Func _SetRunOnStart()
	Local $Key
	If IsAdmin() Then
		$Key = "HKLM"
	Else
		$Key = "HKCU"
	EndIf
	If $RunOnStart = 1 Then
		GUICtrlSetState($RunOnStartItem, $GUI_CHECKED)
		RegWrite($Key & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", $AppName, "REG_SZ", '"' & @ScriptFullPath & '" -Hide')
	Else
		GUICtrlSetState($RunOnStartItem, $GUI_UNCHECKED)
		RegDelete($Key & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", $AppName)
	EndIf
EndFunc   ;==>_SetRunOnStart

; 关键词处理，中间空格替换为%，符合SQLite3的要求
Func _KeyWords($str)
	$str = StringStripWS($str, 7)
	$str = StringReplace($str, " ", "%")
	$str = "%" & $str & "%"
	Return $str
EndFunc   ;==>_KeyWords

; 检查ChemStation 配置文件，检查目录列表中是否已包含数据文件夹
Func _WatchIniFile()
	Local $wNewDirs, $aDirs
	If Not _FileChange($IniFile) Then
		Return
	EndIf
	GetChemIni($IniFile, $DataPaths, $InstNames, $MethPaths)
	$wNewDirs = $IndexDirs & "|" & $DataPaths
	$wNewDirs = CheckDirs($wNewDirs)
	If $wNewDirs <> $IndexDirs Then
		$IndexDirs = $wNewDirs
		MonitorDirectory() ; 暂停监控
		If $IndexDirs <> "" And $AutoIndex = 1 Then
			$aDirs = StringSplit($IndexDirs, "|", 2)
			MonitorDirectory($aDirs)
		EndIf
		IniWrite($SettingsFile, "Settings", "Dirs", $IndexDirs)
	EndIf
	ReduceMemory() ; 整理内存
EndFunc   ;==>_WatchIniFile

; 检查文件是否已修改
Func _FileChange($file)
	Local $Time, $Change
	Static $OldTime = ""
	$Time = FileGetTime($file, 0, 1)
	If $Time = $OldTime Then
		$Change = 0
	Else
		$OldTime = $Time
		$Change = 1
	EndIf
	Return $Change
EndFunc   ;==>_FileChange

; 函数。检查目录是否重叠(目录名之间用｜分隔开)
; 例：
;~ $Dirs = "|D:\lou|D:\lou\2009|D:\lou||D:\lou|C:\AAA|"
;~ MsgBox(0, "", $Dirs & @CRLF & "--> " & CheckDirs($Dirs))
Func CheckDirs($Dirs)
	If $Dirs = "" Then Return
	Local $nDirs
	$Dirs = StringRegExpReplace($Dirs, '\|{2,}', '|')
	$Dirs = StringRegExpReplace($Dirs, '^\|+|\|+$', "") ; 去掉开头或结尾的"|"
	Local $a = StringSplit($Dirs, "|", 2)
	_ArrayQuickSort($a)
	$Dirs = $a[0]
	For $i = 1 To UBound($a) - 1
		$Dirs &= "|" & $a[$i]
	Next
	Local $match[1]
	While 1
		$match = StringRegExp($Dirs, '^[^\|]+', 1) ; 取第一个目录
		If @error Then
			ExitLoop
		EndIf
		If FileExists($match[0]) Then
			$nDirs &= $match[0] & "|"
		EndIf
		$Dirs = StringRegExpReplace($Dirs, '(?i)\Q' & $match[0] & '\E[^\|]*\|?', "")
	WEnd
	$nDirs = StringRegExpReplace($nDirs, '\|$', "") ; 去掉最后一个"|"
	Return $nDirs
EndFunc   ;==>CheckDirs

Func _ArrayQuickSort(ByRef $avArray)
	If UBound($avArray) <= 1 Then Return
	Local $vTmp
	For $i = 1 To UBound($avArray) - 1
		$vTmp = $avArray[$i]
		For $j = $i - 1 To 0 Step -1
			If (StringCompare($vTmp, $avArray[$j]) >= 0) Then ExitLoop
			$avArray[$j + 1] = $avArray[$j]
		Next
		$avArray[$j + 1] = $vTmp
	Next
EndFunc   ;==>_ArrayQuickSort

; 读取ChemStation配置文件信息（_DATAPATH$，InstName等）
Func GetChemIni($IniFile, ByRef $DataPaths, ByRef $InstNames, ByRef $MethPaths)
	Local $i, $_DataPath, $InstName, $_CONFIGMETPATH, $cDataPaths, $cInstNames, $cMethPaths
	Local $Instruments = IniRead($IniFile, "PCS", "Instruments", "") ; Instruments=1,2,3...
	Local $str = StringSplit($Instruments, ",")
	For $i = 1 To $str[0]
		$_DataPath = IniRead($IniFile, "PCS," & $str[$i], "_DATAPATH$", "")
		$InstName = IniRead($IniFile, "PCS," & $str[$i], "InstName", "")
		$_CONFIGMETPATH = IniRead($IniFile, "PCS," & $str[$i], "_CONFIGMETPATH$", "")
		; 仪器名称会显示在ChemStation的标题栏中，读取InstName用来鉴别ChemStation窗口
		If FileExists($_DataPath) Then
			$cDataPaths = $cDataPaths & $_DataPath & "|"
		EndIf
		If $InstName <> "" Then
			$cInstNames = $cInstNames & $InstName & "|"
		EndIf
		If $_CONFIGMETPATH <> "" Then ; _CONFIGMETPATH 路径中带有“\”
			$cMethPaths = $cMethPaths & $_CONFIGMETPATH & "|"
		EndIf
	Next
	$InstNames = StringRegExpReplace($cInstNames, "\|$", "") ; 去掉结尾的“|”
	$DataPaths = StringRegExpReplace($cDataPaths, "\|$", "")
	$MethPaths = StringRegExpReplace($cMethPaths, "\|$", "")
EndFunc   ;==>GetChemIni

;~ 函数。整理内存
;~ http://www.autoitscript.com/forum/index.php?showtopic=13399&hl=GetCurrentProcessId&st=20
; Original version : w_Outer
; modified by Rajesh V R to include process ID
Func ReduceMemory($ProcID = 0)
	Local $ai_GetCurrentProcessId
	If $ProcID = 0 Or ProcessExists($ProcID) = 0 Then ; No process id specified or process doesnt exist - use current process instead.
		$ai_GetCurrentProcessId = DllCall('kernel32.dll', 'int', 'GetCurrentProcessId')
		$ProcID = $ai_GetCurrentProcessId[0]
	EndIf
	Local $ai_Handle = DllCall("kernel32.dll", 'int', 'OpenProcess', 'int', 0x1f0fff, 'int', False, 'int', $ProcID)
	Local $ai_Return = DllCall("psapi.dll", 'int', 'EmptyWorkingSet', 'long', $ai_Handle[0])
	DllCall('kernel32.dll', 'int', 'CloseHandle', 'int', $ai_Handle[0])
	Return $ai_Return[0]
EndFunc   ;==>ReduceMemory

;索引新文件
Func IndexNewFile()
	If $NewCHFiles = "" Then Return
	If $IsRebuilding Then
		Return
	EndIf

	Local $i, $j, $aIndex, $hDb, $idx, $iITEM_COUNT
	$hDb = _SQLite_Open($IndexFile)
	_SQLite_Exec($hDb, "CREATE TABLE IF NOT EXISTS Index1 (样品名称 TEXT,分析方法 TEXT,分析日期 TEXT," & _
			"分析时间 TEXT,数据文件 TEXT UNIQUE,操作者 TEXT,样品信息 TEXT);")
	$NewCHFiles = StringTrimRight($NewCHFiles, 1) ; 去掉最后的 "|"
	Local $arr = StringSplit($NewCHFiles, "|", 2)
	$NewCHFiles = ""
	For $i = 0 To UBound($arr) - 1
		$aIndex = _IndexFile($arr[$i])
		_SQLite_Exec($hDb, "REPLACE INTO Index1 VALUES ('" & $aIndex[0] & "','" & $aIndex[1] & "','" & _
				$aIndex[2] & "','" & $aIndex[3] & "','" & $aIndex[4] & "','" & $aIndex[5] & "','" & $aIndex[6] & "');")
		$idx = _ArraySearch($aSearchResult, $aIndex[4], 1, 0, 0, 1, 1, 4)
		If $idx <> -1 Then ; 更新列表
			For $j = 0 To UBound($aIndex) - 1
				$aSearchResult[$idx][$j] = $aIndex[$j]
			Next
			GUICtrlSendMsg($hMainList, $LVM_REDRAWITEMS, 0, 20)
		Else ; 若新ch文件符合最近一次搜索条件，则加入搜索输出列表
			If Not IsArray($aKeyWords) Then ContinueLoop
			If $aIndex[0] <> "" And Not StringRegExp($aIndex[0], "(?i)" & $aKeyWords[0]) Then ContinueLoop
			If $aIndex[1] <> "" And Not StringRegExp($aIndex[1], "(?i)" & $aKeyWords[1]) Then ContinueLoop
			If $aIndex[2] < $aKeyWords[2] Or $aIndex[2] > $aKeyWords[3] Then ContinueLoop
			If $aIndex[4] <> "" And Not StringRegExp($aIndex[4], "(?i)" & $aKeyWords[4]) Then ContinueLoop
			If $aIndex[6] <> "" And Not StringRegExp($aIndex[6], "(?i)" & $aKeyWords[5]) Then ContinueLoop

			$iITEM_COUNT = UBound($aSearchResult) ; UBound($aSearchResult) -1 + 1 搜索结果条数
			ReDim $aSearchResult[$iITEM_COUNT + 1][7]
			For $j = 0 To 6
				$aSearchResult[$iITEM_COUNT][$j] = $aIndex[$j]
			Next
			GUICtrlSendMsg($hMainList, $LVM_SETITEMCOUNT, $iITEM_COUNT, 0)
			GUICtrlSendMsg($hMainList, $LVM_REDRAWITEMS, 0, 20)
		EndIf
		_GUICtrlStatusBar_SetText($hStatus, @HOUR & ":" & @MIN & ":" & @SEC & " 更新: " & $aIndex[0] & " - " & $aIndex[4], 2)
	Next
	_SQLite_Close($hDb)
	_ShowIndexStatus($IndexFile, $hStatus)
	GUICtrlSetState($hSearch, $GUI_ENABLE)
	ReduceMemory()
EndFunc   ;==>IndexNewFile



#Region =========================== FUNCTION MonitorDirectory() ==============================
#cs
Description:     Monitors the user defined directories for file activity.
Original:        http://www.autoitscript.com/forum/index.php?showtopic=69044&hl=folderspy&st=0
Modified:        Jack Chen
Syntax:          MonitorDirectory($Dirs = "", $Subtree = True, $ext = "", $TimerMs = 1000)
Parameters: 	 $Dirs		- Optional: Array of directories to be monitored.
$Subtree	     Subtrees will be monitored if $Subtree = True.
$ext             file extention filter like txt, docx, xlsx...
$TimerMs         Timer to register changes in milliseconds
Remarks:		Call MonitorDirectory() without parameters to stop monitoring all directories.
THIS SHOULD BE DONE BEFORE EXITING SCRIPT AT LEAST.
#ce
Func MonitorDirectory($Dirs = "", $Subtree = True, $ext = ".ch", $TimerMs = 1000)
	Local Static $nMax, $hBuffer, $hEvents, $aSubtree, $fileex
	Local Static $aDirHandles[0], $aOverlapped[0], $aDirs[0]

	If IsArray($Dirs) Then
		;ConsoleWrite("Start dir monitoring... " & @CRLF)

		$aDirs = $Dirs
		$nMax = UBound($aDirs)
		ReDim $aDirHandles[$nMax], $aOverlapped[$nMax]
		$aSubtree = $Subtree
		$fileex = $ext

		$hBuffer = DllStructCreate("byte[65536]")
		For $i = 0 To $nMax - 1
			If StringRight($aDirs[$i], 1) <> "\" Then $aDirs[$i] &= "\"
			; http://msdn.microsoft.com/en-us/library/aa363858%28VS.85%29.aspx
			Local $aResult = DllCall("kernel32.dll", "hwnd", "CreateFile", "Str", $aDirs[$i], _
					"Int", 0x1, "Int", BitOR(0x1, 0x4, 0x2), "ptr", 0, "int", 0x3, "int", BitOR(0x2000000, 0x40000000), "int", 0)
			$aDirHandles[$i] = $aResult[0]
			$aOverlapped[$i] = DllStructCreate("ulong_ptr Internal;ulong_ptr InternalHigh;dword Offset;dword OffsetHigh;handle hEvent")
			For $j = 1 To 5
				DllStructSetData($aOverlapped[$i], $j, 0)
			Next
			_MonitorDir($aDirHandles[$i], $hBuffer, $aOverlapped[$i], True, $aSubtree)
		Next
		$hEvents = DllStructCreate("hwnd hEvent[" & $nMax & "]")
		For $j = 1 To $nMax
			DllStructSetData($hEvents, "hEvent", DllStructGetData($aOverlapped[$j - 1], "hEvent"), $j)
		Next
		AdlibRegister("_ReadDirChanges", $TimerMs)

	ElseIf $Dirs = "ReadDirChanges" Then
		;ConsoleWrite("Read dir changes... " & @CRLF)

		Local $aMsg = DllCall("User32.dll", "dword", "MsgWaitForMultipleObjectsEx", "dword", $nMax, _
				"ptr", DllStructGetPtr($hEvents), "dword", -1, "dword", 0x4FF, "dword", 0x6)
		Local $i = $aMsg[0]
		If $i >= 0 And $i < $nMax Then
			DllCall("Kernel32.dll", "Uint", "ResetEvent", "uint", DllStructGetData($aOverlapped[$i], "hEvent"))

			Local $hFileNameInfo, $hFileName, $pBuffer, $filepath, $ex, $ActionID, $hDb
			Local $nOffset = 0, $nNext = 1
			$pBuffer = DllStructGetPtr($hBuffer)
			While $nNext <> 0
				$hFileNameInfo = DllStructCreate("dword NextEntryOffset;dword Action;dword FileNameLength", $pBuffer + $nOffset)
				$hFileName = DllStructCreate("wchar FileName[" & DllStructGetData($hFileNameInfo, "FileNameLength") / 2 & "]", _
						$pBuffer + $nOffset + 12)
				$filepath = $aDirs[$i] & DllStructGetData($hFileName, "FileName")
				$ex = StringMid($filepath, StringInStr($filepath, ".", 0, -1))
				If Not $fileex Or $ex = $fileex Then
					$ActionID = DllStructGetData($hFileNameInfo, "Action")
					If $ActionID = 0x1 Or $ActionID = 0x3 Or $ActionID = 0x5 Then ; $FILE_ACTION_ADDED, ; $FILE_ACTION_MODIFIED, $FILE_ACTION_RENAMED_NEW_NAME
						;ConsoleWrite("File added - " & $ActionID & ": " & $filepath & @CRLF)

						If Not StringInStr($NewCHFiles, $filepath & "|") Then
							$NewCHFiles &= $filepath & "|"
						EndIf
					ElseIf $ActionID = 0x2 Or $ActionID = 0x4 Then ; $FILE_ACTION_REMOVED, $FILE_ACTION_RENAMED_OLD_NAME
						;ConsoleWrite("File removed - " & $ActionID & ": " & $filepath & @CRLF)

						; 从数据库中删除记录
						$hDb = _SQLite_Open($IndexFile)
						_SQLite_Exec($hDb, "DELETE FROM Index1 WHERE 数据文件='" & $filepath & "';")
						_SQLite_Close($hDb)
						_ShowIndexStatus($IndexFile, $hStatus)
						_GUICtrlStatusBar_SetText($hStatus, @HOUR & ":" & @MIN & ":" & @SEC & " 删除: " & StringRegExpReplace($filepath, "\\[^\\]*$", ""), 2)
					EndIf
				EndIf
				$nNext = DllStructGetData($hFileNameInfo, "NextEntryOffset")
				$nOffset += $nNext
			WEnd
			_MonitorDir($aDirHandles[$i], $hBuffer, $aOverlapped[$i], False, $aSubtree)
		EndIf

	ElseIf $Dirs = "" Then
		;ConsoleWrite("Stop dir monitoring... " & @CRLF)

		AdlibUnRegister("_ReadDirChanges")
		DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hEvents)
		For $i = 0 To $nMax - 1
			DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $aOverlapped[$i])
			DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $aDirHandles[$i])
		Next
		$nMax = 0
		$hBuffer = 0
		$hEvents = 0
		ReDim $aDirHandles[0], $aOverlapped[0], $aDirs[0]
		$aSubtree = 0
		$fileex = ""
		Sleep(20) ; 有时停止监视网络文件夹时会出现内存分配错误，加上延时就好了。
	EndIf
EndFunc   ;==>MonitorDirectory
Func _MonitorDir($hDir, $hBuffer, $hOverlapped, $bInitial = False, $bSubtree = True)
	Local $hEvent, $pBuffer, $nBufferLength, $pOverlapped
	$pBuffer = DllStructGetPtr($hBuffer)
	$nBufferLength = DllStructGetSize($hBuffer)
	$pOverlapped = DllStructGetPtr($hOverlapped)
	If $bInitial Then
		$hEvent = DllCall("kernel32.dll", "hwnd", "CreateEvent", "UInt", 0, "Int", True, "Int", False, "UInt", 0)
		DllStructSetData($hOverlapped, "hEvent", $hEvent[0])
	EndIf
	; http://msdn.microsoft.com/en-us/library/aa365465%28VS.85%29.aspx
	; $aResult = DllCall("kernel32.dll", "Int", "ReadDirectoryChangesW", "hwnd", $hDir, "ptr", _
	; $pBuffer, "dword", $nBufferLength, "int", $bSubtree, "dword", _
	; BitOR(0x1, 0x2, 0x4, 0x8, 0x10, 0x20, 0x40, 0x100), "Uint", 0, "Uint", $pOverLapped, "Uint", 0)
	; BitOR(0x1, 0x2, 0x4, 0x8, 0x10, 0x40, 0x100) 的说明：
	; 0x1 - 监视文件新建/删除/重命名, 0x2 - 监视文件夹新建/删除, 0x4 - 监视属性, 0x8 - 监视文件大小变化,
	; 0x10 - 监视文件最后修改时间, 0x40 - 监视文件创建时间, 0x100 - 监视安全描述
	; 若只要监视文件名、文件夹名变化（不监视修改），可改成 BitOR(0x1, 0x2)，
	; 若只要监视文件名变化（忽略文件夹），则可直接改成0x1（不用BitOR了）
	Local $aResult = DllCall("kernel32.dll", "Int", "ReadDirectoryChangesW", "hwnd", $hDir, _
			"ptr", $pBuffer, "dword", $nBufferLength, "int", $bSubtree, "dword", BitOR(0x1, 0x10), "Uint", 0, "Uint", $pOverlapped, "Uint", 0)
	Return $aResult[0]
EndFunc   ;==>_MonitorDir
Func _ReadDirChanges()
	MonitorDirectory("ReadDirChanges")
EndFunc   ;==>_ReadDirChanges
#EndRegion =========================== FUNCTION MonitorDirectory() ==============================
