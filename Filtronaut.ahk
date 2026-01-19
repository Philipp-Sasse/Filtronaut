; the Filtronaut switches to windows or opens recent items with given string

global MyWindowId := 0
global Config := {}
Config.FilterModes := {"!O":"Open windows", "!D":"Directories", "!F":"Favorites", "!B":"Bookmarks", "!R":"Recent", "!W":"Word", "!X":"eXcel", "!P":"Pdf", "!M":"Media"}
Config.Launch := {}
Config.Path["Bookmarks"] := localAppData "\Microsoft\Edge\User Data\Default\Bookmarks"
Config.Modes := {"Recent": ".", "Word": "i)\.doc[xm]?$", "eXcel": "i)\.xls[xm]?$", "Pdf": "i)\.pdf$"}
Config.Actions["!1"] := {monitor: 1, layout: "x"}
Config.Exclude := []
EnvGet, userProfile, USERPROFILE
global favFolder := userProfile "\Favorites"

; search for config in appData, programData and scriptDir
configFile := localAppData "\Filtronaut\Filtronaut.config"
if !FileExist(configFile) {
	 EnvGet, programData, PROGRAMDATA
	 configFile := programData "\Filtronaut\Filtronaut.config"
		  if !FileExist(configFile)
				configFile := A_ScriptDir "\Filtronaut.config"
}
; parse config
Loop
{
	FileReadLine, line, %configFile%, %A_Index%
	if ErrorLevel
		break
	if RegExMatch(line, "^\s*([^# ]\S*) +([^:]*):\s*([^\t]*)(\t+(.*))?$", match)
	{
		;MsgBox, Command "%match1%", target "%match2%", key "%match3%", parameter "%match5%"
		Switch, match1
		  {
			Case "Hotkey":
				Switch, match2 {
					Case "Launch": Config.Launch[match3] := match5
					Case "Save": Config.Hotkey["Save"] := match3
					Case "Class": Config.Classes[match3] := match5
					Default:
						for key, value in Config.FilterModes
							if (value = match2) {
								Config.FilterModes.Delete(key)
								Config.FilterModes[match3] := match2
							}
				}
			Case "Mode": {
				Config.FilterModes[match3] := match2
				Config.Modes[match2] := match5
				}
			Case "Sniplets": {
				Config.FilterModes[match3] := match2
				Config.Sniplets[match2] := ResolvePath(match5)
			}
			Case "Monitor": {
					 if (!Config.Actions.HasKey(match3))
						  Config.Actions[match3] := {}
					 Config.Actions[match3].monitor := match2
					 Config.Actions[match3].layout := match5
				}
			Case "Action": {
					 if (!Config.Actions.HasKey(match3))
						  Config.Actions[match3] := {}
					 Config.Actions[match3].sniplet := match2
					 Config.Actions[match3].command := match5
				}
			Case "Path": Config.Path[match2] := ResolvePath(match3)
			Case "Exclude": Config.Exclude.Push({ scope: match2, filter: match3})
			Default:
				MsgBox, Unknown command %match1% in line %A_Index%
		}
	}
	else if RegExMatch(line, "^\s[^#].*$") ; no comment
		MsgBox, 4, , Line #%A_Index%: Cannot parse "%line%". Continue?
	IfMsgBox, No
		return
}
modeList :=
for hotkey, mode in Config.FilterModes
	modeList .= mode "|"
Menu, Tray, NoStandard
Menu, Tray, Add, Help, ShowHelp
Menu, Tray, Add, Exit, ExitApp
Menu, Tray, Default, Help

if (Config.Launch.Count() = 0)
	Config.Launch := {"!Esc":"Open windows"}
for hotkey, mode in Config.Launch
	Hotkey, %hotkey%, LaunchGUI
return

;====================
LaunchGUI:
{
	if (MyWindowId > 0) {
		WinActivate, ahk_id %MyWindowId%
		WinRestore, ahk_id %MyWindowId%
		Return
	}
	global CachedList := []
	global ItemList := []
	global SelectedIndex := 1
	global PresetIndex := 1
	global CaseSensitive := false
	RecentItemList := []
	RecentItemListBuilt := false
	if (Config.Launch.HasKey(A_ThisHotkey))
		FilterMode := Config.Launch[A_ThisHotkey]
	else
		FilterMode := "Open windows"

	Gui, +AlwaysOnTop +ToolWindow +LastFound
	Gui, Font, s10
	Gui, Add, Edit, x10 y10 w280 vSearchInput gUpdateList
	Gui, Add, Checkbox, x+10 yp w50 h22 vCaseToggle gToggleCase, case
	Gui, Add, DropDownList, x+8 yp w120 vModeSelector gModeChanged, %modeList%
	Gui, Add, Button, x+10 yp w22 h22 gShowHelp, ?
	Gui, Add, ListBox, x10 y+10 w500 h200 vWindowBox gListBoxChanged
	Gui, Show,, Filtronaut
	GuiControl, ChooseString, ModeSelector, %FilterMode%
	MyWindowId := WinExist()
	global PredefinedHotkeys := "!o,!d,!f,!b,!w,!x,!p,!r,!m,!1"
	for hotkey, mode in Config.FilterModes {
		if hotkey not in %PredefinedHotkeys%
		{
			Hotkey, %hotkey%, HandleModeHotkey
			Hotkey, %hotkey%, On
		}
	}
	for hotkey, mode in Config.Classes {
		if hotkey not in %PredefinedHotkeys%
		{
			Hotkey, %hotkey%, HandleModeHotkey
			Hotkey, %hotkey%, On
		}
	}
	for hotkey, action in Config.Actions {
		if hotkey not in %PredefinedHotkeys%
		{
			Hotkey, %hotkey%, selection
			Hotkey, %hotkey%, On
		}
	}
	if (Config.Hotkey.HasKey("Save")) {
		Hotkey, % Config.Hotkey["Save"], SaveSniplets
		Hotkey, % Config.Hotkey["Save"], On
	}
	Gosub, ModeChanged
	return
}

;====================
ShowHelp:
{
	Gui, -AlwaysOnTop
	Gui, New, +AlwaysOnTop +ToolWindow +Resize, Filtronaut help
	Gui, Margin, 12, 10

	; ActiveX Browser-Control
	Gui, Add, ActiveX, vWB w720 h480, Shell.Explorer

	css :=
	(
	"<style>
	  body{font:13px Candara,sans-serif; color:#111; margin:0; padding:0 0 24px 0; background:#fff;}
	  h1{font-size:15px; margin:0 0 12px;}
	  h2{font-size:14px; margin:16px 0 6px}
	  kbd{font:12px Lucida,sans-serif; background:#eee; padding:0px 3px; border:1px solid #888; border-radius:7px}
	  li{padding: 2px}
	  table{border-collapse: collapse}
	  td{padding: 3px; border: 1px solid #ddd}
	</style>"
	)

	html := "<!DOCTYPE html><html><head><meta charset='utf-8'>" css "</head><body><div class='wrap'>"
	html .= "<h1>The Filtronaut has the filters to navigate you to open windows, recent documents, bookmarks and more.</h1>"

	html .= "<h2>How to Use:</h2><ul>"
	html .= "<li>Start typing to filter the list</li>"
	html .= "<li>Press <kbd>Tab</kbd> to auto-complete</li>"
	html .= "<li>Press Down and Up to navigate the list</li>"
	html .= "<li>Press <kbd>Return</kbd> to run or bring to front the current selection</li>"
	html .= "<li>Press <kbd>Alt</kbd>+<kbd>1</kbd> to maximise the current selection on screen 1 (more keys configurable)</li>"
	html .= "<li>Press <kbd>Esc</kbd> to close the Filtronaut window or <kbd>Ctrl</kbd>+<kbd>Esc</kbd> to exit the app</li>"
	html .= "</ul>"

	if (Config.Launch.Count()) {
		html .= "<h2>Launcher</h2>"
		html .= RenderTable(Config.Launch)
	}

	html .= "<h2>Filter modes:</h2>"
	html .= RenderTable(Config.FilterModes)

	html .= "<h2>Other shortcuts</h2>"
	shortcuts := {"!C":"toggle Case sensitive search"
		, "^Backspace":"close the selected window or remove the recent item link or sniplet line"
		, "^+":"add the current filter text to the sniplet collection or item to favorites"
		, "^UpDown":"move the selected sniplet in the list"
		, "!Return":"copy the selection to the filter text; further uses change the copied sniplet line or rename the recent/favorites link"
		, "^H":"show this beautiful little Help"}
	if (Config.Hotkey.HasKey("Save"))
		shortcuts .= {Config.Hotkey["Save"]:"Save sniplets"}
	html .= RenderTable(shortcuts)

	if (Config.Actions.Count()) {
		html .= "<h2>Actions</h2>"
		html .= "<table>"
		for hk, action in Config.Actions {
			if (action.HasKey("monitor")) {
				layout := StrReplace(StrReplace(StrReplace(action.layout, "x", "maximised"), "%,", "with borders "), "c", "centered")
				desc := "open " layout " on monitor " action.monitor
				html .= HelpRow(hk, desc)
			}
			if (action.HasKey("command")) {
				scope := action.sniplet
				desc := (scope = "*") ? "for all sniplets" : "for sniplet " scope
				desc .= " do: <code>" HtmlEsc(action.command) "</code>"
				html .= HelpRow(hk, desc)
			}
		}
		html .= "</table>"
	}
	html .= "</div></body></html>"

	WB.Navigate("about:blank")
	while (WB.ReadyState != 4)
		Sleep, 10
	doc := WB.Document
	doc.Open()
	doc.Write(html)
	doc.Close()

	Gui, Show
	return
}

HelpRow(hotkey, desc) {
	RegExMatch(hotkey, "([!^#+<>]*)([A-Za-z]*.)$", split)
	hotkey := RegExReplace(split1, ".[a-z]*", "<kbd>$0</kbd>")
	hotkey := StrReplace(StrReplace(StrReplace(StrReplace(hotkey
		, "^", "Ctrl")
		, "!", "Alt")
		, "#", "Win")
		, "+", "Shift")
	hotkey .= "<kbd>" split2 "</kbd>"
	return "<tr><td>" StrReplace(hotkey, "><", ">+<") "</td><td>" desc "</td></tr>"
}

RenderTable(Map) {
	table := "<table>"
	for hotkey, mode in Map {
		RegExMatch(hotkey, "([!^#+<>]*)([A-Za-z]*.)$", split)
		StringCaseSense On
		disp := StrReplace(mode, split2, "<b>" split2 "</b>",,1)
		StringCaseSense Off
		table  .= HelpRow(hotkey, disp)
	}
	table .= "</table>"
	return table
}

HtmlEsc(s) {
	s := StrReplace(s, "&",  "&amp;")
	s := StrReplace(s, "<",  "&lt;")
	s := StrReplace(s, ">",  "&gt;")
	s := StrReplace(s, """", "&quot;")
	s := StrReplace(s, "'",  "&#39;")
	return s
}

;====================
IsRecentBased(mode) {
	 ;global Config
	return  Config.Modes.HasKey(mode) || mode = "Directories"
}

ModeChanged:
{
	Gosub, CheckPendingSaves
	CachedList := []
	EditedItem := ""
	SetEditVisual(false)
	GuiControlGet, FilterMode,, ModeSelector
	if (Config.Sniplets.HasKey(FilterMode)) {
		path := Config.Sniplets[FilterMode]
		Loop, Read, %path%
		{
			line := Trim(A_LoopReadLine)
			if (line != "")
				CachedList.Push({title: line, path: path})
		}
	} else if IsRecentBased(FilterMode) {
		GuiControl, , ModeSelector, filtering ...
		GuiControl, ChooseString, ModeSelector, filtering ...
		Gosub, UpdateList
		GuiControl, , ModeSelector, |
		GuiControl, , ModeSelector, %modeList%
		GuiControl, ChooseString, ModeSelector, %FilterMode%
		GuiControl, Focus, SearchInput
		return
	} else if (FilterMode = "Bookmarks") {
		EnvGet, localAppData, LOCALAPPDATA
		bookmarksFile := Config.Path["Bookmarks"]
		if !FileExist(bookmarksFile)
			return
		FileRead, rawJson, %bookmarksFile%
		currentName := ""
		Loop, Parse, rawJson, `n, `r
		{
			line := Trim(A_LoopField)
			if (RegExMatch(line, """name"":\s*""(.*?)""", nameMatch)) {
				currentName := nameMatch1
			} else if (RegExMatch(line, """url"":\s*""(.*?)""", urlMatch)) {
				CachedList.Push({title: currentName, path: urlMatch1})
				currentName := ""  ; Reset für nächsten Block
			}
		}
	} else if (FilterMode = "Favorites") {
		Loop, Files, %favFolder%\*.lnk
		{
			FileGetShortcut, %A_LoopFileFullPath%, target
			if (target = "" || !FileExist(target))
				continue
			FileGetTime, modTime, %A_LoopFileFullPath%, M
			if ErrorLevel
				continue
			displayName := StrReplace(A_LoopFileName, ".lnk", "")
			CachedList.Push({path: target, title: displayName})
		}
	} else if (FilterMode = "Media") {
		wmp := ComObjCreate("WMPlayer.OCX")
		mediaCollection := wmp.mediaCollection
		allItems := mediaCollection.getAll()

		Loop, % allItems.count
		{
			item := allItems.Item(A_Index - 1)
			title := item.getItemInfo("Title")
			artist := item.getItemInfo("Artist")
			path := item.sourceURL
			if (title != "")
				CachedList.Push({title: title " / " artist, path: path})
		}
	} else  {
		WindowClassFilter :=
	}
	Gosub, UpdateList
	GuiControl, Focus, SearchInput
	return
}
ToggleCase:
{
	GuiControlGet, CaseSensitive,, CaseToggle
	GuiControl,, CaseToggle, % CaseSensitive ? "CaSe" : "case"
	Gosub, UpdateList
	GuiControl, Focus, SearchInput
	return
}
ListBoxChanged:
{
	GuiControlGet, selectedTitle,, WindowBox
	SelectedIndex := 0
	Loop, % ItemList.Length()
	{
		if (ItemList[A_Index].title = selectedTitle) {
			SelectedIndex := A_Index
			break
		}
	}
	Gosub, selection
	return
}

;====================
IsExcluded(item, mode) {
	Loop % Config.Exclude.Length() {
		scope := Config.Exclude[A_Index].scope
		filter := Config.Exclude[A_Index].filter
		if (scope = "*" or scope = mode) {
			if (RegExMatch(item, filter))
				return True
			}
	}
	return False
}

ResolvePath(path) {
	 if RegExMatch(path, "^\.\\")
		  return A_ScriptDir "\" SubStr(path, 3)
	 else if RegExMatch(path, "^~\\") {
		  EnvGet, userProfile, USERPROFILE
		  return userProfile "\" SubStr(path, 3)
	 } else if RegExMatch(path, "^appData:\\") {
		  EnvGet, appData, APPDATA
		  return appData "\Filtronaut\" SubStr(path, 3)
	 } else if RegExMatch(path, "^home:\\") {
		  EnvGet, homeShare, HOMESHARE
		  EnvGet, homePath, HOMEPATH
		  return homeShare homePath "\" SubStr(path, 7)
	 } else if RegExMatch(path, "^onedrive:\\") {
		  EnvGet, oneDrive, OneDrive
		  if (oneDrive = "")
				EnvGet, oneDrive, OneDriveCommercial
		  if (oneDrive = "")
				EnvGet, oneDrive, OneDriveConsumer
		  return oneDrive "\" SubStr(path, 10)
	 } else
		  return path
}

RenameFile(oldLink, newName) {
	if (newName = "")
		return ""
	newName := RegExReplace(newName, "[<>:""/\\|?*]+", "_")

	SplitPath, oldLink, , oldDir
	newLink := oldDir "\" newName ".lnk"

	if (FileExist(newLink)) {
		MsgBox, 48, , There is already a favorite called '%newName%'.
		return ""
	}
	if (!FileExist(oldLink)) {
		MsgBox, 48, , Strange: '%EditedItem%' could not be found ('%oldLink%').
		return ""
	}
	FileMove, %oldLink%, %newLink%
	if (ErrorLevel) {
		MsgBox, 48, , Error on renaming '%oldLink%' to '%newLink%').
		return ""
	}
	return newLink
}

SetEditVisual(isOn := true) {
	if (isOn) {
		Gui, Font, Italic
	} else {
		Gui, Font, Norm
	}
	GuiControl, Font, SearchInput
	Gui, Font
}

;====================
UpdateList:
{
	global ItemList, SelectedIndex, CaseSensitive, FilterMode
	GuiControlGet, FilterText,, SearchInput
	ItemList := []
	GuiControl,, WindowBox, |

	if (FilterMode = "Open windows") {
		if WindowClassFilter
			WinGet, idList, List, ahk_class %WindowClassFilter%
		else
			WinGet, idList, List
		Loop, % idList
		{
			this_id := idList%A_Index%
			WinGetTitle, title, ahk_id %this_id%
			WinGetClass, class, ahk_id %this_id%
			if (title != "" && title != "Filtronaut" && title != "Program Manager"
					&& InStr(title, FilterText, CaseSensitive ? 1 : 0)
					&& class != "PopupHost"
					&& !IsExcluded(title, FilterMode)) {
				cleanTitle := StrReplace(title, "|", ">>>")
				ItemList.Push({id: this_id, title: cleanTitle})
				GuiControl,, WindowBox, %cleanTitle%
			}
		}
	} else if (FilterMode = "Favorites" || FilterMode = "Bookmarks" || FilterMode = "Media" || Config.Sniplets.HasKey(FilterMode)) {
		 for index, item in CachedList {
			 if InStr(item.title, FilterText) {
				 ItemList.Push(item)
				 GuiControl,, WindowBox, % item.title
			 }
		 }
	} else if IsRecentBased(FilterMode) {
		filterRegex := Config.Modes[FilterMode]

		recentFolder := A_AppData "\Microsoft\Windows\Recent"
		if (!RecentItemListBuilt) {
			tempList := []

			Loop, Files, %recentFolder%\*.lnk
			{
				FileGetShortcut, %A_LoopFileFullPath%, target
				if (target = "" || !FileExist(target))
					continue
				FileGetTime, modTime, %A_LoopFileFullPath%, M
				if ErrorLevel
					continue

				displayName := StrReplace(A_LoopFileName, ".lnk", "")
				tempList.Push({path: target, title: displayName, time: modTime, link: A_LoopFileFullPath})
			}
			tempList.Sort("time D")
			RecentItemList := tempList
			RecentItemListBuilt := true
		}

		for index, item in RecentItemList {
			if (FilterMode = "Directories" && !InStr(FileExist(item.path), "D"))
				continue
			if (!RegExMatch(item.path, filterRegex))
				continue
			if (!InStr(item.title, FilterText, CaseSensitive ? 1 : 0))
				continue
			if (IsExcluded(item.title, FilterMode))
				continue

			ItemList.Push(item)
			GuiControl,, WindowBox, % item.title
		}
	} else {
		MsgBox, Unhandled FilterMode %FilterMode%
	}
	if (ItemList.Length() > 0) {
		SelectedIndex := PresetIndex
		PresetIndex := 1
		GuiControl, Choose, WindowBox, %SelectedIndex%
	}
	return
}

;====================
Selection:
{
	if (SelectedIndex < 1 || SelectedIndex > ItemList.Length())
		return
	Hotkey := A_ThisHotkey
	selectedItem := ItemList[SelectedIndex]
	windowId := ""

	if (Config.Actions.HasKey(Hotkey)) {
		action := Config.Actions[Hotkey]
		monitorNr := action.monitor
	 } else
		  action := {}
	if (FilterMode = "Open windows") {
		windowId := selectedItem.id
		WinActivate, ahk_id %windowId%
	} else if (Config.Sniplets.HasKey(FilterMode)) {
		MyWindowId := 0
		Gui, Destroy
		  if (FilterMode = action.sniplet || action.sniplet = "*") {
			command := StrReplace(action.command, "%s", selectedItem.title)
				Run, %command%
		  }
		  else
				SendInput, % selectedItem.title
		return
	} else if (FilterMode = "Media") {
		path := selectedItem.path
		Run, wmplayer.exe "%path%",, UseErrorLevel
	} else {
		Run, % selectedItem.path,,, pid
		  if (action.monitor) {
				if (pid != "") {
					 WinWait, ahk_pid %pid%,, 4
					 windowId := WinExist("ahk_pid " pid)
				} else {
					 WinWaitActive,,, 4
					 windowId := WinExist("A")
				}
		}
	}
	if (windowId && action.monitor) {
		dim := StrSplit(action.layout, ",", " ")
		; MsgBox hotkey %Hotkey%, action %action% --> Config.Actions["!1"] :: %monitorNr% ::: %dim%, 3
		SysGet, MonitorCount, MonitorCount
		if (monitorNr <= MonitorCount) {
			SysGet, Mon, MonitorWorkArea, %monitorNr%
			MonWidth := MonRight - MonLeft
			MonHeight := MonBottom - MonTop

			WinRestore, ahk_id %windowId%
			WinGetPos, WinX, WinY, WinW, WinH, ahk_id %windowId%
			WinW := min(WinW, MonWidth)
			WinH := min(WinH, MonHeight)
			NewX := MonLeft + (MonWidth - WinW) // 2
			NewY := MonTop + (MonHeight - WinH) // 2
			; MsgBox, %dim% - %WinX% - %WinY% - %WinW% - %WinH% : %NewX% - %NewY% : %MonLeft% - %MonRight%

			if (dim[1] = "x") {
				WinMove, ahk_id %windowId%, , NewX, NewY, WinW, WinH
				WinMaximize, ahk_id %windowId%
			} else if (dim[1] = "c") {
				WinMove, ahk_id %windowId%, , NewX, NewY, WinW, WinH
			} else if (dim[1] = "%") {
				NewX := MonLeft + (dim[5] * MonWidth) // 100
				WinW := ((100 - dim[5] - dim[3]) * MonWidth) // 100
				NewY := MonTop + (dim[2] * MonHeight) // 100
				WinH := ((100 - dim[2] - dim[4]) * MonHeight) // 100
				WinMove, ahk_id %windowId%, , NewX, NewY, WinW, WinH
			}
		}
	}
	MyWindowId := 0
	Gui, Destroy
	return
}

;====================
SnipletsChanged:
{
	SavePending := FilterMode
	if (!Config.Hotkey.HasKey("Save"))
		Gosub, SaveSniplets
	return
}
CheckPendingSaves:
{
	if (SavePending) {
		MsgBox, 4100,, Save unsaved changes?
		IfMsgBox, Yes
			Gosub, SaveSniplets
	}
	SavePending := ""
	return
}
SaveSniplets:
{
	if !Config.Sniplets.HasKey(SavePending)
		return

	path := Config.Sniplets[SavePending]
	FileDelete, %path%

	Loop, % CachedList.MaxIndex()
	{
		item := CachedList[A_Index]
		FileAppend, % item.title . "`n", %path%
	}
	SavePending := ""
	return
}

;====================
HandleModeHotkey:
{
	if !WinActive("ahk_class AutoHotkeyGUI") {
		SendInput, %A_ThisHotkey%
		return
	}
	if (Config.FilterModes.HasKey(A_ThisHotkey)) {
		FilterMode := Config.FilterModes[A_ThisHotkey]
		GuiControl, ChooseString, ModeSelector, %FilterMode%
		Gosub, ModeChanged
		return
	} else if (Config.Classes.HasKey(A_ThisHotkey)) {
		WindowClassFilter := Config.Classes[A_ThisHotkey]
		Gosub, UpdateList
	} else {
		MsgBox, Unknown hotkey %A_ThisHotkey%
		return
	}
	return
}
#IfWinActive ahk_class AutoHotkeyGUI
!o::
!d::
!r::
!w::
!x::
!p::
!b::
!f::
!m::
	Gosub, HandleModeHotkey
	return

!Enter::
{
	if (EditedItem) {
		if (Config.Sniplets.HasKey(FilterMode)) {
			; update edited sniplet text
			Loop, % CachedList.MaxIndex()
			{
				if (CachedList[A_Index].title = EditedItem) {
					GuiControlGet, FilterText,, SearchInput
					CachedList[A_Index].title := FilterText
					EditedItem := FilterText
					Gosub, SnipletsChanged
					Gosub, UpdateList
					break
				}
			}
		} else if (IsRecentBased(FilterMode)) {
			; rename Recent-Item link
			GuiControlGet, FilterText,, SearchInput
			newLink := RenameFile(RecentItemList[EditedItem].link, FilterText)
			if (newLink = "")
				return
			RecentItemList[EditedItem].title := FilterText
			RecentItemList[EditedItem].link  := newLink
			Gosub, UpdateList
		} else if (FilterMode = "Favorites") {
			; rename favorites link
			GuiControlGet, FilterText,, SearchInput
			oldLink := favFolder "\" EditedItem ".lnk"
			newLink := RenameFile(oldLink, FilterText)
			if (newLink = "")
				return
			Gosub, ModeChanged
			EditedItem := FilterText
		}
	} else if (SelectedIndex >= 1 && SelectedIndex <= ItemList.Length()) {
		SearchInput := ItemList[SelectedIndex].title
		ControlSetText, Edit1, %SearchInput%, Filtronaut
		SendInput, ^a

		if (IsRecentBased(FilterMode)) {
			; remember index to change if edited
			sel := ItemList[SelectedIndex]
			if (sel.HasKey("link") && sel.link != "") {
				Loop % RecentItemList.Length() {
					if (RecentItemList[A_Index].link = sel.link) {
						EditedItem := A_Index
						SetEditVisual(true)
						break
					}
				}
			}
		} else {
			; Sniplets & Favorites: remember old title to replace if edited
			EditedItem := SearchInput
			SetEditVisual(true)
		}
	}
	return
}

~Enter:: Gosub, selection
!1:: Gosub, selection

!c::
{
	CaseSensitive := !CaseSensitive
	GuiControl,, CaseToggle, % CaseSensitive ? 1 : 0
	GuiControl,, CaseToggle, % CaseSensitive ? "CaSe" : "case"
	Gosub, UpdateList
	return
}

^h::
{
	Gosub, showHelp
	return
}

Tab::
{
	ControlGetFocus, focusedControl, A
	if (focusedControl != "Edit1")
		return

	GuiControlGet, FilterText,, SearchInput
	if (ItemList.Length() < 2)
		return

	prefix := FilterText
	Loop
	{
		nextChar := ""
		Loop, % ItemList.Length()
		{
			title := ItemList[A_Index].title
			start := InStr(title, prefix)
			if (!start || StrLen(title) < start + StrLen(prefix))
				return
			char := SubStr(title, start + StrLen(prefix), 1)
			if (nextChar = "")
				nextChar := char
			else if (char != nextChar)
				return ; no more matches
		}
		SendInput, %nextChar%
		Sleep, 5
		prefix .= nextChar
	}

	ControlSetText, Edit1, %prefix%, Filtronaut
	SearchInput := prefix
	return
}

^Backspace::
{
	if (SelectedIndex < 1 || SelectedIndex > ItemList.Length())
		return

	selectedItem := ItemList[SelectedIndex]

	if (FilterMode = "Open windows") {
		windowId := selectedItem.id
		WinClose, ahk_id %windowId%
	} else if (Config.Sniplets.HasKey(FilterMode)) {
		Loop, % CachedList.MaxIndex()
		{
			if (CachedList[A_Index].title = selectedItem.title) {
				CachedList.RemoveAt(A_Index)
				Gosub, SnipletsChanged
				break
			}
		}
	} else if (selectedItem.HasKey("link") && FileExist(selectedItem.link)) {
		FileDelete, % selectedItem.link
		Loop % RecentItemList.Length() {
			if (RecentItemList[A_Index].link = selectedItem.link) {
				RecentItemList.RemoveAt(A_Index)
				break
			}
		}
	} else if (FilterMode = "Favorites") {
		favLink := favFolder "\" selectedItem.title ".lnk"
		if FileExist(favLink) {
			FileDelete, %favLink%
			Gosub, ModeChanged ; to rebuild CachedList
		}
	}
	OldSelection := SelectedIndex
	Gosub, UpdateList
	SelectedIndex := Min(OldSelection, ItemList.Length())
	GuiControl, Choose, WindowBox, %SelectedIndex%
	return
}

^+:: ; Ctrl-Plus to add Sniplet or Favorite
{
	if (Config.Sniplets.HasKey(FilterMode)) {
		GuiControlGet, newText,, SearchInput
		if (newText != "") {
			path := Config.Sniplets[FilterMode]
			CachedList.Push({title: newText, path: path})
			if (Config.Hotkey.HasKey("Save"))
				SavePending := FilterMode
			else
				FileAppend, %newText%`n, %path%
			GuiControl,, SearchInput,
			SearchInput := ""
			PresetIndex := CachedList.Length()
		}
	} else if IsRecentBased(FilterMode) {
		if (SelectedIndex >= 1 && SelectedIndex <= ItemList.Length()) {
			selectedItem := ItemList[SelectedIndex]
			favLink := favFolder "\" selectedItem.title ".lnk"
			FileCreateShortcut, % selectedItem.path, %favLink%
		}
	}
	return
}

Up::
Down::
{
	if (A_ThisHotkey = "Up")
		SelectedIndex := Max(1, SelectedIndex - 1)
	else
		SelectedIndex := Min(ItemList.Length(), SelectedIndex + 1)

	GuiControl, Choose, WindowBox, %SelectedIndex%
	return
}
^Up::
^Down::
{
	if (Config.Sniplets.HasKey(FilterMode) && SelectedIndex >= 1 && SelectedIndex <= ItemList.Length()) {
		selectedTitle := ItemList[SelectedIndex].title
		currentIndex := 0
		newIndex := 1
		Loop, % CachedList.MaxIndex()
		{
			if (CachedList[A_Index].title = selectedTitle) {
				if (A_ThisHotkey = "^Up") {
					movedItem := CachedList.RemoveAt(A_Index)
					CachedList.InsertAt(newIndex, movedItem)
					Gosub, SnipletsChanged
					PresetIndex := Max(1, SelectedIndex - 1)
					Gosub, UpdateList
					return
				} else
					currentIndex := A_Index
			}
			else if (InStr(CachedList[A_Index].title, FilterText, CaseSensitive ? 1 : 0)) {
				newIndex := A_Index
				if (A_ThisHotkey = "^Down" && currentIndex > 0)
					break
			}
		}
	}
	if (newIndex > currentIndex) {
		movedItem := CachedList.RemoveAt(currentIndex)
		CachedList.InsertAt(Min(newIndex, CachedList.Length() + 1), movedItem)
		Gosub, SnipletsChanged
		PresetIndex := Min(SelectedIndex + 1, CachedList.Length())
		Gosub, UpdateList
	}
	return
}
^Esc:: Gosub, ExitApp
Esc::
GuiClose:
	MyWindowId := 0
	Gosub, CheckPendingSaves
	for hotkey, mode in Config.FilterModes {
		if hotkey not in %PredefinedHotkeys%
			Hotkey, %hotkey%, Off
	}
	for hotkey, mode in Config.Classes {
		if hotkey not in %PredefinedHotkeys%
			Hotkey, %hotkey%, Off
	}
	for hotkey, action in Config.Actions {
		if hotkey not in %PredefinedHotkeys%
			Hotkey, %hotkey%, Off
	}
	if (Config.Hotkey.HasKey("Save"))
		Hotkey, % Config.Hotkey["Save"], Off
	Gui, Destroy
	return
#IfWinActive

ExitApp:
	Gosub, CheckPendingSaves
ExitApp
