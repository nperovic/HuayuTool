/************************************************************************
 * @description A tool for Teaching Chinese 
 * @file HuayuTool.ahk
 * @author Nikola Perovic 陸汎宇
 * @date 2023/07/03
 * @version 1.1.0
 ***********************************************************************/

#Requires AutoHotKey v2
#SingleInstance Force
#ErrorStdOut 
#Warn All, Off

A_ScriptName  := "華語教學小工具 v1.1"
A_IconTip     := A_ScriptName
A_MenuMaskKey := "vkFF"
hk            := ""

Class Gui_Ext extends Gui
{
	Class _Edit extends Gui.Edit
	{
		Static __New() => (
			super.Prototype.SetCue := ObjBindMethod(this, "SetCue")
		)

		/**
		 * Sets the textual cue, or tip, that is displayed by the edit control to prompt the user for information.
		 * Link: https://docs.microsoft.com/en-us/windows/win32/controls/em-setcuebanner
		 * @param {any} _obj obj.hwnd
		 * @param {any} string 
		 * @param {number} option
		 * * True  -> if the cue banner should show even when the edit control has focus
		 * * False -> if the cue banner disappears when the user clicks in the control
		 * @returns {number}
		 */
		Static SetCue(_obj, string, option := true)
		{
			Static ECM_FIRST := 0x1500
			Static EM_SETCUEBANNER := ECM_FIRST + 1

			If (DllCall("user32\SendMessage", "ptr", _obj.hwnd, "uint", EM_SETCUEBANNER, "int", option, "str", string, "int"))
				Return true
			Return false
		}
	}

	Class _ComboBox extends Gui.ComboBox
	{
		Static __New()
		{
			Super.Prototype.SetCue := ObjBindMethod(this, "SetCue")
		}

		Static SetCue(_obj, string, option := true)
		{
			; Links ....................:  https://docs.microsoft.com/en-us/windows/win32/controls/cb-setcuebanner
			; Description ..............:  Sets the cue banner text that is displayed for the edit control of a combo box.
			Static CBM_FIRST := 0x1700
			Static CB_SETCUEBANNER := CBM_FIRST + 3

			If (DllCall("user32\SendMessage", "ptr", _obj.hwnd, "uint", CB_SETCUEBANNER, "int", option, "str", string, "int"))
				Return true
			Return false
		}
	}

	Class _Control extends Gui
	{
		Static __New() => (
			Gui.Control.Prototype.SetDarkMode := ObjBindMethod(this, "SetDarkMode")
		)

		Static SetDarkMode(_obj) => (
			DllCall("uxtheme\SetWindowTheme", "ptr", _obj.hwnd, "str", "DarkMode_Explorer", "ptr", 0) ? true : false
		)
	}

	; Thanks to jNizM for the interesting DllCalls
	; https://www.autohotkey.com/boards/viewtopic.php?t=70852

	Static __New()
	{
		Gui.Prototype.SetDarkTitle := ObjBindMethod(this, "SetDarkTitle")
		Gui.Prototype.SetDarkMenu := ObjBindMethod(this, "SetDarkMenu")
	}

	Static SetDarkTitle(_obj)
	{
		If VerCompare(A_OSVersion, "10.0.17763") >= 0
		{
			attr := 19
			If VerCompare(A_OSVersion, "10.0.18985") >= 0
			{
				attr := 20
			}
			If (DllCall("dwmapi\DwmSetWindowAttribute", "ptr", _obj.hwnd, "int", attr, "int*", true, "int", 4))
				Return true
		}
		Return false
	}

	Static SetDarkMenu(_obj)
	{
		uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr"),
			SetPreferredAppMode := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr"),
			FlushMenuThemes := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 136, "ptr"),
			DllCall(SetPreferredAppMode, "int", 4)	; Dark
		Return (DllCall(FlushMenuThemes) ? true : false)
	}
}

readhk() => IniRead("setting.ini", "Section1", "HotKey", "MButton")

InstallMouseHook
InstallKeybdHook

myMenu := Menu()
myMenu.Add "拼音輸入模式`t未開啟", pinyin
myMenu.Add "拼音輸入模式使用說明", (*) => MSGBOX
myMenu.Add "查詢所選文字的筆順", search
myMenu.Add "查詢所選文字的繁、簡、英、拼", superFunc
myMenu.Add
myMenu.Add "Google 搜尋所選內容", search
myMenu.Add "Google 翻譯所選內容", translate
myMenu.Add
myMenu.Add "簡體轉繁體", fanJian
myMenu.Add "繁體轉簡體", fanJian
myMenu.Add
myMenu.Add "修復工作列", reExplorer
myMenu.Add
myMenu.Add "設定快捷鍵`t" (hk := NameHK(readhk())), SetHK
myMenu.Add
myMenu.Add "前往華語中心首頁", (*) => Run('"https://c040e.wzu.edu.tw"')
myMenu.Add
myMenu.Add "贊助開發者", (*) => Run('"https://niko.soci.vip/donate"')
myMenu.Add
myMenu.Add "重新載入", (*) => Reload()
myMenu.Add "結束",   (*) => ExitApp()

For v in [135, 136]
	DllCall DllCall("GetProcAddress", "ptr", DllCall("GetModuleHandle", "str", "uxtheme", "ptr"), "ptr", v, "ptr"), "int", 2

HotKey hk := readhk(), setHKNew

OnMessage 0x404, TrayIconClickEvent

MBB "
(
	如何開啟功能表:
	1. 按下滑鼠中鍵(滾輪)。
	2. 按住左鍵的同時按下右鍵。
	
	註：可在功能表中自訂快捷鍵。

	---
	開發者: 陸汎宇 Nikola Perovic
	Email: 2@u4ni.tk 
)",  A_ScriptName,, Map(
	"自訂快捷鍵", SetHK,
	"關閉視窗", (myctrlObj, *) => myctrlObj.Gui.Destroy()
)

#MaxThreadsPerHotkey 2
~LButton & RButton:: myMenu.Show(,,0)
#MaxThreadsPerHotkey 1

#HotIf GetKeyState('vkFF')
~LButton:: Send "{Blind}{vkFF}"
#HotIf

~LCtrl::
{
	Try if A_TimeSincePriorHotkey < 200 && A_PriorHotkey = A_ThisHotkey
		myMenu.Show(,,0)
}


setHKNew(*)
{
	global hk

	myMenu.Show(,, 0)
	
	if hk ~= "MButton|LButton|RButton"
		KeyWait hk, "T5"
}

TrayIconClickEvent(wParam, lParam, msg, hwnd)
{
	; 0x201: L
	; 0x202: L
	; 0x203: L*2
	; 0x205: R
	; 0x208: M

    switch lParam {
    case 0x202: 

	KeyWait "LButton", "T5"
	myMenu.Show(,,0)

    case 0x208: return
    case 0x205: return
    }
}

MouseIsOver(WinTitle) {
	CoordMode "Mouse"
	SetWinDelay -1
	Try MouseGetPos , , &Win
	Try result := WinExist(WinTitle " ahk_id " . Win)
	Return result
}

superFunc(item?, *)
{
	Static myGuiHwnd   := 0
	       savedClip   := ClipboardAll()
	       A_Clipboard := ""
	
	Sleep 200
	Send "{RCtrl Down}{Ins}{RCtrl Up}"
	cw := ClipWait(1)

	if !cw || A_Clipboard = ""
		Return TrayTip("您未選取任何文字。", A_ScriptName)

	Fan  := A_Clipboard
	Jian := googleTranslate(Fan, "zh-CN").text
	sleep 100
	En   := googleTranslate(Fan, "en-US").text

	pin      := "",
	doneNum  := 20,
	startNum := 1,
	len      := strlen(fan)

	Loop {
		A_Clipboard := ""
		wrap        := UrlUnescape2(len > 20 ? SubStr(fan, startNum, doneNum) : Fan)

		Run '"powershell.exe" -c ' "$url = (Invoke-WebRequest -Uri 'https://crptransfer.moe.gov.tw/index.jsp?SN=" . wrap . "&sound=1#res'" '-UseBasicParsing).Content | Set-Clipboard"'
		,, 'Hide'

		if IsSpace(A_clipboard)
			ClipWait 3

		RegExMatch(A_Clipboard, "mS)(?:<tr><th axis='資料類型'>漢語拼音<\/th><td>)\K.*(?:<\/span>)", &m)

		pin      .= StrReplace(StrReplace((A_Index > 1) ? (A_Space . m[]) : m[], "<span class=long>"), "</span>", A_Space)
		startNum += 20,
		doneNum  += 20
		
		Sleep 30

	} Until len <= startNum

	result := Fan "`n" Jian "`n" pin "`n" En

	if !myGuiHwnd || !WinExist(myGuiHwnd)
		myGuiHwnd := (myGui := MBB(result, A_ScriptName,,,, 18)).hwnd
	else 
		GuiFromHwnd(myGuiHwnd, 1)["Edit1"].Value := result

	SetTimer () => (A_Clipboard := savedClip), -100
}

PowerShell(Command, &StdOut?, &StdErr?)
{
	static shell                 := "",
	       interpreter           := "PowerShell",
	       ATTACH_PARENT_PROCESS := 0xFFFFFFFF
	
    if !shell
	{
        Run interpreter " -NoProfile -Command Exit", , "Hide"
        
        if (!DllCall("AttachConsole", "UInt", ATTACH_PARENT_PROCESS))
		{
            wd := A_WinDelay
			SetWinDelay -1
            DllCall "AllocConsole"
            WinHide DllCall("GetConsoleWindow", "Ptr")
            SetWinDelay wd
        }
        shell := ComObject("WScript.Shell")
    }

    exec   := shell.Exec(interpreter " -ExecutionPolicy 'bypass' -MTA -NoLogo -NonInteractive -Command " Command " | Out-String")
    StdOut := Trim(exec.StdOut.ReadAll(), "`t`n`r`s")
    StdErr := Trim(exec.StdErr.ReadAll(), "`t`n`r`s")
	
    return !!StrLen(StdErr)
}

NameHK(ini)
{
	Global hk
	phk  := readhk()
	kmap := Map("!", "Alt", "+", "Shift", "^", "Ctrl")

	; kmap.Default := A_Space
	For v in kmap
		phk := StrReplace(phk, v, kmap[v] " + ")

	return phk
}

SetHK(item?, pos?, myMenu?)
{
	myGui := Gui()
	myGui.Opt("-MinimizeBox -MaximizeBox +AlwaysOnTop")
	myGui.SetFont("s12", "Microsoft YaHei UI")
	myGui.Add("Hotkey", "vhk x16 y56 w255 h35")
	myGui.Add("Button", "x16 y104 w255 h35", "&OK").OnEvent("Click", OnEventHandler)
	ogcButtonOK := myGui.Add("Button", "x16 y104 w255 h35", "送出")
	myGui.SetFont("s12", "Microsoft YaHei UI")
	myGui.Add("Text", "x16 y8 w255 h35 +0x200", "按下要設定的快捷鍵後請按送出")
	myGui.OnEvent('Close', (mygui) => mygui.Destroy())
	myGui.Title := "設定快捷鍵 | " A_ScriptName
	myGui.Show("w292 h150")

	GuiEscape(GuiHwnd)
	{
		ExitApp()
	}

	OnEventHandler(guictrl, *)
	{
		Global hk
		submittedHK := guictrl.Gui.Submit()


		if !fileexist("setting.ini")
			FileAppend "[Section1]`nHotKey=" guictrl.Gui["hk"].Value, "setting.ini"
		Else
		{
			IniDelete "setting.ini", "Section1", "HotKey"
			IniWrite guictrl.Gui["hk"].Value, "setting.ini", "Section1", "HotKey"
		}

		HotKey readhk(), (*) => myMenu.Show()

		myMenu.rename item, "設定快捷鍵`t" (hk := NameHK(readhk()))

		guictrl.Gui.Destroy()
		MsgBox "快捷鍵已設定為 " hk, "華語教學小工具", "T10"
	}
}

translate(item, *)
{
	savedClip   := ClipboardAll()
	A_Clipboard := ""
	Send "{LCtrl Down}{Ins}{LCtrl Up}"

	ClipWait 1

	result      := GoogleTranslate(A_Clipboard).Text
	A_Clipboard := ""
	A_Clipboard := result

	ClipWait 1

	MBB result "`n`n(已將翻譯內容儲存至剪貼簿)", A_ScriptName
}

fanJian(item, *)
{
	savedClip   := ClipboardAll()
	A_Clipboard := ""
	
	Send "{LCtrl Down}{Ins}{LCtrl Up}"
	ClipWait 1

	result      := googleTranslate(A_Clipboard, item = "簡體轉繁體" ? "zh-TW" : "zh-CN").text
	A_Clipboard := ""
	A_Clipboard := result
	
	ClipWait 1
	Send "{LShift Down}{Ins}{LShift Up}"

	SetTimer () => (A_Clipboard := savedClip), -150
}

pinyinIntro(*)
{
	MsgBox "
(
	1. 請將輸入法切換至「英/數模式」。
	
	輸入方法 (以 a 為例): 
	a + 1 = ā
	a + 2 = á
	a + 3 = ǎ
	a + 4 = à
	a + 5 = a
	
	註: 請先關閉此視窗後再啟動拼音輸入模式。
	
	---
	開發者: 陸汎宇 Nikola Perovic
	Email: 2@u4ni.tk 
)", "拼音輸入模式使用說明 | 華語教學小工具"
}

search(item, *)
{
	Sleep 200

	savedClip   := ClipboardAll()
	A_Clipboard := ""
	
	Send "{LCtrl Down}{Ins}{LCtrl Up}"

here:
	if !ClipWait(1)
	{
		if !(MsgBox("要再試一次嗎?`n若要重試，請在按下重試後，於倒數時間內選取文字或複製文字。(複製比較快)", "未偵測到選取文字", "T10 0x5") ~= "Retry")
			Return

		CoordMode "Mouse"
		SetWinDelay -1

		time      := A_TickCount
		countdown := 5

		Loop
		{
			countdown := (5 - (A_TickCount - time) / 1000)
			MouseGetPos(&x, &y)
			ToolTip dcountDown := Round(countdown), x, y, 1

			if (dcountDown - countdown) = 0
			{
				if A_Clipboard
					break
			}
			Sleep 10
		}
		Until countdown < 0
		ToolTip()

		if !A_Clipboard
		{
			Send "{LCtrl Down}{Ins}{LCtrl Up}"
			goto here
		}
	}

	if item = "Google 搜尋所選內容"
		Run "https://www.google.com/search?q=" UrlUnescape2(A_Clipboard)
	Else
		Run '"https://www.twpen.com/' UrlUnescape2(A_Clipboard) '.html"'
}

reExplorer(item, *)
{
	ProcessSetPriority "Realtime"
	Run A_ComSpec " /C taskkill /f /IM explorer.exe && start explorer.exe", "C:\", "Hide"
}

pinyin(item?, pos?, myMenu?)
{
	static toggle := false
	static ih := InputHook("v B")
	static charmap := Map(
		"a1", "ā", "a2", "á", "a3", "ǎ", "a4", "à", "a5", "a", "e1", "ē", "e2", "é", "e3", "ě", "e4", "è", "e5", "e", "i1", "ī", "i2", "í", "i3", "ǐ", "i4", "ì", "i5", "i", "o1", "ō", "o2", "ó", "o3", "ǒ", "o4", "ò", "o5", "o", "u1", "ū", "u2", "ú", "u3", "ǔ", "u4", "ù", "u5", "u", "v1", "ǖ", "v2", "ǘ", "v3", "ǚ", "v4", "ǜ", "v5", "ü", "A1", "Ā", "A2", "Á", "A3", "Ǎ", "A4", "À", "A5", "A", "E1", "Ē", "E2", "É", "E3", "Ě", "E4", "È", "E5", "E", "I1", "Ī", "I2", "Í", "I3", "Ǐ", "I4", "Ì", "I5", "I", "O1", "Ō", "O2", "Ó", "O3", "Ǒ", "O4", "Ò", "O5", "O", "U1", "Ū", "U2", "Ú", "U3", "Ǔ", "U4", "Ù", "U5", "U", "V1", "Ǖ", "V2", "Ǘ", "V3", "Ǚ", "V4", "Ǜ", "V5", "Ü"
	)

	if toggle := !toggle
	{
		myMenu.Check item
		myMenu.Rename item, "拼音輸入模式`t已開啟"

		ih.OnChar := OnChar
		ih.Start()

		TrayTip "拼音輸入模式已開啟。 (請將輸入法切換至「英文模式」)"

		; HotKey "~Shift", (*) => (IME.switchIME(), pinyin("拼音輸入模式`t已開啟",, myMenu)), "On"
	}
	else
	{
		myMenu.UnCheck item
		myMenu.Rename item, "拼音輸入模式`t未開啟"

		ih.Stop()
		TrayTip "拼音輸入模式已關閉。"
	}

	OnChar(ih, char)
	{
		Static preText := ""

		if charmap.Has(preText . char)
			Send "{bs 2}{Text}" charmap[preText . char]

		preText := char
	}
}

currentKB()
{
	VarSetStrCapacity &pwszKLID, 9
	DllCall "GetKeyboardLayoutName", "Str", &pwszKLID
	return pwszKLID
}

UrlUnescape2(url, flags := 0x00100000) => (DllCall('Shlwapi.dll\UrlUnescape', 'Str', url, 'Ptr', 0, 'UInt', 0, 'UInt', flags, 'UInt') ? '' : url)




GoogleTranslate(text, languageTo := "zh-TW", languageFrom := "auto")
{
    if (!GoogleTranslate.HasOwnProp("extracted"))
    {
        hObject := ComObject("WinHttp.WinHttpRequest.5.1"), 
        hObject.Open("GET", "https://translate.google.com"), 
        hObject.SetRequestHeader("User-Agent", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"), 
        hObject.Send(), 
        lol                       := hObject.ResponseText,
        pos_FdrFJe                := InStr(lol, "FdrFJe", true),
        pos_quote_FdrFJe          := InStr(lol, "`"", true, pos_FdrFJe + 9),
        value_FdrFJe              := SubStr(lol, pos_FdrFJe + 9, pos_quote_FdrFJe - (pos_FdrFJe + 9)),
        pos_cfb2h                 := InStr(lol, "cfb2h", true),
        pos_quote_cfb2h           := InStr(lol, "`"", true, pos_cfb2h + 8),
        value_cfb2h               := SubStr(lol, pos_cfb2h + 8, pos_quote_cfb2h - (pos_cfb2h + 8)),
        GoogleTranslate.extracted := { FdrFJe: value_FdrFJe, cfb2h: value_cfb2h }
    }

    hObject := ComObject("WinHttp.WinHttpRequest.5.1")
    hObject.Open("POST", "https://translate.google.com/_/TranslateWebserverUi/data/batchexecute?rpcids=MkEWBc&source-path=%2F&f.sid=" GoogleTranslate.extracted.FdrFJe "&bl=" GoogleTranslate.extracted.cfb2h "&hl=en-US&soc-app=1&soc-platform=1&soc-device=1&_reqid=" Random(1000, 9999) "&rt=c")
    hObject.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=UTF-8")
    hObject.Send("f.req=" encodeURIComponent(JSON_stringify([[["MkEWBc", JSON_stringify([[text, languageFrom, languageTo, { Base: { __Class: "JSON_false" } }], [{ Base: { __Class: "JSON_null" } }]]), { Base: { __Class: "JSON_null" } }, 'generic']]])))
    lol := hObject.ResponseText

    pos_newline := InStr(lol, "`n", true, 8),
    fwef        := SubStr(lol, 7, pos_newline - 7),
    size        := Integer(SubStr(lol, 7, pos_newline - 7)) - 2,
    jsonTemp    := JSON_parse(SubStr(lol, pos_newline + 1, size)),
    json        := JSON_parse(jsonTemp[1][3])

    return { text: json[2][1][1][6][1][1], fromLanguage: json[3] }
}

JSON_parse(str)
{

    c_ := 1

    return JSON_value()

    JSON_value()
    {

        char_ := SubStr(str, c_, 1)
        Switch char_
        {
            case "{": 
            {
                obj_ := Map()
                c_++
                loop
                {
                    skip_s()
                    if (SubStr(str, c_, 1) == "}")
                    {
                        c_++
                        return obj_
                    }

                    if (SubStr(str, c_, 1) == "`"")
                    {
                        RegExMatch(str, "(?:\\.|.)*?(?=`")", &OutputVar, c_ + 1)
                        key_ := StrReplace(StrReplace(StrReplace(StrReplace(StrReplace(StrReplace(StrReplace(OutputVar.0, "\`"", "`"", true), "\f", "`f", true), "\r", "`r", true), "\n", "`n", true), "\b", "`b", true), "\t", "`t", true), "\\", "\", true)
                        c_   += OutputVar.Len
                    }
                    else
                    {
                        RegExMatch(str, ".*?(?=[\s:])", &OutputVar, c_)
                        key_ := OutputVar.0
                        c_   += OutputVar.Len
                    }

                    c_ := InStr(str, ":", true, c_) + 1
                    skip_s()

                    value_     := JSON_value(),
                    obj_[key_] := value_,
                    obj_.DefineProp(key_, { Value: value_ })

                    skip_s()
                    if (SubStr(str, c_, 1) == ",")
                    {
                        c_++, skip_s()
                    }
                }
            }
            case "[": 
            {
                arr_ := []
                c_++
                loop
                {
                    skip_s()
                    if (SubStr(str, c_, 1) == "]")
                    {
                        c_++
                        return arr_
                    }

                    value_ := JSON_value()
                    arr_.Push(value_)

                    skip_s()
                    char_ := SubStr(str, c_, 1)
                    if (char_ == ",")
                    {
                        c_++, skip_s()
                    }
                }
            }
            case "`"":
            {
                RegExMatch(str, '(?:\\.|.)*?(?=")', &OutputVar, c_ + 1)
                unquoted := StrReplace(StrReplace(StrReplace(StrReplace(StrReplace(StrReplace(StrReplace(OutputVar.0, "\`"", "`"", true), "\f", "`f", true), "\r", "`r", true), "\n", "`n", true), "\b", "`b", true), "\t", "`t", true), "\\", "\", true)
                c_       += OutputVar.Len + 2
                return unquoted
            }
            case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9": 
            {
                RegExMatch(str, "[0-9.eE\-+]*", &OutputVar, c_)
                c_ += OutputVar.Len
                return Number(OutputVar.0)
            }
            case "t": 
            {
                c_ += 4
                return { Base: { __Class: "JSON_true" } }
            }
            case "f": 
            {
                c_ += 5
                return { Base: { __Class: "JSON_false" } }
            }
            case "n": 
            {
                c_ += 4
                return { Base: { __Class: "JSON_null" } }
            }
        }
    }
    skip_s()
    {
        RegExMatch(str, "\s*", &OutputVar, c_)
        c_ += OutputVar.Len
    }
}


JSON_stringify(obj, maxDepth := 5)
{

    stringified := ""

    escape(str)
    {
        rp := Map(
            "0\\",  "\\", 
            "1`t",  "\t",
            "2`b",  "\b",
            "3`n",  "\n",
            "4`r",  "\r",
            "5`f",  "\f",
            "6`"",  "\`""
        )
        For k, v in rp 
            str := StrReplace(str, SubStr(k,2),  v, true)

        return str
    }
    ok(obj, depth)
    {
        switch (Type(obj))
        {
            case 'Map':
            {
                if (depth > maxDepth)
                    stringified .= "`"[DEEP ...Map]`""
                else
                {
                    stringified .= "{"
                    for k, v in obj
                    {
                        (A_Index > 1 && stringified .= ",")
                        stringified .= "`"" escape(k) "`": "
                        ok(v, depth + 1)
                    }
                    stringified .= "}"
                }
            }
            case 'Object': 
            {
                if (depth > maxDepth)
                {
                    stringified .= "`"[DEEP ...Object]`""
                }
                else
                {
                    stringified .= "{"
                    for k, v in obj.OwnProps()
                    {
                        (A_Index > 1 && stringified .= ",")
                        stringified .= "`"" escape(k) "`": "
                        ok(v, depth + 1)
                    }
                    stringified .= "}"
                }
            }
            case 'Array': 
            {
                if (depth > maxDepth)
                {
                    stringified .= "`"[DEEP ...Array]`""
                }
                else
                {
                    stringified .= "["
                    for v in obj
                    {
                        (A_Index > 1 && stringified .= ",")
                        ok(v, depth + 1)
                    }
                    stringified .= "]"
                }
            }
            case 'String'    : (stringified .= "`"" escape(obj) "`"")
            case "Integer"   : (stringified .= obj)
            case "Float"     : (stringified .= obj)
            case "JSON_true" : (stringified .= "true")
            case "JSON_false": (stringified .= "false")
            case "JSON_null" : (stringified .= "null")
                        
        }
    }
    ok(obj, 0)

    return stringified

}

encodeURIComponent(str)
{
    static arr := [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "!", 0, 0, 0, 0, 0, "'", "(", ")", "*", 0, 0, "-", ".", 0, "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", 0, 0, 0, 0, 0, 0, 0, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", 0, 0, 0, 0, "_", 0, "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", 0, 0, 0, "~", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,]

    size := StrPut(str, "UTF-8"),
    buf  := Buffer(size),
    
    StrPut(str, buf, "UTF-8"),

    i_           := 0,
    sizeMinusOne := size - 1,
    finalStr     := ""

    while (i_ < sizeMinusOne)
    {
        uChar := NumGet(buf, i_, "UChar")
        if (type(arr[uChar + 1]) == "String")
            finalStr .= arr[uChar + 1]
        else
            finalStr .= "%" Format("{:02X}", uChar)

        ++i_
    }

    return finalStr
}

MBB(tTextToDisplay     := "- Empty -"
	, Title            := ""
	, iGuiSleep        := unset
	, btnMap           := Map("複製內文", lMyGuiCopyBtn_Click, "關閉視窗", lMyGui_EscapeClose)
	, &editHWND        := unset
	, iFontSize        := 14
	, sFontName        := "Microsoft JhengHei UI"
	, sTheme           := "dark"
	, sGuiPosData      := "center"
	, sEditScrollToBtm := "yes"
	, sGuiDims         := "auto"
)
{

	status := "TIMEOUT"
	StrReplace(tTextToDisplay, "`n", , , &iNoOfLines)
	StrReplace(sGuiPosData, A_Space, , , &iNoOfSpaces)

	; ============================================================================================
	; ------------------------------------- ERROR MESSAGE ----------------------------------------
	; ============================================================================================

	if (iNoOfSpaces = StrLen(tTextToDisplay))
		ErrorMessage("tTextToDisplay", "The message text is blank or only contains spaces")

	if (tTextToDisplay = "")
		ErrorMessage("tTextToDisplay", "No message text has been passed to the funtion")

	if IsSet(iGuiSleep) && (InStr(iGuiSleep, "-") OR !IsInteger(iGuiSleep))
		ErrorMessage("iGuiSleep", "It either has a -ve sign or is not an intger")

	if (iFontSize < 6 or iFontSize > 24)
		ErrorMessage("iFontSize", "It should be: 6 <= iFontSize <= 24")

	if InStr(iFontSize, "-") OR !IsInteger(iFontSize)
		ErrorMessage("iGuiSleep", "It either has a -ve sign or is not an intger")

	if (sTheme != "light") AND (sTheme != "dark")
		ErrorMessage("sTheme", "The specified theme is neither `"light`" or `"dark`"")


	if (iNoOfSpaces = 0) AND (StrLen(sGuiPosData) = 0)
		ErrorMessage("sGuiPosData", "The position data is blank")
	else if (iNoOfSpaces = StrLen(sGuiPosData))
		ErrorMessage("sGuiPosData", "The position data only contains spaces")

	if (sEditScrollToBtm != "yes") AND (sEditScrollToBtm != "no")
		ErrorMessage("sEditScrollToBtm", "The specified scroll should be either `"yes`" or `"no`"")

	if (sTheme = "light")
	{
		hBackColor := "Background" . "",
			hTextColor := ""
	}
	Else if (sTheme = "dark")
	{
		hBackColor := "Background" . "302C29 ",
			hTextColor := "c" . "C0B0A0"
	}

	iMaxNoOfChars           := 0,
	iHalfScreenWidth        := A_ScreenWidth / 2,
	iHalfScreenHeight       := A_ScreenHeight / 2,
	iExtraLinesFromLineWrap := 0,
	iEditWidthDelta         := 20,
	iEditHeightDelta        := 40

	if (sGuiDims = "auto")
	{
		Loop Parse, tTextToDisplay, "`n"
		{
			tNoOfChars := STRLEN(A_LoopField)
			If (tNoOfChars > iMaxNoOfChars)
				iMaxNoOfChars := tNoOfChars
			If ((tNoOfChars * iFontSize) > iHalfScreenWidth)
				iExtraLinesFromLineWrap += 1
		}
		iGuiWidth := iMaxNoOfChars * iFontSize + iEditWidthDelta,
		StrReplace(tTextToDisplay, "`n", , , &iNoOfLines),
		iNoOfLines += 1 + iExtraLinesFromLineWrap,
		iGuiHeight := iNoOfLines * iFontSize * 2 + iEditHeightDelta
	}
	else
	{
		aGuiDims := StrSplit(sGuiDims, A_Space)
		iGuiWidth := aGuiDims[1]
		iGuiHeight := aGuiDims[2]

		if (iGuiWidth < 320)
			MsgBox("Gui width is set <320. Script will set min as 320 and continue.")
		if (iGuiWidth > iHalfScreenWidth)
			MsgBox("Gui width is set >HalfScreenWidth.`n`nScript will set max width as HalfScreenWidth and continue.")
		if (iGuiHeight < 120)
			MsgBox("Gui height is set <120. Script will set min as 120 and continue.")
		if (iGuiHeight > iHalfScreenHeight)
			MsgBox("Gui height is set >iHalfScreenHeight.`n`nScript will set max height as iHalfScreenHeight and continue.")
	}

	if (iGuiWidth < 320)
		iGuiWidth := 320
	else if (iGuiWidth > iHalfScreenWidth)
		iGuiWidth := iHalfScreenWidth

	if (iGuiHeight < 120)
		iGuiHeight := 120
	else if (iGuiHeight > iHalfScreenHeight)
		iGuiHeight := iHalfScreenHeight

	if !InStr(sGuiPosData, "center")
	{
		aGuiPosition := StrSplit(sGuiPosData, A_Space)
		if (aGuiPosition[1] > (A_ScreenWidth - iGuiWidth))
			ErrorMessage("sGuiPosData and/or iGuiWidth"
				, "Gui 的水平位置越過屏幕右邊緣.`n"
				"   更改 sGuiPosData 和/或 iGuiWidth 以將所有 Gui 保留在屏幕上。")
		if (aGuiPosition[2] > (A_ScreenHeight - (iGuiHeight + 44)))
			ErrorMessage("sGuiPosData and/or iGuiHeight"
				, "Gui 的垂直位置越過屏幕底部邊緣.`n"
				"   更改 sGuiPosData 和/或 iGuiHeight 以將所有 Gui 保留在屏幕上。")
	}

	; ==========================================================================
	; ------------------------------ GUI SETTINGS ------------------------------
	; ==========================================================================

	; #region

	MsgBoxGui := Gui(, Title)
	M_btnMap  := Map()

	if (Title == "Saving Clipboard")
	{
		MsgBoxGui.Opt("+AlwaysOnTop"),
		ShowNA      := " NA",
		fontSizeBtn := "s12"
	}
	else
	{
		ShowNA      := "",
		fontSizeBtn := "s18"
	}

	MsgBoxGui.MarginX   := 0,
	MsgBoxGui.MarginY   := 0,
	MsgBoxGui.BackColor := "302C29",
	
	MsgBoxGui.SetFont("c0xffffff", "Microsoft YaHei UI"),
	MsgBoxGui.Opt("+Resize +MinSize -DPIScale"),
	MsgBoxGui.OnEvent("Escape", lMyGui_Escape),
	MsgBoxGui.OnEvent("Close", lMyGui_EscapeClose),
	gEdit := MsgBoxGui.AddEdit("vEdit1 -E0200 R8 w400 " . hBackColor hTextColor, tTextToDisplay)
	gEdit.SetDarkMode()

	gEdit.OnEvent("Change", (gEditobj, *) => (SendMessage(0x115, 7, 0, gEditobj.Hwnd), Critical("Off"), Sleep(-1)))

	gEdit.SetFont("s" iFontSize, sFontName)
	gEdit.GetPos(, , &EditPosW, &EditPosH)
	btn_gap := 2 * (btnMap.Capacity - 1)
	btn_w   := (EditPosW / btnMap.Capacity - btn_gap)
	btn_h   := (EditPosH / 4)

	btnPosFunc(i := A_Index) => (
		(i = 1)
		? Trim(Format("-default xm w{1} h{2}", btn_w, btn_h))
		: Trim(Format("-default x+{1} wp hp", btn_gap))
	)

	For prebtnName, btnFunc in btnMap
	{
		btnName := IsDigit(SubStr(prebtnName, 1, 1)) ? SubStr(prebtnName, 2) : prebtnName
		option := btnPosFunc(A_Index)
		M_btnMap.Set btnName, MsgBoxGui.AddButton(option, btnName)
		M_btnMap[btnName].OnEvent("Click", btnFunc)
		M_btnMap[btnName].SetFont(fontSizeBtn " bold c0xffffff", "Microsoft YaHei UI")
		M_btnMap[btnName].SetDarkMode()
	}

	MsgBoxGui.SetDarkTitle()
	MsgBoxGui.GetClientPos(, , &SizeW, &SizeH)
	MsgBoxGui.OnEvent("Size", lMyGui_Size)
	MsgBoxGui.Show("AutoSize " ShowNA)
	MsgBoxGui.Move(,, 616, 440)
	Sleep(-1)
	Critical("Off")

	editHWND := gEdit.Hwnd

	if (sEditScrollToBtm = "yes")
		SendMessage(0x115, 7, 0, gEdit.Hwnd)
	else if (sEditScrollToBtm = "no")
		SendMessage(0x115, 6, 0, gEdit.Hwnd)

	switch
	{
	Case !IsSet(iGuiSleep), Title ~= "Clipboard":
		Return MsgBoxGui

	Case (iGuiSleep = 0):
		WinWaitClose(MsgBoxGui.Hwnd)
		return status

	default:
		if !WinWaitClose(MsgBoxGui.Hwnd, , iGuiSleep)
		{
			SetTimer((*) => MsgBoxGui.Destroy(), -100)
			return Status := "TimeOut"
		}
	}

	lMyGui_Size(GuiObj, MinMax, Width, Height, *)
	{

		SetControlDelay(0),
			SetWinDelay(-1)

		if (MinMax = -1)
			return

		gEdit.Move(, , Width, btn_y := (Height - btn_h))

		For btn, in M_btnMap
		{
			btn_w := (Width / M_btnMap.Capacity - btn_gap),
			btn_x := (A_Index - 1) * (btn_w + btn_gap),
			M_btnMap[btn].Move(btn_x, btn_y, btn_w)
		}
		; GuiObj.Show("AutoSize")
		Sleep(-1)
		return 1
	}

	ErrorMessage(ParamName, ErrorMsg)
	{
		MsgBox(
			"The function parameter **" ParamName "** is specified incorrectly:`n`n"
			"   " ErrorMsg ".`n`n"
			"Script terminating to allow correction."
		)
		Exit()
	}

	lMyGuiCopyBtn_Click(guiObj, *)
	{
		status := true
		
		A_Clipboard := ""
		A_Clipboard := guiObj.Gui["Edit1"].Value
		if ClipWait(1)
			TrayTip "已將文字複製到剪貼簿。", A_ScriptName
		
		Sleep(-1)
		return 1
	}

	lMyGui_Escape(guiObj, *)
	{
		If GuiObj is Gui
			guiObj.Hide()
		Else
			guiObj.Gui.Hide()

		Sleep(-1)
		return 0
	}

	lMyGui_EscapeClose(guiObj, *)
	{
		status := false
		If GuiObj is Gui
			guiObj.Destroy()
		Else
			guiObj.Gui.Destroy()

		Sleep(-1)
		return 0
	}

	zoomInOut(hk){

		Static guiHwnd   := ""
		Static editCtrl  := ""
		Static fontSize  := iFontSize
		Static myGuiEdit := ""

		CoordMode "Mouse"

		if !WinExist(guiHwnd) || !ControlGetVisible(editCtrl) {
			MouseGetPos(&OutputVarX, &OutputVarY, &guiHwnd, &editCtrl, 2)
			myGuiEdit := GuiCtrlFromHwnd(editCtrl)
		}
		
		myGuiEdit.SetFont(hk = "^WheelUp" ? "s" (++fontSize) : "s" (--fontSize))
	}
}




