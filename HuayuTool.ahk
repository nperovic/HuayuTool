/************************************************************************
 * @description A tool for Teaching Chinese 
 * @file HuayuTool.ahk
 * @author Nikola Perovic 陸汎宇
 * @date 2023/06/25
 * @version 1.0.0
 ***********************************************************************/

#Requires AutoHotKey v2
#SingleInstance
#Warn All, Off

A_ScriptName := "華語教學小工具"
A_IconTip    := "按下滑鼠中鍵(滾輪)以開啟功能表"

myMenu := Menu()
myMenu.Add "拼音輸入模式`t未開啟", pinyin
myMenu.Add "拼音輸入模式使用說明", (*) => MSGBOX
myMenu.Add "查詢所選文字的筆順", search
myMenu.Add
myMenu.Add "Google 搜尋所選內容", search
myMenu.Add "Google 翻譯所選內容", translate
myMenu.Add
myMenu.Add "簡體轉繁體", fanJian
myMenu.Add "繁體轉簡體", fanJian
myMenu.Add
myMenu.Add "修復工作列", reExplorer
myMenu.Add
myMenu.Add "前往華語中心首頁", (*) => Run('"https://c040e.wzu.edu.tw"')
myMenu.Add
myMenu.Add "贊助開發者", (*) => Run('"https://niko.soci.vip/donate"')

MsgBox "
(
	如何開啟功能表:
	1. 按下滑鼠中鍵(滾輪)。
	2. 按住左鍵的同時按下右鍵。

	---
	開發者: 陸汎宇 Nikola Perovic
	Email: 2@u4ni.tk 
)" , "華語教學小工具"


InstallMouseHook

#MaxThreadsPerHotkey 2
~MButton:: myMenu.Show()
~LButton & RButton:: myMenu.Show()

translate(item, *)
{
	savedClip   := ClipboardAll()
	A_Clipboard := ""
	Send "{LCtrl Down}{Ins}{LCtrl Up}"
	
	ClipWait 1

	result      := GTranslate(A_Clipboard)
	A_Clipboard := ""
	A_Clipboard := result
	
	ClipWait 1
	
	MsgBox result "`n`n(已將翻譯內容儲存至剪貼簿)"
}

fanJian(item, *)
{
	savedClip   := ClipboardAll()
	A_Clipboard := ""
	Send "{LCtrl Down}{Ins}{LCtrl Up}"
	ClipWait 1

	result      := GTranslate(A_Clipboard, item = "簡體轉繁體" ? "zh-TW" : "zh-CN")
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
	)" , "拼音輸入模式使用說明 | 華語教學小工具"
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

		} Until countdown < 0
		ToolTip()

		if !A_Clipboard {
			Send "{LCtrl Down}{Ins}{LCtrl Up}"
			goto here
		}
	}

	if item = "Google 搜尋所選內容"
		Run "https://www.google.com/search?q=" UrlUnescape(A_Clipboard)
	Else
		Run '"https://www.twpen.com/' UrlUnescape(A_Clipboard) '.html"'
}

reExplorer(item, *) {
	ProcessSetPriority "Realtime"
	Run A_ComSpec " /C taskkill /f /IM explorer.exe && start explorer.exe", "C:\", "Hide"
}

pinyin(item?, pos?, myMenu?)
{
	static toggle  := false
	static ih      := InputHook("v B")
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

UrlUnescape(url, flags := 0x00100000) => (DllCall('Shlwapi.dll\UrlUnescape', 'Str', url, 'Ptr', 0, 'UInt', 0, 'UInt', flags, 'UInt') ? '' : url)


GTranslate(Text, TargetLanguage := "zh-TW", SourceLanguage := "auto") {
	url := "https://translate.google.com/translate_a/single"

	_headers := Map(
		"User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
		"Referer", "https://translate.google.com/",
		"Accept", "application/json"
	)
	_params := Map(
		"client", "gtx",
		"sl", SourceLanguage,
		"tl", TargetLanguage,
		"dt", "t",
		"q", Text
	)

	headers := ""
	params  := ""
	
	for key, value in _headers
		headers .= key "=" value "&"

	for key, value in _params 
		params .= (A_Index != _params.Count) ? key "=" value "&" : key "=" value

	response := ComObject("WinHttp.WinHttpRequest.5.1")
	response.Open("GET", url . "?" . headers . params, false)
	response.Send()

	if response.Status = 200 {
		Translation       := response.ResponseText
		TranslationObject := Jxon_Load(&Translation)
		data              := Jxon_Dump(TranslationObject)

		VarSetStrCapacity(&output, strLen(Text) * 10)

		for k, v in TranslationObject[1]
			output .= TranslationObject[1][k][1]
		return output
	} else 
		MsgBox("Request failed with status code " response.Status)
}


Jxon_Load(&src, args*) {
	key := "", is_key := false
	stack := [ tree := [] ]
	next := '"{[01234567890-tfn'
	pos := 0
	
	while ( (ch := SubStr(src, ++pos, 1)) != "" ) {
		if InStr(" `t`n`r", ch)
			continue
		if !InStr(next, ch, true) {
			testArr := StrSplit(SubStr(src, 1, pos), "`n")
			
			ln := testArr.Length
			col := pos - InStr(src, "`n",, -(StrLen(src)-pos+1))

			msg := Format("{}: line {} col {} (char {})"
			,   (next == "")      ? ["Extra data", ch := SubStr(src, pos)][1]
			  : (next == "'")     ? "Unterminated string starting at"
			  : (next == "\")     ? "Invalid \escape"
			  : (next == ":")     ? "Expecting ':' delimiter"
			  : (next == '"')     ? "Expecting object key enclosed in double quotes"
			  : (next == '"}')    ? "Expecting object key enclosed in double quotes or object closing '}'"
			  : (next == ",}")    ? "Expecting ',' delimiter or object closing '}'"
			  : (next == ",]")    ? "Expecting ',' delimiter or array closing ']'"
			  : [ "Expecting JSON value(string, number, [true, false, null], object or array)"
			    , ch := SubStr(src, pos, (SubStr(src, pos)~="[\]\},\s]|$")-1) ][1]
			, ln, col, pos)

			throw Error(msg, -1, ch)
		}
		
		obj := stack[1]
        is_array := (obj is Array)
		
		if i := InStr("{[", ch) { ; start new object / map?
			val := (i = 1) ? Map() : Array()	; ahk v2
			
			is_array ? obj.Push(val) : obj[key] := val
			stack.InsertAt(1,val)
			
			next := '"' ((is_key := (ch == "{")) ? "}" : "{[]0123456789-tfn")
		} else if InStr("}]", ch) {
			stack.RemoveAt(1)
            next := (stack[1]==tree) ? "" : (stack[1] is Array) ? ",]" : ",}"
		} else if InStr(",:", ch) {
			is_key := (!is_array && ch == ",")
			next := is_key ? '"' : '"{[0123456789-tfn'
		} else { ; string | number | true | false | null
			if (ch == '"') { ; string
				i := pos
				while i := InStr(src, '"',, i+1)
                {
					val := StrReplace(SubStr(src, pos+1, i-pos-1), "\\", "\u005C")
					if (SubStr(val, -1) != "\")
						break
				}
				if !i ? (pos--, next := "'") : 0
					continue

				pos := i ; update pos

				val := StrReplace(val, "\/", "/")
				val := StrReplace(val, '\"', '"')
				, val := StrReplace(val, "\b", "`b")
				, val := StrReplace(val, "\f", "`f")
				, val := StrReplace(val, "\n", "`n")
				, val := StrReplace(val, "\r", "`r")
				, val := StrReplace(val, "\t", "`t")

				i := 0
				while i := InStr(val, "\",, i+1) {
					if (SubStr(val, i+1, 1) != "u") ? (pos -= StrLen(SubStr(val, i)), next := "\") : 0
						continue 2

					xxxx := Abs("0x" . SubStr(val, i+2, 4)) ; \uXXXX - JSON unicode escape sequence
					if (xxxx < 0x100)
						val := SubStr(val, 1, i-1) . Chr(xxxx) . SubStr(val, i+6)
				}
				
				if is_key {
					key := val, next := ":"
					continue
				}
			} else { ; number | true | false | null
				val := SubStr(src, pos, i := RegExMatch(src, "[\]\},\s]|$",, pos)-pos)
				
                if IsInteger(val)
                    val += 0
                else if IsFloat(val)
                    val += 0
                else if (val == "true" || val == "false")
                    val := (val == "true")
                else if (val == "null")
                    val := ""
                else if is_key {
                    pos--, next := "#"
                    continue
                }
				
				pos += i-1
			}
			
			(is_array ? obj.Push(val) : (obj[key] := val))
			next := ((obj == tree) ? "" : is_array ? ",]" : ",}")
		}
	}
	
	return tree[1]
}

Jxon_Dump(obj, indent:="", lvl:=1) {
	if IsObject(obj) {
        If !(obj is Array || obj is Map || obj is String || obj is Number)
			throw Error("Object type not supported.", -1, Format("<Object at 0x{:p}>", ObjPtr(obj)))
		
		if IsInteger(indent)
		{
			if (indent < 0)
				throw Error("Indent parameter must be a postive integer.", -1, indent)
			spaces := indent, indent := ""
			
			Loop spaces ; ===> changed
				indent .= " "
		}
		indt := ""
		
		Loop indent ? lvl : 0
			indt .= indent
        
        is_array := (obj is Array)
        
		lvl += 1, out := "" ; Make #Warn happy
		for k, v in obj {
			if IsObject(k) || (k == "")
				throw Error("Invalid object key.", -1, k ? Format("<Object at 0x{:p}>", ObjPtr(obj)) : "<blank>")
			
			if !is_array ;// key ; ObjGetCapacity([k], 1)
				out .= (ObjGetCapacity([k]) ? Jxon_Dump(k) : escape_str(k)) (indent ? ": " : ":") ; token + padding
			
			out .= Jxon_Dump(v, indent, lvl) ; value
				.  ( indent ? ",`n" . indt : "," ) ; token + indent
		}

		if (out != "") {
			out := Trim(out, ",`n" . indent)
			if (indent != "")
				out := "`n" . indt . out . "`n" . SubStr(indt, StrLen(indent)+1)
		}
		
		return is_array ? "[" . out . "]" : "{" . out . "}"
	
    } Else If (obj is Number)
        return obj
    
    Else ; String
        return escape_str(obj)
	
    escape_str(obj) {
        obj := StrReplace(obj,"\","\\")
        obj := StrReplace(obj,"`t","\t")
        obj := StrReplace(obj,"`r","\r")
        obj := StrReplace(obj,"`n","\n")
        obj := StrReplace(obj,"`b","\b")
        obj := StrReplace(obj,"`f","\f")
        obj := StrReplace(obj,"/","\/")
        obj := StrReplace(obj,'"','\"')
        
        return '"' obj '"'
    }
}

