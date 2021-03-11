; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 10.03.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
return
; CONTROLS                                                                                                                                                                                                                                        	(32)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; GetClassName                                	Control_GetClassNN                     	GetClassNN                                  	GetFocusedControl                        	GetFocusedControlHwnd
; GetFocusedControlClassNN            	GetChildHWND                            	GetControls                                  	GetButtonType                                	Controls
; ControlFind                                     	ControlGet                                    	ControlGetText                              	ControlGetFocus                            	GuiControlGet
; ControlGetFont									WinSaveCheckboxes                        	Toolbar_GetRect
; ControlGetTabs                               	TabCtrl_GetCurSel                          	TabCtrl_GetItemText
; VerifiedClick                                    	VerifiedCheck                                	VerifiedChoose                              	VerifiedSetFocus                            	VerifiedSetText
; UpSizeControl
; LVM_GetNext                                   	LV_MouseGetCellPos							LV_Select                                        	LV_SortArrow									LV_FindRow
; CaretPos
;_________________________________________________________________________________________________________________________________________________________
GetClassName(hwnd) {                                                                                 	;-- returns HWND's class name without its instance number, e.g. "Edit" or "SysListView32"
		;https://autohotkey.com/board/topic/45515-remap-hjkl-to-act-like-left-up-down-right-arrow-keys/#entry283368
	VarSetCapacity( buff, 256, 0 )
	DllCall("GetClassName", "uint", hwnd, "str", buff, "int", 255 )
return buff
}

Control_GetClassNN(hWnd, hCtrl) {
	; SKAN: www.autohotkey.com/forum/viewtopic.php?t=49471
 WinGet, CH, ControlListHwnd	, % "ahk_id " hWnd
 WinGet, CN, ControlList       	, % "ahk_id " hWnd
 LF:= "`n",  CH:= LF CH LF, CN:= LF CN LF,  S:= SubStr(CH,1,InStr(CH,LF hCtrl LF))
 S:= StrReplace(S,"`n","`n", RplCount)
 LP:= InStr(CN, "`n",,, RplCount) - 1
 Return SubStr(CN,LP+2,InStr(CN,LF,0,LP+2)-LP-2)
}

GetClassNN(Chwnd, Whwnd) {
	global _GetClassNN := {}
	_GetClassNN.Hwnd := Chwnd
	Detect := A_DetectHiddenWindows
	WinGetClass, Class, ahk_id %Chwnd%
	_GetClassNN.Class := Class
	DetectHiddenWindows, On
	EnumAddress := RegisterCallback("GetClassNN_EnumChildProc")
	DllCall("EnumChildWindows", "uint",Whwnd, "uint",EnumAddress)
	DetectHiddenWindows, %Detect%
	return, _GetClassNN.ClassNN, _GetClassNN:=""
}
GetClassNN_EnumChildProc(hwnd, lparam) {
	static N
	global _GetClassNN
	WinGetClass, Class, ahk_id %hwnd%
	if _GetClassNN.Class == Class
		N++
	return _GetClassNN.Hwnd==hwnd? (0, _GetClassNN.ClassNN:=_GetClassNN.Class N, N:=0):1
}

GetFocusedControl()  {                                                                                  	;-- retrieves the ahk_id (HWND) of the active window's focused control.

   ; This script requires Windows 98+ or NT 4.0 SP3+.
   guiThreadInfoSize := 8 + 6*A_PtrSize + 16
   VarSetCapacity(guiThreadInfo, guiThreadInfoSize, 0)
   NumPut(GuiThreadInfoSize, GuiThreadInfo, 0)
   ; DllCall("RtlFillMemory" , "PTR", &guiThreadInfo, "UInt", 1 , "UChar", guiThreadInfoSize)   ; Below 0xFF, one call only is needed
   if (DllCall("GetGUIThreadInfo" , "UInt", 0, "PTR", &guiThreadInfo) = 0) {   ; Foreground thread
		ErrorLevel := A_LastError   ; Failure
		Return 0
   }

; *(addr + 12) + (*(addr + 13) << 8) +  (*(addr + 14) << 16) + (*(addr + 15) << 24)
Return Format("0x{:X}", NumGet(guiThreadInfo, 8+A_PtrSize, "Ptr"))
}

GetFocusedControlHwnd(hwnd:="A") {
	ControlGetFocus, FocusedControl	 , % (hwnd = "A") ? "A" : "ahk_id " hwnd
	ControlGet    	 , FocusedControlId, Hwnd,, % FocusedControl, % (hwnd = "A") ? "A" : "ahk_id " hwnd
return FocusedControlId
}

GetFocusedControlClassNN(hwnd:="A") {
	ControlGetFocus, FocusedControl	 , % (hwnd = "A") ? "A" : "ahk_id " hwnd
	ControlGet    	 , FocusedControlId, Hwnd,, % FocusedControl, % "ahk_id " hwnd
return Control_GetClassNN(hwnd, FocusedControlId)
}

GetChildHWND(ParentWindowID, ChildClassNN) {

	; Link: https://autohotkey.com/board/topic/3369-unique-id-number-of-a-child-control/
	; Returns a blank value if parent or child doesn't exist.
	; Otherwise, the HWND is returned.

	WinGetPos, ParentX, ParentY,,, % "ahk_id " ParentWindowID
	if ParentX =
		return  ; Parent window not found (possibly due to DetectHiddenWindows).

	ControlGetPos, ChildX, ChildY,,, % ChildClassNN, % "ahk_id " ParentWindowID
	if ChildX =
		return  ; Child window not found.

return DllCall("WindowFromPoint", "int", ChildX + ParentX, "int", ChildY + ParentY)
}

GetControls(hwnd, class_filter:="", type_filter:="", info_filter:="") {                  	  	;-- returns an array with ClassNN, ButtonTyp, Position.....

	;class_filter - comma separated list of classes you don't want to store
	;type_filter - comma separated list of classes you don't want to store
	;info_filter - comma separated list of classes you !want! to store

	If StrLen(info_filter) = 0
		info_filter:="hwnd,Pos,Enabled,Visible,Style,ExStyle"

	controls:=[], Control_Style:= "Style", Control_IsEnabled:= "Enabled", Control_IsVisible:= "Visible", Control_ExStyle:= "ExStyle", Control_Pos:= "Pos", Control_Handle:= "hwnd"
	WinGet, classnn  	, ControlList        	,ahk_id %hwnd%
	WinGet, controlId	, controllisthwnd	,ahk_id %hwnd%

	loop, parse, classnn,`n
	{
		controls[A_Index]:=[]
		controls[A_Index]["classNN"]:=A_LoopField
	}

	loop, parse, controlId,`n
	{
			RegExMatch(controls[A_Index]["classNN"], "[a-zA-Z]+", class)
			If class in %class_filter%
                    	continue

			If class in Button
			{
					bTyp:= GetButtonType(A_LoopField)
					if bTyp in %type_filter%
						continue

					If bTyp in Radio,Checkbox
						controls[A_Index]["checked"]	:= ControlGet("Checked", "", "", "ahk_id " A_LoopField)
					else
						controls[A_Index]["text"]:= ControlGetText("", "ahk_id " A_LoopField)

					controls[A_Index]["type"]:= bTyp
			}

			If class in Edit,RichEdit
			{
                   	controls[A_Index]["text"]:= ControlGetText("", "ahk_id " A_LoopField)
                   	controls[A_Index]["linecount"]:= ControlGet("LineCount", "", "", "ahk_id " A_LoopField)
			}

			If Control_Handle  	in %info_filter%
				controls[A_Index]["hwnd"]   	:= A_Loopfield

			If Control_IsEnabled 	in %info_filter%
				controls[A_Index]["Enabled"]	:= ControlGet("Enabled", "", "", "ahk_id " A_LoopField)

			If Control_IsVisible 	in %info_filter%
				controls[A_Index]["Visible"]  	:= ControlGet("visible", "", "", "ahk_id " A_LoopField)

			If Control_Style      	in %info_filter%
				controls[A_Index]["Style"]    	:= ControlGet("Style", "", "", "ahk_id " A_LoopField)

			If Control_ExStyle   	in %info_filter%
				controls[A_Index]["Exstyle"]  	:= ControlGet("ExStyle", "", "", "ahk_id " A_LoopField)

			If Control_Pos        	in %info_filter%
			{
                    ControlGetPos, cx, cy, cw, ch,, ahk_id %A_Loopfield%
                    controls[A_Index]["Pos"]        	:= cx "," cy "," cw "," ch
			}
	}

return controls
}

GetButtonType(hwndButton) {                                                                         	;-- ermittelt welcher Art ein Button ist, liest dazu den Buttonstyle aus
	;Link: https://autohotkey.com/board/topic/101341-getting-type-of-control/
  static types := [ "Button"            	;BS_PUSHBUTTON
                     	, "Button"            	;BS_DEFPUSHBUTTON
                     	, "Checkbox"      	;BS_CHECKBOX
                     	, "Checkbox"      	;BS_AUTOCHECKBOX
                     	, "Radio"             	;BS_RADIOBUTTON
                     	, "Checkbox"      	;BS_3STATE
                     	, "Checkbox"      	;BS_AUTO3STATE
                     	, "Groupbox"      	;BS_GROUPBOX
                     	, "NotUsed"       	;BS_USERBUTTON
                     	, "Radio"             	;BS_AUTORADIOBUTTON
                     	, "Button"            	;BS_PUSHBOX
                     	, "AppSpecific"   	;BS_OWNERDRAW
                     	, "SplitButton"       	;BS_SPLITBUTTON    (vista+)
                     	, "SplitButton"       	;BS_DEFSPLITBUTTON (vista+)
                     	, "CommandLink"	;BS_COMMANDLINK    (vista+)
                     	, "CommandLink"]	;BS_DEFCOMMANDLINK (vista+)

  WinGet, btnStyle, Style, % "ahk_id " hwndButton
 return types[1+(btnStyle & 0xF)]
}

Controls(Control,cmd,WinTitle,HiddenText=1,HiddenWin=1,MatchModeSpeed="slow") {                	;-- Universalfunktion für Steuerelemente

	; ********	    ********		Funktion wächst und gedeiht, gegossen am 16.02.2021
	;***	     *    ***	    ***	dependencies: 	Function: ClientToScreen()
	;***             ***      ***                        	Function: KeyValueObjectFromLists()
	;***             ***      ***                         	Function: VerifiedSetText() - [ ControlGetText() ]
	;***        *    ***      ***                        	Function: VerifiedClick()
	;  *******       ********
	;
	; BESCHREIBUNG: 	diese Funktion nutzt automatisch den richtigen Autohotkey Befehl um Steuerelemente anzusprechen.
	;                              	Es ist nicht mehr notwendig sich Befehlsketten zu merken.
	;                              	Z.b. ersetzt Controls("Edit1", "GetText", WinTitle) die Befehlskette ControlGet, var,, Edit1, % WinTitle
	;                              	oder Controls("Listbox1", "GetText", WinTitle) die Befehlskette ControlGet, var, List, Edit1, % WinTitle
	;                            	Eine Hexadezimalzahl als WinTitle wird automatisch als das Handle des Fensters interpretiert.
	;                              	Ausserdem ließt die Funktion beim ersten Aufruf alle vorhandenen Steuerelement ClassNN und die
	;                            	zugehörigen Handles in ein Objekt.

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Einstellungen für Hidden und TitleMode
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		static knWinTitle, knWinText, Ctrl := Object()

		HiddenTextStatus        	:= A_DetectHiddenText
		HiddenWinStatus        	:= A_DetectHiddenWindows
		MatchModeSpeedStatus	:= A_TitleMatchModeSpeed
		CoordModeWinStatus	:= A_CoordModeMouse
		DetectHiddenText      	, % HiddenText=true ? "On":"Off"
		DetectHiddenWindows	, % HiddenWin=true ? "On":"Off"
		SetTitleMatchMode    	, % MatchModeSpeed
		CoordMode              	, Mouse	, Screen
		CoordMode              	, Pixel   	, Screen
		sleep, 10                    	; CoordMode needs a pause to update - https://www.autohotkey.com/boards/viewtopic.php?f=14&t=38467
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Ermitteln aller Steuerelementklassen und Handle des Fensters
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		RegExMatch(Control, "[a-zA-Z#]+", class)

		WinTitle := Trim(WinTitle)
		If !(knWinTitle = WinTitle) 		{

			knWinTitle	 := WinTitle
			WinTitle	:= RegExMatch(WinTitle, "^0x[\w]+$")	? ("ahk_id " WinTitle)	: (WinTitle)
			RegExMatch(WinTitle, "^\d+$", digits)
			WinTitle	:= StrLen(WinTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WinTitle)

			WinGet, cClasses	, ControlList			, % WinTitle                                                   	; use this for example: "ahk_id " hWin
			WinGet, cHwnds	, ControlListHwnd	, % WinTitle
			Ctrl := KeyValueObjectFromLists(cClasses, cHwnds, "`n", "", "", "", "")                        	; ergibt ein Object mit ClassNN als key und handle als value

		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Befehlsbereich -
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		cmd := Trim(cmd)
	    if        RegExMatch(cmd	, "i)^\s*(Hwnd|ID)"              	)         	{   	; returns the handle for a ClassNN

				; empty Control parameter returns the array with class
					If (StrLen(Control) = 0)
						return Ctrl
				; this returns the handle if Control fits to a key
                    else If Ctrl.HasKey(Control)
						return Ctrl[Control]
				; this returns the handle if there was no match before
					else
					   For ControlClass, ControlHwnd in Ctrl
							If InStr(ControlClass, control)
								return ControlHwnd

		}
		else if RegExMatch(cmd	, "i)^\s*Click"                      	)        	{   	; like ControlClick, you need the exactly classNN name

		; but you can specify the method to use (MouseClick, ControlClick)

				If 	InStr(control, "ToolbarWindow") 	{

					ControlGetPos, cx, cy, cw, ch,, % "ahk_id " Ctrl[Control]
					ClientToScreen(Ctrl[Control], cx, xy)
					If InStr(cmd, "Space")                 	{	; Controls("ToolbarWindow324", "Click use Spacebar", "ahk_exe notepad.exe")
						ControlFocus,, % "ahk_id " Ctrl[Control]
						sleep, 400
						ControlSend,, {Space}, % "ahk_id " Ctrl[Control]
					}
					else if InStr(cmd, "ControlClick"	)	{	; Controls("ToolbarWindow324", "Click use Controlclick left", "ahk_exe notepad.exe")
						RegExMatch(cmd, "i)(?<=ControlClick\s)\w+", Button)
						if Button in Left,Middle,Right
							ControlClick,, % "ahk_id " Ctrl[Control],, % Button, 1, NA
					}
					else if InStr(cmd, "MouseClick"	)	{	;Controls("ToolbarWindow324", "Click use MouseClick", "ahk_exe notepad.exe")
						BlockInput, On
						MouseGetPos, mx, my
						ControlGetPos, cx, cy, cw, ch,, % "ahk_id " Ctrl[Control]
						ClientToScreen(Ctrl[Control], cx, xy)
						RegExMatch(cmd, "i)(?<=MouseClick\s)\w+", Button)
						if Button in Left,Middle,Right
							MouseClick, % Button, % (cx + cw - 50), % (cy + ch//2), 1, 0
						MouseMove, % mx, % my, 0
						BlockInput, Off
					}

				}
				else if InStr(control, "Button")                               	{
					; 3 verschiedene Wege einen Buttonklick auszulösen
						ControlClick,, % "ahk_id " Ctrl[Control],, Left,, NA
						If ErrorLevel	{
							ControlClick,, % "ahk_id " Ctrl[Control],, Left
							If ErrorLevel	{
								PostMessage, 0x201, 1, 0,, % "ahk_id " Ctrl[Control] ;0x201 - WM_Click
					    		If ErrorLevel 	{
									BlockInput, On
									MouseGetPos, mx, my
									ControlGetPos,,, cw, ch,, % "ahk_id " Ctrl[Control]
									ClientToScreen(Ctrl[Control], cx, cy)
									MouseClick, Left, % (cx + cw//2) , % (cy + ch//2), 1, 0
									MouseMove, % mx, % my, 0
									BlockInput, Off
								}
					    	}
						}
				}

		}
		else if RegExMatch(cmd	, "i)^\s*ControlClick"           	)         	{   	; clicks a control by its text or classNN and returns the ErrorLevel

			; möglicher Syntax z.B. Controls("", "ControlClick, Speichern, Button", "YourWinTitle")
				searchText		:= Trim( StrSplit(cmd, ",").2 )
				searchClass	:= Trim( StrSplit(cmd, ",").3 )

				If (StrLen(searchText) > 0)  {
					For ControlClass, ControlHwnd in Ctrl	{
						ControlGetText, ControlText,, % "ahk_id " ControlHwnd

						If ( InStr(ControlClass, searchClass) && InStr(ControlText, searchText) ) {
							SciteOutPut(ControlClass ", " searchClass " - " ControlText ", " searchText)
							return VerifiedClick(ControlClass, WinTitle)
						}
					}
				}
				else
					return VerifiedClick(searchClass, WinTitle)		;searchClass must be the exact ClassNN in this case

		}
		else if RegExMatch(cmd	, "i)^\s*ControlFind"            	)        	{   	; finds a control by its text and returns it's ControlClassNN

			; möglicher Syntax z.B. Controls("", "ControlFind, Speichern, Button, return hwnd", winhwnd)
				searchText 	:= Trim(StrSplit(cmd, ",").2)
				searchClass	:= Trim(StrSplit(cmd, ",").3)
				returnOpt  	:= Trim(StrSplit(cmd, ",").4)

				For ControlClass, ControlHwnd in Ctrl {
					ControlGetText, ControlText,, % "ahk_id " ControlHwnd
					ControlText := RegExReplace(ControlText, "(\w)\&(\w)", "$1$2")
					If InStr(ControlClass, searchClass) && InStr(ControlText, searchText) {
						If RegExMatch(returnOpt, "i)^\s*return\s*both|all")
							return {"class":ControlClass, "hwnd":GetHex(ControlHwnd)}
						else if RegExMatch(returnOpt, "i)^\s*return\s*hwnd|id|handle")
							return GetHex(ControlHwnd)
						else
							return ControlClass
					}
				}

				return

		}
		else if RegExMatch(cmd	, "i)^\s*ControlPos"           	)        	{   	; returns the controls position inside window

			ControlGetPos, x,y,w,h, % Control, % WinTitle, % WinText
			return {"X":x, "Y":y, "W":w, "H":h}

		}
		else if RegExMatch(cmd	, "i)^\s*GetControls"         	)         	{   	; try's to return all subcontrol hwnds (no treeversal)

				Childs := Array(), found := false, HiddenControls := true

				If !RegExMatch(cmd, ".*\+Hidden") {
					GCHiddenTextStatus      := A_DetectHiddenText
					GCHiddenWinStatus      := A_DetectHiddenWindows
					DetectHiddenText      	, Off
					DetectHiddenWindows	, Off
					HiddenControls := false
				}

				If (StrLen(Control) > 0) {

					For ControlClass, ControlHwnd in Ctrl
						If InStr(ControlClass, Control) || InStr(ControlGetText(ControlHwnd), Control) {
							WinTitle := "ahk_id " ControlHwnd, found := true
							break
					}

					If !found
						return 0

				}

				WinGet, childList, ControlListHwnd, % WinTitle

				Loop, Parse, childList, `n
				{
					; auch verborgene Steuerelemente einsammeln
						If HiddenControls
							Childs.Push(A_LoopField)
						else if IsWindowVisible(A_LoopField) ; oder nur sichtbare
							Childs.Push(A_LoopField)
				}

				DetectHiddenText      	, % GCHiddenText=true ? "On":"Off"
				DetectHiddenWindows	, % GCHiddenWin=true ? "On":"Off"

				return Childs

		}
		else if RegExMatch(cmd	, "i)^\s*GetFocus"             	)          	{   	; finds the focused control and returns it's ControlClassNN

			For ControlClass, ControlHwnd in Ctrl		{
				If DllCall("IsWindow", "Ptr", ControlHwnd)	{
						ControlGetFocus, cFocus, % "ahk_id " ControlHwnd
						If (StrLen(cFocus) > 0)
								return cFocus
				}
			}
			return cFocus

		}
		else if RegExMatch(cmd	, "i)^\s*GetText"                   	)         	{   	; get text from any control

			If (StrLen(Ctrl[Control]) = 0)
				return ""

			 If class in Edit,ToolbarWindow,Static
				ControlGetText	, result,	 		, % "ahk_id " Ctrl[Control]
			else if class in ComboBox,ListBox,Listview,DropDownList
				ControlGet    	, result, List ,,, % "ahk_id " Ctrl[Control]
			else if RegExMatch(Control, "i)WindowsForms.*\.(STATIC|BUTTON)")
				ControlGetText	, result,        	, % "ahk_id " Ctrl[Control]

			return result
		}
		else if RegExMatch(cmd	, "i)^\s*(ControlSend|Send)"	)        	{   	; wrapper

			if class in Edit,RichEdit
			{
				RegExMatch(cmd, "i)(ControlSend|Send)[\s,]+(?<eys>.*)", k)
				ControlSend,, % keys, % "ahk_id " Ctrl[Control]
			}

		}
		else if RegExMatch(cmd	, "i)^\s*(ControlSendRaw|SendRaw)"){   	; wrapper

			If class in Edit,RichEdit
			{
				RegExMatch(cmd, "i)(?<=ControlSend\s).*", keys)
				ControlSendRaw,, % keys, % "ahk_id " Ctrl[Control]
			}

		}
		else if RegExMatch(cmd	, "i)^\s*SetFocus"                	)          	{   	;

			if class in Edit,RichEdit,ComboBox,ListBox,ListView,DropDownList,ToolbarWindow
			{
				ControlFocus,, % "ahk_id " Ctrl[Control]
				return ErrorLevel
			}

		}
		else if RegExMatch(cmd	, "i)^\s*SetText"                   	)          	{   	;

			If class in Edit,RichEdit
			{
				RegExMatch(cmd, "i)SetText[\s,]+(?<Text>.*)", New)
				If !NewText
					MsgBox, % "not matching syntax [" NewText "]  for SetText command. " A_LineFile
				return VerifiedSetText("", NewText, "ahk_id " Ctrl[Control], 100)
			}

		}
		else if RegExMatch(cmd	, "i)^\s*Reset"                     	)       	{   	; empty Ctrls Object
           	Ctrl			:= Object()
           	result		:= 1
           	knWinTitle	:= ""
           	return 1
		}
		else if RegExMatch(cmd	, "i)^\s*ControlCount"          	)         	{  	; returns the count of all found controls
			return Ctrl.Count()
		}
		else if RegExMatch(cmd	, "i)^\s*GetActiveMDIChild"	)         	{  	; returns the active MDI child

			RegExMatch(cmd, "i)[\s,]return?[\s,](?<md>[\w]+)", c)
			For ControlClass, ControlHwnd in Ctrl {
				If (ControlClass = "MDIClient1")
					break
			}
			hMDIClient := ControlHwnd
			SendMessage, 0x0229,,,, % "ahk_id " hMDIClient
			hMDIChild := GetHex(ErrorLevel)
			StringCaseSense, Off
			Switch cmd	{

				case "id", "hwnd", "handle":
					return hMDIChild

				case "classnn":
					For ControlClass, ControlHwnd in Ctrl
						If (ControlHwnd = hMDIChild)
							return ControlClass

				default:
					return hMDIChild

			}

		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Zurückstellen des Hidden- und Titlemodus
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		DetectHiddenText			, % HiddenTextStatus
		DetectHiddenWindows	, % HiddenWinStatus
		SetTitleMatchMode		, % MatchModeSpeedStatus

	;}

return result
}

ControlFind(Control, cmd, WinTitle) {				                                                	;-- Controls ist die bessere Version jetzt

		static knWinTitle, knWinText
		static Ctrl	:= Object()

		If RegExMatch(WinTitle, "^0x")
				WinTitle := "ahk_id " WinTitle

		If InStr(cmd, "reset")				{
				VarSetCapacity(Ctrl, 0)
				Ctrl := Object()
				return 1
		}

		If !InStr(knWinTitle, WinTitle)			{
			knWinTitle:= WinTitle
			If WinText=""
              	knWinText:= 0
			else
               	knWinText:= WinText

			WinGet, cClasses	, ControlList			, % WinTitle, ;% WinText                            	; use this for example: "ahk_id " hWin
			WinGet, cHwnds	, ControlListHwnd	, % WinTitle, ;% WinText
			Ctrl	:= KeyValueObjectFromLists(cClasses, cHwnds, "`n", "", "", "", "")					; ergibt ein Object mit ClassNN und dem handle
		}

		If InStr(cmd, "ID")
			return Ctrl[Control]
		else if InStr(cmd, "GetText")	{
			RegExMatch(control, "[a-zA-Z]+", class)
			If class in Edit,ToolbarWindow
				ControlGetText, result,, % "ahk_id " Ctrl[Control]
			else If class in ComboBox,ListBox
				ControlGet, result, Choice,,, % "ahk_id " Ctrl[Control]
		}
		else If InStr(cmd, "GetList") 	{
			If InStr(Control, "ComboBox") || InStr(Control, "ListBox") || InStr(Control, "Listview") InStr(Control, "DropDownList")
				ControlGet, result, List,,, % "ahk_id " Ctrl[Control]  ;Choice
		}

return result
}

ControlGet(Cmd,Value="",Control="",WTitle="",WTxt="",ExTitle="",ExText="") {  	;-- ControlGet als Funktion
	ControlGet, v, % Cmd, % Value, % Control, % WTitle, % WTxt, % ExTitle, % ExText
Return v
}

ControlGetText(Control="", WTitle="", WText="", ExTitle="", ExText="") {          	;-- ControlGetText als Funktion
	ControlGetText, v, % Control, % WTitle, % WText, % ExTitle, % ExText
Return v
}

ControlGetFocus(hwnd) {                                                                                	;-- gibt das Handle des fokussierten Steuerelementes zurück
	ControlGetFocus, FocusedControl, % "ahk_id " hwnd
	ControlGet, FocusedControlId, Hwnd,, %FocusedControl%, % "ahk_id " hwnd
return FocusedControlId
}

GuiControlGet(guiname, cmd, vcontrol) {                                                        	;-- GuiControlGet wrapper

	GuiControlGet, cp, % guiname ": " cmd, % vcontrol
	If (cmd = "Pos")
		return {"X":cpX, "Y":cpY, "W":cpW, "H":cpH}

return cp
}

ControlGetFont(hWnd,ByRef Name,ByRef Size,ByRef Style,IsGDIFontSize=0) {  	;-- Fontname, Größe, Stil eines Controls ermitteln

	; www.autohotkey.com/forum/viewtopic.php?p=465438#465438

    SendMessage 0x31, 0, 0, , ahk_id %hWnd% ; WM_GETFONT
    If (ErrorLevel == "FAIL")
        Return

    hFont := Errorlevel
    VarSetCapacity(LOGFONT, LOGFONTSize := 60 * (A_IsUnicode ? 2 : 1 ))
    DllCall("GetObject", "Ptr", hFont, "Int", LOGFONTSize, "Ptr", &LOGFONT)

    Name := DllCall("MulDiv", "Int", &LOGFONT + 28, "Int", 1, "Int", 1, "Str")

    Style := Trim((Weight := NumGet(LOGFONT, 16, "Int")) == 700 ? "Bold" : (Weight == 400) ? "" : " w" . Weight
    . (NumGet(LOGFONT, 20, "UChar") ? " Italic" : "")
    . (NumGet(LOGFONT, 21, "UChar") ? " Underline" : "")
    . (NumGet(LOGFONT, 22, "UChar") ? " Strikeout" : ""))

    Size := IsGDIFontSize ? -NumGet(LOGFONT, 0, "Int") : Round((-NumGet(LOGFONT, 0, "Int") * 72) / A_ScreenDPI)
}

WinSaveCheckboxes(hWin) {                                                                            	;-- speichert den Status aller Checkbox-Steuerelemente

	idx				:=0
	oControls1	:= Object()
	oControls2	:= Object()

	WinGet, cClasses	, ControlList			, % "ahk_id " hWin
	WinGet, cHwnds	, ControlListHwnd	, % "ahk_id " hWin

	oControls1:= KeyValueObjectFromLists(cClasses, cHwnds, "`n", "Button", "[A-Za-z]+", "", "")

	For key, val in oControls1
	{
			If InStr(GetButtonType(val), "Checkbox") {
                    status:= ControlGet("checked",,, "ahk_id " . val)
                    oControls2[(key)]:= status
                    idx++
			}
	}

	If !idx
		return 0

return oControls2
}

ToolbarGetRect(hCtrl, Pos="", pQ="") {                                                           	;-- ermittelt die Größe eines ToolbarControls
	ListLines, Off
	/*
 Function:  GetRect
 			Get button rectangle

 Parameters:
 			pPos		- Button position. Leave blank to get dimensions of the toolbar control itself.
 			pQ			- Query parameter: set x,y,w,h to return appropriate value, or leave blank to return all in single line.

 Returns:
 			String with 4 values separated by space or requested information
 */

	static TB_GETITEMRECT=0x41D

	if pPos !=
		ifLessOrEqual, Pos, 0, return "Err: Invalid button position"

	VarSetCapacity(RECT, 16)
    SendMessage, TB_GETITEMRECT, Pos-1,&RECT, ,ahk_id %hCtrl%
	IfEqual, ErrorLevel, 0, return A_ThisFunc "> Can't get rect"

	if Pos =
		DllCall("GetClientRect", "uint", hCtrl, "uint", &RECT)

	x := NumGet(RECT, 0), y := NumGet(RECT, 4), r := NumGet(RECT, 8), b := NumGet(RECT, 12)
	return (pQ = "x") ? x : (pQ = "y") ? y : (pQ = "w") ? r-x : (pQ = "h") ? b-y : x " " y " " r-x " " b-y
}

ControlGetTabs(hTab) {                                                                                 	;-- ermittelt die Texte aller TabControls zu diesem hwnd

	; https://autohotkey.com/board/topic/70727-ahk-l-controlgettabs/

    Static MAX_TEXT_LENGTH	:= 260
		   , MAX_TEXT_SIZE     	:= MAX_TEXT_LENGTH * (A_IsUnicode ? 2 : 1)

    WinGet PID, PID, ahk_id %hTab%

    ; Open the process for read/write and query info.
    ; PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_VM_OPERATION | PROCESS_QUERY_INFORMATION
    If !(hProc := DllCall("OpenProcess", "UInt", 0x438, "Int", False, "UInt", PID, "Ptr")) {
        Return
    }

    ; Should we use the 32-bit struct or the 64-bit struct?
    If (A_Is64bitOS) {
        Try DllCall("IsWow64Process", "Ptr", hProc, "Int*", Is32bit := true)
    } Else {
        Is32bit := True
    }

    RPtrSize := Is32bit ? 4 : 8
    TCITEM_SIZE := 16 + RPtrSize * 3

    ; Allocate a buffer in the (presumably) remote process.
    remote_item := DllCall("VirtualAllocEx", "Ptr", hProc, "Ptr", 0
                         , "uPtr", TCITEM_SIZE + MAX_TEXT_SIZE
                         , "UInt", 0x1000, "UInt", 4, "Ptr") ; MEM_COMMIT, PAGE_READWRITE
    remote_text := remote_item + TCITEM_SIZE

    ; Prepare the TCITEM structure locally.
    VarSetCapacity(TCITEM, TCITEM_SIZE, 0)
    NumPut(1, TCITEM, 0, "UInt") ; mask (TCIF_TEXT)
    NumPut(remote_text, TCITEM, 8 + RPtrSize) ; pszText
    NumPut(MAX_TEXT_LENGTH, TCITEM, 8 + RPtrSize * 2, "Int") ; cchTextMax

    ; Write the local structure into the remote buffer.
    DllCall("WriteProcessMemory", "Ptr", hProc, "Ptr", remote_item, "Ptr", &TCITEM, "UPtr", TCITEM_SIZE, "Ptr", 0)

    Tabs := []
    VarSetCapacity(TabText, MAX_TEXT_SIZE)

    SendMessage 0x1304, 0, 0,, ahk_id %hTab% ; TCM_GETITEMCOUNT
    Loop % (ErrorLevel != "FAIL") ? ErrorLevel : 0 {
        ; Retrieve the item text.
        SendMessage, % (A_IsUnicode) ? 0x133C : 0x1305, A_Index - 1, remote_item,, ahk_id %hTab% ; TCM_GETITEM
        If (ErrorLevel == 1) { ; Success
            DllCall("ReadProcessMemory", "Ptr", hProc, "Ptr", remote_text, "Ptr", &TabText, "UPtr", MAX_TEXT_SIZE, "Ptr", 0)
        } Else {
            TabText := ""
        }

        Tabs[A_Index] := TabText
    }

    ; Release the remote memory and handle.
    DllCall("VirtualFreeEx", "Ptr", hProc, "Ptr", remote_item, "UPtr", 0, "UInt", 0x8000) ; MEM_RELEASE
    DllCall("CloseHandle", "Ptr", hProc)

    Return Tabs
}

TabCtrl_GetCurSel(HWND) {                             		                                    	;-- index number of active tab in a gui
   ; Returns the 1-based index of the currently selected tab
   Static TCM_GETCURSEL := 0x130B
   SendMessage, TCM_GETCURSEL, 0, 0, , ahk_id %HWND%
   Return (ErrorLevel + 1)
}

TabCtrl_GetItemText(HWND, Index=0) {                                                        	;-- returns text of a tab

   Static TCM_GETITEM  := A_IsUnicode ? 0x133C : 0x1305 ; TCM_GETITEMW : TCM_GETITEMA
   Static TCIF_TEXT := 0x0001
   Static TCTXTP := (3 * 4) + (A_PtrSize - 4)
   Static TCTXLP := TCTXTP + A_PtrSize
   ErrorLevel := 0
   If (Index = 0)
      Index := TabCtrl_GetCurSel(HWND)
   If (Index = 0)
      Return SetError(1, "")
   VarSetCapacity(TCTEXT, 256 * SizeT, 0)
    VarSetCapacity(TCITEM, (5 * 4) + (2 * A_PtrSize) + (A_PtrSize - 4), 0)
   NumPut(TCIF_TEXT, TCITEM, 0, "UInt")
   NumPut(&TCTEXT, TCITEM, TCTXTP, "Ptr")
   NumPut(256, TCITEM, TCTXLP, "Int")
   SendMessage, TCM_GETITEM, --Index, &TCITEM, , ahk_id %HWND%
   If !(ErrorLevel)
      Return SetError(1, "")
   Else
      Return SetError(0, StrGet(NumGet(TCITEM, TCTXTP, "UPtr")))
}
;{ sub of TabCtrl_GetItemText
SetError(ErrorValue, ReturnValue) {                                                                                             	;--belongs to TabCtrl functions
   ErrorLevel := ErrorValue
   Return ReturnValue
}
;}

;\/\/\/\ Funktionen prüfen die erfolgreiche Durchführung ihrer Interaktion mit Steuerelementen /\/\/\/
VerifiedClick(CName, WTitle="", WText="", WinID="", WaitClose=0) {              	;-- 4 verschiedene Methoden um auf ein Control zu klicken

		tmm := A_TitleMatchMode, cd := A_ControlDelay, EL := 0
		SetTitleMatchMode 2
		SetControlDelay	 -1
		CoordMode, Mouse, Screen

	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		if (StrLen(WinID) > 0)
			WText := "", WTitle := "ahk_id " WinID
		else if RegExMatch(WTitle, "i)^(0x[A-F\d]+|[\d]+)$")
			WText := "", WTitle := "ahk_id " WTitle

	; 3 verschiedene Wege einen Buttonklick auszulösen
		ControlClick, % CName, % WTitle, % WText,,, NA
		If (EL := ErrorLevel) {                                                                            ; Misserfolg = 1 , Erfolg = 0
			ControlClick, % CName, % WTitle, % WText
			If (EL := ErrorLevel) {
               	SendMessage, 0x0201, 1, 0, % CName, % WTitle, % WText                 	;0x0201 - WM_Click
				EL := ErrorLevel = "FAIL" ? 1 : 0
				If (EL) {
					BlockInput, On
					WinGetPos    	, wx, wy,,, % WTitle, % WText
                   	ControlGetPos	, cx, cy, cw, ch, % CName, % WTitle, % WText
                   	MouseGetPos	, mx, my
                   	MouseClick   	, Left, % wx + cx + Floor(cw/2), % wy + cy + Floor(ch/2), 1, 0
                   	MouseMove  	, % mx, % my, 0
					BlockInput, Off
                    EL := 0
				}
			}
		}

		If (WaitClose > 0) {
			WinWaitClose, % WTitle, % WText, % WaitClose
			EL := ErrorLevel                                                                                ; Zeitlimit überschritten = 1, sonst 0
		}

		SetControlDelay	 % cd
		SetTitleMatchMode % tmm

return (EL = 0 ? 1 : 0)
}

VerifiedCheck(CName, WTitle="", WText="", WinID="", CheckIt=true) {          	;-- Fensteraktivierung + ControlDelay auf -1 + Kontrolle ob das Control wirklich checked ist jetzt

		; 04.08.2020 - Code effizienter gemacht
		; 02.08.2018 - neuer Parameter: CheckIt. Wenn dieser true, also gesetzt ist, wird ein Häkchen gesetzt , bei 'false' entfernt.
		; die Funktion prüft nicht, ob das Setzen oder Entfernen überhaupt notwendig ist, wenn es schon gesetzt oder nicht gesetzt ist

		command := CheckIt ? "Check":"UnCheck"
		tmm      	:= A_TitleMatchMode
		cd         	:= A_ControlDelay

		SetTitleMatchMode, 2
		SetControlDelay	, -1

		if WinID
			WTitle := "ahk_id " WinID, WText := ""
		else if RegExMatch(WTitle, "^0x[\w]+$", hex)
			WTitle := "ahk_id " WTitle, WText := "", WinID := hex
		else if RegExMatch(WTitle, "^\d+$", digits)
			WTitle := (StrLen(WTitle) = StrLen(digits) ? "ahk_id " digits : WTitle ), WText := "", WinID := digits

		If (StrLen(CName) > 0)
			ControlGet, hCName, hwnd,, % Trim(CName), % WTitle, % WText

		ButtonType := GetButtonType((hCName ? hCName : WinID))
		If !RegExMatch(ButtonType, "Autocheckbox|Checkbox|Radio")	{
			If (StrLen(CName) = 0)
				ControlGetText, CName,, % WTitle
			PraxTT("Fehler in der Funktion VerifiedCheck()`n`nDas angesprochene Steuerelement (" CName ")`nist keine Standard-Checkbox!", "5 0")
			return 0
		}

		If !WinActive(WTitle, WText) {
			WinActivate	 , % WTitle, % WText
			WinWaitActive, % WTitle, % WText, 1
		}

		Loop {
			Control		, % command,, % CName, % WTitle, % WText
            sleep, 20
			ControlGet, isChecked, checked,, % CName, % WTitle, % WText
		} until (isChecked = CheckIt) || (A_Index > 10)

		SetControlDelay	, % cd
		SetTitleMatchMode, % tmm

return (isChecked = CheckIt ? true : false)
}

VerifiedChoose(CName, WTitle, RxStrOrPos ) {                                                  	;-- wählt einen List- oder Comboboxeintrag

	; das gewünschte Listboxelement kann per Übergabe eines String, RegExString
	; oder direkt über seine Position ausgewählt werden
	; letzte Änderung: 17.01.2021

	; für flexible Übergabe des Fenstertitel, von String, Dezimalzahl oder Hexzahl alles möglich
		If RegExMatch(WTitle, "^0x[\w]+$")
			WTitle	:= RegExMatch(WTitle, "^0x[\w]+$")	? ("ahk_id " WTitle)	: (WTitle)
		else if RegExMatch(WTitle, "^\d+$", digits)
			WTitle	:= StrLen(WTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WTitle)
		else
			WTitle:= "ahk_id " WinID := GetHex(WinExist(WTitle, WText))

	; Funktionsabbruch bei inkompatiblem Steuerelement
		If !RegExMatch(CName, "i)^(Listbox|ComboBox)")
			return 2

	; Funktionsabbruch wenn Steuerelement nicht existiert
		ControlGet, CHwnd, Hwnd,, % CName, % WTitle
		If !CHwnd
			return 3

	; Funktionsabbruch wenn RxStrOrPos leer oder bei Übergabe einer Dezimalzahl kleiner gleich 0
		If (StrLen(RxStrOrPos) = 0)
			return 4

	; ermittelt die Einträge im Steuerelement
		ControlGet, CtrlList, List,,, % "ahk_id " CHwnd
		Items := StrSplit(CtrlList, "`n")

	; Auswahl anhand der Positionsnummer setzen
		If RegExMatch(RxStrOrPos, "^\d+$") {

			;Abbruch wenn die Positionsnummer nicht existiert
			If (Items.MaxIndex() < RxStrOrPos) || (RxStrOrPos <= 0)
				return 5

			Control, Choose, % RxStrOrPos,, % "ahk_id " CHwnd
			return ErrorLevel ? 6 : 1
		}

	; Auswahl anhand des übergebenen String setzen
		For idx, item in Items
			If InStr(item, RxStrOrPos) {
				Control, Choose, % idx,, % "ahk_id " CHwnd
				return ErrorLevel ? 7 : 1
			}

return 0
}

VerifiedSetFocus(CName, WTitle:="", WText:="", WinID:="", activate:=true) {     	;-- setzt den Eingabefokus und überprüft das dieser auch gesetzt wurde

	; Rückgabeparameter: 	erfolgreich - 	das Handle des Controls
	;                                	erfolglos  	-	0

		tmm:= A_TitleMatchMode, cd:=A_ControlDelay, idx := 0
		SetTitleMatchMode, 2
		SetControlDelay, -1

	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		if (StrLen(WinID) > 0)
			WText := "", WTitle := "ahk_id " WinID
		else if RegExMatch(WTitle, "i)^(0x[A-F\d]+|[\d]+)$")
			WText := "", WinID := GetHex(WinTitle), WTitle := "ahk_id " WTitle

		if activate {
			WinActivate	 , % WTitle
			WinWaitActive, % WTitle,, 1
		}

	; Focus setzen und überprüfen das der Focus gesetzt wurde
		If (StrLen(CName) > 0) {
			while !InStr(GetFocusedControlClassNN(WinID), CName) {
               	If (A_Index > 1)
					sleep, 200
              	ControlFocus, % CName, % WTitle
				idx ++
                If (idx > 10)
					break
			}
		}
		else {

			while !(GetFocusedControlHwnd() = WinID) {
				If (A_Index > 1)
                   	sleep, 200
                ControlFocus,, % WTitle
                idx ++
                If (idx > 10)
					break
			}
		}

		SetControlDelay, % cd
		SetTitleMatchMode, % tmm
		SciTEOutput("idx: " idx)

return idx = 0 ? 1 : 0
}

VerifiedSetText(CName="", NewText="", WTitle="", delay=100, WText="") {    	;-- erweiterte ControlSetText Funktion

	; kontrolliert ob der Text tatsächlich eingefügt wurde und man kann noch eine Verzögerung übergeben
	; delay = Zeit in ms zwischen den Versuchen
		abb:= delay > 2000 ? 20 : Floor(2000//delay)	;damit wird höchsten 2 Sekunden versucht den Text in das Control einzutragen

	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		If RegExMatch(WTitle, "^0x[\w]+$")
			WTitle	:= RegExMatch(WTitle, "^0x[\w]+$")	? ("ahk_id " WTitle)	: (WTitle)
		else if RegExMatch(WTitle, "^\d+$", digits)
			WTitle	:= StrLen(WTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WTitle)
		else
			WTitle:= "ahk_id " WinID := GetHex(WinExist(WTitle, WText))

	Loop 	{
		If (A_Index >= abb)
              return 0
		ControlSetText, % CName, % NewText, % WTitle, % WText
		sleep % delay
	} until (ControlGetText(CName, WTitle, WText) = NewText)

return (ControlGetText(CName, WTitle, WText) = NewText ? true : false)
}

UpSizeControl(WinTitle, WinClass, UpSizedControl, ExpandDown                      	;-- changes width and height of a control element and repositions the controls below and to the right of it
, ExpandRight, CenterToWin:=0) {

		static lastSizedWin

	; I don' t want to upsize the same win again
		If (lastSizedWin = hWin:=WinExist(WinTitle " ahk_class " WinClass))
			return 0
		lastSizedWin	:= hWin

	; defining variables
		under        	:= Object()
		right         	:= Object()
		g              	:= GetWindowSpot(hWin)
		hUpSized   	:= Controls(UpSizedControl, "ID", hWin)
		ctrls          	:= Controls("", "ID", hWin)

	; get positions of controls
		ControlGetPos, cx, cy, cw, ch,, % "ahk_id " hUpSized
		For CClass, Chwnd in ctrls
		{
				ControlGetPos, tx, ty, tw, th,, % "ahk_id " Chwnd
				If (ty >= (cy+ch))
					under.Push({"hwnd":Chwnd, "class":CClass, "x":tx, "y":ty, "w":w, "h":th})
				else If (tx >= (cx+cw))
					right.Push({"hwnd":Chwnd, "class":CClass, "x":tx, "y":ty, "w":w, "h":th})
		}

	; resizing and reposition the window
		If !CenterToWin
		{
				mon	:= GetMonitorIndexFromWindow(hWin)
				scr   	:= screenDims(mon)
				SetWindowPos(hWin, Floor(scr.W//2) - Floor((g.W + ExpandRight)//2), Floor(scr.H//2) - Floor((g.H + ExpandDown)//2), (g.W + ExpandRight), (g.H + ExpandDown))
		}
		else
		{
				Win	:= GetWindowSpot(CenterToWin)
				SetWindowPos(hWin, Floor(win.W//2) - Floor((g.W + ExpandRight)//2), Floor(win.H//2) - Floor((g.H + ExpandDown)//2), (g.W + ExpandRight), (g.H + ExpandDown))
		}

	; move the controls below and right of it
		Loop, % max := under.MaxIndex()
		{
				nr := Max - A_Index + 1
				ControlMove,,, % under[nr].y + ExpandDown,,, % "ahk_id " under[nr].hwnd
		}

		Loop, % right.MaxIndex()
			ControlMove,,% right[A_Index].x + ExpandRight,,,, % "ahk_id " right[A_Index].hwnd

	; upsize the control
		ControlMove,,,,% cw + ExpandRight, % ch + ExpandDown, % "ahk_id " hUpSized

	; redraw window now
		WinSet, Redraw,, % "ahk_id " hWin

return
}

;\/\/\/\/\/\/\/\/\/\/ Listview Control Funktionen \/\/\/\/\/\/\/\/\/\/
LVM_GetNext(hLV, rLV=0, oLV=0) {

	; hLV = ListView handle.
	; rLV = 1 based index of the starting row for the flag search. Omit or 0 to find first occurance specified flags.
	; oLV = Combination of one or more LVNI flags. See reference above.
	; LVNI_ALL := 0x0, LVNI_FOCUSED := 0x1, LVNI_SELECTED := 0x2
	;LVIS_SELECTED:=	2

Return DllCall("SendMessage", "uint", hLV, "uint", 4108, "uint", rLV-1, "uint", oLV) + 1 ; LVM_GETNEXTITEM
}

LV_Select(r, Control, hWin) {                                                                             	;-- select/deselect 1 to all rows of a listview (funktioniert nicht in fremder Listview)

	; Modified from http://www.autohotkey.com/board/topic/54752-listview-select-alldeselect-all/?p=343662
	; Examples: LVSel(1 , "SysListView321", "Win Title")   ; Select row 1. (or use +1)
	;           LVSel(-1, "SysListView321", "Win Title")   ; Deselect row 1
	;           LVSel(+0, "SysListView321", "Win Title")   ; Select all
	;           LVSel(-0, "SysListView321", "Win Title")   ; Deselect all
	;           LVSel(+0,                 , "ahk_id " HLV) ; Use listview's hwnd

	LVIS_FOCUSED:=1
	LVIS_SELECTED:=2
	LVM_SETITEMSTATE:=0x102B
	VarSetCapacity(LVITEM, 20, 0) ;to receive LVITEM
	NumPut(LVIS_FOCUSED | LVIS_SELECTED, LVITEM, 12)  ; state
	NumPut(LVIS_FOCUSED | LVIS_SELECTED, LVITEM, 16)  ; stateMask
	RemoteBuf_Open(hLVITEM, hWin, 20)  ; MASTER_ID = the ahk_id of the process owning the SysListView32 control
	RemoteBuf_Write(hLVITEM, LVITEM, 20)
	SendMessage, % LVM_SETITEMSTATE, % r, % RemoteBuf_Get(hLVITEM), % Control, % "ahk_id " hWin
	RemoteBuf_Close(hLVITEM)

}

LV_SortArrow(h, c, d="") {

	; LV_SortArrow by Solar. http://www.autohotkey.com/forum/viewtopic.php?t=69642
	; h = ListView handle
	; c = 1 based index of the column
	; d = Optional direction to set the arrow. "asc" or "up". "desc" or "down".

	static ptr, ptrSize, lvColumn, LVM_GETCOLUMN, LVM_SETCOLUMN
	if (!ptr)
		ptr := A_PtrSize ? ("ptr", ptrSize := A_PtrSize) : ("uint", ptrSize := 4)
		,LVM_GETCOLUMN := A_IsUnicode ? (4191, LVM_SETCOLUMN := 4192) : (4121, LVM_SETCOLUMN := 4122)
		,VarSetCapacity(lvColumn, ptrSize + 4), NumPut(1, lvColumn, "uint")
	c -= 1, DllCall("SendMessage", ptr, h, "uint", LVM_GETCOLUMN, "uint", c, ptr, &lvColumn)
	if ((fmt := NumGet(lvColumn, 4, "int")) & 1024) {
		if (d && d = "asc" || d = "up")
			return
		NumPut(fmt & ~1024 | 512, lvColumn, 4, "int")
	} else if (fmt & 512) {
		if (d && d = "desc" || d = "down")
			return
		NumPut(fmt & ~512 | 1024, lvColumn, 4, "int")
	} else {
		Loop % DllCall("SendMessage", ptr, DllCall("SendMessage", ptr, h, "uint", 4127), "uint", 4608)
			if ((i := A_Index - 1) != c)
				DllCall("SendMessage", ptr, h, "uint", LVM_GETCOLUMN, "uint", i, ptr, &lvColumn)
				,NumPut(NumGet(lvColumn, 4, "int") & ~1536, lvColumn, 4, "int")
				,DllCall("SendMessage", ptr, h, "uint", LVM_SETCOLUMN, "uint", i, ptr, &lvColumn)
		NumPut(fmt | (d && d = "desc" || d = "down" ? 512 : 1024), lvColumn, 4, "int")
	}

return DllCall("SendMessage", ptr, h, "uint", LVM_SETCOLUMN, "uint", c, ptr, &lvColumn)
}

LV_GetColWidth(hLV, ColN) {                                                                        	;-- gets the width of a column

	; from AutoGui
    SendMessage 0x101F, 0, 0,, % "ahk_id " hLV ; LVM_GETHEADER
    hHeader := ErrorLevel
    cbHDITEM := (4 * 6) + (A_PtrSize * 6)
    VarSetCapacity(HDITEM, cbHDITEM, 0)
    NumPut(0x1, HDITEM, 0, "UInt") ; mask (HDI_WIDTH)
    SendMessage, % A_IsUnicode ? 0x120B : 0x1203, ColN - 1, &HDITEM,, % "ahk_id " hHeader ; HDM_GETITEMW

Return (ErrorLevel != "FAIL") ? NumGet(HDITEM, 4, "UInt") : 0
}

LV_EX_GetTopIndex(HLV) {                                                                             	;-- retrieves the index of the topmost visible item when in list or report view
	; Author just me
   ; LVM_GETTOPINDEX = 0x1027 -> http://msdn.microsoft.com/en-us/library/bb761087(v=vs.85).aspx
   SendMessage, 0x1027, 0, 0, , % "ahk_id " . HLV
   Return (ErrorLevel + 1)
}

LV_GetScrollViewPos(hwnd) {

	Loop, % LV_GetCount() {
		SendMessage, 0x10B6, % A_Index - 1,,, % "ahk_id " hwnd 	; LVM_ISITEMVISIBLE -> findet das erste sichtbares Item
		If ErrorLevel {
			SciTEOutput("firstvisible item:" A_Index)
			return A_Index
		}
	}

}

LV_FindRow(LV, col, searchStr) {                                                                     	;-- search for a string in listview col, returns row

	Gui, adm: ListView, % LV

	Loop % LV_GetCount() {
		LV_GetText(cmpStr, A_Index, col)
		If InStr(cmpStr, searchStr)
			return A_Index
		}

return 0
}

CaretPos(ControlId) {                                                                                      	;-- Get start and End Pos of the selected string - Get Caret pos if no string is selected
	;https://autohotkey.com/boards/viewtopic.php?p=27979#p27979
	DllCall("User32.dll\SendMessage", "Ptr", ControlId, "UInt", 0x00B0, "UIntP", Start, "UIntP", End, "Ptr")
	SendMessage, 0xB1, -1, 0, , % "ahk_id" ControlId
	DllCall("User32.dll\SendMessage", "Ptr", ControlId, "UInt", 0x00B0, "UIntP", CaretPos, "UIntP", CaretPos, "Ptr")
	if (CaretPos = End)
		SendMessage, 0xB1, % Start, % End, , % "ahk_id" ControlId	;select from left to right ("caret" at the End of the selection)
	else
		SendMessage, 0xB1, % End, % Start, , % "ahk_id" ControlId	;select from right to left ("caret" at the Start of the selection)
	CaretPos++	;force "1" instead "0" to be recognised as the beginning of the string!
return, CaretPos
}

; Hilfsfunktionen
KeyValueObjectFromLists(keyList, valueList, delimiter:="`n"
, IncludeKeys:="", KeyREx:="", IncludeValues:="", ValueREx:="") {                	;-- Funktion um z.B. zwei Listen aus WinGet zusammenzuführen

	keyArr:= valueArr:= []
	merged:= Object()
	mustMatches:=0

	If !(IncludeKeys = "")
			mustMatches+=1
	If !(IncludeValues = "")
			mustMatches+=1

	keyArr		:= StrSplit(keyList	 , delimiter)
	valueArr	:= StrSplit(valueList, delimiter)

	Loop % keyArr.MaxIndex()
	{
				If (KeyREx = "")
					mkey:= keyArr[A_Index]
				else
					RegExMatch(keyArr[A_Index], KeyREx, mkey)

				If (ValueREx = "")
					mval:= valueArr[A_Index]
				else
					RegExMatch(valueArr[A_Index], ValueREx, mval)

				matched:=0
				If IncludeKeys != ""
						If mkey in %IncludeKeys%
                            	matched:= 1
				else
						matched:= 1

				If IncludeValues != ""
						If mval in %IncludeValues%
                            	matched += 1
				else
						matched += 1

				If (matched>mustMatches)
						merged[(keyArr[A_Index])]:= valueArr[A_Index]
	}

return merged
}

KeyValueObjectFromLists2(keyList, valueList, delimiter:="`n"
, IncludeKeys:="", KeyREx:="", IncludeValues:="", ValueREx:="") {                	;-- wird neue Funktion um z.B. zwei Listen aus WinGet zusammenzuführen

	keyArr := valueArr:= Array()
	merged := Object()

	mustMatches:=0
	mustMatches += (StrLen(IncludeKeys) > 0) 	? 1 : 0
	mustMatches += (StrLen(IncludeValues) > 0)	? 1 : 0

	keyArr		:= StrSplit(keyList	 , delimiter)
	valueArr	:= StrSplit(valueList, delimiter)

	Loop % keyArr.MaxIndex()	{

		If (StrLen(KeyREx) = 0)
			mkey:= keyArr[A_Index]
		else
			RegExMatch(keyArr[A_Index], KeyREx, mkey)

		If (StrLen(ValueRex) = 0)
			mval := valueArr[A_Index]
		else
			RegExMatch(valueArr[A_Index], ValueREx, mval)

		matched:=0
		If (StrLen(IncludeKeys) > 0) {
			If mkey in %IncludeKeys%
				matched := 1
			else
				matched := 1
		}

		If (StrLen(IncludeValues) > 0)
			If mval in %IncludeValues%
				matched += 1
			else
				matched += 1

		If (matched > mustMatches)
			merged[keyArr[A_Index]]:= valueArr[A_Index]

	}

return merged
}

