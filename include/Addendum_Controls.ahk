; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 17.07.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
; CONTROLS                                                                                                                                                                                                                                        	(40)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; GetClassName                                	Control_GetClassNN                     	GetClassNN                                  	GetFocusedControl                        	GetFocusedControlHwnd
; GetFocusedControlClassNN            	GetChildHWND                            	GetControls                                  	GetButtonType                                	Controls
; ControlFind                                     	ControlGet                                    	ControlGetText                              	ControlGetFocus
; WinSaveCheckboxes                        	Toolbar_GetRect                            	Toolbar_GetMaxSize
; ControlGetTabs                               	TabCtrl_GetCurSel                          	TabCtrl_GetItemText
; VerifiedClick                                    	VerifiedCheck                                	VerifiedChoose                              	VerifiedSetFocus                            	VerifiedSetText
; UpSizeControl
; LV_EX_FindString                             	LV_GetItemState                            	LV_GetItemState2                          	LV_GetItemText                             	LV_ItemText
; LVM_GetText                                   	LVM_GetNext                                	LV_MouseGetCellPos							LV_SortArrow
; RichEdit_FindText                             	RE_FindText                                   	RE_GetSel                                     	RE_GetTextLength                          	RE_ReplaceSel
; RE_ScrollCaret                                 	RE_SetSel
;_________________________________________________________________________________________________________________________________________________________
GetClassName(hwnd) {                                                                             	;-- returns HWND's class name without its instance number, e.g. "Edit" or "SysListView32"

		;https://autohotkey.com/board/topic/45515-remap-hjkl-to-act-like-left-up-down-right-arrow-keys/#entry283368
			VarSetCapacity( buff, 256, 0 )
			DllCall("GetClassName", "uint", hwnd, "str", buff, "int", 255 )
			return buff
}

Control_GetClassNN(hWnd, hCtrl) {
	; SKAN: www.autohotkey.com/forum/viewtopic.php?t=49471
 WinGet, CH, ControlListHwnd, ahk_id %hWnd%
 WinGet, CN, ControlList, ahk_id %hWnd%
 LF:= "`n",  CH:= LF CH LF, CN:= LF CN LF,  S:= SubStr( CH, 1, InStr( CH, LF hCtrl LF ) )
 StringReplace, S, S,`n,`n, UseErrorLevel
 StringGetPos, LP, CN, `n, L%ErrorLevel%
 Return SubStr( CN, LP+2, InStr( CN, LF, 0, LP+2 ) -LP-2 )
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
   guiThreadInfoSize := 8 + 6 * A_PtrSize + 16
   VarSetCapacity(guiThreadInfo, guiThreadInfoSize, 0)
   NumPut(GuiThreadInfoSize, GuiThreadInfo, 0)
   ; DllCall("RtlFillMemory" , "PTR", &guiThreadInfo, "UInt", 1 , "UChar", guiThreadInfoSize)   ; Below 0xFF, one call only is needed
   if (DllCall("GetGUIThreadInfo" , "UInt", 0, "PTR", &guiThreadInfo) = 0) {   ; Foreground thread
			ErrorLevel := A_LastError   ; Failure
			Return 0
   }

Return NumGet(guiThreadInfo, 8+A_PtrSize, "Ptr") ; *(addr + 12) + (*(addr + 13) << 8) +  (*(addr + 14) << 16) + (*(addr + 15) << 24)
}

GetFocusedControlHwnd(hwnd:="A") {
	ControlGetFocus, FocusedControl, % (hwnd = "A") ? "A" : "ahk_id " hwnd
	ControlGet, FocusedControlId, Hwnd,, %FocusedControl%, % (hwnd = "A") ? "A" : "ahk_id " hwnd
return FocusedControlId
}

GetFocusedControlClassNN(hwnd:="A") {
	ControlGetFocus, FocusedControl, % hwnd = "A" ? "A" : "ahk_id " hwnd
	ControlGet, FocusedControlId, Hwnd,, % FocusedControl, % "ahk_id " hwnd
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

GetControls(hwnd, class_filter:="", type_filter:=""
, info_filter:="hwnd,Pos,Enabled,Visible,Style,ExStyle") {			                      	  	;-- returns an array with ClassNN, ButtonTyp, Position.....

	;class_filter - comma separated list of classes you don't want to store
	;type_filter - comma separated list of classes you don't want to store
	;info_filter - comma separated list of classes you !want! to store

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

Controls(Control, command, WinTitle
, HiddenText:=true, HiddenWin:=true, MatchModeSpeed:="slow") {                	;-- Universalfunktion für Steuerelemente

	; ********	    ********		Funktion wächst und gedeiht, gegossen am 14.07.2020
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
		CoordMode              	, Pixel	, Screen
		sleep, 10          	; CoordMode needs a pause to update - https://www.autohotkey.com/boards/viewtopic.php?f=14&t=38467
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Ermitteln aller Steuerelementklassen und Handle des Fensters
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		RegExMatch(Control, "[a-zA-Z#]+", class)

		If !InStr(knWinTitle, WinTitle) 		{

				knWinTitle	:= WinTitle
				knWinText := ( WinText = "" ) ? 0 : WinText
				WinTitle	:= RegExMatch(WinTitle, "^0x[\w]+$")	? ("ahk_id " WinTitle)	: (WinTitle)
				RegExMatch(WinTitle, "^\d+$", digits)
				WinTitle	:= StrLen(WinTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WinTitle)

				WinGet, cClasses	, ControlList			, % WinTitle, ;% WinText                            ; use this for example: "ahk_id " hWin
				WinGet, cHwnds	, ControlListHwnd	, % WinTitle, ;% WinText
				Ctrl := KeyValueObjectFromLists(cClasses, cHwnds, "`n", "", "", "", "")                	; ergibt ein Object mit ClassNN als key und handle als value

		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Befehlsbereich -
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	    if RegExMatch(command, "Hwnd|ID")                  	{   	; returns the handle for a ClassNN
				; empty Control parameter returns the array
					If (StrLen(Control) = 0)
						return Ctrl
				; this returns the handle if Control fits to a key
                    else If Ctrl.HasKey(Control)
						return Ctrl[Control]
				; this returns the handle if you don't want to use the ClassNN
					else
					   For ControlClass, ControlHwnd in Ctrl
							If InStr(ControlClass, control)
								return ControlHwnd
								;return Ctrl[ControlClass]
		}
		else if InStr(command, "Click"	)                           	{   	; like ControlClick, you need the exactly name of classNN, but you can specify the method to use (MouseClick, ControlClick)

					If InStr(control, "ToolbarWindow") 			{

							ControlGetPos, cx, cy, cw, ch,, % "ahk_id " Ctrl[Control]
							ClientToScreen(Ctrl[Control], cx, xy)

								If InStr(command, "Space")                            ;Controls("ToolbarWindow324", "Click use Spacebar", "ahk_exe notepad.exe")
								{
										ControlFocus,, % "ahk_id " Ctrl[Control]
										sleep, 400
										ControlSend,, {Space}, % "ahk_id " Ctrl[Control]
								}
								else if InStr(command, "ControlClick"	)		;Controls("ToolbarWindow324", "Click use Controlclick left", "ahk_exe notepad.exe")
								{
										RegExMatch(command, "i)(?<=ControlClick\s)\w+", Button)
										if Button in Left,Middle,Right
												ControlClick,, % "ahk_id " Ctrl[Control],, % Button, 1, NA
								}
								else if InStr(command, "MouseClick"	)		;Controls("ToolbarWindow324", "Click use MouseClick", "ahk_exe notepad.exe")
								{
										BlockInput, On
										MouseGetPos, mx, my
										ControlGetPos, cx, cy, cw, ch,, % "ahk_id " Ctrl[Control]
										ClientToScreen(Ctrl[Control], cx, xy)
										RegExMatch(command, "i)(?<=MouseClick\s)\w+", Button)
										if Button in Left,Middle,Right
											MouseClick, % Button, % (cx + cw - 50), % (cy + ch//2), 1, 0
										MouseMove, % mx, % my, 0
										BlockInput, Off
								}

				}
				else if InStr(control, "Button")                            {

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
		else if InStr(command, "ControlClick")	                	{   	; clicks a control by its text or classNN and returns the ErrorLevel

				; möglicher Syntax z.B. Controls("", "ControlClick, Speichern, Button", "YourWinTitle")
                    searchText		:= Trim( StrSplit(command, ",").2 )
                    searchClass	:= Trim( StrSplit(command, ",").3 )

                    If (searchText <> "")
                    {
                            For ControlClass, ControlHwnd in Ctrl
                            {
                            	ControlGetText, ControlText,, % "ahk_id " ControlHwnd
                            	If ( InStr(ControlClass, searchClass) && InStr(ControlText, searchText) )
                            			return VerifiedClick(ControlClass, WinTitle, "", "")
                            }
                    }
                    else
                    {
                            return VerifiedClick(searchClass, WinTitle, "", "")		;searchClass must be the exact ClassNN in this case
                    }

		}
		else if InStr(command, "ControlFind")	                	{   	; finds a control by its text and returns it's ControlClassNN

				; möglicher Syntax z.B. Controls( "ControlFind, Speichern, Button)" )
                    searchText		:= Trim(StrSplit(command, ",").2)
                    searchClass	:= Trim(StrSplit(command, ",").3)

                    For ControlClass, ControlHwnd in Ctrl
                    {
                    	ControlGetText, ControlText,, % "ahk_id " ControlHwnd
                    	If InStr(ControlClass, searchClass) && InStr(ControlText, searchText)
                            	return ControlClass
                    }

		}
		else if RegExMatch(command, "^ControlPos")       	{   	; returns the controls position inside window
				ControlGetPos, x,y,w,h, % Control, % WinTitle, % WinText
				return {"X":x, "Y":y, "W":w, "H":h}
		}
		else if RegExMatch(command, "^GetControls")     	{   	; try's to return all subcontrol hwnds (no treeversal)

				found	:= false
				HiddenControls := true
				Childs	:= Array()

				If !RegExMatch(command, ".*\+Hidden") {
					GCHiddenTextStatus      := A_DetectHiddenText
					GCHiddenWinStatus      := A_DetectHiddenWindows
					DetectHiddenText      	, Off
					DetectHiddenWindows	, Off
					HiddenControls := false
				}

				If (StrLen(Control) > 0) {

					For ControlClass, ControlHwnd in Ctrl
						If InStr(ControlClass, Control) || InStr(ControlGetText(ControlHwnd), Control) {

							WinTitle := "ahk_id " ControlHwnd
							found := true
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
		else if InStr(command, "GetFocus")	                    	{   	; finds the focused control and returns it's ControlClassNN

					For ControlClass, ControlHwnd in Ctrl
					{
							If DllCall("IsWindow", "Ptr", ControlHwnd)
							{
									ControlGetFocus, cFocus, % "ahk_id " ControlHwnd
									If (StrLen(cFocus) > 0)
									    	return cFocus
							}
					}

					return cFocus
		}
		else if InStr(command, "GetText")                        	{
				 If class in Edit,ToolbarWindow
                    	ControlGetText, result,	 , % "ahk_id " Ctrl[Control]
				else if class in ComboBox,ListBox,Listview,DropDownList
                    	ControlGet, result, List  ,,, % "ahk_id " Ctrl[Control]
		}
		else if InStr(command, "Send")                            	{
				if class in Edit,RichEdit
				{
                    RegExMatch(command, "i)(?<=ControlSend|Send\s+).*", keys)
                    ControlSend,, % keys, % "ahk_id " Ctrl[(Control)]
				}
		}
		else if InStr(command, "SendRaw")                      	{
				If class in Edit,RichEdit
				{
                    	RegExMatch(command, "i)(?<=ControlSend\s).*", keys)
                    	ControlSendRaw,, % keys, % "ahk_id " Ctrl[(Control)]
				}
		}
		else if InStr(command, "SetFocus")                       	{
				if class in Edit,RichEdit,ComboBox,ListBox,ListView,DropDownList,ToolbarWindow
				{
                    	ControlFocus,, % "ahk_id " Ctrl[(Control)]
                    	return ErrorLevel
				}
		}
		else if InStr(command, "SetText")                         	{
				If class in Edit,RichEdit
				{
                    	RegExMatch(command, "i)(?<=SetText\s|\,|\s,).*", NewText)
                    	If !NewText
                            MsgBox, % "not matching command syntax for SetText command. " A_LineFile
                    	return VerifiedSetText("", NewText, "ahk_id " Ctrl[(Control)], 100)
				}
		}
		else if InStr(command, "Reset")                            	{
                    	VarSetCapacity(Ctrl, 0)
                    	Ctrl			:= Object()
                    	result		:= 1
                    	knWinTitle	:= ""
                    	knWinText	:= 0
                    	return 1
		}
		else if InStr(command, "ControlCount")                   {  	; returns the count of all found controls
					return Ctrl.Count()
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

ControlFind(Control, command, WinTitle) {				                                    	;-- Controls ist die bessere Version jetzt

		static knWinTitle, knWinText
		static Ctrl	:= Object()

		If RegExMatch(WinTitle, "^0x")
				WinTitle := "ahk_id " WinTitle

		If InStr(command, "reset")
		{
				VarSetCapacity(Ctrl, 0)
				Ctrl := Object()
				return 1
		}

		If !InStr(knWinTitle, WinTitle)
		{
				knWinTitle:= WinTitle
				If WinText=""
                    	knWinText:= 0
				else
                    	knWinText:= WinText

				WinGet, cClasses	, ControlList			, % WinTitle, ;% WinText                            	; use this for example: "ahk_id " hWin
				WinGet, cHwnds	, ControlListHwnd	, % WinTitle, ;% WinText
				Ctrl	:= KeyValueObjectFromLists(cClasses, cHwnds, "`n", "", "", "", "")					; ergibt ein Object mit ClassNN und dem handle
		}


		If InStr(command, "ID")
				return Ctrl[(Control)]
		else if InStr(command, "GetText")
		{
				RegExMatch(control, "[a-zA-Z]+", class)

				If class in Edit,ToolbarWindow
				{
                    ControlGetText, res,, % "ahk_id " Ctrl[(Control)]
                    return res
				}
				else If class in ComboBox,ListBox
				{
                    ControlGet, res, Choice,,, % "ahk_id " Ctrl[(Control)]
                    return res
				}
		}


		If InStr(command, "GetList") {
				If InStr(Control, "ComboBox") || InStr(Control, "ListBox") || InStr(Control, "Listview") InStr(Control, "DropDownList")
				{
                    ControlGet, res, List,,, % "ahk_id " Ctrl[(Control)]  ;Choice
                    return res
				}
		}


}

ControlGet(Cmd,Value:="",Ctrl:="",WTitle:="",WTxt:="",ExTitle:="",ExTxt:="") {	;-- ControlGet als Funktion
	ControlGet, v, % Cmd, % Value, % Ctrl, % WTitle, % WTxt, % ExTitle, % ExTxt
Return v
}

ControlGetText(Control:="", WinTitle:="", WinText:=""
, ExcludeTitle:="", ExcludeText:="") {                                                                 	;-- ControlGetText als Funktion
	ControlGetText, v, %Control%, %WinTitle%, %WinText%, %ExcludeTitle%, %ExcludeText%
	Return v
}

ControlGetFocus(hwnd) {                                                                                	;-- gibt das Handle des fokussierten Steuerelementes zurück
	ControlGetFocus, FocusedControl, % "ahk_id " hwnd
	ControlGet, FocusedControlId, Hwnd,, %FocusedControl%, % "ahk_id " hwnd
return FocusedControlId
}

WinSaveCheckboxes(hWin) {                                                                            	;-- speichert den Status (Haken gesetzt oder nicht) in ein Objekt, z.B. um den Ursprungszustand wieder herstellen zu können

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

Toolbar_GetMaxSize(hCtrl, ByRef Width, ByRef Height) {                                   	;-- zur Ermitlung der maximalen Größe einer Toolbar

    /*
             Function:   	GetMaxSize
								Retrieves the total size of all of the visible buttons and separators in the toolbar.
             Parameters:	Width, Height		- Variables which will receive size.
             Returns:    	Returns TRUE if successful.
     */

	static TB_GETMAXSIZE = 0x453

	VarSetCapacity(SIZE, 8)
	SendMessage, TB_GETMAXSIZE, 0, &SIZE, , ahk_id %hCtrl%
	res := ErrorLevel, 	Width := NumGet(SIZE), Height := NumGet(SIZE, 4)
	return res
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
VerifiedClick(CName, WTitle="", WText="", WinID="", WaitClose=false) {       	;-- 4 verschiedene Methoden um auf ein Control zu klicken

		tmm:= A_TitleMatchMode, cd:=A_ControlDelay, EL:= 0
		SetTitleMatchMode, 2
		SetControlDelay, -1

	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		if StrLen(WinID) > 0	{
				WTitle := "ahk_id " WinID
				WText := ""
		} else if RegExMatch(WTitle, "^0x[\w]+$") {
				WTitle := RegExMatch(WTitle, "^0x[\w]+$")	? ("ahk_id " WTitle)	: (WTitle)
				WText := ""
		} else if RegExMatch(WTitle, "^\d+$", digits) {
				WTitle := StrLen(WTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WTitle)
				WText := ""
		}

		If !WinActive(WTitle, WText) {
				WinActivate	 , % WTitle, % WText
				WinWaitActive, % WTitle, % WText, 1
		}

	; 3 verschiedene Wege einen Buttonklick auszulösen
		ControlClick, % CName, % WTitle, % WText,,, NA
		If (EL := ErrorLevel)
		{
				ControlClick, % CName, % WTitle, % WText
				If (EL :=ErrorLevel)
				{
                    	SendMessage, 0x0201, 1, 0, % CName, % WTitle, % WText                 	;0x0201 - WM_Click
                    	If (EL:=ErrorLevel)
                    	{
                            	ControlGetPos, cx, cy, cw, ch, % CName, % WTitle, % WText
                            	MouseGetPos, mx, my
                            	MouseClick, Left, % cx + (cw//2), % cy + (ch//2)
                            	MouseMove, % mx, % my, 0
                            	EL := 0
                    	}
				}
		}

		If WaitClose {
			WinWaitClose, % WTitle, % WText, 3
			EL:= ErrorLevel
		}

		SetControlDelay	, % cd
		SetTitleMatchMode, % tmm

return (EL = 0 ? 1 : 0)
}

VerifiedCheck(CName, WTitle="", WText="", WinID="", CheckIt=true) {          	;-- Fensteraktivierung + ControlDelay auf -1 + Kontrolle ob das Control wirklich checked ist jetzt

		; 02.08.2018 - neuer Parameter: CheckIt. Wenn dieser 1, also gesetzt ist, wird ein Häkchen gesetzt , bei 'false' entfernt.
		; die Funktion prüft nicht, ob das Setzen oder Entfernen überhaupt notwendig ist, wenn es schon gesetzt oder nicht gesetzt ist
		; Achtung: diese Funktion macht einen unendlichen Loop wenn

		command	:= "UnCheck"
		tmm      	:= A_TitleMatchMode
		cd         	:= A_ControlDelay

		SetTitleMatchMode, 2
		SetControlDelay, -1

		if WinID {
				WTitle	:= "ahk_id " WinID
				WText	:= ""
		} else if RegExMatch(WTitle, "^0x[\w]+$") {
				WTitle	:= "ahk_id " WTitle
				WText	:= ""
		} else if RegExMatch(WTitle, "^\d+$", digits) {
				WTitle	:= StrLen(WTitle) = StrLen(digits) ? ("ahk_id " digits) : (WTitle)
				WText	:= ""
		}

		ControlGet, hCName, hwnd,, % Trim(CName), % WTitle, % WText

		If !InStr(GetButtonType(hCName), "Checkbox")
		{
				PraxTT("Fehler in der Funktion VerifiedCheck()`n`nDas angesprochene Steuerelement`nist keine Standard-Checkbox!", "2 0")
				return
		}

		If !WinActive(WTitle, WText) {
				WinActivate	 , % WTitle, % WText
				WinWaitActive, % WTitle, % WText, 1
		}

		If CheckIt
			command:= "Check"

		Loop {
			Control	, % command,, % CName, % WTitle, % WText
            sleep, 50
			ControlGet, isChecked, checked,, % CName, % WTitle, % WText
		} until (isChecked = CheckIt)

		SetControlDelay	, % cd
		SetTitleMatchMode, % tmm

		If ChecktIt
			return (isChecked = CheckIt ? true : false)

return ErrorLevel
}

VerifiedChoose(CName, WTitle, EntryNr ) {                                                     	;-- wählt einen List- oder Comboboxeintrag

	; letzte Änderung: 17.07.2020

	; für flexible Übergabe des Fenstertitel, von String, Dezimalzahl oder Hexzahl alles möglich
		If RegExMatch(WTitle, "^0x[\w]+$")
			WTitle	:= RegExMatch(WTitle, "^0x[\w]+$")	? ("ahk_id " WTitle)	: (WTitle)
		else if RegExMatch(WTitle, "^\d+$", digits)
			WTitle	:= StrLen(WTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WTitle)
		else
			WTitle:= "ahk_id " WinID := GetHex(WinExist(WTitle, WText))

	; Funktionsabbruch bei inkompatiblem Steuerelement
		If !RegExMatch(CName, "^Listbox|ComboBox")
			return 0

	; Funktionsabbruch wenn Steuerelement nicht existiert
		ControlGet, CHwnd, Hwnd,, % CName, % WTitle
		If !CHwnd
			return 0

	; Funktionsabbruch wenn EntryNr leer oder bei Übergabe einer Dezimalzahl kleiner gleich 0
		If (StrLen(EntryNr) = 0) || (EntryNr <= 0)
			return 0

	; ermittelt die Einträge im Steuerelement
		ControlGet, CtrlList, List,,, % "ahk_id " CHwnd
		Items := StrSplit(CtrlList, "`n")

	; Auswahl anhand der Positionsnummer setzen
		If RegExMatch(EntryNr, "^\d+$") {

			;Abbruch wenn die Positionsnummer nicht existiert
			If (Items.MaxIndex() < EntryNr)
				return 0

			Control, Choose, % EntryNr,, % "ahk_id " CHwnd
			return ErrorLevel
		}

	; Auswahl anhand des übergebenen String setzen
		For idx, item in Items
			If RegExMatch(item, EntryNr) {
				Control, Choose, % idx,, % "ahk_id " CHwnd
				return ErrorLevel
			}

}

VerifiedSetFocus(CName, WTitle:="", WText:="", WinID:="") {                         	;-- setzt den Eingabefokus und überprüft das dieser auch gesetzt wurde

	; Rückgabeparameter: 	erfolgreich - 	das Handle des Controls
	;                                	erfolglos  	-	0

		tmm:= A_TitleMatchMode, cd:=A_ControlDelay, idx := 0
		SetTitleMatchMode, 2
		SetControlDelay, -1

	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		if WinID {
				WTitle:= "ahk_id " WinID
				WText:= ""
		} else if RegExMatch(WTitle, "^0x[\w]+$") {
				WTitle	:= RegExMatch(WTitle, "^0x[\w]+$")	? ("ahk_id " WTitle)	: (WTitle)
		} else if RegExMatch(WTitle, "^\d+$", digits) {
				WTitle	:= StrLen(WTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WTitle)
		} else {
				WTitle:= "ahk_id " WinID := GetHex(WinExist(WTitle, WText))
		}

		WinActivate	 , % WTitle
		WinWaitActive, % WTitle,, 1

	; Focus setzen und überprüfen das der Focus gesetzt wurde
		If StrLen(CName) > 0
		{
				while !InStr(GetFocusedControlClassNN(WinID), CName)
				{
                    	If A_Index > 1
                            	sleep, 200
                    	ControlFocus, % CName, % WTitle
                    	idx ++
                    	If idx > 10
                            break
				}
		}
		else
		{
				while !(GetFocusedControlHwnd(), WinID)
				{
                    	If A_Index > 1
                            	sleep, 200
                    	ControlFocus,, % WTitle
                    	idx ++
                    	If idx > 10
                            break
				}
		}

		SetControlDelay, % cd
		SetTitleMatchMode, % tmm

return idx = 0 ? 1 : 0
}

VerifiedSetText(CName="", NewText="", WTitle="", delay=200, WText="") {    	;-- erweiterte ControlSetText Funktion

	; kontrolliert ob der Text tatsächlich eingefügt wurde und man kann noch eine Verzögerung übergeben
	; delay = Zeit in ms zwischen den Versuchen
		abb:= delay > 2000 ? 1 : Floor(2000//delay)	;damit wird höchsten 2 Sekunden versucht den Text in das Control einzutragen

	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		If RegExMatch(WTitle, "^0x[\w]+$")
			WTitle	:= RegExMatch(WTitle, "^0x[\w]+$")	? ("ahk_id " WTitle)	: (WTitle)
		else if RegExMatch(WTitle, "^\d+$", digits)
			WTitle	:= StrLen(WTitle) = StrLen(digits)     	? ("ahk_id " digits)  	: (WTitle)
		else
			WTitle:= "ahk_id " WinID := GetHex(WinExist(WTitle, WText))

	Loop 	{

		If (A_Index > abb)
              return 0
		ControlSetText, % CName, % NewText, % WTitle, % WText
		sleep % delay

	} until (ControlGetText(CName, WTitle, WText) = NewText)

return (ControlGetText(CName, WTitle, WText) = NewText ? true : false)
}

UpSizeControl(WinTitle, WinClass, UpSizedControl, ExpandDown                      	;-- changes the width and height of a control element and repositions the controls below and to the right of it
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
LV_EX_FindString(HLV, Str, Start := 0, Partial := False) {                       				;-- gibt die Zeilennummer zurück in welchem sich der gesuchte Text befindet
   ; LVM_FINDITEM -> http://msdn.microsoft.com/en-us/library/bb774903(v=vs.85).aspx
   Static LVM_FINDITEM := A_IsUnicode ? 0x1053 : 0x100D ; LVM_FINDITEMW : LVM_FINDITEMA
   Static LVFISize := 40
   VarSetCapacity(LVFI, LVFISize, 0) ; LVFINDINFO
   Flags := 0x0002 ; LVFI_STRING
   If (Partial)
      Flags |= 0x0008 ; LVFI_PARTIAL
   NumPut(Flags, LVFI, 0, "UInt")
   NumPut(&Str,  LVFI, A_PtrSize, "Ptr")
   SendMessage, % LVM_FINDITEM, % (Start - 1), % &LVFI, , % "ahk_id " . HLV
   Return (ErrorLevel > 0x7FFFFFFF ? 0 : ErrorLevel + 1)
}

LV_GetItemState(HLV, Row) {                                                                         	;-- den Status einer Listviewzeile ermitteln

   Static LVM_GETITEMSTATE := 0x102C
   Static LVIS := {Cut: 0x04, DropHilited: 0x08, Focused: 0x01, Selected: 0x02, Checked: 0x2000}
   Static ALLSTATES := 0xFFFF ; not defined in MSDN
   SendMessage, % LVM_GETITEMSTATE, % (Row - 1), % ALLSTATES, , % "ahk_id " . HLV
   If (ErrorLevel + 0) {
      States := ErrorLevel
      Result := {}
      For Key, Value In LVIS
         Result[Key] := !!(States & Value)
      Return Result
   }

 Return False
}

LV_GetItemState2(HLV, Row) {                                                                        	;-- wie darüber, da LV_GetItemState nicht immer funktionierte

   Static LVM_GETITEMSTATE := 0x102C
   Static LVIS1 := {Cut: 0x04, DropHilited: 0x08, Focused: 0x01, Selected: 0x02, Checked: 0x2000}
   Static LVIS2 := {0x04:Cut, 0x08:DropHilited, 0x01:Focused, 0x02:Selected, 0x2000:Checked}
   Static ALLSTATES := 0xFFFF ; not defined in MSDN
   SendMessage, % LVM_GETITEMSTATE, % (Row - 1), % ALLSTATES, , % "ahk_id " . HLV
   If (ErrorLevel + 0) {
      States := GetHex(ErrorLevel)
      For Key, Value In LVIS2
	 {
			If InStr(States, Value)
                    result:= Value
	}
      Return Result.= " (" Key ")"
   }

 Return False
}

LV_GetItemText(item_index, sub_index, ctrl_id, win_id) {                         				;-- read the text from an item in a TListView

		; https://autohotkey.com/board/topic/18299-reading-listview-of-another-app/    ----  code from Tigerite
		MAX_TEXT:= 260
		item_index -= 1
        VarSetCapacity(szText	, MAX_TEXT, 0)
        VarSetCapacity(szClass	, MAX_TEXT, 0)
        ControlGet, hListView, Hwnd, , % ctrl_id, % "ahk_id " win_id
        DllCall("GetClassName", "UInt",hListView, "Str",szClass, "Int",MAX_TEXT)
        if (DllCall("lstrcmpi", "Str",szClass, "Str","SysListView32") == 0 || DllCall("lstrcmpi", "Str",szClass, "Str","TListView") == 0)
            LV_ItemText(hListView, item_index, sub_index, szText, MAX_TEXT)

return %szText%
}

LV_ItemText(hListView, iItem, iSubItem, ByRef lpString, nMaxCount) {            	;--

        ;const
        LVNULL                            	:= 0
        PROCESS_ALL_ACCESS 	:= 0x001F0FFF
        INVALID_HANDLE_VALUE	:= 0xFFFFFFFF
        PAGE_READWRITE         	:= 4
        FILE_MAP_WRITE             	:= 2
        MEM_COMMIT             	:= 0x1000
        MEM_RELEASE               	:= 0x8000
        LV_ITEM_mask              	:= 0
        LV_ITEM_iItem               	:= 4
        LV_ITEM_iSubItem         	:= 8
        LV_ITEM_state                 	:= 12
        LV_ITEM_stateMask       	:= 16
        LV_ITEM_pszText              	:= 20
        LV_ITEM_cchTextMax      	:= 24
        LVIF_TEXT                     	:= 1
        LVM_GETITEM                	:= 0x1005
        SIZEOF_LV_ITEM             	:= 0x28
        SIZEOF_TEXT_BUF         	:= 0x104
        SIZEOF_BUF                     := 0x120
        SIZEOF_INT                     	:= 4
        SIZEOF_POINTER             	:= 4

        ;var
        result        	:= 0
        hProcess    	:= LVNULL
        dwProcessId	:= 0

        if (lpString <> LVNULL) && (nMaxCount > 0)        {

            DllCall("lstrcpy", "Str",lpString, "Str","")
            DllCall("GetWindowThreadProcessId", "UInt", hListView, "UIntP", dwProcessId)
            hProcess := DllCall("OpenProcess", "UInt", PROCESS_ALL_ACCESS, "Int", false, "UInt", dwProcessId)
            if (hProcess <> LVNULL)  {

                ;var
                lpProcessBuf  	:= LVNULL
                hMap            	:= LVNULL
                hKernel         	:= DllCall("GetModuleHandle", Str,"kernel32.dll", UInt)
                pVirtualAllocEx	:= DllCall("GetProcAddress", UInt,hKernel, Str,"VirtualAllocEx", UInt)

                if (pVirtualAllocEx == LVNULL) {

                    hMap := DllCall("CreateFileMapping", "UInt",INVALID_HANDLE_VALUE, "Int",LVNULL, "UInt",PAGE_READWRITE, "UInt",0, "UInt",SIZEOF_BUF, UInt)
                    if (hMap <> LVNULL)
                        lpProcessBuf := DllCall("MapViewOfFile", "UInt",hMap, "UInt",FILE_MAP_WRITE, "UInt",0, "UInt",0, "UInt",0, "UInt")

                }
                else {

                    lpProcessBuf := DllCall("VirtualAllocEx", "UInt",hProcess, "UInt",LVNULL, "UInt",SIZEOF_BUF, "UInt",MEM_COMMIT, "UInt",PAGE_READWRITE)

                }

                if (lpProcessBuf <> LVNULL)   {

                    ;var
                    VarSetCapacity(buf, SIZEOF_BUF, 0)

                    InsertInteger(LVIF_TEXT, buf, LV_ITEM_mask, SIZEOF_INT)
                    InsertInteger(iItem, buf, LV_ITEM_iItem, SIZEOF_INT)
                    InsertInteger(iSubItem, buf, LV_ITEM_iSubItem, SIZEOF_INT)
                    InsertInteger(lpProcessBuf + SIZEOF_LV_ITEM, buf, LV_ITEM_pszText, SIZEOF_POINTER)
                    InsertInteger(SIZEOF_TEXT_BUF, buf, LV_ITEM_cchTextMax, SIZEOF_INT)

                    if (DllCall("WriteProcessMemory", "UInt",hProcess, "UInt",lpProcessBuf, "UInt",&buf, "UInt",SIZEOF_BUF, "UInt",LVNULL) <> 0)
                        if (DllCall("SendMessage", "UInt",hListView, "UInt",LVM_GETITEM, "Int",0, "Int",lpProcessBuf) <> 0)
                            if (DllCall("ReadProcessMemory", "UInt",hProcess, "UInt",lpProcessBuf, "UInt",&buf, "UInt",SIZEOF_BUF, "UInt",LVNULL) <> 0)  {
                                DllCall("lstrcpyn", "Str",lpString, "UInt",&buf + SIZEOF_LV_ITEM, "Int",nMaxCount)
                                result := DllCall("lstrlen", "Str",lpString)
                            }
                }

                if (lpProcessBuf <> LVNULL)
                    if (pVirtualAllocEx <> LVNULL)
                        DllCall("VirtualFreeEx", "UInt",hProcess, "UInt",lpProcessBuf, "UInt",0, "UInt",MEM_RELEASE)
                    else
                        DllCall("UnmapViewOfFile", "UInt",lpProcessBuf)

                if (hMap <> LVNULL)
                    DllCall("CloseHandle", "UInt",hMap)

                DllCall("CloseHandle", "UInt",hProcess)
            }

        }

return result
}
;{Sub	for LV_GetItemText and LV_GetText

ExtractInteger(ByRef pSource, pOffset = 0, pIsSigned = false, pSize = 4) {

; Original versions of ExtractInteger and InsertInteger provided by Chris
; - from the AutoHotkey help file - Version 1.0.37.04

; pSource is a string (buffer) whose memory area contains a raw/binary integer at pOffset.
; The caller should pass true for pSigned to interpret the result as signed vs. unsigned.
; pSize is the size of PSource's integer in bytes (e.g. 4 bytes for a DWORD or Int).
; pSource must be ByRef to avoid corruption during the formal-to-actual copying process
; (since pSource might contain valid data beyond its first binary zero).

   SourceAddress := &pSource + pOffset  ; Get address and apply the caller's offset.
   result := 0  ; Init prior to accumulation in the loop.
   Loop % pSize { ; For each byte in the integer:
      result := result | (*SourceAddress << 8 * (A_Index - 1))  ; Build the integer from its bytes.
      SourceAddress += 1  ; Move on to the next byte.
   }
   if (!pIsSigned OR pSize > 4 OR result < 0x80000000)
      return result  ; Signed vs. unsigned doesn't matter in these cases.
   ; Otherwise, convert the value (now known to be 32-bit) to its signed counterpart:
   return -(0xFFFFFFFF - result + 1)
}

InsertInteger(pInteger, ByRef pDest, pOffset = 0, pSize = 4) {
; To preserve any existing contents in pDest, only pSize number of bytes starting at pOffset
; are altered in it. The caller must ensure that pDest has sufficient capacity.

   mask := 0xFF  ; This serves to isolate each byte, one by one.
   Loop % pSize {  ; Copy each byte in the integer into the structure as raw binary data.
      DllCall("RtlFillMemory"	, "UInt", &pDest + pOffset + A_Index - 1, "UInt", 1              	; Write one byte.
			                        	, "UChar", (pInteger & mask) >> 8 * (A_Index - 1))              	; This line is auto-merged with above at load-time.
      mask := mask << 8  ; Set it up for isolation of the next byte.
   }

}

;}

LVM_GetText(h, r, c=1) {

	;https://autohotkey.com/board/topic/41650-ahk-l-60-listview-handle-library-101/
	r -= 1                                                     	; convert to 0 based index

	VarSetCapacity(t, 511, 1)
	VarSetCapacity(lvItem, A_PtrSize * 7)
	NumPut(1 	, lvItem, "uint")                   	; mask
	NumPut(r   	, lvItem, A_PtrSize, "int")      	; iItem
	NumPut(c-1	, lvItem, A_PtrSize * 2, "int") 	; iSubItem
	NumPut(&t	, lvItem, A_PtrSize * 5, "ptr") 	; pszText
	NumPut(512	, lvItem, A_PtrSize * 6)         	; cchTextMax

	If (A_IsUnicode)
		DllCall("SendMessage", "uint", h, "uint", 4211, "uint", r, "ptr", &lvItem) ; LVM_GETITEMTEXTW
	Else
		DllCall("SendMessage", "uint", h, "uint", 4141, "uint", r, "ptr", &lvItem) ; LVM_GETITEMTEXTA

Return t
}

LVM_GetNext(hLV, rLV=0, oLV=0) {

	; hLV = ListView handle.
	; rLV = 1 based index of the starting row for the flag search. Omit or 0 to find first occurance specified flags.
	; oLV = Combination of one or more LVNI flags. See reference above.
	;LVIS_SELECTED:=	2

Return DllCall("SendMessage", "uint", hLV, "uint", 4108, "uint", rLV-1, "uint", oLV) + 1 ; LVM_GETNEXTITEM
}

LV_MouseGetCellPos(ByRef LV_CurrRow, ByRef LV_CurrCol, LV_LView) {

	/*                              	DESCRIPTION

			Link: https://autohotkey.com/board/topic/30486-listview-tooltip-on-mouse-hover/

	*/

	static LVIR_LABEL                           := 0x0002                                                                    	; LVM_GETSUBITEMRECT constant - get label info
	static LVM_GETITEMCOUNT      	:= 4100                                                                       	; gets total number of rows
	static LVM_SCROLL                    	:= 4116                                                                       	; scrolls the listview
	static LVM_GETTOPINDEX          	:= 4135                                                                       	; gets the first displayed row
	static LVM_GETCOUNTPERPAGE 	:= 4136                                                                       	; gets number of displayed rows
	static LVM_GETSUBITEMRECT    	:= 4152                                                                       	; gets cell width,height,x,y

	ControlGetPos	, LV_lx, LV_ly, LV_lw, LV_lh 			, , % "ahk_id" LV_LView                          	; get info on listview
	SendMessage	, LVM_GETITEMCOUNT		, 0, 0, , % "ahk_id" LV_LView
	LV_TotalNumOfRows		:= ErrorLevel                                                                                 	; get total number of rows
	SendMessage	, LVM_GETCOUNTPERPAGE	, 0, 0, , % "ahk_id" LV_LView
	LV_NumOfRows 			:= ErrorLevel                                                                                 	; get number of displayed rows
	SendMessage	, LVM_GETTOPINDEX			, 0, 0, , % "ahk_id" LV_LView
	LV_topIndex               	:= ErrorLevel                                                                                 	; get first displayed row

	mMode := A_CoordModeMouse
	CoordMode, MOUSE, RELATIVE
	MouseGetPos, LV_mx, LV_my
	LV_mx -= LV_lx, LV_my -= LV_ly
	VarSetCapacity(LV_XYstruct, 16, 0)                                                                                     	; create struct

	Loop,% LV_NumOfRows + 1                                                                                              	; gets the current row and cell Y,H
	{	LV_which := LV_topIndex + A_Index - 1                                                                         	; loop through each displayed row
		NumPut(LVIR_LABEL, LV_XYstruct, 0)                                                                               	; get label info constant
		NumPut(A_Index - 1, LV_XYstruct, 4)                                                                               	; subitem index
		SendMessage, LVM_GETSUBITEMRECT, %LV_which%, &LV_XYstruct,, ahk_id %LV_LView% 	; get cell coords
		LV_RowY 				:= NumGet(LV_XYstruct,4)                                                                 	; row upperleft y
		LV_RowY2 			:= NumGet(LV_XYstruct,12)                                                               	; row bottomright y2
		LV_currColHeight 	:= LV_RowY2 - LV_RowY                                                                    	; get cell height
		If(LV_my <= LV_RowY + LV_currColHeight)                                                                   	; if mouse Y pos less than row pos + height
		{	LV_currRow   := LV_which + 1                                                                                   	; 1-based current row
			LV_currRow0 := LV_which                                                                                          	; 0-based current row, if needed
			LV_currCol	:= 0                                                                                                     	; LV_currCol is not needed here, so I didn't do it! It will always be 0.
																																				; See my ListviewInCellEditing function for details on finding LV_currCol if needed.
			return LV_currRow
			Break
		}
	}
	CoordMode, MOUSE, % mMode

return
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

;\/\/ RICHEDIT \/\/
Rich_FindText(hEdit, Text, Mode:="WHOLEWORD") {

	static EM_FINDTEXT:= 1080
	Static FR:= {DOWN: 1, WHOLEWORD: 2, MATCHCASE: 4}
	Flags := 0

	For Each, Value In Mode
         If FR.HasKey(Value)
            Flags |= FR[Value]

	Sel := RE_GetSel(hEdit)
    Min := (Flags & FR.DOWN) ? Sel.E : Sel.S
	Max := (Flags & FR.DOWN) ? -1 : 0

	VarSetCapacity(FT, 16 + A_PtrSize, 0)
	NumPut(Min	  	,   FT, 0, "Int")
	NumPut(Max	  	,   FT, 4, "Int")
	NumPut(&Text	,   FT, 8, "Ptr")

	SendMessage, EM_FINDTEXT, %Flags%, &FT,, % "ahk_id " hEdit
	S := NumGet(FTX, 8 + A_PtrSize, "Int"), E := NumGet(FTX, 12 + A_PtrSize, "Int")
	If (S = -1) && (E = -1)
         Return False

	RE_SetSel(hEdit, S, E)
	RE_ScrollCaret(hEdit)

Return ErrorLevel=4294967295 ? -1 : ErrorLevel
}

RE_FindText(hEdit, sText, cpMin=0, cpMax=-1, flags="") {
	static EM_FINDTEXT=1080,WHOLEWORD=2,MATCHCASE=4		 ;WM_USER + 56
	hFlags := 0
	loop, parse, flags, %A_Tab%%A_Space%,%A_Space%%A_Tab%
		if (A_LoopField != "")
			hFlags |= %A_LOOPFIELD%
	VarSetCapacity(FT, 12)
	NumPut(cpMin,  FT, 0)
	NumPut(cpMax,  FT, 4)
	NumPut(&sText, FT, 8)
	SendMessage, EM_FINDTEXT, hFlags, &FT,, ahk_id %hEdit%
Return ErrorLevel
}

RE_GetSel(hEdit) {                                                                                         	;-- Funktionen von HiEdit.ahk - diese funktionieren mit dem RichEdit-Control in Albis
	static EM_GETSEL=176
	VarSetCapacity(s, 4), VarSetCapacity(e, 4)
	SendMessage, EM_GETSEL, &s, &e,, ahk_id %hEdit%
	s := NumGet(s), e := NumGet(e)
Return {S: s, E: e}
}

RE_GetTextLength(hEdit) {
	static WM_GETTEXTLENGTH=14
	SendMessage, WM_GETTEXTLENGTH, 0, 0,, ahk_id %hEdit%
	Return ErrorLevel
}

RE_ReplaceSel(hEdit, text=""){
	static  EM_REPLACESEL=194
	SendMessage, EM_REPLACESEL, 0, &text,, ahk_id %hEdit%
Return ErrorLevel
}

RE_ScrollCaret(hEdit){
	static EM_SCROLLCARET=183
	SendMessage, EM_SCROLLCARET, 0, 0,, ahk_id %hEdit%
	Return ErrorLevel
}

RE_SetSel(hEdit, nStart=0, nEnd=-1) {
	static EM_SETSEL=0x0B1
	SendMessage, EM_SETSEL, nStart, nEnd,, ahk_id %hEdit%
Return ErrorLevel
}



