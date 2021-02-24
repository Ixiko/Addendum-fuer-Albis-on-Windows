; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 20.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
; FENSTER                                                                                                                                                                                                                                                   	(49)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; ----- Get -----
; GetAncestor                                	GetLastActivePopup                   	GetParentList	                            	GetParent                                 	GetNextWindow
; GetWindowInfo								GetWindowSpot							GetWindow									FindChildWindow							FindWindow
; WinForms_GetClassNN               	WinForms_GetElementID
; ClientToScreen								RectOverlapsRect
; IsClosed											IsResizable									IsWindow                                    	IsWindowVisible                        	CheckWindowStatus
; WinGetTitle                                 	WinGetClass                             	WinGetText                               	WinGet                                  		WinGetMinMaxState
; WinIsBlocked

; ----- Set -----
; SetWindowPos                             	WinMoveZ									MoveWinToCenterScreen				UpdateWindow								Redraw
; SetParentByID
; WaitForNewPopUpWindow            	WaitAndActivate                        	ActivateAndWait
; FullScreenToggleUnderMouse      	AnimateWindow                        	VerifiedWindowClose

; ----- Process -----
; GetProcessName                         	GetProcessProperties                 	GetProcessNameFromID

; ----- Monitor -----
; MonitorFromWindow                   	GetMonitorInfo   				           	IsInsideVisibleArea

; ----- Helpers -----
; Win32_SendMessage

; ---- get informations ----
GetAncestor(hWnd, Flag=2) {                                                                                                    	;--
	;1 - Parent , 2 - Root
    Return DllCall("GetAncestor", "Ptr", hWnd, "UInt", Flag)
}

GetLastActivePopup(hwnd) {                                                                                                 			;-- get the last active popup window
	return DLLCall("GetLastActivePopup", "uint", hwnd)
}

GetParentList(ChildHwnd) {                                                                                                        	;-- returns a list of comma separated WinTitles and WinClasses of all parent windows

	;15.05.2019: Code shortend - using extra functions and while loop
	while, (pHwnd:= GetParent(ChildHwnd)) {
			List .= WinGetTitle(pHwnd) " - " WinGetClass(pHwnd) ","
			ChildHwnd:= pHwnd
	}

return List
}
GetParentClassList(ChildHwnd, WinHwnd) {                                                                                  	;-- returns a list of comma separated WinTitles and WinClasses of all parent windows

	;15.05.2019: Code shortend - using extra functions and while loop
	List := Control_GetClassNN(WinHwnd, ChildHwnd) "|"
	while, (pHwnd:= GetParent(ChildHwnd)) {
			List := List . Control_GetClassNN(WinHwnd, pHwnd) "|"
			ChildHwnd:= pHwnd
	}

return RTrim(List, "|")
}
_GetParentList(ChildHwnd) {                                                                                                      	;-- returns a list of comma separated WinTitles and WinClasses of all parent windows

	Loop {
			pHwnd:= GetParent(ChildHwnd)
			If !pHwnd
					break
			WinGetTitle	, WTitle		, % "ahk_id " pHwnd
			WinGetClass	, WClass	, % "ahk_id " pHwnd
			List.= WTitle . " - " . WClass "`,"
			ChildHwnd:= pHwnd
	}

return List
}

GetParent(hWnd) {                                                                                                                     	;-- ermittelt das Parent Fenster
	return GetHex(DllCall("GetParent", "Ptr", hWnd, "Ptr"))
}

GetNextWindow(hwnd, wCmd) {                                                                                                	;-- ermittelt die Fenster-Z Ordnung

/* wCMD
		GW_HWNDNEXT = 2		Returns a handle to the window below the given window.
		GW_HWNDPREV = 3		Returns a handle to the window above the given window.
*/
	return DllCall("GetNextWindow", "Ptr", hWnd, "uint", wCMD, "Ptr")
}

GetWindow(hWnd, uCmd) {
	return DllCall( "GetWindow", "Ptr", hWnd, "uint", uCmd, "Ptr")
}

GetWindowInfo(hWnd) {                                                                                                            	;-- returns an Key:Val Object with the most informations about a window
	; returns an Key:Val Object with the most informations about a window (Pos, Client Size, Style, ExStyle, Border size...)
    NumPut(VarSetCapacity(WININFO, 60, 0), WININFO)
    DllCall("GetWindowInfo", "Ptr", hWnd, "Ptr", &WININFO)
    wInfo := Object()
    wInfo.WindowX := NumGet(WININFO, 4	, "Int")
    wInfo.WindowY	:= NumGet(WININFO, 8	, "Int")
    wInfo.WindowW:= NumGet(WININFO, 12, "Int") 	- wInfo.WindowX
    wInfo.WindowH	:= NumGet(WININFO, 16, "Int") 	- wInfo.WindowY
    wInfo.ClientX		:= NumGet(WININFO, 20, "Int")
    wInfo.ClientY		:= NumGet(WININFO, 24, "Int")
    wInfo.ClientW 	:= NumGet(WININFO, 28, "Int") 	- wInfo.ClientX
    wInfo.ClientH 	:= NumGet(WININFO, 32, "Int") 	- wInfo.ClientY
    wInfo.Style   		:= NumGet(WININFO, 36, "UInt")
    wInfo.ExStyle		:= NumGet(WININFO, 40, "UInt")
    wInfo.Active 		:= NumGet(WININFO, 44, "UInt")
    wInfo.BorderW 	:= NumGet(WININFO, 48, "UInt")
    wInfo.BorderH 	:= NumGet(WININFO, 52, "UInt")
    wInfo.Atom    	:= NumGet(WININFO, 56, "UShort")
    wInfo.Version 	:= NumGet(WININFO, 58, "UShort")
    Return wInfo
}

GetWindowSpot(hWnd) {                                                                                                           	;-- like GetWindowInfo, but faster because it only returns position and sizes
    NumPut(VarSetCapacity(WININFO, 60, 0), WININFO)
    DllCall("GetWindowInfo", "Ptr", hWnd, "Ptr", &WININFO)
    wi := Object()
    wi.X    	:= NumGet(WININFO, 4	, "Int")
    wi.Y    	:= NumGet(WININFO, 8	, "Int")
    wi.W   	:= NumGet(WININFO, 12, "Int") 	- wi.X
    wi.H    	:= NumGet(WININFO, 16, "Int") 	- wi.Y
    wi.CX  	:= NumGet(WININFO, 20, "Int")
    wi.CY  	:= NumGet(WININFO, 24, "Int")
    wi.CW	:= NumGet(WININFO, 28, "Int") 	- wi.CX
    wi.CH  	:= NumGet(WININFO, 32, "Int") 	- wi.CY
	wi.S    	:= NumGet(WININFO, 36, "UInt")
    wi.ES   	:= NumGet(WININFO, 40, "UInt")
	wi.Ac  	:= NumGet(WININFO, 44, "UInt")
    wi.BW 	:= NumGet(WININFO, 48, "UInt")
    wi.BH  	:= NumGet(WININFO, 52, "UInt")
	wi.A    	:= NumGet(WININFO, 56, "UShort")
    wi.V    	:= NumGet(WININFO, 58, "UShort")
Return wi
}

FindChildWindow(Parent, Child, DetectHiddenWindow="On") {                                                  	;{-- finds childWindow Hwnds of the parent window

/*                                                                                     	READ THIS FOR MORE INFORMATIONS
                                			    	a function from AHK-Forum : https://autohotkey.com/board/topic/46786-enumchildwindows/
                                                                                      it has been modified by IXIKO on May 09, 2018

	-finds childWindow handles from a parent window by using Name and/or class or only the WinID of the parentWindow
	-it returns a comma separated list of hwnds or nothing if there's no match

	-Parent parameter is an object(). Pass the following {Key:Value} pairs like this - WinTitle: "Name of window", WinClass: "Class (NN) Name", WinID: ParentWinID
	FindChildWindow({"ID":hwnd, "exe":"albis"}, {"class":"#32770"})
*/

		detect:= A_DetectHiddenWindows
		global ChildTitle, ChildNN, ChildClass, ChildExe, active_id
		global ChildHwnds

		ChildHwnds := ""

	; build ParentWinTitle parameter from ParentObject
		If Parent.WinID
			ParentWinTitle:= "ahk_id " Parent.ID
		else
			ParentWinTitle:= Parent.Title " ahk_class " Parent.Class

		ChildTitle  	:= Child.Title
		ChildClass	:= Child.class
		ChildNN   	:= Child.classnn
		ChildExe    	:= Child.Exe

		;DetectHiddenWindows, % DetectHiddenWindow  ; Due to fast-mode, this setting will go into effect for the callback too.
		DetectHiddenWindows, Off
		WinGet, active_id, ID, % ParentWinTitle

	; For performance and memory conservation, call RegisterCallback() only once for a given callback:
		if !EnumAddress  ; Fast-mode is okay because it will be called only from this thread:
			EnumAddress := RegisterCallback("EnumChildWindow") ; , "Fast")

		result:= DllCall("EnumChildWindows", "UInt", active_id, "UInt", EnumAddress, "UInt", 0)

		DetectHiddenWindows, % detect

return RTrim(ChildHwnds, ";")
}
; sub of FindChildWindow
EnumChildWindow(hwnd, lParam) {                                                                                             	;--sub function of FindChildWindow

	global ChildHwnds
	global ChildTitle, ChildNN, ChildClass, ChildExe, active_id

	If ChildNN
		classMatched := InStr(GetClassNN(hwnd, active_id), ChildNN) 	? 1 : 0
	else if ChildClass
		classMatched := InStr(WinGetClass(hwnd), ChildClass)           	? 1 : 0
	else
		classMatched := 1

	ProcName := WinGet(hwnd, "ProcessName")

	If InStr(WinGetTitle(hwnd), ChildTitle) && classMatched && InStr(ProcName, ChildExe)
		ChildHwnds.= GetHex(hwnd) "`;"

return true  ; Tell EnumWindows() to continue until all windows have been enumerated.
}
;}

FindWindow(WTitle,WClass="",WText="",PTitle="",PClass="",HiddenWins="on",HiddenText="on") {	;-- Finds the requested window,and return it's ID

	; 0 if it wasn't found or chosen from a list
	; originally from Evan Casey Copyright (c) under MIT License.
	; changed for my purposes for Addendum for ALBIS On Windows , Ixiko on April-06-2018
	; this version searches for ParentWindows if there is no WinText to check changed on April-27-2018

	HWins := A_DetectHiddenWindows, HText := A_DetectHiddenText
	DetectHiddenWindows	, % HiddenWins
	DetectHiddenText			, % HiddenText

	If Instr(WClass, "Afx:")
		SetTitleMatchMode, RegEx
	else		{
		SetTitleMatchMode, 2
		SetTitleMatchMode, slow
	}

	if !WClass
		sSearchWindow := WTitle
	else
		sSearchWindow := WTitle " ahk_class " WClass

	WinGet, nWindowArray, List, % sSearchWindow, % WText

	;Loop for more windows - this looks for ParentWindow
	if (nWindowArray > 1)	{
		Loop % nWindowArray
			if prev := DllCall("GetWindow", "ptr", hwnd, "uint", GW_HWNDPREV:=3, "ptr")	{		;GetParentWindowID
				WinGetTitle	, ltitle	, % "ahk_id " prev
				WinGetClass	, lclass	, % "ahk_id " prev
				If (ltitel == PTitle) && (lclass == PClass)  {
					sSelectedWinID := % nWindowArray%A_Index%
					break
				}
			}
	}
	else if (nWindowArray = 1)
		sSelectedWinID := nWindowArray1
	else if (nWindowArray = 0)
		sSelectedWinID := 0

	DetectHiddenWindows	, % HWins
	DetectHiddenText			, % HTexts

return sSelectedWinID
}

WinForms_GetClassNN(WinID, fromElement, ElementName) {                                                    	;-- schaut nach, welchen ClassNN ein WinForms-Element hat

	; function by Ixiko 2018 last_change 28.01.2018
	/* Funktionsinfo: Deutsch
		;Achtung: da manchmal 2 verschiedene Elemente den gleichen WindowsForms.Namen haben können, hat die Funktion einen zusätzlichen Parameter: "fromElement"
		;die Funktion untersucht ob das gesuchte Element im ClassNN enthalten ist zb. Button in WindowsForms10.BUTTON.app.0.378734a2
		;die Groß- und Kleinschreibung ist nicht zu beachten
	*/

	/* function info: english
		;Caution: sometimes 2 and more different elements in a gui can contain the same WindowsForms.name, therefore the function has an additional parameter: "fromElement"
		;it examines whether the element specified here is contained in the ClassNN, eg. Button in WindowsForms10.BUTTON.app.0.378734a2
		;this function is: case-insensitive
	*/

	WinGet, CtrlList, ControlList, ahk_id %WinID%

	Loop, Parse, CtrlList, `n
	{
			ClassNN:= A_LoopField
			ControlGetText, Name, %ClassNN% , ahk_id %WinID%
			If Instr(Name, ElementName) and Instr(ClassNN, fromElement)
                                                    return classnn
	}

}

WinForms_GetElementID(WinID, ElementName) {  	                                                                   	;-- schaut nach welche ID das gesuchte Element hat

	;-- schaut nach welche ID das gesuchte Element hat - z.B. Name eines Buttons kann hier eingetragen werden

	WinGet, CtrlListHwnd, ControlListHwnd, ahk_id %WinID%
	Loop, Parse, CtrlListHwnd, `n
	{
			buttonID:= A_LoopField
			ControlGetText, Name, ahk_id %buttonID% , ahk_id %WinID%
				If Instr(Name, ElementName, false)
                                	break
		}

return buttonID
}

ClientToScreen(hwnd, ByRef x, ByRef y) {
    VarSetCapacity(pt, 8)
    NumPut(x, pt, 0)
    NumPut(y, pt, 4)
    DllCall("ClientToScreen", "uint", hwnd, "uint", &pt)
    x := NumGet(pt, 0, "int")
    y := NumGet(pt, 4, "int")
}

RectOverlapsRect(vX1, vY1, vW1, vH1, vX2, vY2, vW2, vH2, vOpt="") {                                          	;-- check if rectangles (windows) overlap

		/*    	DESCRIPTION

				Link:             https://autohotkey.com/boards/viewtopic.php?t=30809
				Author:        	jeeswg 20 Apr 2017, 11:43

		*/

	;if (A is left of B) [rightmost A is left of leftmost B]
	;|| (A is right of B) [leftmost A is right of righttmost B]
	;|| (A is is above B) [bottom of A is is above top of B]
	;|| (A is is below B) [top of A is is below bottom of B]
	;then A and B do not overlap

	if InStr(vOpt, "x") ;reduce width and height values by 1
		vW1 -= 1, vH1 -= 1, vW2 -= 1, vH2 -= 1

	if InStr(vOpt, "e") ;(A = B) A equals B
		if !(vX1 = vX2) || !(vY1 = vY2) || !(vW1 = vW2) || !(vH1 = vH2)
			return 0
		else
			return 1

	if InStr(vOpt, "r") ;coordinates are XYRB 'RECT style'
		vR1 := vW1, vB1 := vH1, vR2 := vW2, vB2 := vH2
	else
		vR1 := vX1+vW1, vB1 := vY1+vH1, vR2 := vX2+vW2, vB2 := vY2+vH2

	if InStr(vOpt, "c") ;(A contains B) A contains all of or is equal to B
		if (vX2 < vX1) || (vY2 < vY1) || (vR2 > vR1) || (vH2 > vH1)
			return 0
		else
			return 1

	if InStr(vOpt, "w") ;(A within B) A is within or is equal to B
		if (vX1 < vX2) || (vY1 < vY2) || (vR1 > vR2) || (vH1 > vH2)
			return 0
		else
			return 1

	;(A overlaps B)
	if (vR1 < vX2) || (vX1 > vR2) || (vB1 < vY2) || (vY1 > vB2)
		return 0
	else
		return 1
}

IsClosed(win, wait) {                                                                                                                  	;-- waits until the specific window is closed
	WinWaitClose, ahk_id %win%,, %wait%
	return ((ErrorLevel = 1) ? False : True)
}

IsResizable() {                                                                                                                            	;-- feststellen ob das untersuchte Fenster in der Größe änderbar ist
    WinGet, Style, Style
    return (Style & 0x40000) ; WS_THICKFRAME
}

IsWindow(hWnd) {                                                                                                                    		;-- wrapper for IsWindow DllCall
    Return DllCall("IsWindow", "Ptr", hWnd)
}

IsWindowVisible(hWnd) {                                                                                                            	;-- ist Fenster sichtbar
	return DllCall("IsWindowVisible","Ptr", hWnd)
}

CheckWindowStatus(hwnd, timeout:=100) {                                                               					;-- check's if a window is responding or not responding (hung or crashed
	NR_temp := 0 ; init
return DllCall("SendMessageTimeout", "UInt", hwnd, "UInt", 0x0000, "Int", 0, "Int", 0, "UInt", 0x0002, "UInt", TimeOut, "UInt *", NR_temp)
}

WinGetMinMaxState(hwnd) {                                                                                                			;-- get state if window ist maximized or minimized

	; this function is from AHK-Forum: https://autohotkey.com/board/topic/13020-how-to-maximize-a-childdocument-window/
	; it returns z for maximized("zoomed") or i for minimized("iconic")
	; it's also work on MDI Windows - use hwnd you can get from FindChildWindow()

	zoomed:= DllCall("IsZoomed", "UInt", hwnd)		; Check if maximized
	iconic	:= DllCall("IsIconic"	, "UInt", hwnd)		; Check if minimized

return (zoomed>iconic) ? "z":"i"
}

WinGetTitle(hwnd) {                                                                                                                   	;-- schnellere Fensterfunktion
	;if (hwnd is not Integer)
	;		hwnd :=GetDec(hwnd)
	vChars := DllCall("user32\GetWindowTextLengthW", "Ptr", hWnd) + 1
	VarSetCapacity(sClass, vChars << !!A_IsUnicode, 0)
	DllCall("user32\GetWindowTextW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass) + 1)
	wtitle := sClass, sClass := ""
Return wtitle
}

WinGetClass(hwnd) {                                                                                                                	;-- schnellere Fensterfunktion
	;if (hwnd is not Integer)
	;		hwnd :=GetDec(hwnd)
	VarSetCapacity(sClass, 80, 0)
	DllCall("GetClassNameW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass)+1)
	wclass := sClass
	sClass =
Return wclass
}

WinGetText(hwnd) {                                                                                                                  	;-- Wrapper
	WinGetText, wtext, % "ahk_id " hwnd
Return wtext
}

WinGet(hwnd, cmd) {                                                                                                                	;-- Wrapper
	WinGet, res, % cmd, % "ahk_id " hwnd
return res
}

WinIsBlocked(hwnd) {                                                                                                                 	;-- Fenster ist blockiert
 ; WS_DISABLED:= 0x8000000
	WinGet, Style, Style, % "ahk_id " hwnd
return (Style & 0x8000000) ? true : false
}

; set
AnimateWindow(hWnd, Duration, Flag) {                                                                                     	;-- DllCall Wrapper für Windows interne Fensteranimation

;{ example DllCall("AnimateWindow,

;	    BLEND					:= 0x00080000	; Uses a fade effect.
;	    SLIDE					:= 0x00040000	; Uses slide animation. By default, roll animation is used.
;	    CENTER				:= 0x00000010	; Animate collapse/expand to/from middle.
;	    HIDE                   	:= 0x00010000	; Hides the window. By default, the window is shown.
;	    ACTIVATE				:= 0x00020000	; Activates the window. Do not use this value with AW_HIDE.
;	    HOR_POSITIVE		:= 0x00000001	; Animates the window from left to right.
;	    HOR_NEGATIVE	:= 0x00000002	; Animates the window from right to left.
;	    VER_POSITIVE		:= 0x00000004	; Animates the window from top to bottom.
;	    VER_NEGATIVE		:= 0x00000008	; Animates the window from bottom to top.

;}

	Return DllCall("AnimateWindow", "UInt",hWnd, "Int",Duration, "UInt",Flag)
}

FullScreenToggleUnderMouse(WT) {                                                                                            	;-- for a pseudo fullscreen of a window

		DetectHiddenWindows, On
		MouseGetPos,,,WinUnderMouse
		WinGetTitle, WTm, %WinUnderMouse%
		WinSet, Style, ^0xC00000, ahk_id %WinUnderMouse%
		WinSet, AlwaysOnTop, Toggle, ahk_id %WinUnderMouse%
		PostMessage, 0x112, 0xF030,,, ahk_id %WinUnderMouse% ;WinMaximize
		;PostMessage, 0x112, 0xF120,,, Fenstertitel, Fenstertext 	;WinRestore
		WinGet, Style, Style, ahk_class Shell_TrayWnd
			If (Style & 0x10000000) {
				  WinShow ahk_class Shell_TrayWnd
				  WinShow Start ahk_class Button

			} Else {
				WinHide ahk_class Shell_TrayWnd
				  WinHide Start ahk_class Button
			}

}

SetParentByID(ParentID, ChildID) {                                                                                                 ;-- title text is the start of the title of the window, gui number is e.g. 99
    Return DllCall("SetParent", "uint", childID, "uint", ParentID) ; success = handle to previous parent, failure =null
}

SetWindowPos(hWnd, x, y, w, h, hWndInsertAfter := 0, uFlags := 0x40) {                                		;--works better than the internal command WinMove - why?

	/*  ; https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-setwindowpos

	SWP_ASYNCWINDOWPOS	:= 0x4000	; This prevents the calling thread from blocking its execution while other threads process the request.
	SWP_DEFERERASE                	:= 0x2000	; Prevents generation of the WM_SYNCPAINT message.
	SWP_DRAWFRAME            	:= 0x0020	; Draws a frame (defined in the window's class description) around the window.
	SWP_FRAMECHANGED     	:= 0x0020	; Applies new frame styles set using the SetWindowLong function.
	SWP_HIDEWINDOW         	:= 0x0080	; Hides the window
	SWP_NOACTIVATE   	        	:= 0x0010	; Does not activate the window.
	SWP_NOCOPYBITS            	:= 0x0100	; Discards the entire contents of the client area.
	SWP_NOMOVE                 	:= 0x0002	; Retains the current position (ignores X and Y parameters).
	SWP_NOOWNERZORDER 	:= 0x0200	; Does not change the owner window's position in the Z order.
	SWP_NOREDRAW             	:= 0x0008	; Does not redraw changes.
	SWP_NOREPOSITION        	:= 0x0200	; Same as the SWP_NOOWNERZORDER flag.
	SWP_NOSENDCHANGING	:= 0x0400	; Prevents the window from receiving the WM_WINDOWPOSCHANGING message.
	SWP_NOSIZE                       	:= 0x0001	; Retains the current size (ignores the cx and cy parameters).
	SWP_NOZORDER              	:= 0x0004	; Retains the current Z order (ignores the hWndInsertAfter parameter).
	SWP_SHOWWINDOW        	:= 0x0040	; Displays the window.

	 */

Return DllCall("SetWindowPos", "Ptr", hWnd, "Ptr", hWndInsertAfter, "Int", x, "Int", y, "Int", w, "Int", h, "UInt", uFlags)
}

WinMoveZ(hWnd, C, X, Y, W, H, Redraw:=0) {                                                                           	;-- WinMoveZ v0.5 by SKAN on D35V/D361 @ tiny.cc/winmovez
Local V:=VarSetCapacity(R,48,0), A:=&R+16, S:=&R+24, E:=&R, NR:=&R+32, TPM_WORKAREA:=0x10000
  C:=( C:=Abs(C) ) ? DllCall("SetRect", "Ptr",&R, "Int",X-C, "Int",Y-C, "Int",X+C, "Int",Y+C) : 0
  DllCall("SetRect", "Ptr",&R+16, "Int",X, "Int",Y, "Int",W, "Int",H)
  DllCall("CalculatePopupWindowPosition", "Ptr",A, "Ptr",S, "UInt",TPM_WORKAREA, "Ptr",E, "Ptr",NR)
  X:=NumGet(NR+0,"Int"),  Y:=NumGet(NR+4,"Int")
Return DllCall("MoveWindow", "Ptr",hWnd, "Int",X, "Int",Y, "Int",W, "Int",H, "Int",Redraw)
}

MoveWinToCenterScreen(hWin) {                                                                                               	;-- moves a window to center of screen if its position is outside the visible screen area

	; dependencies: GetWindowSpot(), GetMonitorIndexFromWindow(), screenDims()

	w:= GetWindowSpot(hWin)

	If !IsInsideVisibleArea(w.X, w.Y, w.W, w.H)
	{
		mon	:= GetMonitorIndexFromWindow(hWin)
		scr	:= screenDims(mon)
		SetWindowPos(hWin, (scr.W//2) - (w.X//2), (scr.H//2) - (w.Y//2), w.W, w.H)
	}

return
}

ActivateAndWait(WinTitle, MaxSecondsToWait) {                                                                            	;-- activates a window and wait for activation
	If !WinActive(WinTitle)
		WinActivate, % WinTitle
           	WinWaitActive, % WinTitle, , % MaxSecondsToWait
return ErrorLevel
}

WaitAndActivate(WinTitle, WinText="", wait:= 3) {                                                             				;-- wait for a window and then activate it

		WinWait, % WinTitle,, % delay
		while !WinExist(WinTitle, WinText) {
			MsgBox, 0x40004, % A_ScriptName, % "Das Fenster " WinTitle " hat sich nicht geöffnet.`nÖffne es bitte manuell und drücke dann ok."
			IfMsgBox, No
				return 0
			sleep, 200
		}

		WinActivate    	, % WinTitle
		WinWaitActive	, % WinTitle,, % delay

return ErrorLevel
}

WaitForNewPopUpWindow(ParentWinID, LastWinID, WaitTime, RequestedTitle:="") {   					;-- waits for a new PopUpWindow

		; for a given parent window and returns Title, Class, Text and Hwnd

	/*		DESCRIPTION FOR WAITFORNEWPOPUPWINDOW        **last change: 22.06.2019**

			This function waits for a new PopUpWindow for a given parent window ID and returns Title, Class, Text and Hwnd of the new window.
			It's been written to handle problems when you can't know what window will open after your call. For example: you've been sending some text to a search dialog and there 2 possible windows.
			One that display the result and the other one shows you a notify message that your search fails.

			It's important to insert the following before you call a command or function to open a new window, ParentWinID - must be the hwnd of the parent window
			LastPopUpWin:= DLLCall("GetLastActivePopup", "uint", ParentWinID)
			ParentWinID is converted to integer format so it can be used with DLLCall

			Wait - is time in seconds to wait for a popup

	*/

		ParentWinID :=  GetHex(ParentWinID) ;16
		StartTime		:= A_TickCount

  ; Loop that waits for a new popup window id
		Loop
		{
				sleep, 500
				PopWinID	:= GetHex(DLLCall("GetLastActivePopup", "uint", ParentWinID))

				Time			:= A_TickCount - StartTime
				if (Time >= WaitTime * 1000)
				{
                    	WinGetPos, pwinX, pwinY, , , % "ahk_id " NewPopUpWin
                    	ToolTip, % "No new PopUpWindow found at specific time `(" WaitTime "s`)", % pwinX +30, % pwinY + 30, 18
                    	SetTimer, WFNPUPWTT, -4000
                    	return {"Title":"", "Class":"", "Text":"", "Hwnd":"", "error":1}
				}

				WinGetTitle, ParentTitle, % "ahk_id " ParentWinID
				If (RequestedTitle != "") && InStr(ParentTitle, RequestedTitle )
                    	return {"Title":"", "Class":"", "Text":"", "Hwnd":"", "error":0}

				If (PopWinID = 0) || (PopWinID = ParentWinID)
                    	continue

		} until (PopWinID <> LastWinID)

		WinGetTitle	, WPop_title    	, % "ahk_id " PopWinID
		WinGetClass	, WPop_Class	, % "ahk_id " PopWinID
		WinGetText	, WPop_Text    	, % "ahk_id " PopWinID

return {"Title":WPop_title, "Class":WPop_Class, "Text":WPop_Text, "Hwnd":PopWinID, "error":0}

WFNPUPWTT:
	ToolTip,,,, 18
return
}

Redraw( hwnd ) {                                                                                                                       	;-- redraw's a window
   static RDW_ALLCHILDREN	:= 0x80, RDW_ERASE    	:= 0x4  , RDW_ERASENOW	:=0x200, RDW_FRAME                 	:=0x400	, RDW_INTERNALPAINT	:=0x2   	, RDW_INVALIDATE	:=0x1
   static RDW_NOCHILDREN	:= 0x40, RDW_NOERASE 	:= 0x20, RDW_NOFRAME 	:=0x800, RDW_NOINTERNALPAINT	:=0x10 	, RDW_UPDATENOW	:=0x100	, RDW_VALIDATE  	:=0x8

   style := RDW_INVALIDATE | RDW_ERASE  | RDW_FRAME | RDW_INVALIDATE | RDW_ERASENOW | RDW_UPDATENOW | RDW_ALLCHILDREN
   return DllCall("RedrawWindow", "uint", hwnd, "uint", 0, "uint", 0, "uint", style)
}

UpdateWindow(hwnd) {                                                                                                                	;-- sends WM_Paint to Update a window
return dllcall("UpdateWindow", "Ptr", hwnd)
}

VerifiedWindowClose(hwnd) {

	If WinExist("ahk_id " hwnd)	{
		WinClose, % "ahk_id " hwnd
		WinWaitClose, % "ahk_id " hwnd,, 2
		If ErrorLevel {
			SendMessage 0x112, 0xF060,,, % "ahk_id " hwnd 		; WMSysCommand + SC_Close
			WinWaitClose                         , % "ahk_id " hwnd,, 2
			If ErrorLevel {
				SendMessage 0x10, 0,,,        % "ahk_id " hwnd         ; WM_Close
				WinWaitClose                  	  , % "ahk_id " hwnd,, 2
				If ErrorLevel {
					SendMessage 0x2, 0,,,     	% "ahk_id " hwnd 	   	; WM_Destroy
					WinWaitClose              	  , % "ahk_id " hwnd,, 2
					If ErrorLevel
						Process, Close          	  , % "ahk_id " hwnd
				}
			}
		}
	}

return ErrorLevel
}


; ---- Process ----
getProcessName(PID) { 			                                                            		                 					;-- get running processes with search using comma separated list

		s := 100096  ; 100 KB will surely be HEAPS
		array := []

	;Get the handle of this script with PROCESS_QUERY_INFORMATION (0x0400)
		h := DllCall("OpenProcess", "UInt", 0x0400, "Int", false, "UInt", PID, "Ptr")
	;Open an adjustable access token with this process (TOKEN_ADJUST_PRIVILEGES = 32)
		DllCall("Advapi32.dll\OpenProcessToken", "Ptr", h, "UInt", 32, "PtrP", t)
		VarSetCapacity(ti, 16, 0)  ; structure of privileges
		NumPut(1, ti, 0, "UInt")  ; one entry in the privileges array...
	;Retrieves the locally unique identifier of the debug privilege:
		DllCall("Advapi32.dll\LookupPrivilegeValue", "Ptr", 0, "Str", "SeDebugPrivilege", "Int64P", luid)
		NumPut(luid, ti, 4, "Int64")
		NumPut(2, ti, 12, "UInt")  ; enable this privilege: SE_PRIVILEGE_ENABLED = 2
	;Update the privileges of this process with the new access token:
		r := DllCall("Advapi32.dll\AdjustTokenPrivileges", "Ptr", t, "Int", false, "Ptr", &ti, "UInt", 0, "Ptr", 0, "Ptr", 0)
		DllCall("CloseHandle", "Ptr", t)  ; close this access token handle to save memory
		DllCall("CloseHandle", "Ptr", h)  ; close this process handle to save memory
	;Increase performance by preloading the library
		hModule := DllCall("LoadLibrary", "Str", "Psapi.dll")
	;Open process with: PROCESS_VM_READ (0x0010) | PROCESS_QUERY_INFORMATION (0x0400)
	   h := DllCall("OpenProcess", "UInt", 0x0010 | 0x0400, "Int", false, "UInt", PID, "Ptr")
	   if !h
	      return 0
	   VarSetCapacity(n, s, 0)  ; a buffer that receives the base name of the module:
	   e := DllCall("Psapi.dll\GetModuleBaseName", "Ptr", h, "Ptr", 0, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
	   if !e    ; fall-back method for 64-bit processes when in 32-bit mode:
	      if e := DllCall("Psapi.dll\GetProcessImageFileName", "Ptr", h, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
	         SplitPath n, n
	   DllCall("CloseHandle", "Ptr", h)  ; close process handle to save memory
	  DllCall("FreeLibrary", "Ptr", hModule)  ; unload the library to free memory
	return n
}

GetProcessProperties(hwnd) {

	Process:= Object()
	WinGet PID, PID, % "ahk_id " hWnd
    StrQuery := "SELECT * FROM Win32_Process WHERE ProcessId=" . PID
    Enum := ComObjGet("winmgmts:").ExecQuery(StrQuery)._NewEnum
    If (Enum[Process])
        ExePath := Process.ExecutablePath


Return Process
}

GetProcessNameFromID(hwnd) {
	Process:= Object()
	Process:= GetProcessProperties(hwnd)
	return Process.Name
}

; ---- Monitor ----
MonitorFromWindow(Hwnd := 0) {
 return DllCall("User32.dll\MonitorFromWindow", "Ptr", Hwnd, "UInt", Hwnd?2:1)
}

GetMonitorInfo(hMonitor) {

    VarSetCapacity(MONITORINFOEX, 104)
    NumPut(104, &MONITORINFOEX, "UInt")
    if (!DllCall("User32.dll\GetMonitorInfoW", "Ptr", hMonitor, "UPtr", &MONITORINFOEX))
        return FALSE
    return {  L:    	     NumGet(&MONITORINFOEX+ 4      	, "Int")
		    	, T:    	     NumGet(&MONITORINFOEX+ 8      	, "Int")
		    	, R:   	     NumGet(&MONITORINFOEX+12     	, "Int")
		    	, B:    	     NumGet(&MONITORINFOEX+16     	, "Int")
		    	, WL: 	     NumGet(&MONITORINFOEX+20     	, "Int")
		    	, WT: 	     NumGet(&MONITORINFOEX+24     	, "Int")
		    	, WR:	     NumGet(&MONITORINFOEX+28     	, "Int")
				, WB: 	     NumGet(&MONITORINFOEX+32     	, "Int")
		    	, Primary:   NumGet(&MONITORINFOEX+36    	, "UInt")
		    	, Name: 	    StrGet(&MONITORINFOEX+40,64	, "UTF-16") }

}

IsInsideVisibleArea(x,y,w,h) {

  isVis:=0
  SysGet, MonitorCount, MonitorCount
  Loop, % MonitorCount   {
    SysGet, Monitor%A_Index%, MonitorWorkArea, %A_Index%
    if (x+w-10>Monitor%A_Index%Left) and (x+10<Monitor%A_Index%Right) and (y+20>Monitor%A_Index%Top) and (y+20<Monitor%A_Index%Bottom)
      isVis:=1
  }

return, IsVis
}

;helper
Win32_SendMessage(win) {                                                                                                       	;-- for closing a window via SendMessage - win is a Hwnd
	static wm_msgs := {"WM_CLOSE":0x0010, "WM_QUIT":0x0012, "WM_DESTROY":0x0002}
	for k, v in wm_msgs {
		SendMessage, %v%, 0, 0,, ahk_id %win%
		if(IsClosed(win, 1))
			break
	}
	if(IsClosed(win, 1))
		return true
	return false
}

