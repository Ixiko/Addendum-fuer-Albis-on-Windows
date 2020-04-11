global DebugConsole:= true

DebugConsole(Text) {

	global
	if !DebugConsole
		return
	if (ahk_id hDebugControl = "") {
		Gui, DebugConsole:Font, s8, Tahoma
		Gui, DebugConsole: +LastFound +AlwaysOnTop +Resize E0x513100C4
		Gui, DebugConsole:Add, Edit, Multi Wrap hwndhDebugControl vDebugControl w600 h300
		y := A_ScreenHeight-750
		x := A_ScreenWidth-1050
		Gui, DebugConsole:Show, x%x% y%y% , DebugConsole
		hDbgConsole:= WinExist("DebugConsole")
		WinSet, AlwaysOnTop, On, ahk_id %hDbgConsole%
		Sleep, 200
	}
	Text .= "`n`n"
    SendMessage, 0x000E, 0, 0,, ahk_id %hDebugControl% ;WM_GETTEXTLENGTH
    SendMessage, 0x00B1, ErrorLevel, ErrorLevel,, ahk_id %hDebugControl% ;EM_SETSEL
    SendMessage, 0x00C2, false, &Text,, ahk_id %hDebugControl% ;EM_REPLACESEL
	return

DebugConsoleGuiSize:

	if ErrorLevel = 1  ; The window has been minimized.  No action needed.
    return
			; Otherwise, the window has been resized or maximized. Resize the Edit control to match.
	Critical
	WinGetPos, wax, way, waw, wah, ahk_id %hDbgConsole%

	hNew:= A_GuiHeight-10
	wNew:= A_GuiWidth-10

	GuiControl, Move, DebugControl, w%wNew% h%hNew%
	WinSet, ReDraw,, ahk_id %AC1h%
	Critical, Off

return


}