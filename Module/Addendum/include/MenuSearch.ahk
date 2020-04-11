MenuSearch:                                                                      ;{ this script is originally by Lexikos - thx
OnMessage(0x100, "GuiKeyDown")
OnMessage(0x6, "GuiActivate")

Gui +LastFoundExist
if WinActive()
    goto GuiEscape
Gui Destroy

Gui Font, s11
Gui Margin, 0, 0
Gui Add, Edit, x20 w500 vQuery gType
Gui Add, Text, x5 y+2 w15, 1`n2`n3`n4`n5`n6`n7`n8`n9
Gui Add, ListBox, x+0 yp-2 w500 r21 vCommand gSelect AltSubmit
Gui Add, StatusBar
Gui +ToolWindow +Resize +MinSize +MinSize200x +MaxSize +MaxSize%A_ScreenWidth%x
window := WinExist("A")

if !(cmds := MenuGetAll(window)) {
    Send {Alt}
    return
}

gosub Type
WinGetTitle title, ahk_id %window%
title := RegExReplace(title, ".* - ")
Gui Show,, Searching menus of:  %title%
GuiControl Focus, Query
return

Type:
SetTimer Refresh, -10
return

;---------------------------------------------------------------
Refresh:

		GuiControlGet Query
		r := cmds

		if (Query != "") {

			StringSplit q, Query, %A_Space%
			Loop % q0
				r := Filter(r, q%A_Index%, c)

		}

		rows := ""
		row_id := []

		Loop Parse, r, `n
		{
			RegExMatch(A_LoopField, "(\d+)`t(.*)", m)
			row_id[A_Index] := m1
			rows .= "|"  m2
		}

		GuiControl,, Command, % rows ? rows : "|"

		if (Query = "")
			c := row_id.MaxIndex()

;---------------------------------------------------------------
Select:

		GuiControlGet Command
		if !Command
			Command := 1
		Command := row_id[Command]
		SB_SetText("Total " c " results`t`tID: " Command)
		if (A_GuiEvent != "DoubleClick")
			return

;----------
Confirm:

		if !GetKeyState("Shift") {
			gosub GuiEscape
			WinActivate ahk_id %window%
		}

		DllCall("SendNotifyMessage", "ptr", window, "uint", 0x111, "ptr", Command, "ptr", 0)

return

;-------------------------------------------------------------------
GuiEscape:
Gui Destroy
cmds := r := ""
OnMessage(0x100, "")
OnMessage(0x6, "")
return

GuiSize:
GuiControl Move, Query, % "w" A_GuiWidth-20
GuiControl Move, Command, % "w" A_GuiWidth-20
return
;}

GuiActivate(wParam) {
	ListLines, Off
    if (A_Gui && wParam = 0)
        SetTimer GuiEscape, -5
}

GuiKeyDown(wParam, lParam) {
	ListLines, Off
    if !A_Gui
        return
    if (wParam = GetKeyVK("Enter"))
    {
        gosub Confirm
        return 0
    }
    if (wParam = GetKeyVK(key := "Down")
     || wParam = GetKeyVK(key := "Up"))
    {
        GuiControlGet focus, FocusV
        if (focus != "Command")
        {
            GuiControl Focus, Command
            if (key = "Up")
                Send {End}
            else
                Send {Home}
            return 0
        }
        return
    }
    if (wParam >= 49 && wParam <= 57 && !GetKeyState("Shift"))
    {
        SendMessage 0x18E,,, ListBox1
        GuiControl Choose, Command, % wParam-48 + ErrorLevel
        GuiControl Focus, Command
        gosub Select
        return 0
    }
    if (wParam = GetKeyVK(key := "PgUp")
     || wParam = GetKeyVK(key := "PgDn"))
    {
        GuiControl Focus, Command
        Send {%key%}
        return
    }
}

Filter(s, q, ByRef count) {
	ListLines, Off
    if (q = "")
    {
        StringReplace s, s, `n, `n, UseErrorLevel
        count := ErrorLevel
        return s
    }
    Li := 1
    match := ""
    result := ""
    count := 0
    while Li := RegExMatch(s, "`ami)^.*\Q" q "\E.*$", match, Li + StrLen(match))
    {
        result .= match "`n"
        count += 1
    }
    return SubStr(result, 1, -1)
}


