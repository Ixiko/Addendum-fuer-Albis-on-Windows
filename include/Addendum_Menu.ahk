;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 07.10.2019 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;		Menu                                                                                                                                                                                                                                                	(03)
;		~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		MenuGetAll 	                                	MenuGetAll_sub                            	CallMenuWaitForWindow
;		_________________________________________________________________________________________________________________________________________________________
MenuGetAll(hwnd) {

    if !menu := DllCall("GetMenu", "ptr", hwnd, "ptr")
        return ""
    MenuGetAll_sub(menu, "", Lcmds)

 return Lcmds
}

MenuGetAll_sub(menu, prefix, ByRef cmds) {

    Loop % DllCall("GetMenuItemCount", "ptr", menu) {

        VarSetCapacity(itemString, 2000)

        if !DllCall("GetMenuString", "ptr", menu, "int", A_Index-1, "str", itemString, "int", 1000, "uint", 0x400)
            continue

        StringReplace itemString, itemString, &
        itemID := DllCall("GetMenuItemID", "ptr", menu, "int", A_Index-1)
        if (itemID = -1)
        if subMenu := DllCall("GetSubMenu", "ptr", menu, "int", A_Index-1, "ptr") {

            MenuGetAll_sub(subMenu, prefix itemString " > ", cmds)
            continue

        }
        cmds .= itemID "`t" prefix RegExReplace(itemString, "`t.*") "`n"
    }

}

CallMenuWaitForWindow(mcommand, hwnd, WinTitle, WinClass:="", WinText:="") {	;-- Menupunkt z.B. in Albis aufrufen, versucht bis zu 2Sek. das Fenster aufzurufen

	class	:= WinClass = "" ? "" : " ahk_class " WinClass
	If !(hcalled:=WinExist(WinTitle . class, WinText) )
	{
			PostMessage, 0x111, %mcommand%,,, ahk_id %hwnd%
			WinWaitActive, % WinTitle . class, % WinText, 3		                                    	; **NEUER EINFALL!** Warte doch einfach mal auf das zu erwartende Fenster, das kann Doppelaufrufe bestimmt verhindern
			hcalled:=WinExist(WinTitle . class, WinText)
			If hcalled
				return hcalled
	}

/* vielleicht brauche ich das hier dann nicht mehr
	while !(hcalled:=WinExist(WinTitle . class, WinText) )
	{
			If (Mod(A_Index, 30) = 0) 	;alle 300ms erfolgt ein erneuter Aufruf
			{
                    PostMessage, 0x111, %mcommand%,,, ahk_id %hwnd%
                    WinWaitActive, % WinTitle . class, % WinText, 3
                    hcalled:=WinExist(WinTitle . class, WinText)
                    If hcalled
                            return hcalled
			}
			sleep, 10
			If (A_Index>200)
				return 0
	}
*/

return hcalled
}
