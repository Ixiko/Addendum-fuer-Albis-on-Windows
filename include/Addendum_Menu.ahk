;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 08.07.2020 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;		Menu                                                                                                                                                                                                                                                	(05)
;		~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		GetMenuHandle                            	GetMenuCount                                	MenuGetAll 	                                	MenuGetAll_sub                            	CallMenuWaitForWindow
;		_________________________________________________________________________________________________________________________________________________________
GetMenuHandle(hwnd) {                                                                                        	;-- ermittelt das Menuhandle vom Fensterhandle
	SendMessage, 0x01E1,,,, % "ahk_id " hWnd
return ErrorLevel
}

GetMenuCount(hMenu) {                                                                                        	;-- gibt die Anzahl der Menupunkte zurück
Return DllCall("GetMenuItemCount", Ptr,hMenu)
}

MenuGetAll(hwnd) {                                                                                              	;-- Liste von Menupunkten (kein Submenu)

    if !hmenu := DllCall("GetMenu", "ptr", hwnd, "ptr")
        return "noMenu"
    MenuGetAll_sub(hmenu, "", Lcmds)

 return Lcmds
}

MenuGetAll_sub(hmenu, prefix, ByRef cmds) {                                                         	;-- Liste von Submenupunkten

    Loop % DllCall("GetMenuItemCount", "ptr", hmenu) {

        VarSetCapacity(itemString, 2000)

        if !DllCall("GetMenuString", "ptr", hmenu, "int", A_Index-1, "str", itemString, "int", 1000, "uint", 0x400)
            continue

        StringReplace itemString, itemString, &
        itemID := DllCall("GetMenuItemID", "ptr", hmenu, "int", A_Index-1)
        if (itemID = -1)
        if hsubMenu := DllCall("GetSubMenu", "ptr", hmenu, "int", A_Index-1, "ptr") {

            MenuGetAll_sub(hsubMenu, prefix itemString " > ", cmds)
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
