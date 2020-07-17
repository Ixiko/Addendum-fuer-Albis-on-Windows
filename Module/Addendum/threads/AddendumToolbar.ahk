; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . .                                                                                                                                                          	. . . . . . . . . .
; . . . . . . . . .                                                             	ADDENDUM  TOOLBAR                                                     	. . . . . . . . . .
											                			 Version:="0.75" , vom:="11.06.2020"
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . .  ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"	. . . . . . . . . .
; . . . . . . . . .         BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE         	. . . . . . . . . .
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . .        RUNS WITH AUTOHOTKEY_H AND AUTOHOTKEY_L IN 32 OR 64 BIT UNICODE VERSION         	. . . . . . . . . .
; . . . . . . . . .                            PLEASE USE ONLY THE NEWEST VERSION OF AUTOHOTKEY!!!                              	. . . . . . . . . .
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . .                         SORRY THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE                           	. . . . . . . . . .
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . .         FOR THE BEST VIEW USE ⊡ FUTURA BK BT ⊡ (SORRY I'DONT LIKE MONOSPACE FONTS)          	. . . . . . . . . .
; . . . . . . . . .                                                                                                                                                         	. . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .


; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
;   HINWEIS: dieses Skript wird von Addendum.ahk als eigener Thread gestartet! Es läßt sich aber auch allein ausführen.
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

; 	Skripteinstellungen                                                 	;{
	#SingleInstance Force
	#NoEnv
	#NoTrayIcon

	SetWorkingDir 					, % A_ScriptDir
	SetBatchLines					, -1
	CoordMode						, ToolTip, Screen

	DetectHiddenWindows   	, On
	SetTitleMatchMode        	, 2	              	;Fast is default
	SetTitleMatchMode        	, Fast        		;Fast is default

	OnError("FehlerProtokoll")                      	; Funktion findet sich in \include\Addendum_Protocol.ahk

;}

;	VARIABLEN werden festgelegt                                 	;{
	hIBitmap:= Create_Addendum_ico(true)                                                                          	; Script Icon laden
	global AlbisID       	:= WinExist("ahk_class OptoAppClass")
	global AddendumDir	:= FileOpen("C:\albiswin.loc\AddendumDir","r").Read()
	global oATb				:= Object()
	global compname 	:= StrReplace(A_ComputerName, "-")                                                	; der Name des Computer auf dem das Skript läuft
	global Auth           	:= Object()
			  Auth.Module	:= Object()
;}

;	verhindert eine mehrfache Ausführung des Skriptes 	;{
	If WinExist("Addendum AlbisToolbar ahk_class AutoHotkeyGui")
	    ExitApp
;}

;	startet die Toolbar
	Addendum_Toolbar()

return

#If ToolbarWasMoved()
#If

Addendum_Toolbar() {                                                           	; Toolbar zeichnen

		global

	; wartet bis die Albistoolbars angezeigt werden
		while !ControlGet("Hwnd",, "AfxControlBar1401", "ahk_id " AlbisWinID()) {
				If (A_Index > 100)  ; Sicherheitsabbruch des Skriptes bei fehlender Albis Toolbar nach gewisser Zeit
					ExitApp
				Sleep, 2000
		}

	; Clientindividuelle Position der Toolbar auslesen
		IniRead, TbIniPos, % AddendumDir "\Addendum.ini", % compname, % "AddendumToolbar_Position", % "x741 y0"

	; App Authorisierung aus der Ini-Datei laden
		Auth:= AddendumClientAuth()

	; App-Einstellungen für den Client laden
		TbApps:= LoadClientApps(Auth, TbButtonNames)

	; lädt die Image-Liste
		ImageList:= LoadImageList(TbApps, Auth)

	; Gui zeichnen
		Gui, ATB: -MinimizeBox -MaximizeBox -SysMenu +ToolWindow -Caption +HWNDhAdmTbGui
		Gui, ATB: Margin, 0, 0
		Gui, ATB: Add	, Picture, x0 y0 vATBMove, % "M:\Praxis\Skripte\Skripte Neu\Addendum für AlbisOnWindows\assets\ICONS\others\ToolbarMove.png"
		Gui, ATB: Show	, % TbIniPos " w300 h24", Addendum AlbisToolbar

	; Style anpassen
		WinSet, Style   	, 0x94000000, % "ahk_id " hAdmTbGui
		WinSet, ExStyle	, 0x00000084, % "ahk_id " hAdmTbGui

	; Toolbar hinzufügen
		GuiControlGet, ATBMover, ATB: Pos, ATBMove
		AdmTb                  	:= ToolbarCreate("OnToolbar", "ATB", TbButtonNames, ImageList, "AutoSize Flat List Tooltips", "", "x" (ATBMoverX + ATBMoverW + 1) " y0 h25")
		oATb.hAdmTb        	:= hAdmTbGui
		oATb.hAdmToolbar	:= AdmTb.hwnd
		ControlGetPos, tbwX,,,,, % "ahk_id " oATb.hAdmToolbar
		oATb.TbW             	:= AdmTb.Width                           	;Toolbar width only
		oATb.Y                  	:= 0
		oATb.W                	:= tbwX + AdmTb.Width + 2
		oATb.H                 	:= AdmTb.Height
		oATb.AlbisBW       	:= GetWindowSpot(AlbisWinID()).BW
		oATb.hAlbisToolbars := ControlGet("Hwnd",, "AfxControlBar1401", "ahk_id " AlbisWinID())

	; Gui Position anpassen
		WinMove, % "ahk_id " oATb.hAdmTb,,,, % (tbwX + oATb.TbW + 2)

	; wird nur fortgesetzt wenn das Albisprogrammfenster vorhanden ist (wartet 3min auf Albis)
		while, !WinExist("ahk_class OptoAppClass")		{
				Sleep, 1000
				If (A_Index > 180)
					ExitApp
		}

	; wartet bis das Albisfenster fertig erstellt worden ist (AfxControlbar1401 ist vorhanden)
		Loop {
			; wartet bis die Albistoolbar angezeigt wird
				ControlGet, AlbisOwner, HWND,, AfxControlBar1401, ahk_class OptoAppClass
				If AlbisOwner
					break
			; 500ms warten, aber unterbrechbar lassen
				Loop {
					Sleep, 50
					If (A_Index > 9)
						break
				}
		}

	; neue Toolbar mit Albis verbinden
		SetParentByID(oATb.hAlbisToolbars, oATb.hAdmTb)

	; wird benötigt um das manuelle Verschieben der Toolbar zu ermöglichen
		OnMessage(0x200, "OnMouseMove")
		OnMessage(0x201, "OnLButtonDown")

	; Änderung der Albisfensterposition erkennen, damit die Toolbar nach Minimieren von Albis neu gezeichnet wird
		DllCall("RegisterShellHookWindow", UInt, oATb.hAdmTb)
		MsgNum := DllCall("RegisterWindowMessage", Str, "SHELLHOOK")
		OnMessage(MsgNum, "RedrawToolbar")

	; Hotkey festlegen
		fnCheckOverlap := Func("CheckOverlap")
		Hotkey, If, ToolbarWasMoved()
		Hotkey, ~LButton Up, % fnCheckOverlap
		Hotkey, If

	; prüft die Position der Toolbar, das diese nicht andere Toolbars verdeckt
		;SetTimer, % fnCheckOverlap, -60
		CheckOverlap()

	; jetzt anzeigen
		Gui, ATB: Show

Return

ATBGuiEscape:
ATBGuiClose:
ExitApp
}

ToolbarWasMoved() {                                                            	; dummy Funktion
	return oATb.WasMoved
}

ToolbarShowHide(TbStatus) {                                                    	; Addendum Toolbar anzeigen oder verstecken

		Gui, ATB: Show	, % TbStatus

return
}

OnToolbar(hWnd, Event, Text, Pos, Id) {                                   	; startet Module/Programme bei Klick auf ein Icon

    If (Event != "Click") {
        Return
    }

	For i, Modul in Auth.Module
	{
			If InStr(Modul.SkriptName, Text)
				Run, % Modul.SkriptPath
	}


    If (Text == "Einstellungen") {

    }
}

OnLButtonDown(wparam, lparam, msg, hWMLbd) {                	; zum Verschieben der Toolbar
	oATb.WasMoved := true
    PostMessage, 0xA1, 2,,, % "ahk_id " oATb.hAdmTb
return
}

OnMouseMove() {	                                                                   	; aktiviert die Toolbar sobald die Maus über dieser steht, damit die ToolTips angezeigt werden

	static fnCheckMousePos:= Func("OnMouseMove")

	MouseGetPos,,, hWinMouseOver

	If (GetDec(hWinMouseOver) = GetDec(oATb.hAdmTb))
	{
			If !WinActive("ahk_id " oATb.hAdmTb)
				WinActivate, % "ahk_id " oATb.hAdmTb

			SetTimer, % fnCheckMousePos, 100
	}
	else if (GetDec(hWinMouseOver) = GetDec(AlbisWinID()))
	{
			If !WinActive("ahk_id " AlbisWinID())
				WinActivate, % "ahk_id " AlbisWinID()

			SetTimer, % fnCheckMousePos, Off

			If oATb.WasMoved
				CheckOverlap()

			oATb.WasMoved	:= false
	}

}

RedrawToolbar(lParam, wParam) {                                         	; zeichnet die Toolbar neu und speichert die Position in der Addendum.ini

	static fnCheckOverlap:= Func("CheckOverlap")

	hactiveWin:= WinExist("A")
	WinGetClass, activeClass, % "ahk_id " wParam
	;If !InStr(activeClass, "OptoAppClass")
	;	return 0

	;ToolTip, % "lParam: " lParam "`n" GetHex(wParam) "`n" activeClass

	If InStr(activeClass, "OptoAppClass") {

			If lParam in 3,5,6
			{

				; Albisfenster minimiert dann Toolbar ausblenden und umgekehrt
					aw := GetWindowSpot(AlbisWinID())
					If (aw.X < -30) && (aw.Y < -30) || !WinExist("ahk_class OptoAppClass") {
						ToolbarShowHide("Hide")
						return 0
					} else {
						ToolbarShowHide("")
					}

				; Toolbar mit dem Albisfenster verschieben und die Position wieder anpassen
					DllCall("SetWindowPos", "Ptr", oATb.hAdmTb, "Ptr", 0, "Int", oATb.X, "Int", 0, "Int", oATb.W, "Int", oATb.H, "UInt", 0x0004)
					CheckOverlap()
					If (oATb.X > 0) && (oATb.X < 1720)
						IniWrite, % "x" oATb.X " y" oATb.Y, % AddendumDir "\Addendum.ini", % compname, % "AddendumToolbar_Position"

			}

	}

	If (lParam = 2)   {                                                                     	; Albisfenster wurde geschlossen

			If !WinExist("ahk_class OptoAppClass")
				ToolbarShowHide("Hide")

	} else if (lParam = 1)   {                                                                     	; Albisfenster wurde gestartet

			If WinExist("ahk_class OptoAppClass")
				ToolbarShowHide("")

	}


return 0
}

CheckOverlap() {                                                                  	; verhindert das die Albis Toolbar Steuerelemente von der Toolbar überlagert werden

		Tbars := Array()

	; Albis ist minimiert wenn x<-20 oder y<-20 ist, dann braucht die Funktion nicht ausgeführt werden
		aw := GetWindowSpot(AlbisWinID())
		If (aw.X < -20) || (aw.Y < -20)
			return

	; akutelle Addendum Toolbar Position und Größe
		WinGetPos, tbX, tbY, tbW, tbH, % "ahk_id " oATb.hAdmTb
		tbY := 51 ; y position der ersten möglichen ToolbarWindow Reihe

	; sichtbare Albis Toolbars mit deren Positionen ermitteln und in Reihenfolge sortieren
		WinGet, childList, ControlListHwnd, % "ahk_id " oATb.hAlbisToolbars
		Loop, Parse, childList, `n
			If IsWindowVisible(A_LoopField) {

					ControlGetPos, CX, CY, CW, CH,, % "ahk_id " A_LoopField

					If (tbY >= CY) && (tbY <= CY + CH) {									; && !(A_LoopField = oATb.hAdmTb)

							If !Tbars.MaxIndex()
								Tbars.Push({"id":A_LoopField, "X":CX, "Y":CY, "W":CW, "H":CH})
							else
							{
									Loop % Tbars.MaxIndex() {
											i:= A_Index
											If (CX < Tbars[i].X) {
												Tbars.InsertAt(i, {"id":A_LoopField, "X":CX, "Y":CY, "W":CW, "H":CH})
												break
											}
									}
									If (i+1 > Tbars.MaxIndex())
										Tbars.Push({"id":A_LoopField, "X":CX, "Y":CY, "W":CW, "H":CH})
							}

					}

			}

			tbX_new := 1

	; auf Überlappungen mit vorhanden Toolbars in Albis prüfen
		Loop % Tbars.MaxIndex() {

			If Overlaps(Tbars[A_Index].X, Tbars[A_Index].Y, Tbars[A_Index].W, Tbars[A_Index].H, tbX_new, tbY, tbW, tbH)
				tbX_new := Tbars[A_Index].X + Tbars[A_Index].W + 1 ;- oATb.AlbisBW

		}

	; Fenster an die berechnete Position verschieben
		oATb.X	:= tbX_new - oATb.AlbisBW

		DllCall("SetWindowPos", "Ptr", oATb.hAdmTb, "Ptr", 0, "Int", oATb.X, "Int", oATb.Y, "Int", oATb.W, "Int", oATb.H, "UInt", 0x0004)


return
}

LoadImageList(TbApps, Auth) {                                              	; lädt die Icon's

		ImageList := IL_Create(appCount+1)
		Loop, Parse, TbApps, `,
				IL_Add(ImageList, Auth.Module[A_LoopField].IconPath)

		IL_Add(ImageList, AddendumDir "\assets\ICONS\ICONS_SMALL\Einstellungen.ico", 1)

return ImageList
}

LoadClientApps(Auth, ByRef Buttons) {                                    	; erstellt die Button Liste für ToolbarCreate

		If !IsObject(Auth)
			ExitApp

		IniRead, TbApps, % AddendumDir "\Addendum.ini", % compname, TbApps
		If InStr(TbApps, "Error") || !RegExMatch(TbApps, "\d\,*")
		{
				TbApps:= "9,1,2"
				IniWrite, % TbApps, % AddendumDir "\Addendum.ini", % compname, TbApps
		}

		Loop, Parse, TbApps, `,
			Buttons .= Auth.Module[A_LoopField].SkriptName "`n"

		Buttons .= "-`nEinstellungen"

return TbApps
}

AddendumClientAuth() {                                                           	; liest alle verfügbaren Addendum-Apps und die Authorisierung in ein Objekt

	IniRead, AuthListe, % AddendumDir "\Addendum.ini", % compname, Module
	If InStr(AuthListe, "Error")
		AuthListe := ""

	Auth.AuthListe := AuthListe

	Loop {

			IniRead, Modul, % AddendumDir "\Addendum.ini", Module, % "Modul" SubStr("0" A_Index, -1)
			If InStr(Modul, "Error")
				break
			else
			{
					Modul := StrSplit(Modul, "|")

				; relative Pfade werden als Pfade ins Addendum-Verzeichnis gewertet, alle anderen Pfade werden als absolut angesehen
				; es lassen sich somit auch andere Programme von der Addendum Toolbar starten
					If !RegExMatch(Modul[3], "^[A-Z]\:")
						Modul[3] 	:= AddendumDir "\" Modul[3]
					If !RegExMatch(Modul[4], "^[A-Z]\:")
						Modul[4]    	:= AddendumDir "\" Modul[4]

					Auth.Module.Push({"Auth": Modul[1], "SkriptName": Modul[2], "SkriptPath": Modul[3], "IconPath": Modul[4]})
					;t.= "Auth: " Modul[1] "`nSkriptName: " Modul[2] "`nSkriptPath: " Modul[3] "`nIconPath: " Modul[4] "`n--------------------------------------------------------------------------------------------`n"
			}


	}
	;MsgBox, % t

return Auth
}

SettingsGui() {


}

;{  Funktionen
ToolbarCreate(Handler, GuiName, Buttons, ImageList := "", Options := "Flat List ToolTips", Extra := "", Pos := "") {

    Static TOOLTIPS := 0x100, WRAPABLE := 0x200, FLAT := 0x800, LIST := 0x1000, TABSTOP := 0x10000,  BORDER := 0x800000, TEXTONLY := 0
    Static BOTTOM := 0x3, ADJUSTABLE := 0x20, NODIVIDER := 0x40, VERTICAL := 0x80
    Static CHECKED := 1, HIDDEN := 8, WRAP := 32, DISABLED := 0 ; States
    Static CHECK := 2, CHECKGROUP := 6, DROPDOWN := 8, AUTOSIZE := 16, NOPREFIX := 32, SHOWTEXT := 64, WHOLEDROPDOWN := 128 ; Styles

    StrReplace(Options, "SHOWTEXT", "", fShowText, 1)
    fTextOnly	:= InStr(Options, "TEXTONLY")

    Styles := 0
    Loop Parse, Options, %A_Tab%%A_Space%, %A_Tab%%A_Space% ; Parse toolbar styles
        IfEqual A_LoopField,, Continue
        Else Styles |= A_LoopField + 0 ? A_LoopField : %A_LoopField%

    If (Pos != "") {
        Styles |= 0x4C ; CCS_NORESIZE | CCS_NOPARENTALIGN | CCS_NODIVIDER
    }

    Gui, %GuiName%: Add, Custom, ClassToolbarWindow32 hWndhWnd g_ToolbarHandler -Tabstop %Pos% %Styles% %Extra%
    _ToolbarStorage(hWnd, Handler)

    TBBUTTON_Size := A_PtrSize == 8 ? 32 : 20
    Buttons := StrSplit(Buttons, "`n")
    VarSetCapacity(TBBUTTONS, TBBUTTON_Size * Buttons.Length() , 0)

    Index := 0
    Loop % Buttons.Length()
	{
        Button := StrSplit(Buttons[A_Index], ",", " `t")

        If (Button[1] == "-")
		{
            iBitmap       	:= 0
            idCommand	:= 0
            fsState       	:= 0
            fsStyle        	:= 1 ; BTNS_SEP
            iString       	:= -1
        }
		Else
		{
            Index++
            iBitmap      	:= (fTextOnly) ? -1 : (Button[2] != "" ? Button[2] - 1 : Index - 1)
            idCommand	:= (Button[5]) ? Button[5] : 10000 + Index
            fsState       	:= InStr(Button[3], "DISABLED") ? 0 : 4 ; TBSTATE_ENABLED

            Loop Parse, % Button[3], %A_Tab%%A_Space%, %A_Tab%%A_Space% ; Parse button states
                IfEqual A_LoopField,, Continue
                Else fsState |= %A_LoopField%

            fsStyle := fTextOnly || fShowText ? SHOWTEXT : 0
            Loop Parse, % Button[4], %A_Tab%%A_Space%, %A_Tab%%A_Space% ; Parse button styles
                IfEqual A_LoopField,, Continue
                Else fsStyle |= %A_LoopField%

            iString := &(ButtonText%Index% := Button[1])
        }

        Offset := (A_Index - 1) * TBBUTTON_Size
        NumPut(iBitmap    	, TBBUTTONS, Offset     	, "Int")
        NumPut(idCommand	, TBBUTTONS, Offset + 4	, "Int")
        NumPut(fsState      	, TBBUTTONS, Offset + 8	, "UChar")
        NumPut(fsStyle      	, TBBUTTONS, Offset + 9	, "UChar")
        NumPut(iString      	, TBBUTTONS, Offset + (A_PtrSize == 8 ? 24 : 16), "Ptr")
    }

    ExtendedStyle := 0x9 ; (mixed buttons, draw dropdown arrows)
    SendMessage, 0x454, 0, % ExtendedStyle	,, % "ahk_id " hWnd                                                                  	; TB_SETEXTENDEDSTYLE
    SendMessage, 0x430, 0, % ImageList     	,, % "ahk_id " hWnd                                                                	; TB_SETIMAGELIST
    SendMessage, % A_IsUnicode ? 0x444 : 0x414, % Buttons.Length(), % &TBBUTTONS,, % "ahk_id " hWnd  	; TB_ADDBUTTONS

    If (InStr(Options, "VERTICAL")) {
        VarSetCapacity(SIZE, 8, 0)
        SendMessage 0x453, 0, &SIZE,, % "ahk_id " hWnd  ; TB_GETMAXSIZE
    } Else {
        SendMessage 0x421, 0, 0,, % "ahk_id " hWnd  ; TB_AUTOSIZE
    }

	; added by Ixiko returns width and height
		VarSetCapacity(SIZE, 8, 0)
		SendMessage 0x453, 0, &SIZE,, % "ahk_id " hWnd  ; TB_GETMAXSIZE
		w:= NumGet(SIZE, 0	, "Int")
		h:= NumGet(SIZE, 4	, "Int")


Return {"hwnd": hWnd, "width": w, "height": h}
}
_ToolbarStorage(hWnd, Callback := "") {
    Static o := {}
    Return (o[hWnd] != "") ? o[hWnd] : o[hWnd] := Callback
}
_ToolbarHandler(hWnd) 	{

    Static n := {-2: "Click", -5: "RightClick", -20: "LDown", -713: "Hot", -710: "DropDown"}

    Handler	:= _ToolbarStorage(hWnd)
    Code    	:= NumGet(A_EventInfo + 0, A_PtrSize * 2, "Int")

    If (Code != -713) {
        ButtonId := NumGet(A_EventInfo + (3 * A_PtrSize))
    } Else {
        ButtonId := NumGet(A_EventInfo, A_PtrSize == 8 ? 28 : 16, "Int") ; NMTBHOTITEM idNew
    }

    SendMessage 0x419, ButtonId,,, ahk_id %hWnd% ; TB_COMMANDTOINDEX
    Pos := ErrorLevel + 1

    VarSetCapacity(Text, 128)
    SendMessage, % A_IsUnicode ? 0x44B : 0x42D, ButtonId, &Text,, ahk_id %hWnd% ; TB_GETBUTTONTEXT

    Event := (n[Code] != "") ? n[Code] : Code

    %Handler%(hWnd, Event, Text, Pos, ButtonId)
}
GetWindowSpot(hWnd)  	{                                                   	;-- like GetWindowInfo, but faster because it only returns position and sizes
    NumPut(VarSetCapacity(WINDOWINFO, 60, 0), WINDOWINFO)
    DllCall("GetWindowInfo", "Ptr", hWnd, "Ptr", &WINDOWINFO)
    wi := Object()
    wi.X   	:= NumGet(WINDOWINFO, 4	, "Int")
    wi.Y   	:= NumGet(WINDOWINFO, 8	, "Int")
    wi.W  	:= NumGet(WINDOWINFO, 12, "Int") 	- wi.X
    wi.H  	:= NumGet(WINDOWINFO, 16, "Int") 	- wi.Y
    wi.CX	:= NumGet(WINDOWINFO, 20, "Int")
    wi.CY	:= NumGet(WINDOWINFO, 24, "Int")
    wi.CW 	:= NumGet(WINDOWINFO, 28, "Int") 	- wi.CX
    wi.CH  	:= NumGet(WINDOWINFO, 32, "Int") 	- wi.CY
	wi.S   	:= NumGet(WINDOWINFO, 36, "UInt")
    wi.ES 	:= NumGet(WINDOWINFO, 40, "UInt")
	wi.Ac	:= NumGet(WINDOWINFO, 44, "UInt")
    wi.BW	:= NumGet(WINDOWINFO, 48, "UInt")
    wi.BH	:= NumGet(WINDOWINFO, 52, "UInt")
	wi.A    	:= NumGet(WINDOWINFO, 56, "UShort")
    wi.V  	:= NumGet(WINDOWINFO, 58, "UShort")
Return wi
}
IsWindowVisible(hWnd)  	{
	return DllCall("IsWindowVisible","Ptr", hWnd)
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
AlbisWinID() {                            		                                       	;-- gibt die ID des übergeordneten Albisfenster zurück
	While !(AID := WinExist("ahk_class OptoAppClass"))
	{
			sleep, 50
			if (A_Index > 40)
				break
	}
return GetHex(AID)
}
Overlaps(vX1, vY1, vW1, vH1, vX2, vY2, vW2, vH2) {

	; X Rect2 liegt rechts neben X+W Rect1
	If (vX2 > vX1+vW1)
		return 0
	; X Rect2 liegt rechts von X Rect1, aber nicht ausserhalb von W Rect1
	If (vX2 >= vX1) && (vX2 <= vX1+vW1)
		return 1
	; X Rect2 liegt links von X Rect1, aber X+W Rect2 überlappt X Rect1
	If ((vX2 < vX1) && (vX2 + vW2 > vX1))
		return 1

}
GetHex(hwnd) {
return Format("0x{:x}", hwnd)
}
GetDec(hwnd) {
return Format("{:u}", hwnd)
}
SetParentByID(ParentID, ChildID) { 																								;-- title text is the start of the title of the window, gui number is e.g. 99
    Return DllCall("SetParent", "uint", childID, "uint", ParentID) ; success = handle to previous parent, failure =null
}
ControlGet(Cmd, Value:= "", Control:= "", WinTitle:= "", WinText:= "", ExcludeTitle:= "", ExcludeText:= "") {                    	;-- ControlGet als Funktion
	ControlGet, v, % Cmd, % Value, % Control, % WinTitle, % WinText, % ExcludeTitle, % ExcludeText
Return v
}
Create_Addendum_ico(NewHandle := False) {                                                                        	;-- erstellt das Trayicon
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAD/////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////oo//oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBf/oo//oo//oo//oo/tl4VTMyJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/6noxCKBdCKBdkPi3/oo//oo//oo//oo9CKBdCKBdCKBeJVkT+oo7/oo//oo//oo//oo+kZ1VCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd7TTr/oo//oo//oo//oo//oo9CKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdDKRj/oo//oo//oo//oo//oo//oo9GKhpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBf/oo/+oY7/oo//oo//oo//oo+PWUdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBduRTP/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo/+oY7/oo//oo//oo/7oI11SDZCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdnQC/0nIn/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9aNyZCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdeOyn/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9LLh1CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdYNiX/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo+QWkiQWkiQWkiQWkiQWkiRW0n/oo//oo//oo//oo//oo+QWkiQWkiQWkiQWkiQWkiQWkj/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo+1cmC1cmC1cmC3c2H/oo//oo//oo//oo//oo+1cmC1cmC1cmC2c2H/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+0cWBCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdILBrkkX7/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo9UNCNCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBeeY1H/oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+rbFpCKBdCKBdCKBf/oo//oo//oo/+oY7/oo9CKBdCKBdGKhr/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo9NLx5CKBdCKBdlPy3/oo//oo//oo9DKRhCKBdCKBeWXk3/oo//oo//oo/7oI13SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+gZVJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhn/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdRMiH/oo//oo//oo//oo//oo9GKxlCKBdCKBdCKBdCKBdCKBdCKBdCKBeRWkn+oY7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+WXkxCKBdCKBdCKBdCKBdCKBdCKBdDKRj/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdVNCP/oo//oo//oo//oo//oo9EKhlCKBdCKBdCKBdCKBdCKBeJVkT+oY7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+MWEZCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdYNiX/oo//oo//oo//oo//oo9DKRhCKBdCKBdCKBeEUkD9oI7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/9oI2BUD9CKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdcOSf/oo//oo//oo//oo//oo9CKBdCKBdCKBf7n43/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/7oI13SjlCKBd3Sjn/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdhOyr/oo//oo//oo//oo//oo+HVEP7n4z/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo/7oIz/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPy3/oo//oo//oo//oo//oo//oo//oo/7oYx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdoQS//oo//oo//oo//oo//oo/7oYx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBduRTP/oo//oo//oo/7oIx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdxRzX5n4r7oIx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3/oo//oo//oo//oo//oo//oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo//oo//////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}
;}

#include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk


