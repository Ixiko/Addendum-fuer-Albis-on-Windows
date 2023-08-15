PraxTT(Textmsg:="", Params:="3 0 B") {						 												;--Tooltip Ersatz im Addendum/Nutzer Design mit off-Timer Feature

	/* PraxTT - ToolTip Ersatz für Addendum
	   *
		*	Textmsg: 	Inhalt der Nachricht
		*	Param	: 	ein mit Space zu trennender String für verschiedene Parameter
		*					Position 1:		Zeit bis zum Ausblenden des Fenster, mit 0 als Parameter bleibt die Gui bis diese per "off" beendet wird
		*					Position 2:		Zoomlevel der Gui kleiner als -1 sollte nicht versucht werden, Schrift ist nicht mehr lesbar
		*					Position 3:		zum Positionieren des Fenster
		*										B:	Bottom - unten , T: Top - oben, L: Left - links, R: Rechts - rechts, M: Middle - Mitte
		*   									A: Attached - relativ zu einem angegeben Fenster, S: Static - wird nicht ausgeblendet
		*
		*					Innerhalb des Textes vergrößert #2 die Schriftgröße um 2 Punkte usw.                                      	(neu seit d. 18.05.2019)
		*					ein #-2 verkleinert die Schriftgröße um 2 Punkte                                                                        	(neu seit d. 15.10.2020)
		*
		*	changes: 	01.06.2019	- 	Funktion gibt die Anzahl der Textzeilen der Gui zurück, damit können die Textzeilen später geändert werden
		*  											ohne eine neue Gui erstellen zu müssen
		*	             	11.06.2019	- 	Timer wird bei erneutem Funktionsaufruf beendet um das Schließen einer neuen Gui während des Aufbaus zu verhindern
		*	             	18.06.2020	- 	Params kann jetzt ein Objekt sein (noch nicht vollständig umgesetzt!).
		*                                   	   	Beispiel: PraxTT("Textnachricht", {timeout:3, zoom:0, position:"Bottom",
		*											parent:"ahk_class OptoAppClass"} - OptoAppClass ist das Albisfenster
		*                                    	- 	der ToolTip wird standardmäßig innerhalb des Albisfenster eingeblendet
		*					15.02.2021	- 	Performanceverbesserungen, Code gekürzt
		*					15.06.2021	- 	Performanceverbesserungen, Code gekürzt, mehr Kommentare
		*					07.11.2021	- 	Text der in geschweiften Klammern {} gepackt wurde, wird für den Titel verwendet z.B. "{WARNUNG: Einstellungen fehlen}`n"
		*                                      	   	ergibt als Titel:                 WARNUNG: Einstellungen fehlen [Addendum]
		*					13.09.2022	- 	Code effizienter gemacht
		*
	   *
	*/

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; COORDMODE EINSTELLUNGEN
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		CoordMode, Mouse	, Screen
		CoordMode, ToolTip	, Screen
		sleep 10			;CoordMode needs a pause to update - https://www.autohotkey.com/boards/viewtopic.php?f=14&t=38467
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; VARIABLEN DEFINIEREN
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		static   FontSize, FntCol1, FntCol2, BgCol1, BgCol2, MrgX_Def, MrgY_Def
				, ModulName, T, BLEND, SLIDE, CENTER, ACTIVATE, HOR_POSITIVE, HOR_NEGATIVE, VER_POSITIVE
				, VER_NEGATIVE, DefaultGui, tFont

		global PraxTT, PraxTThGui, PraxTTRunning, PraxTTDuration, PraxTTCnt1, PraxTTCnt2, PraxTTitle

		ParentID	:= 0
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; EINSTELLUNGEN AUS DER INI EINLESEN, PARAMETER VERARBEITEN
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	  ; einmaliges Initialisieren bestimmter Parameter
		If !PraxTTrunning {

			IniRead, FntCol1  	, % Addendum.Ini, Addendum, PraxTTFontColor1	, White
			IniRead, FntCol2  	, % Addendum.Ini, Addendum, PraxTTFontColor2	, Black
			IniRead, BgCol1	, % Addendum.Ini, Addendum, PraxTTBgColor1 	, 202842
			IniRead, BgCol2	, % Addendum.Ini, Addendum, PraxTTBgColor2 	, 6d8fff

			ModulName        	:= StrReplace(A_ScriptName, ".ahk", "")
			MrgX_Def            	:= 4                    	; x margin
			MrgY_Def            	:= 4                    	; y margin
			BLEND					:= 0x00080000	; Uses a fade effect.
			SLIDE					:= 0x00040000	; Uses slide animation. By default, roll animation is used.
			CENTER				:= 0x00000010	; Animate collapse/expand to/from middle.
			HIDE						:= 0x00010000	; Hides the window. By default, the window is shown.
			ACTIVATE				:= 0x00020000	; Activates the window. Do not use this value with AW_HIDE.
			HOR_POSITIVE		:= 0x00000001	; Animates the window from left to right.
			HOR_NEGATIVE	:= 0x00000002	; Animates the window from right to left.
			VER_POSITIVE		:= 0x00000004	; Animates the window from top to bottom.
			VER_NEGATIVE		:= 0x00000008	; Animates the window from bottom to top.
			PraxTTDuration  	:= 300                 ; Duration of Animation in milliseconds
			PraxTTrunning    	:= 1                		; flag
			tFont                 	:= 1.15             	; tFont - In- oder Dekrement für Titelfont des Fenster

		  ; nur für's ScanPool Skript (nicht mehr in Benutzung)
			DefaultGui :=  InStr(A_ScriptName, "ScanPool") ? "BO" : ""

		}

	  ; die Parameter verrechnen, erster Parameter ist wird ein statisches Fenster geschlossen
		If !IsObject(Params) {
			lifetime						:= StrSplit(Params, " ").1
			FontSizeMultiplicator	:= StrSplit(Params, " ").2
			ScreenPosition			:= StrSplit(Params, " ").3
		} else {
			lifetime						:= Params.timeout
			FontSizeMultiplicator	:= Params.zoom
			ScreenPosition			:= Params.position
			ParentID                	:= RegExMatch(Params.parent, "^(0x\w+|\d+)$") ? Params.parent
											: 	 RegExMatch(Params.parent, "i)^(ahk_class|ahk_exe)\s+\w+") ? WinExist(Params.parent)
											: 	 AlbisWinID()
		}

	  ; Auflösung >2k dann wird der Gui-Text automatisch größer angezeigt
		If (A_ScreenWidth > 1920)
			FontSizeMultiplicator := FontSizeMultiplicator * 2

	  ; ZoomLevel kleiner 0 ist gesperrt
		MrgAddX	:= FontSizeMultiplicator < 0 ? 0 : Floor(FontSizeMultiplicator//2)
		MrgX    	:= MrgX_Def + MrgAddX
		MrgY    	:= MrgY_Def
		cFontSize	:= Addendum.Default.FontSize + FontSizeMultiplicator

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; AUSBLENDEN WENN LIFETIME OFF IST, TIMER BEENDEN WENN FUNKTIONSAUFRUF VORZEITIG KOMMT
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		If (InStr(lifetime, "off") || StrLen(MsgText . Params) = 0) {
			gosub PraxTTOff
			return
		} else if WinExist("PraxTT-Info ahk_class AutoHotkeyGUI") {
			SetTimer, PraxTTOff, Delete
			gosub PraxTTOff
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; GUI ERSTELLEN ODER INHALTE AUFFRISCHEN
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		Gui, PraxTT: New      	, AlwaysOnTop +ToolWindow -Caption -DPIScale +Disabled +Owner +E0x8000000 HWNDPraxTThGui
		Gui, PraxTT: Color     	, % "c" BgCol2
		Gui, PraxTT: Margin  	, 0, 0

	  ; ADD: Progress als andere Hintergrundfarbe für den Titel
		Gui, PraxTT: Add       	, Progress, % "x0 y0 h" Floor(cFontSize*tFont)+10 " Background" BgCol1 " HWNDhProgress" , Addendum für AlbisOnWindows
		Control, ExStyle, -0x20000, , % "ahk_id " hProgress

	  ; ADD: Title
		Gui, PraxTT: Font   		, % "q5 s" Floor(cFontSize*tFont) " c" FntCol1, % Addendum.Default.BoldFont
		Gui, PraxTT: Add, Text	, % "x0 y0 w300 BackgroundTrans Center vPraxTTitle HWNDPraxTTHTitle" , % ModulName
		Gui, PraxTT: Font       	, % "s" cFontSize " c" FntCol2, % Addendum.Default.Font
	  ; zentrieren
		ControlGetPos	,TitleX,, TitleW, TitleH,                     	, % "ahk_id " PraxTThTitle
		ControlMove 	,, 0, 0, % TitleW+MrgX, % TitleH+3	, % "ahk_id " hProgress

	  ; ADD: Textzeilen hinzufügen
 		cwMax  	:= 0 ;TitleW + (2*MrgX)					; maximal mögliche Breite des Gui
		For LNr, printline in StrSplit(TextMsg, "`n", "`r") {

		  ; Zeilenabstand - leere Zeilen erzeugen etwas größere Zeilenabstände
			linespace	:= Floor(cFontSize * ((A_Index > 1 && StrLen(printline) = 0) ? 0.5 : 0.2))

		  ; Titleergänzung
			If RegExMatch(printline, "\{(?<Add>.*?)\}", Title) {
				GuiControl, PraxTT:, PraxTTitle, % TitleAdd " [" ModulName "]"
				continue
			}

		  ; z.B. #2 erhöht die Schriftgröße um X-Punkte für den danach folgenden Text der gesamten Zeile
			If RegExMatch(printline, "(?<=#)\-*\d+", plusSize)
				printline := StrReplace(printline, "#" plusSize, "")
			else
				plusSize := 0

		  ; Zeile als Static-Steuerelement hinzufügen
			Gui, PraxTT: Font       	, % "s" (cFontSize + plusSize) " c" FntCol2
			Gui, PraxTT: Add, Text	, % "x" MrgX " y+" linespace " Center BackgroundTrans hwndPraxTThStatic", % printline

		  ; wenn cw größer als cwMax denn soll cwMax=cw sein, ansonsten gilt  cwMax=cwMax
			ControlGetPos,,, cw,,, % "ahk_id " PraxTThStatic
			cwMax	:= cw > cwMax ? cw : cwMax

		}

		Gui, PraxTT: Add, Text 	, % "x" MrgX " y+" cFontSize/4 " BackgroundTrans", % " "
		Gui, PraxTT: Show     	, % "AutoSize Center NA Hide"

	  ; Zentrieren des Textes
		Prax       	:= GetWindowSpot(PraxTThGui)
		controls	:= GetControls(PraxTThGui)
		Loop % controls.MaxIndex() {
			Control, Style, +0x1,		   	,	% "ahk_id " controls[A_Index].Hwnd
			ControlMove,,,, % cwMax  ,, 	% "ahk_id " controls[A_Index].Hwnd
		}
		GuiControl, PraxTT: MoveDraw, PraxTTitle, % "x0"

	  ; Positionieren der Gui oberhalb der Taskbar
		Prax	:= GetWindowSpot(PraxTThGui)
		ov		:= ParentID ? GetWindowSpot(ParentID) : GetWindowSpot(AlbisWinID())

	; Fenster im sichtbaren Monitorbereich positionieren
		If !(ov.X < -30)
			SetWindowPos(PraxTThGui, ov.X+Floor(ov.CW/2)-Floor(Prax.CW/2), ov.Y+ov.H-Prax.H-19, Prax.W, Prax.H)
		else {
			WinGetPos,, stY,,, % "ahk_class Shell_TrayWnd"
			SetWindowPos(PraxTThGui, Prax.X, stY-Prax.H-18, Prax.W, Prax.H)
		}

	  ; erkennt das ein Counter gewünscht ist, matched zahl und ein unmittelbares 's' ...' , matched nicht 'Der Prozeß hat 13.66s gebraucht'
		Prax	:= GetWindowSpot(PraxTThGui)
		Gui, PraxTT: Font	, % "s" cFontSize*0.8 " c" FntCol2, % Addendum.Default.Font
		Gui, PraxTT: Add	, Text, % "x5 y" Prax.H - 15 " vPraxTTCnt1 Center BackgroundTrans", % "---"
		Gui, PraxTT: Add	, Text, % "x5 y" Prax.H - 15 " vPraxTTCnt2 Center BackgroundTrans", % "---"

		If RegExMatch(TextMsg, "i)(?<=\s)\d+(?=\s*s\s)", time) {
			GuiControl, PraxTT:, PraxTTCnt1, % time "s"
			GuiControl, PraxTT:, PraxTTCnt2, % time "s"
		} else {
			GuiControl, PraxTT:, PraxTTCnt1, % "   "
			GuiControl, PraxTT:, PraxTTCnt2, % "   "
		}

		GuiControl, MoveDraw, % hProgress  	, % "w" (Prax.CW + MrgX) " h" ( TitleH + 3 )

	  ; Einblenden mit Animation wenn keine Teamviewerverbindung besteht
		Gui, PraxTT: Show, % "AutoSize NA", % "PraxTT-Info"

		If Addendum.PraxTTDebug
			SciTEOutput("PraxTT: " StrReplace(TextMsg, "`n", " | "))

	  ; verhindern das die PraxTT-Gui die Default-Gui ist
		If GuiDefault
			Gui, % DefaultGui ": Default"

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; SETTIMER SCHLIESST PRAXTT GUI
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	  ; wenn lifeTime = 0 bleibt die GUI solange sichtbar bis explizit PraxTT() (ohne Argumente oder PraxTT("", "0")) aufgerufen wurde
		If lifetime
			SetTimer, PraxTToff, % "-" (lifetime * 1000)
	;}

return controls.MaxIndex()

PraxTTOff:																																									;{
	SetTimer, PraxTTOff, Delete
	If !WinExist("ahk_class TV_Control") && !IsRemoteSession()
		AnimateWindow(PraxTThGui, PraxTTDuration, BLEND)
	Gui, PraxTT: Destroy
	If GuiDefault
		Gui, % DefaultGui ": Default"
return
;}
}

OSDTIP_Pop(P*) {                                                                                                     	;-- OSDTIP_Pop v0.55 by SKAN on D361/D36E @ tiny.cc/osdtip
Local
Static FN:="", ID:=0, PM:="", PS:=""

  If !IsObject(FN)
    FN := Func(A_ThisFunc).Bind(A_ThisFunc)

	If (P.Count()=0 || P[1]==A_ThisFunc) {
		OnMessage(0x202, FN, 0),  OnMessage(0x010, FN, 0)                   ; WM_LBUTTONUP, WM_CLOSE
		SetTimer, %FN%, OFF
		DllCall("AnimateWindow", "Ptr",ID, "Int",100, "Int",0x50004)          	; AW_VER_POSITIVE | AW_SLIDE
		Progress, 10:OFF                                                    ; | AW_HIDE
	Return ID:=0
  }

  MT:=P[1], ST:=P[2], TMR:=P[3], OP:=P[4], FONT:=P[5] ? P[5] : "Segoe UI"
  Title := (TMR=0 ? "0x0" : A_ScriptHwnd) . ":" . A_ThisFunc

  If (ID) {
    Progress, 10:, % (ST=PS ? "" : PS:=ST), % (MT=PM ? "" : PM:=MT), %Title%
    OnMessage(0x202, FN, TMR=0 ? 0 : -1)                                ; v0.55
    SetTimer, %FN%, % Round(TMR)<0 ? TMR : "OFF"
    Return ID
  }

  If ( InStr(OP,"U2",1) && FileExist(WAV:=A_WinDir "\Media\Windows Notify.wav") )
    DllCall("winmm\PlaySoundW", "WStr",WAV, "Ptr",0, "Int",0x220013)    ; SND_FILENAME | SND_ASYNC
                                                                        ; | SND_NODEFAULT
  DetectHiddenWindows, % ("On", DHW:=A_DetectHiddenWindows)             ; | SND_NOSTOP | SND_SYSTEM
  SetWinDelay, % (-1, SWD:=A_WinDelay)
  DllCall("uxtheme\SetThemeAppProperties", "Int",0)

  Progress, 10:C00 ZH1 FM9 FS10 CWF0F0F0 CT101010 %OP% B2 M1 HIDE,% PS:=ST, % PM:=MT, %Title%, %FONT%

  DllCall("uxtheme\SetThemeAppProperties", "Int",7)                     ; STAP_ALLOW_NONCLIENT
                                                                        ; | STAP_ALLOW_CONTROLS
  WinWait, %Title% ahk_class AutoHotkey2                                ; | STAP_ALLOW_WEBCONTENT
  WinGetPos, X, Y, W, H
  SysGet, M, MonitorWorkArea
  WinMove,% "ahk_id" . WinExist(),,% MRight-W,% MBottom-(H:=InStr(OP,"U1",1) ? H : Max(H,100)), W, H
  If ( TRN:=Round(P[6]) & 255 )
    WinSet, Transparent, %TRN%
  ControlGetPos,,,,H, msctls_progress321
  If (H>2) {
    ColorMQ:=Round(P[7]),  ColorBG:=P[8]!="" ? Round(P[8]) : 0xF0F0F0,  SpeedMQ:=Round(P[9])
    Control, ExStyle, -0x20000,        msctls_progress321               ; v0.55 WS_EX_STATICEDGE
    Control, Style, +0x8,              msctls_progress321               ; PBS_MARQUEE
    SendMessage, 0x040A, 1, %SpeedMQ%, msctls_progress321               ; PBM_SETMARQUEE
    SendMessage, 0x0409, 1, %ColorMQ%, msctls_progress321               ; PBM_SETBARCOLOR
    SendMessage, 0x2001, 1, %ColorBG%, msctls_progress321               ; PBM_SETBACKCOLOR
  }
  DllCall("AnimateWindow", "Ptr",WinExist(), "Int",100, "Int",0x40008)  ; AW_VER_NEGATIVE | AW_SLIDE
  SetWinDelay, %SWD%
  DetectHiddenWindows, %DHW%
  If (Round(TMR)<0)
    SetTimer, %FN%, %TMR%
  OnMessage(0x202, FN, TMR=0 ? 0 : -1),  OnMessage(0x010, FN)           ; WM_LBUTTONUP,  WM_CLOSE
Return ID:=WinExist()
}
OSDTIP(hWnd:="") {
Local OSDTIP
  If (hWnd="")
     Return A_ScriptHwnd . ":OSDTIP_" . "ahk_class AutoHotkey2"
  If !WinExist("ahk_id" . hWnd)
     Return 0
  WinGetTitle, OSDTIP
  OSDTIP := StrSplit(OSDTIP,":")
  If ( OSDTIP[1] = A_ScriptHwnd )
       OSDTIP[2]()

}






