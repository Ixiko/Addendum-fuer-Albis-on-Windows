PraxTT(Textmsg, Params:="3 0 B") {																				;--Tooltip Ersatz im Addendum/Nutzer Design mit off-Timer Feature


	/* FUNKTIONSBESCHREIBUNG,  last change: 25.09.2019
	   *
		*	Textmsg: 	Inhalt der Nachricht
		*	Param	: 	ein mit Space zu trennender String für verschiedene Parameter
		*					Position 1:		Zeit bis zum Ausblenden des Fenster, mit 0 als Parameter bleibt die Gui bis diese per "off" beendet wird
		*					Position 2:		Zoomlevel der Gui kleiner als -1 sollte nicht versucht werden, Schrift ist nicht mehr lesbar
		*					Position 3:		zum Positionieren des Fenster
		*										B:	Bottom - unten , T: Top - oben, L: Left - links, R: Rechts - rechts, M: Middle - Mitte
												A: Attached - relativ zu einem angegeben Fenster, S: Static - wird nicht ausgeblendet
		*					Innerhalb des Textes addiert vergrößert ein #2 die Fontgröße um 2 Punkte usw. (neu seit d. 18.05.2019)
		*
		*	changes: 	01.06.2019	- Funktion gibt die Anzahl der Textzeilen der Gui zurück, damit können die Textzeilen später geändert werden ohne eine neue Gui erstellen zu müssen
		*	             	11.06.2019	- Timer wird bei erneutem Funktionsaufruf beendet um das Schließen einer neuen Gui während des Aufbaus zu verhindern
		*	             	18.06.2020	- Params kann jetzt ein Objekt sein (noch nicht vollständig umgesetzt!).
		*                                   	   Beispiel: PraxTT("Textnachricht", {timeout:3, zoom:0, position:"Bottom ahk_class OptoAppClass"} - OptoAppClass ist das Albisfenster
		*                                    	- der ToolTip wird standardmäßig innerhalb des Albisfenster eingeblendet
		*
	   *
	*/

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; COORDMODE EINSTELLUNGEN
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		CoordMode, Mouse, Screen
		CoordMode, ToolTip, Screen
		sleep, 10			;CoordMode needs a pause to update - https://www.autohotkey.com/boards/viewtopic.php?f=14&t=38467
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; VARIABLEN DEFINIEREN
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		static   Font, BoldFont, FontSize, FontColor1, FontColor2, GuiBgColor1, GuiBgColor2, MrgX_Def, MrgY_Def
				, ModulName, T, BLEND, SLIDE, CENTER, ACTIVATE, HOR_POSITIVE, HOR_NEGATIVE, VER_POSITIVE
				, VER_NEGATIVE

		global PraxTT, hPraxTT, PraxTTRunning, PraxTTDuration, PraxTTCounter1, PraxTTCounter2

		Prax	   	:= Object()
		controls	:= []
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; EINSTELLUNGEN AUS DER INI EINLESEN, PARAMETER VERARBEITEN
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	  ; einmaliges Initialisieren bestimmter Parameter
		If !PraxTTrunning {
			IniRead, Font                	, % Addendum.AddendumIni, Addendum, StandardFont     	, Futura Bk Bt
			IniRead, BoldFont      	, % Addendum.AddendumIni, Addendum, StandardBoldFont	, Futura Mk Md
			IniRead, FontSize       	, % Addendum.AddendumIni, Addendum, StandardFontSize	, 9
			IniRead, FontColor1    	, % Addendum.AddendumIni, Addendum, PraxTTFontColor1	, White
			IniRead, FontColor2    	, % Addendum.AddendumIni, Addendum, PraxTTFontColor2	, Black
			IniRead, GuiBgColor1	, % Addendum.AddendumIni, Addendum, PraxTTBgColor1 	, 202842
			IniRead, GuiBgColor2	, % Addendum.AddendumIni, Addendum, PraxTTBgColor2 	, 6d8fff
			ModulName        	:= StrReplace(A_ScriptName, "`.ahk", "")
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
			If InStr(A_ScriptName, "ScanPool")
				DefaultGui := "BO"
			else
				DefaultGui := ""
	}

	  ; die Parameter verrechnen, erster Parameter ist wird ein statisches Fenster geschlossen
		If !IsObject(Params) {
			lifetime						:= StrSplit(Params, A_Space).1
			FontSizeMultiplicator	:= StrSplit(Params, A_Space).2
			ScreenPosition			:= StrSplit(Params, A_Space).3
		} else {
			lifetime						:= Params.timeout
			FontSizeMultiplicator	:= Params.zoom
			ScreenPosition			:= Params.position
		}

	  ; Auflösung >2k dann wird der Gui-Text automatisch größer angezeigt
		If A_ScreenWidth > 1920
			FontSizeMultiplicator := FontSizeMultiplicator * 2

	  ; ZoomLevel kleiner 0 ist gesperrt
		MrgAddX	:= FontSizeMultiplicator < 0 ? 0 : (FontSizeMultiplicator//2)
		MrgX    	:= MrgX_Def + MrgAddX
		MrgY    	:= MrgY_Def
		cFontSize	:= FontSize + FontSizeMultiplicator

	  ; tFont - In- oder Dekrement für Titelfont des Fenster
		tFont:= 1.25

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; AUSBLENDEN WENN LIFETIME OFF IST, TIMER BEENDEN WENN FUNKTIONSAUFRUF VORZEITIG KOMMT
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		If InStr(lifetime, "off") {
			SetTimer, PraxTToff, -40
			return
		} else if WinExist("ahk_id " hPraxTT) {
			SetTimer, PraxTTOff, Off
			gosub PraxTTOff
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; GUI ERSTELLEN ODER INHALTE AUFFRISCHEN
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		Gui, PraxTT: New, AlwaysOnTop +ToolWindow -Caption -DPIScale +Disabled +E0x8000000 HWNDhPraxTT 	;	Border				;-MaximizeBox -MinimizeBox
		Gui, PraxTT: Color, % "c" GuiBgColor2
		Gui, PraxTT: Margin, 0, 0

	  ; ADD: Progress als andere Hintergrundfarbe für den Titel
		Gui, PraxTT: Add, Progress, % "x0 y0 h" Floor(cFontSize*tFont)+10 " Background" GuiBgColor1 " HWNDhProgress" , Addendum für AlbisOnWindows
		Control, ExStyle, -0x20000, , % "ahk_id " hProgress

	  ; ADD: Title
		Gui, PraxTT: Font				, % "q5 s" Floor(cFontSize*tFont) " c" FontColor1, %BoldFont%
		Gui, PraxTT: Add, Text		, % "x" MrgX " y0 BackgroundTrans Center HWNDhTitle" , % ModulName

		ControlGetPos,TitleX,, TitleW, TitleH,, ahk_id %hTitle%
		ControlMove,, 0, 0, % TitleW+MrgX, % TitleH+3, ahk_id %hProgress%

	  ; Add: Informationstexte
		linespace := Floor(cFontSize * 0.5)
		Gui, PraxTT: Font, % "q5 s" cFontSize " c" FontColor2, %Font%
		cwMax := TitleW + MrgX					;ermittelt den breitesten Titel
		Loop, Parse, TextMsg, `n
		{
					printline:= A_LoopField
					If (A_Index > 1)
							linespace := Floor(cFontSize * 0.2)
					else if (StrLen(printline) = 0)
							linespace := Floor(cFontSize * 0.5)

					RegExMatch(printline, "(?<=#)\d+", plusSize)
					If plusSize
						printline := StrReplace(printline, "#" plusSize, "")
					else
						plusSize := 0

					Gui, PraxTT: Font, % "q5 s" cFontSize + plusSize " c" FontColor2, %Font%
					Gui, PraxTT: Add, Text, % "x" MrgX " y+" linespace " Center BackgroundTrans", % printline

					ControlGetPos,,, cw,, % "Static" A_Index+1, ahk_id %hPraxTT%
					cwMax	:= cw > cwMax ? cw : cwMax															;wenn cw größer als cwMax denn soll cwMax=cw sein, ansonsten gilt  cwMax=cwMax
		}

		Gui, PraxTT: Add, Text, % "x" MrgX " y+" cFontSize/4 " BackgroundTrans", % " "
		Gui, PraxTT: Show, % "AutoSize Center NoActivate Hide"

	  ; ermitteln der Fenstergrößen und Zentrieren des Textes
		Prax := GetWindowSpot(hPraxTT), controls := GetControls(hPraxTT)
		Loop % controls.MaxIndex() {
				If (A_Index > 2) {
						Control, Style, +0x1,				   , % "ahk_id " controls[A_Index].Hwnd
						ControlMove,, 0,, % Prax.W	  ,, % "ahk_id " controls[A_Index].Hwnd										;w - (2*MrgX)
				}
		}

	  ; Abstand der letzten Textzeile zum unteren Fensterrand festlegen
		max := controls.MaxIndex()
		ControlGetPos,, LConY,,LConH,, % "ahk_id " controls[max].Hwnd

	  ; Positionieren der Gui oberhalb der Taskbar
		Prax	:= GetWindowSpot(hPraxTT)
		If InStr(WinGetMinMaxState(AlbisWinID()), "z") { ; Albisfenster ist maximiert
			Albis	:= GetWindowSpot(AlbisWinID())
			SetWindowPos(hPraxTT, Albis.X + Floor(Albis.CW/2) - Floor(Prax.CW/2), Albis.Y + Albis.H - Prax.H - 19, Prax.W, Prax.H) ;(LConY+cFontSize/2) )
		} else {
			WinGetPos,, stY,,, ahk_class Shell_TrayWnd
			SetWindowPos(hPraxTT, Prax.WindowX, ( stY - Prax.WindowH - 18), Prax.WindowW, Prax.WindowH) ;(LConY+cFontSize/2) )
		}


	  ; erkennt das ein Counter gewünscht ist, matched 'Warte bis zu 10s ...' , matched nicht 'Der Prozeß hat 13.66s gebraucht'
		Gui, PraxTT: Font, % "q5 s" cFontSize*0.8 " c" FontColor2, %Font%
		Gui, PraxTT: Add, Text, % "x5 y" Prax.H - 15 " vPraxTTCounter1 Center BackgroundTrans", ---
		Gui, PraxTT: Add, Text, % "x5 y" Prax.H - 15 " vPraxTTCounter2 Center BackgroundTrans", ---

		If RegExMatch(TextMsg, "i)(?<=\s)\d+(?=s\s)", time) {
			GuiControl, PraxTT:, PraxTTCounter1, % time "s"
			GuiControl, PraxTT:, PraxTTCounter2, % time "s"
		} else {
			GuiControl, PraxTT:, PraxTTCounter1, % "---"
			GuiControl, PraxTT:, PraxTTCounter2, % "---"
		}

		GuiControl, MoveDraw, %hProgress%, % "w" (Prax.CW + MrgX) " h" ( TitleH + 3 )
		GuiControl, MoveDraw, %hTitle%		, % "x" ((Prax.CW+ MrgX)//2-TitleW//2)

	  ; Einblenden mit Animation wenn keine Teamviewerverbindung besteht
		Gui, PraxTT: Show, % "AutoSize NoActivate", % "PraxTT-Info"

	  ; verhindern das die PraxTT-Gui die Default-Gui ist
		If GuiDefault
			Gui, %DefaultGui%: Default

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; TIMER ZUM ENTFERNEN DES FENSTER
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	  ; wenn lifeTime = 0 bleibt das Fenster bis es geschlossen wird mit einem leeren Argument PraxTT("","")
		If lifetime
			SetTimer, PraxTToff, % "-" (lifetime * 1000)
	;}

return controls.MaxIndex()

PraxTTOff:																																									;{
	If !WinExist("ahk_class TV_Control") and !IsRemoteSession()
				AnimateWindow(hPraxTT, PraxTTDuration, BLEND)
	Gui, PraxTT: Destroy
	If GuiDefault
			Gui, %DefaultGui%: Default
return
;}
}







