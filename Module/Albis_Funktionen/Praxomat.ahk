
  A1:= "#####################################################"
  A2:= " "
  A3:= "ADDENDUM für AlbisOnWindows"
  A4:= "PRAXOMAT " . PraxoVersion:= "V0.77"
  A5:= " "
  A6:= "#####################################################"
  A7:= " "
  A8:= "Addendum für AlbisOnWindows"
  A9:= "written by Ixiko  this version is from 08.04.2019"
A10:= "please report errors to: Ixiko@mailbox.org"
A11:= "use subject: 'Addendum' so that you don't end up in the spam folder"
A12:= "GNU Lizenz can be found in main directory 2017"


;{																												TODO LISTE
;
;			1. apk nicht eintragen wenn schon bestimmte Ziffern an dem jeweiligen Tag eingetragen sind
;			7.a Inhalt der GVU Liste anzeigen
;			8. about Fenster wird nicht richtig dargestellt und kann nicht geschlossen werden
;			11. Icon ins Script einbinden
;			erledigt 05.06.2018 - List Global Vars per Hotkey und Pause,Toggle, öffnet das Gui und bei erneutem drücken wird es auch wieder geschlossen
;												und das Script läuft weiter
;			12. Timer läuft manchmal ohne Namen eines Patienten
;			13. Albisabsturz wird nicht erkannt oder es war keiner und er wird scheinbar erkannt
;			14. Praxomat hängt sich auf
;			15. TenMinutes Gui funktioniert mal wieder nicht
;
;
;			fehlerhaft --------	06.10.2018 F+ 	geriatrische Basisassement, setzt die richtige Ziffer (03360 oder 03362) in Abhängigkeit der beim letzten mal eingetragenen Ziffer,
;										die Funktion nutzt dabei eine Kürzelnotiz im Cave!von Fenster des Patienten z.B. GB 18.3B (2018 , 3.Quartal Ziffer B = 03362) V0.77
;
;}

;	   Nr		Nummerierung für schnelleres Auffinden, Reservenummern: <8,9	13,14	17,18>
;
;{	1. 	Ablaufeinstellungen des Scriptes

	#NoEnv
	#Persistent
	#SingleInstance, Force

	#InstallKeybdHook
	#MaxHotkeysPerInterval 500
	#HotkeyInterval 1000
	#MaxThreadsPerHotkey 3

	SetTitleMatchMode, 2
	DetectHiddenWindows, On
	DetectHiddenText, On
	SendMode Input					; Recommended for new scripts due to its superior speed and reliability.
	SetBatchLines -1            		; Script nicht ausbremsem (Default: "sleep, 10" alle 10 ms)

	SetControlDelay, -1         		; Wartezeit beim Zugriff auf langsame Controls abschalten ### Teste mal ohne - soll besser sein so
	SetWinDelay, -1             		; Verzögerung bei allen Fensteroperationen abschalten ### Teste mal ohne - soll besser sein so

	CoordMode, Tooltip, Screen
	CoordMode, Mouse, Screen
	CoordMode, Pixel, Screen
	CoordMode, Menu, Screen

	FileEncoding, UTF-8

	OnExit, MaxScite

;}

;{	2. 	Tray Menu
			global AddendumDir	:= FileOpen("C:\albiswin\AddendumDir-DO_NOT_DELETE","r").Read()

		  ;--Tray-Menu erstellen
			hPraxomatIcon:= Create_Praxomat_ico(False)
			Menu Tray, Icon, HICON:%hPraxomatIcon%
			Menu, Tray, NoStandard
			Menu Tray, Tip, % "      Praxomat " PraxoVersion "`n-----------------------`n         Addendum   `nfür Albis on Windows"
			Menu Tray, Add, Praxomat, NoNo
			Menu Tray, Add,
			Menu Tray, Add, Zeige HotKey's, Hotkeys
			Menu Tray, Add, Zeige Skript-Variablen, ZeigeVariablen
			Menu Tray, Add, Skript Neu Starten, SkriptReload
			Menu, Tray, Add, Beenden, MaxScite

;}

;{	3. 	Variablen Definitionen & Auslesen der Addendum.ini

			;--alle Superglobalen Variablen bis auf AddendumDir
			global AlbisWinID:= AlbisWinID()
			global ATitle, Alert:= 0
			global Fos, Fis                                                                           	; Zähler für Behandlungen am heutigen Tage und
			global PatDBCount                                                                    	; Zähler für die Anzahl der Patienten in der Patientendatenbank (Variable wird per WM_Datacopy vom Addendum.ahk-Skript abgerufen)
			global TPCount                                                                         	; Zähler für die Anzahl der Patienten im heutigen Tagesprotokoll (Variable wird per WM_Datacopy vom Addendum.ahk-Skript abgerufen)
																										     	; TPCount - geplant: zeigt die gesamte Zahl aller geöffneten Akten des Tages an (egal ob diese anwesend waren oder nicht)
			global HLPid:= 0,
			global GVUAnzahl 																		; ein Objekt das die Anzahl abgeschlossener Vorsorgeuntersuchungen enthält
			global QNow																				; das aktuelle Quartal - eingestelltes Datum des Computer
			global PtGuiX, PtGuiY, PtGuiW, PtGuiH, Mon1Right                    	; die Größenangaben des Praxomat Gui
			global PtDelta:= PtDelta1:= 19                        	                          	; Praxomat mit sich öffendem Zwischenteil der dann Zugriff auf andere Skripte ermöglicht
			global HwndPraxomat                                                                	; handle der Praxomat Gui
			global hHookedWin                                                                    	; hwnd des Hook-Fensters

			;{3.a 	Ini-Read Größe und Position der Praxomat-GUI, andere Einstellungen

			SysGet, Mon1, Monitor																												    	        			; Achtung läuft Albis bei Ihnen nicht auf Monitor1 müssen Sie hier alles anpassen
			IniRead, SGroesse				, %AddendumDir%\Addendum.ini, Praxomat, Gui_StandardFontSize, 35		        			; SGroesse = Schriftgröße der Praxomat GUI
			IniRead, xVerschiebung		, %AddendumDir%\Addendum.ini, Praxomat, Gui_xVerschiebung	 , 55					     	; xVerschiebung = delta x für die Praxomat GUI
			IniRead, PtGuiX					, %AddendumDir%\Addendum.ini, Praxomat, Gui_PosXGui				 , 1300						; PtGuiX = x Position der Praxomat GUI
			IniRead, PtGuiY					, %AddendumDir%\Addendum.ini, Praxomat, Gui_PosYGui				 , 23							; PtGuiY = y Position der Praxomat GUI
			IniRead, PtGuiW				, %AddendumDir%\Addendum.ini, Praxomat, Gui_WGui					 , 300						; PtGuiW = Breite der Praxomat GUI
			IniRead, PtGuiH					, %AddendumDir%\Addendum.ini, Praxomat, Gui_HGui					 , 74							; PtGuiH = Höhe der Praxomat GUI
			IniRead, StnFont				, %AddendumDir%\Addendum.ini, Praxomat, Gui_StandardFont		 , Futura Bk Bt		    ; StnFont = Standardfont des GUI
			IniRead, UnsichtbareFarbe	, %AddendumDir%\Addendum.ini, Praxomat, UnsichtbareFarbe		 , EEAA99 				; ausgestanzte Farbe
			IniRead, HLPColor				, %AddendumDir%\Addendum.ini, Praxomat, HLP_Color				 , 449999					; HLPColor = Hintergrundfarbe der Helper (Hilfs) GUI
			BgColor:=HLPColor

			;---------------------------- Anpassen der x&y-Position des Praxomat Gui bei Autoeinstellung ---------------- UIA_AutomationID = Item 61472 ACC_Info = Type:  Child  ▪  Id:  13  ▪  Parent child count:  15
			If InStr(PtGuiX, "Auto") {
					PtAutoXY = 1
					acc:= Object(), loc:= Object()
					acc := Acc_Get("Object", "3", 13, "ALBIS ahk_class OptoAppClass")
					loc := Acc_Location(acc, 13)
					PtGuiX:= loc.x - PtGuiW + 3, PtGuiY:= loc.y
			}
			;------------------------------------------- ScanPool BefundOrdner/PDFReader -----------------------------------
			;IniRead, BefundOrdner, %AddendumDir%\Addendum.ini, ScanPool, BefundOrdner							; BefundOrdner = Scan-Ordner für neue Befundzugänge
			;IniRead, PDFReaderName, %AddendumDir%\Addendum.ini, ScanPool, PDFReaderName			    	; der Name des verwendeten PDF Anzeigeprogramms benötigt für eine Abfrage
			;IniRead, PDFReaderFullPath, %AddendumDir%\Addendum.ini, ScanPool, PDFReaderFullPath				; der PDFReader der verwendet werden soll, Achtung: meine Automation funktioniert nur mit dem FoxitReader

			;}

			;{3.b 	Farbverlauf Praxomat-GUI

			TimeCode:=TimeCode1:=0																											; TimeCode = Minutenzähler
			FileRead, cols, %AddendumDir%\assets\colors.col
			fadeCols:=[]
			Loop, Parse, cols, `,
			{
				fadeCols[A_Index]:= A_LoopField
			}
			cols:=""																																			; leeren der Variable um RAM zu sparen

			;}

			;{3.c 	Positions/Größenabhängige Berechnungen - Praxomat GUI/Helfer Gui

			;Einstellungen für Timeroverlay
				faktor				:= Round(SGroesse*10)					    																;Grössenfaktor für die Praxomat/Helfer GUI , abhängig von der eingestellten Schriftgröße
				PosXHGui  		:= PtGuiX - faktor + 25																						;x Position der Helfer GUI errechnet sich aus
				xHelpGui			:= PtGuiX 																										;x Position der Helfer GUI
				yHelpGui			:= PtGuiY + 50																								;y Position der Helfer GUI
				SHelpGroesse	:= 20																												;Schriftgröße in der Helfer GUI
				hHelpGui			:= Round(SGroesse*10)

			;}

			;{3.d 	Ex-/Styles Albis und Praxomat GUI
				AlbisWin_Style 			:= 0x15CF8000
				AlbisWin_ExStyle 		:= 0x00000100
				PraxomatWin_Style 	:= 0x940A0000
				PraxomatWin_ExStyle	:= 0x00080008
											xe := 0
			;}

			;{3.e 	weitere Variablen - Praxomat GUI
				Werbung:= "Addendum für Albis on Windows " . PraxoVersion													; Standardtext für Infobereich im Praxomat-GUI
				tid1					:= "00:00:00"																									; Default der Praxomat GUI
				InfoPraxoText	:= Werbung
				InfoPraxoTime 	:= off
				PatName			:= ""																													; PatName = Patienten Name der gerade in Behandlung ist
				PatFound 			:= 0																													; PatFound = Status ob ein Patient in Behandlung gesetzt wurde
				notification		:= ""
			;}

			;{3.f		ToolTip-Anzeige (behandelte Patienten...)
				heutelog = %A_DD%-%A_MM%-%A_yyyy%
				IniRead, tempvar, %AddendumDir%\logs'n'data\Behandlung.log, Behandlungen , %heutelog%		;--gespeicherte Fos, Fis Variablen aus dem log bekommen
				If (Instr(tempvar,"Error")) {
					Fos	:=0																																; Fos, Fis = Zählervariable für Anzahl der Behandlungen
					Fis	:=0
				} else {
					Fos	:= StrSplit(tempvar, "|").1
					Fis	:= StrSplit(tempvar, "|").2
				}

				tmpvar			:= ""																														; Variable leeren
				Behandlung 	:= 0																														; Behandlung = Zähler für Anzahl der behandelten Patienten am heutigen Tag

			;}

			;{3.g 	Variablen für das Helfer GUI
				PHstatus		:= "out"																															; Anfangsstatus der Helfer GUI - ist "aus"
			;}

			;{3.h 	Variablen für das TENMINUTES GUI
				TMWidth 	:= 500
				TMHeight 	:= 200
				GOPz 		:= 30
				GOPh 		:= 25
				GOPw 		:= 245
				GOPx 		:= 50
				GOPy1		:= 90
				GOPy2		:= GOPy1 + 1*GOPz
				GOPy3		:= GOPy1 + 2*GOPz
				FGOPw		:= GOPh + 5
				FGOPx		:= GOPx + 122
				GOPrx		:= GOPx + 160
				GOPrw		:= GOPh - 8
				Bw			:= 200
				Bx				:= TMWidth - Bw - 50
				Bh			:= GOPh

				Ziffer1 		:= "03230"
				Ziffer2 		:= "35100"
				Ziffer3 		:= "35110"
			;}

			;{3.i		Auslesen spezifischer Einstellungen für den ausführenden Client

			CompName:= StrReplace(A_ComputerName, "-")
			IniRead, Comp, %AddendumDir%\Addendum.ini, Computer, %CompName%
			Loop, Parse, Comp, `|																												; ermitteln von User und Pass mittels Parsen
			{
				Login%A_Index% := A_LoopField
			}
			IniRead, WZStatistikFile	, %AddendumDir%\Addendum.ini, Statistiken	, WZStatistikFile
			IniRead, AlbisExeDir		, %AddendumDir%\Addendum.ini, Albis			, AlbisExeDir
			IniRead, AlbisExe			, %AddendumDir%\Addendum.ini, Albis			, AlbisExe
			IniRead, AlbisWorkDir	, %AddendumDir%\Addendum.ini, Albis			, AlbisWorkDir
			IniRead, StatDir				, %AddendumDir%\Addendum.ini, Statistiken	, StatDir
			;}

			;{3.k		sonstige Variablen

			taste					:= 1
			AlbisKeyHelpA	:= ""
			AlbisKeyHelpB	:= Der Start von allem																							; für den ersten Start brauchen die beiden Variablen unterschiedliche Inhalte
			StopStatus 		:= 0																														;Variable für Stop Status des Timer
			Pausett 				:= 0
			MPause 			:= 0																														; MPause = manuelle Pause Status (durch drücken der Pause-Taste ist der Timer anhaltbar)
			goPLoop			:= 0																														;goPatName = Zähler für Loop

			;}

			;{3.l		GVU Liste scannen und Anzahl der Patienten feststellen

			QNow				:= GetQuartal("heute")																							; QNow = derzeitiges Quartal
			GVUAnzahl		:= Object()																											; GVUAnzahl = Anzahl der Patienten in der Tagesprotokolle\0317-GVU.txt
			gvufile 				:= AddendumDir "\Tagesprotokolle\" QNow "-GVU.txt"   										; gvufile = sollte nur einmal deklariert und definiert werden
			Loop, Read, % gvufile
						GVUAnzahl[QNow]	:=A_Index
			;}

			TopToolTip()																																	; erste Info's anzeigen


;}

;{	4. 	Checkt den Albis Prozeß, startet Albis bei Bedarf, ermittelt die AlbisWinID (unter 3. muss AlbisWinID (ist ein Handle) global definiert sein)

  ;Fehlermeldungen werden in StarteAlbis() Funktion als Key, Valuewerte geschrieben um entsprechende Hinweisfenster einfacher generieren zu können
	result  := WMIEnumProcessExist("albis")
	result+= WMIEnumProcessExist("albisCS")
	If !result
		AlbisNeuStart(CompName, Login1, Login2, "Praxomat", 0)

	AlbisWinID := AlbisWinID()
;}

;{	5. 	Erstellung Praxomat GUI , vorbereitende Funktionen für die sofortige Datenanzeige im Gui (Wartezeitdateien) erstmalig auslesen

		;--holt sich hier die Daten von Botti der auf dem Server die Wartezeit berechnet
	gosub, Wartezeit

	file:= AddendumDir . "\assets\PraxomatGui300x31.png"
		;-- Zähler Startzeit - jetzt
	Startt:=A_TickCount
		;--generieren der PraxomatGui - hoffe die Umstellung auf eine Labelroutine verhindert das hierher zurück gesprungen werden kann
	gosub, PraxomatGui
		;--erstes Auffrischen der Timer GUI damit da etwas vernünftiges steht
	gosub, AktualisiereOSD
		;--AlbisFenster aktivieren sonst muss man erstmal klicken
	;WinActivate, ahk_class OptoAppClass

;}

;{	6. 	Ende des AutoExecutebereich, SetWinEventHook für Positionsänderungen des AlbisFenster, Timer setzen (Infopraxo einmalig - sonst nach Bedarf, Aktualisierungstimer für das Gui, Wartezeitdaten Timer)

	/*
  ;ein Hook nur für Albis der muss allerdings bei einem Restart von Albis eigentlich erneuert werden
	AlbisProcID 					:= AlbisPID()
	ExcludeScriptMessages 	:= 1 ; 0 to include
	ExcludeGuiEvents 			:= 1 ; 0 to include
	dwFlags 						:= ( ExcludeScriptMessages = 1 ? 0x1 : 0x0 )
	HookProcAdr 				:= RegisterCallback( "Hook_Albis", "F" )
	hWinEventHook				:= SetWinEventHook(0x800b	, 0x800b	, 0, HookProcAdr,  AlbisProcID, 0, dwFlags )
	hWinEventHook				:= SetWinEventHook(0x16		, 0x17		, 0, HookProcAdr,  AlbisProcID, 0, dwFlags )					;0x16 Event_System_MinimizeStart | 0x17 Event_System_MinimizeEnd
	hWinEventHook				:= SetWinEventHook(0xB		, 0xB			, 0, HookProcAdr,  AlbisProcID, 0, dwFlags )

 ;Registercallback für einen Aufruf einer Funktion nach Doppelklick
	HookProcAdr1 				:= RegisterCallback( "LM_DoubleClick", "F" )
	hWinEventHook1			:= SetWinEventHook(0xA3, 0xA3, 0, HookProcAdr1, 0, 0, dwFlags )

	;OnMessage(WM_NCLBUTTONDBLCLK:=0xA3, "LM_DoubleClick")
	;OnMessage(WM_LBUTTONDBLCLK:=0x203, "LM_DoubleClick")
	OnMessage(0x201, "PtWM_LBUTTONDOWN")
	OnMessage(0x200, "PtWM_MouseMove")
	*/

  ;-Skriptkommunikation
	OnMessage(0x4a, "Receive_WM_COPYDATA")

  ;-0x555 - für Kommunikation mit Addendum.ahk, Befehl|Name des abfragenden Skriptes, Name des empfangenen Skriptes
	AddendumID := GetAddendumID()
	Send_WM_COPYDATA("PatDBCount|" A_ScriptHwnd , AddendumID)
	while !InStr(Infopraxotext, "InComing message")
			Sleep, 100
	Send_WM_COPYDATA("TPCount|" A_ScriptHwnd , AddendumID)

  ;-InfoPraxo zeigt kurze Statusmeldungen zu laufenden Aktionen in der PraxomatGui an
	Infopraxotext := HwndPraxomat
	SetTimer, InfoPraxo, 6000

  ;-ich frage mich ob hier eine If Abfrage jemals wieder angesprungen wird, ansonsten ist dies der Timer um die Info`s in der PraxomatGui neu zu zeichnen
	If TimeCode < 30
		SetTimer, AktualisiereOSD, 50

  ;-Botti aktualisiert die Wartezeit alle 5 min, es reicht also den Timer für das holen der Daten ebenso auf 5min zu setzen
	SetTimer, Wartezeit, 300000

	WinActivate, ahk_id %AlbisWinID%

return
;}

;{	7. 	Hotkeybereich / HotStrings

; 7a)++ Interaktionshotkeys für das Praxomatskript	;{
LControl & ä::																							;# Praxomat beenden
		SetTimer, AktualisiereOSD, Off
		ExitApp
	return

^#ö::																							;# zeige alle Variablen
	ListVars
return

^#ü::																							;# zeige das ListLinesfenster
	ListLines
return

Pause:: gosub, PauseToggleTimer																;# Praxomat-Timer pausieren
;}

; 7b)++ Hotkey`s die nur bei aktiviertem Albisfenster funktionieren 	;{
#IfWinActive ahk_class OptoAppClass
/* abgeschaltet! 21.05.2019 - besser über Addendum
;# den aktuellen Patienten in die GVU-Liste eintragen (Datum der Programmeinstellung wird genutzt)
^F7::
	progdate=1
	gosub, GVUListe
return

;# den aktuellen Patienten in die GVU-Liste eintragen (Zeilendatum in der geöffnet Akte wird genutzt - Mauszeiger sollte über der entsprechenden Zeile stehen)
 !F7::
	progdate=0
	gosub, GVUListe
return
*/

;# jetzt eine Gesundheitsvorsorge und ein Hautkrebsscreeningformular ausfüllen
^F8::run, %AddendumDir%\Module\GVU ausfuellen.ahk

;# Praxomat-Timer Neustart. Die Zeit für den aktuellen Patienten wird gespeichert.
^F11::
	SendInput, {ESC}
	TopToolTip()
	gosub, Speichern
return

;# Praxomat-Timer Neustart und aktuellen Patienten übernehmen. Die abgelaufene Zeit wird verworfen!
^F12::
	Infopraxotext:= "Suche nach aktuellem Patient!"
	gosub, GoPatName
	;SetTimer, GoPatName, 1000
	Gosub, Erneuern
return

;# eine Ziffer auswählen (Funktion für schnellere Mehrfachauswahl)
^Up::gosub, ZifferWaehlen
;# die vorher ausgewählte Ziffer einfügen
^Left::gosub, ZifferEinfuegen

#IfWinActive

;}

; 7c)++ Hotkey der nur bei geöffnetem Wartezimmer eine Funktion hat ;{
#IfWinActive, Wartezimmer ahk_class OptoAppClass

;# wählt und startet den Timer für den ausgewählten Patienten (Wartezimmer muss aktiviert sein)
Enter::gosub, OeffnePatientenAkte

#IfWinActive
;}

; 7d)++ Hotstrings für Erleichertungen bei der Eingabe von Ziffern ;{

/*
#IfWinActive, ALBIS ahk_class OptoAppClass

::03360::
::03362::
::#gb::																									;{ überprüft welcher geriatrische Basiskomplex angesetzt werden kann

	ControlGetFocus, CFocus, ALBIS ahk_class OptoAppClass
	If InStr(CFocus, "RichEdit20A1") {

			gbStr:=""
			CZeile:= AlbisGetCaveZeile(9, 0, 0, 0)
			SendInput, {Esc}
			StrParts:= StrSplit(CZeile, "`,", A_Space)
			Loop % StrParts.MaxIndex()
			{
					If InStr(StrParts[A_Index], "GB") {
							gbStr:= StrParts[A_Index]
							break
					}
			}

				;wenn das Kürzel GB vorhanden ist, geht er durch die folgenden Zeilen
			If !(gbStr="") {

							;ermittelt aus der Eintragung das Jahr, Quartal und welche Ziffer zuletzt eingetragen wurde
						gblStat:= 	SubStr(gbStr, StrLen(gbStr), 1)				;A steht für 03360 , B für 03362
						gblQ:= 	SubStr(gbStr, StrLen(gbStr)-1, 1)				;ist das Quartal der letzten Eintragung
						gblJ:= 		SubStr(gbStr, StrLen(gbStr)-4, 2)				;ist das Jahr der letzten Eintragung

							;zum Vergleichen wird das aktuelle Quartal und Jahr ermittelt
						gbaQ:=	StrReplace(SubStr(QNow, 1, 2), "0")
						gbaJ:=		SubStr(QNow, StrLen(QNow)-1, 2)

							;berechnet den Quartalsabstand = (aktuelles Jahr - letztes eingetragenes Jahr) * 4 (mögliche Quartal) - letztes Quartal + aktuelles Quartal
						QA:=		((gbaJ - gblJ) * 4) - gblQ + gbaQ

							;A oder B für das aktuelle Quartal ?
						If (Mod(QA, 2) = 0) {						;gerade Zahl 	->  A, wenn voher A,      B, wenn vorher B
								gbaStat:= gblStat
						} else {												;ungerade Zahl ->  B, wenn vorher A,	 A, wenn vorher B
								If InStr(gblStat, "A")
										gbaStat:="B"
								else
										gbaStat:="A"
						}

			} else {
						gbAStat:="A"
			}

			gbStr:= "GB " gbaJ "`." gbaQ . gbAStat
			CZeile:=""

			Loop % StrParts.MaxIndex() - 1
			{
					CZeile.= StrParts[A_Index] . "`, "
			}

				;neue Zeile dem CaveVon Fenster wieder hinzufügen
			AlbisSetCaveZeile(9, CZeile . gbStr)

			SendInput, {Ins}
				Sleep, 100
			SendInput, lko{Tab}
				Sleep, 100

			If (gbAStat="A") {
					SendInput, 03360{Tab}
			} else {
					SendInput, 03362{Tab}
			}

	}

	return
;}



#IfWinActive
;}
*/

;}

;}

;{	8. 	Label für die Generierung der PraxomatGui

PraxomatGui:	;--Erstellen des Praxomatfensters

	If (!pToken := Gdip_Startup()) 	{
						MsgBox, 0x40048, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
						ExitApp
	}

	; Create a layered window (+E0x80000 : must be used for UpdateLayeredWindow to work!) that is always on top (+AlwaysOnTop), has no taskbar entry or caption 0x0x50020000
	Gui, 22: -Caption +ToolWindow +E0x8029C +AlwaysOnTop +HWndHwndPraxomat ;   0x960A0000 0x0008001C

	Gui, 22: Show, % "x" PtGuiX " y" PtGuiY " w" PtGuiW " h" PtGuiH Hide, PraxomatGui

	hbm1 := CreateDIBSection(PtGuiW, PtGuiH)
	hdc:= GetDC(HwndPraxomat)
	GGhdc := CreateCompatibleDC(hdc)
	obm1 := SelectObject(GGhdc, hbm1)
	pBitmap1 := Gdip_CreateBitmapFromFile(file)

	; Check to ensure we actually got a bitmap from the file, in case the file was corrupt or some other error occured
	If !pBitmap1 {
		MsgBox, 0, File loading error!, Could not load '%file%'	;0x40048
		ExitApp
	}

	Font = Futura Bk Bt
	If !Gdip_FontFamilyCreate(Font) {
					MsgBox, 48, Font error!, The font you have specified does not exist on the system
					ExitApp
	}

	G1 := Gdip_GraphicsFromHDC(GGhdc)
	Gdip_SetSmoothingMode(G1, 7)
	IWidth := Gdip_GetImageWidth(pBitmap1), IHeight := Gdip_GetImageHeight(pBitmap1)

	pBrush1 := Gdip_BrushCreateSolid("0xaa" . BgColor)
	Gdip_FillRectangle(G1, pBrush1, 0, 0, PtGuiW-3, PtGuiH - 30)
	Gdip_DeleteBrush(pBrush1)

	pBrush1 := Gdip_BrushCreateSolid("0xaa172842")
	Gdip_FillRectangle(G1, pBrush1, 0, PtGuiH -30, PtGuiW-3, PtGuiH)
	Gdip_DeleteBrush(pBrush1)

	Gdip_DrawImage(G1, pBitmap1, 0, 0, PtGuiW, PtGuiH, 0, 0, PtGuiW, PtGuiH)

	txtColor:= FadeCols[1]
	Options1:=  "x70 y23 w160 h35 +Center Border cFF" FadeCols[1] " ow2 oc66FFFFFF r4 s30"
	Font1 = Futura Md Bt
	Text1 = %tid1%
	;--------------------
	Options2 = x220 y26 cFFffffff r4 s12
	Font2 = Futura Bk Bt
	Text2 = WZeit: %WZeit%
	;--------------------
	Options3 = x220 y40 cFFffffff r4 s12
	Font3 = Futura Bk Bt
	Text3 = WPat : %WPat%
	;--------------------
	Options4 = X3 y26 w78 h30 Wrap cFFffffff r4 s10
	Font4 = Futura Bk Bt
	Text4 = %PatName%
	;--------------------
	Options5 = x0 y56 h18 w300 +Center ccFFffff00 r4 s12
	Font5 = Futura Bk Bt
	Text5 = %Infopraxotext%
	;--------------------
	Options6 = x12 y5 h18 w60 +Center ccFFffff00 r4 s10
	Font6 = Futura Bk Bt
	Text6:= "Q" . GetQuartal("heute")

	Gdip_TextToGraphics2(G1, text1, Options1, Font1, PtGuiW, PtGuiH)

	Loop, 5	{
				AIndex:= A_Index+1
				Gdip_TextToGraphics(G1, text%AIndex%, Options%AIndex%, Font%AIndex%, PtGuiW, PtGuiH)
			}

	UpdateLayeredWindow(HwndPraxomat, GGhdc, PtGuiX, PtGuiY, PtGuiW, PtGuiH)
	WinShow, ahk_id %HwndPraxomat%
	  ;Option für später eingeblendete Infos's
	Font7 = Futura Bk Bt
	OptionsTTpline:= "x2 y" (PtGuiH - 19) "h18 w" (PtGuiW-5) " +Center cFFFFFFFF r4 s12"

return

;}

;###################################################
;{	10.	Labels - Patient in Behandlung setzen

OeffnePatientenAkte: ;{
	;ListLines
	If (PatName = "")
	{
						MsgBox, 0x40004, Patientstatus, Patient in Behandlung setzen?
						IfMsgBox, Yes
						{
								Startt:=A_TickCount , Pausett:= 0 , PatName:="", PatFound:= 0 , Behandlung:= 1
								Fos ++
								TopToolTip()

								Infopraxotext:= "Suche nach aktuellem Patientennamen!"
								SetTimer, InfoPraxo, 8000
								PatName:=AlbisInBehandlungSetzen()
								Infopraxotext:= "Habe Patientennamen gefunden!"
								SetTimer, InfoPraxo, 8000

								SetTimer, ColorIndexing, 500
						}
						IfMsgBox, No
								SendInput, {Enter}
	}
	  else
	{
					; wenn es noch einen eingestellten PatNamen gibt dann sollte zuerst dessen Zeit gesichert werden
						MsgBox, 0x40003, Achtung!, Es läuft noch ein Timer für %PatName%.`nMöchtest Du die Zeit für diesen Patienten sichern?
						IfMsgBox, No
						{
								Startt:=A_TickCount , Pausett:= 0 , PatName:="", PatFound:= 0 , Behandlung:= 0		;alles wichtige zurücksetzen
								goto, OeffnePatientenAkte
						}
						IfMsgBox, Cancel
						{
								return
						}
						IfMsgBox, Yes
						{
								gosub, Speichern
						}
	}

return
;}

GoPatName: ;{

		;--es braucht öfter mehrere Sekunden das er den aktiven Patienten findet, der dann für den akutellen Timer gilt

		if !Instr(AlbisGetActiveWindowType(), "Akte") And (PatFound = 0) {
					SetTimer, GoPatName, off
					PatName:= AlbisCurrentPatient(), PatFound:= 1
					Infopraxotext:= "Habe Patientennamen gefunden!"
					SetTimer, InfoPraxo, -8000
		} else {
					Sleep, 500
					If ( (goPLoop++)>20 )
						return
					else
						goto GoPatName
		}

		goPLoop:=0

return
;}

checkDoubleClick: ;{

loop, 100
{
sleep, 1
lp:= A_Index
If (GetKeyState("LButton", "D") = 1)
		break
}

return
;}

;}

;{	11. 	Labels - für Infoeinblendungen
InfoPraxo: ;{

	If (!Alert) {
		InfoPraxoText:= Werbung
		InfoPraxoTimer = 0
		SetTimer, InfoPraxo, off
	}

return
;}

SplashTextAus: ;{
	SplashTextOff
	ToolTip, , , , %TTnum%
	SetTimer, SplashTextAus, Off

return
;}

;}

;{	12.	Labels - Zusatzfunktionen                	- GVU Liste erstellen, Ziffern wählen, Ziffern einfügen, MessageWorker

GVUListe: ;{										; SubModul - Eintragen eines Patienten in die GVU Liste

	PatID:= AlbisAktuellePatID()																	;Zusammenkürzen auf die richtige Länge , brauche nur die PatNr.

	If (progdate=0) {
			QLDatum:= AlbisLeseZeilenDatum(300)
			MsgBox, % QLDatum
			If QLDatum = 0
			{
					DatumEingeben:
					InputBox, QLDatum, Addendum für Albis on Windows, % "Es gab ein Problem beim Auslesen des Datums.`nBitte geben Sie das Datum in der Form TT.MM.JJJJ ein!"
					If !RegExMatch(QLDatum, "\d\d\.\d\d\.\d\d\d\d")
							goto DatumEingeben
					else if QLDatum = ""
							return
			}
	} else {
			QLDatum:= AlbisLeseProgrammDatum()										;abhängig vom Tagesdatum lassen sich die Daten in die GVUListe eintragen
	}

	QListe:= GetQuartal(QLDatum)
	GVUtoAddfile = %AddendumDir%\Tagesprotokolle\%QListe%-GVU.txt

		;wird immer komplett durchsucht, um die zum Quartal zugehörigen Untersuchungen zählen zu können
	 GVUvorhanden = 0
	 Loop, Read, %GVUtoAddfile%
	{
				LineIdx:= A_Index
				If  A_LoopReadLine contains %PatID%
				{
					GVUvorhanden = 1
					break
				}

	}

	If (GVUvorhanden = 1) {
				MsgBox, 0x40001, GVUListe, Patient mit der ID: %PatID% wurde schon übertragen., 3
	} else {
				GVUAnzahl[QListe]:= LineIdx + 1
				FileAppend, % QLDatum . "`;" . SubStr(QListe, 1, 2) . "/" . SubStr(QListe, 3, 2) . "`;" . PatID . "`n", % GVUtoAddfile
				GVUListPatName:= "Pat.: " . AlbisCurrentPatient() . "`nin GVU-Q Liste gespeichert!."
				Bezier_MouseMove2( 1, 1, 700, 100, 1920, 500, 1200, 0, GVUListPatName )
				TopToolTip()
	}

	Infopraxotext:= "Die GVU Liste " . SubStr(QListe, 1, 2) . "/" . SubStr(QListe, 3, 2) . " enthält " . GVUAnzahl[QListe] . " Untersuchungen."
	SetTimer, InfoPraxo, 6000

	;If !progdate
	;	SendInput, {LControl Down}{F4}{LControl Up}

return
;}

ZifferWaehlen: ;{								; unfertig
	w=300
	h=100

return
;}

ZifferEinfuegen: ;{								; unfertig

	SendInput, {INS}
	SendRaw, lko
	SendInput, {TAB}
	SendRaw, 03230
	SendInput, {TAB}
	SendInput, {ESC}

 return
;}

ZEGui: ;{											; unfertig

	CoordMode, Mouse, Screen
	MouseGetPos, mx, my
	Gui, ZE:NEW																					;kein Rand, kein TaskbarIcon, immer OnTop
	Gui, ZE:Color, Green
	Gui, ZE:+LastFound +AlwaysOnTop +ToolWindow
	Gui, ZE:-Caption
	Gui,Ze:Margin, 0, 0
	Gui, ZE:Font, S8 CDefault, Futura Bk Md, 5
	part1:= Pfeil links 				 Ziffer einfügen
	part2:= Pfeil nach unten 	 Akte schließen
	part3:= Escape Taste		 Einfügen Beenden
	part:= part1 . "`n" . part2 . "`n" . part3
	Gui, ZE:Add, Text, x25 y10 Center vZEinfueg cBlack ,  `n
	mx:= mx+100
	my:= my+50
	Gui, ZE:Show, x%mx% y%my% AutoSize

return
;}

MessageWorker: ;{
	MWmsg:= StrSplit(InComing, "|")
	If InStr(MWmsg[1], "PatDBCount") {
				PatDBCount:= MWmsg[2]
				TopToolTip()
	}
	If InStr(MWmsg[1], "TPCount") {
				TPCount:= MWmsg[2]
				TopToolTip()
	}
	Infopraxotext:= "InComing message : " InComing
	SetTimer, InfoPraxo, 12000
return
;}
;}

;###################################################
;{	15.	Labels - Timerfunktion                   	- Wartezeit, Aktualisierte OSD

Wartezeit: ;{

		FileReadLine, WZeit, %WZStatistikFile%, 1
		FileReadLine, WPat, %WZStatistikFile%, 2
			if (ErrorLevel = 1) {
				Infopraxotext:= "Statistikdateien sind nicht lesbar!"
				Alert:= 1
				SetTimer, InfoPraxo, 10000
			}
			;da Botti die Wartezeit nicht sehr schön darstellt, erfolgt das jetzt hier eben
		StringSplit, WZeit, WZeit, :
		WZeit:= SubStr("00" . WZeit1, -1) . ":" . SubStr("00" . WZeit2, -1)
		WPat:= SubStr("00" . WPat, -1)
		FileDelete, %WZStatistikFile%
		FileAppend, %WZeit%`n%WPat% ,%WZStatistikFile%
		WZeit1:="", WZeit2:=""

return
;}

AktualisiereOSD: ;{

	ListLines, Off
	diffms:=Round((A_TickCount - Startt))													; Millisekunden seitdem der Timer gestartet wurde
	Sekunden := (diffms//1000)			;+ 1000										; Millisekunden in Sekunden

	HilfeMenu(HwndPraxomat, StnFont, SGroesse, xVerschiebung, BgColor)
	AlbisWinID:= AlbisWinID()																	; stürzt Albis ab sollte man regelmäßig die Albis Fensterid erneuern

	;{der Timer bleibt stehen oder läuft weiter

	AlbisWindowState:= InStr(AlbisGetActiveWindowType(), "WZ")

	If (AlbisWindowState = 1) or (MPause = 1)
	{
			If (StopStatus = 0) {
					Pausett += Sekunden
					diffms_old:= diffms
					StopStatus = 1																				;Timer stoppt bei Focuswechsel; die StartZeit wird danach nach erneutem Focuswechsel neu geschrieben,
			}
	}
	  else if (AlbisWindowState = 0) or (MPause = 0)
	{
				If (StopStatus = 1) {
						Startt:= A_TickCount	 																;Timer war angehalten, dann beginnen wir mit neuer Zählung die angesammelte Zeit hatten wir ja aufgehoben (Pausett)
						StopStatus = 0
				}
			    	;und hier läuft sie weiter
				tid1=% FormatSeconds(Sekunden + Pausett) 										;formatierten String ausgeben
				TimeCode:= Floor(SubStr(tid1, 4, 1)/5) + 1										;TimeCode - alle 5 Minuten ergibt sich eine neue Zahl für die Hintergrundfarbänderung
				If (Behandlung = 1) {
					Options1:=  "x70 y23 w160 h35 +Center Border cFF" FadeCols[ColorIndex] " ow2 oc66FFFFFF r4 s30"
				} else {
					Options1:=  "x70 y23 w160 h35 +Center Border cFF" FadeCols[1] " ow2 oc66FFFFFF r4 s30"
					SetTimer, ColorIndexing, Off
				}
	}

	;}

	;{Auffrischen der Praxomat Gui

		Text1:= tid1, Text2:= "WZeit: " . WZeit, Text3:= "WPat : " . WPat, Text4:= PatName, Text5:= Infopraxotext
			;1.Farbhintergrund (oberer)
		pBrush1 := Gdip_BrushCreateSolid("0xaa" . BgColor)
		Gdip_FillRectangle(G1, pBrush1, 0, 0, PtGuiW-3, PtGuiH - 20)
		Gdip_DeleteBrush(pBrush1)
			;2. Farbhintergrund (unterer)
		pBrush1 := Gdip_BrushCreateSolid("0xaa172842")
		Gdip_FillRectangle(G1, pBrush1, 0, PtGuiH - 20, PtGuiW-3, 20)
		Gdip_DeleteBrush(pBrush1)
			;2.a. hellblauer Rahmen um den 2.Farbhintergrund
		pPen := Gdip_CreatePen("0xaa5598FC",2)
		Gdip_DrawRectangle(G1, pPen, 1, PtGuiH - 20, PtGuiW-5, 19)
		Gdip_DeletePen(pPen)
			;3. Picture-Overlay
		Gdip_DrawImage(G1, pBitmap1, 0, 0, PtGuiW, PtGuiH, 0, 0, PtGuiW, PtGuiH)
		Gdip_TextToGraphics2(G1, text1, Options1, Font1, PtGuiW, PtGuiH)
		Loop, 5	{
				AIndex:= A_Index+1
				Gdip_TextToGraphics(G1, text%AIndex%, Options%AIndex%, Font%AIndex%, PtGuiW, PtGuiH)
			}
		Gdip_TextToGraphics2(G1, TTpline, OptionsTTpline, Font6, PtGuiW, PtGuiH)
		UpdateLayeredWindow(HwndPraxomat, GGhdc, PtGuiX, PtGuiY, PtGuiW, PtGuiH)

	;}

	;{prüft ob Albis noch läuft
	;Process, Exist , albisCs.exe
		If !AlbisWinID {
						MsgBox, 0x40004, Neustart, Albis scheint nicht mehr zu funktionieren.`nMöchten Sie es neu starten?
						IfMsgBox, Yes
								RunWait, %AlbisExeDir%\%AlbisExe%, %AlbisWorkDir% , Max
						IfMsgBox, No
								ExitApp
					}
	;}

	;If TimeCode > 30
	;	SetTimer, AktualisiereOSD, Off									; Timer-Gui aktualisieren

return
;}

SaveTimerWindow: ;{

	SetTimer, SaveTimerWindow, Off
	WinGetPos, PtGuiX, PtGuiY, , , ahk_id %HwndPraxomat%
	IniWrite, %PtGuiX%, %AddendumDir%\Addendum.ini, Praxomat, Gui_PosXGui
	IniWrite, %PtGuiY%, %AddendumDir%\Addendum.ini, Praxomat, Gui_PosYGui
	iTHappens:= ""

return
;}

Speichern: ;{

	wzflag:=0
	MsgBox, 0x40004,Addendum für AlbisOnWindows, Der aktuelle Patient ist: %PatName%.`nWollen Sie die Daten zum Patienten jetzt speichern?. , 10
	IfMsgBox, No
		return

	Fis ++

	tvar:= StrSplit(PatName, "`,", " ")
	If !AlbisAkteGeoeffnet(tvar[1], tvar[2])
				AlbisAkteOeffnen(PatName)

	Infopraxotext:= "Abspeichern der Gesprächszeit für: " . PatName
	SetTimer, InfoPraxo, -2000
	TopToolTip()

	FileAppend, % A_DD . A_MM . A_YYYY . "|" . A_Hour . ": " . A_Min . "|" . tid1 . "|" PatAll . "`n", %AddendumDir%\logs'n'data\PraxomatGZeit.log
	TenMinutes := Floor((Sekunden//60)/10)*10									;ermittle den 10 minuten Takt

		;Versicherungsart feststellen - 1 = privat , 2- gesetzlich
	Vstatus:= AlbisVersArt()

		;ermitteln der Gesprächszeit im 10 Minutentakt
	If (TenMinutes > 0) And (Vstatus = 2)
	{
					;ausschalten der TimerGui Aktualisierung, aufrufen der TenMinutesGui
					gosub TenMinutesGui
					if (lko <> 0) {
							FLevel:=AlbisPrepareInput("KKk")									;Setze den Cursor auf das Eingabefeld(Kürzel)
							if (FLevel=99) {
									Infopraxotext:= "Kein Zugriff auf Eingabebereich der Akte erhalten!"
										SetTimer, InfoPraxo, -2500
							} else {
									AlbisSendInputLT("l", lko, "lkü", "lko")						;schreibt lko oder lkn ein und dann die Ziffer
							}
					}

	}
	  else if (TenMinutes = 0) And (Vstatus = 2)
	{
					lko:= "apk"
					FLevel:=AlbisPrepareInput("KKk")									;Setze den Cursor auf das Eingabefeld(Kürzel)
					if (FLevel=99) {
								Infopraxotext:= "Kein Zugriff auf Eingabebereich der Akte erhalten!"
								SetTimer, InfoPraxo, -2500
					} else {
								AlbisSendInputLT("l", lko, "lkü", "lko")								;schreibt lko oder lkn ein und dann die Ziffer
					}
	}

	;--den Privatversicherten schreibe ich die Gesprächszeit in die Akte
	If (Vstatus = 1)
	{
					FLevel:=AlbisPrepareInput("KKk")
					if (FLevel=99) {
								Infopraxotext:= "Kein Zugriff auf Eingabebereich der Akte erhalten!"
								SetTimer, InfoPraxo, -2500
					} else {
								AlbisSendInputLT("info", "Gesprächszeit: " . tid1, "iic", "info")
					}
	}

	;--alles auf Null was die Zeit und die flags dazu betrifft
	Startt:=A_TickCount, Pausett:= 0, Patname:="", PatFound:= 0, Behandlung:= 0, ColorIndex:=1
	Options1:=  "x70 y23 w160 h35 +Center Border c" FadeCols[1] " ow2 oc66FFFFFF r4 s30"

	;--zurückstellen der Anzeige auf Null mit der Startfarbe die unter AC1BgCol1 gespeichert ist
	diffms:=Round((A_TickCount - Startt)) , Sekunden := (diffms//1000) , tid1:= FormatSeconds(Sekunden + Pausett),
		Sleep, 500

	Infopraxotext:= "Gesprächszeit für Pat.: " . PatName . " gespeichert."
	SetTimer, InfoPraxo, -8000

	If (wzflag=0) {
					wzflag:=AlbisOeffneWarteZimmer()
					AlbisWZPatientEntfernen(tvar[1], tvar[2])
	}

	PatAll:= "", PatName:=""
	VarSetCapacity(tvar, 0)

return
;}

Erneuern: ;{

	Startt:=A_TickCount, Pausett:= 0, Patname:="", PatFound:= 0, Behandlung:= 0
	If ColorIndex < 3001
			SetTimer, AktualisiereOSD, 25									; Timer-Gui aktualisieren

return
;}

ShowPraxomatGui: ;{

	SetTimer, ShowPraxomatGui, Off
	WinShow, ahk_id %HwndPraxomat%

return
;}

PauseToggleTimer: ;{

		;--Routine für das manuelle pausieren des Timer - diese Routine wird mit der Pausetaste als Hotkey aufgerufen
    If (MPause = 0) and (result=0) 	; manuelles Pausieren geht nur wenn nicht schon durch das Aufrufen des Wartezimmer der Timer pausiert wurde
	{
					;--ein neues Fenster mit Hinweis erstellen
					Gdiw = 400, Gdih = 216, Gdix:= Floor(A_ScreenWidth - Gdiw), Gdiy:= Floor(A_ScreenHeight - Gdih), MPause:= 1
					SetTimer, ColorIndexing, Off
					hIdleWindow:=PIC_GUI("IW", AddendumDir . "\assets\Pause.png", Gdix/2, Gdiy/2, Gdiw, Gdih, "IdleWindow")
					WinActivate,  ahk_id %AlbisWinID%
	}
	  else if (MPause = 1)
	{
					MPause = 0
					Gui, IW: Hide
					SetTimer, ColorIndexing, On
					WinActivate,  ahk_id %AlbisWinID%
	}

return

IdleWindowOff:

	Gdip_Shutdown(pToken1)
	Gui, 20:Destroy
	MPause = 0
	WinActivate,  ahk_id %AlbisWinID%

return
;}

TenMinutesGui:	 ;{


	;-------------------------------------- Gesprächsziffernauswahl --------------------------------------------
	;----------------------------------- 03230 - 35100 oder 35110 -------------------------------------------

	;{a Berechnung und Variablenzuweisungen für GUI Anzeige
		Leistungsfaktor:= Round(TenMinutes/10,0)					;faktor hieß die Variable zuvor - ist aber schon in Gebrauch gewesen
		lko:= "03230(x:" . Leistungsfaktor . ")"
		return
		lko = 0																		;weil lko manchmal schon eine Ziffer enthielt hat er das GUI nicht mehr aufgerufen
	;}

	;{15.8.b GUI Aufbau mit WinWaitClose - damit das GUI nicht ohne Userinteraktion geschlossen wird

		;static Edit1, Edit2, Edit3, RadioB1, RadioB2, RadioB3, uebernehmen, abbrechen
		Gui, TM: NEW
		Gui ,TM: -MinimizeBox -MaximizeBox -SysMenu +AlwaysOnTop +HwndTMid
		Gui ,TM: Color, 0x172842 ;0xAAAAFF
		Gui ,TM: Margin, 5,5
		Gui ,TM: Font, s14 q5 cWhite, Futura Bk Md
		Gui ,TM: Add, Text, xCenter y10 w%TMWidth% h33 +0x200 Center +BackgroundTrans, %TenMinutes% min. Gesprächsdauer mit Patient(in): %PatName%.
		Gui ,TM: Font, s9 q5, Futura Bk Bt
		Gui ,TM: Add, Text, xCenter y35 w%TMWidth% h28 +0x200 Center +BackgroundTrans, Die Möglichkeit für das Ansetzen einer Gesprächsziffer ist gegeben.
		Gui ,TM: Font, s14 q5
		GOPLine = GOP %Ziffer1%`(x:   `)`nGOP %Ziffer2%`(x:   `)`nGOP %Ziffer3%`(x:   `) +E0x80000
		Gui ,TM: Add, Text, x%GOPx% y%GOPy1% w%GOPw% h%GOPh% +BackgroundTrans, GOP %Ziffer1%`(x:     `)
		Gui ,TM: Add, Text, x%GOPx% y%GOPy2% w%GOPw% h%GOPh% +BackgroundTrans, GOP %Ziffer2%`(x:     `)
		Gui ,TM: Add, Text, x%GOPx% y%GOPy3% w%GOPw% h%GOPh% +BackgroundTrans, GOP %Ziffer3%`(x:     `)
		Gui ,TM: Font, s14 q5 cBlack
		Gui ,TM: Add, Edit, vEdit1 gTenMin x%FGOPx% y%GOPy1% w%FGOPw% h%GOPh% +Number -Vscroll Left +BackgroundTrans, %Leistungsfaktor%
		Gui ,TM: Add, Edit, vEdit2 gTenMin x%FGOPx% y%GOPy2% w%FGOPw% h%GOPh% +Number -Vscroll Left +BackgroundTrans Disabled 0, %Leistungsfaktor%
		Gui ,TM: Add, Edit, vEdit3 gTenMin x%FGOPx% y%GOPy3% w%FGOPw% h%GOPh% +Number -VScroll Left +BackgroundTrans Disabled 0, %Leistungsfaktor%
		Gui ,TM: Font, s10 q5 cWhite
		Gui ,TM: Add, Radio, vRadioB1 gTenMin x%GOPrx% y%GOPy1% w%GOPrw% h%GOPh% +BackgroundTrans +Checked
		Gui ,TM: Add, Radio, vRadioB2 gTenMin x%GOPrx% y%GOPy2% w%GOPrw% h%GOPh% +BackgroundTrans
		Gui ,TM: Add, Radio, vRadioB3 gTenMin x%GOPrx% y%GOPy3% w%GOPrw% h%GOPh% +BackgroundTrans
		Gui ,TM: Font, s12 q5 cWhite
		Gui ,TM: Add, Button, vuebernehmen gTenMin x%Bx% y%GOPy1% w%Bw% h%Bh% +Default +BackgroundTrans, In die Akte &übernehmen
		Gui ,TM: Font, s12 q5
		Gui ,TM: Add, Button, vabbrechen gTenMin x%Bx% y%GOPy3% w%Bw% h%Bh% +BackgroundTrans, Vorgang &abbrechen
		Gui ,TM: Show, xCenter yCenter w%TMWidth% h%TMHeight%, Praxomat Timer Hinweis
		Gui, TM: Default

		;ein Versuch neu am 05.04. - dieses GUI wird manchmal nicht geöffnet, vielleicht weil dieses Label zu früh verlassen wurde
		;WinWaitClose, ahk_id %TMid%
return

	;}

	;{15.8.c GUI User Interaktionsroutinen (gLabel)

TenMin:

	CName:= A_GuiControl
	Gui, TM: Submit, NoHide
	MsgBox, % CName

	Infopraxotext:= CName " gedrückt."
							SetTimer, InfoPraxo, 8000

	If (InStr(CName, "RadioB")) {
			StringRight, num, CName, 1
			Loop 3 	{
						If (num <> A_Index) {
							GuiControl, Disable, % "Edit" (A_Index)
						} else {
							GuiControl, Enable,  % "Edit" (A_Index)
							GuiControl, Enable,  % "Edit" (A_Index+1)
						}

			}
	}

	If  (InStr(CName, "uebernehmen")) {
			Loop 3 {

				i:= A_Index
				ControlGet, Estatus, Enabled, , Edit%i%

				If (EStatus = 1) {

						ControlGet, num, Line, 1, Edit%i%
						Ziffer:= % Ziffer%i%

						if (num = 1) {
								lko = %ziffer%
							} else {
									lko = %ziffer%`(x`:%num%`)
											}

						goto TMGuiClose

					}
			}
	}

	if (Instr(CName, "abbrechen")) {
			goto TMGuiClose
	}

return

TMGuiClose:
		Gui, TM: Destroy
		Gui, 22: Default		;ist das Praxomat-Gui
		TenMinutes = 0
return

;}

;}

ColorIndexing: ;{

	If ColorIndex < 3001
			ColorIndex++

return
;}

;}

;{ 16.	TrayIcon - Prozesse

Hotkeys: ;{

	;--Anzeige aller benutzten Hotkeys mit Beschreibung durch auslesen des Skriptes
	script:= A_ScriptDir . "\" . A_ScriptName



return
;}

SkriptReload: ;{

	Script:= A_ScriptDir "\" A_ScriptName
	scriptPID := DllCall("GetCurrentProcessId")
	run,  Autohotkey.exe /f "%AddendumDir%\include\RestartScript.ahk" "%Script%" "2" "%A_ScriptHwnd%" "%scriptPID%"

return
;}

NoNo: ;{

	B:="`n"
	Loop 12 {
	B .= A%A_Index% "`n"
	}

	Gui UberPt: New, -Caption HWNDhUberPt
	Gui UberPt: Font, S12 CDefault, Futura Bk Bt
	Gui UberPt: Add, Edit, r20 Center ReadOnly -WantCtrlA -WantReturn -Wrap HWNDhUberTextPT, %B%
	Gui UberPt: Add, Button, ys vUberOk gUberEnde, OK
	Gui UberPt: Show, AutoSize, Praxomat
	ControlFocus, UberOk, ahk_id %hUberPt%
	SendInput, {Left}
	DllCall("HideCaret","Int",hUberTextPT)

return

UberEnde:

	If (A_GuiControl=="UberOk") {
			B:=""									;--Variable B wieder leeren um RAM zu sparen
			Gui UberPt: Destroy
	}

return


;}

ZeigeVariablen: ;{
	ListVars
return
;}

;}

;###################################################
;{	19.	Funktion bei Beendigung des Skriptes

MaxScite:

	OnExit

	If hWinEventHook1 {
			UnhookWinEvent(hWinEventHook, HookProcAdr)
	}

	WinClose, ahk_id %HwndPraxomat%
	WinClose, ahk_id %HLPHg%

	;Visualisierung das Programm geschlossen wird, als Hinweis das es nicht fehlerbedingt beendet wird
	Progress, B2 cW00FFFF cB008080 w300 s18, , Du hast Praxomat beendet., Praxomat-Ende, %StnFont%
	Loop 33 {
		p:= 100 - (A_Index*3)
		Progress,  %p%
		Sleep, 1
	}

	Progress, Off


ExitApp

;}

;{ 20.	Funktionen / zusätzliche Gui's /

Receive_WM_COPYDATA(wParam, lParam) {
	global InComing
    StringAddress := NumGet(lParam + 2*A_PtrSize)
	InComing := StrGet(StringAddress)
	SetTimer, MessageWorker, -10
    return true
}

LM_DoubleClick(hWinEventHook1, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime ) {            									;-- Doppelklick abfangen um Patiententimer zu starten

	Critical
	ToolTip, DoubleClick received
	hPtAncestor:= GetHex(GetAncestor(hwnd,4))					;Owner
	if (hPtAncestor=AlbisWinID()) { ; LButton
		WinGetClass, class, % "ahk_id " hwnd
		If InStr(class, "Listview") {
					ControlGet, result, List, Selected , % classNN, % "ahk_id" hwnd
					fpos:= RegExMatch(result, "O)\w+\,\s\W+", Match)
					MsgBox, % name:= Match.Value()
					fpos:= RegExMatch(result, "O)(?<=\()\d+(?=\,)", Match)
					id:= Match.Value()
					lvObj:= {"name": name, "id": id}
		}
	}

}

PtWM_MOUSEMOVE(wparam, lparam, msg, hwnd) {

	if (hwnd=HwndPraxomat) { ; LButton
			CoordMode, Mouse, Client
			MouseGetPos, mx, my
			If ( ((mx>1) and (mx<PtGuiW)) and ( (my>(PtGuiH-PtDelta)) and (my<PtGuiH) )  )
			{
					ToolTip, Drücke hier die linke Maustaste um Hilfe zu bekommen!, % PtGuiX, % (PtGuiY+PtGuiH+2), 14 		;ToolTip, % mx " : 0-" PtGuiW "`n" my " : 0-" PtGuiH
					SetTimer, ToolTip14Off, -1000
			}
			CoordMode, Mouse, Screen
	}
	return

	ToolTip14Off:
		ToolTip,,,, 14
	return
}

PtWM_LBUTTONDOWN(wparam, lparam, msg, hWMLbd) {

	if (hWMLbd=HwndPraxomat) { ; LButton
			CoordMode, Mouse, Client
			MouseGetPos, mx, my
			If ( ((mx>1) and (mx<PtGuiW)) and ( (my>(PtGuiH-PtDelta)) and (my<PtGuiH) )  )
								Help_HotKeyGui(HwndPraxomat, "Futura Bk Bt", 12, 55, "cGreen")		;ToolTip, % mx " : 0-" PtGuiW "`n" my " : 0-" PtGuiH
			CoordMode, Mouse, Screen
	}

}

FormatSeconds(Sekunden) {
	Return SubStr("0" . Sekunden // 3600, -1) . ":"
        . SubStr("0" . Mod(Sekunden, 3600) // 60, -1) . ":"
        . SubStr("0" . Mod(Sekunden, 60), -1)
}

Bezier_MouseMove2( X1, Y1, X2, Y2, X3, Y3, TimeMS = 200, Relative = 0, TT_Text = "Liste fehlt!" ) {

; Code Source: http://www.autohotkey.com/forum/viewtopic.php?p=374641#374641

	Critical
	MouseGetPos, X0, Y0
	If (Relative)
		X1 += X0, X2 += X0, X3 += X0, Y1 += Y0, Y2 += Y0, Y3 += Y0
	TE := ( TI := A_TickCount ) + ( TimeMS := Abs( TimeMS ) )
	Loop
		If ( A_TickCount < TE )
		{
			t := ( A_TickCount - TI ) / TimeMS
			u := ( TE - A_TickCount ) / TimeMS
			bx := Round( X0 * u**3 + 3 * X1 * t * u**2 + 3 * X2 * u * t**2 + X3 * t**3 )
			by := Round( Y0 * u**3 + 3 * Y1 * t * u**2 + 3 * Y2 * u * t**2 + Y3 * t**3 )
			ToolTip, %TT_Text% , bx, by, 9
			Sleep 1
		}
		Else Break
	ToolTip, , X3, Y3, 9
	Critical, Off
}

IfWindowOpen(WinName) {

	DetectHiddenText, Off
	DetectHiddenWindows, off
	WinGet, AllWinsHwnd, List
	DetectHiddenText, On
	DetectHiddenWindows, On
	result = 0
	Loop, % AllWinsHwnd
		{
				if not IsWindow(AllWinsHwnd%A_Index%)
					continue
				WinGetTitle, CurrentWinTitle, % "ahk_id " AllWinsHwnd%A_Index%
				If Instr(CurrentWinTitle, WinName, 1) {
							result = 1
							break
					}
		}

return result
}

TopToolTip() {

	global TTpline
	heutelog:= A_DD "-" A_MM "-" A_yyyy
	IniWrite, % Fos "|" Fis "`," AddendumDir "\logs'n'data\Behandlung`.log", Behandlungen , % heutelog
	sub:= AlbisAktuellePatID()
	TTpline:= "Beh:" . SubStr("0" . Fos, -1) . "|Sp:" . SubStr("0" . Fis, -1) . "|P:" . SubStr("00000" . sub, -4) . "|GL:" . SubStr("000" . GVUAnzahl[QNow], -2) . "|DB:" PatDBCount . "|TP:" TPCount

}

HilfeMenu(HwndPraxomat, StnFont, SGroesse, xVerschiebung, BgColor) {

	static PHStatus, HLPHg
	global WTitleL

	WTitle:= MouseGetWinTitle()
	WinGetClass, wclass, WTitle
	If ( Instr(WTitle,"PraxomatGui") and (PHstatus="out") ) 							;  and InStr(wclass, "AutoHotkeyGUI" )
	{
			PHstatus:= "in"
		  	HLPHg:=Help_HotKeyGui(HwndPraxomat, StnFont, SGroesse, xVerschiebung, BgColor)
			Fade("in", HLPHg, 100, 255)
	}
	  else if ( !Instr(WTitle,"PraxomatGui") and (WTitle <> "PraxomatHilfe") and (PHstatus="in") )
	{
			PHstatus:="out"
			Fade("out", HLPHg, 100, 0)
			WinKill, ahk_id %HLPHg%
	}

}

Help_HotKeyGui(HwndPraxomat, StnFont, SGroesse, xVerschiebung, HLPColor) {

		static index, HLPControls:= Object()

		If !IsObject(HLPControls) {
				HLPControls:= {1: {Hkey: "{LControl Down}F12}{LControl Up}"		, display: "Strg + F12    	| Stoppuhr neustarten"}
										,2: {Hkey: "{LControl Down}{F11}{LControl Up}"		, display: "Strg + F11    	| Stoppuhrzeit speichern"}
										,3: {Hkey: "{LControl Down}{F7}{LControl Up}"		, display: "Strg + F7      	| Pat. in die GVU Liste"}
										,4: {Hkey: "{LControl Down}{F8}{LControl Up}"		, display: "Strg + F8      	| Formular Assistent (GVU)"}
										,5: {Hkey: "{LControl}{Up}"										, display: "Strg + Hoch 	| Modul Zifferneinfügen"}
										,6: {Hkey: "{LControl Down}{Left}{LControl Up}"		, display: "Strg + Links  	| gewählte Ziffer einfügen"}
										,7: {Hkey: "{LControl}{Down}"									, display: "Strg + Runter	| Akte schließen"}
										,8: {Hkey: "{LControl Down}{Esc}{LControl Up}"		, display: "Strg + Esc     	| Ziffern einfügen beenden"}
										,9: {Hkey: "{LControl Down}{F10}{LControl Up}"		, display: "Strg + F10    	| Pat.ins Wartezimmer"}
										,10: {Hkey: ""															, display: "Space           	| (abwesend markiert)"}
										,11: {Hkey: "{LControl Down}ä{LControl Up}"			, display: "Strg + ä       	| Praxomat starten/beenden"}}
		}

		If !index {
					For index, obj in HLPControls
							HLPBLNr:= index
		}

		;Einstellungen des kleinen Hilfefensters  - dieses wird über eine Funktion eingeblendet
		faktor:= Round(SGroesse*10)
		WinGetPos, acx, acy, acw, ach, ahk_id %HwndPraxomat%
		ax:= acx , ay:= acy + ach + 1, acw:= acw - 3, bnplus:= 30

		Gui, HLP: NEW 																												;kein Rand, kein TaskbarIcon, immer OnTop
		Gui, HLP: Color, %HLPColor% ;Green
		Gui, HLP: +AlwaysOnTop +ToolWindow +HwndHLPHg
		Gui, HLP: -Caption +Owner
		Gui,	HLP: Margin, 10, 5
		Gui, HLP: Font, S8 CDefault Wnorm q5, %StnFont%, 6
		Loop, % HLPBLNr
		{
				HLPBName:= "HLPB" . A_Index
				Gui, HLP:Add, Button, % "yp+" bnplus " v" HLPBName " gHLPB Left", % HLPControls[A_Index].display
		}

		HLPBmaxW:=0, HLPLastButtonName:= "HLPB" HLPBLNr
		GuiControlGet, HLPBLast, Pos, % HLPLastButtonName

		Loop % HLPBLNr
		{
				GuiControlGet, HLPB, Pos, % "HLPB" A_Index
				If (HLPBW>HLPBmaxW)
						HLPBmaxW:= HLPBW
		}

		Loop % HLPBLNr
		{
				GuiControl, HLP: Move, % "HLPB" A_Index, % "x" (( acw/2) - (HLPBmaxW/2) ) " w" HLPBmaxW
		}

		Gui, HLP:Font, s10 bold italic q5, Symbol, 8
		Gui, HLP:Add, Button, % "x10 	y" (HLPBLastY + HLPBLastH + bnplus*2) " cBlack vHLPShowMoreL"  , <---------
		Gui, HLP:Add, Button, % "x130 	y" (HLPBLastY + HLPBLastH + bnplus*2) " cBlack vHLPShowMoreR" , --------->
		GuiControlGet, HLPBSize, Pos, % "HLPShowMoreR"
		GuiControl, HLP: Move, % "HLPShowMoreR", % "x" (acw - HLPBSizeW - 10)
		Gui, HLP:Show, x%ax% y%ay% w%acw% , PraxomatHilfe

return HLPHg

HLPB:

	  ;sendet eigentlich nur einen Hotkey, so kann ich weiterhin einen eigenen Thread per Aufruf starten
	keyCmd:= StrReplace(A_GuiControl, "HLPB")
	WinActivate, ahk_id %AlbisWinID%
		WinWaitActive, ahk_id %AlbisWinID%
	SendInput, % HLPControls[keyCmd].HKey

return

}

PraxomatInfo(WinName="Addendum für Albis on Windows", WinText="", PIx=1000, PIy=1000) {
	Pos = Instr(WinName, %A_Space%)
	WinSearchTitle = SubStr(WinName, 1, %Pos%)
	SplashTextOn, 400, 50, %WinName%, %WinText%
	WinMove, %WinSearchTitle%, , %PIx%, %PIy%
	WinGet, SplashPraxID, ID, %WinSearchTitle%
return SplashPraxID
}

SetTextAndResize(controlHwnd, newText, fontOptions := "", fontName := "") {
    Gui 9:Font, %fontOptions%, %fontName%
    Gui 9:Add, Text, R1, %newText%
    GuiControlGet T, 9:Pos, Static1
    Gui 9:Destroy

    GuiControl,, %controlHwnd%, %newText%
    GuiControl Move, %controlHwnd%, % " w" TW					; "h" TH
}

Hook_Albis(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime ) {

		SetTimer, AktualisiereOSD, Off
		Critical
		hHookedWin:=GetHex(hwnd)
		WinGetTitle, hookTitle, ahk_id %hHookedWin%
		WinGetClass, hookClass, ahk_id %hHookedWin%

		If InStr(hookTitle, "ALBIS") and InStr(hookClass, "OptoAppClass")
		{
					  ;Albis Fenster wurde verschoben oder die Größe wurde geändert (nicht minimiert oder maximiert)
					If (event=11) or (event=23) or (event=32779) {
						acc:= Object(), loc:= Object()
						acc := Acc_Get("Object", "3", 13, "ALBIS ahk_class OptoAppClass")
						loc := Acc_Location(acc, 13)
						PtGuiX:= loc.x - PtGuiW + 3, PtGuiY:= loc.y
						AlbisPosIsChanged:=1
					} else if (event=10) {
						PtGuiX:= Mon1Right - PtGuiW, PtGuiY:= 3
						AlbisPosIsChanged:=1
					}

					If AlbisPosIsChanged {
						SetWindowPos(HwndPraxomat, PtGuiX, PtGuiY, PtGuiW, PtGuiH)
					}
		}

		SetTimer, AktualisiereOSD, On

return
}


Create_Praxomat_ico(NewHandle := False) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAAAAAAAAAD/oo//oo//oo//oo//oo//oo/zk5X7nY//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAAAAAAAAAD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo8AAAD/oo//oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo//oo//oo//oo//oo//oo/9oY7/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdUMyKFU0GrbFnHfmzZinflkX7lkX7ZinfHfmuqa1mEU0FTMyJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdJLRuOWUfThnP+oY7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/+oY7ThXKNWEdJLRtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdGKhqWXk3ul4T/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/ul4SVXUxGKhpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjHgjnv/oo//oo//oo//oo//oo//oo/+oo/3nYrxmYbul4Tul4TxmYb3nYr+oo//oo//oo//oo//oo//oo//oo/fjntqQjBCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeVXUz9oY7/oo//oo//oo//oo/4novYiXa6dWNILBpCKBdCKBdCKBdCKBdCKBdCKBdILBq6dWPYiXb8oI3/oo//oo//oo//oo/7n42UXUtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRjXiHb/oo//oo//oo//oo/8oI3MgW5JLBxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBemaVfmkn/8oI3/oo//oo//oo//oo+rbFpDKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBfXiHb/oo//oo//oo//oo+/eGdSMiFCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeqa1nwmIb/oo//oo//oo//oo+rbFpCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBfMgW7/oo//oo//oo//oo+UXktCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeiZlTmkn//oo//oo//oo//oo+UXUtCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBe3c2H+oo//oo//oo/+oY6FU0FCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfEfGr/oo//oo//oo/9oY5qQjBCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdGKxnwmIb/oo//oo//oo+UXUtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfMgW7/oo//oo//oo/fjntGKhpCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBfNgm//oo//oo//oo+/eWZCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfhjnz/oo//oo//oo+WXkxCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdJLBz3nYr/oo//oo/ymodSMiFCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBepa1n8oI3/oo//oo/ul4RJLRtCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBeOWUj/oo//oo//oo+UXktCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfmkn//oo//oo//oo+NWEdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBfVh3T/oo//oo/wmIZJLBxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdcOSiqa1nIf2xoQC9KLRz8oI3/oo//oo/ThXJCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdVNST+oY7/oo//oo+ublxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRh1STjGfWv/oo//oo//oo/EfGpCKBfsloP/oo//oo/+oY5TMyJCKBdCKBf/oo//oo//oo//oo9CKBdCKBeFVEL/oo//oo//oo9xRjRCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdLLh2QW0nwmIb/oo//oo//oo//oo//oo+LV0ZCKBdyRzX/oo//oo//oo+EU0FCKBdCKBf/oo//oo//oo//oo9CKBdCKBesbVr/oo//oo/8oI5HKxpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdbOSenaVf5nov/oo//oo//oo//oo/7n4z/oo/vmIVCKBdCKBdILBr+oY//oo//oo+rbFlCKBdCKBf/oo//oo//oo//oo9CKBdCKBfymof/oo//oo/ci3lCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdLLh1LLh1uRTTgjnv+oY7/oo//oo//oo//oo/ymYaUXktLLh3ymof/oo9UMyJCKBdCKBf3nYr/oo//oo/HfmtCKBdCKBf/oo//oo//oo//oo9CKBdCKBf2nYr/oo//oo/EfGlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeZYE79oY7/oo//oo//oo//oo//oo//oo/Ngm97TTtDKRhCKBdCKBfij3z/oo9qQjBCKBdCKBfij33/oo//oo/ainhCKBdCKBf/oo//oo//oo//oo9CKBdCKBf5nov/oo//oo+5dWJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdMLh3/oo//oo//oo//oo//oo/5n4urbFlbOSdCKBdCKBdCKBdCKBdCKBfXiXb/oo+AUD9CKBdCKBe6dmP/oo//oo/kkH1CKBdCKBf/oo//oo//oo//oo9CKBdCKBf5nov/oo//oo+5dWJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdMLh3/oo//oo//oo//oo+VXkxLLh1CKBdCKBdCKBdCKBdCKBdCKBdCKBfUhnT/oo+HVENCKBdCKBfdjHn/oo//oo/kkH1CKBdCKBf/oo//oo//oo//oo9CKBdCKBftl4T/oo//oo/EfGlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfqlIL/oo//oo+ycV5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfdjHn/oo90SDdCKBdCKBfij33/oo//oo/ainhCKBdCKBf/oo//oo//oo//oo9CKBdCKBflkX7/oo//oo/bjHhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfZiXb/oo//oo91SThCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeiZlTqlIH/oo9gOylCKBdCKBful4T/oo//oo/HfmxCKBdCKBf/oo//oo//oo//oo9CKBdCKBesbVr/oo//oo/8oI1HKxpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfZiXb/oo//oo91SThCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBepa1n+oY73nYpEKhlCKBdILBr+oY//oo//oo+rbFlCKBdCKBf/oo//oo//oo//oo9CKBdCKBeGVEP/oo//oo//oo9wRTRCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfZiXb/oo//oo91SThCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfIf2z/oo+zcV9CKBdCKBdyRzX/oo//oo//oo+EU0FCKBdCKBf/oo//oo//oo//oo9CKBdCKBdVNST/oo//oo//oo+tbVtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfZiXb/oo//oo91SThCKBdCKBdCKBdCKBdCKBdCKBdCKBemaVb1nIn/oo9uRDNCKBdCKBfYiXb/oo//oo/+oY5UMyJCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBfVh3X/oo//oo/vmIVJLRtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfZiXb/oo//oo91SThCKBdCKBdCKBdCKBdCKBdCKBeiZlTShXP/oo/ij3xCKBdCKBdJLBz8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"
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

;{ 21.	Includes

	#Include %A_ScriptDir%\..\..\include\gdip_all.ahk
	#Include %A_ScriptDir%\..\..\include\Gdip_Ext.ahk
	#Include %A_ScriptDir%\..\..\lib\Addendum_Functions.ahk
	#Include %A_ScriptDir%\..\..\lib\Addendum_DB.ahk
	#Include %A_ScriptDir%\..\..\include\ini.ahk
	#Include %A_ScriptDir%\..\..\include\FindText.ahk
	#Include %A_ScriptDir%\..\..\include\ACC.ahk
	#Include %A_ScriptDir%\..\..\include\IPC.ahk
	#Include %A_ScriptDir%\..\..\include\Sift.ahk
	#Include %A_ScriptDir%\..\..\lib\Gui\PraxTT.ahk


;}

;Debug Lager
;FoxitReader_CheckOPenPDfs(): ToolTip, % "RowText: " RetrievedText "`nFoxitMI: " Foxit.MaxIndex() "`, WinTitle: " WT "`npdf: "  "`nnRow: " i "`nTT:1", 800, 80, 8
;BOCheckPDFReader: ToolTip, % "PDFReaderExe=" PDFReaderExe "`, ErrorLevel:" pexst "`, pdfshow= " pdfshow "`nFoxitID: " FoxitID, 800,30, 7
;PDFChecker: ToolTip, % "PDFrow: " PDFrow "`,  is checked: " LVRowState.checked "`, ErrorLevel was: " elLV


