;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                      	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                    	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                    	by Ixiko started in September 2017 - letzte Änderung 11.09.2021 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
return

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                              	BESCHREIBUNG DER FUNKTIONSBIBLIOTHEK ZUR RPA SOFTWARE -- ADDENDUM FÜR ALBIS on WINDOWS --
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;
;       FUNKTIONSDESIGN:
;
;       Die Funktionen sollen eine einfach durchzuführende und möglichst fehlerfreie RPA (robotic process automation) ermöglichen.
;   	Nachfolgend beschreibe ich das dafür verwendete Design:
;
;        - Datenerfassung                               	 -	erfolgt über das Auslesen des Albisfenstertitel, des Inhaltes eines Fenster oder durch Ermittlung eines
;                                                                      	bestehenden Eingabefocus
;
;        - das Albis Fensterhandle                   	 -	wird bei jedem Aufruf neu ermittelt, damit ein neues Albishandle, z.B. nach einem Absturz oder Neustart
;                                                                 		von Albis erkannt wird
;
;        - wenige Parameter		                     	 - viele Parameter = viele Variablen = viel mehr Möglichkeiten? Möglicherweise ja, aber auch eine deutlich
;                                                                     	höhere Wahrscheinlichkeit durch einen Tippfehler Probleme mit seinem Programm zu bekommen
;
;        - keine Wartezeit nach Funktionsende 	 - 	nachdem eine Automatisierungsfunktion beendet ist, ist kein sleep Befehl notwendig um auf
;                                                                   	Fenster oder anderes zu warten
;                                                        			  	Die nächsten Befehle können unmittelbar erfolgen.
;                                                                	    Es ist nicht unbedingt notwendig die Dauer eines sleep-Befehles innerhalb einer Funktion zu ändern !!
;                                                        			  	Die Zeiten haben ich durch tägliche Nutzung meiner Skripte im Laufe der Jahre optimiert. In meiner
;                                                                   	Praxis gibt es schnelle sowie auch langsame Computer. Auf allen funktioniert das Timing der Skripte
;                                                                    	gut.
;
;        - Ausführungsüberprüfung                 	- 	Funktionen welche Menupunkte oder Fenster in Albis öffnen, überprüfen ob Albis im Moment durch
;                                                                   	geöffnete Fenster blockiert ist und schliessen diese automatisch.  Die meisten dieser Funktionen prüfen
;                                                                     	auch ob das zu erwartende Dialogfenster geöffnet wurde. Damit wird verhindert das nachfolgende
;                                                                     	Automatisierungsfunktionen nicht ins Leere arbeiten oder durch wiederholte Abfolgen von Befehlen
;                                                                   	den Nutzer vom Zugriff auf Albis blockieren. Erreicht wird dies z.B. durch folgende Regel:
;                                                                	- 	Daten die aus einem Skript in das Albisfenster übergeben werden, werden sofort wieder ausgelesen
;                                                                   	um den Erfolg der Übergabe zu überprüfen.
;
;        - Fensterhandling                               	 -	Fehlerausgaben oder Hinweisfenster werden ausgewertet und es wird entsprechend reagiert
;
;        - Ereigniserkennung -                            -	ein großer Teil der Funktionalität auf dem Einsatz von Hooks. Hooks sind eine Art Systemrückruf ("Callback")
;                                                         			  	in Windows um Nachrichten in anderen Prozessen/Programmen "Callback's" mitlesen/abfangen zu können.
;                                                                	  	Diese Technik minimiert die CPU-Belastung und ermöglicht praktisch sofort auf Veränderungen im Albisfenster zu reagieren.
;                                                        			  	Außerdem ermöglicht diese Technik eine wesentlich flexiblere Reaktion auf Ereignisse.
;
;        - Interaktion mit der Oberfläche			- 	die zwar relativ einfach einzusetzende, aber sehr unzuverlässige Simulation von Tasten- oder Mauseingaben von Autohotkey wird
;                                                               	  	in den Automatisierungsfunktionen, soweit es möglich ist, nicht eingesetzt. Der Aufruf z.B. des Kassenrezept-Formulares erfolgt
;                                                                	  	nicht durch das Senden von Tastaturkürzeln, sondern erfolgt über das Senden von Nachrichten-ID's (siehe Send- oder Postmessage)

;                                                                	  	Eine  weitere Technik  bei RPA-Software  ist die  Verwendung  von Pixelsuch- oder Bildvergleichsfunktionen (z.B. Sikuli). Auch diese
;                                                               	  	Technik  wird nicht  eingesetzt.   Für eine neuere Technik (Optical Character Recognition), d.h. Texte oder Textbereiche zu erkennen
;                                                                  	  	fehlt es Autohotkey an  Geschwindigkeit. Andere Software nutzt dabei meist Tesseract von Google. Meine eigenen Versuche   mit
;                                                               	  	Tesseract haben eine relativ geringe Verarbeitungsgeschwindigkeit gezeigt und das bei zu geringer Genauigkeit der Zeichenerkennung.
;

;                                                                	  	Diese Techniken  sind aber auch  nicht notwendig, da die  Albisoberflächen  über interne  Windowsfunktionen gezeichnet werden.
;                                                                	  	Microsoft hat gute Funktionen für den Zugriff auf Oberflächen bereitgestellt und Autohotkey ist genau darauf spezialisiert.
;                                                                	  	Somit ist die Interaktion  mit fast allen  Elementen der  Albisoberfläche  und sämtlichen  Albisfenstern nach  kurzer Einarbeitung in
;                                                                 	  	diese Skriptsprache möglich.
;
;	-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -    -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
;
;																																												                  				IXIKO 2021
;
;	-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  - -  -  -  -  -  -  -  -  -  -  -  -  -  - ;}

Init_Albis() {                                                                                       	;-- initialisiert wichtige Konstanten

	; manche Konstanten werden von verschiedenen Funktionen genutzt.
	; ändert die Compugroup die Klassenbezeichnungen von Steuerelementen muss die Zuordnung nicht in allen Funktionen korrigiert werden

	static _ := Init_Albis()           ; static wird zur Ladezeit ausgeführt und ruft sein Funktion auf

	global KK            	:= {"Arzt":"Edit4", "Datum":"Edit5", "Kuerzel":"Edit6", "Inhalt":"RichEdit20A1"}
	global KKrv       	:= {"Edit4":"Arzt", "Edit5":"Datum", "Edit6":"Kuerzel", "RichEdit20A1":"Inhalt"}    ; reversed
	global afxMDI   	:= "AfxMDIFrame140"
	global afxView   	:= "AfxFrameOrView140"
	global IdentData	:= ["Dokument", "Edit4,Edit5,Edit6,RichEdit20A1", "#32770", "AfxFrameOrView", "AfxMDIFrame"]

}

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; HAUPTFENSTER INFO's                                                                                                                                                                  	(05)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetActiveWinTitle          	(02) AlbisWinID                         	(03) AlbisPID                                 	(04) AlbisExist
; (05) AlbisStatus
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisGetActiveWinTitle() {						                                    	;-- Fenstertitel v. Albishauptfenster
	WinGetTitle, iWT, ahk_class OptoAppClass
	RegExMatch(iWT, "(?<=\[).*(?=\])", Match)
	For key, val in StrSplit(Match, "/") 			     					    		;entfernt überflüssige Spacezeichen
		LWT .= Trim(val) "/"
return RTrim(LWT, "/")
}
;02
AlbisWinID() {                            			                                    	;-- gibt die ID des übergeordneten Albisfenster zurück

	; letzte Änderung: 07.02.2021
	Loop
		If (AID := WinExist("ahk_class OptoAppClass"))
			return GetHex(AID)
		else If (A_Index > 40)
			return 0
		else if (A_Index <= 40)
			sleep, 20

return GetHex(AID)
}
;03
AlbisPID() {                            				                                    	;-- gibt die Prozeß-ID des Albisprogrammes aus
	WinGet, AlbisPID, PID, % "ahk_class OptoAppClass"
return AlbisPID
}
;04
AlbisExist() {                            				                                    	;-- schaut nach ob Albis läuft, gibt wahr oder falsch zurück
	;result = 0 wenn kein Albis Prozess läuft , andernfalls die PID(Prozess-ID) des Albis Prozesses
	result:= WMIEnumProcessExist("Albis")
	result+= WMIEnumProcessExist("AlbisCS")
return result
}
;05
AlbisStatus() {                                                                            	;-- ermittelt ob Albis bereit ist um Befehle zu empfangen und gibt eine Fehlermeldung aus

	while !CheckWindowStatus(AlbisWinID(), 200)	{
		MsgBox,1, % "Addendum für Albis on Windows",
					(LTrim
						Albis scheint nicht zu reagieren!
						Bevor Sie weiter machen,
						überprüfen Sie bitte das Albisprogramm!
					)
		IfMsgBox, Cancel
		{
				MsgBox, 1, Addendum für Albis on Windows,
								(LTrim
								Wollen Sie wirklich das laufende
								Skript (%A_ScriptName%) abbrechen?
								Möglicherweise verlieren Sie dabei Daten.

								Vorschlag, starten Sie Albis erneut und
								drücken Sie dann auf OK.
								)
				IfMsgBox, Cancel
						ExitApp
		}
}

}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; MDI CONTROL FUNKTIONEN                                                                                                                                                      	(13)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisMDIClientHandle          	(02) AlbisMDIWartezimmerID      	(03) AlbisMDIClientWindows       	(04) AlbisMDIChildActivate
; (05) AlbisMDIChildHandle           	(06) AlbisMDIChildHandle           	(07) AlbisMDIChildTitle              	(08) AlbisMDIChildWindowClose
; (09) AlbisMDIMinMaxStatus         	(10) AlbisMDITabActive              	(11) AlbisMDITabActivate             	(12) AlbisMDITabHandle
; (13) AlbisMDITabNames
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01    	;-- MDI Control Funktionen
AlbisMDIClientHandle() {                                                           	;-- ermittelt das Handle des MDIClient (Basishandle für alle Unterfenster)
	; letzte Änderung: 10.02.2021
	ControlGet, hMdi, HWND,, MdiClient1, ahk_class OptoAppClass
return hMdi
}
;02
AlbisMDIWartezimmerID() {                                                        	;-- ermittelt das Handle des Wartezimmerfenster innerhalb des MDI-Controls
return FindChildWindow({"Class": "OptoAppClass"}, {"Title": "Wartezimmer"}, "Off")				;ID des Wartezimmer-Fenster
}
;03
AlbisMDIClientWindows() {                                                         	;-- globales Objekt, welches Namen, Klassen und Handles aller geöffneten MDI Fenster enthält

	; Mdi	- [key : value]	- Object > global Mdi - im aufrufenden Skript / aufrufender Funktion einfügen!
	;         	- key             	- ist WinTitle
	;       	- value          	- ist die Fensterklasse, extra Key ist "MdiHwnd" für das Roothandle des MDI-Fenster
	;
	; letzte Änderung: 10.02.2021

	global Mdi
	Mdi := Object()   ; das Mdi-Objekt wird nur von dieser Funktion zurückgesetzt

	WinGet, MDIClientWinList, ControlListHWND, % "ahk_id " AlbisMDIClientHandle()
	Loop, Parse, MDIClientWinList, `n
		If InStr(class := WinGetClass(A_LoopField), "Afx:")
			Mdi[WinGetTitle(A_LoopField)]:= {"class": class, "ID": A_LoopField}

return Mdi
}
;04
AlbisMDIChildActivate(MDITitle) {		                                        	;-- aktiviert ein MDI-Child Fenster (nicht Tab im MDI-Child!)

	; MDITitle: Name(Titel) des MDI child oder das Handle (Dezimal oder Hexadezimal)
	; return value: bei Erfolg das Handle des MDI-Child ansonsten ein Objekt mit dem ErrorLevel, dem letzten Fehler und das MDIClient-Handle
	; letzte Änderung: 20.02.2021

	If RegExMatch(MDITitle, "i)^\s*(?<wnd>(0x)[0-9A-F]+|[\d]+)\s*$", h)
		hMdiChild := hwnd
	else
		hMdiChild	:= AlbisMDIChildHandle(MDITitle)

	SendMessage, 0x222, % hMdiChild,,, % "ahk_id " (hMdi := AlbisMDIClientHandle())

return (ErrorLevel = 0 ? hMdiChild : {"ErrorLevel": ErrorLevel, "LastError": A_LastError, "hMdi": hmdi})
}
;05
AlbisMDIChildGetActive() {                                                        	;-- ermittelt das Handle des aktuellen MDI-Childfensters
	; letzte Änderung: 10.02.2021
		SendMessage, 0x0229,,,, % "ahk_id " AlbisMDIClientHandle() 	; WM_MDIGETACTIVE:= 0x0229
return GetHex(ErrorLevel)
}
;06
AlbisMDIChildHandle(MDITitle) {		                                        	;-- ermittelt das Handle eines sub oder child Fenster innerhalb des Albis-MDI-Controls
return GetHex(FindChildWindow({"Class": "OptoAppClass"}, {"Title": MDITitle}, "Off"))
}
;07
AlbisMDIChildTitle(hMdiChild) {                                                  	;-- gibt den Namen (Text) eines MDI Childfensters zurück
	ControlGetText, MDITitle,, % "ahk_id " hMdiChild
return MDITitle
}
;08
AlbisMDIChildWindowClose(MDITitle) {                                        	;-- schließt ein per Titel identifiziertes MDIClient Fenster

	; MDITitle - darf ein Teil des Namens, die komplette Patienten-ID, Geburtsdatum, .... sein

	; wird nur ausgeführt wenn der Titel gefunden werden kann
		If (hMdiChild := AlbisMDIChildHandle(MDITitle)) 	{
			SendMessage, 0x0221, % hMdiChild,,, % "ahk_id " AlbisMDIClientHandle() ;WM_MDIDESTROY
			return ErrorLevel ? 1 : 0
		}
		else
			return 0

return 1
}
;09
AlbisMDIMinMaxStatus(MDITitle) {                                               	;-- stellt fest ob das gewählte MDI Fenster maximiert und im Vordergrund ist

	; MDITitle - Titel des gesuchten MDI Fenster
	; gibt 1 zurück wenn es maximiert ist und im Vordergrund , Beispiel: AlbisMDIMinMaxStatus("Wartezimmer")
	; letzte Änderung: 10.02.2021

	global MDI

	MDI := AlbisMDIClientWindows()

	If RegExMatch(MDITitle, "^0x[\0-9A-F]+") {
		For title, mditab in MDI
			If InStr(mditab.id, MDIChild)
				return DllCall("IsZoomed", "UInt", val.ID)
	}
	else 	{
		For title, mditab in MDI
			If InStr(mditab, MDIChild)
				return DllCall("IsZoomed", "UInt", val.ID)
	}

}
;10                                                        > MDI Tab Funktionen <
AlbisMDITabActive(MDITitle, ret:="Index") {                                 	;-- gibt die Nummer oder den Titel des aktiven Tab zurück
	SendMessage, 0x130B,,,, % "ahk_id " (hMdiTab := AlbisMDITabHandle(MDITitle))
	tabIndex := ErrorLevel
	If (ret = "Index")
		return tabIndex
	else if (ret = "Name") {
		tabs := AlbisMDITabNames(MDITitle)
		return tabs[tabIndex+1]
	}
}
;11
AlbisMDITabActivate(MDITitle, TabName:="") {                             	;-- aktiviert ein MDI-Tab, merkt sich welches TabControl zuletzt aufgerufen wurde

	; ruft man die Funktion mehrfach nacheinander unter demselben MDITitle auf, ermittelt er nicht jedesmal alle Handles neu (Speicherung in tabs - Object)
	; Beispiel: result:= AlbisMDITabActivate("Wartezimmer", "Arzt")
	; Rückgabewert ist 0 wenn er keinen passenden Tab finden konnte, >0 wenn die Funktion erfolgreich war
	; MDITitle: darf ein Fensterhandle sein
	; letzte Änderung: 29.03.2021

		static last_MDITitle, hMDITab

	; keinen TabNamen übergeben
		If (StrLen(TabName) = 0)
			return 0

	; Änderung des MDI Fenstertitels?
		If (last_MDITitle <> MDITitle) {
			last_MDITitle := MDITitle
			hMDITab	:= AlbisMDITabHandle(MDITitle)
		}

	; kein handle gefunden
		If !hMDITab
			return 0

	; Position des Tab finden und dann fokussieren
		For idx, tab in AlbisMDITabNames(MDITitle)
			If InStr(tab, TabName)
				SendMessage, 0x1330, % idx - 1 ,,, % "ahk_id " hMDITab     ; 0x1330 is TCM_SETCURFOCUS.

return ErrorLevel
}
;12
AlbisMDITabHandle(MDITitle) {                                                     	;-- ermittelt das Handle eines spezifischen MDI-TabControls

	; MDITitle kann der Fenstertitel oder das Handle sein
	; letzte Änderung: 29.03.2021

	hSpecMDI	:= RegExMatch(MDITitle, "i)^(0x[A-F\d]+|[\d]+)$") ? MDITitle : GetHex(AlbisMDIChildHandle(MDITitle))
	oCtrl     	:= GetControls(hSpecMDI)
	If (key := ObjFindValue(oCtrl, "SysTabControl32"))
		return oCtrl[key].Hwnd

	PraxTT(A_ThisFunc " (" (A_LineNumber - 3) "): SysTabControl32 nicht gefunden:`n" MDITitle, "2 1")

return 0
}
;13
AlbisMDITabNames(MDITitle) {                                                   	;-- ermittelt die Namen aller Tabs eines SysTabControls321

  ; Funktion kann ohne die anderen MDI Funktionen aufgerufen werden
  ; MDITitle: darf ein Fensterhandle sein
  ; Abhängigkeiten: GetHex(), GetControls, ObjFindValue(), ControlGetTabs()
  ; letzte Änderung: 29.03.2021

	hSpecMDI	:= RegExMatch(MDITitle, "i)^(0x[A-F\d]+|[\d]+)$") ? MDITitle : GetHex(AlbisMDIChildHandle(MDITitle))
	oCtrl     	:= GetControls(hSpecMDI)
	key			:= ObjFindValue(oCtrl, "SysTabControl32")
	If (hMdiTab := oCtrl[key].Hwnd)
		return ControlGetTabs(hMdiTab)

	PraxTT(A_ThisFunc " (" (A_LineNumber - 3) "): SysTabControl321 nicht gefunden:`n" MDITitle, "2 1")

return
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; INFO's VON ANDEREN STEUERELEMENTEN                                                                                                                                   	(03)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetActiveControl           	(02) AlbisGetActiveWindowType   	(03) AlbisLVContent                      	(04) #AlbisGetStammPos
; (05) #AlbisAktuellesKuerzel
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisGetActiveControl(cmd, matchCase:="") {                                    	;-- ermittelt den aktuellen Eingabefocus und gibt zusätzliche Informationen zur Identifikation zurück

	/* 		FUNKTIONSBESCHREIBUNG: AlbisGetActiveControl() letzte Änderung: 17.06.2021
	   *
		* Steuerelemente in Fenster können die gleichen Namen besitzen. Um ein Steuerelement exakt zu identifizieren ermittelt diese Funktion die Eltern(Parent)-Elemente
		* wieviele Eltern ermittelt werden müssen, hängt von der jeweiligen Fensterstruktur ab
		*
		* gedacht für Contextabhängige Hotstrings, Hotkeys oder andere Dateneingaben zu ermöglichen
		*
		* --------------------------------------------------------------------------------------------------------------------------------
		* Bezeichner	|            Ebenenelemente      	|                          Elternelemente						   	   	        	*
		* Bezeichner	|      ClassNN1,ClassNN2,... 	|ParentClass1|ParentClass2			|ParentClass3                	*
		* Dokument	|Edit4,Edit5,Edit6,RichEdit20A1	|#32770		 |AfxFrameOrView	|AfxMDIFrame              	*
		* --------------------------------------------------------------------------------------------------------------------------------
		* 		Bezeichner 			- ihre Bezeichnung für den Steuerelement-Zusammenhang
		*		Ebenenelemente	- eine durch Komma getrennte Liste an ClassNN Namen von Steuerelementen, diese Steuerelemente haben einen inhaltlichen Zusammenhang z.B.
		*		Elternelemente		- eine durch | getrennte Liste an ClassNN Namen von Eltern-Steuerelementen, diese lassen die eindeutige Identifikation der Ebenenelemente zu
		*
		* --------------------------------------------------------------------------------------------------------------------------------
		* Parameter
		* --------------------------------------------------------------------------------------------------------------------------------
		*   	cmd                      - hwnd, classNN, identify, content, contraction, ContractionIsEqual
		*		matchCase         	- Vergleichsstring für Abgleich mit Zeilenkürzel
		*
		* - 11.07.2020: AfxFrameOrView und AfxMDIFrame wurden intern in Albis umbenannt, aus AfxFrameOrView90 und AfxMidiFrame90 wurde AfxFrameOrView140 und AfxMidiframe140
		* - 17.06.2021: Sourcecode modernisiert, hat dadurch vielleicht auch etwas Geschwindigkeit gewonnen
		* - 27.06.2021: die Definition der Steuerelement Klassenbezeichnungen ausgelagert um Änderungen durch die Compugroup leichter anpassen zu können
	   *
		;static IdentData	:= ["Dokument", "Edit1,Edit2,Edit3,RichEdit20A1", "#32770", "AfxFrameOrView", "AfxMDIFrame"]		;Dokument ist meine Bezeichnung für die 4-Controls (Edit1...)
	 */

		global IdentData, KK, KKrv

	; Ermitteln von handle und classnn des aktuellen fokussierten Steuerelementes im Albisfenster
		hFocused	:= Chwnd := GetFocusedControlHwnd()
		cNN	    	:= Control_GetClassNN(AlbisWinID(), hFocused)

		If InStr(cmd, "hwnd")
			return GetHex(ChWnd)
		else if InStr(cmd, "classNN")
			return cNN

	; Übereinstimmungen suchen
		hFirstParent := DllCall("user32\GetAncestor", "Ptr", hFocused, "UInt", 1, "Ptr")
		If InStr(IdentData[2], cNN)		{

			; die ersten zwei Elemente gehören nicht zur Fensteridentifikation
				aacCounter := 0
				Loop % (IdentData.MaxIndex() - 2) {
					Chwnd := DllCall("user32\GetAncestor", "Ptr", Chwnd, "UInt", 1, "Ptr")
					If InStr(WinGetClass(CHwnd), IdentData[A_Index + 2])
						aacCounter++
				}

			; keine Übereinstimmungskette gefunden
				If !(aacCounter = IdentData.MaxIndex() - 2)
					return false

			; ab hier können spezielle weitere Befehle programmiert werden
				Switch cmd            {

				  ; Inhalt der alle Steuerelemente also der gesamten Zeile zurückgeben
					Case "content":
						If InStr(IdentData[1], "Dokument") {
							ControlGetText, Arzt  	, % KK.Arzt    	, % "ahk_class OptoAppClass"    	; "ahk_id " hFirstParent
							ControlGetText, Datum	, % KK.Datum	, % "ahk_class OptoAppClass"    	; "ahk_id " hFirstParent
							ControlGetText, kuerzel	, % KK.kuerzel	, % "ahk_class OptoAppClass"    	; "ahk_id " hFirstParent
							ControlGetText, inhalt	, % KK.inhalt 	, % "ahk_class OptoAppClass"    	; "ahk_id " hFirstParent
							return {	"FHwnd"         	: GetHex(hFocused)
									, 	"FClassNN"   	: cNN
									,	"FName"          	: KKrv[cNN]
									, 	"fpar"            	: GetHex(hFirstParent)
									, 	"fparClassNN"  	: GetClassNN(hFirstParent, AlbisWinID())
									,	"Identifier"      	: IdentData[1]
									, 	"Edit1"            	: Arzt                ; Edit1,2,3,Rich... aus Kompatibilitätsgründen mit meinen älteren Skripten behalten
									, 	"Edit2"            	: Datum
									, 	"Edit3"           	: kuerzel
									, 	"RichEdit"       	: inhalt
									, 	"Arzt"             	: Arzt
									, 	"Datum"           	: Datum
									, 	"Kuerzel"          	: kuerzel
									, 	"Inhalt"          	: inhalt}
						}

				  ; für Kürzelvergleiche (Feststellung des richtigen Eingabezusammenhanges)
					Case "contraction":
						return ControlGetText(KK.Kuerzel, "ahk_class OptoAppClass")       ; war vorher Edit3

				  ; für Kürzelvergleiche (Feststellung des richtigen Eingabezusammenhanges)
					Case "ContractionIsEqual":
						contraction := ControlGetText(KK.Kuerzel, "ahk_class OptoAppClass")        ; war vorher Edit3
						return (contraction = matchCase) ? true : false

				  ; für Kürzelvergleiche (Feststellung des richtigen Eingabezusammenhanges)
					Case "identify":                              	; identifizieren
						return IdentData[1] "|" cNN

				}

		}

return false
}
;02
AlbisGetActiveWindowType(CBReturn=false) {                                	;-- ermittelt die aktuelle Ansicht Wartezimmer, Terminplaner, Akte, ....

	; diese Funktion kategorisiert die Art des von Albis angezeigten Inhaltes
	; Patientenakte ("Akte"), Wartezimmer ("WZ") , freie Statistik (fS), Tagesprotokoll, Terminplaner oder nicht in der unteren Liste dann ("other")
	; CBReturn: gibt den vollständigen Eintrag der ComboBox zurück z.B. um nach automatisiertem Schalten der Auswahl die ursprüngliche Ansicht wiederherzustellen
	; keine Unterscheidung ob ein PopUp Dialog z.B. Rezeptformular geöffnet ist, da es in dieser Funktion ausschließlich um den Inhalt des Fenstertitels geht
	;
	; letzte Änderung: 01.08.2021

	WinGetTitle, LWT, % "ahk_class OptoAppClass"

    If InStr(LWT, "Wartezimmer")
		return "WZ"
	else if InStr(LWT, "freie Statistik")
		return "fS"
	else if InStr(LWT, "Prüfung EBM/KRW")
		return "PEK"
	else if InStr(LWT, "Tagesprotokoll")
		return "TProt"
	else if InStr(LWT, "Terminplaner")
		return "TP"
	else if RegExMatch(LWT, "\d\d\.\d\d\.\d\d\d\d")	{
		ControlGet, hTbar		, Hwnd		,, ToolbarWindow321	, % "ahk_id " AlbisMDIChildGetActive()
		ControlGet, Auswahl	, Choice	,, ComboBox1				, % "ahk_id " hTbar
		If CBReturn
			return Auswahl
		If InStr(Auswahl, "Abrechnung")
			return "Karteikarte|Patientenakte|Abrechnung"
		else If RegExMatch(Auswahl, "i)^P\s+(Priva|Stand)\s")
			return "Karteikarte|Patientenakte|Privatabrechnung"
		else
			return "Karteikarte|Patientenakte|" Auswahl
	}

return "other"
}
;03
AlbisLVContent(hWin, LVClass:="SysListView321", Column:="1") {     	;-- Listview eines Fenster auslesen (Universalfunktion)

	;Achtung das Auslesen eines Listview Elements funktioniert bisher nur für das Wartezimmer

		/*             	Beschreibung

				neuer Versuch für eine relativ universell einsetzbare Funktion zum Auslesen von Listviewfenstern
				es wird ein Objekt so erzeugt das man per Übergabe von Zeile und Spalte direkt auslesen kann (inhalt := LVRow[Zeile][Spalte])
		*/

		LVRow := Object()

		ControlGet, result, List,, % LVClass, % "ahk_id " hWin	;Col%A_Index%

		Loop, Parse, result, `n
			If (StrLen(A_LoopField) > 0) {

				        	rowNr	:= A_Index
				LVRow[rowNr] 	:= Object()
				            	col 	:= StrSplit(A_LoopField, "`t")

				Loop, % col.MaxIndex()
					LVRow[rowNr].Push(col[A_Index])

			}

return LVRow
}
;04
AlbisGetStammPos() {                                                                    	;-- ##liest aus der Local.ini die Positionen der Stammdatenfenster aus

	; Abhängigkeiten: \lib\ini.ahk

	obj := Object()

	If !FileExist(LocalIni := "C:\albiswin\Local.ini")
		If !FileExist(LocalIni := "C:\albiswin.loc\Local.ini") {
			MsgBox, % "Addendum für AlbisOnWindows", % 	"Die Local.ini befindet sich in keinem der`n"
														        					. 	"Standard-Ordner auf dem Laufwerk (C:)!`n"
																        			.	"Das Skript (" A_ScriptName ") wird beendet."
				ExitApp
	}

	;ini_Load(ini, LocalIni)
	;keys := ini_getAllKeyNames(Ini, "STAMM_POS")

	Loop, Parse, keys, `,
	{
		;~ If InStr(val := ini_getValue(ini, "STAMM_POS", A_LoopField), ",") {
			;~ val := StrSplit(StrReplace(val, " ") , ",")
			;~ obj[A_LoopField] := {"x":val.1, "y":val.2, "w":val.3, "h":val.4}
		;~ }

	}

return obj
}
;05
AlbisAktuellesKuerzel() {                                                                 	;°~° begonnen

		global KK

		aControl := {}

		If !InStr(AlbisGetActiveWindowType(), "Karteikarte")
			return ""

		ControlGetFocus, hactiveCtrl, ahk_class OptoAppClass
		If !InStr(hactiveCtrl, KK.inhalt)
			return 0

		aControl:= AlbisGetActiveControl("content")

}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; DATEN VOM AKTUELLEN PATIENTEN                                                                                                                                             	(09)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAktuellePatID                 	(02) AlbisCurrentPatient              	(03) AlbisPatientGeschlecht           	(04) AlbisPatientGeburtsdatum
; (05) AlbiPatientVersicherung         	(06) AlbisVersArt                        	(07) AlbisTitle2Data                      	(08) AlbisAbrechnungsscheinVorhanden
; (09) AlbisAbrechnungsscheinAktuell
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01		;-- Patienten bezogene Informationen
AlbisAktuellePatID() {                                                                 	;-- liest aus dem Fenstertitel nur die PatientenID aus
	RegExMatch(AlbisGetActiveWinTitle(), "O)^\d+", Match)
return Trim(Match.Value())
}
;02
AlbisCurrentPatient()	{                                                                	;-- sucht den aktuellen Patientennamen raus
return Trim(StrSplit(AlbisGetActiveWinTitle(), "/").2)
}
;03
AlbisPatientGeschlecht() {                            					     			;-- entnimmt dem Fenstertitel das Geschlecht
return Trim(StrSplit(AlbisGetActiveWinTitle(), "/").3)
}
;04
AlbisPatientGeburtsdatum() {					                                    	;-- entnimmt dem Fenstertitel das Geburtsdatum
return Trim(StrSplit(AlbisGetActiveWinTitle(), "/").4)
}
;05
AlbisPatientVersicherung() {                                                        	;-- die Versicherung ist die letzte Angabe im Fenstertitel
return Trim(StrSplit(AlbisGetActiveWinTitle(), "/").5)
}
;06
AlbisVersArt() {                             			                                    	;-- ermittelt die Versicherungsart (GKV, PKV)
;gibt eine 1 für privat und eine 2 für Kassenversichert zurück
return (Instr(StrSplit(AlbisGetActiveWinTitle(), "/").5, "Privat") ? 1 : 2)
}
;07
AlbisTitle2Data(WTAlbis:="") {                            				         	;-- erstellt ein Objekt aus den Daten des Fenstertitel

	;WTAlbis leer lassen dann liest die Funktion den Fenstertitel
	If !InStr(AlbisGetActiveWindowType(), "Akte")
		return ""

	If (StrLen(WTAlbis) = 0)
		WTAlbis := AlbisGetActiveWinTitle()

	Splitstr				:= StrSplit(WTAlbis, "/", A_Space)
	SplitStr2			:= StrSplit(SplitStr[2], ",", A_Space)

return {"ID": SplitStr[1], "Nn": SplitStr2[1], "Vn": SplitStr2[2], "Gt": SplitStr[3], "Gd": SplitStr[4], "Kk": SplitStr[5]}
}
;08
AlbisAbrechnungsscheinVorhanden(param) {                              	;-- gibt Index eines gesuchten Abrechnungsscheines in ComboBox zurück

	; Parameter:      	param	- 	kann als Quartal im Format [0]1/17 übergeben werden für Kassenabrechnungsscheine
	;                                   	-	wird kein Quartal übergeben wird automatisch von der Übergabe eines RegEx-Suchmusters
	;                                      		ausgegangen z.B. "Privat.*?BE"   (eine Rechnung welche noch nicht gedruckt wurde)
	;
	; Rückgabewert:  	ist die Zeilennummer in der ComboBox in welcher der passende Eintrag gefunden wurde
	;
	; letzte Änderung: 17.06.2021

	; Combox mit Abrechnungsscheinen auslesen
		ControlGet, CBox, List,, ComboBox1, % "ahk_id " AlbisWinID()
		If (StrLen(CBox) = 0) || (StrLen(param) = 0)
			return 0

	; je nach übergebenem Suchparameter die Combobox durchsuchen
		If RegExMatch(param, "^\d+\/\d+$") {

			Quartal:= LTrim(param, "0")
			Quartal:= RegExReplace(Quartal, "[^\d]")
			Quartal:= "(" SubStr(Quartal, 1, 1) "/" SubStr(Quartal, 2, 2) ")"

			For CBPos, zeile in StrSplit(CBox, "`n")
				if ( Instr(zeile, Quartal) || InStr(zeile, "aktuell") )
					return CBPos

		}
		else {

			For CBPos, zeile in StrSplit(CBox, "`n")
				If RegExMatch(zeile, param)
					return CBPos

		}

return 0
}
;09
AlbisAbrechnungsscheinAktuell() {            	                 		         	;-- aktuelle Art des Abrechnungsscheines (Abrechnung, Privat, Notfall, Überweisung ...)

	; gibt alle noch nicht abgerechneten Abrechnungsscheinarten zurück

	static Scheinarten 	:= {"Abrechnung":"^Abrechnung", "Überweisung":"^Überweisung", "Belegarzt":"^Belegarzt", "Notfall":"^Notfall", "Privat":"^P\sPriva.*?BE" }
	AbrSchein := Array()

	ControlGet, CBox, List,, ComboBox1, % "ahk_id " AlbisWinID()			;ahk_class OptoAppClass
	CBLines := StrSplit(CBox, "`n", "`r")
	For CBIndex, line in CBLines {

		For Scheinart, rxASAString in Scheinarten
			If RegExMatch(line, rxASAString) {

				If (Scheinart <> "Privat") {
					RegExMatch(line, "(?<Q>\d\/\d+)", Abr)
					AbrSchein.Push({"Scheinart":Scheinart, "Quartal":AbrQ})
					break
				}
				else If (Scheinart = "Privat") {
					RegExMatch(line, "\d\d\.\d\d.\d\d\d\d)", PDatum)
					AbrSchein.Push({"Scheinart":Scheinart, "Datum":PDatum})
					break
				}

			}

	}

return AbrSchein
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; FENSTERELEMENTE AUSLESEN UND BEHANDELN                                                                                                                          	(14)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetCaveZeile               	(02) AlbisGetCave                     	(03) AlbisCaveAddRows                    	(04) AlbisCaveGetCellFocus
; (05) AlbisCaveGetCellFocus      	(06) AlbisCaveUnFocus               	(07) AlbisSetCaveZeile                      	(08) )AlbisMuster30ControlList()
; (09) AlbisOptPatientenfenster    	(10) AlbisHeilMittelKatologPosition	(11) AlbisSortiereDiagnosen              	(12) AlbisReadFromListbox
; (13) AlbisResizeDauerdiagnosen 	                                                  	(14) AlbisResizeLaborAnzeigegruppen
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisGetCaveZeile(nr, SuchString="", CloseCave=false) {             	;-- Auslesen einer oder aller Zeilen im Cave! von Dialog

	; letzte Änderung: 01.10.2020 - Suchstring kann ein RegEx-String sein

	; CloseCave=true Fenster wird anschließend nicht geschlossen

	; Cave! von Fenster aufrufen
		Albismenu(32778, "Cave! von ahk_class #32770", 4)

	; Auslesen aller Zeilen des Cave! von - Dialoges
		ControlGet, result, List, Col4, SysListView321, % "Cave! von ahk_class #32770"
		CaveItems := StrSplit(result, "`n")
		itemfound := false
		Zeile      	:= ""

	; ein SuchString wurde übergeben
		If (StrLen(SuchString) > 0) {

			For nrZeile, Zeile in CaveItems
				If RegExMatch(Zeile, SuchString)	{
					itemfound := true
					If RegExMatch(SuchString, "i)GVU|HKS|KVU") && (nrZeile <> 9) {        ; nur für GVU -> Verschieben auf Zeile 9
						AlbisSetCaveZeile(nrZeile, ""   	, false)
						AlbisSetCaveZeile(9     	 , Zeile	, false)
						nrZeile := 9
					}
					break
				}

			If !itemfound
				zeile := "", nrZeile := 0

		}
		else {
			Zeile := CaveItems[nr]
		}

	; wenn CloseCave = true ist, wird der Cave! von - Dialog geschlossen
		If CloseCave	{

			WinActivate	 , % "ahk_id " CaveVonID
			WinWaitActive, % "ahk_id " CaveVonID,, 1
			while (CaveVonID := WinExist("Cave! von")) && (A_Index <= 45)	{
				VerifiedClick("Button4", CaveVonID)    	; ich wählte Button4 = Abbrechen ; Programmfehler können so keinen Schaden anrichten, alternativ ginge es auch mit zweimal Escape hintereinander
				Sleep, 70
				If (A_Index > 40)               					;es dauert auch mal länger als 4s bis das Fenster erscheint - 40x100 ms in diesem Fall
					MsgBox, 0x40000, Frage, % "Es wurde mehrmals probiert den 'Cave! von' Fenster zu schließen`nBitte schließen Sie das Fenster von Hand.`nIm Anschluß läuft diese Funktion automatisch weiter."
			}

		}

return Trim(Zeile)
}
;02
AlbisGetCave(CloseCave=true) {       		                                  	;-- alle Zeilen des Cave!von Fenster auslesen

		AlbisActivate(1)

	; Cave! von Fenster aufrufen
		If !(CaveVonID := Albismenu(32778, "Cave! von ahk_class #32770", 4))
			return 0

	; Auslesen aller Zeilen des Cave! von - Dialoges
		ControlGet, result, List, Col4, SysListView321, % "ahk_id " CaveVonID

	; CaveVonFenster schließen
		If CloseCave {

			WinActivate	 , % "ahk_id " CaveVonID
			WinWaitActive, % "ahk_id " CaveVonID,, 1
			while (CaveVonID:= WinExist("Cave! von"))	&& (A_Index <= 30)	{
				VerifiedClick("Button4", "Cave! von")   	; wähle Button4 = Abbrechen
																			; Programmfehler können so keinen Schaden anrichten
																			; alternativ geht auch mit zweimal Esc hintereinander
				Sleep, 70
				If (A_Index > 10)
					MsgBox	, 0x40000
								, Frage
								, % 	"Es wurde mehrfach versucht das Cave! von Fenster zu schließen`n"
									. 	"Bitte schließen Sie das Fenster von Hand.`n"
									.	"Im Anschluß läuft diese Funktion automatisch weiter."
			}

		}

return result
}
;03
AlbisCaveAddRows(hCave, RowsToAdd) {                                    	;-- fügt Cave! von neue Zeilen hinzu

	ControlGet, hCaveLV, HWND,, SysListView321, % "ahk_id " hCave
	ControlGet, rowCount, List, Count,, % "ahk_id " hCaveLV
	MaxRow := rowCount + RowsToAdd

	If !AlbisCaveSetCellFocus(hCave, "last", "last")
		return 0

	Table := AlbisCaveGetCellFocus(hCave)

	SetKeyDelay, 30, 40
	SendMode, Input

	Loop {

		ControlGet, hEdit, HWND,, Edit1, % "ahk_id " hCave
		If !(EditVisible := DllCall("IsWindowVisible","Ptr", hEdit))
			return "-1"
		ControlFocus	, Edit1, % "ahk_id " hCave
		sleep 50
		Send, {TAB}
		sleep 200
		ControlFocus	, Edit1, % "ahk_id " hCave
		ControlGetFocus, FControl, % "ahk_id " hCave
		Table := AlbisCaveGetCellFocus(hCave)
		;ToolTip, % A_Index ": r" Table.row ", c" Table.col ", hEdit: " hEdit ", FC: " FControl
		If !IsObject(Table)
			return "-2"

	} until (Table.row = MaxRow)

}
;04
AlbisCaveGetCellFocus(hCave) {                                                   	;-- ermittelt die aktuell ausgewählte Reihe und Spalte

	; letzte Änderung: 10.07.2021
	; ein Eingabefocus ist ein Edit-Steuerelement (nennt sich immer Edit1)
	; die Funktion errechnet die Reihe und Tabellenspalte in der sich gerade das Edit-Steuerlement befindet

		CCol := Array()

	; Listview Daten auslesen
		ControlGet, hCaveLV     	, HWND,, SysListView321	, % "ahk_id " hCave
		ControlGet, hEdit           	, HWND,, Edit1                	, % "ahk_id " hCaveLV
		If !hEdit
			return 2
		ControlGet, RowIndex    	, List, Count Focused,      	, % "ahk_id " hCaveLV
		ControlGet, LVColCount	, List, Count Col,            	, % "ahk_id " hCaveLV

	; Spaltenbreite und x Position errechnen
		Loop % LVColCount {
			W	:= DllCall("SendMessage", "uint", hCaveLV, "uint", 4125, "uint", A_Index-1, "uint", 0)
			X	:= A_Index = 1 ? W : CCol[A_Index-1].X + CCol[A_Index-1].W
			CCol.Push({"X":X, "W":W})
		}

	; Position und Breite des Edit1 Steuerelementes ermitteln
		ControlGetPos, EditX,, EditW,, Edit1, % "ahk_id " hCaveLV

	; Spaltenposition des Edit-Steuerelementes errechnen
		For ColIndex, colpos in CCol
			If (EditX > colpos.X) && (EditX < colpos.X+colpos.W)
				return {"row":RowIndex, "col":ColIndex+1}

return -1
}
;05
AlbisCaveSetCellFocus(hCave, row, column) {                               	;-- setzt den Eingabefokus in Zeile und Spalte

	; Parameter: wenn row := "last", wird die letzte Zeile gewählt
	; letzte Änderung: 10.07.2021

	static CaveColumns := {"Arzt":2, "Datum":3, "Beschreibung":4}

	; Listviewdaten ermitteln
		ControlGet, hCaveLV 	, HWND,, SysListView321, % "ahk_id " hCave
		ControlGet, rowCount	, List, Count,, % "ahk_id " hCaveLV
		ControlGet, colCount	, List, Count Col,, % "ahk_id " hCaveLV

	; column in eine Zahl umwandeln
		If RegExMatch(column, "\d")
			col := column
		else if InStr(column, "last")
			col := colCount
		else
			col := CaveColumns[column]

	; row in Zahl umwandeln
		If InStr(row, "last")
			row := rowCount

	; sind überhaupt ausreichend Zeilen vorhanden?
		If (rowCount < row) || (colCount < col)
			return 2

	; wenn es ein Edit1 Steuerelement (Eingabe von Daten wäre möglich) wird es geschlossen
		If AlbisCaveUnFocus(hCave)
			return 3

	; Zeile (row) auswählen
		ControlFocus, SysListview321, % "ahk_id " hCave
		xRow 	:= Floor(row/10)
		nRow	:= row - xRow*10
		Loop xRow {                      ; einmal die Taste 1 = Zeile 1, das zweite mal ist die Zeile 10
			If (A_Index > 1)
				sleep 50
			ControlSend, SysListview321, % "{1}", % "ahk_id " hCave
		}
		If nRow                            	; dann noch den Rest senden
 			ControlSend, SysListview321, % "{" nRow "}", % "ahk_id " hCave

	; Kontrolle der Ausführung
		ControlGet, FocusedRow, List, Count Focused,, % "ahk_id " hCaveLV
		If (FocusedRow <> row)
			return 4

	; Spalte (col) wählen
		ControlSend,, {Space}, % "ahk_id " hCaveLV
		sleep 50
		Table := AlbisCaveGetCellFocus(hCave)
		If !IsObject(Table)
			return 5

		Loop {
			Table := AlbisCaveGetCellFocus(hCave)
			If !IsObject(Table)
				return 6
			If (Table.col = col)
				return 1
			ControlFocus	, Edit1, % "ahk_id " hCave
			ControlSend	, Edit1, {TAB}, % "ahk_id " hCave
			sleep 50
			Table := AlbisCaveGetCellFocus(hCave)
			If !IsObject(Table)
				return 7
			If (Table.col = col)
				return 1
		} until (Table.col = col)

return 1
}
;06
AlbisCaveUnFocus(hCave:="") {                                                   	;-- entfernt den Eingabefokus

	; letzte Änderung 10.07.2021

	hCave :=  !hCave ? WinExist("Cave! von ahk_class #32770") : hCave
	ControlGet, hCaveLV, HWND,, SysListView321, % "ahk_id " hCave
	ControlGet, hEdit   	, HWND,, Edit1, % "ahk_id " hCaveLV
	EditVisible := DllCall("IsWindowVisible","Ptr", hEdit)

	while (EditVisible && A_Index < 5) {
		If (A_Index > 1)
			sleep 300
		ControlFocus, SysListview321, % "ahk_id " hCave
		ControlGet, hEdit, HWND,, Edit1, % "ahk_id " hCaveLV
		EditVisible := DllCall("IsWindowVisible","Ptr", hEdit)
	}

return EditIsVisible
}
;07
AlbisSetCaveZeile(nr, txt, CloseCave=false) {	                             	;-- überschreibt eine Zeile im cave! von - Fenster

	; letzte Änderung 01.10.2020:
	;
	; 		- SendInput z.T. ersetzt durch ControlSend damit die Tastaturbefehle an das richtige Fenster gesendet werden
	;    	- fügt selbstständig fehlende Listenzeilen hinzu

		CaveVon 	:= "Cave! von ahk_class #32770"
		slTime   	:= 100
		AlbisActivate(1)

	; Cave! von Fenster aufrufen
		If !(hCave := Albismenu(32778, CaveVon, 4)) {
    		MsgBox, 0x40000, Addendum für Albis on Windows, % 	"Dialogfenster:`n 'Cave! von <" AlbisCurrentPatient() ">`nhat sich nicht aufrufen lassen!"
			return 0
		}

	; Cave! von Fenster aktivieren
		WinActivate, % CaveVon
		WinWaitActive, % CaveVon,, 2

	; Focus ins SysListView geben
		ControlGetFocus, FControl, % CaveVon
		while !InStr(FControl, "SysListview") 	|| !InStr(FControl, "Edit")	{
			ControlFocus 	 , SysListView321, % CaveVon
			ControlGetFocus, FControl     	 , % CaveVon
			If InStr(FControl, "SysListview") || !InStr(FControl, "Edit") || (A_Index > 10)
				break
			sleep % slTime * 2
		}

	; Verzögerung bei bestehender Remotedesktopverbindung (die Darstellung der Fenster ist dann langsamer)
		If InStr(FControl, "Edit") {
			SendInput, {Escape}
			sleep % slTime * 3
		} else {
			ControlClick, SysListView321, % CaveVon,, Left, 1
			sleep % slTime
		}

		SendInput, {Home}
		sleep % slTime*2

	; bei neue Patienten ohne Eintragungen im cave Dialog kann eine beliebige Zeile nicht direkt angesprungen werden
	; Der folgende Code prüft ob die gewünschte Zeile schon angelegt wurde. Wenn nicht wird diese angelegt.
		ControlGet, caveZeilenMax, List, Count, SysListview321, % CaveVon
		If (caveZeilenMax < nr) {

		; Fenster nochmals aktivieren, falls inaktiv
			If !WinActive(CaveVon)
				WinActivate, % CaveVon
		; letzte Listenzeile auswählen
			ControlSend, SysListview321, % caveZeilenMax, % CaveVon
			sleep % slTime
		; Editfeld (Arzt) fokusieren
			ControlSend, SysListview321, {Space}, % CaveVon
			sleep % slTime
		; solange TAB senden bis die gewünschte Listenzeile erreicht ist
			Loop {

				ControlGetFocus, FControl, % CaveVon
				If InStr(FControl, "Edit")
					SendInput, {TAB}
				sleep % slTime

				ControlGet, caveZeilenMax, List, Count, SysListview321, % CaveVon
				ControlGetText, CtrlText, Edit1, % CaveVon
				If (caveZeilenMax = nr) && (StrLen(CtrlText) = 0)       	; Eingabefokus soll ins Beschreibungsfeld, dieses enthält keinen Text
					break
				else if (caveZeilenMax > nr) || (A_Index > nr * 4)    	; Notfallabbruch um Endlosschleife zu vermeiden
					break

			}

		}
		else { 	; Listenzeile existiert, dann direkter Sprung

			; im 'Cave! von' - Fenster zur n. Zeile springen (ACC Path 4.1.4.9)
				ControlSend, SysListview321, % nr, % CaveVon
				sleep % slTime
			; 1.Editfeld (Arzt) fokusieren
				ControlSend, SysListview321, % "{Space}", % CaveVon
				sleep % slTime
			; 2.Editfeld (Datum) fokusieren
				SendInput, {TAB}
				sleep % slTime
			; 3.Editfeld (Beschreibung) fokusieren
				SendInput, {TAB}
				sleep % slTime

	}

	; in dieses Feld den Text schreiben
		If WinExist(CaveVon) {
			ControlSend, Edit1, % "{Raw}" txt, % CaveVon
			sleep % slTime * 2
		}

	; kontrollieren ob der Text eingefügt wurde, wenn nicht wird der Nutzer zur manuellen Intervention aufgefordert
		If !(res:= VerifiedSetText("Edit1", txt, CaveVon, 500)) 	{
			Clipboard:= txt
			Clipboard:= txt
			MsgBox,
				(LTrim
				Der neue Text konnte nicht ins
				Cave! von Fenster kopiert werden.
				Dieser ist im Clipboard jetzt vorhanden!
				Bitte manuell per Tastenkombination Steuerung (links)+v in die Zeile %nr% eintragen!
				Drücke wenn Du fertig bist Ok!"
				)
		}

	; das Cave! von - Fenster schliessen, muss eventuell mehrfach trotz spezieller eigener Funktion wiederholt werden,
	; da das Cave! von - Fenster per wm_command Aufruf selbst nach dem Schließen einer Patientenakte geöffnet bleiben kann. Dies beeinträchtigt diese Funktion!
		If CloseCave {
			while WinExist(CaveVon) {
				VerifiedClick("Button3", CaveVon)
	     		Sleep % slTime * 2
				If (A_Index > 5)
					MsgBox, 0x40000, % "Addendum für Albis on Windows - " A_ScriptName, Das "Cave! von" Fenster läßt sich nicht schließen.`nBitte schließen Sie es von Hand!
			}
		}

return ErrorLevel
}
;08
AlbisMuster30ControlList() {                                                       	;-- Funktion um Änderungen von Check-/Radiobuttons zu erkennen - diese speziell für das Muster30 Formular

/*                                                                                                      BESCHREIBUNG
	AlbisMuster30ControlList()
		Hinweis: 	diese Funktion speichert nur Werte zu einem Control wenn es angehakt werden kann (funktioniert bei Radiobuttons und Checkbuttons also)
						zurück wird ein Object mit einem KeyValue Paar gegeben der Aufbau sieht als Beispiel so aus:
                        {"Button1": 1, "Button2": 1, "Button3": 0, "Button4": 1}
						die Strings sind die internen Namen der Button , 1 steht für checked und 0 für not checked
*/


	WinGet, Liste, ControlList, Muster 30 ahk_class #32770
	Controls:= Object()
	Loop, Parse, Liste, `n
	{
		If InStr(A_LoopField, "Button") {
			ControlGet, ischecked, checked, , %A_LoopField%, Muster 30 ahk_class #32770
			Controls[A_LoopField]:= ischecked
		}
	}

	return Controls
}
;09
AlbisOptPatientenfenster(nr) {                                                    	;-- öffnet das Menu-Fenster Optionen/Patientenfenster und holt den Tab mit der 'nr' nach vorne

			;Errorhandling
		if (nr not Integer) {
			throw Exception("Parameter (nr) must be an integer.", -1)
		} else if (nr<0) {
			throw Exception("Parameter (nr) must be a positive integer value.", -1)
		}

		nr--

		AlbisActivate(1)

		;Vorgang wiederholen bis das Einstellungs-Gui für das Patientenfenster geöffnet wurde
		while !(foundWinID := WinExist("Patientenfenster")) {
			If (A_Index>4)
				return
			PostMessage, 0x111, 32886,,, % "ahk_id " AlbisWinID()					;Opt/Patientenfenster
			WinWait, Patientenfenster,, 2
			If WinExist("Patientenfenster")
					break
			sleep, 500
		}

		ControlGet, SysTabID, Hwnd,, SysTabControl321, ahk_id %foundWinID%
			ControlFocus, ahk_id %SysTabID%, ahk_id %foundWinID%

		;im gefundenen Fenster - Tab(nr) (die Tabzählung beginnt bei 0 deshalb nr-1) - also den Tab Nach Öffnen auswählen
		SendMessage, 0x1300, 1, 0x0,, ahk_id %SysTabID% 	; 0x1330 is TCM_SETCURFOCUS
			sleep 0
		SendMessage, 0x130c, 1, 0x0,, ahk_id %SysTabID%	; 0x133c is TCM_SETCURSEL
		SendMessage, 0x1300, %nr%, 0x0,, ahk_id %SysTabID% 	; 0x1330 is TCM_SETCURFOCUS
			sleep 0		;nicht entfernen sonst funktioniert es nicht
		SendMessage, 0x130c, %nr%, 0x0,, ahk_id %SysTabID%	; 0x133c is TCM_SETCURSEL

			sleep, 1500


		;Tab rechts davon auswählen und zurück, sonst wird
		ControlFocus, ahk_id %SysTabID%, ahk_id %foundWinID%
		SendInput, {Shift Down}{Tab}{Shift Up}
		SendInput, {Right}{Left}

		return foundWinID
}
;10
AlbisHeilMittelKatologPosition() {                                                	;-- verschiebt den Heilmittelkatalog in die Mitte des aktuellen Monitors

	/* LÖSUNG: FÜR DAS FEHLENDE ANZEIGEN  DER AUSFÜLLHILFE (CGM HEILMITTELKATALOG) NACH DRÜCKEN VON F3

		 seit Einführung der sogenannten "Ausfüllhilfe" gab es das Problem das der CGM Heilmittelkatalog, den man üblicherweise per F3 - Taste aufruft, immer mal wieder nicht
		 angezeigt wird. Mein Albis Vertriebspartner hatte keine Idee zu Lösung. Weder ein Albisneustart noch ein Windows Neustart haben das Anzeigen immer zuverläßig
		 wieder herstellen können.
		 Was scheinbar niemand wusste ist das die Ausfüllhilfe angezeigt wird, deren Koordinaten allerdings außerhalb des sichtbaren Bildschirmbereiches liegen. Diese Funktion
		 verschiebt das Fenster zurück in die Mitte des aktuellen Monitors.

	 */

	m:= GetWindowSpot(WinExist("Heilmittelverordnung für"))
	w:= GetWindowSpot(CGMHK:=WinExist("CGM HEILMITTELKATALOG ahk_class Qt5QWindowIcon"))
	SetWindowPos(CGMHK, m.X + (m.W - w.W)//2, m.Y + (m.H - w.H)//2, w.W, w.H)

return
}
;11
AlbisSortiereDiagnosen() {                                                         	;-- ## funktioniert nicht - nur zum Sortieren von Diagnosen im Dauerdiagnosenfenster

	DoSort:= 0

	If !hDia:= GetHex(WinExist("Dauerdiagnosen von ahk_class #32770"))
			return -1
	LVRow:= AlbisLVContent(hDia, "SysListView321", "1")
	ControlGet, hLV, HWND,, SysListView321, ahk_id %hDia%
	Loop, % LVRow.MaxIndex()	{

			If RegExMatch(LVRow[A_Index], "^Behandlung")
					thisrow:= "Behandlung"
			else if RegExMatch(LVRow[A_Index], "^anamnestisch")
					thisrow:= "anamnestisch"

			If InStr(lastrow, "Behandlung") && InStr(thisrow, "anamnestisch")
			{
					DoSort:=1
					break
			}

			Acc:= Acc_ObjectFromWindow(hLV)

			;LV_Select(A_Index, "SysListview321", hDia)
			;ToolTip, % A_Index
			;sleep, 1000
	}
	ToolTip, % "State: " Acc_GetStateText(Acc.accState(hLV)) ;"`nSelection: " Acc.accSelection "`nFocus: " acc.accFocus, 1200, 5, 3		;"accName: " Acc.accName(0)
	If DoSort
	{
			MsgBox, 4, Addendum für Albis on Windows, % "Die Dauerdiagnosen könnten eine Sortierung vertragen!`nSoll ich diese jetzt durchführen?", 6
			IfMsgBox, No
				return
			Loop, % LVRow.MaxIndex()
			{
					 if RegExMatch(LVRow[A_Index], "^anamnestisch")
						nil:=1
			}
	}

}
;12
AlbisReadFromListbox(WTitle, ListboxNr, EntryNr:=1) {                	;-- liest eine Zeile oder alle Zeilen aus einer Listbox

	If RegExMatch(WTitle, "^0x[\w]+$", WinID)
		WTitle	:= "ahk_id " WinID
	else if RegExMatch(WTitle, "^\d+$", digits)
		If StrLen(WTitle) = StrLen(digits)
			WinID := "ahk_id " digits
	else
		WTitle:= "ahk_id " GetHex(WinExist(WTitle))

	;ControlGet, LBList, List,, % "Listbox" ListboxNr, % WTitle
	ControlGet, LBList, List,, Listbox%ListboxNr% , % WTitle

	If EntryNr <> 0
		Loop, Parse, LBList, `n
			If (A_Index = EntryNr)
				return A_LoopField


return LBList
}
;13
AlbisResizeDauerdiagnosen(Options:= "xCenterScreen yCenterScreen w0 h0") { ;-- spezielle Funktion für den Dauerdiagnosendialog

	static WT:= "Dauerdiagnosen von ahk_class #32770"
	static hStored

	hWin := WinExist(WT)
	If !hWin || (hStored = hWin)
		return 0

	opt := ParseGuiOptions(Options)

	hStored	:= hWin
	hMed	:= Controls("SysListView322", "ID", hWin)
	hDia 	:= Controls("SysListView321", "ID", hWin)
	dv      	:= GetWindowSpot(hWin)
	med    	:= GetWindowSpot(hMed)
	dia     	:= GetWindowSpot(hDia)

	Control, Hide,, ThunderRT6UserControlDC1, % WT
	Control, Hide,, Internet Explorer_Server1, % WT

	ControlGetPos,, Btn1Y, Btn1W,, Button1, % WT
	ControlGetPos,, Btn2Y, Btn2W,, Button2, % WT
	ControlMove, Button1, % dv.CW - Btn1W,,,, % WT
	ControlMove, Button2, % dv.CW - Btn2W,,,, % WT

	ControlMove, SysListView322,,, 400 ,, % WT
	inCW := dv.CW - Btn1W - 5
	ControlMove, SysListView321, 430,, % inCW - 440 ,, % WT

return 1
}
ParseGuiOptions(optString) {

	RegExMatch(Options, "i)x(?<X>[A-Z0-9]+)(\s+|$)"  	, opt)
	RegExMatch(Options, "i)y(?<Y>[A-Z0-9]+)(\s+|$)"  	, opt)
	RegExMatch(Options, "i)w(?<W>[A-Z0-9]+)(\s+|$)"	, opt)
	RegExMatch(Options, "i)h(?<H>[A-Z0-9]+)(\s+|$)"  	, opt)

	If !optX && !opY && !opW && !opH
		return

return {"X":optX, "Y":optY, "W":optW, "H":optH}
}
;14
AlbisResizeLaborAnzeigegruppen(Options:= "xCenterScreen yCenterScreen w0 h0") { ;-- Fensterhöhe verändert für mehr Übersicht

	static WT := "Labor - Anzeigegruppen ahk_class #32770"
	static hStored

	hWin := WinExist(WT)
	If !hWin || (hStored = hWin)
		return 0

	opt := ParseGuiOptions(Options)

	hStored	:= hWin
	hMGroup	:= Controls("Listbox1", "ID", hWin)
	hPara   	:= Controls("Listbox2", "ID", hWin)
	hIGroup   	:= Controls("Listbox3", "ID", hWin)

	albis        	:= GetWindowSpot(AlbisWinID())
	win          	:= GetWindowSpot(hwin)
	mg          	:= GetWindowSpot(hMGroup)
	pa        	:= GetWindowSpot(hPara)
	ig         	:= GetWindowSpot(hIGroup)

	monIndex	:= GetMonitorIndexFromWindow(AlbisWinID())
	Mon     	:= ScreenDims(monIndex)

	ControlGetPos,,,, TaskbarH	, ReBarWindow321	, % "ahk_class Shell_TrayWnd"
	ControlGetPos,,,, Btn1H     	, Button1 	            	, % WT
	ControlGetPos,,,, LB1H      	, ListBox1                	, % WT
	ControlGetPos,,,, LB2H      	, ListBox2                	, % WT
	ControlGetPos,,,, LB3H      	, ListBox3                	, % WT

	HeightPlus := (Mon.H-TaskbarH) - win.H - win.BH

	win.X := Floor(albis.W/2 - win.W/2)
	win.Y := (albis.Y<0?0:albis.Y)
	SetWindowPos(hwin, win.X, win.Y, win.W, Mon.H-TaskbarH)

	ControlMove, Button1	,,,, % Btn1H	+ HeightPlus - 10	, % WT
	ControlMove, ListBox1	,,,, % LB1H	+ HeightPlus - 10	, % WT
	ControlMove, ListBox2	,,,, % LB2H	+ HeightPlus     	, % WT
	ControlMove, ListBox3	,,,, % LB3H	+ HeightPlus     	, % WT


}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; DATUMSFUNKTIONEN                                                                                                                                                                  	(04)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisLeseZeilenDatum           	(02) AlbisSetzeProgrammDatum 	(03) AlbisLeseProgrammDatum     	(04) AlbisSchliesseProgrammDatum
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
; (01)
AlbisLeseZeilenDatum(Sleeptime=80, closeFocus=true) {                                        	;-- liest das Zeilendatum der ausgewählten Zeile in der Akte aus

	; Achtung eine Zeile in der Akte muss nicht und sollte nicht ausgewählt sein, der Mauspfeil sollte aber über der Zeile stehen
	; dies funktioniert nur in der aktuellen Patientenakte (es darf noch keine Abrechnung stattgefunden haben)
	; letzte Änderung: 19.03.2021 - Erkennung des aktiven Fokus verbessert, Codesyntax modernisiert

		global KK

		Datum := 0

	;falls ein Eingabefocus (blinkendes Caretzeichen) schon gesetzt ist, kann das Datum sofort ausgelesen werden
		If !InStr(AlbisGetActiveWindowType(), "Karteikarte")
				return 0

		hACtrl := GetFocusedControl(), hActive := GetHex(GetParent(hACtrl)), cf := GetClassNN(hACtrl, hActive)
		If InStr(cf, "Edit") ;|| InStr(cf, "RichEdit")
			ControlGetText, Datum, Edit2, % "ahk_id " hActive

	;falls kein Focus besteht, dann jetzt einen Focus erzeugen
		while !RegExMatch(Datum, "\d\d\.\d\d\.\d\d\d\d") && (A_Index < 6) {
			SendInput, {Space}
			Sleep % sleeptime
			ControlGetText, Datum, Edit2, % "ahk_id " hActive ;ALBIS -
		}

	;fragt den Focus ab, Esc beendet den Eingabemodus in der Akte, es sollte danach ein anderer Focus auslesbar sein
		If closeFocus
			while (cf = cf1)  && (A_Index < 10)	{
				SendInput, {Esc}
				Sleep % sleeptime
				ControlGetFocus, cf1, % "ahk_id " hActive
			}

return StrLen(Datum) = 0 ? 0 : Datum
}
; (02)
AlbisSetzeProgrammDatum(Datum:="") {                                                                  	;-- öffnet das Fenster Programmdatum einstellen und setzt das übermittelte Datum ein

		; Achtung: Datum muss absofort zwinged folgendes Format aufweisen: dd.mm.yyyy z.B. 24.12.2019
		; bei leerem Datum wird das aktuelle Tagesdatum eingetragen

		If (StrLen(Datum) = 0)
			FormatTime, Datum,, dd.MM.yyyy

	; prüft auf Richtigkeit des vorliegenden Datumformates
		if !RegExMatch(Datum, "\d\d\.\d\d\.\d\d\d\d") 		{
			Throw exception("'" Datum "' hat kein gültiges Format. [dd.MM.yyyy]" )
			return
		}

	; Fenster Programmdatum einstellen aufrufen
		If !(hWin:= Albismenu(32852, "Programmdatum einstellen", 3)) {
			PraxTT(A_ScriptName "`nDer Dialog 'Programmdatum' ließ sich nicht aufrufen! `nDie Funktion wird jetzt beendet", "5 1")
 			AlbisSchliesseProgrammDatum()
			return 0
		}

		ControlGetText, eingestelltesDatum, Edit1, % "Programmdatum einstellen ahk_class #32770"

		VerifiedSetText("Edit1", Datum, "Programmdatum einstellen ahk_class #32770", 150)
		AlbisSchliesseProgrammDatum()

return eingestelltesDatum
}
; (03)
AlbisLeseProgrammDatum(closeWin:=true) {                                                            	;-- öffnet das Fenster Programmdatum einstellen und setzt das übermittelte Datum ein

	; V1.2 mit Überprüfung ob das Fenster geöffnet ist, das Auslesen des Datum funktioniert und ob das Fenster auch wieder geschlossen werden konnte

	; Fenster Programmdatum einstellen aufrufen
		If !(hWin:= Albismenu(32852, "Programmdatum einstellen", 3)) {
			PraxTT(A_ScriptName ": Das Programmdatum ließ sich nicht anzeigen!", "1 0")
			return 0
		}
		while (StrLen(Datum) = 0) && (A_Index <= 40) {
			If (A_Index > 1)
				Sleep 50
			ControlGetText, Datum, Edit1, % "Programmdatum einstellen ahk_class #32770"
		}

	; Fenster schliessen
		If closeWin
			AlbisSchliesseProgrammDatum()

return Datum
}
; (04)
AlbisSchliesseProgrammDatum() {                                                                          	;-- schließt das Programmdatumsfenster

	static WinTitle := "Programmdatum einstellen"

	while WinExist(WinTitle)	&& (A_Index < 3)	{

		If (A_Index = 1) {
			If !VerifiedClick("Button1", WinTitle) {
				WinActivate, % WinTitle
				SendInput, {ENTER}
			}
		} else If (A_Index = 2) {
			SendInput, {ENTER}
		}

		If WinExist("Albis ahk_class #32770", "Fehlerhaftes Datumsformat")
			VerifiedClick("Button1", "Albis ahk_class #32770", "Fehlerhaftes Datumsformat")

		Sleep, 50

	}

return WinExist(WinTitle) ? 0 : 1
}
; (05)
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; DATENEINGABE                                                                                                                                                                           	(13)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisPrepareInput                 	(02) AlbisSendInputLT                	(03) AlbisWriteProblemMed           	(04) AlbisKarteikarteAktivieren
; (05) AlbisKarteikarteFocus            	(06) AlbisKarteiKarteEingabeStatus	(07) AlbisSchreibeLkMitFaktor      	(08) AlbisSchreibeInKarteikarte
; (09) AlbisFehlendeLkoEintragen    	(10) AlbisKopiekosten                  	(11) AlbisSchreibeSequenz           	(12) AlbisZeilendatumSetzen
; (13) AlbisSendText
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisPrepareInput(Name) {                                                                                    	;-- bereitet das Schreiben von Daten in die Akte vor


		; letzte Änderung 02.09.2019:
		; nach dem Click in die Akte (#327702) , wird die Escape-Taste gesendet, dies befreit das Caret vom derzeitigen Eingabefeld.
		; Fehleranzeige bei Übergabe eines nicht erlaubten Parameter (Kontrollvariable ist benannt als 'allowed')

		/*                            						BESCHREIBUNG

					vorbereitende Funktion um danach andere Aktionen in der Akte auszulösen - z.b. Schreiben oder auslesen
					erstellt dafür eine freie Zeile und bei Angabe eines Kürzel wird dieses in die Akte geschrieben
					Name kann sein: KKk, KKD, TD, cave, DV, DD, DM, scan, ...
					KKk - Karteikartenkürzel um z.b. lko 03230 einzugeben
					KKD - Karteikartendatum ,
					TD - Tagesdatum öffnen,
					cave - das cave Fenster,
					DV - Daten von Patient Fenster,
					DD - Dauerdiagnosen,
					DM - Dauermedikamente (letztere nicht impl.)
					dies funktioniert nur mit Abrechnungsziffern an einem Datum im aktuellen Quartal (es darf noch keine Abrechnung stattgefunden haben)
		*/

		error:= 0
	; erlaubte Kürzel für die Patientenakte, anders übergebene werden verworfen, * öffnet nur die Zeile übergibt aber keinen Inhalt
		allowed:= "lko|lkü|bild1|scan|medrp|medpn|fhv13|fhv18|KKk|*"					;den String können Sie mit ihren eigenen Karteikartenkürzeln füllen

		while !AlbisExist() {
				MsgBox, 0x40001 , Addendum für Albis On Windows, % 	"WARNUNG!`nBitte warten Sie auf den Neustart von Albis!`n"
																									.	"Drücken Sie dann erst auf 'OK'!`n"
																									.	"'Abbrechen' beendet die aufgerufene Funktion."
				IfMsgBox, Cancel
						return 2
		}

	; Albishauptfenster aktivieren und von eventuell blockierenden Fenstern befreien
		AlbisActivate(5)
		AlbisIsBlocked(AlbisWinID(), 2)

	; schaltet auf die Karteikarte des aktuellen Patienten um, falls noch nicht geschehen
		AlbisKarteikarteZeigen()

	; eine freie Zeile in der Akte schaffen
		If Instr(allowed, Name)		{
			localIndex := 0
			Loop					{
					; ein Klick auf den Eingabebereich der Patientenakte um dieses Fenster zu aktivieren, sonst können dort keine Daten eingetragen werden!
						SendInput, {Esc}
						sleep 300
						Mdihwnd:= AlbisMDIChildGetActive()
						VerifiedClick("#327702","","", MdiHwnd)
						If InStr(GetClassName(GetFocusedControl()), "edit")
								SendInput, {Escape}
						sleep, 300
						SendInput, {End}
						sleep 300
						SendInput, {End}
						sleep 300
						SendInput, {Down}
						sleep 300
					; prüft welches Feld jetzt fokusiert ist und ob ein Textinhalt vorhanden ist
						;ControlGetFocus, lCFocus, % "ahk_id " MdiHwnd
						;ControlGetText, lFtext, %lCFocus%, % "ahk_id " MdiHwnd
						ControlGetFocus, lCFocus, % "ahk_class OptoAppClass"
						ControlGetText, lFtext, %lCFocus%, % "ahk_class OptoAppClass "
						If Instr(lCFocus,"Edit3") and (lFtext="")
								error:= 0, break

					;Funktion nicht erfolgreich in der Ausführung
						localIndex++
						if (localIndex>10)
								error:= 99, break

			} until Instr(lCFocus,"Edit3") and (lFtext="")			;dieses Feld ist vorhanden und leer, wenn die Befehle im Loop funktioniert haben
		}
		else		{
			; Fehlerbehandlung bei Eingabe eines falschen Parameter
				ErrorMessage:= "Error at script line: " . scriptline . "`n" . ScriptText . "`n`n-> Function call with not allowed parameter(s)! <-"
				ExceptionHelper("include`\AddendumFunctions.ahk", "If Instr`(allowed`, Name`) `{", ErrorMessage, A_LineNumber)
				return 90				;falsche Parameterbehandlung
		}


	; einzelne Abhandlungen für übergebene Kürzel
		If Instr(Name, "bild1") || Instr(Name, "scan")		{
			SendRaw, %Name%
			SendInput, {TAB}
			WinWait, Grafischer Befund ahk_class #32770,, 8
			If !WinExist("Grafischer Befund ahk_class #32770") {
					error:= ErrorLevel
					MsgBox, 0x1000, Fensterfehler, Das Fenster 'Grafischer Befund hat sich nicht geöffnet', 4
			}
		}

		If Instr(Name, "lko") || Instr(Name, "lkü") 	{
			SendRaw, %Name%
			SendInput, {TAB}
		}

return error
}
;02
AlbisSendInputLT(kk, inhalt, kk_Ausnahme, kk_voll) {                                              	;-- in die Karteikarte schreiben

		; VERALTETE FUNKTION!!
		;AlbisSendInputLeistungstext
		;schreibt ein Leistungskürzel (lko/lkn/lkü) und anschließend die Ziffer in die Akte, wenn eine Ausnahme(kk_Ausnahme) besteht wird das volle Kürzel (kk_voll) in die Akte geschrieben
		;Grund für dieses Vorgehen ist, das es kein einfaches Vorgehen (auslesen der Albis Oberfläche mittels Autohotkey) gibt den Status eines Patienten (Notfall/Überweisung) auszulesen


		SendInput, % kk	                                    	;die Eingabe eines "L" und anschließend eines Tab reicht - Albis ergänzt den Rest, außer wenn ein Überweisungsschein angelegt ist lkü
			sleep, 100
		SendInput, {TAB}
			sleep, 100
		SendRaw, % inhalt
			sleep, 100
		SendInput, {TAB}

return
}
;03
AlbisWriteProblemMed(medname) {                                                                       	;-- Dauermedikamente schreiben

	;falls noch kein Problemmedikamentfeld vorhanden ist
	static1:= "#### Problemmedikamente / Allergien #####", lenstatic1:= StrLen(static1)					;41 Zeichen
	static2:= "#############################", lenstatic2:= StrLen(static2)



}
;04
AlbisKarteikarteAktivieren() {                                                                                  	;-- Karteikartenfenster aktivieren

	; ohne Einsatz dieser Funktion läßt sich keine Dateneingabe in die Akte durchführen
	; gibt das hwnd des Karteikartenfensters zurück

	Mdihwnd:= AlbisMDIChildGetActive()		                   	;neu seit 27.3.2019 - hoffentlich zuverlässig
	VerifiedClick("#327702","","", MdiHwnd)

return Mdihwnd
}
;05
AlbisKarteiKarteEingabeStatus(Title:="") {                                                                   	;-- ermittelt ob die akt. Karteikarte  für Eingaben bereit ist

	; Funktion ermittelt das HWND des RichEdit-Steuerelementes (RichEdit20A1) und prüft ob dieses sichtbar ist,
	; sichtbar bedeutet, das Eingaben gemacht werden können

		global KK

		If !title
			hMDIChild := AlbisMDIChildGetActive()
		else
			hMDIChild := AlbisMDIChildHandle(Title)
		ControlGet, hRE, hwnd,, % KK.INHALT, % "ahk_id " hMDIChild
		WinGet, Style, Style, % "ahk_id" hRE

return Style & 0x10000000 ? true : false
}
;06
AlbisKarteikarteFocus(Control, Karteikarte:="") {                                                        	;-- gezielter Schreibzugriff auf die Karteikarte

	; letzte Änderung 27.08.2020
	; Control: es kann direkt die ClassNN des Steuerelementes übergeben werden oder einer der Alias-Bezeichnungen aus KK
	; KK findet man in der Funktion Init_Albis()

		global Mdi:= Object()
		global KK

	; Alias-Bezeichnung wird durch ClassNN des Steuerelementes ersetzt
		If KK.HasKey(Control)
			Control := KK[Control]

	; prüft und zeigt die Karteikarte an, falls etwas anderes beim Patienten angezeigt wird
		AlbisActivate(2)
		AlbisKarteikarteZeigen()

	; ermittelt das richtige Control um dort die Eingaben durchzuführen
		hChild          	:= AlbisMDIChildGetActive()
		hEingabeWin	:= Controls("#327702", "ID", hChild)

	; Tastaturfokus auf den Karteibereich und Tastenfolge senden
		ControlFocus,, % "ahk_id " hEingabeWin
		SendInput, {Escape}
		sleep, 100
		SendInput, {Escape}
		sleep, 100
		SendInput, {End}
		Sleep, 100
		SendInput, {End}
		Sleep, 100
		SendInput, {Down}
		sleep, 200

	; Sicherheitsüberprüfung
		Loop {

			ControlGetFocus, fctrl, % "ahk_id " hEingabeWin
			If InStr(fctrl, "Edit") || (A_Index > 10)	;If InStr(fctrl, Control) || (A_Index > 5)
				break
			else
				SendInput, {Down}

			sleep, 100
		}

	; gewünschten Teil fokussieren
		ControlFocus, % Control, % "ahk_id " hEingabeWin
		sleep, 100

		ControlGetFocus, fctrl, % "ahk_id " hEingabeWin

return ((fctrl = Control) ? GetHex(hEingabeWin) : "0")
}
;07
AlbisSchreibeLkMitFaktor(Leistungskomplex, StandardFaktor:="") {                          	;-- vereinfachtes Senden von LK per Hotstring

	; z.B. AlbisSchreibeLkMitFaktor("40144", "2") - selektiert automatisch den Faktor damit dieser geändert werden kann
	; ist Standardfaktor leer oder null wird nur der Leistungskomplex gesendet

	sleep, 200

	If Standardfaktor {
		SendInput, % "{Text}" Leistungskomplex "(x:" Standardfaktor ")"
		sleep 100
		SendInput, % "{Left " 1 + StrLen(Standardfaktor) "}"
		SendInput, % "{LShift Down}{Right " StrLen(Standardfaktor) "}{LShift Up}"
	}
	else
		SendRaw, % Leistungskomplex

return
}
;08
AlbisSchreibeInKarteikarte(ArztKennung:="", Datum:="", Kuerzel:="", Eintrag:="") {   	;-- schreibt eine gesamte Zeile in die Karteikarte

	/* BESCHREIBUNG DER FUNKTION > AlbisSchreibeInKarteikarte() <

		- letzte Änderung 02.10.2019

		* diese Funktion soll auf einfache Weise das Eintragen von Daten in eine Karteikarte unterstützen
		* es gibt keine Überprüfung ob die korrekte Patienten-Karteikarte geöffndet ist, nutzen Sie dazu andere Funktionen aus dieser Funktionsbibliothek
		* das Albis Fenster wird aktiviert, damit es in den Vordergrund kommt und der Ablauf überwacht werden kann
		* die Funktion setzt dann den Eingabefokus, auch Cursor oder Caret genannt in den Eingabebereich der Karteikarte.
		   Begonnen wird mit der Spalte, deren Parameter keine leere Zeichenkette enthält.
		   ACHTUNG: die Parameter Kuerzel und Eintrag - sollten nicht leer übergeben werden!
		* sind alle übergebenen Parameter eingetragen wird die Eingabe durch ein zusätzliches Senden der TAB-Taste abgeschlossen
		* der erfolgreiche Abschluss der Eingabe wird durch die Änderung des aktiven (Eingabe-) Steuerelementfokus erkannt

		Parameter:
		 die Namen der Parameter sind der Albisdokumentation entnommen (siehe Karteikarte 4.1.1)
		 - Arztkennung	=	Spalte 1 - Arztkürzel oder auch Arztkennung
		 - Datum        	=	Spalte 2 - Datum
		 - Kuerzel       	=	Spalte 3 - Karteikartenkürzel
		 - Eintrag       	=	Spalte 4 - Karteikarteneintrag

		Rückgabeparameter:
		 - keine

		 Abhängigkeiten:
		  - AlbisKarteikarteFocus()
		  - GetFocusControlClassNN()

	*/

		AlbisActivate(1)

		If (StrLen(ArztKennung) > 0) 	{
			hEingabeFocus := AlbisKarteikarteFocus("Edit1")
			VerifiedSetText("Edit1", Nutzer, hEingabefocus)
			SendInput, {Tab}
		}

		 If (StrLen(Datum) > 0)	{
			If !hEingabefocus
				hEingabeFocus:= AlbisKarteikarteFocus("Edit2")
			VerifiedSetText("Edit2", Datum, hEingabefocus)
			SendInput, {Tab}
		}

		If !hEingabefocus
			hEingabeFocus:= AlbisKarteikarteFocus("Edit3")

		VerifiedSetText("Edit3", Kuerzel, hEingabefocus)
		SendInput, {Tab}
		VerifiedSetText("RichEdit20A1", Eintrag, hEingabefocus)
		SendInput, {Tab}
		sleep, 100

		SendInput, {Esc}
		sleep, 100
		SendInput, {Esc}

		;~ While Instr(GetFocusedControlClassNN(AlbisWinID()), "RichEdit20A1") 	{
				;~ SendInput, {Tab}
				;~ SendInput, {Esc}
				;~ sleep, 600
				;~ If !Instr(GetFocusedControlClassNN(AlbisWinID()), "RichEdit20A1")
					;~ break
				;~ If (A_Index > 3)
					;~ MsgBox, Bitte beenden Sie die Eingabe manuell!
		;~ }

return GetFocusedControl()
}
;09
AlbisFehlendeLkoEintragen(AbrechnungslistenPfad="", Arbeitsfrei="") {                     	;-- automatisiert Eintragungen von Komplexziffern

	; gehört zum Modul Abrechnungshelfer - sollte aber auch universell einsetzbar sein
	; ACHTUNG: ergänzt im Moment nur die Ziffern 03220, 03221
	; Funktion umgestellt am 01.07.2020! : die Automatisierungsdatei ist absofort eine Textdatei im Ini-Format
	; -----------------------------------------------------------------------------------------------------------------------------------------------------
	; Parameter:
	;
	;   	AbrechnungslistenPfad:	vollständiger Windows-Pfad zur Automatisierungsdatei
	; 		Arbeitsfrei:                 	eine Komma getrennte Liste von Tagen (Tag und Monat im Format dd.MM.) an denen die Praxis geschlossen war
	;
	; -----------------------------------------------------------------------------------------------------------------------------------------------------
	; -----------------------------------------------------------------------------------------------------------------------------------------------------

		static hac, Next, BAusschluss, hBAusschluss, BNext, hBNext, NextPos, OkX, OkY, OkW, OkH
		static NextPat, index, obj, PatIndex, bearbeitet

		global oGB

	; Fehlerausgabe bei nicht vorhandener Datei
		If !FileExist(AbrechnungslistenPfad) {
			SplitPath, AbrechnungslistenPfad, OutFileName
			MsgBox, 1, % "Addendum für Albis on Windows - " A_ScriptName, % "Die übergebene Abrechnungsliste`n" OutFileName "`nist nicht vorhanden."
			return 0
		}

	; liest die Abrechnungsliste und wandelt diese in ein Objekt um ;{
		oGB      	:= Array()
		PatIndex 	:= 0
		bearbeitet	:= 0

		IniRead, val, % AbrechnungslistenPfad, Abrechnung, Quartal
		RegExMatch(val, "(?<Jahr>\d+)-(?<Nr>\d+)", Quartal)
		firstMonth	:= ((Quartal - 1) * 3) + 1
		lastmonth	:= FirstMonth + 2
		firstDay 	:= "01." firstMonth "."

		Loop {

			IniRead, val, % AbrechnungslistenPfad, Patientenliste, % "Patient" A_Index
			If InStr(val, "ERROR")
				break
			else if (StrLen(val) = 0)
				continue

			PatIndex ++
			data := StrSplit(val, "|")
			bearbeitet := (data[1] = 1) ? (bearbeitet + 1) : bearbeitet

			oGB.Push({ 	"bearbeitet"        	: data[1]
										, 	"PatID"              	: data[2]
										, 	"PatName"         	: data[3]
										, 	"Behandlungstage"	: data[4]
										, 	"Ziff1"                	: data[5]
										, 	"Ziff2"                	: data[6]})

		}

	;}

	; zeigt eine Liste mit den Sprechtagen an ;{


		;Gui, PraxKal: new


	;}

	; FOR - SCHLEIFE ARBEITET DIE DATEN IM OBJEKT AB ;{

		For index, obj in oGB {

				;SciTEOutput("PatIndex: " index)

				If (obj.bearbeitet = 1)
					continue

				maxoGB	:= PatIndex - bearbeitet
				maxoGBL	:= StrLen(maxoGB)

			; Karteikarte des Patienten öffnen
				If !Instr(AlbisAktuellePatID(), obj.PatID) {
						PraxTT("Die Karteikarte von:`n" obj.PatName " (" obj.PatID ")`nwird geöffnet", "0 3")
						AlbisAkteOeffnen(obj.PatID)
						PraxTT("", "off")
				}

			; Datumsvorschläge für Zifferneinträge berechnen ;{
				Ziff1Tag 	:= Ziff2Tag := ""
				BTag     	:= StrSplit(obj.Behandlungstage, ";")
				ziffmax  	:= obj.Ziff1 + obj.Ziff2

				If !obj.Ziff1 {
					AlbisSchreibeInKarteikarte("", BTag[1], "lko", "03220")
					Ziff1Tag := SubStr(BTag[1], 1, 5)
					obj.Ziff1 := 1
				}
				else
					Ziff1Tag := "eingetragen am: " SubStr(BTag[1], 1, 5)

				If !obj.Ziff2 && (StrLen(BTag[2]) > 0)
					Ziff2Tag := SubStr(BTag[2], 1, 5)
				else if !obj.Ziff2 && (StrLen(BTag[2]) = 0)
					Ziff2Tag := "kein 2.Behandlungstag vorhanden"

				;}

			; Nutzerabfrage  ;{
				PatName	:= obj.PatName
				PatUebrig	:= SubStr("000" index, -1 * (maxoGBL-1)) "/" maxoGB

				info =
				(LTrim
				Patient: %PatName% (%PatUebrig%)
				---------------------------------------------------------------
				                 Vorschläge für Behandlungstage
				---------------------------------------------------------------
				1. lko 03220 - %Ziff1Tag%
				2. lko 03221 - %Ziff2Tag%
				weitere Behandlungstage:
				)

				; weitere Behandlungstage zusammenstellen
					WIndex := 0, WStart := ""
					Loop % BTag.MaxIndex()
						If RegExMatch(BTag[A_Index], "\d\d\.\d\d.\d\d\d\d") && (A_Index > 2) {   ; && !InStr(BTag[A_Index], Ziff2Tag)
								WStart := (StrLen(WStart) = 0) ? A_Index : WStart
								WIndex ++
								info .= WIndex ": " BTag[A_Index] ", "
						}

					info := LTrim(info, ", ") "`n---------------------------------------------------------------`n"

					If (WIndex = 0)
						info .= "Da keine weiteren Behandlungstage vorhanden sind.`nBitte manuelle Eingabe eines weiteren Behandlungsdatums!`n!!Datum ohne Jahrzahl!!"

					If RegExMatch(Ziff2Tag, "\d\d\.\d\d", match)
						UserDatum := StrReplace(Ziff2Tag, ".")
					else
						Userdatum := ""

				; Eingabedialog anzeigen
					Inputser:
					SetTimer, AddButtonsToInputbox, -50
					InputBox, UserDatum, Addendum für Albis on Windows - Behandlungstage, % info,, 600, 350,,,,, % UserDatum
					If ErrorLevel && !NextPat {
						MsgBox, Bin raus!
						break
					} else if NextPat {                                         	; Button "nächster Patient" oder "Patient als bearbeitet markieren" wurde gedrückt
						NextPat := 0
						;AlbisAkteSchliessen()
						continue
					}

				; prüft ob die Eingabe ein Datumsformat hat
					If RegExMatch(Userdatum, "(?<Tag>\d+)\.(?<Monat>\d+)$", Input)
						dateStr := SubStr("00" InputTag, -1) SubStr("00" InputMonat, -1) QuartalJahr
					else if RegExMatch(Userdatum, "(?<Tag>\d\d)(?<Monat>\d\d)$", Input)
						dateStr := InputTag InputMonat QuartalJahr
					else
						goto Inputser

				; war am eingegeben Tag die Praxis geschlossen
					dateProof := InputTag "." InputMonat "."
					If dateProof in % Arbeitsfrei
					{
						MsgBox, An diesem Tag war die Praxis geschlossen!`nBitte die Eingabe korrigieren.
						goto Inputser
					}

			;}

			; zweite Ziffer eintragen
				focused := AlbisSchreibeInKarteikarte("", dateStr, "lko", "03221")
				;SciTEOutput("focused control classNN: " focused)
				obj.Ziff2 := 1

				MsgBox, 0x1004, % "Addendum für Albis - " A_ScriptName,
				(Ltrim
				Sind alle Ziffern korrekt eingetragen worden?
				Dann drücken Sie auf 'Ok' für den nächsten Patienten:
				%NextPatName%
				wählen Sie 'Nein' um diesen Patienten als unbearbeitet
				zu markieren.
				)
				IfMsgBox, Yes
				{
					IniWrite, % "1|" obj.PatID "|" obj.PatName "|" obj.Behandlungstage "|" obj.Ziff1 "|" obj.Ziff2, % AbrechnungslistenPfad, Patientenliste, % "Patient" Index
					obj.bearbeitet:= 1
					bearbeitet ++
				}

			; bearbeitete Karteikarte schliessen
				;AlbisAkteSchliessen()
		}

	;}

return 1
AddButtonsToInputbox: ;{

	NextPat := 0
	hac   	:= WinExist("A")
	ib     	:= GetWindowSpot(hac)
	NextPatName:= oGB[(index+1)].PatName

	ControlGetPos	, OkX , OkY, OkW, OkH, Button1                              	, % "ahk_id " hac
	ControlMove 	, Button1, % ib.BW,,,                                                  	, % "ahk_id " hac
	ControlSetText	, Button2, Beenden                                                    	, % "ahk_id " hac
	ControlGetPos	, Ok2X , Ok2Y, Ok2W, Ok2H, Button2                      	, % "ahk_id " hac
	ControlMove 	, Button2, % (ib.W-ib.BW) - Ok2W - 10,, % Ok2W + 5,	, % "ahk_id " hac

	Gui, next: New, -Caption +ToolWindow -SysMenu +HWNDhNext Parent%hac%
	Gui, next: Margin, 0, 0
	Gui, next: Add, Button, % "x0 y0 vBNext       	HWNDhBNext       	gToNext", % "nächster Patient >> " NextPatName " "
	Gui, next: Add, Button, % "x+5 	vBAusschluss HWNDhBAusschluss 	gToNext", % "Pat. als bearbeitet markieren"
	Gui, next: Show, % "Hide"
	ControlGetPos,,, NextPosW,, Button1, % "ahk_id " hNext
	ControlGetPos,,, NextAusW,, Button2, % "ahk_id " hNext
	Gui, next: Show, % "x" ib.BW + OkW + 10 " y" OkY - (ib.H - ib.CH - ib.BH)

	WinSet, AlwaysOnTop, On, % "ahk_id" hac

	;Voreineinstellung im Edit-Steuerelement selektieren
	ControlFocus, Edit1, % "ahk_id " hac
	ControlGetText, text, Edit1, % "ahk_id " hac
	SendMessage, 0xB1, 0, -1, Edit1, % "ahk_id " hac

return ;}
ToNext: ;{

	If InStr(A_GuiControl, "BNext") {
		NextPat := 1
	} else if InStr(A_GuiControl, "BAusschluss") {
		IniWrite, % "1|" obj.PatID "|" obj.PatName "|" obj.Behandlungstage "|" obj.Ziff1 "|" obj.Ziff2, % AbrechnungslistenPfad, Patientenliste, % "Patient" Index
		obj.bearbeitet := 1
		bearbeitet ++
		NextPat := 1
	}

	Gui, next: Destroy
	ControlClick, Button2, % "ahk_id " hac

return ;}
}
;10
AlbisKopiekosten(Kopiepreis, Kopien:="", nurString:=false) {                                      	;-- Kopiekosten für Privatrechnung berechnen

	; berechnet Kopiekosten und schreibt einen GOÄ fähigen Text in die Karteikarte oder Privatschein
	;
	; Kopiepreis:  in Euro pro Seite als Satz wie z.B. im Brief des LASV: "0.50 ab 50 Seiten 0.15"
	; Kopien:     	verwenden wenn die Funktion ohne Inputbox aufrufen möchte
	; nurString: 	berechnete Gebühren werden nicht in den Abrechnungsschein übertragen, der Gebührentext ist der Rückgabeparameter
	;
	; letzte Änderung 11.09.2021

		static kDelay, kDuration, CooCt, cFocus
		sendString := []

	  ; prüft ob eine Privatabrechnung angelegt ist, gibt einen Hinweis falls nicht
		If !nurString {

			If !AlbisAbrechnungsscheinVorhanden("P\s+(Priva|BG).*?BE") {

				MsgBox	, 0x2030
							, % StrReplace(A_ScriptName, ".ahk")
							, % "Es sollen Sachkosten angesetzt werden.`n"
							. 	 "Dazu muss eine Privatabrechnungsschein angelegt sein.`n"
							.	 "Bitte legen Sie diesen zunächst an und`n wiederholen Sie die Kopiekostenberechnung!`n"
							, 8
				return 0

			}

		  ; bisherige Einstellungen speichern
			kDelay  	:= A_KeyDelay
			kDuration	:= A_KeyDuration
			CooCt  	:= A_CoordModeCaret

		  ; neue Einstellungen setzen
			CoordMode	, Caret, Screen
			SetKeyDelay, 0, 0

		  ; Eingabefocus und Position ermitteln
			ControlGetFocus                 	, 	  cFocus, % "ahk_class OptoAppClass"
			ControlGetPos, cx, cy, cw, ch	, % cFocus	, % "ahk_class OptoAppClass"
			mx := A_CaretX, my := A_CaretY + ch + 5
			my := ((my + 130) > A_ScreenHeight) ? (A_ScreenHeight - 150) : my

		}

	  ; Inputbox-Timer passt die Gui etwas an
		If !Kopien {
			SetTimer, KopiekostenFokus, 100
			InputBox, Eingabe, Addendum für Albis on Windows, Wieviele Kopien wurden gedruckt?,, 300, 130, % mx , % my, Locale, 40, 1
			RegExMatch(Eingabe, "(?<Kopien>\d+)", Anzahl)
		} else {
			AnzahlKopien := Kopien
		}

		AnzahlKopien := RegExReplace(AnzahlKopien, "[^\d]", "")
		If (StrLen(AnzahlKopien) = 0) || !RegExMatch(AnzahlKopien, "\d+")
			AnzahlKopien := 1

	  ; Kopiegebühr aus Standardsatz in Behördenbriefen extrahieren (Preise in Euro angeben)
	  ; z.B.: "0.50 Euro ab 50 Seiten 0.15" = Preis1 = 0.5, AbSeite = 50 , Preis2 = 0.15
		RegExMatch(Kopiepreis, "i)^(?<Preis1>[\d\.\,]+).*?a*b*\s*(?<AbSeite>[\d]*)\s*(Seiten*)\s*(?<Preis2>[\d\.\,]*)", k)
		kPreis1 := StrReplace(kPreis1, ",", ".") + 0
		kPreis2 := StrReplace(kPreis2, ",", ".") + 0

	  ; Kopiegebuehr berechnen, zu sendene Strings erstellen
		If (kAbSeite > 0) && (AnzahlKopien > kAbSeite) {
			Minderungsseiten 	:= AnzahlKopien - kAbSeite
			AnzahlKopien    	:= kAbSeite
			sendString.2      	:= "(sach:" (Minderungsseiten=1 ? "Kopie ": "Kopien ") Minderungsseiten "x a " Round(kPreis2*100) " cent:" Format("{:.2f}", (Minderungsseiten*kPreis2)) ")"
		}
		sendString.1 := "(sach:" (AnzahlKopien=1 ? "Kopie ": "Kopien ") AnzahlKopien
							. 	"x a " Round(kPreis1*100) " cent:" Format("{:.2f}", (AnzahlKopien*kPreis1)) ")"

	; Gebührentext zurückgegeben falls angefordert
		If nurString
			return sendString

	 ; Eingabefocus ins Steuerelement in Albis legen
		ControlFocus, % cFocus, % "ahk_class OptoAppClass"
		Sleep 200

	  ; Sachkostentext(e) senden
		For sStrNr, stringToSend in SendString {

			If (sStrNr > 1) {

			  ; in welcher Ansicht befinden wir uns (Abrechnungsschein oder Karteikarte?)
				If InStr(AlbisGetActiveWindowType(), "Privatabrechnung")
					acView := 1
				else if AlbisGetActiveControl("ContractionIsEqual", "lp")
					acView := 2
				else {
					PraxTT("Ups hier ist was schief gegangen.`nEs ließen sich nicht alle Sachkosten eintragen.", "1 3")
					SetKeyDelay, % kDelay, % kDuration
					CoordMode, Caret, % CooCt
					return 0
				}

              ; sicheres finden des korrekten zweiten Eingabefeldes
				while (A_Index <= 10) {

						ControlGetFocus, cFocus, % "ahk_class OptoAppClass"
						CFocusQueue .= CFocus "|"

						If (acView = 1) {                       ; Privatabrechnungsschein -gesuchtes Steuerelement ClassNN ist Edit6

							If  (cFocus = "Edit3") {
								break
							}

							If (A_Index < 20) {
								Send, {Tab}
								sleep 500
							} else If (A_Index >= 20) {
								PraxTT("Ups hier ist was schief gegangen.`nDas korrekte Eingabefeld konnte nicht angesteuert werden..", "1 3")
								SetKeyDelay, % kDelay, % kDuration
								CoordMode, Caret, % CooCt
								return 0
							}


						}
						else If (acView = 2) {

							; bei Eingabe in der Abrechnungsscheinansicht
							If  (cFocus = "Edit6") {                             	; "lp" ergänzen falls es fehlt
								kuerzel := Trim(AlbisGetActiveControl("contraction"))
								If (kuerzel <> "lp") {
									Send, lp
									Sleep, 300
								}
							}
							; bei Eingabe in der Karteikarte
							else if (CFocus <> "RichEdit20A1") {			; Eingabebereich gefunden
								break
							}

							If (A_Index < 20) {                                	; nicht gefunden, mit Tab ein Steuerelement weiter
								Send, {Tab}
								sleep 500
							}
							else If (A_Index >= 20) {                       	; zuviele Versuche - Abbruch
								PraxTT("Ups hier ist was schief gegangen.`nDas korrekte Eingabefeld konnte nicht angesteuert werden..", "1 3")
								SetKeyDelay, % kDelay, % kDuration
								CoordMode, Caret, % CooCt
								return 0
							}

						}

				}


			}

			SendRaw, % stringToSend    	; Format erzeugt hier immer eine Zahl mit 2 Stellen nach dem Komma
			sleep, 300                          	; wartet damit alles gesendet wurde
			Send, {TAB}                       	; ein Eingabefeld weiter vorrücken
			sleep, 300                          	; wartet damit alles gesendet wurde

		}

		Send, {Escape}
		sleep, 200
		Send, {Escape}

		SetKeyDelay, % kDelay, % kDuration
		CoordMode, Caret, % CooCt


		SciTEOutput("CFocus:" CFocusQueue)
		;~ SciTEOutput(AnzahlKopien ", " kAbSeite ", " Minderungsseiten ", " kPreis1 ", " kPreis2)

return
KopiekostenFokus:
	SetTimer, KopiekostenFokus, Off
	ControlFocus, Edit1, % "ahk_class #32770", % "Wieviele Kopien wurden gedruckt"
return
}
;11
AlbisSchreibeSequenz(sequenz) {                                                                             	;-- Blockeintragungen realisieren

	/* Beschreibung

		letzte Änderung: 27.08.2020

		Funktion ist dafür gedacht ganze Blöcke mit verschiedenen Karteikartenkürzeln und Texte mittels einer Übergabe am Stück zu senden.
		Ähnlich der Albis ToDo-Liste, wobei die Funktion ausschließlich Text in die Karteikarte schreibt.
		Die Funktion schreibt nur Daten für ein Tagesdatum, dieses wird aus der aktuell fokussierten Karteikartenzeile bestimmt

		Parameter:
			sequenz	- ein indizierter Array, wobei jeder String für jeweils ein Karteikartenkürzel und dessen Text steht
							  sendesequenz := ["Karteikartenkürzel1|Karteikartentext1"
														,"Karteikartenkürzel2|Karteikartentext2"
														, usw .... ]
								*	Verwenden sie als Karteikartenkürzel nur die Kürzel, welche keine Formulare aufrufen, sondern ausschließlich für die Erfassung von Text, Leistungen
									oder Werten gedacht sind
								*	um mit dem Karteikartentext einen Zeilenendezeichen zu senden, verwenden sie entweder die Autohotkey- alternativ die RegEx-Schreibweise
									dementsprechend also `n oder \n
	*/

	; Fehlerausgabe bei falschem Parameter
		If (StrLen(sequenz) = 0) || !IsObject(sequenz) {
			PraxTT("Der übergebene Parameter für die Funktion:`n" A_ThisFunc "`nist leer oder kein Object.`nDie Aufgabe wird beendet", "6 2")
			return 0
		}

	; SetKeyDelay Einstellungen speichern und andere setzten
		keyDelay   	:= A_KeyDelay
		keyDuration	:= A_KeyDuration
		SendModus	:= A_SendMode
		SendMode Event
		SetKeyDelay, 50, 50

	; Zeilendatum auslesen
		Zeilendatum := AlbisLeseZeilenDatum()
		If !RegExMatch(Zeilendatum, "\d{2}\.\d{2}\.\d{4}") {
			PraxTT("Das Zeilendatum konnte nicht erfasst werden.`nWiederholen Sie die Eingabe nocheinmal.", "6 2")
			return 0
		}

	; Daten werden zeilenweise in die Karteikarte geschrieben
		Loop, % sequenz.Count() {

			; Parameter erstellen
				kuerzel	:= StrSplit(sequenz[A_Index], "|").1
				text     	:= StrSplit(sequenz[A_Index], "|").2
				If InStr(text, "\n")	; \n mit `n erstetzen
					text := StrReplace(text, "\n", "`n")

			; Eingabefokus in die Karteikarte setzen
				If !AlbisKarteikarteFocus("Datum") {
					PraxTT("Beim Übertragen der Daten ist ein Problem aufgetreten.`nEs konnte kein Eingabefocus in die Karteikarte gesetzt werden.", "6 2")
					return 0
				}

			; Daten schreiben
				AlbisSchreibeInKarteikarte("", Zeilendatum, kuerzel, text)

		}

	; SetKeyDelay wiederherstellen
		SendMode % SendModus
		SetKeyDelay, % keyDelay, % keyDuration


}
;12
AlbisZeilendatumSetzen(DateChoosed) {                                                                 	;-- Zeilendatum in der Karteikarte ändern

	; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	; Karteikarte anzeigen
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If InStr(AlbisGetActiveWindowType(), "Karteikarte")
			If !AlbisKarteikarteZeigen() {
				MsgBox, Die Karteikarte ließ sich nicht Anzeigen
				return
			}

	; Karteikarte aktivieren (anklicken)
		AlbisKarteikarteAktivieren()

	; Karteikartenzeile lesen
		KKarte	:= AlbisGetActiveControl("content")

	; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
	; ist kein Eingabefokus eröffnet, wird die Eingabe eröffnet
	; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
		If !IsObject(KKarte) {
			SendInput, {INSERT}
			Sleep 50
			KKarte	:= AlbisGetActiveControl("content")    ; jetzt kann die Zeile gelesen werden
		}

	; Zeilendatum lesen
		ZDatum 	:= AlbisLeseZeilenDatum()
		If (ZDatum <> DateChoosed) {
			AlbisKarteikarteFocus("Datum")
			If !VerifiedSetText(KK.DATUM, DateChoosed, "ahk_class optoAppClass") {
				SendInput, {Home}
				SendInput, {LShift Down}{End}{Shift Up}
				Sleep 100
				Send, % DateChoosed
				Sleep 100
				SendInput, {TAB}
			}
		}

	; Tastatureingabe ins Kürzelfeld
		AlbisKarteikarteFocus("Kuerzel")

return AlbisLeseZeilenDatum() <> DateChoosed ? false : true
}
; 13
AlbisSendText(classnn, Text, hParent:=0) {

	; weiterer Versuch Texteingaben gezielt, zuverlässig und insbesondere schnell in die Karteikarte zu bekommen
	; **End angeben für das Beenden der Eingaben (sendet TAB und ESC)
	;BlockInput, On

	If (hParent = 0)
		hParent := GetParent(GetFocusedControl())
	VerifiedSetFocus(classnn,,, hParent)

	If (Text <> "**End") {
		VerifiedSetText(classnn, Text, hParent)
		ControlSend, % classnn, {Blind}{TAB}, % "ahk_id " hParent
		while (A_Index < 20) {
			If (A_Index > 1)
				Sleep 20
			hFocused	:= GetFocusedControl(), hParent := GetParent(hFocused)
			fclassnn	:= Control_GetClassNN(hParent, hFocused)
			If (classnn = fclassnn)
				break
		}
		SciTEOutput(fclassnn)
		If (classnn <> fclassnn){
			;BlockInput, Off
			return 0
		}
	}
	If (Text = "**End") {
		sleep 100
		ControlSend, % classnn, {ESC}, % "ahk_id " hParent := GetParent(GetFocusedControl())
		while (A_Index < 20) {
			If (A_Index > 1)
				Sleep 20
			hFocused	:= GetFocusedControl(), hParent := GetParent(hFocused)
			fclassnn	:= Control_GetClassNN(hParent, hFocused)
			If (classnn = fclassnn)
				break
		}
		SciTEOutput(fclassnn)
		If (classnn <> fclassnn){
			;BlockInput, Off
			return 0
		}
	}
	;BlockInput, Off
	sleep 200

return hParent
}


;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; FORMULARE                                                                                                                                                                                  	(15)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisDruckeBlankoFormular  	(02) AlbisRezeptHelfer                	(03) AlbisRezeptHelferGui            	(04) AlbisRezeptFelderLoeschen
; (05) AlbisRezeptFelderAuslesen       	(06) AlbisRezeptSchalterLoeschen 	(07) AlbisDruckePatientenAusweis 	(08) AlbisRezeptAutIdem
; (09) AlbisRezept_DauermedikamenteAuslesen                                    	(10) AlbisFormular                       	(11) AlbisHautkrebsScreening
; (12) AlbisFristenRechner                 	(13) AlbisFristenGui                   	(14) IfapVerordnung                   	(15) AlbisWeitereMedikamente
; (16*) AlbisVerordnungsplan         	(17*) AlbisAusfuellhilfe
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;1
AlbisDruckeBlankoFormular(Art:="KR1", BnPlus:=0) {			                                    	;-- Funktion zum Ausdrucken von Blankoformularen

	; Beispiel für Art KR steht für Kassenrezept, PR für Privatrezept, GR grünes Rezept, BR Btm, UB ist die Überweisung, KH - KRankenhauseinweisung, KB - Krankentransport
	; zum Übergeben von Formular und Menge Kürzel dahinter gleich die Anzahl schreiben, zwischen den Formalaren in Leerzeichen lassen
	; AlbisDruckeBlankoFormular("KR3 PR2 KH1 KB1 UB1")
	; um die anderen 3 Buttons (Spooler, Speichern, Abbrechen) zu kommen ist der zweite Parameter da in der Reihenfolge ist Spooler +1 , Speichern +2 und Abbrechen + 3

	AlbisActivate(1)

	MatchMode       	:= A_TitleMatchMode
	MatchModeSpeed	:= A_TitleMatchModeSpeed

	SetTitleMatchMode, 2		;Fast is default
	SetTitleMatchMode, Fast

	static BM_Click	:= 0x00F5
	static form

	If !IsObject(form) {
		form			:= Object()
		form.KR	:= {"MenuID" : 32797, "CheckBox": "Button92", "Edit":"Edit46"	, "WinTitle":"Muster 16"                   	, "classnn": "#32770", "Print":"Button"  (29 +  BnPlus)	, "FName":"Kassenrezept"					}
		form.PR 	:= {"MenuID" : 32845, "CheckBox": "Button68", "Edit":"Edit43"	, "WinTitle":"Privatrezept für Patient" 	, "classnn": "#32770", "Print":"Button"  (19 +  BnPlus)	, "FName":"Privatrezept"						}
		form.GR	:= {"MenuID" : 34219, "CheckBox": "Button92", "Edit":"Edit46"	, "WinTitle":"Muster 16"	                 	, "classnn": "#32770", "Print":"Button"  (29 +  BnPlus)	, "FName":"grünes Rezept"					}
		form.BR	:= {"MenuID" : 34771, "CheckBox": "Button92", "Edit":"Edit46"	, "WinTitle":"Muster 16"                   	, "classnn": "#32770", "Print":"Button"  (29 +  BnPlus)	, "FName":"Btm Rezept"						}
		form.UB	:= {"MenuID" : 32813, "CheckBox": "Button28", "Edit":"Edit8"	, "WinTitle":"Muster 6"	                	, "classnn": "#32770", "Print":"Button"  (1   +  BnPlus)	, "FName":"Überweisung"					}
		form.KH	:= {"MenuID" : 32837, "CheckBox": "Button25", "Edit":"Edit9"	, "WinTitle":"Muster 2"                    	, "classnn": "#32770", "Print":"Button"  (9   +  BnPlus)	, "FName":"Krankenhauseinweisung"	}
		form.KB	:= {"MenuID" : 32926, "CheckBox": "Button37", "Edit":"Edit10"	, "WinTitle":"Muster 4"	                	, "classnn": "#32770", "Print":"Button"  (25 +  BnPlus)	, "FName":"Krankenbeförderung"		}
	}

	items:= StrSplit(Art, A_Space)
	Loop, % items.MaxIndex()	{

			PraxTT("Drucke das " A_Index ". Formular von " items.MaxIndex() " Dokumenten.`n                              " form[A_Index].FName, "60 2 0")
			RegExMatch(items[A_Index], "[a-zA-Z]+", fn)				;fn - Formular Name
			RegExMatch(items[A_Index], "\d", fz)                            ;fz	 - Formular Anzahl
			If !fz
				continue

		;  verhindern das diesem Parameter eine Zahl größer als 4 enthält
			BnPlus:= BnPlus > 4 ? 4 : fz

		; Fenstertitel zusammenstellen
			WinTitle:= form[fn].WinTitle " ahk_class " form[fn].classnn

		; Aufrufen und Prüfen des Formulares
			while !WinExist(WinTitle)			{
				PostMessage, 0x111, % form[fn].MenuID,,, % "ahk_class OptoAppClass"
				WinWaitActive, % WinTitle,, 3
				If hWin:= WinExist(WinTitle)
					break
				If A_Index > 3
					return
			}

		; Datum Checkbox - Haken entfernen wenn gesetzt (Datumsfeld wird nicht mit einem Datum bedruckt)
			hCheckBox:= GetHex(ControlGet("Hwnd",, form[fn].Checkbox, "ahk_id " hWin))
			If ControlGet("Checked",,, "ahk_id " hCheckbox)
				VerifiedCheck("", hCheckbox, "", "", 0)	; 0 = unCheck

		; Anzahl der Formular einsetzen
			VerifiedSetText(form[fn].Edit, fz, WinTitle)

		; Schließen des Formulares durch Ausführen einer Funktion (Drucken, Spoolerdruck, Speichern, Abbrechen)
			VerifiedClick(form[fn].Print, WinTitle)
			WinWaitClose, % WinTitle,, 5
			If ErrorLevel			{
				MsgBox, 4, Addendum für Albis on Windows, % "folgendes Fenster konnte nicht geschlossen werden:`n" form[fn].FName "`n`nBitte schliessen Sie es manuell.`n Der Vorgang wird danach fortgesetzt."
				IfMsgBox, Cancel
					break
			}
	}

	SetTitleMatchMode, % MatchMode
	SetTitleMatchMode, % MatchModeSpeed

return
}
;2
AlbisRezeptHelfer(tag, Schalter) {                                                                            	;-- (Hilfsmittel-/Kassen-/Btm-) Rezepte schreiben ohne eine Liste aufrufen zu müssen
	; Schalter - kann ein String oder eine Zahl sein z.B. Rezepthilfe("ATStrumpf")

	Muster16:= "Muster 16 ahk_class #32770"

	If RegExMatch(Schalter, "BVG")
		Schalter:= 6
	else If RegExMatch(Schalter, "HiM")
		Schalter:= 7
	else If RegExMatch(Schalter, "I")
		Schalter:= 8
	else If RegExMatch(Schalter, "Spr")
		Schalter:= 9
	else If RegExMatch(Schalter, "HeM")
		Schalter:= 10
	else If RegExMatch(Schalter, "BtM")
		Schalter:= 11
	else If RegExMatch(Schalter, "OTC")
		Schalter:= 12
	else
		RegExMatch(Schalter, "\d+", Schalter)

	Loop, 7
		If (Schalter = A_Index + 5)
			VerifiedCheck("Button" 37 + A_Index, Muster16, "", "", 1)
		else
			VerifiedCheck("Button" 37 + A_Index, Muster16, "", "", 0)

	If (StrLen(tag) > 0)	{
		VerifiedSetText("Edit2"	, Rezept[tag][1] , Muster16)
		VerifiedSetText("Edit8"	, Rezept[tag][2] , Muster16)
		VerifiedSetText("Edit14"	, Rezept[tag][3] , Muster16)
	}

return
}
;3
AlbisRezeptHelferGui(dbfile) {                                                                                  	;{- Schnellrezept - Standardrezepte für bestimmte Diagnosen/Hilfsmittel mit einem Klick erstellen

	; blendet ein kleines Listbox-Steuerelement in einem Kassenrezeptformular ein
	; diese Funktion braucht die Bibliothek >> lib\class_JSON.ahk <<
	; letzte Änderung: 23.03.2021 - Anpassung Rezeptfelder - wurden von CG anders nummeriert

	; VARIABLEN                                                                                                                                                        	;{

		global RHStatus    				                    	; flag damit bei Fokusänderungen die Gui nicht neu erstellt wird
		global RHRetry := 0
		global hRezeptHelfer, RH, RHBtn1
		global hMuster16, hMuster16old,
		global Muster16 := Object()
		global Mdi

		static DDLValues, RezeptWahl, wahl, isVisible, hMdiChild, APos, RPos, TPos, CtrlList, hAfxFOV, WerbungH
		static dX, dY, dW, dH, dxE, lpX, NewX, NewY, NewW, NewH
		static RezeptFeld	:= [2,9,16,23,30,37]	; Edit2, Edit8, ....
		static RezeptZusatz	:= [3,7,11,15,19,23]	; Button2, Button6, ....
		static Schalter    	:= {	48 : "BVG"
		                            	, 	49 : "Hilfsmittel"
										,	50 : "Impfstoff"
										,	51 : "Sprechstunden-Bedarf"
										,	52 : "Heilmittel"
										,	53 : "BTM"
										,	54 : "OTC"}
		static Rezepte                                        	; Objekt enthält den Datensatz für die Verordnungen
		static SchnellRezept                                	; speichert die Nummer des aktuell ausgewählten Schnellrezeptes
		static dbfileLastEdit                                	; enthält die letzte Dateiänderungszeit
		static DDLWidth	:= 380

		If RHStatus
			return

		RezeptIstZentriert := false
		WerbungH := 0
	;}


	; VERHINDERT ERNEUTEN AUFRUF DER REZEPTHELFER GUI                                                                                 	;{
		If !(hMuster16 := WinExist("Muster 16 ahk_class #32770")) {

			RHStatus := false
			RHRetry ++
			sleep, 30
			SetTimer, RezepthelferUpdate, Off
			If (RHRetry < 30)
				AlbisRezeptHelferGui(dbfile)

			RHRetry := 0
			;SciTEOutput("Muster 16 nicht gefunden")
			return
		}

		If WinExist("Rezepthelfer ahk_class AutohotkeyGui") && (hMuster16old = hMuster16 )
			return

		hMuster16old := hMuster16
		RHStatus           := true

	;}

	; ERMITTLUNG VON HWND's UND STEUERELEMENT- UND FENSTERPOSITIONEN                                             	;{
	; gebraucht für die Positionierung des Rezeptfensters innerhalb von Albis

		res                   	      	:= Controls("", "Reset", "")
		;Muster16.Werbung	:= Controls("ThunderRT6UserControlDC1", "HWND"                                  	, hMuster16)
		Muster16.StaticText	:= Controls("", "ControlFind, Werbung, Static, return hwnd"                        	, hMuster16)
		Muster16.ArzneiDB	:= Controls("", "ControlFind, Arzneimitteldatenbank, Button, return hwnd"	, hMuster16)
		Muster16.Drucken		:= Controls("", "ControlFind, Drucken, Button, return hwnd"                     	, hMuster16)
		Muster16.EVO   		:= Controls("", "ControlFind, Einnahmeverordnung, Button, return hwnd"  	, hMuster16)
		Muster16.VOP       	:= Controls("", "ControlFind, Verordnungsplan, Button, return hwnd"        	, hMuster16)

		;SciTEOutput(Muster16.EVO ", " Muster16.VOP)

		ControlGet, isVisible     	, Visible ,, ThunderRT6UserControlDC1	, % "ahk_id " hMuster16
		ControlGet, hWerbung	, HWND,, ThunderRT6UserControlDC1	, % "ahk_id " hMuster16
		ControlGet, hDMed   	, HWND,, ListBox1                             	, % "ahk_id " hMuster16
		ControlGet, hStaticW 	, HWND,, ListBox2                             	, % "ahk_id " hMuster16
		Muster16.DMed    	:= hDMed
		Muster16.StaticW  	:= hStaticW
		Muster16.Werbung 	:= hWerbung

		APos                     	:= GetWindowSpot(AlbisWinID())
		RPos                        	:= GetWindowSpot(hMuster16)

		WinGet, CtrlList, ControlListHWND, % "ahk_id " AlbisMDIChildGetActive()
		For idx, ctrlHwnd in StrSplit(CtrlList, "`n") {
			If InStr(GetClassName((hwnd := ctrlHwnd)), "AfxFrameOrView140") {
				TPos := GetWindowSpot(ctrlHwnd)
				break
			}
		}

	;}

	; WERBUNGSANZEIGEN IM REZEPT AUSSCHALTEN - Rezeptfenstergröße wird automatisch angepasst                  	;{

		; Werbeanzeige 1 wird ausgeblendet
		If Muster16.StaticText && Muster16.StaticW && Muster16.DMed {
			Control         	, Hide,,                      	, % "ahk_id " Muster16.StaticText
			ControlGetPos	,, dY,, dH,                  	, % "ahk_id " Muster16.StaticW
			ControlGetPos	,, cY,, cH,                  	, % "ahk_id " Muster16.DMed
			Control         	, Hide,,                     	, % "ahk_id " Muster16.StaticW
			ControlMove   	,,,,, % dY + dH - cY    	, % "ahk_id " Muster16.DMed
		}

		; Werbeanzeige 2 wird ausgeblendet
		If Muster16.Werbung && isVisible {
			ControlGetPos	,,,, WerbungH,	, % "ahk_id " Muster16.Werbung
			Control         	, Hide,,         	, % "ahk_id " Muster16.Werbung
			WerbungH := RPos.H - WerbungH + 17
		}

		;}

	; REZEPT WIRD INNERHALB VON ALBIS ZENTRIERT                                                                                             	;{

		NewX	:= APos.X + Floor(APos.W/2) - Floor(RPos.W/2)
		NewY	:= TPos.Y
		NewW	:= RPos.W
		NewH	:= WerbungH > 0 ? WerbungH : RPos.H
		SetWindowPos(hMuster16, NewX, NewY, NewW, NewH, 0, 0x50)
		;SciTEOutput("> " NewX " := Floor((" APos.W " - " APos.X ")/2) - Floor(" RPos.W "/2)`n> " NewY " := " TPos.Y "`n> " NewW " := " RPos.W "`n>" NewH " := " RPos.H)

	;}

	; SCHNELLREZEPTDATEN EINLESEN                                                                                                                   	;{
	; überprüft ob die Datei verändert wurde, liest die Daten dann neu ein

		FileGetTime, fileTime, % dbfile, M
		If (fileTime <> dbfileLastEdit)
			Rezepte := ""

		DDLValues := "Auswahl eines Schnellrezeptes...|neues Schnellrezept anlegen"
		If !IsObject(Rezepte)
			If FileExist(dbfile) {

				Rezepte	:= JSON.Load(FileOpen(dbfile, "r").Read())
				For i, val in Rezepte
					DDLValues .= "|" val.Bezeichner

			} else {
				PraxTT("keine Schnellrezepte für Rezeptverordnungen vorhanden!", "0 2")
				;Rezepte := Object()
			}

	;}

	; GUI - ZUSÄTZLICHE STEUERELEMENTE IM REZEPT ANZEIGEN                                                                             	;{

		ControlGetPos, cx,,,               	,, % "ahk_id " Muster16.Drucken           	; unterhalb von Drucken
		ControlGetPos, dX, dY, dW, dH	,, % "ahk_id " Muster16.ArzneiDB          	; Schnell-Rezeptposition wird unterhalb des Steuerelementes Arzneimitteldatenbank angezeigt

		try {
			Gui, RH: New, -Caption -DPIScale +ToolWindow +HWNDhRezeptHelfer +E0x00000004 -0x04020000 Parent%hMuster16%
		}
		catch {
			sleep, 200
			RHRetry ++
			SetTimer, RezepthelferUpdate, Off
			If (RHRetry < 30)
				AlbisRezeptHelferGui(dbfile)
			RHRetry := 0
			return
		}
		Gui, RH: Margin, 0, 0
		Gui, RH: Color	, cFFFFFF
		Gui, RH: Font, s8 q5 cBlack, MS Sans Serif

		Gui, RH: Add, Button , % "x0                                	gRezeptVOPlan                                                                       	"  	, % "Verordnungsplan drucken"
		Gui, RH: Add, Button , % "x300  y0                     	gRezeptLeeren       	vRHBtn1     	HWNDhRHelferBtn1             	"	, % "alle Felder leeren"            	; temporäre Position
		Gui, RH: Add, DDL 	, % "x500 y1 w" DDLWidth " 	gRezeptausgewaehlt	vRezeptwahl HWNDhRezeptWahl AltSubmit"	, % DDLValues                        	; temporäre Position

		lp	:= GuiControlPos("RH", "RHBtn1")
		GuiControl, RH: Move, RHBtn1      	, % "x" (lpX := dx + dw - lp.W - 16)
		GuiControl, RH: Move, RezeptWahl	, % "x" (lpX - DDLWidth - 5)
		;SciTEOutput(A_LineNumber ": lp "  cx ", " dx ", " dY ", " dw ", " dH "`n`t  Btn1: " lpX "=" dx "+" dw "-" lp.W "`n`t  gui: x" cx - 8 " y" dY + dH - 25 )

		Gui, RH: Show, % "x" cx - 8 " y" dY + dH - 25 " NoActivate", Rezepthelfer
		GuiControl, RH: Choose, Rezeptwahl, 1

	  ; GUI AKTIVIEREN, GANZ NACH OBEN SETZEN UND NEUZEICHEN - SONST FUNKTIONIERT ES NICHT!
		WinActivate   	 , % "ahk_id " hRezeptHelfer
		WinSet, Top   	,, % "ahk_id " hRezeptHelfer
		WinSet, Redraw	,, % "ahk_id " hRezeptHelfer
		SetTimer, RezepthelferUpdate, 100

		RHStatus := true

	;}

return

RezeptVOPlan:                	;{

		Gui, RH: Destroy
		WinActivate, % "Muster 16 ahk_class #32770"

		VerifiedCheck("",,, Muster16.VOP)
		VerifiedCheck("",,, Muster16.EVO)
		VerifiedClick("",,, Muster16.Drucken)

		WinWait, % "Auswahl weiterer Medikamente ahk_class #32770",, 15
		If !ErrorLevel
			AlbisVerordnungsplan()
		;~ else
			;~ SciTEOutput("Das hat zu lange gedauert")

return ;}

RezeptAusgewaehlt:       	;{

		Gui, RH: Submit, NoHide
		Gui, RH: Default

	; ANLEGEN EINES NEUEN SCHNELLREZEPTES (VORLAGE) ZUR VERWENDUNG BEI ALLEN PATIENTEN
		If (RezeptWahl = 2)		{

			; Felder werden in ein Object eingelesen
				RezeptDaten:= AlbisRezeptFelderAuslesen(hMuster16)

			; Abbruch der Funktion, wenn keine Eintragungen vorhanden sind
				If !RezeptDaten.Medikamente.MaxIndex() 	{
					MsgBox, 4096, % "Addendum Rezepthelfer", % "Das Rezept enthält keine Eintragungen.`nEin Schnellrezept wurde nicht angelegt!", 6
					return
				}

			; falls zuvor ein Schnellrezept ausgewählt war, nachfragen ob dieses geändert werden soll
				If (SchnellRezept > 0) {

					MsgBox, 4, % "Addendum Rezepthelfer", % "Möchten Sie das Schnellrezept '" Rezept[SchnellRezept].Bezeichner "' überschreiben?"
					IfMsgBox, Yes
					{
							Rezept[SchnellRezept].Medikamente	:= RezeptDaten.Medikamente
							Rezept[SchnellRezept].Zusaetze        	:= RezeptDaten.Zusaetze
							Rezept[SchnellRezept].RezeptTyp        	:= RezeptDaten.RezeptTyp

					}

				} else {

				; Comboboxauswahl auf den ersten Eintrag zurücksetzen
					PostMessage, 0x014E, 0, 0, ComboBox1, % "ahk_id " hRezeptHelfer

					MsgBox, 4, % "Addendum Rezepthelfer", % RezeptDaten.Medikamente.MaxIndex() " ausgefüllte Medikamentenfelder sind vorhanden.`nMöchten Sie dieses Rezept als Schnellrezept speichern?"
					IfMsgBox, No
						return

					InputBox, Bezeichner, % "Addendum Rezepthelfer", % "Geben Sie dem neuen Schnellrezept einen eindeutigen Namen"

					Rezept.Push({"Beschreibung"	: " "
									, 	"Bezeichner"  	: Bezeichner
									, 	"Medikamente"	: RezeptDaten.Medikamente
									,	"Zusaetze"      	: RezeptDaten.Zusaetze
									,	"Optionen"      	: ""
									,	"RezeptTyp"      	: RezeptDaten.RezeptTyp})

					rz := FileOpen(dbfile, "w")
					rz.Write(JSON.Dump(Rezept))
					rz.Close()

				}

		}

	; AUSGEWÄHLTES SCHNELLREZEPT ANZEIGEN
		else If (RezeptWahl > 2)		{

			RezeptNr := RezeptWahl - 2

		; Medikamentenfelder und Schalter werden gelöscht
			AlbisRezeptFelderLoeschen(hMuster16)

		; Medikamente eintragen
			For i, medikament in Rezepte[RezeptNr].medikamente 			{
				ControlSetText, % "Edit" RezeptFeld[i], % medikament, % "ahk_id " hMuster16
				If InStr(medikament, "...")		{
					focusFeld	:= "Edit" RezeptFeld[i]
					focusPos	:= InStr(medikament, "...")
				}
			}

		; Schalter entsprechend setzen
			RezeptTyp:= Rezepte[RezeptNr].RezeptTyp
			If (StrLen(RezeptTyp) > 0)
				For Nr, SchalterBezeichnung in Schalter
					If InStr(SchalterBezeichnung, RezeptTyp)
						VerifiedCheck("Button" Nr, "", "",  hMuster16, 1), 	break

		; Schnellrezept Nr zwischenspeichern
			SchnellRezept := RezeptNr

		; Rezepthelfer-Gui nach vorne holen
		   	WinActivate   	 , % "ahk_id " hRezeptHelfer
			WinSet, Top   	,, % "ahk_id " hRezeptHelfer
        	WinSet, Redraw	,, % "ahk_id " hRezeptHelfer

		; Tastaturfokus auf das Feld setzen welches '...' enthält - Nutzer möchte dort noch eine weitere Eingabe eintragen
			If (StrLen(FocusFeld) > 0) {

					VerifiedSetFocus(FocusFeld, "", "", hMuster16)
					ControlGetText	, text, % FocusFeld, % "ahk_id " hMuster16

					select	:= "..."
					cpos 	:= InStr(text, select) - 1
					SelW 	:= StrLen(select)

     				SendMessage 0xB1, % cpos, % cpos+SelW, % FocusFeld, % "ahk_id " hMuster16 ; EM_SETSEL

			}
	}

return ;}

RezeptLeeren:                	;{
	MsgBox, 4, Addendum für Albis on Windows, Möchten Sie wirklich alle Rezeptfelder löschen?
	IfMsgBox, No
		return
	AlbisRezeptFelderLoeschen(hMuster16)
return ;}

RezepthelferUpdate:      	;{
	If !WinExist("Muster 16 ahk_class #32770") {
		AlbisRezeptHelferGuiClose()
		RHStatus := false
	}
return ;}
}
;3.1
AlbisRezeptHelferGuiClose() {

	global RHStatus, RH, hRezeptHelfer, RHRetry

	SetTimer, RezepthelferUpdate, Off
	Gui, RH: Destroy
	IfWinExist, Addendum Rezepthelfer ahk_class AutohotkeyGui
		WinClose, Addendum Rezepthelfer ahk_class AutohotkeyGui
	RHStatus	:= false
	RHRetry 	:= 0

return
}
;3.2
GuiControlPos(GuiName, ControlName) {
	GuiControlGet, cp, % GuiName ": Pos", % ControlName
return {"X":cpX, "Y":cpY, "W":cpW, "H":cpH}
}   ;}
;4
AlbisRezeptFelderLoeschen(hMuster16) {                                                                	;-- (primär) Hilfsfunktion für die Rezepthelfer Gui

	static RezeptFeld	:= [2,8,14,20,26,32]	; Button
	static RezeptZusatz	:= [3,7,11,15,19,23]	; Button

	; löscht die Zusatztexte
		Loop, % RezeptZusatz.MaxIndex() 	{

			ControlGetText, cText, % "Button" RezeptZusatz[A_Index], % "ahk_id " hMuster16
			If InStr(cText, "zus")		{
				VerifiedClick("Button" RezeptZusatz[A_Index], "ahk_id " hMuster16)
				WinWait, Medikamentenzusätze ahk_class #32770,, 5
				If !ErrorLevel	{
					ControlSetText, % "Edit1", % "", Medikamentenzusätze ahk_class #32770
					ControlSetText, % "Edit2", % "", Medikamentenzusätze ahk_class #32770
					VerifiedClick("Button2", "Medikamentenzusätze ahk_class #32770")
					WinWaitClose, Medikamentenzusätze ahk_class #32770,, 5
					If ErrorLevel	{
						while WinExist("Medikamentenzusätze ahk_class #32770")
							MsgBox, Bitte das Fenster 'Medikamentenzusätze' schließen!
					}
					ControlSetText, % "Button" RezeptZusatz[A_Index], % "...", % "ahk_id " hMuster16
				}
			}

		}

	; löscht die Medikamentenfelder
		Loop, % RezeptFeld.MaxIndex() {
			nr := A_Index
			ControlSetText, % "Edit" RezeptFeld[nr],, % "ahk_id " hMuster16
			Loop, 4
				ControlSetText, % "Edit" RezeptFeld[nr] + A_Index,, % "ahk_id " hMuster16
		}

	; löscht die Schalter
		AlbisRezeptSchalterLoeschen(hMuster16)

}
;5
AlbisRezeptFelderAuslesen(hMuster16) {                                                                 	;-- (primär) Hilfsfunktion für die Rezepthelfer Gui

		static RezeptFeld	:= [2,8,14,20,26,32]	; Edit
		static RezeptZusatz	:= [2,6,10,14,18,22]	; Button
		static Schalter    	:= {38:"BVG", 39:"Hilfsmittel", 40:"Impfstoff", 41:"Sprechstunden-Bedarf", 42:"Heilmittel", 43:"BTM", 44:"OTC"}
				RezeptDaten 	:= {"Medikamente":[], "Zusaetze":[], "RezeptTyp":""}

		Loop, % RezeptZusatz.MaxIndex() 	{

			ControlGetText, val	, % "Edit" RezeptFeld[A_Index]     	, % "ahk_id " hMuster16
			ControlGetText, cText, % "Button" RezeptZusatz[A_Index]	, % "ahk_id " hMuster16

			If (StrLen(val) > 0)
				RezeptDaten.Medikamente.Push(val)

			If InStr(cText, "zus") {
				VerifiedClick("Button" RezeptZusatz[A_Index], "ahk_id " hMuster16)
				ControlGetText, val, % "Edit1", % "", Medikamentenzusätze ahk_class #32770
				RezeptDaten.Zusatz.Push(val)
				VerifiedClick("Button2", "Medikamentenzusätze ahk_class #32770")
				WinWaitClose, Medikamentenzusätze ahk_class #32770,, 5
				If ErrorLevel 	{
					while WinExist("Medikamentenzusätze ahk_class #32770")
						MsgBox, Bitte das Fenster 'Medikamentenzusätze' schließen!
				}
			}
		}

	; bricht nach dem ersten gesetzten Schalter im Rezept ab
		For Nn, SchalterBezeichnung in Schalter
			If ControlGet("Checked", "", "Button" Nn, "ahk_id " hMuster16) {
				RezeptDaten.RezeptTyp := SchalterBezeichnung
				break
			}
	;

return RezeptDaten
}
;6
AlbisRezeptSchalterLoeschen(hMuster16) {                                                                	;-- (primär) Hilfsfunktion für die Rezepthelfer Gui

	static Schalter := {38:"BVG", 39:"Hilfsmittel", 40:"Impfstoff", 41:"Sprechstunden-Bedarf", 42:"Heilmittel", 43:"BTM", 44:"OTC"}

	For Nn, SchalterBezeichnung in Schalter
		If ControlGet("Checked", "", "Button" Nn, "ahk_id " hMuster16)
			VerifiedCheck("Button" Nn, "", "", hMuster16, 0)

return
}
;7
AlbisDruckePatientenAusweis() {                                                                             	;-- druckt den Patientenausweis

		static BM_Click	:= 0x00F5

		MatchMode       	:= A_TitleMatchMode
		MatchModeSpeed	:= A_TitleMatchModeSpeed
		SetTitleMatchMode, 2		;Fast is default
		SetTitleMatchMode, Fast

		AlbisCloseLastActivePopups(AlbisWinID())

		If !WinExist("Patientenausweis ahk_class #32770")
			PostMessage, 0x111, 33003,,, % "ahk_id " AlbisWinID()

		WinWait, Patientenausweis ahk_class #32770,, 3
		If ErrorLevel 	{
			nichts := 0
		}

		;PostMessage, % BM_Click,,,, % "ahk_id " ControlGet("Hwnd", "Button23", "Patientenausweis ahk_class #32770")
		VerifiedClick("", ControlGet("Hwnd", "Button23", "Patientenausweis ahk_class #32770"))
		WinWaitClose, Patientenausweis,, 10
		while WinExist("Patientenausweis ahk_class #32770")	{
			MsgBox,, Addendum für Albis on Windows, Bitte das Patientenausweis-Fenster schließen,`nbevor weiter gemacht werden kann!
			If (A_Index > 10)
				break
		}

		SetTitleMatchMode, % MatchMode
		SetTitleMatchMode, % MatchModeSpeed

return
}
;8a
AlbisRezeptAutIdem() {                                                                                            	;-- setzt automatisch ein 'aut idem' - Häkchen

	; !!Voraussetzung!! - In den Dauermedikamenten ist im letzten Dosisfeld (Nacht) ein 'A' vermerkt
		static Dosisfeld  	:= [6,12,18,24,30,36]  	; Edit
		static RezeptFeld	:= [2,8,14,20,26,32]    	; Edit
		static Idemfeld   	:= [52,53,54,55,56,57]	; Button
		static f, blinks, hMuster16

		global hRezeptOverlay, AutIdem
		global RezeptOverlay:= Object()

		Felder       	:= Object()
		hMuster16	:= GetHex(WinExist("Muster 16 ahk_class #32770"))
		RPos         	:= GetWindowSpot(hMuster16)
		hActive      	:= GetHex(WinActive("A"))
		f                	:= 1
		blinks        	:= 0


	/*  Untersucht die Steuerelemente im Rezeptformular, setzt bei Bedarf das Aut-Idem Häkchen
		 die Aut-Idem Steuerelemente im Rezeptfenster sind keine Standard-Checkboxen, mit dem Control-Befehl läßt sich der Status nicht auslesen
		 es wird deshalb mit SendMessage gearbeitet
	 */
		Loop, % Dosisfeld.MaxIndex() {
			ControlGetText, Nacht       	, % "Edit"    	Dosisfeld[A_Index] 	, % "ahk_id " hMuster16
			ControlGetText, Medikament	, % "Edit"   	Rezeptfeld[A_Index]	, % "ahk_id " hMuster16
			ControlGet, hAutIdem, hwnd,, % "Button" 	Idemfeld[A_Index]  	, % "ahk_id " hMuster16
			SendMessage, 0xF2, 0, 0,, % "ahk_id " hAutIdem      ; BM_GETSTATE
			isChecked := ErrorLevel
			If RegExMatch(Nacht, "[A]") && !isChecked && (StrLen(Medikament) > 0)				{
				Control, Check,,, % "ahk_id " hAutIdem
				t.= Medikament "`n"
				ControlGetPos, ax,ay,aw,ah, % "Button" Idemfeld[A_Index]	, % "ahk_id " hMuster16
				ControlGetPos, rx, ry, rw, rh, % "Edit"    Rezeptfeld[A_Index]	, % "ahk_id " hMuster16
				Felder.Push({"Idemfeld":Idemfeld[A_Index]
									, "ax": (ax + 1)	, "ay": (ay+1)		, "aw": (aw - 2)	, "ah": (ah - 2)
									, "rx" : (rx + 1)  	, "ry" : (ry +1) 	, "rw":  (rw - 3)  	, "rh":  (rh - 3)})
			}
		}

	; Anzeige eines Hinweisfensters das ein Aut-Idem Kreuz gesetzt wurde
		If (Felder.MaxIndex() > 0) {

			RPos         	:= GetWindowSpot(hMuster16)
			Font          	:= "Futura Bk Bt"
			Options    	:= "cFFFFFFFF italic s9 s5"

			Gui, AutIdem: New	 , -Caption +AlwaysOnTop +E0x080800AC +HWNDhRezeptOverlay ;Parent%hMuster16%
			Gui, AutIdem: Margin , 0, 0
			Gui, AutIdem: Show	 , % "x" RPos.X " y" RPos.Y " NoActivate"
			WinSet, ExStyle	, 0x080800AC	, % "ahk_id " hRezeptOverlay
			WinSet, Style 	, 0x54020000	, % "ahk_id " hRezeptOverlay
			Winset, Disable	,                   	, % "ahk_id " hRezeptOverlay
			WinSet, Top   	,                    	, % "ahk_id " hRezeptOverlay

			If !pToken
					pToken	:= Gdip_Startup()

			If !hFamily := Gdip_FontFamilyCreate(Font)  {
				MsgBox, % " font: " Font " does not Exist on this system"
				ExitApp
			  }

			hbm      	:= CreateDIBSection(RPos.W, RPos.H)
			hdc       	:= CreateCompatibleDC()
			obm      	:= SelectObject(hdc, hbm)
			pGraphics	:= Gdip_GraphicsFromHDC(hdc)
			pBrush  	:= Gdip_BrushCreateSolid(0x33FF0000)
			RezeptOverlay := {"hdc":hdc, "hbm":hbm, "obm":obm, "pGraphics":pGraphics, "pBrush":pBrush }


			Gdip_SetSmoothingMode(pGraphics, 5)
			Gdip_SetInterpolationMode(pGraphics, 7)

			ControlGetPos, tx,ty,tw,th	, % "Button51" 	, % "ahk_id " hMuster16
			ControlGetPos, ex,,ew,   	, % "Edit2"	    	, % "ahk_id " hMuster16
			Gdip_FillRoundedRectangle(pGraphics, pBrush, tx+tw+3, ty-2, ex+ew-tx-tw-5, 24, 3)
			Gdip_TextToGraphics(pGraphics, "'aut idem' Kreuze wurden automatisch gesetzt !", "x" tx+tw+5 " y" ty 	" w" ex+ew-tx-tw-10 " Center cFF000000 s18 s5", Font)

			For i, feld in Felder				{
				Gdip_FillRoundedRectangle(pGraphics, pBrush, feld.ax	, feld.ay	, feld.aw	, feld.ah	, 2)
				Gdip_FillRoundedRectangle(pGraphics, pBrush, feld.rx	, feld.ry 	, feld.rw	, feld.rh 	, 2)
				Gdip_TextToGraphics(pGraphics, "aut idem"   	, "x" feld.rx+1 " y" feld.ry-1                	" w" feld.rw-3 " Right " Options, Font)
				Gdip_TextToGraphics(pGraphics, "Medikament"	, "x" feld.rx+1 " y" feld.ry+feld.rh-11	" w" feld.rw-3 " Right " Options, Font)
			}

			UpdateLayeredWindow(hRezeptOverlay, hdc, RPos.X, RPos.Y , RPos.W, RPos.H)
			SetTimer, AutIdemBlinking, 600

		}

return

AutIdemBlinking:    	;{

	If !WinExist("ahk_id " hMuster16)
		AlbisAutIdemGuiClose()

	f:= (-1)*f
	alpha:= f > 0 ? 0 : 255
	Blinks ++

	Loop, 10 {
		alpha += f*25
		UpdateLayeredWindow(hRezeptOverlay, RezeptOverlay.hdc, RPos.X, RPos.Y , RPos.W, RPos.H, alpha)
		Sleep, 10
	}

	If !WinExist("ahk_id " hMuster16)
		AlbisAutIdemGuiClose()

	If Blinks > 20
		AlbisAutIdemGuiClose()

return ;}
}
AlbisAutIdemGuiClose() {

	SetTimer, AutIdemBlinking, Off
	Gui, AutIdem: Destroy
	SelectObject(RezeptOverlay.hdc, RezeptOverlay.obm)
	DeleteObject(RezeptOverlay.hbm)
	DeleteDC(RezeptOverlay.hdc)
	Gdip_DeleteGraphics(RezeptOverlay.pGraphics)

return
}
;8b
AlbisRezeptHook(hMuster16) {                                                                               	;-- gehört zu AlbisRezeptAutIdem - zuständig für die Erkennung von Änderungen im Rezeptformular

	global RezeptHook
	global RezeptProcAdr
	global hRezept

	static EVENT_SKIPOWNTHREAD				:= 0x0001
	static EVENT_SKIPOWNPROCESS			:= 0x0002
	static EVENT_OUTOFCONTEXT				:= 0x0000
	;EVENT_OBJECT_LOCATIONCHANGE := 0x800B

	; Hook soll nur einmal pro hwnd gesetzt werden
	If (hMuster16 = hRezept)
		return

	hRezept         	:= hMuster16
	AlbisPID        	:= AlbisPID()
	idThread        	:= DllCall("GetWindowThreadProcessId", "Int", hMuster16, "UInt*", AlbisPID)

	RezeptProcAdr	:= RegisterCallback("AlbisRezeptHookProc", "F")
	RezeptHook 	:= SetWinEventHook( 0x800E, 0x800E, 0, RezeptProcAdr, AlbisPID, idThread, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
	RezeptHook 	:= SetWinEventHook( 0x0003, 0x0003, 0, RezeptProcAdr, AlbisPID, idThread, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
	RezeptHook 	:= SetWinEventHook( 0x8001, 0x8001, 0, RezeptProcAdr, AlbisPID, idThread, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS ) ;RezeptWindow is Closed
	RezeptHook 	:= SetWinEventHook( 0x800B, 0x800B, 0, RezeptProcAdr, AlbisPID, idThread, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS ) ;RezeptWindow is Closed

	If !ErrorLevel {
		RegExMatch(WinGetTitle(hMuster16), "\<.*?\>", match)
		return "Rezeptformularhook: " match
	}

return "Rezeptformularhook: Fehler"
}
;8c
AlbisRezeptHookProc(hHook, event, hwnd, idObject, idChild) {                                 	;-- zugehöriger HookProc-Prozeß zu AlbisRezeptHook

	; 32782 = 0x800E

	global RezeptHook
	global RezeptProcAdr
	global hRezept

	If !WinExist("ahk_id " hRezept) {

		; Hooks beenden
			UnhookWinEvent(RezeptHook, RezeptProcAdr)
			RezeptHook   	:= ""
			RezeptProcAdr	:= ""

		; AutIdem Overlay Gui beenden
			If WinExist("ahk_id " hRezeptOverlay)
					AlbisAutIdemGuiClose()

		; AutIdem Overlay Gui beenden
			If WinExist("ahk_id " hRezeptHelfer)
					AlbisRezeptHelferGuiClose()

			TrayTip("Rezept wurde geschlossen", "HelferGui und Overlay wurden geschlossen", 2)
			Menu, Tray, Tip, % StrReplace(A_ScriptName, ".ahk") " V." Version " vom " vom "`nRezeptformularhook: entfernt!"
			return
	}

	If (event = 3) || (event = 32782) {
		;If WinExist("Verordnungen von") || WinExist("Hinweise")
		;	return 0
		TrayTip("Rezeptüberprüfung", "Änderungen im Rezept werden überprüft`n" hRezept)
		AlbisRezeptAutIdem()
	} else if (event = 32779) {
		If WinExist("ahk_id " hRezeptOverlay)
			 AlbisAutIdemGuiClose()
	}

return 0
}
;9
AlbisRezept_DauermedikamenteAuslesen(hMuster16) {                                           	;-- liest die Listbox1 (Dauermedikamente) im Rezeptfenster in ein Objekt

		caveBereich	:= false
		Med          	:= Object()
		Med.Cave   	:= Object()
		Med.Dauer  	:= Object()

	; Listbox als Text einlesen, parsen und ein Objekt mit den Daten erstellen
		ControlGet, medlist, List,, ListBox1, % "ahk_id " hMuster16

		Loop, Parse, medList, `n, `r
		{
			If RegExMatch(A_LoopField, "#+\s+([A-Za-zÄÖÜäöüß]+)\s", Gruppierung) {
				If InStr(Gruppierung1, "Problem") || InStr(Gruppierung1, "Allergie")
					caveBereich := true
				else
					caveBereich := false
				continue
			}

		/*  Problemmedikamentenbereich erkannt dann

			 Ich nutze die ersten Zeilen des Dauermedikamentenfenster um dort Medikamente anzuzeigen, auf welche der Patient allergisch reagiert hat
			 oder ich notiere dort Medikamente, welche anderweitig Probleme verursacht haben oder aber verursachen könnten.
			 Beispiel für das Aussehen im Dauermedikamentenfenster:

				############## Problemmedikamente #########*
				Penicillin,Amoxicillin,Cotrim, Cefpodoxim - Quincke Ödem*
				Diclofenac, Etoricoxib, Lidocain - Quincke Ödem*
				ACC, Imeron 350(Coro),  ACE-Hemmer - Quincke Ödem*
				Ciprofloxacin (Fluorchinolone) - zentralnervöse Störung + Psychose!!*
				Metoprolol nicht absetzen - Tachykardie sonst*
				Tomaten, Nüsse* - Allergie bis hinzu Quincke Ödem*
				############### Schilddrüse  ###############*
				L Thyroxin 75 Henning TAB N3 100 St (1---)
				...
				...
				...

			** Bei manchen Patienten kommen dermassen viele Medikamente zusammen. Dies ist der Grund warum ich vom Bundeseinheitlichen Medikationsplan abrate.
				Dieser hat keine Möglichkeit die vielen Allergien anzuzeigen. Ich drucke Medikamente als Stammblatt für den Patient. Das geht erstens sehr schnell. Zweitens
				sind alle Zeilen enthalten und drittens kann man die Dauerdiagnosen ebenfalls ausdrucken. Das Eintragen der Problemmedikamente unter den Dauermedikamten
				hat außerdem den entscheidenden Vorteil das diese beim Erstellen eines Briefes (z.B. Krankenhauseinweisung mit kurzer Epikrise) vom Krankenhausarzt an
				prominenter Stelle gelesen werden können.

		 */
			If caveBereich {
				RegExMatch(A_LoopField, "(?<Medikament>[A-Za-zÄÖÜäöüß\-\,\s\(\))]+)\s\-\s(?<Text>.*)", cave)
				caveText	:= Trim(caveText)
				caveText	:= RTrim(caveText, "*")
				If StrLen(caveText) = 0
					continue

				medis      	:= StrSplit(caveMedikament, ",")
				caveExist 	:= false

				For index, cave in Med.Cave
					If InStr(cave.text, caveText)	{

						Loop % medis.MaxIndex()
							cave.Medikament.Push(Trim(medis[A_Index]))

						caveExist := true
						break
					}

					If !caveExist
						Med.Cave.Push({"medikament": medis, "text": caveText})
			}
		; die eigentlichen Medikamente werden extra gesichert
			else
				Med.Dauer.Push(A_LoopField)
		}

return Med
}
;10
AlbisFormular(Formularbezeichnung) {                                                                    	;-- zum Aufrufen von Formularen per Kurzbezeichner

	static formular := {"eHKS_ND"                       	: {"command" : 34505, "WinTitle" : "Hautkrebsscreening - Nichtdermatologe ahk_class #32770"}
								,"GVU"                              	: {"command" : 33212, "WinTitle" : "Muster 30 ahk_class #32770"}
								,"Cave"                              	: {"command" : 32778, "WinTitle" : "Cave! von ahk_class #32770"}
								,"Abrechnung_vorbereiten"    	: {"command" : 32944, "WinTitle" : "Abrechnung KVDT vorbereiten ahk_class #32770"}}

	If !formular.HasKey(Formularbezeichnung)	{
		throw Exception("Formularbezeichnung: " Formularbezeichnung " ist unbekannt!")
		return 0
	}

return Albismenu(formular[Formularbezeichnung].command, formular[Formularbezeichnung].WinTitle)
}
;11
AlbisHautkrebsScreening(Pathologie:="opb", PlusGVU=true, WinClose=true ) {       	;-- befüllt das eHautkrebsScreening (nicht Dermatologen) Formular

	; das Hautkrebsscreeningformular für Nichtdermatologen wurde mehrfach geändert in den letzten Quartalen
	; die ClassNN der Buttons hat sich dabei jedesmal geändert, damit wurden nicht die richtigen Buttons angesprochen und somit war das Formular nicht korrekt ausgefüllt
	; oder ließ sich nicht speichern da sich auch für den Speichern Button die ClassNN geändert hatte
	; letzteres versuche ich nun mittels Übergabe des Namen/Text des Speichern-Buttons treffsicherer zu machen
	; prinzipiell müsste diese Funktion wesentlich flexibler gestaltet werden

		eHKS_Title := "Hautkrebsscreening - Nichtdermatologe ahk_class #32770"

		If !(hHKS := WinExist(eHKS_Title))		{
			PraxTT("Das Formular 'Hautkrebsscreening - Nichtdermatologe' ist nicht geöffnet!`nDie Funktion wird abgebrochen.", "3 1")
			return 0
		}

	; markiert alle Verdachtsdiagnosen als nein
		If InStr(Pathologie, "opB")		{
			VerifiedClick("Button28", hHKS)	; nein - "Verdachtsdiagnose"  - warum Fehler wenn nicht ausgewählt
			VerifiedClick("Button10", hHKS)	; nein - "Malignes Melanom"
			VerifiedClick("Button12", hHKS)	; nein - "Basalzellkarzinom"
			VerifiedClick("Button14", hHKS)	; nein - "Spinozelluläres Karzinom"
			VerifiedClick("Button24", hHKS)	; nein - "Anderer Hautkrebs"
			VerifiedClick("Button26", hHKS)	; nein - "Sonstiger dermatologisch abklärungsbedürftiger Befund"
			VerifiedClick("Button30", hHKS)	; nein - "Screening-Teilnehmer wird an einen Dermatologen überwiesen:"
			VerifiedClick("Button28", hHKS)	; nein - "Verdachtsdiagnose"  - warum Fehler wenn nicht ausgewählt
		}

	; Häkchen bei "gleichzeitige Gesundheitsvorsorge setzen"
		If PlusGVU
			VerifiedClick("Button16", hHKS)

	; "Speichern" und damit auch Schließen des Formulares
		If WinClose		{
			ResultClose := VerifiedClick("Speichern", hHKS)
			Winwait, % "Albis ahk_class #32770", % "Folgende Fehler sind bei der Plausibilitätsprüfung", 1
			If WinExist("Albis ahk_class #32770", "Folgende Fehler sind bei der Plausibilitätsprüfung")
				return 0

			; falls sich das Formular nicht speichern/schließen ließ, wird keine 1 als Erfolgsmeldung zurück senden
			If WinExist(eHKS_Title)
				return 0
		}

return 1
}
;12
AlbisFristenRechner(AUSeit:="", AUBis:="") {                                                           	;-- errechnet das Datum des Beginns und des Ende der Krankengeldfortzahlung

	/*
		TageBKg   	= Tage bis Krankengeldzahlungsbeginn
		WochenBKg	= Wochen bis Krankengeldzahlungsbeginn
		ZeitBKg     	= Zeit bis Krankengeldzahlungsbeginn (Tage)
		ZeitAKg     	= Zeit ab Krankengeldzahlungsbeginn

		letzte Änderung: 24.06.2021
	*/

		static Muster1a 	:= "Muster 1a ahk_class #32770"

		bgStr1       	:= "ab d."
		endStr1	    	:= ""
		bgStr2      	:= "bis zum"
		endstr2     	:= ""
		Fristen			:= Object()

		If !AUSeit || !AUBis 	{
			If !WinExist(Muster1a)
				return 0
			ControlGetText, AUSeit	, Edit1, % Muster1a
			ControlGetText, AUBis	, Edit2, % Muster1a
			If (AUSeitO = AUSeit) && (AUBisO = AUBis)
				 return
			AUSeitO := AUSeit, AUBisO := AUBis
		}

		startdate:= StrSplit(AUSeit, ".").3 StrSplit(AUSeit, ".").2 StrSplit(AUSeit, ".").1
		enddate:= StrSplit(AUBis, ".").3 StrSplit(AUBis, ".").2 StrSplit(AUBis, ".").1
		KrankengeldStart	:= startdate + 0
		KrankengeldEnde	:= startdate + 0
		KrankengeldStart 	+= 41 - 1 	 , days		; Fristenrechner BARMER,AOK,TKK rechnen einen Tag ab, deshalb - 1
		KrankengeldEnde 	+= (78*7) - 1, days
		FormatTime, KrankengeldStart, % KrankengeldStart, dd.MM.yyyy
		FormatTime, KrankengeldEnde, % KrankengeldEnde, dd.MM.yyyy
		FormatTime, Heute, , dd.MM.yyyy
		TageBisBeginn	:= DateDiff("dd", Heute, KrankengeldStart) + 0
		TageBisAblauf	:= DateDiff("dd", Heute, KrankengeldEnde) + 0
		KgStartVorbei	:= TageBisBeginn	< 0 ? 1 : 0
		KgEndeVorbei	:= TageBisAblauf	< 0 ? 1 : 0
		TageBisBeginn	:= TageBisBeginn	< 0 ? -1 * TageBisBeginn : TageBisBeginn
		TageBisAblauf	:= TageBisAblauf 	< 0 ? -1 * TageBisAblauf : TageBisAblauf

		If (TageBisBeginn >= 21) && !KGStartVorbei		{
			TageBKg   	:= Mod(TageBisBeginn, 7)
			WochenBKg	:= Floor((TageBisBeginn - TageBkg)/7)
			ZeitBKg     	:= WochenBKg = 0 ? "(" : WochenBkg = 1 ? "(1 Woche": "(" WochenBKg " Wochen"
			ZeitBKg     	.= TageBKg = 0 ? ") " : TageBkg = 1 ? "  u. 1 Tag) ": " u. " TageBKg " Tage) "
		}
		else if (TageBisBeginn < 21) && !KGStartVorbei
			ZeitBKg      	:= "(" TageBisBeginn " Tage) "

		if KGStartVorbei		{
			bgStr1      	:= "seit d."
			endstr1     	:= ""
		}

		If (TageBisAblauf >= 21) && !KgEndeVorbei 	{
			TageAKg   	:= Mod(TageBisAblauf, 7) + 0
			WochenAKg	:= Floor((TageBisAblauf - TageAKg)/7) + 0
			ZeitAKg     	:= WochenAKg = 0 ? "(" : WochenAkg = 1 ? "(1 Woche": "(" WochenAKg " Wochen"
			ZeitAKg     	.= TageAKg = 0 ? ") " : TageAkg = 1 ? " u. 1 Tag) ": " u. " TageAKg " Tage) "
		}
		else if (TageBisAblauf < 21) && !KgEndeVorbei
			ZeitAKg     	:= "(" TageBisAblauf " Tage)"

		If KgEndeVorbei	{
			ZeitAKg := ZeitBKg := bgStr1 := endStr1 := ""
			bgStr2      	:= "ist am"
			endstr2     	:= "ausgelaufen"
		}

		Fristen.KgStart   	:= Trim(KrankengeldStart)
		Fristen.KgEnde  	:= Trim(KrankengeldEnde)
		Fristen.ZeitBKg  	:= ZeitBKg
		Fristen.ZeitBKgS		:= RegExReplace(ZeitBKg, "(\d+\s+)Wochen\s+u\.\s(\d+\s+)Tage", "$1W $2d")
		Fristen.ZeitAKg  	:= ZeitAKg
		Fristen.ZeitAKgS  	:= RegExReplace(ZeitAKg, "(\d+\s+)Wochen\s+u\.\s(\d+\s+)Tage", "$1W $2d")
		Fristen.BKgStr1  	:= bgStr1
		Fristen.BKgStr2  	:= endStr1
		Fristen.AKgStr1  	:= bgStr2
		Fristen.AKgStr2  	:= endStr2
		Fristen.WinTitle  	:= "KGZ vom " Fristen.KgStart " " Fristen.ZeitBKgS " bis " Fristen.KgEnde " " Fristen.ZeitAKgS
		Fristen.Anzeige   	:= Fristen.BKgStr1 " " Fristen.KgStart " " Fristen.ZeitBKg Fristen.AKgStr1 " " Fristen.KgEnde " " Fristen.ZeitAKg " " Fristen.AKgStr2

return Fristen
}
;13
AlbisFristenGui() {                                                                                                  	;-- zeigt die Fristen auf dem 'Muster 1a' - Arbeitsunfähigkeitsbescheinigung an

	; diese Funktion wird durch einen WinEventHook gestartet, das Gui wird erst geschlossen wenn der 'Muster 1a' Dialog geschlossen wird
	; schließt bei Bedarf auch den Shift+F3 Kalender
	; letzte Änderung: 24.06.2021

		global hMyCal, newCal
		static cTT       	:= Object()
		static Fristen  	:= Object()
		static Muster1a 	:= "Muster 1a ahk_class #32770"
		static Start, Ende, hStart, hEnde, hFrist, Termine, Frist, hOver, info, FristHInfo
		static AUSeitO, AUBisO, fU:= 0
		static hInfoIcon 	:= ImageFromBase64(true, "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAA"
																		. "lwSFlzAAAA5QAAAOUBj+WbPAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vu"
																		. "PBoAAAFcSURBVDiNlZPPSgJRFMa/O84gOIt8h0pS0PzzDA75ALqwJ7EnyEX2DC2qhZsINChoZ"
																		. "YsEE0ZCBm1hgua4choXcu94W0ST03ihOavvcr7zu9+Fe8h910zrnU5d7/V2KaUE/yhFUXgqlR"
																		. "pmM5kiOb94MKrV0xhfc4+pcFTATjQKAPi0LDSaDU+fSASVyokhG31j/+8wIQTFUgnhcBgAsF"
																		. "qt0LxrgvNfH19z9F/7MYk5/ticc0w/pu55Mpl4hn+KOZRIoneas5mr5+ZcZIMQYFmWqxfWIjjAt"
																		. "peuXtp2cIDDmKupw0Q2MYBuABilwQFsA+AwRwiQRY3WUwvj8TsAYDQaBU+gRiLQ8hq0vAZ"
																		. "VVYMnKJePkcvlvg+EoFY7255ADin+LwbAnJtbtef2kMLleOJgcHNLfMt0fXmFt+EQANB+bvuGi"
																		. "UQQT8QN8vgyO+x0u/Weru8FWudkcpDOZItfoQSYA7zlEWMAAAAASUVORK5CYII=")
		If !info {
		info =
		(LTrim
		Entgeltfortzahlung
		Der Ereignistag für die Berechnung der Fristen für die Entgeltfortzahlung wird wie folgt bestimmt:
		Wenn die Arbeitsunfähigkeit schon vor Beginn der Arbeit eingetreten ist,
		ist der Ereignistag der Tag vor dem 1. Arbeitsunfähigkeitstag.
		Hat der Arbeitnehmer am ersten Tag der Arbeitsunfähigkeit noch gearbeitet,
		entspricht der Ereignistag dem 1. Arbeitsunfähigkeitstag

		Krankengeld
		1. Der Anspruch auf Krankengeld entsteht bei stationärer Behandlung von ihrem Beginn an und im Übrigen von dem Tag
			der ärztlichen Feststellung der Arbeitsunfähigkeit an. Die Frist für den Beginn auf Krankengeld kann vom Beginn der Frist
			auf Entgeltfortzahlung abweichen. Anrechenbare Vorerkrankungen sind in den Ausführungen nicht berücksichtigt.
		2. Die Krankenversicherung übernimmt ab der 6. Woche die Krankengeldfortzahlung.
		3. Grundsätzlich gilt, dass das Krankengeld wegen derselben Erkrankung erst einmal relativ lange läuft – nämlich 78 Wochen
			oder 19,5 Monate lang innerhalb von drei Jahren (§ 48 SGB V). Dabei müssen Sie nicht am Stück krankgeschrieben sein.
			Die Zeiträume werden zusammengezählt.
		4. Auf das Krankengeld müssen Sie keine Steuern zahlen. Es unterliegt jedoch dem Progressionsvorbehalt (§ 32b EStG).
			Dadurch wird das Krankengeld zum versteuernden Einkommen hinzugerechnet. Der somit ermittelte höhere Steuersatz
			wird auf das zu versteuernde Einkommen angewandt. So vermeidet der Fiskus, dass Versicherte, die Krankengeld bezogen
			haben, einen geringeren Steuersatz haben als Versicherte, die kein Krankengeld bekommen haben.
		)
		}

		ControlGetPos, tX, tY,, tH, Button8, % Muster1a
		hAU 	:= GetHex(WinExist(Muster1a))
		If (GetDec(hAU) = 0)
			return

		AU        	:= GetWindowSpot(hAU)
		Fristen   	:= AlbisFristenRechner()
		FristX    	:= AU.BW, FristY := AU.BH, FristW := AU.CW - AU.BW*2, FristH := 20

		try {
			Gui, Frist: New, -Caption -DPIScale -AlwaysOnTop +HWNDhFrist Parent%hAU% -0x04020000 +E0x00000004
			Gui, Frist: Margin, 0, 0
			Gui, Frist: Color, % "c" Addendum.Default.BGColor
			Gui, Frist: Font, s8 q5 cWhite, % Addendum.Default.Font
			Gui, Frist: Add, Text    	, % "x4 y3 BackgroundTrans vTermine                    	"                                 	 , % "Krankengeldzahlung "
			GuiControlGet, t, Frist: Pos, Termine
			Gui, Frist: Font, s8 q5 Normal cWhite
			Gui, Frist: Add, Text		, % "x+0 w" FristW-tW-tX-40 " BackgroundTrans vStart		+HWNDhStart"     	, % Fristen.Anzeige
			Gui, Frist: Add, Picture	, % "x" FristW-18 " y2 +0x4000000  gAlbisFristenInfo hwndFristHInfo"        	 	, % "hBitmap: " hInfoIcon
			Gui, Frist: Show, % "x" FristX " y" FristY " w" FristW " h" FristH " NoActivate"                                           	, % Fristen.WinTitle
			WinSet, Redraw,, % "ahk_id " hFrist
			SetTimer, AlbisFristenUpdate, 100
		}

return

AlbisFristenUpdate: 	;{

	If !WinExist(Muster1a)	{
		Gui, Frist: Destroy
		If WinExist("ahk_id " hMyCal) || WinExist("erweiterter Kalender ahk_class AutohotkeyGui")
			Gui, newCal: Destroy
		SetTimer, AlbisFristenUpdate, Delete
		return
	}

	ControlGetText, AUSeit	, Edit1	, % Muster1a
	ControlGetText, AUBis	, Edit2	, % Muster1a
	If (AUSeitO = AUSeit) && (AUBisO = AUBis)
		return
	Fristen := AlbisFristenRechner((AUSeitO := AUSeit), (AUBisO := AUBis))
	ControlSetText	,, % Fristen.Anzeige, % "ahk_id " hStart
	WinSet          	, Redraw	,, % "ahk_id " hFrist
	WinSet          	, Top     	,, % "ahk_id " FristHInfo

return ;}

AlbisFristenInfo:     	;{
	MouseGetPos, mx, my,, mhWin, 2
	ToolTip, % info, % mx+20, % my, 14
	SetTimer, AlbisFristenInfoOff, -12000
return
AlbisFristenInfoOff:
	ToolTip,,,, 14
return ;}
}
;14
IfapVerordnung(Medikament, printsilent:=true) {                                                    	;-- Medikament über ifap verordnen und Dosiszettel (MS Word Dokument) automatisch ausdrucken

	; MedDocumentDir 	- muss in der Addendum.ini vermerkt sein - wird im Addendum.ahk - Skript eingelesen
	; ifapListe                	- wird später eine im AdditionalData-Pfad im JSON Format gespeicherte Datei mit Informationen zu den Medikamenten und den zu druckenden Dokumenten sein
	;								  die Daten werden beim ersten Aufruf dieser Funktion eingelesen, nur bei Änderung der Datei werden die Daten neu eingelesen
	;
	; ein Beispiel Code ist im Addendum.ahk Skript unter Hotstrings: #IfWinActive, Muster 16 ahk_class #32770 zu finden

	static RezeptWin        	:= "Muster 16 ahk_class #32770"
	static RezeptFelder     	:= "2,8,14,20,26,32"
	static ifapMainWin     	:= "ifap ahk_class TForm_ipcMain"
	static ifapHinweis       	:= "Hinweise zu ahk_class TFormVOHinweis"
	static ifapListe            	:= {"Tryasol 30 ml" : {"ProductName":"Tryasol Codein Forte TEI N2 30 ml", "PZN":"06304304", "Doc":"Tryasol.doc", "Hotstring":"Trya"}}

	DetectHiddenWinStatus	:= A_DetectHiddenWindows
	DetectHiddenTextStatus	:= A_DetectHiddenWindows

	Result                        	:= 0

	If !WinExist(RezeptWin)
		return Result

	DetectHiddenWindows, On
	DetectHiddenText, On

	; sucht nach dem passenden Eintrag im Objekt und nimmt dann die Automatisierung vor
	For MedKey, data in ifapListe
		If InStr(MedKey, Medikament)		{

			; prüft die Eingabeposition
				ControlGetFocus, cFocus, % RezeptWin
				RegExMatch(cFocus, "Edit(?<Nr>\d+)", Feld)
				If !RegExMatch(cFocus, "Edit") || (FeldNr not in 2,8,14,20,26,32)  {
						MsgBox, 1, Sie haben das Medikament nicht in die Rezeptfelder eingegeben.`nDie automatische Rezepterstellung wird jetzt abgebrochen!, 4
						DetectHiddenWindows, % DetectHiddenWinStatus
						DetectHiddenText, % DetectHiddenTextStatus
						return Result
				}

			; PZN senden
				SendRaw, % ifapListe[MedKey].PZN
				Sleep, 300

			; Arzneimitteldatenbank (ifap) aufrufen, Senden der F3-Taste
			; SendMessage (WM_Menu) wäre deutlich zuverlässiger, löst mitunter einen COM-Fehler aus nachdem im Anschluss ifap und Albis neugestartet werden müssen
				SendInput, {F3}
				WinWaitActive, ifap ahk_class TForm_ipcMain,, 4

			; entspricht dem Menupunkt 'Rezept/Markieren des aktuelles Medikament'
				SendMessage, 0x111, 89,,, % ifapMainWin	; entspricht dem Menupunkt 'Rezept/Markieren des aktuelles Medikament'
				Sleep, 500

			; Hinweisdialog abfangen und schliessen
				If WinExist("Hinweise zu ahk_class TFormVOHinweis")
					VerifiedClick("TiButton1", "Hinweise zu ahk_class TFormVOHinweis")

			; entspricht dem Menupunkt 'Datei/Rezeptübernahme'
				SendMessage, 0x111, 33,,, % ifapMainWin

			; wenn ein Dokument vorhanden ist, fragen ob es gedruckt werden soll, wenn ja dann wird es gedruckt
				If (StrLen(ifapListe[MedKey].Doc) > 0)	{

						If FileExist(Addendum.DosisDokumentenPfad "\" ifapListe[MedKey].Doc) {

								DocDruckFrage := "Sie haben eine Dosierungsanweisung für dieses Medikament hinterlegt.`nSoll diese gedruckt werden?"
								MsgBox, 0x40004, ifap Medikamentensuche, % DocDruckFrage
								IfMsgBox, Yes
								{
										oWord := ComObjCreate("Word.Application") ; create MS Word object
										oWord.Visible := !printsilent ; unecessary. handy for turning on for debug purposes
										oDoc := oWord.Documents.Open(Addendum.DosisDokumentenPfad "\" ifapListe[MedKey].Doc) ; create new document
										oWord.DisplayAlerts := 0  	; turns off alerts
										oDoc.PrintOut(0,,,,,,,1)      	; if it prints in the background, it does interfere with closing the application, especially if printing many copies
										oWord.DisplayAlerts := -1 	; turns them back on
										oDoc.Close                       	; close the file
										oWord.Quit                      	; close the application
								}

						}
						else {

							PraxTT("Es konnte das MS Word-Dokument:`n" ifapListe[MedKey].Doc "`nnicht gefunden werden!", "1 2")

						}
				}

			; Rezeptfenster wieder aktivieren
				WinActivate, % RezeptWin

			Result := 1
		}

	DetectHiddenWindows, % DetectHiddenWinStatus
	DetectHiddenText, % DetectHiddenTextStatus

return Result
}
;15
AlbisWeitereMedikamente(KatChar="#", CaveTag="Problem") {                             	;-- Daten aus dem Dialog "Auswahl weiterer Medikamente" erhalten

	/*  Beschreibung

		ich unterteile Medikamente in Gruppen für eine bessere Übersicht
		z.B. so :  	#####                  Blutdruck                  #####

		Allergien oder aufgetretene Nebenwirkungen werden so gekennzeichnet
						#####         Problemmedikamente        #####
						Metformin - KI Niereninsuffizienz
						KEINE NSAR WEGEN NIERENINSUFFIZIENZ
						#############################

		diese zusätzlichen Angaben beginnen mit einem '#' Zeichen. Die Funktion kann diese Informationen filtern.
		KatChar 	- ist für die Identifaktion der Kategorien und
		CaveTag	- ist ein zusätzlicher Filter für den Allergien oder Nebenwirkungsbereich


	 */

		static WMed := "Auswahl weiterer Medikamente ahk_class #32770"

		VOP          	:= Object()
		VOP.DMed	:= Array()
		VOP.WMed	:= Array()
		MedProblem := false

	; Daten des Fensters auslesen
		ControlGet, DMed	, List,, ListBox1       	, % WMed
		ControlGet, WMed, List,, SysListView321, % WMed

	; Daten parsen und in ein AHK-Objekt umwandeln
		DMed	:= StrSplit(DMed, "`n")
		Loop % DMed.MaxIndex() {                                            ; Dauermedikamente

			DMed[A_Index] := RegExReplace(Trim(DMed[A_Index]), "\*$")
			RegExMatch(Trim(DMed[A_Index]), "(?<Date>\d{2}\.\d{2}\.\d{4})\s+(?<ikament>.*[\)\[]*)|(?<Kategorie>" KatChar ".*)", Med)

			;SciTEOutput("MedDatum:  " ", Medikament:" Medkategorie ", Dosis:" Dosis ", Pos:" A_Index)

			If ((StrLen(MedKategorie) > 0) && (A_Index = 1)) || MedProblem {

				If RegExMatch(MedKategorie, "i)" CaveTag ) {
					MedProblem := true
					VOP.DMed.Push({"MedDatum":"", "Medikament":Medkategorie	, "Dosis":"" , "cave":true, "Pos":A_Index})
				} if RegExMatch(MedKategorie, "^" KatChar ) {
					MedProblem := false
					VOP.DMed.Push({"MedDatum":"", "Medikament":Medkategorie	, "Dosis":"" , "cave":true, "Pos":A_Index})
				} else
					VOP.DMed.Push({"MedDatum":"", "Medikament":DMed[A_Index]	, "Dosis":"" , "cave":true, "Pos":A_Index})

			}
			else if (StrLen(MedDate) > 0) {

				RegExMatch(Medikament, "(?<Name>.*)[\(\[]", Med)
				RegExMatch(Medikament, "\((?<Dosis>.*)\)", Med)
				RegExMatch(MedDosis, "(?<DosisF>.*)\-(?<DosisM>.*)\-(?<DosisA>.*)\-(?<DosisN>.*)", Med)
				;SciTEOutput("MedDatum:  " MedDate ", Medikament: " MedName ", Dosis: " MedDosisF "- " MedDosisM "-" MedDosisA "-" MedDosisN  ", Pos: " A_Index)
				Dosis := {"Morgens":MedDosisF, "Mittags":MedDosisM, "Abends":MedDosisA, "MedDosisN":MedDosisN}

				VOP.DMed.Push({"MedDatum":MedDate, "Medikament":MedName, "MedDosis":Dosis, "cave":false, "Pos":A_Index})

			}

		}

		WMed	:= StrSplit(WMed, "`n")
		Loop % WMed.MaxIndex() {                                            ; weitere Medikamente

			if RegExMatch(WMed[A_Index], "(?<Befreit>\w*)\t+(?<autidem>X*)\t+(?<Date>\d{2}\.\d{2}\.\d{4})\t+(?<RezeptArt>\w*)\s+(?<ikament>.*)\t+(?<Zusatz>.*)", Med) {

				RegExMatch(Medikament, "(?<Name>.*)\((?<DosisF>.*)\-(?<DosisM>.*)\-(?<DosisA>.*)\-(?<DosisN>.*)\)", Med)
				Dosis := {"Morgens":MedDosisF, "Mittags":MedDosisM, "Abends":MedDosisA, "MedDosisN":MedDosisN}

				VOP.WMed.Push({"befreit":MedBefreit, "autidem": autidem, "MedDatum":MedDate, "Rezeptart":MedRezeptArt, "Medikament":MedName, "MedDosis":Dosis, "Zusatz":MedZusatz})

			}

		}

return VOP
}
;16
AlbisVerordnungsplan() {

		VOP := AlbisWeitereMedikamente()
		ControlFocus, Listbox1, % "Auswahl weiterer Medikamente ahk_class #32770"
		;SciTEOutput(VOP.Count())
		;~ For index, med in VOP.DMed {

			;~ SendMessage, 0x187

				;~ PostMessage, 0x185, 1, -1, ListBox1, Fenstertitel  ; Wählt alle ListBox-Einträge aus. 0x185 ist LB_SETSEL.

		;~ }

		;Control, Choose, 1, Listbox1, % "Auswahl weiterer Medikamente ahk_class #32770"
		;VerifiedClick("Abbruch", WMed,,, 3)

}
;17
AlbisAusfuellhilfe(Formular, cmd) {                                                                          	;-- #### angefangen

	/* AlbisAusfuellhilfe - Beschreibung

			geplant als universell einsetzbare Funktion für schnelleres Ausfüllen von editierbaren Feldern in Albis

	*/

		static Formulare := {"Muster12" : {"from"	: [6,12,17,22,27,32,37,43,52,59,65,71,78,83]
														,	"to"   	: [7,13,18,23,28,33,38,44,53,60,66,72,79,84]}}
		static fcLast, awLast, wt, entryNr

		hk 	:= A_ThisHotkey
		aw	:= WinExist("A")
		fc  	:= GetFocusedControlClassNN(aw)
		sf		:= true                                                      	; flag bleibt true wenn der Eingabefokus noch im gleichen Steuerelement ist (sf = static focus)

	; Zurücksetzen von Variablen wenn ein anderes Formular bearbeitet wird
		If !(aw = awLast)
			wt	:= WinGetTitle(aw), awLast := aw, fcLast := ""

	; anderer Eingabefokus
		If !(fcLast = fc)
			sf := false, fcLast := fc

	; fokussiertes Steuerelement ist nicht in der Ausfuellhilfsliste (Formulare) dann zurück
		If !(ctrlLabel := ControlFilter(Formulare[Formular], fc))
			return

	switch Formular
	{
		case "Muster12":

			If (hk = "UP") {

			}

		; ermittelt alle schon vorhanden Datumsangaben in den "vom" und "bis" Feldern
			If !sf {

				Formulare["Muster12"].Dates := Array()
				Loop, % Formulare["Muster12"]["from"].Count() {

					from := Formulare["Muster12"]["from"](A_Index)
					to 	:= Formulare["Muster12"]["to"](A_Index)
					date := Array()
					ControlGetText, date1, % "Edit" from	, % "ahk_id " aw
					ControlGetText, date2, % "Edit" to   	, % "ahk_id " aw
					date[1] := date1, date[2] := date2

					Loop, 2
						If !(RegExMatch(date[A_Index], "\d{2}\.\d{2}\.\d{2,4}") || RegExMatch(date[A_Index], "\d{6,8}"))
							date[A_Index] := ""

					If (StrLen((Dates := date.1 "," date.2)) = 1)
						continue

					If !DatesInArr(Dates, Formulare["Muster12"].Dates)
						Formulare["Muster12"].Dates.Push(Dates)

					}
				}


			;SciTEOutput(Formulare["Muster12"]["fromDates"])
			;SciTEOutput(Formulare["Muster12"]["toDates"])

		}


}


;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; WARTEZIMMER                                                                                                                                                                           	(05)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisWZPatientEntfernen       	(02) AlbisWZOeffnen	                 	(03) AlbisWZKommentar             	(04) AlbisWZTabSelect
; (05) AlbisWZListe
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;1
AlbisWZPatientEntfernen(Nachname, Vorname) {                                                     	;--##### benötigt einen REWRITE ###### entfernt einen Patienten aus dem Wartezimmer

	;es wird der Patient entfernt der aktuell ausgewählt ist, noch kein direktes entfernen über den Patientennamen möglich!
	wzflag:=2

	If  !InStr(AlbisGetActiveWindowType(), "WZ") {
			MsgBox, 0x40004 , Frage, Soll ich den zuletzt behandelten Patienten`naus der Warteliste entfernen?
			IfMsgBox, Yes
			{
					SendInput, {DEL}
						sleep, 250
					SendInput, J
						sleep, 250
					wzflag:=1
			}
	}

return wzflag
}
;2
AlbisWZOeffnen(Maximize:=true) {                                                                          	;-- öffnet das Wartezimmer

	; https://docs.microsoft.com/en-us/windows/desktop/winmsg/wm-mdimaximize
	static WM_MDIMAXIMIZE:= 0x0225, WM_MDIGETACTIVE:= 0x0229				;wParam: A handle to the MDI child window to be maximized.

	If !(hWZ:= AlbisMDIWartezimmerID())	{
		PostMessage, 0x111, 32924,,, % "ahk_id " AlbisWinID()                                                      	;ruft das Wartezimmer auf
		sleep, 200
		while !(hWZ:= AlbisMDIWartezimmerID())	{
			If (A_Index > 11)
				return 0 ; kein Erfolg!
			sleep, 200
			PostMessage, 0x111, 32924,,, % "ahk_id " AlbisWinID()
		}
	}

	; Wartezimmer wird maximiert
		If Maximize
			PostMessage, % WM_MDIMAXIMIZE, % hWZ,,, % "ahk_id " AlbisMDIClientHandle()

return hWZ		;oder war schon geöffnet
}
;3
AlbisWZKommentar(WZ, Kommentar, Anwesenheit) {                                             	;-- automatisiert Wartezimmer - Kommentare eingeben/ändern

	; Beispiele:
	; ::^!Numpad1::AlbisWzKommentar("Labor", "Blutabnahme"	, "Anwesend")
	; ::^!Numpad2::AlbisWzKommentar("Arzt"	, "Sprechzimmer 1", "Anwesend")
	; ::^!Numpad3::AlbisWzKommentar("Arzt"	, "Pat. bitte um Rückruf", "Abwesend")

	; letzte Änderung: 29.03.2021

		wzkom := "Wartezimmer - Kommentar ahk_class #32770"

	; nur ausführbar bei angezeigter Karteikarte
		If !RegExMatch(AlbisGetActiveWindowType(), "i)Karteikarte")
			return 0

		hKarteikarte 	:= AlbisMDIChildGetActive()
		Patient       	:= AlbisCurrentPatient()

	; Wartezimmer - Kommentar aufrufen
		If !(hWin := Albismenu(32945, wzkom, 3)) {
			PraxTT(A_ScriptName ": Fenster [" RegExReplace(wzkom, "ahk_class.*") "]ließ sich nicht anzeigen!", "1 0")
			return 0
		}

		 Controls("", "reset", "")
		hAnwesenheit := Controls("", "ControlFind," Anwesenheit ", Button, return Hwnd", hWin)
		VerifiedChoose("Listbox1", hWin, WZ)                     	; Wartezimmer wählen
		 VerifiedSetText("Edit1", Kommentar, hWin)             	; Kommentar einsetzen
		  VerifiedCheck("", hAnwesenheit)                            	; Anwesenheit setzen
		    VerifiedClick("OK", hwin)

}
;4
AlbisWZTabSelect(TabName, hWZ:=0, activate:=false) {                                        	;-- aktiviert ein bestimmtes Wartezimmer

	; Funktion überprüft während der Ausführung das der richtige Tab aktiviert wurde
	; letzte Änderung 29.03.2021

	; handle des Wartezimmer ermitteln
		If !(hWZ := hWZ ? hWZ : AlbisWZOeffnen(false))
			return -1

	; derzeit aktiven Tab ermitteln
		If !(TabindexA 	:= AlbisMDITabActive(hWZ, "index"))
			return -2
		If !(TabNameA	:= AlbisMDITabActive(hWZ, "name"))
			return -3

	; gewünschter Tab = aktueller Tab?
		If InStr(TabNameA, TabName)
			return TabindexA + 1             	; +1 damit ich 0 bei einem Fehler zurückgeben kann

	; Tab wird nicht angezeigt, Tabindex suchen
		tabMatched := false
		If !tabMatched && !RegExMatch(TabName, "^\d+$")
			For tabindex, Tab in AlbisMDITabNames(hWZ)
				If InStr(Tab, TabName) {
					tabMatched := true
					break
				}
		If !tabMatched
			return -4

	; handle des SystabControls erhalten
		If !(hTab := AlbisMDITabHandle(hWZ))
			return -5

	; per Postmessage auf das gewünschte Tab umschalten
		PostMessage, 0x1330, % tabindex-1,,, % "ahk_id " hTab
		Sleep 100
		while !InStr(AlbisMDITabActive(hWZ, "name"), TabName) && (A_Index < 20) {
			If (A_Index > 1)
				Sleep 50
			PostMessage, 0x1330, % tabindex-1,,, % "ahk_id " hWZ
		}
		If !InStr(AlbisMDITabActive(hWZ, "name"), TabName)
			return -6

	; Wartezimmer aktivieren (nach vorne holen) bei Bedarf
		If activate
			If !IsObject(err := AlbisMDIChildActivate(hWZ))
				return err   ; handle des Wartezimmers

return tabindex-1
}
;5
AlbisWZListe(TabName, hWZ:=0, activate:=false) {                                                	;-- den Inhalt eines Wartezimmers auslesen

	; handle des Wartezimmer ermitteln
		If !(hWZ := hWZ ? hWZ : AlbisWZOeffnen(false))
			return -1

	; auf das Wartezimmer umschalten
		tabIndex := AlbisWZTabSelect(TabName, hWZ, activate)

	; handle der Listview und Auslesen des Inhaltes
		ControlGet, hLV, hwnd,, SysListview321, % "ahk_id " hWZ

	; Listview-Header auslesen
		SendMessage 0x101F, 0, 0,, % "ahk_id " hLV ; LVM_GETHEADER
		oheader := GetHeaderInfo(ErrorLevel)

	; Syslistview auslesen
		ControlGet, csv, List,,, % "ahk_id " hLV
		csv := RegExReplace(csv, "[\t]", "|")

	; Tabelle in ein key:value Objekt umwandeln
		wz := Array()
		For Lnr, line in StrSplit(csv, "`n") {
			wz.Push(Object())
			For Enr, element in StrSplit(line, "|") {
				key := oHeader[Enr].Text
				wz[Lnr][key] := element
			}
		}

return wz
}
;6
AlbisWZTabAktuell(TabName:="") {                                                                         	;-- gibt das aktuell angezeigte Wartezimmer zurück

	If !InStr(AlbisGetActiveWindowType(), "WZ")
		return 0

	hMdiChild	:= AlbisMDIChildGetActive()
	WZTitle 	:= AlbisMDIChildTitle(hMdiChild)

	If TabName
		return (AlbisMDITabActive(WZTitle, "Name") = TabName ? 1 : 0)

return AlbisMDITabActive(WZTitle, "Name")
}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; KARTEIKARTE/AKTE, ABRECHNUNG oder ALBIS-MENU                                                                                                                  	(18)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAkteSchliessen               	(02) AlbisAkteOeffnen                	(03) AlbisAkteGeoeffnet              	(04) AlbisPatientAuswaehlen
; (05) AlbisInBehandlungSetzen       	(06) AlbisSucheInAkte                  	(07) AlbisMenu                          	(08) AlbisKarteikarteZeigen
; (09) AlbisOeffnePatient                	(10) AlbisDateiAnzeigen             	(11) AlbisDateiSpeichern             	(12) AlbisKarteikarteAktiv
; (13) AlbisSaveAsPDF                       	(14) AlbisKarteikartenAnsicht        	(15) AlbisNeuerSchein                	(16) AlbisLeseDatumUndBezeichnung
; (17) AlbisAbrScheinCOVIDPrivat    	(18) AlbisAbrechnungAktiv
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisAkteSchliessen(CaseTitle="") {                                                             	;-- schließt eine Karteikarte ## seit Albis20.20 mit Problemen

	;Rückgabewerte 0 - die Funktion konnte keine Karteikarte schliessen, 1 - war nicht erfolgreich, 2 - es war keine Akte zum Schliessen geöffnet

	oACtrl			:= Object()			;active control object
	AlbisWinID	:= AlbisWinID()

	If InStr(AlbisGetActiveWindowType(), "Karteikarte")	{

		; aktuellen FensterTitel auslesen wenn keine Daten der Funktion übergeben wurden
			If (StrLen(CaseTitle) = 0)
				CaseTitle:= AlbisGetActiveWinTitle()


		; ein Eingabefocus in der Patientenakte muss erkannt und beendet werden (senden einer Escape- oder Tab-Tasteneingabe)
			while InStr(AlbisGetActiveWinTitle(), CaseTitle)				{
				aCtrl:= AlbisGetActiveControl("content")
				If ( (aCtrl.Identifier = "Dokument") && (StrLen(aCtrl.RichEdit) = 0) )	     	{					;leeres RichEdit - dann steht das Caret im Edit-Control (das ginge doch auch anders?!)
						AlbisKarteikarteAktivieren()
						SendInput, {Esc}
						sleep 100
				}
				else If ( (aCtrl.Identifier = "Dokument") && !(StrLen(aCtrl.RichEdit) = 0) ) 	{
						AlbisKarteikarteAktivieren()
						SendInput, {Tab}
						sleep 100
						SendInput, {Esc}
						sleep 100
				}

			; schliesst die Akte mit passendem Title
				err := AlbisMDIChildWindowClose(CaseTitle)
			}

		return err

	}

return 0
}
AlbisAkteSchliessen2(CaseTitle="") {                                                             	;-- neue Funktion

	;Rückgabewerte 0 - die Funktion konnte keine Karteikarte schliessen, 1 - war nicht erfolgreich, 2 - es war keine Akte zum Schliessen geöffnet

	; aktuellen FensterTitel auslesen wenn keine Daten der Funktion übergeben wurden
		If !CaseTitle
			CaseTitle := AlbisGetActiveWinTitle()
		else {
			If !AlbisMDIChildHandle(CaseTitle)            ; die Akte gibt es nicht
				return -1
		}

	; falls ein Steuerelement in der Karteikarte durch eine nicht beendete Eingabe blockiert ist
		If AlbisKarteiKarteEingabeStatus(CaseTitle) {

			SciTEOutput(A_ThisFunc ": Eingabefeld ist noch geöffnet" )

			AlbisMDIChildActivate(CaseTitle)                   	; Akte sichtbar machen
			AlbisKarteikarteZeigen(false)                      	; Karteikarte anzeigen
			hKK := AlbisKarteikarteAktivieren()               	; Karteikarte ativieren

			ControlSend,, {TAB}, % "ahk_id " hKK
			sleep 300
			ControlGetFocus, fctrl, ahk_class OptoAppClass
			If InStr(fctrl, "Edit") {
				SendInput {TAB}
				sleep 300
			}

			ControlSend,, {Escape}, % "ahk_id " hKK
			sleep 200
			ControlGetFocus, fctrl, ahk_class OptoAppClass
			If InStr(fctrl, "Edit") {
				SendInput {Escape}
				sleep 200
			}

		}
		else
			AlbisMDIChildActivate(CaseTitle)

		SciTEOutput(AlbisGetActiveWinTitle() " =>> "  InStr(AlbisGetActiveWinTitle(), CaseTitle) )
		If InStr(AlbisGetActiveWinTitle(), CaseTitle) {
			SciTEOutput(A_ThisFunc ": Karteikarte schliessen per Menu" )
			Albismenu(57602)
		}

		while (InStr(AlbisGetActiveWinTitle(), CaseTitle) || A_Index < 100)
			sleep 50


return InStr(AlbisGetActiveWinTitle(), CaseTitle) ? 0 : 1
}
;02
AlbisAkteOeffnen(CaseTitle, PatID="") {                                                      	;-- öffnet eine Patientenakte über Name, ID oder Geburtsdatum

	/*		BESCHREIBUNG

			AlbisAkteOeffnen - zum Öffnen einer Patientenakte nach Übergabe eines Namens, der Patienten ID oder eines Geburtsdatums

			Die Funktion ist in der Lage sämtliche Fenster abzufangen und entsprechend zu behandeln. Es wird bei erfolgreichem Aufruf der gewünschten Akte eine 1 ansonsten eine 0
			zurückgegeben. Falls sich ein Fenster mit einem Listview öffnet - Auswahl eines Namens, wenn mehr als einer vorhanden ist , wird eine 2 zurück gegeben. Nur das Listview
			Fenster wird nicht geschlossen. Da dieses ausgelesen werden muss, um in der Auswahl nach dem entsprechenden Namen suchen zu können.

			letzte Änderung 07.04.2021

	*/

	; Variablen                                                                                        	;{
		global Mdi
		static WarteFunc, AlbisTitleFirst
		static Win_PatientOeffnen := "Patient öffnen ahk_class #32770"

		If !AlbisWinID() {
			PraxTT(A_ThisFunc ": Albis wird nicht ausgeführt!", "2 1")
			return 0
		}

		name    	:= []
		CaseTitle	:= Trim(CaseTitle)
	;}

	; RegExString und anderen Suchstring erstellen                                     	;{
		If RegExMatch(Trim(PatID), "^\d+$") {

			If IsObject(oPat)
				PName := oPat[PatID].Nn ", " oPat[PatID].Vn, GD := oPat[PatID].Gd
			else if IsObject(PatDB)
				PName := PatDB[PatID].Name ", " PatDB[PatID].Vorname , GD := PatDB[PatID].Geburt

			rxStr      	:= "\[" Trim(PatID) "\s\/"
			sStr       	:= Trim(PatID)
			CaseTitle 	:= PName

		}
		else if (StrLen(CaseTitle) > 0) {
			rxStr	:= RegExReplace(CaseTitle, "\,\s+", ",\s+")
			sStr	:= RegExReplace(CaseTitle, "\,\s+", ", ")
		}
		else {
			PraxTT(A_ThisFunc ": es wurden keine Parameter übergeben`nDer Karteikartenaufruf wird abgebrochen.", "2 1")
			return 0
		}
	;}

	; Karteikarte bereits geöffnet?                                                             	;{
		hMdi	 := AlbisMDIClientHandle()
		oMDI := AlbisMDIClientWindows()
		For MDITitle, thisMdi in oMDI
			If RegExMatch(MDITitle, "i)" rxStr) {
				SendMessage, 0x222, % thisMdi.ID,,, % "ahk_id " hMdi
				PraxTT("Die gewünschte Karteikarte wird jetzt angezeigt.", "5 0")
				return 1
			}
	;}

	; Öffne Patient Dialog aufrufen                                                           	;{
		PraxTT("Geöffnet wird die Karteikarte des Patienten:`n#2" PName ", geb.am " ConvertDBASEDate(GD) " (" PatID ")`n`n(Warte bis zu 10s auf die Karteikarte)", "10 0")
		AlbisTitleFirst := WinGetTitle(AlbisWinID())               	; aktuellen Albisfenstertitel auslesen
		AlbisDialogOeffnePatient()                                     	; Aufruf des Fenster 'Patient öffnen'
	;}

	; Übergeben des Parameter an das Albisdialogfenster                         	;{
		If !VerifiedSetText("Edit1", sStr, Win_PatientOeffnen, 200) {
			PraxTT("Dialog: 'Patient öffnen' konnte nicht behandelt werden.", "2 1")
			VerifiedClick("Button3", Win_PatientOeffnen)
			return 0
		}

		while WinExist(Win_PatientOeffnen)		{

			If (A_Index > 1)
				sleep 100

			If (A_Index = 1)
				VerifiedClick("Button2", Win_PatientOeffnen)            	; Versuch 1: Button OK drücken
			else If (A_Index = 2)
				ControlSend, Edit1, {Enter}, % Win_PatientOeffnen	; Versuch 2: Enter simulieren
			else {
				PraxTT(	"Die Karteikarte des Patienten:`n#2" PName ", geb.am " GD "(" PatID ")`nkonnte nicht geöffnet werden", "2 1")
				WinClose, % Win_PatientOeffnen
				return 0
			}

		}

	;}

	; Loop der in den nächsten 10Sekunden auf die neue Karteikarte wartet	;{
		Loop	{

			If !(hwnd:= DLLCall("GetLastActivePopup", "uint", AlbisWinID()))
				hwnd := AlbisWinID()
			newTitle	:= WinGetTitle(hwnd)
			newClass	:= WinGetClass(hwnd)
			newText	:= WinGetText(hwnd)

		; Karteikarte ist geöffnet
			If RegExMatch(newTitle, "i)" rxStr) {
				PraxTT("", "Off")
				return 1
			}

		; Pat hat in diesem Quartal seine Chipkarte ...
			AlbisKeineChipkarte("Ja")

		; Dialog Patient <.....,......> nicht vorhanden
			If Instr(newTitle, "ALBIS") && Instr(newClass, "#32770") && Instr(NewText, "nicht vorhanden") {
				VerifiedClick("Button1", "ALBIS ahk_class #32770", "nicht vorhanden")			;Abbrechen
				PraxTT("Albis konnte keine Karteikarte finden!", "5 3")
				if WinExist(Win_PatientOeffnen)
					VerifiedClick("Button3", Win_PatientOeffnen)
				PraxTT("", "Off")
				return 0
			}
			else if Instr(newTitle, "Patient") && Instr(newClass, "#32770") && Instr(newText, "List1") {
				PraxTT("", "Off")
				return 2
			}

			If (A_Index > 50) 	{
				PraxTT(	"Die Karteikarte des Patienten:`n#2" PName ", geb.am " GD "(" PatID ")`nkonnte nicht geöffnet werden", "6 3")
				if WinExist(Win_PatientOeffnen)
					VerifiedClick("Button3", Win_PatientOeffnen)
				PraxTT("", "Off")
				return 0
			}

			sleep, 200
		}

		PraxTT("", "Off")
	;}

return 1
}
;----
AlbisWarteAufKarteikarte(AlbisTitleFirst, CaseTitle, waitingtime) {

		static Win_PatientOeffnen := "Patient öffnen ahk_class #32770"

		hwnd		:= WinExist("A")
		newTitle	:= WinGetTitle(hwnd)
		newClass	:= WinGetClass(hwnd)
		newText	:= WinGetText(hwnd)

	; Dialog Patient <.....,......> nicht vorhanden
		If Instr(newTitle, "ALBIS") && Instr(newClass, "#32770") && Instr(NewText, "nicht vorhanden") {
			VerifiedClick("Button1", "ALBIS ahk_class #32770", "nicht vorhanden")			;Abbrechen
			SetTimer, % WarteFunc, Off
			return 0
		}
		else if Instr(newTitle, "Patient") && Instr(newClass, "#32770") && Instr(newText, "List1") {
			SetTimer, % WarteFunc, Off
			return 2
		}

		If (AlbisTitleFirst = newTitle) && !InStr(newTitle, CaseTitle)
			return

		SetTimer, % WarteFunc, Off
	;Namen splitten - bei Schreibfehlern, Komma fehlt zwischen den Namen, kann er nicht erkennen das die Akte geöffnet ist
		If !RegExMatch(CaseTitle, "\d+")	{
			name := ExtractNamesFromString(CaseTitle)
			If !InStr(AlbisGetActiveWinTitle(), name[3]) && !InStr(AlbisGetActiveWinTitle(), name[4]) {
				MsgBox,, % "Addendum für AlbisOnWindows - " A_ScriptName, % "Die angeforderte Patientenakte ließ sich nicht öffnen!", 25
				return 0				;fungiert als ErrorLevel
			}
		}
		else	{
			If !InStr(AlbisGetActiveWinTitle(), CaseTitle)	{
				MsgBox,, Addendum für AlbisOnWindows - %A_ScriptName%, % "Achtung:`nDie angeforderte Patientenakte ließ sich nicht öffnen!", 25
				return 0				;fungiert als ErrorLevel
			}
		}

return 			;erfolgreich dann 1
}
;03
AlbisAkteGeoeffnet(Nachname, Vorname, GebDatum:="", PatID:="") {       	;-- wird eine Akte angezeigt

	;sucht nach den Namen und Geburtstag, Name und Vorname, Geburtstag können vertauscht sein, es wird nur geprüft ob die alle Bedingungen positiv sind
	;letzte Änderung 27.08.2018 - leere Parametererkennung hinzugefügt, Erkennungsfunktion gekürzte Fassung

	;Name, Vorname, Geburtstag können leer sein, doch nicht alle zusammen, das Suchergebnis wäre immer ein erfolgreich, deshalb
	If (Nachname = "") && (Vorname = "") && (GebDatum = "") && (PatID = "") {
		ExceptionHelper("include`\AddendumFunctions`.ahk", "AlbisAkteGeoeffnet", "Error at script line: " scriptline "`n" ScriptText "`n`n3 empty parameters are not allowed!", A_LineNumber)
		return 2
	}

	;Erkennungsfunktion
	AlbisTitle:= AlbisGetActiveWinTitle()
	If ( InStr(AlbisTitle, Nachname) && InStr(AlbisTitle, Vorname) && Instr(AlbisTitle, GebDatum) && InStr(AlbisTitle, PatID) )
		return 1            	;ist geöffnet
	else
		return 0            	;nicht geöffnet

return 2        	;wird die Funktion jemals hierher kommen?
}
;04
AlbisPatientAuswaehlen(Nachname, Vorname, GebDatum) {                       	;-- ~#fehlerhaft# Mehrfachauswahl eines Patienten
	win       	:= "Patient auswählen ahk_class #32770"
	hwnd     	:= WinExist("Patient auswählen ahk_class #32770", "List1")
return AlbisLVContent(hwnd, "SysListView321", "Name|Vorname|Geb.-Datum")
}
;05
AlbisInBehandlungSetzen() {                                                                      	;-- Pat. ins Wartezimmer in Behandlung

	;für die Zunkunft könnte hier die aktuelle Auswahl im Listview des Wartezimmerfenster interessant sein

	FControl:= GetFocusedControl()

	while !( hWZ:= GetHex(WinExist("Wartezimmer - Kommentar")) ) {
		SendInput, {F2}
		If hWZ:= GetHex(WinExist("Wartezimmer - Kommentar"))
			break
		sleep, 300
	}

	WinActivate, ahk_id %hWZ%
	RegExMatch(A_ComputerName, "\d+", Nr)                            	;das geht so einfach weil die Computer entweder Sp1 oder Sp2 heißen

	while WinExist("ahk_id " . hWZ)	{
		ControlGetText, WZKommentar, Edit1, ahk_id %hWZ%
		ControlSetText, Edit1, % RegExReplace(WZKommentar, "([0-9])", "") . Nr, ahk_id %hWZ%
		VerifiedClick("Button7"  , "Wartezimmer - Kommentar")
		VerifiedClick("Button11", "Wartezimmer - Kommentar")
	}

	while InStr(AlbisGetActiveWindowType(), "WZ") {
		SendInput, {Enter}
		Sleep, 250
	}

return AlbisCurrentPatient()
}
;06
AlbisSucheInAkte(inhalt, Richtung="up", Vorkommen="") {                          	;-- ~# funktioniert mäßig # suchen in der Akte, öffnet das Suchfenster

	; wenn Vorkommen = "" oder einfach frei gelassen wird, dann gibt die Funktion die Gesamtzahl des gesuchten Inhaltes zurück
	; Vorkommen = 1 und Aufwärtssuche (up) wie eine vorhergehende GVU Ziffer finden
	; Vorwärts und Aufwärtssuche ist mit Richtungsoption forward oder backward möglichch

	AlbisKarteikarteZeigen()

	;ok, die Karteikarte ist geöffnet, dann aktiviere ich sie jetzt und sie bekommt auch den Focus, danach kann sofort die Suchfunktion gestartet werden
	Loop {

		while !WinExist("Suchen","Suchen &nach:") {
			WinActivate    	, ahk_class OptoAppClass
			WinWaitActive	, ahk_class OptoAppClass,,4
			VerifiedClick("#327704", "ALBIS ahk_class OptoAppClass") 	;ein Klick in die Karteikarte damit entweder an den Anfang oder an das Ende gesprungen werden kann
			SendInput, {Up}
			If (Richtung="up")
               	SendInput, {End}
			else if (Richtung="down")
               	SendInput, {Home}
			sleep, 100
			SendInput, {Ctrl down}s{Ctrl up}
			sleep, 100
		}

		;jetzt werden die Parameter gesetzt
		If (Richtung="up")
			VerifiedCheck("Button4", "Suchen")
		else if (Richtung="down")
				VerifiedCheck("Button5", "Suchen")

		sleep, 100
		ControlSetText, Edit1, %inhalt%, "Suchen ahk_class #32770"				;jetzt noch den Suchtext einsetzen
		sleep, 100
		VerifiedClick("Button9", "Suchen ahk_class #32770")					;und jetzt geht es los mit dem Suche
		sleep, 100
		AlbisKarteikarteZeigen()
		sleep, 100

		If WinExist("ALBIS", "Der gesuchte Text") {
			WinClose, Albis, Der gesuchte Text
			WinWaitClose, Albis, Der gesuchte Text, 5
			If WinExist("Suchen ahk_class #32770") {
				WinClose, Suchen ahk_class #32770
				WinWaitClose, Suchen ahk_class #32770,, 5
			}
            break
		}

		localIndex := A_Index

	}	until localIndex = Vorkommen

return localIndex
}
;07
Albismenu(mcmd, FTitel:="", WZeit:=2, methode:=1) {                              	;-- Aufrufen eines Menupunktes oder Toolbarkommandos

	; letzte Änderung: 17.01.2021

	/* Albismenu() - Dokumentation

		mcmd:  		WM_Command - die Zahlen findet man im include Ordner AlbisMenu.json
		FTitel:			zu erwartender Fenstertitel, dient der Erfolgskontrolle des Menuaufrufes
							neu seit dem 06.05.2020 -
	                		FTitel kann ein Objekt sein mit zwei Fenstertiteln und Fenstertexten
							Parameterübergaben: Albismenu(11111, ["Fenstertitel ahk_class WinClass", "Fenstertext", "Alternativer Fenstertitel ahk_class WinClass2", "Alternativer Text"])

		WZeit:	    	Zeit wie lange auf das Fenster gewartet werden soll
		methode:  	Aufruf per Post- (1) oder Sendmessage (2)
							Methode 1 und 2 wirkt sich nur bei leerem FTitel aus, da nur SendMessage die Variable ErrorLevel setzt
	 */

		InfoMsg := false, WinText := ""

		If IsObject(FTitel) {
			WinTitle	:= FTitel.1
			WinText 	:= FTitel.2
			AltTitle  	:= FTitel.3
			AltText   	:= FTitel.4
		}
		else
			WinTitle	:= FTitel

	; Menuaufruf ohne auf ein Fenster zu warten
		If (StrLen(WinTitle) = 0) {
			If (methode = 1) {
				PostMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
				return 1
			}
			else if (methode = 2) { ; ACHTUNG: Sendmessage wartet auf eine Antwort, diese kann lange dauern (bis 5s zum Teil)
				SendMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
				return ErrorLevel
			}
		}

	; Menuaufruf. Wartet auf bis zu 2 verschiedene Fenster
		Loop 	{

			; Prüfen ob erwartete Fenster schon geöffnet sind
				If !IsObject(FTitel) {
					If (pHwnd := WinExist(WinTitle, WinText))
						return GetHex(pHwnd)
				} else {
					If (pHwnd := WinExist(WinTitle, WinText))
						return GetHex(pHwnd)
					else if (pHwnd := WinExist(AltTitle, AltText))
						return {"hwnd": GetHex(pHwnd)}
				}

			; der Menu command Befehl öffnet manchmal nicht das gewollte Fenster (Bug von Albis?)
				AlbisIsBlocked(AlbisWinID())

			; Menupunkt aufrufen
				If (methode = 1)
					PostMessage	, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
				else if (methode = 2)
					SendMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"

			; Fenster abwarten
				while (A_Index * 50 <= WZeit * 1000) {

					If !IsObject(FTitel) {
						If (pHwnd := WinExist(WinTitle, WinText))
							return GetHex(pHwnd)
					} else {
						If (pHwnd := WinExist(WinTitle, WinText))
							return GetHex(pHwnd)
						else if (pHwnd := WinExist(AltTitle, AltText))
							return {"hwnd": GetHex(pHwnd)}
					}

					sleep, 50

				}

				If (A_Index > 2) { ; Abbruch nach dem 2.Versuch
					AlbisIsBlocked(AlbisWinID())
					return 0
				}

		}

return 0
}
;---
AlbisInvoke(menu, start=1) {                                                                     	;-- ~# fehlerhalt ## Vereinfachung eines Menupunktaufrufes

	; ~# fehlerhalt ## Vereinfachung eines Menupunktaufrufes - Teil- oder Ganzstring Übergabe wie im Albis Menu selbst

	; diese Funktion benötigt die Datei AlbisMenu.json aus dem AddendumDir Verzeichnis \include, die Variable AddendumDir muss Superglobal im aufrufenden Skript definiert sein
	; übergebe einen Menu-Namen z.b. AlbisInvoke("Wartezimmer",1) erhälst du die CommandID zurück und der Menupunkt wird bei start=1
	; durch PostMessage aufgerufen, start=0 gibt nur die CommandID zurück
	; folgende Syntaxe sind möglich: 	1. vollständiger Pfad "Formular/BG/F1050  - Ärztliche Unfallmeldung	(A13)" , hier gehen auch Teilstrings "Formular/BG/F1050" oder "Form/BG/A13"
	;                            						2.	Teile der Suche "Formular, F1050" oder "BG, A13" - dies dauert länger da die Funktionen mitunter jedes Key:Value Paar durchsuchen muss

	static oAlbisMenu

	If !IsObject(oAlbisMenu) {

		AlbismenuPath := Addendum.Dir "\include\Daten\AlbisMenu.json"
		If !FileExist(AlbismenuPath) {
			throw A_ThisFunc ": AlbisMenu.json - konnte nicht gefunden werden."
		}

		oAlbisMenu := JSON.Load(FileOpen(Addendum.Dir "\include\Daten\AlbisMenu.json", "r").Read())

	}

	For key, val in oAlbisMenu	{
		If ( Instr(key, menu) && (Start = 1) ) {
			PostMessage, 0x111, % val,,, % "ahk_id " AlbisWinID()
			return
		} else {
			return val
		}
	}

}
;08
AlbisKarteikarteZeigen(Info=true) {                                                               	;-- schaltet zur Karteikartenansicht

	; !!WICHTIG!!: ein Erfolg der wird per Rückgabe einer 1 bestätigt, die aufrufende Funktion sollte den Rückgabewert prüfen bevor andere Funktionen fortgesetzt werden!
	; letzte Änderung: 05.04.2021


	; eine Karteikarte wird schon angezeigt
		If RegExMatch(AlbisGetActiveWindowType(), "i)Karteikarte\s*$")
			return 1

	; es wird gar keine Patientenakte angezeigt - dann ist die Funktionsausführung nicht möglich
		If !InStr(AlbisGetActiveWindowType(), "Patientenakte")
			return 0

	; Info anzeigen
		If Info
			PraxTT("Karteikarte anzeigen", "1 2")

	; Albis aktivieren
		AlbisActivate(1)

	; wechselt zur Karteikarte, muss dazu begonnene Eingaben beenden oder abbrechen.
	; es werden hier zwei Dialogfenster abgefangen und entsprechend behandelt
		Controls("", "Reset", "")
		while !RegExMatch(AlbisGetActiveWindowType(), "i)Karteikarte\s*$")	{

			If (A_Index > 1)
				sleep 50
			PostMessage, 0x111, 33033,,, % "ahk_id " AlbisWinID()
			sleep, 50

			Loop {

				If (hAlbisHinweis2 := WinExist("ALBIS ahk_class #32770", "Gebühren"))
					VerifiedClick("Button1", hAlbisHinweis2)

				while, (hAlbisHinweis1 := WinExist("ALBIS ahk_class #32770", "Die aktuelle Zeile wurde nicht gespeichert")) {

					WinSetTitle, % "ahk_id " hAlbisHinweis1,, % "ALBIS [ Zeileneingabe wird automatisch geschlossen in " Round(5-(0.1*(A_Index-1)),1) "s]"
					Sleep, 50
					If (A_Index > 30)	{
						VerifiedClick("Button1", "ALBIS ahk_class #32770", "Die aktuelle Zeile wurde nicht gespeichert")
						cFocus := Controls("", "GetFocus", (hActiveMDIWin:= AlbisMDIChildGetActive()))
						If (StrLen(cFocus) > 0) {
							SendInput, {Esc}
						} else {
							ControlFocus,, % "ahk_id " Controls("#327702", "ID", hActiveMDIWin)
							SendInput, {Esc}
						}
						break
					}

				}

				If (A_Index > 10)
					return 0

				AlbisActivate(1)
				Controls("", "Reset", "")
				cFocus := Controls("", "GetFocus", (hActiveMDIWin:= AlbisMDIChildGetActive()))
				If (StrLen(cFocus) > 0)
					SendInput, {Esc}

			}

			If InStr(AlbisGetActiveWindowType(), "Karteikarte")
				return 1
			else if (A_Index > 10)
				return 0

		}

return 1
}
;09
AlbisDialogOeffnePatient(command:="invoke", pattern:="" ) {                       	;-- startet Dialog zum Öffnen einer Patientenakte

	; more commands are here:
	; 	abort/close - um das Fenster zu schliessen ohne eine Suche zu durchzuführen
	;	serach/set/open [Namens-/Suchmuster]- übernimmt den eingegeben Text und kann gleichzeitig das Suchmuster eintragen

		static Win_PatientOeffnen := "Patient öffnen ahk_class #32770"

	; Aufrufen des Dialogfensters
		If !InStr(command, "close")
			id:= Albismenu(32768, Win_PatientOeffnen, 3)

	; command parsen
		If InStr(command, "invoke")
			return ID
		else If InStr(command, "close") {

			If WinExist(Win_PatientOeffnen)
				return VerifiedClick("Button3", Win_PatientOeffnen)
			else
				return

		} else If InStr(command, "open") || InStr(command, "set")  {

			; kein Suchmuster übergeben dann wird nur auf Ok gedrückt
				If (StrLen(Pattern) > 0)
					If !VerifiedSetText("Edit1", Pattern, Win_PatientOeffnen, 200)
						return 0

			; Suchmuster wurde als Parameter übergeben, aber
				If InStr(command, "set")
					return ID

				; Akte wird jetzt geöffnet durch drücken von OK
					while WinExist(Win_PatientOeffnen)					{
							; Button OK drücken
								VerifiedClick("Button2", Win_PatientOeffnen)
								WinWaitClose, % Win_PatientOeffnen,, 1
							; Fenster ist immer noch da? Dann sende ein ENTER.
								if WinExist(Win_PatientOeffnen)								{
									WinActivate, % Win_PatientOeffnen
									ControlFocus, Edit1, % Win_PatientOeffnen
									SendInput, {Enter}
								}

								If (A_Index > 10)
										return
								sleep, 100
					}

		}


return id
}
;10
AlbisDateiAnzeigen(FullFilePath) {                                                              	;-- öffnet eine Datei zur Ansicht in Albis

		AlbisActivate(1)
		SplitPath, FullFilePath, filepath, filename

	; Dateiformat prüfen (die erste Zeile einer in Albis anzeigbaren Textdatei beginnt mit "\P")
		FileReadLine, fileline, % FullFilePath, 1
		If !InStr(fileline, "\P")		{
			MsgBox, 0, Addendum für Albis on Windows, % "Die übergebene Datei kann mit Albis nicht angezeigt werden."
			return 1
		}

	; Aufruf des Menupunktes Patient/Datei anzeigen...
		If !hWin := Albismenu("33030", "Öffnen ahk_class #32770")		{
			MsgBox, 0, Addendum für Albis on Windows, % "Ein Fehler ist bei der Automatisierung aufgetreten!`n"
			. "`nDer Aufruf des Dialoges ''Datei anzeigen'' war nicht erfolgreich.`nDie Funktion wird jetzt beendet!"
			return 1
		}

	; schreibt zuerst den Dateipfad in das Fenster, sendet dann Enter und schreibt den Dateinamen in selbiges Feld
		VerifiedSetFocus("Edit1"	, hWin)
		  VerifiedSetText("Edit1"	, FullFilePath, hWin, 200)
		VerifiedSetFocus("Edit1"	, hWin)
		ControlSend, Edit1, {Enter}, % "ahk_id " hWin
		If ErrorLevel		{
			MsgBox, 0, Addendum für Albis on Windows, % "Ein Fehler ist bei der Automatisierung aufgetreten!`nDie aufgerufene Funktion wird jetzt beendet."
			return 1
		}

return 0
}
;11
AlbisDateiSpeichern(FullFilePath, overwrite:= false) {                                     	;-- speichert Auswertungen, Protokolle, Statistiken oder Listen

	; prüft ob der Dateiname schon existiert, setzt dem Dateinamen eine Indizierung hinzu und speichert so unter einem anderen Namen
		If !overwrite && FileExist(FullFilePath)		{
			SplitPath, FullFilePath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
			Loop				{
				fname:= OutDir "\" OutNameNoExt "(" A_Index ")." OutExtension
				If !FileExist(fname)
					break
			}

			FullFilePath:= fname
		}

	; Tastatur- und Mauseingriffe durch den Nutzer kurzzeitig sperren
		hk(1, 0, "!ACHTUNG!`n`nDie Tastatureingaben sind für maximal 10 Sekunden gesperrt", 10)

	; Tagesprotokoll speichern und die Tagesprotokollausgabe schließen
		If !hSaveAsWin := Albismenu(33014, "Speichern unter ahk_class #32770")		{
			MsgBox, 0, Addendum für Albis on Windows, Erwarteter 'Speichern unter' - Dialog konnte nicht abgefangen werden!`nDie Funktion wird abgebrochen
			return 0
		}

	; Steuerelemente befüllen
		Controls("Edit1", "SetText, " FullFilePath, hSaveAsWin)

	; Speichern unter drücken
		Controls("Button2", "Click use Controlclick", hSaveAsWin)
		sleep 200

	; Warten auf einen eventuellen "Speichern unter bestätigen" - Dialog
		If overwrite		{
			PraxTT("Warte 5 Sekunden auf einen möglichen Speichern unter... Dialog!", "0 5")
			WinWait, % "Speichern unter bestätigen ahk_class #32770",, 5
			PraxTT("", "Off")
			If !ErrorLevel
				Controls("Button1", "Click use ControlClick", "Speichern unter bestätigen ahk_class #32770")
		}

	; Tastatur- und Mauseingriffe wieder entsperren
		hk(0, 0, "Die Tastaturfunktionen sind wieder hergestellt!", 2)

return FullFilePath	; Rückgabe des vollständigen Pfades
}
;12
AlbisKarteikarteAktiv() {                                                                               	;-- gibt true zurück wenn eine Karteikarte angezeigt wird

	If InStr(AlbisGetActiveWindowType(), "Patientenakte")
		return true

return false
}
;13
AlbisSaveAsPDF(filePath, printer="", Laborblattdruck=false) {                         	;-- Ausdruck/Speichern per PDF-Druckertreiber

	; ermöglicht den Export von Listen, Protokollen, dem Laborblatt über die Verwendung eines PDF-Druckertreibers
	; ACHTUNG: 	die Funktion kann auch mit dem gesonderten Druckdialog des Laborblatt Drucks umgehen,
	;                    	dazu muss der Dialog "Laborblatt Druck" allerdings zwingend vorher aufgerufen worden sein.
	;                    	Die Funktion Albismenu() kann zwar auf 2 unterschiedliche Fenstertitel reagieren,
	;						allerdings kann man keine Einstellungen im "Laborblatt Druck" Fenster vornehmen

		printer 	:= !printer 	? "Microsoft Print to PDF"            	: printer
		filepath	:= !filepath 	? A_Temp "\AlbisSaveAsPDF.pdf" 	: filepath

	; Albis-Laborblatt Druckdialog anzeigen falls noch nicht geschehen
		If Laborblattdruck {

			If !(hLaborblattdruck := WinExist("Laborblatt Druck ahk_class #32770"))
				If !(hLaborblattdruck := Albismenu(57607, "Laborblatt Druck ahk_class #32770", 3, 1))
					return 2

			while (!WinExist("Drucken ahk_class #32770") && A_Index <= 5) {

				VerifiedClick("Drucker...", "Laborblatt Druck ahk_class #32770")
				while (!WinExist("Drucken ahk_class #32770") && A_Index <= 20)
					sleep 20

			}

		}

	; handle des Drucker-Dialogs
		If !(hprintdialog := WinExist("Drucken ahk_class #32770"))
			return 3

	; PDF-Drucker wählen
		err := VerifiedChoose("ComboBox1", hprintdialog, printer)
		If (err <> 1) {
			PraxTT("PDF Drucker konnte nicht gefunden werden.`nFehlercode: " err, "2 1")
			sleep 2000
			return 4
		}

	; Debug
		;~ If ((res := Weitermachen("Func: " A_ThisFunc "`Laborblattdruck: " Laborblattdruck "`nprinter: " printer "`nhprd: " hprintdialog)) <> 1)
			;~ return res

		;~ sleep 8000
		If !(res := VerifiedClick("OK", hprintdialog))
			return 7

	; Laborblatt Druck aktiv dann ist noch ein OK notwendig
		If Laborblattdruck
			If !VerifiedClick("OK", hLaborblattdruck)
				return 8

	; Microsoft Print to PDF Dialog bearbeiten
		Druckausgabe := "Druckausgabe speichern unter ahk_class #32770"
		WinWait, % Druckausgabe,, 5
		hDruckausgabe := WinExist(Druckausgabe)
		res := VerifiedSetText("Edit1", filePath, hDruckausgabe)
		sleep 200
		res := VerifiedClick("Button2", hDruckausgabe)
		sleep 200
		WinWaitClose, % Druckausgabe,, 5

	; Fenster: "Albis", Text: "Drucken" erscheint bei größeren Dokumenten
		WinWait        	, % "Albis ahk_class #32770", "Drucken", 5
		WinWaitClose	, % "Albis ahk_class #32770", "Drucken", 20

return ErrorLevel ? 8 : 1
}
;14
AlbisKarteikartenAnsicht(PatientFensterCB) {                                                	;-- Ansicht Karteikarte, Abrechnun, Laborblatt

	; verändert aktuelle Auswahl der Combobox in der PatientFensterToolbar (Karteikarte,Abrechnung,Laborblatt...)

	ControlGet, hTbar	, Hwnd	,, ToolbarWindow321, % "ahk_id " AlbisMDIChildGetActive()
	ControlGet, hCB  	, Hwnd	,, ComboBox1			, % "ahk_id " hTbar
	If !hCB
		return 0

	; String oder Nummer wird unterschieden
	;SciTEOutput(hTbar)
return VerifiedChoose("ComboBox1", hTbar, PatientFensterCB)
}
;15
AlbisNeuerSchein(Zeigen=true) {                                                               	;-- behandelt den Dialog "Neuen Schein für  aufnehmen"

	; Zeigen schließt Fenster bei Übergabe von "OK" oder "Abbruch"

		PopupWin := ["Wollen Sie wirklich einen neuen Schein anlegen", "Sie haben für diesen Patienten in diesem Quartal"]
		AlbisBlanko := "ALBIS ahk_class #32770"
		neuerSchein:= "Neuen Schein für ahk_class #32770"

	; Zeigen = false , Dialog wird geschlossen
		hwnd := WinExist(neuerSchein)
		If !RegExMatch(Zeigen, "1")
			If hwnd {
				VerifiedClick(Zeigen, neuerSchein,,, true)
				WinWaitClose, % neuerSchein,, 3
				return ErrorLevel
			}

	; Dialog ist bereits geöffnet
		If (hwnd := WinExist(neuerSchein))
			return hwnd

		AlbisCloseLastActivePopups(AlbisWinID())

	; Menubefehl Patient/Schein/Neu
		PostMessage	, 0x111, 32788,,, % "ahk_class OptoAppClass"

	; Fenster abwarten
		while (A_Index <= 60) {

			For pwinNr, poptext in PopupWin
				If (WinExist(AlbisBlanko, poptext))
					If !VerifiedClick("Ja", AlbisBlanko, poptext,, true)
						Sleep 100

			if (hwnd := WinExist(neuerSchein))
				return GetHex(hwnd)

			Sleep 50
		}

return 0
}
;16
AlbisLeseDatumUndBezeichnung(MouseX, MouseY) {                                   	;-- Mausposition abhängige Karteikarten Informationen

	while (StrLen(KKDatum) = 0) {
		MouseClick, Left, MouseX, MouseY                                                                     	; ein Mausklick setzt den Cursor
		Sleep, 200
		Content 	:= AlbisGetActiveControl("Content")
		RichText  	:= Content.RichEdit
		KKDatum	:= Content.Edit2
		If (StrLen(KKDatum) > 0)                                                                                      	; Datum konnte ausgelesen werden, dann Abbruch
			break
		If (A_Index > 4)
			return ""
	}
	SendInput, {Escape}                                                                                               	; Karteikartenzeile freigeben
	PdfTitel  	:= RegExReplace(PdfTitel, "\.*pdf$", "")                                                    	; eventuell noch vorhandene pdf Endung aus dem Karteikartentext entfernen

return {"Text": RichText, "Datum":KKDatum, "Content":Content}
}
;17
AlbisAbrScheinCOVIDPrivat(Quartal:="") {                                                    	;-- Abrechnungsschein für Private bei COVID-19 Impfungen

		/* Beschreibung


				nur für die Impforgie bei uns Hausärzten. Legt in einem Durchlauf einen Abrechnungsschein für Privatpatienten an
				, damit die Impfdosis über die KV abgerechnet werden kann

			-- techn. Notizen

				Ersatzverfahren / Manuelle Eingabe der Versichertendaten
				- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Edit1 		(VKNR) 				38825
				Edit3 		(IK-Nummer) 	100038825
				Edit5 		(Kasse)          	BAS
				Button6 	M						Mitglied
				Edit9     	Gültig von:		kann leer bleiben
				Edit10     	bis:               	kann leer bleiben
				Button29	OK

				Rechnung für <....,....> aufnehmen
				- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Button3		&Abrechnungsschein
				Button17	E&rsatzverfahren
				Edit12     	Gültig von:    	kann leer bleiben
				Edit13     	bis:                	kann leer bleiben

	 */

		dht 	:= A_DetectHiddenText
		dhw	:= A_DetectHiddenWindows
		DetectHiddenText      	, Off
		DetectHiddenWindows	, Off
		BlockInput, On

		AlbisActivate(1)

	; Abbruch wenn Quartal kein Objekt ist (Addendum_Datum.ahk/QuartalTage())
		If !IsObject(Quartal)
			throw A_ThisFunc ": Qartal muss ein Objekt sein (Rückgabeparameter von QuartalTage() / Addendum_Datum.ahk)"

	; Abbruch wenn keine Karteikarte oder die eines gesetzl. Versicherten geöffnet ist
		If !AlbisKarteikarteAktiv() || (AlbisVersArt() <> 1) {
			DetectHiddenText      	, % dht
			DetectHiddenWindows	, % dhw
			return 0
		}
	; prüft ob Coronaabrechnungsschein schon angelegt wurde
		If AlbisAbrechnungsscheinVorhanden(Quartal.Quartal "/" SubStr(Quartal.Jahr, -1)) {
			DetectHiddenText      	, % dht
			DetectHiddenWindows	, % dhw
			return 1
		}

	; jetzt blockierende Fenster entfernen
		AlbisIsBlocked()

	; nochmal prüfen ob Fenster geöffnet sind
		If WinExist("Ersatzverfahren ahk_class #32770")
			VerifiedClick("Abbruch", "Ersatzverfahren ahk_class #32770")
		If WinExist("Rechnung für ahk_class #32770")
			VerifiedClick("Abbruch", "Rechnung für ahk_class #32770")

	; Strg+Shift gedrückt halten bis das Fenster: Rechnung für <...,...> erscheint
		BlockInput, On
 		SendInput, {LControl Down}{LShift Down}

	; sendet Menubefehl für "Neuer Schein"
		PostMessage, 0x111, 32788,,, ahk_class OptoAppClass

	; mögliche blockierende Dialogfenster schließen
		while !(hneuerSchein:=WinExist("Rechnung für ahk_class #32770")) {
			if WinExist("ALBIS ahk_class #32770", "existiert noch eine nicht ausgedruckte Privatrechnung")
				VerifiedClick("Ja", "ALBIS ahk_class #32770", "existiert noch eine nicht ausgedruckte Privatrechnung")
			else If WinExist("ALBIS ahk_class #32770", "Es ist ein alternativer Rechnungs")
				VerifiedClick("Nein", "ALBIS ahk_class #32770", "Es ist ein alternativer Rechnungs")
		}

	; gedrückte Tasten freigeben
		SendInput, {LControl Up}{LShift Up}
		BlockInput, Off

	; Eintragungen in "Rechnung für <...,...>" vornehmen
		Controls("", "Reset", "")
		res	:= VerifiedClick("Abrechnungsschein",,, hneuerSchein)
		;~ VerifiedSetText("Edit12"	, Quartal.TDBeginn  	, hneuerSchein)		; Gueltig von
		;~ VerifiedSetText("Edit13"	, Quartal.TDEnde  	, hneuerSchein)		; bis
		hErsatz 	:= Controls("", "ControlFind, Ersatzverfahren, Button, return hwnd", hneuerSchein)
		res2    	:= VerifiedClick("",,, hErsatz)

	; Ersatzverfahren Dialog abfangen
		WinWait, Manuelle Eingabe der Versichertendaten ahk_class #32770,, 6
		If !(hEingabe := WinExist("Ersatzverfahren / Manuelle Eingabe ahk_class #32770")) {
			DetectHiddenText      	, % dht
			DetectHiddenWindows	, % dhw
			return 0
		}

	; Schein im Ersatzverfahren anlegen
		VerifiedCheck("eGK",,, hEingabe, false)				        			; eGK aus
		VerifiedSetText("Edit1"	, "38825"               	, hEingabe)		; VKNR
		VerifiedSetText("Edit3"	, "100038825"        	, hEingabe)		; IK-Nummer
		VerifiedSetText("Edit5"	, "BAS"                    	, hEingabe)		; Kasse
		VerifiedSetText("Edit9"	, Quartal.TDBeginn  	, hEingabe)		; Gueltig von
		VerifiedSetText("Edit10"	, Quartal.TDEnde  	, hEingabe)		; bis
		VerifiedClick("Button6",,, hEingabe, false)                             	; M (itglied)

	; Ersatzverfahrendialog schliessen
		 VerifiedClick("OK",,, hEingabe)                                           	; OK

	; Rechnung für ... schließen und Schein ist angelegt.
		VerifiedClick("OK", "Rechnung für < ahk_class #32770")
		AlbisKarteikarteZeigen()

	; Fenstereinstellungen zurücksetzen
		DetectHiddenText      	, % dht
		DetectHiddenWindows	, % dhw

return 1
}
;18
AlbisAbrechnungAktiv(EditFocus:=false) {                                                    	;-- wird eine Kassenabrechnung angezeigt

	; nur bei aktivem Albisfenster und angezeigter Abrechnung
	; mit wählbarer Option (EditFocus) für aktiven o. inaktiven Eingabefokus
	;
	; Parameter:
	; EditFocus - darf ein Editelement in der Abrechnungsanzeige sichtbar sein

	If !WinActive("ahk_class OptoAppClass")
		return false

	;~ DHWMode := A_DetectHiddenWindows
	;~ DetectHiddenWindows, Off

	If !EditFocus {
		hMDIChild 	:= AlbisMDIChildGetActive()
		hSysTV      	:= Controls("SysTreeView321", "HWND", hMDIChild	, false, false)
								 Controls("", "Reset", "")
		hEdit          	:= Controls("Edit"              	, "HWND", hSysTV 	 	, false, false)
		ControlGet,  cStyle, Style,,, % "ahk_id " hEdit
		EditIsVisible := cStyle & 	0x10000000                ; WS_VISIBLE
	}

	;~ DetectHiddenWindows, % DHWMode

	If InStr(AlbisGetActiveWindowType(), "Abrechnung")
		If !EditFocus && !EditIsVisible
			return true
		else if EditFocus
			return true

return false
}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; ANDERE DIALOGE                                                                                                                                                                      	(2+4)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAusindisKorrigieren
; (02) class PatientenDaten
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
AlbisAusIndisKorrigieren(IndisToRemove:="", IndisToHave:="", Chroniker:=true, GB:=true) {  ; Ausnahmeindikationen: bestehende EBM-Ziffern tauschen oder hinzufügen

	; aktueller Patient
		currPatName := AlbisCurrentPatient()
		If !(PatID := AlbisAktuellePatID())
			return

	; Variablen
		If !IndisToRemove || !IsObject(IndisToHave)
			return 0

	; Ausnahmeindikationen auslesen
		ausIndisAlt := ausIndis := PatientenDaten.weitereInfos("GetText Ausnahmeindikation")
		;SciTEOutput("ausIndis bei [" currPatName "]`nA: " ausIndis)

	; Indikationen entfernen
		ausIndis := RegExReplace(ausIndis, IndisToRemove)
		ausIndis := RegExReplace(ausIndis, "\-{2,}", "-")
		ausIndis := RegExReplace(ausIndis, "\-\s*$", "")

	; Chroniker und GB Ziffer-Indikationen hinzufügen
		If Chroniker && IstChronischKrank(PatID)
			If !RegExMatch(ausIndis, "03220")
				ausIndis := "03220-" ausIndis
		If GB && IstGeriatrischerPatient(PatID) {
			If !RegExMatch(ausIndis, "03362")
				ausIndis := "03362-" ausIndis
			If !RegExMatch(ausIndis, "03360")
				ausIndis := "03360-" ausIndis
		}

	; Quartalsziffern hinzufügen
		For idx, ebm in IndisToHave
			If !RegExMatch(ausIndis, ebm "(\-|\s*$)")
				ausIndis := ebm "-" ausIndis
		ausIndis := RegExReplace(ausIndis, "\-{2,}", "-")
		ausIndis := RegExReplace(ausIndis, "\-\s*$", "")

	; Liste sortieren ([N]umerisch [U]Doppelte aussortieren [D]Trennzeichen -)
		Sort, ausIndis, N U D-

	; Dialoge schliessen
		If (ausIndisAlt <> ausIndis)
			PatientenDaten.weitereInfos("SetText Ausnahmeindikation", ausIndis)
		If (PatientenDaten.weitereInfos("Close", "OK") = 1)
			res := PatientenDaten.CloseDialog("Personalien")

	; Ausgabe
		If (ausIndisAlt <> ausIndis) {
			FileAppend, % "[" currPatName "] (" ausIndisAlt ") (" ausIndis ")`n", % Addendum.DBPath "\sonstiges\Ausnahmeindikationen.txt"  , UTF-8
			return {"alt":ausIndisAlt, "neu":ausIndis}
		}

return  {"alt":ausIndisAlt, "neu":""}
}

class PatientenDaten {               ; Menu Patient/Stammdaten - Dialoge aufrufen und bearbeiten

	static stammdaten := {"Abrechnungsassistent": {"cmd":34811, "WinTitle":"Abrechnungsassistent"}
								, 	  "Personalien"           	: {"cmd":32774, "WinTitle":"Daten von"}
								,	  "Dauerdiagnosen"    	: {"cmd":32776, "WinTitle":"Dauerdiagnosen von"}
								,	  "Dauermedikamente" 	: {"cmd":32777, "WinTitle":"Dauermedikamente von"}
								,	  "Cave"                     	: {"cmd":32778, "WinTitle":"Cave! von"}
								,	  "Kontrolltermine"       	: {"cmd":32775, "WinTitle":"Kontrolltermine"}
								,	  "Krankengeschichte"    	: {"cmd":32830, "WinTitle":"Krankengeschichte"}
								,	  "Patienteneinwilligung"	: {"cmd":35247, "WinTitle":"Patienteneinwilligung"}
								,	  "Patientenbild"          	: {"cmd":33317, "WinTitle":"Patientenbild"}
								,	  "Patientengruppen"     	: {"cmd":34362, "WinTitle":"Patientengruppen"}
								,	  "Familie"                   	: {"cmd":34705, "WinTitle":"Familie"}
								,	  "Therapiesitzungen"   	: {"cmd":34199, "WinTitle":"Therapiesitzungen"}
								,	  "Antikoagulantien"    	: {"cmd":34869, "WinTitle":"Antikoagulantion-Pass von"}}

	ShowDialog(dialog)                         	{

		If !AlbisIsBlocked(2)
			return Albismenu(this.stammdaten[dialog].cmd, this.stammdaten[dialog].WinTitle " ahk_class #32770")

	return 0
	}

	CloseDialog(dialog, Options:="OK") 	{
		If !(hwnd := WinExist(this.stammdaten[dialog].WinTitle " ahk_class #32770"))
			return 1
		WinActivate, % "ahk_id " hwnd
		If !VerifiedClick(Options, hwnd)
			If !VerifiedClick(Options, hwnd)
				return -1
	return 1
	}

	Personalien(cmd)                            	{

		Switch cmd {

			; Dialog aufrufen
				case "Show":

					this.Dialog 	:= "Personalien"
					this.hDialog 	:= this.ShowDialog(this.dialog)

				return this.hDialog

			; Subdialog weitere Informationen aufrufen
				case "weitere Informationen":

					this.Dialog := "Personalien"
					If !(this.hDialog := WinExist(this.stammdaten.personalien.WinTitle " ahk_class #32770"))
						If !(this.hDialog := this.ShowDialog(this.Dialog))
							return -2
					Controls("","Reset","")
					If !(hwnd := Controls("", "ControlClick, weitere Informationen, Button", this.hDialog))
						return -3
					WinWait, % "ahk_class #32770", % "Adresse des", 4
					this.hwInfos := WinExist("ahk_class #32770", "Adresse des")

				return this.hwInfos

		}

	}

	weitereInfos(cmd, options:="")        	{

		; direkter Aufruf mit Auslesen der Ausnahmeindikation: PatientenDaten.weitereInfos("GetText Ausnahmeindikationen")
		; alle anderen Funktionen in der Klasse werden bei Bedarf nacheinander ausgeführt

		If !(hAdresse := WinExist("ahk_class #32770", "Adresse des"))
			If !(hAdresse := this.Personalien("weitere Informationen"))
				return -4

		Switch cmd {

			; Text aus dem Feld Ausnahmeindikation auslesen
				case "GetText Ausnahmeindikation":
					If !(hAusnahmeindikation := ControlGet("hwnd",, "Edit14", "ahk_id " hAdresse))
						return -5
					ControlGetText	, Ausnahmeindikation	,			, % "ahk_id " hAusnahmeindikation
				return Trim(Ausnahmeindikation)

			; Text in das Feld Ausnahmeindikation schreiben
				case "SetText Ausnahmeindikation":
					If !(hAusnahmeindikation := ControlGet("hwnd",, "Edit14", "ahk_id " hAdresse))
						return -6
					If !VerifiedSetText("Edit14", Options, hAdresse)
						return -7
				return 1

			; Dialogfenster "weitere Informationen" schliessen
				case "Close":
					If !VerifiedClick(Options, hAdresse)
						return -8
					WinWaitClose, % "ahk_class #32770", % "Adresse des", 1
					while WinExist("ahk_class #32770", "Adresse des") && (A_Index < 800) {
						If (hwnd:= WinExist("ahk_class #32770", "ist keine Ausnahmeindikation"))
							VerifiedClick("Button2", hwnd)
						Sleep 20
					}
				return WinExist("ahk_class #32770", "Adresse des") ? -9 : 1

		}

	}

}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; LABOR ABRUF AUTOMATION / LABOR DIVERSES                                                                                                                           	(11)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) vianovaInfoBoxWebClient      	(02) AlbisLaborWaehlen            	(03) AlbisLaborbuchOeffnen       	(04) AlbisLaborImport
; (05) AlbisGNRAnforderungChilds  	(06) AlbisLaborblattDrucken          	(07) AlbisLaborblattExport           	(08) AlbisLaborblattZeigen
; (09) AlbisLaborAuswählen             	(10) AlbisLaborDaten                    	(11) AlbisLaborAlleUebertragen
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
vianovaInfoBoxWebClient(infoBoxID=0) {                                                               	;-- WebClient Automation (Labor IMD)

	; letzte Änderung vom 17.06.2021
	; 17.06.2021: SetTimer Aufruf fehlerhaft
	; 30.01.2021: flexiblere Steuerelementerkennung
	; Funktion benötigt Addendum.ahk

	;##Hersteller des WebClienten haben die Gui des WebClienten geändert, Button haben andere ClassNN bekommen
		global fcRLStat
		static MCS_WinClass  	:= "ahk_class WindowsForms10.Window.8.app.0.378734a"  ; WindowsForms10.Window.8.app.0.34f5582_r6_ad1
		If !infoBoxID
			infoBoxID	:= WinExist("vianova infoBox ahk_class WindowsForms10.Window.8.app")                              	;übergebenes Hookhandle ist nicht immer das Fensterhandle

	; Web-Client wurde abgefangen, Verarbeitungsflags werden gesetzt                                                                	;{
		Addendum.Labor.AbrufDaten:= false
		Addendum.Labor.AbrufVoll  	:= false

		MsgBox, 0x1024, Addendum für Albis on Windows,                                                                                	;{
									(LTrim
									Laborabruf erkannt!

									Möchten Sie alle weiteren Schritte bis hin zum
									Übertragen der Labordaten in die Patienten-
									akten automatisch durchführen lassen?
									)
		IfMsgBox, No
			return 0

		;}

		Addendum.Labor.AbrufVoll := true

	;}

	; Button "Abruf" drücken                                                                                                                                   	;{
		ctrlClass := Controls("", "ControlFind, Abruf, Button, return classNN", infoBoxID)
		If (StrLen(ctrlClass) > 0) {
			If !VerifiedClick(ctrlClass, infoBoxID) {
				Addendum.Labor.AbrufDaten:= false
				Addendum.Labor.AbrufVoll	:= false
				Addendum.Labor.AbrufStatus	:= "fail"
				PraxTT("Das Auslösen des Abrufes (simulierter Mausklick auf Abruf)`nhat nicht funktioniert.`nDie Funktion bricht hier ab", "0 10")
				return 0
			}
		}
		else {
			Addendum.Labor.AbrufDaten:= false
			Addendum.Labor.AbrufStatus	:= "fail"
			Addendum.Labor.AbrufVoll	:= false
			PraxTT("Das Auslösen des Abrufes (klick auf Abruf)`nhat nicht funktioniert.`nDie Funktion bricht hier ab", "0 10")
			return 0
		}

	;}

	; wartet bis keine Daten mehr hinzukommen                                                                                                   	;{
		Loop {

			oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4" , 0, "ahk_id " infoBoxID)
			t := ""
			Loop % oAcc.accChildCount()	{
				 oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4." A_Index, 0, "ahk_id " infoBoxID)
				 t .= oAcc.accValue(2) "`n"
				 If Instr(t, "Keine Daten zum Abruf") {
					Addendum.Labor.AbrufDaten:= false
					break
				} else if Instr(t, "Es liegen neue Labordaten vor") {
					Addendum.Labor.AbrufDaten:= true
					break
				}
			}
			If (A_Index > 250) || Addendum.Labor.AbrufDaten
				break
			sleep 100

		}
	;}

	; keine Labordaten dann                                                                                                                                	;{
		If !Addendum.Labor.AbrufDaten {

			msg :=" - es liegen keine neuen Labordaten vor`n`nDer Laborabruf wird nicht fortgesetzt!"
			Addendum.Labor.AbrufStatus:= "no_data"
			Addendum.Labor.AbrufDaten:= false
			Addendum.Labor.AbrufVoll	:= false

		}
		else {

			NoCopy 	:= 0
			msg      	:= "- neue Labordaten vorhanden"
			Addendum.Labor.AbrufStatus := "data_received"
			PraxTT(msg, "20 2")

		; sucht nach den LDT Dateinamen (es können mehrere sein)
			oT := StrSplit(t, "`n")
			Loop % oT.MaxIndex() 	{
				if RegExMatch(oT[A_index], "\|\s*Sichere\s*\<(?<File>.*?)\>\s*ins\s*Archiv\|", LDT) {	; Sichere <X01492.LDT> ins Archiv
					msg .= "`n- Sichere LDT-Datei: " LDTFile
					PraxTT(msg, "20 2")
					NoCopy += AlbisLaborLDTBackup(LDTFile)
				}
			}

			If NoCopy
				msg .= "`n- die " (NoCopy > 1 ? "LDT-Dateien (" NoCopy ") konnten ":" LDT-Datei konnte ") "nicht gesichert werden"
			else
				msg .= "`n- ein Backup der erhaltenen Labordatendatei" (NoCopy > 1 ? "en":"") " wurde erstellt!"

		}

		PraxTT(msg, "20 2")
	;}

	; Button "Schliessen" drücken                                                                                                                          	;{
		ctrlClass := Controls("", "ControlFind, Schließen, Button, return classNN", infoBoxID)
		If (StrLen(ctrlClass) > 0)
			VerifiedClick(ctrlClass , infoBoxID)
		else
			WinClose, % "ahk_id " infoBoxID
	;}

	; Schließen des Fensters abwarten, wenn das Fenster nicht schließt über andere Befehle das Fenster schliessen  	;{
		WinWaitClose, % MCS_WinClass,, 6
		If ErrorLevel 	{

			; Prozess ID ermitteln
				Process, Exist, infoBoxWebClient.exe
				iBoxWebPID := ErrorLevel

			; mittels PID zunächst ein normales WinClose und falls nicht erfolgreich ein schließen des Prozeß versuchen
				WinClose, % "ahk_pid " iBoxWebPID
				WinWaitClose, % "ahk_pid " iBoxWebPID,, 1
				If ErrorLevel {
					SendMessage 0x112, 0xF060,,, % "ahk_pid " iBoxWebPID 			; WMSysCommand + SC_Close
					WinWaitClose, % "ahk_pid " iBoxWebPID,, 1
				If ErrorLevel {
					SendMessage 0x10, 0,,, % "ahk_pid " iBoxWebPID 	            		; WM_Close
					WinWaitClose, % "ahk_pid " iBoxWebPID,, 1
				If ErrorLevel {
					SendMessage 0x2, 0,,, % "ahk_pid " iBoxWebPID 	             		; WM_Destroy
					WinWaitClose, % "ahk_pid " iBoxWebPID,, 1
				If ErrorLevel
					Process, Close, % "ahk_pid " iBoxWebPID
					EL := ErrorLevel
				}
				}
				}
		}
	;}

	; mit Aufruf des Dialogfensters 'Daten holen' wird der automatisierte Vorgang fortgesetzt
		PraxTT("`n- die Anwendung 'vianovaWebClient " (EL ? "konnte nicht beendet werden" : "wurde beendet"), "10 2")

		If (GetDec(Addendum.Labor.InvokedHWND  := AlbisLaborwaehlen()) = 0)
			ResetLaborabrufStatus()

		Addendum.Labor.AbrufStatus := "invoke_dialog_LaborDatenHolen"
		Addendum.Labor.Reset      	:= fcRLStat := Func("ResetLaborabrufStatus")
		SetTimer, % fcRLStat, -120000   ; nach 120 Sekunden wird der Laborstatus und alle zugehörigen Variablen zurückgesetzt

return 1
}
ResetLaborabrufStatus() {                                                                                       	;-- Variablen des Laborabrufes zurücksetzen

		global fcRLStat

	; Timerausführung beenden, falls die Funktion nicht per Timeraufruf gestartet wurde
	; falls doch sollte isFunc das Löschen eines nicht mehr vorhanden Timer verhindern (Fehlermeldung sonst)
		If IsFunc(fcRLStat)
			SetTimer, % fcRLStat, Delete

		Addendum.Labor.AbrufStatus	:= ""
		Addendum.Labor.AbrufDaten:= false
		Addendum.Labor.AbrufVoll  	:= false

return
}
AlbisLaborLDTBackup(LDTDateiName) {                                                                 	;-- empfangene Labordaten sichern

	; -------------------------------------------------------------------------------------------------------------
	;                 !!! EINSTELLUNG DES RICHTIGEN ENCODINGS FÜR DIE LDT-DATEI !!!
	; -------------------------------------------------------------------------------------------------------------
	;
	;	nach Spezifikation durch die KVB ist ein bestimmtes Zeichenkodierungsformat zu verwenden.
	; 	Vorgabe ist: ISO-8859-15 dies entspricht unter Windows Latin 9 (Westeuropa)
	;	der CodePage Identifier ist -- 28605 -- / siehe URL-Link:
	; 	https://docs.microsoft.com/de-de/windows/win32/intl/code-page-identifiers
	;	weiteres auf den Seiten des - Qualitätsring Medizinische Software e. V.
	;	https://www.qms-standards.de/mitgliederbereich/arbeitsgruppe-ldt/
	;	in Notepad++ lässt sich dies am besten mit North European/OEM 865 darstellen
	; -------------------------------------------------------------------------------------------------------------
	;
	;	ihr Stammlaborverzeichnis (LDT Verzeichnis) - hinterlegen Sie in der Addendum.ini!
	;	[Labor_abrufen]
	; 	LDTDirectory = C:\Labor

	FileEncoding, CP28605
	FileCopy, % Addendum.LDTDirectory "\" LDTDateiName, % Addendum.DBPath "\Labordaten\" A_YYYY "-" A_MM "_" A_DD "-" A_Hour "_" A_Min "-Ldt.txt"
	err := ErrorLevel
	FileEncoding, UTF-8

return err
}
;02
AlbisLaborWaehlen() {                                                                                            	;-- öffnet das Fenster Daten holen
; Menupunkt 32965 - Extern\Labor\Daten importieren ....
return Albismenu(32965, "Labor auswählen ahk_class #32770")
}
;03
AlbisLaborbuchOeffnen(hinweis=true) {                                                                   	;-- oeffnet das Laborbuch

	; ermittelt alle 100ms den Fenstertitel von Albis
	; letzte Änderung 20.02.2021

	; ist das Laborbuch schon geöffnet wird es aktiviert
		If (hLabbuch := AlbisMDIChildHandle("Laborbuch")) {
			AlbisMDIChildActivate("Laborbuch")
			return hLabbuch
		}

	; blockierende Fenster schliessen
		AlbisIsBlocked(AlbisWinID())
		If hinweis
			PraxTT("Befehl:`n- Laborbuch anzeigen -`nwird ausgeführt", "6 1")

	; Öffnen des Laborbuches per WM_Command Befehl - Laborbuch:= 34162
		PostMessage, 0x111, 34162,,, % "ahk_class OptoAppClass"
		while !InStr(AlbisGetActiveWinTitle(), "Laborbuch") {
			If (A_Index > 99)
				break
			else if (A_Index > 1)
				sleep 50
		}

	; Hinweis
		If !(hLabbuch := AlbisMDIChildHandle("Laborbuch")) {
			If hinweis
				PraxTT("Das Laborbuch konnte nicht angezeigt werden.", "2 1")
			return 0
		}

return hLabbuch
}
;04
AlbisLaborGNRHandler(nr, hwnd) {                                                                          	;-- automatisiert 'GNR der Anford.-Ident übernehmen'

	; letzte Änderung: 13.03.2021

		static WinGNRAnforderung := "GNR der Anford ahk_class #32770"
		static rxStatics1 := {"AfNr" : "^\d+$", "BArt" : "^[A-Z][\sa-z]+$", "Pat" : "^(?<Name>.*,.*)\s+\((?<ID>\d+)\)$"}
		static rxStatics2 := {"ETag":"^\d{2}\.\d{2}\.\d{4}$", "ATag":"^\d{2}\.\d{2}\.\d{4}$"}

		Statics       	:= Object()
		mcount        	:= 0

	; Daten für die Identifizierung auslesen
		Controls("", "Reset", "")
		sleep 500                                                                                                                          	; Zeichnen des Fensters abwarten
		Statics.AfNr 	:= Controls("Static2"  	, "GetText", WinGNRAnforderung)									; AnforderungsNr identifiziert das Fenster
		Statics.BArt  	:= Controls("Static8"  	, "GetText", WinGNRAnforderung)									; BefundArt
		Statics.ETag 	:= Controls("Static4"  	, "GetText", WinGNRAnforderung)									; Eingangs-Datum:
		Statics.ATag	:= Controls("Static11"	, "GetText", WinGNRAnforderung)									; Abnahme-Datum:
		Statics.Pat    	:= Controls("Static6"  	, "GetText", WinGNRAnforderung)									; Patient:

	; per RegExMatch die Werte prüfen, Veränderungen der ClassNN der Steuerelemente können so erkannt werden
	; Übereinstimmungen müssen der Anzahl der Steuerelemente entsprechen
		For staticArt, rxStr in rxStatics1 {
			If RegExMatch(Statics[staticArt], rxStr, Pat)
				mcount ++
			else
				failedClassNN .= staticArt " (" StaticText "),"
		}
		If (mcount <> rxStatics1.Count()) {
			Addendum.Labor.AbrufStatus := "Failure: mismatching classNNs"
			FileAppend	, % datestamp() "|" A_ThisFunc "()`t " Addendum.Labor.AbrufStatus "  (" RTrim(failedClassNN, ",") ")`n"
								, % Addendum.DBPath "\Labordaten\LaborimportLog.txt"
			return 0
		}	else {
			PraxTT("Anford.-Nr: " Statics.AfNr, "1 1")
			FileAppend	, % datestamp() "|" A_ThisFunc "()`t [" Statics.AfNr "] " Statics.BArt ", " Statics.ETag ", " Statics.ATag ", " Statics.Pat "`n"
								, % Addendum.DBPath "\Labordaten\LaborimportLog.txt"
		}

	; ist der Dialog neu dann die Daten im Protokoll sichern
		If (Addendum.Labor.AfNrOld <> Statics.AfNr) {
			Addendum.Labor.AbrufStatus := "Success: window data read out are correct"
			Addendum.Labor.AfNrOld := Statics.AfNr
		}	else {
			PraxTT("#2Labor importieren`n`nDas vorhergehende Anfordungsfenster ist noch geöffnet!`nDie Importfunktion wird abgebrochen.", "3 1")
			Addendum.Labor.AbrufStatus := "Failure: last dialog window is not closed!"
			FileAppend	, % datestamp() "|" A_ThisFunc "()`t  " Addendum.Labor.AbrufStatus "`n"
						   		, % Addendum.DBPath "\Labordaten\LaborimportLog.txt"
			return 0
		}

	; alle GNR übernehmen, der Dialog schließt dann
		hGNRA	:= GetHex(WinExist(WinGNRAnforderung))
		If !VerifiedClick("Alle GNR", WinGNRAnforderung,,, 1)                                                     	; und das Schließen des Dialoges abwarten
			If (hwnd := WinExist(WinGNRAnforderung))
				If InStr(WinGetText(Hwnd), Statics.AfNr) {
					PraxTT("Dialog 'GNR der Anforderung' konnte nicht geschlossen werden.", "1 1")
					Addendum.Labor.AbrufStatus := "Failure: action on button ''Alle GNR übernehmen'' failed!"
					FileAppend	, % datestamp() "|" A_ThisFunc "()`t " Addendum.Labor.AbrufStatus "`n"
										, % Addendum.DBPath "\Labordaten\LaborimportLog.txt"
					return 0
				}

	; abhängige Dialoge während der Übertragung behandeln (das bisher vohandene)
		Addendum.Labor.AbrufStatus := "Wait: popup window"
		Loop {

			If (A_Index > 30)
				break
			else if (A_Index > 1)
				sleep 100

			If (hGNRA	= GetHex(WinExist(WinGNRAnforderung))) {
				childs := AlbisLaborGNRChilds(hGNRA)
				If (childs.Count() > 0)
					For idx, win in childs {
						If InStr(win.text, "Soll der unbekannte Parameter") {
							VerifiedClick("Ja", "ALBIS ahk_class #32770", "Soll der unbekannte Parameter")
						} else If InStr(win.text, "Labordaten liegen als Vor-") {
							VerifiedClick("Nein", "ALBIS ahk_class #32770", "Labordaten liegen als")
						} else If InStr(win.text, "GO-Nr. werden nicht übertragen") {
							VerifiedClick("OK", "ALBIS ahk_class #32770", "GO-Nr. werden nicht übertragen")
						} else if InStr(win.title, "Bitte warten") {
							sleep 100
						} else if !InStr(win.title, "GNR der Anford") {
							Addendum.Labor.AbrufStatus := "Failure: unknown dialog [" win.Title ", " SubStr(RegExReplace(win.Text, "[\n\r]", "|"), 1, 50) "]"
							FileAppend	, % datestamp() "|" A_ThisFunc "()`t " Addendum.Labor.AbrufStatus "`n"
												, % Addendum.DBPath "\Labordaten\LaborimportLog.txt"
							return 0
						}
					}
			}

		}

	; 'alle übertragen' per wm_command aufrufen
	; Albis überträgt nie alle Labordaten bei mehr als einem Datensatz aufeinmal. Deshalb wird nach jeder erfolgreichen
	; Übertragung nochmals der Befehl 'alle übertragen' gesendet.
		Addendum.Labor.AbrufStatus := ""
		If !WinExist(WinGNRAnforderung)
			PostMessage, 0x111, 34157,,, % "ahk_id " AlbisMDIChildHandle("Laborbuch")
		else If (hGNRA <> GetHex(WinExist(WinGNRAnforderung)))
			AlbisLaborGNRHandler(1, WinExist(WinGNRAnforderung))

return 1
}
;05
AlbisLaborGNRChilds(hGNRA) {                                                                            	;-- alle Childfenster v. 'GNR der Anforderung..' erhalten

	; gehört unmittelbar zu AlbisRestrainLabWindows()
	; die Funktion soll alle abhängigen Dialogfenster beim Übertragen von Labordaten aus dem Laborbuch ins Laborblatt erkennen
	; Die Schwierigkeit bestand bis dahin in der Erkennung und Behandlung bisher unbekannter Dialogfenster. Dies führte zu fehlerhaften
	; RPA Prozessen mit teilweise schleifenartigem Verhalten ohne Möglichkeit eines Abbruchs.

	childs 	:= []
	hwnds	:= StrSplit(FindChildWindow({"ID":hGNRA},{"class":"#32770", "exe":"albis"}, "off"), ";")
	For idx, hwnd in hwnds
		If (hwnd <> hGNRA)
			childs.Push({"Title": WinGetTitle(hwnd), "Text":WinGetText(hwnd), "Hwnd": hwnd})

return childs
}
;06
AlbisLaborblattDrucken(Spalten, Drucker="", SaveDir="", PrintAnnotation=1) {        	;-- Automatisiertes Drucken eines Laborbefundes (## AlbisLaborblattExport() ist besser!)

	; im Moment kann nur die Anzahl der auszudruckenden Spalten als Option übergeben werden
	; Drucker := "pdf" für PDF Drucker - einzustellen in den Addendum.ini
	; Drucker := "Standard" für den eingestellten Standard Drucker - zu hinterlegen in der Addendum.ini

		static oDrucker:= Object()
		static printInit, Fenster_Druckeinrichtung, Fenster_Laborblatt_Druck

	; Erstinitalisierung
		If !printInit		{
			IniRead, PDF       	, % AddendumDir "\Addendum.ini", % CompName, Drucker_pdf
			IniRead, Standard	, % AddendumDir "\Addendum.ini", % CompName, Drucker_Standard
			Fenster_Druckeinrichtung	:= "Druckeinrichtung ahk_class #32770"
			Fenster_Laborblatt_Druck	:= "Laborblatt Druck ahk_class #32770"
			Fenster_Drucken             	:= "Drucken ahk_class #32770"
			Fenster_Druckausgabe   	:= "Druckausgabe speichern unter"
			oDrucker.PDF                	:= PDF
			oDrucker.Standard        	:= Standard
			printInit                         	:= true
		}

	; Kontext prüfen
		If !InStr(AlbisGetActiveWindowType(), "Patientenakte")		{
			PraxTT("Der Laborblattdruck kann nur bei einer geöffneten`nKarteikarte ausgeführt werden!.", "6 3")
			return 0
		}

	; bei Druckausgabe in eine PDF-Datei einen neuen Dateinamen finden
		If InStr(Drucker, "pdf")	{
			Name:= StrSplit(AlbisCurrentPatient(), ",", A_Space)
			; noch nicht vorhandenen Dateinamen finden
				Loop		{
					filename:= "BW-" SubStr(Name[1], 1, 2) SubStr(Name[2], 1, 2) ( A_Index = 1 ? "" : "-" A_Index - 1 ) "_Ausdruck_vom_" A_DD "." A_MM "." A_YYYY
					If !FileExist(SaveDir "\" filename . ".pdf")
						break
				}
		}

	; Auslesen des eingestellten Druckers und Zwischenspeichern der Einstellungen für späteres Wiederherstellen
		Albismenu(57606, Fenster_Druckeinrichtung, 3, 1)
		oDrucker.ActivePrinter:= { "Name"     	: ControlGet("Choice" 	,, "ComboBox1"	, Fenster_Druckeinrichtung)
		                                		, "Size"        	: ControlGet("Choice" 	,, "ComboBox2"	, Fenster_Druckeinrichtung)
			                                 	, "Source"   	: ControlGet("Choice" 	,, "ComboBox3"	, Fenster_Druckeinrichtung)
				                                , "Portrait"    	: ControlGet("Checked"	,, "Button5"     	, Fenster_Druckeinrichtung)
					                           , "Landscape" : ControlGet("Checked"	,, "Button6"     	, Fenster_Druckeinrichtung)}

	; Einrichten des gewünschten Drucker
		If InStr(Drucker, "pdf") || (Drucker = "")
				Dokumentdrucker:= oDrucker.PDF
		else if InStr(Drucker, "Standard")
				Dokumentdrucker:= oDrucker.Standard

		Control, ChooseString, % Dokumentdrucker, ComboBox1, % Fenster_Druckeinrichtung
		sleep, 200

		VerifiedCheck("Button5", Fenster_Druckeinrichtung)       	;Checkbox	: Hochformat
		  VerifiedClick("Button7", Fenster_Druckeinrichtung)       	;Button      	: Ok

	; Laborblatt öffnen und Druckvorgang ausführen
		If AlbisLaborblattZeigen() {

			; Druckaufruf ausführen
				If !(hwnd := Albismenu(57607, Fenster_Laborblatt_Druck, 3, 1)) {
					PraxTT("Der Dialog 'Laborblattdruck' konnte nicht aufgerufen werden.`nDie Funktion wird jetzt abgebrochen!", "6 3")
					return 0
				}
				;sleep 500

			; Überprüft nochmals den ausgewählten Drucker
				VerifiedClick("Button11", Fenster_Laborblatt_Druck) 	                    	; Drucker...
				WinWaitActive, % Fenster_Drucken,, 10
				sleep, 500
				If !InStr(ControlGet("Choice",, "ComboBox1", Fenster_Drucken), Dokumentdrucker)
					Control, ChooseString, % Dokumentdrucker, ComboBox1, % Fenster_Drucken
				sleep, 300
				If !VerifiedClick("Button11", Fenster_Drucken) {                                	; OK - Button im Fenster 'Drucken'
					PraxTT("Der Druckereinstellungsdialog hat sich nicht schließen lassen.`nDie Funktion wird jetzt abgebrochen!", "6 3")
					return
				}

			; Druck vorbereiten
				VerifiedCheck("Button2", Fenster_Laborblatt_Druck)                            	; letzte  ..... Spalten
				VerifiedSetText("Edit1", Spalten, Fenster_Laborblatt_Druck, 200)          	; Spaltenanzahl
				VerifiedCheck("Button5", Fenster_Laborblatt_Druck,,, PrintAnnotation)	; Anmerkungen und Probedaten
				sleep, 500

			; Drucken
				VerifiedClick("Button9", Fenster_Laborblatt_Druck)                              	; OK - Button

			; Druckausgabe speichern unter Fenster abfangen
				If InStr(Drucker, "pdf") {
						WinWait , % Fenster_DruckAusgabe,, 10
						If ErrorLevel		{
							PraxTT("Das Speichern-Dialog für die Ausgabe als PDF-Datei`nwurde nicht detektiert.`nDie Funktion wird jetzt abgebrochen!", "6 3")
							return
						}
						WinActivate, % Fenster_DruckAusgabe
						sleep, 200

					; Ordner in das Feld 'Dateiname' einsetzen
						If !VerifiedSetText("Edit1", SaveDir, Fenster_DruckAusgabe, 200) {
								PraxTT("Der Ordner für die Ausgabe in eine PDF-Datei`nkonnte nicht gesetzt werden.`nDie Funktion wird jetzt abgebrochen!", "6 3")
								return
						}
						If !VerifiedSetFocus("Edit1", Fenster_DruckAusgabe) {
								PraxTT("Der Ordner für die Ausgabe in eine PDF-Datei`nkonnte nicht übergeben werden.`nDie Funktion wird jetzt abgebrochen!", "6 3")
								return
						}
						ControlSend, Edit1, {Enter}, % Fenster_DruckAusgabe
						sleep, 200

					; Dateinamen der PDF in das Feld "Dateiname' schreiben
						If !VerifiedSetText("Edit1", filename, Fenster_DruckAusgabe, 200) {
								PraxTT("Der Name für die PDF-Datei`nkonnte nicht gesetzt werden.`nDie Funktion wird jetzt abgebrochen!", "6 3")
								return
						}
						If !VerifiedSetFocus("Edit1", Fenster_DruckAusgabe) {
								PraxTT("Der Name für die PDF-Datei`nkonnte nicht übergeben werden.`nDie Funktion wird jetzt abgebrochen!", "6 3")
								return
						}
						ControlSend, Edit1, {Enter}, % Fenster_DruckAusgabe
						sleep, 200
				}

			; Karteikarte wieder öffnen
				AlbisKarteikarteZeigen()

			; Wiederherstellen der vorherigen Druckeinstellungen
				Albismenu(57606, Fenster_Druckeinrichtung, 3, 1)

				while !( ControlGet("Choice", "", "ComboBox1", Fenster_Druckeinrichtung) = oDrucker.ActivePrinter.Name )	{
					Control, ChooseString, % oDrucker.ActivePrinter.Name, ComboBox1, % Fenster_Druckeinrichtung
					If (A_Index > 10)
						break
					sleep, 200
				}

				Control, ChooseString, % oDrucker.ActivePrinter.Size 	, ComboBox2, % Fenster_Druckeinrichtung
				Control, ChooseString, % oDrucker.ActivePrinter.Source, ComboBox3, % Fenster_Druckeinrichtung

				If (oDrucker.ActivePrinter.Portrait = 1)
					VerifiedCheck("Button5", Fenster_Druckeinrichtung)	;Checkbox	: Hochformat
				else
					VerifiedCheck("Button6", Fenster_Druckeinrichtung)	;Checkbox	: Querformat

				VerifiedClick("Button7", Fenster_Druckeinrichtung)          	;Button      	: Ok
		}

return 1
}
;07
AlbisLaborblattExport(PrintRange, SaveAs="", Printer="", PrintAnnotation=1) {          	;-- PDF Export oder Druckausgabe des Laborblattes

	; ACHTUNG: 	Funktion setzt nur die Inhalte der Steuerelemente anhand der übergebenen Parameter
	;                    	der Export/Druck als PDF Datei erfolgt über AlbisSavePDF()
	;
	; letzte Änderung am 11.09.2021

	/* PARAMETER BESCHREIBUNGEN

		PrintRange     	:= Datum als String im Format "DD.MM.YYYY-DD.MM.YYYY"
							    	 oder eine Zahl für "letzte Spalten"
							     	 oder das Wort "Alles" um alle Spalten auszugeben
		SaveAs          	:= Dateipfad mit Dateinamen (Dateierweiterungen wie .pdf werden entfernt )
		Printer           	:= empfohlener Drucktreiber  "Microsoft Print to PDF"
		PrintAnnotation	:= 1 oder true um Anmerkungen/Probedaten zu drucken

	 */

	; Parameter parsen und prüfen ;{
		PrintRange	:= Trim(PrintRange)
		SaveAs      	:= RegExReplace(SaveAs, "i)\.[a-z\d]+$")
		;~ Printer       	:= Printer ? Printer : "Microsoft Print to PDF"

	  ; SAVEAS
		If RegExMatch(SaveAs, "[A-Z]\:\\") {

			tmpSaveAs :=  SavesAs ".pdf"
			SplitPath, tmpSaveAs, SaveName, SaveDir
			If !InStr(FileExist(SaveDir "\"), "D") {
				PraxTT(A_ThisFunc ": Der übergebene Dateiordner existiert nicht`n[" SaveDir "]", "3 1")
				return 0
			}

			If FileExist(saveAs ".pdf")
				FileDelete, % saveAs ".pdf"
		}
		else
			SaveAs := ""

	  ; PrintRange
		If !RegExMatch(PrintRange, "\d{2}\.\d{2}\.\d{4}\-\d{2}\.\d{2}\.\d{4}") && !RegExMatch(PrintRange, "^\d+$") && (PrintRange != "Alles") {
			PraxTT(A_ThisFunc ": PrintRange - Es muss ein Datumsbereich, eine Zahl oder das Wort 'Alles' übergeben werden!`n[" PrintRange "]", "3 1")
			sleep 3000
			return 0
		}

	  ; Printer
		If (StrLen(Printer) = 0) {
			PraxTT(A_ThisFunc ": Es wurde kein Drucker übergeben!", "3 1")
			sleep 3000
			return 0
		}

	;}

	; Laborblatt anzeigen falls möglich
		If !InStr(AlbisGetActiveWindowType(), "Laborblatt")	{
			If !InStr(AlbisGetActiveWindowType(), "Patientenakte")		{
				PraxTT("Der Laborblattdruck kann nur bei einer geöffneten`nKarteikarte ausgeführt werden!.", "3 1")
				sleep 3000
				return 0
			}
			If !AlbisLaborblattZeigen(true) {
				PraxTT("Das Anzeigen des Laborblattes ist fehlgeschlagen.", "3 1")
				return 0
			}
		}

	; startet den Druckdialog
		If !(hLaborblattdruck := Albismenu(57607, "Laborblatt Druck ahk_class #32770", 3, 1)) {
			PraxTT(A_ThisFunc ": Der Laborblatt Druck konnte nicht ausgeführt werden!", "3 1")
			sleep 3000
			return 0
		}

	; Einstellungen eintragen
		If RegExMatch(PrintRange, "(?<Von>\d\d\.\d\d\.\d\d\d\d)\-(?<Bis>\d\d\.\d\d\.\d\d\d\d)", rx) || (PrintRange = "Alles") {
			If !VerifiedClick("Button3", hLaborblattdruck)                    ; Zeitraum
				return 0
			sleep 200
			If !(PrintRange = "Alles") {
				If !VerifiedSetText("Edit2", rxVon	, hLaborblattdruck)
					return 0
				sleep 200
				If !VerifiedSetText("Edit3", rxBis	, hLaborblattdruck)
					return 0
				sleep 200
			}
		}
		else {
			If !VerifiedClick("Button2", hLaborblattdruck)                 	 ; letzte
				return 0
			sleep 200
			If !VerifiedSetText("Edit1", (PrintRange = 0 ? 1 : PrintRange), hLaborblattdruck)
				return 0
			sleep 200
		}

	 ; Anmerkungen und Probedate
		VerifiedCheck("Button5", hLaborblattdruck,,, PrintAnnotation)

return AlbisSaveAsPDF(saveAs, Printer, true)   ; (Dateipfad, Names des Druckers, Laborblattdruck = ja)
}
;08
AlbisLaborblattZeigen(Info:=true) {                                                                           	;-- schaltet auf das Laborblatt

	; letzte Änderung: 24.02.2021

	; Info anzeigen
		If Info
			PraxTT("bitte warten ...`nLaborblatt wird aufgerufen!", "2 1")

	; Albis aktivieren
		AlbisActivate(2)

	; prüfen auf geöffnete Karteikarte
		If !InStr(activeWT:= AlbisGetActiveWindowType(), "Patientenakte")
			return 0

	; Laborblatt aufrufen und Erfolg prüfen
		Loop {

			PostMessage, 0x111, 33034,,, % "ahk_id " AlbisWinID()
			while !InStr(AlbisGetActiveWindowType(), "Laborblatt") 	{
				if  (A_Index > 50)
					break
				else
					sleep 100
			}

			If InStr(AlbisGetActiveWindowType(), "Laborblatt") {
				sleep 400 		; Zeichnen abwarten
				return 1
			}
			else If (A_Index > 4)
				return 0

		}

return 0
}
;09
AlbisLaborAuswählen(Laborname) {                                                                       	;-- für das Fenster Labor auswählen

	; übergebener Laborname wird im Dialog ausgewählt und anschliessend mit Ok bestätigt
	; vermutlich erscheint dieses Fenster nicht wenn man nur ein Labor zur Auswahl hat
		If (StrLen(Laborname) > 0) {
			ControlGet, currSel, Choice,, ListBox1, Labor auswählen ahk_class #32770
			If !InStr(currSel, Laborname) 	{
				Control, ChooseString, % Laborname, ListBox1, Labor auswählen ahk_class #32770
				sleep, 200
				If (A_Index > 10) 	{
					MsgBox, 1, Addendum für Albis on Windows, % "Das eingestellte Labor: " Laborname "`nkonnte nicht ausgewählt werden.", 6
					return 0
				}
			}
		}

		err:= VerifiedClick("Button1", "Labor auswählen ahk_class #32770")
		PraxTT("Warte auf den nächsten Laborabruf-Dialog", "2 1")
		WinWait, ALBIS, Keine Datei(en) im Pfad, 2
		If (ErrorLevel = 0) 	{
				err:= VerifiedClick("Button1", "ALBIS", "Keine Datei(en) im Pfad")
				Addendum.Labor.AbrufDaten	:= false
				Addendum.Labor.AbrufStatus	:= false
				Addendum.Labor.AbrufVoll    	:= false
				return err
		}

return err
}
;10
AlbisLaborDaten()	{                                                                                                 	;-- bearbeitet das "Labordaten importieren" Fenster und öffnet im Anschluss das Laborbuch
	VerifiedCheck("Button5"	, "Labordaten ahk_class #32770",,, 1)
	VerifiedClick("Button1"	, "Labordaten ahk_class #32770")
	sleep, 200
	AlbisLaborbuchOeffnen()
}
;11
AlbisLaborAlleUebertragen() {                                                                                	;-- Alle Übertragen im Laborbuch auslösen

	; Laborbuch aufrufen falls nicht angezeigt
		hLabbuch := AlbisLaborbuchOeffnen()

	; wm_command, wParam = 34157 = 'alle übertragen' (ToolbarWindow321)
		SendMessage, 0x111, 34157,,, % "ahk_class OptoAppClass"
		EL := ErrorLevel

	; Bestätigungsdialog abwarten und auf Ja drücken
		WinTitle	:= "ALBIS ahk_class #32770"
		WinText 	:= "Möchten Sie die zugeordneten Anforderungen"
		WinWait, % WinTitle, % WinText, 30
		If (hDialog := WinExist(WinTitle, WinText)) {
			If !(res := VerifiedClick("Ja",,, hDialog, true))
				(res := VerifiedClick("Button1",,, hDialog, true))
		}

return {"EL":EL, "click":res}
}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; GRAFISCHER BEFUND                                                                                                                                                                   	(05)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisOeffneGrafischerBefund	(02) AlbisUebertrageGrafischenBefund                                           	(03) AlbisImportierePdf
; (04) AlbisImportiereBild               	(05) AlbisBrief
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisOeffneGrafischerBefund() {                                                             	;-- öffnet den Dialog 'Grafischer Befund' - importieren von Bilddateien (jpg z.B.)
return Albismenu(32960, "Grafischer Befund ahk_class #32770")
}
;02
AlbisUebertrageGrafischenBefund(ImgName, KKText:="") {                      	;-- Funktion mit Sicherheitsüberprüfung falls ein Focus verloren geht

		; letzte Änderung: 17.02.2021

		static WinText1	:= "Bitte überprüfen Sie den Namen bzw. Pfad."
		static GRBFD 	:= "Grafischer Befund ahk_class #32770"

	; ein fehlender Dateibez. wird automatisch ersetzt
		If !RegExMatch(ImgName, "\.\w+$")
			ImgName .= ".pdf"

	; KKText ist leer
		If (KKText = "")
			KKText := SubStr(ImgName, 1, StrLen(ImgName) - 4)

	; Dialog aufrufen falls noch nicht geschieht
		If !(hGBef:= WinExist("Grafischer Befund ahk_class #32770"))
			return 2
		WinActivate   	, % "ahk_id " hGBef
		WinWaitActive	, % "ahk_id " hGBef,, 4

	; der Haken bei Vorschau wird entfernt, dann muss ich keine PDF Viewer oder Imageviewer Fenster abfangen und abhandeln mehr
		VerifiedCheck("Button1", "", "", hGBef, 0)
		sleep, 200

	;Daten in das Fenster eintragen
		VerifiedSetText("Edit1", ImgName	, GRBFD, 100)
		VerifiedSetText("Edit2", KKText 		, GRBFD, 100)

	;auf Ok drücken - ein mögliches Fehlerfenster abfangen und schliessen (beendet den Importvorgang)
		while WinExist(GRBFD)		{

			VerifiedClick("Button2", GRBFD)
			sleep, 300
			If WinExist(GRBFD)                                         	{
				WinWaitClose, % GRBFD,, 2
				sleep, 300
			}
			If WinExist("ALBIS ahk_class #32770", WinText1)	 {
				If !VerifiedClick("Button1", "ALBIS ahk_class #32770", WinText1)
					If WinExist("ALBIS ahk_class #32770", WinText1)
						AlbisClosePopups()
				return 3
			} else {
				return 1
			}


			If (A_Index > 2)
				return 4
			else
				Sleep, 500

		}

return 1
}
;03
AlbisImportierePdf(PdfName, KKText:="", DocDate:="") {                          	;-- Importieren einer PDF-Datei in eine Patientenakte

	/* AlbisImportierePdf()	-	letzte Änderung: 17.02.2021

		benötigt ein globales Objekt mit dem Namen SPData -> Aufbau siehe ScanPool.ahk
		oPat : globals Object, enthält die Addendum interne Patientendatenbank
		Addendum: globales Object - enthält Daten zu Albis-Ordnern
		Das Datum des Karteikarteneintrages wird entweder aus DocDate oder wenn leer dem Dateierstellungsdatum entnommen.
		Ist DocDate im Format dd.MM.yyyy wird das Erstellungsdatum der Datei verwendet

	 */

		suggestions:= Object()

	; zurück falls Datei nicht mehr vorhanden ist
		If !FileExist(Addendum.BefundOrdner "\" PdfName) {
			PraxTT("Datei: " PdfName " ist nicht vorhanden!", "1 0")
			return 0
		}

	; Eingangsdatum des Befundes auslesen
		PraxTT("Importiere PDF-Datei:`n[" PdfName "]", "15 2")
		BlockInput, On
		If !DocDate || !RegExMatch(DocDate, "\d\d\.\d\d\.\d\d\d\d")
			DocDate := FormatedFileCreationTime(Addendum.BefundOrdner "\" PdfName)

	; Programmdatum ändern
		AlbisSetzeProgrammDatum(DocDate)

	; Eingabefokus in die Karteikarte setzen
		If !AlbisKarteikarteFocus("Edit3") {
			PraxTT("Beim Importieren ist ein Problem aufgetreten.`nDer Eingabefocus konnte nicht gesetzt werden!", "6 2")
			AlbisSetzeProgrammDatum()
			BlockInput, Off
			return 0
		}
		SendRaw, Scan
		sleep, 100
		SendInput, {Tab}
		WinWait, % "Grafischer Befund ahk_class #32770",, 3

	; Erstellen des Karteikartentextes
		If !KKText
			KKText := string.Karteikartentext(PdfName)

	; Übertragen des Befundes in die Akte
		err := AlbisUebertrageGrafischenBefund(PdfName, KKText)

	; Karteikartenfokus entfernen, ans Ende der Karteikarte springen
		AlbisKarteikarteAktivieren()
		SendInput, {Escape}
		sleep, 300
		SendInput, {End}
		sleep, 50

	; auf aktuelles Tagesdatum zurücksetzen und Tastaturzugriff wieder zulassen
		AlbisSetzeProgrammDatum()
		BlockInput, Off

	; Hinweis anzeigen
		If (err > 1)
			PraxTT("Beim Übertragen des Befundes`nist ein Problem aufgetreten.`nFehlercode: " err, "6 2")
		else if (err = 1)
			PraxTT("Der PDF-Befund wurde importiert.", "3 2")

return (err > 1 ? 0 : DocDate)
}
;04
AlbisImportiereBild(ImageName, KKText:="", DocDate:="") {                    	;-- Importieren einer JPG-Datei in eine Patientenakte

	; Änderung des Tagesdatum wie bei AlbisImportierePdf() beschrieben
	; letzte Änderung: 17.02.2021

	; zurück falls Datei nicht mehr vorhanden ist
		If !FileExist(Addendum.BefundOrdner "\" ImageName) {
			PraxTT("Datei: " ImageName " ist nicht vorhanden!", "1 0")
			return 0
		}

	; Dialogfenster aufrufen über das Albismenu
		BlockInput, On
		If !(hGBef := AlbisOeffneGrafischerBefund()) 	{
			PraxTT(	"Der Dialog zum Importieren der Bilddatei:`n<" ImageName ">`n"
					. 	"konnte nicht geöffnet werden.`nDer Importvorgang wird abgebrochen.", "2 2")
			BlockInput, Off
			return 0
		}
		PraxTT("Importiere Bild:`n[" ImageName "]", "15 2")

	; Dialog aktivieren
		WinActivate   	, % (GRBFD := "Grafischer Befund ahk_class #32770")
		WinWaitActive	, % GRBFD,, 4

	; Dokumentdatum prüfen
		If !DocDate || !RegExMatch(DocDate, "\d\d\.\d\d\.\d\d\d\d")
			DocDate := FormatedFileCreationTime(Addendum.BefundOrdner "\" ImageName)

	; Tagesdatum ändern
		AlbisSetzeProgrammDatum(DocDate)

	; Erstellen des Karteikartentextes
		If !KKText
			KKText := string.Karteikartentext(ImageName)

	; Importieren des Bildes in die Akte
		err := AlbisUebertrageGrafischenBefund(ImageName, KKText)

	; auf das aktuelle Tagesdatum zurücksetzen
		AlbisSetzeProgrammDatum()
		BlockInput, Off
		If !err
			PraxTT("Beim Importieren des Bildes ist ein Problem aufgetreten.", "6 2")
		else
			PraxTT("Der Bildbefund wurde importiert.", "3 2")

return (!err ? 0 : DocDate)
}
;05
AlbisBrief(Dokumentname) {                                                                  	;-- Textvorlagen anzeigen

	/* BESCHREIBUNG

		öffnet ein Dokument über den Textvorlagen-Dialog
		z.B. AlbisBrief("arztbrief.docx")

	*/

	; ohne geöffnete Karteikarte nicht ausführen
		If !AlbisKarteikarteAktiv()
			return 0

	; Menu extern/Arztbrief aufrufen
		AlbisClosePopups()
		hVorlagen := Albismenu(32829, "Vorlagen ahk_class #32770")

	; Dokument auswählen und OK drücken
		If (EL := VerifiedChoose("ComboBox1"	, hVorlagen, "Alle anzeigen") <> 1)
			return EL
		If (EL := VerifiedChoose("ListBox1"    		, hVorlagen, Dokumentname) <> 1)
			return EL
		If (EL := VerifiedClick("OK", "Vorlagen ahk_class #32770",,, true) <> 1)
			return EL

return 1
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; PROTOKOLLE, STATISTIKEN - ERSTELLUNG UND AUSWERTUNG                                                                                                 	(03)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisErstelleTagesprotokoll       	(02) AlbisErstelleZiffernStatistik     	(03) AlbisListeSpeichern
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
; 01
AlbisErstelleTagesprotokoll(Options:="", SaveFolder:="", CloseProtokoll:=1) {                                                       	;--	Erstellen und Speichern von Tagesprotokollen

	;Parameter Periode: 	    1. leerer String                         	- setzt den Zeitraum auf den ersten und letzten Tag des aktuellen Quartals
	;                            		2. mmyy (z.B. 0219)                 	- es wird ein Tagesprotokoll mit dem ersten und letzten Tag des übergebenen Quartals erstellt
	;                            		3. "dd.dd.dddd[,/-]dd.dd.dddd"	- von bis Datumsübergabe ist auch als String mit zwei Tagesdaten welche durch ein "-" oder "," getrennt sind möglich
	;									4. Übergabe eines Objektes		- Schlüssel: Patienten (Alle, mit_Einträgen, aktiver_Patient), Kürzel (Alle, Filtername), Periode (wie oben)
	;
	; Beispiel               :		AlbisErstelleTagesprotokoll("04.01.2016-06.01.2016", AddendumDir "\Tagesprotokolle\TP-Abrechnungshelfer", 1)
	;
	; letzte Änderung	:		11.09.2021

		static Periode
		static TPWin 	:= "Tagesprotokoll ahk_class #32770"
		Quartal     	:= Object()
		SaveFolder	:= RTrim(SaveFolder, "\")

	; DATEN DES TAGESPROTOKOLLFENSTER IN Controls LEEREN
		Controls("","Reset", "")

	; TAGESPROTOKOLLFENSTER AUFRUFEN
		hTprotWin	:= Albismenu(32802, TPWin)
		If !WinExist(TPWin)
			return "invoke error"

	; OPTIONS PARSEN
		If !IsObject(Options)
			Periode := Options
		else
			Periode := Options.Periode

	; DATUM PARSEN
		Periode:= RegExReplace(periode, "[\s\/]")
		If  RegExMatch(Periode, "^\d{3,4}$", QuartalAkt)		{  ; 0120 oder 221 z.B.
				Quartal.aktuell	:= SubStr("0" QuartalAkt, -1)
				Quartal         	:= QuartalTage(Quartal)
				QNr             	:= SubStr(QuartalAkt, 2, 1)
				ErsterTag 	:= Quartal.TDBeginn
				LetzterTag	:= Quartal.TDEnde
		}
		else if (StrLen(Periode) = 0)                                    	{
				Quartal.aktuell	:= GetQuartal("heute")
				Quartal         	:= QuartalTage(Quartal)
				QNr             	:= SubStr(QuartalAkt, 2, 1)
				ErsterTag       	:= Quartal.TDBeginn
				LetzterTag     	:= Quartal.TDEnde
		}
		else If RegExMatch(Periode, "^\s*(?<1>\d{1,2}.\d{1,2}.\d{2,4})[,\s\-]+(?<2>\d{1,2}.\d{1,2}.\d{2,4})", Tag) {
				Quartal.ErsterTag 	:= Tag1
				Quartal.LetzterTag	:= Tag2
				ErsterTag 	:= Tag1
				LetzterTag	:= Tag2
		}
		else	{
			throw A_ThisFunc " (" A_LineFile ") : Error in function call, a date must be passed in the following format`n - dd.mm.yyyy[ -|,]dd.mm.yyyy or qq[// ]yy"
			return 1
		}

	; FELDER VON UND BIS MIT DEN JEWEILS ÜBERGEBEN DATEN FÜLLEN
		If !Controls("Edit1", "SetText," ErsterTag, hTprotWin) {
			PraxTT("Das Anfangsdatum (von) konnte nicht gesetzt werden.", "6 3")
			return 1
		}
		If !Controls("Edit2", "SetText," LetzterTag, hTprotWin) {
			PraxTT("Das Enddatum (bis) konnte nicht gesetzt werden.", "6 3")
			return 1
		}

	; STEUERELEMENTE IM TAGESPROTOKOLL BEARBEITEN
		If IsObject(Options) {

		; KÜRZEL
			optFilter := Options.filter
			If (optFilter = "Alle")                                                     		; Alle
				VerifiedClick("Button4", hTprotWin)
			else {
				VerifiedClick("Button5", hTprotWin)                                	; Filter
				sleep 80
				If !VerifiedChoose("ComboBox1", hTprotWin, optFilter) {
					PraxTT("Fehler bei der Kürzelauswahl", "1 3")
					return -3
				}
			}

		; PATIENTEN
			optPat := Options.Patienten
			Switch optPat 			{

				Case "Alle":
					If !VerifiedClick("Button9", hTprotWin)
						return -4

				Case "Mit_Einträgen":
					If !VerifiedClick("Button10", hTprotWin)
						return -5

				Case "aktiver_Patient":
					If !VerifiedClick("Button11", hTprotWin)
						return -6
			}

		; AUSGABE
			If (Options.HasKey("anamnestisch") || Options.HasKey("Behandlung"))
				Options.Dauerdiagnosen := true

			If Options.HasKey("Dauerdiagnosen") {
				VerifiedCheck("Button23", hTprotWin,,, Options.Dauerdiagnosen)
				If Options.Dauerdiagnosen {
					If Options.HasKey("anamnestisch")
						VerifiedCheck("Button33", hTprotWin,,, Options.anamnestisch)
					If Options.HasKey("Behandlung")
						VerifiedCheck("Button34", hTprotWin,,, Options.Behandlung)
				}
			}

			If Options.HasKey("Dauermedikamente")
				VerifiedCheck("Button24", hTprotWin,,, Options.Dauermedikamente)

			If Options.HasKey("IK_Vers_Nr")
				VerifiedCheck("Button25", hTprotWin,,, Options.IK_Vers_Nr)

			If Options.HasKey("Cave")
				VerifiedCheck("Button26", hTprotWin,,, Options.Cave)

			If Options.HasKey("Hinweis_bei_fehlender_Diagnose")
				VerifiedCheck("Button28", hTprotWin,,, Options.Hinweis_bei_fehlender_Diagnose)

			If Options.HasKey("Sortierung_nach_Namen")
				VerifiedCheck("Button29", hTprotWin,,, Options.Sortierung_nach_Namen)

			If Options.HasKey("Diagnosen_mit_Scheinbezug")
				VerifiedCheck("Button30", hTprotWin,,, Options.Diagnosen_mit_Scheinbezug)

			If Options.HasKey("Telematikinfrastruktur_Hinweistext")
				VerifiedCheck("Button35", hTprotWin,,, Options.Telematikinfrastruktur_Hinweistext)

		}

	; NUTZERINTERAKTION NOTWENDIG
		MsgBox, 0x4, % StrReplace(A_ScriptName, ".ahk"), % "Bitte die Einstellungen und prüfen und mit JA bestätigen."
		IfMsgBox, No
			return

	; EINEN MAUSCLICK AN DEN 'OK' BUTTON DES FENSTER SENDEN
		If !Controls("", "ControlClick, OK, Button", hTprotWin)
			Controls("Button31", "Click use ControlClick", hTprotWin)
		WinWait, % "Bitte warten... ahk_class #32770",, 10
		If ErrorLevel && WinExist(TPWin) {
			PraxTT("Der ausgelöste Click auf den 'OK' Button`nim Tagesprotokoll-Dialog ist fehlgeschlagen.`nEs wurde als letztes ein simulierter Mausklick versucht.", "1 3")
			WinClose, % TPWin
			return 1
		} else if ErrorLevel && !WinExist(TPWin) {
			PraxTT("Es ist ein unbekannter Fehler aufgetreten,`nwelcher die Erstellung des Tagesprotokoll verhindert hat.", "1 3")
			return 1
		}

	; DATEN ZUM TAGESPROTOKOLLFENSTER IN Controls LEEREN
		Controls("","Reset", hTprotWin)

	; AUF DAS FERTIG ERSTELLTE TAGESPROTOKOLL WARTEN
		bitteWarten:= false
		PraxTT("Warte auf die vollständige Erstellung des Tagesprotokoll!", "0 3")
		while !WinExist("ALBIS - [Tagesprotokoll") {
				Sleep, 500
			; läßt den Loop pausieren
				If hBWin := WinExist("Bitte warten ahk_class #32770") {
					WinGetText, wText, % "ahk_id " hBWin
					If InStr(wText, "Tagesprotokoll") {
							bitteWarten := true
							BWin	:= GetWindowSpot(hBWin)
							TPWin	:= GetWindowSpot(WinExist(TPWin))
							SetWindowPos(hBwin, TPWin.X + Floor(TPWin.W//2) - Floor(BWin.W//2), TPWin.Y + TPWin.H + 5, BWin.W, BWin.H)
							WinSet, ExStyle, +0x8, % "ahk_id " hBWin	;+ WS_EX_TOPMOST
							WinWaitClose, % "Bitte warten... ahk_class #32770"
					}
				}
		}
		PraxTT("", "off 3")

	; SPEICHERN DES ANGEZEIGTEN TAGESPROTOKOLL AUF FESTPLATTE
		If SaveFolder		{

				Controls("","Reset", "")

			; Tastatur- und Mauseingriffe durch den Nutzer kurzzeitig sperren
				hk(1, 0, "!ACHTUNG!`n`nDie Tastatureingaben sind für maximal 10 Sekunden gesperrt", 10)

			; Tagesprotokoll speichern und die Tagesprotokollausgabe schließen
				hSaveAsWin	:= Albismenu(33014, "Speichern unter")
				WinWaitActive, % "Speichern unter ahk_class #32770",, 10
				If ErrorLevel	{
					MsgBox, 0	, % StrReplace(A_ScriptName, ".ahk")
									, % "Erwarteter 'Speichern unter' - Dialog konnte nicht abgefangen werden!`n"
									. 	  "Die Funktion wird abgebrochen."
					return 1
				}

			; Dateinamen generieren
				If !IsObject(Options)
					ProtokollFileName := SaveFolder "\" (StrLen(Quartal.aktuell) = 0 ? Quartal.ErsterTag "-" Quartal.LetzterTag : Quartal.Jahr "-" QNr) ".txt"
				else
					ProtokollFileName := SaveFolder "\" Options.filename ".txt"

			; Ordner und Dateinamen im Speichern unter Dialog eintragen
				If FileExist(ProtokollFileName)
					FileDelete, % ProtokollFileName

				Controls("Edit1", "SetText," ProtokollFileName, hSaveAsWin)

			; Speichern unter drücken
				Controls("Button2", "Click use Controlclick", hSaveAsWin)
				sleep 200

			; Warten auf einen eventuellen "Speichern unter bestätigen" - Dialog
				PraxTT("Warte bis zu 10 Sekunden auf einen möglichen Speichern unter... Dialog!", "10 2")
				WinWait, % "Speichern unter bestätigen ahk_class #32770",, 10
				If !ErrorLevel
					Controls("Button1", "Click use ControlClick", "Speichern unter bestätigen ahk_class #32770")

			; Tastatur- und Mauseingriffe wieder entsperren
				hk(0, 0, "Tastatur- und Mausfunktion sind wieder hergestellt!", 2)
		}

	; SCHLIESSEN DES PROTOKOLLS, WENN PARAMETER GESETZT
		If CloseProtokoll
			AlbisMDIChildWindowClose("Tagesprotokoll")

		PraxTT("Das Tagesprotokoll wurde erstellt und gespeichert!", "6 3")

return ProtokollFileName		; ein vollständiger Dateipfad als Rückgabe = erfolgreicher Funktionsablauf
}
; 02
AlbisErstelleZiffernStatistik(Quartal="aktuell",Zeitraum="",Tag="",Arztwahl="",SaveToFile="",CloseStatistik:=true) {	;-- Erstellen von Ziffernstatistiken

	; Menu Statistik\Leistungsstatistik\EBM2000Plus\2009Ziffernstatistik
	; ACHTUNG: Funktion kommuniziert mit Addendum.ahk - das Skript sollte vorher aufgerufen sein
	; Parameter Arztwahl: String in der Form "BSNR 123456678" oder "Arzt Dr.Mustergültig" oder "Person Dr.Kontrolle" zur Auswahl der Arztwahlfelder

		static WM_MDIMAXIMIZE:= 0x0225

		M5UserBreak:= false
		Erstellungszeitraum := "das akutelle Quartal"

	; kein Überprüfen des erfolgreichen Dialogaufrufes über die AlbisMenu() Funktion, da der User als Administrator eingeloggt sein muss
		If !WinExist("Ziffernstatistik ahk_class #32770")
			hZfStat:= Albismenu(34293)		                                    	;EBM 2000plus/2009 Ziffernstatistik = 34293

	; prüft ob als Administrator eingeloggt wurde, sonst kann keine Ziffernstatistik erstellt werden, wenn ja wird der der Dialog aufgerufen
		Loop		{
			If WinExist("Ziffernstatistik ahk_class #32770")
				break
			else if WinExist("ALBIS ahk_class #32770", "Zugang verweigert")	{
						MsgBox, 4, % "Abrechnungshelfer",
						(LTrim
						Um eine Ziffernstatistik zu erstellen müssen Sie sich
						mit Ihrem Administrator oder Hauptkennwort anmelden!
						Soll ich Sie dafür jetzt aus Albis ausloggen?
						)
						VerifiedClick("Button1", "ALBIS ahk_class #32770", "Zugang verweigert")
						IfMsgBox, No
							return
						IfMsgBox, Yes
						{
							;- Skriptkommunikation -0x555 - für Kommunikation mit Addendum.ahk, Befehl|Name des abfragenden Skriptes, Name des empfangenen Skriptes
								If (AddendumID := GetAddendumID())	{
									MessageReceived := ""
									OnMessage(0x4a, "Receive_WM_COPYDATA")
									Send_WM_COPYDATA("AutoLogin Off|" A_ScriptHwnd, AddendumID)
									while !InStr(MessageReceived, "AutoLogin disabled")
										Sleep, 100
									PraxTT("AutoLogin ist ausgestellt.", "3 1")
									MessageReceived := ""
								}

							; Ausloggen und warten auf den erneuten Login
								AlbisLogout()

								PraxTT("Warte auf den Abschluß des Login-Vorganges.`nSie können an dieser Stelle die weitere`nSkriptausführung mit >>Escape<< abbrechen.", "0 3")
								while, WinExist("ALBIS - Login ahk_class #32770") 	{
									If GetKeyState("Escape")	{
										M5UserBreak:= true
										break
									}
									Sleep, 50
								}

								PraxTT("", "Off")

							; AutoLogin wieder einschalten
								If (AddendumID := GetAddendumID())	{
									Send_WM_COPYDATA("AutoLogin On|" A_ScriptHwnd, AddendumID)
									while !InStr(MessageReceived, "AutoLogin enabled")
										Sleep, 100
									PraxTT("AutoLogin ist wieder aktiviert.", "3 1")
									MessageReceived := ""
								}

							; Selbstaufruf der Funktion für den Neustart
								AlbisErstelleZiffernStatistik(Quartal, Zeitraum, Tag)
						}
				}
			else If WinExist("ALBIS ahk_class #32770", "Fehler beim Aufruf dppivassist")
				VerifiedClick("Button1","ALBIS ahk_class #32770", "Fehler beim Aufruf dppivassist")

			Sleep, 100
		}

	; Abbruch wenn das Dialogfenster nicht geöffnet werden könnte
		If !WinExist("Ziffernstatistik ahk_class #32770") {
				PraxTT("Das Ziffernstatistik Fenster konnte nicht geöffnet werden`nDie Funktion wird jetzt beendet.", "5 1")
				return 0
		}

	; Daten werden jetzt in die Steuerelemente des Dialogfensters eingetragen
		PraxTT("Ich starte jetzt mit der Erstellung der Ziffernstatistik", "3 1")
		VerifiedClick("Button1", "Ziffernstatistik ahk_class #32770") ; Quartal
		If (StrLen(Zeitraum) > 0)	                                	{

				RegExMatch(Zeitraum, "(?<von>[\d\.]+)\s*,\s*(?<bis>[\d\.]+)", Datum)
				If !RegExMatch(DatumVon, "\d{2}\.\d{2}\.\d{4}") || !RegExMatch(DatumBis, "\d{2}\.\d{2}\.\d{4}") {
					throw Exception("Datumsformat(TT.MM.JJJJ) wurde nicht eingehalten <" Zeitraum ">")
					return 0
				}

				Erstellungszeitraum := "den Zeitraum " DatumVon " bis zum " DatumBis

				VerifiedClick("Button2", "Ziffernstatistik ahk_class #32770") ; Leistungen im Zeitraum
				VerifiedSetText("Edit1", DatumVon, "Ziffernstatistik ahk_class #32770")
				VerifiedSetText("Edit2", DatumBis, "Ziffernstatistik ahk_class #32770")
		}
		else If (StrLen(Tag) > 0)                                 		{
				If !RegExMatch(Tag, "\d\d\.\d\d\.\d\d\d\d") {
						throw Exception("Datumsformat(TT.MM.JJJJ) wurde nicht eingehalten <" Tag ">")
						return 0
				}

				Erstellungszeitraum := "den Tag " Tag

				VerifiedClick("Button3", "Ziffernstatistik ahk_class #32770") ; Tag
				VerifiedSetText("Edit3", Tag, "Ziffernstatistik ahk_class #32770")
		}
		else If !RegExMatch(Quartal, "^\s*aktuell\s*$")		{
			If !RegExMatch(Quartal, "\d\/\d\d") {
				throw Exception("Quartalsformat(Q/JJ) wurde nicht eingehalten <" Quartal ">")
				return 0
			}
			Erstellungszeitraum := "das Quartal <" Quartal ">"
			Control, ChooseString, % Quartal, ComboBox1, Ziffernstatistik ahk_class #32770
		}

		VerifiedClick("Button22", "Ziffernstatistik ahk_class #32770")	; %Durchschnitt
		VerifiedClick("Button24", "Ziffernstatistik ahk_class #32770")	; Optionen berücksichtigen
		VerifiedClick("Button23", "Ziffernstatistik ahk_class #32770")	; FG Vergleich
		VerifiedClick("Button25", "Ziffernstatistik ahk_class #32770")	; Leistungstexte zeigen
		; den Button für Optionen und Einstellungen habe ich noch nie genutzt

	; drückt den Button Ok um die Statistik zu erstellen
		VerifiedClick("Button26", "Ziffernstatistik ahk_class #32770")	 ; OK

	; wartet bis die Erstellung abgeschlossen ist
		while !AlbisMDIChildHandle("Ziffernstatistik")		{
			If (hwnd:= WinExist("ahk_class #32770", "Scheine einlesen"))
				WinSetTitle, % "ahk_id " hwnd,, % "Ziffernstatistik für " Erstellungszeitraum " wird erstellt (" A_Index "s)"
			Sleep, 1000
		}

	; wenn abgeschlossen, dann wird das Statistikfenster innerhalb des Albisfenster maximiert
		AlbisActivate(1)
		If (hStatistikWin:= AlbisMDIChildHandle("Ziffernstatistik"))
			If !AlbisMDIMinMaxStatus(hStatistikWin)                                                                                       	; MDI Fenster ist nicht vergrößert dann maximieren
				PostMessage, % WM_MDIMAXIMIZE, % hStatistikWin,,, % "ahk_id " AlbisMDIClientHandle()


	; kurze Pause damit Windows wieder reagieren kann
		sleep 1000

	; wenn angegeben wird das erstellte Protokoll gespeichert
		If (StrLen(SaveToFile) > 0)
			savedAs:= AlbisDateiSpeichern(SaveToFile, true)

	; wenn CloseStatistik = true wird das Statistikfenster geschlossen
		If CloseStatistik
			AlbisMDIChildWindowClose("Ziffernstatistik")

	; wenn savedAs vorhanden ist wird dieser String zurückgegeben
		If (StrLen(savedAs) > 2)
			return savedAs

return 1 	; ansonsten wird nur eine 1 zurückgegeben
}
; 03
AlbisListeSpeichern(ListTitle, FilePath, ext:= "csv", overwrite:= true) {	                                                                  	;--	Albislisten speichern

	; Parameter:
	; 	ListTitle      	- sollte der Titel einer in Albis geöffneten Liste sein, z.B. offene Posten. Achtung: Der Listentitel wird als Dateiname weiterverwendet!
	;	FilePath     	- Dateiverzeichnis in dem diese Liste gespeichert werden soll
	;	ext            	- Dateiendung kann jede in Windows nutzbare sein z.B. txt, log, csv.
	;	overwrite   	- flag als true oder false falls eine bestehende Datei ohne nachfragen überschrieben werden soll, ansonsten erfolgt eine fortlaufende Nummerierung
	;
	; Rückgabewert:
	;	FullFilePath	- der komplette Pfad und der Dateiname der gespeicherten Datei


	; Aktivieren der Liste in Albis
		result := AlbisMDIChildActivate(ListTitle)
		If IsObject(result)
			return false	; muss ergänzt werden

	; Tastatur- und Mauseingriffe durch den Nutzer kurzzeitig sperren
		hk(1, 0, "!ACHTUNG!`n`nDie Tastatureingaben sind für maximal 10 Sekunden gesperrt", 10)

	; Tagesprotokoll speichern und die Tagesprotokollausgabe schließen
		If !(hSaveAsWin	:= Albismenu(33014, "Speichern unter ahk_class #32770")) {
				MsgBox, 0, Addendum für Albis on Windows, % "Erwarteter Albisdialog (Speichern unter)`nkonnte nicht abgefangen werden!`nDie Funktion wird hier abgebrochen", 10
				return 1
		}

	; ermitteln ob diese Datei schon erstellt wurde
		FullFilePath := FilePath "\" StrReplace(ListTitle, " ", "_") "." RegExReplace(ext, "^\.", "")
		If (!overwrite && FileExist(FullFilePath))
			FileExists:=1  ; #### dummy - Code ergänzen!
		else if (overwrite && FileExist(FullFilePath))
			FileDelete, % FullFilePath

		PraxTT("Warte auf das Ende des Speichervorgangs.", "0 5")

	; Dateipfad dem Albisdialog übergeben
		Controls("Edit1", "SetText, " FullFilePath, hSaveAsWin)

	; Speichern unter drücken
		Controls("Button2", "Click use Controlclick", hSaveAsWin)
		sleep 200

	; Warten auf einen eventuellen "Speichern unter bestätigen" - Dialog
		WinWait, % "Speichern unter bestätigen ahk_class #32770",, 6
		If WinExist("Speichern unter bestätigen ahk_class #32770")
			Controls("Button1", "Click use ControlClick", "Speichern unter bestätigen ahk_class #32770")

	; Tastatur- und Mauseingriffe wieder entsperren
		hk(0, 0, "Tastatur- und Mausfunktion sind wieder hergestellt!", 2)

return FullFilePath
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; ABRECHNUNG - KASSE UND PRIVAT                                                                                                                                              	(02)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAbrechnungVorbereiten 	(02) AlbisBehandlungsliste
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
AlbisAbrechnungVorbereiten(Set) {	                                                                                             	;-- bearbeitet den Dialog Abrechnung vorbereiten

	/*  Beispielaufruf:

		Set =
		(LTrim
		AODT, Check
		GNR-Regelwerkskontrolle, Check
		KRW-Regelprüfung, Check
		Obligat, Check
		Fakultativ, Check
		Ok, Click
		)
		AlbisAbrechnungVorbereiten(Set)

	 */

	Buttons := {	"Ok"                         		: {"Type":"Button"
										, "ClassNN": 	"Button1"}
		, "Abbruch"                          	        	: {"Type":"Button"
										, "ClassNN": "Button2"}
		, "Optionen"                        	        	: {"Type":"Button"
										, "ClassNN": "Button3"}
		, "ADT"                              	        	: {"Type":"Checkbox"
										, "ClassNN": "Button5"
										, "Standard": "Checked"
										, "Action": {"ToggleHide": {1:	"Static1"
														        	    	, 	2:	"Button6"
														            		, 	3:	"Button7"
														            		, 	4:	"Button8"
														            		, 	5:	"Button16"
														            		, 	6:	"Button17"
															            	, 	7:	"Button"}}}
		, "inkl."                            		        	: {"Type":"Checkbox"
     										, "ClassNN":	"Button6"
	    									, "Standard":	"Unchecked"}
		, "exkl."       	                    	        	: {"Type":"Checkbox"
		    								, "ClassNN":	"Button7"
			    							, "Standard":	"Checked"}
		, "ausschließlich"	                        	: {"Type":"Checkbox"
				    						, "ClassNN": 	"Button8"
					    					, "Standard": "Unchecked"}
		, "AODT"       						        	: {"Type":"Checkbox"
						    				, "ClassNN": 	"Button9"
							    			, "Standard": "UnChecked"
					    					, "Action":  	{"Unhide": {1: "Edit3", 2: "Edit4"}}}
		, "GNR-Regelwerkskontrolle" 	        	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button13"
											, "Standard": "Hide Unchecked"}
       	, "KRW-Regelprüfung" 	                	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button14"
											, "Standard": "Hide Unchecked"}
       	, "Obligat"                         	        	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button15"
											, "Standard": "Unchecked"}
       	, "Fakultativ"                      	        	: {"Type":"Checkbox"
								     		, "ClassNN": 	"Button16"
								    		, "Standard": "Unchecked"}
		, "Nein-Scheine"	                            	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button18"
	    							    	, "Standard": "Unchecked"}
		, "Vorquartalsscheine"                    	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button19"
									    	, "Standard": "Unchecked"}
		, "Scheine ohne Einlesedatum"        	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button20"
									    	, "Standard": "Unchecked"}
		, "Patienten mit mehreren Scheinen"	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button23"
									    	, "Standard": "Unchecked"}
		, "Patienten mit Quartalsquittung"   	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button24"
									    	, "Standard": "Unchecked"}
		, "Ringversuchszertifikate"                 	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button25"
									    	, "Standard": "Unchecked"}
		, "Aktive &HzV-/FaV-Teilnehmer"       	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button28"
									    	, "Standard": "Unchecked"}
		, "Teilabrechnung"                         	: {"Type":"Checkbox"
									    	, "ClassNN": 	"Button33"
									    	, "Standard": "Unchecked"}}

	AlbisActivate(1)
	hWin := AlbisFormular("Abrechnung_vorbereiten")

	Loop, Parse, Set, `n, `r
	{
			RegExMatch(A_LoopField, "^(.*)?\,(.*)", match)
			ButtonName	:= Trim(match1)
			Command   	:= Trim(match2)

			If (StrLen(ButtonName) = 0)
				continue

			If !Buttons.HasKey(ButtonName)
				throw Exception("You called an unknown button: '" ButtonName "'")

			If (Buttons[ButtonName].Type = "CheckBox") {
					If RegExMatch(Command, "i)Check")
						VerifiedCheck(Buttons[ButtonName].ClassNN, "ahk_id " hWin,,, 1)
					else
						VerifiedCheck(Buttons[ButtonName].ClassNN, "ahk_id " hWin,,, 0)
			}
			else If (Buttons[ButtonName].Type = "Button") {
				VerifiedClick(Buttons[ButtonName].ClassNN, hWin)
			}
			else If (Buttons[ButtonName].Type = "Edit") {
				VerifiedSetText(Buttons[ButtonName].ClassNN, TextValue, hWin)
			}

	}

return
}

AlbisBehandlungsliste(Options:="") {                                                                                          	;-- Privat/Behandlungsliste anzeigen

	; Dialog Behandlungsliste aufrufen
		hBhlg := Albismenu(32891, ["Behandlungsliste ahk_class #32770", "", "ALBIS ahk_class #32770", "Zugang verweigert"])
		If IsObject(hBhlg) {
			MsgBox, 1, Addendum für Albis on Windows, % "Sie haben nicht die Berechtigung`ndie Behandlungsliste anzuzeigen!", 6
			VerifiedClick("Button1", "ALBIS ahk_class #32770", "Zugang verweigert")
			return 0
		}

	; Steuerelement Felder anhand der übergebenen Optionen setzen
		doCount := 0, done := 0
		For do, what in Options {

			doCount ++
			switch do
			{
					case "Bearbeitung":
							done ++
							If (what = "alle_Ärzte")
								VerifiedCheck("Button2",,, hBhlg)
							else if (what = "Arztgruppe")
								VerifiedCheck("Button3",,, hBhlg)
					case "Arztname":
							done ++
							Control, ChooseString, % Options[Arztname], ListBox1, % "ahk_id " hBhlg
					case "Sortierung":
							done ++
							If (what = "Patienten-Nr.")
								VerifiedCheck("Button25",,, hBhlg)
							else if (what = "alphabetisch")
								VerifiedCheck("Button26",,, hBhlg)
					case "Gruppieren":
							done++
							If (what = "Status")
								VerifiedCheck("Button30",,, hBhlg)
							else if (what = "Arzt")
								VerifiedCheck("Button31",,, hBhlg)
							else if (what = "nicht gruppieren")
								VerifiedCheck("Button33",,, hBhlg)
					case "Behandlungsart":
							done++
							If InStr(what, "ambulant")
								VerifiedCheck("Button12",,, hBhlg, true) 	; Check
							else
								VerifiedCheck("Button12",,, hBhlg, false)	; Uncheck
							If InStr(what, "stationär")
								VerifiedCheck("Button13",,, hBhlg, true)
							else
								VerifiedCheck("Button13",,, hBhlg, false)
					case "Kassenart":
							done++
							If (what = "beide")
								VerifiedCheck("Button16",,, hBhlg)
							else if (what = "nur Kasse")
								VerifiedCheck("Button17",,, hBhlg)
							else if (what = "nur Privat")
								VerifiedCheck("Button18",,, hBhlg)
			}

		}

	; Dialog schliessen
		VerifiedClick("OK",,, hBhlg)
		WinWaitClose, Behandlungsliste ahk_class #32770,, 5

return ErrorLevel
}



;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; SONSTIGE AUTOMATISIERUNGSFUNKTIONEN                                                                                                                              	(15)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisWaitAndActivate            	(02) AlbisNeuStart                     	(03) AlbisCaveVonToolTip          	(04) AlbisHotKeyHilfe
; (05) AlbisIsBlocked                        	(06) AlbisCloseLastActivePopups   	(07) AlbisDefaultLogin               	(08) AlbisAutoLogin
; (09) AlbisLogout                       		(10) AlbisActivate                        	(11) AlbisCopyCut                     	(12) AlbisIsElevated
; (13) AlbisSelectAll                          	(14) CheckAISConnector         		(15) AlbisKeineChipkarte
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisWaitAndActivate(WinTitle, Debug=1, DbgWhwnd=0) {                                    	;-- ## könnte gelöscht werden, Fehlermeldungsfenster automatisch schliessen

	; Debug= 0 dann erfolgt keine Ausgabe in ein Listviewfenster, doch noch immer läßt sich der Errorlevel auslesen
	; (GNR-Vorschlag, Prüfung EBM/KRW, Überprüfung der Versichertendaten)
		If (Debug =1) {
            	Gui, %DbgWhwnd%:Default
            	LV_Add("", "      Warte auf Fenster: " . WinTitle)
		}

		WinActivate, %WinTitle%,,1
            WinWaitActive, %WinTitle%,,5

		while !WinExist(WinTitle) {

            		phwnd:= DLLCall("GetLastActivePopup", "uint", AlbisWinID())
            		WinGetTitle, WT, ahk_id %phwnd%

            		If InStr(WT, "GNR-Vorschlag") {
                        	WinActivate, %WT%
                        		VerifiedClick("Button1", "GNR-Vorschlag")
                                    Sleep, %sleeptime%
                                    If (Debug =1) {
                                    	LV_Add("", "Fenster: GNR Vorschlag - geschlossen")
                                    }
            		}

            		If InStr(WT, "ALBIS - [Prüfung EBM/KRW]") {
                        	WinActivate, %WT%
                        		SendInput, {ESC}
                                    Sleep, 100
                                    If (Debug =1) {
                                    	LV_Add("", "Fenster: Prüfung EBM/KRW geschlossen")
                                    }
                        }

            		If Instr(WT, "Überprüfung der Versichertendaten") {
                        		VerifiedClick("Button1", "Überprüfung der Versichertendaten")
                                    Sleep, 100
                                    If (Debug =1) {
                                    	LV_Add("", "Fenster: Überprüfung der Versichertendaten wurde ignoriert.")
                                    }
                        }

            		While !WinExist(WinTitle)
                        	MsgBox, 0x40000, Achtung, Das Fenster %WinTitle%	hat sich nicht geöffnet.`nöffne es bitte manuell und drücke dann ok.

            		sleep, 100
		}

		IfWinNotActive, %WinTitle%
            	WinActivate, %WinTitle%
                        WinWaitActive, %WinTitle%, , 3

return ErrorLevel
}
;02
AlbisNeuStart(CompName, User, Pass, CallingProcess, Auto=0) {                             	;-- starten von Albis oder Überprüfen ob sämtliche Dialoge geschlossen sind

	;**********************WICHTIG******************************
	; die Variablen AlbisWinID, AddendumDir werden von dieser Funktion benötigt
	; da viele Funktionen diese Variablen brauchen sollten sie in allen aufrufenden Skripten als "global" definiert sein

	;{ 1. Variablen und Einstellungen

	  ; globale Variablen / Thread Einstellungen
		SetTitleMatchMode, 2
		DetectHiddenWindows, On
		DetectHiddenText, On

	  ; schaut nochmal nach ob Albis läuft und holt sich das Fensterhandle
		AlbisWinID:= AlbisWinID()
		newrun:= 0                                                ;Variable die auf 1 gesetzt wird wenn Albis neu gestartet werden muss
		MDIWin:={}

	 ; Auslesen von wichtigen Daten aus der Addendum.ini
		IniRead, Comp, %AddendumDir%\Addendum.ini, Computer, %CompName%
		Loop, Parse, Comp, `|                        ; ermitteln von User und Pass mittels Parsen
			Login%A_Index% := A_LoopField

		IniRead, AlbisExeDir	 , % "AddendumDir" "\Addendum.ini", Albis, AlbisExeDir
		IniRead, AlbisExe   	 , % "AddendumDir" "\Addendum.ini", Albis, AlbisExe
		IniRead, AlbisWorkDir, % "AddendumDir" "\Addendum.ini", Albis, AlbisWorkDir
	;}

	;{ 2. Auto=0 oder 1
		If (Auto=0)                                                                                                             ;--wenn kein Autostart (Auto=0) , dann kommt hier eine Nachfrage ob Albis wirklich gestartet werden soll
		{
				SplashTextOn, 400, 50, Addendum für Albis on Windows, Starte Albis auf dem Computer %CompName%`nBenutze den Usernamen`: %User% für das Login.
				WinMove, Addendum, , A_ScreenWidth - 400, A_ScreenHeight - 120
				WinGet, SplashPraxID, ID, Addendum

				MsgBox, 4, Addendum für Albis on Windows,
				(LTrim
				Die Albis Praxissoftware ist nicht gestartet!
				%CallingProcess%  ist als Addon dafür konzipiert.
				Soll ich Albis jetzt starten?"
				)

				IfMsgBox, Yes
					Result := "Ja"
				else
					result := "Nein"
		}
		else if (Auto=1)                                                                                                 	;--bei gesetztem Autostart - setzt er nur das Ergebnis der Abfrage auf ja
				Result:="Ja"
	;}

	;{ 3. Albis Start- und Verifizierungsprozeduren, a. Albis starten wenn nicht vorhanden, b. Prüfen ob Login notwendig, wenn ja dann einloggen, c. Öffnen des Wartezimmerfenster, 4. Maximieren des Wartezimmerfensters
	If InStr(Result, "Ja") 	{
		user_result :=1

		;{a ------------- falls Albis nicht gestartet ist, dann wird es jetzt gestartet
            If !AlbisWinID()			{
					Msgbox,1, Addendum für ALBIS - %CallingProcess%, Albis muss gestartet sein`nDamit dieses Skript verwendet werden kann!, 10
					Run, %AlbisExeDir%\%AlbisExe%, %AlbisWorkDir%, UserErrorLevel, AlbisPID
					If ErrorLevel
						FileAppend, % TimeCode(1) "Function:  AlbisNeuStart ErrorLevel: " ErrorLevel "|Error: " A_LastError "|AlbisPID: " AlbisPID "`n", %AddendumDir%\logs'n'data\ErrorLogs\Errorbox.txt, UTF-8
					else
						newrun := 1
            }
		;}

		;{b ------------ wartet auf die Erstellung des Albisfenster ---
            WinWait, ALBIS ahk_class OptoAppClass

		 ; wichtig die FensterID von Albis
           	AlbisWinID:= AlbisWinID()
           	WinGetPos, ASx, ASy, ASw, ASh, % "ahk_id " AlbisWinID()

		 ; Mitteilung das Albis gestartet wurde
            if (Auto=0)
			{
					ControlSetText, static1,  Albis ist gestartet!`nWarte auf das Albis Login Fenster, Addendum
					Addy:= Asy + Ash, Addx:= Asx//2 - 200
					WinMove, Addendum, % Addx, % Addy
            }
            ;}

		;{c ------------- Albis is blocked? --- welches Fenster blockiert?
            	;prüfe ob das Login schon stattgefunden hat, wenn nicht führe den folgenden Code aus
            if AlbisIsBlocked(AlbisWinID()) || (newrun = 1)
			{
                        ;welches Fenster blockiert Albis? ---
            		phwnd:= DLLCall("GetLastActivePopup", "uint", AlbisWinID())
            		WinGetTitle, title, % "ahk_id " phwnd
            		MsgBox, % title
            Pause
            		WinWait, ALBIS - Login ahk_class #32770
                        WinActivate, ALBIS - Login ahk_class #32770

                        ;Benachrichtigung
            		if (Auto=0)
                        ControlSetText, static1, gebe Logindaten ein, starte Albis

                        ;das Loginfenster ist da - greift sich die ID ab um Nutzer und Passwort ohne SendInput in die Felder einzutragen
            		If (hLogin:= WinExist("ALBIS - Login ahk_class #32770"))
					{
                        	ControlSetText, Edit1, %Login1%, % "ahk_id " hLogin
                        		sleep 100
                        	ControlSetText, Edit2, %Login2%, % "ahk_id" hLogin
                        		sleep 100
                        	VerifiedClick("Button1", "", "", hLogin)		;schliessen des Loginfenster

                        	if (Auto=0)
                        		SplashTextOn, 400, 50, Addendum für Albis on Windows, In Albis eingeloggt.`nÖffne die Wartezimmerliste.
            		}
            }
		;}

		;{d -------------- Wartezimmer öffnen --- falls noch nicht geöffnet
          	;Benachrichtigung
				if (Auto=0)
					ControlSetText, static1, Öffne das Wartezimmer, starte Albis

           	;nach dem Login eine kurze Pause einräumen
				sleep 15000

			;Vorbereitungen für das Öffnen des Wartezimmer
				WinActivate, ahk_class OptoAppClass
            	PostMessage, 0x111, 32924,,, % "ahk_id " AlbisWinID()
				Parent:={WinID: AlbisWinID(), WinTitle: "", WinClass: ""}
				;Parent:={WinID: "", WinTitle: "ALBIS", WinClass: "OptoAppClass"}
				i:=0, WartezimmerID:=0

			;durchläuft den Loop solange bis das Wartezimmer geöffnet wurde
				While !(WartezimmerID:= FindChildWindow(Parent, "Wartezimmer", "On"))         ;das Wartezimmer wird per WM_Command aufrufen - Patient/Wartezimmer = 32924
				{
						PostMessage, 0x111, 32924,,, % "ahk_id " AlbisWinID()
						sleep 300
						i++
						if i>20
                            break
				}

           	;wurde zu lange gesucht, gab es wahrscheinlich ein Problem, deshalb bricht die Funktion hier ab
				If i = 20
				{
						MsgBox, 1, Achtung, Konnte das Wartezimmer nicht öffnen!, 5
						goto CloseSplashWindow
				}

            ;wenn das Wartezimmerfenster nicht maximiert ist dann maximiere es jetzt
            WinMaximize, % "ahk_id " WartezimmerID
		;}

	}
	else if InStr(Result, "Nein")
            user_result := 0
	;}

		goto SAExit

CloseSplashWindow: ;{
If WinExist("Addendum für Albis")
		SetTimer, SplashOff, 1000
return

SplashOff:
	SplashTextOff
	ToolTip, , , , %TTnum%
	SetTimer, SplashOff, Off
return
;}

SAExit:
return user_result
}
;03
AlbisCaveVonToolTip(compname, CaveVonID) {                                                     	;-- zeigt Infos an was man neues im Cave Fenster machen kann

		SetTitleMatchMode, 2
		CoordMode, Pixel, Screen
		CoordMode, ToolTip, Screen

;{GUI

	Global ImpfLogo1, ImpfLogo2, ImpfLogo3, CTTExist
	static BuchstabeK


	CaveTT=
(
Lieber Anwender,

Achte bei diesem Fenster darauf was Du wie einschreibst!
Die Zeilen haben in unserer Praxis eine bestimmte Bedeutung!
eine kleine Anleitung folgt jetzt:

1: z, DMP Untersuchung
2: [unbenutzt im Moment]
3: Impfungen - folgende Schreibweisen sind erlaubt:
     Td, TdPP, TdPert, Po, Pneu, MMR, MMR-V, MenC, GSI, Vari,
	 FSME, VZV, Röteln, Mumps, VChol
     Bsp: TdPP-w 17 (Wiederholung),
            TdPP-2 03/16 (2.Impfung März 16)
4: [auch Impfungen wenn Zeile 3 nicht reicht]
5: [unbenutzt im Moment]
6: [unbenutzt im Moment]
7: [unbenutzt im Moment]
8: RöntgenThorax, EKG , Lufu
       RöThorax 02/15, EKG 01/11, Lufu 06/16
9: Gastro, Colo, Vorsorgeuntersuchung, Geriatrischer Basiskomplex
       01740, (C08 K 18, G14 iP K 15), GVU/HKS 04/15, GB 18.1A
       "Coloskopie 2008 Kontrolle 2008, Gastroskopie 2014
	   intestinale Metaplasie Kontrolle 2015, erklärt sich selbst,
       GB 18.1A = Quartal I A=03360 oder B=03362)
10: wichtiges
	Chemopatient, anderweitige Kontrollen notwendig
)

ImpfText1 = Erinnerung an Impfungen!
ImpfText2 =
(
Bei diesem Patienten sind noch ⋘KEINE IMPFUNGEN⋙ eingetragen!

Bitte frage ihn danach oder errinnere ihn an seinen Impfausweis!
)


	NewX:= CaveVon[compname].x
	NewY:= CaveVon[compname].y
	NewW:= CaveVon[compname].w
	NewH:= CaveVon[compname].h

	WinGetPos, CvX, CvY, CvW, CvH, ahk_id %CaveVonID%
	;wird das Fenster nicht gefunden, soll er die Funktion verlassen
		If (CvX="") {
                        return
                        		}

	;wenn die Größe nicht passt. Wird die Größe angepasst. Auf manchen Rechnern fehlt immer die letzte Zeile.
	If (CvX <> NewX OR CvY <> NewY OR CvW <> NewW OR CvH <> NewH) {
		WinMove, ahk_id %CaveVonID%,, %NewX%, %NewY%, %NewW%, %NewH%
	}

	;ist das Gui da sorge dafür das es oberhalb des cave von Fenster bleibt
	if (CTTExist=1) {

		WinSet, AlwaysOnTop, On, ahk_id %hImpf%
            return

	} else {

		CTTExist=1
		CaveX:= CvX + NewW + 10
		;MsgBox, X: %X% - cx: %CaveX% --  Y: %Y%

		zeile3:= AlbisGetCaveZeile(3, "", 0) ; 3.Zeile auslesen, keine Ausgabe in ein Debugfenster (0, 0) und Fenster nicht schließen
		ToolTip, %CaveTT%`n%zeile3%, %CaveX%, %CvY%, 10

		If (zeile3 = "") {

		;Transparente Hintergrund Gui
            Gui, IMPFb: New, +HWNDhImpfb 0x94000000 Ex0x00180000 -SysMenu -Caption
            Gui, IMpfb: +Owner
            Gui, IMPFb: Color, c172842
            Gui, IMPFb: +Disabled
            WinSet, Transparent, 220, ahk_id %hImpfb%

		;Textoverlay Gui
            Gui, IMPF: New, +HWNDhImpf -SysMenu -Caption +AlwaysOnTop ;0x54020000 Ex0x00080028
            Gui, IMPF: Color, c000010 ;c3F627F                  %ColRow%
            Gui, IMPF: Margin, 20, 20
            Gui, IMPF: Add, GroupBox, x10 y5 w710 h80

            Gui, IMPF: Font, s32 q5 c00FF00 cWhite, Futura Bk Md
            Gui, IMPF: Add, Text, xm ym w690 h56 Center, %ImpfText1%

            Gui, IMPF: Font, s16 q5 c00FF00 cWhite, Futura Bk Bt
            Gui, IMPF: Add, GroupBox, x10 y90 w710 h120
            Gui, IMPF: Add, Text, xm y115 w680 h80 Center, %ImpfText2%

            Gui, IMPF: Add, Picture, x230 y240 w148 h148 BackGroundTrans vImpfLogo1 gLogoClick, %logo1%
            Gui, IMPF: Add, Picture, x230 y240 w148 h148 BackGroundTrans HIDE vImpfLogo2 gLogoClick, %logo2%
            Gui, IMPF: Add, Picture, x230 y240 w148 h148 BackGroundTrans HIDE vImpfLogo3 gLogoClick, %logo3%

            Gui, IMPF: Font, s146 q5 c00FF00 cWhite, Futura Bk Md
            Gui, IMPF: Add, Text, x380 y196 BackGroundTrans vBuchstabeK gLogoClick , K

            WinSet, AlwaysOnTop, On, ahk_id %hImpf%
            	WinSet, TransColor, Off, ahk_id %hImpf%
            		WinSet, TransColor, 000010, ahk_id %hImpf%

            Gui, IMPFb: Show, % "x" NewX " y" NewY " w" NewW " h" NewH, Impferrinnerungback
            Gui, IMPF: Show, % "x" NewX " y" NewY " w" NewW " h" NewH, Impferrinnerung            ; "x" NewX " y" NewY " w" NewW " h" NewH

            GuiControl, +Redraw, ImpfLogo1

            	OnMessage(0x200, "MMOVE_CVTT")
            WinWaitClose, ahk_id %hImpf%
            	OnMessage(0x200, "")
            	CTTExist:=0
            	ToolTip,,,,10
		}

	}

return
;}

LogoClick: ;{
	CName:= A_GuiControl
	Gui, Submit, nohide

	If (Instr(CName, "ImpfLogo")) {

            OnMessage(0x200, "")
            Gui, IMPF:Default
            guicontrol, Hide, ImpfLogo1
            guicontrol, Show, ImpfLogo2
            guicontrol, Hide, ImpfLogo3

            sleep, 300
            Gui, IMPF: Destroy
            Gui, IMPFb: Destroy

	}

return
}

;}
MMOVE_CVTT() {                                                                                                  	;-- gehört zu AlbisCaveVonToolTip

	CName:= A_GuiControl
	MouseGetPos,,,, GCtrl, 2
	;ToolTip %GCtrl%

	If (Instr(CName , "ImpfLogo")) {                        		;(Instr(CName , "ImpfLogo"))

            GuiControl, Hide, ImpfLogo1
            GuiControl, Hide, ImpfLogo2
            GuiControl, Show, ImpfLogo3
            GuiControl, +Redraw, ImpfLogo3


	} else if !Instr(CName , "ImpfLogo") {

            GuiControl, Show, ImpfLogo1
            GuiControl, +Redraw, ImpfLogo1
            GuiControl, Hide, ImpfLogo2
            GuiControl, Hide, ImpfLogo3
	}

}
;04
AlbisHotKeyHilfe(AddendumHelp:="", PraxomatHelp:="") {                                       	;-- blendet die aktuellen zusätzlichen Hotkeys in der Albisstatusbar ein

	; letzte Änderung: 07.02.2021

	static AlbisStatusbarClass := "Afx:00400000:0:00010003:00000010:000000001"

	If !WinActive("ahk_class OptoAppClass")
		return

	ControlGetText, AlbisKeyHelp, % AlbisStatusbarClass

	If PraxomatHelp && !Instr(AlbisKeyHelp, PraxomatHelp)
		ControlSetText, % AlbisStatusbarClass, % AlbisKeyHelp "  " PraxomatHelp
	else if AddendumHelp && !InStr(AlbisKeyHelp, AddendumHelp)
		ControlSetText, % AlbisStatusbarClass, % AlbisKeyHelp "  " AddendumHelp

}
;05
AlbisIsBlocked(AlbisWinID:=0, autoclose:=2) {	                                                      	;-- befreit Albis von blockierenden Fenstern

	; stellt fest ob Albis durch ein ChildWindow blockiert ist und wie dieses heißt, kann dieses auch sofort schließen, autoclose ist default

/* 		BESCHREIBUNG

		#### AlbisIsBlocked() Funktion : die Parameterliste
		die Funktion bietet 3 Möglichkeiten
		1. autoclose = 0
            ausschließlich Rückgabe eines Objektes mit zwei Key:Value Paaren, a) true/false für ist blockiert und b) Name des blockierenden Fensters
		2. autoclose = 1
            Rückgabe des oben genannten Objektes und Rückfrage an den User ob die blockierenden Fenster geschlossen werden dürfen
        3. autoclose = 2
            alle blockierenden Fenster werden ohne Rückfrage geschlossen, es werden keine Werte zurück gegeben

 		von der Funktion AlbisCloseLastActivePopups() bekommt sie den Rückgabewert (errorStr), den sie an die aufrufende Prozeß weiterreicht
		dieser String gibt den Erfolg oder Mißerfolg zurück , damit der Prozeß Fehler erkennen kann um sich wenn notwendig zu beenden oder
        andere Maßnahmen einzuleiten die zum Erfolg führen könnten

		letzte Änderung: 	11.06.2019	- Fehlerkorrektur bei WinGet, Ausbrechen aus dem AlbisClosePopups-Loop durch eine zusätzliche Bedingung möglich, schätze das sich hier manchmal unendliche Loops ergeben haben

*/


		If !AlbisWinID
			AlbisWinID := AlbisWinID()

		If (autoclose = 0)   		{
			WinGetTitle, ltitle, % "ahk_id " phwnd:= DLLCall("GetLastActivePopup", "uint", AlbisWinID:= AlbisWinID())
			return {"isblocked": WinIsBlocked(AlbisWinID()), "blockWinT": ltitle, "errorStr": errorStr}
		}
		else if (autoclose = 1)	{
				;platzhalter erstmal
		}
		else if (autoclose = 2)	{
			If WinIsBlocked( AlbisWinID:= AlbisWinID() )
				AlbisClosePopups()
		}

return
}
AlbisClosePopups() {                                                                                               	;-- schließt alle PopUp-Fenster in Albis

		AlbisActivate(1)
		AlbisWinID := AlbisWinID()

		Loop		{

		  ; welches Fenster blockiert Albis? ---
			phwnd		:= GetHex(DllCall("GetLastActivePopup", "uint", AlbisWinID()))
			If InStr(phwnd, AlbisWinID())
				break

		  ; 3 verschiedene Versuche das Fenster zu schließen von "sanft" bis "kraftvoll"
			SendMessage, 0x112, 0xF060,,, % "ahk_id " phwnd  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
			t.= " EL1: " ErrorLevel "`n"
			If ErrorLevel	{
				WinClose, % "ahk_id " phwnd
				t .= " EL2: " ErrorLevel "`n"
				If ErrorLevel	{
					WinKill, % "ahk_id " phwnd
					t .= " EL3: " ErrorLevel "`n"
					If ErrorLevel	{
						WinGetTitle, ltitle, % "ahk_id " phwnd
						throw Exception("Can not close entire window : '" . ltitle . "' `n. The function AlbisCloseLastActivePopups stops here.", -1)
						errorStr := "noClose"
						break
					}
				}
			}

			sleep 150

		; Loop endet hier wenn Albis nicht mehr blockiert ist
			If !WinIsBlocked(AlbisWinID())
				break
		}

return
}
;06
AlbisCloseLastActivePopups(AlbisWinID:="") {                                                            	;-- schließt alle PopUp-Fenster bis keines mehr da ist

	;diese Funktion wird benötigt wenn ein Skript

	If !WinIsBlocked(AlbisWinID := AlbisWinID())
		return

	Loop	{

		; mehr als 10 zu schließende Dialoge gibt es nicht (Sicherheits return , damit die Schleife irgendwann verlassen wird)
			If (A_Index > 9)
				return 2         	; 2 = Problem beim Schliessen

		; ist fertig wenn Popup handle und Albis handle identisch sind
			phwnd := GetHex(DLLCall("GetLastActivePopup", "uint", AlbisWinID()))
			If (pHwnd = AlbisWinID())
					return 1

		;3 verschiedene Versuche das Fenster zu schließen von "sanft" bis "kraftvoll"
			SendMessage, 0x112, 0xF060,,, % "ahk_id " phwnd  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
			ELevel := ErrorLevel

		; schießt manchmal versehentlich über das Ziel hinaus und will Albis schließen, der "Hiermit beenden Sie Albis" Dialog wird erkannt und
		; wieder geschlossen, die Schleife beendet
			If WinExist("ALBIS", "Hiermit beenden Sie Albis")			{
				VerifiedClick("Button2", "ALBIS", "Hiermit beenden Sie Albis")
				return 1
			}

		; so neuer Versuch
			If ELevel {

				WinClose, % "ahk_id " phwnd
				If ErrorLevel	{
					WinKill, % "ahk_id " phwnd
					If ErrorLevel {
						WinGetTitle, ltitle, % "ahk_id " pHwnd
                        throw Exception("Can not close entire window : '" ltitle  "'`n. The function AlbisCloseLastActivePopups stops here.", -1)
                        return "Das Schließen geöffneter Dialoge war nicht möglich!"
					}
				}

			}


		; alle Popupfenster sind geschlossen, dann zurück
			If (AlbisWinID() = GetHex(DLLCall("GetLastActivePopup", "uint", AlbisWinID())))
				return 1

			sleep 200
	}

return 1
}
;07
AlbisDefaultLogin(key) {                                                                                         	;-- liest die Default User und Passwordeinstellung aus und gibt ein Array zurück (als Class wäre das auch nicht schlecht)

	static DLogIn

	If !IsObject(DLogIn) {
		DLogIn:= Object()
		IniRead, Comp, %AddendumDir%\Addendum.ini, Computer, % compname
		;MsgBox, % compname ", " Comp
		Utmp:= StrSplit(Comp, "|")
		DLogin:= {"User": (Utmp[1]), "Password": (Utmp[2])}
	}

return DLogin[key]
}
;08
AlbisAutoLogin(MsgBoxZeit:=10) {                                                                            	;-- loggt den jeweiligen Nutzer automatisch ein

		; letzte Änderung: 10.09.2020 - Vorgang beschleunigt
		; der Click mittels einer anderen Technik hat das Login-Fenster nicht geschlossen und den Hook für die Erkennung des Login-Fenster erneut ausgelöst

		static LoginTime	:= 0
		static LoginNo   	:= false

	;wenn ein automatischer Login verneint kann der nächste automatische Loginvorgang erst in 60 Sekunden erfolgen
		If LoginNo && ((A_TickCount - LoginTime) > 60000)
			LoginNo := false

	;Routine zum Einloggen, nach drücken auf Nein erkennt der Hook das Loginfenster erneut, deswegen muss eine erneute Abfrage unterdrückt werden
		If !LoginNo && WinExist("ALBIS - Login ahk_class #32770")		{

			MsgBox, 4, Addendum für AlbisOnWindows, Soll das automatische Einloggen`ndurchgeführt werden?, % MsgBoxZeit
			IfMsgBox, No
			{
					LoginNo 	:= true
					LoginTime := A_TickCount
					return
			}

			If VerifiedSetText("Edit1", AlbisDefaultLogin("User"), "ALBIS - Login ahk_class #32770", 100)
				If VerifiedSetText("Edit2", AlbisDefaultLogin("Password"), "ALBIS - Login ahk_class #32770", 100)	{
						VerifiedClick("Ok", "ALBIS - Login ahk_class #32770")
						sleep 100
						;~ c:= GetWindowSpot((hcontrol := ControlGet("HWND", "", "Button1", "ALBIS - Login ahk_class #32770")))
						;~ MouseMove, % c.X +5 , % c.Y + 5
						;~ MouseClick, Left
				}

		; Backupvorgang wenn VerifiedClick versagt
			If (hcontrol:= ControlGet("HWND", "", "Button1", "ALBIS - Login ahk_class #32770"))	{
					c:= GetWindowSpot(hcontrol)
					MouseMove, % c.X +5 , % c.Y + 5
					MouseClick, Left
			}

			AlbisWZOeffnen()
		}

return
}
;09
AlbisLogout() {                                                                                                      	;-- loggt den Nutzer einfach aus
	Albismenu(32922)
}
;10
AlbisActivate(waitingtime) {                                                                                       	;-- aktiviert das Albishauptfenster
	WinActivate		, % "ahk_id " (AlbisWinID := WinExist("ahk_class OptoAppClass"))
	WinWaitActive	, % "ahk_id " AlbisWinID,, % waitingtime
return ErrorLevel
}
;11
AlbisCopyCut() {                                                                                                    	;-- speichert alles was mit Strg+c oder Strg+x in Albis kopiert wird in eine Datei (#überarbeiten#))

		static AlbisCText_old


		while (AlbisCText = "") 	{

				ControlGetFocus, CFocus, % "ahk_class OptoAppClass"
				If ErrorLevel
						CGF := "#no control focus "

				ControlGet, AlbisCText, Selected,, % CFocus, % "ahk_class OptoAppClass"
				If ErrorLevel
						CGF .= "#, error retrieving text from control " CFocus

				Sleep, 100
				If (A_Index > 3)
					break

		}

		If (StrLen(CGF) > 1)
			FileAppend, % CGF "`n", % AddendumDir "\logs'n'data\CopyAndPaste.log"
		else if (AlbisCText = AlbisCText_old)
			AlbisCText     	:= ""
		else 		{
			AlbisCText_old	:= AlbisCText
			FileAppend, % AlbisCText "`n", % AddendumDir "\logs'n'data\CopyAndPaste.log"
		}

return AlbisCText
}
;12
AlbisIsElevated() {                                                                                              		;-- stellt fest ob Albis mit UAC Virtualisierung gestartet wurde

	Process, Exist, albisCS.exe
	If !AlbisPID:= ErrorLevel
	{
			Process, Exist, albis.exe
			albisPID := ErrorLevel
	}

	If AlbisPID
		If IsProcessElevated(AlbisPID)
				return 1
		else
				return 0

return 2 ; - means Albis is not running
}
;13
AlbisSelectAll() {                                                                                                     	;-- kompletten Text markieren

	Send, {PgUp}
	sleep, 100
	Send, {LControl Down}{LShift Down}{Down}
	sleep, 100
	Send, {LControl Up}
	sleep, 50
	Send, {LShift Up}

}
;14
CheckAISConnector() {                                                                                				;-- sieht nach ob der AIS Connector (Laborverbindungsprogramm) läuft und startet es bei Bedarf neu

	; durch Verwendung von A_Username müsste diese Funktion auf allen Computern unter allen Benutzernamen funktionieren

	for Prozess in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	   prozesse .= Prozess.Name "|"

	If !Instr(prozesse, "AISConnector") {

		TrayTip, % "Hinweis!", % "Der AIS Connector ist`nnicht mehr verfügbar`nund wird neu gestartet!", 4
		run, % "C:\Users\" A_UserName "\Desktop\AIS Connector.appref-ms"
		If ErrorLevel {
			TrayTip, Fehler!, % "Der AIS Connector `nließ sich nicht starten!`nVersuche es später nocheinmal!", 4
			return 0
		}
		else
			return 1
	}

return 1
}
;15
AlbisKeineChipkarte(Button:="Ja") {                                                                        	;-- Dialogfenster: 'Patient hat in diesem Quartal seine Chipkarte...' schließen
	If WinExist("ALBIS ahk_class #32770", "Patient hat in diesem Quartal")
		VerifiedClick((Button="Ja" ? "Button1":"Button2"), "ALBIS ahk_class #32770", "Patient hat in diesem Quartal",, true)
}
;}
;-------------------------------------------------------- Hilfsfunktionen für die Albisfunktionen -------------------------------------------------------------------------------
; (01) WMIEnumProcessExist            	(02) ExtractNamesFromString     	(03) FormatedFileCreationTime	    	(04) ObjFindValue
; (05) hk 	                                       	(06) isActiveFocus
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;1
WMIEnumProcessExist(ProcessSearched) {                                                            		;-- logische Funktion zur Suche ob ein bestimmter Prozeß existiert

    ;Funktion gibt nur 0 - für nicht vorhanden und 1 für vorhanden aus (als schnell Variante gedacht)
	;zudem ist der Funktion die exakte Schreibweise des Prozeß egal (GROSS/klein oder nur ein Teil davon)

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	   liste .= process.Name "`n"

	Sort, liste, U

	Loop, parse, liste, `n
		If Instr(A_LoopField, ProcessSearched, false ,1)
			return 1

	return 0
}
;2
ExtractNamesFromString(Str) {                                                                        			;-- die Namen müssen die ersten zwei Worte im String sein

	; im Parameter Str muss Wort1 und Wort2 den Patientennamen enthalten

	Name := Array()

	;Ermitteln des Patientennamen aus dem PDF-Dateinamen
	Pat := StrReplace(Str, ",", " ")
	;zuviele Space Zeichen hintereinander machen Probleme
	Pat := RegExReplace(Pat, "\s{2,}", " ")
	Pat := StrSplit(Pat, " ")
	Name[1] := Pat[1] ", " Pat[2]
	Name[2] := Pat[2] ", " Pat[1]
	Name[3] := Pat[1]
	Name[4] := Pat[2]

return Name
}
;3
FormatedFileCreationTime(filepath) {                                                                     	;-- formatiert ins deutsche Format 18.05.2019
	;bei leerem creationtime String gibt FormatTime automatisch das Datum von heute aus
	FileGetTime, creationtime, % filepath, C
	FormatTime, creationtime, % creationtime, dd.MM.yyyy
return creationtime
}
;4
ObjFindValue(arr, value) {                                                                                     	;-- Array in Array Suche

	; findet den Wert in mehreren key:value Objekten, welche in einem indizierten Array enthalten sind, gibt den Index key zurück

	For OFVIndex, obj in arr
		For key, val in arr[OFVIndex]
			If InStr(val, value)
                return OFVIndex

return 0
}
;5
hk(keyboard:=0, mouse:=0, message:="", timeout:=3) {                                          	;-- Tastatur- und/oder Mauseingriffe abschalten

	;https://www.autohotkey.com/boards/viewtopic.php?f=6&t=33925
	;!F1::hk(1,1,"Keyboard keys and mouse buttons disabled!`nPress Alt+F2 to enable")   ; Disable all keyboard keys and mouse buttons
	;!F2::hk(0,0,"Keyboard keys and mouse buttons restored!")         ; Enable all keyboard keys and mouse buttons
	;!F3::hk(1,0,"Keyboard keys disabled!`nPress Alt+F2 to enable")   ; Disable all keyboard keys (but not mouse buttons)
	;!F4::hk(0,1,"Mouse buttons disabled!`nPress Alt+F2 to enable")   ; Disable all mouse buttons (but not keyboard keys)

   static AllKeys
   static optKeyboard, optMouse

   if !AllKeys {
      s := "||NumpadEnter|Home|End|PgUp|PgDn|Left|Right|Up|Down|Del|Ins|"
      Loop, 254
         k := GetKeyName(Format("VK{:0X}", A_Index))
       , s .= InStr(s, "|" k "|") ? "" : k "|"
      For k,v in {Control:"Ctrl",Escape:"Esc"}
         AllKeys := StrReplace(s, k, v)
      AllKeys := StrSplit(Trim(AllKeys, "|"), "|")
   }
   ;------------------
   For k,v in AllKeys {
      IsMouseButton := Instr(v, "Wheel") || Instr(v, "Button")
      Hotkey, *%v%, Block_Input, % (keyboard && !IsMouseButton) || (mouse && IsMouseButton) ? "On" : "Off"
   }
   if (StrLen(message) > 0) {
      Progress, B1 M FS12 ZH0, %message%
	  optKeyboard	:= keyboard
	  optMouse    	:= mouse
      SetTimer, HkTimeoutTimer, % -1000*timeout
   }
   else
      Progress, Off

Block_Input:
Return

hkTimeoutTimer:

   Progress, Off
   If (optKeyboard=1) || (optMouse=1)
		hk(0, 0, "Die Sperrung des Nutzereingriffes ist aufgrund des übergebenen Zeitintervalles aufgehoben worden!" )

Return
}
;6
isActiveFocus(conditions) {                                                                                        	;-- Hotkey/Hotstring #If Bedingung Funktion

	; negative Vergleiche sind noch nicht möglich

	If !WinActive("ahk_class OptoAppClass")
		return false

	If StrLen(conditions) = 0
		Conditions := "contraction=lp|bg, activeWinType=Privatabrechnung"

	RegExReplace(conditions, "[+\-]", "", conditionsMax)
	conditionsMax	+= 1
	condMatched 	:= 0
	conditions     	:= StrSplit(conditions, ",")
	;conditionsMax	:= conditions.MaxIndex()
	For condNr, condition in conditions {

		RegExMatch(condition, "^\s*(?<lc>[+\-])*\s*(?<1>\w+)?\=(?<2>[A-Za-z\|]+)\s*$", param)
		If (param1 = "contraction")
			If RegExMatch(AlbisGetActiveControl("contraction"), param2)
				condMatched ++
		else if (param1 = "activeWinType")
			If InStr(AlbisGetActiveWindowType(), param2)
				condMatched ++

		if (condMatched = conditionsMax)
			return true

	}

return (condMatched = conditionsMax ? true : false)
}
;{ Hilfsfunktionen Ausfuellhilfe
ControlFilter(CtrlObj, classnn) { 	    	; Hilfsfunktion für AlbisAusfuellhilfe - gibt Steuerelementbezeichnung zurück

	; überprüft im Moment nur Edit Steuerelement

	For ctrlLabel, class in CtrlObj
		Loop, % class.Count()
			If (classnn = "Edit" class[A_Index])
				return ctrlLabel

return false
}

DatesInArr(Dates, Arr) {	                     	; Hilfsfunktion für AlbisAusfuellhilfe

	For idx, val in Arr
		If (val = Dates)
			return true

return false
}

IndexInArr(s, Arr) {	                        	; Hilfsfunktion für AlbisAusfuellhilfe

	For idx, val in Arr
		If (val = s)
			return idx

return 0
}

;}

;}


