;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                         	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                       	!diese Bibliothek wird von fast allen Skripten benötigt!
;                       	by Ixiko started in September 2017 - letzte Änderung 29.09.2022 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
return

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                	BESCHREIBUNG DER FUNKTIONSBIBLIOTHEK ZUR RPA SOFTWARE -- ADDENDUM FÜR ALBIS on WINDOWS --
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;
;       FUNKTIONSDESIGN:
;
;       Die Funktionen sollen eine einfach durchzuführende und möglichst fehlerfreie RPA (robotic process automation) ermöglichen.
;   	Nachfolgend beschreibe ich das dafür verwendete Design:
;
;        - Datenerfassung                   	 -	erfolgt über  das Auslesen des Albisfenstertitel,  des Inhaltes eines Fensters oder  durch Ermittlung eines be-
;                                                          	stehenden Eingabefocus
;
;        - das Albis Fensterhandle       	 -	wird bei  jedem Aufruf neu ermittelt,  damit ein neues Albishandle,  z.B. nach einem  Absturz oder Neustart
;                                                     		von Albis erkannt
;
;        - wenige Parameter		         	 -	viele Parameter = viele Variablen = viel mehr Möglichkeiten?  Möglicherweise ja, aber auch eine deutlich
;                                                         	höhere Wahrscheinlichkeit durch einen Tippfehler Probleme mit seinem Programm zu bekommen
;
;        - keine Wartezeit                    	 - 	nachdem eine Automatisierungsfunktion beendet ist,  ist kein sleep Befehl  notwendig um auf Fenster oder
;          nach Funktionsende                  	anderes zu warten. Die nächsten Befehle können unmittelbar erfolgen.
;                                                      	    Es ist nicht unbedingt notwendig die Dauer der Warte-Befehle zu ändern! Ich nutze schnelle und langsame
;                                                        	Computer in der Praxis.    Im Laufe der Jahre habe ich die Zeiten so optimiert das bei gleichbleibender Zu-
;                                                    	  	verlässigkeit für alle Computer selten eine übermäßig lange Wartezeit entsteht.    Es wird eher so sein, daß
;                                                         	die Geschwindigkeit weit über der menschlichen, visuellen Erfassungsgeschwindigkeit liegt.
;
;        - Ausführungsüberprüfung       	- 	Funktionen die Menupunkte  oder Fenster in Albis öffnen,  überprüfen ob Albis im durch geöffnete Popup-
;                                                         	Fenster blockiert ist und schliessen diese automatisch.   Die meisten dieser Funktionen prüfen auch ob das
;                                                          	angeforderte  Dialogfenster geöffnet  wurde.  Damit wird  verhindert das nachfolgende Automatisierungs-
;                                                          	funktionen  ins  Leere  arbeiten  oder  durch  wiederholte   Abfolgen  von  Befehlen  den  Zugriff auf  Albis
;                                                         	blockieren.    Erreicht  wird dies z.B. dadurch  das die Daten  die  von einem Skript  nach Albis übertragen
;                                                           wurden, sofort wieder ausgelesen und mit der Eingabe verglichen werden.    So lassen sich die gemachten
;                                                         	Änderungen zuverlässiger kontrollieren.
;
;        - Fensterhandling                    	-	Fehlerausgaben oder Hinweisfenster werden ausgewertet und es wird entsprechend reagiert
;
;        - Ereigniserkennung -               - 	ein großer Teil der Funktionalität von Addendum beruht auf dem Einsatz sogenannter Hooks. Hooks sind
;                                                    	  	Rückruffunktionen  oder Callbacks des Betriebssystems.    Windows erlaubt  sich in den Nachrichtenstrom
;                                                           von und zu oder innerhalb anderer Prozesse "einzuhaken" um dort mitlesen zu können.       Diese Technik
;                                                      	  	minimiert die CPU-Belastung erheblich  und ermöglicht außerdem  sofort auf Veränderungen z.B. in Albis
;                                                   	    zu reagieren.   Weiterhin  läßt sich mit Hilfe dieser  Technik  eine wesentlich  flexiblere Reaktion auf unter-
;                                                           schiedliche Ereignisse realisieren.
;
;        - Interaktion mit der       			- 	die zwar relativ einfach einzusetzende , aber sehr unzuverlässige Simulation  von Tasten- oder Mausein-
;       	Oberfläche                            	gaben von Autohotkey wird 	in den Automatisierungsfunktionen, soweit es möglich ist, nicht eingesetzt.
;                                                          	Der Aufruf z.B. des Kassenrezept-Formulares erfolgt nicht durch Senden von Tastaturkürzeln, sondern er-
;                                                        	folgt über das Senden von Nachrichten-ID's (siehe Send- oder Postmessage)

;                                                    	  	Eine weitere Technik bei RPA-Software  ist die Verwendung von Pixelsuch- oder Bildvergleichsfunktionen
;                                                         	(z.B. Sikuli). Auch diese Technik wird von Addendum nicht eingesetzt.  Für eine neuere Technik (Optical
;                                                      	  	Character Recognition),   d.h. Texte oder Textbereiche wie bei einem Textscanner zu erkennen,  fehlt es
;                                                      	  	Autohotkey an Geschwindigkeit.
;                                                      	  	Die Techniken sind nicht notwendig,  da die Albisoberfläche über Windowsfunktionen gezeichnet wird.
;                                                      	  	Microsoft hat gute Funktionen für den Zugriff auf Oberflächen bereitgestellt  und Autohotkey ist genau
;                                                      	  	darauf spezialisiert. Eine Interaktion ist mit fast allen Elementen der Albisoberfläche schon nach kurzer
;                                                      	  	Einarbeitungszeit in diese Skriptsprache möglich.
;
;	-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -    -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
;
;																																												                  				IXIKO 2021
;
;	-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  - -  -  -  -  -  -  -  -  -  -  -  -  -  - ;}

Init_Albis() {                                                                                       	;-- initialisiert wichtige Konstanten

	; manche Konstanten werden von verschiedenen Funktionen genutzt.
	; ändert die Compugroup die Klassenbezeichnungen von Steuerelementen muss die Zuordnung nicht in allen Funktionen korrigiert werden

	static _ := Init_Albis()           ; static wird zur Ladezeit ausgeführt und ruft sich selbst auf

	global KK            	:= {"Arzt":"Edit4", "Datum":"Edit5", "Kürzel":"Edit6", "Kuerzel":"Edit6", "Inhalt":"RichEdit20A1"}
	global KKx           	:= {"Arzt":"Edit1", "Datum":"Edit2", "Kürzel":"Edit3", "Inhalt":"RichEdit20A1"}	; bei Verwendung des direkten Elternfensters (#327702)
	global KKrv       	:= {"Edit4":"Arzt", "Edit5":"Datum", "Edit6":"Kuerzel", "RichEdit20A1":"Inhalt"}	; reversed
	global afxMDI   	:= "AfxMDIFrame140"
	global afxView   	:= "AfxFrameOrView140"
	global afxKKarte	:= "#327702"
	global IdentData	:= ["Dokument", "Edit(\d+),Edit(\d+),Edit(\d+),RichEdit20A\d+", "#32770", "AfxFrameOrView", "AfxMDIFrame"]
	global cparents 	:= {"Karteikarte"            	: {"parents" 	: "#32770\d+,AfxFrameOrView140\d+,AfxMDIFrame140\d+,"
																							. "Afx:00400000:b:0001000\d+,MDIClient\d+,OptoAppClass,ALBIS"
			        		     						            	,	"childs"		: {"Zeilenauswahl":{"Arzt":"Edit1", "Datum":"Edit2"
																							, 	"Kürzel":"Edit3", "Inhalt":"RichEdit20A1"}}}
					            ,	  "Stammdaten"         	: {"parents" 	: "#32770\d+,AfxMDIFrame140\d+,"
																							. "Afx:00400000:b:0001000\d+,MDIClient\d+,OptoAppClass,ALBIS"}
								,	  "Laborblatt"             	: {"parents" 	: "AfxFrameOrView140\d+,AfxFrameOrView140\d+,AfxMDIFrame140\d+,"
																							. "Afx:00400000:b:0001000,MDIClient,OptoAppClass,ALBIS"}
					            ,     "Daten"                   	: {"parents" 	: "#32770,Daten"
																		,	"childs"   	: {"Anrede"         	: "Edit1"		, "Titel"               	: "Edit2"		, "Zusatz"            	: "Edit3"
																							, 	"Vors. Wort"     	: "Edit4"
																							,	"Name"          	: "Edit5"		, "Vorname"       	: "Edit6" 	, "Geb.Datum"   	: "Edit7"
																							, 	"Straße_Straße"	: "Edit8" 	, "Straße_Nr"     	: "Edit9" 	, "Straße_Zusatz"	: "Edit10"
																							,	"Land"             	: "Edit11"	, "PLZ"               	: "Edit12"	, "Ort"               	: "Edit13"
																							,	"Postfach"       	: "Edit14"	, "Postfach_Land"	: "Edit15"	, "Postfach_PLZ" 	: "Edit16"
																							, 	"Postfach_Ort"	: "Edit17"	, "Nationalität"   	: "Edit18"
																							, 	"Telefon-Nr"   	: "Edit19"	, "2.Telef.-Nr"       	: "Edit20"	, "Telefax-Nr"     	: "Edit21"
																							, 	"E-Mail"          	: "Edit22"	, "Arbeitgeber"     	: "Edit23"	, "Patient seit"    	: "Edit24"
																							, 	"Entfernung"   	: "Edit25"	, "Hausarzt"       	: "Edit26"
																							, 	"Interne Zuordnung":"Combobox1", "BG/KH"	: "Edit27"}}
					            ,     "Dauermedikamente" 	: {"parents" 	: "#32770,Dauermedikamente"}
					            ,     "Dauerdiagnosen"       	: {"parents" 	: "#32770,Dauerdiagnosen"}
					            ,     "Patientengruppen"  	: {"parents" 	: "#32770,Patientengruppen"}
					            ,     "Familiendaten"          	: {"parents" 	: "#32770,Familiendaten"}
					            ,     "Abrechnungsassistent"	: {"parents" 	: "#32770,Abrechnungsassistent"}
					            ,     "Cave!"                     	: {"parents" 	: "#32770,Cave!"}}

}

;                                                                                                                                                                                               	  (161)
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
	For key, val in StrSplit(Match, "/") 			     					    		;entfernt überflüssige Leerzeichen
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

	If !AlbisWinID()
		return 0

	If !CheckWindowStatus(AlbisWinID(), 200)	{

		MsgBox, 0x1036, % "Addendum für Albis on Windows",
					(LTrim
						Albis scheint nicht zu reagieren!
						Bevor Sie weiter machen,
						überprüfen Sie bitte das Albisprogramm!
					), 20

		IfMsgBox, Cancel
		{
				MsgBox, 0x1036, Addendum für Albis on Windows,
								(LTrim
								Wollen Sie wirklich das laufende
								Skript (%A_ScriptName%) abbrechen?
								Möglicherweise verlieren Sie dabei Daten.

								Vorschlag, starten Sie Albis erneut und
								drücken Sie dann auf OK.
								), 20
				IfMsgBox, Cancel
					return 0
		}

		IfMsgBox, timeout
			return 0

return AlbisPID()
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

	hMdiChild := RegExMatch(MDITitle, "i)^\s*(?<wnd>(0x)[0-9A-F]+|[\d]+)\s*$", h) ? hwnd : AlbisMDIChildHandle(MDITitle)
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
	; Rückgabeparameter ist 1 wenn es funktioniert hat, 0 wenn nicht und nichts wird zurückgegeben wenn das Childfenster nicht vorhanden war

	; wird nur ausgeführt wenn der Titel gefunden werden kann
		If (hMdiChild := AlbisMDIChildHandle(MDITitle)) 	{
			SendMessage, 0x0221, % hMdiChild,,, % "ahk_id " AlbisMDIClientHandle() ;WM_MDIDESTROY
			return ErrorLevel ? 0 : 1 	;? "1A" : "0A"
		}

return
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
; INFO's VON ANDEREN STEUERELEMENTEN                                                                                                                                   	(04)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetActiveControl           	(02) AlbisGetActiveWindowType   	(03) AlbisLVContent                      	(04) AlbisGetFocus
; (05) #AlbisGetStammPos
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisGetActiveControl(cmd, matchCase:="") {                                   	;-- ermittelt den aktuellen Eingabefocus und gibt zusätzliche Informationen zur Identifikation zurück

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
		* 		Bezeichner 			- 	ihre Bezeichnung für den Steuerelement-Zusammenhang
		*		Ebenenelemente	- 	eine durch Komma getrennte Liste an ClassNN Namen von Steuerelementen,
												diese Steuerelemente haben einen inhaltlichen Zusammenhang z.B.
		*		Elternelemente		- 	eine durch | getrennte Liste an ClassNN Namen von Eltern-Steuerelementen,
												diese lassen die eindeutige Identifikation der Ebenenelemente zu
		*
		* --------------------------------------------------------------------------------------------------------------------------------
		* Parameter
		* --------------------------------------------------------------------------------------------------------------------------------
		*   	cmd                      - hwnd, classNN, identify, content, contraction, ContractionIsEqual
		*		matchCase         	- Vergleichsstring für Abgleich mit Zeilenkürzel
		*
		* - 11.07.2020: 	AfxFrameOrView und AfxMDIFrame wurden intern in Albis umbenannt, aus AfxFrameOrView90 und
									AfxMidiFrame90 wurde AfxFrameOrView140 und AfxMidiframe140
		* - 17.06.2021: 	Sourcecode modernisiert, hat dadurch vielleicht auch etwas Geschwindigkeit gewonnen
		* - 27.06.2021: 	die Definition der Steuerelement Klassenbezeichnungen ausgelagert um Änderungen durch die Compugroup leichter anpassen zu können
	   *
		; static IdentData	:= ["Dokument", "Edit1,Edit2,Edit3,RichEdit20A1", "#32770", "AfxFrameOrView", "AfxMDIFrame"]
		;									Dokument ist meine Bezeichnung für die 4-Controls (Edit1...)
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

			; Dokument
				If InStr(IdentData[1], "Dokument")
					KKZeile := AlbisKarteikarteLesen(hFirstParent)

			; ab hier können spezielle weitere Befehle programmiert werden
				Switch cmd            {

				  ; Inhalt der alle Steuerelemente also der gesamten Zeile zurückgeben
					Case "content":
						If InStr(IdentData[1], "Dokument")
							return {	"fhwnd"         	: GetHex(hFocused)
									, 	"fclassNN"     	: cNN
									,	"fname"          	: KKrv[cNN]
									, 	"fpar"            	: GetHex(hFirstParent)
									, 	"fparClassNN"  	: GetClassNN(hFirstParent, AlbisWinID())
									,	"Identifier"      	: IdentData[1]
									, 	"Edit1"            	: KKZeile.Arzt.text                ; Edit1,2,3,Rich... aus Kompatibilitätsgründen mit meinen älteren Skripten behalten
									, 	"Edit2"            	: KKZeile.Datum.text
									, 	"Edit3"           	: KKZeile.kuerzel.text
									, 	"RichEdit"       	: KKZeile.inhalt.text
									, 	"Arzt"             	: KKZeile.Arzt.text
									, 	"Datum"           	: KKZeile.Datum.text
									, 	"Kuerzel"          	: KKZeile.kuerzel.text
									, 	"Inhalt"          	: KKZeile.inhalt.text
									,	"EditArzt"        	: KKZeile.Arzt.classnn
									,	"EditDatum"      	: KKZeile.Datum.classnn
									,	"EditKuerzel"     	: KKZeile.kuerzel.classnn
									, 	"EditInhalt"      	: KKZeile.inhalt.classnn}

				  ; für Kürzelvergleiche (Feststellung des richtigen Eingabezusammenhanges)
					Case "contraction":
						If InStr(IdentData[1], "Dokument")
							return KKZeile.kuerzel.text

				  ; für Kürzelvergleiche (Feststellung des richtigen Eingabezusammenhanges)
					Case "ContractionIsEqual":
						contraction := ControlGetText(KK.Kuerzel, "ahk_class OptoAppClass")        ; war vorher Edit3
						return (KKZeile.kuerzel.text = matchCase) ? true : false

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
	; letzte Änderung: 09.01.2022

	static winTypes := {"Wartezimmer":"WZ","freie\s+Statistik":"fS", "Prüfung\s+EBM":"PEK", "Tagesprotokoll":"TProt", "Terminplaner":"TP"
								, "Fehlerliste":"FL", "Laborbuch":"Laborbuch"}

	WinGetTitle, LWT, % "ahk_class OptoAppClass"

	if RegExMatch(LWT, "\d\d\.\d\d\.\d\d\d\d")	{
		ControlGet, hTbar		, Hwnd		,, ToolbarWindow321	, % "ahk_id " AlbisMDIChildGetActive()
		ControlGet, Auswahl	, Choice	,, ComboBox1				, % "ahk_id " hTbar
		If CBReturn
			return Auswahl
		If RegExMatch(Auswahl, "i)(Abrechnung|Überweisung)")
			return "Karteikarte|Patientenakte|Abrechnung"
		else If RegExMatch(Auswahl, "i)^P\s+(Priva|Stand)\s")
			return "Karteikarte|Patientenakte|Privatabrechnung"
		else
			return "Karteikarte|Patientenakte|" (Auswahl = "Karteikarte" ? "Karteifenster":Auswahl) "|iWin"   ; Bei dieser Auswahl kann das Infofenster eingefügt werden
	}

	For win, wintype in winTypes
		If RegExMatch(LWT, "i)" win)
			return winType


return "other"
}
;03
AlbisLVContent(hWin, LVClass:="SysListView321", Columns:="1") {    	;-- Listview eines Fenster auslesen (Universalfunktion)

	;Achtung das Auslesen eines Listview Elements funktioniert bisher nur für das Wartezimmer

		/*             	Beschreibung

				universell einsetzbare Funktion zum Auslesen von Listviewfenstern
				es wird ein Objekt ein 2 dimensionaler Array (Tabelle) erzeugt
		*/

		Items := Array()
		ControlGet, result, List,, % LVClass, % "ahk_id " hWin
		For row, line in StrSplit(result, "`n")
			If Trim(line)
		       	Items[row] := StrSplit(line, "`t")

return Items
}
;04
AlbisGetFocus(mainFocus:="", subFocus:="", ByRef retval:="", AutoActivate:=true) {   	;-- smarte Erkennung des Eingabefokus

	/* FUNKTIONSBESCHREIBUNG: AlbisGetFocus() - letzte Änderung: 01.07.2022

		Ähnlich wie AlbisGetActiveControl() sollen Information zum momentanen Tastatur/Mauseingabefokus
		zurückgegeben werden. Das Ziel hier ist eine einfache Klassifikation eines Eingabefokus z.B. des Eingabe-
		ortes in der Karteikarte als "Karteikarte" und "Kürzel". AlbisGetActiveControl() dagegen, ist spezialisiert auf
		das Handling der Karteikarte und gibt andere Informationen wie z.B. das aktive Zeilenkürzel zurück egal
		welches Steuerelement der Karteikarte aktuell fokussiert ist.

		Technisch ermittelt AlbisGetFocus() als erstes den Klassennamen des fokussierten Steuerelementes und
		anschließend die Klassennamen der zugehörigen Eltern-Elemente. Die MDI-Fenster (Afx:) haben feste
		Klassennamen mit einem variablen Teil am Ende. Um eine eindeutige Liste zu erhalten muss der
		variable Teil entfernt werden. Mit einer solchen eindeutigen Liste läßt sich der Eingabeort und das
		Eingabefeld sicher klassifizieren.

		Als Ergebnis erhält man dann z.B. folgendes (*Beispiel 1):
		"RichEdit20A,#32770,AfxFrameOrView140,AfxMDIFrame140,Afx:00400000:b:0001000,MDIClient,OptoAppClass,ALBIS"

		Übergabeparameter:
		-------------------------------------------------------------------------
		[mainfocus,     	- wenn mainFocus und/oder subFocus Strings sind, wird entweder wahr oder falsch zurückgeben,
		 subFocus]      	  wenn der Fokus der Eingabe überstimmt. Damit kann die Funktion auch direkt mit IF-Abfragen verwendet werden.

		[AutoActivate]	- es ist besser das Albisfenster oder die Albisdialoge zu aktivieren bevor man Daten zum aktiven
					     		  Eingabefokus auslesen möchte. GetGuiThreadInfo() welches für das Auslesen der Daten verwendet wird,
								  ermittelt nur Daten des aktiven Eingabefokus. Mit der Option AutoActivate akitiviert die Funktion das im
								  Vordergrund liegende Fenster.

		Rückgabeparameter:
		-------------------------------------------------------------------------
		werden keine Parameter übergeben, erhält man ein Objekt mit folgenden Daten:

				parentList   	(String)  	: die Liste der Eltern-Elemente + Fenstertitel des Eltern-Fensters
				fcontrol      	(String)    	: Klassenname des fokussierten Steuerelementes
				hCaret       	(Hex)    	: Handle des fokussierten Steuerelementes
				hWin         	(Hex)    	: Handle des Eltern-Fensters (parent)
				CaretExist    	(Boolean) 	: wahr oder falsch je nachdem ob das Caret sichtbar oder versteckt ist
				mainFocus   	(String)    	: Gruppenbezeichnung des Eingabefokus
				subFocus   	(String)    	: eindeutigere Bezeichnung eines Steuerelementes

		AlbisGetFocus unterscheidet deutlich mehr Eingabefoci als AlbisGetActiveFocus() und der Erkennungsumfang ist leicht erweiterbar.
		In der Funktion Init_Albis() sind im Objekt cparents bereits einige Fokuszuordnungen vorbereitet. Mittels zusätzlicher Daten kann
		prinzipiell jedes Win32-Steuerelement klassifiziert werden.

		das Rückgabe Objekt enthält z.B. diese Informationen:

		{
		  ;___________________ CARET ___________________________
			"CaretExist": 1,                                                 	; Caret ist vorhanden
			"CaretH": 15,                                                 	; Höhe des Caret
			"CaretW": 1,                                                  	; Breite des Caret
			"CaretX": 3,                                                     	; x Position
			"CaretY": 2,                                                    	; y Positition
		  ;_______________ sub Steuerelemente ___________________
			"childs": {                                                      	; Kind-Elemente (childs)
				"Arzt": {                                                       	; 1.Feld mit dem Buchstaben des Arztes, welcher die Karteikarte bearbeitet
					"classNN": "Edit12",                                 	; ClassNN - wäre eigentlich Edit1 im Normalfall (die Integration meines Infofenster verschiebt die Nummerierung)
					"hwnd":"0xf234d1"
					"text": "A"                                                  	; Steuerelementinhalt/text
				},
				"Datum": {                                                  	; 2.Feld mit dem Datum der Eintragung
					"classNN": "Edit13",                                  	; ClassNN - wäre eigentlich Edit2 im Normalfall (die Integration meines Infofenster verschiebt die Nummerierung)
					"hwnd":"0xab0066"
					"text": "21.12.2021"                                  	; Steuerelementinhalt/text
				},
				"Kuerzel": {                                                 	; 3.Feld mit dem verwendeten Kürzel
					"classNN": "Edit14",                                  	; ClassNN - wäre eigentlich Edit3 im Normalfall (die Integration meines Infofenster verschiebt die Nummerierung)
					"hwnd":"0x450ff4"
					"text": "lko"                                                	; Steuerelementinhalt/text
				}
				"Inhalt": {                                                  	; 4.Feld mit ihren Eintragungen
					"classNN": "RichEdit20A1",                        	; ClassNN  - RichEdit20A1 bleibt immer gleich
					"hwnd":"0x10676"
					"text": "03230"                                          	; Steuerelementinhalt/text
				},
				"hKarteikarte": "0xbf23fb",                          	; Handle des Elternelementes
			},
		    ;___________________________________________________
			"fcontrol": "Edit14",                                         	; classnn des Steuerelementes mit dem Tastaturfokus
			"flags": "0x1",                                                  	;
			"hActive": "0x772106",                                      	; Handle des aktiven Elementes oder Fensters
			"hCapture": "0x0",                                            	;
			"hCaret": "0x13B1976",                                     	; Handle des Caret oder des aktiven Steuerelementes
			"hFirstParent": "0x9900DA",                               	; Handle des Elternelementes (Caret)
			"hFocus": "0x13B1976",                                    	;
			"hMenuOwner": "0x0",                                     	;
			"hMoveSize": "0x0",                                         	;
			"hWin": "0x772106",                                         	; Handle des aktiven Fensters
			"states": "GUI_CARETBLINKING",
			"mainFocus": "Karteikarte",                               	; Was zeigt Albis gerade an?                             	- in diesem Fall eine Karteikarte
			"subFocus": "KUERZEL",                                     	; Wo befindet sich der Tastatur/Eingabefokus? 	- im Kürzelfeld
			"subName": "Zeilenauswahl",                             	;
			"subText": "",                                                   	;
			"wtitle": "ALBIS"                                              	;
			"parentList": "Edit14,#327702,AfxFrameOrView1401,AfxMDIFrame1401,Afx:00400000:b:00010005,MDIClient1,OptoAppClass,ALBIS",
		}

	*/

	global cparents  ;Init_Albis()
	static rxrplWins := "(\s-\s\[.*\])|(\(.*\).*)|(\svon.*)|(\sfür.*)|(\<.*?\>\s*)"

  ; GetGuiThreadInfo() ermittelt keine Daten von Popup Fenstern (Dialogfenster) wenn diese nicht aktiviert sind
  ; es muss zunächst festgestellt werden ob ein Dialogfenster angezeigt wird. Um von diesem Daten zu erhalten muss es zuerst aktiviert werden.
  ; Kommentar/Vermutung: die Aktivierung des Albisfenster muss vermieden werden, da ein erneutes Aktivieren bei vorhandenem Eingabefokus, den Fokus schließen kann. (nicht sicher)
	If (AutoActivate=true) {
		hwinToActivate := (hPopUp := GetWindow(hAlbis:=AlbisWinID(), 6)) && hPopUp<>hAlbis  ? hPopUp : hAlbis
		If (hwinToActivate <> hAlbis) && !WinActive("ahk_id " hwinToActivate) {
			WinActivate    	, % "ahk_id " hwinToActivate
			WinWaitActive	, % "ahk_id " hwinToActivate,, 1
		}
	}

  ; Daten des aktiven (oder gerade aktivierten) Fensters ermittelt
	gthread 	:= GetGUIThreadInfo()
	hCaret  	:= gthread.hCaret ? gthread.hCaret : gthread.hFocus
	hWin      	:= gthread.hActive
	wtitle      	:= RegExReplace(WinGetTitle(hWin), rxrplWins)
	parentList 	:= GetParentClassList(hCaret, hWin) ","                                                   ; durch Komma separierte List
	parentList 	:= RegExReplace(parentList, ":\w+:\w+\,", ",") WinGetClass(hWin) "," wtitle

	RegExMatch(parentList, "^(?<control>.+?),", f)
	retval                 	:= gthread
	retval.parentList 	:= parentList
	retval.fcontrol    	:= fcontrol
	retval.hWin        	:= hWin
	retval.wtitle        	:= wtitle
	retval.hFirstParent 	:= GetParent(hCaret)
	retval.CaretExist		:= gthread.hCaret ? true : false

  ; mit den Daten aus cparents vergleichen
	For focusedWin, match in cparents
		If RegExMatch(parentList, match.parents) {

			retVal.mainFocus := focusedWin

		; kein subfocus wenn keine child controls in cparent vorhanden sind
			If !IsObject(match.childs)
				break

		; subfocus ermitteln
			If (focusedWin <> "Karteikarte") {
				For focusedControl, classNn in match.childs
					If !IsObject(classnn) {
						If RegExMatch(parentList, "^" classNn ",") {
							retVal.subFocus := focusedControl
							break
						}
					}
					else {
						For subname, controlclass in classnn
							If (controlclass == fcontrol) {
								retVal.subFocus := focusedControl
								retval.subName	:= subname
								break
							}
					}
			}

		; anderes Vorgehen bei dem subfocus der Karteikarte
			else {

				hactiveMDI 	:= AlbisMDIChildGetActive()
				retval.childs	:= AlbisKarteikarteLesen(hactiveMDI)
				For focusedControl, classNn in match.childs
					For subcontrol, control in retval.childs
						If (control.classnn == fcontrol) {
							retVal.subFocus	:= subcontrol
							retval.subName	:= focusedControl
							retval.subText  	:= control.Text
							break
						}
					}

				If retVal.subFocus
					break

			}

   ; Rückgabeparameter für Benutzung in If-Abfragen
	If mainFocus {
		mF	:= RegExMatch(retval.mainFocus, "i)" mainFocus)                    	? 1 : 0
		sF 	:= !subFocus ? 1 : (RegExMatch(retval.subFocus, "i)" subFocus)	? 1 : 0)
	return mF & sF 	; subFocus ? mF << sF : mF
	}

return retval
}
;05
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
;06
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; PATIENTENDATEN                                                                                                                                                                      	(09)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAktuellePatID                 	(02) AlbisCurrentPatient              	(03) AlbisPatientGeschlecht           	(04) AlbisPatientGeburtsdatum
; (05) AlbisPatientVersicherung         	(06) AlbisVersArt                        	(07) AlbisTitle2Data                      	(08) AlbisAbrechnungsscheinVorhanden
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
	If !(AlbisGetActiveWindowType() ~= "(Karteikarte|Patientenakte|Karteifenster|iWin)")
		return

	WTAlbis := !WTAlbis ? AlbisGetActiveWinTitle() : WTAlbis

	Splitstr				:= StrSplit(WTAlbis	, "/", A_Space)
	SplitStr2			:= StrSplit(SplitStr[2]	, ",", A_Space)

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
; Cave! von - Fenster                                                                                                                                                                        	(08)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetCaveZeile                	(02) AlbisGetCave                     	(03) AlbisCaveAddRows                	(04) AlbisCaveGetCellFocus
; (05) AlbisCaveGetCellFocus         	(06) AlbisCaveGVU                     	(07) AlbisCaveUnFocus               	(08) AlbisSetCaveZeile
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

			WinActivate	 , % "ahk_id " hCave
			WinWaitActive, % "ahk_id " hCave,, 1
			while (hCave := WinExist("Cave! von")) && (A_Index <= 45)	{
				VerifiedClick("Button4", hCave)    	; ich wählte Button4 = Abbrechen ; Programmfehler können so keinen Schaden anrichten, alternativ ginge es auch mit zweimal Escape hintereinander
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
		If !(hCave := Albismenu(32778, "Cave! von ahk_class #32770", 4))
			return 0

	; Auslesen aller Zeilen des Cave! von - Dialoges
		ControlGet, result, List,, SysListView321, % "ahk_id " hCave

	; CaveVonFenster schließen
		If CloseCave {

			WinActivate	 , % "ahk_id " hCave
			WinWaitActive, % "ahk_id " hCave,, 1
			while (hCave:= WinExist("Cave! von ahk_class #32770"))	&& (A_Index <= 30)	{
				VerifiedClick("Button4", "Cave! von ahk_class #32770")   	; wähle Button4 = Abbrechen
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
AlbisCaveGetCellFocus(hCave) {                                                  	;-- ermittelt die aktuell ausgewählte Reihe und Spalte

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
AlbisCaveGVU(GVString, GVUDatum:="") {                                	;-- hilft Hinweise zu Vorsorgeuntersuchungen im Cave! von Fenster einzutragen

  ; letzte Änderung: 03.07.2022

	static rxVorsorge := 	"^\s*(?<Colo1>01740)*[\s,]*(?<Colo2>\(\s*C[\pL\s\d]+\))*[\s,]*"
			        		.    	"(?<Aorta>A|AORTA)*[\s,]*"
			        		.		"(?<GVU1>GVU)*\/*(?<HKS1>HKS)*\/*(?<KVU1>KVU)*\s*(?<GHK>\d+.\d+)*[\s,]*"
			        		.		"(?<GVU2>GVU\s*\d+.\d+)*[\s,]*"
			        		.		"(?<HKS2>HKS\s*\d+.\d+)*[\s,]*"
			        		.		"(?<KVU2>KVU\s*\d+.\d+)*[\s,]*"
			        		. 		"(?<GB>GB)*\s*(?<GBy>\d+)*\.*(?<GBq>\d)*(?<GBn>[AB])*"
	static rxVSShorts := "i)(GVU|HKS|KVU|01740|C\d+|GB\s*\d+\.\d|Aorta|[,\s]+A[\s,]+|Aorta)"
	static hfocused, hCave, newGVString

	dbg := false
	newGVString := ""
	hCave := hfocused := 0

  ; prüft den übergeben Parameter - GVString
	If !RegExMatch(GVString, "Oi)" rxVorsorge, cv)
		return 0

	If dbg
		SciTEOutput("cv: " GVString ", cv.GHK: " cv.GHK)

  ; Cave!von öffnen
	If !(hCave := Albismenu(32778, "Cave! von ahk_class #32770", 4)) {
		ControlFocus,, % "ahk_class OptoAppClass"
		ControlSend,, {LControl Down}4{LControl Up}, % "ahk_class OptoAppClass"
		WinWait, % "Cave! von ahk_class #32770",, 1
		hCave := WinExist("Cave! von ahk_class #32770")
		MsgBox, 0x1031, Achtung!, Bitte das Cave! von Fenster manuell öffnen und danach Ok drücken.
		IfMsgBox, Cancel
			ExitApp
		If !(hCave := WinExist("Cave! von ahk_class #32770"))
			ExitApp
	}

  ; prüft ob das Handle hCave zum Cave! von Fenster gehört
	while !RegExMatch(cvTitle := WinGetTitle(hCave), "i)Cave!\s+von\s+\<") && A_Index < 10 {
		If (A_Index>1)
			Sleep 50
		hCave := WinExist("Cave! von ahk_class #32770")
	}

  ; Cave!von auslesen
	LVCave  	:= AlbisGetCave(false)

  ; Zeile mit den Vorsorgeuntersuchungen finden
	For cvIndex, line in StrSplit(LVCave, "`n") {
		description := StrSplit(line, A_Tab).4
		If RegExMatch(description, rxVSShorts) && RegExMatch(description, "Oi)" rxVorsorge, cvO) {
			GVIndex := cvIndex
			If FilePathCreate(Addendum.DBPath "\logs")
				FileAppend, % GVUDatum " | [" AlbisAktuellePatID() "] " AlbisCurrentPatient() " | " description "`r`n", % Addendum.DBPath "\logs\Cave von.txt", UTF-8
			break
		}
	}

 ; GVU/HKS/KVU String erstellen                                                                      	;{
	GVU1 := (cv.GVU1 ? "GVU" : "") (cv.GVU1&& cv.HKS1 ? "/HKS" : cv.HKS1? "HKS") ((cv.GVU1 || cv.HKS1)
			&& 	cv.KVU1 ? "/KVU" : cv.KVU1 ? "KVU" : "") (cv.GHK ? " " cv.GHK : "")

  ; Ausgliedern von Untersuchungen die aktuell nicht durchgeführt wurden
	GVU2 := (cvO.GVU1 && !cv.GVU1 ? "GVU" : "")                            	; GVU/
	GVU2 .= (cvO.HKS1 && !cv.HKS1 ? (GVU2 ? "/" : "") "HKS" : "")        	; GVU/HKS/
	GVU2 .= (cvO.KVU1 && !cv.KVU1 ? (GVU2 ? "/" : "") "KVU" : "")      	; GVU/HKS/KVU
	GVU2 := RTrim(GVU2, "/")
	GVU2 .= (GVU2 && cvO.GHK ? " " cvO.GHK : "")                           	; GVU/HKS/KVU 06^22

  ; einzeln stehende Untersuchungen welche
	GVU3 := (cvO.GVU2 && !cv.GU1 ? cvO.GVU2 ", " : "") (cvO.HKS2 && !cv.HKS1 ? cvO.HKS2 ", " : "") (cvO.KVU2 && !cv.KVU1 ? cvO.KVU2 ", " : "")
  ;}

  ; geriatrisches Basisassement                                                                       	;{
	GBold := (cvO.GBy cvO.GBq) + 0
	GBnow := (cv.GBy cv.GBq) + 0
	If (GBold && GBnow)
		GB :=  GBold < GBnow ?  "GB " cv.GBy "." cv.GBq cv.GBn
	else if GBnow
		GB := "GB " cv.GBy "." cv.GBq cv.GBn
	else if GBold
		GB := "GB " cvO.GBy "." cvO.GBq cvO.GBn
  ;}

 ; kompletten Cave String zusammenfügen                                                     	;{
	newGVString .= (cvO.Colo1 ? cvO.Colo1 : "")
	newGVString .= (cv.Colo2 ? cv.Colo2 ", " : cvO.Colo2 ", ")
	newGVString	.= (cv.Aorta ? cv.Aorta "," : cvO.Aorta ? cvO.Aorta ", " : "")
	newGVString .= GVU1 (GVU1 ? ", " : "") GVU2 (GVU2 || GVU1 ? ", " : "") GVU3
	newGVString := RTrim(newGVString, ", ") ((GVU1||GVU2||GVU3) && GB ? ", " : "") GB
	newGVString := RTrim(newGVString, ", ")
	newGVString := RegExReplace(newGVString, "[,\s]+$")
	newGVString := RegExReplace(newGVString, "^[,\s]+")

	If dbg
		SciTEOutput(newGVString)
  ;}

  ; den Fokus aus dem Cave!von Listview-Steuerelement nehmen
	unfocused	:= AlbisCaveUnFocus()

  ; Fenster aktivieren
	WinActivate, % "ahk_id " hCave
	WinWaitActive, % "Cave! von ahk_class #32770",, 1

  ; die Zeile mit dem Vorsorgeeintrag auswählen, den Eintrag für Eingaben freischalten und den neuen String als ToolTip unter dem Steuerelement einblenden
	If GVIndex {

		WinActivate, % "ahk_id " hCave
		WinWaitActive, % "Cave! von ahk_class #32770",, 1

	  ; Zeile auswählen
		ControlSend, SysListview321, % GVIndex, % "ahk_id " hCave
		Sleep 500

	  ; prüft das die richtige Zeile gewählt wurde, falls nicht erfolgt ein zweiter Versuch
		ControlGet, hLV, hwnd,, SysListview321, % "ahk_id " hCave
		selRow := DllCall("SendMessage", "uint", hLV, "uint", 4108, "uint", 0, "uint", 0x3) + 1
		If (selRow != GVIndex) {
			ControlSend, SysListview321, % GVIndex, % "ahk_id " hCave
			Sleep 500
		}

	  ; jetzt das Edit1 Steuerelement (Spalte Beschreibung) freischalten
		ControlSend, SysListview321, {Space}, % "ahk_id " hCave
		Sleep 500
		ControlSend, SysListview321, {Tab}, % "ahk_id " hCave
		Sleep 400
		ControlSend, SysListview321, {Tab}, % "ahk_id " hCave

	 ; Edit1 - Steuerelement sollte offen liegen
		ControlGet, hfocused, hwnd,, Edit1, % "ahk_id " hCave
		classNN := Control_GetClassNN(hCave, hfocused)
		If (classNN = "Edit1") {

			description := ControlGetText("Edit1", "ahk_id " hCave)
			If RegExMatch(description, rxVSShorts) && RegExMatch(description, "Oi)" rxVorsorge) {

              ; ToolTip einblenden
				cvpos := GetWindowSpot(hfocused)
				ToolTip, % newGVString "`n(Übernahme mit -Enter- , Abbrechen mit -Escape-)", % cvpos.X, % cvpos.Y+cvpos.H, 4

             ; Hotkey's ind sauberer
				Hotkey, IfWinActive, % "Cave! von ahk_class #32770"
				Hotkey, $Escape	, AlbisCaveGVUHotkey	, On
				Hotkey, $Enter		, AlbisCaveGVUHotkey 	, On
				Hotkey, IfWinActive

				SetTimer, AlbisCaveGVUCheck, 100
			}
		}
	}

return -1

AlbisCaveGVUHotkey: ;{

	SetTimer, AlbisCaveGVUCheck, Off
	ToolTip,,,, 4

	If (A_ThisHotkey ~= "i)Escape") {

		If WinExist("Cave! von ahk_class #32770")
			If !VerifiedClick("Button4", hCave)
				VerifiedClick("Abbruch", hCave)

	}
	else If (A_ThisHotkey ~= "i)Enter") {

		If WinExist("Cave! von ahk_class #32770") {

		  ; fokussieren
			ControlFocus,, % "ahk_id " hfocused
			Sleep 50

		  ; löschen
			SendInput, {BS}
			Sleep 100

		  ; Eintrag einfügen
			ControlSendRaw, Edit1, % newGVString, % hCave

		  ; prüfen und weiterer Versuch: Textinhalt löschen, dann neuen Eintrag einfügen
			If (newGVString != ControlGetText("Edit1", "ahk_id " hCave)) {

			  ; fokussieren
				ControlFocus,, % "ahk_id " hfocused
				Sleep 50

			  ; löschen
				SendInput, {BS}
				Sleep 200

			  ; prüfen, alles auswählen, löschen
				If ControlGetText("Edit1", "ahk_id " hCave) {
					SendInput, {LControl Down}a{LControl Up}
					Sleep 100
					SendInput, {BS}
					Sleep 100
				}

			  ; Buchstabe für Buchstabe einfügen
				ControlFocus, Edit1, % "ahk_id " hCave
				For each, key in StrSplit(newGVString) {
					SendRaw, % key
					Sleep 20
				}

			}

		 ; Texteingabe abschließen
			ControlFocus, Edit1, % "ahk_id " hCave
			Sleep 50
			ControlSend,, {Tab}, % "ahk_id " hfocused
			Sleep 200

		  ; Cave!von Fenster schließen
			If !VerifiedClick("Button3", hCave)
				VerifiedClick("OK", hCave)

		 ; Fenster nicht geschlossen?
			If WinExist("Cave! von ahk_class #32770") {
				ControlFocus,, % "ahk_id " hCave
				Sleep 50
				ControlSend,, {Enter}, % "ahk_id " hCave
			}

		}

	}

AlbisCaveGVUClip:

	Hotkey, IfWinActive, % "Cave! von ahk_class #32770"
	Hotkey, $Escape	, AlbisCaveGVUHotkey	, Off
	Hotkey, $Enter		, AlbisCaveGVUHotkey	, Off
	Hotkey, IfWinActive

  ; newGVString ins Clipboard kopieren
	Clipboard := newGVString
	ClipWait, 2
	PraxTT("Die neue Zeichenkette`n#2<" newGVString ">`n wurde ins Clipboard kopiert.", "6 1")

return ;}

AlbisCaveGVUCheck:  ;{
	ControlGet, hEdit, hwnd,, Edit1, % "ahk_id " hCave
	classNN := Control_GetClassNN(hCave, hEdit)
	If !WinExist("Cave! von ahk_class #32770") || (classNN != "Edit1") {
		SciTEOutput("Focus lost")
		SetTimer, AlbisCaveGVUCheck, Off
		ToolTip,,,, 4
		goto AlbisCaveGVUClip
	}
return ;}
}
;07
AlbisCaveUnFocus(hCave:="") {                                                   	;-- entfernt den Eingabefokus

	; letzte Änderung 02.07.2022

	hCave :=  !hCave ? WinExist("Cave! von ahk_class #32770") : hCave
	ControlGet, hCaveLV, HWND,, SysListView321, % "ahk_id " hCave
	ControlGet, hEdit   	, HWND,, Edit1, % "ahk_id " hCaveLV

	while (DllCall("IsWindowVisible","Ptr", hEdit) && A_Index < 5) {
		If (A_Index > 1)
			sleep 300
		ControlFocus, SysListview321, % "ahk_id " hCave
		ControlGet, hEdit, HWND,, Edit1, % "ahk_id " hCaveLV
	}

return DllCall("IsWindowVisible","Ptr", hEdit)
}
;08
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
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; FENSTERELEMENTE AUSLESEN UND BEHANDELN                                                                                                                          	(07)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisMuster30ControlList       	(02) AlbisOptPatientenfenster       	(03) AlbisHeilMittelKatologPosition	  	(04) AlbisSortiereDiagnosen
; (05) AlbisReadFromListbox             	(06) AlbisResizeDauerdiagnosen  	(07) AlbisResizeLaborAnzeigegruppen
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisMuster30ControlList() {                                                       	;-- Änderungen von Check-/Radiobuttons erkennen - speziell für das Muster30 Formular

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
;02
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
;03
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
;04
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
;05
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
;06
AlbisResizeDauerdiagnosen(Options:= "xCenterScreen yCenterScreen w0 h0") { ;-- spezielle Funktion für den Dauerdiagnosendialog

	static WT:= "Dauerdiagnosen von ahk_class #32770"
	static hStored

	hWin := WinExist(WT)
	If (!hWin || hStored = hWin)
		return 0

	opt := ParseGuiOptions(Options)

	hStored	:= hWin
	hMed	:= Controls("SysListView322", "ID", hWin)
	hDia 	:= Controls("SysListView321", "ID", hWin)
	dv      	:= GetWindowSpot(hWin)
	med    	:= GetWindowSpot(hMed)
	dia     	:= GetWindowSpot(hDia)


  ; seit Albisupdate Q2/2022 nicht mehr vorhanden?
	Control, Hide,, ThunderRT6UserControlDC1	, % WT
	Control, Hide,, Internet Explorer_Server1     	, % WT

	ControlGetPos,, Btn1Y, Btn1W,, Button1      	, % WT
	ControlGetPos,, Btn2Y, Btn2W,, Button2     	, % WT
	ControlMove, Button1, % dv.CW - Btn1W,,,	, % WT
	ControlMove, Button2, % dv.CW - Btn2W,,,	, % WT

  ; bei Privatpatienten werden keine Dauermedikamente angezeigt, der Dialog muss nicht vergrößert werden
	If hMed {
		ControlMove, SysListView322,,, 400 ,                  	 , % WT
		inCW := dv.CW - Btn1W - 5
		ControlMove, SysListView321, 430,, % inCW - 440 ,, % WT
	}

  ; Fensterbreite so anpassen das der gesamte Dauerdiagnosentext sichtbar ist
  ; die Summe der Spaltenbreite minus der aktuellen Listviewbreite ergibt die zu addierende Breite für das Dialogfenster
	lvW       	:= 0
	dv         	:= GetWindowSpot(hWin)
	dia        	:= GetWindowSpot(hDia)
	hHeader  	:= GetHex(DllCall("SendMessage", "uint", hDia, "uint", 4127))
	header  	:= GetHeaderInfo(hHeader)
	For index, col in header
		lvW += col.Width, colText .= (A_Index > 1 ? "," : "") col.Text

	;~ SciTEOutput(hDia ", " hHeader ": " lvW ", " colText)

	If (lvW > dia.W) {
		newWinWidth := dv.W + lvW + 2*dv.BW - dia.W
		SetWindowPos(hWin, dv.X, dv.Y, newWinWidth, dv.H)
	}

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
;07
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
; DATUMSFUNKTIONEN                                                                                                                                                                  	(05)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisZeilenDatumLesen           	(02)  AlbisZeilenDatumSetzen     	(03) AlbisSetzeProgrammDatum 	(04) AlbisLeseProgrammDatum
; (05) AlbisSchliesseProgrammDatum
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
; (01)
AlbisZeilenDatumLesen(sleeptime=80, closeFocus=true) {                	;-- liest das Zeilendatum der ausgewählten Zeile in der Akte aus

	; Achtung eine Zeile in der Akte muss nicht und sollte nicht ausgewählt sein, der Mauspfeil sollte aber über der Zeile stehen
	; dies funktioniert nur in der aktuellen Patientenakte (es darf noch keine Abrechnung stattgefunden haben)
	; letzte Änderung: 12.01.2022 - Erkennung des aktiven Fokus verbessert, Codesyntax modernisiert

		global KK, KKx

	; Abbruch wenn keine Karteikarte angezeigt wird
		If !InStr(AlbisGetActiveWindowType(), "Karteikarte")
			return -1
		If !WinActive("ahk_class OptoAppClass")
			AlbisActivate(1)

	; falls nur eine Zeile in der Karteikarte blau markiert ist, muss zunächst eine Eingabezeile angefordert werden
		If !AlbisGetFocus("Karteikarte", ".", KKarte) {
			SciTEOutput("ups das hab ich nicht erwartet.. " res)
			while (A_Index < 10) {
				SendInput, {Insert}
				If AlbisGetFocus("Karteikarte", ".", KKarte)
					KKInput := true
					break
				If (A_Index>1)
					sleep % sleeptime
			}

		; Abbruch wenn sich keine Eingabezeile öffnen ließ
			If !KKInput
				return -2

		}

	; Datum auslesen
		ControlGetText, Datum, % KK.Datum, % "ahk_id " KKarte.hFirstParent
		If !RegExMatch(Datum, "\d\d\.\d\d\.\d\d\d\d")
			ControlGetText, Datum, % KKx.Datum, % "ahk_id " KKarte.hFirstParent

	; fragt den Focus ab, Esc beendet den Eingabemodus in der Akte, es sollte danach ein anderer Focus auslesbar sein
		If closeFocus
			while (A_Index < 4) {
				If (A_Index > 1)
					sleep % sleeptime
				while (A_Index < 5)	{
					SendInput, {Esc}
					Sleep % sleeptime
					If !AlbisGetFocus("Karteikarte", ".", KKarte)
						break
				}
			}

return RegExMatch(Datum, "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})") ? Datum : ""
}
; (02)
AlbisZeilenDatumSetzen(DateChoosed) {                                           	;-- Zeilendatum in der Karteikarte ändern

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
		ZDatum 	:= AlbisZeilenDatumLesen()
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

return AlbisZeilenDatumLesen() <> DateChoosed ? false : true
}
; (03)
AlbisSetzeProgrammDatum(Datum:="", closeWin=true, dbg=false) { 	;-- ändert das Programmdatum

		; Achtung: Datum muss absofort zwingend folgendes Format aufweisen: dd.mm.yyyy z.B. 24.12.2019
		; bei Übergabe eines leeren Datumstrings wird das aktuelle Tagesdatum eingetragen
		; letzte Änderung: 26.06.2022

		static Prgdatum := "Programmdatum einstellen ahk_class #32770"
		static rxDate := "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})"

		dhw := A_DetectHiddenWindows, dht := A_DetectHiddenText
		tmm := A_TitleMatchMode, tmms := A_TitleMatchModeSpeed
		DetectHiddenWindows	, On
		DetectHiddenText      	, On
		SetTitleMatchMode     	, 2
		SetTitleMatchMode     	, slow

		If !RegExMatch(Datum, rxDate)
			FormatTime, Datum,, dd.MM.yyyy

	; prüft auf ein korrektes Datumformat
		Datum := StrReplace(Datum, " ")
		if !RegExMatch(Datum, rxDate) 		{
			PraxTT("Das übergebene Datum [" Datum "] hat kein gültiges Format. [dd.MM.yyyy]", "2 1")
			return -4
		}

	; Dialog 'Programmdatum einstellen' aufrufen
		AlbisActivate(1)
		Albismenu(32852, "Programmdatum einstellen", 3)
		sleep 100

	  ; warte bis das Fenster erscheint (prüft genau ob es richtig erfasst wurde)
	    found := false
		Loop {

			If (A_Index > 80)
				break

		  ; tiefergehende Prüfung 	- 	wie ist der Fenstertitel zum erhaltenen Fensterhandle
		  ;                                 	 	und wie ist die ClassNN des Edit Steuerelementes
			If (hProgDate := WinExist(Prgdatum)) {
				            hEdit 	:= ControlGet("HWND" ,, "Edit1", "ahk_id " hProgDate)
				ProgDateTitle	:= WinGetTitle(hProgDate)
				If (hEdit && ProgDateTitle ~= "i)Programmdatum\s+einstellen") {
					ProgDateEditNN := Control_GetClassNN(hProgDate, hEdit)
					If (ProgDateEditNN = "Edit1") {
						eingestelltesDatum := ControlGetText("", "ahk_id " hEdit)
						If dbg
							SciTEOutput(" [" A_ThisFunc "] (" SubStr("0" A_Index, -1) ") hProgDate = " hProgDate
																			. "`n                                ProgDateTitle = "  ProgDateTitle
																			. "`n                                hEdit = " hEdit " {" ProgDateEditNN "}"
																			. "`n                                Datum = " eingestelltesDatum  "`n")
						found := true
						break
					}
				}
			}

			sleep 100

		}

		If dbg && (!ProgDateTitle || !hEdit)
			SciTEOutput(" [" A_ThisFunc "] (" SubStr("0" A_Index, -1) ") hProgDate = " hProgDate
																			. "`n                                ProgDateTitle = "  ProgDateTitle
																			. "`n                                hEdit = " hEdit " {" ProgDateEditNN "}"
																			. "`n                                Datum = " eingestelltesDatum "`n")

	  ; handles ließen sich nicht ermitteln dann muss die Funktion hier beendet werden
		If !found {
			PraxTT(A_ScriptName "`nDas Dialogfenster 'Programmdatum einstellen' konnte nicht aufgerufen werden! `nDie Funktion wird beendet", "2 1")
			DetectHiddenWindows	, % dhw
			DetectHiddenText      	, % dht
			SetTitleMatchMode     	, % tmm
			SetTitleMatchMode     	, % tmms
			If dbg
				SciTEOutput(" [" A_ThisFunc "] hProgDate = " (!hProgDate ? "0x0" : hProgDate) " , ProgDateTitle = " (!ProgDateTitle ? " % " : ProgDateTitle) "´n")
			return -3
		}

	; aktuelles Programmdatum ermitteln und sichern
		If !RegExMatch(eingestelltesDatum, rxDate) {
			eingestelltesDatum := ControlGetText("Edit1", "ahk_id " hProgDate)
			sleep 100
		}

	  ; prüft das ermittelte Datum, bei einem Fehler wird versucht das Datum durch simuliertes auswählen des Inhalts und senden von Strg+C zu erhalten
		If !RegExMatch(eingestelltesDatum, rxDate) {
			ControlSend	, % "Edit1", % "{LControl Down}a{LControl Up}", % "ahk_id " hProgDate
			sleep 50
			ControlSend	, % "Edit1", % "{LControl Down}c{LControl Up}", % "ahk_id " hProgDate
			ClipWait, 1
			ControlSend	, % "Edit1", % "{LControl Up}", % "ahk_id " hProgDate
			sleep 50
			eingestelltesDatum := Clipboard
		}

		If dbg
			SciTEOutput(" [" A_ThisFunc "] (A)  eingestelltesDatum = " eingestelltesDatum  "`n")

	  ; nochmaliges Überprüfen. Wenn wiederrum nicht korrekt mit heutigem Datum füllen.
		If !RegExMatch(eingestelltesDatum, rxDate) {
			PraxTT("Das eingestellte Programmdatum konnte nicht ausgelesen werden!`nEs wird stattdessen das Tagesdatum verwendet", "3 1")
			eingestelltesDatum := A_DD "." A_MM "." A_YYYY
		}

	; übergebenes Datum setzen
		VerifiedSetText("Edit1", Datum, hProgDate)
		neuesDatum := ControlGetText("Edit1", "ahk_id " hProgDate)
		If (Datum <> neuesDatum) {
			VerifiedSetText("Edit1", Datum, hProgDate)
			ControlGetText, neuesDatum, Edit1, % Prgdatum
		}

	; Dialogfenster schliessen
		If closeWin
			res := AlbisSchliesseProgrammDatum()

	; DetectHidden Modi zurücksetzen
		DetectHiddenWindows	, % dhw
		DetectHiddenText      	, % dht
		SetTitleMatchMode     	, % tmm
		SetTitleMatchMode     	, % tmms

return eingestelltesDatum? eingestelltesDatum : ""
}
; (04)
AlbisLeseProgrammDatum(closeWin:=true) {                                    	;-- liest das aktuell eingestellte Programmdatum

	; V1.2 mit Überprüfung ob das Fenster geöffnet ist, prüft das Auslesen des Datum und wird erst beendet wenn der Dialog geschlossen wurde

	; Fenster Programmdatum einstellen aufrufen
		If !(hWin:= Albismenu(32852, "Programmdatum einstellen", 3)) {
			PraxTT(A_ScriptName ": Das Programmdatum ließ sich nicht anzeigen!", "1 0")
			return 0
		}

	; Datum lesen
		ControlGetText, Datum, Edit1, % "Programmdatum einstellen ahk_class #32770"
		while (!Datum && A_Index <= 40) {
			If (A_Index > 1)
				Sleep 50
			ControlGetText, Datum, Edit1, % "Programmdatum einstellen ahk_class #32770"
		}

	; Fenster schliessen
		If closeWin
			AlbisSchliesseProgrammDatum()

return Datum
}
; (05)
AlbisSchliesseProgrammDatum() {                                                  	;-- schließt das Programmdatumsfenster

	; letzte Änderung: 03.05.2022

	static Prgdatum := "Programmdatum einstellen ahk_class #32770"
	static Fehlerhaft := {"wtitle":"ALBIS ahk_class #32770", "wtext":"Fehlerhaftes Datumsformat"}
	static press := ["OK", "Abbruch"]

	btntoclick := press.1
	while WinExist(Prgdatum) && (A_Index < 6)	{

		If (A_Index<>2)
			VerifiedClick(btntoclick, Prgdatum,, 1)
		else If (A_Index = 2)
			ControlSend, Edit1, {ENTER}, % Prgdatum

	  ; das Schließen des Dialogfenster schneller erkennen
		while WinExist(Prgdatum) && (A_Index < 100) {

			sleep 20
			If WinExist(Fehlerhaft.wtitle, Fehlerhaft.wtext) {
				btntoclick := press.2
				While (hFehlerhaft := WinExist(Fehlerhaft.wtitle, Fehlerhaft.wtext)) {
					If !VerifiedClick("OK",,, hFehlerhaft, 1)
						If !VerifiedClick("Button1",,, hFehlerhaft, 1)
							WinClose, % "ahk_id " hFehlerhaft,, 1
				}
			}

		}

	}

  ; Abbruch wegen fehlerhaftem Datumsformat zurückmelden
	If (btntoclick = press.2)
		return -2

return WinExist(Prgdatum) ? true : -1
}
; (06)
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; DATENEINGABE                                                                                                                                                                           	(13)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisPrepareInput                 	(02) AlbisSendInputLT                	(03) AlbisWriteProblemMed          	(04) AlbisKarteikarteAktivieren
; (05) AlbisKarteikarteEingabeStatus  	(06) AlbisKarteikarteEingabe       	(07) AlbisSchreibeLkMitFaktor      	(08) AlbisSchreibeInKarteikarte
; (09) AlbisFehlendeLkoEintragen    	(10) AlbisKopiekosten                  	(11) AlbisSchreibeSequenz           	(12) AlbisSendText
; (13) AlbisVordatierer
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisPrepareInput(Name) {                                                                                    	;-- bereitet das Schreiben von Daten in die Akte vor


		; letzte Änderung 27.09.2021:
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
				Exception := {"file":A_ScriptFullPath, "what": "wrong parameter call!", "line":"If Instr(allowed, Name) {", "extra":A_ThisFunc}
				FehlerProtokoll(Exception, false)
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

	active := 0
	VerifiedClick("#327702","","", hKarteikarte:= AlbisMDIChildGetActive()) 	;neu seit 27.3.2019 - hoffentlich zuverlässig
	while (!WinActive("ahk_id " hKarteikarte) && A_Index<11) {
		sleep 10
	}

return hKarteikarte
}
;05
AlbisKarteikarteEingabeStatus(KKTitle:="") {                                                              	;-- ermittelt ob die akt. Karteikarte für Eingaben bereit ist

	; Funktion ermittelt das HWND des RichEdit-Steuerelementes (RichEdit20A1) und prüft es auf Sichtbarkeit.
	; (Sichtbarkeit im Sinne das Eingaben gemacht werden können)
	; KKTitle = PatNr oder Name des Pat. um die Karteikarte finden zu können
	; letzte Änderung 26.09.2021

		global KK

		ControlGet, hRE, hwnd,, % KK.INHALT, % "ahk_id " (KKTitle ? AlbisMDIChildHandle(KKTitle) : AlbisMDIChildGetActive())
		WinGet, Style, Style, % "ahk_id" hRE

return Style & 0x10000000 ? true : false
}
;06
AlbisKarteikarteEingabe(Control, AutoActivate:=false) {                                          	;-- gezielter Schreibzugriff auf die Karteikarte

	; Control: es kann direkt die ClassNN des Steuerelementes übergeben werden oder einer der Alias-Bezeichnungen aus KK
	; KKx wird über die Funktion Init_Albis() bereitgestellt
	;
	; letzte Änderung 01.07.2022

		global Mdi:= Object(), KK
		static hKarteikarte

	; Control Optionen parsen
		If RegExMatch(Control, "O)(?<name>[\pL\d\_]+)\s*\=\s*(?<text>.+)$", pCtrl)
			Control := pCtrl.name, ControlText := Trim(pCtrl.text)
		else
			Control := RegExReplace(Control, "[\=\s]")


	; Alias-Bezeichnung wird durch ClassNN des Steuerelementes ersetzt
		ControlNN := KK.HasKey(Control) ? KK[Control] : Control

	; prüft und zeigt die Karteikarte an, falls etwas anderes beim Patienten angezeigt wird
		If AutoActivate {
			AlbisActivate(2)
			If !AlbisKarteikarteZeigen()
				return 0
		}

	; ermittelt das richtige hwnd des Steuerelementes der Karteikarte
		hKarteikarte 	:= AlbisKarteikarteGetActive()
		Unfocus    	:= true

	; eventuell vorhandenen Eingabefokus entfernen
		If AlbisKarteikarteEingabeStatus()  {	 	; prüft ob Eingaben möglich sind

		  ; nicht schließen falls der Inhalt des Steuerelementes (Control) mit ControlText  übereinstimmt
			If ControlText {
				Kkarte := AlbisGetFocus()
				;~ If IsObject(Kkarte)
					;~ SciTEOutput("A: " Control ", <" ControlText "> " AlbisKarteikarteEingabeStatus() "`n" cJSON.Dump(Kkarte, 1))
				If (Kkarte.childs[Control].Text = ControlText)
					Unfocus := false
			}

		}

	; Tastaturfokus auf den Karteikartenbereich setzen und Tastenfolge senden
		If Unfocus {

		 ; Escape senden und Eingabefokus überwachen
			fcontrol := AlbisKarteikarteUnfocus(hKarteikarte, "fcontrol")

			ControlSend,, {End}, % "ahk_id " hKarteikarte
			sleep 75
			ControlSend,, {Down}, % "ahk_id " hKarteikarte
			sleep 100

		; zweiter Versuch
			Kkarte := AlbisGetFocus()
			If (Kkarte.subName != "Zeilenauswahl") {
				VerifiedClick("", "ahk_id " hKarteikarte)
				ControlSend,, {End}  	, % "ahk_id " hKarteikarte
				sleep 75
				ControlSend,, {Down}	, % "ahk_id " hKarteikarte
				sleep 100
			}

		}

	; Tastaturfokus (subFocus) prüfen und eventuell Fokus setzen
		If (Kkarte.subFocus <> Control) {   	; das Caret ist im Kürzelfeld, wenn eine neue Eingabezeile aufgemacht wurde
			ControlNN 	:= Kkarte.childs[Control].classNN
			hControl     	:= Kkarte.childs[Control].hwnd
			If !VerifiedSetFocus(ControlNN, hKarteikarte)
				If hControl
					VerifiedSetFocus("", hControl)
		}

return AlbisGetFocus("Karteikarte", Control, Kkarte) ? Kkarte.hCaret : 0  ; hwnd des Steuerelementes zurückgeben bei Erfolg
;~ return Kkarte.subFocus = Control ? GetHex(Kkarte.hCaret) : 0  ; hwnd des Steuerelementes zurückgeben bei Erfolg
}
AlbisKarteikarteFocus(Control, AutoActivate:=false) {                                               	;-- gezielter Schreibzugriff auf die Karteikarte

	; Control: es kann direkt die ClassNN des Steuerelementes übergeben werden oder einer der Alias-Bezeichnungen aus KK
	; KKx wird über die Funktion Init_Albis() bereitgestellt
	;
	; letzte Änderung 20.10.2021

		global Mdi:= Object(), KK
		static hKarteikarte

	; Alias-Bezeichnung wird durch ClassNN des Steuerelementes ersetzt
		Control := KK.HasKey(Control) ? KK[Control] : Control

	; prüft und zeigt die Karteikarte an, falls etwas anderes beim Patienten angezeigt wird
		If AutoActivate {
			AlbisActivate(2)
			If !AlbisKarteikarteZeigen()
				return 0
		}

	; ermittelt das richtige hwnd des Steuerelementes der Karteikarte
		hKarteikarte := Controls("#327702", "ID", AlbisMDIChildGetActive())

	; Tastaturfokus auf den Karteikartenbereich setzen und Tastenfolge senden
		ControlFocus,, % "ahk_id " hKarteikarte
		sleep 100
		fctrl := AlbisKarteikarteUnfocus(hKarteikarte)            	; behandelt auch Ausnahmen
		If (fctrl = Control)                                                     	; Eingabefokus steht an der richtigen Stelle
			return GetHex(hKarteikarte)
		else if (RegExMatch(fctrl, "i)^(Edit|RichEdit)")) {       	; der Eingabefokus liegt in einem anderen Editfeld
			SendInput, {Escape}                                          	; nochmals Esape senden
			sleep 100
		}
		SendInput, {End}                                                   	; mit End springt man ans Ende der Karteikarte
		Sleep 50
		SendInput, {End}
		Sleep 100

	; Sicherheitsüberprüfung
		fctrl := ""
		ControlGetFocus, fctrl, % "ahk_id " hKarteikarte
		while (!RegExMatch(fctrl, "i)^(Edit|RichEdit)") && A_Index < 6) {
			SendInput, {Down}                                            	; Pfeil runter - ermöglicht die Tast
			sleep 200
			ControlGetFocus, fctrl, % "ahk_id " hKarteikarte
		}

	; angefordertes Steuerelement fokussieren
		If !(res := VerifiedSetFocus(Control, hKarteikarte))
			SciTEOutput(A_ThisFunc ": Setzen des Eingabefocus auf " Control " ist fehlgeschlagen.")
		ControlGetFocus, fctrl, % "ahk_id " hKarteikarte

return ((fctrl = Control) ? GetHex(hKarteikarte) : 0)
}
AlbisKarteikarteUnfocus(hKarteikarte, retval:="fcontrol") {                                       	;-- gehört zu AlbisKarteikarteFocus()

	; versucht den Tastaturfokus aus der Karteikarte zu entfernen
	; bisher nicht gespeicherte Eingaben werden erkannt und es wird versucht die Eingabe abzuschliessen
	; letzte Änderung: 01.07.2022

	static DlgTitle := "ALBIS ahk_class #32770"
	static DlgText := ["Die aktuelle Zeile wurde nicht gespeichert", "Änderungen des Befundkürzels"]

	AlbisGetFocus("Karteikarte",, Kkarte)
	If Kkarte.subFocus  {

	  ; TAB wenn Inhalte in der Zeile geschrieben wurden, Escape wenn leer
		ControlSend,, % Kkarte.childs.Inhalt.Text ? "{TAB}" : "{Escape}", % "ahk_id " Kkarte.hCaret  ; in diesem Steuerelement befindet sich der aktuelle Eingabefokus

	  ; wartet bis sich der Eingabefokus verändert hat
		while (GetFocusedControlHwnd() = Kkarte.hCaret) && (A_Index <= 20) {
			If WinExist(DlgTitle, DlgText.1)
				break
			Sleep 20
		}

	 ; neuen Eingabefokus ermitteln
		AlbisGetFocus("Karteikarte",, Kkarte)

	 ; Fenster abfangen
		If WinExist(DlgTitle, DlgText.1)
			If VerifiedClick("Ja", DlgTitle, DlgText.1,, true) {

				ControlFocus, % "ahk_id " Kkarte.hCaret
				ControlSend,, {TAB}, % "ahk_id " Kkarte.hCaret
				while (GetFocusedControlHwnd() = Kkarte.hCaret) && (A_Index <= 10) {
					If WinExist(DlgTitle, DlgText.2)     ; Befundkürzel Abbruch
						break
					Sleep 20
				}

			  ; Die aktuelle Zeile wurde nicht gespeichert
				while (WinExist(DlgTitle, DlgText.2) && A_Index < 4) {
					VerifiedClick("OK", DlgTitle, DlgText.2,, true)
					SendInput, {Escape}
					sleep, 150
					while (WinExist(DlgTitle, DlgText.1) && A_Index < 4)
						VerifiedClick("Nein", DlgTitle, DlgText.1,, true)
				}

			  ; Änderungen des Befundkürzels
				ControlGetFocus, focused, % "ahk_id " hKarteikarte
				while (RegExMatch(focused, "i)^(Edit|RichEdit)") && A_Index < 21) {
					SendInput, {TAB}
					sleep 100
					If (A_Index > 1)
						sleep 50
					ControlGetFocus, focused, % "ahk_id " hKarteikarte
				}

			}
	}

	AlbisGetFocus("Karteikarte",, Kkarte)

return Kkarte[retval]
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
					IniWrite, % StrUtf8BytesToText("1|" obj.PatID "|" obj.PatName "|" obj.Behandlungstage "|" obj.Ziff1 "|" obj.Ziff2), % AbrechnungslistenPfad, Patientenliste, % "Patient" Index
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
		IniWrite, % StrUtf8BytesToText("1|" obj.PatID "|" obj.PatName "|" obj.Behandlungstage "|" obj.Ziff1 "|" obj.Ziff2), % AbrechnungslistenPfad, Patientenliste, % "Patient" Index
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

							If  (cFocus = "Edit3")
								break

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
							else if (CFocus <> "RichEdit20A1") 			; Eingabebereich gefunden
								break

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
		Zeilendatum := AlbisZeilenDatumLesen()
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
; 12
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
		;~ SciTEOutput(fclassnn)
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
		;~ SciTEOutput(fclassnn)
		If (classnn <> fclassnn){
			;BlockInput, Off
			return 0
		}
	}
	;BlockInput, Off
	sleep 200

return hParent
}
; 13
AlbisVordatierer() {                                                                                                	;-- vordatieren ohne Reue

  ; nimmt das Datum der aktuellen Karteikartenzeile
  ; letzte Änderung: 26.06.2022

  ; Hinweisdialog schließen
	If WinExist("ahk_class #32770", "Datum darf nicht") {
		VerifiedClick("OK", "ALBIS ahk_class #32770", "Datum darf nicht")
		sleep 200
	}

  ; aktuelles Zeilendatum auslesen
	Kkarte := AlbisGetFocus()
	VorDatierDatum := Kkarte.childs.Datum.text
	If !RegExMatch(VorDatierDatum, "\d{2}\.\d{2}\.\d{4}") {
		PraxTT("Das Karteikartendatum konnte für die Vordatierung nicht ausgelesen werden!", "2 1")
		return 0
	}

  ; Programmdatum auf das aktuelle Zeilendatum ändern
	PraxTT("Vordatierung durchführen...", "20 1")
	letztesDatum := AlbisSetzeProgrammDatum(VorDatierDatum)
	VerifiedSetFocus("", "", "", Kkarte.hFocus)
	Sleep % IsRemoteSession() ? 100 : 50
	SendInput, {Tab}
	Sleep % IsRemoteSession() ? 600 : 200
	ControlGetFocus, cFocused, % "ahk_class OptoAppClass"
	If RegExMatch(cFocused, "(i)Edit|RichEdit)") {
		SendInput, {Escape}
		Sleep % IsRemoteSession() ? 600 : 200
	}
	If !IsRemoteSession() {
		AlbisSetzeProgrammDatum(letztesDatum)
		Sleep 100
	}
	PraxTT("Vordatierung auf den " VorDatierDatum " ist erfolgt.", "2 1")

return 1
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; FORMULARE                                                                                                                                                                                  	(16)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisDruckeBlankoFormular  	(02) AlbisRezeptHelfer                	(03) AlbisRezeptHelferGui            	(04) AlbisRezeptFelderLoeschen
; (05) AlbisRezeptFelderAuslesen       	(06) AlbisRezeptSchalterLoeschen 	(07) AlbisDruckePatientenAusweis 	(08) AlbisRezeptAutIdem
; (09) AlbisRezept_DauermedikamenteAuslesen                                    	(10) AlbisFormular                       	(11) AlbisHautkrebsScreening
; (12) AlbisFristenRechner                 	(13) AlbisFristenGui                   	(14) IfapVerordnung                   	(15) AlbisWeitereMedikamente
; (16*) AlbisVerordnungsplan         	(17*) AlbisAusfuellhilfe                  	(18) AlbisGVU
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
AlbisRezeptHelferGui(dbfile) {                                                                                  	;-- Schnell/Standardrezepte für bestimmte Diagnosen/Hilfsmittel mit einem Klick

	; blendet ein kleines Listbox-Steuerelement in einem Kassenrezeptformular ein
	; diese Funktion braucht die Bibliothek >> lib\class_cJSON.ahk <<
	; letzte Änderung: 22.02.2022 - funktioniert auch bei Privatrezepten

	; VARIABLEN                                                                                                                                                        	;{

		global Rezept, RzFeld, RzZusatz, RzSchalter
		global hMuster16, hMuster16O, hRezeptHelfer, RH, RHBtn1, Mdi, RState
		global RHRetry 			:= 0
		global Muster16		:= Object()

		static DDLValues, RezeptWahl, wahl, isVisible, hMdiChild, APos, RPos, TPos, CtrlList, hAfxFOV, WerbungH
		static dX, dY, dW, dH, dxE, lpX, NewX, NewY, NewW, NewH
		static dbfileLastEdit 	:= 0	                       	; enthält die letzte Dateiänderungszeit
		static DDLWidth  		:= 380
		static SchnellRezept                                	; speichert die Nummer des zuletzt ausgewählten Schnellrezeptes
		static Rezepte                                        	; Objekt enthält den Datensatz für die Verordnungen


		RezeptIstZentriert 	:= false
		WerbungH        	:= 0

		If !IsObject(Rezept) {
			Rezept := {	"Kasse"	: { "Feld"		: [2,9,16,23,30,37]
											, 	"Zusatz"		: [3,7,11,15,19,23]
											, 	"Menge"   	: [1,8,15,22,26,36]
											, 	"Dosis"		: [[3,4,5,6], [10,11,12,14], [17,18,19,20], [24,25,26,27], [31,32,33,34], [38,39,40,41]]
											,	"Dj"         	: [5,10,15,20,25,30]
											, 	"Schalter"	: {48:"BVG", 49:"Hilfsmittel", 50:"Impfstoff", 51:"Sprechstunden-Bedarf", 52:"Heilmittel", 53:"BTM", 54:"OTC"}
											, 	"classnn" 	: "Muster 16 ahk_class #32770"}
							,	"Privat"	: {"Feld"	   	: [2,8,14,20,26,32]
											, 	"Zusatz"		: [1,5,9,13,17,21]
											, 	"Menge"   	: [1,7,13,19,25,31]
											, 	"Dosis"		: [[3,4,5,6], [9,10,11,12], [15,16,17,18], [21,22,23,24], [27,28,29,30], [33,34,35,36]]
											,	"Dj"         	: [5,10,15,20,25,30]
											, 	"classnn" 	: "Privatrezept ahk_class #32770"}}
			RzSchalter := Rezept.Kasse.Schalter
			RzFeld    	:= Array()
			RzZusatz 	:= Array()
		}

		If (!WinExist(Rezept.Kasse.classnn) && !WinExist(Rezept.Privat.classnn)) {
			RState   	:= ""
			RHRetry 	:= 0
			SetTimer, RezepthelferUpdate, Off
			return
		}

		If RState
			return
	;}

	; Fensterdaten zuordnen
		If (hMuster16 := WinExist(Rezept.Kasse.classnn))
			RState := "Kasse"
		else if (hMuster16 := WinExist(Rezept.Privat.classnn))
			RState := "Privat"
	; verhindert erneuten Aufruf
		else {
			RState := ""
			If (RHRetry < 30) {
				RHRetry ++
				fn := Func("AlbisRezeptHelferGui").Bind(dbfile)
				SetTimer, % fn, -50
			}
			else If (RHRetry >= 30) {
				RHRetry := 0
				SetTimer, RezepthelferUpdate, Off
			}
			return
		}

		If WinExist("Rezepthelfer ahk_class AutoHotkeyGUI") && (hMuster16O = hMuster16 )
			return
		hMuster16O	:= hMuster16

	; globale Variablen dem aktuellen Rezepttyp anpassen
		Loop % RzFeld.Count()
			RzFeld.Pop()
		For each, val in Rezept[RState].Feld
			RzFeld.Push(val)
		Loop % RzZusatz.Count()
			RzZusatz.Pop()
		For each, val in Rezept[RState].Zusatz
			RzZusatz.Push(val)

	; SCHNELLREZEPTDATEN EINLESEN
	; überprüft ob die Datei verändert wurde, liest die Daten dann neu ein
		FileGetTime, fileTime, % dbfile, M
		If (fileTime <> dbfileLastEdit) {
			dbfileLastEdit	:= fileTime
			DDLValues 	:= "Auswahl eines Schnellrezeptes...|neues Schnellrezept anlegen"
			If FileExist(dbfile) {
				JSONstr	:= FileOpen(dbfile, "r", "UTF-8").Read()
				Rezepte	 	:= cJSON.Load(JSONstr)
				For i, val in Rezepte
					DDLValues .= "|" val.Bezeichner
			} else {
				PraxTT("keine Schnellrezepte für Rezeptverordnungen vorhanden!", "0 2")
				;Rezepte := Object()
			}
		}

	; ERMITTELN VON STEUERELEMENTHANDLES, STEUERELEMENT- UND FENSTERPOSITIONEN                              	;{
	; gebraucht für die Positionierung des Rezeptfensters innerhalb von Albis
		res                   	      	:= Controls("", "Reset", "")
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

							APos		:= GetWindowSpot(AlbisWinID())
							RPos		:= GetWindowSpot(hMuster16)

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
			Control         	, Hide                         ,,, % "ahk_id " Muster16.StaticW
			ControlMove   	,,,,, % dY+dH-cY       	, % "ahk_id " Muster16.DMed
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

	; GUI - ZUSÄTZLICHE STEUERELEMENTE IM REZEPT ANZEIGEN                                                                             	;{
		ControlGetPos, cx,,,               	,, % "ahk_id " Muster16.Drucken           	; unterhalb von Drucken
		ControlGetPos, dX, dY, dW, dH	,, % "ahk_id " Muster16.ArzneiDB          	; Schnell-Rezeptposition unterhalb d. Steuerelementes Arzneimitteldatenbank angezeigt

		try {
			Gui, RH: New, -Caption -DPIScale +ToolWindow +HWNDhRezeptHelfer +E0x00000004 -0x04020000 Parent%hMuster16%
		}
		catch {
			sleep 100
			RHRetry ++
			RState := ""
			SetTimer, RezepthelferUpdate, Off
			If (RHRetry < 30)
				AlbisRezeptHelferGui(dbfile)
			else
				RHRetry := 0
			return
		}
		Gui, RH: Margin, 0, 0
		Gui, RH: Color	, cFFFFFF
		Gui, RH: Font, s8 q5 cBlack, MS Sans Serif

		Gui, RH: Add, Button , % "x0                                	gRezeptVOPlan                                                                       	"  	, % "Verordnungsplan drucken"
		Gui, RH: Add, Button , % "x300  y0                     	gRezeptLeeren       	vRHBtn1     	HWNDhRHelferBtn1             	"	, % "alle Felder leeren"            	;
		Gui, RH: Add, DDL 	, % "x500 y1 w" DDLWidth " 	gRezeptAusgewaehlt	vRezeptwahl HWNDhRezeptWahl AltSubmit"	, % DDLValues

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


	;}

return

RezeptVOPlan:                	;{

	Gui, RH: Destroy
	WinActivate, % RState

	VerifiedCheck(""	,,, Muster16.VOP)
	VerifiedCheck(""	,,, Muster16.EVO)
	VerifiedClick(	""	,,, Muster16.Drucken)

	RState := ""
	WinWait, % "Auswahl weiterer Medikamente ahk_class #32770",, 15
	If !ErrorLevel
		AlbisVerordnungsplan()

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

					MsgBox, 0x1004, % "Addendum Rezepthelfer", % RezeptDaten.Medikamente.MaxIndex() " ausgefüllte Medikamentenfelder sind vorhanden.`nMöchten Sie dieses Rezept als Schnellrezept speichern?"
					IfMsgBox, No
						return

					SchnellRezeptBezeichner:
					InputBox, Bezeichner, % "Addendum Rezepthelfer", % "Geben Sie dem neuen Schnellrezept einen eindeutigen Namen"
					If !Bezeichner {
						MsgBox, 0x1000, % "Addendum Rezepthelfer", % "Sie haben keine Bezeichnung für ihr Schnellrezept vergeben!"
						goto SchnellRezeptBezeichner
					}

					Rezepte.Push({	"Beschreibung"	: " "
								    	, 	"Bezeichner"  	: Bezeichner
								    	, 	"Medikamente"	: RezeptDaten.Medikamente
								    	,	"Zusaetze"      	: RezeptDaten.Zusaetze
								    	,	"Optionen"      	: ""
								    	,	"RezeptTyp"      	: RezeptDaten.RezeptTyp})

					FileOpen(dbfile, "w").Write(JSON.Dump(Rezepte))

				}

		}

	; AUSGEWÄHLTES SCHNELLREZEPT ANZEIGEN
		else If (RezeptWahl > 2)		{

		; Schnellrezept Nr zwischenspeichern
			SchnellRezept := RezeptNr := RezeptWahl - 2

		; Medikamentenfelder und Schalter werden gelöscht
			AlbisRezeptFelderLoeschen(hMuster16)

		; Medikamente eintragen
			For i, medText in Rezepte[RezeptNr].medikamente 			{

				medText := AlbisRezeptReplaceDates(medText)

				If RegExMatch(medText, "^\s*(?<cc>[\d]+)\s*x", Med) {
					medText := RegExReplace(medText, Med)
					ControlSetText, % "Edit" Rezept[RState].Menge[i], % RegExReplace(Med, "\s*x"), % "ahk_id " hMuster16
				}

				If RegExMatch(medText, "i)\s*Diagnose:")
					VerifiedCheck("Button"Rezept[RState].Dj[i],,, hMuster16, true)

				If InStr(medText, "...")		{
					focusFeld	:= "Edit" Rezept[RState].Feld[i]
					focusPos	:= InStr(medText, "...")
				}

				ControlSetText, % "Edit" Rezept[RState].Feld[i], % medText, % "ahk_id " hMuster16
			}

		; Schalter entsprechend setzen
			If (RState = "Kasse") {
				RezeptTyp:= Rezepte[RezeptNr].RezeptTyp
				If (StrLen(RezeptTyp) > 0)
					For Nr, SchalterBezeichnung in Schalter
						If InStr(SchalterBezeichnung, RezeptTyp)
							VerifiedCheck("Button" Nr, "", "",  hMuster16, 1), 	break
			}


		; Rezepthelfer-Gui nach vorne holen
		   	WinActivate   	 , % "ahk_id " hRezeptHelfer
			WinSet, Top   	,, % "ahk_id " hRezeptHelfer
        	WinSet, Redraw	,, % "ahk_id " hRezeptHelfer

		; Tastaturfokus auf das Feld setzen welches '...' enthält - Nutzer möchte dort noch eine weitere Eingabe eintragen
			If (StrLen(FocusFeld) > 0) {
				VerifiedSetFocus(FocusFeld, "", "", hMuster16)
				ControlGetText	, text, % FocusFeld, % "ahk_id " hMuster16
   				SendMessage 0xB1, % focusPos, % focusPos+StrLen(focusPos), % FocusFeld, % "ahk_id " hMuster16 ; EM_SETSEL
			}
	}

return ;}

RezeptLeeren:                	;{
	MsgBox, 0x1004, Addendum für Albis on Windows, Möchten Sie wirklich alle Rezeptfelder löschen?
	IfMsgBox, No
		return
	AlbisRezeptFelderLoeschen(hMuster16)
return ;}

RezepthelferUpdate:      	;{
	If !WinExist(Rezept[RState].classnn)
		AlbisRezeptHelferGuiClose()
return ;}
}
;3.1 ;{
AlbisRezeptHelferGuiClose() {

	global RState, RHRetry, RH, hRezeptHelfer

	RHRetry 	:= 0
	RState   	:= ""

	Gui, RH: Destroy

	SetTimer, RezepthelferUpdate, Off
	IfWinExist, Addendum Rezepthelfer ahk_class AutoHotkeyGUI
		WinClose, Addendum Rezepthelfer ahk_class AutoHotkeyGUI

return
}
;3.2
GuiControlPos(GuiName, ControlName) {
	GuiControlGet, cp, % GuiName ": Pos", % ControlName
return {"X":cpX, "Y":cpY, "W":cpW, "H":cpH}
}
;3.3
AlbisRezeptReplaceDates(item) {

	item := RegExReplace(item, "TT\."   	, A_DD ".")
	item := RegExReplace(item, "MM\."	, A_MM ".")
	item := RegExReplace(item, "\.YYYY"	, "." A_YYYY)

	while RegExMatch(item, "(?<d>\d+)(?<p>[\+\-])(?<a>\d+)", X)
		item := StrReplace(item, X, (Xp="+" ? Xd+Xa : Xd-Xa))

return item
} ;}
;4
AlbisRezeptFelderLoeschen(hMuster16) {                                                                	;-- (primär) Hilfsfunktion für die Rezepthelfer Gui

	global RzFeld, RzZusatz, RState
	static medzusaetze := "Medikamentenzusätze ahk_class #32770"

  ; löscht die Zusatztexte
	Loop, % RzZusatz.MaxIndex() 	{

		ControlGetText, cText, % "Button" RzZusatz[A_Index], % "ahk_id " hMuster16
		If InStr(cText, "zus")		{
			VerifiedClick("Button" RzZusatz[A_Index], "ahk_id " hMuster16)
			WinWait, % medzusaetze,, 5
			If !ErrorLevel	{
				ControlSetText, % "Edit1", % "", % medzusaetze
				ControlSetText, % "Edit2", % "", % medzusaetze
				VerifiedClick("Button2", medzusaetze)
				WinWaitClose, % medzusaetze,, 5
				If ErrorLevel	{
					while WinExist(medzusaetze)
						MsgBox, Bitte das Fenster 'Medikamentenzusätze' schließen!
				}
				ControlSetText, % "Button" RzZusatz[A_Index], % "...", % "ahk_id " hMuster16
			}
		}

	}

  ; löscht die Medikamentenfelder
	Loop, % RzFeld.MaxIndex() {
		nr := A_Index
		ControlSetText, % "Edit" RzFeld[nr],, % "ahk_id " hMuster16
		Loop, 4
			ControlSetText, % "Edit" RzFeld[nr]+A_Index,, % "ahk_id " hMuster16
	}

  ; löscht die Schalter
	If (RHState = "Kasse")
		AlbisRezeptSchalterLoeschen(hMuster16)

}
;5
AlbisRezeptFelderAuslesen(hMuster16) {                                                                 	;-- (primär) Hilfsfunktion für die Rezepthelfer Gui

	global RzFeld, RzZusatz, RzSchalter, RState

	RezeptDaten 	:= {"Medikamente":[], "Zusaetze":[], "RezeptTyp":""}
	Loop, % RzZusatz.MaxIndex() 	{

		ControlGetText, val	, % "Edit"    	RzFeld[A_Index]    	, % "ahk_id " hMuster16
		ControlGetText, cText, % "Button" 	RzZusatz[A_Index]	, % "ahk_id " hMuster16

		If (StrLen(val) > 0)
			RezeptDaten.Medikamente.Push(val)

		If InStr(cText, "zus") {
			VerifiedClick("Button" RzZusatz[A_Index], "ahk_id " hMuster16)
			ControlGetText, val, % "Edit1", % "", % medzusaetze
			RezeptDaten.Zusatz.Push(val)
			VerifiedClick("Button2", medzusaetze)
			WinWaitClose, % medzusaetze,, 5
			If ErrorLevel 	{
				while WinExist(medzusaetze)
					MsgBox, Bitte das Fenster 'Medikamentenzusätze' schließen!
			}
		}
	}

; bricht nach dem ersten gesetzten Schalter im Rezept ab
	If (RState = "Kasse")
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

	global RzSchalter, RState

	If (RState = "Kasse")
		For Nn, SchalterBezeichnung in RzSchalter
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
		static RzFeld      	:= [2,8,14,20,26,32]    	; Edit
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
			ControlGetText, Medikament	, % "Edit"   	RzFeld[A_Index]	, % "ahk_id " hMuster16
			ControlGet, hAutIdem, hwnd,, % "Button" 	Idemfeld[A_Index]  	, % "ahk_id " hMuster16
			SendMessage, 0xF2, 0, 0,, % "ahk_id " hAutIdem      ; BM_GETSTATE
			isChecked := ErrorLevel
			If RegExMatch(Nacht, "[A]") && !isChecked && (StrLen(Medikament) > 0)				{
				Control, Check,,, % "ahk_id " hAutIdem
				t.= Medikament "`n"
				ControlGetPos, ax,ay,aw,ah, % "Button" Idemfeld[A_Index]	, % "ahk_id " hMuster16
				ControlGetPos, rx, ry, rw, rh, % "Edit"    RzFeld[A_Index]	, % "ahk_id " hMuster16
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
AlbisRezeptHook(hMuster16) {                                                                               	;-- gehört zu AlbisRezeptAutIdem - Erkennung von Änderungen im Rezeptformular

	global RezeptHook, RezeptProcAdr, hRezept
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
	RezeptHook 	:= SetWinEventHook( 0x800E	, 0x800E	, 0, RezeptProcAdr, AlbisPID, idThread, 0x3)
	RezeptHook 	:= SetWinEventHook( 0x0003	, 0x0003	, 0, RezeptProcAdr, AlbisPID, idThread, 0x3)
	RezeptHook 	:= SetWinEventHook( 0x8001	, 0x8001	, 0, RezeptProcAdr, AlbisPID, idThread, 0x3)	; RezeptWindow is Closed
	RezeptHook 	:= SetWinEventHook( 0x800B	, 0x800B	, 0, RezeptProcAdr, AlbisPID, idThread, 0x3)	; RezeptWindow is Closed

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

	static Formular := {"eHKS_ND"                       	: {"command" : 34505, "WinTitle" : "Hautkrebsscreening - Nichtdermatologe ahk_class #32770"}
								,"GVU"                              	: {"command" : 33212, "WinTitle" : "Muster 30 ahk_class #32770"}
								,"Cave"                              	: {"command" : 32778, "WinTitle" : "Cave! von ahk_class #32770"}
								,"Abrechnung_vorbereiten"    	: {"command" : 32944, "WinTitle" : "Abrechnung KVDT vorbereiten ahk_class #32770"}}

	If !Formular.HasKey(Formularbezeichnung)	{
		throw Exception("Formularbezeichnung: " Formularbezeichnung " ist unbekannt!")
		return 0
	}

	dhw := A_DetectHiddenWindows, dht := A_DetectHiddenText
	tmm := A_TitleMatchMode, tmms := A_TitleMatchModeSpeed
	DetectHiddenWindows   	, On
	DetectHiddenText          	, On
	SetTitleMatchMode        	, 2
	SetTitleMatchMode        	, slow

	hWin1:= Albismenu(Formular[Formularbezeichnung].command, Formular[Formularbezeichnung].WinTitle)
	RegExMatch(Formular[Formularbezeichnung].WinTitle, "i)^\s*(?<out>.*?)\s*ahk_class\s*(?<class>.*)$", with)
	without := RegExReplace(without, "\s", "\s+")
	without := RegExReplace(without, "\-", "\-")
	while (A_Index < 40) {
		hwnd := GetHex(DllCall("GetLastActivePopup", "uint", AlbisWinID()))
		If (hwnd <> GetHex(AlbisWinID())) {
			t := " [" A_ThisFunc "]           (" SubStr("0" A_Index, -1) ")  " (atitle := WinGetTitle(hwnd)) " {" hwnd "}"
			If RegExMatch(atitle, "i)^" without) {
				hWin := hwnd
				break
			}
			;~ SciTEOutput(t "`n")
		}
		sleep 50
	}

	DetectHiddenWindows   	, % dhw
	DetectHiddenText          	, % dht
	SetTitleMatchMode        	, % tmm
	SetTitleMatchMode        	, % tmms

return hWin ? hWin : "-1"
}
;11
AlbisHautkrebsScreening(Pathologie:="opb", PlusGVU=true, closeWin=true, hWin=0 ) {      	;-- befüllt das eHautkrebsScreening (nicht Dermatologen) Formular

	; das Hautkrebsscreeningformular für Nichtdermatologen wurde mehrfach geändert in den letzten Quartalen
	; die ClassNN der Buttons hat sich dabei jedesmal geändert, damit wurden nicht die richtigen Buttons angesprochen und somit war das Formular nicht korrekt ausgefüllt
	; oder ließ sich nicht speichern da sich auch für den Speichern Button die ClassNN geändert hatte
	; letzteres versuche ich nun mittels Übergabe des Namen/Text des Speichern-Buttons treffsicherer zu machen
	; prinzipiell müsste diese Funktion wesentlich flexibler gestaltet werden

	; letzte Änderung: 03.07.2022

		static B := "Button"

	; Fenstermatchmodes sichern
		dhm	:= A_DetectHiddenWindows
		dht	:= A_DetectHiddenText
		stm	:= A_TitleMatchMode
		stms	:= A_TitleMatchModeSpeed

	; Fenstermatchmodes ändern
		;~ DetectHiddenWindows   	, On
		;~ DetectHiddenText          	, On
		;~ SetTitleMatchMode        	, 2
		;~ SetTitleMatchMode        	, slow

	; wenn kein Handle übergeben wurde
		If !hWin {
			WinActivate    	, % (eHKS_Title := "Hautkrebsscreening ahk_class #32770")
			WinWaitActive	, % eHKS_Title,, 2
			hHKS := WinExist(eHKS_Title)
			If !hHKS {
				PraxTT("Das Formular 'Hautkrebsscreening - Nichtdermatologe' ist nicht geöffnet!`nDie Funktion wird abgebrochen.", "3 1")
				return 0
			}
		}
		else
			hHKS := hWin

	; prüfen des eHKS-Dialog Handles
		while !RegExMatch(WTitle := WinGetTitle(hHKS), "i)Hautkrebsscreening\s+\-\s+Nichtdermatologe") && (A_Index <= 10) {
			If (A_Index > 1)
				sleep 100
			hHKS := WinExist(eHKS_Title)
			hHKS := WinExist("Hautkrebsscreening - Nichtdermatologe ahk_class #32770")
			If !hHKS
				WinGet, hHKS, ID, % "Hautkrebsscreening - Nichtdermatologe ahk_class #32770"
			WTitle := WinGetTitle(hHKS)
			SciTEOutput(A_Index ": " WTitle)

		}
		If !RegExMatch(WTitle := WinGetTitle(hHKS), "i)Hautkrebsscreening\s+\-\s+Nichtdermatologe")
			PraxTT("Addendum hat Schwierigkeiten das Hautkrebsscreeningdialog zu ermittlen", "4 1")

	; Dialog aktivieren
		WinActivate    	, % eHKS_Title
		WinWaitActive	, % eHKS_Title,, 2

	; markiert alle Verdachtsdiagnosen als nein
		If (Pathologie ~= "i)^\s*opB")		{
			    ;~ VerifiedClick(B "28", hHKS)	; nein - "Verdachtsdiagnose"  - warum Fehler wenn nicht ausgewählt
			If !VerifiedClick(B "10", hHKS)	; nein - "Malignes Melanom"
				fail .= ",Malignes Melanom"
			If !VerifiedClick(B "12", hHKS)	; nein - "Basalzellkarzinom"
				fail .= ",Basalzellkarzinom"
			If !VerifiedClick(B "14", hHKS)	; nein - "Spinozelluläres Karzinom"
				fail .= ",Spinozelluläres Karzinom"
			If !VerifiedClick(B "24", hHKS)	; nein - "Anderer Hautkrebs"
				fail .= ",Anderer Hautkrebs"
			If !VerifiedClick(B "26", hHKS)	; nein - "Sonstiger dermatologisch abklärungsbedürftiger Befund"
				fail .= ",Sonstiger abklärungsbedürftiger Befund"
			If !VerifiedClick(B "30", hHKS)	; nein - "Screening-Teilnehmer wird an einen Dermatologen überwiesen:"
				fail .= ",wird überwiesen"
			If !VerifiedClick(B "28", hHKS)	; nein - "Verdachtsdiagnose"  - warum Fehler wenn nicht ausgewählt
				fail .= ",Verdachtsdiagnose"
		}

	; Häkchen bei "gleichzeitige Gesundheitsvorsorge setzen"
		If PlusGVU {
			If !VerifiedClick(B "16", hHKS)
				fail .= ",gleichzeitige Gesundheitsvorsorge ja"
		}
		else {
			If !VerifiedClick(B "17", hHKS)
				fail .= ",gleichzeitige Gesundheitsvorsorge nein"
		}

		If fail
			SciTEOutput(A_ThisFunc ": " fail)

	; Wiederhestellen der Einstellungen
		DetectHiddenWindows   	, % dhm
		DetectHiddenText          	, % dht
		SetTitleMatchMode        	, % stm
		SetTitleMatchMode        	, % stms

	; "Speichern" und damit auch Schließen des Formulares
		If closeWin		{
			ResultClose := VerifiedClick("Speichern", hHKS,,, true)
			Winwait, % "Albis ahk_class #32770", % "Folgende Fehler sind bei der Plausibilitätsprüfung", 1
			If WinExist("Albis ahk_class #32770", "Folgende Fehler sind bei der Plausibilitätsprüfung")
				return -2

			; falls sich das Formular nicht speichern/schließen ließ, wird keine 1 als Erfolgsmeldung zurück senden
			If WinExist(eHKS_Title)
				return -3
		}

return fail ? -1 : 1
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

	 ; eAU Hinweis auisschalten, habe ich oft genug gelesen jetzt
		ControlMove, Static26,, 20, 10, 10, % "ahk_id " hAU
		Control, Hide,, Static26, % "ahk_id " hAU

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
IfapVerordnung(Medikament, printsilent:=true) {                                                    	;-- Medikament über ifap verordnen u. Dosishinweise (MS Word) ausdrucken

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
		VerifiedClick("Auswahl umkehren", "Auswahl weiterer Medikamente ahk_class #32770")
		VerifiedClick("OK", "Auswahl weiterer Medikamente ahk_class #32770",,, 3)

		WinWait, Vorlagen ahk_class #32770,, 10
		If !ErrorLevel {
			VerifiedClick("OK", "Vorlagen ahk_class #32770")
		}


		;~ ControlFocus, Listbox1, % "Auswahl weiterer Medikamente ahk_class #32770"
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
;18
AlbisGVU(dbg:=false, xFee:="") {                                                                           	;-- GVU/HKS abrechnen und Formular erstellen

	; letzte Änderung: 03.07.2022

		static rxDate := "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})"

	; doppelten Aufruf vermeiden
		If Addendum.Flags.AlbisGVUruns
			return

	; GVU Abrechnungsflags setzen damit Addendum Ruhe gibt         	;{
		Addendum.Flags.AlbisGVUruns	:= true
		Addendum.Flags.NoTimerDelay	:= true
	;}

	; zusätzliche Gebühren abrechnen                                          		;{
		xFee := RegExReplace(xFee, "\s")
		xFee := RegExReplace(xFee, "\-$")
	;}

	; fokussiertes Steuerelement (classNN) ermitteln                          	;{
		kbf := AlbisGetFocus()
		cChanged	:= cFocused := kbf.fcontrol
	;}

	; Zeilendatum auslesen                                                             		;{
		GVUDatum := AlbisZeilenDatumLesen(60, false)
		If !RegExMatch(GVUDatum, rxDate) {
			GVUDatum := AlbisZeilenDatumLesen(60, false)
			If !RegExMatch(GVUDatum, rxDate) {
				PraxTT("Das Zeilendatum konnte nicht ermittelt werden.#2Abbruch der Vorsorgeformularerstellung")
				Addendum.Flags.NoTimerDelay := false
				Addendum.Flags.AlbisGVUruns := false
				return 0
			}
		}
	;}

	; Gebühren eintragen                                                             		;{
		If IsObject(kbf := AlbisGetFocus()) {
			If (kbf.mainFocus = "Karteikarte" && kbf.subFocus = "Inhalt") {

			  ; sind bereits Ziffern eingetragen?
				cFeeO := cFee := ControlGetText("", "ahk_id " kbf.hCaret)
				cFee := RegExReplace(cFee, "\s")
				cFee := RegExReplace(cFee, "\-$") "-"

			 ; eingetragene Ziffern (cFee) und übergebene Ziffern (xFee) vergleichen um doppelte Eintragungen zu vermeiden
				If xFee && cFee {
					For each, fee in StrSplit(xFee, "-")
						If !RegExMatch(cFee, fee "\-")        	; anhängen wenn nicht gefunden
							cFee .= fee "-"
				}
				cFee := RegExReplace(cFee, "\-$")

			  ; bisher bestehende Eintragung löschen (verschiedene Möglichkeiten)
				If cFeeO  {

				  ; Eingabefocus kontrollieren
					If !CheckFocus(kbf.hCaret)
						return 0

				  ; alte Eingaben entfernen
					ControlSend,, {Home}                                    	, % "ahk_Id " kbf.hCaret
					sleep 100
					ControlSend,, {LShift Down}{End}{LShift Up}	, % "ahk_Id " kbf.hCaret
					sleep 100
					ControlSend,, {BS}                                           	, % "ahk_Id " kbf.hCaret
					sleep 0
					sleep 50

					If (cText := ControlGetText("", "ahk_id " kbf.hCaret))
						Loop % StrLen(cFeeO) {
							ControlSend,, {BS}                                  	, % "ahk_Id " kbf.hCaret
							sleep 20
						}

				  ; noch Text vorhanden?
					If (cText := ControlGetText("", "ahk_id " kbf.hCaret)) {

					  ; Eingabefocus kontrollieren (beendet die Funktion wenn nicht wiederherstellbar)
						If !CheckFocus(kbf.hCaret)
							return 0

                      ; alle Eintragungen entfernen
						VerifiedSetText("", "", kbf.hCaret)

					  ; noch Text vorhanden?
						If (cText := ControlGetText("", "ahk_id " kbf.hCaret)) {

						 ; Eingabefocus kontrollieren
							If !CheckFocus(kbf.hCaret)
								return 0

							ControlSend,, {End}                                   	, % "ahk_Id " kbf.hCaret
							sleep 100
							Loop % StrLen(cText) {
								If (A_Index>1)
									sleep 5
								SendInput {BS}
							}
						}

					}

					Sleep 0
					Sleep 50

				}

			  ; alles eintragen
				toSend := (cFee ? RegExReplace(cFee, "\-$") "-" : "") "01732-01746"
				For i, char in StrSplit(toSend) {
					SendRaw, % char
					Sleep 5
				}
				Sleep 50
				SendInput, {Tab}
				WaitFocusChanged(kbf.hCaret, kbf.hWin, "`nauf Abschluß der Gebührenübernahme")

			}
		}
	;}

	; EBM/KRW Fenster abwarten                                                     	;{
		PraxTT("Warte Fenster EBM/KRW Hinweise ab", "10 1")
		while !WinExist("Prüfung EBM/KRW ahk_class OptoAppClass") && (A_Index < 20)
			sleep 100

		If WinExist("Prüfung EBM/KRW ahk_class OptoAppClass") {

			AlbisActivate(1)
			PraxTT("Warte das Schließen der EBM/KRW Hinweise ab", "10 1")

			Loop 2 {
				AlbisMDIChildWindowClose("Prüfung EBM/KRW")
				while WinExist("Prüfung EBM/KRW ahk_class OptoAppClass") && (A_Index < 20)
					sleep 100
				If !WinExist("Prüfung EBM/KRW ahk_class OptoAppClass")
					break
			}

		}

		PraxTT("", "off")

		;~ If (hMDIChild := AlbisMDIChildHandle("Prüfung EBM/KRW")) {
			;~ SciTEOutput("hMDI: " hMDIChild)
			;~ ControlSend,, {Escape}, % "ahk_id " hMDIChild
			;~ sleep 400
		;~ }

		If dbg
			SciTEOutput(" [" A_ThisFunc "]               (A)  Zeilendatum " GVUDatum  "`n")
	;}

	; Programmdatum ändern für das Hautkrebsscreening-Formular 	;{
		If !RegExMatch((ProgrammDatum := AlbisSetzeProgrammDatum(GVUDatum, true, dbg)), rxDate)
			ProgrammDatum := A_DD "." A_MM "." A_YYYY
		WinWaitClose, % "Programmdatum einstellen ahk_class #32770",, 1
	;}

	; Hautkrebsscreening ausfüllen                                                     	;{
		If (hWin := AlbisFormular("eHKS_ND"))
			res := AlbisHautkrebsScreening("opb", true, true, hWin)

		If dbg
			SciTEOutput(" [" A_ThisFunc "]               (B)  ProgrammDatum = " ProgrammDatum "`n")
	;}

	; Programmdatum zurücksetzen                                                  	;{
		AlbisSetzeProgrammDatum(ProgrammDatum, true, dbg)
	;}

	; Eingabefokus aus der Karteikarte nehmen                                 	;{
		If AlbisKarteikarteAktivieren() {
			SendInput, {Escape}
			sleep 200
			SendInput, {Escape}
		}
	;}

	; GVU String für Cave ins Clipboard                                            	;{

	  ; stellt den CAVE String zusammen
		GVString := RegExMatch(toSend "-", "01740(\-)") ? "01740, " : ""
		GVString .= RegExMatch(toSend "-", "01747(\-)") ? "A, " : ""
		GVString .= RegExMatch(toSend "-", "01731(\-)") ? "GVU/HKS/KVU " : "GVU/HKS "
		GVString .= StrSplit(GVUDatum, ".").2 "^" SubStr(StrSplit(GVUDatum, ".").3, -1)
	;}

    ; helfen den GVU String in den CaveVon Dialog zu schreiben
		AlbisCaveGVU(GVString, GVUDatum)

	  ; Flags löschen
		Addendum.Flags.NoTimerDelay := false
		Addendum.Flags.AlbisGVUruns := false

return 1
}
CheckFocus(hFocusedControl) {                                                                                	;-- Eingabefokus kontrollieren und evtl. neu setzen

	If (GetFocusedControlHwnd() != hFocusedControl)
		If !VerifiedSetFocus("","","", hFocusedControl) {
			PraxTT("Der Eingabefocus ist nicht mehr vorhanden. Die Eintragung muss abgebrochen werden", "3 1")
			return 0
	}

return 1
}
WaitFocusChanged(fcontrol, fwinHwnd, msg:="", wait:=6) {                                   	;-- wartet bis der Tastaturfokus sich geändert hat

	slTime 	:= 50
	mxRnds	:= wait*1000 ;/slTime

	ControlGetFocus, fchanged, % "ahk_id " fwinHwnd
	while (fcontrol = fchanged && A_Index < mxRnds) {
		If (A_Index > 1)
			Sleep % slTime
		ToolTip, % "warte noch " (mxRnds)-(A_Index*slTime) " ms" . msg
		ControlGetFocus, fchanged, % "ahk_id " fwinHwnd
	}

return fcontrol<>fchanged ? true: false
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; WARTEZIMMER                                                                                                                                                                           	(06)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisWZPatientEntfernen       	(02) AlbisWZOeffnen	                 	(03) AlbisWZKommentar             	(04) AlbisWZTabSelect
; (05) AlbisWZListe                         	(06) AlbisWZHeader
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

return hWZ		; oder war schon geöffnet
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
				return -7   ; handle des Wartezimmers

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
		HDInfo := AlbisWZHeader()

	; Syslistview auslesen
		ControlGet, csv, List,,, % "ahk_id " hLV
		csv := RegExReplace(csv, "[\t]", "|")

	; Tabelle in ein key:value Objekt umwandeln
		wzTable := Array()
		For LNr, line in StrSplit(csv, "`n") {
			cols := Object()
			For ENr, element in StrSplit(line, "|") {
				key := HDInfo[ENr]
				cols[key] := element
			}
			wzTable.Push(cols)
		}

return wzTable
}
;6
AlbisWZHeader() {	               																					;-- liest die Spaltennamen des aktuellen Wartezimmers

    Static MAX_TEXT_LENGTH := 260
         , MAX_TEXT_SIZE := MAX_TEXT_LENGTH * (A_IsUnicode ? 2 : 1)
		 , headerNames:= {"Notfall":"Notfall"	     				     				;1
									, "Name":"Name"											;2
									, "Anwesenheit":"Anwesenheit"						;3
									, "Datum":"Datum"											;4
									, "Kommt um":""												;5
									, "Geht um":""												;6
									, "Wartezeit":"Wartezeit"									;7
									, "Zeit rel. Termin":"relZeit"								;8
									, "Pager":"Pager"											;9
									, "Beh.-zeit":"BZeit"											;10
									, "Kommentar":"Kommentar"							;11
									, "Gruppen":"Gruppen"									;12
									, "Geburtsdatum(Alter)":"Geb"}						;13

	HDInfo := []

	ControlGet, hHeader, Hwnd,, SysHeader321, % "ahk_id " hWZ:= AlbisMDIWartezimmerID()
    WinGet PID, PID, ahk_id %hHeader%

    ; Open the process for read/write and query info.
    ; PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_VM_OPERATION | PROCESS_QUERY_INFORMATION
		If !(hProc := DllCall("OpenProcess", "UInt", 0x438, "Int", False, "UInt", PID, "Ptr"))
				Return

    ; Should we use the 32-bit struct or the 64-bit struct?
		If (A_Is64bitOS)
			Try DllCall("IsWow64Process", "Ptr", hProc, "int*", Is32bit := true)
		Else
			Is32bit := True

		RPtrSize := Is32bit ? 4 : 8,  cbHDITEM := (4 * 6) + (RPtrSize * 6)

    ; Allocate a buffer in the (presumably) remote process.
		remote_item := DllCall("VirtualAllocEx", "Ptr", hProc, "Ptr", 0, "uPtr", cbHDITEM + MAX_TEXT_SIZE, "UInt", 0x1000, "UInt", 4, "Ptr") ; MEM_COMMIT, PAGE_READWRITE
		remote_text := remote_item + cbHDITEM

    ; Prepare the HDITEM structure locally.
		VarSetCapacity(HDITEM, cbHDITEM, 0)
		NumPut(0x3, HDITEM, 0, "UInt") ; mask (HDI_WIDTH | HDI_TEXT)
		NumPut(remote_text, HDITEM, 8, "Ptr") ; pszText
		NumPut(MAX_TEXT_LENGTH, HDITEM, 8 + RPtrSize * 2, "Int") ; cchTextMax

    ; Write the local structure into the remote buffer.
		DllCall("WriteProcessMemory", "Ptr", hProc, "Ptr", remote_item, "Ptr", &HDITEM, "uPtr", cbHDITEM, "Ptr", 0)

		VarSetCapacity(HDText, MAX_TEXT_SIZE)

	; Get count of items
		SendMessage 0x1200, 0, 0,, ahk_id %hHeader% ; HDM_GETITEMCOUNT

	; Read every item and get its size
		Loop % (ErrorLevel != "FAIL" ? ErrorLevel : 0)
		{
			SendMessage, % (A_IsUnicode ? 0x120B : 0x1203), A_Index - 1, remote_item,, % "ahk_id " hHeader ; HDM_GETITEMW				-  Retrieve the item text.
			If (ErrorLevel == 1) 				 ; Success
				DllCall("ReadProcessMemory", "Ptr", hProc, "Ptr", remote_text, "Ptr", &HDText, "uPtr", MAX_TEXT_SIZE, "Ptr", 0)
			Else
				HDText := ""

			HDInfo.Push(HDText)
			;~ HDInfo.Push({"Text": HDText})
		}


    ; Release the remote memory and handle.
    DllCall("VirtualFreeEx", "Ptr", hProc, "Ptr", remote_item, "UPtr", 0, "UInt", 0x8000) ; MEM_RELEASE
    DllCall("CloseHandle", "Ptr", hProc)

Return HDInfo
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
; KARTEIKARTE | PATIENTENAKTE, ABRECHNUNG oder MENU                                                                                                         	(14)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAkteSchliessen               	(02) AlbisAkteOeffnen                	(03) AlbisAkteGeoeffnet              	(04) AlbisPatientAuswaehlen
; (05) AlbisInBehandlungSetzen       	(06) AlbisSucheInAkte                  	(07) AlbisKarteikarteZeigen        	(08) AlbisDialogOeffnePatient
; (09) AlbisKarteikarteAktiv               	(10) AlbisKarteikartenAnsicht        	(11) AlbisLeseDatumUndBezeichnung
; (12) AlbisKarteikarteGetActive      	(13) AlbisKarteikarteLesen             	(14) AlbisKarteikarteDokumentsuche
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisAkteSchliessen(CaseTitle="") {                                                             	;-- schließt eine Karteikarte ## seit Albis20.20 mit Problemen

	;Rückgabewerte 0 - die Funktion konnte keine Karteikarte schliessen, 1 - war nicht erfolgreich, 2 - es war keine Akte zum Schliessen geöffnet

	oACtrl			:= Object()			;active control object
	AlbisWinID	:= AlbisWinID()

	If InStr(AlbisGetActiveWindowType(), "Karteikarte")	{

		; aktuellen FensterTitel auslesen wenn keine Daten der Funktion übergeben wurden
			CaseTitle:= !CaseTitle ? AlbisGetActiveWinTitle() : CaseTitle

		; ein Eingabefocus in der Patientenakte muss erkannt und beendet werden (senden einer Escape- oder Tab-Tasteneingabe)
			while InStr(AlbisGetActiveWinTitle(), CaseTitle)				{
				aCtrl:= AlbisGetActiveControl("content")
				If ((aCtrl.Identifier = "Dokument") && (StrLen(aCtrl.RichEdit) = 0) )	     	{					;leeres RichEdit - dann steht das Caret im Edit-Control (das ginge doch auch anders?!)
						AlbisKarteikarteAktivieren()
						SendInput, {Esc}
						sleep 100
				}
				else If ((aCtrl.Identifier = "Dokument") && !(StrLen(aCtrl.RichEdit) = 0) ) 	{
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

		   ; diese Karteikarte gibt es nicht
			If !AlbisMDIChildHandle(CaseTitle)
				return 0

		   ; MDI-Tab nach vorne holen
			AlbisMDIChildActivate(CaseTitle)

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

		;~ SciTEOutput(AlbisGetActiveWinTitle() " =>> "  InStr(AlbisGetActiveWinTitle(), CaseTitle) )
		If InStr(AlbisGetActiveWinTitle(), CaseTitle) {
			SciTEOutput(A_ThisFunc ": Karteikarte schliessen per Menu" )
			Albismenu(57602)
		}

		while (InStr(AlbisGetActiveWinTitle(), CaseTitle) || A_Index < 100)
			sleep 50


return InStr(AlbisGetActiveWinTitle(), CaseTitle) ? 0 : 1
}
AlbisAkteSchliessen3(CaseTitle="", showerror:=false) {                                  	;-- 3.Version da 1. und 2. Version nicht funktionieren

	AlbisActivate(1)
	If IsObject(res := AlbisMDIChildActivate(CaseTitle)) {
		If showerror
			MsgBox, 0x1000, Addendum für Albis on Windows
						, % "Karteikarte von <" CaseTitle ">ist nicht geöffnet?`n"
						   .  "ErrorLevel:    " res.ErrorLevel "`n"
						   .  "LastError:     " res.LastError "`n"
						   .  "Mdi handle: " res.hMdi "`n"
		return 0
	}
	hKarteikarte := AlbisKarteikarteAktivieren()
	If AlbisKarteikarteEingabeStatus(CaseTitle) {
		Sleep 50
		ControlSend,, {Escape}, % "ahk_id " hKarteikarte
		Sleep 200
		If AlbisKarteikarteEingabeStatus(CaseTitle) {
			AlbisKarteikarteUnfocus(hKarteikarte)
			Sleep 200
		}
	}
	If InStr(WinGetTitle(AlbisWinID()), CaseTitle)
		Albismenu(57602)
	else {
		If showerror
			PraxTT("Karteikarte ist nicht geöffnet?`n" CaseTitle, "3 0")
		return 0
	}

}
;02
AlbisAkteOeffnen(CaseTitle="", PatID="") {                                                   	;-- öffnet eine Patientenakte über Name, ID oder Geburtsdatum

	/*		BESCHREIBUNG

			AlbisAkteOeffnen - zum Öffnen einer Patientenakte nach Übergabe eines Namens, der Patienten ID oder eines Geburtsdatums

			Die Funktion ist in der Lage sämtliche Fenster abzufangen und entsprechend zu behandeln. Es wird bei erfolgreichem Aufruf der
			gewünschten Akte eine 1 ansonsten eine 0 zurückgegeben. Falls sich ein Fenster mit einem Listview öffnet - Auswahl eines Namens,
			wenn mehr als einer vorhanden ist , wird eine 2 zurück gegeben. Nur das Listview Fenster wird nicht geschlossen. Da dieses
			ausgelesen werden muss, um in der Auswahl nach dem entsprechenden Namen suchen zu können.

			letzte Änderung 03.07.2022

	*/

	; Variablen                                                                                        	;{
		global Mdi, hMdi
		static WarteFunc, AlbisTitleFirst
		static Win_PatientOeffnen := "Patient öffnen ahk_class #32770"

		If !AlbisWinID() {
			PraxTT(A_ThisFunc ": Albis wird nicht ausgeführt!", "2 1")
			return 0
		}

		name    	:= []
		CaseTitle	:= Trim(CaseTitle)
		sStr := PatID := Trim(PatID)
	;}

	; RegExStrings, Anzeigenamen, Geburtsdatum und anderen Suchstring erstellen                         	;{
		If RegExMatch(PatID, "^\d+$") {

		  ; zusammengesetzter Name aus Nachname, Vorname
			PName := IsObject(cPat) ? cPat.NAME(PatID)         	: IsObject(PatDB) ? PatDB[PatID].Name ", " PatDB[PatID].Vorname 	: ""
		  ; Geburtsdatum
			GD   	:= IsObject(cPat) ? cPat.GEBURT(PatID, true)	: IsObject(PatDB) ? ConvertDBASEDate(PatDB[PatID].Geburt)        	: ""
		  ; Geschlecht m/w/o oder 1,2, ...
			GS     	:= IsObject(cPat) ? cPat.Geschl(PatID)         	: IsObject(PatDB) ? PatDB[PatID].Geschl                                       	: ""
		  ; Suchstrings
			rxStr      := "\[" PatID "\s\/"

		}
		else if RegExMatch(CaseTitle, "(?<surname>[\pL\-\s]+)*[\s,]*(?<prename>[\pL\-\s]+)*[\s,]*(?<Birth>\d{1,2}\.\d{1,2}\.(\d{4}|\d{2}))*", case) {

			PName	:= caseSurname ", " casePrename
			GD      	:= FormatDateEx(caseBirth, "DMY", "dd.MM.yyyy")
			rxStr     	:= (PName ? RegExReplace(PName, "\,\s+", ",\s+") "\s\/\s\w\s\/\s": "") (GD ? GD : "")
			sStr      	:= PName (PName && GD ? ", " : "") GD

		}
		else if (!CaseTitle && !PatID) {
			PraxTT(A_ThisFunc ": es wurden keine Parameter übergeben`nDer Karteikartenaufruf wird abgebrochen.", "2 1")
			return 0
		}
	;}

	; Karteikarte bereits geöffnet?                                                             	;{
		If (hMDIChild := AlbisMDIChildActivate(PatID " / " PName)) {
			PraxTT("Die Karteikarte " (GS~="(m|1)" ? "des Patienten " : "der Patientin ") PName ", geb. am " GD " wird angezeigt.", "5 1")
			return hMDIChild
		}
	;}

	; Öffne Patient Dialog aufrufen                                                           	;{
		PraxTT("Geöffnet wird die Karteikarte des Patienten:`n#2[" PatID "] " PName ", geb.am " GD "`n`n(Warte bis zu 10s auf die Karteikarte)", "10 0")
		AlbisTitleFirst := WinGetTitle(AlbisWinID())               	; aktuellen Albisfenstertitel auslesen
		hPOeffnen 	:= AlbisDialogOeffnePatient()               	; Aufruf des Fenster 'Patient öffnen'
	;}

	; Übergeben des Parameter an das Albisdialogfenster                         	;{
		If !VerifiedSetText("Edit1", sStr, Win_PatientOeffnen, 200)
			If (hPOeffnen && !VerifiedSetText("Edit1", sStr, hPOeffnen, 200)) {
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
				PraxTT(	"Die Karteikarte des Patienten:`n#2" PName ", geb.am " GD " (" PatID ")`nkonnte nicht geöffnet werden", "2 1")
				WinClose, % Win_PatientOeffnen
				return 0
			}

		}

	;}

	; Loop der in den nächsten 10 Sekunden auf die neue Karteikarte wartet	;{
		Loop	{

			If !(hwnd:= DLLCall("GetLastActivePopup", "uint", AlbisWinID()))
				hwnd := AlbisWinID()
			newTitle	:= WinGetTitle(hwnd)
			newClass	:= WinGetClass(hwnd)
			newText	:= WinGetText(hwnd)

		; Karteikarte ist geöffnet
			If RegExMatch(newTitle, "i)" rxStr) {
				PraxTT("", "Off")
				return AlbisMDIChildGetActive()
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
				return AlbisMDIChildGetActive()
			}

			If (A_Index > 50) 	{
				PraxTT(	"Die Karteikarte des Patienten:`n#2[" PatID "] " PName ", geb. am " GD "`nkonnte nicht geöffnet werden", "6 3")
				if WinExist(Win_PatientOeffnen)
					VerifiedClick("Button3", Win_PatientOeffnen)
				PraxTT("", "Off")
				return 0
			}

			sleep, 200
		}

		PraxTT("", "Off")
	;}

return 0
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
AlbisAkteGeoeffnet(Nachname, Vorname, GebDatum:="", PatID:="") {       	;-- ist die Akte des gesuchten Patienten aktuell geöffnet

	; vergleicht den Titel des Albisfenster mit den gesuchten Namen, Geburtstag und der PatID
	; letzte Änderung 27.09.2021

	; Name, Vorname, Geburtstag können leer sein, doch nicht alle zusammen, das Suchergebnis wäre immer ein erfolgreich, deshalb
		If (!Nachname && !Vorname && !GebDatum && !PatID) {
			Exception := {"file":A_ScriptFullPath, "what": "Empty parameter call! a minimum of 1 parameter is needed!", "line":A_LineNumber-7, "extra":A_ThisFunc}
			FehlerProtokoll(Exception, false)
			return 2
		}

	; Erkennungsfunktion
		AlbisTitle:= AlbisGetActiveWinTitle()

return (InStr(AlbisTitle, Nachname) && InStr(AlbisTitle, Vorname) && Instr(AlbisTitle, GebDatum) && InStr(AlbisTitle, PatID)) ? true : false
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
AlbisKarteikarteZeigen(Info=false) {                                                              	;-- schaltet zur Karteikartenansicht

	; !!WICHTIG!!: ein Erfolg der wird per Rückgabe einer 1 bestätigt, die aufrufende Funktion sollte den Rückgabewert prüfen bevor andere Funktionen fortgesetzt werden!
	; letzte Änderung: 12.01.2022

	; eine Karteikarte wird schon angezeigt
		If RegExMatch(AlbisGetActiveWindowType(), "i)Karteifenster")
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
		while !RegExMatch(AlbisGetActiveWindowType(), "i)Karteifenster")	{

			If (A_Index > 1)
				sleep 50
			PostMessage, 0x111, 33033,,, % "ahk_id " AlbisWinID()
			sleep 75

			Loop {


				If (hAlbisHinweis2 := WinExist("ALBIS ahk_class #32770", "Gebühren"))
					VerifiedClick("Button1", hAlbisHinweis2)


				while, (hAlbisHinweis1 := WinExist("ALBIS ahk_class #32770", "Die aktuelle Zeile wurde nicht gespeichert")) {

					WinSetTitle, % "ahk_id " hAlbisHinweis1,, % "ALBIS [ Zeileneingabe wird automatisch geschlossen in " Round(5-(0.1*(A_Index-1)),1) "s]"
					Sleep 50
					If (A_Index>30)	{
						VerifiedClick("Button1", "ALBIS ahk_class #32770", "Die aktuelle Zeile wurde nicht gespeichert")
						cFocus := Controls("", "GetFocus", (hActiveMDIWin:= AlbisMDIChildGetActive()))
						If (StrLen(cFocus) = 0)
							ControlFocus,, % "ahk_id " Controls("#327702", "ID", hActiveMDIWin)
						SendInput, {Esc}
						break
					}


				}

				If (A_Index > 10)
					return 0

				Controls("", "Reset", "")
				cFocus := Controls("", "GetFocus", (hActiveMDIWin:= AlbisMDIChildGetActive()))
				If (StrLen(cFocus) > 0)
					SendInput, {Esc}

			}

			If InStr(AlbisGetActiveWindowType(), "Karteifenster")
				return 1
			else if (A_Index > 10)
				return 0

		}

return 1
}
;08
AlbisDialogOeffnePatient(command:="invoke", pattern:="" ) {                       	;-- startet Dialog zum Öffnen einer Patientenakte

	; more commands are here:
	; 	abort/close - um das Fenster zu schliessen ohne eine Suche zu durchzuführen
	;	serach/set/open [Namens-/Suchmuster]- übernimmt den eingegeben Text und kann gleichzeitig das Suchmuster eintragen
	; letzte Änderung: 30.06.2022

		static Win_PatientOeffnen := "Patient öffnen ahk_class #32770"

	; Aufrufen des Dialogfensters
		If !InStr(command, "close") {
			hwnd 	:= Albismenu(32768, "Patient öffnen", 3)
			wT 		:= WinGetTitle(hwnd)
			If !InStr(wT, "Patient öffnen") {
				WinWait, % Win_PatientOeffnen,, 1
				If !WinExist(Win_PatientOeffnen) {
					PraxTT("Patient öffnen - Dialog konnte nicht aufgerufen werden.", "3 1")
					return 0
				}
				hwnd := WinExist(Win_PatientOeffnen)
			}
		}

	; command parsen
		If InStr(command, "invoke")
			return hwnd
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
					return hwnd

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


return hwnd
}
;09
AlbisKarteikarteAktiv() {                                                                               	;-- gibt true zurück wenn eine Karteikarte angezeigt wird
return InStr(AlbisGetActiveWindowType(), "Patientenakte") ? true : false
}
;10
AlbisKarteikartenAnsicht(PatientFensterCB) {                                                	;-- Ansicht Karteikarte, Abrechnung, Laborblatt

	; verändert aktuelle Auswahl der Combobox in der PatientFensterToolbar (Karteikarte,Abrechnung,Laborblatt...)

	ControlGet, hTbar	, Hwnd	,, ToolbarWindow321, % "ahk_id " AlbisMDIChildGetActive()
	ControlGet, hCB  	, Hwnd	,, ComboBox1			, % "ahk_id " hTbar
	If !hCB
		return 0

	; String oder Nummer wird unterschieden
	;SciTEOutput(hTbar)
return VerifiedChoose("ComboBox1", hTbar, PatientFensterCB)
}
;11
AlbisLeseDatumUndBezeichnung(MouseX, MouseY) {                                   	;-- Mausposition abhängige Karteikarten Informationen

	while (StrLen(KKarte.Datum.Text) = 0 && A_Index < 2) {
		MouseClick, Left, MouseX, MouseY                                                                     	; ein Mausklick setzt den Cursor
		Sleep 200
		KKarte := AlbisKarteikarteLesen()
	}
	SendInput, {Escape}                                                                                               	; Karteikartenzeile freigeben

return {"Text": KKarte.Inhalt.Text, "Datum":KKarte.Datum.Text, "Content":KKarte}
}
;12
AlbisKarteikarteGetActive() {                                                                      	;-- Handle der aktiven Karteikarte

	hactiveMDI   	:= AlbisMDIChildGetActive()
	hAfxFOrView	:= Controls("AfxFrameOrView140", "hwnd", hactiveMDI)
	hKarteikarte	:= Controls("#32770", "hwnd", hAfxFOrView)

return hKarteikarte
}
;13
AlbisKarteikarteLesen(hKarteikarte:=0) {                                                       	;-- Daten der aktiven Karteikartenzeile erhalten

	; letzte Änderung 20.02.2022

	If !hKarteikarte
		If !(hKarteikarte := AlbisKarteikarteGetActive())
			return

	presorted   	:= [], sorted := []
	baseclass   	:= 0
	hactiveMDI   	:= AlbisMDIChildGetActive()
	controls     	:= GetControls(hactiveMDI)

	For each, control in controls {

		If RegExMatch(control.classNN, "i)^Edit(?<N>\d+)", N) {
			baseclass := baseclass < NN ? NN : baseclass
			sorted.Push({"classNN": control.classnn, "text":control.Text, "hwnd":control.hwnd})
		}
		else if If RegExMatch(control.classNN, "i)^Richedit(?<N>\d+)")
			richedit := {"classNN": control.classnn, "text":control.Text, "hwnd":control.hwnd}

		If (sorted.Count() > 3)
			sorted.RemoveAt(1)

	}
	baseclass -= 2

	If (sorted.Count() = 3) {
		sortedtext := ""
		For each, edit in sorted
			sortedtext .= edit.text "#"
		If RegExMatch(sortedtext, "[A-Z]{1,2}#\d\d\.\d\d\.\d\d\d\d#(\w{1,5})*#")  ; zur Sicherheit wird der Inhalt geprüft
			return {"Arzt":sorted.1, "Datum":sorted.2, "Kuerzel":sorted.3, "Inhalt":richedit, "hKarteikarte":GetHex(hKarteikarte)}
	}

return
}
_AlbisKarteikarteLesen(hKarteikarte:=0) {                                                      	;-- _Daten einer Karteikarte erhalten

	static fcall := 0

	If !hKarteikarte
		If !(hKarteikarte := AlbisKarteikarteGetActive())
			return

	presorted   	:= [], sorted := []
	baseclass   	:= 0
	hactiveMDI   	:= AlbisMDIChildGetActive()
	controls     	:= GetControls(hactiveMDI)

	For each, control in controls
		If RegExMatch(control.classNN, "i)^Edit(?<N>\d+)", N)
			baseclass := baseclass < NN ? NN : baseclass
	baseclass -= 2

	If IsObject(controls)
		SciTEOutput("sorted[" baseclass "]: " cJSON.Dump(controls, 1))

	controls := GetControls(hKarteikarte)
	For each, control in controls
		If RegExMatch(control.classNN, "i)^Edit(?<N>\d+)", N) {
			If (NN >= baseclass)
				presorted[NN] := control.Text
		}
		else if RegExMatch(control.classNN, "i)^RichEdit")
			richedit := {"text":control.Text, "classNN":control.classNN}

	For NN, ctext in presorted
		sorted.Push({"text": ctext, "classNN": "Edit" NN})

	If IsObject(controls)
		SciTEOutput("sorted[" baseclass "]: " cJSON.Dump(controls, 1))

	If (sorted.Count() = 3) {
		sortedtext := ""
		For each, edit in sorted
			sortedtext .= edit.text "#"
		If RegExMatch(sortedtext, "[A-Z]{1,2}#\d\d\.\d\d\.\d\d\d\d#(\w{1,5})*#")
			return {"Arzt":sorted.1, "Datum":sorted.2, "Kuerzel":sorted.3, "Inhalt":richedit, "hKarteikarte":GetHex(hKarteikarte)}
	}

return
}
;14
AlbisKarteikarteDokumentsuche(hSearch:=0) {

	global hexs, exsEdtSearch, exsBtnSearch, exsDocList, exsGroup
	static hSearchO, swH
	static GuiH := 300

	If (WinGetTitle(hSearch) <> "Suchen")
		If !(hSearch := WinExist("Suchen ahk_class #32770"))
			return

	hSearchO := hSearch
	sw := GetWindowSpot(hSearch)
	If (sw.H < 500) {
		swH := sw.H
		SetWindowPos(hSearch, sw.X, sw.Y, sw.W, sw.H+GuiH)
	}

	If !WinExist("Addendum Volltextsuche ahk_class AutoHotkeyGUI1") {

		Gui, exs: New, -Caption -DPIScale +AlwaysOnTop +HWNDhexs Parent%hSearch%  ;0x50000000 E0x0000000C

		Gui, exs: Margin, 0, 0
		Gui, exs: Font, s8 q5, MS Sans Serif

		opt := "0x50000007 E0x4 "
		Gui, exs: Add, GroupBox	, % "x5 y0 w" sw.W-25 " h" GuiH+10 " vexsGroup " opt                                 	, Addendum - Volltextsuche in PDF und Word Dateien -
		Gui, exs: Add, Edit         	, % "x10 y15  w330         	vexsEdtSearch  	gexsHandler"                             	,
		Gui, exs: Add, Button        	, % "x+5 yp+-2                	vexsBtnSearch  	gexsHandler"                              	, Suchen

		cp := GuiControlGet("exs", "Pos", "exsGroup")
		dp := GuiControlGet("exs", "Pos", "exsEdtSearch")
		w := cp.W-10, h := GuiH-dp.Y-dp.H-5
		Gui, exs: Add, Listview     	, % "x10 y+5 w" w " h" h " 	vexsDocList   	gexsHandler"                             	, Datum|Bezeichnung|Fund
		Gui, exs: Show, % "x0 y" 175 " NoActivate", Addendum Volltextsuche

		SetTimer, exsCheck, 200

	}

return
exsHandler:

return
exsCheck: ;{

	If !WinExist("Suchen ahk_class #32770") {
		SetTimer, exsCheck, off
		If WinExist("Addendum Volltextsuche ahk_class AutoHotkeyGUI1")
			Gui, exs: Destroy
	}

return ;}

}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; DIALOGE (ANDERE)                                                                                                                                                                    (12+4)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAusindisKorrigieren		(02) AlbisDateiAnzeigen                   	(03) AlbisDateiSpeichern              	(04) AlbisSaveAsPDF
; (05) AlbisMenu                        	(06) AlbisLoescheEmpfaenger           	(07) AlbisAdressfelderZusatz       	(08) AlbisRehaDialog
; (09) AlbisPrintSettings                	(10) AlbisPrintSettings                      	(11) AlbisGNRVorschlag             	(12) AlbisMuster13Position
; class PatientenDaten                	(1) ShowDialog 	(2) CloseDialog 	(3) Personalien  	(4) weitereInfos
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisAusIndisKorrigieren(IndisToRemove:="", IndisToHave:="", opt:="") {     	;-- Ausnahmeindikationen: bestehende EBM-Ziffern tauschen oder hinzufügen

	; letzte Änderung: 26.06.2022

	; aktueller Patient
		currPatName := AlbisCurrentPatient()
		If !(PatID := AlbisAktuellePatID())
			return

	; Variablen
		If !IndisToRemove || !IsObject(IndisToHave)
			return 0

	; Optionen parsen
		Chroniker	:= RegExMatch(opt, "i)ChronischKrank\s*\=\s*1")    	? true : false
		GB 	    	:= RegExMatch(opt, "i)Geriatsch\s*\=\s*1")             	? true : false
		KeinBMP	:= RegExMatch(opt, "i)KeinBMP\s*\=\s*(1|ja|true)")	? true : false

	; Ausnahmeindikationen auslesen
		If ((ausIndisAlt := ausIndis := PatientenDaten.weitereInfos("GetText Ausnahmeindikation"))<0)
			ausIndisAlt := ausIndis := ""

	; Pat. wünscht keinen CGM BMP
		If !VerifiedCheck("Pat. wünscht keinen CGM BMP", "ahk_class #32770", "Adresse des",, KeinBMP)
			VerifiedCheck("Button19", "ahk_class #32770", "Adresse des",, KeinBMP)

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

	; Ausnahmeindikationen einsetzen
		If (ausIndisAlt <> ausIndis)
			PatientenDaten.weitereInfos("SetText Ausnahmeindikation", ausIndis)

	; Dialoge schliessen
		If (PatientenDaten.weitereInfos("Close", "OK") = 1)
			res := PatientenDaten.CloseDialog("Personalien")

	; Ausgabe
		If (ausIndisAlt <> ausIndis) {
			If !FilePathCreate(Addendum.DBPath "\logs")
				PraxTT("Protokolldateipfad: `n" Addendum.DBPath "\logs`nkonnte nicht angelegt werden.", "1 1" )
			else
				FileAppend, % "[" currPatName "] (" ausIndisAlt ") (" ausIndis ")`n", % Addendum.DBPath "\logs\Ausnahmeindikationen.txt"  , UTF-8
			return {"alt":ausIndisAlt, "neu":ausIndis}
		}

return  {"alt":ausIndisAlt, "neu":""}
}
;02
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
;03
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
;04
AlbisSaveAsPDF(filePath, printer="", Laborblattdruck=false) {                         	;-- Ausdruck/Speichern per PDF-Druckertreiber

	; ermöglicht den Export von Listen, Protokollen, dem Laborblatt über die Verwendung eines PDF-Druckertreibers
	; ACHTUNG: 	- die Funktion kann auch mit dem gesonderten Druckdialog des Laborblatt Drucks umgehen,
	;                    	  dazu muss der Dialog "Laborblatt Druck" allerdings zwingend vorher aufgerufen worden sein.
	;                    	- Albismenu() kann zwar auf 2 unterschiedliche Fenstertitel reagieren,
	;						  allerdings kann man keine Einstellungen im "Laborblatt Druck" Fenster vornehmen
	;                   	- das Dialogfenster macht erhebliche Schwierigkeiten wenn man per RDP arbeitet. Es läßt sich der
	;                         richtige Druckertreiber einstellen, allerdings wird auf den lokale Druckort umgeschaltet. PDF Druckausgaben
	;                     	  verschwinden oder öffnen sich auf dem RDP Client über welchen man gerade zugreift.
	;                      	  Anderen Weg der Manipulation der ComboBox entworfen am 21.11.2021
	;
	; letzte Änderung: 21.11.2021

		printer 	:= !printer 	? "Microsoft Print to PDF"            	: printer
		filepath	:= !filepath 	? A_Temp "\AlbisSaveAsPDF.pdf" 	: filepath

	; Albis-Laborblatt Druckdialog anzeigen falls noch nicht geschehen
		If Laborblattdruck {

			If !(hLaborblattdruck := WinExist("Laborblatt Druck ahk_class #32770"))
				If !(hLaborblattdruck := Albismenu(57607, "Laborblatt Druck ahk_class #32770", 3, 1))
					return "SAPDF2"

			while (!WinExist("Drucken ahk_class #32770") && A_Index <= 5) {
				VerifiedClick("Drucker...", "Laborblatt Druck ahk_class #32770")
				while (!WinExist("Drucken ahk_class #32770") && A_Index <= 20)
					sleep 20
			}

		}

	; handle des Drucker-Dialogs
		If !(hprintdialog := WinExist("Drucken ahk_class #32770"))
			return "SAPDF3"

	; PDF-Drucker wählen
	; mehrere Kontrollen (in einer Remotedesktopsitzung wird der Standort des virtuellen Druckers nicht erkannt)
	; die Überprüfung des Standortes wird nur bei MS Print to PDF vorbereitet!

	  ; Anzahl der Comboboxelemente
		SendMessage, 0x0146,,, % "ComboBox1", % "ahk_id " hprintdialog
		CBItems := ErrorLevel

	  ; erstes Element auswählen
		VerifiedSetFocus("ComboBox1", hprintdialog)
		VerifiedChoose("ComboBox1", hprintdialog, Printer)
		If (printer = "Microsoft Print to PDF") {
			Sleep, 50
			SendInput, {Right}
			Sleep, 100
			SendInput, {Left}
			Sleep, 100
			VerifiedChoose("ComboBox1", hprintdialog, Printer)
			sleep, 50
			ControlGetText, typ    	, Static5, % "ahk_id " hprintdialog
			ControlGetText, location, Static7, % "ahk_id " hprintdialog
			If  (typ <> printer || !InStr(location, "PORTPROMPT")) {
				PraxTT(	"Probleme bei der Auswahl des PDF-Druckertreiber:`n"
						. 	"<< " Printer " >>`n"
						. 	"Der Druckertreiber wurde nicht gefunden oder es gab Probleme bei Ermittlung des Standortes!", "2 1")
				return "SAPDF4bcText " location
			}
		}

	; OK und weiter
		If !(res := VerifiedClick("OK", hprintdialog))
			return "SAPDF7"

	; Laborblatt Druck aktiv dann ist noch ein OK notwendig
		If Laborblattdruck
			If !VerifiedClick("OK", hLaborblattdruck)
				return "SAPDF8"

	; Microsoft Print to PDF Dialog bearbeiten
		Druckausgabe := "Druckausgabe speichern unter ahk_class #32770"
		SpeichernBestaetigen := "Speichern unter bestätigen ahk_class #32770"
		WinWait, % Druckausgabe,, 5
		hDruckausgabe := WinExist(Druckausgabe)
		res := VerifiedSetText("Edit1", filePath, hDruckausgabe)	; Speicherpfad
		res := VerifiedClick("Button2", hDruckausgabe)              	; Speichern

	; Dialog Speichern unter bestätigen abfangen bei Bedarf
		while (WinExist(Druckausgabe) && A_Index < 30) {
			If (A_Index > 1)
				sleep 50
			If WinExist(SpeichernBestaetigen)                              	; Speichern unter bestätigen
				If !VerifiedClick("Ja", SpeichernBestaetigen)
					VerifiedClick("Button1", SpeichernBestaetigen)
		}
		If WinExist(SpeichernBestaetigen)
			return "SAPDF9"
		else if WinExist(Druckausgabe)
			return "SAPDF10"


	; Fenster: "Albis", Text: "Drucken" erscheint bei größeren Dokumenten
		WinWait        	, % "Albis ahk_class #32770", "Drucken", 5
		If WinExist("Albis ahk_class #32770", "Drucken")
			WinWaitClose	, % "Albis ahk_class #32770", "Drucken", 20

return ErrorLevel ?  "SAPDF11" : 1
}
;05
Albismenu(mcmd, FTitel:="", wait:=2, methode:=1) {                                	;-- Aufrufen eines Menupunktes oder Toolbarkommandos

	; letzte Änderung: 05.05.2022

	/* Albismenu() - Dokumentation

		mcmd:  		WM_Command - die Zahlen findet man im include Ordner AlbisMenu.json

		FTitel:			zu erwartender Fenstertitel, dient der Erfolgskontrolle des Menuaufrufes
	                		FTitel kann ein Objekt mit zwei Fenstertiteln/Fenstertexten sein
							Parameterübergaben: Albismenu(11111, ["Fenstertitel ahk_class WinClass", "Fenstertext", "Alternativer Fenstertitel ahk_class WinClass2", "Alternativer Text"])

		wait:	        	Zeit in Sekunden wie lange auf das Fenster gewartet werden soll

		methode:  	Aufruf per Post- (1) oder Sendmessage (2)
							Methode 1 und 2 wirkt sich nur bei leerem FTitel aus, da nur SendMessage die Variable ErrorLevel setzt
	 */

		InfoMsg := false

		If IsObject(FTitel) {
			WinTitle	:= FTitel.1
			WinText 	:= FTitel.2
			AltTitle  	:= FTitel.3
			AltText   	:= FTitel.4
		}
		else
			WinTitle	:= FTitel

	; prüft ob eines der erwarteten Fenster bereits geöffnet ist
		If (hwin := WinExist(WinTitle, WinText)) || (IsObject(FTitle) && (hAltwin := WinExist(AltTitle, AltText)))
			return hwin ? GetHex(hwin) : GetHex(hAltwin)

	; der Menu command Befehl öffnet manchmal nicht das gewollte Fenster (Bug von Albis?)
		AlbisIsBlocked(AlbisWinID())

	; Menuaufruf, wenn kein Fenstername übergeben wurde
		If !WinTitle {

			If RegExMatch(methode, "i)(1|Post)") {
				PostMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
				return 1
			}
			else { ; ACHTUNG: Sendmessage wartet auf eine Antwort, diese kann lange dauern (z.T. bis zu 5s)
				SendMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
				return ErrorLevel
			}

		}

	; Menuaufruf, bei Übergabe eines oder zweier Fensternamen
		If RegExMatch(methode, "i)(1|Post)")
			PostMessage	, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
		else
			SendMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"

	; Fenster abwarten
		maxRounds := wait*1000/20
		while (A_Index <= maxRounds) {
			sleep 20
			hWin 	:= WinExist(WinTitle	, WinText)
			hAltwin	:= WinExist(AltTitle	, AltText)
			If hWin || hAltwin
				return hwin ? GetHex(hwin) : GetHex(hAltwin)
		}

	hwin 	:= WinExist(WinTitle	, WinText)
	hAltwin	:= WinExist(AltTitle	, AltText)

return hWin ? GetHex(hwin) : hAltwin ? GetHex(hAltwin) : 0
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
;06
AlbisLoescheEmpfaenger(ask:=true) {                                                          	;-- Adresse des Rechnungsempfängers löschen

	; Daten von ... / weitere Informationen / Adresse des Rechnungsempfängers löschen
	; bei Adressen ohne Straßenangabe, kann das Fenster 'weitere Informationen' nicht ohne weiteres geschlossen werden

	If ask {
		MsgBox, 0x1024, Addendum für Albis on Windows, Rechnungsempfänger löschen?
		IfMsgBox, No
			return
	}

 ; leert alle Edit-Felder
	Loop 9
		VerifiedSetText("Edit" A_Index, "", "ahk_class #32770", 200, "Adresse des Rechnungsempfängers")

return VerifiedClick("OK", "ahk_class #32770", "Adresse des Rechnungsempfängers")
}
;07
AlbisAdressfelderZusatz() 	{                                                                      	;-- zeigt eine Schaltfläche zum Löschen der Adressfelder (weitere Informationen....)

	; letzte Änderung: 26.11.2021

	static hwInfo, wTittle, arr, hadr, hgroup, adrContent, adrDel, adrHDel, adrUndo := Object()

	If WinExist("Addendum Adressfelder leeren ahk_class AutoHotkeyGUI1")
		return

	hwInfo	:= WinExist("ahk_class #32770", "Adresse des Rechnungsempfängers")
	wTittle	:= WinGetTitle(hwInfo)
	ControlGet, hgroup     	, hwnd,, Button1	, % "ahk_id " hwInfo
	ControlGet, hgebDatum, hwnd,, Static8 	, % "ahk_id " hwInfo
	ControlGet, hZusatzver	, hwnd,, Button3	, % "ahk_id " hwInfo

	ws	:= GetWindowSpot(hwInfo)
	cs 	:= GetWindowSpot(hgebDatum)
	ds 	:= GetWindowSpot(hZusatzver)
	adrX	:= cs.CX 	- ws.CX - ws.BW - 8
	adrY	:= ds.CY 	- ws.CY - 2*ws.BH - 1

	arr := Object(), adrContent := false
	Loop 9
		If (cText := ControlGetText("Edit" A_Index, "ahk_id " hwInfo))
			arr[A_Index] 	:= cText

	Try {
		Gui, adr:  -Caption -DPIScale +AlwaysOnTop +HWNDhadr Parent%hgroup%  0x50000000 E0x0000000C
		Gui, adr: Margin	, 0, 0
		Gui, adr: Font 	, s8 q5 Normal, MS Sans Serif
		Gui, adr: Add  	, Button, % "x0 y0 w130 h" ds.H " vadrDel gadrLabel"	, % (arr.Count()                    	? "Felder leeren"
																													: adrUndo[wTitle].Count() 	? "Felder wiederherstellen" : "Felder leeren")
		Gui, adr: Show  	, % "x" adrX " y" adrY " NA"                                    	, % "Addendum Adressfelder leeren"
		SetTimer, adrCheck, 200
	}

	If arr.Count()
		adrUndo[wTitle] 	:= arr

return
adrLabel: ;{

	If arr.Count() {
		VerifiedSetText("Button1", "Felder wiederherstellen", hadr)
		Loop 9  ; leert alle Felder
			VerifiedSetText("Edit" A_Index, "", "ahk_id " hwInfo, 50)
		adrUndo[wTitle] 	:= arr
		arr := ""
	}
	else if adrUndo[wTitle].Count() {
		VerifiedSetText("Button1", "Felder leeren", hadr)
		Loop 9  ; schreibt den letzten Inhalt wieder zurück
			VerifiedSetText("Edit" A_Index, adrUndo[WTitle][A_Index], "ahk_id " hwInfo, 50)
		arr := adrUndo[wTitle]
		adrUndo.Delete(wTitle)
	}

return ;}
adrCheck: ;{

	If !WinExist("ahk_class #32770", "Adresse des Rechnungsempfängers") {
		SetTimer, adrCheck, off
		If WinExist("Addendum Adressfelder leeren ahk_class AutoHotkeyGUI1")
			Gui, adr: Destroy
	}

return ;}

}
;08
AlbisRehaDialog(hReha) {                                                                             	;-- Dialog aufziehen, wenn der Monitor genügend Auflösung hat

	MonNr	:= GetMonitorIndexFromWindow(hReha)
	wMon	:= ScreenDims(MonNr)
	wReha	:= GetWindowSpot(hReha)

  ;  Vergrößern wenn genug Höhe vorhanden ist
	if (wMon.H >= 1740)
		SetWindowPos(hReha, wReha.X, Floor(wMon.H//2 - 1740//2), wReha.W, 1740)
  ; Dialog an Auflösungshöhe des Monitor anpassen, wenn nicht genug Höhe vorhanden ist
	else
		SetWindowPos(hReha, wReha.X, 10, wReha.W, wMon.H-50)

}
;09
class PatientenDaten {                                                                               	;-- Menu Patient/Stammdaten - Dialoge aufrufen und bearbeiten

	static stammdaten := {"Abrechnungsassistent": {"cmd":34811, "WinTitle":"Abrechnungsassistent"}
								, 	  "Personalien"           	: {"cmd":32774, "WinTitle":"Daten von"}
								,	  "Dauerdiagnosen"    	: {"cmd":32776, "WinTitle":"Dauerdiagnosen von"}
								,	  "Dauermedikamente" 	: {"cmd":32777, "WinTitle":"Dauermedikamente von"}
								,	  "Cave"                     	: {"cmd":32778, "WinTitle":"Cave! von"}
								,	  "Kontrolltermine"       	: {"cmd":32775, "WinTitle":"Kontrolltermine"}
								,	  "Krankengeschichte"   	: {"cmd":32830, "WinTitle":"Krankengeschichte"}
								,	  "Patienteneinwilligung"	: {"cmd":35247, "WinTitle":"Patienteneinwilligung"}
								,	  "Patientenbild"          	: {"cmd":33317, "WinTitle":"Patientenbild"}
								,	  "Patientengruppen"     	: {"cmd":34362, "WinTitle":"Patientengruppen"}
								,	  "Familie"                   	: {"cmd":34705, "WinTitle":"Familie"}
								,	  "Therapiesitzungen"   	: {"cmd":34199, "WinTitle":"Therapiesitzungen"}
								,	  "Antikoagulantien"    	: {"cmd":34869, "WinTitle":"Antikoagulantion-Pass von"}}

	ShowDialog(dialog)                         	{	;-- eines der Dialogfenster aus dem stammdaten-Objekt aufrufen
		If !AlbisIsBlocked(2)
			return Albismenu(this.stammdaten[dialog].cmd, this.stammdaten[dialog].WinTitle " ahk_class #32770")
	return 0
	}

	CloseDialog(dialog, Options:="OK") 	{ 	;-- Dialogfenster schliessen
		If !(hwnd := WinExist(this.stammdaten[dialog].WinTitle " ahk_class #32770"))
			return 1
		WinActivate, % "ahk_id " hwnd
		If !VerifiedClick(Options, hwnd)
			If !VerifiedClick(Options, hwnd)
				return -1
	return 1
	}

	Personalien(cmd)                            	{	;-- Personaliendialog und weitere Info's

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
					this.hWInfos := WinExist("ahk_class #32770", "Adresse des")

				return this.hWInfos

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
					ControlGetText	, Ausnahmeindikation,, % "ahk_id " hAusnahmeindikation
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
;10
AlbisPrintSettings(newprinter:="Microsoft Print to PDF") {                              	;-- stellt den Drucker über das Menu Patient/Druckereinrichtung um

	hPrintSettings := AlbisMenu(57606, "Druckeinrichtung ahk_class #32770")
	ControlGet, hPrinterCB, hwnd,, ComboBox1, % "ahk_id " hPrintSettings
	ControlGetText, lastprinter,, % "ahk_id " hPrinterCB
	ControlGet, printerlist, List,,, % "ahk_id " hPrinterCB
	ControlClick, ComboBox1, % "ahk_id " hPrintSettings,, Left, 1
	Sleep 50
	For entryNr, printer in StrSplit(printerlist, "`n")
		If (printer == newprinter) {

			printerfound := true

			ControlSend, ComboBox1, {Home},  % "ahk_id " hPrintSettings
			Sleep 50
			ControlGetText, activeprinter,, % "ahk_id " hPrinterCB
			If (activeprinter = newprinter)
				break

			Loop % entryNr-1 {
				ControlSend, ComboBox1, {Down},  % "ahk_id " hPrintSettings
				Sleep 50
				ControlGetText, activeprinter,, % "ahk_id " hPrinterCB
				If (activeprinter = newprinter)
					break
			}

			break
		}
	Sleep 50
	ControlClick, ComboBox1, % "ahk_id " hPrintSettings,, Left, 1
	VerifiedClick,("OK", hPrinterSettings,,, true)

return printerfound ? lastprinter : ""
}
;11
AlbisGNRVorschlag(quartal:="auto", close:=true) {                                     	;-- Auswahl des aktuell bearbeiteten Quartals ##

	; im Moment wird nur bei einer Ziffer automatisch ausgewählt

	hGNR := WinExist("GNR-Vorschlag ahk_class #32770")
	If IsObject(EditControls	:= Controls("", "GetControls", hKarteikarte := AlbisKarteikarteGetActive())) {
		For each, hEdit in EditControls
			If RegExMatch(ControlGetText("", "ahk_id " hEdit), "^[\d\.]+", KKDate)
				break
		quartal	:= GetQuartalEx(KKdate ? ConvertToDBASEDate(KKDate) : A_YYYY A_MM A_DD, "Q/YY")
		LBItems := AlbisLVContent(hGNR, "Listbox1")
		If (LBItems = 1)
			For row, cols in LBItems
				For col, item in cols
					If (item ~= "\(" quartal "\)") {
						SendMessage, 0x0186, % row-1,, ListBox1, % "ahk_id " hGNR
						EL := ErrorLevel
						If (EL=0 && close)
							VerifiedClick("OK",,, hGNR, true)
						break
					}
				}
}
;12
AlbisMuster13Position() {                                                                           	;-- Heilmittelformular wird am linken Bildschirmrand positioniert
	albisP 	:= GetWindowSpot(AlbisWinID())
	WinP 	:= GetWindowSpot(hHookedWin)
	WinMoveZ(hHookedWin, 0, albisP.X<0?0:albisP.X, Floor(albisP.H/2-WinP.H/2), winP.W, winP.H)
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; LABOR                                                                                                                                                                                        	(17)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) vianovaInfoBoxWebClient      	(02) AlbisLaborWaehlen             	(03) AlbisLaborbuchOeffnen       	(04) AlbisLaborGNRHandler
; (05) AlbisGNRAnforderungChilds  	(06) AlbisLaborblattDrucken          	(07) AlbisLaborblattPDFDruck		(08) AlbisLaborblattExport
; (09) AlbisLaborblattZeigen            	(10) AlbisLaborAuswählen           	(11) AlbisLaborDaten                    	(12) AlbisLaborAlleUebertragen
; (13) AlbisLaborImport                  	(14) AlbisLaborKeineDateien         	(15) AlbisLabBuchRowSelected		(16) AlbisLabBuchUnknownID
; (17) AlbisLabBuchZuordnen
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
vianovaInfoBoxWebClient(infoBoxID=0) {                                                               	;-- WebClient Automation (Labor IMD)

	; letzte Änderung vom 17.06.2021
	; 17.06.2021: SetTimer Aufruf fehlerhaft
	; 30.01.2021: flexiblere Steuerelementerkennung
	; Achtung: volle Funktionsfähigkeit nur wenn Addendum läuft (Fensterhooks)

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
	;	mit Notepad++ lässt sich eine LDT am besten anzeigen, Encoding mit: North European/OEM 865
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
	; letzte Änderung 13.09.2022

	; blockierende Fenster schliessen
		AlbisIsBlocked(AlbisWinID())
		If hinweis
			PraxTT("Befehl:`n- Laborbuch anzeigen -`nwird ausgeführt", "6 1")

	; ist das Laborbuch schon geöffnet wird es aktiviert
		If (hLabbuch := AlbisMDIChildHandle("Laborbuch")) {
			AlbisMDIChildActivate("Laborbuch")
			return hLabbuch
		}

	; Öffnen des Laborbuches per WM_Command Befehl - Laborbuch:= 34162
		PraxTT("Warte auf das Laborbuch", "0 0")
		WarteLBuch := bitteWarten := 0
		PostMessage, 0x111, 34162,,, % "ahk_class OptoAppClass"   ; Menu Laborbuch
		SciTEOutput("Menu Laborbuch ErrorLevel: " ErrorLevel)

	; wartet etwas länger auf das Laborbuch wenn Albis auf das Zusammenstellen von Daten hinweist
		TStart := A_TickCount, stopAt := 180
		while !InStr(AlbisGetActiveWinTitle(), "Laborbuch") {
			If (hBW := WinExist("Bitte warten... ahk_class #32770")) || (hAW := WinExist("Albis ahk_class #32770", "Archiv"))
				stopAt := 90, bitteWarten += 1
		   ; es wird länger gewartet wenn Dialogfenster zusätzliche Datenverarbeitung durch Albis angezeigt hatten
			wTime := Round((A_TickCount-TStart)/1000, 1)
			GuiControl, PraxTT:, PraxTTCnt1, % stopAt-wTime "s"
			GuiControl, PraxTT:, PraxTTCnt2, % Round((bitteWarten*50)/1000, 1) "s"
			If (wTime >= 180) || (bitteWarten && wTime >= 90)
				break
			sleep 50
		}
		SciTEOutput("Bitte warten/Achivieren.... wurde für " wTime " Sekunden angezeigt", "8 1")

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
AlbisLaborblattDrucken(Spalten, Drucker="", SaveDir="", PrintAnnotation=1) {        	;-- drucken eines Laborbefundes


	; im Moment kann nur die Anzahl der auszudruckenden Spalten als Option übergeben werden
	; Drucker := "pdf" für PDF Drucker - einzustellen in den Addendum.ini
	; Drucker := "Standard" für den eingestellten Standard Drucker - zu hinterlegen in der Addendum.ini
	; (## AlbisLaborblattExport() ist besser!)

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
AlbisLaborblattPDFDruck(Columns, PrintAnnotation:=true)      	{

	static AlbisView

	AlbisView 	:= AlbisGetActiveWindowType(true)
	savePath 	:= Addendum.ExportOrdner "\Laborwerte von " RegExReplace(AlbisCurrentPatient(), "[\s]") ".pdf"
	If (Columns = "Alles") {
		MsgBox, 0x1004, % "Sie sind gerade im Begriff alle Laborwerte zu versenden.`nSind Sie sich sicher?"
		IfMsgBox, No
			return "LBExNo"
	}

  ; Drucker wird über das Albismenu geändert
	If !(lastprinter := AlbisPrintSettings("Microsoft Print to PDF")) {
		PraxTT("Das Umstellen des Drucker auf Microsoft Print to PDF ist fehlgeschlagen.`nEin EMailversand von Laborwerten ist nicht möglich.", "3 0")
		return "LBEx6"
	}

  ; startet den Druckdialog
	If !(hLaborblattdruck := Albismenu(57607, "Laborblatt Druck ahk_class #32770", 3, 1)) {
		PraxTT(A_ThisFunc ": Der Laborblatt Druck konnte nicht ausgeführt werden!", "3 1")
		sleep 3000
		return "LBEx7"
	}

  ; auf Spaltendruck umstellen
	If !VerifiedClick("Button2", hLaborblattdruck)                 	 ; letzte
		return "LBEx8"

  ; Spaltenzahl eingeben
	If !VerifiedSetText("Edit1", Columns, hLaborblattdruck)
		return "LBEx9"

  ; Anmerkungen und Probedate
	If !VerifiedCheck("Button5", hLaborblattdruck,,, PrintAnnotation)
		return "LBEx10"

	If !VerifiedClick("Drucker", hLaborblattdruck)
		return "LBEx11a"

	WinWait, % "Drucken ahk_class #32770",, 5
	If !WinExist("Drucken ahk_class #32770")
		return "LBEx11b"

	Sleep 200
	If !VerifiedClick("OK", "Drucken ahk_class #32770",,, true)
		return "LBEx11c"

  ; OK drücken -> startet Ausgabevorgang
	If !VerifiedClick("OK", "Laborblatt Druck ahk_class #32770",,, true)
		return "LBEx12"

  ; Microsoft Print to PDF Dialog bearbeiten
	Druckausgabe := "Druckausgabe speichern unter ahk_class #32770"
	SpeichernBestaetigen := "Speichern unter bestätigen ahk_class #32770"
	WinWait, % Druckausgabe,, 5
	hDruckausgabe := WinExist(Druckausgabe)
	res := VerifiedSetText("Edit1", savePath, hDruckausgabe)	; Speicherpfad
	res := VerifiedClick("Button2", hDruckausgabe)              	; Speichern

  ; Dialog Speichern unter bestätigen abfangen bei Bedarf
	while (WinExist(Druckausgabe) && A_Index < 30) {
		If (A_Index > 1)
			sleep 50
		If WinExist(SpeichernBestaetigen)                              	; Speichern unter bestätigen
			If !VerifiedClick("Ja", SpeichernBestaetigen)
				VerifiedClick("Button1", SpeichernBestaetigen)
	}
	If WinExist(SpeichernBestaetigen)
		return "LBEx13"
	else if WinExist(Druckausgabe)
		return "LBEx14"

  ; vorherige Druckereinstellung wiederherstellen
	If !(lastprinter := AlbisPrintSettings(lastprinter)) {
		PraxTT("Das Wiederherstellen des ursprünglich eingestellten Druckers ist fehlgeschlagen.`nBitte kontrollieren Sie die Einstellung!", "9 0")
	}

  ;~ ; Albisansicht wiederherstellen
	;~ If (AlbisView <> AlbisGetActiveWindowType(true))
		;~ AlbisKarteikartenAnsicht(AlbisView)

	Sleep 2000

return savePath
}
;08
AlbisLaborblattExport(PrintRange, SaveAs="", Printer="", PrintAnnotation=1) {          	;-- PDF Export oder Druckausgabe des Laborblattes

	; ACHTUNG: 	Funktion setzt nur die Inhalte der Steuerelemente anhand der übergebenen Parameter
	;                    	der Export/Druck als PDF Datei erfolgt über AlbisSavePDF()
	;
	; letzte Änderung am 20.11.2021

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
				return "missing path: " SaveDir
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
			return "print range problem"
		}

	  ; Printer
		If !Printer {
			PraxTT(A_ThisFunc ": Es wurde kein Drucker übergeben!", "3 1")
			sleep 3000
			return "missing printer device name"
		}

	;}

	; Laborblatt anzeigen falls möglich
		If !InStr(AlbisGetActiveWindowType(), "Laborblatt")	{
			If !InStr(AlbisGetActiveWindowType(), "Patientenakte")		{
				PraxTT("Der Laborblattdruck kann nur bei einer geöffneten`nKarteikarte ausgeführt werden!.", "3 1")
				sleep 3000
				return "no case"
			}
			If !AlbisLaborblattZeigen(true) {
				PraxTT("Das Anzeigen des Laborblattes ist fehlgeschlagen.", "3 1")
				return "problem switching to labview"
			}
		}

	; startet den Druckdialog
		If !(hLaborblattdruck := Albismenu(57607, "Laborblatt Druck ahk_class #32770", 3, 1)) {
			PraxTT(A_ThisFunc ": Der Laborblatt Druck konnte nicht ausgeführt werden!", "3 1")
			sleep 3000
			return "print impossible"
		}

	; Einstellungen eintragen
	  ; Datumsbereich
		If RegExMatch(PrintRange, "(?<Von>\d\d\.\d\d\.\d\d\d\d)\-(?<Bis>\d\d\.\d\d\.\d\d\d\d)", rx) || (PrintRange = "Alles") {
			If !VerifiedClick("Button3", hLaborblattdruck)                    ; Zeitraum
				return "LBEx5"
			If !(PrintRange = "Alles") {
				If !VerifiedSetText("Edit2", rxVon	, hLaborblattdruck)
					return "problem setting printrange from"
				If !VerifiedSetText("Edit3", rxBis	, hLaborblattdruck)
					return "problem setting printrange to"
			}
		}
	  ; Anzahl von Spalten
		else {
			If !VerifiedClick("Button2", hLaborblattdruck)                 	 ; letzte
				return "LBEx8"
			If !VerifiedSetText("Edit1", (PrintRange = 0 ? 1 : PrintRange), hLaborblattdruck)
				return "LBEx9"
		}

	 ; Anmerkungen und Probedate
		VerifiedCheck("Button5", hLaborblattdruck,,, PrintAnnotation)

return AlbisSaveAsPDF(saveAs, Printer, true)   ; (Dateipfad, Names des Druckers, Laborblattdruck = ja)
}
;09
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
;10
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
				Addendum.Labor.AbrufStatus 	:= false
				Addendum.Labor.AbrufVoll    	:= false
				return err
		}

return err
}
;11
AlbisLaborDaten()	{                                                                                                 	;-- Labordaten importieren

  ; +öffnet im Anschluss das Laborbuch
  ; letzte Änderung: 28.09.2021
	static WinLabDat := "Labordaten ahk_class #32770"

	If !VerifiedCheck("Button5"	, WinLabDat,,, true)
		If !VerifiedClick("Button5", WinLabDat)
			return 0
	VerifiedClick("Button1"	, WinLabDat,,, true)
	WinWaitClose, % WinLabDat,, 20
	If !WinExist(WinLabDat) {
		AlbisLaborbuchOeffnen()
		return 2
	}

return 1
}
;12
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
				res := VerifiedClick("Button1",,, hDialog, true)
		}

return {"EL":EL, "click":res, "hwnd":hDialog}
}
;13
AlbisLaborImport(LbName) {			         					                                				;-- Labordaten importieren

	; 1. 	Aufruf des Albis Menu: Extern/Labor/Daten importieren...
	; 2. 	Bearbeitung des Laborauswahl Dialogs:  LbName ist das zu wählende Labor
	; 3. 	Bearbeitung des Labordaten Sammelimport Dialogs:
	;		a) Haken bei Sammelimport vormerken setzen
	;		b) OK Button drücken
	; 4. 	Popup-Dialoge (Fortschrittsanzeigen des Datenimports) abwarten bis das Laborbuch angezeigt wird
	;
	; !WICHTIG!: 	in Albis im Menu: 'Option/Labor/Import' muss der Haken bei 'Laborbuch nach Import automatisch öffnen' gesetzt sein
	;						, damit die Automatisierung des Laborimportes durch Addendum komplett durchgeführt werden kann
	;
	; letzte Änderung: 07.01.2022

		static WinLabordaten     	:= "Labordaten ahk_class #32770"
		static WinLaborWaehlen 	:= "Labor auswählen ahk_class #32770"

		global isLabScript := RegExMatch(A_ScriptName, "i)Laborabruf") ? true : false
		global amsg

	; ―――――――――――――――――――――――――――――――――――――――――――
	; [Menuaufruf] Aufruf des Dialoges Labor wählen über das Menu Extern\Labor\Daten importieren ....
	; ―――――――――――――――――――――――――――――――――――――――――――
		while (!hLaborwaehlen := AlbisLaborwaehlen() && A_Index < 200)         	; bis 10s wird gewartet
			sleep 50
		If (!hLaborwaehlen || !VerifiedChoose("ListBox1", hLaborwaehlen, LbName)) {
			AlbisLabImportMsg(A_ThisFunc "() `t- Labor wählen: Listboxeintrag " LbName " konnte nicht ausgewählt werden.")
			return 0
		}

	; ―――――――――――――――――――――――――――――――――――――――――――
	; OK drücken und 2 Sekunden auf das Schliessen des Fenster warten
	; der Hinweisdialog 'Keine Datei(en) im Pfad' wird abgefangen und geschlossen, das Skript bricht dann hier ab
	; ―――――――――――――――――――――――――――――――――――――――――――
		err := VerifiedClick("Button1", hLaborwaehlen,,, 2)
		If isLabScript
			amsg .= A_ThisFunc "() `t- Warte auf den nächsten Laborabruf-Dialog"
		WinWait, % "ALBIS", % "Keine Datei(en) im Pfad", 2
		If !ErrorLevel 	{
			AlbisLabImportMsg(A_ThisFunc "() `t- Labor wählen: keine Daten im Ordner " adm.Labor.LDTDirectory)
			If !(err := VerifiedClick("Button1", "ALBIS", "Keine Datei(en) im Pfad"))
				AlbisLabImportMsg(datestamp(2) "|" A_ThisFunc "() `t- Fenster 'Keine Datei(en) im Pfad' konnte nicht geschlossen werden.")
			return err
		}

	; ―――――――――――――――――――――――――――――――――――――――――――
	; [Labordaten Sammelimport]
	; ―――――――――――――――――――――――――――――――――――――――――――
		If isLabScript
			amsg .= A_ThisFunc "() `t- Warte auf Labordaten - Sammelimport"
		while !WinExist(WinLabordaten, "Sammelimport") && (A_Index < 800) {    ; = 40 Sekunden
			Sleep 50
		}
		If !(hSammelImport := WinExist(WinLabordaten, "Sammelimport")) {
			AlbisLabImportMsg(A_ThisFunc "() `t- Labordaten Sammelimport: hat sich nicht geöffnet`n")
			return 0
		}

	; ―――――――――――――――――――――――――――――――――――――――――――
	; kurze Pause damit sich das Fenster noch komplett aufbauen kann
	; ―――――――――――――――――――――――――――――――――――――――――――
		Sleep 4000

	; ―――――――――――――――――――――――――――――――――――――――――――
	; [Labordaten Sammelimport] Checkbox 'für Sammelimport vormerken' anhaken
	; ―――――――――――――――――――――――――――――――――――――――――――
		If !VerifiedCheck("Button5", hSammelImport,,, true)
			If !VerifiedClick("Button5", hSammelImport)
				If VerifiedSetFocus("Button5", hSammelImport) {
					BlockInput, On
					SendInput, {Space}
					BlockInput, Off
				}
				else {
					AlbisLabImportMsg(A_ThisFunc "() `t- Labordaten Sammelimport: Checkbox ''für Sammelimport vormerken'' konnte nicht gesetzt werden")
					return 0
				}

	; ―――――――――――――――――――――――――――――――――――――――――――
	; [Labordaten Sammelimport] Button - 'Sammelimport' drücken
	; ―――――――――――――――――――――――――――――――――――――――――――
		If !VerifiedClick("Button1", WinLabordaten,,, 3)
			If !VerifiedClick("Button1", hSammelImport,,, 3)
				If !VerifiedClick("Sammelimport", WinLabordaten,,, 3) {
					AlbisLabImportMsg(A_ThisFunc "() `t- Labordaten Sammelimport: konnte nicht geschlossen werden.")
					return 0
				}
		If isLabScript
			amsg .= A_ThisFunc "() `t- Warte auf die Anzeige des Laborbuch"

	; ―――――――――――――――――――――――――――――――――――――――――――
	; [Importvorgang] 	Fortschrittsanzeige detektieren und solange warten bis der Import abgeschlossen ist
	; 		            		wartet auf Popupfenster und wartet bis dieses geschlossen wurde
	;		            		oder bricht ab wenn das Laborbuch angezeigt wird
	; ―――――――――――――――――――――――――――――――――――――――――――
		writeStr2 	:= "" A_ThisFunc "() `t- Labordaten Sammelimport: Wartezeit für Laborbuch überschritten."
		writeStr3 	:= "" A_ThisFunc "() `t- Labordaten Sammelimport: letztes Fenster (Importfortschritt) konnte nicht abgefangen werden."
		PopupWin	:= false, waitdialog := albisonly := 0
		win1 := "Warten ahk_class #32770", win2 := "ALBIS ahk_class #32770", "Warten"
		Loop {

		  ; Laborbuch wird angezeigt
			If InStr(AlbisGetActiveWinTitle(), "Laborbuch")
				return 1

			PopTile 	:= WinGetTitle(hPopup := DllCall("GetLastActivePopup", "uint", AlbisWinID())) " "
			PopText  	:= WinGetText(hPopup) " "
			PopClass 	:= WinGetClass(hPopup)

		  ; Bitte warten... Dialog abfangen
			If (   RegExMatch(PopTitle . PopClass, "i)(Bitte\swarten)|(ALBIS)\s+#32770")
			  || RegExMatch(PopText . PopClass, "i)Bitte\swarten\s+#32770")
			  || WinExist(win1) || WinExist(win2))
				waitdialog	:= (!waitdialog ? true : waitdialog), albisonly	:= 0

		  ; Abbruch nach insgesamt 200 Loops ohne Wartedialog oder 100 mit Wartedialog
			else if (AlbisWinID() = hPopUp) {
				albisonly ++
				If (!waitdialog && albisonly >= 200) || (waitdialog && albisonly >= 100) {
					ts := GetTimestrings(albisonly*1000, 1)
					LeerLaufzeit 	:= " Leerlaufzeit:  "       	(ts.min>0 ? ts.min "min ":"") ts.sec "." Floor(ts.ms/100) "s"
					ts := GetTimestrings(A_Index*1000, 1)
					Wartezeit  	:= ", Gesamtwartezeit: " 	(ts.min>0 ? ts.min "min ":"") ts.sec "." Floor(ts.ms/100) "s"
					AlbisLabImportMsg(writeStr2 . LeerLaufzeit ", " Wartezeit)
					return 1
				}
			}

		  ; Abbruch nach circa 2 Minuten
			else if (A_Index > 1200) {
				AlbisLabImportMsg(writeStr3 " [1]")
				return 0
			}

			Sleep 100

		}

}
AlbisLabImportMsg(msg) {                                                                                      	;-- gehört zu AlbisLaborimport
	global isLabScript
	If isLabScript
		FileAppend, % (amsg .= datestamp(2) "|" msg) "`n", % adm.LogFilePath
	else
		PraxTT(msg, "5 1")
}
;14
AlbisLaborKeineDateien(hwnd=0) {                                                                          	;-- schließt Albisdialog: Keine Datei(en) im Pfad

	If !WinExist("ALBIS ahk_class #32770", "Keine Datei(en) im Pfad")
		return 1

	If !hwnd
		err := VerifiedClick("Button1", "ALBIS ahk_class #32770", "Keine Datei")
	else
		err := VerifiedClick("Button1", hwnd)

return err
}
;15
AlbisLabBuchRowSelected() {                                                                                    	;-- gibt die ausgewählte Zeile im Laborbuch zurück

	hLabBuch := AlbisMDIChildHandle("Laborbuch")
	SendMessage, 0x100C, -1, 0x2, SysListview321, % "ahk_id " hLabBuch   ; LVM_GETNEXTITEM
	rowSelected := ErrorLevel + 1          ; if is ErrorLevel -1, in this case + 1, is zero nothing is selected

return rowSelected
}
;16
AlbisLabBuchUnknownID(rowSelected:=0) {                                                            	;-- schlägt den wahrscheinlichsten Patienten vor

		global aDB

		static unknown := Array()
		static rxlabTable := "^\s*(?<Patient>[\pL\-]+,\s+[\pL\-]+)*\s*\(*(?<PatID>\d+)*\)*\s+(?<ANFNR>\d+)\s+(?<dd>\d\d)\.(?<mm>\d\d)\.(?<yyyy>\d\d\d\d)"

		cPat := new PatDBF("","","moredata=true")
		If !cPat.ItemsCount()
			throw A_ThisFunc " [codeline: " A_LineNumber-9 "]: braucht cPat als Klassen-Objekt!`n"

		hLabBuch := AlbisMDIChildHandle("Laborbuch")
		VarSetCapacity(LBuchText, 26000)
		ControlGet, LBuchText, List,, SysListview321, % "ahk_id " hLabBuch
		tableRows := StrSplit(LBuchText, "`n")
		If rowSelected {
			RegExMatch(tableRows[rowSelected], rxlabTable, lab)
			SelectedANFNR := labANFNR
			SelectedLabDay := labyyyy . labmm . labdd
		}

	; weitere namenlose Einträge erfassen
		newitems := 0
		anforderungen := labdays := "rx:("
		For row, rowText in StrSplit(LBuchText, "`n")
			If RegExMatch(rowText, rxlabTable, lab) {
				If !labPatient && !unknown.haskey(labANFNR) {
					newitems              	+= 1
					labDay                  	:= labyyyy . labmm . labdd
					anforderungen         	.= (labANFNR) "|"
					labDays                 	.= (labDay) "|"
					unknown[labANFNR] := {"labDay":labday, "row":row}
				}
			}
		anforderungen 	:= RTrim(anforderungen, "|") ")"
		labDays        	:= RTrim(labDays, "|") ")"

 	; Datenbank öffnen, auslesen und Lesezugriff beenden (nur wenn neue Einträge vorhanden sind)
		If (newitems) {

			VarSetCapacity(data, 10x1024)
			labDB 	:= new DBASE(Addendum.AlbisDBPath "\LABBUCH.dbf", 0)
			res     	:= labDB.OpenDBF()
			data 	:= labDB.Search({"EINDATUM":labDays, "ANFNR":anforderungen}, 0,, {"LogicalComparison":"and"})
			res     	:= labDB.CloseDBF()
			labDB 	:= ""

		; aussortieren nicht passender Einträge
			If (data.Count() = 0)
				return 0

			For idx, m in data
				If m.PATGEB {

				  ; findet mitunter mehrere Patient mit diesem Geburtsdatum
					If !IsObject(PatIDs := cPat.PatID({"GEBURT":m.PatGEB})) {
						SciTEOutput(m.PATGEB " nicht gefunden")
						continue
					}

					For each, PatID in PatIDs  {      ; Geburtstage sind noch im DBASE Format yyyymmdd
						matchpriority := 0
						matchpriority += (unknown[m.ANFNR].labDay <= cPat.LASTBEH(PatID)) ? 2 : 0
						matchpriority += InStr(cPat.Get(PatID, "PRIVAT"), "t") ? 1 : 0
						If !matchpriority
							continue
						unknown[m.ANFNR][PatID] := { "Name"       	: cPat.NAME(PatID, true)
																	, 	"LASTBEH"    	: cPat.LASTBEH(PatID)
																	,	"Privat"      	: (cPat.Get(PatID, "PRIVAT")="t" ? true : false)
																	, 	"PATGEB"    	: m.PATGEB
																	, 	"matchPerc"	: matchpriority}
					}
				}

			VarSetCapacity(dbfdata, 0)
		}

	; besten Treffer heraussuchen
		bestmatches := Array(), maxPerc := 0
		If (unknown[SelectedANFNR].labday = SelectedLabDay)  {
				For PatID, data in unknown[SelectedANFNR] {
					If data.matchPerc && (maxPerc <= data.matchPerc ) {
						maxPerc := data.matchPerc
						bestmatches.Push({"Name": StrReplace(data.NAME, "*", ""), "PATID":PatID, "LASTBEH": data.LASTBEH})
					}
				}
		}


return (bestmatches.Count()=1) ? bestmatches.1.NAME : bestmatches.Count()>1 ? bestmatches : ""
}
;17
AlbisLabBuchZuordnen() {                                                                                     	;-- ## bearbeitet das Fenster Anforderung zuordnen

	; CGM benutzt die LABREAD, LABBUCH.dbf nicht mehr - daher sind nur Daten bis zu einem bestimmten Tag verwendbar

	global aDB

	hAnforderung := WinExist("Anforderung zuordnen ahk_class #32770")
	ControlGet, hEdit, hwnd,, Edit1, % "ahk_id " hAnforderung
	anf := GetWindowSpot(hEdit)

	Gui, bl1: New, Hwndhbl1 -Caption -DPIScale +AlwaysOnTop
	Gui, bl1: Margin, 0, 0
	Gui, bl1: Font, s10 q5 italic bold, Courier New
	Gui, bl1: Add, Text, % "x0 y0 w" anf.W " h" anf.H*2 " Center", % "...bitte warten...`nPatientendaten werden ermittelt"
	Gui, bl1: Show, % "x" anf.X " y" anf.Y " NoActivate", Addendum Autozuordnung läuft

  ; Datenbankzugriffe öffen
	aDB := new AlbisDB(Addendum.AlbisDBPath)
	If (rowSelected := AlbisLabBuchRowSelected()) {
		If !IsObject(predictPat := AlbisLabBuchUnknownID(rowSelected))
			ControlSetText, Edit1, % RegExReplace(predictPat, ",*\s*\d+\.\d+\.\d+\s*"), % "Anforderung zuordnen ahk_class #32770"
		else {
			for index, Pat in predictPat
				p .= index ". [" PatID "] " Pat.Name ", letzte Behandlung am: " Pat.LASTBEH "`n"
			MsgBox, 0x1000, Addendum für Albis on Windows, % "Es kommen " predictPat.Count() " Patienten in Frage:`n" p
		}
	}
	aDB := ""

	Gui, bl1: Destroy

}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; GRAFISCHER BEFUND                                                                                                                                                                   	(05)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisOeffneGrafischerBefund	(02) AlbisUebertrageGrafischenBefund                                           	(03) AlbisImportierePdf
; (04) AlbisImportiereBild               	(05) AlbisBrief
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisOeffneGrafischerBefund() {                                                             	;-- Dialog 'Grafischer Befund' - importieren von Bilddateien (jpg z.B.)
return Albismenu(32960, "Grafischer Befund ahk_class #32770")
}
;02
AlbisUebertrageGrafischenBefund(ImgName, KKText:="") {                      	;-- Funktion mit Sicherheitsüberprüfung falls ein Focus verloren geht

		; letzte Änderung: 17.02.2021

		static WinText1	:= "Bitte überprüfen Sie den Namen bzw. Pfad."
		static WinText2	:= "Der Patient wurde bereits als verstorben"
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

	; der Haken bei Vorschau wird entfernt, dann muss ich keine PDF- oder Bild-Anzeigeprogramme abfangen
		VerifiedCheck("Button1", "", "", hGBef, 0)

	;Daten in das Fenster eintragen
		VerifiedSetText("Edit1", ImgName	, GRBFD, 100)
		VerifiedSetText("Edit2", KKText 		, GRBFD, 100)

	;auf Ok drücken - ein mögliches Fehlerfenster abfangen und schliessen (beendet den Importvorgang)
		while WinExist(GRBFD)		{

			VerifiedClick("Button2", GRBFD)
			sleep 300

			If WinExist(GRBFD)
				WinWaitClose, % GRBFD,, 1

			If WinExist("ALBIS ahk_class #32770", WinText1)	 {
				If !VerifiedClick("Button1", "ALBIS ahk_class #32770", WinText1)
					If WinExist("ALBIS ahk_class #32770", WinText1)
						AlbisClosePopups()
				return 3
			} else
				return 1


			If (A_Index > 2)
				return 4
			else
				Sleep 500

		}

return 1
}
;03
AlbisImportierePdf(PdfName, KKText:="", DocDate:="") {                          	;-- Importieren einer PDF-Datei in eine Patientenakte

	/* AlbisImportierePdf()

		Abhängigkeiten: 		                  	- Addendum (globales Object) enthält Addendum Skripteinstellungen
		Datum des Karteikarteneintrages: 	- wird entweder aus DocDate oder wenn leer dem Dateierstellungsdatum entnommen
		                                                   	- Ist DocDate nicht im Format dd.MM.yyyy wird das Erstellungsdatum der Datei verwendet
		Rückgabewert:                            	- entweder das benutzte Karteikartendatum oder 0 bei Mißerfolg

															-	letzte Änderung: 26.09.2021

	 */

		global KK

		ScanKuerzel := Addendum.PDF.ScanKuerzel ? Addendum.PDF.ScanKuerzel : "scan"

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

	; die Karteikarte aktivieren
		AlbisActivate(1)
		AlbisKarteikarteAktivieren()

	; Eingabefokus in die Karteikarte ins Editfeld für Kürzeleingaben
		hInput := AlbisKarteikarteEingabe("Kuerzel")
		Kkarte := AlbisGetFocus()
		If (Kkarte.subFocus <> "Kuerzel") {
			AlbisSetzeProgrammDatum()
			PraxTT("Beim Importieren ist ein Problem aufgetreten.`nDer Eingabefocus konnte nicht ins Kürzelfeld gesetzt werden!", "6 2")
			BlockInput, Off
			return 0
		}
		SendRaw, % ScanKuerzel  ; Albis Karteikartenkürzel hinterlegt in der Addendum.ini ScanPool/Scan
		sleep, 100
		SendInput, {Tab}
		WinWait, % "Grafischer Befund ahk_class #32770",, 3

	; Erstellen des Karteikartentextes
		If !KKText
			KKText := xstring.Karteikartentext(PdfName)

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
			KKText := xstring.Karteikartentext(ImageName)

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
; PROTOKOLLE, STATISTIKEN                                                                                                                                                          	(03)
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
		WinWait, % TPWin,, 1
		If !(hTPtmp := GetHex(WinExist(TPWin)))
			return "invoke error"
		SciTEOutput("hTprotWin: " hTprotWin " = " hTPtmp " ?")

	; PRÜFT DAS HANDLE
		If (hTprotWin <> hTPtmp)
			hTprotWin := hTPtmp


	; OPTIONS PARSEN
		Periode := !IsObject(Options) ? Options : Options.Periode

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
		If !Controls("Edit1", "SetText," ErsterTag	, hTprotWin) {
			PraxTT("Das Anfangsdatum (von) konnte nicht gesetzt werden.", "6 3")
			return 1
		}
		If !Controls("Edit2", "SetText," LetzterTag	, hTprotWin) {
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
		MsgBox, 0x1004, % StrReplace(A_ScriptName, ".ahk"), % "Bitte die gemachten Einstellungen überprüfen!`nTagesprotokoll erstellen JA bestätigen. Mit NEIN wird kein Protokoll erstellt."
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
		;~ while !WinExist("ALBIS - [Tagesprotokoll] ahk_class OptoAppClass") {
		while !RegExMatch(WinGetTitle(AlbisWinID()), "ALBIS\s+\-\s+\[Tagesprotokoll") {
				Sleep, 100
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
; ABRECHNUNG - KASSE UND PRIVAT                                                                                                                                              	(06)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAbrechnungAktiv            	(02) AlbisAbrechnungVorbereiten 	(03) AlbisAbrScheinCOVIDPrivat	(04) AlbisBehandlungsliste
; (05) AlbisPrivatliquidation             	(06) AlbisNeuerSchein
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;01
AlbisAbrechnungAktiv(EditFocus:=false) {                                                        	;-- wird eine Kassenabrechnung angezeigt

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
;02
AlbisAbrechnungVorbereiten(Set) {	                                                                	;-- bearbeitet den Dialog Abrechnung vorbereiten

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
;03
AlbisAbrScheinCOVIDPrivat(Quartal="", BASKT="", IK="") {                             	;-- Abrechnungsschein anlegen für Private bei COVID-19 Impfungen

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

		BASKT 	:= !BASKT 	? "38825"     	: BASKT		;~
		IK     	:= !IK       	? "103609999" : IK

		AlbisActivate(1)
		Controls("", "Reset", "")

	; Abbruch wenn Quartal kein Objekt ist (Addendum_Datum.ahk/QuartalTage())
		If !IsObject(Quartal)
			throw A_ThisFunc ": Qartal muss ein Objekt sein (Rückgabeparameter von QuartalTage() / Addendum_Datum.ahk)"

	; Abbruch wenn keine Karteikarte oder die eines gesetzl. Versicherten geöffnet ist
		If (!AlbisKarteikarteAktiv() || AlbisVersArt() <> 1) {
			DetectHiddenText      	, % dht
			DetectHiddenWindows	, % dhw
			return 1
		}
	; prüft ob Coronaabrechnungsschein schon angelegt wurde
		If AlbisAbrechnungsscheinVorhanden(Quartal.Quartal "/" SubStr(Quartal.Jahr, -1)) {
			DetectHiddenText      	, % dht
			DetectHiddenWindows	, % dhw
			return 1
		}

	; blockierende Fenster schließen
		AlbisIsBlocked()

	; nochmal prüfen ob Fenster geöffnet sind
		If WinExist("Ersatzverfahren ahk_class #32770")
			VerifiedClick("Abbruch", "Ersatzverfahren ahk_class #32770")
		If WinExist("Rechnung für ahk_class #32770")
			VerifiedClick("Abbruch", "Rechnung für ahk_class #32770")

	; evtl. Modifier-Tasten lösen
		SendInput, {LControl Up}{LShift Up}
		sleep 200

	; Strg+Shift gedrückt halten bis das Fenster: Rechnung für <...,...> erscheint
		ControlSend,, {LControl Down}{LShift Down}, % "ahk_class OptoAppClass"
		sleep 200

	; sendet Menubefehl für "Neuer Schein"
		PostMessage, 0x111, 32788,,, ahk_class OptoAppClass

	; mögliche blockierende Dialogfenster schließen
		while !(hneuerSchein := WinExist("Rechnung für ahk_class #32770")) {
			if WinExist("ALBIS ahk_class #32770", "existiert noch eine nicht ausgedruckte Privatrechnung")
				VerifiedClick("Ja", "ALBIS ahk_class #32770", "existiert noch eine nicht ausgedruckte Privatrechnung")
			else If WinExist("ALBIS ahk_class #32770", "Es ist ein alternativer Rechnungs")
				VerifiedClick("Nein", "ALBIS ahk_class #32770", "Es ist ein alternativer Rechnungs")
		}

	; gedrückte Tasten freigeben
		SendInput, {LControl Up}{LShift Up}
		sleep 200

	; Eintragungen in "Rechnung für <...,...>" vornehmen
		;~ hneuerSchein := WinExist("Schein von ahk_class #32770")
		VerifiedClick("Abrechnungsschein",,, hneuerSchein)
		SendMessage, 0x00F0,,, Button3, % "ahk_id " hneuerSchein
		If !(ischecked := ErrorLevel) {
			VerifiedClick("Abrechnungsschein", "Rechnung für ahk_class #32770")
			SendMessage, 0x00F0,,, Button3, % "ahk_id " hneuerSchein
			If !(ischecked := ErrorLevel) {
				VerifiedClick("Button3", "Rechnung für ahk_class #32770")
				SendMessage, 0x00F0,,, Button3, % "ahk_id " hneuerSchein
				ischecked := ErrorLevel
				while (!ischecked) {
					If (A_Index = 6) {
						MsgBox, 0x1000, % "Achtung [2]", % "Okay Abbruch nach 5 Versuchen!", 10
						return 0
					}
					MsgBox, 0x1000, % "Achtung [1]", % "Bitte Abrechnungsschein auswählen"
					SendMessage, 0x00F0,,, Button3, % "ahk_id " hneuerSchein
					ischecked := ErrorLevel
				}
			}
		}

		sleep 200
		;~ SciTEOutput(WinGetTitle(hneuerSchein) "- Checkbox Abrechnungsschein: " (!isChecked ? "nicht":"") " gesetzt")
		;~ If !isChecked
			;~ MsgBox, 0x1000, % "Achtung [2]", % "Bitte Abrechnungsschein auswählen"

		Controls("", "Reset", "")
		hErsatz	:= Controls("", "ControlFind, Ersatzverfahren, Button, return hwnd", hneuerSchein)
		res2    	:= VerifiedClick("",,, hErsatz)

	; Ersatzverfahren Dialog abfangen
		PraxTT("Warte auf Dialog '" (ersatzV := "Manuelle Eingabe der Versichertendaten") "' ", "6 0")
		WinWait, % ersatzV " ahk_class #32770",, 6
		If !(hEingabe := WinExist(ersatzV " ahk_class #32770")) {
			DetectHiddenText      	, % dht
			DetectHiddenWindows	, % dhw
			return 0
		}
		sleep 300

	; Schein im Ersatzverfahren anlegen
		VerifiedCheck("eGK",,, hEingabe, false)				        			; eGK aus
		VerifiedSetText("Edit1"	, BASKT                  	, hEingabe)		; VKNR
		VerifiedSetText("Edit3"	, IK                       	, hEingabe)		; IK-Nummer
		VerifiedSetText("Edit5"	, "BAS"                    	, hEingabe)		; Kasse
		VerifiedSetText("Edit9"	, Quartal.TDBeginn  	, hEingabe)		; Gueltig von
		VerifiedSetText("Edit10"	, Quartal.TDEnde  	, hEingabe)		; bis
		VerifiedClick("Button6",,, hEingabe, false)                             	; M (itglied)

	; Ersatzverfahrendialog schliessen
		PraxTT("Schließen des Dialogfensters 'Ersatzverfahren' abwarten....", "10 0")
		VerifiedClick("OK",,, hEingabe)                                           	; OK
		Loop {
			sleep % (A_Index=1 ? 300:100)
			If WinExist("Abgleich Daten der Chipkarte mit dem Schein ahk_class #32770") {
				WinWaitClose, % "Abgleich Daten der Chipkarte mit dem Schein ahk_class #32770",, 30
				If (WinExist("Abgleich Daten der Chipkarte mit dem Schein ahk_class #32770") && !WinExist("Ersatzverfahren ahk_class #32770")) || (A_Index > 100) {
					PraxTT("", "off")
					DetectHiddenText      	, % dht
					DetectHiddenWindows	, % dhw
					return 1
				}
				VerifiedClick("OK",,, hEingabe)                                           	; OK
			}
			If (A_Index > 60)
				break
		}
		WinWaitClose, % "Ersatzverfahren ahk_class #32770",, 3

	; Rechnung für ... schließen und Schein ist angelegt.
		PraxTT("restliche Dialoge abwarten....", "2 0")
		VerifiedClick("OK",,, hEingabe)                                          	; OK
		VerifiedClick("OK", "Rechnung für ahk_class #32770")
		WinWaitClose, % "Rechnung für ahk_class #32770",, 3
		AlbisKarteikarteZeigen()
		PraxTT("", "off")

	; Fenstereinstellungen zurücksetzen
		DetectHiddenText      	, % dht
		DetectHiddenWindows	, % dhw

return 1
}
;04
AlbisBehandlungsliste(Options:="") {                                                               	;-- Behandlungslisten anzeigen/speichern

	; letzte Änderung: 27.12.2021

	; Dialog Behandlungsliste aufrufen
		If !(hLiqui := AlbisPrivatliquidation("Behandlungsliste")) {
			MsgBox, 0x1000, Addendum für Albis on Windows, % "Der Aufruf der Behandlungsliste ist gescheitert! " hLiqui ", " IsObject(hLiqui), 6
			return 0
		}

		; hLiqui	:= Albismenu(AMenu, ["Behandlungsliste ahk_class #32770", "", "ALBIS ahk_class #32770", "Zugang verweigert"])
		If IsObject(hLiqui) {
			MsgBox, 0x1000, Addendum für Albis on Windows, % "Sie haben nicht die Berechtigung`ndie Behandlungsliste anzuzeigen!", 6
			VerifiedClick("Button1", "ALBIS ahk_class #32770", "Zugang verweigert")
			return 0
		}

		If !IsObject(Options) {
			MsgBox, 0x1000, Addendum für Albis on Windows, % "AlbisBehandlungsliste(Options): Parameter 1 (Options) muss ein Objekt mit Einstellungsdaten sein!", 6
			return
		}

	; Steuerelement Felder anhand der übergebenen Optionen setzen
		doCount := 0, done := 0
		For do, what in Options {

			doCount ++
			switch do 	{

				case "Bearbeitung":
						done ++
						If (what = "alle Ärzte")
							VerifiedCheck("Button2",,, hLiqui)
						else if (what = "Arztgruppe")                 ; Auswahl von Listboxeinträgen fehlt noch
							VerifiedCheck("Button3",,, hLiqui)

				case "Rechnungsfilter":
						done ++
						If RegExMatch(do, "i)(Alle|Nur\sBG|Alle\sohne\sBG|DKGNT|EMB)")
							VerifiedClick(do, hLiqui)

				case "Arztname":
						done ++
						Control, ChooseString, % Options[Arztname], ListBox1, % "ahk_id " hLiqui

				case "Sortierung":
						done ++
						If (what = "Patienten-Nr.")
							VerifiedCheck("Button25",,, hLiqui)
						else if (what = "alphabetisch")
							VerifiedCheck("Button26",,, hLiqui)

				case "Gruppieren":
						done++
						If (what = "Status")
							VerifiedCheck("Button30",,, hLiqui)
						else if (what = "Arzt")
							VerifiedCheck("Button31",,, hLiqui)
						else if (what = "nicht gruppieren")
							VerifiedCheck("Button33",,, hLiqui)

				case "Behandlungsart":
						done++
						VerifiedCheck("Button12",,, hLiqui, InStr(what, "ambulant") 	? true : false) 	; Check
						VerifiedCheck("Button13",,, hLiqui, InStr(what, "stationär") 	? true : false)

				case "Anlegedatum von":
						done++
						VerifiedSetText("Edit2", what, hLiqui)

				case "Anlegedatum bis":
						done++
						VerifiedSetText("Edit3", what, hLiqui)

				case "Kassenart":
						done++
						If (what = "beide")
							VerifiedCheck("Button16",,, hLiqui)
						else if (what = "nur Kasse")
							VerifiedCheck("Button17",,, hLiqui)
						else if (what = "nur Privat")
							VerifiedCheck("Button18",,, hLiqui)

			}

		}

	; Pause
		MsgBox, 0x1000, Addendum für Albis on Windows, % "erledigt: " done "/" doCount "`nMit OK gehts weiter!"

	; Dialog durch drücken von "OK" schliessen und das schliessen abwarten
		VerifiedClick("OK", "Behandlungsliste ahk_class #32770",,, true)

	; Datenaufbereitungsdialog abwarten
		PraxTT("Datenaufbereitung wird abgewartet.", "0 2")
		WinWait, % "bitte warten ahk_class #32770",, 1
		If WinExist("bitte warten ahk_class #32770")
			WinWaitClose, % "bitte warten ahk_class #32770",, 15
		while (!CheckWindowStatus(AlbisWinID(), 200) || A_Index < 20) ;|| A_Index <
			Sleep, 100

	; Speichern der Liste, wenn angefordert
		sPath := PrivatListe.savePath
		SplitPath, sPath,, sDir, sExt, sNameNoExt
		If FilePathExist(sDir) {
			PraxTT("Dialogverarbeitung`nDer Listeninhalt wird gespeichert unter`n>>" sNameNoExt ".csv<<", "0 2")
			ListFileName := AlbisListeSpeichern("Behandlungsliste" , sDir "\" sNameNoExt, "csv", true)
		}

	; Schliessen der Liste, wenn angefordert
		If PrivatListe.Schliessen {
			PraxTT("Die Behandlungsliste wird geschlossen.", "0 2")
			If (hPList := AlbisMDIChildHandle("Behandlungsliste"))
				result := AlbisMDIChildWindowClose("Behandlungsliste") ; result = 1 - MDIChild konnte geschlossen werden
		}

return result
}
;05
AlbisPrivatliquidation(PrivatListe) {                                                                    	;-- unkompliziert Aufruf für alle Privat/Listen

	; z.B. PrivatListe := "Behandlungsliste"

	static	Privat := {"Behandlungsliste"	:{	"cmd"	: 32891
												     	,  	"name"	: "Behandlungsliste"
											     		,  	"titles"	: ["Behandlungsliste ahk_class #32770", "", "ALBIS ahk_class #32770", "Zugang verweigert"]}
				    	,	"OffenePosten"    	:{	"cmd"	: 32892
									    				, 	"name"	: "Offene Posten"
										      			, 	"titles"	: ""}}

	; Liste ist schon geöffnet, dann aktualisieren
		If (hLiquidation := AlbisMDIChildHandle(PrivatListe)) {
			PraxTT("aktiviere die bereits geöffnet Liste -" PrivatListe "-", "0 2")
			Sleep, 500
			AlbisMDIChildActivate(PrivatListe)
			PraxTT("Sende eine Aktualisierungsanfrage an die Liste", "0 2")
			Sleep, 500
			ControlSend,, % "{F5}", % "ahk_id " hOPosten
		}
	; Aufruf per Menubefehl
		else 		{
			PraxTT("Öffne die Liste`n📄 " PrivatListe " 📄`nper Menuaufruf.", "0 2")
			If !(hLiquidation := Albismenu(Privat[PrivatListe].cmd, Privat[PrivatListe].titles)) {
				MsgBox, 0x1000, Addendum für Albis on Windows, % "Die ᔓ" Privat[PrivatListe].name "ᔕ Liste konnte nicht aufgerufen werden." , 6
				return 0
			}
		}


return hLiquidation
}
;06
AlbisNeuerSchein(Zeigen=true) {                                                                    	;-- behandelt den Dialog "Neuen Schein für ... aufnehmen"

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

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; SONSTIGE AUTOMATISIERUNGSFUNKTIONEN                                                                                                                              	(15)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisWaitAndActivate            	(02) AlbisNeuStart                      	(03) AlbisCaveVonToolTip          	(04) AlbisHotKeyHilfe
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
AlbisNeuStart(Client, User="", Pass="", CallingProcess="", AutoStart=0) {                 	;-- startet Albis

	; letzte Änderung 18.11.2021

	;{ Variablen und Einstellungen

	  ; globale Variablen / Thread Einstellungen
		SetTitleMatchMode, 2
		DetectHiddenWindows, On
		DetectHiddenText, On

	  ; schaut nochmal nach ob Albis läuft und holt sich das Fensterhandle
		result := 0
	;}

	  ; result = 0 wenn kein Albis Prozess läuft , andernfalls die PID(Prozess-ID) des Albis Prozesses
		result:= WMIEnumProcessExist("Albis")
		result+= WMIEnumProcessExist("AlbisCS")
		If result
			return -1

		SetTimer, SplashOff, -30000
		SplashTextOn, 400, 50, % "Addendum für Albis on Windows", % "Starte Albis auf dem Computer " Client "`ngenutzter Loginname: " AlbisDefaultLogin("User")
		SplashPraxID := WinExist("Addendum", "Starte Albis")
		WinMove, % "ahk_id " SplashPraxID,, A_ScreenWidth - 400, A_ScreenHeight - 120

	;{ Autostart - Abbruch
		If !AutoStart {                                                                                                             ;--wenn kein Autostart (Auto=0) , dann kommt hier eine Nachfrage ob Albis wirklich gestartet werden soll

			MsgBox, 4, % "Addendum für Albis on Windows", % "Die Albis Praxissoftware ist nicht gestartet!`nAuf Ja wird " CallingProcess " Albis gestartet."
			IfMsgBox, No
				return -2
			IfMsgBox, TimeOut
				return -3

		}
	;}

	;{ Albis Start- und Verifizierungsprozeduren

		;{ ------------- falls Albis nicht gestartet ist, dann wird es jetzt gestartet
            If !AlbisWinID()			{
				Run, % q . Addendum.Albis.MainPath "\" Addendum.Albis.Exe . q, % Addendum.Albis.LocalPath, UserErrorLevel, AlbisPID
				If LE := (EL := ErrorLevel) ? A_LastError : ""
				sleep 5000
				If (!AlbisPID() || EL) {
					FileAppend, % TimeCode(1) ": AlbisNeuStart " ( EL ? "ErrorLevel: " EL " | LastError: " LE " | AlbisPID: " AlbisPID
										: " fehlgeschlagen ohne ErrorLevel keine AlbisPID") "`n", % Addendum.LogPath "\ErrorLogs\Errorbox.txt", UTF-8
					return -4
				}
            }
			else
				return -4
		;}

		;{ ------------ wartet auf die Erstellung des Albishauptfenster ---

            WinWait, ahk_class OptoAppClass
           	WinGetPos, ASx, ASy, ASw, ASh, % "ahk_id " AlbisWinID()

		 ; Mitteilung das Albis gestartet wurde
            if !AutoStart {
				ControlSetText, static1, % "Albis ist gestartet!`nWarte auf das Albis Login Fenster", % "ahk_id " SplashPraxID
				WinMove, % "Addendum", % "Starte Albis", % Asy+Ash, % Asx//2 - 200
            }
		;}

		;{ ------------- Albis is blocked? --- welches Fenster blockiert?
		; prüfe ob das Login schon stattgefunden hat, wenn nicht führe den folgenden Code aus
			AlbisCloseLastActivePopups()

			  ;Benachrichtigung
            	if !AutoStart
					ControlSetText, static1, gebe Logindaten ein, starte Albis

			  ;das Loginfenster ist da - greift sich die ID ab um Nutzer und Passwort ohne SendInput in die Felder einzutragen
            	If (hLogin:= WinExist("ALBIS - Login ahk_class #32770"))	{
					If VerifiedSetText("Edit1", AlbisDefaultLogin("User"), hLogin, 50)
						If VerifiedSetText("Edit2", AlbisDefaultLogin("Password"), hLogin, 50)	{
							res := VerifiedClick("Ok",,, hLogin, true)
							SciTEOutput("Albis-Login operation [" res "]: " (res ? "success ":"failure "))
						}
					if !AutoStart
						ControlSetText, static1, % "Albis-Login abgeschlossen.`nRufe das Wartezimmer auf.", % "ahk_id " SplashPraxID
				}
		;}

		;{ -------------- Wartezimmer öffnen --- falls noch nicht geöffnet
          	;Benachrichtigung
				if !AutoStart
					ControlSetText, static1, % "Öffne das Wartezimmer", % "ahk_id " SplashPraxID

			;Vorbereitungen für das Öffnen des Wartezimmer
				WinActivate, % "ahk_class OptoAppClass"
            	PostMessage, 0x111, 32924,,, % "ahk_id " AlbisWinID()

			; das Wartezimmer wird per WM_Command aufrufen - Patient/Wartezimmer = 32924
				While (!AlbisMDIWartezimmerID() && A_Index <= 20) {
					If (A_Index > 1)
						sleep 300
					PostMessage, 0x111, 32924,,, % "ahk_id " AlbisWinID()
				}

           	;wurde zu lange gesucht, gab es wahrscheinlich ein Problem, deshalb bricht die Funktion hier ab
				WartezimmerID := AlbisMDIWartezimmerID()
				If WartezimmerID {
					WinMaximize, % "ahk_id " WartezimmerID
					ControlSetText, static1, % "Wartezimmer geöffnet und maximiert!", % "ahk_id " SplashPraxID
				} else {
					ControlSetText, static1, % "Wartezimmer konnte nicht geöffnet werden!", % "ahk_id " SplashPraxID
				}

				SetTimer, SplashOff, -3000
		;}

	;}

return

SplashOff: ;{
	SplashTextOff
	ToolTip,,,, % TTnum
	SetTimer, SplashOff, Off
return ;}
}
;03
AlbisCaveVonToolTip(compname, hCave) {                                                            	;-- zeigt Infos an was man neues im Cave Fenster machen kann

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

	WinGetPos, CvX, CvY, CvW, CvH, ahk_id %hCave%
	;wird das Fenster nicht gefunden, soll er die Funktion verlassen
		If (CvX="") {
                        return
                        		}

	;wenn die Größe nicht passt. Wird die Größe angepasst. Auf manchen Rechnern fehlt immer die letzte Zeile.
	If (CvX <> NewX OR CvY <> NewY OR CvW <> NewW OR CvH <> NewH) {
		WinMove, ahk_id %hCave%,, %NewX%, %NewY%, %NewW%, %NewH%
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

		AlbisWinID := 	!AlbisWinID ? AlbisWinID() : AlbisWinID

		If (autoclose = 0)   		{
			WinGetTitle, ltitle, % "ahk_id " (phwnd := DLLCall("GetLastActivePopup", "uint", AlbisWinID))

			return {"isblocked"	: InStr(ltitle, "Bitte warten...") 	? 2
										: WinIsBlocked(AlbisWinID) 	? 1 : 0
										, "blockWinT": ltitle
										, "errorStr": errorStr}
		}
		else if (autoclose = 1)	{
				;platzhalter erstmal
		}
		else if (autoclose = 2 && WinIsBlocked(AlbisWinID))
			AlbisClosePopups()

return
}
AlbisClosePopups() {                                                                                               	;-- schließt alle PopUp-Fenster in Albis

  ; letzte Änderung 13.05.2022

	AlbisWinID := GetHex(AlbisWinID())

	Loop		{

	  ; welches Fenster blockiert Albis? ---
		phwnd := GetHex(DllCall("GetLastActivePopup", "uint", AlbisWinID))
		If (phwnd = AlbisWinID)   ; kein Fenster blockiert Albis
			break

	  ; 3 verschiedene Versuche das Fenster zu schließen von "sanft" bis "kraftvoll"
		SendMessage, 0x112, 0xF060,,, % "ahk_id " phwnd  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
		t.= " EL1: " ErrorLevel "`n"
		If ErrorLevel	{
			WinClose, % "ahk_id " phwnd
			t .= " EL2: " ErrorLevel "`n"
			If ErrorLevel	{
				WinKill, % "ahk_id " phwnd               ; ErrorLevel wird nicht gesetzt
				WinWaitClose, % "ahk_id " phwnd,, 1
				WinGetTitle, ltitle, % "ahk_id " phwnd
				If ltitle	{
					SciTEOutput(A_ThisFunc " - Can not close entire window (" ltitle "):`n" t )
					return 0
				}
				else
					return 1
			}
		}

		sleep 150

	; Loop endet hier wenn Albis nicht mehr blockiert ist
		If !WinIsBlocked(AlbisWinID)
			break

	}

return 1
}
;06
AlbisCloseLastActivePopups(AlbisWinID:="") {                                                            	;-- schließt alle PopUp-Fenster

	If !WinIsBlocked(AlbisWinID := AlbisWinID())
		return

	Loop	{

		; mehr als 10 zu schließende Dialoge gibt es nicht (Sicherheits return , damit die Schleife irgendwann verlassen wird)
			If (A_Index > 9)
				return 0         	; 2 = Problem beim Schliessen

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
					WinWaitClose, % "ahk_id " phwnd,, 1
					WinGetTitle, ltitle, % "ahk_id " phwnd
					If ltitle	{
						SciTEOutput(A_ThisFunc " - Can not close entire window (" ltitle "):`n" t )
						return 0
					}
					else
						return 1
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
AlbisDefaultLogin(key) {                                                                                         	;-- liest Default User und Password

	; , gibt Array zurück (als Class wäre das auch nicht schlecht)

	static DLogIn

	If !IsObject(DLogIn) {
		IniRead, Comp, % Addendum.Ini, Computer, % compname
		Utmp	:= StrSplit(Comp, "|")
		DLogin	:= {"User": Utmp.1, "Password": Utmp.2}
	}

return DLogin[key]
}
;08
AlbisAutoLogin(MsgBoxZeit:=10) {                                                                           	;-- loggt den jeweiligen Nutzer automatisch ein

		; letzte Änderung: 17.11.2021 - Ablauf beschleunigt
		; der Click mittels einer anderen Technik hat das Login-Fenster nicht geschlossen und den Hook für die Erkennung des Login-Fenster erneut ausgelöst

		static LoginTime    	:= 0
		static LoginNo       	:= false
		static AlbisLoginClass := "ALBIS - Login ahk_class #32770"

	;wenn ein automatischer Login abgelehnt wurde, wartet Addendum 2 Minuten bevor es erneut nachfragt
		If LoginNo && (A_TickCount-LoginTime > 120000)
			LoginNo := false

	;Routine zum Einloggen, nach drücken auf Nein erkennt der Hook das Loginfenster erneut, deswegen muss eine erneute Abfrage unterdrückt werden
		If !LoginNo && WinExist(AlbisLoginClass)		{

			MsgBox, 4, Addendum für AlbisOnWindows, % "Soll das automatische Einloggen`ndurchgeführt werden?", % MsgBoxZeit
			LoginNo 	:= true, LoginTime := A_TickCount
			IfMsgBox, No
				return
			LoginNo 	:= false, LoginTime := A_TickCount


			If VerifiedSetText("Edit1", AlbisDefaultLogin("User"), AlbisLoginClass, 50)
				If VerifiedSetText("Edit2", AlbisDefaultLogin("Password"), AlbisLoginClass, 50)	{
					res := VerifiedClick("Ok", AlbisLoginClass,,, true)
					SciTEOutput("Albis-Login operation [" res "]: " (res ? "success ":"failure "))
				}

		; Backupvorgang wenn VerifiedClick versagt
			If (hcontrol:= ControlGet("HWND", "", "Button1", AlbisLoginClass))	{
				c:= GetWindowSpot(hcontrol)
				MouseMove, % c.X +5 , % c.Y + 5
				MouseClick, Left
			}

			AlbisWZOeffnen()
		}

return
}
;09
AlbisLogout(CloseAlbis:=false, ForceClose:=true) {                                                	;-- logt aus oder beendet Albis per Menuaufruf

	; Funktion erweitert Dialoginteraktion um nicht gespeicherte Eingaben eigentlich zu sichern.
	; wählt man jedoch während des Logouts "JA" im Dialog mit der Nachfrage die Eingabe zu sichern dann erscheint das Loginfenster
	; und keine weitere Karteikarte wurde geschlossen. Damit der Loginvorgang mit einem leeren Albisfenster abgeschlossen werden kann
	; wählt die Funktion "Nein" aus.
	; letzte Änderung: 17.11.2021

	; Patient/Beenden oder Logout
	Albismenu(CloseAlbis ? 57665 : 32922)

	stime := 100
	timeToStop := Floor((40*1000)/stime)
	Loop {

		If (hZeile := WinExist("ALBIS ahk_class #32770", "Die aktuelle Zeile wurde nicht gespeichert")) {
			If (!changed && VerifiedClick((!CloseAlbis && ForceClose ? "Nein" : "JA"),,, hZeile, true)) {
					SendInput, {TAB}
					sleep 100
			} else if changed {
				VerifiedClick("Nein",,, hZeile, true)
			}
			changed := false
		}
		else If (hChange := WinExist("ALBIS ahk_class #32770", "Änderungen des Befundkürzels")) {
			If VerifiedClick("OK",,, hChange, true) {
				changed := true
				SendInput, {Esc}
				sleep 100
			}
		}
		else if (hExit := WinExist("ALBIS ahk_class #32770", "Hiermit beenden Sie ALBIS")) {
			VerifiedClick("OK",,, hExit, true)
		}

		If CloseAlbis && !WinExist("ahk_class OptoAppClass")
			return true
		else if !CloseAlbis && WinExist("ALBIS - Login ahk_class #32770")
			return true

		sleep % stime
		If (A_Index > timeToStop)
			return false

	}

return false
}
;10
AlbisActivate(waitingtime) {                                                                                       	;-- aktiviert das Albishauptfenster
	WinActivate		, % "ahk_class OptoAppClass"
	WinWaitActive	, % "ahk_class OptoAppClass",, % waitingtime
return ErrorLevel
}
;11
AlbisCopyCut() {                                                                                                    	;-- speichert alles was mit Strg+c oder Strg+x in Albis kopiert wird in eine Datei

		; letzte Änderung: 17.11.2021

		static AlbisCText_old

		while (!AlbisCText && A_Index < 4) 	{
			ControlGetFocus, cFocused, % "ahk_class OptoAppClass"
			If cFocused
				ControlGet, AlbisCText, Selected,, % cFocused, % "ahk_class OptoAppClass"
			If AlbisCText
				break
			Sleep, 100
		}

		If (StrLen(AlbisCText)>0 && AlbisCText<>AlbisCText_old)		{
			ControlGet, hFocused, hwnd,, % cFocused, % "ahk_class OptoAppClass"
			hOwner := GetWindow(hwnd, 4)
			WinGetTitle	, WTitle 	, % "ahk_id " hOwner
			WinGetClass	, WClass	, % "ahk_id " hOwner
			AlbisCText_old	:= AlbisCText
			FileAppend, % "[" A_YYYY "." A_MM "." A_DD " " A_Hour ":" A_Min ":" A_Sec "] [" WTitle " | " WClass " | " cFocused "]`n" AlbisCText "`n", % Addendum.LogPath "\CopyAndPaste.log"
		}
		else
			AlbisCText := ""

return AlbisCText
}
;12
AlbisIsElevated() {                                                                                              		;-- stellt fest ob Albis mit UAC Virtualisierung gestartet wurde

	Process, Exist, albisCS.exe
	If !(AlbisPID := ErrorLevel)	{
		Process, Exist, albis.exe
		albisPID := ErrorLevel
	}

	If AlbisPID
		return IsProcessElevated(albisPID) ? true : false

return 2 ; - means Albis is not running
}
;13
AlbisSelectAll() {                                                                                                     	;-- kompletten Text markieren

	SendInput, {PgUp}
	sleep, 100
	SendInput, {LControl Down}{LShift Down}{Down}
	sleep, 100
	SendInput, {LControl Up}
	sleep, 50
	SendInput, {LShift Up}

}
;14
CheckAISConnector(LabConnector:="AIS Connector ") {                                				;-- prüft ob das Laborverbindungsprogramm läuft und startet es bei Bedarf neu

	; durch Verwendung von A_Username müsste diese Funktion auf allen Computern unter allen Benutzernamen funktionieren
	; AIS Connector


	for Prozess in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	   prozesse .= Prozess.Name "|"

	If !Instr(prozesse, "LabConnector") {

		TrayTip, % "Hinweis!", % "Der AIS Connector ist`nnicht mehr verfügbar`nund wird neu gestartet!", 4
		Run, % "C:\Users\" A_UserName "\Desktop\AIS Connector.appref-ms"
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

	If !conditions
		Conditions := "contraction=i)(lp|bg), activeWinType=i)Privatabrechnung"

	RegExReplace(conditions, "[+\-]", "", conditionsMax)
	conditions     	:= StrSplit(conditions, ",")
	conditionsMax	:= conditions.Count()
	condMatched 	:= 0

	For i, condition in conditions {

		RegExMatch(condition, "^\s*(?<lc>[+\-])*\s*(?<focus>\w+)?\=(?<name>[A-Za-z\|]+)\s*$", p)
		If (pFocus = "contraction") {
			albisfocus := AlbisGetFocus()
			SciTEOutput("childs: " albisfocus.childs.Kuerzel.Text)
			If RegExMatch(albisfocus.childs.Kuerzel.Text, pName)
				condMatched ++
		}
		else if (pFocus = "activeWinType")
			If RegExMatch(AlbisGetActiveWindowType(), pName)
				condMatched ++

	}

return condMatched = conditionsMax ? true : false
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


