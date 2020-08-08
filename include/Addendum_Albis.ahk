;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 05.07.2020 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
return
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                  	BESCHREIBUNG DER FUNKTIONSBIBLIOTHEK ZUR RPA SOFTWARE -- ADDENDUM FÜR ALBIS on WINDOWS --
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ;{
;
;       FUNKTIONDESIGN:
;
;       Die Funktionen sollen eine einfach durchzuführende und möglichst fehlerfreie RPA (robotic process automation) ermöglichen.
;   	Nachfolgend beschreibe ich das dafür verwendete Design:
;
;        - Datenerfassung                               	 -	erfolgt über das Auslesen des Albisfenstertitel, des Inhaltes eines Fenster oder durch Ermittlung eines bestehenden Eingabefocus
;
;        - das Albis Fensterhandle                   	 -	wird bei jedem Aufruf neu ermittelt, damit ein neues Albishandle, z.B. nach einem Absturz oder Neustart von Albis erkannt wird
;
;        - wenige Parameter		                     	 -
;
;        - keine Wartezeit nach Funktionsende 	 - 	nachdem eine Automatisierungsfunktion beendet ist, ist kein sleep Befehl notwendig um auf Fenster oder anderes zu warten
;                                                        			  	Die nächsten Befehle können unmittelbar erfolgen.
;                                                                	    Es ist nicht nötig die Dauer eines sleep-Befehles innerhalb einer Funktion zu ändern !!
;                                                        			  	Die Zeiten haben ich durch tägliche Nutzung meiner Skripte im Laufe der Jahre optimiert. In meiner Praxis gibt es schnelle sowie
;                                                                  	  	auch langsame Computer. Auf allen funktioniert das Timing der Skripte gut.
;
;        - Ausführungsüberprüfung                 	- 	Funktionen welche Menupunkte oder Fenster in Albis öffnen, überprüfen ob Albis im Moment durch geöffnete Fenster blockiert ist
;                                                                	  	und schliessen diese automatisch. Die meisten dieser Funktionen prüfen auch ob das zu erwartende Dialogfenster geöffnet wurde.
;                                                        			  	Damit wird verhindert das weitere Automatisierungsroutinen nicht ins Leere laufen.
;                                                                	- 	Daten die aus einem Skript in das Albisfenster übergeben werden, werden sofort wieder ausgelesen um den Erfolg der Über-
;                                                                	  	gabe zu überprüfen.
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
;	-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
;
;                                                                                        	WARUM KEINE KLASSENBIBLIOTHEK ?
;
;	Ich bin kein IT-Profi. Das notwendige Wissen habe ich mir über das Autohotkey-Forum angeeignet. Die Bereitstellung von Klassenbibliotheken, anstatt wie hier von
;	einzelnen Funktionen, ist mir bis heute fremd geblieben.
;
;	IXIKO 2019
;
;	-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  - ;}

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; HAUPTFENSTER INFO's                                                                                                                                                                                                                            	(05)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetActiveWinTitle          	(02) AlbisWinID                         	(03) AlbisPID                                 	(04) AlbisExist                            	(05) AlbisStatus
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;01
AlbisGetActiveWinTitle() {						                                    	;-- ermittelt den Fenstertitel v. Albishauptfenster, nützlich um z.B. zu bestimmen mit welchem Patienten gerade gearbeitet wird
	WinGetTitle, iWT, ahk_class OptoAppClass
	RegExMatch(iWT, "(?<=\[).*(?=\])", Match)
	For key, val in StrSplit(Match, "/") 			     					    		;entfernt überflüssige Spacezeichen vor und hinter dem String, CG hatte die Darstellung des Albis Fenstertitel verändert
		LWT.= Trim(val) . "/"
return RTrim(LWT, "/")
}
;02
AlbisWinID() {                            			                                    	;-- gibt die ID des übergeordneten Albisfenster zurück
	While !(AID := WinExist("ahk_class OptoAppClass"))
	{
			sleep, 50
			if (A_Index > 40)
				break
	}
return GetHex(AID)
}
;03
AlbisPID() {                            				                                    	;-- gibt die Prozeß-ID des Albisprogrammes aus
	WinGet, AlbisPID, PID, ahk_class OptoAppClass
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

	while !CheckWindowStatus(AlbisWinID(), 200)
	{
				MsgBox,1, Addendum für Albis on Windows, Albis scheint nicht zu reagieren!`nBevor Sie weiter machen,`nüberprüfen Sie bitte das Albisprogramm!
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
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; MDI CONTROL FUNKTIONEN                                                                                                                                                                                                                 	(11)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetMDIClientHandle     	(02) AlbisGetWartezimmerID      	(03) AlbisGetAllMDIWin                 	(04) AlbisGetMDIMaxStatus        	(05) AlbisGetActiveMDIChild
; (06) AlbisGetSpecificHMDI           	(07) AlbisGetAllMDITabNames   	(08) AlbisGetHMDITab                   	(09) AlbisActivateMDIChild         	(10) AlbisActivateMDITab
; (11) AlbisCloseMDITab
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;01    	;-- MDI Control Funktionen
AlbisGetMDIClientHandle() {                                                      	;-- ermittelt das Handle des MDIClient (Basishandle für alle Unterfenster)

	global Mdi

	If !IsObject(Mdi)
		Mdi := Object()

	ControlGet, hMdi, HWND,, MdiClient1, ahk_class OptoAppClass
	;Mdi.MDIClientHandle := hMdi

return hMdi
}
;02
AlbisGetWartezimmerID() {                                                        	;-- ermittelt das Handle des Wartezimmerfenster innerhalb des MDI-Controls
return FindChildWindow({Class: "OptoAppClass"}, {Title: "Wartezimmer"}, "Off")				;ID des Wartezimmer-Fenster
}
;03
AlbisGetAllMDIWin() {                                                                	;-- erstellt ein globales Objekt, welches Namen, Klassen und Handles aller geöffneten MDI Fenster enthält

	;Mdi - key:value Object - key ist WinTitle, value ist die Fensterklasse, extra Key ist "MdiHwnd" für das Roothandle des MDI-Fenster
	; global Mdi:= Object() - im aufrufenden Skript / aufrufender Funktion einfügen!

	global Mdi
	Mdi := Object()   ; Mdi-Objekt wird nur von dieser Funktion zurückgesetzt

	hMdi:= AlbisGetMDIClientHandle()
	WinGet, MDIClientWinList, ControlListHWND, % "ahk_id " hMdi
	Loop, Parse, MDIClientWinList, `n
		If InStr(class:= WinGetClass(A_LoopField), "Afx:")
				Mdi[WinGetTitle(A_LoopField)]:= {"class": (class), "ID": (A_LoopField)}

	;Mdi.MDIClientHandle:= hMdi

return Mdi
}
;04
AlbisGetMDIMaxStatus(MDIChild) {                                            	;-- stellt fest ob das gewählte MDI Fenster maximiert und im Vordergrund ist

	;gibt 1 zurück wenn es maximiert ist und im Vordergrund , Beispiel: AlbisGetMdiMaxStatus("Wartezimmer")
	global Mdi

	Mdi := AlbisGetAllMDIWin()

	If RegExMatch(MDIChild, "^0x[\w]+") ;|| RegExMatch(MDIChild, "^\d+$")
	{
		For title, val in Mdi
			If InStr(val.id, MDIChild)
				return DllCall("IsZoomed", "UInt", val.ID)
	}
	else
	{
		For title, val in Mdi
			If InStr(title, MDIChild)
				return DllCall("IsZoomed", "UInt", val.ID)
	}

}
;05
AlbisGetActiveMDIChild() {                                                        	;-- ermittelt das Handle des aktuellen MDI-Childfensters

	global Mdi

	If !IsObject(Mdi)
		Mdi := Object()

	; WM_MDIGETACTIVE:= 0x0229
	hMdi:= AlbisGetMDIClientHandle()
	SendMessage, 0x0229,,,, % "ahk_id " hMdi
	hMdiChild := GetHex(ErrorLevel)
	;Mdi.MDIClientHandle:= hMdi

return hMdiChild
}
;06
AlbisGetSpecificHMDI(MDITitle) {		                                        	;-- ermittelt das Handle eines sub oder child Fenster innerhalb des Albis-MDI-Controls
return GetHex(FindChildWindow({Class: "OptoAppClass"}, {Title: (MDITitle)}, "Off"))
}
;07
AlbisGetAllMDITabNames(MDITitle) {                                         	;-- ermittelt die Namen aller Tabs eines SysTabControls321

  ; Funktion kann ohne die anderen MDI Funktionen aufgerufen werden
  ; Abhängigkeiten: GetHex(), GetControls, ObjFindValue(), ControlGetTabs()

	global Mdi

	oCtrl			:= Object()
	hSpecMDI	:= GetHex(AlbisGetSpecificHMdi(MDITitle))
	oCtrl			:= GetControls(hSpecMDI)
	key			:= ObjFindValue(oCtrl, "SysTabControl321")
	hMdiTab	:= oCtrl[key]["Hwnd"]
	If !key {
		MsgBox,, Addendum für AlbisOnWindows, % "AlbisGetMDITabs (" A_LineNumber "): Could not find SysTabControl321`n for this window :" WinTitle, 5
		return ""
	}

return ControlGetTabs(hMdiTab)
}
;08
AlbisGetHMDITab(MDITitle) {                                                     	;-- ermittelt das Handle eines spezifischen MDI-TabControls

	global Mdi

	oCtrl     	:= Object()
	hSpecMDI	:= GetHex(AlbisGetSpecificHMDI(MDITitle))
	oCtrl			:= GetControls(hSpecMDI)
	key			:= ObjFindValue(oCtrl, "SysTabControl321")
	If !key {
		MsgBox,, Addendum für AlbisOnWindows, % "AlbisGetMDITabs (" A_LineNumber "): Could not find SysTabControl321`n for this window :" WinTitle, 5
		return ""
	}

return oCtrl[key]["Hwnd"]
}
;09
AlbisActivateMDIChild(MDITitle) {		                                        	;-- aktiviert ein MDI-Child Fenster (nicht Tab im MDI-Child!)

	; return value: bei Erfolg das Handle des MDI-Child ansonsten ein Objekt mit dem ErrorLevel, dem letzten Fehler und das MDIClient-Handle
	global Mdi

	hMdi     	:= AlbisGetMDIClientHandle()
	hMdiChild	:= AlbisGetSpecificHMDI(MDITitle)

	SendMessage, 0x222, % hMdiChild,,, % "ahk_id " hMdi

return (ErrorLevel = 0 ? hMdiChild : {"ErrorLevel": ErrorLevel, "LastError": A_LastError, "hMdi": hmdi})
}
;10
AlbisActivateMDITab(MDITitle, TabName:="") {                             	;-- aktiviert ein MDI-Tab, merkt sich welches TabControl zuletzt aufgerufen wurde

	; ruft man die Funktion mehrfach nacheinander unter demselben MDITitle auf, ermittelt er nicht jedesmal alle Handles neu (Speicherung in tabs - Object)
	; Beispiel: result:= AlbisActivateMDITab("Wartezimmer", "Arzt")
	; Rückgabewert ist 0 wenn er keinen passenden Tab finden konnte, >0 wenn die Funktion erfolgreich war

	global Mdi
	static tabs := Object()
	static last_MDITitle, hTab

	If !InStr(last_MDITitle, MDITitle) {
		last_MDITitle := MDITitle
		hTab	:= AlbisGetHMDITab(MDITitle)
		tabs	:= AlbisGetAllMDITabNames(MDITitle)
	}

	If (StrLen(TabName) > 0)
		For idx, tab in tabs
			If InStr(tab, TabName)
				SendMessage, 0x1330, % idx - 1 ,,, % "ahk_id " hTab     ; 0x1330 is TCM_SETCURFOCUS.

return ErrorLevel
}
;11
AlbisCloseMDITab(MDITitle) {                                                    	;-- schließt ein per Titel identifiziertes MDIClient Fenster

	; MDITitle - darf ein Teil des Namens, die komplette Patienten-ID, Geburtsdatum, .... sein
	global Mdi

	If (hMdiChild:= AlbisGetSpecificHMDI(MDITitle)) 	{
		SendMessage, 0x0221, % hMdiChild,,, % "ahk_id " AlbisGetMDIClientHandle() ;WM_MDIDESTROY
		EL := ErrorLevel
		sleep, 400
	}
	else
		EL := 0

return (EL = 0 ? 1 : EL)
}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; INFO's VON ANDEREN STEUERELEMENTEN                                                                                                                                                                                             	(05)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetActiveControl           	(02) AlbisGetActiveWindowType   	(03) AlbisLVContent                       	(04) AlbisGetStammPos             	(05) AlbisAktuellesKuerzel
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;01
AlbisGetActiveControl(cmd) {                                                            	;-- ermittelt den aktuellen Eingabefocus und gibt zusätzliche Informationen zur genauen Identifikation des Controls zurück

	/* 		FUNKTIONSBESCHREIBUNG: AlbisGetActiveControl() letzte Änderung: 20.06.2020
	   *
		* Steuerelemente in Fenster können die gleichen Namen besitzen. Um ein Steuerelement exakt zu identifizieren ermittelt diese Funktion die Eltern(Parent)-Elemente
		* wieviele Eltern ermittelt werden müssen, hängt von der jeweiligen Fensterstruktur ab
		*
		* wozu braucht man das: z.B. um Contextabhängige Hotstrings, Hotkeys oder andere Dateneingaben zu ermöglichen
		*
		* --------------------------------------------------------------------------------------------------------------------------------
		* Bezeichner	|            Ebenenelemente      	|                          Elternelemente						   	   	        	*
		* Bezeichner	|      ClassNN1,ClassNN2,... 	|ParentClass1|ParentClass2			|ParentClass3                	*
		* Dokument	|Edit1,Edit2,Edit3,RichEdit20A1	|#32770		 |AfxFrameOrView	|AfxMDIFrame              	*
		* --------------------------------------------------------------------------------------------------------------------------------
		* 		Bezeichner 			- ihre Bezeichnung für den Steuerelement-Zusammenhang
		*		Ebenenelemente	- eine durch Komma getrennte Liste an ClassNN Namen von Steuerelementen, diese Steuerelemente haben einen inhaltlichen Zusammenhang z.B.
		*		Elternelemente		- eine durch | getrennte Liste an ClassNN Namen von Eltern-Steuerelementen, diese lassen die eindeutige Identifikation der Ebenenelemente zu
		*
		* --------------------------------------------------------------------------------------------------------------------------------
		* WICHTIGES                                                                                                                                        	*
		* --------------------------------------------------------------------------------------------------------------------------------
		* Der Aufruf der Funktion holt immer das Albisfenster in den Vordergrund!

		* - 07.01.2020: AfxFrameOrView und AfxMDIFrame wurden intern in Albis umbenannt, aus AfxFrameOrView90 und AfxMidiFrame90 wurde AfxFrameOrView140 und AfxMidiframe140
		* - 11.07.2020: AfxFrameOrView und AfxMDIFrame wurden intern in Albis umbenannt, aus AfxFrameOrView90 und AfxMidiFrame90 wurde AfxFrameOrView140 und AfxMidiframe140
	   *
	 */

		static AACInit, IdentData
		aacCounter := 0, idx := 2

	; hier kommen noch Identifikationsstrings später hinzu , im Moment brauche ich diese Funktion nur für Eintragungen im Dokumentfenster
		If !AACInit
		{
				IdentData	:= []
				IdentData	:= StrSplit("Dokument|Edit1,Edit2,Edit3,RichEdit20A1|#32770|AfxFrameOrView|AfxMDIFrame", "|")		;Dokument ist meine Bezeichnung für die 4-Controls (Edit1...)
				AACInit 	:= true
		}

	; Ermitteln des aktuellen ControlFocus hwnd und classNN
		hFocused:= Chwnd:= GetFocusedControlHwnd()
		cNN	:= Control_GetClassNN(AlbisWinID(), hFocused)

		If InStr(cmd, "hwnd")
			return GetHex(ChWnd)
		else if InStr(cmd, "classNN")
			return cNN

	; Übereinstimmungen suchen
		firstparent:= DllCall("user32\GetAncestor", "Ptr", hFocused, "UInt", 1, "Ptr")

		If InStr(IdentData[2], cNN)		{

					Loop % (IdentData.MaxIndex() - 2) {					; die ersten zwei Felder gehören nicht zur Fensteridentifikation
					       Chwnd:= DllCall("user32\GetAncestor", "Ptr", Chwnd, "UInt", 1, "Ptr")
                            If InStr(WinGetClass(CHwnd), IdentData[A_Index + 2])
                            			aacCounter++
					}

				; keine Übereinstimmungskette gefunden
					If !(aacCounter = IdentData.MaxIndex() - 2)
                            		return false

				; ab hier können spezielle weitere Befehle programmiert werden
					If InStr(cmd, "content")     	{

                            If InStr(IdentData[1], "Dokument") {
									Loop, 3
										ControlGetText, Edit%A_Index%, Edit%A_Index%, % "ahk_id " firstparent

									ControlGetText, RichEdit, RichEdit20A1, % "ahk_id " firstparent
									return {"FHwnd":GetHex(hFocused), "FClassNN":cNN, "fpar":firstparent,"Identifier":IdentData[1], "Edit1":Edit1, "Edit2":Edit2, "Edit3":Edit3, "RichEdit":RichEdit}
                            }
					}
					else if InStr(cmd, "contraction") 	{
                            If InStr(IdentData[1], "Dokument")
                        			return ControlGetText("Edit3", "ahk_id " firstparent)
					}
					else if InStr(cmd, "identify")     	{
                            return IdentData[1] "|" cNN
					}

		}

return 0
}
;02
AlbisGetActiveWindowType() {                                                         	;-- ermittelt den aktuell bearbeiteten Dokumenttyp - Wartezimmer, Terminplaner, Akte, ....

	;diese Funktion unterscheidet die 3 möglichen Fenstertitel von Albis
	; Patientenakte ("Akte"), Wartezimmer ("WZ") , freie Statistik (fS) oder alles geschlossen ("other")
	; keine Unterscheidung ob ein PopUp Dialog z.B. Rezeptformular geöffnet ist, da es in dieser Funktion ausschließlich um den Inhalt des Fenstertitels geht
	;
	;letzte Änderung: 20.04.2020

	WinGetTitle, LWT, ahk_class OptoAppClass

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
	else if RegExMatch(LWT, "\d\d\.\d\d\.\d\d\d\d")
	{
			ControlGet, hTbar		, Hwnd		,, ToolbarWindow321	, % "ahk_id " AlbisGetActiveMDIChild()
			ControlGet, Auswahl	, Choice	,, ComboBox1				, % "ahk_id " hTbar
			;ControlGet, Auswahl	, Choice	,, ComboBox1				, % "PatientfensterToolbar"

			If InStr(Auswahl, "Abrechnung")
						return "Patientenakte|Abrechnung"
			else
						return "Patientenakte|" Auswahl
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
AlbisGetStammPos() {                                                                    	;-- liest aus der Local.ini die Positionen der Stammdatenfenster aus

	; Abhängigkeiten: \lib\ini.ahk

	obj := Object()

	If !FileExist(LocalIni:= "C:\albiswin\Local.ini")
			If !FileExist(LocalIni:= "C:\albiswin.loc\Local.ini") {
					MsgBox, Addendum für AlbisOnWindows, Die Local.ini befindet sich in keinem der`nStandard-Ordner auf dem Laufwerk (C:)!`nDas Skript (%A_ScriptName%) wird beendet.
					ExitApp
			}


	ini_Load(ini, LocalIni)
	keys := ini_getAllKeyNames(Ini, "STAMM_POS")

	Loop, Parse, keys, `,
	{

				If InStr(val:= ini_getValue(ini, "STAMM_POS", A_LoopField), ",") {

                            val:= StrSplit(StrReplace(val, " ") , ",")
                            obj[(A_LoopField)] := {"x":(val[1]), "y":(val[2]), "w":(val[3]), "h":(val[4])}

				}

	}

return obj
}
;05
AlbisAktuellesKuerzel() {                                                                 	;°~° begonnen

		aControl := {}

		If !InStr(AlbisGetActiveWindowType(), "Karteikarte")
				return 0

		ControlGetFocus, hactiveCtrl, ahk_class OptoAppClass
		If !InStr(hactiveCtrl, "RichEdit201A")
				return 0

		aControl:= AlbisGetActiveControl("content")

}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; DATEN VOM AKTUELLEN PATIENTEN                                                                                                                                                                                                      	(08)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAktuellePatID                 	(02) AlbisCurrentPatient              	(03) AlbisPatientGeschlecht()          	(04) AlbisPatientGeburtsdatum   	(05) AlbisPatientVersicherung
; (06) AlbisVersArt                          	(07) AlbisTitle2Data                   	(08) AlbisAbrechnungsscheinVorhanden
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
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
	data := StrSplit(AlbisGetActiveWinTitle(), "/")
	If Instr(data[5], "Privat")
					Vstat = 1
return (Vstat= 1 ? 1 : 2)
}
;07
AlbisTitle2Data(WTAlbis:="") {                            				         	;-- erstellt ein Objekt aus den Daten des Fenstertitel

	;WTAlbis leer lassen dann liest die Funktion den Fenstertitel
	If !InStr(AlbisGetActiveWindowType(), "Akte")
		return ""

	If (WTAlbis="")
		WTAlbis:= AlbisGetActiveWinTitle()

	Splitstr				:= StrSplit(WTAlbis, "/", A_Space)
	SplitStr2			:= StrSplit(SplitStr[2], "`,", A_Space)

return {"ID": SplitStr[1], "Nn": SplitStr2[1], "Vn": SplitStr2[2], "Gt": SplitStr[3], "Gd": SplitStr[4], "Kk": SplitStr[5]}
}
;08
AlbisAbrechnungsscheinVorhanden(Quartal) {                            	;-- Funktion nutzt das ComboBox1 Element auf der Hauptoberfläche

	; Achtung: das Quartal muss mit folgendem Trennzeichen "/" übergeben sein, wenn z.b. das Quartal:="01/17" ist, kürzt die Funktion die 0 am Anfang weg
	; Rückgabewert ist die Zeilennummer in der ComboBox wenn die Abrechnung existiert - kein Abrechnungsschein vorhanden dann 0

	Quartal:= LTrim(Quartal, "0")
	If !InStr(Quartal, "/") {
		Quartal:= SubStr(Quartal, 1, 1) . "/" . SubStr(Quartal, 2, 2)
	}
	Quartal:= "`(" . Quartal . "`)"
		;MsgBox, % Quartal

	AlbisActivate(1)
	ControlGet, ComBox, List, , ComboBox1, % "ahk_id " AlbisWinID()			;ahk_class OptoAppClass
	Loop, Parse, ComBox, `n
	{
		zeile:= A_LoopField
		if ( Instr(zeile, Quartal) or InStr(zeile, "aktuell") ) {
			found:= A_Index
			break
		}
	}

	return found
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; FENSTERELEMENTE AUSLESEN UND BEHANDELN                                                                                                                                                                                    	(08)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisGetCaveZeile                	(02) AlbisGetCave                     	(03) AlbisMuster30ControlList()       	(04) AlbisOptPatientenfenster     	(05) AlbisHeilMittelKatologPosition
; (06) AlbisSortiereDiagnosen         	(07) AlbisReadFromListbox				(08) AlbisResizeDauerdiagnosen
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;01
AlbisGetCaveZeile(nr, SuchString:="", NoClose:= false) {              	;-- auslesen einer oder aller Zeilen im Cave! von Dialog

	;NoClose=1 Fenster wird anschließend nicht geschlossen
		Zeile := ""

	; Cave! von Fenster aufrufen
		Albismenu(32778, "Cave! von Ahk_class #32770", 4)

	; Auslesen aller Zeilen des Cave! von - Dialoges
		ControlGet, result, List, Col4, SysListView321, Cave! von ahk_class #32770

	; ein SuchString wurde übergeben
		If !(StrLen(SuchString) = 0)
			Loop, Parse, result, `n
				If InStr(A_LoopField, SuchString)
				{
						Zeile    	:= A_LoopField
						nrZeile	:= A_Index
						break
				}

	; falls der GVU-Text nicht in Zeile 9 steht dorthin kopieren
		If (nrZeile <> 9) && InStr(SuchString, "GVU") && (StrLen(Zeile) > 0)
		{
				AlbisSetCaveZeile(nrZeile, "", true)
				AlbisSetCaveZeile(9, Zeile, true)
		}

	; wenn keine Zeile gefunden wurde, welche den gesuchten String enthält und nach GVU gesucht wurde, wird die Zeile ausgelesen
	; Backup für Gesundheitsvorsorgeliste.ahk - hat bei nicht gefundenem GVU String nichts zurück gegeben, die Zeile war dann futsch!
		If (StrLen(Zeile) =0) && InStr(SuchString, "GVU") && (StrLen(nr) = 0 || nr = 0)
			Loop, Parse, result, `n
				If (A_Index = 9)
				{
    					Zeile    	:= A_LoopField
						nrZeile	:= A_Index
						break
				}

	; wenn NoClose = 0 ist, wird der Cave! von - Dialog wieder geschlossen
		localIndex:=0
		If !NoClose
		{
				while ( CaveVonID:= WinExist("Cave! von") )
				{
						WinActivate	 , % "ahk_id " CaveVonID
						WinWaitActive, % "ahk_id " CaveVonID,, 1
						VerifiedClick("Button4", "Cave! von")   	; ich wählte Button4 = Abbrechen ; Programmfehler können so keinen Schaden anrichten, alternativ geht auch mit zweimal Esc hintereinander
						Sleep, 50
						localIndex++
						If (localIndex>40)               					;es dauert auch mal länger als 4s bis das Fenster erscheint - 40x100 ms in diesem Fall
								MsgBox, 0x40000, Frage, Es wurde mehrmals probiert das Cave! von Fenster zu schließen`nBitte schließen Sie das Fenster von Hand.`nIm Anschluß läuft diese Funktion automatisch weiter.
				}
		}

return Trim(Zeile)
}
;02
AlbisGetCave(WinClose:=true) {       		                                    	;-- alle Zeilen des Cave!von Fenster auslesen

		AlbisActivate(1)

	; Cave! von Fenster aufrufen
		If !CaveVonID := Albismenu(32778, "Cave! von Ahk_class #32770", 4)
			return 0

	; Auslesen aller Zeilen des Cave! von - Dialoges
		ControlGet, result, List, Col4, SysListView321, % "ahk_id " CaveVonID

	; CaveVonFenster schließen
		If WinClose {
			while, CaveVonID:= WinExist("Cave! von")
			{
					WinActivate	 , % "ahk_id " CaveVonID
					WinWaitActive, % "ahk_id " CaveVonID,, 1
					VerifiedClick("Button4", "Cave! von")   	; ich wählte Button4 = Abbrechen ; Programmfehler können so keinen Schaden anrichten, alternativ geht auch mit zweimal Esc hintereinander
					Sleep, 50
					localIndex++
					If (localIndex>40)               					;es dauert auch mal länger als 4s bis das Fenster erscheint - 40x100 ms in diesem Fall
							MsgBox, 0x40000, Frage, Es wurde mehrmals probiert das Cave! von Fenster zu schließen`nBitte schließen Sie das Fenster von Hand.`nIm Anschluß läuft diese Funktion automatisch weiter.
			}
		}

return List
}
;03
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
;04
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
;05
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
;06
AlbisSortiereDiagnosen() {                                                         	;-- ## funktioniert nicht - nur zum Sortieren von Diagnosen im Dauerdiagnosenfenster

	DoSort:= 0

	If !hDia:= GetHex(WinExist("Dauerdiagnosen von ahk_class #32770"))
			return -1
	LVRow:= AlbisLVContent(hDia, "SysListView321", "1")
	ControlGet, hLV, HWND,, SysListView321, ahk_id %hDia%
	Loop, % LVRow.MaxIndex()
	{
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
;07
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
;--
LV_Select(r, Control, hWin) {                                                      	;-- select/deselect 1 to all rows of a listview (funktioniert nicht in fremder Listview)

	; Modified from http://www.autohotkey.com/board/topic/54752-listview-select-alldeselect-all/?p=343662
	; Examples: LVSel(1 , "SysListView321", "Win Title")   ; Select row 1. (or use +1)
	;           LVSel(-1, "SysListView321", "Win Title")   ; Deselect row 1
	;           LVSel(+0, "SysListView321", "Win Title")   ; Select all
	;           LVSel(-0, "SysListView321", "Win Title")   ; Deselect all
	;           LVSel(+0,                 , "ahk_id " HLV) ; Use listview's hwnd

	LVIS_FOCUSED:=1
	LVIS_SELECTED:=2
	LVM_SETITEMSTATE:=0x102B
	VarSetCapacity(LVITEM, 20, 0) ;to receive LVITEM
	NumPut(LVIS_FOCUSED | LVIS_SELECTED, LVITEM, 12)  ; state
	NumPut(LVIS_FOCUSED | LVIS_SELECTED, LVITEM, 16)  ; stateMask
	RemoteBuf_Open(hLVITEM, hWin, 20)  ; MASTER_ID = the ahk_id of the process owning the SysListView32 control
	RemoteBuf_Write(hLVITEM, LVITEM, 20)
	SendMessage, % LVM_SETITEMSTATE, % r, % RemoteBuf_Get(hLVITEM), % Control, % "ahk_id " hWin
	RemoteBuf_Close(hLVITEM)
}
;08
AlbisResizeDauerdiagnosen(Options:= "xCenterScreen yCenterScreen w0 h0") { ;-- spezielle Funktion für den Dauerdiagnosendialog

	static WT:= "Dauerdiagnosen von ahk_class #32770"
	static hStored

	hWin := WinExist(WT)
	If !hWin || (hStored = hWin)
		return 0

	RegExMatch(Options, "x(?<X>[A-Z0-9]+)\s+"	, Opt)
	RegExMatch(Options, "y(?<Y>[A-Z0-9]+)\s+"	, Opt)
	RegExMatch(Options, "w(?<W>[A-Z0-9]+)\s+"	, Opt)
	RegExMatch(Options, "h(?<H>[A-Z0-9]+)\s+"	, Opt)

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
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; DATUMSFUNKTIONEN                                                                                                                                                                                                                           	(04)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisLeseZeilenDatum           	(02) AlbisSetzeProgrammDatum 	(03) AlbisLeseProgrammDatum      	(04) AlbisSchliesseProgrammDatum
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
; (01)
AlbisLeseZeilenDatum(Sleeptime:=200) {                                                               	;-- liest das Zeilendatum der ausgewählten Zeile in der Akte aus

	; Achtung eine Zeile in der Akte muss nicht und sollte nicht ausgewählt sein, der Mauspfeil sollte aber über der Zeile stehen
	; dies funktioniert nur in der aktuellen Patientenakte (es darf noch keine Abrechnung stattgefunden haben)
	; letzte Änderung: 27.06.2020 - Erkennung des aktiven Fokus verbessert, Codesyntax modernisiert

	;falls ein Eingabefocus (blinkendes Caretzeichen) schon gesetzt ist, kann das Datum sofort ausgelesen werden
		If !InStr(AlbisGetActiveWindowType(), "Karteikarte")
				return 0

		hACtrl := GetFocusedControl()
		hActive := GetHex(GetParent(hACtrl))
		cf := GetClassNN(hACtrl, hActive)

		If InStr(cf, "Edit") || InStr(cf, "RichEdit")
				ControlGetText, Datum, Edit2, % "ahk_id " hActive

	;falls kein Focus besteht, dann jetzt einen Focus erzeugen
		while !RegExMatch(Datum, "\d\d\.\d\d\.\d\d\d\d") {

				SendInput, {Space}
				Sleep % sleeptime
				ControlGetText, Datum, Edit2, % "ahk_id " hActive ;ALBIS -
				If (A_Index > 3)
					return 0

		}

	;fragt den Focus ab, Esc beendet den Eingabemodus in der Akte, es sollte danach ein anderer Focus auslesbar sein
		while (cf = cf1) 		{

				SendInput, {Esc}
				Sleep % sleeptime
				ControlGetFocus, cf1, % "ahk_id " hActive
				If (A_Index > 10)
					return 0

		}

return Datum
}
; (02)
AlbisSetzeProgrammDatum(Datum:="") {                                                                     	;-- öffnet das Fenster Programmdatum einstellen und setzt das übermittelte Datum ein

		; Achtung: Datum muss absofort zwinged folgendes Format aufweisen: dd.mm.yyyy z.B. 24.12.2019
		; bei leerer Übergabe wird das Tagesdatum eingetragen

		If (StrLen(Datum) = 0)
			FormatTime, Datum,, dd.MM.yyyy

	; prüft auf Richtigkeit des vorliegenden Datumformates
		if !RegExMatch(Datum, "\d\d\.\d\d\.\d\d\d\d") 		{
					Throw exception("`"" Datum . "`" ist kein reguläres Datumsformat." )
					return
		}

	; Fenster Programmdatum einstellen aufrufen
		hWin:= Albismenu(32852, "Programmdatum einstellen", 3)
		If !hWin
		{
				MsgBox, 1, Addendum für Albis on Windows - %A_ScriptName%, Der Dialog 'Programmdatum' ließ sich nicht aufrufen! `nDie Funktion wird jetzt beendet, 5
				AlbisSchliesseProgrammDatum()
				return 0
		}

		ControlGetText, eingestelltesDatum, Edit1, Programmdatum einstellen ahk_class #32770

		VerifiedSetText("Edit1", Datum, "Programmdatum einstellen", 200)
		AlbisSchliesseProgrammDatum()

return eingestelltesDatum
}
; (03)
AlbisLeseProgrammDatum() {                                                                                	;-- öffnet das Fenster Programmdatum einstellen und setzt das übermittelte Datum ein

	;V1.2 mit Überprüfung ob das Fenster geöffnet, das Auslesen des Datum funktioniert und ob das Fenster auch wieder geschlossen wurde

	; Fenster Programmdatum einstellen aufrufen
		hWin:= Albismenu(32852, "Programmdatum einstellen", 3)
		If !hWin
		{
					MsgBox, 1, Addendum für Albis on Windows - %A_ScriptName%, Das Programmdatum ließ sich nicht einstellen! `nDie Funktion wird jetzt beendet, 5
					return 0
		}
					;PostMessage, 0x111, 32852, 0, , ALBIS ahk_class OptoAppClass;WinWaitActive, Programmdatum einstellen,, 3;If WinActive("Programmdatum einstellen");break;If A_Index > 5;		return;else	;		Sleep, 50


		while (Datum="") {
			ControlGetText, Datum, Edit1, Programmdatum einstellen ahk_class #32770
			Sleep, 100
			If A_Index > 10
					break
		}

		AlbisSchliesseProgrammDatum()
		AlbisActivate(2)

return Datum
}
; (04)
AlbisSchliesseProgrammDatum() {                                                                          	;-- schließt das Programmdatumsfenster

		while WinExist("Programmdatum einstellen")
		{
					If A_Index = 1
                            VerifiedClick("Button1", "Programmdatum einstellen")
					If A_Index = 2
                            SendInput, {ENTER}

					If WinExist("Albis ahk_class #32770", "Fehlerhaftes Datumsformat")
                            VerifiedClick("Button1", "Albis ahk_class #32770", "Fehlerhaftes Datumsformat")

					If A_Index > 2
                            return 0

					Sleep, 50
		}

return 1
}
; (05)
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; DATENEINGABE                                                                                                                                                                                                                                       	(09)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisPrepareInput                 	(02) AlbisSendInputLT                	(03) AlbisSetCaveZeile                   	(04) AlbisWriteProblemMed        	(05) AlbisKarteikarteAktivieren
; (06) AlbisKarteikartenFocusSetzen  	(07) AlbisSchreibeLkMitFaktor      	(08) AlbisSchreibeInKarteikarte       	(09) AlbisFehlendeLkoEintragen
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;01
AlbisPrepareInput(Name) {                                                                                    	;-- bereitet das Schreiben von Daten in die Akte vor


			;letzte Änderung 02.09.2019: nach dem Click in die Akte (#327702) , wird zunächst die Escape-Taste gesendet,
			; dies befreit den Cursor vom derzeitigen Eingabefeld, die Taste Ende konnte so manchmal nichts erreichen
			;                            				 Fehleranzeige bei Übergabe eines nicht erlaubten Parameter (Kontrollvariable ist benannt als 'allowed')

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

			while !AlbisExist()
			{
					MsgBox, 0x40001 , Addendum für Albis On Windows, WARNUNG!`nBitte warten Sie auf den Neustart von Albis!`nDrücken Sie dann erst auf 'OK'!`n'Abbrechen' beendet die aufgerufene Funktion.
					IfMsgBox, Cancel
                            return 2
			}

		; Albishauptfenster aktivieren und von eventuell blockierenden Fenstern befreien
			AlbisActivate(5)
			AlbisIsBlocked(AlbisWinID(), 2)

		; schaltet auf die Karteikarte des aktuellen Patienten um, falls noch nicht geschehen
			AlbisZeigeKarteikarte()

		; eine freie Zeile in der Akte schaffen
			If Instr(allowed, Name)
			{
						localIndex := 0
						Loop
						{
                            	; ein Klick auf den Eingabebereich der Patientenakte um dieses Fenster zu aktivieren, sonst können dort keine Daten eingetragen werden!
                            		SendInput, {Esc}
                            		sleep 300
                            		Mdihwnd:= AlbisGetActiveMDIChild()
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
			else
			{
						; Fehlerbehandlung bei Eingabe eines falschen Parameter
                            ErrorMessage:= "Error at script line: " . scriptline . "`n" . ScriptText . "`n`n-> Function call with not allowed parameter(s)! <-"
                            ExceptionHelper("include`\AddendumFunctions.ahk", "If Instr`(allowed`, Name`) `{", ErrorMessage, A_LineNumber)
                            return 90				;falsche Parameterbehandlung
			}


		; einzelne Abhandlungen für übergebene Kürzel
			If Instr(Name, "bild1") OR Instr(Name, "scan")
			{
						SendRaw, %Name%
						SendInput, {TAB}
						WinWait, Grafischer Befund ahk_class #32770,, 8
						If !WinExist("Grafischer Befund ahk_class #32770") {
                            	error:= ErrorLevel
                            	MsgBox, 0x1000, Fensterfehler, Das Fenster 'Grafischer Befund hat sich nicht geöffnet', 4
						}
			}

			If Instr(Name, "lko") OR Instr(Name, "lkü")
			{
						SendRaw, %Name%
						SendInput, {TAB}
			}

return error
}
;02
AlbisSendInputLT(kk, inhalt, kk_Ausnahme, kk_voll) {                                              	;-- schreibt zuerst ein Kürzel ein und dann den Inhalt für dieses z.B. 1.' lko ' 2. ' 03221 '

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
AlbisSetCaveZeile(nr, txt, NoClose:=false) {	                                                           	;-- überschreibt eine Zeile im cave! von - Fenster

	; letzte Änderung 05.07.2020:
	;
	; 		- SendInput z.T. ersetzt durch ControlSend damit die Tastaturbefehle an das richtige Fenster gesendet werden
	;    	- fügt selbstständig fehlende Listenzeilen hinzu

		CaveVon 	:= "Cave! von ahk_class #32770"
		slTime   	:= 50
		AlbisActivate(1)

	; Cave! von Fenster aufrufen
		If !(hwnd := WinExist(CaveVon))
			Albismenu(32778, CaveVon, 4)

	; Cave! von Fenster aktivieren
		WinActivate, % CaveVon
		WinWaitActive, % CaveVon,, 2

	; Focus ins SysListView geben
		ControlGetFocus, FControl, % CaveVon
		while !InStr(FControl, "SysListview") 	|| !InStr(FControl, "Edit")	{
			ControlFocus 	 , SysListView321, % CaveVon
			ControlGetFocus, FControl     	 , % CaveVon
			If InStr(FControl, "SysListview") || !InStr(FControl, "Edit") || (A_Index > 6)
				break
			sleep % slTime * 2
		}

	; Verzögerung bei bestehende Remotedesktopverbindung (die Darstellung der Fenster ist dann langsamer)
		ControlClick, SysListView321, % CaveVon,, Left, 1
		sleep % slTime

	; bei neue Patienten ohne Eintragungen im cave Dialog kann eine beliebige Zeile nicht direkt angesprungen werden
	; Der folgende Code prüft ob die gewünschte Zeile schon angelegt wurde. Wenn nicht wird diese angelegt.
		ControlGet, caveZeilenMax, List, Count, SysListview321, % CaveVon
		If (caveZeilenMax < nr) {

		; letzte Listenzeile auswählen
			ControlSend, SysListview321, % caveZeilenMax, % CaveVon
			sleep % slTime
		; Editfeld (Arzt) fokusieren
			ControlSend, SysListview321, {Space}, % CaveVon
			sleep % slTime
		; solange TAB senden bis die gewünschte Listenzeile erreicht ist
			Loop {

				If !WinActive(CaveVon)
					WinActivate, % CaveVon

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
				ControlSend, SysListview321, % caveZeilenMax, % CaveVon
				sleep % slTime
			; 1.Editfeld (Arzt) fokusieren
				ControlSend, SysListview321, {Space}, % CaveVon
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
			sleep % slTime
		}

	; kontrollieren ob der Text eingefügt wurde, wenn nicht wird der Nutzer zur manuellen Intervention aufgefordert
		If !VerifiedSetText("Edit1", txt, CaveVon, 500) || (caveZeilenMax > nr)	{
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
		If !NoClose {
			while WinExist(CaveVon) {
				VerifiedClick("Button3", CaveVon, "", "")
	     		Sleep % slTime * 2
				If (A_Index > 5)
					MsgBox, 1, % "Addendum für Albis on Windows - " A_ScriptName, Das "Cave! von" Fenster läßt sich nicht schließen.`nBitte schließen Sie es von Hand!
			}
		}

return ErrorLevel
}
;04
AlbisWriteProblemMed(medname) {                                                                       	;-- diese Funktion soll später eigenständig Eintragungen unter den Dauermedikamenten vornehmen

	;falls noch kein Problemmedikamentfeld vorhanden ist
	static1:= "#### Problemmedikamente / Allergien #####", lenstatic1:= StrLen(static1)					;41 Zeichen
	static2:= "#############################", lenstatic2:= StrLen(static2)



}
;05
AlbisKarteikarteAktivieren() {                                                                                  	;-- aktiviert durch einen simulierten Mausklick die Karteikarte

	; ohne Einsatz dieser Funktion läßt sich keine Dateneingabe in die Akte durchführen
	; gibt das hwnd des Karteikartenfensters zurück

	Mdihwnd:= AlbisGetActiveMDIChild()		                   	;neu seit 27.3.2019 - hoffentlich zuverlässig
	VerifiedClick("#327702","","", MdiHwnd)
	sleep, 200

return Mdihwnd
}
;06
AlbisKarteikartenFocusSetzen(Control) {                                                                   	;-- beste Funktion für gezielten Schreibzugriff auf die Karteikarte

		global Mdi:= Object()

	; prüft und zeigt die Karteikarte an, falls etwas anderes beim Patienten angezeigt wird
		AlbisActivate(2)
		AlbisZeigeKarteikarte()

	; ermittelt das richtige Control um dort die Eingaben durchzuführen
		hChild          	:= AlbisGetActiveMDIChild()
		hEingabeWin	:= Controls("#327702", "ID", hChild)

	; Tastaturfokus auf den Karteibereich und Tastenfolge senden
		ControlFocus,, % "ahk_id " hEingabeWin
		SendInput, {Escape}
		sleep, 300
		SendInput, {End}
		Sleep, 300
		SendInput, {Down}
		sleep, 300

	; Sicherheitsüberprüfung
		Loop {
			ControlGetFocus, fctrl, % "ahk_id " hEingabeWin
			If InStr(fctrl, Control) || (A_Index > 2)
				break
			else
				SendInput, {Down}
			sleep, 300
		}

	; gewünschten Teil fokussieren
		ControlFocus, % Control, % "ahk_id " hEingabeWin
		sleep, 100

		ControlGetFocus, fctrl, % "ahk_id " hEingabeWin

return ((fctrl = Control) ? GetHex(hEingabeWin) : "0")
}
;07
AlbisSchreibeLkMitFaktor(Leistungskomplex, StandardFaktor:="") {                          	;-- Vereinfachung für das Senden von Leistungskomplexen per Hotstring

	; z.B. AlbisSchreibeLkMitFaktor("40144", "2") - selektiert automatisch den Faktor damit dieser geändert werden kann
	; ist Standardfaktor leer oder null wird nur der Leistungskomplex gesendet
	sleep, 200

	If Standardfaktor
	{
		sendText:= Leistungskomplex "(x:" Standardfaktor ")"
		SendInput, {Text}%SendText%
		;SendRaw, (x:
		;SendInput, StandardFaktor
		;SendRaw, )
		sleep 200
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

		Funktion erstellt am 02.10.2019 - letzte Änderung 02.10.2019

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
		  - AlbisKarteikartenFocusSetzen()
		  - GetFocusControlClassNN()

	*/

		AlbisActivate(1)

		If (StrLen(ArztKennung) > 0) 	{
			hEingabeFocus := AlbisKarteikartenFocusSetzen("Edit1")
			VerifiedSetText("Edit1", Nutzer, hEingabefocus)
			SendInput, {Tab}
		}

		 If (StrLen(Datum) > 0)	{
			If !hEingabefocus
				hEingabeFocus:= AlbisKarteikartenFocusSetzen("Edit2")
			VerifiedSetText("Edit2", Datum, hEingabefocus)
			SendInput, {Tab}
		}

		If !hEingabefocus
			hEingabeFocus:= AlbisKarteikartenFocusSetzen("Edit3")

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

				SciTEOutput("PatIndex: " index)

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
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; FORMULARE                                                                                                                                                                                                                                           	(14)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisDruckeBlankoFormular  	(02) AlbisRezeptHelfer                	(03) AlbisRezeptHelferGui            		(04) AlbisRezeptFelderLoeschen  	(05) AlbisRezeptFelderAuslesen
; (06) AlbisRezeptSchalterLoeschen 	(07) AlbisDruckePatientenAusweis	(08) AlbisRezeptAutIdem                	(09) AlbisRezept_DauermedikamenteAuslesen
; (10) AlbisFormular                       	(11) AlbisHautkrebsScreening     	(12) AlbisFristenRechner                    	(13) AlbisFristenGui                   	(14) IfapVerordnung
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
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

	Loop, % items.MaxIndex()
	{
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
				while !WinExist(WinTitle)
				{
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
				If ErrorLevel
				{
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

	If (StrLen(tag) > 0)
	{
			VerifiedSetText("Edit2"	, Rezept[tag][1] , Muster16)
			VerifiedSetText("Edit8"	, Rezept[tag][2] , Muster16)
			VerifiedSetText("Edit14"	, Rezept[tag][3] , Muster16)
	}

return
}
;3
AlbisRezeptHelferGui(dbfile) {                                                                                  	;-- Schnellrezept - Standardrezepte für bestimmte Diagnosen/Hilfsmittel mit einem Klick erstellen

	; blendet ein kleines Listbox-Steuerelement in einem Kassenrezeptformular ein
	; diese Funktion braucht die Bibliothek >> lib\class_JSONFile.ahk <<
	; letzte Änderung: 09.06.2020

		global RHStatus    				                    	; flag damit bei Fokusänderungen die Gui nicht neu erstellt wird
		global hRezeptHelfer, RH
		global Mdi

		static DDLValues, RezeptWahl, wahl, RHelferBtn1
		static RezeptFeld	:= [2,8,14,20,26,32]	; Edit2, Edit8, ....
		static RezeptZusatz	:= [3,7,11,15,19,23]	; Button2, Button6, ....
		static Schalter    	:= {	39 : "BVG"
		                            	, 	40 : "Hilfsmittel"
										,	41 : "Impfstoff"
										,	42 : "Sprechstunden-Bedarf"
										,	43 : "Heilmittel"
										,	44 : "BTM"
										,	45 : "OTC"}
		static Rezepte                                        	; Objekt enthält den Datensatz für die Verordnungen
		static SchnellRezept                                	; speichert die Nummer des aktuell ausgewählten Schnellrezeptes
		static hMuster16                                    	; window-handle des Albis Rezeptfenster
		static hMuster16_stored                            	; window-handle des Albis Rezeptfenster
		static dbfileLastEdit                                	; enthält die letzte Dateiänderungszeit
		static RHRetry := 0

		RezeptIstZentriert := false

	; verhindern eines doppelten Aufrufes der Rezepthelfer Gui
		hMuster16:= WinExist("Muster 16 ahk_class #32770")
		If !WinExist("ahk_id " hMuster16) {
				RHRetry ++
				SciTEOutput("kein Rezeptfenster" RHRetry)
				RHStatus := false
				If WinExist("ahk_id " hRezeptHelfer) {
					Gui, RH: Destroy
					SetTimer, RezepthelferUpdate, Off
				}
				return
		}

		If WinExist("ahk_id " hRezeptHelfer)
			return

	; wartet kurz bis das Rezeptfenster vollständig gezeichnet wurde
		sleep, 200

	; für die Positionierung des Rezeptformulars
		RHStatus           := true
		APos             	:= GetWindowSpot(AlbisWinID())
		RPos                	:= GetWindowSpot(hMuster16)

	; im Moment nur so ermittelbare Position des Karteikarten Filter Steuerelementes (SysTabControl32[nr] - nr variiert)
	; gebraucht für die Positionierung des Rezeptfensters innerhalb von Albis
		hMdiChild     	:= AlbisGetActiveMDIChild()
		WinGet, CtrlList, ControlListHWND, % "ahk_id " hMdiChild
		Loop, Parse, CtrlList, `n
		{
				class := GetClassName(hwnd := A_LoopField)
				If InStr(class, "AfxFrameOrView140") {
					hAfxFOV := hwnd
					TPos    	:= GetWindowSpot(hAfxFOV)
					break
				}
		}

	; Position berechnen
		NewX             	:= Floor((APos.W - APos.X)/2) - Floor(RPos.W/2)
		NewY				:= TPos.Y, NewW := RPos.W, NewH	:= RPos.H

	; Werbefenster im Rezept ausschalten - Rezeptfenstergröße wird automatisch angepasst
		ControlGet, hWerbung	, HWND,, Shell Embedding1, % "ahk_id " hMuster16
		ControlGet, isVisible  	, Visible ,, Shell Embedding1, % "ahk_id " hMuster16
		If ((hWerbung > 0) && isVisible) {
				ControlGetPos,,,, WerbungH,, % "ahk_id " hWerbung
				Control, Hide,,, % "ahk_id " hWerbung
				NewH := RPos.H - WerbungH + 17
		}

		WinMoveZ(hMuster16, 0, NewX, NewY, NewW, NewH, 1)

	; Vergleichen der letzten Änderung der Schnellrezeptdatenbank, damit neue Rezeptevorschläge angezeigt werden können
		FileGetTime, fileTime, % dbfile, M
		If (fileTime <> dbfileLastEdit)
			Rezepte := ""

	; Schnellrezeptdaten einlesen (überprüft ob die Datei verändert wurde, liest die Daten dann neu ein)
		If !IsObject(Rezepte) && FileExist(dbfile) {

				Rezepte := Object()
				Rezepte	:= new JSONFile(dbfile)

		} else if !IsObject(Rezepte) && !FileExist(dbfile) {

				PraxTT("keine Schnellrezepte für Rezeptverordnungen vorhanden!", "0 2")
				Rezepte := Object()

		}

	; DDL Auswahl erstellen
		DDLValues := "Auswahl eines Schnellrezeptes...|neues Schnellrezept anlegen"
		For i, val in Rezepte.Object()
			DDLValues .= "|" val.Bezeichner

	; Schnell-Rezeptposition wird unterhalb des Steuerelementes Arzneimitteldatenbank angezeigt
		ControlGetPos, cx, cy, cw, ch, Button76, % "ahk_id " hMuster16

	; Gui erstellen
		Gui, RH: New, -Caption -DPIScale +ToolWindow +HWNDhRezeptHelfer +E0x00000004 -0x04020000 Parent%hMuster16%
		Gui, RH: Margin, 0, 0
		Gui, RH: Font, s8 q5 cBlack, MS Sans Serif
		Gui, RH: Add, DDL 	, % "x+5 w380 AltSubmit vRezeptwahl gRezeptausgewaehlt HWNDhRezeptWahl" , % DDLValues
		Gui, RH: Add, Button , % "x+5 gRezeptLeeren HWNDhRHelferBtn1"                                                    	, % "alle Felder leeren"
		Gui, RH: Show, % "x" cx + cw + 35 " y" cY + cH + 5 " NoActivate", Rezepthelfer
		GuiControl, RH: Choose, Rezeptwahl, 1

		RHStatus := true

	; Gui aktivieren, ganz nach oben setzen und neuzeichen - sonst funktioniert es nicht!
		WinActivate   	 , % "ahk_id " hRezeptHelfer
		WinSet, Top   	,, % "ahk_id " hRezeptHelfer
		WinSet, Redraw	,, % "ahk_id " hRezeptHelfer
		SetTimer, RezepthelferUpdate, 100

return

Rezeptausgewaehlt:       	;{

		Gui, RH: Submit, NoHide
		Gui, RH: Default

	; --------------------------------------------------------------------------------------------------------
	; Anlegen eines neuen Schnellrezeptes (Vorlage) zur Verwendung bei allen Patienten
	; --------------------------------------------------------------------------------------------------------
		If (RezeptWahl = 2)
		{
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

					} else {

						InputBox, Bezeichner, % "Addendum Rezepthelfer", % "Geben Sie dem neuen Schnellrezept einen eindeutigen Namen"

					}

				}

			; Comboboxauswahl auf den ersten Eintrag zurücksetzen
				PostMessage, 0x014E, 0, 0, ComboBox1, % "ahk_id " hRezeptHelfer

				MsgBox, 4, % "Addendum Rezepthelfer", % RezeptDaten.Medikamente.MaxIndex() " ausgefüllte Medikamentenfelder sind vorhanden.`nMöchten Sie dieses Rezept als Schnellrezept speichern?"
				IfMsgBox, No
					return

				RezeptDaten := ""
		}
	; --------------------------------------------------------------------------------------------------------
	; ausgewähltes Schnellrezept anzeigen
	; --------------------------------------------------------------------------------------------------------
		else If (RezeptWahl > 2)
		{

			RezeptNr := RezeptWahl - 2

		; Medikamentenfelder und Schalter werden gelöscht
			AlbisRezeptFelderLoeschen(hMuster16)

		; Medikamente eintragen
			For i, medikament in Rezepte[RezeptNr].medikamente
			{
					ControlSetText, % "Edit" RezeptFeld[i], % medikament, % "ahk_id " hMuster16
					If InStr(medikament, "...")
					{
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

	If !WinExist("Muster 16 ahk_class #32770")
	{
			SetTimer, RezepthelferUpdate, Off
			Gui, RH: Destroy
			If WinExist("Rezepthelfer ahk_class AutohotkeyGui")
				WinClose, % "Rezepthelfer ahk_class AutohotkeyGui"
			RHStatus:= false
			return
	}

return ;}
}
;3.1.
AlbisRezeptHelferGuiClose() {

	global RHStatus, RH, hRezeptHelfer

	SetTimer, RezepthelferUpdate, Off
	Gui, RH: Destroy
	IfWinExist, Addendum Rezepthelfer ahk_class AutohotkeyGui
		WinClose, Addendum Rezepthelfer ahk_class AutohotkeyGui
	RHStatus:= false

return
}
;4
AlbisRezeptFelderLoeschen(hMuster16) {                                                                	;-- (primär) Hilfsfunktion für die Rezepthelfer Gui

	static RezeptFeld	:= [2,8,14,20,26,32]	; Button
	static RezeptZusatz	:= [3,7,11,15,19,23]	; Button

	; löscht die Zusatztexte
		Loop, % RezeptZusatz.MaxIndex() 	{

				ControlGetText, cText, % "Button" RezeptZusatz[A_Index], % "ahk_id " hMuster16
				If InStr(cText, "zus")
				{
						VerifiedClick("Button" RezeptZusatz[A_Index], "ahk_id " hMuster16)
						WinWait, Medikamentenzusätze ahk_class #32770,, 5
						If !ErrorLevel
						{
								ControlSetText, % "Edit1", % "", Medikamentenzusätze ahk_class #32770
								ControlSetText, % "Edit2", % "", Medikamentenzusätze ahk_class #32770
								VerifiedClick("Button2", "Medikamentenzusätze ahk_class #32770")
								WinWaitClose, Medikamentenzusätze ahk_class #32770,, 5
								If ErrorLevel
								{
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

		Loop, % RezeptZusatz.MaxIndex()
		{
				ControlGetText, val, % "Edit" RezeptFeld[A_Index], % "ahk_id " hMuster16

				If StrLen(val) > 0
					RezeptDaten.Medikamente.Push(val)

				ControlGetText, cText, % "Button" RezeptZusatz[A_Index], % "ahk_id " hMuster16
				If InStr(cText, "zus")
				{
						VerifiedClick("Button" RezeptZusatz[A_Index], "ahk_id " hMuster16)
						ControlGetText, val, % "Edit1", % "", Medikamentenzusätze ahk_class #32770
						RezeptDaten.Zusatz.Push(val)
						VerifiedClick("Button2", "Medikamentenzusätze ahk_class #32770")
						WinWaitClose, Medikamentenzusätze ahk_class #32770,, 5
						If ErrorLevel
						{
								while WinExist("Medikamentenzusätze ahk_class #32770")
									MsgBox, Bitte das Fenster 'Medikamentenzusätze' schließen!
						}
				}
		}

	; bricht nach dem ersten gesetzten Schalter im Rezept ab
		For Nn, SchalterBezeichnung in Schalter
			If ControlGet("Checked", "", "Button" Nn, "ahk_id " hMuster16) {
				RezeptDaten.RezeptTyp:= SchalterBezeichnung
				break
			}
	;

return RezeptDaten
}
;6
AlbisRezeptSchalterLoeschen(hMuster16) {                                                                	;-- (primär) Hilfsfunktion für die Rezepthelfer Gui

	static Schalter    	:= {38:"BVG", 39:"Hilfsmittel", 40:"Impfstoff", 41:"Sprechstunden-Bedarf", 42:"Heilmittel", 43:"BTM", 44:"OTC"}

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
	If ErrorLevel
	{

	}

	;PostMessage, % BM_Click,,,, % "ahk_id " ControlGet("Hwnd", "Button23", "Patientenausweis ahk_class #32770")
	VerifiedClick("", ControlGet("Hwnd", "Button23", "Patientenausweis ahk_class #32770"))
	WinWaitClose, Patientenausweis,, 10
	while WinExist("Patientenausweis ahk_class #32770")
	{
			MsgBox,, Addendum für Albis on Windows, Bitte das Patientenausweis-Fenster schließen,`nbevor weiter gemacht werden kann!
			If A_Index > 10
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
				ControlGetText, Nacht       	, % "Edit"    Dosisfeld[A_Index], % "ahk_id " hMuster16
				ControlGetText, Medikament	, % "Edit"  Rezeptfeld[A_Index], % "ahk_id " hMuster16
				ControlGet, hAutIdem, hwnd,, % "Button" Idemfeld[A_Index], % "ahk_id " hMuster16
				SendMessage, 0xF2, 0, 0,, % "ahk_id " hAutIdem      ; BM_GETSTATE
				isChecked := ErrorLevel
				If RegExMatch(Nacht, "[A]") && !isChecked && StrLen(Medikament) > 0
				{
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

				For i, feld in Felder
				{
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
				If caveBereich
				{
						RegExMatch(A_LoopField, "(?<Medikament>[A-Za-zÄÖÜäöüß\-\,\s\(\))]+)\s\-\s(?<Text>.*)", cave)
						caveText	:= Trim(caveText)
						caveText	:= RTrim(caveText, "*")
						If StrLen(caveText) = 0
							continue

						medis      	:= StrSplit(caveMedikament, ",")
						caveExist 	:= false

						For index, cave in Med.Cave
							If InStr(cave.text, caveText)
							{
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

	If !formular.HasKey(Formularbezeichnung)
	{
		throw Exception("Formularbezeichnung: " Formularbezeichnung " ist unbekannt!")
		return 0
	}

return Albismenu(formular[Formularbezeichnung].command, formular[Formularbezeichnung].WinTitle)
}
;11
AlbisHautkrebsScreening(Pathologie:="opb", PlusGVU:=true, WinClose:=true ) {       	;-- befüllt das eHautkrebsScreening (nicht Dermatologen) Formular

	; das Hautkrebsscreeningformular für Nichtdermatologen wurde mehrfach geändert in den letzten Quartalen
	; die ClassNN der Buttons hat sich dabei jedesmal geändert, damit wurden nicht die richtigen Buttons angesprochen und somit war das Formular nicht korrekt ausgefüllt
	; oder ließ sich nicht speichern da sich auch für den Speichern Button die ClassNN geändert hatte
	; letzteres versuche ich nun mittels Übergabe des Namen/Text des Speichern-Buttons treffsicherer zu machen
	; prinzipiell müsste diese Funktion wesentlich flexibler gestaltet werden

		eHKS_Title := "Hautkrebsscreening - Nichtdermatologe ahk_class #32770"

		If !hHKS:= WinExist(eHKS_Title)
		{
				PraxTT("Das Formular 'Hautkrebsscreening - Nichtdermatologe' ist nicht geöffnet!`nDie Funktion wird abgebrochen.", "3 1")
				return 0
		}

	; markiert alle Verdachtsdiagnosen als nein
		If InStr(Pathologie, "opB")
		{
				VerifiedClick("Button28", eHKS_Title)	; nein - "Verdachtsdiagnose"  - warum Fehler wenn nicht ausgewählt
				VerifiedClick("Button10", eHKS_Title)	; nein - "Malignes Melanom"
				VerifiedClick("Button12", eHKS_Title)	; nein - "Basalzellkarzinom"
				VerifiedClick("Button14", eHKS_Title)	; nein - "Spinozelluläres Karzinom"
				VerifiedClick("Button24", eHKS_Title)	; nein - "Anderer Hautkrebs"
				VerifiedClick("Button26", eHKS_Title)	; nein - "Sonstiger dermatologisch abklärungsbedürftiger Befund"
				VerifiedClick("Button30", eHKS_Title)	; nein - "Screening-Teilnehmer wird an einen Dermatologen überwiesen:"
		}

	; Häkchen bei "gleichzeitige Gesundheitsvorsorge setzen"
		If PlusGVU
			VerifiedClick("Button16", eHKS_Title)

	; "Speichern" und damit auch Schließen des Formulares
		If WinClose
		{
				;VerifiedClick("Button19", eHKS_Title)
				ControlClick, Speichern, Hautkrebsscreening - Nichtdermatologe ahk_class #32770,,,, NA
				; falls sich das Formular nicht speichern ließ, wird keine 1 als Erfolgsmeldung zurück gesendet
				If WinExist(eHKS_Title)
					return 0
		}

return 1
}
;12
AlbisFristenRechner(AUSeit:="", AUBis:="") {                                                           	;-- errechnet das Datum des Beginns und des Ende der Krankengeldfortzahlung

		info =
		(
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

		bgStr1       	:= "ab d."
		endStr1	    	:= ""
		bgStr2      	:= "bis zum"
		endstr2     	:= ""

		Fristen			:= Object()

		If !AUSeit || !AUBis
		{
				If !WinExist("Muster 1a ahk_class #32770")
				{
					 return 0
				 }
				ControlGetText, AUSeit, Edit1, Muster 1a
				ControlGetText, AUBis, Edit2, Muster 1a
				If (AUSeitO = AUSeit) && (AUBisO = AUBis)
					 return
				AUSeitO 	:= AUSeit
				AUBisO 	:= AUBis
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

		If (TageBisBeginn >= 21) && !KGStartVorbei
		{
				TageBKg   	:= Mod(TageBisBeginn, 7)
				WochenBKg	:= Floor((TageBisBeginn - TageBkg)/7)
				ZeitBKg     	:= WochenBKg = 0 ? "(" : WochenBkg = 1 ? "(1 Woche": "(" WochenBKg " Wochen"
				ZeitBKg     	.= TageBKg = 0 ? ") " : TageBkg = 1 ? "  u. 1 Tag) ": " u. " TageBKg " Tage) "
		}
		else if (TageBisBeginn < 21) && !KGStartVorbei
				ZeitBKg      	:= "(" TageBisBeginn " Tage) "

		if KGStartVorbei
		{
				bgStr1      	:= "seit d."
				endstr1     	:= ""
		}

		If (TageBisAblauf >= 21) && !KgEndeVorbei
		{

				TageAKg   	:= Mod(TageBisAblauf, 7) + 0
				WochenAKg	:= Floor((TageBisAblauf - TageAKg)/7) + 0
				ZeitAKg     	:= WochenAKg = 0 ? "(" : WochenAkg = 1 ? "(1 Woche": "(" WochenAKg " Wochen"
				ZeitAKg     	.= TageAKg = 0 ? ") " : TageAkg = 1 ? " u. 1 Tag) ": " u. " TageAKg " Tage) "
		}
		else if (TageBisAblauf < 21) && !KgEndeVorbei
				ZeitAKg     	:= "(" TageBisAblauf " Tage)"

		If KgEndeVorbei
		{
				ZeitAKg     	:= ""
				ZeitBKg     	:= ""
				bgStr1      	:= ""
				endStr1     	:= ""
				bgStr2      	:= "ist am"
				endstr2     	:= "ausgelaufen"
		}

		Fristen.KgStart   	:= Trim(KrankengeldStart)
		Fristen.KgEnde  	:= Trim(KrankengeldEnde)
		Fristen.ZeitBKg  	:= ZeitBKg
		Fristen.ZeitAKg  	:= ZeitAKg
		Fristen.BKgStr1  	:= bgStr1
		Fristen.BKgStr2  	:= endStr1
		Fristen.AKgStr1  	:= bgStr2
		Fristen.AKgStr2  	:= endStr2
		Fristen.info        	:= info
		Fristen.Anzeige   	:= Fristen.BKgStr1 " " Fristen.KgStart " " Fristen.ZeitBKg Fristen.AKgStr1 " " Fristen.KgEnde " " Fristen.ZeitAKg " " Fristen.AKgStr2

return Fristen
}
;13
AlbisFristenGui() {                                                                                                  	;-- zeigt die Fristen auf dem 'Muster 1a' - Arbeitsunfähigkeitsbescheinigung an

	; diese Funktion wird durch einen WinEventHook gestartet, das Gui wird erst geschlossen wenn der 'Muster 1a' Dialog geschlossen wird
	; schließt bei Bedarf auch den Shift+F3 Kalender

	static Start, Ende, hStart, hEnde, hFrist, Termine, Frist, hOver, Info
	static AUSeitO, AUBisO, fU:= 0
	static hInfoIcon 	:= ImageFromBase64(true, "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAA5QAAAOUBj+WbPAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAFcSURBVDiNlZPPSgJRFMa/O84gOIt8h0pS0PzzDA75ALqwJ7EnyEX2DC2qhZsINChoZYsEE0ZCBm1hgua4choXcu94W0ST03ihOavvcr7zu9+Fe8h910zrnU5d7/V2KaUE/yhFUXgqlRpmM5kiOb94MKrV0xhfc4+pcFTATjQKAPi0LDSaDU+fSASVyokhG31j/+8wIQTFUgnhcBgAsFqt0LxrgvNfH19z9F/7MYk5/ticc0w/pu55Mpl4hn+KOZRIoneas5mr5+ZcZIMQYFmWqxfWIjjAtpeuXtp2cIDDmKupw0Q2MYBuABilwQFsA+AwRwiQRY3WUwvj8TsAYDQaBU+gRiLQ8hq0vAZVVYMnKJePkcvlvg+EoFY7255ADin+LwbAnJtbtef2kMLleOJgcHNLfMt0fXmFt+EQANB+bvuGiUQQT8QN8vgyO+x0u/Weru8FWudkcpDOZItfoQSYA7zlEWMAAAAASUVORK5CYII=")
	;static InfoIcon  :=	"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAA5QAAAOUBj+WbPAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAFcSURBVDiNlZPPSgJRFMa/O84gOIt8h0pS0PzzDA75ALqwJ7EnyEX2DC2qhZsINChoZYsEE0ZCBm1hgua4choXcu94W0ST03ihOavvcr7zu9+Fe8h910zrnU5d7/V2KaUE/yhFUXgqlRpmM5kiOb94MKrV0xhfc4+pcFTATjQKAPi0LDSaDU+fSASVyokhG31j/+8wIQTFUgnhcBgAsFqt0LxrgvNfH19z9F/7MYk5/ticc0w/pu55Mpl4hn+KOZRIoneas5mr5+ZcZIMQYFmWqxfWIjjAtpeuXtp2cIDDmKupw0Q2MYBuABilwQFsA+AwRwiQRY3WUwvj8TsAYDQaBU+gRiLQ8hq0vAZVVYMnKJePkcvlvg+EoFY7255ADin+LwbAnJtbtef2kMLleOJgcHNLfMt0fXmFt+EQANB+bvuGiUQQT8QN8vgyO+x0u/Weru8FWudkcpDOZItfoQSYA7zlEWMAAAAASUVORK5CYII="
	static cTT       	:= Object()
	static Fristen  	:= Object()

	global hMyCal, newCal

	;hInfoIcon      	:= ImageFromBase64(true, InfoIcon)

	ControlGetPos, tX, tY,, tH, Button8, Muster 1a ahk_class #32770
	hAU      	:= WinExist("Muster 1a ahk_class #32770")
	If (GetDec(hAU) = 0)
		return
	AU        	:= GetWindowSpot(hAU)
	Fristen   	:= AlbisFristenRechner()
	FristX    	:= AU.BW
	FristY    	:= AU.BH
	FristW    	:= AU.CW - AU.BW*2
	FristH    	:= 20

	Gui, Frist: New, -Caption -DPIScale -AlwaysOnTop +HWNDhFrist Parent%hAU% -0x04020000 +E0x00000004  ;0x00000004 = NoParentNotify
	Gui, Frist: Margin, 0, 0
	Gui, Frist: Color, % "c172842"   ;"c128576"
	Gui, Frist: Font, s8 q5 cWhite, Futura Md Bt
	Gui, Frist: Add, Text    	, % "x4 y3 BackgroundTrans vTermine                    	"            	 , % "Krankengeldzahlung "
	GuiControlGet, t, Frist: Pos, Termine
	Gui, Frist: Font, s8 q5 Normal cWhite, Futura Bk Bt
	Gui, Frist: Add, Text		, % "x+0 w" FristW - tW - tX - 40 " BackgroundTrans vStart		+HWNDhStart	", % Fristen.Anzeige
	Gui, Frist: Add, Picture	, % "x" FristW - 18 " y2 +0x4000000  vInfo gAlbisFristenInfo"                      	 	 , % "hBitmap: " hInfoIcon	;w16 h16
	Gui, Frist: Show, % "x" FristX " y" FristY " w" FristW " h" FristH " NoActivate", AdmFristen
	WinSet, Redraw,, % "ahk_id " hFrist
	SetTimer, AlbisFristenUpdate, 100

return

AlbisFristenUpdate: ;{

	If !WinExist("Muster 1a ahk_class #32770")
	{
			Gui, Frist: Destroy
			If WinExist("ahk_id " hMyCal) || WinExist("erweiterter Kalender ahk_class AutohotkeyGui")
				Gui, newCal: Destroy
			SetTimer, AlbisFristenUpdate, Off
			return
	}

	ControlGetText, AUSeit, Edit1, Muster 1a ahk_class #32770
	ControlGetText, AUBis, Edit2, Muster 1a ahk_class #32770
	If (AUSeitO = AUSeit) && (AUBisO = AUBis)
    		 return
	AUSeitO 	:= AUSeit
	AUBisO 	:= AUBis
	Fristen   	:= AlbisFristenRechner(AUSeit, AUBis)
	ControlSetText,, % Fristen.Anzeige, % "ahk_id " hStart
	WinSet, Redraw,, % "ahk_id " hFrist

return ;}

AlbisFristenInfo: ;{

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
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; WARTEZIMMER                                                                                                                                                                                                                                        	(03)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisWZPatientEntfernen       	(02) AlbisOeffneWarteZimmer    	(03) AlbisWartezeit
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;1
AlbisWZPatientEntfernen(Nachname, Vorname) { 				                                    	;--##### benötigt einen REWRITE ###### entfernt einen Patienten aus dem Wartezimmer

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
AlbisOeffneWarteZimmer() {                            					                                    	;-- öffnet das Wartezimmer

	; https://docs.microsoft.com/en-us/windows/desktop/winmsg/wm-mdimaximize
	static WM_MDIMAXIMIZE:= 0x0225, WM_MDIGETACTIVE:= 0x0229				;wParam: A handle to the MDI child window to be maximized.
	localIndex:=0

	If !(hWZ:= AlbisGetWartezimmerID())
	{
				PostMessage, 0x111, 32924,,, % "ahk_id " AlbisWinID()                                                      	;ruft das Wartezimmer auf
				sleep, 200
				while !(hWZ:= AlbisGetWartezimmerID())
				{
						localIndex++
						If (localIndex > 11)
							return 0 ; kein Erfolg!
						sleep, 200
						PostMessage, 0x111, 32924,,, % "ahk_id " AlbisWinID()
				}
	}

	; Wartezimmer wird maximiert
	PostMessage, % WM_MDIMAXIMIZE, %hWZ%,,, % "ahk_id " AlbisGetMDIClientHandle()

return 1		;oder war schon geöffnet
}

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; KARTEIKARTE/AKTE oder ALBIS-MENU                                                                                                                                                                                                        	(13)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAkteSchliessen               	(02) AlbisAkteOeffnen                	(03) AlbisAkteGeoeffnet                 	(04) AlbisPatientAuswaehlen       	(05) AlbisInBehandlungSetzen
; (06) AlbisSucheInAkte                  	(07) AlbisMenu                          	(08) AlbisZeigeKarteikarte              	(09) AlbisOeffnePatient              	(10) AlbisDateiAnzeigen
; (11) AlbisDateiSpeichern              	(12) AlbisOeffneAkte                    	(13) AlbisKarteikarteAktiv
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;01
AlbisAkteSchliessen(CaseTitle="") {                                                             	;-- schließt eine Karteikarte ## seit Albis20.20 mit Problemen

	;Rückgabewerte 0 - die Funktion konnte keine Karteikarte schliessen, 1 - war nicht erfolgreich, 2 - es war keine Akte zum Schliessen geöffnet

	oACtrl			:= Object()			;active control object
	AlbisWinID	:= AlbisWinID()

	If InStr(AlbisGetActiveWindowType(), "Karteikarte")
	{
			; aktuellen FensterTitel auslesen wenn keine Daten der Funktion übergeben wurden
				If (StrLen(CaseTitle) = 0)
					CaseTitle:= AlbisGetActiveWinTitle()

			; das Albishauptfenster aktivieren
				AlbisActivate(1)

			; ein Eingabefocus in der Patientenakte muss erkannt und beendet werden (senden einer Escape- oder Tab-Tasteneingabe)
				while InStr(AlbisGetActiveWinTitle(), CaseTitle)
				{
                            oACtrl:= AlbisGetActiveControl("content")
                            If ( (oACtrl.Identifier = "Dokument") && (StrLen(oACtrl.RichEdit) = 0) )	     	{					;leeres RichEdit - dann steht das Caret im Edit-Control (das ginge doch auch anders?!)
                            		AlbisKarteikarteAktivieren()
                            		SendInput, {Esc}
									sleep 100
                            }
                            else If ( (oACtrl.Identifier = "Dokument") && !(StrLen(oACtrl.RichEdit) = 0) ) 	{
                            		AlbisKarteikarteAktivieren()
                            		SendInput, {Tab}
									sleep 100
                            }

						; schliesst ausschließlich die Akte in deren Titel sich der Stringparameter -CaseTitle- findet
                            err := AlbisCloseMDITab(CaseTitle)
				}

			return err
	}

return 0
}
;02
AlbisAkteOeffnen(CaseTitle, PatID="") {                                                      	;-- öffnet eine Patientenakte über Name, ID oder Geburtsdatum

	/*				DESCRIPTION

			AlbisAkteOeffnen - zum Öffnen einer Patientenakte nach Übergabe eines Namens, der Patienten ID oder eines Geburtsdatums, letzte Änderung 06.06.2019

			Die Funktion ist in der Lage sämtliche Fenster abzufangen und entsprechend zu behandeln. Es wird bei erfolgreichem Aufruf der gewünschten Akte eine 1 ansonsten eine 0
			zurückgegeben. Falls sich ein Fenster mit einem Listview öffnet - Auswahl eines Namens, wenn mehr als einer vorhanden ist , wird eine 2 zurück gegeben. Nur das Listview
			Fenster wird nicht geschlossen. Da dieses Ausgelesen werden muss um in der Auswahl nach dem entsprechenden Namen suchen zu können.
	*/

	; Variablen ;{
		global Mdi
		static WarteFunc, AlbisTitleFirst
		name    	:= []
		CaseTitle	:= Trim(CaseTitle)
	;}

	; ist die Akte bereits geöffnet? ;{
		oMDI := AlbisGetAllMDIWin()
		hMdi	 := AlbisGetMDIClientHandle()
		sStr	 := StrLen(PatID) > 0 ? PatID "\s\/" : CaseTitle
		For MDITitle, thisMdi in oMDI
		{
				If RegExMatch(MDITitle, sStr) {
						SendMessage, 0x222, % thisMdi.id,,, % "ahk_id " hMdi
						PraxTT("Die gewünschte Karteikarte wird angezeigt.", "5 0")
						return 1
				}
		}

	;}

		If PatID
			PraxTT("Geöffnet wird die Karteikarte des Patienten:`n#2" oPat[PatID].Nn ", " oPat[PatID].Vn ", geb.am " oPat[PatID].Gd "(" PatID ")`n`n(Warte bis zu 10s auf die Karteikarte)", "10 0")

		AlbisTitleFirst:= AlbisHookTitle       	; aktuellen Albisfenstertitel auslesen
		AlbisDialogOeffnePatient()             	; Aufruf des Fenster 'Patient öffnen'

	; Übergeben des Parameter an das Albisdialogfenster ;{
		If (StrLen(PatID) > 0) {

			if VerifiedSetText("Edit1", PatID, "Patient öffnen ahk_class #32770", 200) {
				VerifiedClick("Button2", "Patient öffnen ahk_class #32770")
				goto AlbisWaitAkte
			} else {
				VerifiedClick("Button3", "Patient öffnen ahk_class #32770")
				return 0
			}

		}
		else {

			if !VerifiedSetText("Edit1", CaseTitle, "Patient öffnen ahk_class #32770", 200)
			return 0

		}

	;}

	; Aufrufen der Akte durch Schließen des -Patient öffnen- Dialoges ;{
		while WinExist("Patient öffnen ahk_class #32770")
		{
				; Button OK drücken
					VerifiedClick("Button2", "Patient öffnen ahk_class #32770")
					WinWaitClose, Patient öffnen ahk_class #32770,, 1
				; Fenster ist immer noch da? Dann sende ein ENTER.
					if WinExist("Patient öffnen ahk_class #32770")
					{
                            WinActivate, Patient öffnen ahk_class #32770
                            ControlFocus, Edit1, Patient öffnen ahk_class #32770
                            SendInput, {Enter}
					}

				sleep, 500
		}
	;}

	; Loop der in den nächsten 10Sekunden auf die neue Karteikarte wartet ;{
		AlbisWaitAkte:
		Loop 		{

				hwnd		:= WinExist("A")
				newTitle	:= WinGetTitle(hwnd)
				newClass	:= WinGetClass(hwnd)
				newText	:= WinGetText(hwnd)

			; Dialog Patient <.....,......> nicht vorhanden
				If Instr(newTitle, "ALBIS") && Instr(newClass, "#32770") && Instr(NewText, "nicht vorhanden") {
						VerifiedClick("Button1", "ALBIS ahk_class #32770", "nicht vorhanden")			;Abbrechen
						PraxTT("Albis konnte keine Karteikarte finden!", "5 3")
						if WinExist("Patient öffnen ahk_class #32770")
							VerifiedClick("Button2", "Patient öffnen ahk_class #32770")
						return 0
				}
				else if Instr(newTitle, "Patient") && Instr(newClass, "#32770") && Instr(newText, "List1") {
						return 2
				}

				If InStr(newTitle, CaseTitle)
					break

				If (A_Index > 20) 	{

						PraxTT("Die Karteikarte des Patienten:`n#2" oPat[PatID].Nn ", " oPat[PatID].Vn ", geb.am " oPat[PatID].Gd "(" PatID ")`nkonnte nicht geöffnet werden", "6 3")
						if WinExist("Patient öffnen ahk_class #32770")
							VerifiedClick("Button2", "Patient öffnen ahk_class #32770")
						return 0

				}

				sleep, 500
		}
	;}

		PraxTT("", "Off")

return 1
}
;----
AlbisWarteAufKarteikarte(AlbisTitleFirst, CaseTitle, waitingtime) {

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
		If !RegExMatch(CaseTitle, "\d+", dummy)
		{
				name:= ExtractNamesFromString(CaseTitle)
				If !InStr(AlbisGetActiveWinTitle(), name[3]) && !InStr(AlbisGetActiveWinTitle(), name[4])
				{
						MsgBox,, Addendum für AlbisOnWindows - %A_ScriptName%, % "Achtung:`nDie angeforderte Patientenakte ließ sich nicht öffnen!", 25
						return 0				;fungiert als ErrorLevel
				}
		}
		else
		{
				If !InStr(AlbisGetActiveWinTitle(), CaseTitle)
				{
						MsgBox,, Addendum für AlbisOnWindows - %A_ScriptName%, % "Achtung:`nDie angeforderte Patientenakte ließ sich nicht öffnen!", 25
						return 0				;fungiert als ErrorLevel
				}
		}

return 			;erfolgreich dann 1
}
;03
AlbisAkteGeoeffnet(Nachname, Vorname, GebDatum:="", PatID:="") {       	;-- kontrolliert ob die Akte mit den übergebenen Namen geöffnet ist, alle Bedingungen müssen erfüllt sein

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
AlbisPatientAuswaehlen(Nachname, Vorname, GebDatum) {                       	;-- ~# fehlerhaft ## wenn sich die Mehrfachauswahl eines Patienten bei häufigen Nachnamen z.B. ergibt

		Patienten:= Object()

		win = Patient auswählen ahk_class #32770
		hwnd:= WinExist("Patient auswählen ahk_class #32770", "List1")

		Patienten:= AlbisLVContent(hwnd, "SysListView321", "Name|Vorname|Geb.-Datum")


return
}
;05
AlbisInBehandlungSetzen() {                                                                      	;-- Praxomat Funktion - setzt den Pat. im Wartezimmer in Behandlung und danach wird der Timer gestartet

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

	while WinExist("ahk_id " . hWZ)
	{
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


	AlbisZeigeKarteikarte()

	;ok, die Karteikarte ist geöffnet, dann aktiviere ich sie jetzt und sie bekommt auch den Focus, danach kann sofort die Suchfunktion gestartet werden
	Loop {

		while !WinExist("Suchen","Suchen &nach`:") {
			WinActivate, ahk_class OptoAppClass
				WinWaitActive, ahk_class OptoAppClass,,4
					VerifiedClick("#327704","ALBIS ahk_class OptoAppClass") 	;ein Klick in die Karteikarte damit entweder an den Anfang oder an das Ende gesprungen werden kann
                            	SendInput, {Up}
						If (Richtung="up") {
                            	SendInput, {End}
						} else if (Richtung="down") {
                            	SendInput, {Home}
						}

                            sleep, 100
						SendInput, {Ctrl down}s{Ctrl up}
                            sleep, 100
                            		}

		;jetzt werden die Parameter gesetzt
		If (Richtung="up") {
				VerifiedCheck("Button4", "Suchen")
		} else if (Richtung="down") {
				VerifiedCheck("Button5", "Suchen")
		}

				sleep, 100

		ControlSetText, Edit1, %inhalt%, "Suchen ahk_class #32770"				;jetzt noch den Suchtext einsetzen
				sleep, 100
		VerifiedClick("Button9", "Suchen ahk_class #32770")					;und jetzt geht es los mit dem Suche
				sleep, 100
		AlbisZeigeKarteikarte()
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

		localIndex:= A_Index

	}	until localIndex = Vorkommen

return localIndex
}
;07
Albismenu(mcmd, FTitel:="", WZeit:=2, methode:=1) {                              	;-- Aufrufen eines Menupunktes oder Toolbarkommandos, wartet je nach Parameter auf ein sich öffnendes Fenster

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

		InfoMsg:= false

		If IsObject(FTitel) {
			WinTitle	:= FTitel.1
			WinText		:= FTitel.2
			AltTitle  	:= FTitel.3
			AltText   	:= FTitel.4
		}
		else
			WinTitle	:= FTitel

	; aktiviert das Albishauptfenster
		AlbisActivate(2)

	; schließt blockierende Fenster
		AlbisIsBlocked(AlbisWinID())
		;sleep, 500

	; wenn ein zu erwartender Fenstertitel übergeben wurde, wird der Titel zur Kontrolle des Menuaufrufes genutzt. Ansonsten bleibt nur SendMessage - siehe unten
		If (StrLen(WinTitle) > 0)
		{
				Loop 	{

						; Menu Aufruf durchführen
							If !WinExist(WinTitle) {
								; Methodenwahl
									If (methode = 1)
										PostMessage, 0x111, % mcmd,,, ahk_class OptoAppClass
									else if (methode = 2)
										SendMessage, 0x111, % mcmd,,, ahk_class OptoAppClass
							}

						; Fenster abwarten
                            Loop {

									If WinExist(WinTitle, WinText)	{
										WinActivate   	, % WinTitle, % WinText
										WinWaitActive	, % WinTitle, % WinText, % WZeit
										return GetHex(WinExist(WinTitle, WinText))
									}
									else if IsObject(FTitel) && WinExist(AltTitle, AltText)
										return {"hwnd": GetHex(WinExist(AltTitle, AltText))}

									sleep, 100
									If (A_Index *100 >= WZeit * 1000)
											break

							}

							If (A_Index > 2) ; Abbruch nach dem 3.Versuch
                      			return 0
					}

		}
		else
		{
				If (methode = 1) {
						PostMessage, 0x111, % mcmd,,, ahk_class OptoAppClass
						return 1
				}
				else if (methode = 2) { ; ACHTUNG: Sendmessage wartet auf eine Antwort, diese kann lange dauern (bis 5s zum Teil)
						SendMessage, 0x111, % mcmd,,, ahk_class OptoAppClass
						return ErrorLevel
				}
		}


return 0
}
;---
AlbisInvoke(menu, start=1) {                                                                     	;-- ~# fehlerhalt ## Vereinfachung eines Menupunktaufrufes - Teil- oder Ganzstring Übergabe wie im Albis Menu selbst

	;--diese Funktion benötigt die Datei AlbisMenu.json aus dem AddendumDir Verzeichnis \include, die Variable AddendumDir muss Superglobal im aufrufenden Skript definiert sein
	;--übergebe einen Menu-Namen z.b. AlbisInvoke("Wartezimmer",1) erhälst du die CommandID zurück und der Menupunkt wird bei start=1 durch PostMessage aufgerufen, start=0 gibt nur die CommandID zurück
	;--folgende Syntaxe sind möglich: 	1. vollständiger Pfad "Formular/BG/F1050  - Ärztliche Unfallmeldung	(A13)" , hier gehen auch Teilstrings "Formular/BG/F1050" oder "Form/BG/A13"
	;                            						2.	Teile der Suche "Formular, F1050" oder "BG, A13" - dies dauert länger da die Funktionen mitunter jedes Key:Value Paar durchsuchen muss
	static oAlbisMenu:=Object(Json2Obj(FileOpen(Addendum.AddendumDir "\include\AlbisMenu.json", "r").Read()))

	For key, val in oAlbisMenu
	{
		If ( Instr(key, menu) and (Start=1) ) {
				PostMessage, 0x111, % val,,, % "ahk_id " AlbisWinID()
				return
		} else {
				return val
		}
	}

}
;08
AlbisZeigeKarteikarte() {                                                                             	;-- schaltet zur Karteikartenansicht

	; !!WICHTIG!!: ein Erfolg der funktioniert wird per Rückgabeparameter 1 bestätigt, die aufrufende Funktion sollte den Rückgabewert prüfen bevor andere Funktionen fortgesetzt werden!

	; es wird keine Patientenakte angezeigt - dann ist die Funktionsausführung nicht möglich
		If !InStr(AlbisGetActiveWindowType(), "Patientenakte")
				return 0

	; wechselt zur Karteikarte, muss dazu begonnene Eingaben beenden oder abbrechen.
	; es werden hier zwei Dialogfenster abgefangen und entsprechend
		while !InStr(AlbisGetActiveWindowType(), "Karteikarte")
		{

				PostMessage, 0x111, 33033,,, % "ahk_id " AlbisWinID()
				sleep, 200

				Loop {

					hAlbisHinweis1 := WinExist("ALBIS ahk_class #32770", "Die aktuelle Zeile wurde nicht gespeichert")
					hAlbisHinweis2 := WinExist("ALBIS ahk_class #32770", "Gebühren")

					If hAlbisHinweis2
						VerifiedClick("Button1", "ALBIS ahk_class #32770", "Gebühren")

					while, WinExist("ahk_id " hAlbisHinweis1)
					{
								WinSetTitle, % "ahk_id " hAlbisHinweis1,, % "ALBIS ...wird automatisch geschlossen in " Round(5-(0.1*(A_Index-1)),1) "s"
								Sleep, 100
								If (A_Index > 50)
								{
										VerifiedClick("Button1", "ALBIS ahk_class #32770", "Die aktuelle Zeile wurde nicht gespeichert")
										AlbisActivate(1)
										cFocus := Controls("", "GetFocus", hActiveMDIWin:= AlbisGetActiveMDIChild())
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
					cFocus := Controls("", "GetFocus", hActiveMDIWin:= AlbisGetActiveMDIChild())
					If (StrLen(cFocus) > 0)
						SendInput, {Esc}

					Sleep, 100
				}

				If InStr(AlbisGetActiveWindowType(), "Karteikarte")
					break
				else if (A_Index > 10)
					return 0

		}

return 1
}
;09
AlbisDialogOeffnePatient(command:="invoke", pattern:="" ) {                       	;-- startet den Dialog zum Öffnen einer Patientenakte per Tastatureingabe

	; more commands are here:
	; 	abort/close - um das Fenster zu schliessen ohne eine Suche zu durchzuführen
	;	serach/set/open [Namens-/Suchmuster]- übernimmt den eingegeben Text und kann gleichzeitig das Suchmuster eintragen

	; Aufrufen des Dialogfensters
		If !InStr(command, "close")
			id:= Albismenu(32768, "Patient öffnen ahk_class #32770", 3)

	; command parsen
		If InStr(command, "invoke")
			return ID
		else If InStr(command, "close") {

				If WinExist("Patient öffnen ahk_class #32770")
				return VerifiedClick("Button3", "Patient öffnen ahk_class #32770")
			else
				return

		} else If InStr(command, "open") || InStr(command, "set")  {

			; kein Suchmuster übergeben dann wird nur auf Ok gedrückt
				If (StrLen(Pattern) > 0)
					If !VerifiedSetText("Edit1", Pattern, "Patient öffnen ahk_class #32770", 200)
						return 0

			; Suchmuster wurde als Parameter übergeben, aber
				If InStr(command, "set")
					return ID

				; Akte wird jetzt geöffnet durch drücken von OK
					while WinExist("Patient öffnen ahk_class #32770")
					{
							; Button OK drücken
								VerifiedClick("Button2", "Patient öffnen ahk_class #32770")
								WinWaitClose, Patient öffnen ahk_class #32770,, 1
							; Fenster ist immer noch da? Dann sende ein ENTER.
								if WinExist("Patient öffnen ahk_class #32770")
								{
										WinActivate, Patient öffnen ahk_class #32770
										ControlFocus, Edit1, Patient öffnen ahk_class #32770
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
AlbisDateiAnzeigen(FullFilePath) {                                                              	;-- öffnet eine Datei zur Ansicht in Albis, überprüft auf das korrekte Dateiformat

		AlbisActivate(1)
		SplitPath, FullFilePath, filepath, filename

	; Dateiformat prüfen (die erste Zeile einer in Albis anzeigbaren Textdatei beginnt mit "\P")
		FileReadLine, fileline, % FullFilePath, 1
		If !InStr(fileline, "\P")
		{
				MsgBox, 0, Addendum für Albis on Windows, % "Die übergebene Datei kann mit Albis nicht angezeigt werden."
				return 1
		}

	; Aufruf des Menupunktes Patient/Datei anzeigen...
		If !hWin := Albismenu("33030", "Öffnen ahk_class #32770")
		{
				MsgBox, 0, Addendum für Albis on Windows, % "Ein Fehler ist bei der Automatisierung aufgetreten!`n"
				. "`nDer Aufruf des Dialoges ''Datei anzeigen'' war nicht erfolgreich.`nDie Funktion wird jetzt beendet!"
				return 1
		}

	; schreibt zuerst den Dateipfad in das Fenster, sendet dann Enter und schreibt den Dateinamen in selbiges Feld
		VerifiedSetFocus("Edit1"	, hWin)
		  VerifiedSetText("Edit1"	, FullFilePath, hWin, 200)
		VerifiedSetFocus("Edit1"	, hWin)
		ControlSend, Edit1, {Enter}, % "ahk_id " hWin
		If ErrorLevel
		{
				MsgBox, 0, Addendum für Albis on Windows, % "Ein Fehler ist bei der Automatisierung aufgetreten!`nDie aufgerufene Funktion wird jetzt beendet."
				return 1
		}

return 0
}
;11
AlbisDateiSpeichern(FullFilePath, overwrite:= false) {                                     	;-- speichert von Albis erstellte Auswertungen, Protokolle, Statistiken oder Listen

	; prüft ob der Dateiname schon existiert, setzt dem Dateinamen eine Indizierung hinzu und speichert so unter einem anderen Namen
		If !overwrite && FileExist(FullFilePath)
		{
				SplitPath, FullFilePath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
				Loop
				{
						fname:= OutDir "\" OutNameNoExt "(" A_Index ")." OutExtension
						If !FileExist(fname)
							break
				}

				FullFilePath:= fname
		}

	; Tastatur- und Mauseingriffe durch den Nutzer kurzzeitig sperren
		hk(1, 0, "!ACHTUNG!`n`nDie Tastatureingaben sind für maximal 10 Sekunden gesperrt", 10)

	; Tagesprotokoll speichern und die Tagesprotokollausgabe schließen
		If !hSaveAsWin := Albismenu(33014, "Speichern unter ahk_class #32770")
		{
				MsgBox, 0, Addendum für Albis on Windows, Erwarteter 'Speichern unter' - Dialog konnte nicht abgefangen werden!`nDie Funktion wird abgebrochen
				return 0
		}

	; Steuerelemente befüllen
		Controls("Edit1", "SetText, " FullFilePath, hSaveAsWin)

	; Speichern unter drücken
		Controls("Button2", "Click use Controlclick", hSaveAsWin)
		sleep 200

	; Warten auf einen eventuellen "Speichern unter bestätigen" - Dialog
		If overwrite
		{
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
AlbisOeffneAkte(Pat, PdfPath:="") {                                                            	;-- ursprüngl. ScanPool.ahk - Akte lässt sich auch öffnen, selbst wenn der Name nicht ganz korrekt geschrieben ist

		AllreadyOpen       	:= 0
		CurrPat    				:= Object()
		CurrPat    				:= VorUndNachname(AlbisCurrentPatient())

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Überprüfen ob die Patientenakte schon geöffnet ist
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		If (Pat[5] = "")
		{
				PraxTT("Öffne die Akte des Patienten`n#2" Pat[1], "3 3")
				AllreadyOpen := FuzzyNameMatch(CurrPat, Pat)
		}
		else
		{
				If InStr(AlbisAktuellePatID(), Pat[5])
						AllreadyOpen := 1
		}

		If AllreadyOpen
		{
				AlbisZeigeKarteikarte() 		; sicherstellen das die Karteikarte angezeigt wird
				PraxTT("Die Patientenakte ist schon geöffnet!", "3 3")
				return 1
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Patient öffnen - Dialog: Aufruf des Fensters und Texteingabe
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		AlbisActivate(2)

	; 32768 Menu Kommando für den 'Patient öffnen' Auswahldialog
		while !WinExist("Patient öffnen ahk_class #32770")
		{
				PostMessage, 0x111, 32768,,, % "ahk_id " AlbisWinID()
				Sleep, 200
				If (A_Index > 10) {
						PraxTT("Der Dialog zum Öffnen einer Patientenakte`nließ sich nicht aufrufen.`nDer Vorgang wurde abgebrochen!", "5 0")
						return 0
				}
		}

		WinActivate, Patient öffnen ahk_class #32770
		WinWaitActive, Patient öffnen ahk_class #32770,, 3

		If VerifiedSetText("Edit1", (Pat[5]="" ? Pat[1] : Pat[5]), "Patient öffnen ahk_class #32770", 200)			;Pat enthält nur die ID wenn es kein Array ist
		{
				MsgBox, 1, Addendum für Albis on Windows - %A_ScriptName%,
                             (LTrim
                            			   Der Patientenname konnte nicht eingetragen werden.

                            		1. Bitte manuell per Strg+c Taste einfügen (Kopie aus dem Clipboard).
                            		2. Drücken Sie dann auf 'Ok" im 'Patient öffnen' - Dialog
                            		3. Warten Sie das Öffnen der Akte ab,
                            		4. Drücken Sie erst jetzt auf den 'Ok"-Button in diesem Dialog!
                            	)
		}
		else
		{
			  ; Patient öffnen - Dialog: Ok oder Enter drücken
				LastPopUpWin:= GetLastActivePopup(AlbisWinID())
				while WinExist("Patient öffnen ahk_class #32770")
				{
						;Button OK drücken
						VerifiedClick("Button2", "Patient öffnen ahk_class #32770")
						WinWaitClose, Patient öffnen ahk_class #32770,, 1

						;Fenster ist immer noch da? Dann sende ein ENTER.
						if WinExist("Patient öffnen ahk_class #32770")
						{
                            	WinActivate, Patient öffnen ahk_class #32770
                            		WinWaitActive, Patient öffnen ahk_class #32770,, 1
                            	ControlFocus, Edit1, Patient öffnen ahk_class #32770
                            	SendInput, {Enter}
						}

						Sleep, 300
				}
		}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; 2 Folgedialoge sind möglich - 1. Patient nicht vorhanden , 2. Listview-Auswahl -> den will ich durch möglichst vermeiden
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		PopWin:= WaitForNewPopUpWindow(AlbisWinID(), LastPopUpWin, 10)

	  ;Dialog Patient <.....,......> nicht vorhanden schließen.
		If ( Instr(PopWin.Title, "ALBIS") AND Instr(PopWin.Class, "#32770") AND  Instr(PopWin.Text, "nicht vorhanden") ) {
				LastPopUpWin:= GetLastActivePopup(AlbisWinID())
				VerifiedClick("Button1", "", "", PopWin.Hwnd)			;soll das Fenster schließen
				PopWin:= WaitForNewPopUpWindow(AlbisWinID(), LastPopUpWin, 10)
		}
	  ;Dialog Patient auswählen mit eingebettetem ListView-Fenster
		If ( Instr(PopWin.Title, "Patient") AND Instr(PopWin.Class, "#32770") AND  Instr(PopWin.Text, "List1") ) {
				;Scite_OutPut("Listenfenster ist aufgetaucht", 0, 1, 0)
			return 0
		}
	  ;Dialog Patient öffnen schließen der sich nach dem Schließen des "<Patient ... nicht vorhanden>" Dialoges erneut öffnet
		If ( Instr(PopWin.Title, "Patient öffnen") AND Instr(PopWin.Class, "#32770") ) {
				LastPopUpWin:= GetLastActivePopup(AlbisWinID())
				VerifiedClick("Button3", "", "", PopWin.Hwnd)			;soll das Fenster schließen
				PopWin:= WaitForNewPopUpWindow(AlbisWinID(), LastPopUpWin, 10)
				return 0
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; wartet auf das Öffnen der angeforderten Akte
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		PraxTT( "Warte bis zu 10s auf die Patientenakte (" Name[5] ", " Name[1] ")", "10 3" )
		start:= A_TickCount
		Loop
		{
				If FuzzyNameMatch(CurrPat, Pat) or InStr(AlbisAktuellePatID(), Pat[5])
				{
						PraxTT("Die Patientenakte ist jetzt bereit!", "3 3")
						return 1
				}

				Sleep, 50
				timesince:= (A_TickCount - start)/1000
				GuiControl, PraxTT:, PraxTTCounter2, % Round( (10-timesince), 1 ) "s"
				GuiControl, PraxTT:, PraxTTCounter1, % Round( (10-timesince), 1 ) "s"

		} until timesince > 9.9

		PraxTT("Die richtige Patientenakte konnte nicht geöffnet werden!!", "3 3")

	;}

return 0
}
;13
AlbisKarteikarteAktiv() {                                                                               	;-- gibt true zurück wenn eine Karteikarte angezeigt wird

	If InStr(AlbisGetActiveWindowType(), "Patientenakte")
		return true

return false
}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; LABOR ABRUF AUTOMATION / LABOR DIVERSES                                                                                                                                                                                     	(08)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) MCSVianova_WebClient       	(02) AlbisLaborDatenholen        	(03) AlbisOeffneLaborbuch            	(04) AlbisRestrainLabWindow     	(05) AlbisDruckeLaborblatt
; (06) AlbisZeigeLaborBlatt              	(07) AlbisLaborAuswählen          	(08) AlbisLaborDaten
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;01
MCSVianova_WebClient(MCSinfoBoxID) {                                                              	;-- WebClient Automation (Labor IMD)

	; letzte Änderung vom 08.06.2020

	;##Hersteller des WebClienten haben die Gui dieses WebClienten geändert, Button haben andere ClassNN bekommen, brauche eher eine Suchfunktion für das ClassNN
		static MCS_WinClass  	:= "ahk_class WindowsForms10.Window.8.app.0.378734a"
		MCSinfoBoxID           	:= WinExist("MCS vianova infoBox-webClient")		                                    	;übergebenes Hookhandle ist nicht immer das Fensterhandle

	; Web-Client wurde abgefangen, Verarbeitungsflags werden gesetzt                                                                	;{
		Addendum.LaborAbruf.Status	:= true
		Addendum.LaborAbruf.Daten:= false
		Addendum.LaborAbruf.Voll  	:= false

		MsgBox, 0x1024, Addendum für Albis on Windows,
		(LTrim
		Laborabruf erkannt!

		Möchten Sie alle weiteren Schritte bis hin zum
		Übertragen der Labordaten in die Patienten-
		akten automatisch durchführen lassen?
		)

		IfMsgBox, No
		{
			Addendum.LaborAbruf.Status	:= false
			return
		}

		Addendum.LaborAbruf.Voll := true
	;}

	; WindowsForms10.BUTTON.app.0.34f5582_r6_ad12
	; Button "Abruf" drücken                                                                                                                                   	;{
		If err := Controls("", "ControlClick, &Abruf, BUTTON", MCS_WinClass)
			If !VerifiedClick("WindowsForms10.BUTTON.app.0.34f5582_r6_ad12", MCS_WinClass)	{
					Addendum.LaborAbruf.Status	:= false
					Addendum.LaborAbruf.Voll	:= false
					return 0
			}

		msg := "Das Auslesen des MCSVianova Fenster`nhat nicht funktioniert, wenn man das`nhier lesen kann"
	;}

	; 4.1.4.5.4.1.4.4.4.1.2
	; wartet bis keine Daten mehr hinzukommen                                                                                                   	;{
		oAcc     	:= Acc_Get("Object", "4.1.4.5.4.1.4.4.4" , 0, "ahk_id " MCSinfoBoxID)
		stopAt   	:= 50
		stopIndex	:= 0
		Loop {

			If (oAcc.accChildCount() <> lastChildCount) {
				lastChildCount := oAcc.accChildCount()
				stopIndex := 0
			}

			sleep 100
			stopIndex ++
			If (stopIndex >= stopAt)
					break
		}
	;}

	; liest den Fensterinhalt in einen String und durchsucht dann den Inhalt                                                            	;{
		t := "|"
		Loop % oAcc.accChildCount()
		{
				oAcc := Acc_Get("Object", "4.1.4.5.4.1.4.4.4." A_Index, 0, "ahk_id " MCSinfoBoxID)
				t .= oAcc.accValue(2) "`n"
		}

		; keine Daten - Abbruch
		If RegExMatch(t, "Keine Daten zum Abruf") {
				msg :="Es liegen keine neuen Labordaten vor!`nDer Laborabruf wird nicht fortgesetzt."
				Addendum.LaborAbruf.Daten:= false
				Addendum.LaborAbruf.Status	:= false
				Addendum.LaborAbruf.Voll	:= false
				PraxTT(msg, "0 4")
				return 0
		}

		PraxTT("Es liegen neue Labordaten vor!`nDie Daten werden für die spätere Auswertung gesichert.", "0 4")
		Addendum.LaborAbruf.Daten := true
		NoCopy := 0

		; sucht nach den LDT Dateinamen (es können mehrere sein)
		Loop, Parse, t, `n
		{
				if RegExMatch(A_LoopField, "\|\s*Sichere\s*\<(?<File>.*?)\>\s*ins\s*Archiv\|", LDT) {	; Sichere <X01492.LDT> ins Archiv
						PraxTT("Sichere LDT-Datei:`n" LDTFile, "0 4")
						NoCopy += AlbisLaborLDTBackup(LDTFile)
				}
		}

		If NoCopy
			PraxTT(NoCopy ((NoCopy > 1) ? " LDT-Dateien konnten ":" LDT-Datei konnte ") "nicht gesichert werden!", "0 4")
		else
			PraxTT("Ein Backup der übermittelten Labordatendatei`nwurde erstellt!", "6 4")

		sleep, 4000

	;}

	; Button "Schliessen" drücken                                                                                                                          	;{
		err:= Controls("", "ControlClick, Schließen, BUTTON", MCS_WinClass)
		VerifiedClick("WindowsForms10.BUTTON.app.0.378734a1", MCS_WinClass)
	;}

	; Schließen des Fensters abwarten, wenn das Fenster nicht schließt über andere Befehle das Fenster schliessen  	;{
		WinWaitClose, % MCS_WinClass,, 6
		If ErrorLevel 	{

			; Prozess ID ermitteln
				Process, Exist, infoBoxWebClient.exe
				iBoxWebPID:= ErrorLevel

			; mittels PID zunächst ein normales WinClose und falls nicht erfolgreich ein schließen des Prozeß versuchen
				WinClose, % "ahk_pid " iBoxWebPID
				WinWaitClose, % "ahk_pid " iBoxWebPID,, 1
				If ErrorLevel {
					SendMessage 0x112, 0xF060,,, % "ahk_pid " iBoxWebPID 			; WMSysCommand + SC_Close
					WinWaitClose, % "ahk_pid " iBoxWebPID,, 1
				If ErrorLevel {
					SendMessage 0x10, 0,,, % "ahk_pid " iBoxWebPID 			; WM_Close
					WinWaitClose, % "ahk_pid " iBoxWebPID,, 1
				If ErrorLevel {
					SendMessage 0x2, 0,,, % "ahk_pid " iBoxWebPID 			; WM_Destroy
					WinWaitClose, % "ahk_pid " iBoxWebPID,, 1
				If ErrorLevel
					Process, Close, % "ahk_pid " iBoxWebPID
					If ErrorLevel
						PraxTT("Das Fenster: MCSVianova_WebClient`nkonnte nicht geschlossen werden!", "10 2")
				}
				}
				}
		}
	;}

	; mit Aufruf des Dialogfensters 'Daten holen' wird der automatisierte Vorgang fortgesetzt
		PraxTT("", "off")
		AlbisLaborDatenholen()


return 1 ; 1 steht für erfolgreich
}
;01A
AlbisLaborLDTBackup(LDTDateiName) {

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
AlbisLaborDatenholen() {                                                                                       	;-- öffnet das Fenster Daten holen
	Albismenu(32965, "Labordaten ahk_class #32770") 	; Menupunkt 32965 - Extern\Labor\Daten ?????
return
}
;03
AlbisOeffneLaborbuch() {                                                                                      	;-- oeffnet das Laborbuch
	AlbisActivate(2)
	AlbisIsBlocked(AlbisWinID())
	PostMessage, 0x111, 34162,,, % "ahk_id " AlbisWinID()		;Öffnen des Laborbuches per WM_Command Befehl - Laborbuch:= 34162
	WinWait, Laborbuch ahk_class OptoAppClass,,6
return ErrorLevel
}
;04
AlbisRestrainLabWindow(nr, hwnd) {                                                                      	;-- liest Daten eines Fenster im Laborbuch bei Aufruf von 'ins Laborbuch übertragen' ein

	; liest Daten eines Fenster im Laborbuch bei Aufruf von 'ins Laborbuch übertragen' ein und speichert diese in eine Datei zur späteren Verarbeitung
	;  (!Funktionsumfang ist noch nicht vollständig)
		static Labordaten

	; setze ich auf 0 zurück, da ich nicht wie weiß wie ich feststellen kann das alle Laborwerte in die Akten sortiert wurden
		Addendum.LaborAbruf.Status := false

		If (nr=1) {

				idx:=0
				ControlGetText, AfNr     	, Static2	, % "ahk_id " hwnd                                                 	;AnforderungsNr
				ControlGetText, BStatus 	, Static8	, % "ahk_id " hwnd                                                 	;End- oder Teilbefund diese Daten möchte ich
				ControlGetText, EDatum	, Static4	, % "ahk_id " hwnd                                                 	;Eingangsdatum
				ControlGetText, ADatum	, Static11	, % "ahk_id " hwnd                                                 	;Abnahmedatum
				ControlGetText, PNam   	, Static6	, % "ahk_id " hwnd                                                  	;Patientenname

				pre:= SubStr(ADatum, StrLen(ADatum)-3, 4) . SubStr(ADatum, StrLen(ADatum)-6, 2)

				Labordaten .=  A_DD "." A_MM "." A_YYYY ";" AfNr ";" BStatus ";" EDatum ";" ADatum ";" PName "`n"
				FileAppend, % Labordaten, % AddendumDir "\Module\Albis_Funktionen\Laborabruf\Labordaten\Laborbuch\" pre "-Labordaten.txt"

				ControlGet, content, List,, ListBox1, % "ahk_id " hwnd
				Loop, Parse, content, `n
						idx++

				If (idx>1)	{

						searchstr:= SubStr(ADatum, StrLen(ADatum)-5, 1) . "/" . SubStr(ADatum, StrLen(ADatum)-1, 2)
						Loop, Parse, content, `n
                            	If InStr(A_LoopField, searchstr)
                            			Control, Choose, % A_Index, ListBox1, % "ahk_id " hwnd
				}

				sleep, 1000

				VerifiedClick("Button8","","", hwnd)
				return Labordaten

		}

return ""
}
;05
AlbisDruckeLaborblatt(Spalten, Drucker:="", SaveDir:="") {                                     	;-- Automatisiertes Drucken eines Laborbefundes

	; im Moment kann nur die Anzahl der auszudruckenden Spalten als Option übergeben werden
	; Drucker := "pdf" für PDF Drucker - einzustellen in den Addendum.ini
	; Drucker := "Standard" für den eingestellten Standard Drucker - zu hinterlegen in der Addendum.ini

		If !InStr(AlbisGetActiveWindowType(), "Patientenakte")
		{
				PraxTT("Der Laborblattdruck kann nur bei einer geöffneten`nPatientenakte ausgeführt werden!.", "6 3")
				return 0
		}

		static oDrucker:= Object()
		static printInit, Fenster_Druckeinrichtung, Fenster_Laborblatt_Druck

	; bei Druckausgabe in eine PDF-Datei einen neuen Dateinamen finden
		If InStr(Drucker, "pdf")
		{
				Name:= StrSplit(AlbisCurrentPatient(), ",", A_Space)
				If !(SubStr(SaveDir, -1) = "\")
						SaveDir .= "\"

			; noch nicht vorhandenen Dateinamen finden
				Loop
				{
						filename:= "BW-" SubStr(Name[1], 1, 2) SubStr(Name[2], 1, 2) ( A_Index = 1 ? "" : "-" A_Index - 1 ) "_Ausdruck_vom_" A_DD "." A_MM "." A_YYYY
						If !FileExist(SaveDir . filename . ".pdf")
                            	break
				}
		}


	; Erstinitalisierung
		If !printInit
		{
				IniRead, PDF       	, % AddendumDir "\Addendum.ini", % CompName, Drucker_pdf
				IniRead, Standard	, % AddendumDir "\Addendum.ini", % CompName, Drucker_Standard
				Fenster_Druckeinrichtung	:= "Druckeinrichtung ahk_class #32770"
				Fenster_Laborblatt_Druck	:= "Laborblatt Druck ahk_class #32770"
				Fenster_Drucken             	:= "Drucken ahk_class #32770"
				Fenster_Druckausgabe   	:= "Druckausgabe speichern unter"
				oDrucker.PDF                	:= PDF
				oDrucker.Standard        	:= Standard
				printInit                         	:= 1

		}

	; Auslesen des eingestellten Druckers und Zwischenspeichern der Einstellungen für späteres Wiederherstellen
		Albismenu(57606, Fenster_Druckeinrichtung, 3, 1)
		oDrucker.ActivePrinter:= { "Name"     	:ControlGet("Choice" 	,, "ComboBox1"	, Fenster_Druckeinrichtung)
		                                		, "Size"        	:ControlGet("Choice" 	,, "ComboBox2"	, Fenster_Druckeinrichtung)
			                                 	, "Source"   	:ControlGet("Choice" 	,, "ComboBox3"	, Fenster_Druckeinrichtung)
				                                , "Portrait"    	:ControlGet("Checked"	,, "Button5"     	, Fenster_Druckeinrichtung)
					                           , "Landscape":ControlGet("Checked"	,, "Button6"     	, Fenster_Druckeinrichtung)}

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
		If AlbisZeigeLaborBlatt()
		{
				; Druckaufruf ausführen
					Albismenu(57607, Fenster_Laborblatt_Druck, 3, 1)
					sleep, 500

				; Überprüft nochmals den ausgewählten Drucker
					VerifiedClick("Button11", Fenster_Laborblatt_Druck) 	                    	; Drucker...
					WinWaitActive, % Fenster_Drucken,, 10
					sleep, 500
					If !InStr(ControlGet("Choice",, "ComboBox1", Fenster_Drucken), Dokumentdrucker)
                            Control, ChooseString, % Dokumentdrucker, ComboBox1, % Fenster_Drucken
					sleep, 300
					If !VerifiedClick("Button11", Fenster_Drucken)                                 	; OK - Button im Fenster 'Drucken'
					{
                            PraxTT("Der Druckereinstellungsdialog hat sich nicht schließen lassen.`nDie Funktion wird jetzt abgebrochen!", "6 3")
                            return
					}

				; Druck vorbereiten
					VerifiedCheck("Button2", Fenster_Laborblatt_Druck)                        	; letzte  ..... Spalten
					VerifiedSetText("Edit1", Spalten, Fenster_Laborblatt_Druck, 200)    	; Spaltenanzahl
					VerifiedCheck("Button5", Fenster_Laborblatt_Druck)                        	; Anmerkungen und Probedaten
					sleep, 500

				; Drucken
					VerifiedClick("Button9", Fenster_Laborblatt_Druck)                          	; OK - Button

				; Druckausgabe speichern unter Fenster abfangen
					If InStr(Drucker, "pdf") {
                            WinWait , % Fenster_DruckAusgabe,, 10
                            If ErrorLevel
                            {
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
					AlbisZeigeKarteikarte()

				; Wiederherstellen der vorherigen Druckeinstellungen
					Albismenu(57606, Fenster_Druckeinrichtung, 3, 1)

					while !( ControlGet("Choice", "", "ComboBox1", Fenster_Druckeinrichtung) = oDrucker.ActivePrinter.Name )
					{
                            Control, ChooseString, % oDrucker.ActivePrinter.Name, ComboBox1, % Fenster_Druckeinrichtung
                            sleep, 200
					}

					Control, ChooseString, % oDrucker.ActivePrinter.Size 	, ComboBox2, % Fenster_Druckeinrichtung
					Control, ChooseString, % oDrucker.ActivePrinter.Source, ComboBox3, % Fenster_Druckeinrichtung

					If oDrucker.ActivePrinter.Portrait = 1
                            VerifiedCheck("Button5", Fenster_Druckeinrichtung)	;Checkbox	: Hochformat
					else
                            VerifiedCheck("Button6", Fenster_Druckeinrichtung)	;Checkbox	: Querformat

					VerifiedClick("Button7", Fenster_Druckeinrichtung)          	;Button      	: Ok
		}

	; das Wartezimmer wieder anzeigen
	;	AlbisOeffneWarteZimmer()

return
}
;06
AlbisZeigeLaborBlatt() {                                                                                         	;-- schaltet auf das Laborblatt

		If !InStr(activeWT:= AlbisGetActiveWindowType(), "Patientenakte")
				return 0

		AlbisActivate(2)

		while !InStr(AlbisGetActiveWindowType(), "Laborblatt")
		{
				PostMessage, 0x111, 33034,,, % "ahk_id " AlbisWinID()
				sleep, 100
				If InStr(AlbisGetActiveWindowType(), "Laborblatt")
                            break
				If A_Index > 10
						return 0
				sleep, 200
		}

return 1
}
;07
AlbisLaborAuswählen(Laborname) {                                                                       	;-- für das Fenster Labor auswählen

	; übergebener Laborname wird im Dialog ausgewählt und anschliessend mit Ok bestätigt
	; vermutlich erscheint dieses Fenster nicht wenn man nur ein Labor zur Auswahl hat
		If Laborname <> ""
		{
				ControlGet, currSel, Choice,, ListBox1, Labor auswählen ahk_class #32770
				If !InStr(currSel, Laborname)
				{
						Control, ChooseString, % Laborname, ListBox1, Labor auswählen ahk_class #32770
						sleep, 200
						If A_Index > 10
						{
                            	MsgBox, 1, Addendum für Albis on Windows, % "Das eingestellte Labor: " Laborname "`nkonnte nicht ausgewählt werden.", 6
                            	return 0
						}
				}
		}

		err:= VerifiedClick("Button1", "Labor auswählen ahk_class #32770")

		WinWait, ALBIS, Keine Datei(en) im Pfad, 2
		If ErrorLevel = 0
		{
				err:= VerifiedClick("Button1", "ALBIS", "Keine Datei(en) im Pfad")
				LaborAbruf_Status:= 0
				return err
		}

return err
}
;08
AlbisLaborDaten()	{                                                                                                 	;-- bearbeitet das "Labordaten importieren" Fenster und öffnet im Anschluss das Laborbuch
	VerifiedCheck("Button5", "Labordaten ahk_class #32770",,, 1)
	VerifiedClick("Button1", "Labordaten ahk_class #32770")
	AlbisOeffneLaborbuch()
}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; GRAFISCHER BEFUND                                                                                                                                                                                                                             	(04)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisOeffneGrafischerBefund	(02) AlbisUebertrageGrafischenBefund                                               	(03) AlbisImportierePdf          	(04) AlbisImportiereBild
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;01
AlbisOeffneGrafischerBefund() {                                                             	;-- öffnet den Dialog 'Grafischer Befund' - importieren von Bilddateien (jpg z.B.)
return Albismenu(32960, "Grafischer Befund ahk_class #32770")
}
;02
AlbisUebertrageGrafischenBefund(ImgName, Karteikartentext:="") {           	;-- Funktion mit Sicherheitsüberprüfung falls ein Focus verloren geht

	; ein fehlender Dateibez. wird automatisch ersetzt
		If !RegExMatch(ImgName, "\.\w+$")
				ImgName .= ".pdf"

	; Karteikartentext ist leer
		If (Karteikartentext = "")
				Karteikartentext := SubStr(ImgName, 1, StrLen(ImgName) - 4)

		WinActivate   	, Grafischer Befund ahk_class #32770
		WinWaitActive	, Grafischer Befund ahk_class #32770,,4

	;der Haken bei Vorschau wird entfernt, dann muss ich keine PDF Viewer oder Imageviewer Fenster abfangen und abhandeln mehr
		If !(hGBef:= WinExist("Grafischer Befund ahk_class #32770"))
			return

		VerifiedCheck("Button1", "", "", hGBef, 0)
		sleep, 200

	;Daten in das Fenster eintragen
		VerifiedSetText("Edit1", ImgName       	, "Grafischer Befund ahk_class #32770", 100)
		VerifiedSetText("Edit2", Karteikartentext	, "Grafischer Befund ahk_class #32770", 100)

	;auf Ok drücken - ein mögliches Fehlerfenster abfangen und schliessen (beendet den Vorgang insgesamt)
		while WinExist("Grafischer Befund ahk_class #32770")		{

				VerifiedClick("Button2", "Grafischer Befund ahk_class #32770")
				WinWaitClose, Grafischer Befund ahk_class #32770,, 1

				If WinExist("ALBIS ahk_class #32770", "Bitte überprüfen Sie den Namen bzw. Pfad.") {
						VerifiedClick("Button1", "ALBIS ahk_class #32770", "Bitte überprüfen Sie den Namen bzw. Pfad.")
						return 0
				}

				If (A_Index > 2)
						return 0

				Sleep, 1000

		}

return 1
}
;03
AlbisImportierePdf(PdfName, LVRow:=0) {                                             	;-- Importieren einer PDF-Datei in eine Patientenakte

	; benötigt ein globales Objekt mit dem Namen SPData -> Aufbau siehe ScanPool.ahk
	; oPat : globals Object, enthält die Addendum interne Patientendatenbank
	; Addendum: globales Object - enthält Daten zu Albis-Ordnern

		suggestions:= Object()

	; für das ScanPool-Skript, ansonsten wird davon ausgegangen das die richtige Patientenakte geöffnet ist
		If InStr(A_ScriptName, "ScanPool") {

			; Pdf Previewer ausschalten
				PrevBackup:= dpiShow, dpiShow:= "Aus"

			; Patientenname ermitteln
				Name := ExtractNamesFromFileName(PdfName)

			; Prüfen ob der Patient in der Datenbank vorhanden ist und ermitteln einer Patienten ID wenn möglich
				If oPat.Count()	{

						suggestions:= PatInDB_SiftNgram(Name)
						If IsObject(suggestions)	{
									If suggestions.MaxIndex() = 1
											Name[5]:= suggestions.1.PatID
						}
						else	                        	{
								InputBox, newName, ScanPool, % "Patient: " Name[1] "konnte in der Addendum-Datenbank nicht gefunden werden.`nBitte geben Sie hier den korrekten Namen ein!",, 400,,,,,, % Name[1]
								Name:= ExtractNamesFromFileName(newName)
						}

				}

			; Patientenakte öffnen
				PraxTT("Öffne die Akte von: `n" . Name[1] , "3 2")
				If !AlbisOeffneAkte(Name) {
						PraxTT("Die Akte des Patienten`n" . Name[1] . "`nließ sich nicht öffnen! Der angeforderte Vorgang wird jetzt abgebrochen", "6 2")
						return 0
				}

		}

	; Eingangsdatum des Befundes auslesen
		If InStr(A_ScriptName, "ScanPool")
			creationtime := FormatedFileCreationTime(SPData.BefundOrdner "\" PdfName)
		else
			creationtime := FormatedFileCreationTime(Addendum.BefundOrdner "\" PdfName)

	; öffnet den Dialog 'Grafischer Befund' durch Eingabe eines Kürzel in die Akte - falls Sie ein anderes Kürzel verwenden - müssen Sie dies hier ändern
		PraxTT("Importiere den Befund:`n" PdfName, "15 2")
		AlbisSetzeProgrammDatum(creationtime)

	; Eingabefokus in die Karteikarte setzen
		If !AlbisKarteikartenFocusSetzen("Edit3") {
			PraxTT("Beim Übertragen des Befundes ist ein Problem aufgetreten.`nEs konnte kein Eingabefocus in die Karteikarte gesetzt werden.", "6 2")
			return 0
		}

		SendRaw, Scan
		sleep, 100
		SendInput, {Tab}

	; Erstellen des Karteikartentextes
		KarteikartenText := StrReplace(PdfName, ".pdf")            	;entfernen von '.pdf'

		If InStr(A_ScriptName, "ScanPool") {
			KarteikartenText := StrReplace(KarteikartenText, Name[3], "")
			KarteikartenText := StrReplace(KarteikartenText, Name[4], "")
		} else {
			KarteikartenText := RegExReplace(KarteikartenText, "^[\w\p{L}-]+\s*\,*\s*[\w\p{L}]+\s*\,*\s*", "")
		}

		KarteikartenText := RegExReplace(KarteikartenText, "^\s*\,*\s*")
		KarteikartenText := StrReplace(KarteikartenText, "   ", "°")
		KarteikartenText := StrReplace(KarteikartenText, "  ", "°")
		KarteikartenText := StrReplace(KarteikartenText, "°", " ")

	; Übertragen des Befundes in die Akte
		If !AlbisUebertrageGrafischenBefund(PdfName, KarteikartenText) 	{
				PraxTT("Beim Übertragen des Befundes`nist ein Problem aufgetreten.", "6 2")
				return 0
		}

	; ScanPool nutzt derzeit noch andere Variablen
		If InStr(A_ScriptName, "ScanPool") {

			; Daten sichern (PdfData ist global - ein return nicht notwendig)
				PdfData.KarteikartenText 	:= KarteikartenText
				PdfData.PdfFullPath       	:= PdfName
				PdfData.ReaderID 	        	:= ReaderID

			; Zurücksetzen von gemachten Einstellungen
				DpiShow	:= PrevBackup
		}
		else	{

				SPData.KarteikartenText 	:= KarteikartenText
				SPData.PdfFullPath        	:= PdfName
				SPData.ReaderID 	        	:= ReaderID

		}

	; auf aktuelles Tagesdatum zurücksetzen
		AlbisSetzeProgrammDatum()

		PraxTT("Der Befund wurde importiert.", "3 2")

return creationtime
}
;04
AlbisImportiereBild(ImageName, KarteikartenText) {                                	;-- Importieren einer JPG-Datei in eine Patientenakte

	; Dialogfenster aufrufen über das Albismenu
		If !(hGBef := AlbisOeffneGrafischerBefund()) 	{
				PraxTT("Der Dialog zum Importieren der Bilddatei:`n" ImageName "konnte nicht geöffnet werden.`nDer nächste Vorgang wird abgebrochen.", "2 2")
				return 0
		}

	; Dialog aktivieren
		WinActivate   	, Grafischer Befund ahk_class #32770
		WinWaitActive	, Grafischer Befund ahk_class #32770,,4

	; Tagesdatum ändern
		AlbisSetzeProgrammDatum(creationtime := FormatedFileCreationTime(Addendum.BefundOrdner "\" ImageName))

	; Importieren des Bildes in die Akte
		If !AlbisUebertrageGrafischenBefund(ImageName, KarteikartenText) 	{
				PraxTT("Beim Übertragen des Befundes`nist ein Problem aufgetreten.", "6 2")
				return 0
		}

	; auf das aktuelle Tagesdatum zurücksetzen
		AlbisSetzeProgrammDatum()

return creationtime
}
;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; PROTOKOLLE, STATISTIKEN - ERSTELLUNG UND AUSWERTUNG                                                                                                                                                                	(03)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisErstelleTagesprotokoll       	(02) AlbisErstelleZiffernStatistik     	(03) AlbisListeSpeichern
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
; 01
AlbisErstelleTagesprotokoll(Periode:="", SaveFolder:="", CloseProtokoll:=1)     	;{--	Erstellen und Speichern von Tagesprotokollen
{
	;Parameter Periode: 	    1. leerer String                         	- setzt den Zeitraum auf den ersten und letzten Tag des aktuellen Quartals
	;                            		2. mmyy (z.B. 0219)                 	- es wird ein Tagesprotokoll mit dem ersten und letzten Tag des übergebenen Quartals erstellt
	;                            		3. "dd.dd.dddd[,/-]dd.dd.dddd"	- von bis Datumsübergabe ist auch als String mit zwei Tagesdaten welche durch ein "-" oder "," getrennt sind möglich
	;
	; Beispiel               :		AlbisErstelleTagesprotokoll("04.01.2016-06.01.2016", AddendumDir "\Tagesprotokolle\TP-Abrechnungshelfer", 1)

		Quartal     	:= Object()
		SaveFolder	:= RTrim(SaveFolder, "\")

	; TAGESPROTOKOLLFENSTER AUFRUFEN
		hTprotWin	:= Albismenu(32802, "Tagesprotokoll")

	; PARAMETER PARSEN
		periode:= StrReplace(periode, "/")
		periode:= StrReplace(periode, " ")

		If RegExMatch(Periode, "^\d{4}$", QuartalAkt)
		{
				Quartal.aktuell	:= QuartalAkt
				Quartal         	:= QuartalTage(Quartal)
				QNr             	:= SubStr(QuartalAkt, 2, 1)
		}
		else If RegExMatch(Periode, "^\d{3}$", QuartalAkt)
		{
				Quartal.aktuell	:= "0" QuartalAkt
				Quartal         	:= QuartalTage(Quartal)
				QNr              	:= SubStr(QuartalAkt, 1, 1)
		}
		else if StrLen(Periode) = 0
		{
				Quartal.aktuell	:= GetQuartal("heute")
				Quartal         	:= QuartalTage(Quartal)
				QNr             	:= SubStr(QuartalAkt, 2, 1)
		}
		else If RegExMatch(Periode, "^\d\d\.\d\d\.\d\d\d\d", ErsterTag) && RegExMatch(Periode, "(?<=-|,)\d\d\.\d\d\.\d\d\d\d", LetzterTag)
		{
				Quartal.ErsterTag  	:= ErsterTag
				Quartal.LetzterTag	:= LetzterTag
		}
		else
		{
				throw A_ThisFunc " (" A_LineFile ") : Error in function call, a date must be passed in the following format`n - dd.mm.yyyy[-/,]dd.mm.yyyy or qq[// ]yy"
				return 1
		}

	; FELDER VON UND BIS MIT DEN JEWEILS ÜBERGEBEN DATEN FÜLLEN
		If !Controls("Edit1", "SetText " Quartal.ErsterTag, hTprotWin)
				PraxTT("Das Anfangsdatum (von) konnte nicht gesetzt werden.", "6 3"), return 1
		If !Controls("Edit2", "SetText " Quartal.LetzterTag, hTprotWin)
				PraxTT("Das Enddatum (bis) konnte nicht gesetzt werden.", "6 3"), return 1

	; EINEN MAUSCLICK AN DEN 'OK' BUTTON DES FENSTER SENDEN
		If !Controls("", "ControlClick, OK, Button", hTprotWin)
				Controls("Button31", "Click use ControlClick", hTprotWin)
		WinWait, % "Bitte warten... ahk_class #32770",, 10
		If ErrorLevel && WinExist("Tagesprotokoll") {
				PraxTT("Der ausgelöste Click auf den 'OK' Button`nim Tagesprotokoll-Dialog ist fehlgeschlagen.`nEs wurde als letztes ein simulierter Mausklick versucht.", "6 3")
				WinClose, Tagesprotokoll ahk_class #32770
				return 1
		} else if ErrorLevel && !WinExist("Tagesprotokoll") {
				PraxTT("Es ist ein unbekannter Fehler aufgetreten,`nwelcher die Erstellung des Tagesprotokoll verhindert hat.", "6 3")
				return 1
		}

	; DATEN ZUM TAGESPROTOKOLLFENSTER IN Controls LEEREN
		Controls("","Reset", hTprotWin)

	; AUF DAS FERTIG ERSTELLTE TAGESPROTOKOLL WARTEN
		bitteWarten:= false
		PraxTT("Warte auf das vollständig erstellte Tagesprotokoll!", "0 3")
		while !WinExist("ALBIS - [Tagesprotokoll") {
				Sleep, 500
			; läßt den Loop pausieren
				If hBWin := WinExist("Bitte warten ahk_class #32770") {
					WinGetText, wText, % "ahk_id " hBWin
					;ToolTip, % wText
					If InStr(wText, "Tagesprotokoll") {
							bitteWarten := true
							BWin	:= GetWindowSpot(hBWin)
							TPWin	:= GetWindowSpot(WinExist("Tagesprotokoll ahk_class #32770"))
							SetWindowPos(hBwin, TPWin.X + Floor(TPWin.W//2) - Floor(BWin.W//2), TPWin.Y + TPWin.H + 5, BWin.W, BWin.H)
							WinSet, ExStyle, +0x8, % "ahk_id " hBWin	;+ WS_EX_TOPMOST
							WinWaitClose, % "Bitte warten... ahk_class #32770"
					}
				}
		}
		PraxTT("", "off 3")

	; SPEICHERN DES ANGEZEIGTEN TAGESPROTOKOLL AUF FESTPLATTE
		If SaveFolder
		{
			; Tastatur- und Mauseingriffe durch den Nutzer kurzzeitig sperren
				hk(1, 0, "!ACHTUNG!`n`nDie Tastatureingaben sind für maximal 10 Sekunden gesperrt", 10)

			; Tagesprotokoll speichern und die Tagesprotokollausgabe schließen
				hSaveAsWin	:= Albismenu(33014, "Speichern unter")
				WinWaitActive, % "Speichern unter ahk_class #32770",, 10
				If ErrorLevel
				{
						MsgBox, 0, Addendum für Albis on Windows, Erwarteter 'Speichern unter' - Dialog konnte nicht abgefangen werden!`nDie Funktion wird hier abgebrochen
						return 1
				}

			; ermitteln ob ein Tagesprotokoll mit selben Datumsangaben schon einmal erstellt wurde
				If FileExist(SaveFolder "\" Quartal.LetztesJahr "(" SubStr(Quartal.ErsterTag, 1, 6) "-" SubStr(Quartal.LetzterTag, 1, 6) ".txt")
					tprotFileExists:=1

			; Dateinamen erstellen
				If StrLen(Quartal.aktuell) = 0
					ProtokollFileName:= SaveFolder "\" Quartal.ErsterTag "-" Quartal.LetzterTag ".txt"
				else
					ProtokollFileName:= SaveFolder "\" Quartal.Jahr "-" QNr ".txt"

			; Ordner und Dateinamen im Speichern unter Dialog eintragen
				If FileExist(ProtokollFileName)
						FileDelete, % ProtokollFileName

				Controls("Edit1", "SetText, " ProtokollFileName, hSaveAsWin)

			; Speichern unter drücken
				Controls("Button2", "Click use Controlclick", hSaveAsWin)
				sleep 200

			; Warten auf einen eventuellen "Speichern unter bestätigen" - Dialog
				PraxTT("Warte 5 Sekunden auf einen möglichen Speichern unter... Dialog!", "0 5")
				WinWait, % "Speichern unter bestätigen ahk_class #32770",, 5
				If !ErrorLevel
					Controls("Button1", "Click use ControlClick", "Speichern unter bestätigen ahk_class #32770")

			; Tastatur- und Mauseingriffe wieder entsperren
				hk(0, 0, "Tastatur- und Mausfunktion sind wieder hergestellt!", 2)
		}

	; SCHLIESSEN DES PROTOKOLLS, WENN PARAMETER GESETZT
		If CloseProtokoll
			AlbisCloseMDITab("Tagesprotokoll")

	PraxTT("Das Tagesprotokoll wurde erstellt und ist gespeichert!", "6 3")

return ProtokollFileName		; ein vollständiger Dateipfad als Rückgabe = erfolgreicher Funktionsablauf
}
QuartalTage(Quartal) {

	If !IsObject(Quartal)
	{
			throw Exception(A_ThisFunc " (" A_LineFile ") : Error in function call, parameter of function must be an object!")
			return 1
	}

	Quartal.Quartal    	:= SubStr(Quartal.aktuell, 1, 2)
	Quartal.Jahr	         	:= SubStr(Quartal.aktuell, 3, 2)

	Quartal.Jahr          	:= Quartal.Jahr > SubStr(A_YYYY, 3, 2) ? ("19" Quartal.Jahr) : ("20" Quartal.Jahr)
	Quartal.ErsterMonat	:= (Quartal.Quartal - 1) * 3 + 1
	Quartal.LetzterMonat	:= SubStr("0" . Quartal.ErsterMonat + 2, -1)
	Quartal.ErsterTag	   	:= "01." SubStr("0" . Quartal.ErsterMonat, -1) "." Quartal.Jahr
	Quartal.LetzterTag    	:= days_in_month(Quartal.Jahr Quartal.LetzterMonat) "." Quartal.LetzterMonat "." Quartal.Jahr

return Quartal
}
days_in_month(date:="") {
    date := (date = "") ? (a_now) : (date)
    FormatTime, year,  % date, yyyy
    FormatTime, month, % date, MM
    month += 1                 ; goto next month
    if (month > 12)
        year += 1, month := 1  ; goto next year, reset month
    month := (month < 10) ? (0 . month) : (month)  ; 0 to 01
    new_date := year . month
    new_date += -1, days       ; minus 1 day
    return subStr(new_date, 7, 2)
}
;} end sub AlbisTagesprotokollErstellen
; 02
AlbisErstelleZiffernStatistik(Quartal:="aktuell", Zeitraum:="", Tag:=""                	;--	Erstellen von Ziffernstatistiken
, Arztwahl:="", SaveToFile:="", CloseStatistik:=true) {

	; Menu Statistik\Leistungsstatistik\EBM2000Plus\2009Ziffernstatistik
	; ACHTUNG: Funktion kommuniziert mit Addendum.ahk - das Skript sollte vorher aufgerufen sein
	; Arztwahl wird folgend funktionieren. Übergabe eines String mit "BSNR 123456678" oder "Arzt Dr.Mustergültig" oder "Person Dr.Kontrolle" zur Auswahl der Arztwahlfelder

		static WM_MDIMAXIMIZE:= 0x0225

		M5UserBreak:= false
		Erstellungszeitraum := "das akutelle Quartal"

	; kein Überprüfen des erfolgreichen Dialogaufrufes über die AlbisMenu() Funktion, da der User als Administrator eingeloggt sein muss
		If !WinExist("Ziffernstatistik ahk_class #32770")
			hZfStat:= Albismenu(34293)			;EBM 2000plus/2009 Ziffernstatistik = 34293

	; prüft ob als Administrator eingeloggt wurde, sonst kann keine Ziffernstatistik erstellt werden, wenn ja wird der der Dialog aufgerufen ;{
		Loop
		{
				If WinExist("Ziffernstatistik ahk_class #32770")
					break
				else if WinExist("ALBIS ahk_class #32770", "Zugang verweigert")
				{
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
								If AddendumID := GetAddendumID()
								{
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
								while, WinExist("ALBIS - Login ahk_class #32770")
								{
										If GetKeyState("Escape")
										{
												M5UserBreak:= true
												break
										}
										Sleep, 50
								}

								PraxTT("", "Off")

							; AutoLogin wieder einschalten
								If AddendumID := GetAddendumID()
								{
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
		;}

	; Abbruch wenn das Dialogfenster nicht geöffnet werden könnte
		If !WinExist("Ziffernstatistik ahk_class #32770") {
				PraxTT("Das Ziffernstatistik Fenster konnte nicht geöffnet werden`nDie Funktion wird jetzt beendet.", "5 1")
				return 0
		}

	; Daten werden jetzt in die Steuerelemente des Dialogfensters eingetragen
		PraxTT("Ich starte jetzt mit der Erstellung der Ziffernstatistik", "3 1")

		VerifiedClick("Button1", "Ziffernstatistik ahk_class #32770") ; Quartal

		If (StrLen(Zeitraum) > 0)
		{
				RegExMatch(Zeitraum, "(?<von>[\d\.]+)\s*,\s*(?<bis>[\d\.]+)", Datum)
				If !RegExMatch(DatumVon, "\d\d\.\d\d\.\d\d\d\d") || !RegExMatch(DatumBis, "\d\d\.\d\d\.\d\d\d\d") {
						throw Exception("Datumsformat(TT.MM.JJJJ) wurde nicht eingehalten <" Zeitraum ">")
						return 0
				}

				Erstellungszeitraum := "den Zeitraum " DatumVon " bis zum " DatumBis

				VerifiedClick("Button2", "Ziffernstatistik ahk_class #32770") ; Leistungen im Zeitraum
				VerifiedSetText("Edit1", DatumVon, "Ziffernstatistik ahk_class #32770")
				VerifiedSetText("Edit2", DatumBis, "Ziffernstatistik ahk_class #32770")
		}
		else If (StrLen(Tag) > 0)
		{
				If !RegExMatch(Tag, "\d\d\.\d\d\.\d\d\d\d") {
						throw Exception("Datumsformat(TT.MM.JJJJ) wurde nicht eingehalten <" Tag ">")
						return 0
				}

				Erstellungszeitraum := "den Tag " Tag

				VerifiedClick("Button3", "Ziffernstatistik ahk_class #32770") ; Tag
				VerifiedSetText("Edit3", Tag, "Ziffernstatistik ahk_class #32770")
		}
		else If !RegExMatch(Quartal, "^\s*aktuell\s*$")
		{

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
		; den Button für Optionen und Einstellungen dort habe ich noch nie genutzt

	; drückt den Button Ok um die Statistik zu erstellen
		VerifiedClick("Button26", "Ziffernstatistik ahk_class #32770")	 ; OK

	; wartet bis die Erstellung abgeschlossen ist
		while !AlbisGetSpecificHMDI("Ziffernstatistik")
		{
				If hwnd:= WinExist("ahk_class #32770", "Scheine einlesen")
					WinSetTitle, % "ahk_id " hwnd,, % "Ziffernstatistik für " Erstellungszeitraum " wird erstellt (" A_Index "s)"
				Sleep, 1000
		}

	; wenn abgeschlossen, dann wird das Statistikfenster innerhalb des Albisfenster maximiert
		AlbisActivate(1)

		If hStatistikWin:= AlbisGetSpecificHMDI("Ziffernstatistik") {
			If !AlbisGetMDIMaxStatus(hStatistikWin) ; MDI Fenster ist nicht vergrößert dann maximieren
					PostMessage, % WM_MDIMAXIMIZE, % hStatistikWin,,, % "ahk_id " AlbisGetMDIClientHandle()
		}

	; kurze Pause damit Windows wieder reagieren kann
		sleep 1000

	; wenn angegeben wird das erstellte Protokoll gespeichert
		If (StrLen(SaveToFile) > 0)
		{
				savedAs:= AlbisDateiSpeichern(SaveToFile, true)
		}

	; wenn CloseStatistik = true wird das Statistikfenster geschlossen
		If CloseStatistik
			AlbisCloseMDITab("Ziffernstatistik")

	; wenn savedAs vorhanden ist wird dieser String zurückgegeben
		If (StrLen(savedAs) > 2)
			return savedAs

return 1 	; ansonsten wird nur eine 1 zurückgegeben
}
; 03
AlbisListeSpeichern(ListTitle, FilePath, ext:= "csv", overwrite:= true) {	            	;--	speichert eine in Albis geöffnete Liste in ein gewünschtes Verzeichnis

	; Parameter:
	; 	ListTitle      	- sollte der Titel einer in Albis geöffneten Liste sein, z.B. offene Posten. Achtung: Der Listentitel wird als Dateiname weiterverwendet!
	;	FilePath     	- Dateiverzeichnis in dem diese Liste gespeichert werden soll
	;	ext            	- Dateiendung kann jede in Windows nutzbare sein z.B. txt, log, csv.
	;	overwrite   	- flag als true oder false falls eine bestehende Datei ohne nachfragen überschrieben werden soll, ansonsten erfolgt eine fortlaufende Nummerierung
	;
	; Rückgabewert:
	;	FullFilePath	- der komplette Pfad und der Dateiname der gespeicherten Datei


	; Aktivieren der Liste in Albis
		result := AlbisActivateMDIChild(ListTitle)
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
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; ABRECHNUNG - KASSE UND PRIVAT                                                                                                                                                                                                       	(01)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisAbrechnungVorbereiten
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
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
		, "Nein-Scheine"	                        	: {"Type":"Checkbox"
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

;}
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; SONSTIGE AUTOMATISIERUNGSFUNKTIONEN                                                                                                                                                                                          	(14)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) AlbisWaitAndActivate            	(02) AlbisNeuStart                     	(03) AlbisCaveVonToolTip              	(04) AlbisHotKeyHilfe             	(05) AlbisIsBlocked
; (06) AlbisCloseLastActivePopups   	(07) AlbisDefaultLogin                	(08) AlbisAutoLogin                       	(09) AlbisLogout                       	(10) AlbisActivate
; (11) AlbisCopyCut                        	(12) AlbisIsElevated                      	(13) AlbisSelectAll                          	(14) CheckAISConnector
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
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
	If InStr(Result, "Ja")
	{
            user_result :=1

		;{a ------------- falls Albis nicht gestartet ist, dann wird es jetzt gestartet
            If !AlbisWinID()
			{
					Msgbox,1, Addendum für ALBIS - %CallingProcess%, Albis muss gestartet sein`nDamit dieses Skript verwendet werden kann!, 10
					Run, %AlbisExeDir%\%AlbisExe%, %AlbisWorkDir%, UserErrorLevel, AlbisPID
					If ErrorLevel
					{
                            MonitorScreenShot(1, "Function AlbisNeuStart", AddendumDir . "\logs'n'data\ErrorLogs")
                            FileAppend, % TimeCode(1) "Function:  AlbisNeuStart ErrorLevel: " ErrorLevel "|Error: " A_LastError "|AlbisPID: " AlbisPID "`n", %AddendumDir%\logs'n'data\ErrorLogs\Errorbox.txt, UTF-8
					}
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

	IfWinActive, ahk_class OptoAppClass
	{
            		If !(PraxomatHelp = "")
					{
                            	ControlGetText, AlbisKeyHelp, Afx:00400000:0:00010005:00000010:000000001
                            	If !Instr(AlbisKeyHelp, PraxomatHelp)
                            		ControlSetText, Afx:00400000:0:00010005:00000010:000000001, % AlbisKeyHelp "  " PraxomatHelp
            		}
					  else if !(AddendumHelp = "")
					{
                            	static AlbisKeyHelp_old
                            	ControlGetText, AlbisKeyHelp, Afx:00400000:0:00010005:00000010:000000001
                            	If !InStr(AlbisKeyHelp, AddendumHelp)
                            			ControlSetText, Afx:00400000:0:00010005:00000010:000000001, % AlbisKeyHelp "  " AddendumHelp
            		}
	}

}
;05
AlbisIsBlocked(AlbisWinID, autoclose:=2) {	                                                            	;-- befreit Albis von blockierenden Fenstern

	; stellt fest ob Albis durch ein ChildWindow blockiert ist und wie dieses heißt, kann dieses auch sofort schließen, autoclose ist default

/* 		DESCRIPTION

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

		If (autoclose = 0)
		{
				WinGetTitle, ltitle, % "ahk_id " phwnd:= DLLCall("GetLastActivePopup", "uint", AlbisWinID:= AlbisWinID())
				return {"isblocked": WinIsBlocked(AlbisWinID()), "blockWinT": ltitle, "errorStr": errorStr}
		}
		else if (autoclose = 1)
		{
				;platzhalter erstmal
		}
		else if (autoclose = 2)
		{
				If WinIsBlocked( AlbisWinID:= AlbisWinID() ) {
					AlbisClosePopups()
				}
				return
		}

return
}

AlbisClosePopups() {                                                                                               	;-- schließt alle PopUp-Fenster in Albis

		AlbisActivate(1)
		AlbisWinID := AlbisWinID()

		Loop
		{
				; welches Fenster blockiert Albis? ---
					phwnd		:= GetHex(DllCall("GetLastActivePopup", "uint", AlbisWinID()))
					t.= "hwnd" A_Index ": " phwnd ", AlbisWinID: " AlbisWinID()  "`n"

					If InStr(phwnd, AlbisWinID())
						break

				; 3 verschiedene Versuche das Fenster zu schließen von "sanft" bis "kraftvoll"
					SendMessage, 0x112, 0xF060,,, % "ahk_id " phwnd  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
					t.= " EL1: " ErrorLevel "`n"
					If ErrorLevel
					{
                            WinClose, % "ahk_id " phwnd
							t.= " EL2: " ErrorLevel "`n"
                            If ErrorLevel
                            {
                            		WinKill, % "ahk_id " phwnd
									t.= " EL3: " ErrorLevel "`n"
                            		If ErrorLevel
                            		{
                            				WinGetTitle, ltitle, % "ahk_id " phwnd
                            				throw Exception("Can not close entire window : '" . ltitle . "' `n. The function AlbisCloseLastActivePopups stops here.", -1)
                            				errorStr:="noClose"
                            				break
                            		}
                            }
					}

					sleep 300
				; Loop endet hier wenn Albis nicht mehr blockiert ist
					If !WinIsBlocked( AlbisWinID:= AlbisWinID() )
                            break
		}

		;ToolTip, % t, 200, 1, 17

return
}

WinIsBlocked(hwnd) {
 ; WS_DISABLED:= 0x8000000
	WinGet, Style, Style, % "ahk_id " hwnd
return (Style & 0x8000000) ? true : false
}
;06
AlbisCloseLastActivePopups(AlbisWinID) {                                                               	;-- schließt alle PopUp-Fenster bis keines mehr da ist

	;diese Funktion wird benötigt wenn ein Skript

	If !WinIsBlocked(AlbisWinID:= AlbisWinID())
		return

	Loop
	{
			phwnd:= GetHex(DLLCall("GetLastActivePopup", "uint", AlbisWinID()))

		; ist fertig wenn Popup handle und Albis handle identisch sind
			If (pHwnd = AlbisWinID())
					return 1

		;3 verschiedene Versuche das Fenster zu schließen von "sanft" bis "kraftvoll"
			SendMessage, 0x112, 0xF060,,, % "ahk_id " phwnd  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
			ELevel:= ErrorLevel

		; schießt manchmal versehentlich über das Ziel hinaus und will Albis schließen, der "Hiermit beenden Sie Albis" Dialog wird erkannt und
		; wieder geschlossen, die Schleife beendet
			If WinExist("ALBIS", "Hiermit beenden Sie Albis")
			{
					VerifiedClick("Button2", "ALBIS", "Hiermit beenden Sie Albis")
					return 1
			}

		; so neuer Versuch
			If ELevel
			{
					WinClose, % "ahk_id " phwnd
					If ErrorLevel
					{
                            WinKill, % "ahk_id " phwnd
                            If ErrorLevel
                            {
                            	WinGetTitle, ltitle, ahk_id %pHwnd%
                            	throw Exception("Can not close entire window : '" . ltitle . "' `n. The function AlbisCloseLastActivePopups stops here.", -1)
                            	return "Das Schließen geöffneter Dialoge war nicht möglich!"
                            }
					}
			}

			sleep 200

		; mehr als 10 zu schließende Dialoge gibt es nicht (Sicherheits return , damit die Schleife irgendwann verlassen wird)
			If A_Index = 10
				return 2   ; 2 = Problem beim Schliessen
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

		; letzte Änderung: 12.09.2019 - der Klick auf den OK Button wird jetzt per Mausklick-Simulation durchgeführt,
		; der Click mittels einer anderen Technik hat das Login-Fenster nicht geschlossen und den Hook für die Erkennung des Login-Fenster erneut ausgelöst

		static LoginTime	:= 0
		static LoginNo   	:= false

	;wenn ein automatischer Login verneint kann der nächste automatische Loginvorgang erst in 60 Sekunden erfolgen
		If LoginNo && ((A_TickCount - LoginTime) > 60000)
					LoginNo := false

	;Routine zum Einloggen, nach drücken auf Nein erkennt der Hook das Loginfenster erneut, deswegen muss eine erneute Abfrage unterdrückt werden
		If !LoginNo && WinExist("ALBIS - Login ahk_class #32770")
		{
					MsgBox, 4, Addendum für AlbisOnWindows, Soll das automatische Einloggen`ndurchgeführt werden?, % MsgBoxZeit
					IfMsgBox, No
					{
							LoginNo 	:= true
							LoginTime := A_TickCount
							return
					}

					If VerifiedSetText("Edit1", AlbisDefaultLogin("User"), "ALBIS - Login ahk_class #32770", 200)
						If VerifiedSetText("Edit2", AlbisDefaultLogin("Password"), "ALBIS - Login ahk_class #32770", 200)
						{
                            	hcontrol:= ControlGet("HWND", "", "Button1", "ALBIS - Login ahk_class #32770")
                            	c:= GetWindowSpot(hcontrol)
                            	MouseMove, % c.X +5 , % c.Y + 5
                            	MouseClick, Left
						}

					If hcontrol:= ControlGet("HWND", "", "Button1", "ALBIS - Login ahk_class #32770")
					{
                            c:= GetWindowSpot(hcontrol)
                            MouseMove, % c.X +5 , % c.Y + 5
                            MouseClick, Left
					}

					AlbisOeffneWarteZimmer()
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

	ListLines, Off
	; durch Verwendung von A_Username müsste diese Funktion auf allen Computern unter allen Benutzernamen funktionieren

	for Prozess in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	   prozesse.= Prozess.Name . " "

	Sort, prozesse, U

	If !Instr(prozesse, "AISConnector") {

		ErrorBox("Der AIS Connector ist nicht mehr verfügbar und wird neu gestartet!", "", 0)
		TrayTip, Hinweis!, Der AIS Connector ist`nnicht mehr verfügbar`nund wird neu gestartet!, 6
			sleep, 6000
		run C:\Users\%A_UserName%\Desktop\AIS Connector.appref-ms
		If ErrorLevel
			ErrorBox("Der AIS Connector ließ sich nicht starten!", "", 0)
			TrayTip, Fehler!, Der AIS Connector `nließ sich nicht starten!`nVersuche es später nocheinmal!, 6
				sleep, 6000
	}


}

;}
;----------------------------------------------------- Unterfunktionen der Albisfunktionen (kein manipulativer Zugriff auf die Albisoberfäche!) -------------------------------------------------------
; GetQuartal                                    	WMIEnumProcesses                    	WMIEnumProcessID                    	WMIEnumProcessExist                	ExtractFromString
; --WartezeitInfo--                            	--CheckAlbisHealth--                 	ExtractNamesFromString            	SleepOnRdPSession
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{
;1
GetQuartal(Datum, Format:="") {                                                                           	;--zum Errechnen zu welchem Quartal das übergeben Datum gehört

	; Funktionsbeschreibung:
	; Datum im folgenden Format ist erlaubt: 13.02.2017 oder 13.2.17 - in dieser Form der Übergabe müssen die Punkte im Übergabestring vorhanden sein
	; 14.5.2018: absofort ist die Übergabe des Begriffs heute möglich, es wird damit das aktuelle Quartal berechnet
	; Format ist ein Trennzeichen zwischen Quartal und Jahr z.B. Format="`/" dann wäre die Ausgabe so z.B. 01/18 , Format kann jedes beliebige Zeichen oder mehrere enthalten

	If InStr(Datum, "heute") {
		Monat	:= A_MM
		Jahr		:= SubStr(A_Year, 3, 2) ;die letzten zwei Zeichen
	} else {
		split		:= StrSplit(Datum, "`.")
		Monat	:= split[2]
		Jahr		:= Substr(split[3], StrLen(split3) - 1, 2)
	}

	;MsgBox, % Monat ", " Jahr "`n" SubStr("0" . Ceil(Monat/3), -1) . Format . Jahr
return SubStr("0" . Ceil(Monat/3), -1) . Format . Jahr
}
;2
WMIEnumProcesses() {                                                                                          	;--returns process list
	;==== ein deutlich besserer Ersatz für Process, Exist nur 32bit ======
	;funktioniert zuverlässiger, in der Original Library waren zudem Fehler bei den ID Bezeichnungen
	; gibt eine Liste existierender Prozesse samt PID zurück, die Liste muss nachfolgend ausgewertet werden

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
        list .= process.ProcessID . "|" . process.Name . "`n"

    return list
}
;3
WMIEnumProcessID(ProcessSearched) {

	ProcList:= Object()
	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		   ProcList[process.Name]:= process.ProcessID

	return ProcList.ProcessSearched
}
;4
WMIEnumProcessExist(ProcessSearched) {                                                            		;-- logische Funktion zur Suche ob ein bestimmter Prozeß existiert

    ;Funktion gibt nur 0 - für nicht vorhanden und 1 für vorhanden aus (als schnell Variante gedacht)
	;zudem ist der Funktion die exakte Schreibweise des Prozeß egal (GROSS/klein oder nur ein Teil davon)

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		   liste .= process.Name . "`n"

	Sort, liste, U

	Loop, parse, liste, `n
		If Instr(A_LoopField, ProcessSearched, false ,1)
			return 1

	return 0
}
;6
WartezeitInfo(command) {                                                                                    		;-- können Ausgaben von Funktionen unter den Skripten gemeinsam genutzt werden?



}
;7
CheckAlbisHealth() {                                                                                                	;-- eine Funktion die helfen soll durch Albis ausgelöste Verzögerungen zu erkennen, die nicht im Zusammenhang mit einem Programmabsturz stehen


}
;8
ExtractNamesFromString(Str) {                                                                        			;-- die Namen müssen die ersten zwei Worte im String sein

	Name:=[]
	;diese Funktion basiert auf der manuellen Erstellung des Dateinamens während des Scanvorganges, es ist sehr kompliziert dem Computer die Erkennung von Eigennamen beizubringen

	;Ermitteln des Patientennamen aus dem PDF-Dateinamen
	Pat:= StrReplace(Str, "`,", A_Space)
	;zuviele Space Zeichen hintereinander machen Probleme
	Pat:= StrReplace(Pat, A_Space . A_Space . A_Space, A_Space)
	Pat:= StrReplace(Pat, A_Space . A_Space, A_Space)
	Pat:= StrSplit(Pat, A_Space)
	Name[1]:= Pat[1] . ", " . Pat[2]
	Name[2]:= Pat[2] . ", " . Pat[1]
	Name[3]:= Pat[1]
	Name[4]:= Pat[2]

return Name
}
;9
SleepOnRdPSession(ms, altms) {                                                                               	;-- unterschiedliche sleep Zeiten ms = wenn eine Remotedesktop-Verbindung besteht, altms = wenn nicht

	; manchmal notwendig wegen verzögertem Übertragung einer Fensterdarstellung während einer bestehenden Remotesitzung. Manche Fenster flickern lange und man kann nichts lesen.
	; wenn altms = 0 kehrt die Funktion sofort zurück

	If altms = 0
			return

	If WMIEnumProcessExist("RdClient.Windows.exe")
			sleep % ms
	else
			sleep % altms

return
}
;}


