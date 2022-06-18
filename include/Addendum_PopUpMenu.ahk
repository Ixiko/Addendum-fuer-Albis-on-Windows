; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                                            	Addendum_PopUpMenu
;
; 		Verwendung:        -	eigene Menupunkte in Rechtsklickmenu's in Albis einblenden und verwenden
;                                	- 	Karteikarte: blendet weitere Menupünkte in Abhängigkeit zum aktuellen Kürzel ein, z.B. 'scan' - Drucken, Exportieren,
;                                		- exportiert wird in einen eigenen Ordner mit Patientennamen z.B. für komfortablen Aktenexport
;										- fehlbezeichnete Dokument, z.B. ist es das Dokument eines anderen Patienten lassen sich in den Befundordner nach
;										  Abfrage des geänderten Patientennamen exportieren
;                                	-	oder ein Rezept/Formular nochmal drucken ohne manuellen Eingriff
;
;		Basisskript:        	-	Addendum.ahk
;
;       Abhängigkeiten:   	-	Addendum_Menu.ahk, Addendum_Window.ahk, Addendum_Albis.ahk, Addendum_PDFReader.ahk
;                                	-	benötigt wird außerdem ein ShellHook-Callback, welcher im Addendum.ahk Skript dafür angelegt ist
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;
;       Addendum_PopUpMenu started: 20.06.2020 | last change: 09.02.2022
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Addendum_PopUpMenu(hWin, hMenu) {                                                                               	;-- fügt neue Menupunkte dem Rechtsklickmenu der Karteikarte hinzu

		static ExtMenu, cmdGroups

		global pm 	:= Object()
		pm.menu  	:= Array()

	; initialisiert die Menubefehle für jedes Karteikartenkürzel
		If !IsObject(ExtMenu) {

			; Befehle und Menutexte
				const	 := {	1: {"cmd":"fax",      	"menutext":"als Fax versenden"}
							,	2: {"cmd":"mail",    	"menutext":"per Mail versenden"}
							,  	3: {"cmd":"tgram",    	"menutext":"per Telegram versenden"}
							,	4: {"cmd":"show",    	"menutext":"Anzeigen"}
							, 	5: {"cmd":"print",     	"menutext":"Drucken"}
							, 	6: {"cmd":"export1",	"menutext":"Exportieren"}
							, 	7: {"cmd":"export2",	"menutext":"Exportieren als anderer Patient"}
							, 	8: {"cmd":"open",    	"menutext":"Bearbeiten"} }

			; bei Karteikartenkürzel die ein Leerzeichen enthalten, ist dieses durch ein # zu ersetzen
				ExtMenu := {"scan"    	: [1, 2, 3, 4, 5, 6, 7]
								, 	"epi"      	: [1, 2, 3, 7, 5, 6]
								, 	"besch"   	: [1, 2, 3, 7, 5, 6]
								, 	"medrp"  	: [7, 5]
								, 	"medh"   	: [7, 5]
								, 	"medpn"	: [7, 5]
								, 	"fau"      	: [7, 5]
								, 	"fhv13"  	: [7, 5]
								, 	"fkh"      	: [7, 5]
								, 	"füb"     	: [7, 5]}

			; ExtMenu wird für die Verarbeitung verändert
				For key, arr in ExtMenu	{
					tmpArr := Array()
					For idx, val in arr
						tmpArr.Push(const[val])
					ExtMenu[key] := tmpArr
				}

		}

	; Maus Position festhalten, Menu wird durch das Anfügen weiterer Punkte länger und bei Erreichen des unteren Bildschirmrandes nach oben verschoben
	; die eigentliche Klickposition ist dann nicht mehr feststellbar
		CoordMode	,	Mouse, Screen
		MouseGetPos,	MouseScreenX, MouseScreenY
		pm.MouseX :=	MouseScreenX
		pm.MouseY :=	MouseScreenY

	; Kontextmenu - alle Menupunkte ermitteln
		MenuGetAll_sub(hmenu, "", Lcmds)

	; zusätzliche Menupunkte werden individuell erstellt
		If RegExMatch(Lcmds, "i)filter\:\s*(\w+)", filter) {                                  	; Rechtsklick Menu der Karteikarte erkannt

				If !ExtMenu.haskey(filter1) || (StrLen(filter1) = 0)
						return

				Addendum_MenuAppendSeparator(hWin, hMenu)                       	; Menu-Separator anfügen

				pmPosO   	:= GetWindowSpot(hWin)                                      	; Position des Menu vor der Veränderung
				pm.MenuX	:= pmPosO.X
				pm.MenuY	:= pmPosO.Y

				Addendum_MenuAppendLabels(hWin, hMenu, ExtMenu[filter1])		; Menupunkte hinzufügen

				pmPosN    	:= GetWindowSpot(hWin)                                      	; Position des Menu nach der Veränderung
				newY        	:= pmPosO.Y + pmPosO.H                                  	; Y-Position des angefügten Menu's
				H              	:= Floor((pmPosN.H - pmPosO.H) / ExtMenu[filter1].MaxIndex())
				PatID        	:= AlbisAktuellePatID()

				Loop % ExtMenu[filter1].MaxIndex() 	{                                      	; stellt die zusätzlichen Menupunkte zusammen

						cmd := ExtMenu[filter1][A_Index].cmd
						If InStr(cmd, "tgram") && !oPat[PatID].tgChatID
							continue
						else if InStr(cmd, "fax") && (InStr(Addendum.Drucker.FAX, "ERROR") || StrLen(Addendum.Drucker.FAX) = 0)
							continue

						pm.menu.Push({ 	"filter"	: filter1                                         	; Karteikartenkürzel der Auswahl
								        		, 	"cmd"	: cmd                                           	; auszuführender Befehl
								        		,	"StartY"	: (newY+H*(A_Index-1))             	; y1 Position des Menupunktes
								        		, 	"EndY"	: (newY+H*(A_Index)-1)})           	; y2 Position des Menupunktes

				}

				return pm

		}

return
}

Addendum_PopUpMenuItem(pm) {                                                                                         	;-- wertet den Nutzerklick aus

	; Menu-Events in Prozessen welche nicht zum eigenen Programm/Skript gehören lassen sich systembedingt nur mittels "Code-Injection" abfangen.
	; Dafür "injiziert" man fremden Code (z.B. eine DLL) in einen anderen Prozess.  Mit Autohotkey läßt sich dies mittels Boardmittel nicht bewerkstelligen.
	; Diese Art von Code wird aus guten Gründen als Schadsoftware von Virenscanner bewertet. Da ich keinen Code verwenden möchte,
	; den man nicht überprüfen kann, verwende ich eine einfache Methode auf Basis der Position des Mauszeigers. Dies funktioniert in den allermeisten Fällen gut.
	;
	; cmdGroups:		- statische Variable, die verschiedene Aktenkürzel zu einer Aktionsausführung zusammenfasst
	;
	; Parameter: 	pm	- ist ein Objekt mit Daten zu jedem angefügten Menupunkt. Gespeichert ist auch die relative y Position innerhalb des Menu, so
	;                         	  daß mit einfachen Bestimmung der Position des Mauszeiger auf dem Bilderschirm der ausgewählte Menupunkt errechnet werden kann

	; Kürzel und Befehlsgruppierung
	static cmdGroups := {	1:{	"group" 	: "medrp,medh,medpn,medbm,fau,fhv13,fhv18,fkh,füb"
								        	, 	"WinTitle"	: "Muster ahk_class #32770"
							        		, 	"open"   	: "{F3}"
							        		, 	"print"    	: "Drucken"}}
					    		;~ , 	2:{	"group" 	: "fau"
						    			;~ ,	"WinTitle"	: ""}}

	static DiaUebernahme := "Diagnosen auf Schein ahk_class #32770"

		CoordMode, Mouse, Screen
		MouseGetPos, MouseScreenX, MouseScreenY  ; Int vars: 4 octets

	; dies ist der Trick ohne Prozess-Injection doch die angeklickte Menuposition zu erhalten
		For idx, menu in pm.menu
			If (MouseScreenY >= menu.StartY) && (MouseScreenY <= menu.EndY) {
				cmd := menu.cmd, filter := menu.filter
				break
			}
	 ; command und filter sind leer, dann nichts machen
		If (StrLen(cmd . filter) = 0)
				return

	; für Formulare braucht es keine weiteren Funktionen
		For idx, obj in cmdGroups		{

			thisgroup := obj.group
			If filter in %thisgroup%
			{
					If InStr(cmd, "open") {                    ; Formular bearbeiten - Öffnen per SendInput
						SendInput, % obj.open
						return
					}
					else if InStr(cmd, "print") {	        	; Formular drucken - Öffnen und simulierter Mausklick auf Drucken
						SendInput, % obj.open
						WinWait, % obj.WinTitle,, 6
						VerifiedClick(obj.print, obj.WinTitle,,, true)
						WinWait, % DiaUebernahme,, 3
						If WinExist(DiaUebernahme)
							VerifiedClick("Abbruch", DiaUebernahme,,, true)
						return

					}
			}
		}

		If IsFunc("admMenu_" filter)
			admMenu_%filter%(cmd, pm)

}

; ~~~~~ Karteikartenkürzel Funktionen ~~~~~~
admMenu_scan(cmd, pm) {                                                                                                   	;-- wird beim Aktenkürzel scan aufgerufen

	; Variablen zur Identifikation von Fenstern
	static foxitPrint := { "Dialog"	:  "i)Print|Drucken ahk_class i)classFoxitReader ahk_exe i)FoxitReader"
							,	 "Control"	: {"Properties"            	: "Button1"
												,  	"Advanced"            	: "Button42"
												,	"Cancel"                	: "Button45"
												,	"Copies"                 	: "Edit1"
												,	"Ok"                      	: "Button44"
												,	"Pages_All"            	: "Button15"
												,	"Pages_Range"       	: "Button16"
												,	"Pages_Range_Edit"	: "Edit4"
												,	"Printers"               	: "ComboBox1"}}
	
	; liest zunächst die Karteikartenzeile aus
	If RegExMatch(cmd, "i)^(export1|export2|print|fax)$") 	{                                                       
	
		AlbisActivate(1)
		BlockInput, On                                                                                                            	; Nutzerinteraktion verhindern
		res := AlbisLeseDatumUndBezeichnung(pm.MouseX, pm.MouseY)                                 	; holt sich das Zeilendatum und den eingetragenen Text
		SendInput, {Escape}       
		If !IsObject(res) {                                                                                                        	; Fehlerbehandlung
			BlockInput, Off
			PraxTT("Der Karteikartentext konnte nicht ermittelt werden.`nDer Dokumentdruck ist fehlgeschlagen", "3 0")
			return
		}
		
		; Patientendaten zusammenstellen
			PatNamePath := Addendum.ExportOrdner "\" "(" 
								. 	 (PatID      	:= AlbisAktuellePatID()) ") " 
								. 	 (PatName	:= AlbisCurrentPatient())
			
		; Karteikartentext von Patientennamen und anderem befreien
			PdfTitle := RegExReplace(res.Text	, "i)" StrSplit(PatName, ", ").1 ".*?" StrSplit(PatName, ", ").2)
			PdfTitle := RegExReplace(PdfTitle 	, "i)" StrSplit(PatName, ", ").2 ".*?" StrSplit(PatName, ", ").1)
			PdfTitle := RegExReplace(PdfTitle 	, "([a-zäöüß])([A-ZÄÖÜ])", "$1 $2" )
			PdfTitle := RegExReplace(PdfTitle 	, "\.pdf")
			
		; Jahr.Monat.Tag - Befunde lassen sich im Explorer nach Datum sortieren
			KKDatum 	:= ConvertToDBASEDate(res.Datum)                                 
		
	}

	If InStr(cmd, "show")                                                   	{
		SendInput, {F3}
		return
	}
	else if InStr(cmd, "export1")                                          	{                                                   	; Exportieren in den Versandordner

		/* Beschreibung "export1"

			- 	holt sich das Datum und die Bezeichnung des Befundes aus dem Eintrag in der Albis Patientenkarteikarte
			- 	erstellt hieraus aus einen zusammengesetzten Dateinamen
			- 	erstellt einen mit Patienten ID und Namen versehenen Ordner in dem in der Addendum.ini hinterlegten Export-Ordner
			- 	öffnet das zu sende Dokument mittels Tastatureingabe (senden der Taste F3)
			- 	per ShellHook(Addendum.ahk Skript) wird ein Callback zur Funktion admMenu_SaveAs() oder admMenu_print() ausgeführt.
				wird das Dokument in dem in Albis vom Nutzer hinterlegten PDF-Anzeigeprogramm anzeigt (der Shellhook regiert nur auf den Foxitreader
				und Sumatra) wird eine der beiden oberen Funktionen aufgerufen

		*/

		; Export Ordner anlegen für diesen Patienten falls nicht vorhanden
			If !InStr(FileExist(PatNamePath), "D")
				FileCreateDir % PatNamePath

		; anderen Dateinamen geben, falls der Name schon vergeben wurde
			PdfFullFilePath := PatNamePath "\" KKDatum " - " PdfTitle ".pdf"                                         		 	
			while FileExist(PdfFullFilePath)                                                                                                       	; sucht einen noch nicht vorhandenen Dateinamen
				PdfFullFilePath := RegExReplace(PdfFullFilePath, "\(*\d*\)*\.pdf$", "(" A_Index ").pdf")

		; Taste F3 an Albis senden um PDF-Datei mit dem in Albis hinterlegten PDF-Reader anzuzeigen
		; die Shellhook-Prozedure im Addendum Hauptskript ruft die als Parameter übergebene Funktion auf
			Addendum.PopUpMenuCallback := "admMenu_PdfSaveAs|" PdfFullFilePath
			SendInput, {F3}
			BlockInput, Off

	}
	else if InStr(cmd, "export2")                                          	{                                                     	; Exportieren eines unter falschem Patienten abgelegten Dokumentes

		/* 	Beschreibung export2

			- 	Funktionsweise in etwa wie bei "export1" beschrieben
			- 	Dokument wird aber in den BefundOrdner gespeichert
			- 	der Karteikarteneintrag wird nicht automatisch entfernt
			- 	das mit neuem Namen versehene Dokument wird nicht automatisch in eine andere Karteikarte importiert!

		*/
		
		; Patientennamen abfragen (Zeilimit der Inputbox - falls Nutzer den Dialog nicht beachtet)
			BlockInput, Off
			prompt :=	"  DOKUMENT UNTER ANDEREM PATIENTENNAMEN SPEICHERN`n`n"
			.	"Geben Sie dem Dokument den zugehörigen Namen des Patienten.`n"
        	. 	"benutzen Sie folgende Schreibweise: Nachname, Vorname"
			fnAutoIB := Func("AutoSizeInputBox")
			SetTimer, % fnAutoIB, -100
			InputBox, PatName, Addendum für Albis on Windows, % prompt,, 420, 186,,, Locale, 60, % "Mustermann, Max"
			If ErrorLevel                                                                                                               	; Zeitlimit erreicht oder 'Abbrechen' wurde gedrückt
				return

		; Mauspfeil auf den Karteikarteneintrag bewegen
			BlockInput, On
			AlbisActivate(1)
			MouseClick, Left, pm.MouseX, pm.MouseY                                                                   	; ein Mausklick über der alten Position selektiert den ursprünglichen Eintrag
			Sleep, 200
			SendInput, {Escape}                                                                                                   	; nur einfaches Auswählen (blau hinterlegter Eintrag)
			sleep, 200

		; gespeichert wird in den Ordner für Befundeingänge (Callback Funktionsaufruf in Addendum.ahk)
			Addendum.PopUpMenuCallback := "admMenu_PdfSaveAs|" Addendum.BefundOrdner "\" PatName ", " PdfTitle 
																. (!RegExMatch(PdfTitle, "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})") ? " vom " KKDatum : "") ".pdf"
			SendInput, {F3}
			BlockInput, Off

	}
	else if RegExMatch(cmd, "i)(print|fax)$")                         	{                                                     	; Drucken oder Faxversand einer PDF-Datei

		; der Callback Funktion wird Text für die Protokollierung und der Name des Standard-A4 Druckers übergeben
			Addendum.PopUpMenuCallback := "admMenu_PdfPrint|(Karteikartendatum: " res.Datum ", Dokumentname: " res.Text ")"
																. "#" (cmd = "print" ? Addendum.Drucker.StandardA4 : Addendum.Drucker.FAX)
			SendInput, {F3}
			BlockInput, Off                                                                                                            	; Nutzerinteraktion zulassen

	}

}

admMenu_medrp(cmd, MenuX, MenuY) {                                                                                	;-- wird ausgefügt beim Karteikartenkürzel "medrp"

	;AlbisGet
	rezept := "Muster 16 ahk_class #32770"

	SendInput, {F3}
	WinWait, % rezept,, 2
	while WinExist(rezept) {
		VerifiedClick("Drucken", rezept,,, true)
		If (A_Index > 5)
			break
		Sleep, 200
	}


}

; ~~~~~ Windows Menu Funktionen ~~~~~~~
Addendum_MenuAppendSeparator(hWin, hMenu) {                                                                	;-- fügt einen Separator ein
	DllCall("AppendMenu", "Ptr", hMenu, "UInt", 0x800	, "UPtr", 0, "Str", "")
	DllCall("DrawMenuBar", "Ptr", hWin)
}

Addendum_MenuAppendLabels(hWin, hMenu, labels) {                                                            	;-- fügt neue Menupunkte ein

	; labels - Array
		pm  		:= Array()
		PatID 	:= AlbisAktuellePatID()

	; Patient ist nicht bei Telegram, dann wird der Menupunkt entfernt
		If !oPat[PatID].tgChatID
			Loop % labels.MaxIndex()
				If InStr(labels[A_Index].cmd, "tgram") {
					labels.RemoveAt(A_Index)
					break
				}

	; Menupunkte werden dem Albis Kontextmenu hinzugefügt
		Loop % labels.MaxIndex()
			DllCall("AppendMenu", "Ptr", hMenu, "UInt", 0x0, "UPtr", (90000+A_Index-1), "Str", labels[A_Index].menutext)
		DllCall("DrawMenuBar", "Ptr", hWin)

	; virtuelles Menu wird erstellt
		pmPosN	:= GetWindowSpot(hWin)
		newY    	:= pmPosO.Y + pmPosO.H
		H           	:= Floor((pmPosN.H - pmPosO.H) / labels.MaxIndex())
		Loop % labels.MaxIndex()
			pm.Push({ 	"filter"	: filter1                                            	; Karteikartenkürzel der Auswahl
							, 	"cmd"	: labels[A_Index].cmd                       	; auszuführender Befehl
							,	"StartY"	: (newY+H*(A_Index-1))                 	; y1 Position des Menupunktes
							, 	"EndY"	: (newY+H*(A_Index)-1)})              	; y2 Position des Menupunktes

return pm
}

; ~~~~~ PDF Reader/Viewer Funktionen ~~~~~
admMenu_PdfSaveAs(PdfFullFilePath, PDFViewerClass, PDFViewerHwnd) {                               	;-- Callback Funktion für Befundexporte (_scan)

	; Hinweis
		PraxTT("PDF Datei wird exportiert.", "20 0")

	; Foxit/Sumatra (universelle Funktion)
		FExists := PdfSaveAs(PdfFullFilePath, PDFViewerClass, PDFViewerHwnd)

	; Albis wieder aktivieren
		AlbisActivate(2)

	; Hinweis über Erfolg oder Mißerfolg
		SplitPath, PdfFullFilePath, OutFileName, OutDir
		PraxTT(FExists 	? "(" FExists ") Der Befund wurde im Verzeichnis:`n" OutDir "`nunter dem Namen:`n" OutFileName "`ngespeichert." 
								: "Der Befund konnte nicht nach`n" OutDir "`nexportiert werden!.", "6 0")

}

admMenu_PdfPrint(Param, PDFViewerClass, PDFViewerHwnd) {                                               	;-- Callback Funktion für Pdf-Druck (_scan)

	; Funktion bedient den Druck auf einen Hardwaredrucker und auch für virtuelle Druckertreiber z.B. Fax

		PraxTT("Pdf Dokument wird gedruckt.`nBitte warten!", "20 3")

		PatID		:= AlbisAktuellePatID()
		LogText		:= StrSplit(Param, "#").1
		Printer   	:= StrSplit(Param, "#").2
		PdfPages	:= PdfPrint(Printer, PDFViewerClass, PDFViewerHwnd)
		LogText	 	.= ", Dokument mit " PdfPages.max " Seite(n) gedruckt"                      	; LogText präzisieren
		admGui_PatLog(PatID, "AddToLog", LogText)                                                  	; die ausgedruckte Datei inkl. Seitenzahl protokollieren

		PraxTT("", "off")                                                                                                	; Info-Anzeige aus
		AlbisActivate(2)                                                                                               	; zurück zu Albis

}

; ~~~~~ zusätzliche Funktionen ~~~~~~~~~
AutoSizeInputBox() {                                                                                                                	;-- eine Standardinputbox wird mittig innerhalb des Albisfenster zentriert

	local a, i, hwnd

	WinSet, AlwaysOnTop, On, % "ahk_id " (hwnd := WinExist("A"))
	a	:= GetWindowSpot(AlbisWinID())
	i	:= GetWindowSpot(hwnd)

	; soll mittig innerhalb des Albisfenster positioniert werden
	x 	:= a.X + Floor((a.CW - a.CX)/2) - Floor(i.W/2)
	y 	:= a.Y + Floor((a.CH - a.CY)/2) - Floor(i.H/2)
	SetWindowPos(hwnd, x, y, i.W, i.H)

}

