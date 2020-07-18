; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                                            	Addendum_PopUpMenu
;
; 		Verwendung:        -	eigene Menupunkte in Rechtsklickmenu's in Albis einblenden und verwenden
;                                	- 	Karteikarte: blendet weitere Menupünkte in Abhängigkeit zum aktuellen Kürzel ein, z.B. 'scan' - Drucken, Exportieren,
;                                		exportiert wird dann in einen eigenen Ordner mit Patientennamen für Aktenexport
;                                	-	oder ein Rezept/Formular nochmal drucken ohne manuellen Eingriff
;
;		Basisskript:        	-	Addendum.ahk
;       Abhängigkeiten:   	-	Addendum_Menu, Addendum_Window, Addendum_Albis
;                                	-	benötigt wird außerdem ein ShellHook-Callback, welcher im Addendum.ahk Skript dafür angelegt ist
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_PopUpMenu started:    	20.06.2020
;       Addendum_PopUpMenu last change:	16.07.2020
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Addendum_PopUpMenu(hWin, hMenu) {                                                                               	;-- fügt neue Menupunkte dem Rechtsklickmenu der Karteikarte hinzu

		static ExtMenu, cmdGroups
		pm := Object()
		pm.menu := Array()

	; initialisiert die Menubefehle für jedes Karteikartenkürzel
		If !IsObject(ExtMenu) {

			; Befehle und Menutexte
				const	 := {	1: {"cmd":"fax",  	"menutext":"als Fax versenden"}
							,	2: {"cmd":"mail",	"menutext":"per Mail versenden"}
							,  	3: {"cmd":"tgram",	"menutext":"per Telegram versenden"}
							,	4: {"cmd":"show",	"menutext":"Anzeigen"}
							, 	5: {"cmd":"print", 	"menutext":"Drucken"}
							, 	6: {"cmd":"export",	"menutext":"Exportieren"}
							, 	7: {"cmd":"open",	"menutext":"Bearbeiten"} }

			; bei Karteikartenkürzel die ein Leerzeichen enthalten, ist dieses durch ein # zu ersetzen
				ExtMenu := {"scan"    	: [1, 2, 3, 4, 5, 6]
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
				For key, arr in ExtMenu
				{
						tmpArr := Array()
						For idx, val in arr
							tmpArr.Push(const[val])

						ExtMenu[key] := tmpArr
				}

		}

	; Maus Position festhalten, Menu wird durch das Anfügen weiterer Punkte länger und bei Erreichen des unteren Bildschirmrandes nach oben verschoben
	; die eigentliche Klickposition ist dann nicht mehr feststellbar
		CoordMode, Mouse, Screen
		MouseGetPos, MouseScreenX, MouseScreenY
		pm.MouseX := MouseScreenX
		pm.MouseY := MouseScreenY

	; Kontextmenu - alle Menupunkte ermitteln
		MenuGetAll_sub(hmenu, "", Lcmds)

	; zusätzliche Menupunkte werden individuell erstellt
		If RegExMatch(Lcmds, "i)filter\:\s*(\w+)", filter) { ; Rechtsklick Menu der Karteikarte erkannt

				If !ExtMenu.haskey(filter1) || (StrLen(filter1) = 0)
						return

				Addendum_MenuAppendSeparator(hWin, hMenu)                       	; Menu-Separator anfügen

				pmPosO   	:= GetWindowSpot(hWin)                                      	; Position des Menu vor der Veränderung
				pm.MenuX	:= pmPosO.X
				pm.MenuY	:= pmPosO.Y

				Addendum_MenuAppendLabels(hWin, hMenu, ExtMenu[filter1])		; zusätzliche Menupunkte hinzufügen

				pmPosN    	:= GetWindowSpot(hWin)                                      	; Position des Menu nach der Veränderung
				newY        	:= pmPosO.Y + pmPosO.H                                  	; Y-Position des angefügten Menu's
				H              	:= Floor((pmPosN.H - pmPosO.H) / ExtMenu[filter1].MaxIndex())
				PatID        	:= AlbisAktuellePatID()

				Loop % ExtMenu[filter1].MaxIndex() 	{

						cmd := ExtMenu[filter1][A_Index].cmd
						If InStr(cmd, "tgram") && !oPat[PatID].tgChatID
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
	; Dafür "injiziert" man fremden Code (z.B. eine DLL) in einen anderen Prozess.  Mit Autohotkey läßt sich dies per Boardmittel nicht bewerkstelligen.
	; Diese Art von Code wird aus guten Gründen als Schadsoftware von Virenscanner bewertet. Da ich keinen Code verwenden möchte,
	; den man nicht überprüfen kann, verwende ich hier eine einfache Methode auf Basis der Position des Mauszeigers.
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

	For idx, menu in pm.menu
		If (MouseScreenY >= menu.StartY) && (MouseScreenY <= menu.EndY) {
			cmd := menu.cmd, filter := menu.filter
			break
		}

	If (StrLen(cmd . filter) = 0)                                                                                ; command und filter sind leer, dann nichts machen
			return

	For idx, obj in cmdGroups
	{
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

; Karteikartenkürzel Funktionen
admMenu_scan(cmd, pm) {

	; Variablen zur Identifikation von Fenstern
	static foxitPrint := { "Dialog"	:  "i)Print|Drucken ahk_class i)classFoxitReader ahk_exe i)FoxitReader"
							,	 "Control"	: {"Properties"            	: "Button1"
												,  	"Advanced"           	: "Button42"
												,	"Cancel"                	: "Button45"
												,	"Copies"                 	: "Edit1"
												,	"Ok"                     	: "Button44"
												,	"Pages_All"            	: "Button15"
												,	"Pages_Range"      	: "Button16"
												,	"Pages_Range_Edit"	: "Edit4"
												,	"Printers"               	: "ComboBox1"}}

	KKDatum := ""

	If InStr(cmd, "show") {

			SendInput, {F3}

	}
	else if InStr(cmd, "export") {

		/* Beschreibung

				- holt sich das Datum und die Bezeichnung des Befundes aus dem Eintrag in der Albis Patientenkarteikarte
				- erstellt hieraus aus einen zusammengesetzten Dateinamen
				- erstellt einen mit Patienten ID und Namen versehenen Ordner in dem in der Addendum.ini hinterlegten Export-Ordner

		*/

		; Patientendaten zusammenstellen
			PatID            	:= AlbisAktuellePatID()
			PatName      	:= StrReplace(AlbisCurrentPatient(), ", ", ",")
			PatNamePath	:= Addendum.ExportOrdner "\" "(" PatID ")" PatName

		; Export Ordner anlegen für diesen Patienten falls nicht vorhanden
			If !(FileExist(PatNamePath) = "D")
				FileCreateDir % PatNamePath

		; Pdf Benennung aus der Karteikarte holen
			BlockInput, On                                                                                                            	; Nutzerinteraktion verhindern
			AlbisActivate(1)
			result := AlbisLeseDatumUndBezeichnung(pm.MouseX, pm.MouseY)
			If !IsObject(result) {
				BlockInput, Off
				PraxTT("Der Karteikartentext konnte nicht ermittelt werden.`nDer Befundexport ist fehlgeschlagen", "3 0")
				SendInput, {Escape}
				return
			}
			PdfTitel  	:= result.Text                                                                                              	; eventuell noch vorhandene pdf Endung aus dem Karteikartentext entfernen
			KKDatum	:= result.Datum
			KKDatum 	:= SubStr(KKDatum, 7, 4) "." SubStr(KKDatum, 4, 2) "." SubStr(KKDatum, 1, 2)	; Jahr.Monat.Tag - Befunde lassen sich im Explorer nach Datum sortieren
			SendInput, {Escape}                                                                                                   	; Zeile wieder freigeben

		; anderen Dateinamen geben, falls der Name schon vergeben wurde
			PdfFullFilePath := PatNamePath "\" KKDatum "-" PdfTitel ".pdf"
			while % FileExist(PdfFullFilePath) {
				If RegExMatch(PdfFullFilePath, "\(\d+\)\.pdf")
					PdfFullFilePath := RegExReplace(PdfFullFilePath, "\(\d+\)\.pdf$", "(" A_Index ").pdf")
				else
					PdfFullFilePath := RegExReplace(PdfFullFilePath, "\.pdf$", "(" A_Index ").pdf")
			}

		; Taste F3 an Albis senden um PDF-Datei mit dem in Albis hinterlegten PDF-Reader anzuzeigen
		; die Shellhook-Prozedure im Addendum Hauptskript ruft die als Parameter übergebene Funktion auf
			Addendum.PopUpMenuCallback := "admMenu_PdfSaveAs|" PdfFullFilePath
			SendInput, {F3}
			BlockInput, Off

	} else if InStr(cmd, "print") {

			BlockInput, On                                                                                                            	; Nutzerinteraktion verhindern
			result := AlbisLeseDatumUndBezeichnung(pm.MouseX, pm.MouseY)
			If !IsObject(result)
				return
			Addendum.PopUpMenuCallback := "admMenu_PdfPrint|(Karteikartendatum: " result.Datum ", Dokumentname: " result.Text ")"
			SendInput, {F3}
			BlockInput, Off                                                                                                            	; Nutzerinteraktion zulassen

	}

}

admMenu_medrp(cmd, MenuX, MenuY) {

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

; Windows Menu Funktionen
Addendum_MenuAppendSeparator(hWin, hMenu) {

	DllCall("AppendMenu", "Ptr", hMenu, "UInt", 0x800	, "UPtr", 0, "Str", "")
	DllCall("DrawMenuBar", "Ptr", hWin)

}

Addendum_MenuAppendLabels(hWin, hMenu, labels) {

	; labels - Array

		pm  	:= Array()
		PatID := AlbisAktuellePatID()

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

; PDF Reader Funktionen
admMenu_PdfSaveAs(PdfFullFilePath	, PDFViewerClass, PDFViewerHwnd) {                               	;-- Callback Funktion für Befundexporte

		static foxitSaveAs    	:= "Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe"
		static sumatraSaveAs	:= "Speichern unter ahk_class #32770 ahk_exe SumatraPDF.exe"

	; Pfad und Dateiname
		SplitPath, PdfFullFilePath, OutFileName, OutDir

	; verschiedene RPA-Routinen je nach benutztem PDFReader (im Moment FoxitReader und
		If InStr(PDFViewerClass, "Foxit") {                ; FoxitReader

			FoxitInvoke("SaveAs", PDFViewerHwnd)                    	; 'Speichern unter' - Dialog
			WinWait, % foxitSaveAs,, 6                                       	; wartet 6 Sekunden auf das Dialogfenster
			hfoxitSaveAs := GetHex(WinExist(foxitSaveAs))        	; 'Speichern unter' - Dialog handle
			VerifiedSetText("Edit1", PdfFullFilePath , hfoxitSaveAs)	; Speicherpfad eingeben
			VerifiedClick("Button3", hfoxitSaveAs,,, true)               	; Speichern Button drücken
			FoxitInvoke("Close", PDFViewerHwnd)                        	; dieses Dokument schließen
			AlbisActivate(2)                                                        	; Albis wieder aktivieren

		}
		else If InStr(PDFViewerClass, "Sumatra") {	; Sumatra PdfReader

			SumatraInvoke("SaveAs", PDFViewerHwnd)                	; 'Speichern unter' - Dialog
			WinWait, % sumatraSaveAs,, 6                                  	; wartet 6 Sekunden auf das Dialogfenster
			hSaveAs := GetHex(WinExist(sumatraSaveAs))        	; 'Speichern unter' - Dialog handle
			VerifiedSetText("Edit1", PdfFullFilePath , hSaveAs)     	; Speicherpfad eingeben
			VerifiedClick("Button2", hSaveAs,,, true)                  	; Speichern Button drücken
			SumatraInvoke("Close", PDFViewerHwnd)                  	; dieses Dokument schließen
			AlbisActivate(2)                                                        	; Albis wieder aktivieren

		}

	; Wartet bis Datei fertig gespeichert wurde
		while !(FExists := FileExist(PdfFullFilePath)) {
			Sleep, 200
			If (A_Index > 40)
				break
		}

	; Hinweis über Erfolg oder Mißerfolg
		If FExists
			PraxTT("(" FExists ") Der Befund wurde im Verzeichnis:`n" OutDir "`nunter dem Namen:`n" OutFileName "`ngespeichert." , "6 0")
		else
			PraxTT("Der Befund konnte nicht exportiert werden!. `nOutput: " FExists , "6 0")

}

admMenu_PdfPrint(LogText, PDFViewerClass, PDFViewerHwnd) {                                               	;-- Callback Funktion für Pdf-Druck

		static foxitprint    	:= "i)[(Print)|(Drucken)]+ ahk_class i)#32770 ahk_exe i)FoxitReader.exe"
		static sumatraprint	:= "Speichern unter ahk_class #32770 ahk_exe SumatraPDF.exe"

		SciTEOutput("PdfPrint: " LogText " ," PDFViewerClass ", " PDFViewerHwnd )

		If InStr(PDFViewerClass, "Foxit") {                ; FoxitReader

			PatID	 := AlbisAktuellePatID()
			PdfPage := FoxitReader_GetPages(PDFViewerHwnd)
			admGui_PatLog(PatID, "AddToLog", LogText)

			OldMatchMode := A_TitleMatchMode
			SetTitleMatchMode, RegEx
			FoxitInvoke("Print", PDFViewerHwnd)                       	; 'Drucken' - Dialog
			WinWait, % foxitPrint,, 6                                          	; wartet 6 Sekunden auf das Dialogfenster
			hfoxitPrint := GetHex(WinExist(foxitPrint))                	; 'Speichern unter' - Dialog handle
			SciTEOutput("A4Drucker: " Addendum.Drucker.StandardA4 ", hfoxitPrint: " hfoxitPrint)
			VerifiedChoose("ComboBox1", hfoxitPrint, Addendum.Drucker.StandardA4)
			;~ VerifiedClick("Button44", hfoxitPrint,,, true)            	; Speichern Button drücken
			;FoxitInvoke("Close", PDFViewerHwnd)                       	; dieses Dokument schließen
			If WinExist("ahk_id " PDFViewerHwnd)
				WinMinimize, % PDFViewerHwnd
			AlbisActivate(2)                                                        	; Albis wieder aktivieren
			SetTitleMatchMode, % OldMatchMode

		}
		else If InStr(PDFViewerClass, "Sumatra") {	; Sumatra PdfReader

			;~ SumatraInvoke("Print", PDFViewerHwnd)                  	; 'Drucken' - Dialog
			;~ WinWait, % sumatraSaveAs,, 6                                  	; wartet 6 Sekunden auf das Dialogfenster
			;~ hSaveAs := GetHex(WinExist(sumatraSaveAs))        	; 'Speichern unter' - Dialog handle
			;~ VerifiedSetText("Edit1", PdfFullFilePath , hSaveAs)     	; Speicherpfad eingeben
			;~ VerifiedClick("Button2", hSaveAs,,, true)                  	; Speichern Button drücken
			;~ SumatraInvoke("Close", PDFViewerHwnd)                  	; dieses Dokument schließen
			;~ AlbisActivate(2)                                                        	; Albis wieder aktivieren

		}


}

; Karteikartenfunktionen
AlbisLeseDatumUndBezeichnung(MouseX, MouseY) {                                                                	;-- Mausposition abhängige Karteikarten Informationen

		while (StrLen(KKDatum) = 0) {

			MouseClick, Left, MouseX, MouseY                                                                     	; Mausklick setzt den Cursor
			Sleep, 300

			Content 	:= AlbisGetActiveControl("Content")
			PdfTitel  	:= Content.RichEdit
			KKDatum	:= Content.Edit2

			If (StrLen(KKDatum) > 0)                                                                                      	; Datum konnte ausgelesen werden, dann Abbruch
				break
			If (A_Index > 4)
				return ""

		}

		SendInput, {Escape}

		PdfTitel  	:= RegExReplace(PdfTitel, "\.*pdf$", "")                                                            ; eventuell noch vorhandene pdf Endung aus dem Karteikartentext entfernen

return {"Text": PdfTitel, "Datum":KKDatum}
}


