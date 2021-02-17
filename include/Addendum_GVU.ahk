; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                          	   ausgelagerte Funktionen für die Automatisierung der Formularbefüllung der GVU - z.Z. nicht in Benutzung
;                                                            	by Ixiko started in September 2017 - last change 01.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;		func_ZeigeGVUs              	:= Func("GVU_Gui")
;		Menu, SubMenu4, Add, % "Vorsorgeuntersuchungen"         	, % func_ZeigeGVUs

;   	Hotkey, #F7  					, % func_ZeigeGVUs                  			;= Albis: GVUListe anzeigen, bearbeiten und ausführen lassen
;		GVUListeAnzeigen:        	;{
;			gosub GVU_GUI
;		return ;}

GVU_GUI() {

	; Variablen
		;global GVU, hGVu, hLV_GVU, LBQuartale, GVULV, LVCount
		global 	GVUTextProgress, GVUProgress, GVU_LV, GVURun, hGVU_LV, GVUInfo, ZText
		;local 	LBQuartale
		Quartale     	:= ""
		GVUListe   	:= Object()
		aktuelleListe	:= IniReadExt("GVUListe", "aktuelles_GVUFile")

	; Inhalt vorhandener Dateien in ein Objekt einlesen
		Loop, Files, % AddendumDir "\TagesProtokolle\*-GVU.txt"
		{
				SplitPath, A_LoopFileFullPath,,,, filename
				Quartal:= StrReplace(filename, "-GVU")
				Lnr:= A_Index
				GVUListe[Lnr]:= {"Quartal": Quartal, "Liste":[]}
				Quartale.= Quartal "|"
				FileRead, file, % A_LoopFileFullPath
				Loop, Parse, file, `n, `r
				{
					If StrLen(A_LoopField) > 0
						GVUListe[Lnr].Liste.Push(A_LoopField)
				}
		}

		Quartale        	:= RTrim(Quartale, "|")
		GVUPrg_Width	:= 680
		GVULV_Width	:= 800
		GVULVOptions	:= "xm w" GVULV_Width " r30 BackgroundAAAAFF Grid NoSort"

		Gui, GGVU: New  	, +AlwaysOnTop +HwndhGVU -DPIScale
		Gui, GGVU: Font  	, S10 Normal q5, % Addendum.Default.Font
		Gui, GGVU: Margin	, 5, 5
		Gui, GGVU: Add   	, Text    	, xm ym+3                                                                               	, % "angezeigtes Quartal:"
		Gui, GGVU: Add    	, ListBox	, x+5 ym+3 r1 gGVU_LBEvent vLBQuartale                              	, % Quartale
		Gui, GGVU: Add   	, Button  	, x+5 ym w200 gGVU_Ablaufstarten vGVU_Run +0x00000300  	, % "Formularerstellung starten"
		Gui, GGVU: Font  	, s8 Normal q5, % Addendum.Default.Font
		Gui, GGVU: Add   	, Text     	, x+10 y10 w250 vZText                                                          	, % "Vorsorgen unbearbeitet: " GVUListe[Lnr].Liste.MaxIndex()
		Gui, GGVU: Font  	, S10 Normal q5, % Addendum.Default.Font
		Gui, GGVU: Add   	, Progress	, % "xm w" GVUPrg_Width " h20 CBLime BackgroundEEEEEE cBlue vGVUProgress", % "0"
		Gui, GGVU: Add   	, Text    	, % "x+5 w" (GVULV_Width - GVUPrg_Width - 5) " vGVUTextProgress"                                                 	, % ""
		Gui, GGVU: Add   	, Listview	, % GVULVOptions " vGVU_LV HWNDhGVU_LV"                    	, % "U-Nr|bearbeitet|PatientenID|Name|Geburtsdatum|Untersuchungsdatum|Hinweis"
		GuiControl, GGVU: ChooseString, LBQuartale, % aktuelleListe

		gosub GVU_LVFuellen

		Gui, GGVU: Show 	, xCenter yCenter AutoSize   	, % "Anzeige der Patienten in der GVU Liste"


return

GGVUGuiEscape:     	;{
GGVUGuiClose:
	If InStr(Running, "Vorsorgeautomatisierung")
	{
			Running:= "Vorsorgeautomatisierung beendet"
			GuiControl, GGVU: Text, GVUInfo, % "Einen Augenblick bitte ...`n`nDas Fenster wird nach`nvollständiger Formularerstellung`ngeschlossen."
			GuiControl, GGVU: Show, GVUInfo
			while, !StrLen(Running) = 0
			{
					Sleep, 20
					If A_Index > 3000 ; 1min
						break
			}
	}
	Gui, GGVU: Destroy
return ;}

GVUListView:            	;{

return ;}

GVU_LBEvent:           	;{
	Gui, GGVU: Default
	LV_Delete()
	gosub GVU_LVFuellen
return ;}

GVU_LVFuellen:        	;{

	Gui, GGVU: Default
	Gui, GGVU: Submit, NoHide

	For Lnr, data in GVUListe
		If InStr(data.Quartal, LBQuartale)
			break

	Loop, % GVUListe[Lnr].Liste.MaxIndex()
	{
			arr:= StrSplit(GVUListe[Lnr].Liste[A_Index], ";")
			If StrLen(arr[4])>0
				bearbeitet:= "ja"
			else
				bearbeitet:= "nein"
			LV_Add("", A_Index, bearbeitet, arr[3], oPat[Arr[3]].Nn ", " oPat[Arr[3]].Vn, oPat[Arr[3]].Gd, arr[1])
	}

	LV_ModifyCol()
	LV_ModifyCol(7, 300)
	LV_ModifyCol(2, "SortDesc")

return ;}

GVU_Ablaufstarten:   	;{

	Gui, GGVU: Default
	GuiControlGet, GVURunStatus, GGVU:, GVURun, Text
	If InStr(GVURunStatus, "Starten")
	{
			Gui, GGVU: Default
			GuiControl , GGVU: Text, GVURun, % "Formularerstellung pausieren"
		; unbearbeitete Sammeln und einer Funktion übergeben
			IDListe 	:= Object()
			IDIndex	:= 0
			PraxTT("Formularlauf ist gestartet!", "5 3")

			Gui, GGVU: Default
			Gui, GGVU: ListView, GVU_LV
			Loop, % LV_GetCount()
			{
					LV_GetText(bearbeitet, A_Index, 2)
					If InStr(bearbeitet, "nein")
					{
							IDIndex ++
							LV_GetText(UNr	    	, A_Index, 1)
							LV_GetText(ID         	, A_Index, 3)
							LV_GetText(Gd         	, A_Index, 5)
							LV_GetText(UDatum	, A_Index, 6)
							IDListe[IDIndex] := Object()
							IDListe[IDIndex] := {"Unr": A_Index, "PatientenID": ID, "Geburtsdatum": Gd, "UDatum": UDatum}
					}
			}

			If IDIndex = 0
			{
					MsgBox, Es sind alle Untersuchungen angelegt und abgerechnet worden.
					return
			}

			GVU_Automat(IDListe)
	}
	else If InStr(GVURunStatus, "pausieren")
	{
			Gui, GGVU: Default
			GuiControl, GGVU: Text, GVURun, % "Formularerstellung fortsetzen"
			GuiControl, GGVU: Show, GVUInfo
			Running:= "Vorsorgeautomatisierung pausiert"
	}
	else If InStr(GVURunStatus, "fortsetzen")
	{
			Gui, GGVU: Default
			GuiControl, GGVU: Text, GVURun, % "Formularerstellung pausieren"
			GuiControl, GGVU: Hide, GVUInfo
			Running:= "Vorsorgeautomatisierung"
	}

return ;}

}

GVU_Automat(IDListe) {

	Running:= "Vorsorgeautomatisierung"
	z:= []
	PatToDo:= IDListe.MaxIndex()
	lenToDo:= StrLen(PatToDo)
	Sleep, 3000

	Gui, GGVU: Default
	Gui, GGVU: ListView, GVU_LV
	GuiControl, GGVU:, GVUTextProgress, % SubStr("0000", -1 * lenToDo) "/" PatToDo

	For Index, data in IDListe	{

		PatID:= data.PatientenID
		z[1]:= "akt. Patient     : " oPat[PatID].Nn ", " oPat[PatID].Vn ", geb. am " oPat[PatID].Gd
		z[2]:= "Health-Status  : " Running
		z[3]:= "A_Index          : " A_Index
		t:=""
		Loop % z.MaxIndex()
			t.= z[A_Index] "`n"
		Gui, GGVU: Default
		GuiControl, GGVU: Text, ZText, % t

		; Überprüfen des Patientenalters und Abbruch falls notwendig
			age:= Floor(DateDiff( "YY", data.Geburtsdatum, data.UDatum))
			If (GVU[minAlter] > age)	{

				PraxTT("Patientenalter: " age ", Mindestalter: " GVU[minAlter] "`nPatient unterschreitet das Mindesalter", "5 3")
				GVU_DatenSichern(data.PatientenID, data.UDatum, ">>>Das Alter des Patienten unterschreitet das Mindestalter<<<", "")
				Sleep, 4000
			  ; GUI Anzeige Fortschritt zeigen
				Gui, GGVU: Default
				Gui, GGVU: ListView, GVU_LV
				Loop % LV_GetCount()	{
					LV_GetText(ID, A_Index, 3)
					If (ID = data.PatientenID)		{
						LV_Modify(A_Index,,, "ja")
						break
					}
				}
				Sleep, 200
				LV_ModifyCol(2, "SortDesc")
				GuiControl, GGVU:, GVUTextProgress, % SubStr("0000" Index, -1 * lenToDo) "/" PatToDo
				GuiControl, GGVU:, GVUProgress, % Floor(Index*100/PatToDo)
				continue

			}

			If InStr(Running, "Vorsorgeautomatisierung pausiert")
				GVU_Pause()
			else If InStr(Running, "Vorsorgeautomatisierung beendet")
				break

			AlbisAkteOeffnen(data.PatientenID, data.PatientenID)
			DatumZuvor := AlbisSetzeProgrammDatum(data.UDatum)

			If InStr(Running, "Vorsorgeautomatisierung pausiert")
				GVU_Pause()

		; Albis GVU Makro starten (über ein Kürzel mit dem Namen GVU wird erst das GVU- dann das HKS Formular aufgerufen und automatisch die Ziffern eingetragen)
			AlbisKarteikartenFocusSetzen("Edit2")
			VerifiedSetText("Edit2", data.UDatum, "ahk_class OptoAppClass")
			SendInput, {Tab}
			Sleep, 100
			ControlFocus, Edit3, ahk_class OptoAppClass
			VerifiedSetText("Edit3", "GVU", "ahk_class OptoAppClass")
			SendInput, {Tab}

		; wartet auf Abschluß der gesamten Prozedur
			WinGetTitle, AlbisT1, ahk_class OptoAppClass
			while !(StrLen(DatumZuvor) = 0) {

				If GetKeyState("Escape")
					return
				Sleep, 150
				WinGetTitle, AlbisT2, ahk_class OptoAppClass
				If (AlbisT1 <> AlbisT2)
					break
			}

		; GUI Anzeige Fortschritt zeigen
			Gui, GGVU: Default
			Gui, GGVU: ListView, GVU_LV
			LV_Modify(data.UNr,,, "ja")
			Sleep, 50
			LV_ModifyCol(2, "SortDesc")
			GuiControl, GGVU:, GVUTextProgress, % SubStr("0000" Index, -1 * lenToDo) "/" PatToDo
			GuiControl, GGVU:, GVUProgress, % Floor(Index*100/PatToDo)

			If InStr(Running, "Vorsorgeautomatisierung pausiert")
					GVU_Pause()

		; Fortfahren
			MsgBox, 4, Addendum, Möchten Sie mit dem nächsten Patienten fortfahren?
			IfMsgBox, No
				break

		; Akte schließen
			WinGetTitle, AlbisT1, ahk_class OptoAppClass
			Albismenu(57602)
			while, (AlbisT1 = AlbisT2)				{

				WinGetTitle, AlbisT2, ahk_class OptoAppClass
				If AlbisT1 <> AlbisT2
					break
				Sleep, 50

			}

	}

	PraxTT("Der Formularlauf ist beendet!", "5 3")
	Running:= ""

return
}

GVU_Pause() {

	Loop	{

		If RegExMatch(Running, "^Vorsorgeautomatisierung$") || InStr(Running, "beendet") || (StrLen(Running) = 0)
			break
		Sleep, 50
		If RegExMatch(Running, "^Vorsorgeautomatisierung$") || InStr(Running, "beendet") || (StrLen(Running) = 0)
			break
		Sleep, 50

	}

return
}

GVU_CaveVonEintragen(UDatum, SaveData:=true) {                                               	;-- zum Ändern der Cave! von Zeile 9 (hier bewahre ich Daten der letzten Untersuchungen auf)

		RegExMatch(UDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
		GVUMonat	:= aGVU1
		GVUJahr   	:= aGVU2

	; Ändern des Kürzeltextes der letzten GVU
		GVUZeile:= AlbisGetCaveZeile("", "GVU", true)

	; Sichern der ausgelesenen Cave! von Zeile 9
		PatientenID:= AlbisAktuellePatID()
		FileAppend, % A_DD "." A_MM "." A_Year ";" PatientenID ";" AlbisCurrentPatient() ";" GVUZeile "`n", % AddendumDir . "\Tagesprotokolle\cave9.txt", UTF-8

	; Erstellen der neuen Zeile ;{
		If RegExMatch(GVUZeile, "[gvuGVU]+.[hksHKS]+\/*[kvuKVU]*\s\d+.\d+")
			neueZeile:= RegExReplace(GVUZeile, "[gvuGVU]+.[hksHKS]+\/*[kvuKVU]*\s*\d+.\d+", "GVU/HKS " GVUMonat "^" GVUJahr)
		else	{

			If InStr(GVUZeile, ",")				{

				GVUTeile:= StrSplit(GVUZeile, ",")
				MaxIndex:= GVUTeile.MaxIndex()
				If RegExMatch(GVUTeile[MaxIndex], "GB\s*\d+\.\d\s*[AB]")	{

					GVUTeile[MaxIndex+1] := GVUTeile[MaxIndex]
					GVUTeile[MaxIndex]  	:= "GVU/HKS " GVUMonat "^" GVUJahr
					Loop, % MaxIndex + 1
						neueZeile .= GVUTeile[A_Index] ", "

				}
				else
					neueZeile := GVUZeile ", GVU/HKS " GVUMonat "^" GVUJahr

			}
			else				{

				if StrLen(GVUZeile) = 0
					neueZeile := "GVU/HKS " GVUMonat "^" GVUJahr ", "
				else
					neueZeile := GVUZeile ", GVU/HKS " GVUMonat "^" GVUJahr

				}

		}

		neueZeile:= RTrim(neueZeile, ", ")
		;}

	; Schreiben der neuen Zeile
		AlbisSetCaveZeile(9, neueZeile)
		Clipboard:= neueZeile

		If SaveData
			GVU_DatenSichern(PatientenID, UDatum, GVUZeile, neueZeile)

return
}

GVU_DatenSichern(PatientenID, UDatum, GVUAlt:="", GVUNeu:="") {

		PatVorhanden:= false

	; Sichern der Daten in die GVU Liste
		RegExMatch(UDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
		QListe:= GetQuartal("01." aGVU1 "." aGVU2)
		GVUtoAddfile:= AddendumDir "\Tagesprotokolle\" QListe "-GVU.txt"

	; Anlegen einer Sicherheitskopie der *gvu.txt Datei
		;Loop
		; {
		; 		backupFile:= AddendumDir "\TagesProtokolle\" SubStr(GVUFile, 1, StrLen(GVUFile) - 4) . A_Index . ".txt"
		; } until !FileExist(backupFile)
		; FileCopy, % AddendumDir "\TagesProtokolle\" GVUFile, % backupFile

		FileRead, gfile, % GVUtoAddfile
		Loop, Parse, gfile, `n, `r
		{
				If StrLen(A_LoopField) = 0
					continue
				else If InStr(A_LoopField, PatientenID)
				{
					PatVorhanden := true
					If RegExMatch(GVUAlt, "^\>\>\>")
						newgFile.= A_LoopField ";" GVUAlt "`n"
					else
						newgFile.= A_LoopField ";Alt:" GVUAlt ";Neu:" GVUNeu "`n"
				}
				else
					newgFile.= A_LoopField "`n"
		}

		If !PatVorhanden
			newgFile.= UDatum ";" SubStr(QList, 1, 2) "/" SubStr(QList, 3, 2) ";" PatientenID ";Alt:" GVUAlt ";Neu:" GVUNeu "`n"

		newgFile:= RTrim(newgFile, "`n")

		FileDelete	, % GVUtoAddfile
		FileAppend, % newgFile, % GVUtoAddfile, UTF-8

		If oPat.HasKey(PatientenID)
		{
				oPat[PatientenID].letzteGVU := UDatum
				PatDBSave(AddendumDBPath)
		}

return
}

GVU_AlteFormulardatenUebernehmen() {

	; Auslesen des letzten Abrechnungsdatum einer GVU
		entry := AlbisReadFromListbox("Alte Formulardaten übernehmen ahk_class #32770", 1, 1)
		RegExMatch(entry, "\d+\.(\d+)\.\d\d(\d+)", lastGVU)
		RegExMatch(EditDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
	; Berechnung der Monate zwischen zwei Vorsorgeuntersuchungen
		Untersuchungsabstand := ((aGVU2*12)+aGVU1) - ((lastGVU2*12)+lastGVU1)

		If (Untersuchungsabstand < Addendum.GVUAbstand)		{
			;MsgBox, % "Zeile: " A_LineNumber "`ndelta: " Untersuchungsabstand "`nAbstand: " GVU[Abstand] ; GVU[Abstand]
			ClickReady := 1
			PraxTT("Bei diesem Patienten kann im Moment noch`nkeine weitere GVU abgerechnet werden.`nEs fehlen noch " 24 - Untersuchungsabstand " Monate!", "6 3")
			Sleep, 2000
			VerifiedClick("Button2",  "Alte Formulardaten übernehmen ahk_class #32770")	; Button 2 - [Abbruch]
			GVU_DatenSichern(AlbisAktuellePatID(), EditDatum, ">>>Mindestabstand zwischen zwei Vorsorgeuntersuchung nicht errreicht<<<", "")
			GVU_CaveVonEintragen(lastGVU, false)
		}
		else
		{
				ClickReady := 4
				PraxTT("Eine neue GVU wird angelegt.", "3 3")
				GVU_FormularDatenAuswaehlen()
		}


return ClickReady
}

GVU_FormularDatenAuswaehlen() {

		VerifiedSetFocus("Listbox1", "Alte Formulardaten übernehmen ahk_class #32770", 1)
		SendInput, {Space}
		Sleep, 100
		VerifiedClick("Button1", "Alte Formulardaten übernehmen ahk_class #32770")
		WinWaitClose, % "Alte Formulardaten übernehmen ahk_class #32770", , 3

return
}

GVU_KeineAlteDatenVorhanden() {

		VerifiedClick("Button1", "ALBIS ahk_class #32770", "alten Daten vorhanden") ; schließen

		hSeite1:= WinExist("Muster 30 (01.2009), Gesundheitsuntersuchung (Seite 1)")
		VerifiedCheck("Button52", hSeite1) ; Beweg.apparat
		VerifiedCheck("Button57", hSeite1) ; 140/90
		  VerifiedClick("Button61", hSeite1)

		WinWait, % "Muster 30 (01.2009), Gesundheitsuntersuchung (Seite 2) ahk_class #32770",,3
		hSeite2:= WinExist("Muster 30 (01.2009), Gesundheitsuntersuchung (Seite 2)")
		VerifiedCheck("Button37", hSeite2) ; orthopädische Erkrankung
		VerifiedCheck("Button57", hSeite2) ; sonstiges

		VerifiedClick("Button61", hSeite2)	;Button Weiter

return
}

GVU_LeistungsketteBestaetigen() {

		VerifiedSetFocus("Listview321", "Leistungskette bestätigen ahk_class #32770")
		SendInput, {Space}
		Sleep, 100
		SendInput, {Space}
	; Dialog schliessen
		VerifiedClick("Button1", "Leistungskette bestätigen ahk_class #32770")
		WinWaitClose, Leistungskette bestätigen ahk_class #32770,, 3

return
}

GVU_HKSFormularBefuellen() {

	; Verdachtsdiagnose "nein" - warum Fehler wenn nicht ausgewählt
		VerifiedClick("Button28", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Malignes Melanom"
		VerifiedClick("Button10", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Basalzellkarzinom"
		VerifiedClick("Button12", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Spinozelluläres Karzinom"
		VerifiedClick("Button14", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Anderer Hautkrebs"
		VerifiedClick("Button24", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Sonstiger dermatologisch abklärungsbedürftiger Befund"
		VerifiedClick("Button26", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Screening-Teilnehmer wird an einen Dermatologen überwiesen:"
		VerifiedClick("Button30", "Hautkrebsscreening - Nichtdermatologe")
	; Häkchen bei "gleichzeitige Gesundheitsvorsorge setzen"
		VerifiedClick("Button16", "Hautkrebsscreening - Nichtdermatologe")
	; "Speichern" und damit auch Schließen des Formulares
		VerifiedClick("Button19", "Hautkrebsscreening - Nichtdermatologe")

return
}

SetClickReadyBack:                                                                                              	;{ Zurücksetzen einer bei der GVU Automatisierung benutzten Variable
	ClickReady:= 1
return ;}

GVU_WarteAufAlteDaten:                                                                                    	;{

		If (A_TickCount - GVU_AlteDatenStartZeit > 6000)	{
			SetTimer, GVU_WarteAufAlteDaten, Off
			PraxTT("Alte Daten wurden eingetragen", "5 3")
			ClickReady := 4
			VerifiedClick("Button61", "Muster 30 ahk_class #32770")	; Button61 = [Weiter]   	 - GVU Seite 1
			return
		}

		If (ClickReady <> 2) {
			SetTimer, GVU_WarteAufAlteDaten, Off
			PraxTT("Alte Daten wurden übernommen.", "5 3")
		}

return ;}

PatDBSave(AddendumDBPath) {                                                                       	;-- zum Sichern der Patientendatenbank

	dbFile:= FileOpen(AddendumDbPath "\Patienten.txt", "w", "UTF-8")
	For PatID, obj in oPat
	{
			line           	:=  PatID ";" obj.Nn ";" obj.Vn ";" obj.Gt ";" obj.Gd ";" obj.Kk ";" obj.letzteGVU
			dataToWrite	.=  RTrim(line, ";") "`n"
	}
	dbFile.Write(dataToWrite)
	dbFile.Close()

return
}

/*  Addendum.ahk - Label EventHook_WinHandler:

	If        InStr(EHproc1, "albis")                                                                      	{

			...
			...
			...
			...
			.
			.
			.


			;-------------------- GVU und HKS Formular Automatisierung -----------------------------------------
			else If Addendum.GVUAutomation && InStr(EHWT  	, "Muster 30")                                             	&& (ClickReady > 0)     		{

					RegExMatch(EHWT, "(?<=Seite\s)\d", FormularSeite)
					If (FormularSeite = 1)     				{

							If (ClickReady = 1)							{

									ClickReady := 2
								; überprüft auf welche Art das GVU Formular aufgerufen wurde, verhindert die Automatisierung schon angelegter Formulare
								; funktioniert nur wenn man sich ein Albismakro mit dem Namen GVU anlegt! In Edit3 steht dann GVU
									EditDatum := EditKuerzel := ""
									ControlGetText, EditDatum , Edit2, ahk_class OptoAppClass
									ControlGetText, EditKuerzel, Edit3, ahk_class OptoAppClass
									If (StrLen(EditDatum) = 0) || !InStr(EditKuerzel, "GVU")
											DatumZuvor := "", ClickReady := 1

								; Quartal, Untersuchungsmonat und Jahr ermitteln
									AbrQuartal := LTrim(GetQuartal(EditDatum, "/"), "0")
									RegExMatch(EditDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
									GVU_AlteDatenStartZeit := A_TickCount
									PraxTT("Warte auf alte Daten", "0 3")
									SetTimer, GVU_WarteAufAlteDaten, 250
									If !VerifiedClick("Alte Daten", "Muster 30 ahk_class #32770")
										VerifiedClick("Button63", "Muster 30 ahk_class #32770")	        	; Button63 = [Alte Daten] - GVU Seite 1

							}
							else if (ClickReady = 3)							{

									SetTimer, GVU_WarteAufAlteDaten, Off
									PraxTT("Abbruch", "2 3")
									VerifiedClick("Button62", "Muster 30 ahk_class #32770")	        	; Button62 = [Abbruch]	 - GVU Seite 1
									If RegExMatch(DatumZuvor, "\d{2}\.\d{2}\.\d{4}")
										AlbisSetzeProgrammDatum(DatumZuvor)
									DatumZuvor := ""
									SetTimer, SetClickReadyBack, -4000

							}
							else if (ClickReady = 4)							{

									SetTimer, GVU_WarteAufAlteDaten, Off
									PraxTT("Weiter", "2 3")
									VerifiedClick("Button61", "Muster 30 ahk_class #32770")	        	; Button61 = [Weiter]   	 ->	GVU Seite 2

							}

					}
					else if (FormularSeite = 2)				{
						ClickReady := 5
						SetTimer, GVU_WarteAufAlteDaten, Off
						PraxTT("GVU Formular Seite 2 schließen", "2 3")
						VerifiedClick("Button61", "Muster 30 ahk_class #32770", "", "", 3)      	; Button61 = [Weiter]   	 ->	GVU Seite 2

					}

			}
			else If Addendum.GVUAutomation && InStr(EHWText	, "keine alten Daten vorhanden")                	&& (ClickReady = 2)			{
					ClickReady := 0
					PraxTT("", "off")
					SetTimer, GVU_WarteAufAlteDaten, Off
					GVU_KeineAlteDatenVorhanden()
					ClickReady := 5
			}
			else If Addendum.GVUAutomation && InStr(EHWT  	, "Alte Formulardaten übernehmen")           	&& (ClickReady = 2)			{
						ClickReady := 0 ; auf 0 gesetzt damit nichts anderes den Ablauf hier behindern kann
						SetTimer, GVU_WarteAufAlteDaten, Off
						ClickReady := GVU_AlteFormulardatenUebernehmen()
			}
			else If Addendum.GVUAutomation && InStr(EHWText	, "Soll das Makro abgebrochen werden")      	&& (ClickReady > 1)			{
			    	VerifiedClick("Button1", "ALBIS", "Soll das Makro abgebrochen werden")
					ClickReady	:= 1
					SetTimer, SetClickReadyBack, off
			}
			else If Addendum.GVUAutomation && InStr(EHWT  	, "GNR-Vorschlag zur Befundung")              	&& (ClickReady = 5)			{

						entries := ""
						while (StrLen(entries) = 0)		{

							entries := AlbisReadFromListbox("GNR-Vorschlag zur Befundung", 1, 0)
							If (StrLen(entries) > 0)
								break
							;SciteOutPut("entrieTry: " A_Index "`tQuartal: " AbrQuartal "`tEinträge: " entries)
							Sleep 300
							If (A_Index > 5) 	{
								PraxTT("Die Erkennung des Abrechnungsquartals im Dialog hat nicht funktioniert`nBitte wählen Sie das richtige Abrechnungsquartal manuell aus!", "8 4")
								ClickReady:= 6
							}

						}

						Loop, Parse, entries, `n
							If InStr(A_LoopField, AbrQuartal)	{
								PostMessage, 0x186, % A_Index - 1, 0, ListBox1, % "ahk_id " WinExist("GNR-Vorschlag zur Befundung ahk_class #32770") ;LB_SetCursel
								entrieIndex:= A_Index - 1
								break
							}

						;ToolTip, % "entrieId: " entrieIndex "`nQuartal: " AbrQuartal "`nEinträge: " entries, 1250, 1, 5
						If !VerifiedClick("Button1", "GNR-Vorschlag zur Befundung ahk_class #32770", "", "", 3) ;GNR-Vorschlag schliessen
								MsgBox, Bitte schließen Sie den Dialog`nGNR-Vorschlag zur Befundung

						ClickReady := 6
						WinWait, Hautkrebsscreening - Nichtdermatologe ahk_class #32770,, 8
						If ErrorLevel
							WinActivate, Hautkrebsscreening - Nichtdermatologe ahk_class #32770
						else
							Albismenu(34505, "Hautkrebsscreening - Nichtdermatologe")     	;"eHautkrebs-Screening Nicht-Dermatologe": "34505" wird aufgerufen



			}
			else If Addendum.GVUAutomation && InStr(EHWT  	, "Hautkrebsscreening - Nichtdermatologe")	&& (ClickReady = 6)			{
						GVU_HKSFormularBefuellen()
						ClickReady:= 7
			}
			else If Addendum.GVUAutomation && InStr(EHWT  	, "Leistungskette bestätigen")                        	&& (ClickReady = 7)			{
					ClickReady:= 1
					GVU_LeistungsketteBestaetigen()
					GVU_CaveVonEintragen(EditDatum)
					If RegExMatch(DatumZuvor, "\d{2}\.\d{2}\.\d{4}")
						AlbisSetzeProgrammDatum(DatumZuvor)
					DatumZuvor := ""
			}
			else If Addendum.GVUAutomation && InStr(EHWText	, "Übertrage Gebühren-Nummer(n)...")		                                        	{
					VerifiedClick("Button1", "ALBIS", "Übertrage Gebühren-Nummer")
			}
			else If Addendum.GVUAutomation && InStr(EHWText	, "außerhalb des Quartals in dem der Schein gültig ist")		                	{        	;<-- überarbeiten - mehr Code notwendig
					VerifiedClick("Button1", hHookedWin)
			}



*/
