; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                     INFOFENSTER - 	Funktion:   	wird in die Patientenakte im Albisprogramm eingeblendet
;																					   							Basisskript: 	Addendum.ahk
;
;	                                                                                                                Addendum für Albis on Windows
;                                                            	by Ixiko started in September 2017 - last change 14.07.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; Infofenster
AddendumGui() {

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		global

		LVPatOpt    	 := " r" (Addendum.InfoWindow.LVScanPool.r)
		LVPatOpt   	 .= " gadmPatLV 	    	vAdmPdfReports 	HWNDhPdfReports 	AltSubmit -Hdr	ReadOnly	-LV0x10 BackgroundF0F0F0 E0x20000"
		LVJournalOpt .= " gadmJournal     	vAdmJournal 		HWNDhJournal 	 	AltSubmit                      	-LV0x10 BackgroundF0F0F0 E0x20200"
		LVTProtokoll 	 := " gadmTProtokollLV	vAdmTProtokoll 	HWNDhTProtokoll 	AltSubmit      	ReadOnly	-LV0x10 BackgroundF0F0F0 E0x20000"

		local LVContent
		local ImageListID, adm_init := false, hbmPdf_Ico, hbmImage_Ico
		local adm, adm2, adm3, admWidth, admHeight, TabSize, TabSizeX, TabSizeY, TabSizeW, TabSizeH
		local APos, SPos, rpos, mox, moy, moWin, moCtrl, regExt, Aktiv, col1W, col2W, col3W, Pat
		local admReCheck := 0

		If !adm_init {

				adm_init  		:= true
				fc_admGui	:= Func("AddendumGui")

				ImageListID  	:= IL_Create(2)
				IL_Add(ImageListID, Addendum.AddendumDir "\assets\ModulIcons\pdf.ico" 	, 0xFFFFFF, 0) 		;1
				IL_Add(ImageListID, Addendum.AddendumDir "\assets\ModulIcons\image.ico", 0xFFFFFF, 0)		;2

				;hbmPdf_Ico   	:= Create_PDF_ico()
				;hbmImage_Ico	:= Create_Image_ico()
				;~ IL_Add(ImageListID, "HBITMAP: " hbmPdf_Ico 		, 0xFFFFFF, 0) 		;1
				;~ IL_Add(ImageListID, "HBITMAP: " hbmImage_Ico	, 0xFFFFFF, 0)		;2

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Kontextmenu (Rechtsklickmenu) für das Journal
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				fc_admOpen 	:= Func("admGui_CM").Bind("JOpen")
				fc_admImport	:= Func("admGui_CM").Bind("JImport")
				fc_admRename	:= Func("admGui_CM").Bind("JRename")
				fc_admDelete	:= Func("admGui_CM").Bind("JDelete")
				fc_admView		:= Func("admGui_CM").Bind("JView")
				fc_admRefresh	:= Func("admGui_CM").Bind("JRefresh")

				Menu, admJCM, Add, Patient öffnen                  	, % fc_admOpen
				Menu, admJCM, Add, Datei importieren             	, % fc_admImport
				Menu, admJCM, Add, Datei umbennen              	, % fc_admRename
				Menu, admJCM, Add, Datei löschen                    	, % fc_admDelete
				Menu, admJCM, Add, Datei anzeigen                	, % fc_admView
				Menu, admJCM, Add, Befundordner neu indizieren, % fc_admRefresh
			;}

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Hotkey's für alle Tab's
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				;Hotkey, $F1, admTestung
				Hotkey, IfWinActive, AddendumGui ; ahk_class AutoHotkeyGUI"
				Hotkey, Enter       	, % fc_admView
				Hotkey, F3         	, % fc_admView
				Hotkey, F2        	, % fc_admRename
				Hotkey, F5        	, % fc_admRefresh
				Hotkey, BackSpace, % fc_admDelete
				Hotkey, IfWinActive
			;}

		}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Vorbereitungen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	; Albis wurde beendet -return-
		If !WinExist("ahk_class OptoAppClass")
			return

	; mehr als 3 Versuche das MDIFrame zu finden führen zum Abbruch des Selbstaufrufes der Funktion
		If (admReCheck > 9) {
			admReCheck := 0
			return
		}

	; noch evtl. vorhandene Infofenster schliessen
		If WinExist("ahk_id " hadm)
			Gui, adm: Destroy
		If WinExist("ahk_id " hadm2)
			Gui, adm2: Destroy

		Aktiv := Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		hMDIFrame	:= Controls("AfxMDIFrame1401", "ID", AlbisGetActiveMDIChild())
		hStamm		:= Controls("#327701", "ID", hMDIFrame)
		If !RegExMatch(Aktiv, "i)Karteikarte|Laborblatt|Biometriedaten") || (GetDec(hMDIFrame) = 0) {
			SetTimer, % fc_admGui, -1000
			admReCheck ++
			return
		}

	; legt die Höhe der Gui anhand der Größe des MDIFrame-Fenster fest
		APos := GetWindowSpot(AlbisWinID())
		SPos := GetWindowSpot(hStamm)

		Addendum.InfoWindowX	:= APos.CW - Addendum.InfoWindow.W ; ClientWidth passt besser!
		admHeight 	:= Addendum.InfoWindow.H	:= SPos.CH - 10
		admWidth 	:= Addendum.InfoWindow.LVScanPool.W + 10

	; der Albisinfobereich (Stammdaten, Dauerdiagnosen, Dauermedikamente) wird für das Einfügen der Gui verkleinert
		ControlMove,,,, % (Addendum.InfoWindowX - SPos.X - 8),, % "ahk_id " hStamm

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------
	; Gui zeichnen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	; umgeht die seltene Problematik der MDIFrame ein anderes handle bekommen hat, die aktuelle Patientenakte geschlossen wurde
		try {
			Gui, adm: New 	, +HWNDhadm -Caption 0x50020000 E0x0802009C +Parent%hMDIFrame%
		} catch {
			SetTimer, % fc_admGui, -1000
			admReCheck ++
			return
		}

		Gui, adm: Margin	, 0, 0
		Gui, adm: Add  	, Tab     	, % "x0 y0  	w" admWidth " h" admHeight " hwndhadmTab vadmTabs"                           	, Patient|Journal|Protokoll|Info

	;-: Tab1 :- Patient                          	;{
		Gui, adm: Tab  	, 1
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text      	, % "x10 y27 w" admWidth - 15 " BackgroundTrans vAdmPdfReportTitel Section"               	, (es wird gesucht ....)
		Gui, adm: Font  	, s6 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth-55 " y23 w50 HWNDhAdmButton1 vAdmButton1 gBefundImport"         	, Importieren
		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+2 	w" admWidth - 6	" " LVPatOpt                                                                    	, % "S|Befundname"
		LV_SetImageList(ImageListID)
		GuiControlGet, cpos, adm: Pos, AdmPdfReports
		gpHeight := admHeight - cPosX - cposH - 45
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text, % "x2 y+0 	w" admWidth - 6 " Center vAdmGP1"  	                                                             	, Notizen/Erinnerungen
		GuiControlGet, cpos, adm: Pos, AdmGP1
		Gui, adm: Add  	, Edit     	, % "x2 y" cposY+12 " w" admWidth - 6 " h" admHeight - cposH - 35 " gAdmNotes vAdmNotes"
		;}

	;-: Tab2 :- Journal                         	;{
		Gui, adm: Tab  	, 2
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text      	, % "x10 y25  	w" admWidth - 22 " BackgroundTrans vAdmJournalTitel Section"               	, (es wird gesucht ....)
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth-145 " y23 w140 h16 HWNDhAdmButton2 vAdmButton2 gadmJournal"   	, obersten Befund Importieren
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+5 		w" admWidth - 6 " h" admHeight - 55 " " LVJournalOpt                             	, % "Befund|Eingang|TimeStamp"
		LV_SetImageList(ImageListID)
		col2W := 70, col1W := admWidth - col2W, col3W := 0
		LV_ModifyCol(1, col1W " Left NoSort")
		LV_ModifyCol(2, col2W " Left NoSort")
		LV_ModifyCol(3, col3W " Left Integer")              	; versteckte Spalte enthält Integerzeitwerte für die Sortierung des Datums
		;}

	;-: Tab3 :- Protokoll (letzte Patienten)	;{
		Gui, adm: Tab  	, 3

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		ProtText := "[" compname "] im Tagesprotokoll: " TProtokoll.MaxIndex() " Patienten"
		Gui, adm: Add  	, Text      	, % "x10 y25  	w" admWidth - 22 " BackgroundTrans vAdmTProtokollTitel"                     	, % ProtText

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth - 145 " y23 h16 vadmTPZurueck gadmTP"                                              	, % "<"

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x+2 y23 h16 vadmTPVor gadmTP"                                                                             	, % ">"
		Gui, adm: Add  	, Text      	, % "x+5 y25 BackgroundTrans vAdmTPTag"                                                                    	, % "Heute"

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+5   	w" admWidth - 6	 " h" admHeight - 55 " " LVTProtokoll                           	, % "RF|Nachname, Vorname|Geburtstag|PatID"

		col1W := 30, col2W := 160, col3W := 70, col4W := 70
		Gui, adm: ListView, AdmTProtokoll
		LV_ModifyCol(1, col1W )
		LV_ModifyCol(2, col2W )
		LV_ModifyCol(3, col3W )
		LV_ModifyCol(4, col4W )
		;}

	;-: Tab4 :- Info                              	;{
		Gui, adm: Tab  	, 4
		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		InfoText	:= "" ;{
		InfoText	.= "Sprechstunde heute: `n"
		if StrLen(Addendum.Praxis.Sprechstunde[A_DDDD]) > 0
			InfoText	.= A_DDDD ",der " A_DD "." A_MM "." A_YYYY " von " Addendum.Praxis.Sprechstunde[A_DDDD] " Uhr`n`n"
		else
			InfoText	.= "Keine Sprechstunde.`n`n"
		InfoText	.= "im Tagesprotokoll: " (TProtokoll.MaxIndex() = "" ? 0 : TProtokoll.MaxIndex()) " Patienten"
		;}
		Gui, adm: Add  	, Edit     	, % "x5 y25 w" admWidth - 10 " h" admHeight - 36 "t2 t6 t12", % InfoText
	;-:          :-
		Gui, adm: Tab
		;}

		Gui, adm: Show	, % "x" Addendum.InfoWindowX " y" Addendum.InfoWindow.Y " w" Addendum.InfoWindow.W  " h" Addendum.InfoWindow.H " NA ", AddendumGui
		Gui, adm: Default

		WinSet, Style, 0x50020000, % "ahk_id " hadm
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Tabs - Listviews mit Daten befüllen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		Addendum.InfoWindow.PatID := AlbisAktuellePatID()
		If (Addendum.InfoWindow.PatID <> Addendum.InfoWindow.lastPatID) {
				Addendum.InfoWindow.ReIndex := true
				Addendum.InfoWindow.lastPatID:= Addendum.InfoWindow.PatID
		}
		else {
				Addendum.InfoWindow.ReIndex := false
		}

	; Inhalte auffrischen
		BefundOrdner_Indizieren()
		PdfReports := admGui_Reports()
		admGui_Journal()
		admGui_TProtokoll()

	; Gui aktivieren damit es angezeigt wird
		WinActivate, % "ahk_id " hadm
		WinSet, AlwaysOnTop, On ,	% "ahk_id " hadm
		WinSet, Top                    	,,	% "ahk_id " hadm
		AlbisActivate(1)

	;}

		admReCheck := 0
		gosub toptop
		SetTimer, toptop, % Addendum.InfoWindow.RefreshTime

return

admTestung:
	SciTEOutput("Gut")
return
}

admNotes:						;{

return ;}

admPatLV:                    	;{ Patienten Tab 	- Listview Handler

	;Critical
	LV_GetText(admFile, EventInfo:= A_EventInfo, 1)

  ;PDF Vorschau
	If Instr(A_GuiEvent, "DoubleClick")
			admGui_View(admFile)

return ;}

admJournal:                 	;{ Journal Tab   	- Listview Handler

	admGui_Default("admJournal")

	If (A_EventInfo = 0) && !InStr(A_GuiControl, "AdmButton2")
		return

	LV_GetText(admFile, EventInfo:= A_EventInfo, 1)

	If Instr(A_GuiEvent, "DoubleClick") {
			admGui_View(admFile)
	} else if InStr(A_GuiEvent, "ColClick") {
			admGui_Sort(EventInfo)
	} else if InStr(A_GuiEvent, "RightClick") {
			MouseGetPos, mx, my, mWin, mCtrl, 2
			If (GetHex(mCtrl) = GetHex(hJournal))
				Menu, admJCM, Show, % mx - 20, % my + 10
	} else If InStr(A_GuiControl, "AdmButton2") {
			Loop, % LV_GetCount() {		; Dokumente ohne Patientenbezeichnung nicht importieren
				LV_GetText(admFile, A_Index, 1)
				If !RegExMatch(admFile, "^\s*\d")
					break
			}
			If (StrLen(admFile) = 0)
				return
			If FuzzyKarteikarte(admFile) {
				admGui_ImportFromJournal(admFile)
				SendMessage, 0x1330, 1,,, % "ahk_id " hadmTab ; Journal-Tab aktivieren
			} else {
				PraxTT("off")
				MsgBox, 0, Addendum für Albis on Windows, Es konnte kein passender Patient gefunden werden.
			}
	}


return ;}

admTP:                          	;{ Protokoll Tab 	- Button Handler

	If InStr(A_GuiControl, "admTPZurueck") {
			nil:=0
	} else If InStr(A_GuiControl, "admTPVor") {

	}


return ;}

admTProtokollLV:             	;{ Tagesprotokoll - Listview Handler

	If Instr(A_GuiEvent, "DoubleClick") {
		LV_GetText(PatID, EventInfo:= A_EventInfo, 4)
		AlbisAkteOeffnen("", PatID)
	}

return ;}

admGuiDropFiles:         	;{ nicht erkannte Dokumente einfach in das Posteingangfenster ziehen! klappt genial!

	PatName 	:= StrReplace(AlbisCurrentPatient(), " ")
	PatName 	:= StrReplace(PatName, "-")
	PatName 	:= StrSplit(PatName, ",")

	; übergebene Dateien liegen als `n getrennte Liste in A_GuiEvent vor (genial einfach))
		Loop, Parse, A_GuiEvent, `n
		{
				SplitPath, A_LoopField, filename

			; Datei muss zunächst sicherheitshalber in den Befundordner kopiert werden, wenn sie dort nicht vorliegt
				If !InStr(A_LoopField, Addendum.BefundOrdner )
					FileCopy, % A_LoopField, % Addendum.BefundOrdner "\" filename

				If RegExMatch(filename, "\.pdf")
				{
						PdfReports.Push(StrReplace(filename, ".pdf"))
						;LV_Add("ICON" 1,, RegExReplace(filename, "^[\w\p{L}-]+[\s,]+[\w\p{L}]+\s*\,*\s*", ""))
						LV_Add("ICON" 1,, filename)
				}
				else If RegExMatch(filename, "\.jpg")
				{
						PdfReports.Push(filename)
						LV_Add("ICON" 2,, RegExReplace(filename, "^[\w\p{L}-]+\s*[,]*\s*[\w\p{L}]+\,*\s*", ""))
				}
		}

	; Importbuttons aktivieren und später nach Befundart ausschalten
		admGui_InfoText("admReports")

return ;}

BefundImport:               	;{ läuft wenn Befunde oder Bilder importiert werden sollen

		If (PdfReports.MaxIndex() = 0)                                	; versehentlichen Aufruf bei leerem Array verhindern
			return

		Addendum.Importing := true                             		; globale flag für Importvorgang setzen
		PreReaderListe := admGui_GetAllPdfWin()              	; derzeit geöffnete Readerfenster einlesen
		SetTimer, toptop, off                                               	; Neuzeichnen kurz stoppen
		admGui_ImportGui()                                             	; Hinweisfenster anzeigen
		Imports:= admGui_FileImport()                                	; Importvorgang starten
		admGui_RemoveImports(Imports)                            	; PdfReports - Importe entfernen
		BefundOrdner_Indizieren()	                                    	; Befundordner neu indizieren und Index speichern
		Gui, adm2: Destroy                                                	; Hinweisfenster wieder schliessen
		;admGui_Default("admPdfReports")
		admGui_InfoText("Reports")                                    	; Kurzinfo aktualisieren
		admGui_Journal()                                                 	; Journalinhalt auffrischen
		admGui_ShowPdfWin(PreReaderListe)                    	; holt den/die PdfReaderfenster in den Vordergrund
		SetTimer, toptop, % Addendum.InfoWindow.RefreshTime
		Addendum.Importing := false                                 	; globalen flag für Importvorgang ausschalten

return ;}

toptop:                          	;{ zeichnet das Gui neu

	; Gui schliessen unter bestimmten Bedingungen
		Aktiv := Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		If !RegExMatch(Aktiv, "i)Karteikarte|Laborblatt|Biometriedaten") || !WinExist("ahk_class OptoAppClass") {

				Gui, adm: 	Destroy
				Gui, adm2: 	Destroy
				SetTimer, toptop, Off
				return

		}

	; Fenster ab und zu - neu zeichnen lassen
		If WinExist("ahk_id " hadm) {

				WinSet, Style, 0x50020000, % "ahk_id " hadm
				WinSet, Top                   	,, % "ahk_id " hadm
				WinSet, Redraw             	,, % "ahk_id " hadm

		} else {

			SetTimer, toptop, Off
			return

		}

	; ist Albis inaktiv, dann wird das wiederholte neuzeichnen der Gui zeitlich reduziert
		If !WinActive("ahk_class OptoAppClass")
				SetTimer, toptop, % Addendum.InfoWindow.RefreshTime * 2
		else
				SetTimer, toptop, % Addendum.InfoWindow.RefreshTime

return ;}

admGui_PatLog(PatID, cmd, data) {                                                                                  	; individuelles Patienten Logbuch

		;PatID    	:= AlbisAktuellePatID()
		BaseNum	:= Round((PatID / 1000) ) * 1000
		IDBase  	:= PatID - BaseNum <= 0 ? BaseNum : BaseNum + 1000
		PatientDir 	:= Addendum.DBPath "\PatData\" IDBase "\" PatID
		If !InStr(FileExist(PatientDir), "D")
			FileCreateDir, % PatientDir

		If InStr(cmd, "AddToLog") {

			timeStamp := A_DD "." A_MM "." A_YYYY " " A_Hour ":" A_Min ":" A_Sec "`t"
			FileAppend, % timeStamp . data "`n", % PatientDir "\log.txt"

		}

}

admGui_Reports() {	                                                                                                    	; befüllt die Listview mit Pdf und Bildbefunden aus dem Befundordner

		static PdfImport := false
		static PatName
		global PdfReports
		ImageImport := false

		admGui_Default("admPdfReports")
		LV_Delete()

	; Pdf Befunde des Patienten ermitteln und entfernen bestimmter Zeichen aus dem Patientennamen für die fuzzy Suchfunktion
		;PatName	:= StrReplace(AlbisCurrentPatient(), " ")
		RegExMatch(AlbisCurrentPatient(), "(?<Nachname>[\pL-]+)\,\s(?<Vorname>[\pL-]+)", Pat)
		PatName	:= PatNachname "," PatVorname
		PdfReports	:= ScanPoolArray("Reports", PatName, "")
		PatName 	:= StrReplace(PatName, "-")
		PatName 	:= StrSplit(PatName, ",")

	; Pdf Befunde anzeigen
		Loop, % PdfReports.MaxIndex() 	{
			If FileExist(Addendum.BefundOrdner "\" PdfReports[A_Index] ".pdf")
				LV_Add("ICON" 1,, RegExReplace(PdfReports[A_Index], "^[\pL-]+[\s\,]+[\pL-]+\,*\s*", ""))
		}

	; Bilddateien aus dem Befundordner einlesen
		Loop, Files, % Addendum.BefundOrdner "\*.*"
			If RegExMatch(A_LoopFileName, "\.jpg|png|tiff|bmp|wav|mov|avi$") {

					RegExMatch(A_LoopFileName, "(?<Nachname>[\pL-]+)[\,\s]+(?<Vorname>[\pL-]+)", Such)

					SuchName := SuchNachname SuchVorname
					a	:= StrDiff(SuchName, PatName[1] PatName[2])
					b	:= StrDiff(SuchName, PatName[2] PatName[1])

					If ((a < 0.11) || (b < 0.11)) {

						LV_Add("ICON" 2,, RegExReplace(A_LoopFileName, "^[\pL-]+[\s,]+[\pL-]+\,*\s*", ""))
						PdfReports.Push(A_LoopFileName)

					}
			}

	; InfoText auffrischen
		admGui_InfoText("Reports")

return PdfReports
}

admGui_Journal() {	                                                                                                    	; befüllt die Listview mit Pdf und Bildbefunden aus dem Befundordner

		global Journal := Object()

		ImageImport := false
		PdfImport   	:= false

		admGui_Default("admJournal")
		LV_Delete()

		BefundOrdner_Indizieren()

	; Pdf Befunde des Patienten ermitteln und entfernen bestimmter Zeichen aus dem Patientennamen für die fuzzy Suchfunktion
		For key, val in ScanPool
		{
				BefundName := StrSplit(val, "|").1
				FileGetTime, timeStamp, % Addendum.BefundOrdner "\" BefundName
				FormatTime, filetime  	, % timeStamp, dd.MM.yyyy
				FormatTime, timeStamp	, % timeStamp, yyyyMMdd

				If RegExMatch(BefundName, "i)\.pdf$")
					LV_Add("ICON" 1, BefundName, filetime, timeStamp)

		}

		Loop, Files, % Addendum.BefundOrdner "\*.*"
			If RegExMatch(A_LoopFileName, "\.jpg|png|tiff|bmp|wav|mov|avi$") {

					FileGetTime, timeStamp, % A_LoopFileName
					FormatTime, filetime  	, % timeStamp, dd.MM.yyyy
					FormatTime, timeStamp	, % timeStamp, yyyyMMdd
					SplitPath, A_LoopFileName, BefundName
					LV_Add("ICON" 2, BefundName, filetime, timeStamp)

			}

	; InfoText auffrischen
		admGui_InfoText("Journal")

	; Anzeige wie zuletzt sortiert wurde
		admGui_Sort(0, true)

return
}

admGui_TProtokoll(day:="") {                                                                                           	; zeigt alle aufgerufenen Patienten des Tages an

	global hTProtokoll

	admGui_Default("admTProtokoll")
	LV_Delete()
	TProtMax 	:= TProtokoll.MaxIndex()

	For idx, PatID in TProtokoll
		LV_Insert(1,, SubStr("00" idx, -1) , oPat[PatID].Nn ", " oPat[PatID].Vn, oPat[PatID].Gd, PatID)

	LV_ModifyCol(2, "Auto")
	colWidth := LV_GetColWidth(hTProtokoll, 2)
	If (colWidth < 115)
		LV_ModifyCol(2, 115)

	admGui_InfoText("AdmTProtokollTitel")

}

admGui_InfoText(TabTitel) {                                                                                           	; Zusammenfassungen aktualisieren

	global PdfReports
	global Journal

	Gui, adm: Default

	If InStr(TabTitel, "Journal") {

		InfoText := (ScanPool.MaxIndex() = 0) ? " keine Dokumente" : (ScanPool.MaxIndex() = 1) ? "1 Dokument" : ScanPool.MaxIndex() " Dokumente"
		GuiControl, adm: , AdmJournalTitel, % InfoText

	}

	If InStr(TabTitel, "Reports") {

		InfoText := "Posteingang: (" (!PdfReports.MaxIndex() ? "keine neuen Befunde)" : PdfReports.MaxIndex() = 1 ? "1 Befund)" : PdfReports.MaxIndex() " Befunde)") ", insgesamt: " ScanPool.MaxIndex()
		GuiControl, adm: , AdmPdfReportTitel, % InfoText
		If (PdfReports.MaxIndex() > 0)
			GuiControl, adm: Enable, AdmButton1
		else
			GuiControl, adm: Disable, AdmButton1

	}

	If InStr(TabTitel, "TProtokoll") {

		InfoText := "[" compname "] im Tagesprotokoll: " TProtokoll.MaxIndex() " Patienten"
		GuiControl, adm: , AdmTProtokollTitel, % InfoText

	}

}

admGui_Sort(EventNr, LV_Init:=false)  {                                                                         	; sortiert das Journal

		; Funktion wird gebraucht für die Wiederherstellung der letzten Sortierungseinstellung und für das Sichern der Einstellungen bei Nutzerinteraktion
		; LV_Init - nutzen um gespeicherte Sortierungseinstellung wiederherzustellen

		global 	hJournal
		static 	JCol1Dir, JCol2Dir

		admGui_Default("admJournal")

		If LV_Init {

			RegExMatch(Addendum.InfoWindow.JournalSort, "\s*(\d)\s+(\d)", JSort)

			EventNr := JSort1

			If (JSort1 = 1) {
				If (JSort2 = 1)
					JCol1Dir := true
			}
			else {
				If (JSort2 = 1)
					JCol2Dir := true
			}

			If JCol1Dir
				JCol2Dir := false

			If JCol2Dir
				JCol1Dir := false
		}

		; Sortierung je nach gewählter Spalte vornehmen
		; Idee von: https://www.autohotkey.com/boards/viewtopic.php?t=68777
			If (EventNr = 2) {
				LV_ModifyCol(3, (JCol2Dir := !JCol2Dir) ? "Sort" : "SortDesc")
				LV_SortArrow(hJournal, 2, (JCol2Dir ? "up":"down"))
				LVSortStr := "2 " ((JCol2Dir = true) ? "1":"0")
			}
			else If (EventNr = 1) {
				LV_ModifyCol(1, (JCol1Dir := !JCol1Dir) ? "Sort" : "SortDesc")
				LV_SortArrow(hJournal, 1, (JCol1Dir ? "up":"down"))
				LVSortStr := "1 " ((JCol1Dir = true) ? "1":"0")
			}

	; veränderte Sortierung in der Addendum.ini unter den PC Namen speichern
		If !LV_Init
		{
				Addendum.InfoWindow.JournalSort := LVSortStr
				IniWrite, % LVSortStr, % Addendum.AddendumIni, % compname, % "Infofenster_JournalSortierung"
		}

}

admGui_CM(MenuName)                   {                                                                         	; Kontextmenu des Journal

		; fehlt: nachschauen ob im Befundordner ein Backup-Verzeichnis angelegt ist
		; auch für Tastaturkürzelbefehle in allen Tabs (01.07.2020)

		global hJournal, hPdfReports, admFile, rcEvInfo, hadm
		static newadmFile

		SciTEOutput("A_Gui..: " A_GuiControl ", " A_GuiControlEvent ", " A_GuiEvent)

		If RegExMatch(MenuName, "^J")
				admGui_Default("admJournal")

		If        InStr(MenuName, "JDelete")   	{

				MsgBox,0x1024, Addendum für AlbisOnWindows, % "Wollen Sie die Datei:`n" admFile "`nwirklich löschen?"
				IfMsgBox, No
				{
						PraxTT(	"Der Löschvorgang wurde abgebrochen!", "5 2")
						return
				}

				If FileExist(Addendum.BefundOrdner "\" admFile) 	{

						;zu löschende Datei in das Backup-Verzeichnis verschieben
							FileMove , % Addendum.BefundOrdner "\" admFile, % Addendum.BefundOrdner "\backup\" admFile
							If !ErrorLevel 	{

								If RegExMatch(MenuName, "^J")
									admGui_Default("admJournal")

								Loop % LV_GetCount() 	{
									LV_GetText(rowText, A_Index, 1)
									If InStr(rowText, admFile) {
											LV_Delete(A_Index)
											break
									}
								}

							} else {
									PraxTT(	"Die Datei konnte nicht gelöscht werden,`nda ein anderes Programm diese geöffnet hat.", "5 2")
									return
							}
				}

				; Ordner neu einlesen
				admGui_Journal()

		}
		else if InStr(MenuName, "JRename") 	{

		; testet ob der Schreibzugriff auf das Dokument möglich ist. Die Datei kann ansonsten nicht umbenannt werden.
			If FileExist(Addendum.BefundOrdner "\" admFile) {

				FileMove , % Addendum.BefundOrdner "\" admFile, % Addendum.BefundOrdner "\backup\" admFile
				If ErrorLevel {
					FileMove, % Addendum.BefundOrdner "\backup\" admFile , % Addendum.BefundOrdner "\" admFile
					PraxTT(	"Die Datei kann nicht umbenannt werden,`nda ein anderes Programm diese geöffnet hat.", "5 2")
					return
				}
				FileMove, % Addendum.BefundOrdner "\backup\" admFile , % Addendum.BefundOrdner "\" admFile

				WinGetPos, wx, wy, ww, wh, % "ahk_id " hadm
				x := wx + ww - 405, y := wy + wh + 5

				SplitPath, admFile,,, fileext, fileout

				InputBox, newadmFile, Addendum für AlbisOnWindows, % "Ändern Sie den Dateinamen und drücken Sie 'Ok' oder die Enter Taste.",, 400, 125, % x, % y,,, % fileout
				newadmFile .= "." fileext
				If (newadmFile = admFile) {
						PraxTT("Sie haben den Dateinamen nicht geändert!", "5 2")
						return
				}
				else if (StrLen(newadmFile) = 0) {
						PraxTT("Ihr neuer Dateiname enthält keine Zeichen.`nDie Datei kann deshalb nicht umbenannt werden!", "5 2")
						return
				}

				FileMove, % Addendum.BefundOrdner "\" admFile, % Addendum.BefundOrdner "\" newadmFile, 1
				If ErrorLevel {
						MsgBox, 0x1024, Addendum für Albis on Windows, % "Das Umbenennen der Datei wurde von Windows`nmit folgendem Fehler abgelehnt:`n" A_LastError
						return
				}

				admGui_Journal()
				admGui_Reports()

			}


		}
		else if InStr(MenuName, "JOpen")   	{
			If (StrLen(admFile) > 0)
				FuzzyKarteikarte(admFile)
		}
		else if InStr(MenuName, "JView")     	{
			admGui_View(admFile)
		}
		else if InStr(MenuName, "JRefresh") 	{
			admGui_Journal()
		}
		else if InStr(MenuName, "JImport")  	{
			admGui_ImportFromJournal(admFile)
		}

return
}

admGui_View(filename) {                                                                                              	; Befund-/Bildanzeigeprogramm aufrufen

	If (StrLen(filename) = 0)
		return 0

	If RegExMatch(filename, "\.jpg|png|tiff|bmp|wav|mov|avi$") && FileExist(Addendum.BefundOrdner "\" filename) {
					PraxTT("Anzeigeprogramm wird geöffnet", "8 3")
					Run % q Addendum.BefundOrdner "\" filename q
					PraxTT("", "off")
	}
	else {
					PraxTT("PDF-Anzeigeprogramm wird geöffnet", "8 3")
					Run % Addendum.PDFReaderFullPath " " q Addendum.BefundOrdner "\" filename q
					PraxTT("", "off")
	}

}

admGui_Default(LVName) {                                                                                          	; Gui Listview als Default setzen

	Gui, adm: Default
	Gui, adm: ListView, % LVName

return
}

admGui_ImportGui(ShowGUI=true) {                                                                               	; Hinweisfenster für laufenden Importvorgang

	global adm2, hadm, hadm2, AdmImporter, AdmJournal

	If ShowGUI {

		GuiControlGet, rpos, adm: Pos, AdmPdfReports

		Gui, adm2: New	, +HWNDhadm2 +ToolWindow -Caption +Parent%hadm%
		Gui, adm2: Color	, c880000
		Gui, adm2: Font	, s14 q5 cFFFFFF Bold, % Addendum.StandardFont
		Gui, adm2: Add	, Text, % "x" 0 " y" Floor(rposH//4) " w" rposW " Center vAdmImporter"	, Importvorgang läuft...`nbitte nicht unterbrechen!
		Gui, adm2: Show	, % "x" rposX " y" rposY " w" rposW " h" rposH " Hide"                        	, AdmImportLayer

		GuiControlGet, cpos, adm2: Pos, AdmImporter
		GuiControl, adm2: Move, AdmImporter, % "y" Floor(rposH // 2 - cposH // 2)
		Gui, adm2: Show

		WinSet, Style         	, 0x50020000	, % "ahk_id " hadm2
		WinSet, ExStyle	    	, 0x0802009C	, % "ahk_id " hadm2
		WinSet, Transparent	, 100            	, % "ahk_id " hadm2

	} else {

		Gui, adm2: Destroy

	}

}

admGui_ImportFromJournal(Report) {                                                                            	; einzel Importfunktion für das Journal

	global hPdfReports

	SetTimer, MoveAdmMsgBox, -100
	message := "Wollen Sie die Datei:`n" (StrLen(Report) >30 ? SubStr(Report, 1, 30) "..." : Report) "`nin die aktuelle geöffnete Karteikarte von`n" AlbisCurrentPatient() "`nimportieren ?"
	MsgBox, 0x1024, Addendum für AlbisOnWindows, % message
	IfMsgBox, Yes
	{

			admGui_Default("admJournal")

		; Pdf Befund importieren
			If RegExMatch(report, "\.pdf") 	{
				If (AlbisImportierePdf(report ) != 0) {
						ScanPoolArray("Remove", report)
				}
			}

		; Bild importieren
			If RegExMatch(report, "\.jpg|png|tiff|bmp|wav|mov|avi$")	{
				If (AlbisImportiereBild(report, RegExReplace(report, "^[\w\p{L}]+)[\,\s]+[\w\p{L}]+") != 0)) {
						AlbisActivate(1)
				}
			}

			BefundOrdner_Indizieren()
			admGui_Journal()
			admGui_Reports()

	}

return

MoveAdmMsgBox:  ;{ verschiebt die Messagebox in die Nähe des Importbuttons

		hwnd := WinExist("A")
		m   	 := GetWindowSpot(hwnd)
		p  	 := GetWindowSpot(hPdfReports)

		SetWindowPos(hwnd, p.X + (p.W/2 - m.W/2), p.Y + (p.H/2 - m.H/2) + 105, m.w, m.H)

return ;}
}

admGui_FileImport() {                                                                                                    	; Befundimport für die Befunde eines Patienten

	global PdfReports

	admGui_Default("admPdfReports")                         	; Journallistview als Default setzen

	Loop, % PdfReports.MaxIndex() 		{

				imported	:= ""
				report 	 	:= PdfReports[A_Index]

				If RegExMatch(report, "\.jpg|png|tiff|bmp|wav|mov|avi")	{	; Image importieren

					If (AlbisImportiereBild(report, RegExReplace(report, "^[\w\p{L}]+)[\,\s]+[\w\p{L}]+") = 0)) {
						SciTEOutput("Datei (" Report ") wurde nicht importiert.", 0, 1)
						continue
					} else {
						imported := Report
						AlbisActivate(1)
					}

				}
				else {                                                                                    ; Pdf importieren

					If (AlbisImportierePdf(report ".pdf") = 0) {
						SciTEOutput("Datei (" Report ") wurde nicht importiert.", 0, 1)
						continue
					}	else {
						ScanPoolArray("Remove", report)
						imported := Report
						AlbisActivate(1)
					}

				}

				If (StrLen(imported) > 0)                   			{                        	; importierten Eintrag aus der Listview entfernen

						Imports .= report "`n"
						admGui_Default("admPdfReports")
						Loop, % LV_GetCount() {
							LV_GetText(Text, A_Index, 2)
							If InStr(imported, Text) {
								LV_Delete(A_Index)
								break
							}
						}

				}

		}

return Imports
}

admGui_RemoveImports(ImportList) {                                                                              	; PdfReport-Array aufräumen

	global PdfReports

	Loop, Parse, % RTrim(ImportList, "`n"), `n
		Loop % PdfReports.MaxIndex()
			If InStr(PdfReports[A_Index], A_LoopField) {
				PdfReports.RemoveAt(A_Index)
				continue
			}

}

admGui_GetAllPdfWin() {                                                                                              	; Namen aller angezeigten PdfReader ermitteln

	WinGet, tmpListe, List, % "ahk_class " Addendum.PDFReaderWinClass
	Loop % tmpListe
		PreReaderListe .= tmpListe%A_Index% "`n"

return PreReaderListe
}

admGui_ShowPdfWin(PreReaderListe) {                                                                          	; PdfReaderfenster nach vorne holen

	; funktioniert nicht wenn man nur ein Fenster (eine Instanz) des PdfReader eingestellt hat
	; die Tabs im FoxitReader bekomme ich nicht ermittelt und somit auch nicht angesteuert

	WinGet, tmpListe, List, % "ahk_class " Addendum.PDFReaderWinClass
	Loop % tmpListe
		ReaderListe .= tmpListe%A_Index% "`n"
	ReaderListe := PreReaderListe "`n" ReaderListe
	Sort, ReaderListe, U                              	; entfernt die vor dem Importvorgang geöffneten Readerfenster aus der Aktivierungsliste
	Loop, Parse, ReaderListe, `n
	{
		WinActivate    	, % "ahk_id " A_LoopField
		WinWaitActive	, % "ahk_id " A_LoopField,, 1
	}

}

BefundOrdner_Indizieren() {                                                                                          	; ruft die externe Funktion ScanPoolArray auf

		BefundIndex_Old  	:= ScanPool.MaxIndex()
		BefundIndex_New 	:= ScanPoolArray("Refresh")
		If (BefundIndex_Old <> BefundIndex_New)
				ScanPoolArray("Save")

}

ImportZaehler(creationtime, Report) {                                                                             	; ### zuviele Fehler !!

	RegExMatch(Report, "(?<Name>.*)(?<extension>\.\w+$)", Case)
	RegExMatch(creationtime, "(?<Day>\d\d)\.(?<Month>\d\d)\.(?<Year>\d\d\d\d)", creation)

	If InStr(CaseExtension, "pdf")
	{
			DocName := "Briefe"
			DocKey		:= ScanPoolArray("Find", Report)
			Pages   	:= Trim(StrSplit(ScanPool[DocKey], "|").3)
			Pages     	:= StrLen(Pages) = 0 ? 1 : Pages
			RegExMatch(IniReadExt("ScanPool", "importierte_" DocName "_" creationYear, "0 0S"), "^(?<Count>\d+)\s(?<Pages>\d+)S", Ini)
			ToolTip, % (IniCount +1) " " (IniPages + Pages)
			IniWrite, % (IniCount +1) " " (IniPages + Pages) "S" , % Addendum.AddendumIni, ScanPool, % "importierte_" DocName "_" creationYear
	}
	else
	{
			DocName := "Bilder"
			BefundNr := IniReadExt("ScanPool", "importierte_" DocName "_" creationYear, 0) + 1
			IniWrite, % BefundNr, % Addendum.AddendumIni, ScanPool, % "importierte_" DocName "_" creationYear
	}


}


; ------------------------
LV_GetColWidth(hLV, ColN) {                                                                                           	;-- gets the width of a column

	; from AutoGui
    SendMessage 0x101F, 0, 0,, % "ahk_id " hLV ; LVM_GETHEADER
    hHeader := ErrorLevel
    cbHDITEM := (4 * 6) + (A_PtrSize * 6)
    VarSetCapacity(HDITEM, cbHDITEM, 0)
    NumPut(0x1, HDITEM, 0, "UInt") ; mask (HDI_WIDTH)
    SendMessage, % A_IsUnicode ? 0x120B : 0x1203, ColN - 1, &HDITEM,, % "ahk_id " hHeader ; HDM_GETITEMW

Return (ErrorLevel != "FAIL") ? NumGet(HDITEM, 4, "UInt") : 0
}


