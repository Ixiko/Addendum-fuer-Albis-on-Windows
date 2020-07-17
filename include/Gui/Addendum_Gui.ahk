

; Infofenster
AddendumGui() {

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		global

		local LVPatOpt    	 := " r" (Addendum.InfoWindow.LVScanPool.r)
				LVPatOpt   	 .= " gadmPatLV 	    	vAdmPdfReports 	HWNDhPdfReports 	AltSubmit -Hdr ReadOnly	-LV0x10 BackgroundF0F0F0 E0x20000"
		local LVJournalOpt .= " gadmJournal    	vAdmJournal 		HWNDhJournal 	 	AltSubmit                      	-LV0x10 BackgroundF0F0F0 E0x20200"
		local LVTProtokoll	 := " gadmTProtokollLV	vAdmTProtokoll 	HWNDhTProtokoll 	AltSubmit -Hdr ReadOnly	-LV0x10 BackgroundF0F0F0 E0x20000"

		local LVContent
		local ImageListID, adm_init := false, hbmPdf_Ico, hbmImage_Ico
		local adm, adm2, adm3, admWidth, admHeight, TabSize, TabSizeX, TabSizeY, TabSizeW, TabSizeH
		local APos, SPos, rpos, mox, moy, moWin, moCtrl, regExt, Aktiv, col1W, col2W, col3W, Pat
		local admReCheck := 0

		If !adm_init {

				fn_adm := Func("AddendumGui")

				;hbmPdf_Ico   	:= Create_PDF_ico()
				;hbmImage_Ico	:= Create_Image_ico()

				adm_init        	:= true
				ImageListID    	:= IL_Create(2)
				IL_Add(ImageListID, Addendum.AddendumDir "\assets\ModulIcons\pdf.ico" 	, 0xFFFFFF, 0) 		;1
				IL_Add(ImageListID, Addendum.AddendumDir "\assets\ModulIcons\image.ico", 0xFFFFFF, 0)		;2

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

		}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Vorbereitungen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	; Albis wurde beendet -return-
		If !WinExist("ahk_class OptoAppClass")
			return

	; mehr als 3 Versuche das MDIFrame zu finden führen zum Abbruch des Selbstaufrufes der Funktion
		If (admReCheck > 3) {
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
			SetTimer, % fn_adm, -3000
			admReCheck ++
			return
		}

	; legt die Höhe der Gui anhand der Größe des MDIFrame-Fenster fest
		APos := GetWindowSpot(AlbisWinID())
		SPos := GetWindowSpot(hStamm)

		Addendum.InfoWindowX	:= APos.CW - Addendum.InfoWindow.W ; ClientWidth passt besser!
		admHeight 	:= Addendum.InfoWindow.H	:= SPos.CH
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
			SetTimer, % fn_adm, -5000
			admReCheck ++
			return
		}

		Gui, adm: Margin	, 0, 0
		Gui, adm: Add  	, Tab     	, % "x0 y0  	w" admWidth " h" admHeight " hwndhadmTab vadmTabs"                   	, Patient|Journal|Protokoll|Info

	;-: Tab 1 :- Patient
		Gui, adm: Tab  	, 1
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text      	, % "x10 y27 w" admWidth - 15 " BackgroundTrans vAdmPdfReportTitel Section"       	, (es wird gesucht ....)
		Gui, adm: Font  	, s6 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth-55 " y23 w50 HWNDhAdmButton1 vAdmButton1 gBefundImport" 	, Importieren
		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+2 	w" admWidth - 6	" " LVPatOpt                                                            	, % "S|Befundname"
		LV_SetImageList(ImageListID)
		GuiControlGet, cpos, adm: Pos, AdmPdfReports
		gpHeight := admHeight - cPosX - cposH - 45
		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, GroupBox, % "x2 y+0 	w" admWidth - 6 " h" gpHeight " Section Center vAdmGP1"  	               	, Notizen/Erinnerungen
		GuiControlGet, cpos, adm: Pos, AdmGP1
		Gui, adm: Add  	, Edit     	, % "x5 y" cposY+15 " w" admWidth - 12 " h" cposH - 20 " gAdmNotes vAdmNotes"

	;-: Tab 2 :- Journal
		Gui, adm: Tab  	, 2
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text      	, % "x10 y25  	w" admWidth - 22 " BackgroundTrans vAdmJournalTitel Section"        	, (es wird gesucht ....)
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth-145 " y23 w140 h16 HWNDhAdmButton2 vAdmButton2 gadmJournal"   	, obersten Befund Importieren
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+5 		w" admWidth - 6 " h" admHeight - 55 " " LVJournalOpt                     	, % "Befund|Eingang|TimeStamp"
		LV_SetImageList(ImageListID)
		col2W := 70, col1W := admWidth - col2W, col3W := 0
		LV_ModifyCol(1, col1W " Left NoSort")
		LV_ModifyCol(2, col2W " Left NoSort")
		LV_ModifyCol(3, col3W " Left Integer")              	; versteckte Spalte enthält Integerzeitwerte für die Sortierung des Datums

	;-: Tab 3 :- letzte Patienten
		Gui, adm: Tab  	, 3
		Gui, adm: Add  	, Text      	, % "x10 y25  	w" admWidth - 22 " BackgroundTrans vAdmTProtokollTitel Section"     	, % "[" compname "] im Tagesprotokoll: " TProtokoll.MaxIndex() " Patienten"
		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+5   	w" admWidth - 6	 " h" admHeight - 55 " " LVTProtokoll                     	, % "RF|Nachname|Vorname|Geburtstag|PatID"

	;-: Tab 4 :- Info
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

admNotes:						;{

return
;}

;admTabs:                     	;{

;return ;}

admPatLV:                    	;{ Patienten Tab 	- Listview Handler

	;Critical
	LV_GetText(fname, EventInfo:= A_EventInfo, 1)

  ;PDF Vorschau
	If Instr(A_GuiEvent, "DoubleClick")
			admGui_View(fname)

return ;}

admJournal:                 	;{ Journal Tab

	If InStr(A_GuiControl, "admJournal") {		; Listview

			LV_GetText(fname, EventInfo:= A_EventInfo, 1)

			If Instr(A_GuiEvent, "DoubleClick") {
					admGui_View(fname)
			} else if InStr(A_GuiEvent, "ColClick") {
					admGui_Sort(EventInfo)
			} else if InStr(A_GuiEvent, "RightClick") {
					MouseGetPos, mx, my, mWin, mCtrl, 2
					If (GetHex(mCtrl) = GetHex(hJournal))
						Menu, admJCM, Show, % mx - 20, % my + 10
			}

	}


return ;}

admTProtokollLV:             	;{

	LV_GetText(txt, EventInfo:= A_EventInfo, 1)
	RegExMatch(Txt, "\((?<ID>\d+)\)", Pat)
	;MsgBox, % PatID

	AlbisAkteOeffnen(PatID)

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

		If (PdfReports.MaxIndex() = 0)
			return

	; globale flag für Importvorgang setzen
		Addendum.Importing:= true
		Imports := ""

	; derzeit geöffnete Readerfenster einlesen
		WinGet, tmpListe, List, % "ahk_class " Addendum.PDFReaderWinClass
		Loop % tmpListe
			PreReaderListe .= tmpListe%A_Index% "`n"

	; Neuzeichnen kurz stoppen
		SetTimer, toptop, off

	; -----------------------------------------------------------------------------------------------
	; Hinweisfenster für laufenden Importvorgang einblenden
	; -----------------------------------------------------------------------------------------------;{
		GuiControlGet, rpos, adm: Pos, AdmPdfReports

		Gui, adm2: New, +HWNDhadm2 +ToolWindow -Caption +Parent%hadm%
		Gui, adm2: Color, c880000
		Gui, adm2: Font, s14 q5 cFFFFFF Bold, % Addendum.StandardFont
		Gui, adm2: Add, Text, % "x" 0 " y" Floor(rposH//4) " w" rposW " Center vAdmImporter", Importvorgang läuft...`nbitte nicht unterbrechen!
		Gui, adm2: Show, % "x" rposX " y" rposY " w" rposW " h" rposH " Hide", AdmImportLayer

		GuiControlGet, cpos, adm: Pos, AdmImporter
		cposY := Floor(rposH // 2 - cposH // 2)
		GuiControl, adm: Move, AdmImporter, % "y" cposY
		Gui, adm2: Show

		WinSet, Style  	, 0x50020000, % "ahk_id " hadm2
		WinSet, ExStyle	, 0x0802009C, % "ahk_id " hadm2
		WinSet, Transparent, 100, % "ahk_id " hadm2

		admGui_Default("admPdfReports")
		;}

	; -----------------------------------------------------------------------------------------------
	; Befunde/Bilder importieren
	; -----------------------------------------------------------------------------------------------;{
		Loop, % PdfReports.MaxIndex() 		{

				imported	:= false
				report 	 	:= PdfReports[A_Index]

			; Importvorgang
				If RegExMatch(report, "\.jpg|png|tiff|bmp|wav|mov|avi")	{

					; Image importieren
						If (AlbisImportiereBild(report, RegExReplace(report, "^[\w\p{L}]+)[\,\s]+[\w\p{L}]+") = 0)) {
							SciTEOutput("Datei (" Report ") wurde nicht importiert.", 0, 1)
							continue
						} else {
							imported := Report
							AlbisActivate(1)
						}

				}
				else {

						If (AlbisImportierePdf(report ".pdf") = 0) {
							SciTEOutput("Datei (" Report ") wurde nicht importiert.", 0, 1)
							continue
						}	else {
							ScanPoolArray("Remove", report)
							imported := Report
							AlbisActivate(1)
						}

				}

			; importierten Eintrag aus der Listview entfernen
				If (StrLen(imported) > 0)                   			{

						admGui_Default("admPdfReports")
						Imports .= report "`n"
						ControlGet, LVContent, List,,, % "ahk_id " hPdfReports
						Loop, Parse, LVContent, `n
							If InStr(imported, A_LoopField) {		; If InStr(report, Trim(StrReplace(A_LoopField, ".pdf"))) {
								SciTEOutput(Report ", " A_LoopField, 0, 1)
								LV_Delete(A_Index)
								break
							}

				}

		}
	;}

	; -----------------------------------------------------------------------------------------------
	; PdfReports - Importe entfernen
	; -----------------------------------------------------------------------------------------------;{
		Loop, Parse, % RTrim(Imports, "`n"), `n
		{
				imported := A_LoopField
				Loop % PdfReports.MaxIndex()
					If InStr(PdfReports[A_Index], imported) {
						PdfReports.RemoveAt(A_Index)
							continue
					}
		}
	;}

	; Befundordner neu indizieren und Index speichern
		BefundOrdner_Indizieren()

	; -----------------------------------------------------------------------------------------------
	; Anzeigen auffrischen
	; -----------------------------------------------------------------------------------------------;{
		Gui, adm2: Destroy
		ImageImport := PdfImport	:= false

		admGui_Default("admPdfReports")
		admGui_InfoText("Reports")
		admGui_Journal()

		gosub toptop
		SetTimer, toptop, % Addendum.InfoWindow.RefreshTime

		;}

	; -----------------------------------------------------------------------------------------------
	; alle Dokumente nach vorne holen (für das Lesen sinnvoll)
	; -----------------------------------------------------------------------------------------------;{
		; dieser Code bringt alle Readerfenster in den Vordergrund
		WinGet, tmpListe, List, % "ahk_class " Addendum.PDFReaderWinClass
		Loop % tmpListe
			ReaderListe .= tmpListe%A_Index% "`n"
		ReaderListe := PreReaderListe "`n" ReaderListe
		Sort, ReaderListe, U                              	; entfernt die vor dem Importvorgang geöffneten Readerfenster aus der Aktivierungsliste
		Loop, Parse, ReaderListe, `n
			WinActivate, % "ahk_id " A_LoopField
	;}

	; globalen flag für Importvorgang ausschalten
		Addendum.Importing:= false

return ;}

toptop:                          	;{ zeichnet das Gui neu

	Aktiv := Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
	If !RegExMatch(Aktiv, "i)Karteikarte|Laborblatt|Biometriedaten") || !WinExist("ahk_class OptoAppClass") {

			Gui, adm: 	Destroy
			Gui, adm2: 	Destroy
			SetTimer, toptop, Off
			return

	}

	If WinExist("ahk_id " hadm) {

			WinSet, Style, 0x50020000, % "ahk_id " hadm
			WinSet, AlwaysOnTop, On	 , % "ahk_id " hadm
			WinSet, Top                  	,, % "ahk_id " hadm
			WinSet, Redraw             	,, % "ahk_id " hadm

	} else {

		SetTimer, toptop, Off
		return

	}

	If !WinActive("ahk_class OptoAppClass")
			SetTimer, toptop, % Addendum.InfoWindow.RefreshTime * 2
	else
			SetTimer, toptop, % Addendum.InfoWindow.RefreshTime

return ;}
}

admGui_Reports() {	                                                                                                    	; befüllt die Listview mit Pdf und Bildbefunden aus dem Befundordner

		static PdfImport := false
		static PatName
		global PdfReports
		ImageImport := false

		admGui_Default("admPdfReports")
		LV_Delete()

	; Pdf Befunde des Patienten ermitteln und entfernen bestimmter Zeichen aus dem Patientennamen für die fuzzy Suchfunktion
		PatName	:= StrReplace(AlbisCurrentPatient(), " ")
		PdfReports:= ScanPoolArray("Reports", PatName, "")
		PatName 	:= StrReplace(PatName, "-")
		PatName 	:= StrSplit(PatName, ",")

	; Pdf Befunde anzeigen
		Loop, % PdfReports.MaxIndex() 	{
			If FileExist(Addendum.BefundOrdner "\" PdfReports[A_Index] ".pdf")
				LV_Add("ICON" 1,, RegExReplace(PdfReports[A_Index], "^[\w\p{L}-]+[\s,]+[\w\p{L}]+\,*\s*", ""))
		}

	; Bilddateien aus dem Befundordner einlesen
		Loop, Files, % Addendum.BefundOrdner "\*.*"
			If RegExMatch(A_LoopFileName, "\.jpg|png|tiff|bmp|wav|mov|avi$") {

					RegExMatch(A_LoopFileName, "(?P<Nachname>[\w\p{L}]+)[\,\s]+(?P<Vorname>[\w\p{L}]+)", Such)

					SuchName := SuchNachname SuchVorname
					a	:= StrDiff(SuchName, PatName[1] PatName[2])
					b	:= StrDiff(SuchName, PatName[2] PatName[1])

					If ((a < 0.11) || (b < 0.11)) {

						LV_Add("ICON" 2,, RegExReplace(A_LoopFileName, "^[\w\p{L}-]+[\s,]+[\w\p{L}]+\,*\s*", ""))
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

admGui_TProtokoll() {                                                                                                    	; zeigt alle aufgerufenen Patienten des Tages an

	admGui_Default("admTProtokoll")
	LV_Delete()
	TProtMax 	:= TProtokoll.MaxIndex()

	For idx, PatID in TProtokoll
		LV_Insert(1,, SubStr("00" idx, -1) , oPat[PatID].Nn, oPat[PatID].Vn, oPat[PatID].Gd, "(" PatID ")" )

	LV_ModifyCol()

	admGui_InfoText("AdmTProtokollTitel")

}

admGui_InfoText(TabTitel) {                                                                                            	; Zusammenfassungen aktualisieren

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

admGui_Sort(EventNr, LV_Init:=false)	{                                                                          	; sortiert das Journal

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

admGui_CM(MenuName)                 	{                                                                          	; Kontextmenu des Journal

		; fehlt: nachschauen ob im Befundordner ein Backup-Verzeichnis angelegt ist

		global hJournal, hPdfReports, fname, rcEvInfo, hadm
		static newfname

		If RegExMatch(MenuName, "^J")
				admGui_Default("admJournal")

		If        InStr(MenuName, "JDelete")   	{

				MsgBox,0x1024, Addendum für AlbisOnWindows, % "Wollen Sie die Datei:`n" fname "`nwirklich löschen?"
				IfMsgBox, No
				{
						PraxTT(	"Der Löschvorgang wurde abgebrochen!", "5 2")
						return
				}

				If FileExist(Addendum.BefundOrdner "\" fname) 	{

						;zu löschende Datei in das Backup-Verzeichnis verschieben
							FileMove , % Addendum.BefundOrdner "\" fname, % Addendum.BefundOrdner "\backup\" fname
							If !ErrorLevel 	{

								If RegExMatch(MenuName, "^J")
									admGui_Default("admJournal")

								Loop % LV_GetCount() 	{
									LV_GetText(rowText, A_Index, 1)
									If InStr(rowText, fname) {
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

		; testet durch FileMove ob die Datei evtl. durch den Zugriff eines anderen Programmes blockiert ist
			If FileExist(Addendum.BefundOrdner "\" fname) {

				FileMove , % Addendum.BefundOrdner "\" fname, % Addendum.BefundOrdner "\backup\" fname
				If ErrorLevel {
					FileMove, % Addendum.BefundOrdner "\backup\" fname , % Addendum.BefundOrdner "\" fname
					PraxTT(	"Die Datei kann nicht umbenannt werden,`nda ein anderes Programm diese geöffnet hat.", "5 2")
					return
				}
				FileMove, % Addendum.BefundOrdner "\backup\" fname , % Addendum.BefundOrdner "\" fname

				WinGetPos, wx, wy, ww, wh, % "ahk_id " hadm
				x := wx + ww - 405
				y := wy + wh + 5
				SplitPath, fname,,, fileext, fileout

				InputBox, newfname, Addendum für AlbisOnWindows, % "Ändern Sie den Dateinamen und drücken Sie 'Ok' oder die Enter Taste.",, 400, 125, % x, % y,,, % fileout
				newfname .= "." fileext
				If (newfname = fname) {
						PraxTT("Sie haben den Dateinamen nicht geändert!", "5 2")
						return
				}
				else if (StrLen(newfname) = 0) {
						PraxTT("Ihr neuer Dateiname enthält 0 Zeichen.`nDie Datei kann deshalb nicht umbenannt werden!", "5 2")
						return
				}

				FileMove, % Addendum.BefundOrdner "\" fname, % Addendum.BefundOrdner "\" newfname, 1
				If ErrorLevel {
						MsgBox, 0x1024, Addendum für Albis on Windows, % "Das Umbenennen der Datei wurde von Windows`nmit folgender Fehlerangabe abgelehnt:`n" A_LastError
						return
				}

				admGui_Journal()
				admGui_Reports()

			}


		}
		else if InStr(MenuName, "JOpen")   	{

			If (StrLen(fname) > 0)
				FuzzyKarteikarte(fname)

		}
		else if InStr(MenuName, "JView")     	{

			admGui_View(fname)

		}
		else if InStr(MenuName, "JRefresh") 	{

			admGui_Journal()

		}
		else if InStr(MenuName, "JImport")  	{

			Report := fname
			MsgBox, 0x1024, Addendum für AlbisOnWindows, % "Wollen Sie die Datei:`n" (StrLen(Report) >30 ? SubStr(Report, 1, 30) "..." : Report) "`nin die aktuelle geöffnete Karteikarte von`n" AlbisCurrentPatient() "`nimportieren ?"
			IfMsgBox, Yes
			{
				; Pdf Befund importieren
					If RegExMatch(report, "\.pdf") 	{
						If (AlbisImportierePdf(report ) != 0) {
								admGui_Default("admJournal")

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

BefundOrdner_Indizieren() {                                                                                          	; ruft die externe Funktion ScanPoolArray auf

		BefundIndex_Old  	:= ScanPool.MaxIndex()
		BefundIndex_New 	:= ScanPoolArray("Refresh")
		If (BefundIndex_Old <> BefundIndex_New)
				ScanPoolArray("Save")

}

ImportZaehler(creationtime, Report) {

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
