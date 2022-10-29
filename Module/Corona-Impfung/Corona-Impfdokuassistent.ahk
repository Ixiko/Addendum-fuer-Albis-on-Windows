; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                    Addendum Assistent für die Corona Impfdokumentation
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      Funktion: 				 ⚬	Erstellungen aller Formulare
;                                	 ⚬	Eintragen aller Abrechnungsziffern, Kasse sowie Privat
;                                	 ⚬	Eintragen der Impfdiagnose
;
;
;		Hinweis:
;
;		Abhängigkeiten:		siehe includes
;
;      	begonnen:       	    	01.05.2021
; 		letzte Änderung:	 	02.05.2021
;
;	  	Addendum für Albis on Windows by Ixiko started in September 2017
;      	- this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣

  ; Einstellungen
	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines, -1
	ListLines    	, Off

	global PatDB            	; Patientendatenbank
	global Addendum
	global CIProps         	; Gui Einstellungen und anderes

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

  ; Tray Icon erstellen
	If (hIconImpfen := Create_Impfen_ico())
    	Menu, Tray, Icon, % "hIcon: " hIconImpfen

  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	Addendum                 	:= Object()
	Addendum.Dir            	:= AddendumDir
	Addendum.Ini              	:= AddendumDir "\Addendum.ini"
	Addendum.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	Addendum.DBasePath  	:= AddendumDir "\logs'n'data\_DB\DBase"
	Addendum.AlbisDBPath	:= AlbisPath "\DB"
	Addendum.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft
	Addendum.propsPath 	:= A_ScriptDir "\Corona-Impfdoku.json"

	SciTEOutput()

  ; Patientendatenbank
	outfilter 	:= ["NR", "PRIVAT", "ANREDE", "NAME", "VORNAME", "GEBURT", "PLZ", "ORT", "STRASSE", "TELEFON", "GESCHL", "TELEFON2", "TELEFAX", "LAST_BEH", "DEL_DATE", "MORTAL", "HAUSNUMMER"]
	PatDB 		:= ReadPatientDBF(Addendum.AlbisDBPath,, outfilter)

  ; Gui und notwendige Daten laden z.B. die Chargennummern
	If FileExist(Addendum.propsPath)
		CIProps := JSONData.Load(Addendum.propsPath, "", "UTF-8")
	else {
		CIProps := {"CNR":[]}
		; Beispielbefüllung
		CIProps.CNR[14] := ["ET3045",""]
		CIProps.CNR[15] := ["EX3510", "ABW7189"]
		CIProps.CNR[16] := ["EX3510",""]
	}

  ; Gui ausführen
	CoronaImpfung()


return


CoronaImpfung() {

	; VARIABLEN                                                          ;{
		global CI, hCI, Impflinge
		global KK := {"Datum":"Edit5", "Kuerzel":"Edit6", "Inhalt":"RichEdit20A1"}

		global hCtrlActive, CITL1, CITL2, CIDate, CIBTN, CIBTE, CIBT1, CIBT2, CIAZN, CIAZE, CIAZ1, CIAZ2
		global CIPGS1, CIPGS2, CIPGS3, CIPGS4

		static PatID, PatName
		static TitleW := 600, col1W := 190, chNrW := 14*11
		static cpX, cpY, cpW, cpH, dpX, dpY, dpW, dpH, cx, cy, PG2Y

		static tr          	:= "Right "
		static tc         	:= "Center "
		static bgt       	:= "BackgroundTrans ", OCW := "cWhite", OCB := "cBlack"
		static gH       	:= " gCIButtonHandler "

		static CNRListe	:= []
		static btn          	:= ["CIDate", "CIBTN", "CIBTE", "CIBT1", "CIBT2", "CIAZN", "CIAZE", "CIAZ1", "CIAZ2"]
		static disButtons	:= ["CIBT", "CIAZ"]
		static CNRCtrls 	:= ["CIBTE", "CIAZE"]

	;}

	; aktuellen Fokus des Albisfenster merken
		ControlGetFocus, hCtrlActive, ahk_class OptoAppClass

	; akutelle Kalenderwoche berechnen
		Tagesdatum := A_DD "." A_MM "." A_YYYY
		If RegExMatch(Tagesdatum, "(?<day>\d\d)\.(?<month>\d\d)\.(?<year>\d\d\d\d)", _) {
			guiDate	:= _year _month _day
			weekNr	:= SubStr("0" WeekOfYear(guiDate), -1)
		}

	; ComboxStrings erstellen
		For CNRWeek, VacList in CIProps.CNR {
			For VacNr, CNRs in VacList
				For idx, CNR in StrSplit(CNRs, "|")
					CNRListe[VacNr] .= CNRWeek ": " CNR "|"
		}
		Loop % CIProps.CNR.Count() {
			liste := RTrim(CNRListe[A_Index], "|")
			Sort, liste, D| U N R
			CNRListe[A_Index] := liste
		}

	; aktuelle Daten vom Albisfenster
		PatID     	:= AlbisAktuellePatID()
		PatName 	:= AlbisCurrentPatient()

	; Nachfragen bei als verstorben markierten Patienten
		If PatDB[PatID].MORTAL {
			MsgBox, 0x1024, % RegExReplace(A_ScriptName, "\.*$"), % "Möchten Sie wirklich bei einem`nverstorbenen Patienten eine Impfdokumentation anlegen?"
			IfMsgBox, No
				ExitApp
		}

	; Impfliste aus Albisdatenbankeinträgen erstellen
		Impflinge	:= new ImpfListe
		ImpfNr  	:= Impflinge.Vaccinations(PatID)
		ImpfTag	:= Impflinge.VaccinationDates(PatID)
		Impfstoff	:= Impflinge.Vaccine(PatID)
		Impfzahl	:= Impflinge.VaccinationsCount()


	; GUI ZEICHNEN                                                	;{

	;-: Steuerelemente ;{
		Gui, CI: New 	, -DPIScale +HWNDhCI +AlwaysOnTop
		Gui, CI: Margin	, 0, 0
		Gui, CI: Font		, % "s13 cWhite", % "Futura Md Bt"
		Gui, CI: Add  	, Progress	, % "x0 y0 w" TitleW+10 " h50 C007300 vCIPGS1 "                 	, 100
		Gui, CI: Add  	, Text    	, % "x5 y5 w" TitleW " " tc bgt " vCITL1"                                      	, % "IMPFUNG GEGEN COVID-19"
		Gui, CI: Font		, % "s12 cWhite", % "Futura Bk Bt"
		Gui, CI: Add  	, Text    	, % "y+0 w" TitleW " " tc bgt " vCITL2"                                        	, % "[" weekNr ". Kalenderwoche " _year "]"
		GuiControlGet, cp, CI: Pos, CIPGS1
		PG2Y := cpY + cpH

		Gui, CI: Add  	, Progress	, % "x0 y" PG2Y " w" TitleW+10 " h60 C0CF09C vCIPGS2 "     	, 100
		Gui, CI: Font		, % "s12 cBlack"	, % "Futura Bk Bt"
		Gui, CI: Add  	, Text    	, % "x5 y" PG2Y+5 " w" col1W  " " tr bgt                                  	, % "Patient:"
		Gui, CI: Add  	, Text    	, % "x+5 " bgT                                                                       	, % "[" PatID "] " PatName ",  Impfungen: " ImpfNr
		Gui, CI: Font		, % "s11 cBlack"	, % "Futura Bk Bt"
		Gui, CI: Add  	, Text    	, % "x5 y+5 w" col1W  " " tr bgt                                                	, % "Tagesdatum:"
		Gui, CI: Add  	, Edit	    	, % "x+5 yp-3 w100 Center vCIDate gCITagesdatum "  bgt       	, % Tagesdatum
		GuiControlGet, cp, CI: Pos, CIPGS2
		PG2Y := cpY+cpH

		Gui, CI: Add  	, Progress    	, % "x0 y" PG2Y " w" TitleW+10 " h95 C16A8D9 vCIPGS3 " 	, 100
		Gui, CI: Font		, % "s11"    	, % "Futura Bk Bt"
		Gui, CI: Add  	, Text        	, % "x5 y" PG2Y+20 " w" col1W " vCIBTN " tr bgt                 	, % "Biontech [COMIRNATY Ch.:"
		GuiControlGet, cp, CI: Pos, CIBTN
		Gui, CI: Add  	, ComboBox 	, % "x+5 y" cpY-3 " w" chNrW " vCIBTE"                              	, % CNRListe.1
		GuiControl, CI: Choose, CIBTE, 1
		Gui, CI: Add  	, Text        	, % "x+2 y" cpY " " bgt                                                         	, % "]"
		Gui, CI: Add  	, Button     	, % "x+5 y" cpY-6 " vCIBT1" gH tc                                          	, % "1.Impfdosis      "
		Gui, CI: Add  	, Button     	, % "x+5 y" cpY-6 " vCIBT2" gH tc                                          	, % "2.Impfdosis      "

		Gui, CI: Add  	, Text        	, % "x5 y+10 w" col1W " vCIAZN "	tr bgt                            	, % "AstraZeneca [Astrazid Ch.:"
		GuiControlGet, cp, CI: Pos, CIAZN
		Gui, CI: Add  	, ComboBox 	, % "x+5 y" cpY-3 " w" chNrW " vCIAZE"                                 	, % CNRListe.2
		GuiControl, CI: Choose, CIAZE, 1
		Gui, CI: Add  	, Text        	, % "x+2 " bgt                                                                     	, % "]"
		Gui, CI: Add  	, Button     	, % "x+5 y" cpY-6 " vCIAZ1" gH tc                                          	, % "1.Impfdosis      "
		Gui, CI: Add  	, Button     	, % "x+5 y" cpY-6 " vCIAZ2" gH tc                                          	, % "2.Impfdosis      "
		GuiControlGet, cp, CI: Pos, CIPGS3
		PG2Y := cpY+cpH

		Gui, CI: Add  	, Progress    	, % "x0 y" PG2Y " w" TitleW+10 " h35 C0E6D8C vCIPGS4 " 	, 100
		Gui, CI: Font		, % "s11 cWhite"	, % "Segoe UI"
		Gui, CI: Add   	, Text            , % "x5 y" PG2Y+5 " w" col1W " " tr bgt                              		, % "durchgeführte Impfungen:"
		Gui, CI: Add   	, Text            , % "x+5 " bgt                                                                 		, % Impfzahl
		;}

	;-: Steuerelementinteraktionen einschränken wenn schon geimpft oder der Impfabstand zu gering ist
		Loop % ImpfNr {

			Nr := A_Index
			For btnIndex, btnName in disButtons {

				; wurde geimpft
					GuiControl, CI: Disable, % (btnName . Nr)
					If (btnIndex = Impfstoff[Nr])
						GuiControl, CI: Text, % (btnName . Nr), % ConvertDBASEDate(ImpfTag[Nr])
					else
						GuiControl, CI: Hide, % (btnName . Nr)

			}

		}

	;-: Wiederholungsimpfungen
		gosub CITagesdatum

		Gui, CI: Show	, AutoSize,	 Impfassistent

	; Progresselemente stören die Interaktion mit anderen Elementen. Alle nicht Progresselemente werden nach vorn geholt oder aktiviert ;{
		WinGet, CList, ControlListHwnd, % "ahk_id " hCI
		For idx, hwnd in StrSplit(CList, "`n") {
			WinGetClass, classnn, % "ahk_id " hwnd
			If !RegExMatch(classnn, "(msctls_progress)") {
				WinSet, Top    	,, % "ahk_id " hwnd
			}
		}
		Sleep 200
		For idx, hwnd in StrSplit(CList, "`n") {
			WinGetClass, classnn, % "ahk_id " hwnd
			If RegExMatch(classnn, "ComboBox")
				SendMessage, 0x0007,,,, % "ahk_id " hwnd
		}
		;}

	;}

return

CIButtonHandler:

return

CITagesdatum:  	;{

		Gui, CI: Submit, NoHide

	; wird nur ausgeführt wenn das Datumformat stimmt
		If !RegExMatch(CIDate, "(?<day>\d\d)\.(?<month>\d\d)\.(?<year>\d\d\d\d)", _)
			return

	; Datum umformatieren und damit das Wochendatum bestimmen
		guiDate	:= _year _month _day
		weekNr	:= SubStr("0" WeekOfYear(guiDate), -1)
		GuiControl, CI:, CITL2, % "[" weekNr ". Kalenderwoche " _year "]"

	; die Chargennummern werden anhand der Wochenkennung ausgewählt
		For ctrlIndex, vctrl in CNRCtrls {
			cbox	:= GuiComboBoxPos("CI", vctrl, "^" weekNr ":")
			GuiControl, CI: Choose, % vctrl, % cbox.Pos
		}

	; Buttons 2.Impfdosis werden aktiviert oder inaktiviert
		If (ImpfNr > 0) {

			ImpfIntervalle 	:= Impflinge.VaccinationIntervals(PatID, guiDate)
			ImpfWieder		:= Impflinge.NextVaccinationDate(PatID)

			For btnIndex, btnName in disButtons {

				iv := LTrim(ImpfIntervalle[btnIndex], "#") ; # Array speichert weder 0 noch negative Zahlen
				;SciTEOutput("iv: " iv ", d: " ImpfWieder[btnIndex])
				If (iv = 0)   {                                 	; kann geimpft werden
					GuiControl, CI: Text   	, % (btnName . "2"), % "2.Impfdosis"
					GuiControl, CI: Enable	, % (btnName . "2")
				}
				else if (iv < 0) {                             	; sollte noch gewartet werden
					GuiControl, CI: Text   	, % (btnName . "2"), % "ab " ConvertDBASEDate(ImpfWieder[btnIndex])
					GuiControl, CI: Disable	, % (btnName . "2")
				}
				else if (iv > 0) {                             	; Impfzeitraum überschritten
					GuiControl, CI: Text   	, % (btnName . "2"), % "2.Impfdosis +" iv
					GuiControl, CI: Enable	, % (btnName . "2")
				}

			}

		}
		else if (ImpfNr = 0) {

			For btnIndex, btnName in disButtons
				GuiControl, CI: Disable	, % (btnName . "2")

		}

return ;}

CIGuiClose:      	;{
CIGuiEscape:
	Gui, CI: Destroy
	JSONData.Save(Addendum.propsPath, CIProps, true,, 1, "UTF-8")
	ExitApp
return ;}

CIMakro:          	;{

	ControlFocus, % KK.Kuerzel, ahk_class OptoAppClass
	VerifiedSetText(KK.Kuerzel, "dia", "ahk_class OptoAppClass")
	ControlFocus, % KK.Inhalt, ahk_class OptoAppClass
	VerifiedSetText(KK.Inhalt, , "ahk_class OptoAppClass")
	SendInput, {Tab}

return ;}
}

class ImpfListe {                                                             	;	erleichtert das Handling einer Impfliste für Corona Impfungen

		__New()                                                    	{			;	Anlegen einer Impfling-Liste

			this.Vaccines 	:= {"#88331":"1", "#88332":"2", "Biontech":"88331"}  ; 1 = Biontech , 2 = AstraZeneca
			this.VacPeriods	:= {1:"6", 2:"8-12"}
			this.Impflinge 	:= this.CoronaImpfungen()

		}

		CoronaImpfungen()                                   	{        	;  Daten aller Geimpften einlesen

		starttime    	:= A_TickCount
		aDB          	:= new AlbisDB(Addendum.AlbisDBPath)
		BefundIndex	:= aDB.IndexRead("BEFUND",, false)
		recordStart  	:= BefundIndex["#20212"].2
		CImpfungen 	:= aDB.GetDBFData("BEFUND", ".*lko.*(88331|88332|88333|88334)", ["PATNR", "DATUM", "KUERZEL", "INHALT", "removed"], recordStart, 0)
								;, {"KUERZEL"	: "rx:.*lko", "INHALT" : "rx:.*(88331|88332|88333|88334)"}, recordStart, 1, "callback=dBaseProgress")

		; nur die Patienten-ID's zusammentragen
			this.Impflinge := Object()
			EBMziffer := 0
			For idx, set in CImpfungen {

				If set.removed {
					continue
				}

				If InStr(set.INHALT, "88331")
					EBMziffer ++

				If RegExMatch(set.KUERZEL, "i)lko") {
					Impfstatus := RegExMatch(set.INHALT, "(A|G|V)") ? 1 : 2
					If RegExMatch(set.INHALT, "i)(88331|88332|88333|88334)[A-Z]", EBM) {
						If !this.Impflinge.haskey(set.PATNR)
							this.Impflinge[set.PATNR] := {"Impftage":{}}
						If !this.Impflinge[set.PATNR].haskey(set.DATUM)
							this.Impflinge[set.PATNR].ImpfTage[set.DATUM] := {}
						this.Impflinge[set.PATNR].Status := ImpfStatus
						this.Impflinge[set.PATNR].ImpfTage[set.DATUM].gebuehr := EBM
					}
				}
				else if RegExMatch(set.KUERZEL, "i)dia") {
					If RegExMatch(set.INHALT, "i)\s+Ch.*\s*(?<NR>\w+)\s*\{", C) {
						If !this.Impflinge.haskey(set.PATNR)
							this.Impflinge[set.PATNR] := {"Impftage":{}}
						If !this.Impflinge[set.PATNR].haskey(set.DATUM)
							this.Impflinge[set.PATNR].ImpfTage[set.DATUM] := {}
						this.Impflinge[set.PATNR].ImpfTage[set.DATUM].CNR := CNR
					}
				}
			}

		SciTEOutput("time: " Round((A_TickCount-starttime)/1000, 2) "s")

	return this.Impflinge
	}


	; Objektinformationen
		VaccinationsStatistic()                                  	{         	;	erstellt eine Gesamtstatistik aller Impfungen

			If IsObject(this.stats)
				return this.stats

			this.stats := {"countAll":0, "vaccines":[0,0,0,0], "days":0, "weeks":0, "months":0}
			For PatID, Impf in this.Impflinge {
				this.stats.countAll += Impf.Status
			}


		return this.stats
		}

		VaccinationsCount() 		                        		{			; 	Anzahl der Impfungen bisher

			stats := this.VaccinationsStatistic()

		return stats.countAll
		}


	; Patienteninformationen
		PatExist(PatID)                                           	{   		;	PatID ist in der Impfliste

			If this.Impflinge.haskey(PatID)
				return true

		return false
		}

		Vaccinations(PatID)                                    	{			;	wieviele Impfungen hat der Patient erhalten

			If !this.PatExist(PatID)
				return 0

		return this.Impflinge[PatID].Status
		}

		VaccinationDates(PatID)                             	{			;	gibt die Tage zurück an denen geimpft wurde

			If !this.PatExist(PatID)
				return 0

			VacDays := []
			For VacDate, data in this.Impflinge[PatID].ImpfTage
				VacDays.Push(VacDate)

		return VacDays
		}

		Vaccine(PatID)                                           	{			;	gibt die/den Impfstoffnamen die geimpft wurden zurück

			If !this.PatExist(PatID)
				return 0

			Vacs := []
			For VacDate, data in this.Impflinge[PatID].ImpfTage {
				RegExMatch(data.gebuehr, "\d+", gebuehr)
				Vacs.Push(this.Vaccines["#" gebuehr])
			}

		return Vacs
		}


	; Berechnungen
		NextVaccinationDate(PatID)                         	{			;	Rückgabe des frühstmöglichen Datums der Wiederholungsimpfung nach Impfstoff

			FirstVacDate 	:= this.VaccinationDates(PatID).1
			nxtVDays   	:= []
			For vaccineNr, period in this.VacPeriods {
				RegExMatch(period, "(?<from>\d+)\-*(?<till>\d+)*", W)
				nextdate := RTrim(DateAdd(FirstVacDate, Wfrom*7, "days"), "0")
				nxtVDays.Push(nextdate)
			}

		return nxtVDays
		}

		VaccinationIntervals(PatID, DateNow)          	{			;	eingestelltes Datum wird mit Impfintervall verglichen

			; Rückgabeparameter: 	0 				= kann geimpft werden
			; 									kleiner 0 	= Anzahl der Wochen bis zur nächsten Impfung
			; 									größer 0 	= Anzahl der Wochen die der Impfzeitraum überschritten wurde

			FirstVacDate 	:= this.VaccinationDates(PatID).1
			If RegExMatch(DateNow, "\d{2}\.\d{2}\.\d{4}")
				DateNow 	:= ConvertToDBASEDate(DateNow)

			;SciTEOutput("fd: " FirstVacDate "`ntd: " DateNow )
			intervals := []
			For vaccineNr, period in this.VacPeriods {

				RegExMatch(period, "(?<from>\d+)\-*(?<till>\d+)*", W)
				days 	:= DaysBetween(FirstVacDate, DateNow)
				weeks 	:= Floor(days/7)
				;SciTEOutput("nr: " vaccineNr ", d: " days ", w: " weeks)

				If (weeks < Wfrom)
					intervals.Push("#" -1*(Wfrom-weeks))
				else if (weeks >= Wfrom) {

					If !WTill
						intervals.Push("#0")
					else if WTill && (weeks <= WTill)
						intervals.Push("#0")
					else if WTill && (weeks > WTill)
						intervals.Push("#" WTill - weeks)

				}

			}

		return intervals
		}

}

dBaseProgress(index, maxIndex, len, matchcount)	{

	global	DBSearchW
	;prgPos := Floor((index*100)/maxIndex)
	prgPos := "Datensätze gefunden: " matchcount
	ToolTip, % prgPos
	SetTimer, dBaseProgressOff, -2000

return
dBaseProgressOff:
	ToolTip
return
}

ShowImpflinge(Impflinge) {

	For PatID, impfdaten in Impflinge {
		tage := ""
		For DATUM, o in Impfdaten
			tage := ConvertDBASEDate(DATUM) " EBM: " o.EBM ", "
		SciTEOutput(" [" PatID "] " PatDB[PatID].NAme ", " PatDB[PatID].Vorname "`t Status: " Impfdaten.Status ". | " RTrim(Tage, ", "))
	}
	SciTEOutput(Impflinge.Count())

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
GuiComboBoxPos(guiname, vctrl, string) {

	global hCtrlActive, CITL, CIDate, CIBTN, CIBTE, CIBT1, CIBT2, CIAZN, CIAZE, CIAZ1, CIAZ2

	cbox := {}
	cbox.pos := 0

	Gui, % guiname ": Default"
	GuiControlGet, chwnd, % guiname ": Hwnd", % vctrl
	ControlGet, CBList, List,,, % "ahk_id " chwnd
	For cbPos, content in StrSplit(CBList, "`n") {
		cbox.list .= (cbPos > 1 ? "|" : "") content
		If RegExMatch(content, string)
			cbox.pos := cbPos
	}

return cbox
}

Create_Impfen_ico() {

VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fT09P////////////Pz88AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fT09P////////////Pz88AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADU1NT////////////Nzc0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACHh4fm5ub////////Nzc0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjo7j4+PLy8sAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB///////8AAD///////wAAn///////AACP//////8AAMP//////wAA4P//////AADwP/////8AAPgOH////wAA+AQP////AAD8AAf///8AAPwAA////wAA/AAB////AAD+AAD///8AAP8AAH///wAA/wAAP///AAD+AAAf//8AAPwAAA///wAA/AAAB///AAD8AAAD//8AAP4AAAH//wAA/wAAAP//AAD/gAAAf/8AAP/AAAA//wAA/+AAAB//AAD/8AAAD/8AAP/4AAAH/wAA//wAAAP/AAD//gAAAf8AAP//AAAA/wAA//+AAAD/AAD//8AAAP8AAP//4AAB/wAA///wAAP/AAD///gAB/8AAP///AAH/wAA///+AAP/AAD///8AAfAAAP///4BA4AAA////wOBAAAD////h8AEAAP/////4AwAA//////wHAAD//////A8AAP/////4HwAA//////A/AAD/////8H8AAP/////g/wAA//////H/AAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -1
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -2
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("gdiplus\GdipCreateHICONFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hIcon)
DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hIcon
}
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Includes
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DATUM.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PDFHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

#Include %A_ScriptDir%\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}




