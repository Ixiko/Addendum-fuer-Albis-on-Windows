; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 		Addendum Laborjournal
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;      Funktion:           	- 	schneller Überblick über die Laborwerte der letzten Tage in Tabellenform mit farblicher Hervorhebung als Webanwendung.
;									- 	nur klinisch relevante Werte kommen zur Anzeige
;									- 	gesondert zu behandelnde Laborparameter können angegeben werden
;										Kategorien: 	immer   	 :	wenn pathologisch anzeigen
;                                                                             	Parameter die unabhängig von der Höhe ihrer Über- oder Unterschreitung des Grenzwertes angezeigt werden
;                                                                           	z.B. Na oder Kalium
;															exklusiv 	 :	werden stets angezeigt, da die klinische Relevanz unabhängig vom Ergebniswert ist
;                                                                            	z.B. COVID-PCR, HIV, Troponine oder pathologische Leukozytenformen
;															nie			 :	Parameter welche nie angezeigt werden sollen
;                                                                             	z.B. Quick (da der INR relevanter ist)
;                                                       	LaborTexte :	Parameter, deren Ergebnisse in der LABTEXT.dbf gespeichert werden, da diese mehrere Ergebnisse enthalten
;                                                                             	z.B. bei Untersuchungen auf Erreger und Resistenz
;                                                                             	Da jedes Labor andere Texte verschicken wird, bleibt nur die Verwendung von individuellen RegEx-Zeichenketten,
;                                                                           	um hier ohne Anpassung des Skriptes Daten extrahieren und weiterverarbeiten zu können.
;                                                                              	Die Regeln werden bei jedem Skriptstart aus der Addendum.ini gelesen und automatisch angewendet.
;                                                                              	[Funktion ist mit Daten eines Fremdlabores bisher nicht getestet]
;									- 	hat 3 Filtermodi implementiert. Durch die Modi schaltet man sich mit einem Klick auf das Corona-Virus:
;										Modus 1 :	alle Ergebnisse
;										Modus 2 :	nur Corona-Abstrich-/Blutergebnisse
;	                               		Modus 3 : alle Ergebnisse außer jene zu Corona
;
;
;
;		Hinweis:				Nichts ist ohne Fehler. Auch dies hier sicher nicht. Also nicht auf die Ausgaben verlassen! Das Skript soll unterstützen und Denken nicht ersetzen!
;									INR Werte werden angezeigt, wenn ein Pat. Falithrom/Marcumar nimmt, der Wert aber nicht als therapeutischer INR gespeichert wurde.
;									Im Moment müssen Einstellungen noch direkt im Skript oder in der Addendum.ini vorgenommen werden, wie z.B. die gesondert zu
;                                  	behandelnden Laborparameter
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 29.06.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Skripteinstellungen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	#NoEnv
	#Persistent
	#SingleInstance               	, Force
	#KeyHistory                  	, Off

	SetBatchLines                	, -1
	ListLines                         	, Off

	SetKeyDelay, -1, -1
	SetMouseDelay, -1
	SetDefaultMouseSpeed, 0
	SetWinDelay, -1
	SetControlDelay, -1
	SendMode Input


  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Variablen / Einstellungen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	global LabJ := Object()
	global adm := Object()
	global load_Bar, percent
	global PatDB
	global neutron
	global Warnen
	global q := Chr(0x22)

	adm.Dir            	:= AddendumDir
	adm.Ini              	:= AddendumDir "\Addendum.ini"
	adm.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	adm.LabDBPath   	:= AddendumDir "\logs'n'data\_DB\LaborDaten"
	adm.AlbisDBPath 	:= AlbisPath "\DB"

	adm.iconpath      	:= adm.Dir "\assets\ModulIcons\Laborjournal2.svg"

	adm.PathLabj     	:= A_Temp "\labjournal_" A_YYYY "-" A_MM "-" A_DD ".html"
	adm.LabPat      	:= A_Temp "\labPat_" A_YYYY "-" A_MM "-" A_DD ".json"
	adm.LabTexte      	:= A_Temp "\labTexte_" A_YYYY "-" A_MM "-" A_DD ".json"
	adm.COVIDStats  	:= A_Temp "\COVIDStats_" A_YYYY "-" A_MM "-" A_DD ".json"

	adm.compname	:= StrReplace(A_ComputerName, "-")                                                    	; der Name des Computer auf dem das Skript läuft

	workini := IniReadExt(adm.Ini)
	If !Instr(FIleExist(adm.AlbisDBPath), "D")
		adm.AlbisDBPath 	:= IniReadExt("Albis", "AlbisWorkDir") "\Db"

	cJSON.EscapeUnicode := "UTF-8"

  ; Tray Icon erstellen
	hIconLabJournal	:= Laborjournal_ico()
   	Menu, Tray, Icon, % (hIconLabJournal ? "hIcon: " hIconLabJournal : adm.Dir "\assets\ModulIcons\LaborJournalS.ico")

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ;  Einstellungsbereich in der Addendum.ini anlegen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
	If !IniReadExt("LaborJournal", "") {
		backupfailure := ""
		If !FilePathCreate(adm.Dir "\_backup") {
			backupfailure := "Der Backup-Pfad konnte nicht angelegt werden.`n<<" adm.Dir "\_backup>> "
			PraxTT("Der Backup-Pfad konnte nicht angelegt werden!", "4 1")
		}
		else {
			FileCopy, % adm.ini, % adm.Dir "\_backup\" A_YYYY A_MM A_DD "Addendum.ini", 1
			If ErrorLevel
				backupfailure := "Die Addendum.ini konnte nicht nach \_backup\" A_YYYY A_MM A_DD "Addendum.ini kopiert werden."
		}
		If !backupfailure {
			labjournalIni =
			(LTrim
			[Laborjournal]
			htmlresource=%A_ScriptDir%\resources

			;-: verschiedene Warnstufen und Anzeigebedingungen für Laborparameter
			Warnen_Nie=
			Warnen_Immer=
			Warnen_Hoch=
			Warnen_Niedrig=
			Warnen_Exklusiv_Bezeichnung=
			Warnen_Exklusiv1=
			Warnen_Exklusiv2=
			Warnen_Exklusiv3=
			Warnen_Exklusiv4=
			Warnen_Exklusiv5=
			Warnen_Exklusiv6=

			Warnen_Gruppe1=atyp. Leuko::ALYMP,ALYMPDIFF,ALYMPH,ALYMPNEO,ALYMPREA,ALYMPRAM,KERNS,METAM,MYELOC,PROMY,HLAMBDA
			Warnen_Gruppe2=atyp. Ery::DIFANISO,DIFPOLYC,DIFPOIKL,DIFHYPOC,DIFMIKRO,DIFMAKRO,DIFOVALO,DIFECHIN
			Warnen_Gruppe3=COVID-19::COVIPC-A,COVIP-SP,COVIGAB,COVIA,COVIG,COVMU371,COVMU452,COVMU484,COVMU501,COVM6970

			;-: Laborkürzel deren Labortexte geladen werden sollen, um diese mit dem Laborjournal anzuzeigen oder zu verarbeiten
			;-: Autohotkey RegEx Strings um gekürzte Ausgaben der Labortexte zu erhalten z.B. mit Regex: "i)(?<TXT>Ct\s(?<Wert>\d+))",
			;-: Ct wird aus SARS-CoV-2-spezifische RNA nachgewiesen, Ct 25.
			;-: Ein String (Ct-TXT: "Ct 25") für die Ausgabe extrahiert und ein Wert (Ct-Wert: 25) gleichzeitig ausgelesen.
			;-: mehrere RegExStrings sind mit ## voneinander zu trennen
			Warnen_LabText1=

			;-: vordefinierte Filtergruppen zum Ein- und Ausblenden nach bereits erstellter Anzeige
			Tabellenfilter1=Abstriche::COV[\w\-]+
			Tabellenfilter2=Vitamine::VD25

			;-: die Variantenbezeichnungen werden von Labor zu Labor variieren
			SARS-CoV-WHO-Namen=Alpha,Beta,Gamma,Delta,Omikron,Lambda,My
			SARS-CoV-Varianten=COVM6970:Beta,COVMU501:Beta,COVMU452:Delta,COVMU371:Omikron,COVMU484:Beta

			;-: sonstiges
			Statistik_Param=
			LabParam_letztesUpdate=
			Datenabgleich_LaborAbruf=

			[B----
			)
			inifull := FileOpen(adm.ini, "r", "UTF-8").Read()
			inifull := RegExReplace(inifull, "\[B\-{4}", labjournalIni)
			FileOpen(adm.ini, "w", "UTF-8").Write(inifull)
			inifull := ""
		}
		else {
			MsgBox, 0x1000, Laborjournal, % "Die Erstellung eines Backup der Addendum.ini ist aus folgenden Grund fehlgeschlagen:`n" backupfailure
			ExitApp
		}

	}

  ;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ;   Laden der HTML-Vorlage und des Zeitstempel des Parameterupdates
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
	adm.LabJ := Object()
	adm.LabJ.htmlresource	:= IniReadExt("Laborjournal", "htmlresource", adm.Dir "\include\Gui") "\LaborjournalN.html"

	; Prüfen der HTML Datei
	If !FileExist(adm.LabJ.htmlresource)
		throw "Webinterface Datei nicht vorhanden.`n" adm.LabJ.htmlresource
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ;   Laden der "Warnen - Laborparameter"
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
	Warnen := Object()
	Loop 6 {
		iniVal := IniReadExt("Laborjournal", "Warnen_Exklusiv" A_Index)
		If iniVal && !InStr(iniVal, "ERROR")
			tmp .= iniVal ","
	}

	Warnen.exklusiv 	:= RTrim(tmp, ",")
	Warnen.nie       	:= IniReadExt("Laborjournal", "Warnen_Nie"    	, "CHOL,HBA1CIFC")
	Warnen.immer  	:= IniReadExt("Laborjournal", "Warnen_Immer"	, "NA,K,HBK-ST,HBK2-ST,PROCAL")
	Warnen.hoch    	:= IniReadExt("Laborjournal", "Warnen_Hoch" 	, "BNP,NTBNP,CK,CKMB,CKMM,CKBB,DDIM-CP")
	Warnen.niedrig  	:= IniReadExt("Laborjournal", "Warnen_Niedrig"	, "FOLS,VB12")

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Gruppen - Laborparameter für Anzeige zusammenfassen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Warnen.Gruppe	:= { "atyp. Leuko"	: "ALYMP,ALYMPDIFF,ALYMPH,ALYMPNEO,ALYMPREA,ALYMPRAM,KERNS,METAM,MYELOC,PROMY,HLAMBDA"
			    					,	"atyp. Ery" 	: "DIFANISO,DIFPOLYC,DIFPOIKL,DIFHYPOC,DIFMIKRO,DIFMAKRO,DIFOVALO,DIFECHIN"
						    		,	"COVID-19"	: "COVIPC-A,COVIP-SP,COVIGAB,COVIA,COVIG,COVMU371,COVMU452,COVMU484,COVMU501,COVM6970"}

	tmp := Array()
	For groupName, groupParameters in Warnen.Gruppe
		tmp.Push(groupName "::" groupParameters)

	Loop {
		iniVal := IniReadExt("Laborjournal", "Warnen_Gruppe" A_Index)
		If (!iniVal || InStr(iniVal, "ERROR"))
			break
		iniVal := StrSplit(iniVal, "::")
		Warnen.Gruppe[iniVal.1] := iniVal.2
	}
  ;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ;   Laden der "Warnen - LabText Parameter" und von Statistikdaten
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
	Warnen.LabText     	:= Object()
	Warnen.SaveLabText	:= "("

	Loop {

		iniVal := IniReadExt("Laborjournal", "Warnen_LabText" A_Index)
		If (!iniVal || InStr(iniVal, "ERROR"))
			break

	  ; z.B. COVID::^COV[\w\-]+::i)Ct\s+(?<CtWert>\d+)#i)(Verdacht.*?SARS.*?[CoOV\s\-2]+\-.*?Variante)\s+(\(VOC\))*\s*
	  ; 											. (?<Mutant>[\pL\d\-\.]+).*(?<MutantEx>[A-Z]\.\d+\.\s*\d+\.\s*\d+)
		RegExMatch(iniVal, "^(?<Name>.*?)::(?<Param>.*?)::(?<Filter>.*)$", rxLab)

	 ; Filter werden als Array gespeichert
		filter := StrSplit(rxLabFilter, "#")
		Warnen.LabText[rxLabName] := {"rxParam":rxLabParam, "rxFilter":(!IsObject(filter) ? [filter] : filter)}

	  ; alle RegExStrings die Parameter mit Labortexten erfassen in einen RegExString zusammenfassen
		Warnen.SaveLabText .= rxLabParam ? (StrLen(Warnen.SaveLabText)>1 ? "|" : "") rxLabParam : ""
	}

	If (Warnen.LabText.Count() > 0)
		Warnen.SaveLabText .= ")"
	else
		Warnen.SaveLabText := ""

  ; Nutzern eigene Statistiken (Zählungen) ermöglichen
    tmp := IniReadExt("Laborjournal", "Statistik_Parameter")
	tmp := "COVIPC-A:::Es wurden # PCR-Abstriche eingesendet§§COVIPC-A:i)(POSITIV|NEGATIV):i)POSITIV:Bei # lag ein positives Ergebnis vor."

	If !InStr(tmp, "Error") {
		tmp := StrSplit(tmp, "§§")
		Warnen.LabStats 	:= Object()
		Warnen.LabStatsKeys :="i)("
		For index, param in tmp {
			split := StrSplit(param, ":")
			key := split.1
			index := 1
			while IsObject(Warnen.LabStats[key index])
				index := A_Index + 1
			Warnen.LabStats[key index] := {"countIf": split.2, "matchif":split.3, "output":split.4}
			Warnen.LabStatsKeys .= !RegExMatch(Warnen.LabStatsKeys, "\b" key "\b") ? key  "|" : ""
		}
		Warnen.LabStatsKeys := RTrim(Warnen.LabStatsKeys, "|") ")"
	}

	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ;   Laden der "Tabellenfilter"
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		inipointer := 0
		Loop {
			iniVal := IniReadExt("Laborjournal", "Tabellenfilter" A_Index)
			If (iniVal && !InStr(iniVal, "ERROR")) {
				inipointer := 0
				If !IsObject(Warnen.Filter)
					Warnen.Filter := Object()
				iniVal := StrSplit(iniVal, "::")
				Warnen.Filter[iniVal.1] :={"rxFilter":iniVal.2, "state":0}
			}
			else {
				inipointer ++
				If (inipointer > 5)  ; 5 Zahlen Abstand können bei der Nummerierung gemacht werden
					break
			}
		}

	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ;   Laden von SARS-CoV-Bezeichnungen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
	Warnen.Covid := Object()
	Warnen.Covid.Varianten := Object()

	iniVal := IniReadExt("Laborjournal", "SARS-CoV-WHO-Namen")
	If iniVal && !InStr(iniVal, "ERROR")
		Warnen.Covid.WHONamen := StrReplace(iniVal, ",", "|")

	iniVal := IniReadExt("Laborjournal", "SARS-CoV-Varianten")
	If iniVal && !InStr(iniVal, "ERROR")
		For each, mutant in StrSplit(iniVal, ",")
			Warnen.Covid.Varianten[StrSplit(mutant, ":").1] := StrSplit(mutant, ":").2
  ;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ;   prüfen auf neue Labordaten
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
	adm.Labj.Override := Labjournal_NewLabData()
  ;}

	tmp:=key:=index:=mutant:=param:=iniVal:=inifull:=labjournalIni:=backupfailure:=filter:=rxLabParam:=rxLabName:=rxLabFilter:=""
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Skriptstart mit Parameter (Parsen der Parameter)
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If IsObject(A_Args) {

		For ArgIndex, arg in A_Args
			If RegExMatch(arg, "i)preRender\s*\=\s*(Yes|Ja|true|1)")
				adm.LabJ.preRender := true

		If adm.LabJ.preRender && !adm.Labj.Override
			ExitApp

	}

  ; load_Bar anzeigen
	If (!adm.LabJ.preRender), LoadBar_Gui()

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Patientendaten laden
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	If (!adm.LabJ.preRender), load_Bar.Set(0, "lade Patientendatenbank")
	filter  	:= ["NR", "PRIVAT", "GESCHL", "NAME", "VORNAME", "GEBURT", "PLZ", "ORT", "STRASSE", "HAUSNUMMER", "TELEFON", "TELEFON2", "TELEFAX"]
	PatDB 	:= ReadPatientDBF(adm.AlbisDBPath, filter, "allData")
	If (!adm.LabJ.preRender), load_Bar.Set(1, "Patientendatenbank geladen")
  ;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; bereits erstellte Tagesdaten lesen (Standardordner User)
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If FileExist(adm.LabPat) && !adm.Labj.preRender && !adm.Labj.Override {
		LabPat 	    	:= cJSON.Load(FileOpen(adm.LabPat        	, "r", "UTF-8").Read())
		load_Bar.Set(Percent:=50, "Labordaten geladen")

		COVID     	:= cJSON.Load(FileOpen(adm.COVIDStats	, "r", "UTF-8").Read())
		load_Bar.Set(Percent:=60, "COVID-Statistik geladen")

		COVIDStats 	:= LabJournal_CovidStats(COVID)
		load_Bar.Set(Percent:=70, "COVID-Statistik aufbereitet")
		adm.filtering	:= 0

		Properties  	:= LabPat.Delete("Properties")
		adm.LabJ.Tagesanzeige := Properties.Tagesanzeige
	}
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Journal neu zusammenstellen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else {
		LabPat	     	:= AlbisLaborJournal("", "", Warnen, 140, true)
		COVIDStats 	:= AlbisLabCOVID19()
		IniWrite, % adm.Labj.lastUpdate, % adm.Ini, % "Laborjournal", % "letztes_Journalupdate"
	}

 ; Laborjournal anzeigen
	LaborJournal(LabPat, true, Warnen, COVIDStats)


return

~ESC:: ;{
	If WinActive("Laborjournal") {
		SaveGuiPos(labJ.hwnd)
		ExitApp
	}
return ;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Berechnungen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
AlbisLaborJournal(Von="", Bis="", Warnen="", GWA=100, Anzeige=true) {           	;--  klinisch auffällige Werte der letzten Blutabnahmen bereitstellen

	/*  Parameterbeschreibung

		Von       	-	Anfangsdatum, leerer String oder Zahl
							eine Zahl als Wert wird als Anzahl der Werktage die rückwärtig vom aktuellem Tagesdatum angezeigt werden soll gewertet
		Bis         	-	Enddatum oder leer

		Warnen  	-	ein AHK Objekt mit Namen von Laborparametern für die Zuordnung zu einer Warnstufe ()

		GWA    	- 	(G)renz(w)ert(a)bweichung als absoluter Prozentwert
							wird mit der prozentualen (Normal-/) Grenzwertüber- oder unterschreitung verglichen
							nur Laborergebnisse welche größer oder gleich dieses Wertes sind werden erfasst
							ACHTUNG: bei einem zu hoch eingestellten GWA werden eventuell alarmierend hohe Laborwerte nicht erfasst!
											   ein guter GWA, möglichweise gilt das nur für mein Labor oder/und Einzugsbereich, liegt bei 140(%)
											   der Standardwert ist deshalb auf 100(%) eingestellt

	*/

	; Variablen                                                                                                         	;{

		global LabDic
		static Lab

		LabPat          	:= Object()
		LabTexte       	:= Object()
		nieWarnen    	:= "i)\b(" 	StrReplace(Warnen.nie       	, ","	, "|"	) 	")\b"
		immerWarnen 	:= "i)\b(" 	StrReplace(Warnen.immer  	, ","	, "|"	) 	")\b"
		niedrigWarnen 	:= "i)\b(" 	StrReplace(Warnen.niedrig  	, ","	, "|"	) 	")\b"
		hochWarnen 	:= "i)\b(" 	StrReplace(Warnen.hoch    	, ","	, "|"	) 	")\b"
		exklusivWarnen	:= "i)\b(" 	StrReplace(Warnen.exklusiv	, ","	, "|"	) 	")\b"

		If (Warnen.LabText.Count() > 0) {
			WarnLabTexte	:= true
			SaveLabText   	:= "i)" Warnen.SaveLabText
			SkipIfNotDigit	:= "\d+"
			lbarBase1 := 50
			lbarBase2 := 30
		}

		If !WarnLabTexte
			lbarBase1 := 80, lbarBase2 := 0

	  ; dies ist ein RegEx Pattern für Laborergebnisse meines Labors (nicht allgemein anwendbar)
		rxlabtxt 	:= 	{"probe": 	"Material.*?[\n\r]+(?<Material>.*)"
						    				. 	"(MIKROSKOPISCH.*?[\n\r]+(?<Mikroskopisch>.*))"
						    				.	"ERREGER.*?[\n\r]+(?<Erreger>.*)"
						    				. 	"Verfahren:.*?[\n\r](?<Verfahren>.*)"
						    				. 	"ANTIBIOGRAMM\sLegende:(?<Antibiogramm>.*)"}

	;}

	; durchschnittliche Abweichung von den Normwertgrenzen laden oder berechnen 	;{
		If !IsObject(Lab) {
			If !FileExist(adm.LabDBPath "\Laborwertgrenzen.json") {
				load_Bar.Set(1, "ermittle Laborwertgrenzen ...")
				Lab := AlbisLaborwertGrenzen(adm.LabDBPath "\Laborwertgrenzen.json", Anzeige)
			} else {
				load_Bar.Set(1, "lade Laborwertgrenzen ...")
				Lab := cJSON.Load(FileOpen(adm.LabDBPath "\Laborwertgrenzen.json"	, "r", "UTF-8").Read())
			}
		}

		If (!adm.LabJ.preRender), load_Bar.Set(5, "Laborwertgrenzen geladen")
	;}

	; Laborparameter lange Bezeichnungen laden                                                     	;{
		If (!adm.LabJ.preRender), load_Bar.Set(6, "lade Bezeichnungen der Laborparameter")
		LabDic := AlbisLabParam("load", "short")
		If (!adm.LabJ.preRender), load_Bar.Set(7, "Bezeichnungen der Laborparameter geladen")
	;}

	; Suchzeitraum                                                                                                  	;{

		If (!adm.LabJ.preRender), load_Bar.Set(8, "berechne den Suchzeitraum")

		; Von enthält kein Datum, dann werden beide Variablen geleert
			RegExMatch(Von, "^\d+$", Tage)

			Von 	:= Tage ? "" : Von
			Bis 	:= Tage ? "" : Bis

		; die Anzeige wird für x Werktage berechnet
			If !Von && !Bis {

				Tage	    	:= Tage ? Tage : 31                            	; kein Zeitraum, Defaulteinstellung sind 31 Tage
				Von      	:= A_YYYY . A_MM . A_DD                	; aktueller Tag

			; Zahl der Tage berechnen, die kein Werktag sind
				Wochen 	:= Floor(Tage/7)                               	; Tage in ganze Wochen umrechnen
				RestTage	:= Tage-Wochen*7                            	; verbliebene Resttage
				NrWTag	:= WeekDayNr(A_DDD)                		; Nr. des Wochentages
				WMod  	:= Mod(NrWTag-RestTage+1, 7)			; Divisonrest = Nummer des Wochentages minus berechnete RestTage + 1 durch 7
				PDays   	:= WMod = 0 ? 1 : WMod < 0 ? 2 : 0	; Divisonrest = 0 wäre Sonntag +1 Tag, < 0 alle Tage vor Sonntag also +2 Tage
				PDays   	:= PDays + Wochen*2                       	; + die Zahl der Wochenendtage der ganzen Wochen
				PDaysW	:= Floor(PDays/7)                               	; mehr als 7 Tage dann ist mindestens ein Wochende enthalten
				DaysPlus	:= PDays + 2*PDaysW + Tage
				Von     	  += -1*(DaysPlus), Days
				Von 	    	:= SubStr(Von, 1, 8)

				ViewStart 	:= FormatDate(Von, "YMD", "dd.MM.yyyy")
				WTag    	:= GetWeekDay(ViewStart)

				adm.LabJ.Tagesanzeige :=  "ab " WTag ", " StrReplace(ViewStart, A_YYYY) " (" Tage " Werktage)"

				FormatTime, Von, % Von, yyyyMMdd
				VB	:= "Von"

			}
			else
				VB := (Von && !Bis) ? "Von" : (!Von && Bis) ? "Bis" : "VonBis"

		If (!adm.LabJ.preRender), load_Bar.Set(8, "LABBLATT.dbf: öffnen")
	;}

	; Albis: Datenbank laden und entsprechend des Datums die Leseposition vorrücken	;{
		labDB := new DBASE(adm.AlbisDBPath "\LABBLATT.dbf", 0)
		labDB.OpenDBF()
	;}

	; Datei: Startposition berechnen (ein Quartal vor dem Datum)                             	;{
		If (VB = "Von") || (VB = "VonBis") {

			QJ := !QJ ? SubStr(Von, 1, 4) . Ceil(SubStr(Von, 5, 2)/3) : QJ 	; muss das Quartaljahr des Von-Datum sein

		  ; ein Quartal davor
			If !lab.seeks.HasKey(QJ) {
				nYYYY 	:= SubStr(QJ, 1, 4)
				nQ      	:= SubStr(QJ, 5, 1)
				nQ    	:= nQ-1=0 ? 4 : nQ-1
				nYY    	:= nQ=4 ? nYYYY-1 : nYYYY
				QJ     	:= nYYYY nQ
			}

			startrecord	:= Lab.seeks.HasKey(QJ) ? Lab.seeks[QJ] : Lab.seeks[Lab.seeks.MaxIndex()]
			result        	:= labDB._SeekToRecord(startrecord, 1)

		}
	;}

	; Laborwerte: klinisch interessante Werte finden                                                    	;{
		t :=  nieWarnen "`n" exklusivWarnen "`n" immerWarnen "`n`n"
		ShowAt := StrLen(labDB.records)-2 > 0 ? 10**(StrLen(labDB.records)-2) : Floor(labDB.records/10)
		If (!adm.LabJ.preRender), load_Bar.Set(9, "LABBLATT.dbf: lade neue Laborwerte")

	; filtert nach Datum und ab einer Überschreitung von der durchschnittlichen Abweichung von den Grenzwerten anhand
	; der zuvor für jeden Laborwert berechneten durchschnittlichen prozentualen Überschreitung des Grenzwertes aus den
	; Daten der LABBLATT.dbf [AlbisLaborwertGrenzen()]
		while !(labDB.dbf.AtEOF) {

			; Daten zeilenweise auslesen
				data := labDB.ReadRecord(["PATNR", "ANFNR", "DATUM", "BEFUND", "PARAM", "ERGEBNIS", "EINHEIT", "GI", "NORMWERT", "TXTHINW", "ERGTXT", "TEXTE"])

			; Fortschritt
				If !adm.LabJ.preRender && Anzeige && (Mod(A_Index, ShowAt) = 0) {
					tt := "lade neue Laborwerte  [" ConvertDBASEDate(data.Datum) "]  "
						. 	SubStr("00000000" labDB.recordnr, -5) "/"
						. 	SubStr("00000000" labDB.records, -5)
					percent := Round(lbarBase1*(labDB.recordnr/labDB.records),1)
					load_Bar.Set(9+percent, tt)
				}

			; Datum passt nicht, weiter
				Switch VB {
					Case "Von":
						If (data.Datum <= Von)
							continue
					Case "Bis":
						If (data.Datum => Bis)
							continue
					Case "VonBis":
						If (data.Datum < Von || data.Datum > Bis)
							continue
				}

			; IDs zu Texthinweisen bei bestimmten Laborwerten sichern  	;{
				PN := data.PARAM
				If WarnLabTexte && RegExMatch(PN, SaveLabText) {

					If !IsObject(LabTexte[PATNR := data.PATNR]) {
						LabTexte[PATNR] := Object()
						LabTexte[PATNR].Ref := Object()
					}
					If !IsObject(LabTexte[PATNR][ANFNR	:= data.ANFNR])
						LabTexte[PATNR][ANFNR] := Object()
					If !IsObject(LabTexte[PATNR][ANFNR][PN])
						LabTexte[PATNR][ANFNR][PN] := Object()

					LabTexte[PATNR].Ref[data.TEXTE]     	:= [ANFNR, PN, "TEXTE"]
					LabTexte[PATNR].Ref[data.ERGTXT]  	:= [ANFNR, PN, "ERGTXT"]
					LabTexte[PATNR].Ref[data.TXTHINW]	:= [ANFNR, PN, "TXTHINW"]
				 }
			;}

			; kein pathologischer Wert und in keiner Filterliste dann weiter ;{
				PNImmer := PNExklusiv := PNLow := PNHigh := false
				If RegExMatch(PN, nieWarnen)                   {        	; Parameter nieWarnen
					continue
				}
				else if RegExMatch(PN, exklusivWarnen)  	{        	; Parameter exklusivWarnen
					PNExklusiv 	:= true
				}
				else if RegExMatch(data.GI, "[\+\-\!]", GI) 	{ 			; Parameter immer,niedrig und hochWarnen

					If RegExMatch(PN, immerWarnen)      	{
						PNImmer	:= true
					}
					else if RegExMatch(PN, niedrigWarnen)	{
						If !(PNLow := (GI="-") ? true : false)
							continue
					}
					else if RegExMatch(PN, hochWarnen)  	{
						If !(PNHigh := (GI="+")	? true : false)
							continue
					}

				}
				else if !RegExMatch(data.GI, "[\+\-\!]")
					continue

			;}

			; Strings lesbarer machen (oder auch nicht)                        	;{

				; Laborergebnis oder (P)arameter(W)ert
					PWo	:= PW := StrReplace(data.ERGEBNIS, ",", ".")        	; Parameter - Wert
					If !RegExMatch(PW, "i)(\d|positiv|negativ)")                             	;
						continue
					RegExMatch(PW, "[\<\>]+", PWco)                              	; PW cutoff - ein cutoff Ergebnis
					PW	:= RegExReplace(PW, "[\<\>\s]+", "")
					PW	:= PWco ="<" ? 0 : PW

				; NORMWERTGRENZE - AUSWERTEN
				; PV1 kann > oder < sein, PSU - unterer Grenzwert, PSO - oberer Grenzwert
				; 	Beispiele aus der LABBLATT.dbf - Spalte Normwerte
				; 		1. Beispiel: 0,02 - 0,08	 --> PSU	= 0,02	u. PSO=0,08
				; 		2. Beispiel: 	<	0,01  	 --> PV 	= "<" 	u. PSO=0,01 	(path. ist alles darüber)
				; 		3. Beispiel: 	>	1,04    	 --> PV 	= ">" 	u. PSU=1,04   	(path. ist alles darunter)

				  ; Normwert-String mit RegEx aufsplitten
					data.NORMWERT := RegExReplace(data.NORMWERT, "[a-z]+$")  ; entfernt nicht auswertbare Buchstaben
					RegExMatch(data.NORMWERT, "(?<V1>[\<\>])*(?<SU>[\d,\.]+)\s*(?<V2>\-)*\s*(?<SO>[\d,\.]+)*", P)

					PSU  	:= StrReplace(PSU	, ",", ".")                        	; unterer Grenzwert
					PSO  	:= StrReplace(PSO, ",", ".")                        	; oberer Grenzwert
					PSO  	:= !PSO ? PSU : PSO                              	; wenn kein oberer Grenzwert vorhanden ist, dann ist es ein Cut-Wert
					PV    	:= PV1 ? PV1 : PV2                               	; "<", ">" oder "-"
					If (PV = "<")
						PSO := PSU ? PSU : PSO
					else if (PV = ">")
						PSU := PSO ? PSO : PSU

					OD    	:= lab.params[PN].OD                          	; durchschn. Abweichung v. d. oberen Normwertgrenze (oberer Durchschnitt in Prozent)
					UD	  	:= lab.params[PN].UD                            	; durchschn. Abweichung v. d. unteren Normwertgrenze (unterer Durchschnitt in Prozent)
					PE	    	:= data.EINHEIT				                    	; Parameter - Einheit
					ONG	:= UNG := 0                                        	; Obere/Untere Norm-Grenze
			;}

			; außer bei immer anzuzeigenden Parametern ausführen
				If !PNExklusiv {

					If (PW > PSO)                                                   	; Laborwert überschreitet die obere Grenze
						ONG := Floor(Round((PW*100)/PSO, 3))       	; Obere Normgrenze in Prozent
					else if (PW < PSU)                                              	; Laborwert ist kleiner als der untere Grenzwert
						UNG := Floor(Round((PSU*100)/PW, 3))        	; Untere Normgrenze in Prozent

					If !ONG && !UNG                                             	; ONG und UNG sind beide null -> weiter
						continue

					If (ONG > 0)                                                    	; ONG überschritten
						plus := Floor(Round((ONG*100)/OD,1))           	; Überschreitung der ONG in Prozent
					else if (UNG > 0 )                                             	; UNG überschritten
						plus := Floor(Round((UNG*100)/UD,1))           	; Unterschreitung der UNG in Prozent

					If PNImmer && (plus <= 100)                              	; nur aufnehmen bei Überschreitung einer Normgrenzen
						PNImmer := false
					else if PNLow && (plus <= 100)							; werden einfach über PNImmer eingeschleust
						continue
					else if PNHigh && (plus <= 100)
						continue

					PW := PWco ? PWo : PW

				}

			; Laborwert liegt über der Grenzwertabweichung (GWA) oder gehört zu einer Filterliste
				If (plus > GWA || PNImmer || PNExklusiv) {

					PatID 	:= data.PATNR
					Datum 	:= data.Datum
					TH    	:= data.TXTHINW  (data.TEXTE ? "," data.TEXTE : "")  (data.ERGTXT ? "," data.ERGTXT : "")
					PW    	:= RegExReplace(PW, "(\.[0-9]+[1-9])0+$", "$1")        ; Parameter Wert
					AN    	:= LabPat[Datum][PatID][PN].AN                             	; Anforderungsnummer?
					AN    	:= !RegExMatch(AN, "\b" data["ANFNR"] "\b") ? AN "," data["ANFNR"] : AN

				; Subobjekte anlegen
					If !IsObject(LabPat[Datum])
						LabPat[Datum] := Object()
					If !IsObject(LabPat[Datum][PatID])
						LabPat[Datum][PatID] := Object()
					If !IsObject(LabPat[Datum][PatID][PN])
						LabPat[Datum][PatID][PN] := Object()

				; Daten übergeben
					If !RegExMAtch(LabPat[Datum][PatID]["ANFNR"], "\[" data["ANFNR"])
						LabPat[Datum][PatID]["ANFNR"] := data["ANFNR"] " " data["BEFUND"] ", "

					PL 	:= ONG > 0 ? "+" Round(ONG/100, 1)	: (UNG > 0 ? "-" Round(UNG/100,1)   	: "")

					LabPat[Datum][PatID][PN].PatID	:= PatID                                                             	; Patienten-ID
					LabPat[Datum][PatID][PN].AN  	:= LTrim(AN, ",")                                                  	; Labor Anforderungsnummer
					LabPat[Datum][PatID][PN].CV  	:= PNImmer ? 1 : PNExklusiv ? 2 : 0                     	; CAVE-Wert für andere Textfarbe im Journal
					LabPat[Datum][PatID][PN].PN  	:= PN                                                                  	; Parameter - Name
					LabPat[Datum][PatID][PN].PW  	:= RegExReplace(PW, "\.0+$")                              	; Parameter - Wert
					LabPat[Datum][PatID][PN].PV		:= P ? StrReplace(P, ",", ".") : data["NORMWERT"]   	; Parameter - Varianz (Normwertgrenzen)
					LabPat[Datum][PatID][PN].PA		:= ONG > 0 ? OD : (UNG > 0 ? UD : "" )           	; Parameter - Abweichung +/-
					LabPat[Datum][PatID][PN].PE		:= PE                                                                  	; Parameter - Einheit
					LabPat[Datum][PatID][PN].PB   	:= LabDic[PN].1                                                   	; Parameter - Bezeichnung (Langtext)
					LabPat[Datum][PatID][PN].PL 		:= PL                                                                   	; proz. Abweichung v.d. Norm (beide Richtungen)
					LabPat[Datum][PatID][PN].NG 	:= PV                                                                 	; Normgrenzenart (<,>,-,bool)

				}

		}

	;}

	; Lesezugriff beenden                                                                                         	;{
		labDB.CloseDBF()
		LabDB := ""
		If Anzeige
			ToolTip

		percent := 9+percent + 1
		If (!adm.LabJ.preRender), load_Bar.Set(percent, "Labortexte auslesen")
	;}

	; LabTexte: auslesen                                                                                          	;{
		If WarnLabTexte {

			PatIDs := "^("
			For PatID, labtext in LabTexte
				PatIDs .= PatID "|"
			PatIDs := RTrim(PatIDs, "|") ")$"

		; Labtexte.dbf: zum Lesen bereit machen
			labDB := new DBASE(adm.AlbisDBPath "\LABTEXTE.dbf", 0)
			labDB.OpenDBF()
			ShowAt := StrLen(labDB.records)-2 > 0 ? 10**(StrLen(labDB.records)-2) : Floor(labDB.records/10)

		; LabTexte: Daten auslesen
			while !(labDB.dbf.AtEOF) {

			; Daten zeilenweise auslesen
				data := labDB.ReadRecord(["PATNR", "ID", "POS", "TEXT"])
				If (Mod(labDB.recordnr, ShowAt) = 0)
					load_Bar.Set(percent+Round(lbarBase2*(labDB.recordnr/labDB.records), 1), "Labortexte werden gelesen")

			; Textlinks vergleichen
				PATNR := data.PATNR
				If (PN := LabTexte[PATNR].Ref[data.ID].2) {
					ANFNR	:= LabTexte[PATNR].Ref[data.ID].1
					FLabel 	:= LabTexte[PATNR].Ref[data.ID].3
					txt := RegExReplace(data.Text, "([^\w])[\r\n]+", "$1")
					txt := RegExReplace(txt, "[\r\n]+", " ")
					txt := RegExReplace(txt, "([a-zäöüß])([A-ZÄÖÜ])", "$1 $2")
					txt := RegExReplace(txt, "\.([A-ZÄÖÜ])", ". $1")
					LabTexte[PATNR][ANFNR][PN][FLabel] .= txt
				}

			}

		; Labtexte: aufräumen
			For PATNR, obj in LabTexte
				obj.Delete("Ref")

		; LabTexte: Lesezugriff beenden
			labDB.CloseDBF()
			labDB := ""

		; Fortschritt
			percent += lbarBase2
			If (!adm.LabJ.preRender), load_Bar.Set(percent, "Labortexte ausgelesen")

		; Fortschritt
			percent += 1
			If (!adm.LabJ.preRender), load_Bar.Set(percent, "Labortexte gesichert")

		; LabTexte: hinzufügen
			For Datum, Patients in LabPat
				For PATNR, LabParams in Patients {
					For PN, Labval in LabParams {
						IF (PN = "ANFNR") && !RegExMatch(PN, SaveLabText)
							continue
						ANFNR := LabVal.AN
						If IsObject(LabTexte[PATNR][ANFNR][PN]) {
							LabPat[Datum][PATNR][PN].TH := LabTexte[PATNR][ANFNR][PN]
							LabTexte[PATNR][ANFNR].Delete(PN)
							If !LabTexte[PATNR][ANFNR].Count()
								LabTexte[PATNR].Delete(ANFNR)
						}
					}
				}

		}

		; Labtexte: aufräumen
			For PATNR, obj in LabTexte {

				If !obj.MaxIndex() {
					LabTexte[PATNR] := ""
					LabTexte.Delete(PATNR)
					continue
				}

				For ANFNR, val in obj
					If !val.Count()
						obj.Delete(ANFNR)

			}

		percent += 1
		If (!adm.LabJ.preRender), load_Bar.Set(percent, "Datenausgabe erstellt")
	;}

	; LabPat für erneutes Laden ohne Neuberechnung speichern
		LabPat.Properties := Object()
		LabPat.Properties.Tagesanzeige := adm.LabJ.Tagesanzeige
		FileOpen(adm.LabTexte	, "w", "UTF-8").Write(cJSON.Dump(LabTexte, 1))
		FileOpen(adm.LabPat	, "w", "UTF-8").Write(cJSON.Dump(LabPat, 1))
		IniWrite, % adm.Labj.lastUpdate, % adm.Ini, % "Laborjournal", % "letztes_Journalupdate"

return LabPat
}

AlbisLaborwertGrenzen(LbrFilePath, Anzeige=true) {                                             	;-- Berechnung der durchschnittlichen Abweichung v. Grenzwert

	/* 	Beschreibung - AlbisLaborwertGrenzen()

		⚬	Berechnet aus den Daten der LABBLATT.dbf die durchschnittliche prozentuale Über- oder Unterschreitung eines Laborparameters.
		⚬	Für eventuelle Anpassungen wird die maximalste Über- oder Unterschreitung als Einzelwert gespeichert.
		⚬	Durch Nutzung eines Faktors (Prozentwert) erscheinen mir, die durch Annährung erreichten "Warngrenzen", auch bei unterschiedlichen
			Einheiten und altersabhängigen Normwertgrenzen klinisch bedeutsame Laborwertveränderungen sicher herauszufiltern.

		⊛  Die Relevanz von path. Laborwerten ist bei einem Teil ganz deutlich am Grenzwert festzumachen. Als bestes Beispiel sind hier
			Elektrolytwert-Veränderungen zu nennen. Bei anderen Parametern ist dies aber nicht so. Als langjährig tätiger Arzt benötigt man oft
			nur einen Blick die Werte und kann die Relevanz sofort ermessen. Doch wie vermittelt man dies dem Computer?

			Im Laufe meiner Tätigkeit als niedergelassener Arzt circa 300 verschiedene Laborparameter interessant geworden.
			Bei der Menge an Parametern ist eine regelmäßige Wichtung schon aufgrund der vom Laboranbieter in Abständen angepassten und sich
			damit verändernden Normbereiche unsinnig.

			Die Idee war also nur die Über- und Unterschreitung der Grenzwerte zu wichten. Bei der Berechnung werden alle im Normbereich liegenden
			Werte vom Algorithmus ignoriert. Über- und Unterschreitungen werden als relative Abweichung erfasst. Der Einfachheit halber erfolgt dies
			in Prozent (200 % oder das 2fache der Norm z.B.). Für die Abweichungen in beide Richtungen werden für jeden Parameter zwei Werte ge-
			bildet. Erfasst wird zum einen die maximale prozentuale Abweichung. Und aus den Laborwerten aller Patienten wird für jeden Parameter
			ein Durchschnitt berechnet. Ich nenne es die "durchschnittliche Abweichung" vom oberen oder unteren Grenzwert.

			Die berechneten Alarmwerte sind in den meisten Fällen gut und entsprechen meinen Erwartungen. Das Ziel der Reduktion der Informationsfülle
			und die Sicherheit der Patienten konnte mit ein paar wenigen Ausnahmen eingehalten werden.
			Bei einigen Parameter sehe ich noch Einschränkungen. Unzureichend ist im Moment noch die Erkennung von Unterschreitungen.
			Beispiel Hb-Wert: 5.4 mmol/l Normwert 7.2-9.6 mmol/l. Die Alarmgrenze wurde mit 5.2 mmol/l berechnet. Dieser relevante Werte erscheint
			nicht im Laborjournal

			Verlassen Sie sich deshalb nicht gänzlich auf die errechneten Ergebnisse!

	 */

	; Variablen
		Lab  	:= Object()
		seeks:= Object()
		cJSON.EscapeUnicode := "UTF-8"

	; Albis Datenbank laden und entsprechend des Datums die Leseposition vorrücken
		labDB := new DBASE(adm.AlbisDBPath "\LABBLATT.dbf", Anzeige)
		labDB.OpenDBF()

	; filtert die Daten und berechnet die Abweichungen von den Grenzwerten
		while !(labDB.dbf.AtEOF) {

				data := labDB.ReadRecord(["DATUM", "PARAM", "ERGEBNIS", "EINHEIT", "GI", "NORMWERT"])
				QJ := SubStr(data.Datum, 1, 4) . Ceil(SubStr(data.Datum, 5, 2)/3)
				If !seeks.HasKey(QJ)
					seeks[QJ] := labDB.recordnr

				If Anzeige && (Mod(A_Index, 1000) = 0)
					ToolTip, %	"Laborwertgrenzen werden berechnet.`ngefundene Parameter: " lab.Count()
								. 	"`nDatensatz: " SubStr("00000000" A_Index, -5) "/" SubStr("00000000" labDB.records, -5) ;"`n" lfound

				If !RegExMatch(data["GI"], "(\+|\-)") || StrLen(data["NORMWERT"]) = 0
					continue

				PN	:= data.PARAM                                                  	; Parameter Name
				PE		:= data.EINHEIT                                              		; Parameter Einheit
				PW	:= StrReplace(data["ERGEBNIS"]	, ",", ".")             	; PARAMETER WERT
				RegExMatch(data["NORMWERT"], "(?<V>[\<\>])*(?<SU>[\d,\.]+)[\-\s]*(?<SO>[\d,\.]+)*", P)
				PSU	:= StrReplace(PSU	, ",", ".")                                   	; Parameter untere Grenze
				PSO	:= StrReplace(PSO, ",", ".")                                   	; Parameter obere Grenze


			; berechnet die Grenzwertüberschreitungen
				ONG :=UNG := 0
				PSO := StrLen(PSO) > 0 ? PSO : PSU
				If (PW >= PSO)
					ONG := Floor(Round((PW * 100)/PSO, 3))
				else If (PW <= PSU)
					UNG := Floor(Round((PSU * 100)/PW, 3))

				If !ONG && !UNG
					continue

			; Ersterstellung eines Laborwert Datensatzes
				If !Lab.HasKey(PN) {

					Lab[PN] := {	"PZ" 	: 1			; Parameterzähler
									, 	"O" 	: 0			; Summe d. prozentualen Abweichung vom oberen Grenzwert
									, 	"OI"	: 0			; Abweichungszähler (obere Grenzwerte)
									, 	"OD"	: 0			; durchschnittliche prozentuale Abweichung vom oberen Grenzwert (O/OI)
									, 	"OM": 0    		; maximale prozentuale Abweichung o Obergrenze
									,	"N"	: 0          	; Summe der Normwerte
									,	"NI"	: 0			; Normalwertzähler
									,	"ND"	: 0			; Normalwertdurchschnitt
									, 	"U"	: 0			; Summe d. prozentualen Abweichung vom unteren Grenzwert
									, 	"UI"	: 0			; Abweichungszähler
									, 	"UD"	: 0			; durchschittliche prozentuale Abweichung vom unteren Grenzwert (O/OI)
									, 	"UM"	: 0}			; maximale prozentuale Abweichung an der Untergrenze

				}
			; Hinzufügen von Daten
				If (ONG > 0) {
					Lab[PN].O 	+= ONG
					Lab[PN].OI 	+= 1
					Lab[PN].OD	:= Round(Lab[PN].O/Lab[PN].OI)
					Lab[PN].OM	:= Lab[PN].OM >= ONG ? Lab[PN].OM : ONG
				}

				If (UNG > 0) {
					Lab[PN].U 	+= UNG
					Lab[PN].UI 	+= 1
					Lab[PN].UD 	:= Round(Lab[PN].U/Lab[PN].UI)
					Lab[PN].UM 	:= Lab[PN].UM >= UNG ? Lab[PN].UM : UNG
				}

				Lab[PN].PZ ++

		}

	; Lesezugriff beenden
		labDB.CloseDBF()

	; Datenmenge wird verkleinert
		Labwrite := Object()
		Labwrite.params := Object()
		For PN, m in Lab
			Labwrite.params[PN] := {"OM": m.OM, "OD": m.OD, "UD": m.UD, "UM": m.UM}

	; Dateipositionen für eine quartalsweise Verarbeitung werden hinzugefügt
		Labwrite.seeks := Object()
		For QJ, filepos in seeks
			Labwrite.seeks[QJ] := filepos

	; Speichern der Daten
		FileOpen(LbrFilePath "\Laborwertgrenzen.json"	, "w", "UTF-8").Write(cJSON.Dump(Labwrite, 1))
		FileOpen(LbrFilePath "\Laborwertbasis.json"    	, "w", "UTF-8").Write(cJSON.Dump(Lab, 1))

		ToolTip

return Labwrite
}

AlbisLabParam(cmd:="load", data:="short", Anzeige:=false) {                               	;-- Laborparameter als Excel (.csv) Datei speichern

		cJSON.EscapeUnicode := "UTF-8"

	; LABPARAM.dbf wurde verändert dann müssem Inhalte neu erstellt werden
		If (filetime := Labjournal_NewLabParams()) {
			data		:= "short"
			cmd  	:= "rewrite"
			LabDB	:= new DBASE(adm.AlbisDBPath "\LABPARAM.dbf", Anzeige)
			IniWrite, % (adm.Labj.lastParams := filetime), % adm.Ini, % "Laborjournal", % "LabParam_letztesUpdate"
			If (!adm.LabJ.preRender)
				load_Bar.Set(7, "Bezeichnungen der Laborparameter müssen aktualisiert werden")
		}

	; Wörterbuch laden
		If (cmd = "load") {
			If FileExist(adm.LabDBPath "\LabDictionary.json") {
				LabDic := cJSON.Load(FileOpen(adm.LabDBPath "\LabDictionary.json"	, "r", "UTF-8").Read())
				return LabDic
			}
			data := "short"
		}

	; Datenbank öffnen und lesen
		LabDB.OpenDBF()
		LabParam := labDB.GetFields()
		LabDB.CloseDBF()
		LabDB := ""

	; full - ist für andere Auswertungen/Anzeigeeinstellungen gedacht
		If (data = "full") {

			For idx, line in LabParam {
				For key, val in Line {
					t .= val "`t"
					If (idx=1)
						h .= key "`t"
				}
				t := RTrim(t, "`t") "`n"
			}
			h := RTrim(h, "`t") "`n"

			FileOpen(adm.LabDBPath "\LabParam.csv" 	, "w", "UTF-8").Write(h . t)
			FileOpen(adm.LabDBPath "\LabParam.json"	, "w", "UTF-8").Write(cJSON.Dump(LabParam, 1))

			return LabParam
		}
	  ; es wurde noch kein Wörterbuch angelegt oder es soll neu erstellt werden
	  ; Wörterbuch besteht nur aus der Abkürzung (key) und der ausgeschriebenen Parameterbezeichnung (value)
		else if (data = "short") {

			LabDic	:= Object()
			For idx, obj in LabParam {

				If !LabDic.HasKey(obj.NAME) && (StrLen(obj.BEZEICH) > 0)
					LabDic[obj.NAME] := [obj.BEZEICH]
				else if LabDic.HasKey(obj.NAME) && (StrLen(obj.BEZEICH) > 0)
					LabDic[obj.NAME].Push(obj.BEZEICH)

			}
			FileOpen(adm.LabDBPath "\LabDictionary.json"	, "w", "UTF-8").Write(cJSON.Dump(LabDic, 1))

		return LabDic
		}

}

AlbisLabCOVID19(DateRegEx:="", LabParams:="") {                                          	;-- COVID19-PCR Statistik

	COVID                 	:= Object()
	COVID.day           	:= Object()
	COVID.weeks        	:= Object()
	COVID.PCRcount   	:= 0
	COVID.PCRpositive  	:= 0
	COVID.mutants     	:= {}
	COVID.Patients     	:= {}
	lmax := 0

	Y := "202["
	Loop % SubStr(A_YYYY, 4, 1)+1
		Y .= A_Index - 1

	DateRegEx 	:= Y "]\d\d\d\d"   ; 202[012] verlängert sich automatisch
	LabParams 	:= !LabParams ? {"Param" : "(COVIPC-A|COVM\w+)"} : LabParams
	mutants     	:= {"COVM6970" : "Beta", "COVMU501":"Beta", "COVMU452":"Delta", "COVMU371":"Omikron", "COVMU484":"N501Y"}

  ; Datenbank öffnen, auslesen und Lesezugriff beenden
	labDB := new DBASE(adm.AlbisDBPath "\LABBLATT.dbf", 0)
	labDB.OpenDBF()
	dbfdata	:= labDB.Search({"Datum" : "rx:" DateRegEx, "Param" : "rx:" LabParams.Param}, 0,, {"LogicalComparison":"and"})
	labDB.CloseDBF()
	LabDB := ""

  ; Progressvariablen
	maxrecords := dbfdata.Count()
	maxLength := StrLen(maxrecords) - 1

  ; Daten auswerten
	For recordNr, m in dbfdata {

		PatID      	:= m.PATNR
		examday	:= m.DATUM
		examyear	:= SubStr(m.Datum, 1, 4)
		weeknr  	:= WeekOfYear(m.DATUM)

	; Patienten zählen
		If !IsObject(COVID.Patients[PatID])
			COVID.Patients[PatID] := {"PCRcount":0}

	; Tagesstatistik erweitern
		If !IsObject(COVID.day[examday])
			COVID.day[examday] := Object()
		If !IsObject(COVID.day[examday][PatID])
			COVID.day[examday][PatID] := {"PCRpositive":0}
	; Wochenstatistik erweitern
		If !IsObject(COVID.weeks[examyear])
			COVID.weeks[examyear] := {"PCRcount":0, "PCRpositive":0}
		If !IsObject(COVID.weeks[examyear][weeknr])
			COVID.weeks[examyear][weeknr] := {"PCRcount":0, "PCRpositive":0}

	 ; positive Abstriche zählen
		If (m.Param = "COVIPC-A") {

		  ; PCR Abstriche zählen
			COVID.PCRcount += 1
			COVID.Patients[PatID].PCRcount += 1
			COVID.weeks[examyear].PCRcount += 1
			COVID.weeks[examyear][weeknr].PCRcount += 1

		  ; positives Testergebnis?
			isPositive := m.ERGEBNIS="POSITIV" ? 1 : 0

		  ; zählen um später Reinfektionen zu finden
			If isPositive {
				If !IsObject(COVID.Patients[PatID].PCRpositive)
					COVID.Patients[PatID].PCRpositive := {m.Datum:1}
				else If !COVID.Patients[PatID].PCRpositive[m.Datum]
					COVID.Patients[PatID].PCRpositive[m.DATUM] := 1
			}

		 ; positive PCR Ergebnisse zählen
			COVID.PCRpositive += isPositive
			COVID.weeks[examyear].PCRpositive	+= isPositive
			COVID.weeks[examyear][weeknr].PCRpositive	+= isPositive
			COVID.day[examday][PatID].PCRpositive += isPositive

		}
	; Coronavarianten zählen
		else if (RegExMatch(m.Param, "COVM\w+")) {

			variant := mutants[m.Param]

			If (m.ERGEBNIS = "POSITIV") {

			  ; Jahresstatistik
				If !IsObject(COVID.weeks[examyear].Mutants)
					COVID.weeks[examyear].Mutants := Object()
			  ; Wochenstatistik
				If !IsObject(COVID.weeks[examyear][weeknr].Mutants)
					COVID.weeks[examyear][weeknr].Mutants := Object()
			 ; Tagesstatistik
				If !IsObject(COVID.day[examday][PatID].Mutant)
					COVID.day[examday][PatID].Mutant := Object()

			  ; Allzeitstatistik
				COVID.mutants[variant]  := !COVID.Mutants[variant] ?  1 : COVID.Mutants[variant] +1
			  ; Jahresstatistik
				COVID.weeks[examyear].Mutants[variant] := !COVID.weeks[examyear].Mutants[variant] ? 1 : COVID.weeks[examyear].Mutants[variant] + 1
			  ; Wochenstatistik
				COVID.weeks[examyear][weeknr].Mutants[variant] := !COVID.weeks[examyear][weeknr].Mutants[variant] ? 1 : COVID.weeks[examyear][weeknr].Mutants[variant] + 1
			  ; Tagesstatistik
				COVID.day[examday][PatID].Mutant[variant] := !COVID.day[examday][PatID].Mutant[variant] ? 1 : COVID.day[examday][PatID].Mutant[variant] + 1

			}
		}

	}

  ; COVID Objekt speichen
	FileOpen(adm.COVIDStats, "w", "UTF-8").Write(cJSON.Dump(COVID, 1))

return LabJournal_CovidStats(COVID)
}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; grafische Anzeige
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
LaborJournal(LabPat, Anzeige=true, Warnen="", htmltable="") {                         	; erstellt die HTML Ausgabe

	; neutron - object wurde super global gemacht
		global srchdrecords, dbrecords, Tagesanzeige, parentNode

	; Filter vorbereiten                                     	;{
		rxFilter := ""
		For filtername, filter in Warnen.Filter
			rxFilter .= filter.state ? (rxFilter ? "|" : "") filter.rxFilter : ""
		rxFilter := rxFilter ? "i)^(" rxFilter ")" : ""

	  ; Statistikfilter vorbereiten
		rxCtWert   	:= Warnen.LabText.COVID.rxFilter.1
		rxMutant   	:= Warnen.LabText.COVID.rxFilter.2
		rxCovidFilter	:= Warnen.LabText.COVID.rxFilter

	  ; HTML-Knoten nachdem Tabellenzeilen eingefügt werden
		tbodynode := adm.tbodynode
		FirstNodeInsert := false
	;}

	; HTML Vorbereitungen 	                              	;{
		static TDC     	:= "<TD style='text-align:Right;	border-right:1px; '>"                                                        	;
		static TDJ      	:= "<TD style='text-align:Justify;border-right:1px; '>"
		static TDL      	:= "<TD style='text-align:Left; '>"                                                                          	; Textausrichtung links
		static TDL1     	:= "<TD style='text-align:Left; 	 '>"                                                                       	; Textausrichtung links kein Rand links
		static TDLR1     	:= "<TD style='text-align:Left; 	 '>"                                                                       	; Textausrichtung links, kein Rand rechts
		static TDAB      	:= "<TD style='text-align:Right;  '>"                                                                       	; Textausrichtung rechts
		static TDR1     	:= "<TD style='text-align:Right;	 '>"                                                                       	; Textausrichtung rechts, kein Rand rechts
		static TDR2a     	:= "<TD style='text-align:Right;	border-right:1px; '>"
		static TDR2b     	:= "<TD style='text-align:Right;	border-right:1px; '>"                                             	; Tooltip
		static TDR3      	:= "<TD class='tooltip' data-tooltip='#TT1#'"                                                           	; Param
									. " onclick='ahk.LabJournal_MoreTips(" q "#TT2#" q  ")'"
									. " style='text-align:Center; '>"

		static cColW    	:= ["color:Red;         	font-weight:bold; "
									, "color:BlueViolet;	font-weight:bold; "
									, "color:DarkTeal; 	font-weight:bold;   	font-family: sans-serif; "
									, "color:Blue;        	font-weight:normal; "
									, "color:FireBrick;  	font-weight:normal; "
									, "color:Black;       	font-weight:normal; 	font-size:smaller; "
									, "color:BlueViolet;	font-weight:normal; 	font-size:smaller; "]

	  ; für die Corona-Statistiken
		If !adm.filtering
			COVID 	:= {"Mutation":{}, "Ct":0, "CtCount":0, "Positiv":{}, "PCR":{}, "CtMin":999, "CtMax":0}  	; COVID-19 Statistik für Rekordsammler

	; HTML/CSS - Header, Infobereich und
		If !adm.filtering
			html := Laborjournal_BaseHTML(htmltable)

	;}

	; HTML Tabelle wird erstellt                         	;{
		labstats   	:= Object()
		sortDatum	:= Array()
		thisrow		:= 1
		filternumb	:= 1
		trlDf      	:= false                                                                                                                                	; alternierende Farbanzeige
		Firstrow 	:= true
		NewPat 	:= false

      ; nach Tagen der Blutwertbestimmungen sortieren
		For Datum, Patients in LabPat
			sortDatum.InsertAt(1, Datum)

	; Fortschritt
		If (!adm.LabJ.preRender), load_Bar.Set(percent+=1, "erstelle die Journalanzeige")
		pstep := (100-percent-2)/sortDatum.Count()

	; ―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; die Journalausgabe beginnt mit dem Datum der zuletzt eingegangen Labordaten
	; ―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		For idx, Datum in sortDatum {

			If (DatumLast <> Datum)
				UDatum := ConvertDBASEDate(Datum)

		; ―――――――――――――――――――――――――――――――――――――――――――――――――――――
		; Ausgabe der Werte sortiert nach Patient
		; ―――――――――――――――――――――――――――――――――――――――――――――――――――――
			For PatID, parameter in LabPat[Datum] {

				If (PatIDLast <> PatID) || (DatumLast <> Datum) {                            	; Patientenname wird nur einmal pro Tag angezeigt
					Patient  	:= PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
					PatGeburt	:= ConvertDBASEDate(PatDB[PatID].GEBURT)
					ANFNR  	:= RTrim(parameter["ANFNR"], ", ")
					PatIDLast	:= PatID
					NewPat 	:= true                                                                     	; alternierende Änderung der Hintergrundfarbe
				}

			; ―――――――――――――――――――――――――――――――――――――――――――――――――――
			; Parameter Objekt (PO) - Erstellen der Datenzeilen
			; ―――――――――――――――――――――――――――――――――――――――――――――――――――
				For LabParam, PO in parameter {

				  ; ist PO leer, einfach ignorieren
					If !IsObject(PO)
						continue

				  ;  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
				  ; COVID  ‧  ‧  ‧ 	Labortexte und Statistiken
				  ;  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧ ;{
					If IsObject(PO.TH) {

					 ; Texte zusammenfügen
						PTEXT := ""
						For each, txt in PO.TH
							PTEXT .= "##" txt
						PTEXT .= "##"

						if (adm.filtering>1 && Warnen.Filter.Abstriche.state=2) {
							RegExMatch(PTEXT, rxCtWert, Lab_)
							LAB_TXT := LAB_CtWert ? "[" LAB_CtWert "]" : ""
							goto checkFilter
						}

						For rxN, rxStr in rxCovidFilter {

						  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
						  ; Abstrich zählen
						  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
							If !COVID.PCR.HasKey(PatID) {
								COVID.PCR[PatID] := 1
							   ; Anforderungsnummer für Vergleich sichern
								ANFLast := PO.PN "|" PO.AN
							}
						  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
						  ; weiteren Abstrich im untersuchten Zeitraum gefunden
						  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
							else if (ANFLast <> PO.AN) {
								COVID.PCR[PatID] += 1
								ANFLast := PO.PN "|" PO.AN
							}

						  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
						  ; Ct-Wert, Virusvarianten aus den Labortexten lesen
						  ;   [LAB-CtWert, LAB_Mutant und LAB_MutantEx sind
						  ;   die RegEx Variablen mit den Ergebnissen]
						  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
							If RegExMatch(PTEXT, rxStr, LAB_) {

							  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
							  ; Ct-Wert gefunden
							  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
								If LAB_CtWert {

									COVID.Ct         	+=	LAB_CtWert
									COVID.CtCount 	+=	1
									COVID.CtMin    	:=	LAB_CtWert < COVID.CtMin 	? LAB_CtWert : COVID.CtMin
									COVID.CtMax    	:=	LAB_CtWert > COVID.CtMax 	? LAB_CtWert : COVID.CtMax

									For Laborparameter, PNdata in LabPat[Datum][PatID] {
										If RegExMatch(labval.TXT, rxMutant, LAB_) {

										  ; gefundener Name wird in die einfachere Bezeichnung der WHO umbenannt (Delta/Omikron..)
											mutant := LAB_MUTANT ? LAB_MUTANT : LAB_MUTANTEX
											mutant := RegExReplace(mutant, "[\n\r\s\v]")
											mutant := Warnen.Covid.Varianten[mutant]

										 ; Ct-Werrte zählen und summieren
											If !IsObject(COVID.Mutation[mutant])
												COVID.Mutation[mutant] := {"Ct":0, "CtCount":0, "CtMin":999, "CtMax":0, "count":0, "virus":mutant}
											COVID.Mutation[mutant].Ct          	+= 	LAB_CtWert
											COVID.Mutation[mutant].CtCount 	+= 	1
											COVID.Mutation[mutant].CtMin		:= 	LAB_CtWert < COVID.Mutation[mutant].CtMin 	? LAB_CtWert : COVID.Mutation[mutant].CtMin
											COVID.Mutation[mutant].CtMax   	:= 	LAB_CtWert > COVID.Mutation[mutant].CtMax	? LAB_CtWert : COVID.Mutation[mutant].CtMax
										}
									}

								}

							  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
							  ; SARS-CoV-Variante gefunden (Alpha|Beta|Gamma|.....)
							  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
								else If LAB_MUTANT && RegExMatch(LAB_MUTANT "|" LAB_MUTANTEX, "i)(" Warnen.Covid.WHONamen ")") {

									mutant := LAB_MUTANT ? LAB_MUTANT : LAB_MUTANTEX
									mutant := RegExReplace(mutant, "[\n\r\s\v]")
									mutant := Warnen.Covid.Varianten[mutant]
									If !IsObject(COVID.Mutation[mutant])
										COVID.Mutation[mutant] := {"Ct":0, "CtCount":0, "CtMin":999, "CtMax":0, "count":0, "virus":mutant}
									COVID.Mutation[mutant].count += 1

								}

							  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
								LAB_TXT := LAB_CtWert ?	"[" LAB_CtWert "]" : ""                		; LAB_TXT wenn Labortexte ausgegeben werden sollten
								LAB_MUTANT := LAB_MUTANTEX := LAB_CtWert := ""
								break

							}

						}

					}

				  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
				  ; andere Statistiken und noch mehr über COVID-19
				  ; ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
					If (Warnen.LabStatsKeys && !adm.filtering > 0) {
						If RegExMatch(PO.PN, Warnen.LabStatsKeys) {
							Loop {
								key := (PO.PN) . A_Index
								If !IsObject(Warnen.LabStats[key])
									break
								labstats[key] := !labstats.HasKey(key) ? 0 : labstats[key]
								If !(countif := Warnen.LabStats[key].countif)
									labstats[key] += 1
								else if RegExMatch(PO.PW, countif, match)
									If RegExMatch(match, Warnen.LabStats[key].matchif) {
										labstats[key] += 1
										COVID.Positiv[PatID] := !COVID.Positiv.HasKey(PatID) ? 1 : COVID.Positiv[PatID]+1
									}
							}
						}
					}
					;}

				  ; Tabelle wird gefiltert
				  checkFilter:
					If (adm.filtering=2 && RegExMatch(PO.PN, rxFilter, filtered))
						continue
					else if (adm.filtering=3 && !RegExMatch(PO.PN, rxFilter, filtered))
						continue

				  ;  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
				  ; Tabelle  ‧  ‧  ‧ 	einzelne Zellen vorbereiten (CSS Styles den Tabellenfelder zuordnen)
				  ;  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧;{
				  ; temporäre Variablen
					TDPATBUTTONS 	:= (Patient ? " class='table-btn' onclick='ahk.LabJournal_KarteiKarte(event)'" : "")
					TDR4	            	:= StrReplace(TDR3, "#TT1#", PO.PB)         	; Beschreibung des Laborparameter (PB = Parameterbeschreibung)
					TDR4	            	:= StrReplace(TDR4, "#TT2#", "event")         	; Beschreibung des Laborparameter (PB = Parameterbeschreibung)
					html_ID              	:= PatID "_" RegExReplace(Patient, "[,\s]+", "_")

				  ; Filter anwenden um Tabellenzeilen zu kennzeichnen
					rowID := ""
					For filtername, filter in Warnen.Filter
						If RegExMatch(PO.PN, "i)" filter.rxFilter) {
							filter.count := !filter.count ? 1 : filter.count+1
							rowID := filtername . filter.count
							Warnen.Filter.count := !Warnen.Filter.count ? 1 : Warnen.Filter.count+1
							break
						}
					rowID := !rowID ? "tblrow" . thisrow++ : rowID

				  ; Datum                                      Hintergrund=Farbe1        Farbe2
					If Firstrow {
						trdcolor := "tds"
						trdcolor_last := "td1a"
						Firstrow := false
					} else if NewPat {
						trdnr 	:= InStr(trdcolor_last, "1") ? "2" : "1"
						trdcolor	:= "td" trdnr (UDatum ? "a": "b")
						trdcolor_last := "td" trdnr "b"                                    ; immer randlose Zeilen nach Datumwechsel mit Rand einstellen
						NewPat := false
					}  else
						trdcolor := trdcolor_last

					TRDAT   	:= "<TR class='table-row " trdcolor "' id='" rowID "'>"
					TDAFN		:= UDatum ? TDR2a : TDR2b

				  ; ID & Patientenname
					TDPAT := TDPATNR 	:= RTrim(TDLR1, ">") (TDPATBUTTONS ? TDPATBUTTONS " id='" html_ID : "") "'>"

				  ; Parameter Name        PO.CV = CAVE          1=red    2=blueviolet bold   3=     4=blue    5=thin red   6=black      7=blueviolet normal
				  ; für ungruppierte Laborwerte (wenn es nur einen Grenzwert gibt) wenn Wert größer ist aber normalerweiser kleiner sein sollte oder umgekehrt
				  ; dann mit rot pathologisch kennzeichnen
					colNr1      := PO.CV = 0 ? (PO.PL > 0 ? 5 : 4) : (PO.CV = 1 ? 1 : 2)
					colNr3   	:= colNr1 = 2 ? 7 : colNr1
					TDPN     	:= StrReplace(TDR4, "'>", cColW[colNr1] "'>")

				  ; Parameter Wert
					colNr2       := PO.CV = 0 ? (PO.PL > 0 ? 5 : 4) : PO.CV = 1 ? 1 : PO.PW = "NEGATIV" ? 6 : 2
					TDPW		:= StrReplace(TDAB, "'>", cColW[colNr2] "'>")

				  ; Parameter Normwerte
					TDPNW   	:= StrReplace(TDR1, "'>" , cColW[PO.PE ? colNr2 : colNr3] "'>")

				  ; Parameter Einheit
					TDPE     	:= StrReplace(TDL1, "'>" , cColW[colNr2] "'>")

				  ; Parameter Lage (Hinweiszeichen)
					TDPL     	:= StrReplace(TDC, "'>" , cColW[colNr2] "'>")
					If (adm.filtering=3 && Warnen.Filter.Abstriche.state ) {
						tel 	:= "Tel1: " PatDB[PatID].TELEFON "`nTel2: " PatDB[PatID].TELEFON2 "`nFax:" PatDB[PatID].TELEFAX
						rpl	:= cColW[2] "' onclick='alert(" tel ")'>"
						TDPL		:= StrReplace(TDC, "'>" , rpl) "▷"     ; ▽
					}

				; Abweichung oberhalb der Norm und PO.CV (CAVE) nicht wahr (CV - wenn wahr dann andere Farbkennzeichnung)
					If (PO.PL > 250 && !PO.CV) {
						TDAB      	:= StrReplace(TDAB 	, ">"	, " " cColW.3 "'>")
						TDR1X     	:= StrReplace(TDR1	, ">"	, " " cColW.3 "'>")
						TDCX      	:= StrReplace(TDC	, ">"	, " " cColW.3 "'>")
						TDL1X     	:= StrReplace(TDL1	, ">"	, " " cColW.3 "'>")
					}

				; Parameterabweichung
					If PO.PA {
						RegExMatch(PO.PV, "(?<V>[\<\>])*(?<SU>[\d.]+)\-*(?<SO>[\d.]+)*", P)
						PSM  	:= PSO ? PSO : PSU
						PA     	:= RegExReplace(Round(PO.PA * (PSM/100), 1), "\.0+$")
					}

					PE         	:= Trim(PO.PE)	? PO.PE : ""                                                               	;" - - - - -"
					tbDatum	:= UDatum ? SubStr(UDatum, 1, 6) : ""
				;}

				  ;  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
				  ; Tabelle  ‧  ‧  ‧   	eine komplette Zeile schreiben
				  ;  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧;{
					TRstart:= "`t"   	TRDAT                                                                  "`n"
					htmp:=	"`t`t "	TDAFN  	. 	tbDatum                             	"</TD>`n"	; Abnahmedatum                	[Datum]
							. 	"`t`t "	TDPATNR	. 	(PatID ? "[" PatID "]" : "")         	"</TD>`n"	; Patientennummer                [Patient]
							. 	"`t`t "	TDPAT     	.	Patient             	                  	"</TD>`n"	; Patientenname                    [Patient]
							. 	"`t`t "	TDPAT     	.	PatGeburt        	                  	"</TD>`n"	; Patientengeburtstag           	[Patient]
							. 	"`t`t "	TDAFN    	.	ANFNR          	                  	"</TD>`n"	; Anforderungsnummer       	[Patient]
							. 	"`t`t "	TDPN     	.	PO.PN                                 	"</TD>`n"	; Parameter Name              	[Param]
							. 	"`t`t "	TDPW    	.	PO.PW                               	"</TD>`n"	; Parameter Wert               	[Wert]
							.	"`t`t "	TDPNW   	.	PO.PV                                	"</TD>`n"	; Parameter Normwerte        	[Normalwerte]
							.	"`t`t "	TDPE     	.	(PE ? PE : LAB_TXT)            	"</TD>`n"	; Parameter Einheit                [      ]     o./u. Labortext (z.B. CtWert)
							. 	"`t`t "	TDPL     	.	(PO.PL	? PO.PL "x" 	: "")     	"</TD>`n"	; Parameter Lage                   [+/-]
							. 	"`t`t "	TDAB       	.	(PA   	? PA " " PE  	: "") 		"</TD>`n"	; Parameter Abweichung     	[⍉ Abw.]
					TRend:="`t</TR>`n"

					If !adm.filtering
						html .= TRstart . htmp . TRend
					else {

						RegExMatch(TRstart, "i)class\s*\=\s*'(?<class>.*?)'\s+id\s*\=\s*'(?<id>.*?)'", row_)
						tr := neutron.doc.createElement("TR")
						tr.className	:= row_class
						tr.id            	:= row_id
						tr.innerHTML	:= htmp
						parentNode.appendchild(tr)

					}


				; Fortschritt
					If (!adm.LabJ.preRender), load_Bar.Set(percent+=pstep, "erstelle die Journalanzeige")
				;}

					LAB_TXT := PE := PA := UDatum := PatID := Patient := PatGeburt := ANFNR := TRstart := htmp := TRend := tr := ""

				}


			}

			DatumLast := Datum

		}

		if !adm.filtering
			html .= "`n`t`t`t </tbody>`n`t `t </table>`n</div>`n</div>`n</body>`n`n</html>  `n`n"

		;}

	; Statistikausgabe erstellen                          	;{
		if !adm.filtering {

			virusNr := 0, arr := Array(), covtext := Array()
			For key, val in labstats
				If RegExMatch(key, "i)^COVI") {
					arr.Push(val)
					replacewith := arr.Count()=1 ? val : val . " Untersuchungen, somit in " Floor(arr.2*100/arr.1) "% aller Fälle,"
					covtext.Push(StrReplace(Warnen.LabStats[key].output, "#", replacewith))
				}

			anzahlder	:= COVID.PCR.Count()
			Positiv     	:= COVID.Positiv.Count()
			Mutations	:= COVID.Mutation.Count()
			doubles 	:= arr.2 - Positiv
			n1         	:= Positiv>1 ? "en" : ""
			n2a        	:= Mutations>1 || Mutations=" keine" ? "n ": " "
			n2b        	:= Mutations=" keine" ? "n ": " "
			n3         	:= doubles = 1 ? "" : "en"

			RegExMatch(adm.LabJ.Tagesanzeige, "(?<age>\d+)\s+Werktage", T)

		; Statistiktabelle 1
		  ; ab ##abDatum## (##Tage## Werktage)
			html := StrReplace(html, "??", Tage)
			html := StrReplace(html, "#cst1c2r1#", arr.2 " / " arr.1 " (" Floor(arr.2*100/arr.1) "%) [" arr.1-anzahlder " Doppeltest]" )
			html := StrReplace(html, "#cst1c1r2#", COVID.CtCount " Ct-Werte")
			html := StrReplace(html, "#cst1c2r2#", "min:" COVID.CtMin ", max:" COVID.CtMax ", Ø:" Floor(COVID.Ct/COVID.CtCount))
			For mutant, virus in COVID.Mutation {
				tinsert .= "`t`t<tr><td class='cststd1'>" (Mutations = 0 ? " keine" : Mutations) " Virusvariante" (Mutations>1 ? "n":"") ":</td>  "
				tinsert .= "<td>" virus.count "x" mutant "`nCt:" virus.CtMin "-" virus.CtMax ", Ø:" Floor(virus.Ct/virus.CtCount) "</td></tr>`n"
			}
			html := StrReplace(html, "<!-- ##01t## -->", tinsert)

		; als Ausgabetext
			If opt.statstext {
				statstext 	:= "SARS-CoV-2 Statistik des Ausgabezeitraumes:`n`n"
				statstext 	.= "In den letzten 31 Tagen wurden " anzahlder " Personen mittels PCR-Test auf SARS-CoV-2 untersucht. "
				statstext 	.= covtext.1 ". " covtext.2 ". "
				statstext 	.=  doubles > 0 	? doubles " Person" n3 " wurde" n3 " doppelt/mehrfach untersucht. "
															: "Es wurden keine Doppel-/Mehrfachuntersuchungen durchgeführt.`n"
				statstext 	.= "In " COVID.CtCount " Fällen wurde der Ct-Wert vom Labor angegeben. "
				statstext 	.= "Der höchste Ct-Wert lag bei " COVID.CtMax " und der niedrigste bei " COVID.CtMin " (Durchschnitt: " Floor(COVID.Ct/COVID.CtCount) ")."

				statstext 	.= "Es wurde" n2a . Mutations " besorgniserregende" n2b "SARS-CoV-2-Variante" n2a "nachgewiesen.`n"
				For mutant, virus in COVID.Mutation {
					virusNr ++
					statstext .= "Die " mutant "-Variante " (virusNr=1?"wurde":"") " bei " virus.count (virusNr=1?" der " anzahlder " untersuchten Patienten gefunden":" Patienten") ". ["
					statstext .= mutant "-Werte: Ct-Wert bestimmt " virus.CtCount "x, Ct-Min " virus.CtMin ", Ct-Max " virus.CtMax ", CTAvg " Floor(virus.Ct/virus.CtCount) "]`n"
				}
			}

		}
	;}

	; letzte Daten des Laborabrufes hinzufügen    	;{
		if !adm.filtering {
			abrufdaten := Labjournal_LababrufDaten()
			html := StrReplace(html, "<!-- #02la## -->", abrufdaten)
		}
	;}

	; erstellte HTML Seite anzeigen               		;{
		If (!adm.LabJ.preRender && Anzeige && !adm.filtering > 0)
			neutron := LabJournal_Viewer(html)

	; Progress beenden
		If (!adm.LabJ.preRender) {
			load_Bar.Set(100, "Das Laborjounal ist geladen!")
			SetTimer, loader_BarGUIDestroy, -1500
		}

	;}

	; speichern als BMP zum Versenden als Bilddatei
		;hBMP := hWnd_to_hBmp(labJ.hwnd)
		;SavePicture(hBMP, A_Temp "\Laborjournal.jpg")
		;Run % A_Temp "\Laborjournal.bmp"

return

loader_BarGUIDestroy:
	Gui, load_BarGUI: Destroy
	adm.filtering := false
return
}

Laborjournal_Redraw(neutron, filterstate, parentNode) {                                     	; nach Einstellungsänderung z.B. Filter die Tabelle neu zeichnen lassen

	global LabPat, Anzeige, Warten, COVIDStats

	adm.tbodynode := parentNode
	LaborJournal(LabPat, true, Warnen, COVIDStats)

}

Laborjournal_BaseHTML(htmltable) {

		global srchdrecords, dbrecords, Tagesanzeige

		srchdrecords  	:= LTrim(labJ.srchdrecords, "0")
		dbrecords     	:= labJ.records
		Tagesanzeige	:= adm.labJ.Tagesanzeige
		logo               	:= Labjournal_Logo()

	; HTML Seitendaten id='\d+_.*'>[a-zA-ZÄÖÜäöüß\s\-]+,[a-zA-ZÄÖÜäöüß\s\-]+|\[\d-]\)
		htmlheader := FileOpen(adm.LabJ.htmlresource, "r", "UTF-8").Read()
		htmlheader := StrReplace(htmlheader, "<!--##01##-->", "`n" logo "`n")
		htmlheader := StrReplace(htmlheader, "##svgoptions##", "style='width:32px; height:24px; padding-top:2px;margin-left:5px;'")
		htmlheader := StrReplace(htmlheader, "##02##", Tagesanzeige)
		htmlheader := StrReplace(htmlheader, "<!--##03##-->", "`n")
		htmlheader := StrReplace(htmlheader, "<!-- ##02t## -->", htmltable)

		htmlbody =
		(
			<div style='overflow-x:scroll; overflow-y:scroll; width:100`%; height:92`%;'>
				<div class="table-FixHead">
					<table class='LabJ_Table'">
						<thead>
							<tr>
								<th style='text-align:Right'>Datum</th>
								<th style='text-align:Center' colspan='3'>[NR] Name, Vorname  Geburtsdatum</th>
								<th style='text-align:Center'>Anford.Nr</th>
								<th style='text-align:Center'>Param</th>
								<th style='text-align:Right'>Wert</th>
								<th style='text-align:Center' colspan='2'>Normalwerte</th>
								<th style='text-align:Right'>+/-</th>
								<th style='text-align:Right'>⍉ Abw.</th>
							</tr>
						</thead>

						<tbody>
		)

return htmlheader . htmlbody
}

LabJournal_CovidStats(data) {                                                                            	; erstellt einen Teil der HTML SARS-CoV-2 Statistiken

	If !IsObject(data) {
		If !RegExMatch(data, "i)^[A-Z]:\\.*\.json$") {
			MsgBox, 0x1000, % A_ThisFunc ": Der data-Parameter muss ein Objekt mit den statistischen Daten oder ein Pfad zu den Daten im json Format sein!"
			ExitApp
		}
		data := cJSON.Load(FileOpen(data, "r", "UTF-8").Read())
	}

	COVID := data

  ; Texttabelle und Teile der HTML Tabelle erstellen
	For variant, variantCount in COVID.mutants
		lmax := StrLen(variant) > lmax ? StrLen(variant) : lmax
	posv := 0, cols := Object()
	For variant, variantCount in COVID.mutants {
		posv ++
		cols[variant] := posv
		th .= "`t`t`t`t<th class='" (posv=COVID.mutants.Count() ? "thr":"thc") "'>" variant "</th>`n"
	}

	totalPositive := totalPCR := 0, totalVariants := Object()
	Loop % SubStr(A_YYYY, 4, 1)+1 {

		examyear 	:= "202" A_Index-1

		PercentPositive := Round(COVID.weeks[examyear].PCRpositive*100/COVID.weeks[examyear].PCRcount, 1)
		PositiveToPCRcount := COVID.weeks[examyear].PCRpositive " / " COVID.weeks[examyear].PCRcount " (" PercentPositive "%)"
		ws := SubStr("                 ", -1*(Floor((StrLen(PositiveToPCRcount)/4)-1)))

		td .= "`t`t<tr>`n`t`t`t"
		td .= "<td>" examyear "</td>"
		td .= "<td>" COVID.weeks[examyear].PCRpositive "</td>"
		td .= "<td>" COVID.weeks[examyear].PCRcount "</td>"

		totalPositive 	+= COVID.weeks[examyear].PCRpositive
		totalPCR    	+= COVID.weeks[examyear].PCRcount

		Loop % COVID.mutants.Count()
			td .= "<td>#" A_Index "#</td>"

		posv := 1
		For variant, variantCount in COVID.weeks[examyear].mutants {

			If !totalVariants.HasKey(variant)
				totalVariants[variant] := 0

			totalVariants[variant] += variantCount ? variantCount : 0
			td := StrReplace(td, "#" cols[variant] "#", variantCount)

		}

		td := RegExReplace(td, "(#\d+#)")
		td .= "`n`t`t</tr>`n"

	}

	td .= "`t`t<tr class='tb2last'>`n`t`t`t"
	td .= "<td></td><td>" totalPositive "</td><td>" totalPCR "</td>"
	Loop % totalVariants.Count()
		td .= "<td>#" A_Index "#</td>"
	For variant, variantCount in totalVariants
		td := StrReplace(td, "#" cols[variant] "#", variantCount)
	td := RegExReplace(td, "(#\d+#)")

	table =
	(
	<table class='tablecovidstats2'>
		<thead>
			<tr>
				<th class='thl'>Jahr</th>
				<th class='thc'>Positiv</th>
				<th class='thc'>PCR's</th>
%th%
			</tr>
		</thead>
		<tbody>
%td%
		</tbody>
	</table>

	)

return table
}

LabJournal_Viewer(html)  {                                                                                	; übergibt die HTML Daten an den Internet-Explorer

	global neutron
	static winpos

  ; letzte Fensterposition laden
	IniRead, winpos, % adm.ini, % adm.compname, LaborJournal_Position
	If (InStr(winpos, "ERROR") || StrLen(winpos) = 0)
		winpos := "w1050 h500"

	neutron := new NeutronWindow(html, "", "", "Laborjournal" LabJ.maxrecords ")", " +AlwaysOnTop minSize1045x300")
	neutron.Gui("+LabelNeutron")

	obj := neutron.wb.document.getElementById("LaborJournal_Header")	, hcr	:= obj.getBoundingClientRect()
	obj := neutron.wb.document.getElementById("LabJournal_Table")    	, tcr	:= obj.getBoundingClientRect()

	labJ.hwnd := neutron.hwnd
	labJ.HeaderHeight	:= Floor(hcr.Bottom + 1)
	labJ.TableHeight   	:= Floor(tcr.Bottom  + 1)
	labJ.enrolled	    		:= true

  ; Monitordaten erfassen
	hMonitor 	:= MonitorFromWindow(labJ.hwnd)
	mon     	:= GetMonitorInfo(hMonitor)
	TBHeight	:= TaskbarHeight()

  ; Fensterposition anpassen
	npos	:= PosStringToObject(winpos)
	npos.X	:= npos.X 	< mon.L	? mon.L	: npos.X
	npos.Y	:= npos.Y 	< mon.T	? mon.T	: npos.Y
	npos.W := npos.W	< 1045	? 1045 	: npos.W
	npos.H	:= npos.H 	< 600   	? 600
				: 	 npos.H		> mon.B	? mon.B - TBHeight 	: npos.H
	winpos :=	"x" npos.X " y" npos.Y " w" npos.W " h" npos.H
	If !npos.Maximize && !(isInside := IsInsideVisibleArea(npos.X, npos.Y, npos.W, npos.H, CoordInjury)) {
		winpos :=	"x" 	(InStr(CoordInjury, "x") ? "0"     	: npos.X)
					. 	" y" 	(InStr(CoordInjury, "y") ? "0"     	: npos.Y)
					. 	" w" 	(npos.W	< 1045	 ? 1045	: npos.W)
					.	" h" 	(npos.H 	< 600   	 ? 600   	: npos.H)
	}


  ; Breite der Überschrift an die Tabellenbreite und Höhe der Tabelle 1 an Tabelle 2  anpassen
	Table1 := neutron.wb.document.getElementsByClassName("tablecovidstats")[0]
	tb1 := Table1.getBoundingClientRect()
	Table2 := neutron.wb.document.getElementsByClassName("tablecovidstats2")[0]
	tb2 := Table2.getBoundingClientRect()

  ; Überschrift
	TitleWidth	:= Floor(tb2.Right - tb1.Left)
	innerTitle	:= neutron.wb.document.getElementsByClassName("Statistik-Title")[0]
	innerTitle.Style.Width := TitleWidth "px"

  ; Tabelle
	Table1.Style.Height := tb2.Bottom - tb2.Top "px"

  ; Anzeigen
	Gui, load_BarGUI: Destroy

	;~ SciTEOutput(A_ThisFunc "- " labJ.hwnd " = " neutron.hwnd ", " TBHeight " , "  winpos)
	neutron.Show(winpos)
	SetWindowPos(labJ.hwnd, npos.X, npos.Y, npos.W, npos.H)

	;~ OnMessage(0x200, "LabJournal_Eyes")


return neutron
}

LabJournal_ParamToolTip(neutron, event) {

	ToolTip, % event

}

LabJournal_KarteiKarte(neutron, event) {                                                               	; Neutron Klickfunktion: Karteikarte anzeigen

	; event.target will contain the HTML Element that fired the event.
	; Show a message box with its inner text.
		RegExMatch(event.target.id, "^\s*(?<ID>\d+)?_(?<Name>.*)$", Pat)

		PatName := StrReplace(PatName, "_", ", ")
		SciTEOutput(A_ThisFunc ": [" PatID "] " PatName)

		AlbisActivate(1)
		If AlbisAkteOeffnen(PatName, PatID) {
			Sleep 500
			AlbisLaborBlattZeigen()
		}
		else
			PraxTT("Die Patientenakte von <[" (PatID ? PatID : 0) "] " (PatName ? PatName : "---, ---") "> konnte nicht geöffnet werden!")

return
}

LabJournal_Filter(neutron, event) {

		global parentNode

		static viruscolor 	:= ["D47A04", "E18484", "3C9F4D"]
		static colflag  		:= 0

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; COVID Filter ein- oder ausschalten
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		colflag := colflag+1 < 3 ? colflag+1 : 0
		Warnen.Filter.Abstriche.state := colflag
		adm.filtering := colflag + 1

	  ; ―――――――――――――――――――――――――――――――
      ; nach dem ersten Druck trägt der Virus eine Maske sonst nicht
	  ; ―――――――――――――――――――――――――――――――
		mask := neutron.doc.getElementsByClassName("mask")[0]
		try mask.Style.visibility := (colflag=1 ? "visible" : "hidden")
		catch e
			 MsgBox, 16,, % "Ausnahme ausgelöst!`n`nWas: " e.what "`nDatei: " e.file
					. "`nZeile: " e.line "`nNachricht: " e.message "`nZusatz: " e.extra

	  ; ―――――――――――――――――――――――――――――――
	  ; auf rosafarbender Virus bei state=1, grün bei state = 2, sonst orange
	  ; ―――――――――――――――――――――――――――――――
	  ; das Gesicht
		COVID := neutron.doc.getElementsByClassName("facecolor")[0]
		try COVID.Style.fill := "#" viruscolor[colflag+1]
		catch e
			 MsgBox, 16,, % "Ausnahme ausgelöst!`n`nWas: " e.what "`nDatei: " e.file
					. "`nZeile: " e.line "`nNachricht: " e.message "`nZusatz: " e.extra

	  ; ―――――――――――――――――――――――――――――――
	  ; und die Spikes einfärben
	  ; ―――――――――――――――――――――――――――――――
		Spikes := neutron.doc.getElementsByClassName("Spike")
		Loop, % Spikes.length {
			try Spikes[A_Index-1].Style.fill := "#" viruscolor[colflag+1]
			catch e
			 MsgBox, 16,, % "Ausnahme ausgelöst!`n`nWas: " A_Index-1 ": " e.what "`nDatei: " e.file
					. "`nZeile: " e.line "`nNachricht: " e.message "`nZusatz: " e.extra
		}

	  ; ―――――――――――――――――――――――――――――――
	  ; bei state = 2 soll er lächeln
	  ; ―――――――――――――――――――――――――――――――
		mouth := neutron.doc.getElementsByClassName("mouth")[0]
		try mouth.Style.Transform := colflag = 2 ? "translateX(50px) translateY(90px) rotate(165deg)"
																	  : "translateX(0px) translateY(0px) rotate(0deg)"
		catch e
			 MsgBox, 16,, % "Ausnahme ausgelöst!`n`nWas: " e.what "`nDatei: " e.file
					. "`nZeile: " e.line "`nNachricht: " e.message "`nZusatz: " e.extra


		;~ try mouth.Style.animation-name := colflag = 2 ? "moveMouth1" : "moveMouth"
		;~ catch e
			 ;~ MsgBox, 16,, % "Ausnahme ausgelöst!`n`nWas: " e.what "`nDatei: " e.file
					;~ . "`nZeile: " e.line "`nNachricht: " e.message "`nZusatz: " e.extra



		; mouth
		;transform: translateX(50px) translateY(90px) rotate(165deg);
		;translateX(0px) translateY(0px) rotate(0deg);

		;eye-brow
		;transform: translateX(90px) translateY(90px) rotate(175deg);
		;transform: translateX(90px) translateY(87px) rotate(175deg);

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; schnellste Methode um die komplette Tabelle zu entfernen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		node := neutron.doc.getElementsByClassName("table-row")[0]
		parentNode := node.parentNode

		nodelist := parentNode.querySelectorAll("TR")
		Loop, % nodelist.length
			parentNode.removeChild(nodelist[nodelist.length-A_Index])

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 ; so und jetzt die Tabelle je nach Filtereinstellung neu aufbauen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Laborjournal_Redraw(neutron, colflag, parentNode)

}

LabJournal_Eyes(lparam, wparam) {                                                                   	; soll den Corona-Virus animieren

	global neutron, h
	static reye, reyeX, leye, leyeY

	CoordMode, Mouse, Window

	win := GetWindowSpot(labJ.hwnd)

	If !IsObject(reye) {
		reye := neutron.doc.getElementsByClassName("right-eyeball")[0]
		reyeX := reye.cx
		;~ reyeX := reye.getBoundingClientRect().Left
	}
	If !IsObject(leye) {
		leye := neutron.doc.getElementsByClassName("left-eyeball")[0]
		leyeX := leye.cx
		SciTEOutput("x: " leye.Count())
	}

	MouseGetPos, mx, my
	lmx := mx > leyeX ? mx-leyeX : leyeX-mx
	lx 	:= -1 + 6*(mx/1045)
	rx	:= -2 + 10*(mx/1045)
	leye.transform := "translateX(" lx "px)"
	reye.Style.transform := "translateX(" rx "px)"
	ToolTip %  lx "|" rx
}

LabJournal_DragTitleBar(event) {

	If labJ.enrolled {
		PostMessage, 0xA1, 2, 0,, % "ahk_id" labJ.hwnd  ; WM_NCLBUTTONDOWN := 0xA1
	}
	else If (labj.enrolled = "x") {

		lj := GetWindowSpot(labJ.hwnd)
		dh	:= lj.h
		dtb	:= labJ.TableHeight - labJ.HeaderHeight + 10
		dStep:= Floor(dtb/25)
		Loop 25 {
			dh += dStep
			SetWindowPos(labJ.hwnd, lj.x, lj.y, lj.w, dh)
		}

		SetWindowPos(labJ.hwnd, lj.x, lj.y, lj.w, labJ.TableHeight - labJ.HeaderHeight)
		labJ.enrolled := true

	}

return
}

LabJournal_Close(event) {

	SaveGuiPos(labJ.hwnd)
	Gui, % labJ.hwnd ": Destroy"
	ExitApp

}

LabJournal_Grenzen(event) {                                                                               	; Neutron Klickfunktion: AlbisLaborwertGrenzen
  AlbisLaborwertGrenzen(adm.LabDBPath)
}

LabJournal_MoreTips(neutron, event) {

	global LabDic

	LabPB 	:= LabDic[event.target.innerText].1
	RegExMatch(LabPB, "^(?<ledge>.+?)(,|$)", Lab)

	;~ Run, % "https://de.wikipedia.org/wiki/" LabPB
	;~ SciTEOutput("url: https://de.wikipedia.org/wiki/" LabPB)

	;~ ToolTip, % knowledge
	;~ SetTimer, ttaus, -6000

	knowledge := DownloadToString("https://de.wikipedia.org/wiki/" LabPB)
	If InStr(knowledge, "Diese Seite existiert nicht") {
		ToolTip, % "Suchbegriff <" LabPB "> nicht gefunden"
		SetTimer, ttaus, -6000
		return
	}
	RegExMatch(knowledge, "\<p\>(?<PB>.+?)\<\/p\>", know)

	SciTEOutput(knowledge)
	;knowledge := RegExReplace(knowledge, "(\<.*?\>)")
	;~ URLDownloadToFile, % "https://de.wikipedia.org/wiki/" LabPB, % A_Temp "\wikipedia.html"
return knowledge
ttaus:
	ToolTip
return
}

Labjournal_LababrufDaten() {                                                                            	; berechnet die Zeit bis zum nächsten Laborabruf

	IniRead, AClient	, % adm.Ini, % "LaborAbruf"	, % "OnlyRunOnClient"
	If InStr(AClient, "ERROR") || !AClient
		return
	IniRead, ATimer  	, % adm.Ini, % AClient       	, % "Laborabruf_Timer"

	IniRead, LCall1  	, % adm.Ini, % "LaborAbruf"	, % "Letzter_Abruf"
	IniRead, LCall2  	, % adm.Ini, % "LaborAbruf"	, % "Letzter_Abruf_ohne_Daten"
	IniRead, LCall3  	, % adm.Ini, % "LaborAbruf"	, % "Letzter_Abruf_mit_Daten"

	pdate 	:= A_YYYY A_MM A_DD "000000"
	heute 	:= A_MM A_DD
	pdate 	+= -1, days
	gestern	:= SubStr(pdate, 5, 2) SubStr(pdate, 7, 2)
	pdate 	+= +2, days
	morgen	:= SubStr(pdate, 5, 2) SubStr(pdate, 7, 2)

	If RegExMatch(LCall1, "(?<Y>\d{4})-(?<M>\d{2})-(?<D>\d{2})\s+(?<H>\d{2}):(?<Min>\d{2}):(?<S>\d{2})[\s\|]*", T) {
		Abruftag := TM TD
		lastLBCall := (	Abruftag = heute  	? "heute "
							: 	Abruftag = gestern 	? "gestern " : TD "." TM "." )
						. 	", " TH ":" TMin " Uhr<br>"
						. 	(LCall1 = LCall2 ? "keine neue" : LCall1 = LCall3 ? "neue" : "??") " Daten"
	}

	If !ATimer
		nextLBCall := "der zeitgesteuerte Abruf ist aus"
	else {
		IniRead, nextCall, % adm.Ini, % "LaborAbruf", % "naechster_Abruf"
		If RegExMatch(nextCall, "(?<D>\d{2})\.(?<M>\d{2})\.(?<Y>\d{4})\s*,\s*(?<hm>\d+:\d+)", T) {
			Abruftag 	:= TM TD
			nextLBCall	:= ( Abruftag = heute  	? "heute "
								: Abruftag = morgen	? "morgen " :  TD "." TM )	. ", " Thm " Uhr"
		}
	}

	lahtml =
	(
		<div class="LaborAbruf">
			<div class="laRow">
				<tspan class='laInfo1'>letzter Abruf</tspan>
				<br>
				<tspan class='laInfo2'>%lastLBCall%</tspan>
			</div>
			<div class="laEmpty"></div>
			<div class="laRow">
				<tspan class='laInfo1'>nächster Abruf</tspan>
				<br>
				<tspan class='laInfo2'>%nextLBCall%</tspan>
			</div>
	  </div>
	)


return lahtml
}

Labjournal_NewLabData() {

	FileGetTime, tmp, % adm.AlbisDBPath "\LABBLATT.dbf", M

	If !adm.Labj.lastUpdate
		adm.Labj.lastUpdate := IniReadExt("Laborjournal", "letztes_Journalupdate")

	SciTEOutput(A_ThisFunc ": " adm.Labj.lastUpdate ", " tmp)

	If (tmp <> adm.Labj.lastUpdate) {
		adm.Labj.lastUpdate := tmp
		return adm.Labj.Override := true
	}

return false
}

Labjournal_NewLabParams() {

	FileGetTime, tmp, % adm.AlbisDBPath "\LABPARAM.dbf", M

	If !adm.Labj.lastParams
		adm.Labj.lastParams := IniReadExt("Laborjournal", "LabParam_letztesUpdate")

	If (tmp <> adm.Labj.lastParams)
		return tmp

return false
}

Labjournal_Knowledge() {

	/* vielleicht lassen sich Zusammenhänge zwischen einzelnen Veränderungen von Laborwerten analysieren

		Hypalbuminämie
		- 	wenn Albumin i. Serum erniedrigt ist, steigt die Cholinesterase (CHE) an, konsekutiv oder Angabe in Prozent findet sich in der Elektrophorese
			je nach Erniedrigung von Albumin ein höherer Anteil von Beta- und Gamma-Globulinen


	 */

}

LoadKnowlegde(LabPB) {

	LabPB 	:= RegExReplace(LabPB, "zyten", "zyt")
	html  	:= DownloadToString("https://de.wikipedia.org/wiki/" LabPB)
	html  	:= RegExReplace(html, "^.*\<div\sclass\=" q "mw\-parser\-output" q "\>" q)
	RegExMatch(html, "\<p\>(?<ledge>.*?)\<\/p\>", know)

return knowledge
}

Menu_LabJournal(event) {                                                                                 	; im Moment nur Reload

	SaveGuiPos(labJ.hwnd)
	;~ MsgBox, 0x1004, Laborjournal, Tabelle neu erstellen?
	;~ IfMsgBox, Yes
	If Labjournal_NewLabData() 	{
		FileDelete, % adm.PathLabj
		FileDelete, % adm.LabPat
		FileDelete, % adm.COVIDStats
	}
	Reload

}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
SaveGuiPos(hwnd) {

	MinMaxState := WinGetMinMaxState(hwnd)
	If (MinMaxState <> "i") {
		win 			:= GetWindowSpot(hwnd)
		winPos	 	:= "x" win.X " y" win.Y " w" win.CW " h" win.CH   (MinMaxState = "z" ? " Maximize " : " ")
		IniWrite, % winpos	, % adm.Ini, % adm.compname, LaborJournal_Position
	}

}

PosStringToObject(string) {

	p := Object()
	For wIdx, coord in StrSplit("XYWH") {
		RegExMatch(string, "i)" coord "(?<Pos>\d+)", w)
		p[coord] := !wPos ? 100 : wPos
	}

	p.Maximize := InStr(string, "Maximize") ? true : false

return p
}

MessageWorker(InComing) {                                                                                	;-- verarbeitet die eingegangen Nachrichten

		recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
					, 	"opt"     	: (StrSplit(InComing, "|").2)
					, 	"fromID"	: (StrSplit(InComing, "|").3)}

	; Laborabruf_*.ahk
		; per WM_COPYDATA die Daten im Laborjournal auffrischen (Skript wird neu gestartet)
		if RegExMatch(recv.txtmsg, "i)\s*reload") {
			Send_WM_COPYDATA("reload Laborjournal||" labJ.hwnd, recv.fromID)
			Reload
		}

return
}

hWnd_to_hBmp( hWnd:=-1, Client:=0, A:="", C:="" ) {                                       	;-- Capture fullscreen, Window, Control or user defined area of these

	; By SKAN C/M on D295|D299 @ bit.ly/2lyG0sN

	A      := IsObject(A) ? A : StrLen(A) ? StrSplit( A, ",", A_Space ) : {},     A.tBM := 0
	Client := ( ( A.FS := hWnd=-1 ) ? False : !!Client ), A.DrawCursor := "DrawCursor"
	hWnd   := ( A.FS ? DllCall( "GetDesktopWindow", "UPtr" ) : WinExist( "ahk_id" . hWnd ) )

	A.SetCapacity( "WINDOWINFO", 62 ),  A.Ptr := A.GetAddress( "WINDOWINFO" )
	A.RECT := NumPut( 62, A.Ptr, "UInt" ) + ( Client*16 )

	If (DllCall("GetWindowInfo",   "Ptr",hWnd, "Ptr",A.Ptr ) && DllCall("IsWindowVisible", "Ptr",hWnd ) && DllCall("IsIconic", "Ptr",hWnd )=0) {
        A.L	:= NumGet( A.RECT+ 0, "Int" )	, A.X 	:= ( A.1 <> "" ? A.1 : (A.FS ? A.L : 0) )
        A.T	:= NumGet( A.RECT+ 4, "Int" )	, A.Y 	:= ( A.2 <> "" ? A.2 : (A.FS ? A.T : 0 ))
        A.R	:= NumGet( A.RECT+ 8, "Int" )	, A.W	:= ( A.3  >  0 ? A.3 : (A.R - A.L - Round(A.1)) )
        A.B	:= NumGet( A.RECT+12, "Int" )	, A.H 	:= ( A.4  >  0 ? A.4 : (A.B - A.T - Round(A.2)) )

        A.sDC	:= DllCall( Client ? "GetDC" : "GetWindowDC", "Ptr",hWnd, "UPtr" )
        A.mDC	:= DllCall( "CreateCompatibleDC", "Ptr",A.sDC, "UPtr")
        A.tBM 	:= DllCall( "CreateCompatibleBitmap", "Ptr",A.sDC, "Int",A.W, "Int",A.H, "UPtr" )

        DllCall( "SaveDC", "Ptr",A.mDC )
        DllCall( "SelectObject", "Ptr",A.mDC, "Ptr",A.tBM )
        DllCall( "BitBlt",       "Ptr",A.mDC, "Int",0,   "Int",0, "Int",A.W, "Int",A.H, "Ptr",A.sDC, "Int",A.X, "Int",A.Y, "UInt",0x40CC0020 )

        A.R := ( IsObject(C) || StrLen(C) ) && IsFunc( A.DrawCursor ) ? A.DrawCursor( A.mDC, C ) : 0
        DllCall( "RestoreDC", "Ptr",A.mDC, "Int",-1 )
        DllCall( "DeleteDC",  "Ptr",A.mDC )
        DllCall( "ReleaseDC", "Ptr",hWnd, "Ptr",A.sDC )
    }

Return A.tBM
}

SavePicture(hBM, sFile) {                                                                                      	;-- By SKAN on D293 @ bit.ly/2krOIc9
Local V,  pBM := VarSetCapacity(V,16,0)>>8,  Ext := LTrim(SubStr(sFile,-3),"."),  E := [0,0,0,0]
Local Enc := 0x557CF400 | Round({"bmp":0, "jpg":1,"jpeg":1,"gif":2,"tif":5,"tiff":5,"png":6}[Ext])
  E[1] := DllCall("gdi32\GetObjectType", "Ptr",hBM ) <> 7
  E[2] := E[1] ? 0 : DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr",hBM, "UInt",0, "PtrP",pBM)
  NumPut(0x2EF31EF8,NumPut(0x0000739A,NumPut(0x11D31A04,NumPut(Enc+0,V,"UInt"),"UInt"),"UInt"),"UInt")
  E[3] := pBM ? DllCall("gdiplus\GdipSaveImageToFile", "Ptr",pBM, "WStr",sFile, "Ptr",&V, "UInt",0) : 1
  E[4] := pBM ? DllCall("gdiplus\GdipDisposeImage", "Ptr",pBM) : 1
Return E[1] ? 0 : E[2] ? -1 : E[3] ? -2 : E[4] ? -3 : 1
}

RegExMatchObject(Haystack, NeedleRegEx, StartingPos:=1) {                              	;-- wandelt RegExMatch-Objekt in ein Key:Value Objekt um

	rxobj := Object()

  ; O - first removes object option and then added O with or without )
	NeedleRegEx := RegExReplace(NeedleRegEx, "^([^\(]+)O(.*\))", "$1$2")
	NeedleRegEx	:= RegExMatch(NeedleRegEx, "^[^\(]+?\)") ? "O" : "O)"

  ; builds match object
	RegExMatch(Haystack, NeedleRegEx, Match, StartingPos)

  ; converts to Autohotkey key:value object
	Loop % Match.Count()
		rxObj[Match.Name(A_Index)] := Match.Value(A_Index)

return rxObj
}

Labjournal_Logo() {                                                                                           	;-- animiertes Logo (funktioniert nicht mit IE)

	svg =
	(
	<!-- Created with enve https://maurycyliebner.github.io -->
	<svg viewBox="0 0 48 48" ##svgoptions##>
	 <g transform="translate(140.233 188.648)">
	  <g transform="translate(-109.802 -160.482)">
	   <g transform="rotate(0)">
		<g transform="scale(5.2 5.25)">
		 <g transform="skewX(0) skewY(0)">
		  <g transform="translate(-140.233 -188.648)" opacity="1">
		   <g>
			<g transform="translate(140.233 188.648)">
			 <g transform="translate(0 0)">
			  <g transform="rotate(0)">
			   <g transform="scale(1 1)">
				<g transform="skewX(0) skewY(0)">
				 <g transform="translate(-140.233 -188.648)" opacity="1">
				  <path stroke-width="0.229718" fill="#a29f9f" stroke="none">
				   <animate attributeName="d" keyTimes="0;0.15;0.295;0.35;0.65;0.725;1" keySplines="0 0 1 1;0 0 1 1;0 0 1 1;0 0 1 1;0 0 1 1;0 0 1 1" calcMode="spline" dur="8.33333s" values="M140.944 185.562C140.944 185.562 140.974 185.996 140.974 185.996C141.513 187.487 142.107 188.937 142.616 190.18C142.769 190.554 142.752 190.933 142.562 191.28C142.382 191.581 142.102 191.681 141.779 191.732C141.779 191.732 141.654 191.734 141.654 191.734C141.654 191.734 138.767 191.734 138.767 191.734C138.767 191.734 138.642 191.732 138.642 191.732C137.867 191.658 137.567 190.911 137.814 190.187C138.446 188.732 139.096 187.207 139.698 185.959C139.698 185.959 139.713 185.562 139.713 185.562C140.162 185.618 140.499 185.562 140.944 185.562Z;M140.944 185.562C140.944 185.562 140.974 185.996 140.974 185.996C141.513 187.487 142.107 188.937 142.616 190.18C142.769 190.554 142.752 190.933 142.562 191.28C142.382 191.581 142.102 191.681 141.779 191.732C141.779 191.732 141.654 191.734 141.654 191.734C141.654 191.734 138.767 191.734 138.767 191.734C138.767 191.734 138.642 191.732 138.642 191.732C137.867 191.658 137.567 190.911 137.814 190.187C138.446 188.732 139.096 187.207 139.698 185.959C139.698 185.959 139.713 185.562 139.713 185.562C140.162 185.618 140.499 185.562 140.944 185.562Z;M142.251 189.057C142.251 189.057 142.417 189.29 142.417 189.29C142.651 189.778 142.789 190.216 142.964 190.681C143.041 190.887 143.09 191.066 142.9 191.413C142.72 191.714 142.102 191.681 141.779 191.732C141.779 191.732 141.654 191.734 141.654 191.734C141.654 191.734 138.767 191.734 138.767 191.734C138.767 191.734 138.642 191.732 138.642 191.732C137.867 191.658 137.407 191.625 137.31 191.273C137.24 190.997 137.587 190.267 137.713 190.064C137.713 190.064 138.242 189.11 138.242 189.11C139.452 189.379 141.254 189.375 142.251 189.057Z;M142.772 190.382C142.772 190.382 143.001 190.545 143.001 190.545C143.119 190.652 143.085 190.699 143.134 190.868C143.177 191.013 143.218 191.116 143.028 191.463C142.848 191.764 142.102 191.681 141.779 191.732C141.779 191.732 141.654 191.734 141.654 191.734C141.654 191.734 138.767 191.734 138.767 191.734C138.767 191.734 138.642 191.732 138.642 191.732C137.867 191.658 137.884 191.724 137.657 191.513C137.627 191.422 137.461 191.268 137.427 190.955C137.427 190.955 137.656 190.456 137.656 190.456C139.155 190.805 141.565 190.822 142.772 190.382Z;M142.772 190.382C142.772 190.382 143.001 190.545 143.001 190.545C143.119 190.652 143.085 190.699 143.134 190.868C143.177 191.013 143.218 191.116 143.028 191.463C142.848 191.764 142.102 191.681 141.779 191.732C141.779 191.732 141.654 191.734 141.654 191.734C141.654 191.734 138.767 191.734 138.767 191.734C138.767 191.734 138.642 191.732 138.642 191.732C137.867 191.658 137.884 191.724 137.657 191.513C137.627 191.422 137.461 191.268 137.427 190.955C137.427 190.955 137.656 190.456 137.656 190.456C139.155 190.805 141.565 190.822 142.772 190.382Z;M140.944 185.562C140.944 185.562 140.974 185.996 140.974 185.996C141.513 187.487 142.107 188.937 142.616 190.18C142.769 190.554 142.752 190.933 142.562 191.28C142.382 191.581 142.102 191.681 141.779 191.732C141.779 191.732 141.654 191.734 141.654 191.734C141.654 191.734 138.767 191.734 138.767 191.734C138.767 191.734 138.642 191.732 138.642 191.732C137.867 191.658 137.567 190.911 137.814 190.187C138.446 188.732 139.096 187.207 139.698 185.959C139.698 185.959 139.713 185.562 139.713 185.562C140.162 185.618 140.499 185.562 140.944 185.562Z;M140.944 185.562C140.944 185.562 140.974 185.996 140.974 185.996C141.513 187.487 142.107 188.937 142.616 190.18C142.769 190.554 142.752 190.933 142.562 191.28C142.382 191.581 142.102 191.681 141.779 191.732C141.779 191.732 141.654 191.734 141.654 191.734C141.654 191.734 138.767 191.734 138.767 191.734C138.767 191.734 138.642 191.732 138.642 191.732C137.867 191.658 137.567 190.911 137.814 190.187C138.446 188.732 139.096 187.207 139.698 185.959C139.698 185.959 139.713 185.562 139.713 185.562C140.162 185.618 140.499 185.562 140.944 185.562Z" repeatCount="indefinite"/>
				  </path>
				 </g>
				</g>
			   </g>
			  </g>
			 </g>
			</g>
		   </g>
		  </g>
		 </g>
		</g>
	   </g>
	  </g>
	 </g>
	 <g transform="translate(138.992 187.889)">
	  <g transform="translate(-115.018 -163.912)">
	   <g transform="rotate(0)">
		<g transform="scale(5.2 5.25)">
		 <g transform="skewX(0) skewY(0)">
		  <g transform="translate(-138.992 -187.889)" opacity="1">
		   <g>
			<g transform="translate(138.992 187.889)">
			 <g transform="translate(0 0)">
			  <g transform="rotate(0)">
			   <g transform="scale(1 1)">
				<g transform="skewX(0) skewY(0)">
				 <g transform="translate(-138.992 -187.889)" opacity="1">
				  <path stroke-width="0.229718" d="M134.684 183.47C134.62 183.47 134.569 183.522 134.569 183.585L134.531 192.192C134.53 192.255 134.582 192.307 134.645 192.306L137.089 192.307C136.799 192.192 136.749 192.101 136.631 191.937C134.861 191.937 136.523 191.937 134.9 191.937L134.939 183.876L138.669 183.876C138.688 183.712 138.49 183.586 138.466 183.47L134.684 183.47ZM138.669 184.479L135.32 184.48L135.32 184.874L138.669 184.874L138.669 184.479ZM138.669 185.808L135.321 185.809L135.321 186.203L138.556 186.203L138.669 185.994L138.669 185.808ZM137.901 187.415L135.321 187.414L135.321 187.814L137.698 187.815L137.901 187.415ZM136.988 189.251L135.32 189.251L135.321 189.663L136.784 189.663L136.988 189.251ZM136.308 190.986L135.32 190.985L135.32 191.38C135.32 191.38 136.224 191.381 136.35 191.381C136.284 191.187 136.308 190.986 136.308 190.986L136.308 190.986ZM138.982 183.47L139.315 183.813L139.315 185.994C138.618 187.254 137.962 188.599 137.379 189.776L137.097 190.354L137.058 190.441C136.909 190.814 136.949 191.262 137.053 191.578C137.26 192.057 137.642 192.262 138.132 192.304L138.205 192.306L142.197 192.306L142.27 192.304C142.783 192.301 143.114 192.022 143.349 191.578C143.504 191.201 143.476 190.818 143.344 190.441L143.304 190.354L143.067 189.81C142.402 188.35 141.885 187.303 141.31 185.994L141.31 183.813L141.631 183.48L141.643 183.47L138.982 183.47ZM140.878 184.006L140.878 185.981C141.418 187.473 142.003 188.945 142.513 190.187C142.666 190.561 142.752 190.933 142.562 191.28C142.382 191.581 142.102 191.681 141.779 191.732L141.654 191.734L138.767 191.734L138.642 191.732C137.867 191.658 137.655 190.911 137.902 190.187C138.534 188.732 139.184 187.229 139.786 185.981L139.786 184.006C140.235 184.063 140.439 184.076 140.878 184.006L140.878 184.006Z" fill="#d9d9d9" stroke="none"/>
				 </g>
				</g>
			   </g>
			  </g>
			 </g>
			</g>
		   </g>
		  </g>
		 </g>
		</g>
	   </g>
	  </g>
	 </g>
	 <defs/>
	</svg>

	)

return svg
}

Laborjournal_ico(NewHandle := False) {                                                            	;-- Skript-Icon
Static hBitmap := Laborjournal_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "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"
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

LoadBar_Gui(show:=1, opt:="") {

	global load_BarGUI, load_Bar

	If !IsObject(opt)
		opt:={"col": ["0x4D4D4D","0xFFFFFF","0xEFEFEF"], "w":320,	"h":36}

	Gui, load_BarGUI: -Border -Caption +ToolWindow +AlwaysOnTop HWNDhLoad_BarWin
	Gui, load_BarGUI: Color, % opt.col.1, % opt.col.2

	load_Bar := new LoaderBar("load_BarGUI", 3, 3, opt.W, opt.H, "LABORJOURNAL", 1, opt.col.3)

	wW :=load_Bar.Width + 2*load_Bar.X
	wH :=load_Bar.Height + 2*load_Bar.Y

	Gui, load_BarGUI: Show, % "w" wW " h" wH, % "Laborjournal lädt..."
	load_Bar.hWin := hLoad_BarWin

}

LoadBar_Callback() {


}

class LoaderBar {

	__New(GUI_ID:="Default",x:=0,y:=0,w:=280,h:=28,Title:="",ShowDesc:=0,FontColorDesc:="2B2B2B",FontColor:="EFEFEF",BG:="2B2B2B|2F2F2F|323232",FG:="66A3E2|4B79AF|385D87") {
		SetWinDelay,0
		SetBatchLines,-1
		if (StrLen(A_Gui))
			_GUI_ID:=A_Gui
		else
			_GUI_ID:=1
		if ( (GUI_ID="Default") || !StrLen(GUI_ID) || GUI_ID==0 )
			GUI_ID:=_GUI_ID

		this.GUI_ID := GUI_ID
		Gui, %GUI_ID%:Default

		this.X        	:= x
		this.Y        	:= y
		this.Width 	:=w
		this.Height	:=h
		this.BG     	:= StrSplit(BG,"|")
		this.BG.W  	:= w
		this.BG.H  	:= h
		this.FG       	:= StrSplit(FG,"|")
		this.FG.W  	:= this.BG.W - 2
		this.FG.H   	:= (fg_h:=(this.BG.H - 2))
		this.Percent 	:= 0
		fg_x            	:= this.X + 1
		fg_y            	:= this.Y + 1
		this.FontColor := FontColor
		this.ShowDesc := ShowDesc

		;DescBGColor:="4D4D4D"
		DescBGColor:="Black"
		this.DescBGColor := DescBGColor
		this.FontColorDesc := FontColorDesc

		Gui,Font,s10
		Gui, Add, Text, % "x" x " y" y " w" w " h" h " BackgroundTrans 0xE hwndhLoaderBarTitle", % Title
		this.hLoaderBarTitle := hLoaderBarTitle

		Gui,Font,s8
		Gui, Add, Text, % "x" x " y+1 w" w " h" h " 0xE hwndhLoaderBarBG"
		this.ApplyGradient(this.hLoaderBarBG	:= hLoaderBarBG, this.BG.1, this.BG.2, this.BG.3,1)

		Gui, Add, Text, % "x" fg_x " y" fg_y " w0 h" fg_h " 0xE hwndhLoaderBarFG"
		this.ApplyGradient(this.hLoaderBarFG   	:= hLoaderBarFG, this.FG.1, this.FG.2, this.FG.3,1)

		Gui, Add, Text, % "x" x " y" y " w" w " h" h " 0x200 border center BackgroundTrans hwndhLoaderNumber c" FontColor, % "[ 0 % ]"
		this.hLoaderNumber := hLoaderNumber

		if (this.ShowDesc) {
			Gui, Add, Text, % "xp y+2 w" w " h16 0x200 Center border BackgroundTrans hwndhLoaderDesc c" FontColorDesc, Loading...
			this.hLoaderDesc 	:= hLoaderDesc
			this.Height        	:= h+18
		}

		Gui, Font

		Gui, %_GUI_ID%:Default
	}

	Set(p,w:="Loading...") {
		if (StrLen(A_Gui))
			_GUI_ID:=A_Gui
		else
			_GUI_ID:=1
		GUI_ID := this.GUI_ID

		Gui, %GUI_ID%:Default
		GuiControlGet, LoaderBarBG, Pos, % this.hLoaderBarBG

		this.BG.W 	:= LoaderBarBGW
		this.FG.W  	:= LoaderBarBGW - 2
		this.Percent	:= (p>=100) ? p:=100 : p

		PercentNum	:= Round(this.Percent,0)
		PercentBar	:= Floor((this.Percent/100)*(this.FG.W))

		hLoaderBarTitle		:= this.hLoaderBarTitle
		hLoaderBarFG  	:= this.hLoaderBarFG
		hLoaderNumber 	:= this.hLoaderNumber

		GuiControl,Move	,% hLoaderBarFG  	, % "w" PercentBar
		GuiControl,       	,% hLoaderNumber 	, % "[" PercentNum "% ]"

		if (this.ShowDesc) {
			hLoaderDesc := this.hLoaderDesc
			GuiControl,, % hLoaderDesc, % w
		}
		Gui, %_GUI_ID%:Default
	}

	ApplyGradient( Hwnd, LT := "101010", MB := "0000AA", RB := "00FF00", Vertical := 1 ) {
		Static STM_SETIMAGE := 0x172
		ControlGetPos,,, W, H,, ahk_id %Hwnd%
		PixelData := Vertical ? LT "|" LT "|" LT "|" MB "|" MB "|" MB "|" RB "|" RB "|" RB : LT "|" MB "|" RB "|" LT "|" MB "|" RB "|" LT "|" MB "|" RB
		hBitmap := this.CreateDIB( PixelData, 3, 3, W, H, True )
		oBitmap := DllCall( "SendMessage", "Ptr",Hwnd, "UInt",STM_SETIMAGE, "Ptr",0, "Ptr",hBitmap )
		Return hBitmap, DllCall( "DeleteObject", "Ptr",oBitmap )
	}

	CreateDIB( PixelData, W, H, ResizeW := 0, ResizeH := 0, Gradient := 1  ) {
		; http://ahkscript.org/boards/viewtopic.php?t=3203                  SKAN, CD: 01-Apr-2014 MD: 05-May-2014
		Static LR_Flag1 := 0x2008 ; LR_CREATEDIBSECTION := 0x2000 | LR_COPYDELETEORG := 8
			,  LR_Flag2 := 0x200C ; LR_CREATEDIBSECTION := 0x2000 | LR_COPYDELETEORG := 8 | LR_COPYRETURNORG := 4
			,  LR_Flag3 := 0x0008 ; LR_COPYDELETEORG := 8
		WB := Ceil( ( W * 3 ) / 2 ) * 2,  VarSetCapacity( BMBITS, WB * H + 1, 0 ),  P := &BMBITS
		Loop, Parse, PixelData, |
		P := Numput( "0x" A_LoopField, P+0, 0, "UInt" ) - ( W & 1 and Mod( A_Index * 3, W * 3 ) = 0 ? 0 : 1 )
		hBM := DllCall( "CreateBitmap", "Int",W, "Int",H, "UInt",1, "UInt",24, "Ptr",0, "Ptr" )
		hBM := DllCall( "CopyImage", "Ptr",hBM, "UInt",0, "Int",0, "Int",0, "UInt",LR_Flag1, "Ptr" )
		DllCall( "SetBitmapBits", "Ptr",hBM, "UInt",WB * H, "Ptr",&BMBITS )
		If not ( Gradient + 0 )
			hBM := DllCall( "CopyImage", "Ptr",hBM, "UInt",0, "Int",0, "Int",0, "UInt",LR_Flag3, "Ptr" )
		Return DllCall( "CopyImage", "Ptr",hBM, "Int",0, "Int",ResizeW, "Int",ResizeH, "Int",LR_Flag2, "UPtr" )
	}
}

DownloadToString(url, encoding = "utf-8") {
    static a := "AutoHotkey/" A_AhkVersion
    if (!DllCall("LoadLibrary", "str", "wininet") || !(h := DllCall("wininet\InternetOpen", "str", a, "uint", 1, "ptr", 0, "ptr", 0, "uint", 0, "ptr")))
        return 0
    c := s := 0, o := ""
    if (f := DllCall("wininet\InternetOpenUrl", "ptr", h, "str", url, "ptr", 0, "uint", 0, "uint", 0x80003000, "ptr", 0, "ptr"))    {
        while (DllCall("wininet\InternetQueryDataAvailable", "ptr", f, "uint*", s, "uint", 0, "ptr", 0) && s > 0)        {
            VarSetCapacity(b, s, 0)
            DllCall("wininet\InternetReadFile", "ptr", f, "ptr", &b, "uint", s, "uint*", r)
            o .= StrGet(&b, r >> (encoding = "utf-16" || encoding = "cp1200"), encoding)
        }
        DllCall("wininet\InternetCloseHandle", "ptr", f)
    }
    DllCall("wininet\InternetCloseHandle", "ptr", h)
    return o
}

GoogleApis() {

	;~ #Include %A_LineFile%\..\..\Plugin.ahk
	/*
		Usage: Wikipedia <Search>
		Desc: Searches wikipedia for a page and returns the first sentence plus a link
	*/

	;~ UserAgent := "GeekBot by GeekDude (Contact me on GitHub: https://github.com/G33kDude/MyRC)"
	;~ Base := "https://ajax.googleapis.com/ajax/services/search/web?v=1.0"

	;~ Google := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	;~ Google.Open("GET", Base "&q=" UriEncode("site:en.wikipedia.org " Plugin.Param), false)
	;~ Google.SetRequestHeader("User-Agent", UserAgent)
	;~ Google.Send()

	;~ if !(Result := Jxon_Load(Google.ResponseText).responseData.results[1])	{
		;~ Chat(Channel, "No results found")
		;~ ExitApp
	;~ }

	;~ Url := UriDecode(Result.Url)
	;~ SplitPath, Url, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive

	;~ if !(OutDrive ~= "^https?://en\.wikipedia\.org")	{
		;~ Chat(Channel, "Non wikipedia URL found: " OutDrive)
		;~ ExitApp
	;~ }

	;~ Base := "http://en.wikipedia.org/w/api.php?action=query&prop=extracts&exintro=&format=xml"

	;~ Wiki := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	;~ Wiki.Open("GET", Base "&titles=" UriEncode(OutFileName), false)
	;~ Wiki.SetRequestHeader("User-Agent", UserAgent)
	;~ Wiki.Send()

	;~ xml := ComObjCreate("MSXML2.DOMDocument")
	;~ xml.loadXML(Wiki.ResponseText)

	;~ html := ComObjCreate("htmlfile")
	;~ html.write(xml.selectSingleNode("//extract").text)

	;~ Page := html.body.innerText

	;~ if (Pos := InStr(Page, "."))
		;~ Page := SubStr(Page, 1, Pos)
	;~ else
		;~ Page := SubStr(Page, 1, 64)
	;~ Chat(Channel, Page " - " Shorten(Url))
	;~ ExitApp

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Includes
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
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
#Include %A_ScriptDir%\..\..\lib\class_cJSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_Neutron.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}




