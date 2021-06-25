; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 	Addendum Laborjournal
;
;      Funktion:           	- 	schnellerer Überblick über die Laborwerte der letzten Tage in Tabellenform mit farblicher Hervorhebung.
;									- 	nur relevante Werte kommen zur Anzeige
;									- 	gesondert zu behandelnde Laborparameter können angegeben werden
;										Kategorien: 	immer	:	Parameter welche unabhängig der Höhe der Normwertüber- oder unterschreitung angezeigt werden z.B. Troponin
;															exklusiv	:	werden immer angezeigt, ob auffällig oder nicht z.B. COVID-PCR o. HIV
;															nie		:	Parameter die nicht in der Tabelle erscheinen sollen
;
;
;		Hinweis:				Nichts ist ohne Fehler. Auch das hier sicher nicht. Nicht drauf verlassen! Skript soll nur eine Unterstützung sein!
;									Das Skript filtert noch keine Werte heraus die nur positiv oder negativ sein können (z.B. SARS-CoV-2 PCR)
;									INR Werte werden angezeigt, wenn ein Pat. Falithrom/Marcumar nimmt, der Wert aber nicht als therapeutischer INR gespeichert wurde.
;									Im Moment müssen Einstellungen für die Anzeige noch direkt im Skript vorgenommen werden, wie z.B. die gesondert zu behandelnden Laborparameter
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 28.04.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  ; Einstellungen
	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines, -1
	ListLines    	, Off

	global LabJ := Object()
	global adm := Object()
	global load_Bar, percent
	global PatDB

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

  ; Tray Icon erstellen
	If (hIconLabJournal := Create_Laborjournal_ico())
    	Menu, Tray, Icon, % "hIcon: " hIconLabJournal

  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	adm.Dir            	:= AddendumDir
	adm.Ini              	:= AddendumDir "\Addendum.ini"
	adm.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	adm.LabDBPath   	:= AddendumDir "\logs'n'data\_DB\LaborDaten"
	adm.AlbisDBPath 	:= AlbisPath "\DB"
	adm.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft


  ; hier alle Parameter eintragen welche gesondert verarbeitet werden sollen
						  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						  ; nie         	= werden nie angezeigt
						  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Warnen	:= {	  "nie"    	: 	"CHOL,LDL,TRIG,HDL,LDL/HD,HBA1CIFC,nicht au,AUFTRAG,Sprosspilz,Erreger"
						  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						  ; immer 	 	= nur wenn pathologisch
						  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					    	, "immer" 	: 	"NA,K,BNP,NTBNP,TropIs,TROPI,TROPT,TROP,TROPOIHS,CK,CKMB,HBK-ST,DDIM-CP,PROCAL"
						  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						  ; exklusiv 	= wenn pathologisch und bei negativen Befund (z.B. COVIA, HIV)
						  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
							, "exklusiv"	: 	"ALYMP,ALYMPDIFF,ALYMPH,ALYMPNEO,ALYMPREA,ALYMPRAM,KERNS,METAM,MYELOC,PROMY,HLAMBDA," 		; Leukozytenveränderungen
											.	"DIFANISO,DIFPOLYC,DIFPOIKL,DIFHYPOC,DIFMIKRO,DIFMAKRO,DIFOVALO," 					                             	; Erythrozytenveränderungen
											.	"COVIPC-A,COVIP-SP,COVIGAB,COVIA,COVIG,COVMU484,COVMU501,COVM6970,HIV,"									; Infektionen
											.	"KAP,KAP-U,LAM,LAM-U,"																																		; Tumor
											.	"I"                        ; I = ikterisch                                                                                                                            	; sonstige
						  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						  ; Gruppierungen
						  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
							, "Gruppe"   	: {"atyp. Leuko"	: "ALYMP,ALYMPDIFF,ALYMPH,ALYMPNEO,ALYMPREA,ALYMPRAM,KERNS,METAM,MYELOC,PROMY,HLAMBDA"
												,	"atyp. Ery" 	: "DIFANISO,DIFPOLYC,DIFPOIKL,DIFHYPOC,DIFMIKRO,DIFMAKRO,DIFOVALO"
												,	"COVID-19"	: "COVIPC-A,COVIP-SP,COVIGAB,COVIA,COVIG,COVMU484,COVMU501,COVM6970"}}

  ; load_Bar anzeigen
	LoadBar_Gui()

 ; Patientendaten laden
	load_Bar.Set(0, "lade Patientendatenbank")
	PatDB 	:= ReadPatientDBF(adm.AlbisDBPath,, ["NR", "NAME", "VORNAME", "GEBURT"])
	load_Bar.Set(1, "Patientendatenbank geladen")

  ; Laborjournal anzeigen
	LabPat     := AlbisLaborJournal("", "", Warnen, 140, true)
	LaborJournal(LabPat, true)

return

~ESC:: ;{
	If WinActive("Laborjournal") {
		SaveGuiPos(labJ.hwnd)
		ExitApp
	}
return ;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Berechnungen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
AlbisLaborJournal(Von="", Bis="", Warnen="", GWA=100, Anzeige=true) {               	;--  sehr abweichender Laborwerte

	/*  Parameter

		Von       	-	Anfangsdatum, leerer String oder Zahl
							eine Zahl als Wert wird als Anzahl der Werktage die rückwärtig vom aktuellem Tagesdatum angezeigt werden soll gewertet
		Bis         	-	Enddatum oder leer

	*/

		static Lab

	; Variablen                                                                                                         	;{
		LabPat := Object()
		nieWarnen    	:= "i)\b(" StrReplace(Warnen.nie      	, ","	, "|") ")\b"
		immerWarnen 	:= "i)\b(" StrReplace(Warnen.immer 	, ","	, "|") ")\b"
		exklusivWarnen	:= "i)\b(" StrReplace(Warnen.exklusiv	, ","	, "|") ")\b"
		skipErgebnis		:= "i)(Siehe un|Material)"
	;}

	; durchschnittliche Abweichung von den Normwertgrenzen laden oder berechnen 	;{
		If !IsObject(Lab) {
			If !FileExist(adm.LabDBPath "\Laborwertgrenzen.json") {
				load_Bar.Set(1, "ermittle Laborwertgrenzen ...")
				Lab := AlbisLaborwertGrenzen(adm.LabDBPath "\Laborwertgrenzen.json", Anzeige)
			} else {
				load_Bar.Set(1, "lade Laborwertgrenzen ...")
				Lab := JSONData.Load(adm.LabDBPath "\Laborwertgrenzen.json", "", "UTF-8")
			}
		}
		load_Bar.Set(5, "Laborwertgrenzen geladen")
	;}

	; Laborparameter lange Bezeichnungen laden                                                     	;{
		load_Bar.Set(6, "lade Bezeichnungen der Laborparameter")
		LabDic := AlbisLabParam("load","short")
		load_Bar.Set(7, "Bezeichnungen der Laborparameter geladen")
	;}

	; Suchzeitraum                                                                                                  	;{

		load_Bar.Set(7, "berechne den Suchzeitraum")

		; Von enthält kein Datum
			RegExMatch(Von, "^\d+$", Tage)
			Von 	:= Tage ? "" : Von
			Bis 	:= Tage ? "" : Bis

		; die Anzeige wird für x Werktage berechnet
			If !Von && !Bis {

				Tage	    	:= Tage ? Tage : 14                            	; kein Zeitraum, Defaulteinstellung sind 14 Tage
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
				Von      	+= -1*(DaysPlus), Days
				Von 	    	:= SubStr(Von, 1, 8)

				ViewStart 	:= FormatDate(Von, "YMD", "dd.MM.yyyy")
				WTag    	:= GetWeekDay(ViewStart)
				LabJ.Tagesanzeige :=  "ab " WTag ", " StrReplace(ViewStart, A_YYYY) " (" Tage " Werktage)"

				FormatTime, Von, % Von, yyyyMMdd
				QJ	:= A_YYYY . Ceil(A_MM/3)
				VB	:= "Von"

			}
			else
				VB := (Von && !Bis) ? "Von" : (!Von && Bis) ? "Bis" : "VonBis"

		load_Bar.Set(8, "Suchzeitraum berechnet")
	;}

	; Albis: Datenbank laden und entsprechend des Datums die Leseposition vorrücken	;{
		labDB := new DBASE(adm.AlbisDBPath "\LABBLATT.dbf", 0)
		labDB.OpenDBF()
	;}

	; Datei: Startposition berechnen                                                                            	;{
		If (VB = "Von") || (VB = "VonBis") {
			If !QJ
				QJ := SubStr(Von, 1, 4) . Ceil(SubStr(Von, 5, 2)/3)

			SciTEOutput(QJ)

			If !lab.seek.HasKey(QJ) {
				; ein Quartal davor
				nYYYY 	:= SubStr(QJ, 1, 4)
				nQ      	:= SubStr(QJ, 5, 1)
				nQ    	:= nQ-1=0 ? 4 : nQ-1
				nYY    	:= nQ=4 ? nYYYY-1 : nYYYY
				QJ     	:= nYYYY nQ
			}

			SciTEOutput(QJ)

			startrecord := Lab.seeks[QJ]
			res := labDB._SeekToRecord(startrecord, 1)
			;~ SciTEOutput(QJ ": " startrecord ", " res)
		}
	;}

		t :=  nieWarnen "`n" exklusivWarnen "`n" immerWarnen "`n`n"
		ShowAt := Floor(labDB.records/80)
		load_Bar.Set(8, "lade neue Laborwerte")

	; filtert nach Datum und ab einer Überschreitung von der durchschnittlichen Abweichung von den Grenzwerten anhand
	; der zuvor für jeden Laborwert berechneten durchschnittlichen prozentualen Überschreitung des Grenzwertes aus den
	; Daten der LABBLATT.dbf [AlbisLaborwertGrenzen()]
		while !(labDB.dbf.AtEOF) {

			; Daten zeilenweise auslesen
				data := labDB.ReadRecord(["PATNR", "ANFNR", "DATUM", "BEFUND", "PARAM", "ERGEBNIS", "EINHEIT", "GI", "NORMWERT"])

			; Fortschritt
				If Anzeige && (Mod(A_Index, ShowAt) = 0) {
					tt := "lade Laborwerte  [" ConvertDBASEDate(data.Datum) "]  " SubStr("00000000" labDB.recordnr, -5) "/" SubStr("00000000" labDB.records, -5)
					percent := Floor(((labDB.recordnr*100)/labDB.records)*0.8)
					load_Bar.Set(8+percent, tt)
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

			; kein pathologischer Wert und in keiner Filterliste dann weiter ;{
				PNImmer := PNExklusiv := false
				If RegExMatch((PN := data.PARAM ), nieWarnen)  	; Parameter nieWarnen
					continue
				else if RegExMatch(PN, exklusivWarnen)               	; Parameter exklusivWarnen
					PNExklusiv := true
				else if RegExMatch(PN, immerWarnen)                  	; Parameter immerWarnen
					If RegExMatch(data.GI, "(\+|\-)")                      	; und Wert ist pathologisch
						PNImmer 	:= true
					else
						continue
			;}

			; Strings lesbarer machen (oder auch nicht)                        	;{

				; Laborergebnis oder Wert
					PW    	:= StrReplace(data.ERGEBNIS	, ",", ".")        	; Parameter - Wert
					If RegExMatch(PW, skipErgebnis)
						continue

				; NORMWERTGRENZE - AUSWERTEN
				; PV kann > oder < sein, PSU - unterer Grenzwert, PSO - oberer Grenzwert
				; 	Beispiele aus der LABBLATT dBASE Datei Spalte Normwerte
				; 		1. Beispiel: "0,02 - 0,08"	 --> PSU="0,02" u. PSO="0,08"
				; 		2. Beispiel: "< 0,01"  		 --> PV="<" u. PSO="0,01"
				; 		3. Beispiel: "> 1,04"	    	 --> PV=">" u. PSU=0,01
					data.NORMWERT := RegExReplace(data.NORMWERT, "[a-z]+$")  ; entfernt nicht auswertbare Buchstaben
					RegExMatch(data.NORMWERT, "(?<V>[\<\>])*(?<SU>[\d,\.]+)[\-\s]*(?<SO>[\d,\.]+)*", P)   ; Normwert-String mit RegEx aufsplitten
					PSU  	:= StrReplace(PSU	, ",", ".")                            	; unterer Grenzwert
					PSO  	:= StrReplace(PSO, ",", ".")                            	; oberer Grenzwert
					PSO  	:= !PSO ? PSU : PSO                                  	; wenn kein oberer Grenzwert vorhanden ist, dann ist es ein Cut-Wert
					If (PV = "<")
						PSO := PSU ? PSU : PSO
					else if (PV = ">")
						PSU := PSO ? PSO : PSU

					OD    	:= lab.params[PN].OD                              	; durchschnittliche Abweichung von der oberen Normwertgrenze (oberer Durchschnitt in Prozent)
					UD	  	:= lab.params[PN].UD                                	; durchschnittliche Abweichung von der unteren Normwertgrenze (unterer Durchschnitt in Prozent)
					PE	    	:= data.EINHEIT				                        	; Parameter - Einheit
					ONG	:= UNG := 0                                            	; Obere/Untere Norm-Grenze
			;}

			; außer bei immer anzuzeigenden Parametern ausführen
				If !PNExklusiv {

					If (PW > PSO)                                                    	; Laborwert überschreitet die obere Grenze
						ONG := Floor(Round((PW*100)/PSO, 3))        	; Oberere Normgrenze in Prozent
					else if (PW < PSU)                                               	; Laborwert ist kleinerals der untere Grenzwert
						UNG := Floor(Round((PSU*100)/PW, 3))	    	; Untere Normgrenze in Prozent

					If !ONG && !UNG                                              	; ONG und UNG sind beide null -> weiter
						continue

					If (ONG > 0)                                                     	; ONG überschritten
						plus := Floor(Round((ONG*100)/OD,1))           	; Überschreitung der ONG in Prozent
					else if (UNG > 0 )                                              	; UNG überschritten
						plus := Floor(Round((UNG*100)/UD,1))           	; Unterschreitung der UNG in Prozent

					If PNImmer
						If (plus <= 100)
							PNImmer := false


				}

					;~ If plus < 102
						;~ t .= "D:" data["DATUM"] " | P:" data["PARAM"] " | PW:" PW " | NW:"  data["NORMWERT"]
							;~ . 	" | PSU: " PSU " | PSO:" PSO " | UNG:" UNG " | ONG:" ONG " | UD:" UD " | OD:" OD " => plus:" plus "`n"

			; Laborwert liegt über der Grenzwert-Abweichung oder gehört zu einer Filterliste
				If (plus > GWA || PNImmer || PNExklusiv) {

						PatID 	:= data.PATNR
						Datum 	:= data.Datum
						PW    	:= RegExReplace(PW, "(\.[0-9]+[1-9])0+$", "$1")
						AN    	:= LabPat[Datum][PatID][PN].AN
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
						PD	:= ONG > 0 ? "+" Round(ONG - labOD)	: (UNG > 0 ? "-" Round(UNG - labUD)	: "")
						LabPat[Datum][PatID][PN].PatID	:= PatID                                                                                                                  	; Patienten-ID
						LabPat[Datum][PatID][PN].CV  	:= PNImmer ? 1 : PNExklusiv ? 2 : 0                                                                         	; CAVE-Wert für andere Textfarbe im Journal
						LabPat[Datum][PatID][PN].PN  	:= PN                                                                                                                    	; Parameter Name
						LabPat[Datum][PatID][PN].PW  	:= RegExReplace(PW, "\.0+$")                                                                                 	; Parameter Wert
						LabPat[Datum][PatID][PN].PV		:= P ? StrReplace(P, ",", ".") : data["NORMWERT"]                                                      	; Parameter Varianz (Normwertgrenzen)
						LabPat[Datum][PatID][PN].PL 		:= PL                                                                                                                    	; proz. Abweichung von der Normgrenze (oberhalb sowie unterhalb)
						LabPat[Datum][PatID][PN].PA		:= ONG > 0 ? OD : (UNG > 0 ? UD : "" )                                                                	; Parameter Abweichung +/-
						LabPat[Datum][PatID][PN].PD		:= PD                                                                                                                     	;
						LabPat[Datum][PatID][PN].PE		:= PE                                                                                                    	            	; Parameter Einheit
						LabPat[Datum][PatID][PN].AN  	:= LTrim(AN, ",")                                                                                                     	; Labor Anforderungsnummer
						LabPat[Datum][PatID][PN].PB   	:= LabDic[PN].1                                                                                                       	; Parameter Bezeichnung: ausgeschriebene Parameterbezeichnung

					}

		}

	; Lesezugriff beenden
		labDB.CloseDBF()
		LabDB := ""
		If Anzeige
			ToolTip
		percent := percent + 8

		;~ FileOpen(A_Temp "\labdata.txt", "w", "UTF-8").Write(t)
		;~ Run, % A_Temp "\labdata.txt"

return LabPat
}

AlbisLaborwertGrenzen(LbrFilePath, Anzeige=true) {                                               	;-- Berechnung der durchschnittlichen Abweichung v. Grenzwert

	/* 	Beschreibung - AlbisLaborwertGrenzen()

		⚬	Berechnet aus den Daten der LABBLATT.dbf die durchschnittliche prozentuale Über- oder Unterschreitung eines Laborparameters.
		⚬	für eventuelle Anpassungen wird die maximalste Über- oder Unterschreitung als Einzelwert gespeichert
		⚬	Durch Nutzung eines Faktors (Prozent) sind die dadurch mit Annährung erreichte "Warngrenze" auch bei unterschiedlichen
			Einheiten praktikabel, außerdem entfällt das Anpassen an die sich wechselnden Normbereiche.

		⊛ Die Aufgabe war (ist es noch) einen Algorithmus zu finden, welcher mit einer guten Zuverlässigkeit klinisch bedeutsame
			Laborwerte herausfiltern kann. Möglichst wenige Informationen sollten angezeigt werden. Die Relevanz von path. Laborwerten ist manchmal ganz
			deutlich am Grenzwert festzumachen. Als bestes Beispiel sind hier wären hier Elektrolytwert-Veränderungen zu nennen. Bei anderen Parametern ist
			dies aber nicht so. Warum wirft man als langjährig tätiger Arzt manchmal nur einen Blick die Parameter und weiß sofort die Relevanz zu ermessen?
			Mit diesem Programm versuche ich mich unter anderem dieser Fragestellung zu nähern.
			So sind im Laufe meiner Tätigkeit als niedergelassener Arzt circa 300 verschiedene Laborparameter interessant geworden.
			Bei der Menge an Parametern ist eine regelmäßige Wichtung schon aufgrund regelmäßig vom Laboranbieter angepasster Normbereiche unsinnig.
			Die Idee war also nur die Über- und Unterschreitung der Grenzwerte zu wichten. Bei der Berechnung werden alle im Normbereich liegenden
			Werte vom Algorithmus ignoriert. Über- und Unterschreitungen werden als relative Abweichung erfasst.
			Der Einfachheit halber in Prozent (200 % oder das 2fache der Norm z.B.). Für die Abweichungen in beide Richtungen werden für jeden Parameter
			zwei Werte gebildet. Erfasst wird am Ende die maximale prozentuale Abweichung und aus den Laborwerten aller Patienten wird für jeden Parameter
			der Durchschnitt berechnet. Ich nenne es die "durchschnittliche Abweichung" vom oberen oder unteren Grenzwert.
			Die erreichten Alarmwerte sind in vielen Fällen gut und recht nah dran an dem was ich sehen wollte.
			Bei einigen Parameter sehe ich noch Einschränkungen. Fehlerhaft ist im Moment noch die Erkennung von Unterschreitungen.

	 */

	; Variablen
		Lab  	:= Object()
		seeks:= Object()

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
					ToolTip, % "Laborwertgrenzen werden berechnet.`ngefundene Parameter: " lab.Count() "`nDatensatz: " SubStr("00000000" A_Index, -5) "/" SubStr("00000000" labDB.records, -5) ;"`n" lfound

				If !RegExMatch(data["GI"], "(\+|\-)") || StrLen(data["NORMWERT"]) = 0
					continue

				PN	:= data.PARAM
				PE		:= data.EINHEIT
				PW	:= StrReplace(data["ERGEBNIS"]	, ",", ".")                                                                         ; PARAMETER WERT
				RegExMatch(data["NORMWERT"], "(?<V>[\<\>])*(?<SU>[\d,\.]+)[\-\s]*(?<SO>[\d,\.]+)*", P)
				PSU	:= StrReplace(PSU	, ",", ".")
				PSO	:= StrReplace(PSO, ",", ".")


			; berechnet die Grenzwertüberschreitungen
				ONG :=UNG := 0
				PSO := StrLen(PSO) > 0 ? PSO : PSU
				If (PW >= PSO)
					ONG := Floor(Round((PW * 100)/PSO, 3))
				else If (PW <= PSU)
					UNG := Floor(Round((PSU * 100)/PW, 3))

				If (PN = "NA" && PW < 130 ) {
					dcount ++
					SciTEOutput( "D: " data["DATUM"] " | P: " data["PARAM"] " | E: " PW " | N: "  data["NORMWERT"]
									. 	" | U: " PSU " | O: " PSO " | UG: " UNG " | OG: " ONG )
				}

				If !ONG && !UNG
					continue

			; Ersterstellung eines Laborwert Datensatzes
				If !Lab.HasKey(PN) {

					Lab[PN] := {	"PZ" 	: 1			; Parameterzähler
									, 	"O" 	: 0			; Summe d. prozentualen Abweichung vom oberen Grenzwert
									, 	"OI"	: 0			; Abweichungszähler (obere Grenzwerte)
									, 	"OD"	: 0			; durchschnittliche prozentuale Abweichung vom oberen Grenzwert (O/OI)
									, 	"OM": 0    		; maximale prozentuale Abweichung an der Obergrenze
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
		JSONData.Save(LbrFilePath "\Laborwertgrenzen.json", Labwrite, true,, 1, "UTF-8")
		JSONData.Save(LbrFilePath "\Laborwertbasis.json", Lab, true,, 1, "UTF-8")

		ToolTip

return Labwrite
}

AlbisLabParam(cmd:="load", data:="short", Anzeige:=false) {                                  	;-- Laborparameter als Excel (.csv) Datei speichern

	; Albis: Inhalt der LABPARAM.dbf komplett laden
		LabDB := new DBASE(adm.AlbisDBPath "\LABPARAM.dbf", Anzeige)
		LabDB.OpenDBF()
		LabParam := labDB.GetFields("")
		LabDB.OpenDBF()

		If (data = "full") {

			For idx, line in LabParam {
				For key, val in Line
					h .= key "`t"
				break
			}
			h := RTrim(h, "`t") "`n"
			For idx, line in LabParam {
				For key, val in Line
					t .= val "`t"
				t := RTrim(t, "`t") "`n"
			}

			FileOpen(adm.LabDBPath "\LabParam.csv", "w", "UTF-8").Write(h . t)
			JSONData.Save(adm.LabDBPath "\LabParam.json", LabParam, true,, 1, "UTF-8")
			return LabParam

		}
		else {

			; es wurde noch kein Wörterbuch angelegt oder es soll neu erstellt werden
			; Wörterbuch besteht nur aus der Abkürzung (key) und der ausgeschriebenen Parameterbezeichnung (value)
				If !FileExist(adm.LabDBPath "\LabDictionary.json") || (cmd = "rewrite") {

					LabDic	:= Object()
					For idx, obj in LabParam {

						If !LabDic.HasKey(obj.NAME) && (StrLen(obj.BEZEICH) > 0)
							LabDic[obj.NAME] := [obj.BEZEICH]
						else if LabDic.HasKey(obj.NAME) && (StrLen(obj.BEZEICH) > 0)
							LabDic[obj.NAME].Push(obj.BEZEICH)

					}
					JSONData.Save(adm.LabDBPath "\LabDictionary.json", LabDic, true,, 1, "UTF-8")

				}
			; Wörterbuch laden
				else (cmd = "load")
					LabDic := JSONData.Load(adm.LabDBPath "\LabDictionary.json", "", "UTF-8")

			return LabDic
		}

}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; grafische Anzeige
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
LaborJournal(LabPat, Anzeige=true) {

	; letzte Fensterposition laden
		IniRead, winpos, % adm.ini, % adm.compname, LaborJournal_Position
		If (InStr(winpos, "ERROR") || StrLen(winpos) = 0)
			winpos := "w920 h500"

	; Variablen bestücken                                	;{
		srchdrecords  	:= LTrim(labJ.srchdrecords, "0")
		dbrecords     	:= labJ.records
		Tagesanzeige	:= LabJ.Tagesanzeige
		iconpath        	:= adm.Dir "\assets\ModulIcons\LabJournal.svg"
	;}

	; HTML Vorbereitungen 	                              	;{
		static TDL      	:= "<TD style='text-align:Left'>"                                                                                                   	; Textausrichtung links
		static TDL1     	:= "<TD style='text-align:Left; border-left:0px solid'>"                                                                   	; Textausrichtung links kein Rand links
		static TDLR1     	:= "<TD style='text-align:Left; border-right:0px solid; border-left:0px solid'>"                              	; Textausrichtung links, kein Rand rechts
		static TDR      	:= "<TD style='text-align:Right' border-right:0px solid border-left:1px solid'>"                                	; Textausrichtung rechts
		static TDR1     	:= "<TD style='text-align:Right; border-right:0px solid'>"                                                            	; Textausrichtung rechts, kein Rand rechts
		static TDR2a     	:= "<TD style='text-align:Right; border-right:1px'>"
		static TDR2b     	:= "<TD style='text-align:Right; border-right:1px; border-top:0px;border-bottom:0px'>"                	; Tooltip
		static TDR3      	:= "<TD class='tooltip' data-tooltip='##' style='text-align:Right'>"                                              	; Param
		static TDC     	:= "<TD style='text-align:Right' border-right:1px>"                                                                     	;
		static TDJ      	:= "<TD style='text-align:Justify' border-right:1px>"

		static cCol     	:= ["color: Red", "color: BlueViolet", "color: DarkTeal"]
		static FWeight	:= ["font-weight: bold", "font-weight: bold", "font-weight: bold; font-family: sans-serif"]

		If !IsObject(cColW) {
			cColW := []
			For i, htmlcolor in cCol
				cColW[i] := ";" cCol[i] ";" FWeight[i]
		}


	; HTML Seitendaten
		htmlheader := FileOpen(A_ScriptDir "\assets\Laborjournal.html", "r", "UTF-8").Read()
		htmlheader := StrReplace(htmlheader, "##01##", iconpath)
		htmlheader := StrReplace(htmlheader, "##02##", Tagesanzeige)
		htmlbody =
		(

			<div style='overflow-x:scroll; width:100`%; height:100`%;'>
			<table id='LabJournal_Table'>

				<tr>
					<th style='text-align:Right'>Datum</th>
					<th style='text-align:Center' colspan='3'>NR, Name, Vorname, Geburtsdatum</th>
					<th style='text-align:Left'>Anford.Nr</th>
					<th style='text-align:Right'>Param</th>
					<th style='text-align:Right'>Wert</th>
					<th style='text-align:Center' colspan='2'>Normalwerte</th>
					<th style='text-align:Right'>+/-</th>
					<th style='text-align:Right'>⍉ Abw.</th>
				</tr>

			<div style='overflow-y: scroll;'>
		)

		html := RegExReplace(htmlheader . htmlbody, "^\s{16}(.*[\n\r]+)", "$1")

		load_Bar.Set(percent+2, "erstelle die Journalanzeige")
	;}

	; HTML Tabelle wird erstellt                         	;{

		trIDf      	:= false                    	; alternierende Farbanzeige
		sortDatum	:= Array()
		For Datum, Patients in LabPat
			sortDatum.InsertAt(1, Datum)

	; ――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
	; Journalausgabe wird mit absteigendem Datum erstellt
	; ――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
		For idx, Datum in sortDatum {

				If (DatumLast <> Datum)
					UDatum := ConvertDBASEDate(Datum)

			; ―――――――――――――――――――――――――――――――――――――――――――――――――――――
			; Ausgabe der Werte nach Patient
			; ―――――――――――――――――――――――――――――――――――――――――――――――――――――
				For PatID, parameter in LabPat[Datum] {

					If (PatIDLast <> PatID) || (DatumLast <> Datum) {                            	; der Patient wird nur einmal pro Tag angezeigt
						Patient  	:= PatDB[PatID].NAME ", " PatDB[PatID].VORNAME
						PatGeburt	:= ConvertDBASEDate(PatDB[PatID].GEBURT)
						ANFNR  	:= RTrim(parameter["ANFNR"], ", ")
						PatIDLast	:= PatID
						trIDf      	:= !trIDf                                                                        	; Pat. hat sich geändert, alternierende Farbänderung
					}

				; ―――――――――――――――――――――――――――――――――――――――――――――――――――
				; Parameter Objekt (PN) - Erstellung der Datenzeilen
				; ―――――――――――――――――――――――――――――――――――――――――――――――――――
 					For key, PN in parameter {

					  ; ist PN leer, einfach ignorieren
						If !IsObject(PN)
							continue

					  ;  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧
					  ; Tabelle  ‧  ‧  ‧ 	einzelne Zellen vorbereiten (HTML CSS Styles den Tabellenfelder zuordnen)
					  ;  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧  ‧

						; temporäre Variablen
						TDPATBUTTONS 	:= (Patient  	? (trIdf ? " class='table-btn1'" : " class='table-btn2'") " onclick='ahk.LabJournal_KarteiKarte(event)'" : "")
						TDR4	            	:= StrReplace(TDR3, "##", PN.PB)         ; Laborparameterbeschreibung
						html_ID              	:= PatID "_" Patient

					  ; Datum                                                     Farbe1        Farbe2
						TRDAT    	:= "<TR" (UDatum 	? (trIdf ? " id='td1a'" : " id='td2a'") : (trIdf ? " id='td1b'" : " id='td2b'") ) ">"
						TDRDX		:= UDatum ? TDR2a : TDR2b
					  ; ID
						TDPATNR 	:= RTrim(TDLR1, ">")  	TDPATBUTTONS "; id='" html_ID "'>"
					  ; Patientenname
						TDPAT     	:= RTrim(TDLR1, ">") 	TDPATBUTTONS "; id='" html_ID "'>"
					  ; Parameter Name                             PN.CV = CAVE
						TDR1X     	:= PN.CV = 0 ? TDR4 	: (PN.CV = 1 ? StrReplace(TDR4, "'>", cColW.1 "'>") : StrReplace(TDR4 	, "'>", cColW.2 "'>"))
					  ; Parameter Wert
						TDR2X     	:= PN.CV = 0 ? TDR 	: (PN.CV = 1 ? StrReplace(TDR 	, "'>"	, cColW.1 "'>") : StrReplace(TDR  	, "'>", cColW.2 "'>"))
					  ; Parameter Normwerte
						TDR3X     	:= PN.CV = 0 ? TDR1	: (PN.CV = 1 ? StrReplace(TDR1, "'>"	, cColW.1 "'>") : StrReplace(TDR1	, "'>", cColW.2 "'>"))
					  ; Parameter Einheit
						TDL1X     	:= PN.CV = 0 ? TDL1	: (PN.CV = 1 ? StrReplace(TDL1	, "'>"	, cColW.1 "'>") : StrReplace(TDL1	, "'>", cColW.2 "'>"))
					  ; Parameter Lage (Hinweiszeichen)
						TDCX      	:= PN.CV = 0 ? TDC 	: (PN.CV = 1 ? StrReplace(TDC	, "'>"	, cColW.1 "'>") : StrReplace(TDC 	, "'>", cColW.2 "'>"))

					; Abweichung oberhalb der Norm und PN.CV (CAVE) negativ (CV - wenn wahr dann andere Farbkennzeichnung)
						If (PN.PL > 250 && !PN.CV) {
							TDRX      	:= StrReplace(TDR 	, ">"	, " " cColW.3 "'>")
							TDR1X     	:= StrReplace(TDR1	, ">"	, " " cColW.3 "'>")
							TDCX      	:= StrReplace(TDC	, ">"	, " " cColW.3 "'>")
							TDL1X     	:= StrReplace(TDL1	, ">"	, " " cColW.3 "'>")
						}

						If PN.PA {
							RegExMatch(PN.PV, "(?<V>[\<\>])*(?<SU>[\d.]+)\-*(?<SO>[\d.]+)*", P)
							PSM  	:= PSO ? PSO : PSU
							PA     	:= RegExReplace(Round(PN.PA * (PSM/100), 1), "\.0+$")
						}

						PE         	:= Trim(PN.PE)	? PN.PE : "" ;" - - - - -"
						tbDatum	:= UDatum ? SubStr(UDatum, 1, 6) : ""

						html .= "`t" TRDAT "`n"
								.  	"`t`t "	TDRDX  	. 	tbDatum                             	"</TD>`n"	; Abnahmedatum                	[Datum]
								. 	"`t`t "	TDPATNR	.	PatID             	                  	"</TD>`n"	; Patientennummer                [Patient]
								. 	"`t`t "	TDPAT     	.	Patient             	                  	"</TD>`n"	; Patientenname                    [Patient]
								. 	"`t`t "	TDPAT     	.	PatGeburt        	                  	"</TD>`n"	; Patientengeburtstag           	[Patient]
								. 	"`t`t "	TDPAT     	.	ANFNR          	                  	"</TD>`n"	; Anforderungsnummer       	[Patient]
								. 	"`t`t "	TDR1X     	.	PN.PN                                 	"</TD>`n"	; Parameter Name              	[Param]
								. 	"`t`t "	TDR2X    	.	PN.PW                               	"</TD>`n"	; Parameter Wert               	[Wert]
								.	"`t`t "	TDR3X    	.	PN.PV			                     	"</TD>`n"	; Parameter Normwerte        	[Normalwerte]
								.	"`t`t "	TDL1X     	.	PE                                       	"</TD>`n"	; Parameter Einheit                [      ]
								. 	"`t`t "	TDCX    	.	(PN.PL	? PN.PL "x" 	: "")     	"</TD>`n"	; Parameter Lage                   [+/-]
								. 	"`t`t "	TDR       	.	(PA   	? PA " " PE  	: "") 		"</TD>`n"	; Parameter Abweichung     	[⍉ Abw.]
								. 	"`t</TR>`n`n"

						PE := PA := UDatum := PatID := Patient := PatGeburt := ANFNR := ""

					}

				}

				DatumLast := Datum

		}

		html .= "</div></table></div></body></html>"
		;html .= "</tbody></table></div></body></html>"
		;}

	; erstellte HTML Seite wird angezeigt         		;{
		FileOpen(A_Temp "\Laborjournal.html", "w", "UTF-8").Write(html)
		neutron := new NeutronWindow("","","","Laborjournal" LabJ.maxrecords ")", "+AlwaysOnTop minSize920x400")
		neutron.Load(A_Temp "\Laborjournal.html")
		neutron.Gui("+LabelNeutron")
		neutron.Show(winpos)
		hLJ := WinExist("A")
		;WinSet, AlwaysOnTop,, % "ahk_id " hLJ

		obj := neutron.wb.document.getElementById("LaborJournal_Header")	, hcr	:= obj.getBoundingClientRect()
		obj := neutron.wb.document.getElementById("LabJournal_Table")    	, tcr	:= obj.getBoundingClientRect()

		labJ.hwnd := hLJ
		labJ.HeaderHeight	:= Floor(hcr.Bottom + 1)
		labJ.TableHeight   	:= Floor(tcr.Bottom  + 1)
		labJ.enrolled	    		:= true

		npos := PosStringToObject(winpos)
		If !IsInsideVisibleArea(npos.X, npos.Y, npos.W, npos.H)
			winpos := "xCenter yCenter w" npos.W " h" npos.H

		neutron.Show(winpos)

		load_Bar.Set(100, "Das Laborjounal ist geladen!")
		Gui, load_BarGUI: Destroy

	;}

	; speichern als BMP zum Versenden als Bilddatei
		;hBMP := hWnd_to_hBmp(labJ.hwnd)
		;SavePicture(hBMP, A_Temp "\Laborjournal.jpg")
		;Run % A_Temp "\Laborjournal.bmp"

}

LabJournal_ParamToolTip(neutron, event) {

	ToolTip, % event

}

LabJournal_KarteiKarte(neutron, event) {

	; event.target will contain the HTML Element that fired the event.
	; Show a message box with its inner text.
		RegExMatch(event.target.id, "^\s*(?<ID>\d+)?_(?<Name>.*)", thisPat)
		;~ SciTEOutput("ID: " thisPatID ", text: " thisPatName ", " event.target.id)

		If labJ.enrolled ="x" {

			lj  	:= GetWindowSpot(labJ.hwnd)
			dh	:= lj.h
			dtb	:= labJ.TableHeight - labJ.HeaderHeight + 1
			dStep:= Floor(dtb/25)
			Loop 25 {
				dh -= dStep
				SetWindowPos(labJ.hwnd, lj.x, lj.y, lj.w, dh)
			}

			labJ.enrolled := false

		}

		If AlbisAkteOeffnen(thisPatName, thisPatID)
			AlbisLaborBlattZeigen()
		else
			PraxTT("Die Patientenakte konnte nicht geöffnet werden!")

}

LabJournal_DragTitleBar(event) {

	If labJ.enrolled {
		PostMessage, 0xA1, 2, 0,, % "ahk_id" labJ.hwnd  ; WM_NCLBUTTONDOWN := 0xA1
	}
	else If labj.enrolled ="x" {

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

LabJournal_Grenzen(event) {
AlbisLaborwertGrenzen(adm.LabDBPath)
}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
SaveGuiPos(hwnd) {

	win 			:= GetWindowSpot(hwnd)
	winPos	 	:= "x" win.X " y" win.Y " w" win.CW " h" win.CH
	IniWrite, % winpos	, % adm.Ini, % adm.compname, LaborJournal_Position

}

PosStringToObject(string) {

	p := Object()
	For wIdx, coord in StrSplit("XYWH") {
		RegExMatch(string, "i)" coord "(?<Pos>\d+)", w)
		p[coord] := wPos
	}

return p
}

MessageWorker(InComing) {                                                                                    	;-- verarbeitet die eingegangen Nachrichten

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

hWnd_to_hBmp( hWnd:=-1, Client:=0, A:="", C:="" ) {                                          	;-- Capture fullscreen, Window, Control or user defined area of these

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

SavePicture(hBM, sFile) {                                                                                        	;-- By SKAN on D293 @ bit.ly/2krOIc9
Local V,  pBM := VarSetCapacity(V,16,0)>>8,  Ext := LTrim(SubStr(sFile,-3),"."),  E := [0,0,0,0]
Local Enc := 0x557CF400 | Round({"bmp":0, "jpg":1,"jpeg":1,"gif":2,"tif":5,"tiff":5,"png":6}[Ext])
  E[1] := DllCall("gdi32\GetObjectType", "Ptr",hBM ) <> 7
  E[2] := E[1] ? 0 : DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr",hBM, "UInt",0, "PtrP",pBM)
  NumPut(0x2EF31EF8,NumPut(0x0000739A,NumPut(0x11D31A04,NumPut(Enc+0,V,"UInt"),"UInt"),"UInt"),"UInt")
  E[3] := pBM ? DllCall("gdiplus\GdipSaveImageToFile", "Ptr",pBM, "WStr",sFile, "Ptr",&V, "UInt",0) : 1
  E[4] := pBM ? DllCall("gdiplus\GdipDisposeImage", "Ptr",pBM) : 1
Return E[1] ? 0 : E[2] ? -1 : E[3] ? -2 : E[4] ? -3 : 1
}

Create_Laborjournal_ico(NewHandle := False) {                                                    	;-- ICON anzeigen
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 :=	"AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAABwAAAAcAAAAAAAAAAAAAAAAHwAAHwAAHwAAHwAAHwAFIQNUSi+kc1zYjXj0nIj/oo//oo//oo//oo//oo//oo//oo//oo//oo//"
		. 	"oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/znIjYjXijc1tTSi8FIQMAHwAAHwAAHwAAHwAAHwAAHwAAHwAAHwAAHwBRSC3ckHv/oo/ai3isbVqWXk2OWUeOWUeOWUeOWUeOW"
		.	"UeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeWXk2tbFvbi3n/oo/bj3pOSCwAHwAAHwAAHwAAHwAA"
		. 	"HwAAHwACIAGKZk3/oo/HfmtjPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd"
		. 	"CKBdkPi3If2z/oo+JZU0CIAEAHwAAHwAAHwAAHwCLZk78oI2VXUxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKB"
		.	"dCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeWXk38oI2JZU0AHwAAHwAAHwBSSS7/oo+VXUxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd"
		.	"CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeXX03/oo9PRywAHwAGIgPdkHvHfmtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKB"
		.	"dCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfIf2zbj3oFIQNVSzD/oo9jPSxCKBdCKBdCKBdCKBdC"
		.	"KBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo"
		.	"9TSi+ldF3ainhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdC"
		.	"KBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfcjHmiclvaj3qrbFlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdiPCu5dWLKf23Ngm/Ngm/Ngm/Ngm/Ngm/Ngm/Ngm/Ngm/Ngm/Ngm/Ngm/Ngm/Ngm/"
		.	"Ngm/Jf225dWNeOilCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBesbVrZjnnznIiWXk1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeWXk3/oo//oo//oo//oo//oo//oo//oo//oo//"
		.	"oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+RW0lCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeYYE7xmof/oo+OWUdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdHKxp2STi"
		.	"BUT+BUT/ul4T/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/qlIKBUT+BUT90SDdGKhpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeOWUf/oo//oo+OWUdCKBdCKBdCKBdCKBdCKBdCKBdCKB"
		.	"dCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBflkX7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/gjntCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeOWUf/oo//oo+OWU"
		.	"dCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBflkX7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/gjntCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKB"
		.	"dCKBdCKBdCKBeOWUf/oo//oo+OWUdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeZYE6naVenaVenaVenaVenaVenaVenaVenaVenaVenaVeWXk1CKBdCKBdCKBdCKBdC"
		.	"KBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeOWUf/oo//oo+OWUdCKBdCKBdSMiHBemjymofymofymofymofymofymofymofymofymofymofymofymofymofymofymofymofymofymofymofymofymofym"
		.	"ofymofymofymofymofymofymofymofymofymofymofymofymofymofymoe8d2VOLx9CKBdCKBeOWUf/oo//oo+OWUdCKBdCKBfJf23hjnx7TTt0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd"
		.	"0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd0SDd8TjzkkH3BemdCKBdCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9uRDNCKBdCKBdCKBdCKBdCKBdCKBdCK"
		.	"BdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd0SDf+oY5CKBdCKBeOWUf/oo//oo+OWUdCKBdH"
		.	"Kxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRhEKhlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH"
		.	"/oo9CKBdCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBerbFmzcV5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd"
		.	"CKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9CKBdCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRifZFNPMB9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdGKxnwmIbvl4VDK"
		.	"RhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9CKBdCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBemaVf/oo+KV0RCKBdCKBdCKBdCKB"
		.	"dCKBdCKBdCKBdCKBd/Tz7kkX7vmIVtRDNCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9DKRhCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBeE"
		.	"U0Hymoe4dGLAeWdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfGfWudY1G4dGKnaVdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9DKRhCKBeOWUf/oo//oo+OWUdCKBdHKxr"
		.	"/oo9lPy1CKBdCKBdCKBdCKBdCKBdmPi71nIl6TDpnQC70mohGKhpCKBdCKBdCKBdCKBdCKBdCKBdTMyL6n4xYNiV/Tz7gjntCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9DKRh"
		.	"CKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdDKRhmPi5oQC9uRDPqlYKYYE5CKBdDKRjsloNwRTRCKBdCKBdCKBdCKBdCKBdCKBeXX03LgW5CKBdKLRz4nYtdOilCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdC"
		.	"KBdCKBdCKBdCKBdrQjH/oo9DKRhCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdXNSTkkX7mkX/mkX+ycF5CKBdCKBdCKBe4dGKmaVdCKBdCKBdCKBdCKBdCKBdCKBfejXqEU0FCKBdCKBfHfmuXX01CKBdCKBdC"
		.	"KBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9EKhlCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeBUT/ci3lCKBdCKBdCKBdCKBdCKBdpQS/zm4dJ"
		.	"LRtCKBdCKBeNWEfRhXJCKBdCKBdCKBdCKBdFKhnVh3TymofymofymofymofKgG1CKBdCKBdrQjH/oo9FKhlCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdOLx/7oIxWNSR"
		.	"CKBdCKBdCKBdCKBewcF6zcV5CKBdCKBdCKBdVNCP7oIxRMiBCKBdCKBdCKBeLV0bjkH1oQC9oQC9oQC9oQC9XNiRCKBdCKBdrQjH/oo9GKhpCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKB"
		.	"dCKBdCKBdCKBdCKBdCKBfThXKLV0ZCKBdCKBdCKBdGKxnxmYZrQjFCKBdCKBdCKBdCKBfXiHaHVENCKBdCKBdGKxnnkn+CUj9CKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9FKhlCKBeOWUf/oo//oo+OWUdCKBdHK"
		.	"xr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBecY1DBemdCKBdCKBdCKBeCUUDjkH1CKBdCKBdCKBdCKBdCKBeeZFLBemhCKBdCKBeWXkzdjHlDKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9FKh"
		.	"lCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdmPi71m4hGKhpCKBdCKBfKf22ZYE5CKBdCKBdCKBdCKBdCKBdjPiz2nIlILBpLLh3vl4V3SzlCKBdCKBdCKBdCKBdCKBd"
		.	"CKBdCKBdCKBdrQjH/oo9FKhlCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRjslYJxRjRCKBdVNST6n4xVNSRCKBdCKBdCKBdCKBdCKBdCKBfnkn94SzmgZVPThXJ"
		.	"CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9FKhlCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBe2c2GnaVdCKBebYU/If2xCKBdCKBdCKBdCKBd"
		.	"CKBdCKBdCKBesbVrBemf1m4htRDNCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9FKhlCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeAUD7djHl"
		.	"CKBfhj32AUD5CKBdCKBdCKBdCKBdCKBdCKBdCKBdzSDb/oo/GfWtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9FKhlCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBd"
		.	"CKBdCKBdCKBdCKBdCKBdNMB77oIyBUT/ymodHKxpCKBdCKBdCKBdCKBdCKBdCKBdCKBdFKhndjHljPixCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9FKhlCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lP"
		.	"y1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfRhXLvl4WvblxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH/oo9FKhlCKBeO"
		.	"WUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBebYU//oo9oQC9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCK"
		.	"BdCKBdCKBdCKBdrQjH/oo9FKhlCKBeOWUf/oo//oo+OWUdCKBdHKxr/oo9lPy1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdcOSi8d2VCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd"
		.	"CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdrQjH7n41CKBdCKBeOWUf/oo//oo+OWUdCKBdGKxn+oY52SThCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKB"
		.	"dCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd7TTvwmIZCKBdCKBeOWUf/oo/0nIiWXkxCKBdCKBe0cl/tl4SdY1GaYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+"
		.	"aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+aYU+eZFLvmIW0cl9CKBdCKBeXX03ym4faj3urbFlCKBdCKBdJLRujZl"
		.	"TNgm/Ngm/Ngm/Ngm/Ngm/WiHXZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfZinfNgm/KgG2kZ1VILBpCKBdCKBesbVrajnmndV7YinZCKBdCKBdCKBdCKBdCKBdC"
		.	"KBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfai3ikc"
		.	"1xWTDD+oY5iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdC"
		.	"KBdCKBdCKBdCKBdCKBdCKBdjPSz/oo9USi8GIgPekHzFfGpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKB"
		.	"dCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBfHfmvckHsFIQMAHwBUSi//oo+UXUtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd"
		.	"CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeUXkv/oo9RSC0AHwAAHwAAHwCMZ0/8oI2UXUtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd"
		.	"CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBeUXkv8oI2KZk0AHwAAHwAAHwAAHwACIAGMZ0//oo/FfWtiPCtCKBdC"
		.	"KBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSzGfmv/oo+KZk0CIAEAHwAAHwAA"
		.	"HwAAHwAAHwAAHwBUSi/ekHz/oo/ai3irbFqWXkyOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWUeOWU"
		.	"eOWUeOWUeOWUeWXkyrbFnainj/oo/dkHtRSC0AHwAAHwAAHwAAHwAAHwAAHwAAHwAAHwAAHwAGIgNWTDCmdF3ajnr1nIn/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/"
		.	"/oo//oo//oo//oo//oo//oo//oo//oo/0nIjZjnqldF1VSzAFIQMAHwAAHwAAHwAAHwAAHwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
		.	"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
		.	"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
		.	"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
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

	Gui, load_BarGUI: -Border -Caption +ToolWindow HWNDhLoad_BarWin
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

		this.BG     	:= StrSplit(BG,"|")
		this.BG.W  	:= w
		this.BG.H  	:= h
		this.Width 	:=w
		this.Height	:=h
		this.FG       	:= StrSplit(FG,"|")
		this.FG.W  	:= this.BG.W - 2
		this.FG.H   	:= (fg_h:=(this.BG.H - 2))
		this.Percent 	:= 0
		this.X        	:= x
		this.Y        	:= y
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
		this.ApplyGradient(this.hLoaderBarBG	:= hLoaderBarBG,this.BG.1, this.BG.2, this.BG.3,1)

		Gui, Add, Text, x%fg_x% y%fg_y% w0 h%fg_h% 0xE hwndhLoaderBarFG
		this.ApplyGradient(this.hLoaderBarFG   	:= hLoaderBarFG,this.FG.1, this.FG.2, this.FG.3,1)

		Gui, Add, Text, x%x% y%y% w%w% h%h% 0x200 border center BackgroundTrans hwndhLoaderNumber c%FontColor%, % "[ 0 % ]"
			this.hLoaderNumber := hLoaderNumber

		if (this.ShowDesc) {
			Gui, Add, Text, xp y+2 w%w% h16 0x200 Center border BackgroundTrans hwndhLoaderDesc c%FontColorDesc%, Loading...
			this.hLoaderDesc := hLoaderDesc
			this.Height:=h+18
		}

		Gui,Font

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

		this.BG.W := LoaderBarBGW
		this.FG.W := LoaderBarBGW - 2
		this.Percent:=(p>=100) ? p:=100 : p

		PercentNum	:= Round(this.Percent,0)
		PercentBar	:= Floor((this.Percent/100)*(this.FG.W))

		hLoaderBarTitle		:= this.hLoaderBarTitle
		hLoaderBarFG  	:= this.hLoaderBarFG
		hLoaderNumber 	:= this.hLoaderNumber

		GuiControl,Move	,% hLoaderBarFG  	, % "w" PercentBar
		GuiControl,       	,% hLoaderNumber 	, % "[" PercentNum "% ]"

		if (this.ShowDesc) {
			hLoaderDesc := this.hLoaderDesc
			GuiControl,,%hLoaderDesc%, %w%
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
#Include %A_ScriptDir%\..\..\lib\class_Neutron.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}




