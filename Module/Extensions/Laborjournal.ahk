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
;                        			by Ixiko started in September 2017 - last change 16.03.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  ; Einstellungen
	#NoEnv
	#Persistent
	#KeyHistory, Off

	SetBatchLines, -1
	ListLines    	, Off

	global LabJ := Object()
	global adm := Object()

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

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
	Warnen	:= {	"nie"      	: 	"CHOL,LDL,TRIG,HDL,LDL/HD,HBA1CIFC"                                                 	; nie      	= werden nie gezeigt
						,	"immer" 	: 	"NTBNP,TROPI,TROPT,TROP,CKMB,K"                                                        	; immer 	= wenn pathologisch
						,	"exklusiv"	: 	"COVIPC-A,COVIP-SP,COVIGAB,COVIA,COVIG,COVMU501,COVM6970,"  	; exklusiv 	= zeigen auch wenn kein ausgeprägtes path. Ergebnis
											.	"ALYMP,ALYMPDIFF,ALYMPH,ALYMPNEO,ALYMPREA,ALYMPRAM,"               	;					und bei negativem Befund (z.B. COVIA, HIV)
											. 	"KERNS,MYELOC,PROMY,DIFANISO,DIFPOLYC,DIFPOIKL,"
											.	"DIFHYPOC,DIFMIKRO,DIFMAKRO,DIFOVALO,METAM,TROPOIHS,"
											.	"DDIM-CP,HBK-ST,HIV"}

  ; Laborjournal anzeigen
	LabPat     := AlbisLaborJournal("", "", Warnen, 140, false)
	LaborJournal(LabPat, false)


return

ESC:: ;{
	If WinActive("Laborjournal") {
		SaveGuiPos(labJ.hwnd)
		ExitApp
	}
return ;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Berechnungen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
AlbisLaborJournal(Von="", Bis="", Warnen="", GWA=100, Anzeige=true) {                 	;--  sehr abweichender Laborwerte

	/*  Parameter

		Von       	-	Anfangsdatum, leerer String oder Zahl
							eine Zahl als Wert wird als Anzahl der Werktage die rückwärtig vom aktuellem Tagesdatum angezeigt werden soll gewertet
		Bis         	-	Enddatum oder leer

	*/

		static Lab

		LabPat := Object()
		nieWarnen    	:= "\b(" StrReplace(Warnen.nie      	, ","	, "|") ")\b"
		immerWarnen 	:= "\b(" StrReplace(Warnen.immer 	, ","	, "|") ")\b"
		exklusivWarnen	:= "\b(" StrReplace(Warnen.exklusiv	, ","	, "|") ")\b"

	; durchschnittliche Abweichung von den Normwertgrenzen laden oder berechnen ;{
		If !IsObject(Lab) {
			If !FileExist(adm.LabDBPath "\Laborwertgrenzen.json")
				Lab := AlbisLaborwertGrenzen(adm.LabDBPath "\Laborwertgrenzen.json", Anzeige)
			else
				Lab := JSONData.Load(adm.LabDBPath "\Laborwertgrenzen.json", "", "UTF-8")
		}
	;}

	; Laborparameter lange Bezeichnungen laden
		LabDic := AlbisLabParam("load","short")

	; Suchzeitraum ;{

		; Von enthält kein Datum
			RegExMatch(Von, "^\d+$", Tage)
			Von 	:= Tage ? "" : Von
			Bis 	:= Tage ? "" : Bis

		; es werden Werktage berechnet
			If !Von && !Bis {

				Tage	    	:= Tage ? Tage : 14                            	; kein Zeitraum, Defaulteinstellung sind 14 Tage
				Von      	:= A_YYYY . A_MM . A_DD                	; aktueller Tag

			; Zahl der Tage berechnen, welche kein Werktag sind
				Wochen 	:= Floor(Tage/7)                               	; Tage in ganze Wochen umrechnen
				RestTage	:= Tage-Wochen*7                            	; verbliebene Resttage
				NrWTag	:= WeekDayNr(A_DDD)                		; Nr. des Wochentages
				WMod  	:= Mod(NrWTag-RestTage+1, 7)			; Divisonrest = Nummer des Wochentages minus berechnete RestTage + 1 durch 7
				PDays   	:= WMod = 0 ? 1 : WMod < 0 ? 2 : 0	; Divisonrest = 0 wäre Sonntag +1 Tag, < 0 alle Tage vor Sonntag also +2 Tage
				PDays   	:= PDays + Wochen*2                       	; + die Zahl der Wochenendtage der ganzen Wochen
				PDaysW	:= Floor(PDays/7)                               	; mehr als 7 Tage ist mindestens ein Wochende enthalten
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
	;}

	; Albis: Datenbank laden und entsprechend des Datums die Leseposition vorrücken ;{
		labDB := new DBASE(adm.AlbisDBPath "\LABBLATT.dbf", Anzeige)
		labDB.OpenDBF()
	;}

	; Datei: Startposition berechnen ;{
		If (VB = "Von") || (VB = "VonBis") {
			If !QJ
				QJ := SubStr(Von, 1, 4) . Ceil(SubStr(Von, 5, 2)/3)
			startrecord := Lab.seeks[QJ]
			labDB._SeekToRecord(startrecord, 1)
		}
	;}

	; filtert nach Datum und ab einer Überschreitung von der durchschnittlichen Abweichung von den Grenzwerten anhand
	; der zuvor für jeden Laborwert berechneten durchschnittlichen prozentualen Überschreitung des Grenzwertes aus den
	; Daten der LABBLATT.dbf [AlbisLaborwertGrenzen()]
		while !(labDB.dbf.AtEOF) {

				data := labDB.ReadRecord(["PATNR", "ANFNR", "DATUM", "LABID", "PARAM", "ERGEBNIS", "EINHEIT", "GI", "NORMWERT"])
				UDatum := data.Datum

			; Fortschritt
				If Anzeige && (Mod(A_Index, 3000) = 0)
					ToolTip, % ConvertDBASEDate(UDatum) ": " PSU "," PSO "`n" SubStr("00000000" labDB.recordnr, -5) "/" SubStr("00000000" labDB.records, -5) ;"`n" lfound

			; Datum passt nicht, weiter
				Switch VB {

					Case "Von":
						If (UDatum <= Von)
							continue
					Case "Bis":
						If (UDatum => Bis)
							continue
					Case "VonBis":
						If (UDatum < Von) || (UDatum > Bis)
							continue

				}

			; kein pathologischer Wert und in keiner Filterliste dann weiter ;{
				PNImmer := PNExklusiv := false
				PN := data.PARAM                                              	; Parameter-Name

				If RegExMatch(PN, nieWarnen)                            	; Parameter in nieWarnen - weiter
					continue
				If RegExMatch(PN, immerWarnen)                        	; Parameter immerWarnen - aufnehmen
					PNImmer  	:= true
				If RegExMatch(PN, exklusivWarnen)                    	; Parameter exklusivWarnen - aufnehmen
					PNExklusiv	:= true
				If !PNExklusiv && !PNImmer                                	; nicht in exklusiv und immer
					If !InStr(data["GI"], "+")                                    	; und Parameter nicht als path. markiert - weiter
						continue
			;}

			; Strings lesbarer machen (oder auch nicht)                        	;{
				P	    	:= ""
				PW   	:= StrReplace(data["ERGEBNIS"]	, ",", ".") 	; Parameter Wert
				PE	    	:= data.EINHEIT				                    	; Parameter Einheit
				ONG	:= UNG := 0                                        	; Obere/Untere Norm-Grenze
			;}

			; Grenzwertüberschreitungen berechnen
				If !PNExklusiv {

				  ; Normwert-String mit RegEx aufsplitten
					RegExMatch(data["NORMWERT"], "(?<V>[\<\>])*(?<SU>[\d,]+)\-*(?<SO>[\d,]+)*", P)
					PSU  	:= StrReplace(PSU	, ",", ".")                              	; unterer Grenzwert
					PSO  	:= StrReplace(PSO, ",", ".")                              	; oberer Grenzwert
					PSO  	:= PSO ? PSO : PSU                                       	; kein u.GW, dann gibt es nur einen o.GW

					If (PW >= PSO)                                                           	; Laborwert ist größer gleich dem oberen Grenzwert
						ONG := Floor(Round((PW * 100)/PSO, 3))               	; Oberere Normgrenze in Prozent
					else If (PW <= PSU)                                                   	; Laborwert ist kleiner gleich dem unteren Grenzwert
						UNG := Floor(Round((PSU * 100)/PW, 3))		   	     	; Untere Normgrenze in Prozent

					If !ONG && !UNG                                                      	; ONG und UNG sind beide null -> weiter
						continue

					If (ONG > 0)                                                            	; ONG überschritten
						plus := Round((ONG*100)/(lab.params[PN].OD))		; Überschreitung der ONG in Prozent
					else if (UNG > 0 )                                                     	; UNG überschritten
						plus := Round(((lab.params[PN].UD)*100)/UNG)    	; Unterschreitung der UNG in Prozent

				}

			; Laborwert liegt über der Grenzwert-Abweichung oder gehört zu einer Filterliste
				If (plus > GWA || PNImmer || PNExklusiv) {

						PatID 	:= data.PATNR
						Datum 	:= data.Datum
						PW    	:= RegExReplace(PW, "(\.[0-9]+[1-9])0+$", "$1")
						AN    	:= LabPat[Datum][PatID][PN].AN
						AN    	:= !RegExMatch(AN, "\b" datum.ANFNR "\b") ? AN "," datum.AFNR : AN

					; Subobjekte anlegen
						If !IsObject(LabPat[Datum])
							LabPat[Datum] := Object()
						If !IsObject(LabPat[Datum][PatID])
							LabPat[Datum][PatID] := Object()
						If !IsObject(LabPat[Datum][PatID][PN])
							LabPat[Datum][PatID][PN] := Object()

					; durchschnittliche Abweichungen von den Normgrenzen
					; wurde zuvor berechnet aus allen Werten in ihrer Datenbank dieses Parameters (alle Patienten)
						labOD	:= lab.params[PN].OD   	; durchschnittliche Abweichung von der oberen Normwertgrenze (oberer Durchschnitt in Prozent)
						labUD	:= lab.params[PN].UD    	; durchschnittliche Abweichung von der unteren Normwertgrenze (unterer Durchschnitt in Prozent)

					; Daten übergeben
						LabPat[Datum][PatID][PN].PatID	:= PatID
						LabPat[Datum][PatID][PN].CV  	:= PNImmer ? 1 : PNExklusiv ? 2 : 0                                                             	; CAVE-Wert für andere Textfarbe im Journal
						LabPat[Datum][PatID][PN].PN  	:= PN                                                                                                        	; Parameter Name
						LabPat[Datum][PatID][PN].PW  	:= RegExReplace(PW, "\.0+$")                                                                     	; Parameter Wert
						LabPat[Datum][PatID][PN].PV		:= P ? StrReplace(P, ",", ".") : data.NORMWERT                                             	; Parameter Varianz Normwertgrenzen
						LabPat[Datum][PatID][PN].PL 		:= (ONG > 0 ? "+" Round(ONG/100, 1)	: (UNG > 0 ? "-" UNG : ""))                                   	;
						LabPat[Datum][PatID][PN].PA		:= (ONG > 0 ? labOD : (UNG > 0 ? labUD : "" ))                                     	; Parameter Abweichung +/-
						LabPat[Datum][PatID][PN].PD		:= (ONG > 0 ? "+" ONG - labOD : (UNG > 0 ? "-" UNG - labUD : ""))      	;
						LabPat[Datum][PatID][PN].PE		:= PE                                                                                                    		; Einheit
						LabPat[Datum][PatID][PN].AN	:= LTrim(AN, ",")                                                                                          	; Labor Anforderungsnummer
						LabPat[Datum][PatID][PN].PB 	:= LabDic[PN].1                                                                                           	; Parameter Bezeichnung: ausgeschriebene Parameterbezeichnung

					}

		}

	; Lesezugriff beenden
		labDB.CloseDBF()
		LabDB := ""
		ToolTip

return LabPat
}

AlbisLaborJournal_new(Von="", Bis="", Warnen="", GWA=100, Anzeige=true) {       	;-- Laborwerte mit gewisser Überschreitung der Normgrenzen werden erfasst

		static Lab

		LabPat := Object()

		If IsObject(Warnen) {
			nieWarnen    	:= "(" StrReplace(Warnen.nie      	, ","	, "|") ")$"
			immerWarnen 	:= "(" StrReplace(Warnen.immer 	, ","	, "|") ")$"
			exklusivWarnen	:= "(" StrReplace(Warnen.exklusiv	, ","	, "|") ")$"
		}

	; Laborwertgrenzen laden oder berechnen                                                          	;{
		If !IsObject(Lab) {
			If !FileExist(adm.LabDBPath "\Laborwertgrenzen.json")
				Lab := AlbisLaborwertGrenzen(adm.LabDBPath "\Laborwertgrenzen.json", Anzeige)
			else
				Lab := JSON.Load(FileOpen(adm.LabDBPath "\Laborwertgrenzen.json", "r", "UTF-8").Read())
		}
	;}

	; Suchzeitraum                                                                                                  	;{
		If !Von && !Bis {
			Von 	:= A_YYYY A_MM A_DD
			If (A_DDD = "Fr")
				Von += -30, days
			else
				Von += -30, days

			FormatTime, Von, % Von, yyyyMMdd
			QJ 	:= A_YYYY . Ceil(A_MM/3)
			VB  	:= "Von"
		}
		else {
			VB := (Von && !Bis) ? "Von" : (!Von && Bis) ? "Bis" : "VonBis"
			QJ := SubStr(Von, 1, 4) . Ceil(SubStr(Von, 5, 2)/3)   ; Quartalsjahr QQYYYY z.B. 012021
		}

	;}

	; Laborbuchdaten laden                                                                                      	;{
		;~ p := ["ANFNR", "STATUS"]
		;~ s :=
		;~ labbuch := GetDBFData(adm.AlbisDBPath "\LABBUCH.dbf",p,,s)
	;}


	; Albis: Datenbank laden und entsprechend des Datums die Leseposition vorrücken ;{
		labDB := new DBASE(adm.AlbisDBPath "\LABBLATT.dbf", Anzeige)
		labDB.OpenDBF()
	;}

	; Datenbank: Startposition berechnen ;{
		startrecord := 0
		If (VB = "Von") || (VB = "VonBis") {
			startrecord := Lab.seeks[QJ] - 1
			labDB._SeekToRecord(startrecord, 1)
		}
		LabJ.srchdrecords	:= SubStr("00000000" (labDB.records - startrecord), -5)
		LabJ.records     	:= LabDB.records
	;}

	; filtert nach Datum und ab einer Überschreitung von der durchschnittlichen Abweichung von den Grenzwerten anhand
	; der zuvor für jeden Laborwert berechneten durchschnittlichen prozentualen Überschreitung des Grenzwertes aus den
	; Daten der LABBLATT.dbf [AlbisLaborwertGrenzen()]
		while !(labDB.dbf.AtEOF) {

			; eine Zeile lesen
				data := labDB.ReadRecord(["PATNR", "ANFNR", "DATUM", "PARAM", "ERGEBNIS", "EINHEIT", "GI", "NORMWERT"])
				UDatum := data.Datum

			; Fortschritt                                                     	;{
				If Anzeige && (Mod(A_Index, 3000) = 0)
					ToolTip, % ConvertDBASEDate(UDatum) ": " PSU "," PSO "`n" SubStr("00000000" labDB.recordnr, -5) "/" LabJ.srchdrecords ;"`n" lfound
			;}

			; Datum prüfen                                             	;{
				Switch VB {

					Case "Von":
						If (UDatum <= Von)
							continue
					Case "Bis":
						If (UDatum => Bis)
							continue
					Case "VonBis":
						If (UDatum < Von) || (UDatum > Bis)
							continue

				}
			;}

			; Kriterien prüfen                                           	;{
				PNImmer := PNExklusiv := false
				PN := data.PARAM                                            	; Parameter (Name)

				If RegExMatch(PN, "(" nieWarnen ")")
					continue
				If RegExMatch(PN, "(" immerWarnen ")")
					PNImmer := true
				If RegExMatch(PN, "(" exklusivWarnen ")")
					PNExklusiv := true
				If !PNExklusiv && !PNImmer
					If !InStr(data["GI"], "+")
						continue
				;}

			; Strings lesbarer machen (oder auch nicht)    	;{
				P	    	:= ""
				PW   	:= StrReplace(data["ERGEBNIS"]	, ",", ".") 	; Wert
				PE	    	:= data.EINHEIT				                    	; Einheit
			;}

			/* Beschreibung der Variablen

					GWA	- Grenzwertabweichung in %

					ONG	- Abweichung vom oberen Normgrenzwert in %
					UNG	- Abweichung vom unteren Normgrenzwert in %

					P        	- (P)arameter
					P(N)    	- (N)ame
					P(W)    	- (W)ert
					P(E)    	- (E)inheit

					P(SU)	- unterer Grenzwert aus dem Laborblatt
					P(SO)	- oberer Grenzwert aus dem Laborblatt

				# diese Daten werden vorher aus allen Einträgen der LABBLATT.dbf für alle Parameter berechnet

					lab.param[PN].OD 	- durchschnittliche Abweichung vom oberen Grenzwert

			*/

			; Grenzwertüberschreitungen berechnen          	;{
				If !PNExklusiv {           ; exklusiv Parameter werden immer gelistet

					RegExMatch(data["NORMWERT"], "(?<V>[\<\>])*(?<SU>[\d,]+)\-*(?<SO>[\d,]+)*", P)
					PSU  	:= StrReplace(PSU	, ",", ".")                                     	; unterer Grenzwert
					PSO  	:= StrReplace(PSO, ",", ".")                                       	; oberer Grenzwert
					PSO  	:= PSO ? PSO : PSU                                               	; wenn es nur einen Grenzwert gibt
					ONG 	:= (PW >= PSO) 	? Round((PW*100)/PSO	, 2) 	: 0	; Grenzwertüberschreitung in %
					UNG 	:= (PW <= PSU) 	? Round((PSU*100)/PW	, 2)	: 0	; Grenzwertunterschreitung in %

				; weiter wenn keine Grenzwerte über- oder unterschritten wurden
					If !(ONG && UNG)
						continue

				; berechnet die prozentuale Abweichung von der durchschnittlichen Grenzwertabweichung dieses Parameters
					plus := (ONG > 0) ? Round((ONG*100)/lab.params[PN].OD) : (UNG > 0) ? Round((lab.params[PN].UD*100)/UNG) : 0

				}
			;}

			; Datenobjekt erstellen                                    	;{
				If (plus >= GWA || PNImmer || PNExklusiv) {

						PatID 	:= data.PATNR
						Datum 	:= data.Datum
						AFNR	:= data.AFNR

						If !IsObject(LabPat[Datum])
							LabPat[Datum] := Object()
						If !IsObject(LabPat[Datum][PatID])
							LabPat[Datum][PatID] := Object()
						If !IsObject(LabPat[Datum][PatID][PN])
							LabPat[Datum][PatID][PN] := Object()
						If !IsObject(LabPat[Datum][PatID].AFNR) 			; Anforderungsnummern und Teil oder Endbefund zusammen bringen
							LabPat[Datum][PatID].AFNR := Object()

						PW := RegExReplace(PW, "(\.[0-9]+[1-9])0+$", "$1")
						LabPat[Datum][PatID][PN].PatID	:= PatID
						LabPat[Datum][PatID][PN].AFNR	:= AFNR                                                                                                                                        	; Anforderungsnummer
						LabPat[Datum][PatID][PN].CV  	:= PNImmer ? 1 : PNExklusiv ? 2 : 0                                                                                                	; CAVE-Wert für andere Textfarbe im Journal
						LabPat[Datum][PatID][PN].PN  	:= PN                                                                                                                                           	; Parameter Name
						LabPat[Datum][PatID][PN].PW  	:= RegExReplace(PW, "\.0+$")                                                                                                        	; Parameter Wert
						LabPat[Datum][PatID][PN].PV		:= P ? StrReplace(P, ",", ".") : data.NORMWERT                                                                                  	; Parameter Normwert
						LabPat[Datum][PatID][PN].PL 		:= (ONG > 0 ? "+" ONG	: (UNG > 0 ? "-" UNG : ""))                                                                      	; obere Normwertgrenze
						LabPat[Datum][PatID][PN].PA		:= (ONG > 0 ? lab.params[PN].OD : (UNG > 0 ? lab.params[PN].UD : "" ))                                     	; durchschn. Abweichung vom Normwert [+/-] in %
						LabPat[Datum][PatID][PN].PD		:= (ONG > 0 ? "+" ONG - lab.params[PN].OD : (UNG > 0 ? "-" UNG - lab.params[PN].UD : ""))  	; untere Normwertgrenze
						LabPat[Datum][PatID][PN].PE		:= PE                                                                                                                                           		; Einheit

					}
			;}

		}

	; Lesezugriff beenden
		labDB.CloseDBF()
		LabDB := ""
		ToolTip

return LabPat
}

AlbisLaborwertGrenzen(LbrFilePath, Anzeige=true) {                                               	;-- Berechnung der durchschnittlichen Abweichung v. Grenzwert

	/* 	Beschreibung - AlbisLaborwertGrenzen()

		⚬	Berechnet aus den Daten der LABBLATT.dbf die durchschnittliche prozentuale Über- oder Unterschreitung eines Laborparameters.
		⚬	für eventuelle Anpassungen wird die maximalste Über- oder Unterschreitung als Einzelwert gespeichert
		⚬	Durch Nutzung eines Faktors (Prozent) sind die dadurch mit Annährung erreichte "Warngrenze" auch bei unterschiedlichen
			Einheiten praktikabel, außerdem entfällt das Anpassen an die sich wechselnden Normbereiche.

		⊛ Die Ausgabe nur wirklich auffälliger pathologischer Laborwerte war das Ziel. Bei über 300 Laborparametern per Hand einen Alarm-
			wert anzulegen ist aufgrund sich häufig ändernder Normbereich unsinnig. Die Idee war nur die Über- und Unterschreitung der Grenzwerte
			zu beachten. Normwerte werden nicht "gewichtet". Die einfachste Methode ist die Ermittlung des Durchschnitts um minimale und maximale
			Überschreitungen ausgleichen zu können. Die erreichten Alarmwerte sind in vielen Fällen gut und an der ärztlichen Wirklichkeit, in vielen
			anderen Fällen aber auch nicht. Diese Funktion filtert viel unnütze Information heraus, macht dies aber noch zu einheitlich mechanisch.

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

				If !InStr(data["GI"], "+")
					continue

				PN	:= data.PARAM
				PE		:= data.EINHEIT
				PW	:= StrReplace(data["ERGEBNIS"]	, ",", ".")
				RegExMatch(data["NORMWERT"], "(?<V>[\<\>])*(?<SU>[\d,]+)\-*(?<SO>[\d,]+)*", P)
				PSU	:= StrReplace(PSU	, ",", ".")
				PSO	:= StrReplace(PSO, ",", ".")

			; berechnet die Grenzwertüberschreitungen
				ONG :=UNG := 0
				PSO := StrLen(PSO) > 0 ? PSO : PSU
				If (PW >= PSO) {
					ONG := Floor(Round((PW * 100)/PSO, 3))
				}
				else If (PW <= PSU) {
					UNG := Floor(Round((PSU * 100)/PW, 3))
				}

				If !ONG && !UNG
					continue

			; Ersterstellung eines Laborwert Datensatzes
				If !Lab.HasKey(PN) {

					Lab[PN] := {	"PZ" 	: 1			; Parameterzähler
									, 	"O" 	: 0			; Summe d. prozentualen Abweichung vom oberen Grenzwert
									, 	"OI"	: 0			; Abweichungszähler (obere Grenzwerte)
									, 	"OD"	: 0			; durchschittliche prozentuale Abweichung vom oberen Grenzwert (O/OI)
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
		JSONData.Save(LbrFilePath, Labwrite, true,, 1, "UTF-8")

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
			winpos := "w950 h600"

	; Patientendaten laden
		PatDB 	:= PatientDBF(adm.AlbisDBPath, ["NR", "NAME", "VORNAME", "GEBURT"],, Anzeige)

	; Variablen bestücken
		srchdrecords  	:= LTrim(labJ.srchdrecords, "0")
		dbrecords     	:= labJ.records
		Tagesanzeige	:= LabJ.Tagesanzeige

	; HTML Vorbereitungen ;{
		static TDL      	:= "<TD style='text-align:Left'>"
		static TDL1     	:= "<TD style='text-align:Left;border-left:0px solid'>"
		static TDR      	:= "<TD style='text-align:Right'>"
		static TDR1     	:= "<TD style='text-align:Right;border-right:0px solid'>"
		static TDR2a     	:= "<TD style='text-align:Right;border-bottom:0px'>"
		static TDR2b     	:= "<TD style='text-align:Right;border-top:0px;border-bottom:0px'>"
		static TDR3      	:= "<TD class='tooltip' data-tooltip='##' style='text-align:Right'>"                    ; Param
		static TDC     	:= "<TD style='text-align:Right'>"
		static TDJ      	:= "<TD style='text-align:Justify'>"

		static cCol1   	:= ";color: Red"
		static FWeight1	:= ";font-weight: bold"
		static cCol2    	:= ";color: BlueViolet"
		static FWeight2	:= ";font-weight: bold"
		static cCol3   	:= ";color: DarkTeal"
		static FWeight3	:= ";font-weight: bold;font-family: sans-serif"

	; HTML Seitendaten
		htmlheader =
		(
			<!DOCTYPE html>
			<html lang="de">
			<head>
			<meta http-equiv="X-UA-Compatible" content="IE=edge">
			<title>Laborjournal</title>
			<meta charset="utf-8">
			<style>

			header {
				width:        	100`%;
				display:       	flex;
				background: #6e6c6c;
				font-family: 	Segoe UI;

				margin:     	0;
				padding:    	0;
			}

			.title-bar {
				overflow: hidden;
				font-family: Segoe UI;
				flex-grow: 1;
				font-size: 100`%;
			}

			.title-txt {
				font-family: Segoe Script;
				color:#D9D8D8;
				margin-top:2px;
				vertical-align: center;
			}

			.title-btn {
				padding: 0.35em 1.0em;
				cursor: pointer;
				vertical-align: bottom;
				font-family: Webdings;
				font-size: 11pt;
			}

			.title-btn:hover {
			  background: rgba(0, 0, 0, .2);
			}

			.title-btn-close:hover {
			  background: #dc3545;
			}

			.title-menu {
				float:                 	left;
				cursor:              	pointer;
				vertical-align:    	bottom;
				font-family:       	Segoe Script;
				font-size:           	10pt;
				background:     	#999898;
				padding-left:      	12px;
				padding-right:    	12px;
				margin-bottom: 	1px;
				margin-top:      	3px;
				border-radius:    	8px;
				transition-timing-function: ease-in-out;
			}

			.title-menu-btn1 {
				position: absolute;
				left: 70px;
			}

			.title-menu-btn2 {
				position: absolute;
				left: 194px;
			}

			.title-menu-btn3 {
				position: absolute;
				left: 266px;
			}

			.title-menu-active {
				background-color: #70A0A6;
			}

			.title-menu:hover {
			  background: #AEACAC;
			}

			html, body {
				width: 100`%; height: 100`%;
				margin: 0; padding: 0;
				font-family: sans-serif;
			}

			body {
				display: flex;
				flex-direction: column;
			}
		)

		htmltable =
		(

			table {
				font-family: Segoe UI, sans-serif;
				border-collapse: collapse;
				overflow-x:auto;
				width: 100`%;
			}

			.table-btn1 {
			  background: #ffffff;
			}

			.table-btn1:hover {
				background: rgba(0, 0, 0, .2);
				cursor: pointer;
			}

			.table-btn2 {
			  background: #f1f1c1;
			}

			.table-btn2:hover {
				background: rgba(0, 0, 0, .2);
				cursor: pointer;
			}

			.table-btn3a {
			  background: #ffffff;
			}

			.table-btn3a:hover {
				cursor: info;
			}

			.table-btn3b {
			  background: #f1f1c1;
			}

			.table-btn3b:hover {
				cursor: info;
			}

			th {
				border: 1px solid #888888;
				background: #70A0A6;
				text-align: left;
				padding: 4px;
				font-size: 11;
			}

			td {
				border-left: 1px solid #888888;
				border-right: 1px solid #888888;
				padding-right: 4px;
				padding-left: 4px;
				font-size: 90`%;
			}

			#td1a {
				background-color: #ffffff;
				padding-top: 18px;
				padding-bottom: 2px;
				border-top: 3px solid #555555;
				border-left: 1px solid #888888;
				border-right: 1px solid #888888;
				border-bottom: 0px solid #888888;
			}

			#td1b {
				background-color: #ffffff;
				padding-top: 2px;
				padding-bottom: 2px;
				border-top: 0px solid #555555;
				border-left: 1px solid #888888;
				border-right: 1px solid #888888;
				border-bottom: 0px solid #888888;
			}

			#td2a {
				background-color: #f1f1c1;
				padding-top: 18px;
				padding-bottom: 2px;
				border-top: 3px solid #555555;
				border-left: 1px solid #888888;
				border-right: 1px solid #888888;
				border-bottom: 0px solid #888888;
			}

			#td2b {
				background-color: #f1f1c1;
				padding-top: 2px;
				border-top: 0px solid #555555;
				border-left: 1px solid #888888;
				border-right: 1px solid #888888;
				border-bottom: 0px solid #888888;
			}

			TD.tooltip {
				cursor:                      	default;
				position:                    	relative;
				text-decoration:         	none;
			}

			TD.tooltip:hover {
				background: rgba(0, 0, 0, .1);
			}

			TD.tooltip:after {
				content:                    	attr(data-tooltip);
				font-family:                	Segoe UI, sans-serif;
				font-style:                    	normal;
				font-weight:                	normal;
				font-size:                    	80`%;
				position:                   	absolute;
				bottom:                     	0`%;
				left:                              	30`%;
				background:               	#ffcb66;
				padding:                    	3px 8px;
				color:                         	black;
				-webkit-border-radius: 	10px;
				-moz-border-radius : 	10px;
				border-radius :             	10px;
				white-space:              	nowrap;
				opacity:                     	0;
				-webkit-transition:      	all 0.4s ease;
				-moz-transition :          	all 0.4s ease;
				transition :                 	all 0.4s ease;
			}

			TD.tooltip:before {
				content:                     	"";
				position:                     	absolute;
				width:                         	0;
				height:                         	0;
				border-top:                 	20px solid #ffcb66;
				border-left:                 	20px solid transparent;
				border-right:             	20px solid transparent;
				-webkit-transition:      	all 0.4s ease;
				-moz-transition :         	all 0.4s ease;
				transition :                 	all 0.4s ease;
				opacity:                     	0;
				bottom:                     	60`%;
				left:                             	90`%;
			}

			TD.tooltip:hover:after {
				bottom:                      	100`%;
			}

			TD.tooltip:hover:before {
				bottom:                     	70`%;
			}

			TD.tooltip:hover:after, a:hover:before {
				opacity: 1;
			}

			.main {
				display: flex;
				flex-grow: 1;
				padding: 1.5em;
				 justify-content: space-between;
			}

			</style>
			</head>

		)

		htmlbody =
		(
			 <header id='LaborJournal_Header'>
				<div class='title-bar' onmousedown='neutron.DragTitleBar()'>            </div>
				<div class='title-menu title-menu-btn1' onclick='neutron.Menu_LabJournal()'>Laborjournal</div>
				<div class='title-menu title-menu-btn2' onclick='neutron.Menu_LabJ_Suche()'>Suche</div>
				<div class='title-menu title-menu-btn3' onclick='neutron.Menu_LabJ_Props()'>Einstellungen</div>
				<div class='title-txt'>%Tagesanzeige%</div>
				<div class='title-btn' onclick='neutron.Minimize()'>0</div>
				<div class='title-btn' onclick='neutron.Maximize()'>1</div>
				<div class='title-btn title-btn-close' onclick='ahk.LabJournal_Close()'>r</div>
			</header>

			<body>
			<div STYLE='overflow-y:auto;'>
			<table id='LabJournal_Table'>

				<tr>
					<th style='text-align:Right'>Datum</th>
					<th style='text-align:Left'>Patient</th>
					<th style='text-align:Right'>Param</th>
					<th style='text-align:Right'>Wert</th>
					<th style='text-align:Center' colspan='2'>Normalwerte</th>
					<th style='text-align:Right'; >+/-</th>
					<th style='text-align:Right'>⍉ Abw.</th>
				</tr>

		)

		html := RegExReplace(htmlheader . htmltable . htmlbody, "^\s{16}(.*[\n\r]+)", "$1")
	;}

	; HTML Tabelle wird erstellt
		trIDf:= false
		sortDatum := Array()

		For Datum, Patients in LabPat
			sortDatum.InsertAt(1, Datum)

		For idx, Datum in sortDatum {

				If (DatumLast <> Datum)
					UDatum := ConvertDBASEDate(Datum)

				For PatID, parameter  in LabPat[Datum] {

					If (PatIDLast <> PatID) || (DatumLast <> Datum) {
						Patient  	:= PatDB[PatID].NAME ", " PatDB[PatID].VORNAME " (" ConvertDBASEDate(PatDB[PatID].GEBURT) ") [" PatID "]"
						PatIDLast	:= PatID
						trIDf      	:= !trIDf	; flag für alternierenden Farbwechsel
					}

					For key, PN in parameter {

						TRPat    	:= "<TR" (UDatum 	? (trIdf ? " id='td1a'" : " id='td2a'") : (trIdf ? " id='td1b'" : " id='td2b'") ) ">"
						tdPat     	:= "<TD" (Patient  	? (trIdf ? " class='table-btn1'" : " class='table-btn2'") " onclick='ahk.LabJournal_KarteiKarte(event)'>" : ">")
						tbDatum 	:= UDatum ? SubStr(UDatum, 1, 6) : ""
						TDRDX		:= UDatum ? TDR2a : TDR2b

						TDRX      	:= PN.CV = 0 ? TDR 	: (PN.CV = 1 ? StrReplace(TDR 	, "'>"	, cCol1 FWeight1 "'>") : StrReplace(TDR  	, "'>", cCol2 FWeight2 "'>"))
						TDR1X     	:= PN.CV = 0 ? TDR1	: (PN.CV = 1 ? StrReplace(TDR1, "'>"	, cCol1 FWeight1 "'>") : StrReplace(TDR1	, "'>", cCol2 FWeight2 "'>"))

						TDR4		:= StrReplace(TDR3, "##", PN.PB)
						TDR2X     	:= PN.CV = 0 ? TDR4 	: (PN.CV = 1 ? StrReplace(TDR4, "'>", cCol1 FWeight1 "'>") : StrReplace(TDR4 	, "'>", cCol2 FWeight2 "'>"))

						TDCX      	:= PN.CV = 0 ? TDC 	: (PN.CV = 1 ? StrReplace(TDC	, "'>"	, cCol1 FWeight1 "'>") : StrReplace(TDC 	, "'>", cCol2 FWeight2 "'>"))
						TDL1X     	:= PN.CV = 0 ? TDL1	: (PN.CV = 1 ? StrReplace(TDL1	, "'>"	, cCol1 FWeight1 "'>") : StrReplace(TDL1	, "'>", cCol2 FWeight2 "'>"))

					; Abweichung oberhalb der Norm und PN.CV (CAVE) negativ (CV - wenn wahr dann wird andere Farbkennzeichnung)
						If (PN.PL > 250 && !PN.CV) {
							TDRX      	:= StrReplace(TDR 	, "'>"	, cCol3 FWeight3 "'>")
							TDR1X     	:= StrReplace(TDR1	, "'>"	, cCol3 FWeight3 "'>")
							TDCX      	:= StrReplace(TDC	, "'>"	, cCol3 FWeight3 "'>")
							TDL1X     	:= StrReplace(TDL1	, "'>"	, cCol3 FWeight3 "'>")
						}

						If PN.PA {

							RegExMatch(PN.PV, "(?<V>[\<\>])*(?<SU>[\d.]+)\-*(?<SO>[\d.]+)*", P)
							PSM  	:= PSO ? PSO : PSU
							PA     	:= Round(PN.PA * (PSM/100), 1)
							PA     	:= RegExReplace(PA, "\.0+$")

						}

						PE := Trim(PN.PE)	? PN.PE : "" ;" - - - - -"

						html .= "`t`t`t" TRPat "`n"
								.  	" "	TDRDX	. 	tbDatum                             	"</td>"	; Abnahmedatum                	[Datum]
								. 	" "	tdPat  	.	Patient             	                  	"</td>"	; Patientenname                    [Patient]
								. 	" "	TDR2X 	.	PN.PN                                 	"</td>"	; Parameter Name              	[Param]
								. 	" "	TDRX	.	PN.PW                               	"</td>"	; Parameter Wert               	[Wert]
								.	" "	TDR1X	.	PN.PV			                     	"</td>"	; Parameter Normwerte        	[Normalwerte]
								.	" "	TDL1X 	.	PE                                       	"</td>"	; Parameter Einheit                [      ]
								. 	" "	TDCX	.	(PN.PL	? PN.PL "x" 	: "")     	"</td>"	; Parameter Lage                   [+/-]
								. 	" "	TDRX	.	(PA   	? PA " " PE  	: "") 		"</td>"	; Parameter Abweichung     	[⍉ Abw.]
								. 	" </TR>`n`n"

						PE := PA := UDatum := Patient := ""

					}

				}

				DatumLast := Datum

		}

		html .= "</tbody></table></body></html>"

	; erstellte HTML Seite wird angezeigt
		FileOpen(A_Temp "\Laborjournal.html", "w", "UTF-8").Write(html)
		neutron := new NeutronWindow("","","","Laborjournal" LabJ.maxrecords ")", "+AlwaysOnTop minSize900x400")
		neutron.Load(A_Temp "\Laborjournal.html")
		neutron.Gui("+LabelNeutron")
		neutron.Show(winpos)
		hLJ := WinExist("A")
		WinSet, AlwaysOnTop,, % "ahk_id " hLJ

		obj := neutron.wb.document.getElementById("LaborJournal_Header")	, hcr	:= obj.getBoundingClientRect()
		obj := neutron.wb.document.getElementById("LabJournal_Table")    	, tcr	:= obj.getBoundingClientRect()

		labJ.hwnd := hLJ
		labJ.HeaderHeight	:= Floor(hcr.Bottom + 1)
		labJ.TableHeight   	:= Floor(tcr.Bottom  + 1)
		labJ.enrolled	    		:= true

		neutron.Show(winpos)

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
		RegExMatch(event.target.innerText, "^(?<Name>.*)\[(?<ID>\d+)\]", Pat)

		If labJ.enrolled ="x" {

			lj := GetWindowSpot(labJ.hwnd)
			dh	:= lj.h
			dtb	:= labJ.TableHeight - labJ.HeaderHeight + 1
			dStep:= Floor(dtb/25)
			Loop 25 {
				dh -= dStep
				SetWindowPos(labJ.hwnd, lj.x, lj.y, lj.w, dh)
			}

			labJ.enrolled := false

		}

		AlbisAkteOeffnen(PatName, PatID)
		AlbisLaborBlattZeigen()


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


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
SaveGuiPos(hwnd) {

	win 			:= GetWindowSpot(hwnd)
	winPos	 	:= "x" win.X " y" win.Y " w" win.CW " h" win.CH
	IniWrite, % winpos	, % adm.Ini, % adm.compname, LaborJournal_Position

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

PatientDBF(basedir, infilter="", outfilter="", debug=0) {                                            	;-- gibt nur benötigte Daten der albiswin\db\PATIENT.DBF zurück

	; Rückgabeparameter ist ein Objekt mit Patienten Nr. und dazugehörigen Datenobjekten (die key's sind die Feldnamen in der DBASE Datenbank)

		PatDBF := Object()

	; für Abrechnungsüberprüfungen die geschätzt minimal notwendigste Datenmenge
		If !IsObject(infilter)
			infilter := ["NR", "NAME", "VORNAME", "GEBURT", "MORTAL", "LAST_BEH"]

	; liest alle Patientendaten in ein temp. Objekt ein
		database 	:= new DBASE(basedir "\PATIENT.dbf", debug)
		res        	:= database.OpenDBF()
		matches	:= database.GetFields(infilter)
		res         	:= database.CloseDBF()

	; temp. Objekt wird nach PATNR sortiert
		For idx, m in matches {

			strObj	:= Object()
			For key, val in m
				If (key <> "NR")
					strObj[key] := val

			PatDBF[m.NR] := strObj

		}

return PatDBF
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




