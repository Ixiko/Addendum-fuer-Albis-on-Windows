; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 	Addendum Laborjournal
;
;      Funktion:           	zeigt nur die sehr auffälligen Laborwerte seiner Patienten in Form eines HTML Journals an. Das visuelle Durchforsten des Laborbuches soll damit entfallen.
;
;
;		Hinweis:				Nichts ist ohne Fehler. Auch das hier sicher nicht. Nicht drauf verlassen! Skript soll nur eine Unterstützung sein!
;									Das Skript filtert noch keine Werte heraus die nur positiv oder negativ sein können (z.B. SARS-CoV-2 PCR)
;									INR Werte werden angezeigt, wenn ein Pat. Falithrom/Marcumar nimmt, der Wert aber nicht als therapeutischer INR gespeichert wurde.
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 11.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  ; Einstellungen
	#NoEnv
	#Persistent
	#KeyHistory, Off
	;#MaxMem 4096
	SetBatchLines, -1
	ListLines    	, Off

	global LabJ := Object()
	global adm := Object()

  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	adm.Dir            	:= AddendumDir
	adm.Ini              	:= AddendumDir "\Addendum.ini"
	adm.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	adm.AlbisDBPath 	:= AlbisPath "\DB"
	adm.compname	:= StrReplace(A_ComputerName, "-")                    	; der Name des Computer auf dem das Skript läuft

  ; hier alle Parameter hineinschreiben bei denen Sie keine Warnung erhalten wollen
	Warnen		:= {	"nie"      	: 	"CHOL,LDL,TRIG,HDL,LDL/HD,HBA1CIFC"                             	; nie      	= werden nie gezeigt
							,	"immer" 	: 	"NTBNP,TROPI,TROPT,TROP,CKMB,K"                                       	; immer 	= wenn pathologisches Ergebnis
							,	"exklusiv"	: 	"HBK-ST,COVIPC-A,COVIP-SP,COVIGAB,COVIA,COVIG,HIV,"  	; exklusiv 	= zeigen auch wenn kein ausgeprägtes path. Ergebnis
												  .	"KERNS,ALYMPNEO,PROMY,DIFANISO,MYELOC,METAM,"
												  .	"TROPOIHS,DDIM-CP"}

  ; Laborjournal anzeigen
	LabPat      	:= AlbisLaborJournal("", "", Warnen, 105, true)
	LaborJournal(LabPat, false)

return

ESC::
	If WinActive("Laborjournal") {
		SaveGuiPos(labJ.hwnd)
		ExitApp
	}
return

AlbisLaborJournal(Von="", Bis="", Warnen="", GWA=100, Anzeige=true) {               	;-- Laborwerte mit gewisser Überschreitung der Normgrenzen werden erfasst

		static Lab

		LabPat := Object()

		If IsObject(Warnen) {
			nieWarnen    	:= "(" StrReplace(Warnen.nie      	, ","	, "|") ")$"
			immerWarnen 	:= "(" StrReplace(Warnen.immer 	, ","	, "|") ")$"
			exklusivWarnen	:= "(" StrReplace(Warnen.exklusiv	, ","	, "|") ")$"
		}

	; Laborwertgrenzen laden oder berechnen                                                          	;{
		If !IsObject(Lab) {
			If !FileExist(adm.DBPath "\sonstiges\Laborwertgrenzen.json")
				Lab := AlbisLaborwertGrenzen(adm.DBPath "\sonstiges\Laborwertgrenzen.json", Anzeige)
			else
				Lab := JSON.Load(FileOpen(adm.DBPath "\sonstiges\Laborwertgrenzen.json", "r", "UTF-8").Read())
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

				If RegExMatch(PN, "\b(" nieWarnen ")\b")
					continue
				If RegExMatch(PN, "\b(" immerWarnen ")\b")
					PNImmer := true
				If RegExMatch(PN, "\b(" exklusivWarnen ")\b")
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

	; Dateipositionen für quartalsweise Verarbeitung werden hinzugefügt
		Labwrite.seeks := Object()
		For QJ, filepos in seeks
			Labwrite.seeks[QJ] := filepos

	; Speichern der Daten
		FileOpen(LbrFilePath, "w", "UTF-8").Write(JSON.Dump(Labwrite,, 1))

return Labwrite
}

AlbisLabParam() {                                                                                                  	;-- Laborparameter als Excel (.csv) Datei speichern

	; Albis: Inhalt der LABPARAM.dbf komplett laden
		LabDB := new DBASE(adm.AlbisDBPath "\LABPARAM.dbf", Anzeige)
		LabDB.OpenDBF()
		LabParam := labDB.GetFields("")
		LabDB.OpenDBF()

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

		FileOpen(adm.DBPath "\sonstiges\LabParam.csv", "w", "UTF-8").Write(h . t)

}

LaborJournal(LabPat, Anzeige=true) {

	; letzte Fensterposition laden
		IniRead, winpos, % adm.ini, % adm.compname, LaborJournal_Position
		If (InStr(winpos, "ERROR") || StrLen(winpos) = 0)
			winpos := "w950 h600"

	; Patientendaten laden
		PatDB := PatientDBF(adm.AlbisDBPath, ["NR", "NAME", "VORNAME", "GEBURT"],, Anzeige)

	; HTML Vorbereitungen ;{
		static TDL      	:= "<TD style='text-align:Left'>"
		static TDL1     	:= "<TD style='text-align:Left;border-left:0px solid'>"
		static TDR      	:= "<TD style='text-align:Right'>"
		static TDR1     	:= "<TD style='text-align:Right;border-right:0px solid'>"
		static TDR2a     	:= "<TD style='text-align:Right;border-bottom:0px'>"
		static TDR2b     	:= "<TD style='text-align:Right;border-top:0px;border-bottom:0px'>"
		static TDC     	:= "<TD style='text-align:Center'>"
		static TDJ      	:= "<TD style='text-align:Justify'>"

		static cCol1   	:= ";color: Red"
		static FWeight1	:= ";font-weight: bold"
		static cCol2    	:= ";color: BlueViolet"
		static FWeight2	:= ";font-weight: bold"
		static cCol3   	:= ";color: DarkTeal"
		static FWeight3	:= ";font-weight: bold"

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
			  width: 100`%;
			  display: flex;
			  background: #6e6c6c;
			  font-family: Segoe UI;
			  font-size: 90`%;
			}

			.title-bar {
			  font-family: Segoe UI;
			  padding: 0.35em 0.5em;
			  flex-grow: 1;
			  font-size: 90`%;
			}

			.title-txt {
				width: 50 px;
				 font-size: 30`%;
				 vertical-align: bottom;
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
				background: rgba(0, 0, 0, .4);
				cursor: pointer;
			}

			.table-btn2 {
			  background: #f1f1c1;
			}

			.table-btn2:hover {
				background: rgba(0, 0, 0, .4);
				cursor: pointer;
			}

			th {
				border: 1px solid #888888;
				background-color:powderblue;
				text-align: left;
				padding: 8px;
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

			.main {
				display: flex;
				flex-grow: 1;
				padding: 0.5em;
				 justify-content: space-between;
			}

			</style>
			</head>

		)

		srchdrecords := LTrim(labJ.srchdrecords, "0")
		dbrecords 	:= labJ.records
		htmlbody =
		(
			 <header id="LaborJournal_Header" >
				<span class='title-bar' onmousedown='neutron.DragTitleBar()'>Laborjournal (Datensätze durchsucht: %srchdrecords%, Datensätze gesamt: %dbrecords%)</span>
				<span class='title-btn' onclick='neutron.Minimize()'>0</span>
				<span class='title-btn' onclick='neutron.Maximize()'>1</span>
				<span class='title-btn title-btn-close' onclick='ahk.LabJournal_Close()'>r</span>
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
					<th style='text-align:Center'; >+/-</th>
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

				If (UDatumLast <> Datum) {
					UDatum := ConvertDBASEDate(Datum)
					UDatumLast := Datum
				}

				For PatID, parameter  in LabPat[Datum] {

					If (PatIDLast <> PatID) {
						Patient  	:= PatDB[PatID].NAME ", " PatDB[PatID].VORNAME " (" ConvertDBASEDate(PatDB[PatID].GEBURT) ") [" PatID "]"
						PatIDLast	:= PatID
						trIDf      	:= !trIDf	; flag für alternierenden Farbwechsel
					}

					For key, PN in parameter {

						TRPat    	:= "<TR" (UDatum	? (trIdf ? " id='td1a'" : " id='td2a'") : (trIdf ? " id='td1b'" : " id='td2b'") ) ">"
						tdPat     	:= "<TD" (Patient  	? (trIdf ? " class='table-btn1'" : " class='table-btn2'") " onclick='ahk.LabJournal_KarteiKarte(event)'>" : ">")
						tbDatum 	:= UDatum ? SubStr(UDatum, 1, 6) : ""
						TDRDX		:= UDatum ? TDR2a : TDR2b
						TDRX      	:= PN.CV = 0 ? TDR 	: (PN.CV = 1 ? StrReplace(TDR 	, "'>"	, cCol1 FWeight1 "'>") : StrReplace(TDR  	, "'>", cCol2 FWeight2 "'>"))
						TDR1X     	:= PN.CV = 0 ? TDR1	: (PN.CV = 1 ? StrReplace(TDR1, "'>"	, cCol1 FWeight1 "'>") : StrReplace(TDR1	, "'>", cCol2 FWeight2 "'>"))
						TDCX      	:= PN.CV = 0 ? TDC 	: (PN.CV = 1 ? StrReplace(TDC	, "'>"	, cCol1 FWeight1 "'>") : StrReplace(TDC 	, "'>", cCol2 FWeight2 "'>"))
						TDL1X     	:= PN.CV = 0 ? TDL1	: (PN.CV = 1 ? StrReplace(TDL1	, "'>"	, cCol1 FWeight1 "'>") : StrReplace(TDL1	, "'>", cCol2 FWeight2 "'>"))

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
								.  	" "	TDRDX	. 	tbDatum                             	"</td>"	; Untersuchungsdatum
								. 	" "	tdPat  	.	Patient             	                  	"</td>"	; Patientenname
								. 	" "	TDRX  	.	PN.PN                                 	"</td>"	; Parameter
								. 	" "	TDRX	.	PN.PW                               	"</td>"	; Wert
								.	" "	TDR1X	.	PN.PV			                     	"</td>"	; Normalwerte
								.	" "	TDL1X 	.	PE                                       	"</td>"	; Einheit
								. 	" "	TDCX	.	(PN.PL	? PN.PL "%" 	: "")     	"</td>"	; +/-
								. 	" "	TDRX	.	(PA   	? PA " " PE  	: "") 		"</td>"	; ⍉ Abw.
								. 	" </TR>`n`n"

						PE := PA := UDatum := Patient := ""

					}

				}

		}

		html .= "</tbody></table></body></html>"

	; erstellte HTML Seite wird angezeigt
		FileOpen(A_Temp "\Laborjournal.html", "w", "UTF-8").Write(html)
		neutron := new NeutronWindow("","","","Laborjournal" LabJ.maxrecords ")", "+AlwaysOnTop minSize900x400")
		neutron.Load(A_Temp "\Laborjournal.html")
		neutron.Gui("+LabelNeutron")
		;neutron.Show("w950 h600")
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

}

LabJournal_KarteiKarte(neutron, event) {

	; event.target will contain the HTML Element that fired the event.
	; Show a message box with its inner text.
		RegExMatch(event.target.innerText, "^(?<Name>.*)\[(?<ID>\d+)\]", Pat)

		If labJ.enrolled ="x" {

			lj := GetWindowSpot(labJ.hwnd)
			dh	:= lj.h
			dtb	:= labJ.TableHeight - labJ.HeaderHeight + 1
			SciTEOutput(dtb " := " labJ.TableHeight " - " labJ.HeaderHeight " + 1")
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

SaveGuiPos(hwnd) {

	win 			:= GetWindowSpot(hwnd)
	winPos	 	:= "x" win.X " y" win.Y " w" win.CW " h" win.CH
	IniWrite, % winpos	, % adm.Ini, % adm.compname, LaborJournal_Position
	SciTEOutput(winpos "," adm.Ini ", " adm.compname )
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


;{ INCLUDES
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




