; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                         	** Addendum_Mining **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Beschreibung:	    	einfachste Funktionen für Datenextraktionen aus Textdateien
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
; 		Inhalt:
;       Abhängigkeiten:		Addendum_Gui.ahk, Addendum.ahk
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_Calc started:    	14.03.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


FindDoc(Text) {


}

FindDocStrings() {

	; ‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿
	;
	;                  REGEX-STRINGS FÜR DAS FINDEN VON PERSONENNAMEN UND DATUMSZAHLEN IN EPIKRISEN, UNTERSUCHUNGSBEFUNDEN, BRIEFEN
	;				   jede Funktion kann durch globales Deklarieren Zugriff nur auf die benötigten Variablen erhalten
	;
	;⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀
	global FindDocStrings_Init := false
	If FindDocStrings_Init
		return
	else
		FindDocStrings_Init := true

	global ThousandYear:= SubStr(A_YYYY, 1, 2)
	global HundredYear	:= SubStr(A_YYYY, 3, 2)

	global excludeIDs  	:= 	"2"

	global rx               	:= 	"[;,:.\s]+"
	global rx1               	:= 	"[;,:.\s]"
	global ry               	:= 	"[,:.;\s\n\r\f]+"
	global rz               	:= 	"[.,;\s]+.*?"
	global rs               	:= 	"\s+"
	global rp1              	:= 	".*"
	global rp2              	:= 	".*?"

	global rxDay         	:= 	"(?<D>[12]*[0-9])"
	global rxWDay       	:= 	"Mon*t*a*g*|Die*n*s*t*a*g*|Mi*t*t*w*o*c*h*|Do*n*n*e*r*s*t*a*g*|Fr*e*i*t*a*g*|Son*n*a*b*e*n*d*|Sam*s*t*a*g*|Son*n*t*a*g*"
	global rxWDay2     	:= 	"(?<WD>" rxWDay ")"
	global rxMonths     	:=	"Jan*u*a*r*|Feb*r*u*a*r*|Mä*a*rz|Apr*i*l*|Mai|Jun*i*|Jul*i*|Aug*u*s*t*|Sept*e*m*b*e*r*|Okt*o*b*e*r*|Nov*e*m*b*e*r*|Deze*m*b*e*r*"
	global rxMonths2   	:=	"(?<Monat>" rxMonths ")"
	global rxYear			:=	"(?<Jahr>([12][0-9])*[0-9]{2})"                                                                  	; Jahr=(19)*85
	global rxYear2			:=	"(?<JH>[12][0-9])*(?<JZ>[0-9]{2})"                                                            	; JH=19*,  JZ=85
	global rxYear3			:=	"(?<Jahr>(\d{4}|\d{2}))"                                                                         	; Jahr= 1985 o. 85

	global rxPerson     	:= 	"[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL]+)*"
	global rxPerson1    	:= 	"(?<Name1>" rxPerson                                                                             	; ?<Name1>Müller-Wachtendónk*
	global rxPerson2    	:= 	"(?<Name2>" rxPerson                                                                              	; ?<Name2>Müller-Wachtendónk*
	global rxPerson3   	:= 	"[A-ZÄÖÜ][\pL]+([\-\s][A-ZÄÖÜ][\pL]+)*"                                                    	; Müller Wachtendonk*
	global rxPerson4    	:= 	"(?<Name>[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL]+)*)"		                                 	; ?<Name>Müller Wachtendonk*

	global rxName1    	:= 	"(?<Name2>" rxPerson	  ")" . rx . "(?<Name1>" 	rxPerson 	")"
	global rxName2    	:= 	"(?<Name2>" rxPerson2 ")" . rx . "(?<Name1>" 	rxPerson2	")"

	global rxDatum	    	:= 	"\d{1,2}" rx "\w+" rx "\d{2,4}"                                                                     	; 12. July 1981 oder 12.7.1981
	global rxDatumLang 	:= 	"\d{1,2}" rx "(" rxMonths ")" rx "\d{2,4}"                                                        	; 12. Jul. 81
	global rxGb1	    	:= 	"(geb\.*o*r*e*n*\s*a*m*|Geb(urts)*" rx "Datu*m*|\*" rx ")"                              	; geb. am o. Geb. Dat. usw
	global rxGb2        	:= 	"(geboren\sam|geb\.\sam|geb\.*o*r*e*n*|Geburtsdatum|\*)"

	global GName      	:= 	"Geb\.\sName|Geburtsname"
	global rxAnrede     	:= 	"(Betrifft)"
	global rxAnredeM   	:= 	"(Herrn*|Patient|Patienten)"
	global rxAnredeF   	:= 	"(Frau|Patientin)"
	global rxVorname  	:= 	"[VY]o(rn|m)ar*me"
	global rxNachname	:= 	"([Nn]ach)*[Nn]ame"
	global rxNameW    	:= 	"[Nn]ame[\w]*"

	global RxNames     	:= [	rxAnrede    	. 	rx	. 	rxName1  .	rx  	.	rxGb1   . 	ry  .  rxDatum        	; 01	Betrifft: Marx, Karl, geb. am: 14.03.1818
							    		,	rxAnrede   	.	rx	. 	rxName2	.	rx  	.	rxGb1   . 	ry                         	; 02
							    		,	rxName1     	.	rx	.	rxGb1   	.	rs                          	.   rxDatum          	; 03
							    		,	rxName2     	.	rx	.	rxGb1     	.	rs                             	.   rxDatum          	; 04
							    		,	rxName1     	.	rx	.	rxGName	.	rp2	.	rxGb1 	.	rx	.	rxDatum        	; 05
							    		,	rxName2     	.	rx	.	rxGName	.	rz  	.	rxGb1 	.	rx	.	rxDatum        	; 06	Müller-Wachtendonk, Marie-Luise, Geburtsname: Müller, geb. am 12.11.1964

							    		,	rxAnrede   	. 	rx  .  rxName2                                                                 	; 07	Betrifft: Müller-Wachtendonk, Marie-Luise
							    		,	rxAnredeM 	. 	rx  .  rxName2                                                               	; 08	Herr Karl Marx o. Patient Marx, Karl
							    		,	rxAnredeF 	. 	rx  .  rxName2                                                                 	; 09	Frau Marie-Luise Müller-Wachtendonk
							    		,	rxAnredeM 	. 	rx  .  rxName2  .	rx 	.	rxGb2  	.	rx	.	rxDatum       	; 10	Patient Marx, Karl, * 14.03.1818
							    		,	rxAnredeF 	. 	rx  .  rxName2  .	rx 	.	rxGb2  	.	rx	.	rxDatum       	; 11	Patientin Luxemburg, Rosa, * 05.03.1871

							    		,	rxNameW		.	rx	.	rxName1                                                                 	; 12
							    		,	rxNachname	. 	rx	. 	rxPerson2	.	".*?" 	.	rxVorname	. rx . rxPerson1     	; 13	Nachname: Karl, Versichtennummer: 12345678, Vorname: Marx
							    		,	rxVorname	. 	rx	. 	rxPerson2	.	".*?"	.	rxNachname	. rx . rxPerson2     	; 14	Vorname: Marx, Versichtennummer: 12345678, Nachname: Karl

							    		,	"^\s*" rxName2 "\s"]	                                                                              	; 15	^ Muster, Max  ....

	global RxNames2  	:= [	rxNachname	. 	rx	. 	rxPerson2                            										; 01 	Nachname: Müller(-Wachtendonk)*
								    	,	rxVorname	. 	rx	.	rxPerson1]                          										; 02	Vorname: Marie(-Luise)*

	global	rxDates       	:= [	"(?<Tag>\d{1,2})[.,;](?<Monat>\d{1,2})[.,;]" rxYear3 "([^\d]|$)"           	; 01	12.11.64 o. 12.11.1964 o. 2.1.(19)*64
								    	,	"i)(?<Tag>\d{1,2})" rx . rxMonths2 . rx . rxYear3 "([^\d]|$)"                  	; 02	12. November (19)64
								    	,	"(" rxGb1 ")"	rx  ".{0,20}?" . rxDay . rx . rxMonths2 . rx . rxYear]              	; 03   Geburtsdatum: .*(20 beliebige Zeichen) 12. Nov. 1964


	global rxContinue  	:= "(Nachname|Vor[nm]ame|Geburtsdatum)"

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; REGEXSTRINGS FÜR DOKUMENTDATUM (FindDocDate)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	global rxTags	       	:= {	1:	"Druckdatum|Erstellungszeitpunkt|Dokumentdatum|D[ae]tum|Beginn|ausgedruckt|"
											. 	"Probenentnahmedatu*m*|Abnahmedatu*m*|Berichtsdatu*m*|Eingangsdatu*m*|"
											. 	"Eingang\s+am|gedruckt\s+am|angelegt|angelegt\s+am|"
											.	"Anfrage\s+vom|Arztbrief\s+vom|Befund\s+vom|Befundbericht\s+vom|"
											. 	"Behandlung\s+vom|Ebenen\s+vom|Ebenen\/axial\s+vom|"
											.	"Echokardiografie vom|Konsil\s+vom|Labor\s+vom|Laborblatt\s+vom|Nachricht\s+vom|"
											. 	"Startzeit|Aufgezeichnet|Erstellt\s+am|Eing[.,\-\s]+Dat[.,-]+|Ausg[.,\-\s]+Dat[.,-]+"
											.	"Labor(blatt)*\svom"
										,	2:	"Behandlung|haben wir.*|sich"}

	global rxBehandlung	:= [	"i)(" rxTags[2] ")\svom\s(?<Datum1>[\d.]+)\s*(bis\s*z*u*m*)\s*(?<Datum2>[\d.]+)"                	; 	1| (haben wir ...| sich) vom .... bis (zum) .....
										, 	"i)(" rxTags[2] ")\svom\s(?<Datum1>[\d.]+)\s*(\-)\s*(?<Datum2>[\d.]+)"]                                 	; 	2| (haben wir ...| sich) vom ..... - .......

	global rxDokDatum	:=[	"i)\s*(" rxTags[1] ")" rx "(" rxWDay ")*[.,;\s]+(?<Datum1>\d+\.\d+\.\d+)"                                  	;   1| Erstellungszeitpunkt: Do, 02.01.2020
							    		,	"i)\s*(" rxTags[1] ")" rx "(" rxWDay ")*[.,;\s]+(?<Datum1>" rxDatumLang ")"                                  	;   2| Erstellungszeitpunkt: Do. 02. Januar 2020
										, 	"i)(gedruckt|angelegt|sich)\s*am\s*[;:]*\s*(?<Datum1>\d+\.\d+\.\d+)"                                   	;   3| gedruckt am: 02.01.2020
										,	"i)(gedruckt|angelegt)\s*[;:]*\s*(?<Datum1>\d+\.\d+\.\d+)"                                                	;	4| angelegt: 02.01.2020
							    		,	"i)^[\pL\-]+\s[\pL\-]+\s*[;,.]\s*(den)*\s+(?<Datum1>\d+\.\d+\.\d+)\s*$"                              	;   5| Hamburg-Hochburg, 02.01.2020 (nur Leerzeichen folgen)
							    		,	"i)^[\pL\-]+\s[\pL\-]+\s*[;,.]\s*(den)*\s+(?<Datum1>" rxDatumLang ")\s*$"                             	;   6| Hamburg-Hochburg, (den) 02. Januar 2020 (nur Leerzeichen folgen)
							    		,	"^s*[\pL\s\(\)]+\s*[,;.]\s*(den\s)*(?<Datum1>\d{1,2}\.\d{1,2}\.\d{2,4})"                            	;   7| Hamburg, (den) 02.01.2020
							    		,	"^s*[\pL\s\(\)]+\s*[,;.]\s*(den\s)*(?<Datum1>" rxDatumLang ")\s*"                                           	;   8| Hamburg, (den) 2. Januar 2020
							    		,	"^\s*(?<Datum1>\d{2}\.\d{2}\.(\d{2}|\d{4}))\s*$"                                                            	;   9| 02.01.2020 (alleinstehend in Zeile)
							    		,	"^\s*(?<Datum1>" rxDatumLang ")\s*$"]                                                                                 	; 10| 02. Jan(.|uar) 2020 (alleinstehend in Zeile)

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; ZUM VALIDIEREN EINES DATUMS
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	global rxDateOCRrpl		:= "[,;:]"
	global rxDateValidator	:= [ 	"^(?<D>[0]?[1-9]|[1|2][0-9]|[3][0|1]).(?<M>[0]?[1-9]|[1][0-2]).(?<Y>[0-9]{4}|[0-9]{2})$"
											,	"^(?<D>[0]?[1-9]|[1|2][0-9]|[3][0|1])[.;,\s]+(?<M>" rxMonths ")[.;,\s]+(?<Y>[0-9]{4}|[0-9]{2})$"]

}

FindDocNames(Text, debug:=false)                                  	{               	;-- sucht Namen von Patienten im Dokument

	; letzte Änderung: 26.02.2021
		debug := true

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; RegEx-Strings zusammenstellen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
		global FindDocStrings_Init

		If !FindDocStrings_Init
			FindDocStrings()

		; ----------------------------- Ausschluß von bestimmten Patientennummern (PatID)
			global excludeIDs
		; ----------------------------- Worttrennzeichen
			global rx, ry, rz, rs, rp1, rp2
		; ----------------------------- Datum
			global ThousandYear, HundredYear, rxDay, rxWDay, rxWDay2, rxMonths, rxMonths2, rxYear 	; Teile des Datum
			global rxDatum, rxDates                                                                                                         	; Komplett
			global rxGb1, rxGb2                                                                                                    	; Startstrings
		; ----------------------------- Eigen- oder Personennamen
			global rxPerson, rxPerson1, rxPerson2, rxPerson3, rxPerson4, rxName1, rxName2                      	; Teile
			global RxNames, RxNames2                                                                                                  	; Komplett
			global GName, rxAnrede, rxAnredeM, rxAnredeF, rxVorname, rxNachname, rxNameW				; Startstrings
		; ----------------------------- sonstige
			global rxContinue

		; auf Festsplatte schreiben zur Fehlerkorrektur
			If (debug > 1) {
				f := FileOpen(A_ScriptDir "\admMiningRegEx.txt", "w", "UTF-8")
				For i, str in RxNames
					f.WriteLine("[-" SubStr("00" i, -1) " -] " str)
				For i, str in RxDates
					f.WriteLine("[-" SubStr("00" i, -1) " -] " str)
				f.Close()
			}
	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	  PatBuf := Object()
	; Methode 1:	RegEx Strings nutzen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
		If debug
		  SciTEOutput("`n  --------------------------------------`n {Methode 1}")

	; - benutzt RegEx-Strings aus RxNames um Patientennamen zu finden
		For RxStrIndex, rxNameStr in RxNames {

			spos	:= 1
			while (spos := RegExMatch(text, rxNameStr, Pat, spos)) {

				; end of line Zeichen entfernen
					spos  += StrLen(Pat)
					PatName2 := RegExReplace(PatName2, "^.*[\n\r]+", ""), Pat := RegExReplace(Pat, "^.*[\n\r]+", "")

					If (debug = 2)
						SciTEOutput("    [" RxStrIndex " - " rxNameStr "]`n         " PatName2 ", " PatName1)
					else if (debug = 1)
						SciTEOutput("    [" RxStrIndex "]  " PatName2 ", " PatName1  " / " spos " (" Pat ")")

				; Unpassendes entfernen
					PatName1 := StrReplace(PatName1, " "), PatName2 := StrReplace(PatName2, " ")
					If RegExMatch(PatName1, rxContinue) || RegExMatch(PatName2, rxContinue)
						continue

				  ; schaut ob der gefundene Name in der Datenbank ist
					If RegExMatch(rxNameStr, rxAnredeM)
						sex := "m"
					else if RegExMatch(rxNameStr, rxAnredeF)
						sex := "w"
					matches 	:= admDB.StringSimilarityEx(PatName1, PatName2, 0.11)
					mdiff     	:= matches.Delete("diff")
					If IsObject(matches)
						For PatID, Pat in matches {

							; diese ID's ignorieren
								If PatID in %excludeIDs%
									continue

							; zählen oder neue ID hinzufügen
								If !PatBuf.haskey(PatID) {

									PatBuf[PatID] := {"Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd, "hits":1, "method":1}
									If (StrLen(sex) > 0) && (Pat.Gt = sex)
										PatBuf[PatID].hits := 2
									If debug
										SciTEOutput("   ++" PatID " , hits: " PatBuf[PatID].hits ", Name: " PatBuf[PatID].Nn ", "PatBuf[PatID].Vn ", " PatBuf[PatID].Gd)

								}
								else {

									If (StrLen(sex) > 0) && (Pat.Gt = sex)
										PatBuf[PatID].hits += 2
									else
										PatBuf[PatID].hits += 1

									If debug
										SciTEOutput("    +" PatID " , hits: " PatBuf[PatID].hits ", Name: " PatBuf[PatID].Nn ", "PatBuf[PatID].Vn ", " PatBuf[PatID].Gd)
								}

						}

			}

		}

	 ; debugging
		If debug
			SciTEOutput("    gefundene ID's: " PatBuf.Count())

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Methode 1a: zähle die im Text vorhandenen Geburtstage der Patienten (PatBuf) und entferne diese
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		If debug && (PatBuf.Count() > 0)
			SciTEOutput(" {Methode 1a}")

		GdCountSum := 0
		For PatID, Pat in PatBuf {

			rxGeburtsdatum := StrReplace(Pat.Gd, ".", "[.,;]")
			text := RegExReplace(text, rxGeburtsdatum, "", GdCount)
			PatBuf[PatID].hits += (GdCount * 2)
			GdCountSum += GdCount

			If debug
				SciTEOutput("   " PatID " , hits: " PatBuf[PatID].hits ", Datum: " Pat.Gd " +" GdCount " hits")
		}
	  ;}

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Methode 1b: mehrere Patienten gefunden. Welche Patientennamen findest Du im Text
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (PatBuf.MaxIndex() > 1) {

			If debug
			  SciTEOutput(" {Methode 1b}")

			; alle Worte mit einem Großbuchstaben am Anfang finden
			spos := 1, PNames := Array()
			while (spos := RegExMatch(text, "\s" rxPerson4, P, spos)) {
				PNames.Push(PName)
				spos += StrLen(PName)
			}

			; Vor- und Nachnamen mit jedem Wort vergleichen
			For PatID, Pat in PatBuf {
				thisPat := 0
				For pidx, word in PNames {
					DiffA	:= StrDiff(word, Pat.Nn), DiffB := StrDiff(word, Pat.Vn)
					If (DiffA <= diffmin) || (DiffB <= diffmin)
						thisPat += 1
				}
				If (thisPat >= 2) {
					PatBuf[PatID].hits += 1
					If debug
						SciTEOutput("   " PatID " , hits: " PatBuf[PatID].hits)
				}
			}

		}

	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Methode 2:	Suche nach Patientennamen über gefundene Geburtstage
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
		If debug
		  SciTEOutput(" {Methode 2}")
		spos := 1
		For rxDateIdx, rxDateString in rxDates {

			spos := 1
			while (spos := RegExMatch(text, rxDateString, D, spos)) {

				spos += StrLen(D)

			  ; Monat wurde als Wort geschrieben, dann die Nummer des Monats ermitteln
				If !RegExMatch(DMonat, "\d{1,2}")
					For nr, rxMonth in StrSplit(rxMonths, "|")
						If RegExMatch(DMonat, "i)" rxMonth) {
							DMonat := SubStr("00" nr, -1)
							break
						}

			  ; ein Datumsformat wie 01.08.24, wird geändert auf 01.08.1924
			  ; (ist die Jahreszahl (im Beispiel 24) größer als die aktuelle Jahrzahl dann wird daraus 1924,
			  ;   ist diese kleiner wird das Datum verworfen, weil das Datum in der Zukunft liegt)
				If (StrLen(DJahr) = 2) {
					If (DJahr < HundredYear)
						continue
					DJahr := SubStr("00" (ThousandYear - 1), -1) . DJahr
				}

			  ; sucht nach Patienten mit passenden Geburtstagen
				Datum := SubStr("00" DTag, -1) "." SubStr("00" DMonat, -1) "." DJahr
				If debug
					SciTEOutput("    [" rxDateIdx "]  " DTag "." DMonat "." DJahr  "  (" Datum ")")

				matches := admDB.MatchID("gd", Datum), pidx := 1
				If IsObject(matches)
					For PatID, Pat in matches {

					; bestimmte ID's ausschließen
						If PatID in %excludeIDs%
							continue

					; PatID ist bekannt, Zähler erhöhen
						If PatBuf.haskey(PatID) {
							; bekommt zwei Treffer-Punkte wenn zuvor die PatID mit Methode1 gefunden wurde
								If (PatBuf[PatID].method = 1)
									PatBuf[PatID].hits += 2
								else
									PatBuf[PatID].hits += 1
						}
						else { ; PatID ist unbekannt
							PatBuf[PatID] := {"Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd, "hits":1, "method":2}
							If InStr(text, Pat.Nn) && InStr(text, Pat.Vn) ; + weitere 2 wenn der Name im Dokument auftaucht
								PatBuf[PatID].hits += 2
						}

						If debug {
							SciTEOutput("   Patient " (pidx++) ":`t" PatID " , hits: " PatBuf[PatID].hits ", Datum: " RegExReplace(D, "[\n\r]"))
						}

					}
			}
		}

	 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Methode 2a: mehrere Patienten gefunden. Welche Patientennamen findest Du im Text
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (PatBuf.MaxIndex() > 1) {

			If debug
			  SciTEOutput(" {Methode 2a}")

			; alle Worte mit einem Großbuchstaben am Anfang finden
			spos := 1, PNames := Array()
			while (spos := RegExMatch(text, "\s" rxPerson4, P, spos)) {
				PNames.Push(PName)
				spos += StrLen(PName)
			}

			; Vor- und Nachnamen mit jedem Wort vergleichen
			For PatID, Pat in PatBuf {
				thisPat := 0
				For pidx, word in PNames {
					DiffA	:= StrDiff(word, Pat.Nn), DiffB := StrDiff(word, Pat.Vn)
					;SciTEOutput("      word: " word  " A: " DiffA)
					If (DiffA <= 0.21) || (DiffB <= 0.21) {
						thisPat += 1
						SciTEOutput("   PatID:    `t" PatID " word: " word " A:" DiffA " B: " DiffB)
					}
				}
				If (thisPat >= 2) {
					PatBuf[PatID].hits += 1
					If debug
						SciTEOutput("   " PatID " , hits: " PatBuf[PatID].hits)
				}
			}
		}
	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Methode 3:	kommt nur zum Einsatz wenn die Methoden 1 und 2 nichts gefunden haben.
	;                  	Findet zwei aufeinander folgende Worte mit Großbuchstaben am Anfang (Personennamen, Eigennamen ...)
	;               	und vergleicht diese per StringSimiliarity-Algorithmus mit den Namen in der Addendum-Patientendatenbank
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
		If debug
		  SciTEOutput(" {Methode 3}")
		ExclBuf := ""
		If (PatBuf.Count() = 0) {

			spos := 1
			while (spos := RegExMatch(text, "([A-ZÄÖÜ][\pL-]+)[\s,.;:]+([A-ZÄÖÜ][\pL-]+)", name, spos)) {

				; überflüssige Zeichen entfernen
					name1 := RegExReplace(name1, "[\s\n\r\f]")
					name2 := RegExReplace(name2, "[\s\n\r\f]")

				; Bindestrichworte ignorieren
					If RegExMatch(name1, "\-$") {
						spos += StrLen(name1)                    ; ein Wort weiter
						continue
					} else if RegExMatch(name2, "\-$") {
						spos += StrLen(name)                     ; beide Wörter weiter
						continue
					} else if (StrLen(name1) = 0) || (StrLen(name2) = 0) {
						spos += StrLen(name)
						continue
					}

					matches := admDB.StringSimilarityEx(name1, name2)
					If IsObject(matches) {

						spos += StrLen(name)

						For PatID, Pat in matches {

							If PatID in %excludeIDs%
								continue

							If PatBuf.haskey(PatID)
								PatBuf[PatID].hits += 1
							else
								PatBuf[PatID]	:= {"Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd, "hits":1, "method":1}

							If debug
								SciTEOutput("   " PatID " , hits: " PatBuf[PatID].hits)
						}

					}
					else
						spos += StrLen(name1)

			}

	}

		; ---------------------------------------------------------------------------------------------------------------------------------------------------
		; Methode 3a: Es wurden Patienten gefunden. Sieht nach welche Namen tatsächlich im Text vorkommen (untersucht Wort für Wort)
		; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
			If (PatBuf.Count() > 9999) {    ; nicht fertig

			; sammelt alle Wort entsprechend rxPerson in ein Object()
				propernames := Object()
				spos	:= 1
				while (spos := RegExMatch(text, rxPerson, pname, spos)) {

					spos += StrLen(pname)
					If !propernames.HasKey(pname)
						propernames[pname] := 1
					else
						propernames[pname] += 1

				}

			;
				For PatID, Pat in PatBuf {

					;n1 := StrDiff(Text)

				}

		}
		;}

	;}

	; höchste Trefferzahl ermitteln ;{
		maxHits := 0, BestHits := Object()
		For PatID, Pat in PatBuf {
			If (PatID = "Diff")
				continue
			maxHits := Pat.hits > maxHits ? Pat.hits : maxHits
			If debug
				SciTEOutput("   hitpoints:`t" Pat.hits " [(" PatID ") " Pat.Nn ", " Pat.Vn "]")
		}
	;}

	; alle Namen mit gleicher Trefferanzahl werden behalten ;{
		For PatID, Pat in PatBuf
			If (PatID = "Diff")
				continue
			else If (Pat.hits = maxHits) {
				;SciTEOutput(" (" PatID ") " Pat.Nn ", " Pat.Vn ", hits: " Pat.hits)
				BestHits[PatID] := {"Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd}
			}
	;}

return BestHits
}

FindDocDate(Text, names="", debug=false) 						{                	;-- Behandlungstage und/oder Erstellungsdatum des Dokuments

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; RegEx-Strings zusammenstellen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
		global FindDocStrings_Init
		If !FindDocStrings_Init
			FindDocStrings()

		global excludeIDs								; Ausschluß von bestimmten Patientennummern (PatID)
		global rxBehandlung, rxDokDatum  	; Datum

	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Objekte
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
		DocDates := Object()
		DocDates.Behandlung 	:= Object()
		DocDates.Dokument  	:= Object()

		retDates := Object()
		retDates.Behandlung := Array()
		retDates.Dokument  	:= Array()
	;}

	; Text in Zeilen aufspalten
		TLines := StrSplit(Text, "`n", "`r")

	; Behandlung von ... bis ....
		For RxStrIndex, rxDateStr in rxBehandlung {
			For LNr, line in TLines {

				If !RegExMatch(line, rxDateStr, D)
					continue

				DDatum1 := DateValidator(DDatum1, A_YYYY)
				DDatum2 := DateValidator(DDatum2, A_YYYY)

			; prüft ob die Daten Geburtstage sind
				For PatID, Pat in names {
					If (Pat.Gd = Datum1)
						Datum1 := ""
					If (Pat.Gd = Datum2)
						Datum2 := ""
					If (StrLen(Datum1 Datum2) = 0)
						break
				}
				If (StrLen(Datum1 Datum2) = 0)
					continue
				If (StrLen(Datum1) = 0) && (StrLen(Datum2) > 0)
					Datum1 := Datum2, Datum2 := ""

			; Datumstring(s) speichern
				saveDate := DDatum1 (DDatum2 ? "-" DDatum2 : "")
				SciTEOutput(" B2: " saveDate)
				If !DocDates.Behandlung.HasKey(saveDate)
					DocDates.Behandlung[saveDate] := {"fLine":[LNr], "dcount":1}
				else {
					For DDIdx, fLineNr in DocDates.Behandlung[saveDate].fline
						If (fLineNr = LNr)
							continue
					savedLNr := LNr
					DocDates.Behandlung[saveDate].fLine.Push(LNr)
					DocDates.Behandlung[saveDate].dcount += 1
				}
			}
		}

	; Erstellungsdatum des Dokumentes
		For RxStrIndex, rxDateStr in rxDokDatum {
			For LNr, line in TLines {

				If !RegExMatch(line, rxDateStr, D)
					continue

				DDatum1 := DateValidator(DDatum1, A_YYYY)
				DDatum2 := DateValidator(DDatum2, A_YYYY)
				SciTEOutput(" D: " DDatum1 ", " DDatum2)

			; prüft ob die Daten Geburtstage sind
				For PatID, Pat in names {
					If (Pat.Gd = Datum1)
						Datum1 := ""
					If (Pat.Gd = Datum2)
						Datum2 := ""
					If (StrLen(Datum1 Datum2) = 0)
						break
				}
				If (StrLen(Datum1 Datum2) = 0)
					continue
				If (StrLen(Datum1) = 0) && (StrLen(Datum2) > 0)
					Datum1 := Datum2, Datum2 := ""

			; Datumstring(s) speichern
				saveDate := DDatum1 (DDatum2 ? "-" DDatum2 : "")
				SciTEOutput(" D: " saveDate)
				If !DocDates.Dokument.HasKey(saveDate) {
					DocDates.Dokument[saveDate] := {"fLine":[LNr], "dcount":1}
				}
				else {
					For DDIdx, fLineNr in DocDates.Dokument[saveDate].fline
						If (fLineNr = LNr)
							continue
					savedLNr := LNr
					DocDates.Dokument[saveDate].fLine.Push(LNr)
					DocDates.Dokument[saveDate].dcount += 1
				}
			}
		}

	; es wurde kein Behandlungs- noch Erstellungsdatum gefunden
		If (DocDates.Behandlung.Count() = 0) && (DocDates.Dokument.Count() = 0) {

			DocPages	:= FindDocPages(Text)
			PageNr 	:= 1
			For LNr, line in TLines {

				; eine Seite weiter
					If RegExMatch(line, "[\f]") {
						PageNr ++
						continue
					}

				; einzelnes Datum wird auf seine Position innerhalb des Dokumentes geprüft
					If RegExMatch(line, "(?<Datum1>\d{2}\.\d{2}\.\d{2,4})", D) {

						; Datum befindet sich in der "Kopf- oder Fußzeile"
							OberesViertel := Floor(DocPages[PageNr]/5)
							UnteresViertel := Floor(DocPages[PageNr] - DocPages[PageNr]/5)
							If (LNr <= OberesViertel) || (LNr >= UnteresViertel) {

								saveDate := DDatum1 (DDatum2 ? "-" DDatum2 : "")
								If !DocDates.Dokument.HasKey(saveDate) {
									DocDates.Dokument[saveDate] := {"fLine":[LNr], "dcount":1}
								}
								else {
									For DDIdx, fLineNr in DocDates.Dokument[saveDate].fline
										If (fLineNr = LNr)
											continue
									savedLNr := LNr
									DocDates.Dokument[saveDate].fLine.Push(LNr)
									DocDates.Dokument[saveDate].dcount += 1
								}

							}
					}
			}
		}

	; Rückgabe Objekt: alle Daten unter maximaler Häufigkeit aussortieren [Dokument] ;{
		maxdcount := 0
		For KeyDate, docdate in DocDates.Dokument
			maxdcount := docdate.dcount > maxdcount ? docdate.dcount : maxdcount

		For KeyDate, docdate in DocDates.Dokument
			If (docdate.dcount = maxdcount)
				retDates.Dokument.Push(KeyDate)

		If debug && (DocDates.Dokument.Count() <> retDates.Dokument.Count())
			SciTEOutput("  " (DocDates.Dokument.Count() - retDates.Dokument.Count()) " Erstellungsdaten wurde(n) aussortiert!")

	;}

	; Rückgabe Objekt: alle Daten unter maximaler Häufigkeit aussortieren [Dokument] ;{
		maxdcount := 0
		For KeyDate, docdate in DocDates.Behandlung
			maxdcount := docdate.dcount > maxdcount ? docdate.dcount : maxdcount

		For KeyDate, docdate in DocDates.Behandlung
			If (docdate.dcount = maxdcount)
				retDates.Behandlung.Push(KeyDate)

		If debug && (DocDates.Behandlung.Count() <> retDates.Behandlung.Count())
			SciTEOutput("  " (DocDates.Behandlung.Count() - retDates.Behandlung.Count()) " Behandlungstage wurde(n) aussortiert!")

	;}

return retDates
}

FindDocDate_GetMaxHits(datesObj)                              	{

	retArr := Array()

	maxdcount := 0
	For didx, docdate in datesObj
		maxdcount := docdate.dcount > maxdcount ? docdate.dcount : maxdcount

	For KeyDate, docdate in datesObj
		If (docdate.dcount = maxdcount)
			retArr.Push(KeyDate)

return retArr
}

FindDocSender(Text) 															{

	; Dr.Kardiotoid, medilog, Darwin, professional. 		    	- LZ-EKG
	; Arztanfrage, (Muster 52), 											- Muster 52

}

FindDocPages(Text) 															{ 	         		;-- ermittelt die Zeilenzahl jeder Seite

	; DocPage Array ist erstens der Seitenzähler
	; zu jeder Seite wird die Zeilenzahl als Wert übergeben
	; durch

	; Variablen
		DocPage := Array()
		LastPageBreakLineNr := 0

	; Text in Zeilen splitten
		TLines := StrSplit(Text, "`n", "`r")

	; es wird nach Pagebreaks (FF) gesucht und diese gezählt
		For LNr, line in TLines
			If RegExMatch(line, "[\f]") {
				DocPage.Push(LNr - LastPageBreakLineNr)
				LastPageBreakLineNr := LNr
			}

	; ohne Pagebreak ist es eine Seite
		If (DocPage.MaxIndex() = 0)
			DocPage.Push(LNr)

return DocPage
}

FindName_BestHit(PatBuf) 													{

	; höchste Trefferzahl ermitteln
		maxHits := 0, equal := false, BestHits := Object()
		For PatID, Pat in PatBuf {
			;SciTEOutput("  hitpoints:   `t`t" Pat.hits " - (" PatID ") " Pat.Nn ", " Pat.Vn)
			If (Pat.hits > maxHits)
				maxHits := Pat.hits
		}

	; alle Namen mit gleicher Trefferanzahl werden behalten
		For PatID, Pat in PatBuf
			If (Pat.hits = maxHits)
				BestHits[PatID] := {"Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd}

return BestHits
}

FindNameNewWay(txt) 														{

	WBuf := Array()
	txt := RegExReplace(txt, "[\n\r]"   	, " ")
	txt := RegExReplace(txt, "[\s]{2,}"	, " ")
	;~ stext := StrSplit(text, " ")
	;~ For FNIndex, word in sText {


	;~ }
	dates := GetTextDates(txt)


}

FindDocOther(Text)                                                       	{                	;-- Kategorisierung von nicht Patientenbriefen

	; Brief key-Parameter:
	;		Stichworte 	- Wörter die den Brief definieren
	;		Stichzahl    	- wieviele davon sind notwendig damit ein Treffer angerechnet wird
	;		rxFilename	- RegExStrings um Details für den Dateinamen zu finden
	Brief := {"Fortbildung"	:	{	"Stichworte"	: ["Einladung|Anmeldung|neues zur|Programm|LÄK", "Fortbildung|Online-Fortbildung"]
											,	"Stichzahl" 	: 2
											,	"rxFileName"	: [""]}
				, "Rechnung"		:	{	"Stichworte"	: ["Rechnung vom", "Zahlungsaufforderung", "Verwendungszweck", "Rechnung", "Gesamtbetrag"]
											,	"Stichzahl" 	: 2
											,	"rxFileName"	: [""]}}

}

GetTextDates(txt) 																{

	static rxDate  	:= "(?<T>\d{1,2})[\.,](?<M>\d{1,2})[\.,](?<J>\d{1,4})"
	static rxMonths 	:= "\d{1,2}\.\s*(Janu*a*r*|Febr*u*a*r*|Mä*a*rz|Apri*l*|Mai|Juni*|Juli*|Augu*s*t*|Septe*m*b*e*r*|Okto*b*e*r*|Nove*m*b*e*r*|Deze*m*b*e*r*)\s\d{2,4}"

	dates := Array()
	pos := 1
	txt := RegExReplace(txt, "[\n\r]", " ")
	words := StrSplit(txt, " ")

	For wIdx, word in words {

		If RegExMatch(word, rxDate, f) {

			datefound := false
			wsbefore 	:= ""
			datum    	:= SubStr("00" fT, -1) "." SubStr("00" fM, -1) "." fJ

			Loop 2 {
				wbIdx := wIdx - 3 + A_Index
				If (wbIdx > 0) {
					wclean := RegExReplace(words[wbIdx], "i)[^\p{L}]+$")
					wclean := RegExReplace(wclean, "i)^[^\p{L}]+")
				    wsbefore .= wclean (A_Index = 1 ? " " : "")  ; entfernt alle Sonderzeichen am Ende der Wörter
				}
			}

			For DateIdx, obj in dates
				If (obj.Datum = Datum) {
					datefound := true
					dates[DateIdx].Counter += 1
					If (StrLen(wsbefore) > 0) && !InStr(obj.wordsbefore, wsbefore)
						dates[DateIdx].wordsbefore .= "|" wsbefore
					break
				}

			If !dateFound
				dates.Push({"Datum":Datum, "Counter":1, "wordsbefore":wsbefore})

		}

	}

return dates
}


