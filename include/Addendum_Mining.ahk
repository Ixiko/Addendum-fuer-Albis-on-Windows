; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                         	** Addendum_Mining **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Beschreibung:	    	einfachste Funktionen für Datenextraktionen aus Textdateien
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
; 		Inhalt:
;       Abhängigkeiten:		Addendum_Gui.ahk, Addendum.ahk
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_Calc started:    	14.02.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


FindDoc(Text) {


}

FindDocNames(Text, debug:=false)                                  	{               	;-- sucht Namen von Patienten im Dokument

	; letzte Änderung: 10.02.2021
	; ‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿
	;
	;                  REGEX-STRINGS FÜR DAS FINDEN VON PERSONENNAMEN UND DATUMSZAHLEN IN EPIKRISEN, UNTERSUCHUNGSBEFUNDEN, BRIEFEN
	;
	;{⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀

	;debug := true

	; REGEXSTRINGS

	; 	(geboren\sam|geb\.\sam|geb\.|geboren|Geburtsdatum)[;,:.\s]+
	;	. (?<WDay>Mon*t*a*g*|Die*n*s*t*a*g*|Mi*t*t*w*o*c*h*|Do*n*n*e*r*s*t*a*g*|Fr*e*i*t*a*g*|Son*n*a*b*e*n*d*|Sa*m*s*t*a*g*|So*n*n*t*a*g*)[;,:.\s]+
	;	. (?<D>[\s12]*[0-9])[;,:.\s]+(?<M>Jan*u*a*r*|Feb*r*u*a*r*|Mä*a*rz|Apr*i*l*|Mai|Jun*i*|Jul*i*|Augu*s*t*|Septe*m*b*e*r*|Okto*b*e*r*|Nove*m*b*e*r*|Deze*m*b*e*r*)
	; 	. [;,:.\s]+(?<Y>([12][0-9])*[0-9]{2})
	;
	; rxGebAm . rx1 . rxWDay . rx1 . (rxDay) . rx1 . (rxMonths) . rx1 . (rxYear)


		static ThousandYear	:= SubStr(A_YYYY, 1, 2)
		static HundredYear	:= SubStr(A_YYYY, 3, 2)

		static excludeIDs  	:= 	"2"

		static rx              	:= 	"[;,:.\s]+"
		static ry              	:= 	"[,:.;\s\n\r\f]+"
		static rz              	:= 	"[.,;\s]+.*?"
		static rs              	:= 	"\s+"
		static rp1              	:= 	".*"
		static rp2              	:= 	".*?"

		static rxDay        	:= 	"(?<D>[12]*[0-9])"
		static rxWDay       	:= 	"Mon*t*a*g*|Die*n*s*t*a*g*|Mi*t*t*w*o*c*h*|Do*n*n*e*r*s*t*a*g*|Fr*e*i*t*a*g*|Son*n*a*b*e*n*d*|Sam*s*t*a*g*|Son*n*t*a*g*"
		static rxWDay2     	:= 	"(?<WD>Mon*t*a*g*|Die*n*s*t*a*g*|Mi*t*t*w*o*c*h*|Do*n*n*e*r*s*t*a*g*|Fr*e*i*t*a*g*|Son*n*a*b*e*n*d*|Sam*s*t*a*g*|Son*n*t*a*g*)"
		static rxMonths     	:=	"Janu*a*r*|Febr*u*a*r*|Mä*a*rz\.*|Apri*l*|Mai|Juni*|Juli*|Augu*s*t*|Septe*m*b*e*r*|Okto*b*e*r*|Nove*m*b*e*r*|Deze*m*b*e*r*"
		static rxMonths2   	:=	"(?<Monat>Janu*a*r*|Febr*u*a*r*|Mä*a*rz\.*|Apri*l*|Mai|Juni*|Juli*|Augu*s*t*|Septe*m*b*e*r*|Okto*b*e*r*|Nove*m*b*e*r*|Deze*m*b*e*r*)"
		static rxYears			:=	"(?<Jahr>([12][0-9])*[0-9]{2})"

		static rxPerson      	:= 	"[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL]+)*"
		static rxPerson2   	:= 	"[A-ZÄÖÜ][\pL]+([\-\s][A-ZÄÖÜ][\pL]+)*"

		static rxName1	:= 	"(?<Name2>" rxPerson 	")\s*[.,;]+\s*(?<Name1>"	rxPerson	")"
		static rxName2	:= 	"(?<Name2>" rxPerson 	")" . rx . "(?<Name1>"    	rxPerson 	")"
		static rxName3	:= 	"(?<Name2>" rxPerson2 	")" . rx . "(?<Name1>"    	rxPerson2	")"

		static rxDatum		:= 	"\d{1,2}[.,;\s]+\w+[.,\s]+\d{2,4}"
		static rxGebAm		:= 	"geboren\sam|geb\.\sam|geb\.|geboren|Geburtsdatum|Geb[.,\-\s]+Dat[,.]+"
		static rxGebAm2	:= 	"(geboren\sam|geb\.\sam|geb\.|geboren|Geburtsdatum)"
		static GName    	:= 	"Geb\.\sName|Geburtsname"
		static rxAnrede   	:= 	"Betrifft"
		static rxAnredeM   	:= 	"Herrn*|Patient|Patienten"
		static rxAnredeF  	:= 	"Frau|Patientin"

		static RxNames     	:= [	"(" rxAnrede ")" 	rx   .  	rxName1  .		rx  	"(" rxGebAm ")"   . 	ry  .  rxDatum                                	; 01
										,	"(" rxAnrede ")"	rx   .   	rxName2	.		rx  	"("	rxGebAm ")"   . 	ry                                                 	; 02
										,	"(" rxAnrede ")"	rx   .   	rxName3	.		rx  	"("	rxGebAm ")"   . 	ry                                                 	; 03
										,	  	rxName1 	.	rx   	"("	rxGebAm	")"		rs  	.   rxDatum                                                                  	; 04
										,		rxName2 	.	rx   	"("	rxGebAm 	")"		rs  	.   rxDatum                                                                 	; 05
										,		rxName1 	.	rx   	"("	rxGName	")"		rp2	"(" rxGebAm ")"	.	rx	.	rxDatum                                	; 06
										,		rxName3 	.	rx   	"(" rxGName 	")"		rz  	"("	rxGebAm ")"	.	rx	.	rxDatum         	                      	; 07

								       	,	"[Nn]ame[\s\w]*[;:\s]+"	rxName1                                                                                                   	; 08
								       	,	"(" rxAnrede   	")[;,:.\s]+" 	rxName2                                                                                             	; 09
								       	,	"(" rxAnredeM 	")[;,:.\s]+" 	rxName2                                                                                             	; 10
								       	,	"(" rxAnredeF 	")[;,:.\s]+" 	rxName2                                                                                             	; 11

								       	,	"(Nach)*name[:,;.\s]+(?<Name2>" rxPerson ").*Vo(rn|m)ar*me[:,;.\s]+(?<Name1>" rxPerson ")"		; 12
						        		,	"Vo(rn|m)ar*me[:,;.\s]+(?<Name1>" rxPerson ").*(Nach)*name[:,;.\s]+(?<Name2>" rxPerson ")"		; 13

						        		,	"^\s*" rxName2 "\s"]	                                                                                                                    	; 14	^ Muster, Max  ....

		static	rxDates       	:= [	"(?<Tag>\d{1,2})[.,;](?<Monat>\d{1,2})[.,;](?<Jahr>(\d{4}|\d{2}))[^\d]"                                 	; 01	12.11.64 oder 12.11.1964
						               	,	"i)(?<Tag>\d{1,2})[.,;\s]+(?<Monat>" rxMonths ")[.,;\s]+(?<Jahr>(\d{4}|\d{2}[^\d]))[^\d]"        	; 02
										,	"(" rxGebAm ")"	rx  ".{0,20}?" . rxDay . rx . rxMonths2 . rx . rxYears]                                                   	; 03   Geburtsdatum: 12. Nov. 1964


		static rxContinue := "(Nachname|Vorname|Geburtsdatum)"

		;}

		PatBuf	:= Object()

		If debug
		  SciTEOutput("`n  --------------------------------------`n {Methode 1}")

	; auf Festsplatte schreiben zur Fehlerkorrektur
		If (debug > 1) {
			f := FileOpen(A_ScriptDir "\admMiningRegEx.txt", "w", "UTF-8")
			For i, str in RxNames
				f.WriteLine("[-" SubStr("00" i, -1) " -] " str)
			For i, str in RxDates
				f.WriteLine("[-" SubStr("00" i, -1) " -] " str)
			f.Close()
		}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Methode 1:	RegEx Strings nutzen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
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
		If debug {
			SciTEOutput("    gefundene ID's: " PatBuf.Count())
			SciTEOutput(" {Methode 1a}")
		}

	  ; Methode 1a: zähle die im Text vorhandenen Geburtstage der Patienten (PatBuf) und entferne diese
		For PatID, Pat in PatBuf {

			rxGeburtsdatum := StrReplace(Pat.Gd, ".", "\.")
			text := RegExReplace(text, rxGeburtsdatum, "", GdCount)
			PatBuf[PatID].hits += (GdCount * 2)

			If debug
				SciTEOutput("   " PatID " , hits: " PatBuf[PatID].hits ", Datum: " Pat.Gd " +" GdCount " hits")
		}

		spos := 1
		If debug
		  SciTEOutput(" {Methode 2}")
	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Methode 2:	Suche nach Patientennamen über gefundene Geburtstage
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
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

				matches := admDB.MatchID("gd", Datum)
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

						If debug
							SciTEOutput("  Methode 2:`t" PatID " , hits: " PatBuf[PatID].hits ", Datum: " RegExReplace(D, "[\n\r]"))

					}
			}
		}
	;}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Methode 3:	kommt nur zum Einsatz wenn die Methoden 1 und 2 nichts gefunden haben.
	;                  	Findet zwei aufeinander folgende Worte mit Großbuchstaben am Anfang (Personennamen, Eigennamen ...)
	;               	und vergleicht diese per StringSimiliarity-Algorithmus mit den Namen in der Addendum-Patientendatenbank
	; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
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

			ToolTip,,,, 15
	}

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
			If (Pat.hits = maxHits) {
				;SciTEOutput(" (" PatID ") " Pat.Nn ", " Pat.Vn " < hits: " Pat.hits)
				BestHits[PatID] := {"Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd}
			}
	;}

return BestHits
}

FindDocDate(Text, debug=false) 										{                	;-- Behandlungstage und/oder Erstellungsdatum des Dokuments

	static rxWDay           	:= "Mon*t*a*g*|Die*n*s*t*a*g*|Mi*t*t*w*o*c*h*|Do*n*n*e*r*s*t*a*g*|Fr*e*i*t*a*g*|Son*n*a*b*e*n*d*|Sam*s*t*a*g*|Son*n*t*a*g*"
	static rxWDay2       	:= "(?<WD>Mon*t*a*g*|Die*n*s*t*a*g*|Mi*t*t*w*o*c*h*|Do*n*n*e*r*s*t*a*g*|Fr*e*i*t*a*g*|Son*n*a*b*e*n*d*|Sam*s*t*a*g*|Son*n*t*a*g*)"
	static rxMonths      	:= "\d{1,2}[,.]\s*(Jan\.*u*a*r*|Febr\.*u*a*r*|Mä*a*rz|Apr\.*i*l*|Mai|Jun\.*i*|Jul\.*i*|Aug\.*u*s*t*|Sept\.*e*m*b*e*r*|Okt\.*o*b*e*r*|Nov\.*e*m*b*e*r*|Dez\.*e*m*b*e*r*)\s\d{2,4}"
	static rxTags	        	:= {	1:	"Druckdatum|Erstellungszeitpunkt|Dokumentdatum|D[ae]tum|Beginn|ausgedruckt|Probenentnahmedatu*m*|Abnahmedatum|"
											. 	"Berichtsdatum|Eingangsdatum|"
											. 	"Eingang\s+am|gedruckt am|"
											.	"Anfrage\s+vom|Arztbrief\s+vom||Befund\s+vom|Befundbericht\s+vom|Behandlung\s+vom|Ebenen\s+vom|Ebenen\/axial\s+vom|"
											.	"Echokardiografie vom|Konsil\s+vom|Labor\s+vom|Laborblatt\s+vom|Nachricht\s+vom||"
											. 	"Startzeit|Aufgezeichnet|Erstellt\s+am|Eing[.,\-\s]+Dat[.,-]+|Ausg[.,\-\s]+Dat[.,-]+"
										,	2:	"Behandlung|haben wir.*|sich"}

	static rxBehandlung  	:= [	"i)(" rxTags[2] ")\svom\s(?<Datum1>[\d.]+)\s*(bis\s*z*u*m*)\s*(?<Datum2>[\d.]+)"                	; 1| (haben wir ...| sich) vom .... bis (zum) .....
										, 	"i)(" rxTags[2] ")\svom\s(?<Datum1>[\d.]+)\s*(\-)\s*(?<Datum2>[\d.]+)"]                                 	; 2| (haben wir ...| sich) vom ..... - .......

	static rxDokDatum		:=[	"i)(gedruckt|sich)\s*am\s*[;:]*\s*(?<Datum1>\d+\.\d+\.\d+)"                                                	;   1| gedruckt am: 02.01.2020
										,	"i)^[\pL\-]+\s[\pL\-]+\s*[;,.]\s*den\s+(?<Datum1>\d+\.\d+\.\d+)"                                       	;   2|
										,	"i)^[\pL\-]+\s[\pL\-]+\s*[;,.]\s*den\s+(?<Datum1>" rxMonths ")"                                          	;   3|
										,	"i)\s*(" rxTags[1] ")[;:\s]*(" rxWDay ")*[.,\s]+(?<Datum1>\d+\.\d+\.\d+)"                               	;   4| Erstellungszeitpunkt: Do, 02.01.2020
										,	"i)\s*(" rxTags[1] ")[;:\s]*(" rxWday ")*[.,\s]+(?<Datum1>" rxMonths ")"                                        	;   5| Erstellungszeitpunkt: Do. 02. Januar 2020
										,	"^s*[\pL\s\(\)]+\s*[,;.]\s*(den\s)*(?<Datum1>\d{1,2}\.\d{1,2}\.\d{2,4})"                            	;   6| Hamburg, (den) 02.01.2020
										,	"^s*[\pL\s\(\)]+\s*[,;.]\s*(den\s)*(?<Datum1>" rxMonths ")\s*"                                               	;   7| Hamburg, (den) 2. Januar 2020
										,	"^\s*(?<Datum1>\d{2}\.\d{2}\.(\d{2}|\d{4}))\s*$"                                                            	;   8| 02.01.2020 (alleinstehend in Zeile)
										,	"^\s*(?<Datum1>" rxMonths ")\s*$"]                                                                                       	;   9| 02. Januar 2020 (alleinstehend in Zeile)


		DocDates := Object()
		DocDates.Behandlung 	:= Object()
		DocDates.Dokument  	:= Object()
		retDates := Object()
		retDates.Behandlung := Array()
		retDates.Dokument  	:= Array()

	; Text in Zeilen aufspalten
		TLines := StrSplit(Text, "`n", "`r")

	; Behandlung von ... bis ....
		For RxStrIndex, rxDateStr in rxBehandlung {
			For LNr, line in TLines {
				If !RegExMatch(line, rxDateStr, D)
					continue
				DDatum1 := Trim(StrReplace(DDatum1, ",", "."))
				DDatum2 := Trim(StrReplace(DDatum2, ",", "."))
				saveDate := DDatum1 (DDatum2 ? "-" DDatum2 : "")
				If !DocDates.Behandlung.HasKey(saveDate) {
					DocDates.Behandlung[saveDate] := {"fLine":[LNr], "dcount":1}
				}
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
				DDatum1 := Trim(StrReplace(DDatum1, ",", "."))
				DDatum2 := Trim(StrReplace(DDatum2, ",", "."))
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

	; Rückgabe Objekt erstellen


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

	; Mieg, medilog, Darwin, professional. 							- LZ-EKG
	; Arztanfrage, (Muster 52), 											- Muster 52

}

FindDocPages(Text) 															{ 	         		;-- Zeilen je Seite durch Suche nach Pagebreaks im Text

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


