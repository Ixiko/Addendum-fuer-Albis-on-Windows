; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                         	** Addendum_Mining **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Beschreibung:	    	Funktionen für Datamining aus Textdateien
;										nutzt RegEx-Suchmuster um Daten innerhalb eines Textes zu finden
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
; 		Inhalt:
;       Abhängigkeiten:		Addendum_Gui.ahk, Addendum.ahk
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_Calc started:    	12.07.2022
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
return

class DocStrings                                                      	{                	;-- RegExStrings für alles

	; ‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿‿
	;   REGEX-STRINGS FÜR DAS FINDEN VON PERSONENNAMEN UND DATUMSZAHLEN IN EPIKRISEN, UNTERSUCHUNGSBEFUNDEN, BRIEFEN
	;   jede Funktion kann durch globales Deklarieren Zugriff nur auf die benötigten Variablen erhalten
	;
	;⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀⁀;{

	static Init:= this.regexStrings()


	regexStrings() {

	;{  Shorts, Tags und Parts
	ThousandYear:= SubStr(A_YYYY, 1, 2)
	HundredYear	:= SubStr(A_YYYY, 3, 2)

	excludeIDs  	:= 	"2"

	rx               	:= 	"[;,:.\s]+"
	rx1             	:= 	"[;,:.\s]"
	rxS              	:= 	"[;,:.\s]*"
	rxNL            	:= 	"(\s+|\s*[\n\r]+)"                                                                                          ; für Sätze (mit Zeilenumbruch)
	ry               	:= 	"[,:.;\s\n\r\f]+"
	ryNL           	:= 	"[\s\n\r]+"
	rz               	:= 	"[.,;\s]+.*?"
	rs               	:= 	"\s+"
	rp1             	:= 	".*"
	rp2             	:= 	".*?"

	rxDay         	:= 	"(?<D>[12]*[0-9])"
	rxWDay       	:= 	"Mon*t*a*g*|Die*n*s*t*a*g*|Mi*t*t*w*o*c*h*|Do*n*n*e*r*s*t*a*g*|Fr*e*i*t*a*g*|Son*n*a*b*e*n*d*|Sam*s*t*a*g*|Son*n*t*a*g*"
	rxWDay2     	:= 	"(?<WD>" rxWDay ")"
	rxMonths     	:=	"Jan*u*a*r*|Feb*r*u*a*r*|Mä*a*rz|Apr*i*l*|Mai|Jun*i*|Jul*i*|Aug*u*s*t*|Sept*e*m*b*e*r*|Okt*o*b*e*r*|Nov*e*m*b*e*r*|Deze*m*b*e*r*"
	rxMonths2   	:=	"(?<Monat>" rxMonths ")"
	rxYear        	:=	"(?<Jahr>([12][0-9])*[0-9]{2})"                                                                      	; Jahr=(19)*85
	rxYear2      	:=	"(?<JH>[12][0-9])*(?<JZ>[0-9]{2})"                                                                	; JH=19*,  JZ=85
	rxYear3      	:=	"(?<Jahr>(\d{4}|\d{2}))"                                                                             	; Jahr= 1985 o. 85

	rxPerson     	:= 	"[A-ZÄÖÜ][\pL]+([\-\s][A-ZÄÖÜ][\pL]+)*"
	rxPerson1    	:= 	"(?<Name1>" rxPerson ")"                                                                                	; ?<Name1>Müller-Wachtendónk*
	rxPerson2    	:= 	"(?<Name2>" rxPerson ")"                                                                                	; ?<Name2>Müller-Wachtendónk*
	rxPerson3   	:= 	"[A-ZÄÖÜ][\pL]+([\-\s][A-ZÄÖÜ][\pL]+)*"                                                        	; Müller Wachtendonk*
	rxPerson4    	:= 	"(?<Name>[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL]+)*)"		                                     	; ?<Name>Müller Wachtendonk*

	rxGeschl       	:= 	"(?<Sex>\(W|M\))"

	 rxNames1    	:= 	"(?<Name2>" rxPerson	  ")" . rx . "(?<Name1>"	rxPerson ")"                         	; <Name2>Baselitz, <Name1>Georg
	 rxNames2    	:= 	rxPerson2 . rx . 	rxPerson1	                                                                          	; <Name2>Baselitz, <Name1>Georg
	 rxNames3    	:= 	rxPerson1 . rx . 	rxPerson2	                                                                          	; <Name2>Georg, <Name1>Baselitz
	 rxGName   	:= 	"(?<GName2>" rxPerson	  ")" . rx . "(?<GName1>"	rxPerson ")"                    	; <Name2>Kern, <Name1>Hans-Georg

	;~ static rxDate	    	:= 	"\d{1,2}" rx "[\wäöü]+" rx "(\d{4}|\d{2})"                                                   	; 12. Juli 1981 o. 12.7.1981 o. 12.07.1981
	rxDate	    	:= 	"(\d{1,2})" rx "(0[123456789]|1[012]|[a-zäöüß]+)" rx "(\d{4}|\d{2})"              	; 12. Juli 1981 o. 12.7.1981 o. 12.07.1981
	rxDateLong 	:= 	"\d{1,2}" rx "(" rxMonths ")" rx "\d{2,4}"                                                            	; 12. Jul. 81
	rxGb1        	:= 	"(geb(\.|oren)(\sam)*|Geb(urts)*" rxS "Dat(um)*|\*" rxS ")"                               	; geb. am o. Geb. Dat.   und mehr
	rxGb2        	:= 	"(geb(\.|oren)(\sam)*|Geb([.,;]|urts)\s*(D|d)at([.,;]|um)|\*)"                              	;
	rxBirth         	:= 	"(?<Birth>(?<Tag>\d{1,2})" rx "(?<Monat>\d{1,2})" rx "(?<Jahr>\d{4}|\d{2}))"  ; Datum wird als Geburtsdatum interpretiert

	 rxGN           	:= 	"Geb\.\sName|Geburtsname"
	 rxAnrede     	:= 	"Betrifft"
	 rxAnredeM   	:= 	"(Her(rn|m|r)|Patient|Patienten)"
	 rxAnredeF   	:= 	"(Frau|Patientin)"
	 rxAnredeXL 	:=	"(Versicherter*|Mitglied|Betrifft|Her(rn|m|r)|Patiente*n*|Frau|Patientin|PatName)"
	 rxVorname  	:= 	"[VY]o(rn|m)ar*me"
	 rxNachname	:= 	"([Ff]amilien)*([Nn]ach)*[Nn]ame"
	 rxVorUnd     	:= 	rxVorname "\s+und\s+" rxNachname
	 rxNameW    	:= 	"[Nn]ame[\w]*"
	;}

	;{  Name + Geburtsdatum
	this.rxNames     	:= [		rxAnredeXL  	. 	rx	. 	rxNames1 .	rx  	.	rxGb1   . 	ry	.   rxBirth              	; 01	Betrifft: Marx, Karl, geb. am: 14.03.1818
							    		,	rxAnredeXL  	.	rx	. 	rxNames2	.	rx  	.	rxGb1   . 	ry	.	rxBirth              	; 02	Versicherter: Marx, Karl, Geburtsdatum: 14.03.1818
							    		,	rxNames1    	.	rx	.	rxGb1   	.	rs                          	.   rxBirth           	; 03  <Name2>Baselitz, <Name1>Georg, Geb.Datum:
																																							; 		23.01.1938
							    		,	rxAnredeM 	. 	rx .  rxNames2 .	rx 	.	rxGb2  	.	rx	.	rxBirth           	; 04	Patient Marx, Karl, *14.03.1818
							    		,	rxAnredeF 	. 	rx .  rxNames2 .	rx 	.	rxGb2  	.	rx	.	rxBirth            	; 05	Patientin Luxemburg, Rosa, *05.03.1871
										,	rxNames3   	.	rx	.	rxGb1     	.	rx 	. 	rxBirth                                     	; 06	Karl Marx geb. am 14.03.1818
										,	rxNames2   	.	rx .	rxBirth                                                                    	; 07	Karl Marx, 14.03.1818
										,	rxNames2   	.	rx .	rxGeschl 	.	rx		. rxBirth                                        	; 08	Karl Marx (M) 14.03.1818
							    		,	rxNames1    	.	rx	.	rxGN       	.	rp2 . rxGName	.	rxGb1 .  rx . rxBirth   	; 09  <Name2>Baselitz, <Name1>Georg, Geburtsname:
																																							;		Hans-Georg Kern, geb. am 23.01.1938

							    		,	rxVorUnd     	.	rx	. rxAnredeXL . 	rx 	. 	rxNames3 	                           	; 10  Vorname und Name:  Herrn <Name1>Heinrich
																																							; 		<Name2>Hoffmann
							    		,	rxAnredeXL  	. 	rx  .  ryNL      	.	rxNames2                                          	; 11	Versicherte`n Müller-Wachtendonk, Marie-Luise
							    		,	rxAnredeM 	. 	rx  .  rxNames2                                                               	; 12	Herr Karl Marx o. Patient Marx, Karl
							    		,	rxAnredeF 	. 	rx  .  rxNames2                                                                	; 13	Frau Marie-Luise Müller-Wachtendonk

							    		,	rxNameW		.	rx	.	rxNames1                                                                	; 14	Name: Marx, Karl
							    		,	rxNachname	. 	rx	. 	rxPerson2	.	rp2 	.	rxVorname	. rx . rxPerson1     	; 15	Nachname: Karl, Versichertennummer: 12345678,
																																							; 		Vorname: Marx
							    		,	rxVorname	. 	rx	. 	rxPerson2	.	rp2	.	rxNachname	. rx . rxPerson2     	; 16	Vorname: Marx, Versichertennummer: 12345678,
																																							; 		Nachname: Karl

										,	rxNames2   	. ".*?[\n\r]+.*"   	.	rxNachname	.	rx	.	rxVorname			; 17	Karl Marx \n\r Name, Vorname

							    		,	"^\s*" rxNames2 "\s"]	                                                                           	; 18	^ Muster, Max  ....


	this.rxVNR          	:= [ "i)(VNR|Versicherten[\s\-]*N[ume]*r" . rx ")(?<VNR>[A-Z][\s\d]+)"]           	; 01  Versicherten-Nr. G 666 999 666

	this.rxDates       	:= [	"(?<Tag>\d{1,2})[.,;/](?<Monat>\d{1,2})[.,;/]" rxYear3 "([^\d]|$)"         	; 01	12.11.64 o. 12.11.1964 o. 2.1.(19)*64
								    	,	"i)(?<Tag>\d{1,2})" rx . rxMonths2 . rx . rxYear3 "([^\d]|$)"                  	; 02	12. November (19)64
										,	rxGb2 . rxS . rxBirth                                                                                 	; 03	Geburtsdatum: 12.11.64
								    	,	"(" rxGb2 ")"	rx  ".{0,20}?" . rxDay . rx . rxMonths2 . rx . rxYear]              	; 04  Geburtsdatum: .*(20 beliebige Zeichen) 12. Nov. 1964

	this.rxContinue  	:= "(Nachname|Vor[nm]ame|Geburtsdatum)"
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; REGEXSTRINGS FÜR DOKUMENTDATUM (FindDocDate)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	this.rxTags	       	:= {	1	:	"(Druckdatum|Erstellungszeitpunkt|Dokumentdatum|D[ae]tum|Untersuchungsdatum|Beginn|Probenentnahmedatu*m*|"
											. 	"Abnahmedat(\.|um)*|Berichtsdat(\.|um)*|Eingangsdat(\.|um)*|Eingang(\s+am)*|(aus)*gedruckt\s+(am)*|angelegt(\s+am)*|"
											. 	"Ausgedruckt(\s+am)*|Anfrage\s+vom|Arztbrief\s+vom|Befund\s+vom|Befundbericht\s+vom|Behandlung(sbericht)*\s+vom|"
											.	"Ebenen\s+vom|Ebenen\/axial\s+vom|Echokardiogra(f|ph)ie\s+vom|Konsil\s+vom|Labor\s+vom|Laborblatt\s+vom|"
											. 	"Nachricht\s+vom|Startzeit|Aufgezeichnet|Erstellt\s+am|Eing[.,\-\s]+Dat[.,-]+|Ausg[.,\-\s]+Dat[.,-]+Labor(blatt)*\svom|"
											. 	"festgestellt(\s+am)*|Aufnahme\s*vom.*?|festgestellt(\s+am)*)"
									,	2	:	"Behandlung|haben wir|sich"}

										; 	Behandlungszeiträume
	this.rxBehandlung	:= [	"i)\s+vom"   	 rxNL "(?<Datum1>\d+[.,;]+\d+[.,;]+(\d{4}|\d{2})*)" rxNL
										. 	"(bis(\szum)*)" rxNL "(?<Datum2>\d+[.,;]+\d+[.,;]+(\d{4}|\d{2}))"                          ; 	1| vom Datum1 bis (zum) Datum2
										, 	"i)(?<Datum1>[\d.,;]+)\s*(\-)\s*(?<Datum2>[\d.,;]+)"                                             	; 	2| vom Datum1 - Datum2.
										, 	"i)Behandlungstag" rxNL "(?<Datum1>\d+[.,;]+\d+[.,;]+(\d{4}|\d{2})*)"
										,	"i)(wir\s+berichten|ich\s+berichte).*?untersuchung\s+vom\s+(?<Datum1>\d+[.,;]+\d+[.,;]+(\d{4}|\d{2})*)"]

	this.rxDokDatum	:=[		"i)\s*(" rxTags[1] ")" rx "(?<Datum1>\d+\.\d+\.\d+)"                                                	;   1| Eingang 02.03.22
							    		,	"i)\s*(" rxTags[1] ")" rx "(" rxWDay ")*[.,;\s]+(?<Datum1>" rxDateLong ")"                  	;   2| Erstellungszeitpunkt: Do. 02. Januar 2020
										,	"i)\s*(" rxTags[1] ")" rx "(" rxWDay ")*[.,;\s]+(?<Datum1>\d+\.\d+\.\d+)"                  	;   3| Erstellungszeitpunkt: Do, 02.01.2020
										, 	"i)((aus)*gedruckt|angelegt|sich)\s*(am)*" rxS "(?<Datum1>\d+\.\d+\.\d+)"           	;   4| gedruckt am: 02.01.2020
										,	"i)((aus)*gedruckt|angelegt)\s*[;:]*\s*(?<Datum1>\d+\.\d+\.\d+)"                          	;	5| angelegt: 02.01.2020
							    		,	"i)^[\pL\-]+\s[\pL\-]+\s*[;,.]\s*(den)*\s+(?<Datum1>\d+\.\d+[.;,\s]+\d+)\s*$"     	;   6| Hamburg-Hochburg, 02.01.2020 (nur
																																											; 		 Leerzeichen oder Zeilenende folgt)
							    		,	"i)^[\pL\-]+\s[\pL\-]+\s*[;,.]\s*(den)*\s+(?<Datum1>" rxDateLong ")\s*$"             	;   7| Hamburg-Hochburg, (den) 02. Januar 2020
																																											; 		 (nur Leerzeichen oder Zeilenende folgt)
							    		,	"^\s*[\pL\s\(\)]+\s*[,;.]\s*(den\s)*(?<Datum1>\d{1,2}\.\d{1,2}[.;,\s]+\d{2,4})"     	;   8| Hamburg, (den) 02.01.2020
							    		,	"^\s*[\pL\s\(\)]+\s*[,;.]\s*(den\s)*(?<Datum1>" rxDateLong ")\s*"                           	;   9| Hamburg, (den) 2. Januar 2020
							    		,	"^\s*(?<Datum1>\d{2}\.\d{2}\.(\d{2}|\d{4}))\s*$"                                            	; 10| 02.01.2020 (alleinstehend in Zeile)
							    		,	"^\s*(?<Datum1>" rxDateLong ")\s*$"]                                                                  	; 11| 02. Jan(.|uar) 2020 (alleinstehend in Zeile)

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; ZUM VALIDIEREN EINES DATUMS
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	static rxDateOCRrpl		:= "[,;:]"
	static rxDateValidator	:= [ 	"^(?<D>[0]?[1-9]|[1|2][0-9]|[3][0|1]).(?<M>[0]?[1-9]|[1][0-2]).(?<Y>[0-9]{4}|[0-9]{2})$"
										,	"^(?<D>[0]?[1-9]|[1|2][0-9]|[3][0|1])[.;,\s]+(?<M>" rxMonths ")[.;,\s]+(?<Y>[0-9]{4}|[0-9]{2})$"]


	static rxSentence := "[A-Z]\pL.*?\.(?=[\s\n\r]|[A-Z]|$)"
	static rxSatzende := "(?!\pL)\,(?=[\s\n\r]*[A-Z]|$)" ; Komma durch . ersetzen
	}
	;}

	dcsDebug                                                                  	{                 	;-- Eigenschaft setzen für debugging

		get {
			return this.debug
		}

		set {
			this.debug := value
		}

	}

	FindNames(Text, unknow:=false, callbackFunc:="")       	{               	;-- sucht Namen von Patienten im Dokument

		; letzte Änderung: 12.07.2022 - Methode 2 - stringsimilarity Suche und Umlaute ÄÖÜ äöü durch AeOeUe aeueoe erstetzen

			static methode := ["Methode 1: finden über Namen"
									, 	"Methode 2: finden über Geburtstage"
									, 	"Methode 3: finden über Namensvergleich"
									, 	"Methode 4: finden über Adressen (cPat Objekt wird benötigt)"]

			WorkText := RegExReplace(Text, "[\n\r]+", "  ")
			WorkText := RegExReplace(WorkText, "(\pL)[^\pL\d\s\-\.\*\@](\pL)", "$1 $2")
			WorkText := RegExReplace(WorkText, "([a-zäöü])\-\s\s*([a-zäöü])", "$1$2")                                         	; Silbentrennung aufheben
			WorkText := RegExReplace(WorkText, "([A-ZÄÖÜ])([A-ZÄÖÜ][a-zäöüß]{2,}[\s\-\@\.;,:])", "$1 $2")   	; Großbuchstaben abtrennen
																																					; KGeriatrie@Krankenhaus.de -> K Geriatrie....
			WorkText := RegExReplace(WorkText, "(Seite)\s(\d+[\\\/\-v]*\d*)", "$1 $2`n`n")                                	; bei Seitenangaben im Dokument eine Trennung einsetzen

			WorkText := RegExReplace(WorkText, "\s{2,}", " ")
			words := StrSplit(WorkText, A_Space)
			wordsCount := words.Count()

			; auf Festsplatte schreiben zur Fehlerkorrektur
				If (this.debug > 1) {
					f := FileOpen(A_ScriptDir "\admMiningRegEx.txt", "w", "UTF-8")
					For i, str in RxNames
						f.WriteLine("[-" SubStr("00" i, -1) " -] " str)
					For i, str in RxDates
						f.WriteLine("[-" SubStr("00" i, -1) " -] " str)
					f.Close()
				}

		; ---------------------------------------------------------------------------------------------------------------------------------------------------
		  PatBuf := Object()
		; Methode 1:	RegEx Strings nutzen
		; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
			If this.debug {
				SciTEOutput("`n" WorkText)
				SciTEOutput("`n --------------------------------------`n {" methode.1 "}")
		  }

		; - benutzt RegEx-Strings aus RxNames um Patientennamen zu finden
			For RxStrIndex, rxNameStr in this.RxNames {

				spos	:= 1
			  ; zählt alle Vorkommen des RegExStrings im Text wenn ein passender Patient in der Albis Datenbank vorhanden ist
				while (spos := RegExMatch(WorkText, rxNameStr, Pat, spos)) {

				  ; end of line Zeichen entfernen
					spos  += StrLen(Pat)

				; zählt die gefundenen Wörter, bei mehr als 4 gefundenen Worten wird es sich nicht um einen Namen handeln
					pnrpl := RegExReplace(Pat, "[\pL\-]", " ")

					PatName2 := RegExReplace(PatName2, "^.*[\n\r]+", ""), Pat := RegExReplace(Pat, "^.*[\n\r]+", "")

				  ; debug Ausgabe ;{
					If (debug = 2)
						SciTEOutput("    [" RxStrIndex " - " rxNameStr "]`n         " PatName2 ", " PatName1)
					else if (debug = 1)
						SciTEOutput("    [" RxStrIndex "]  " PatName2 ", " PatName1  " / " spos " (" Pat ")")
				  ;}

				  ; Unpassendes entfernen
					PatName1 := StrReplace(PatName1, " "), PatName2 := StrReplace(PatName2, " ")
					If RegExMatch(PatName1, rxContinue) || RegExMatch(PatName2, rxContinue)
						continue

				  ; schaut ob der gefundene Name in der Datenbank ist
					sex        	:= RegExMatch(Pat, rxAnredeM) ? "m" : RegExMatch(Pat, rxAnredeF) ? "w" : ""
					birth     	:= RegExMatch(PatBirth, "\.\d\d$") ? RegExReplace(PatBirth, "\.(\d\d)$", ".\d\d$1") : PatBirth      ; macht aus 19.01.60 > 19.01.\d\d60
					birth  		:= RegExReplace(PatBirth, "\.", "\.")
					matches	:= cPat.StringSimilarityEx(PatName1, PatName2)
					mdiff     	:= matches.Delete("diff")
					If IsObject(matches)
						For PatID, Pat in matches {

						  ; diese ID's ignorieren
							If PatID in %excludeIDs%
								continue

						  ; neue ID hinzufügen oder zählen
							If !PatBuf.haskey(PatID) {
								PatBuf[PatID] := {"Nn"   	: cPat.Get(PatID, "NAME")
														, "Vn"    	: cPat.Get(PatID, "VORNAME")
														, "Gd"   	: cPat.GEBURT(PatID, true)
														, "hits"   	: 1
														, "method"	: 1}
							}
							else
								PatBuf[PatID].hits += 1

							PatBuf[PatID].hits += StrLen(sex)>0 	&& cPat.GESCHL(PatID)=sex ? 1 : 0
							PatBuf[PatID].hits += StrLen(birth)>0	&& RegExMatch(cPat.GEBURT(PatID, true), birth) ? 1 : 0
							If debug
								SciTEOutput("    +" PatID " , hits: " PatBuf[PatID].hits ", Name: " PatBuf[PatID].Nn ", "PatBuf[PatID].Vn ", " PatBuf[PatID].Gd)

						}

				}

			}

		 ; debugging
			If this.debug
				SciTEOutput("    gefundene ID's: " PatBuf.Count())

		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		  ; Methode 1a: zählt vorhandene Geburtstage im Text der PatBuf-Patienten und entfernt diese anschließend
		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			If this.debug && (PatBuf.Count() > 0)
				SciTEOutput("`n {Methode 1a: Geburtstage der Patienten finden}")

			GdCountSum := 0
			For PatID, Pat in PatBuf {

				rxGeburtsdatum := StrReplace(Pat.Gd, ".", "[\.\,\;]")
				rxGeburtsdatum := SubStr(Pat.Gd, 1, 6) "(" SubStr(Pat.Gd, 7, 2) ")" SubStr(Pat.Gd, 9, 2)
				WorkText := RegExReplace(WorkText, rxGeburtsdatum, "", GdCount)
				PatBuf[PatID].hits += (GdCount * 2)
				GdCountSum += GdCount

				If this.debug
					SciTEOutput("   " PatID " , hits: " PatBuf[PatID].hits ", Datum: " Pat.Gd " +" GdCount " hits")
			}


			;}

		;}

		; ---------------------------------------------------------------------------------------------------------------------------------------------------
		; Methode 2:	Suche nach Patienten über im Text enthaltene Datum-Zeichenketten
		;                	ein gefundenes Datum wird mit den Geburtstagen aus der Datenbank verglichen
		; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
			If this.debug
			  SciTEOutput("`n {" methode.2 "}")

			birthdays := "|"
			For rxDateIdx, rxDateString in this.rxDates {

				spos := 1
				while (spos := RegExMatch(WorkText, rxDateString, D, spos)) {

					spos += StrLen(D)
					If IsFunc(callbackFunc)
						%callbackFunc%(methode.2, spos, StrLen(WorkText))

				  ; Monat wurde als Wort geschrieben, dann die Nummer des Monats ermitteln
					If !RegExMatch(DMonat, "\d{1,2}")
						For nr, rxMonth in StrSplit(rxMonths, "|")
							If RegExMatch(DMonat, "i)" rxMonth) {
								DMonat := SubStr("00" nr, -1)
								break
							}

				  ; Zusammenfügen und das Datumsformat anpassen (aus dem 1.1.60 wird 01.01.1960)
					GEBURT := SubStr("00" DTag, -1) "." SubStr("00" DMonat, -1) "." DJahr
					GEBURT := this._FormatDateEx(GEBURT, "DMY", "dd.MM.yyyy")

				  ; doppelt Suche ausschließen
					If InStr(birthdays, "|" GEBURT "|")
						continue
					birthdays .= GEBURT "|"

				  ; gefundenes Datum ausgeben
					If this.debug
						SciTEOutput("    [" rxDateIdx "]"  GEBURT )

				  ; sucht nach Patienten mit passenden Geburtstagen
					PatIDs := cPat.PatID([{"key":"GEBURT", "value":this._ConvertToDBASEDate(GEBURT)}])
					If IsObject(PatIDs) {

						For each, PatID in PatIDs {

						  ; bestimmte ID's ausschließen
							If PatID in %excludeIDs%
								continue

						  ; Name des Patienten muss im Text vorkommen
							NAME        	:= cPat.Get(PatID, "NAME")
							VORNAME	:= cPat.Get(PatID, "VORNAME")
							If RegExMatch(WorkText,  "i)(" NAME ".*?" VORNAME "|" VORNAME ".*?\s" NAME ")", nstr) {

							  ; PatID bekannt, Zähler erhöhen
								If PatBuf.haskey(PatID) {
									addSign := " "
									PatBuf[PatID].hits += PatBuf[PatID].method=1 ? 2 : 1       ; bekommt zwei Treffer-Punkte wenn zuvor die PatID mit Methode1 gefunden wurde
								}
							  ; PatID ist unbekannt
								else {
									addSign := "+"
									PatBuf[PatID] := {"Nn"   	: NAME
															, "Vn"    	: VORNAME
															, "Gd"   	: GEBURT
															, "hits"   	: 1
															, "method"	: 2}
								}

							  ; Treffer ausgeben
								If this.debug {
									SciTEOutput("       gefundener Namensstring: " nstr)
									SciTEOutput("     " addSign "[" PatID "]" NAME "," VORNAME "*" GEBURT " , hits: " PatBuf[PatID].hits)
								}

							}

						}

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

				If this.debug
				  SciTEOutput("`n {" methode.3 "}")

				spos := 1
				while (spos := RegExMatch(WorkText, "([A-ZÄÖÜ][\pL-]+)[\s,.;:\n\r]+([A-ZÄÖÜ][\pL-]+)", name, spos)) {

					; überflüssige Zeichen entfernen
						name1 := RegExReplace(name1, "[\s\n\r\f]")
						name2 := RegExReplace(name2, "[\s\n\r\f]")

						If IsFunc(callbackFunc)
							%callbackFunc%(methode.3, spos, StrLen(WorkText))

					; Bindestrichworte ignorieren
						If RegExMatch(name1, "\-$") {
							spos += StrLen(name1)                    ; ein Wort weiter
							continue
						} else if RegExMatch(name2, "\-$") {
							spos += StrLen(name1)                     ; zwei Worte weiter
							continue
						} else
							spos += StrLen(name1)

					; gefundene Worte mit Namen in der Patientendatenbank vergleichen
						matches	:= cPat.StringSimilarityEx(name1, name2)
						mdiff     	:= matches.Delete("diff")
						If IsObject(matches) {

							For PatID, Pat in matches {

								If PatID in %excludeIDs%
									continue

								If PatBuf.haskey(PatID) {
									addSign := " "
									PatBuf[PatID].hits += 1
								}
								else {
									addSign := "+"
									PatBuf[PatID] := {"Nn"   	: cPat.Get(PatID, "NAME")
															, "Vn"    	: cPat.Get(PatID, "VORNAME")
															, "Gd"   	: cPat.GEBURT(PatID, true)
															, "hits"   	: 1
															, "method"	: 3}
								}

								If this.debug
									SciTEOutput("  " addSign "[" PatID "]" PatBuf[PatID].Nn ", "PatBuf[PatID].Vn " *" PatBuf[PatID].Gd " , hits: " PatBuf[PatID].hits)
							}

						}


				}

		}
		;}

		; ---------------------------------------------------------------------------------------------------------------------------------------------------
		; Methode 4:	finden über Adressen
		; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
			If this.debug
			  SciTEOutput("`n {" methode.4 "}")

			If (PatBuf.Count()=0 && IsObject(cPat)) {

				PatIDs := cPat.GetPatIDsAll()
				For each, PatID in PatIDs {

					If PatID in %excludeIDs%
						continue

					If this.debug {
						If (Mod(each, 300) = 0)
							ToolTip, % "Methode 4:  Patient " each "/" PatIDs.Count() " verarbeitet"
					}

					If (IsFunc(callbackFunc) && Mod(each, 300) = 0)
						%callbackFunc%(methode.4, each, PatIDs.Count())

					If !(Strasse := cPat.Get(PatID, "STRASSE"))
						continue

					Hausnr	:= cPat.Get(PatID, "HAUSNUMMER")
					Ort   	:= cPat.Get(PatID, "ORT")
					PLZ   	:= cPat.Get(PatID, "PLZ")

					If RegExMatch(Hausnr, "Oi)(?<Zahl>\d+)\s*(?<Zusatz>[a-zäöü]{1,2})", Hnr)
						HausNr := Hnr.Zahl "\s*" Hnr.Zusatz

					If (RegExMatch(WorkText " ", strasse "\s*" Hausnr "[\s,.;:\n\r\t]") && InStr(WorkText, Ort) && InStr(WorkText, PLZ)) {

						If !PatBuf.haskey(PatID) {
							PatBuf[PatID]	:= {	"Nn"	    	: cPat.Get(PatID, "NAME")
													, 	"Vn"	    	: cPat.Get(PatID, "VORNAME")
													, 	"Gd"	    	: cPat.Get(PatID, "GEBURT")
													, 	"hits"	    	: 1
													, 	"method"	: 4}
						}

						If InStr(WorkText, Ort)
							PatBuf[PatID].hits += 1
						If InStr(WorkText, PLZ)
							PatBuf[PatID].hits += 1
						If InStr(WorkText, cPat.Get(PatID, "GEBURT"))
							PatBuf[PatID].hits += 1
						If InStr(WorkText, cPat.Get(PatID, "NAME"))
							PatBuf[PatID].hits += 1
						If InStr(WorkText, cPat.Get(PatID, "VORNAME"))
							PatBuf[PatID].hits += 1


						If this.debug
							SciTEOutput("   " strasse ", " Hausnr ", " PLZ " " Ort ":   [" PatID "] " PatBuf[PatID].Nn ", " PatBuf[PatID].Vn " *" PatBuf[PatID].Gd  " , hits: " PatBuf[PatID].hits)

					}

				}

				ToolTip

			}
		;}

		; ---------------------------------------------------------------------------------------------------------------------------------------------------
		; Methode 5: mehrere Patienten gefunden. Welche Patientennamen findest man im Text
		; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
			If (PatBuf.Count() > 1) {

				If this.debug
				  SciTEOutput("`n {Methode 5: mehrere Patienten gefunden, zählen wie oft die Namen vorkommen}")

			  ; alle Worte mit einem Großbuchstaben am Anfang finden
				spos := 1, PNames := Array()
				textB := RegExReplace(WorkText, "[^\pL\-\s]", " ")
				textB := RegExReplace(textB, "\s{2, }", " ") " "

				If this.debug
					SciTEOutput(textB "`n - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - `n")

				while (spos := RegExMatch(textB, "\s" this.rxPerson4, P, spos)) {
					PNames.Push(xstring.TransformGerman(PName))
					spos += StrLen(P)
				}

			  ; Vor- und Nachnamen mit jedem Wort vergleichen und Treffer einzeln zählen
				For PatID, Pat in PatBuf {

					thisPatNn := thisPatVn := 0
					NAME       	:= xstring.TransformGerman(Pat.Nn)
					VORNAME 	:= xstring.TransformGerman(Pat.Vn)

					For pidx, word in PNames {

						DiffA	:= StrDiff(word, NAME), DiffB := StrDiff(word, VORNAME)
						If (DiffA <= 0.21)
							thisPatNn += 1
						If (DiffB <= 0.21)
							thisPatVn += 1

						If (this.debug) && (thisPatVn+thisPatNn>=2) && (thisPatVn = thisPatNn)
							SciTEOutput("   PatID:    `t[" PatID "] " Pat.Nn ", " Pat.Vn ", Position: " pidx ", vollst. Namenstreffer: " thisPatVn)

					}

					If (thisPatVn+thisPatNn>=2) {

						hits := Min(thisPatVn, thisPatNn)
						PatBuf[PatID].hits += hits

						If this.debug
							SciTEOutput("   PatID:    `t[" PatID "] " Pat.Nn ", " Pat.Vn " , hits: " PatBuf[PatID].hits)
					}
				}
			}
		;}


		; höchste Trefferzahl ermitteln ;{
			maxHits := 0, BestHits := Object()
			For PatID, Pat in PatBuf {
				If (PatID = "Diff")
					continue
				maxHits := Pat.hits > maxHits ? Pat.hits : maxHits
				If this.debug
					SciTEOutput("   hitpoints:`t" Pat.hits " [(" PatID ") " Pat.Nn ", " Pat.Vn "]")
			}
		;}

		; alle Namen mit gleicher Trefferanzahl werden behalten ;{
			For PatID, Pat in PatBuf
				If (PatID = "Diff")
					continue
				else If (Pat.hits = maxHits)
					BestHits[PatID] := {"Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd}
		;}

	return BestHits
	}

	FindDates(Text, names="") 					                    	{                	;-- Behandlungstage und/oder Erstellungsdatum des Dokuments

			If this.debug
			  SciTEOutput("`n  --------------------------------------`n {FindDocDates}")

		; ---------------------------------------------------------------------------------------------------------------------------------------------------
		; Objekte
		; ---------------------------------------------------------------------------------------------------------------------------------------------------;{
			DocDates := Object()
			DocDates.Behandlung 	:= Object()
			DocDates.Dokument  	:= Object()

			retDates := Object()
			retDates.Behandlung := Array()
			retDates.Dokument  	:= Array()

			workText := Text
		;}

		; Behandlung von ... bis ....
			For RxStrIndex, rxDateStr in this.rxBehandlung {

				spos := 1
				while (spos := RegExMatch(workText, rxDateStr, D, spos)) {

					spos += StrLen(D)

				; Datum formatieren, fehlerhafte Interpunktion ändern
					DDatum1	:= RegExReplace(DDatum1, "[\.\,\;\:\s]+", ".")
					DDatum1	:= RegExReplace(DDatum1, "\.{2,}", ".")
					DDatum1 := this._FormatDateEx(DDatum1, "DMY", "dd.MM.yyyy")
					DDatum1 := DateValidator(DDatum1, A_YYYY)

					DDatum2	:= RegExReplace(DDatum2, "[\.\,\;\:\s]+", ".")
					DDatum2	:= RegExReplace(DDatum2, "\.{2,}", ".")
					DDatum2 := this._FormatDateEx(DDatum2, "DMY", "dd.MM.yyyy")
					DDatum2 := DateValidator(DDatum2, A_YYYY)

					If this.debug && (DDatum1 || DDatum2)
						SciTEOutput("  von " (DDatum2 ? "| bis " : "") ": " DDatum1 (DDatum2 ? " | "  DDatum2 : "") )

				; prüft ob die Daten Geburtstage sind, löscht diese dann
					if DDatum1 {
						PatID1 := cPat.PatID([{"key":"GEBURT", "value":this._ConvertToDBASEDate(DDatum1)}])
						If this.debug
							SciTEOutput(" DDatum1: " this._ConvertToDBASEDate(DDatum1))
					}
					If DDatum2 {
						PatID2 := cPat.PatID([{"key":"GEBURT", "value":this._ConvertToDBASEDate(DDatum2)}])
						If this.debug
							SciTEOutput(" DDatum2: " this._ConvertToDBASEDate(DDatum2))
					}

					If this.debug
						SciTEOutput(" PatID1: " PatID1.Count()  ", PatID2: " PatID2.Count())
					If (this.debug && PatID1.Count()>0)
						SciTEOutput(" D1: " DDatum1 "  [(" RxStrIndex ") " rxDateStr "] matched with birthday(s)"  this._ConvertToDBASEDate(DDatum1)  )
					If (this.debug && PatID2.Count()>0)
						SciTEOutput(" D2: " DDatum2 "  [(" RxStrIndex ") " rxDateStr "] matched with birthday(s)" this._ConvertToDBASEDate(DDatum2)  )

					DDatum1 := PatID1.Count()>0 ? "" : DDatum1
					DDatum2 := PatID2.Count()>0 ? "" : DDatum2
					If (StrLen(DDatum1 . DDatum2) = 0)
						continue


				; DDatum1 ist leer und DDatum2 enthält eine Datum dann das Datum nach DDatum1 kopieren
					If (!DDatum1 && DDatum2)
						DDatum1 := DDatum2, DDatum2 := ""

				; Datumstring(s) speichern
					saveDate := DDatum1 (DDatum2 ? "-" DDatum2 : "")

					If this.debug && saveDate
						SciTEOutput("   Date" (InStr(saveDate, "-") ? "s" : "") ": " saveDate)

				; ein von bis Datum hat Vorrang, DDatum1 ist kleiner als DDatum2, erstes Datum liegt vor dem 2.Datum,
				; damit Beginn und Ende extrahiert, weitere Datumsuche ist nicht notwendig
					DbaseDate1 := this._ConvertToDBASEDate(DDatum1)
					DbaseDate2 := this._ConvertToDBASEDate(DDatum2)
					If (DDatum1 && DDatum2) && ( DbaseDate1 < DbaseDate2) {

						If this.debug
							SciTEOutput("   " A_ThisFunc ": Behandlungszeitraum gefunden! [" saveDate "]")

					  ; Objekt wird zum Ende hin umgebaut. Da gefunden wurde was gesucht wurde, kann hier das Datum zurück gegeben werden
						retObj := Object()
						retObject["Behandlung"] := [saveDate, DDatum1, DDatum2]

					return saveDate ; retObj
					}

				}

			}

		; Erstellungsdatum des Dokumentes
			For RxStrIndex, rxDateStr in this.rxDokDatum {

				spos := 1
				while (spos := RegExMatch(workText, rxDateStr, D, spos)) {

					spos += StrLen(D)

					DDatum1	:= RegExReplace(DDatum1, "[\.\,\;\:\s]+", ".")
					DDatum1	:= RegExReplace(DDatum1, "\.{2,}", ".")
					DDatum1 := this._FormatDateEx(DDatum1, "DMY", "dd.MM.yyyy")
					If !(DDatum1 := DateValidator(DDatum1, A_YYYY))
						continue

				; prüft ob gefundene Tage Geburtstage sind
					PatID1 := cPat.PatID([{"key":"GEBURT", "value":this._ConvertToDBASEDate(DDatum1)}])
					If this.debug
						SciTEOutput(" PatID1: " (!PatID1.Count()?0:PatID1.Count()))
					If (PatID1.Count()>0) {
						If this.debug
							SciTEOutput(" D1: " DDatum1 "  [(" RxStrIndex ") " rxDateStr "] matched with birthday(s)"  this._ConvertToDBASEDate(DDatum1))
						continue
					}

				; Datumstring(s) speichern
					If !DocDates.Dokument.HasKey(DDatum1) {
						DocDates.Dokument[DDatum1] := {"spos":[spos], "dcount":1}
					}
					else {
						For DDIdx, datepos in DocDates.Dokument[DDatum1]
							If (datepos = (spos-StrLen(D)))
								continue
						savedLNr := LNr
						DocDates.Dokument[DDatum1].spos.Push(spos)
						DocDates.Dokument[DDatum1].dcount += 1
					}
				}
			}

			If (this.debug && IsObject(DocDates.Dokument) && DocDates.Dokument.Count() > 0)
				SciTEOutput("DocDates.Dokument:`n" (IsObject(DocDates.Dokument) ? cJSON.Dump(DocDates.Dokument, 1) : "") "`n")

		; es wurde kein Behandlungs- noch Erstellungsdatum gefunden
			If (DocDates.Behandlung.Count() = 0) && (DocDates.Dokument.Count() = 0) {

				DocPages	:= FindDocPages(workText)
				PageNr 	:= 1

				If this.debugg
					SciTEOutput("  keine Daten gefunden`n  Seiten: " DocPages)

				For LNr, line in TLines {

					; eine Seite weiter
						If RegExMatch(line, "[\f]") {
							PageNr ++
							continue
						}

					; einzelnes Datum wird auf seine Position innerhalb des Dokumentes geprüft
						If RegExMatch(line, "(?<Datum1>\d{2}\.\d{2}\.(\d{4}|\d{2})|" rxDateLong ")", D) {

							; Datum muss sich im Kopf- oder Fußzeilenbereich befinden
								If (LNr <= OberesViertel) || (LNr >= UnteresViertel) {

										OberesViertel 	:= Floor(DocPages[PageNr]/5)
										UnteresViertel	:= Floor(DocPages[PageNr] - DocPages[PageNr]/5)
										saveDate      	:= DateValidator(DDatum1, A_YYYY)

									; Geburtstage aussortieren
										For PatID, Pat in names {
											If (Pat.Gd = DDatum1) {
												DDatum1 := ""
												break
											}
										}
										If (StrLen(DDatum1) = 0)
											continue

									; Datum hinzufügen
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

			If this.debug && (DocDates.Dokument.Count() <> retDates.Dokument.Count())
				SciTEOutput("  " (DocDates.Dokument.Count() - retDates.Dokument.Count()) " Erstellungsdaten wurde(n) aussortiert!")

		;}

		; Rückgabe Objekt: alle Daten unter maximaler Häufigkeit aussortieren [Dokument] ;{
			maxdcount := 0
			For KeyDate, docdate in DocDates.Behandlung
				maxdcount := docdate.dcount > maxdcount ? docdate.dcount : maxdcount

			For KeyDate, docdate in DocDates.Behandlung
				If (docdate.dcount = maxdcount)
					retDates.Behandlung.Push(KeyDate)

			If this.debug && (DocDates.Behandlung.Count() <> retDates.Behandlung.Count())
				SciTEOutput("  " (DocDates.Behandlung.Count() - retDates.Behandlung.Count()) " Behandlungstage wurde(n) aussortiert!")

		;}


	return retDates
	}

	FindDocInhalt(Text)                                                         	{                	;-- Kategorisierung/Autonaming

			description	:= Object()
			workdescr 	:= Array()
			bezeichner  	:= FileOpen(Addendum.DBPath "\Dictionary\Befundbezeichner.txt", "r", "UTF-8").Read()
			bezeichner 	:= RegExReplace(bezeichner, "[\n\r]+([\n\r]+)", "$1")             ; leere Zeilen werden entfernt
			descriptions  	:= StrSplit(bezeichner, "`n", "`r")

		; RegExReplace räumt auf und erstellt ein Arbeitsobjekt der Beschreibungen
			For dNR, line in descriptions {

				line := RegExReplace(line, "Dr\.\s+", "Dr.")
				line := RegExReplace(line, "([^\w])v\.", "$1")
				line := RegExReplace(line, "[\(\)\!\~\$]", "")
				line := RegExReplace(line, "\d[\-\/]\d+", "")
				line := RegExReplace(line, "([A-Z])([A-ZÄÖÜ][a-z]{3,})", "$1 $2")
				line := RegExReplace(line, "\s{2,}", " ")
				descriptions[dNR] := line

			; Wörter/Zahlen entfernen welche eher nicht im Text vorkommen werden
				workline := RegExReplace(line     	, "(^|\s)\d+(\s|$)", " ")                     	; einzeln stehende Zahlen
				workline := RegExReplace(workline, "(^|\s*)\d+\.\d+\.\d+(\s|$)*", " ") 	; Datumzahlen
				workline := RegExReplace(workline, "([^\w])\-", "$1 ")                            	; einzeln stehende Minus Symbol
				workline := RegExReplace(workline, "[\+\,]", " ")                                     	; Plus und Komma  immer
				workline := RegExReplace(workline, "(^|\s)[a-zäöü]+", " ")                 	; Worte mit kleinem Anfangsbuchstaben  immer
				workline := RegExReplace(workline, "(^|\s)v\.(\s|$)", " ")                       	; 'v.'
				workline := RegExReplace(workline, "\s{2,}", " ")
				workdescr[dNR] := workline

			}

		; Bezeichnungen auftrennen nach Wörtern und Wortpositionen - Reihenfolge und Häufigkeit erfassen
			For dNr, line in workdescr {

				For wordPos, word in StrSplit(line, " ") {

					If (StrLen(word) = 0 || RegExMatch(word, "^(\-|[\d\.]+|\w+\.[a-z]+)$") || RegExMatch(word, "^[a-zäöü]"))
						continue

					If (wordPos = 1)  {
						mainword := word
						If !IsObject(description[word])
							description[word] := {"__freq":1, "dNr":dNR}
						else
							description[word].__freq +=1
					}
					else if (wordPos > 1) {
						If !description[mainword].hasKey(word)
							description[mainword][word] := {"__freq":1, "pos":wordPos, "dNr":dNR}
						else
							description[mainword][word].__freq += 1

					}

				}

			}

		; Inhaltserkennung
		; Schritt 1: grobe Zusammenstellung (ein Treffer reicht)
			matches := {}
			For firstword, nextwords in description {

				submatches := {}
				If InStr(Text, firstword) {

					If !matches.hasKey("_" description[firstword].dNR)
						matches["_" description[firstword].dNR] := 1
					else
						matches["_" description[firstword].dNR] += 1

					For subword, obj in nextwords {

						If !IsObject(obj)                                 ; Schlüssel wie __freq, dNR
							continue
						If InStr(Text, subword)
							If !matches.hasKey("_" obj.dNR)
								matches["_" obj.dNR] := 1
							else
								matches["_" obj.dNR] += 1
					}

				}

			}

		; Schritt 2: feinere Zusammenstellung (Vergleich der Wortreihenfolge in Text und Beschreibung)
			validMatches := {}
			For matchNr, freq in matches {

				dNR          	:= StrReplace(matchNr, "_")
				words         	:= StrSplit(workdescr[dNR], " ")
				wordMatches	:= 0
				posMatches  	:= 0
				matchPos  	:= []

			; Schritt 2a: ermittelt die Reihenfolge der Wörter im Text
				For wPos, word in words {

					If (TextPos := InStr(Text, word)) {

						If (wPos = 1) {                                     ; erstes Wort ist die Hauptkategorie und wird als erste Position gesichert
							wordMatches	+= 	1
							posMatches	+= 	1
							matchPos[1]      	:= word
						} else {
							wordMatches	+= 	1
							matchPos[Textpos] := word
						}

					}

				}

			; Schritt 2b: 	vergleicht nun die Reihenfolge der Wörter im Text mit denen in der Beschreibung
			;                	und zählt die Treffer für diese Bedingung
				wPos := 0
				breakFor := false
				For TextPos, TextWord in matchPos {

					wPos ++
					If (wPos = 1)
						continue
					else if (wPos > words.MaxIndex())
						break

					Loop  {

						If (TextWord = words[wPos]) {
							posMatches += 1
							break
						}

						wPos ++
						if (wPos > words.MaxIndex()) {
							breakFor := true
							break
						}

					}

					If breakFor
						break

				}

				validMatches[matchNR] := {"words": words.MaxIndex(), "wordMatches":wordMatches, "posMatches":posMatches, "description": descriptions[dNR], "matchpos":matchpos}
			}

			;~ JSONData.Save(A_Temp "\tmp1.json", description	, true,, 1, "UTF-8")
			JSONData.Save(A_Temp "\tmp2.json", matches      	, true,, 1, "UTF-8")
			JSONData.Save(A_Temp "\tmp3.json", validMatches	, true,, 1, "UTF-8")

			;~ Run % A_Temp "\tmp1.json"
			Run % A_Temp "\tmp2.json"
			Run % A_Temp "\tmp3.json"

	}

	FindDocSender(Text) 															{               	;-- Absender erkennen (funktioniert nicht)

		Sender := []
		; Dr.Kardiotoid, medilog, Darwin, professional. 		    	- LZ-EKG
		; Arztanfrage, (Muster 52), 											- Muster 52
		Sender.Push(["A,Landesamt,Soziales,Versorgung"																																, "LASV"
								, "A,Abgabe,Befundberichtes"																																		, "Errinnerung"])
			Sender.Push(["3,Entlassungsbrief,Anamnese,Diagnose,Aufnahmestatus,EKG,Chefarzt,Oberarzt,Facharzt,Stationsarzt,Assistenzarzt" 	, "Epikrise"])

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

	FindDocEmpfehlung(Text)                                             		{              	;-- Termine, Behandlungsempfehlungen extrahieren

		/*
		Weiterbehandlung:
		Wir empfehlen
		 */
		 static rxEmpfehlung := ["mi)Weiterbehandlung\s*\:*(?<Empfehlung>.+)?(Seite\s*\d*|[\p{L}\-])+\s*\:\s*[\n\r+])"
											,"mi)([\n\r]+|\s)(wir\s)*empfehlen(\swir)*\s*\:*(?<Empfehlung>.+)?(Seite\s*\d*|[\p{L}\-]+\s*\:\s*[\n\r+])"
											,"i)\.[\pL\s]+ kann abgesetzt werden."
											,"i)ASS dauerhaft weiter."
											,"i)Kontrolle.*?[\pL\s]+(in|am|zum)\s(\d+\s+(bis|\-)\s+\d+\s+(Wochen|Monaten)|\d{1,2}\.\d{1,2}\.\d{2,4})"] ; Kontrolle Lipdstoffwechselwerte und ALAT in 4 bis 8 Wochen.

		For rxNR, rxStr in rxEmpfehlung
			If RegExMatch(Text, rxStr, Text) {
				TextEmpfehlung := RegExReplace(TextEmpfehlung, "[\n\r]+\s*[\n\r]+", "`n")
				TextEmpfehlung := RegExReplace(TextEmpfehlung, "[\n\r]+", "`n")
				break
			}

		If (StrLen(TextEmpfehlung) > 0) {

			tmpE := RegExReplace(TextEmpfehlung	, "[\n\r\f]", " ")
			tmpE := RegExReplace(tmp                  	, "\s{2,}", " ")
			tmpL := tmpP := ""
			For idx, word in StrSplit(tmpE	, " ") {
				If (StrLen(tmpL " " word) <= 70)
					tmpL .= " " word
				else {
					tmpP .= "`n    " tmpL
					tmpL := ""
				}
			}

			If this.debug
				SciTEOutput("  - Empfehlung gefunden:`n" tmpP)
		}

	return TextEmpfehlung
	}

	FindDocAlarm(Text)                                                        	{               	;-- Text nach aufälligen Befunden durchsuchen

	  ; Für bessere Lesbarkeit die Leerzeichen lassen. Aus jedem Leerzeichen wird \s+ gemacht.
		phrases =
		(LTrim
		(neu aufgetretener)* Pleuraerguss (re\.|rechts|li\.|links|im|in) *
		(entzündliche|Peribronchiale) \pL* Zeichnungsvermehrung
		Verdacht auf ([\pL\-]+)
		Höhenminderung
		keine Fahreignung
		)

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

	; ## warum Addendum_Datum.ahk nicht erreichbar ist für Funktionen aus dieser Bibliothek kann ich mir nicht erklären
	_FormatDateEx(datestr, dateformat:="DMY", returnformat:="dd.MM.yyyy") {          	;-- formatiert ein Datum in der gewünschten Folge samt Interpunktion

			static rxMDate := {"D"	: "(?<D>\d{1,2})"
									, 	"M"	: "(?<M>\d{1,2})"
									, 	"Y"	: "(?<Y>\d{1,4})"}

		; Delimeter erkennen z.B. ein Punkt oder minus
			RegExMatch(datestr, "O)\w+(\D)\w+(\D)\w", dmtr)
			For idx, interpunctation in dmtr
				datestr := StrReplace(datestr, interpunctuation)

		; dateformat korrigieren
			dateformat := RegExReplace(dateformat, "i)[D]"	, "D")
			dateformat := RegExReplace(dateformat, "i)[M]"	, "M")
			dateformat := RegExReplace(dateformat, "i)[Y]" 	, "Y")

		; RegExMatch String anhand dateformat zusammenstellen
			TF	:= Object()
			TF[InStr(dateformat, "D")]	:= "D"
			TF[InStr(dateformat, "M")]	:= "M"
			TF[InStr(dateformat, "Y")]	:= "Y"

			rxMatch := rxMDate[TF.1] ".*?" rxMDate[TF.2] ".*?" rxMDate[TF.3]

		; inkorrekte Jahreszahl - umwandeln abbrechen
			If !RegExMatch(datestr, rxMatch, T) || (StrLen(TY) = 1) || (StrLen(TY) = 3)
				return  ; "wrong dateformat"

		; Anzahl der Ziffern des Jahres in returnformat berichtigen
			RegExReplace(returnformat, "i)y", "", YearDigits)
			If !(YearDigits = 4)
				returnformat := RegExReplace(returnformat, "i)y+", "yyyy")

		; Eingabedatum formatieren
			TD 	:= SubStr("0" TD	, -1)
			TM 	:= SubStr("0" TM	, -1)
			If (StrLen(TY) = 2) {   ; wendet sich gegen Datumszahlen in der Zukunft!
				YD	:= SubStr(A_YYYY, 3, 2) + 0
				YT	:= SubStr(A_YYYY, 1, 2) + 0
				TY 	:= (TY > YD ? YT-1 : YT) . TY
			}

		; Zeitstring formatieren
			FormatTime, formatedDate, % TY TM TD "000000", % returnformat

	return formatedDate
	}

	_ConvertToDBASEDate(Date) {                                                                             	;-- Datumskonvertierung von DD.MM.YYYY nach YYYYMMDD
		RegExMatch(Date, "((?<Y1>\d{4})|(?<D1>\d{1,2})).(?<M>\d+).((?<Y2>\d{4})|(?<D2>\d{1,2}))", t)
	return (tY1?tY1:tY2) . SubStr("00" tM, -1) . SubStr("00" (tD1?tD1:tD2), -1)
	}

	_PrintRegExStrings() {

		rxVarNames := ["rxNames", "rxVNR", "rxDates", "rxContinue", "rxTags", "rxBehandlung" , "rxDokDatum"]

		SciTEOutput("rxNames:" this.rxNames.Count() "`n")
		For each, rxString in this.rxNames
			SciTEOutput(each ": " rxString)
		SciTEOutput("`n")

		SciTEOutput("rxVNR:`n")
		For each, rxString in this.rxVNR
			SciTEOutput(each ": " rxString)
		SciTEOutput("`n")

		SciTEOutput("rxDates:`n")
		For each, rxString in this.rxDates
			SciTEOutput(each ": " rxString)
		SciTEOutput("`n")

		SciTEOutput("rxContinue:`n")
		For each, rxString in this.rxContinue
			SciTEOutput(each ": " rxString)
		SciTEOutput("`n")

		SciTEOutput("rxTags:`n")
		For each, rxString in this.rxTags
			SciTEOutput(each ": " rxString)
		SciTEOutput("`n")

		SciTEOutput("rxBehandlung:" this.rxBehandlung.Count() "`n")
		For each, rxString in this.rxBehandlung
			SciTEOutput(each ": " rxString)
		SciTEOutput("`n")

		SciTEOutput("rxDokDatum:`n")
		For each, rxString in this.rxDokDatum
			SciTEOutput(each ": " rxString)
		SciTEOutput("`n")

	}

}
