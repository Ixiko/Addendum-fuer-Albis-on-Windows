; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                         	** Addendum_Autonaming **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Beschreibung:	    	Funktionen für die Textklassifizierungen von Dokumenten
;
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;       Abhängigkeiten:
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_Calc started:    	29.09.2022
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
return

class autonamer {

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; statische Variablen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	static rplwrds := [  "alt[e]*", "abgerufen[emnrs]*", "(ä|ae)hnlich[emnrst]*", "allgemein[emnrst]*", "anerkannt[emnrs]*", "(einzig|un)*artig[emnrs]*"
								, "(ä|ae)u(ss|ß)erst[enmrs]*"
								, "ausnahmslos[emnrs]*", "andauernd[emnrs]*", "anderweitig[emnrs]*", "ander[emnrs]+", "anerkannt[emnrs]+", "angesetzt[emnrs]+"
								, "augenscheinlich(st)*[emnrs]*", "ausdr(ü|ue)ck[ent]+", "ausdr(ü|ue)cklich[emnrs]*", "ausgenommen[emnrs]*", "ausgerechnet[emnrs]*"
								, "ausnahmlos[emnrs]*", "(ä|ae)usserst[emnrs]*"
								, "bearbeite[enst]*", "beide[nrs]*", "besser[emnrs]*", "bekanntlich[emnrs]*", "betr(ä|ae)chtlich[emnrs]*", "bisherig[emnrs]*", "denkbar[emnrs]*"
								, "befragt[emnrs]*", "bestimmt[emnrs]*", "betreffend[emnrs]*"
								, "damalig[emnrs]*", "dank[en]+", "(der)*artig[emnrs]*", "derzeitig[emnrs]*", "diesseitig[emnrs]*", "direkt[emnrs]*", "dies[emnrs]*"
								, "dahingehend[emnrs]*", "dein[emnrs]*", "demgegen(ue|ü)ber", "demgem(ae|ä)ss", "d(ü|ue|u)rf(te|e)[nst]*"
								, "eigentlich[emnrs]*", "(aber|ein|drei|mehr|zwei)(seit|mal)ig[emnrs]*", "entsprechend[emnrs]*", "erg(ä|ae)nzen[demnrs]*"
								, "ehe[emnrst]+", "erst[emnrs]*", "einf(ü|ue)hr[ent]+"
								, "(er)*öffne[emnrst]*", "etlich[emnrs]+"
								, "(fort|immer)w(ä|ae)hrend[emnrs]*", "finde[nst]+", "folgen[edmnr]*", "forder[ednrst]+", "fortsetz[ent]+"
								, "g(ä|ae)ngig[emnrs]*", "g(ä|ae)nzlich[emnrs]*", "geehrt[emnrs]*", "genommen[emnrs]*", "gepriesen[emnrs]*", "(un)*gleichzeitig[enrs]*"
								, "gegen(ü|ue)ber", "gerechnet[emnrs]*", "(un)geteilt[emnrs]*", "ganz[emnrs]*", "gratulier[emnrst]+", "gewiss[emnrs]*", "gr(ue|ü)ndlich"
								, "h(ä|ae)ufig[emnrs]*", "heutig[emnrs]*", "hiesig[emnrs]*", "h(ae|a|ä)tt[enst]{0,3}"
								, "ih[ersmnt]*(wegen)*", "insgeheim[emnrs]*", "insgesamt[emnrs]*", "irgendjemand[emnrs]*"
								, "irgend(eine*r*n*m*s*|etwas|wann|was|wer|wo|wie|wen|wohin|welche|wenn)*"
								, "j(ä|ae)hrig[emnrs]*", "jeglich[emnrs]+", "jene[mnrs]*", "jenseitig[emnrs]*", "junge[mnrs]+"
								, "klar[emnrs]*", "klein[emnrs]*", "k(oe|ö|o)nnt[enst]*", "konkret[emnrs]*", "k(ü|ue)nftig[emnrs]*", "k(ü|ue)rzlich[emnrs]*"
								, "letzt[emnrs]*", "(letzt|schlu(ss|ß))*endlich[emnrs]*"
								, "neuerlich[emnrs]*", "neu[emnrs]*", "nach"
								, "meist[emn]*"
								, "oberst[emnrs]+", "offenkundig[emnrs]*", "offensichtlich[emnrs]*"
								, "(un)*pers(ö|oe)nlich[emnrs]*", "pl(ö|oe)tzlich[emnrs]*"
								, "reagier[ernt]*", "regelm(ä|ae)ßig[emnrs]*", "reichlich[emnrs]*", "restlos[emnrs]*", "(richtig|durch|lang)gehend[emnrs]*"
								, "schätz[enst]+", "senk[enst]+", "start[enst]+", "stellenweise[mn]*"
								, "schwerlich[emnrs]*", "selbstredend[emnrs]*", "(ge)*setz[enst]*", "sicher(lich)*[emnrs]*", "solch[emnrs]*", "sonstig[emnrs]*", "soll[enst]*"
								, "tats(ä|ae)chlich[emnrs]*", "teil[ten]+"
								, "unbedingt[enrst]*", "(un)*gleich[emnrst]*", "(un)*ma(ss|ß)geb(end|lich)[emnrs]*", "unsagbar[emnrs]*"
								, "uns(ä|ae)glich[emnrs]*", "unser[emnrs]*", "unerh(ö|oe)rt[emnrs]*", "(un|wo)*(gew|m)(ö|oe)(g|hn)lich[emnrs]*"
								, "umst(ae|ä)ndehalber"
								, "unsagbar[emnrs]*", "uns(ä|ae)glich[emnrs]*", "unser[emnrs]*", "(un)*streitig[enrs]*", "unter[emnrs]*", "(un)*zweifelhaft[emnrs]*"
								, "vergangen[emnrs]*", "vermutlich[emnrs]*", "veröffentlicht[emnrs]*", "vielmalig[emnrs]*", "v(ö|oe)llig[emnrs]*", "vorherig[emnrs]*"
								, "vollst(ä|ae)ndig[emnrs]*", "weitgehend[emnrs]*", "welche[mnrs]*", "weiter[hiermns]*", "weitestgehend[emnrs]*"
								, "(un)*wesentlich[emnrs]*", "wirklich[emnrs]*", "wohlweislich[emnrs]*", "wo(für|fuer|gegen|durch|mit|raus|raufhin|her|hingegen)*"
								, "zahlreich[emrns]*", "zeitweise[emrns]*", "(zweifels)*frei[emrns]*", "ziemlich[emrns]*"
								, "m(ö|oe)glichst[emnrs]*"
								, "jemand[emnrs]*", "manch[emnrs]+", "eur[emnrs]+", "alleine*", "viel[ens]*", "wegen", "erneute*", "V\.a\.", "d\.", "bds\.", "beidseits"]

		static stopwords := Array()
	  ;}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Autonaming Objekt
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	__New(classifierfilepath:="", workpath:="", options:="" ) {

		/* Parameter

			workpath   		.pdfdatapath:	der Pfad zur Metadatei der PDF-Dateien im Befundordner (erstellt Addendum_Gui.ahk)
									.pdfpath:      	Dateipfad mit den Dokumenten

			Beispiel:
			classifier := new autonamer(Addendum.DBPath "\Dictionary"
													, {"pdfpath":Addendum.BefundOrder, "pdfdatapath":Addendum.DBPath "\sonstiges\PdfDaten.json"}
													,  "Debug=true ShowCleanupedText=true RemoveStopWords=true Save_immediately=false")

		*/

	  ; cJSON speichert den Text als UTF-8 ohne HTML Escaping (nannte man es so?))
		cJSON.EscapeUnicode := "UTF-8"

	  ; Parse Options
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		options         	.= !InStr(options, "Debug")              	? " Debug=false"                	: ""
		options         	.= !InStr(options, "removeStopwords")	? " RemoveStopwords=true" 	: ""
		this.options    	:= options
		this.debug     	:= this.optParser("Debug"                     	, false	, true)                 ; option, defaultvalue, is bool?
		this.shwclean 	:= this.optParser("ShowCleanupedText"	, false	, true)
		this.rmwords		:= this.optParser("RemoveStopwords"   	, true	, true)
		this.instant    	:= this.optParser("Save_immediately"  	, true	, true)

	  ; Dokumentverzeichnis und Addendum Metadaten-Datei (PdfDaten.json)
		If IsObject(workpath) {
			SciTEOutput(A_ThisFunc ": " workpath.pdfpath " = " FileExist(workpath.pdfpath)	)
			this.pdfpath  	:= InStr(FileExist(workpath.pdfpath), "D")	? workpath.pdfpath : ""
			pdfdatapath	:= workpath.pdfdatapath
			If !FileExist(pdfdatapath) {
				SplitPath, pdfdatapath,, pdfdataOutDir
				If !InStr(FileExist(pdfdataOutDir), "D")
					throw "Es wird ein Datensicherungspfad benötigt. Der übergebene Pfad ist nicht vorhanden.`n<" pdfdataOutDir ">"
				workpath.pdfdatapath := pdfdataOutDir "\PdfDaten.json"
			}
			this.pdfdatapath	:= workpath.pdfdatapath
		}

		this.Debug ? SciTEOutput( " ["A_ThisFunc "]`n  1. options = " options "`n  2. this.options = " this.options "`n") : ""
		;}

	  ; Stopwörter laden und vorbereiten
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		If !stopwords.Count() {                                                       	; wird nur einmal beim Aufruf der ersten Instanz ausgeführt

			If FileExist(Addendum.Dir "\include\Daten\" (stopwordsFilename := "german_stopwords_full.txt"))
				stopwordstxt := FileOpen(Addendum.Dir "\include\Daten\" stopwordsFilename, "r", "UTF-8").Read()
			else {

				TrayTip, Download..., Lade Stoppwortliste in Deutsch von Github...., 10
				stopwordstxt := this.URLDownloadToVar("https://raw.githubusercontent.com/solariz/german_stopwords/master/" stopwordsFilename)
				FileOpen(Addendum.Dir "\include\Daten\" stopwordsFilename, "w", "UTF-8").Write(stopwordstxt)
				TrayTip, Download..., Stoppwortliste wurde geladen, 4

			}

		  ; Wort-Daten aus geladener Textdatei herausfiltern
			stopwordstmp := "|"
			For each, line in (TextStopwords := StrSplit(stopwordstxt, "`n", "`r"))
				stopwordstmp .= (!RegExMatch(line, "^\s*;") ? line "|" : "" )

		  ; Wörter entfernen
			For each, rxrplStr in autonamer.rplwrds
				stopwordstmp := RegExReplace(stopwordstmp, "i)\|" rxrplStr "(\b|$)", "")
			stopwordstmp := RegExReplace(stopwordstmp, "\|{2,}", "|")
			stopwordstmp := Trim(stopwordstmp, "|")

		  ; Stopwörter zusammenfügen
			For each, word in StrSplit(stopwordstmp, "|") {
				autonamer.stopwords.Push(word)
				wordsall .= word "|"
			}
			For each, word in autonamer.rplwrds {
				autonamer.stopwords.Push(word)
				wordsall .= word "|"
			}

		  ; Debug-Ausgabe
			If (this.debug & 2) {
				SciTEOutput("Länge zuvor: " StrLen(stopwordstxt))
				SciTEOutput("Länge danach: " StrLen(stopwordstmp))
				SciTEOutput("Anzahl Stopwörter zuvor: " TextStopwords.Count() )
				SciTEOutput("Anzahl Stopwörter nach umwandeln: " autonamer.stopwords.Count())
				SciTEOutput(wordsall "`n`n")
			}

		}
		;}

	  ; Dateibezeichnungen laden
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		this.classifierfilepath := classifierfilepath "\Befundbezeichner.json"
		this.autonames := this.load()

	}

	__Delete() {

		;~ res := this.save()
		;~ SciTEOutput("Klassifizierungen gespeichert: " res)

	return res
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Klassifizierdaten
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	load()                                                                                          	{ ; lädt Klassifizierungen

	  ; dieses Objekt hält die letzte Aktualisierungszeit der Klassifizierungsdaten vor
		If !IsObject(this.classifierfile) {

			fE := FileExist(this.classifierfilepath) ? true : false
			cfpath := this.classifierfilepath
			SplitPath, cfpath,, fDir, fExt, fNoExt
			FileGetTime, ftime, % this.classifierfilepath, M
			this.classifierfile := {"name"	: fE ? fNoExt                      	: "-"
										, "path" 	: fE ? fDir                           	: "-"
										, "ext"    	: fE ? fExt                            	: "-"
										, "filetime"	: fE ? this.ClassifierFileTime() : 0}
		}

	  ; Klassifizierungsdaten laden
		fileitems 	:= FileOpen(this.classifierfilepath, "r", "UTF-8").Read()
		try {

			autonames := cJSON.Load(fileitems)
			this.classifierfile.filetime := this.ClassifierFileTime()
			this.classifierfile.mainItems := autonames.GetCapacity()
			subItems := 0
			For maintitle, auto in autonames
				subItems += auto.sub.Count()
			this.classifierfile.subItems := subItems

		} catch {
			MsgBox, 0x1000, % A_ScriptName, % "Die Datei mit den Klassifizierungsdaten ist leer, hat kein json-Format`n"
																. 	" oder befindet sich nicht im Ordner: `n"
																.	 "<" this.classifierfile.path ">`n"
																. 	"Eine Klassifizierung von Dokumenten ist zunächst nicht möglich.`n"
																. 	"Die dafür notwendigen Daten lassen sich aber`n"
																.	"beim Umbenennen von Dokumenten mit`n"
																. 	"'Addendum für Albis On Windows' erstellen."
			autonames := Object()
		}

		this.debug ? SciTEOutput(" [" A_ThisFunc "]`n  mains: " this.classifierfile.mainItems ", subs: " this.classifierfile.subItems "`n") : ""

	return autonames
	}

	save()                                                                                        	{ ; speichert Klassifizierungen im json Format

	  ; ohne zu speichende Daten ist nichts zu speichern
		If (!isObject(this.autonames) || !this.autonames.Count())
			return -1

	  ; Daten sichern
		FileOpen(this.classifierfilepath, "w", "UTF-8").Write(cJSON.Dump(this.autonames, 1))
		this.newdata := ""
		EL := ErrorLevel

	  ; Debugging
		If this.debug {
			FileGetSize, fSize, % this.classifierfilepath, K
			res := "ErrorLevel: "    	ErrorLevel
					. ", A_LastError: " 	A_LastError
					. ", Zeichen: "    	StrLen(fileitems)
					. ", Dateigröße: " 	(fSize/1024 > 10 ? Round(fSize/1024, 2) " MB" : fSize " KB")
			SciTEOutput("gesichert nach: " this.classifierfilepath "`n" res)
		}

	return EL
	}

	update()                                                                                     	{ ; Klassifizierungen auffrischen

	  ; Klassifizierungen werden nur bei geändertem Dateiinhalt (Zeitstempelvergleich der letzten Änderung) aufgefrischt
	  ; dadurch muss nicht das komplette Skript neugestartet werden, wenn der Dateiinhalt z.B. durch einen Texteditor verändert wurde
	  ; zurückgeben wird der Änderungsstatus, also entweder original - für unverändert oder update - für verändert

		If !IsObject(this.classifierfile)
			throw A_ThisFunc ": please initialize class first with ' new autonamer(classifierfilepath)' before calling the class method"

		If !FileExist(this.classifierfilepath) {
			SciTEOutput("Datei nicht vorhanden: " classifierfilepath)
			return "original"
		}

		If (this.ClassifierFileTime() > this.classifierfile.filetime)
			this.autonames := this.load()

	return this.autonames.Count() ? "update" : "empty"
	}

	add(mainclass, subclass:="")                                                      	{ ; fügt Daten hinzu

	  ; erstellt main und bei Bedarf subclass, generiert aus mainclass und subclass die ersten Klassifikationswörter
	  ; weitere Worte werden über .setSub() hinzugefügt

		this.newdata := false

	  ; - - - - - - - - - - - - - - - - - - - - - - -
	  ; Kategorie
		If !this.hasMain(mainclass) {

			this.newdata 	:= this.newdata | 1
			mainwords       	:= RegExReplace(mainclass, "\s{2,}", " ")

			this.autonames[mainclass] := {"main":Object(), "sub":Object()} ; erstellt alle Objekte

			For each, mainword in StrSplit(mainwords, A_Space)
				If Trim(mainword)
					this.autonames[mainclass].main.Push(Trim(mainword))
		}

	  ; - - - - - - - - - - - - - - - - - - - - - - -
	  ; Inhalt
		If subclass && !this.hasSub(mainclass, subclass) {

			this.newdata 	:= this.newdata | 2
			subwords       	:= RegExReplace(subclass, "\s{2,}", " ")

			If !IsObject(this.autonames[mainclass].sub)
				this.autonames[mainclass].sub := Object()
			If !IsObject(this.autonames[mainclass].sub[subclass])
				this.autonames[mainclass].sub[subclass] := Object()

			For each, subword in StrSplit(subwords, A_Space)
				If Trim(subword)
					this.autonames[mainclass].sub[subclass].Push(Trim(subword))

		}

	  ; - - - - - - - - - - - - - - - - - - - - - - -
	  ; Daten speichern
		If this.instant && this.newdata
			this.save()

	return this.newdata	; 1 für mainclass, 2 für subclass , 3 für main- und subclass
	}

	getMains(matchstring:="", ret:="arr")                                            	{ ; alle oder bestimmte Kategorien zurückgeben

		; gibt die Hauptklassifikationen entweder als Array oder als durch ein mit "|" getrennte Liste zurück
		; die Listenausgabe vereinfacht die Handhabung von Combo- und Listboxen, sowie von DropDownListen

		If (!isObject(this.autonames) || this.autonames.Count()=0)
			return

		maintitles := ret = "arr" ? Array() : ret~="i)^[a-zäöüß\d]+$" ? "|" : ret

	 ; gibt nur Kategorien zurück, die den matchstring enthalten
		If matchstring
			For maintitle, sub in this.autonames
				If Instr(maintitle, matchstring)
					If (ret = "arr")
						maintitles.Push(maintitle)
					else
						maintitles .= ret . maintitle

	 ; alle Kategorien zurückgeben, wenn matchstring leer war oder keine Übereinstimmungen gefunden  wurden
		If !matchstring
			For maintitle, sub in this.autonames
				If (ret = "arr")
					maintitles.Push(maintitle)
				else
					maintitles .= ret . maintitle

	return IsObject(maintitles) ? maintitles : LTrim(maintitles, ret)
	}

	getMainwords(mainclass, ret:="arr")                                          		{ ; gibt die klassifizierenden Hauptworte zurück

		; ret := "`n" gibt einen String mit diesem Listenteiler zurück

		If (!isObject(this.autonames) || this.autonames.Count()=0)
			return -1
		If !this.hasMain(mainclass)
			return -2

		If (ret = "arr")
			return this.autonames[mainclass].main

		wordlist := ""
		ret := !ret || ret~="i)^[a-zäöüß\d]+$" ? "|" : ret
		For mIndex, mainword in this.autonames[mainclass].main
			wordlist .= (mIndex>1 ? ret : "") mainword

	return wordlist
	}

	getSubs(mainclass, subclass:="", ret:="arr")                                   	{ ; alle oder einen Teil der Subklassifizierungsdaten für eine Kategorie zurückgeben

	  ; subclass leer lassen um alle Subklassifizierungsobjekte zu erhalten
	  ; ein spezifiziertes subclass als String gibt nur die zugehörigen  Subklassifizierungsworte zurück

		If (!isObject(this.autonames) || this.autonames.Count()=0)
			return -1
		If !this.hasMain(mainclass)
			return -2

	  ; gibt je nachdem ein Array oder einen String zurück (der String ist für Combox, Listboxen oder DropDownListen)
		If (ret = "arr" || subclass)
			return subclass && IsObject(this.autonames[mainclass].sub[subclass]) ? this.autonames[mainclass].sub[subclass] : this.autonames[mainclass].sub

		ret := !ret || ret~="i)^[a-zäöüß\d]+$" ? "|" : ret
		subtitles := ""
		For subtitle, sub in this.autonames[mainclass].sub
			subtitles .= ret . subtitle
		return LTrim(subtitles, ret)
	}

	getSubwords(mainclass, subclass, ret:="arr")                            		{ ; gibt die subklassifizierenden Worte zurück

		; ret := "`n" gibt einen String mit diesem Listenteiler zurück

		If (!isObject(this.autonames) || this.autonames.Count()=0)
			return -1
		else If !this.hasMain(mainclass)
			return -2
		else if !this.hasSub(mainclass, subclass)
			return -3

		If (ret = "arr")
			return this.autonames[mainclass].sub[subclass]

		ret := !ret || ret~="i)^[a-zäöüß\d]+$" ? "|" : ret
		wordlist := ""
		For mIndex, subword in this.autonames[mainclass].sub[subclass]
			wordlist .= (mIndex>1 ? ret : "") subword

	return wordlist
	}

	setSub(mainclass, mainwords, subclass , subwords)                       	{ ; Subklassifizierungsdaten ändern)

	  ; Hauptklasse, klassif. Worte, Subklasse, subklassif. Worte
		static instantState

		instantState := this.instant
		this.instant := false

    ; Subklassifikation bisher unbekannt, dann neu anlegen, subklassifizierende Wörte können danach einfach hinzugefügt werden
	; das Anlegen einer neuer Haupt- und/oder Subklasse muss durch einen .add() Aufruf erfolgen
	; .add() gibt bei Mißerfolg false zurück, so daße
		If !IsObject(this.autonames[mainclass].sub[subclass])
			If (this.add(mainclass, subclass) = 0) {
				this.instant := instantState
				return -2
			}

		this.instant := instantState

	; Haupt- und Subklassifikation ist angelegt, hunzufügen bisher unbekannte subklassifizierender Worte
	; - - - - main - - - -
		mainwords := !IsObject(mainwords) ? StrSplit(mainwords, "`n") : mainwords
		For each, mainword in mainwords {
			mainexists := false
			For idx, mainStored in this.autonames[mainclass].main
				If (mainStored = mainword) {
					mainexists := true
					break
				}
			If !mainexists
				this.autonames[mainclass].main.Push(mainword)
		}

	; - - - - sub - - - -
		subwords := !IsObject(subwords) ? StrSplit(subwords, "`n") : subwords
		For each, subword in subwords {
			subexists := false
			For idx, subStored in this.autonames[mainclass].sub[subclass]
				If (subStored = subword) {
					subexists := true
					break
				}
			If !subexists
				this.autonames[mainclass].sub[subclass].Push(subword)
		}

	; Daten speichern
		If this.instant
			this.save()

	return {"mainwords":mainwords.Count(), "subwords":subwords.Count()}
	}

	hasMain(mainclass:="", mmode:="")                                             	{ ; ist die Kategorie bekannt

		If !mainclass
			return false

	  ; pattern- oder RegEx Matching um Teile einer Bezeichnung zu finden, gibt in diesem Fall die erste gefundene Bezeichnung zurück
 		If (mmode ~= "i)\s*(pattern|RegEx)") {
			For maintitle, data in this.autonames
				If (mmode="pattern" && maintitle = mainclass)
					return maintitle
				else if (mmode="RegEx" && maintitle ~= mainclass)
					return maintitle

			return ""
		}

	return isObject(this.autonames[mainclass])
	}

	hasSub(mainclass:="", subclass:="", mmode:="pattern")                	{ ; Kategorie enthält Subklassifizierung

	  ; für bessere Auswertbarkeit gibt die Funktion neg. Werte als Fehlerwerte zurück
		If !mainclass || !subclass || !this.hasMain(mainclass)
			return !this.hasMain(mainclass) ? false : !mainclass ? -1 : -2

		If (mmode ~= "i)\s*(pattern|RegEx)") {
			For subtitle, data in this.autonames[mainclass].sub
				If (mmode="pattern" && subtitle = subclass)
					return subtitle
				else if (mmode="RegEx" && subtitle ~= subclass)
					return subtitle

		return ""
		}

	return IsObject(this.autonames[mainclass].sub[subclass])
	}

	countSubclassifications()                                                              	{

		counter := 0
		For mainTitle, data in this.autonames {
			For subtitle, sub in data.sub
				counter ++
		}

	return counter + this.autonames.Count()
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Dokumente
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	readDirectory(options:="only_unnamed=true")                            	{ ; liest ein Verzeichnis ein oder nutzt die von AddendumGui erstellten PDF-Daten (.json Format)

	  ; Funktion benötigt Addendum_PdfHelper.ahk und die qpdf.exe, sowie Addendum_DB.ahk (xstring Klasse)

		this.debug ? SciTEOutput(" [" A_ThisFunc "]`n  pdfpath:    " this.pdfpath "`n  pdfdatapath: " this.pdfdatapath) : ""

		If !this.pdfpath || !this.pdfdatapath
			throw A_ThisFunc ": please initialize class first with ' new autonamer(classifierfilepath, workpath)' before calling the class method"

		If !Instr(FileExist(Addendum.Dir "\include\OCR\qpdf"), "D")
			throw A_ThisFunc ": kann das qpdf Verzeichnis nicht finden.`n<<" Addendum.Dir "\include\OCR\qpdf"  ">>"

		only_unnamed := options ~= "i)only_unnamed\s*=\s*true" ? true : false

	  ; Metadatendatei laden oder erstellen
		documents := []
		files := FileOpen(this.pdfdatapath, "r", "UTF-8").Read()
		try
			files := cJSON.Load(files)
		catch {
			this.debug ? SciTEOutput("  cJSON konnte die geladenen PDF-Daten nicht parsen.`n  " this.pdfdatapath) : ""
		}
	  ; vorhandene Datei einlesen
		If IsObject(files) {
			For fNr, file in files
				If FileExist(this.pdfpath "\" file.name) {
					If (only_unnamed && xstring.isFullNamed(file.name))
						continue
					else
						documents.Push(file)
				}
		}
		else {
			Loop, Files, % this.pdfpath "\*.pdf"
			{
				If (only_unnamed && xstring.isFullNamed(A_LoopFileName))
					continue
				FileGetTime	, timestamp	, % A_LoopFileFullPath, M
				FileGetSize	, filesize       	, % A_LoopFileFullPath, K
				documents.Push({	"name"          	: A_LoopFileName
										, 	"isSearchable"	: PDFisSearchable(A_LoopFileFullPath)
										, 	"pages"				: PDFGetPages(A_LoopFileFullPath, Addendum.Dir "\include\OCR\qpdf")
										, 	"timestamp"   	: timestamp
										,	"filesize"           	: filesize})
			}
		}

		this.debug ? SciTEOutput("  gefundene Dokumente (" (only_unnamed ? "nicht " : "") "vollständig benannte): " documents.Count() "`n") : ""

	return documents
	}

	getDocumentText(fpath)                                                             	{ ; PDF/Word Dokumente in Text umwandeln

		If !RegExMatch(fpath, "[A-Z]\:\\") && this.pdfpath
			txt := IFilter(pdfpath := this.pdfpath "\" fpath)
		else if FileExist(fpath)
			txt := IFilter(pdfpath := fpath)

	return txt
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Klassifizierer

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	matchscore(txt, fname:="")                                                          	{ ; der Klassifizierungsalgorithmus

		bestscore := {main:0, sub:0, noscores:0, titles:[]}
		x := "", maxmainscore := 0
		mains := {}

	  ; Text für schnellere Verarbeitungen bereinigen
		txt := " " this.cleanup(txt) " "
		txtlen := StrLen(txt)
		If (this.debug > 1)
			SciTEOutput(A_ThisFunc ": .cleanup() - Text bereinigt")

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ;	Klassifizieren
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Haupt-Klassifizierer
		For maintitle, data in this.autonames {

			 mt := v := ""
			 lastmpos := mainScore := mainScore1 := mainScore2 := mainScore3 := matchedIndex := caseMustMatch := caseMatched := 0

		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		  ; Hauptklassifikation
		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			For mIndex, mainword in data.main {

			  ; Vorbereitungen
				mpos1 := mpos2 := ""
				mt .= (mIndex>1 ? ".*?") mainword
				caseMustMatch += RegExMatch(mainword, "^[A-ZÄÖÜ]") ? 1 : 0

			  ; Anzahl der Worte (EMailadressen, Telefonnummern gelten als ein Wort)
				If !RegExMatch(mainword, "^\s*([\w\.\_\-]+@[\w\.\_\-]+\.\w+|[\d\s\(\)\-]+)\s*") {
					wordstmp := RegExReplace(mainword, "[^\pL\d]", " ")
					wordstmp := RegExReplace(wordtmp, "\s{2,}", " ")
					mwordsCount := StrSplit(wordstmp, A_Space).Count()
				}

			  ; Position innerhalb des Dokumenttextes
				rxmainword := RegExReplace(mainword, "\s{1,}", "\s+")
				rxmainword := RegExReplace(mainword, "\-", "[^\pL\d]+")
				If !(mpos1	:= RegExMatch(txt, "i)\s" rxmainword "\s", rxC1))
					mpos2	:= RegExMatch(txt, "i)\s" rxmainword   	, rxC2)
				If	(mpos := mpos2 ? mpos2 : mpos1) {
					;~ caseMatched   +=	RegExMatch(mainword, "^[A-ZÄÖÜ]") ? 1 : 0
					;~ caseFactor    	:=	RegExMatch(mainword, "^[A-ZÄÖÜ]") ? 2 : 1
					caseMatched   +=	mwordsCount
					caseFactor    	:=	RegExMatch(mainword, "^[A-ZÄÖÜ]") ? 2 : 1
					posFactor      	:= 	Round((10-(mpos/txtlen)*10)/3)
					ppos	            	:= 	posFactor + ((	mIndex=1 					     	? 8
																		: 	mIndex>1 && mIndex<=3 	? 3
																		: 	mIndex>3 					    	? 2 : 1)*caseFactor)
					mainScore1   	+=	ppos
					matchedIndex 	++
					v 	.= mainword "[" posFactor "," ppos "," mainScore1 "],"
				}

			  ; Wort gehört zu einem Satz/Bezeichnung?
				If (mpos && lastmpos < mpos) {
					If lastmpos {
						SubString := SubStr(txt, lastmpos, mpos-lastmpos+1)
						RegExReplace(SubString, "\s+", "", wsCount)
						mainScore2 += wsCount<4 && wsCount>0 ? (4-wsCount)*2 : 0
					}
					lastmpos := mpos > lastmpos ? mpos : 0
				}

			}

		  ; Dokumentenklasse - wird komplett in der Wortreihenfolge gefunden
			If (data.main.Count()>1) {

			  ; in Regex String umformen
				rxmatch    	:= this.RegExTransformGerman(maintitle)
				rxmatch       	:= RegExReplace(Trim(rxmatch), "\s{2,}", " ")   	; doppelte Leerzeichen entfernen
				rxmatchTight := RegExReplace(rxmatch, "\s", "\s+")            		; es wird kein Wort übersprungen, alle müssen exakt treffen!    ➧  ➧  ➧  ➧   !
				rxmatchSkip	:= RegExReplace(rxmatch, "\s", ".*?\s")              	; Worte werden übersprungen. (z.b. Universität ➧ ➧ ? ➧ ➧ Hauptstadt ➤ Universität des Wissens
																											; 	, Wissensplatz 1, 11111 Hauptstadt)

			  ; unterschiedliche Konditionen
				If RegExMatch(txt, rxmatchTight)                  	; exakte Übereinstimmung (Reihenfolge, nur Leerzeichen zwischen den Worten und Buchstabengrößen)
					mainScore3 := 10
				else If RegExMatch(txt, "i)" rxmatchTight)      	; dichte Übereinstimmung (Reihenfolge und nur Leerzeichen zwischen den Worten)
					mainScore3 := 8
				else If RegExMatch(txt, "i)" rxmatchSkip)       	; nahe Übereinstimmung (nur die Reihenfolge der Wörter reicht für einen Treffer)
					mainScore3 := 6

			 ; CaseMatched anpassen
				caseMatched := mainScore3 ? caseMustMatch : caseMatched

			}

			maincaseScore 	:= caseMustMatch=caseMatched ? 5 : 0
			mainScore    	:= mainScore1 + mainScore2 + mainScore3 + maincaseScore

			If mainScore {

				mains[maintitle] := {"mainscore"	: mainScore
											, "main123C"	: mainScore1 "+" mainScore2 "+" mainScore3 "+" maincaseScore
											, "mainwords"	: RTrim(v, ",")
											, "subscore"   	: 0	}

				If (maxMainScore < mainScore) {
					maxMainScore	:= mainScore
					maxmaintitle 	:= maintitle
				}

			} else {
				bestscore.noscores += 1
				continue
			}
			;}

		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		  ; Subklassifikationen durchgehen
		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			maxSubScore := 0, maxSubTitle := maxSubWords := ""

			for subTitle, subwords in data.sub {

				w           	:= ""
				spos     	:= 1
				lastspos	:= subScore := 0
				subwordC := subwords.Count()                                         	; Anzahl der Subklassifizierungen

				For windex, words in subwords {

				  ; in Regex String umformen
					rxmatch := this.RegExTransformGerman(words)                	; words -
					rxmatch := RegExReplace(Trim(rxmatch), "\s{2,}", " ")    		; doppelte Leerzeichen entfernen
					rxmatch := RegExReplace(rxmatch, "\s", "\s+")            		; es wird kein Wort übersprungen, alle müssen exakt treffen
					rxmatch := RegExReplace(rxmatch, "\*", ".*?")               		; * in den Stichworten ist ein Platzhalter, dann Suche bis zum ersten Treffer

					If (npos := RegExMatch(txt, "i)\s" rxmatch)) {

						If (lastspos < npos) {

							SubString := SubStr(txt, lastspos, npos-lastspos+1)
							RegExReplace(SubString, "i)\s+", "", wsCount)
							SubScore1 := wsCount <=3  ? 2 : 1
							wIndexFactor :=(subwordC>1	&& wIndex=1)                                   	? 4
												:	(subwordC>1	&& wIndex>1 && wIndex<subWordC) 	? 3
												:	(subwordC>1	&& wIndex=subWordC)					    	? 2
												:	(subwordC=1	&& wIndex=1)                                   	? 1: 1
							SubScore += SubScore1*wIndexFactor

						}
						else
							SubScore += 1

						spos := npos+StrLen(word)
						lastspos := spos > lastspos ? spos : 0
						w .= words "[" wIndex "/" subwordC "," subScore (wsCount ? wsCount "/" wIndexFactor : "" )  "],"

					}

				}
			;}

			 ; neuer Subscore ist höher als der letzte?
			 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				If (subScore > mains[maintitle].subScore) {
					mains[maintitle].subScore	:= subScore
					mains[maintitle].subTitle 	:= subTitle
					mains[maintitle].subwords 	:= RTrim(w, ", ") ", subclassifier: " data.sub.Count()
				}

			}

		}

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; beste Scores sichern
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		For mTitle, mData in mains {

			mData.subScore := !mData.haskey("subScore") ? 0 : mData.subScore
			thisScore := mData.mainScore + mData.subScore

		; nimmt nimmt nur soviele Titel auf bis keiner der folgenden einen höheren Punktwert erreicht
			If (thisScore > bestscore.max1)   {
				bestscore.titles[thisScore] := {"mainScore"	: mData.mainScore
														, 	"subScore" 	: mData.subScore
														,	"maxScore"	: mData.mainScore + mData.subScore
														, 	"main123C"	: mData.main123C
														, 	"mainwords"	: mData.mainwords
														, 	"subwords"	: mData.subwords
														, 	"mainTitle"  	: mTitle
														, 	"subTitle"    	: mData.subTitle}

				bestscore.max3             	:= bestscore.max2
				bestscore.max2             	:= bestscore.max1
				bestscore.max1               	:= thisScore
				bestscore.maxMainTitle  	:= mTitle
				bestscore.maxSubTitle    	:= mData.subTitle
				bestscore.maxMainScore	:= mData.mainScore
				bestscore.maxsubScore  	:= mData.subScore

			}

		}


		bestscore.txt := txt

		If (this.debug > 1)
			SciTEOutput(cJSON.Dump(mains, 1) "`n`n")

	return bestscore
	}


  ; * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
  ; * * * * * * * * * * * * * * * * * * * * * * * * *     internal    * * * * * * * * * * * * * * * * * * * * * * * * * * *
  ; * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	autoNStrip(item)                                                                           	{ ;-- bereitet Vergleichsdaten auf

		sitem 	:= RegExReplace(item, "\s\-\s*", "###")
		stitle   	:= StrSplit(sitem, "###")

		If !stitle.1
			return

		If !this.autonames.haskey(stitle.1) {
			main := this.removeWords(stitle.1, false)
			mainwords := this.getMatchItems(main)
		}
		else
			mainwords := this.autonames[stitle.1].main

		If !stitle.2
			return {"maintitle": stitle.1, "main" : mainwords}

	  ; Performs word separation in second array, first array will not separated
		sub := this.removeWords(stitle.2, false)
		subwords := this.getMatchItems(sub)

		retObj := {"maintitle": stitle.1, "main": mainwords}
		If stitle.2 {
			retObj.subtitle	:= stitle.2
			retObj.sub     	:= subwords
		}

	return retObj
	}

	removeWords(item, rmdoubles:=false)                                        	{ ;-- bereinigt Klassifikationstexte

		static removers := [{rx: 	"[,;:!\/\(\)\?]"         	, rp: " "          	}
								;, 	{rx: 	"Dr.(\pL)"                 	, rp: " $1"       	}
								, 	{rx: 	"\d+\-\d+"            	, rp: " "          	}
				        		, 	{rx: 	"\d\d\.\d\d\.\d*"   	, rp: " "          	}
								, 	{rx: 	"\d+\."                    	, rp: " "          	}
								, 	{rx: 	"\s\pL(\s|$)"          	, rp: " "          	}
		        				, 	{rx:	"(\pL\.)(\pL)"          	, rp: "$1.\s*$2"	}
								, 	{rx:	"\-(\d+)"                	, rp: "\-*\s*$1"  	}
								, 	{rx:	"(\d+)\-"                	, rp: "$1\-*\s*"	}
								, 	{rx:	"([^\pL])\+([^\pL])"	, rp: "$1 $2"  	}]

									;, {rx: "(\pL)\s+(\d+)\s"	, rp: "$1\s+$2"	}
		item   	:= " " item " "

		If this.rmwords
			For each, rxrpl in this.stopwords
				item := RegExReplace(item, "i)\s" rxrpl "\s", " ")

		For each, remove in removers
			item := RegExReplace(item, remove.rx, remove.rp)

		if rmdoubles
			item 	:= RegExReplace(item, "(?<!\pL|\-|\-\s)\pL{1,3}\.*(\s|$)", " ")               ; entfernt Worte die aus 1 bis max. 3 Zeichen bestehen

		item  	:= RegExReplace(item, "\s{2,}", " ")
		item  	:= Trim(item)

	return item
	}

	getMatchItems(Items)                                                                   	{ ;-- erstellt RegEx Suchstrings

		static replacers := {"(Magnetresonanztomo[ag]ra[fph]+ie|MRT*)\s"                                    	: ""
								, 	"((craniale|kraniale)\s+Magnetresonanztomo[ag]ra[fph]+ie|cMRT*)\s" 	: ""
								, 	"(Computertomographie|CT)\s"                                                          	: ""
								, 	"(Halswirbels(ä|ae)ule|HWS|HWK)\s"                                                  	: ""
								,	"(Brustwirbels(ä|ae)ule|BWS|BWK)\s"                                                    	: ""
								, 	"(Lendenwirbels(ä|ae)ule|LWS|LWK)\s"                                              	: ""
								, 	"(Abdomen|Bauch|Oberbauch|Mittelbauch|Unterbauch)\s"                	: ""
								, 	"(Kopf|Schädel|Cranium|Neurocranium)\s"                                         	: ""
								, 	"(Hüftgelenke*\s|Hüfte\s|Hüft\-)"                                                         	: ""
								, 	"(OP\s|Operation*\s|OP\-)"                                                                 	: ""
								, 	"(TEP|Totalendoprothese|Endoprothese)\s"                                          	: ""
								, 	"(Schultergelenk*|Schulter)\s"                                                              	: ""
								, 	"(Thorax|Brustkorb)\s"                                                                        	: ""
								, 	"(Radiologie|Radiologische Praxis)\s"                                                   	: ""
								, 	"(Daumen|Finger)\s"                                                                          	: ""
								, 	"(Kasse|Krankenkasse|Krankenversicherung|Gesundheitskasse)"          	: ""
								, 	"(DRV|Deutsche\sRentenversicherung)\s"                                              	: ""
								, 	"(OSG|USG|(oberes|unteres)\s+Sprunggelenk)\s"                                	: ""
								, 	"(Ebenen|Ebn\.*|Eb\.*)\s"                                                                   	: ""
								, 	"(re\.|rechts)\s"                                                                                  	: ""
								, 	"(li\.|links)\s"                                                                                     	: ""
								, 	"Rö\s"                                                                                                	: "Röntgen"}


		retItems := Array()
		Items := RegExReplace(Items, "\s{2,}", " ")

	  ; Zahlen als Wort darstellen
		;~ If RegExMatch(Items, "O)\s(?<number>\d+)\s+(?<word>[\pL\-\+°§$']+)", item) {
			;~ ; 105 - einhundertundfünf, 1153 - eintausendeinhundertdreiundfünfzig, 13678 - dreizehntausendsechshundertachtundsiebzig
			;~ d1 := SubStr(item.number, -1)
			;~ Firstnumbs := d1 = 0 ? "" : d1 <= 12 ? numbers.twelves[d1] :
		;~ }

		For each, item in StrSplit(Items, A_Space) {
			item := " " item " "
			For rxstring, replacestring in replacers
				item := RegExReplace(item, "\s" rxstring , (replacestring ? replacestring : rxstring) )
			item := Trim(item, A_Space)
			If item
				retItems.Push(item)
		}

	return retItems
	}

	cleanup(txt)                                                                               	{ ;-- Dokumententext vorbereiten

		; entfernt Stopworte, vereinheitlicht medizinische Bezeichnungen und
		; korrigiert ein paar OCR Zeichenfehler

		static replacers := {"(c|k)raniale\s+Magnetresonanztomo[ag]ra[fph]+ie\s"                      	: "cMRT"
								,	"Magnetresonanztomo[ag]ra[fph]+ie\s"                                              	: "MRT"
								, 	"Computertomogra(ph|f)ie\s"                                                              	: "CT"
								, 	"(Halswirbels(ä|ae)ule|HWK)\s"                                                          	: "HWS"
								,	"(Brustwirbels(ä|ae)ule|BWK)\s"                                                            	: "BWS"
								, 	"(Lendenwirbels(ä|ae)ule|LWK)\s"                                                      	: "LWS"
								;~ , 	"(Abdomen|Bauch|Oberbauch|Mittelbauch|Unterbauch)\s"                	: "Abdomen"
								, 	"(Sch(ä|ae)del|Kalotten*|Cranium|Neurocranium)\s"                          	: "Kopf"
								, 	"(Hüftgelenk(es|s|e)*|Hüfte*)\-*\s"                                                       	: "Hüfte"
								, 	"(OP\s|Operation*\s|OP\-)"                                                                 	: "OP"
								, 	"(TEP|Totalendoprothese|Endoprothese)\s"                                          	: "TEP"
								, 	"Schultergelenk*\s"                                                                             	: "Schulter"
								, 	"Brustkorb\s"                                                                                     	: "Thorax"
								, 	"(Dialysezentrum|Dialysepartner|Nephrologe|Nephrologische*r*s*)\s"   	: "Nephrologie"
								, 	"(Kardiologische Praxis|Kardiolog(en|in|isches|e))"                              	: "Kardiologie"
								, 	"(nuklearmedizin(ische\sPraxis)*|Nuklearmedizin\.)"                            	: "Nuklearmedizin"
								, 	"(Radiologie|Radiologische Praxis|Radiologen|MVZ\s*Diagnostikum|Radiologisches)\s"               	: "Radiologie"
								, 	"Daumen\s"                                                                                         	: "Finger"
								, 	"(Krankenversicherung|Gesundheitskasse|Kasse)"                               	: "Krankenkasse"
								, 	"Deutsche\sRentenversicherung\s"                                                      	: "DRV"
								, 	"obere(s|n)\s+Sprunggelenk(s|es)*\s"                                                 	: "OSG"
								, 	"untere(s|n)\s+Sprunggelenk(s|es)*\s"                                                  	: "USG"
								, 	"(Lungenheilkunde|Lungen-\s+und|Bronchialheilkunde|Lungenarztpraxis|Pneumolog(ie|e|in|en))" : "Pulmologie"
								, 	"(Ebenen|Ebn(\.|\s)|Eb(\.|\s))\s"                                                           	: "Ebenen"
								, 	"re(\.|\s)\s"                                                                                        	: "rechts"
								, 	"li(\.|\s)\s"                                                                                           	: "links"
								, 	"Rö\.*\s"                                                                                               	: "Röntgen"}

		txt := " " txt " "

		If this.rmwords
			For each, rxrplStr in this.stopwords
				txt := RegExReplace(txt, "i)\s" rxrplStr "(\s|$)", " ")

		For rxstring, replacestring in replacers
			txt := RegExReplace(txt, "i)\s" rxstring , replacestring " ")

     ; andere Ersetzungen/Fehlerkorrekturen
	 ; - - - -

	  ; OCR Zeichenfehler korrigieren
		txt := RegExReplace(txt, "([^\pL])(Prof|Priv|Doz|Dr|med|jur|Str|Nr)[,;.]", "$1 $2.")
		txt := RegExReplace(txt, "\s[a-zäöüß]\s", " ")
		;~ txt := RegExReplace(txt, "nfer", "nter")
		txt := RegExReplace(txt, "i)ch[lf]r", "chir")
		txt := RegExReplace(txt, "Giie", "Glied")
		txt := RegExReplace(txt, "arlte", "arite")
		txt := RegExReplace(txt, "i)kardlale", "kardiale")
		txt := RegExReplace(txt, "ı", "i")
		txt := RegExReplace(txt, "Barlin", "Berlin")
		txt := RegExReplace(txt, "hellos", "helios")
		txt := RegExReplace(txt, "i)\bFra[lf]en\b", "Freien")
		txt := RegExReplace(txt, "(D|d)igitate", "$1igitale")
		txt := RegExReplace(txt, "i)Magnetresonanztomo[ag]ra[fph]+ie", "Magnetresonanztomographie")

	  ; Leerzeichen zwischen einem Klein- und Großbuchstaben einfügen
		txt := RegExReplace(txt, "([a-zß\d])([A-ZÄÖÜ][a-zäöüß]+)", "$1 $2")
		txt := RegExReplace(txt, "([a-zß\d][\)\.,;:\!\?]*)([A-ZÄÖÜ][a-zäöüß]+)", "$1 $2")
	  ; alle Zeichen außer die Zeichen in der eckigen Klammer entfernen
		txt := RegExReplace(txt, "[^\pL\d\-\.\/\\\!\?\$%&°§\n]", " ")
	  ; an Bindestrichen angehängte Leerzeichen entfernen
		txt := RegExReplace(txt, "\-\s+", "-")
	  ; Leerzeichen nach einem Punkt (Satzende)
		txt := RegExReplace(txt, "([\pL\d])\.([\pL\d])", "$1. $2  ")
	  ; Wort mit einem Großbuchstaben nach einem anderen Wort abrücken: UniversitätBerlin -> Universität Berlin
		txt := RegExReplace(txt, "([a-zäöüß]{2,})([A-ZÄÖÜ][a-zäöüß])", "$1 $2")
	  ; groß geschriebenes Wort aus mind. 3 aufeinanderfolgenden Großbuchstaben abrücken:  MRTder -> MRT der
		txt := RegExReplace(txt, "([A-ZÄÖÜ]{3,})([a-zäöüß]{2,})", "$1 $2")
	  ; gezielt wissenschaftliche Titel mit einem Punkt versehen
		txt := RegExReplace(txt, "(\S)(Univ|Prof|Dr|med)[\.,;\s]", "$1 $2. ")
	  ; Links/EMail Adressen korrigieren: ,de -> .de
		txt := RegExReplace(txt, "[,;]([a-z]{2,3})", ".$1")
	  ; Komma zu Punkt (Dr, M, Klump > Dr. M. Klump)
		txt := RegExReplace(txt, "([A-ZÄÖÜ][a-zäöüß]{0,1})[\s,,;.]+([A-ZÄÖÜ]{0,1}[a-zäöüß]{0,1})[\s,,;.]+([A-ZÄÖÜ][a-zäöüß]+)", "$1. $2. $3")
      ; fügt 6 oder 8 stellige Zahlenkombinationen welche Leerzeichen enthalten zusammen, damit diese als Datum erkannt werden können
	  ; z.B.: "22. 0 1. 1 979 43 Jahre"  oder "Aufnahme vom Blutdruck 13. 0 7. 2 022 08 44 56 "
		 txt := RegExReplace(txt, "(\d)*\s*(\d)\s*[\.\,]\s*(\d)*\s*(\d)\s*[\.\,]\s*(\d)\s*(\d)\s*(\d)\s*(\d)(?=\D|\pL|\s|$)", "$1$2.$3$4.$5$6$7$8")
		 txt := RegExReplace(txt, "(\d)*\s*(\d)\s*[\.\,]\s*(\d)*\s*(\d)\s*[\.\,]\s*(\d)\s*(\d)\s*(?=\D|\pL|\s|$)", "$1$2.$3$4.$5$6$7$8")

		txt := RegExReplace(txt, "\s{2,}", " ")

		If this.shwclean
			SciTEOutput(A_ThisFunc "`n" txt)

	return txt
	}

	spokennumber(number, uppercase:=true, separator:=" ")            	{ ;-- Darstellung ganzer Zahlen als Zahlwörter (deutsche Sprache)

	  ; z.B. zur Verwendung in RegEx-Strings

	  ; Zahlen werden in Gruppen zu je 3 Zahlen gesprochen (Tausendertrennung)
	  ; 105                           	- Einhundertfünf
	  ; 1.153                        	- Eintausend.einhundertdreiundfünfzig
	  ; 13.678                      	- Dreizehntausend.sechshundertachtundsiebzig
	  ; 416.738                    	- Vierhundertsechzehntausend.siebenhundertachtunddreißig
	  ; 2.416.738                 	- Zweimillionen.vierhundertsechzehntausend.siebenhundertachtunddreißig
	  ; 12.416.738               	- Zwölfmillionen.vierhundertsechzehntausend.siebenhundertachtunddreißig
	  ; 612.416.738               	- Sechshundertzwölfmillionen.vierhundertsechzehntausend.siebenhundertachtunddreißig
	  ; 3.612.416.768            	- Dreimilliarden.sechshundertzwölfmillionen.vierhundertsechzehntausend.siebenhundertachtundsechzig
	  ; 23.003.612.416.768   	- Dreiundzwanzigbillionen.dreimilliarden.sechshundertzwölfmillionen.vierhundertsechzehntausend.siebenhundertachtundsechzig

	  ; zu Fragen bzgl. der Rechtschreibung siehe z.B. hier:
	  ; https://www.sekretaria.de/bueroorganisation/rechtschreibung/zahlwoerter/

		static numbers 	:= { twelves 	: ["ein", "zwei", "drei", "vier", "fünf", "sech", "sieb", "acht", "neun", "zehn", "elf", "zwölf"]
									, 	tens    	: ["zehn", "zwanzig", "dreißig", "vierzig", "fünfzig", "sechzig", "siebzig", "achtzig", "neunzig"]
									,	groups 	: ["tausend", "million", "milliarde", "billion", "billiarde", "trillion", "trilliarde"]}

	  ; die Null
		If (number=0)
			return "Null"

	  ; Vorarbeit - Tausendergruppen separieren
		GroupsOfThousands := Array()
		zerofillednumber      	:= Format("{:024}", number)
		zfnLen                    	:= StrLen(zerofillednumber)
		IgnoreZeros            	:= true
		Loop % (zfnLen/3) {
			ngroup := SubStr(zerofillednumber, (A_Index-1)*3+1, 3)
			If (IgnoreZeros && ngroup>0) || !IgnoreZeros {
				IgnoreZeros := false
				GroupsOfThousands.InsertAt(1, ngroup)
			}
		}

	  ; gruppenweise in Zahlworte umwandeln
		For index, thousands in GroupsOfThousands {

			ngroup := SubStr("000" thousands, -2)
			d1 := SubStr(ngroup, -0)
			d2 := SubStr(ngroup, -1)
			d3 := SubStr(ngroup, 1, 1)

		  ; die Bezeichnung der Zahlengruppe
			spokennumber	:=	(index=1                         	?	""
									:   	 numbers.groups[index-1]
									.		(index>2 && ngroup>1	? (Mod(index, 2)=0 ? "n" : "en") : "")
									.    	(index>1 ? separator : ""))
									.     	 spokennumber

		  ; Zahlen von 1 bis 99
			spokennumber 	:=	(d2=0                           	?	""
									:     	 d2=1                           	?	numbers.twelves[d2] (index=1 ? "s" : index>2 ? "e" : "")
									:     	 d2=2                           	?	numbers.twelves[d2]
									:     	 d2=6                           	?	numbers.twelves[d2] "s"
									:     	 d2=7                           	?	numbers.twelves[d2] "en"
									:     	 d2>=3  	&& d2<=9      	?	numbers.twelves[d2]
									:     	 d2>=10 && d2<=12 	?	numbers.twelves[d2]
									:    	 d2>=13 && d2<=19 	?	numbers.twelves[d1] numbers.tens[1]
									:    	 d1=0                          	?	numbers.tens[SubStr(d2, 1,1)]
									:     	 numbers.twelves[d1] . (d1=6 ? "s" : d1=7 ? "en" :"")
									.        "und" . numbers.tens[SubStr(d2, 1,1)])
									. 		 spokennumber

		  ; die Hunderterstelle
			spokennumber	:=  	(d3=0                             	?	""
									:     	 d3=6                             	?	numbers.twelves[d3] "s"
									:     	 d3=7                             	?	numbers.twelves[d3] "en"
									:    	 numbers.twelves[d3]) . (d3>0 ? "hundert" : "")
									. 		 spokennumber

		}

	  ; erstes Zeichen in Großschreibung
		If uppercase
			spokennumber := Format("{:U}", SubStr(spokennumber, 1, 1)) . SubStr(spokennumber, -1*(StrLen(spokennumber)-2))

	return spokennumber
	}

	RemovePunctuation(str)                                                              	{ ;-- Interpunktion entfernen
	return RegExReplace(str, "[;:_,'\-\.\+\*\(\)\{\}\[\]\/\\]")
	}

	RegExTransformGerman(str)                                                         	{ ;-- für Stringmatching Algorithmen
		str   	:= StrReplace(str, "ä"	, "(ä|ae)")
		str   	:= StrReplace(str, "ö"	, "(ö|oe)")
		str   	:= StrReplace(str, "ü"	, "(ü|ue)")
		str   	:= StrReplace(str, "ß"	, "(ß|ss)")
		str   	:= StrReplace(str, "Ä"	, "(Ä|Ae)")
		str   	:= StrReplace(str, "Ö"	, "(Ö|Oe)")
		str   	:= StrReplace(str, "Ü"	, "(Ü|Ue)")
	return str
	}

	optParser(tag, default:="", isbool:=false)                                     	{ ;-- Stringoptionen parsen

		static rxbool := "true|1|false|0"

		If !RegExMatch(this.options, "i)" tag "\s*=\s*(?<option>" (isbool ? rxbool : ".+?") ")(\s|$)", p_)
			return default

	return isbool && p_option ~="(true|1)" ? true : isbool && p_option ~="(false|0)" ? false :  p_option
	}

	URLDownloadToVar(url)                                                             	{ ;-- lädt Textdateien aus dem Internet
		hObject:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
		hObject.Open("GET", url)
		hObject.Send()
	return hObject.ResponseText
	}

	UCLCRatio(tx)                                                                             	{ ;-- Großbuchstaben/Kleinbuchstabenverhältnis

	  ; Tesseract scheint mehr Großbuchstaben zu erkennen umso schlechter die Qualität der zu scannenden Datei ist
	  ; eine Ratio zwischen 0.09 und 0.25 ist gut zu lesen, ab 0.35 wird es sehr schwierig, ab 1.00 kann man von einer verdrehten Vorlage ausgehen (Text stand auf dem Kopf)

	    txtA   	:= RegExReplace(txt, "[^\pL]")
		txtLC 	:= RegExReplace(txtA, "[A-ZÄÖÜ]")
		LC    	:= StrLen(txtLC)
		UC   	:= StrLen(txtA)-StrLen(txtLC)
		UCLC 	:= Round(UC/LC, 2)

	return UCLC
	}

	ClassifierFileTime()                                                                    	{ ;-- gibt nur die Zeit der letzten Aktualisierung der Klassifizierungsdaten zurück
		FileGetTime, ftime, % this.classifierfilepath, M
	return ftime
	}
}





