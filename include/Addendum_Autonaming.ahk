; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                         	** Addendum_Autonaming **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Beschreibung:	    	Funktionen für Textklassifizierungen von Patientendokumenten
;
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;       Abhängigkeiten:
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_Calc started:    	14.06.2022
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
	__New(classifierfilepath:="", workpath:="", options:="debug=false removeStopwords=true" ) {
	  ; Parse Options
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		this.options 	:= options
		this.debug 	:= this.optParser("debug", false, true)                 ; option, defaultvalue, is bool?
		this.rmwords	:= this.optParser("removeStopwords", true, true)
		this.instant	:= this.optParser("save_immediately", true, true)

	  ; Stopwörter laden und vorbereiten
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		If !stopwords.Count() {                                                       	; wird nur einmal beim Aufruf der ersten Instanz ausgeführt

			If FileExist(Addendum.Dir "\include\Daten\" (stopwordsFilename := "german_stopwords_full.txt"))
				stopwordstxt := FileOpen(Addendum.Dir "\include\Daten\" stopwordsFilename, "r", "UTF-8").Read()
			else {

				TrayTip, Download..., Lade Stoppwortliste in Deutsch von Github...., 10
				stopwordstxt := this.URLDownloadToVar("https://raw.githubusercontent.com/solariz/german_stopwords/master/" stopwordsFilename)
				If this.debug
					SciTEOutput(Addendum.Dir "\include\Daten\" stopwordsFilename)
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
			If this.debug {
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
		this.autonames := {}
		If isObject(autonames := this.load(classifierfilepath)) {
			For maintitle, data in autonames
				this.autonames[maintitle] := data
			autonames := ""
		}

		If this.debug && isObject(this.autonames)
			SciTEOutput(cJSON.dump(this.autonames, 1))

	  ; Dokumente eines Verzeichnisses können eingelesen werden
		If RegExMatch(workpath, "^([A-Z]\:|[A-Za-z_\d\-]+\\)\\") {

			If this.debug
				SciTEOutput("lese Dokumentnamen ein")

			documents := this.readDirectory(workpath)
			If (documents.Count() > 0) {
				this.documents := []
				For each, document in documents
					this.documents.Push(document)
				documents := ""
			}

		}

		If (this.documents.Count() > 0)
			return this.documents

	}

	__Delete() {

		;~ res := this.save()
		;~ SciTEOutput("Klassifizierungen gespeichert: " res)

	return res
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Klassifizierdaten
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	load(filepath)                                                                              	{ ; lädt Klassifizierungen

	  ; prüft ob Daten vorhanden sind
		SplitPath, filepath,, fDir, fExt, fnNoExt
		classifierfilepath	:= !filepath ? Addendum.DBPath "\Dictionary\Befundbezeichner" : fDir "\" fnNoExt
		If !(txtfile := FileExist(classifierfilepath ".txt")) && !(jsonfile := FileExist(classifierfilepath ".json")) {
			PraxTT("es ist weder die Text noch die Json formatierte Datei vorhanden: " classifierfilepath, "5 1")
			return
		}

	  ; die .json Datei wird wenn bereits angelegt verwendet
		classifierfilepath .= FileExist(classifierfilepath ".json") ? ".json" : ".txt"
		SplitPath, classifierfilepath,, fDir, fExt, fnNoExt
		FileGetTime, filetime, % classifierfilepath
		this.classifierfile := {"name":fnNoExt, "path":fDir, "ext":fExt, "filetime":filetime}
		autonames    	:= {}

	  ; Klassifizierungsdaten laden
		fileitems 	:= FileOpen(this.classifierfile.path "\" this.classifierfile.name "." this.classifierfile.ext, "r", "UTF-8").Read()
		If (fExt = "json") {
			autonames := cJSON.Load(fileitems)
		}
		else {

		  ; konvertiert die alte Datendatei (Befundbezeichner.txt - je Zeile ein Dateiname) in ein direkt verwendbares Objekt
			Items := ""
			For i, item in StrSplit(fileItems, "`n", "`r")
				If (StrLen(item := Trim(item)) > 0)
					Items .= (i>1 ? "`n" : "") . item

			Items := RegExReplace(Items, "[\n]{2,}", "`n")
			Items := Trim(Items, "`n")
			Sort, Items, U                                                     ; doppelte Einträge löschen

			For i, item in StrSplit(Items, "`n", "`r")
				If (StrLen(item := Trim(item)) > 0) {

					NStrip := this.autoNStrip(item)

					If !IsObject(autonames[NStrip.maintitle]) {
						autonames[NStrip.maintitle] := Object()
						autonames[NStrip.maintitle].main := NStrip.main
					}

					If NStrip.sub.Count() {
						If !IsObject(autonames[NStrip.maintitle].sub)
							autonames[NStrip.maintitle].sub := Object()
						autonames[NStrip.maintitle].sub[NStrip.subtitle] := NStrip.sub
					}

				}

			If (autonames.Count() > 0)
				this.save(autonames)

		}

	return autonames
	}

	save(classification:="")                                                                	{ ; speichert Klassifizierungen

		If !isObject(this.classifierfile)
			return -1

		cJSON.EscapeUnicode := "UTF-8"
		If (StrLen(fileitems := isObject(classification) ? cJSON.Dump(classification, 1) : cJSON.Dump(this.autonames, 1)) > 0)
			FileOpen(this.classifierfile.path "\" this.classifierfile.name ".json", "w", "UTF-8").Write(fileitems)

		res := ErrorLevel ", " A_LastError ", " StrLen(fileitems)
		this.newdata := ""

		SciTEOutput("gesichert: " this.classifierfile.path "\" this.classifierfile.name ".json")

	return res
	}

	update()                                                                                     	{ ; Klassifizierungen auffrischen

	  ; Klassifizierungen werden nur bei geändertem Dateiinhalt (Zeitstempelvergleich der letzten Änderung) aufgefrischt
	  ; dadurch muss nicht das komplette Skript neugestartet werden, wenn der Dateiinhalt z.B. durch einen Texteditor verändert wurde
	  ; zurückgeben wird der Änderungsstatus, also entweder original - für unverändert oder update - für verändert

		If !IsObject(this.classifierfile) {
			throw A_ThisFunc ": please initialize the class first with ' new autonamer(classifierfilepath)' before calling the class method"
		}

		If !FileExist(classifierfilepath := this.classifierfile.path "\" this.classifierfile.name "." this.classifierfile.ext) {
			SciTEOutput("Datei nicht vorhanden: " classifierfilepath)
			return "original"
		}

		FileGetTime, filetime, % classifierfilepath
		If (filetime > this.classifierfile.filetime) {

			If isObject(autonames := this.load(classifierfilepath)) {

				this.classifierfile.filetime := filetime

				SciTEOutput("start autonames count: " this.autonames.Count())
				For maintitle, data in this.autonames
					this.autonames.Delete(maintitle)
				SciTEOutput("end autonames count: " this.autonames.Count())

				For maintitle, data in autonames
					this.autonames[maintitle] := data
				autonames := ""

				If this.autonames.Count()
					return "update"

			}

		}

	return "original"
	}

	add(mainclass, subclass:="")                                                      	{ ; fügt Daten hinzu

		this.newdata := newdata := false

	  ; Kategorie
		If !this.hasMain(mainclass) {
			main             	:= this.removeWords(mainclass, false)
			mainwords       	:= this.getMatchItems(main)
			this.newdata 	:= newdata := true
			this.autonames[mainclass] := {"main": mainwords}
		}

	  ; Inhalt
		If subclass && !this.hasSub(mainclass, subclass) {

			If !IsObject(this.autonames[mainclass].sub)
				this.autonames[mainclass].sub := Object()

			sub                	:= this.removeWords(subclass, true)
			subwords       	:= this.getMatchItems(sub)
			this.newdata 	:= newdata := true
			this.autonames[mainclass].sub[subclass] := subwords

		}

	  ; Daten speichern
		If this.instant && this.newdata
			this.save()

	return newdata
	}

	getMains(matchstring:="")                                                          	{ ; alle Kategorien zurückgeben

		If (!isObject(this.autonames) || this.autonames.Count()=0)
			return

		maintitles := []

	 ; gibt nur Kategorien zurück welche den matchstring enthalten
		If matchstring {
			For mt, sub in this.autonames
				If Instr(mt, matchstring)
					maintitles.Push(mt)
		}

	 ; alle Kategorien zurückgeben, wenn matchstring war leer oder keine Übereinstimmungen gefunden  wurden
		If (maintitles.Count() = 0) {
			For mt, sub in this.autonames
				maintitles.Push(mt)
		}

	return maintitles
	}

	getMainwords(mainclass)                                                          		{ ; gibt die klassifizierenden Worte zurück
		return this.autonames[mainclass].main
	}

	getSubs(mainclass, subclass:="")                                                 	{ ; alle Subklassifizierungsdaten für eine Kategorie zurückgeben

	  ; subclass leer lassen um alle Subklassifizierungsobjekte zu erhalten
	  ; ein spezifiziertes subclass als String gibt nur die zugehörigen  Subklassifizierungsworte zurück

		If (!isObject(this.autonames) || this.autonames.Count()=0)
			return -1
		If !this.hasMain(mainclass)
			return -2

	return subclass ? this.autonames[mainclass].sub[subclass] : this.autonames[mainclass].sub
	}

	setSub(mainclass, subclass:="")                                                  	{ ; Subklassifizierungsdaten ändern)

	}

	hasMain(mainclass:="")                                                              	{ ; ist die Kategorie bekannt

		If !mainclass
			return false

	return isObject(this.autonames[mainclass])
	}

	hasSub(mainclass:="", subclass:="")                                           	{ ; Kategorie enthält Subklassifizierung

		If !this.hasMain(mainclass)
			return false

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
	readDirectory(workpath)                                                               	{ ; liest ein Verzeichnis ein oder nutzt die von AddendumGui erstellten PDF-Daten (.json Format)

	  ; PDF Dateinamen (Befundordner) einlesen
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If RegExMatch(workpath, "\.json$")  						{   ; schnelles Einlesen durch Addendum PDFDaten.json

			pdfdatapath := RegExMatch(workpath, "^[A-Z]\:\\") ? workpath : Addendum.DBPath "\sonstiges\" workpath
			If !FileExist(pdfdatapath)                                         	{
				throw A_ThisFunc ": Die Dokumentdatendatei ist nicht vorhanden. `n<<" pdfdatapath ">>"
				return -1
			}
			If !Addendum.BefundOrdner  									{
				throw A_ThisFunc ": Das Addendum Dokumentverzeichnis ist nicht angelegt oder wurde nicht angegeben.`nDie Funktion bricht hier ab."
				return -1
			}
			else If !Instr(FileExist(Addendum.BefundOrdner), "D") 	{
				throw A_ThisFunc ": Das Verzeichnis <<" Addendum.Befundordner ">> ist nicht vorhanden"
				return -1
			}

			this.workpath := Addendum.BefundOrdner
			If FileExist(pdfdatapath) {
				files := FileOpen(pdfdatapath, "r", "UTF-8").Read()
				files := cJSON.Load(files)
				documents := []
				For fNr, file in files
					If FileExist(this.workpath "\" file.name)
						documents.Push(file)
			}

		}
		else If workpath && InStr(FileExist(workpath), "D") 	{

			this.workpath := workpath
			documents := []

			If !Instr(FileExist(Addendum.Dir "\include\OCR\qpdf"), "D") {
				throw A_ThisFunc ": kann das qpdf Verzeichnis nicht finden.`n<<" Addendum.Dir "\include\OCR\qpdf"  ">>"
				return -1
			}

			Loop, Files, % this.workpath "\*.pdf"
			{

				FileGetTime	, timestamp	, % A_LoopFileFullPath
				FileGetSize	, filesize       	, % A_LoopFileFullPath
				documents.Push({	"name"          	: A_LoopFileName
										, 	"isSearchable"	: PDFisSearchable(A_LoopFileFullPath)
										, 	"pages"				: PDFGetPages(A_LoopFileFullPath, Addendum.Dir "\include\OCR\qpdf")
										, 	"timestamp"   	: timestamp
										,	"filesize"           	: filesize})
			}

		}
		else {

			throw A_ThisFunc ": notwendige Daten können nicht gefunden werden.`n<<" workpath  ">>"
			return -1

		}

	return documents
	}

	getDocumentText(fpath)                                                             	{ ; PDF/Word Dokumente in Text umwandeln

		If !RegExMatch(fpath, "[A-Z]\:\\") && this.workpath
			txt := IFilter(pdfpath := this.workpath "\" fpath)
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

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Haupt-Klassifizierer
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		For maintitle, data in this.autonames {

			 mt := v := ""
			 lastmpos := mainScore := mainScore1 := mainScore2 := mainScore3 := matchedIndex := caseMustMatch := caseMatched := 0

		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		  ; Hauptklassifikation vergleichen
		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			For mIndex, mainword in data.main {

			  ; Vorbereitungen
				mpos1 := mpos2 := ""
				mt .= (mIndex>1 ? ".*?") mainword
				caseMustMatch += RegExMatch(mainword, "^[A-ZÄÖÜ]") ? 1 : 0

			  ; Position innerhalb des Dokumentes
				rxmainword := RegExReplace(mainword, "\s{1,}", "\s+")
				If !(mpos1	:= RegExMatch(txt, "i)\s" rxmainword "\s"))
					mpos2	:= RegExMatch(txt, "i)\s" rxmainword)
				mpos := mpos2 ? mpos2 : mpos1
				If	mpos {
					caseMatched   +=	RegExMatch(mainword, "^[A-ZÄÖÜ]") ? 1 : 0
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
				rxmatchTight := RegExReplace(rxmatch, "\s", "\s+")
				rxmatchSkip	:= RegExReplace(rxmatch, "\s", ".*?")

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

			maincaseScore 	:= caseMustMatch=caseMatched ? 1 : 0
			mainScore    	:= mainScore1 + mainScore2 + mainScore3 + maincaseScore

			;~ If v
				;~ SciTEOutput(v ": " mainScore ", " matchedIndex)

			If mainScore
				mains[maintitle] := {"mainscore"	: mainScore
											, "main123C"	: mainScore1 "+" mainScore2 "+" mainScore3 "+" maincaseScore
											, "mainwords"	: RTrim(v, ",")
											, "subscore"   	: 0	}

			If (maxMainScore < mainScore) {
				maxMainScore	:= mainScore
				maxmaintitle 	:= maintitle
			}
			If !mainScore {
				bestscore.noscores += 1
				continue
			}

			z := SubStr("00" mainScore, -1)  "  " maintitle "`n"

		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		  ; Subklassifikationen durchgehen
		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			maxSubScore := 0
			maxSubTitle := maxSubWords := ""
			tp := ""
			for subTitle, subwords in data.sub {

				w := ""
				lastspos := subScore := 0
				spos := 1
				subwordC := subwords.Count()
				;~ tp .= "`n" subtitle "[" subwordC "]: "

				For windex, word in subwords {
					tp .= word
					If (npos := RegExMatch(txt, "\s" word)) {             ; ,, spos !!!! ####

						;~ tp .= " " npos

						If (lastspos < npos) {

							SubString := SubStr(txt, lastspos, npos-lastspos+1)
							RegExReplace(SubString, "\s+", "", wsCount)
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
						w .= word "[" wIndex "/" subwordC "," SubScore (wsCount ? wsCount "/" wIndexFactor : "" )  "],"

					}
					;~ tp .= ","
				}

			 ; Subscore ist höher
			 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				If (SubScore > mains[maintitle].subScore) {
					;~ maxMainScore	:= mainScore
					;~ maxmaintitle 	:= maintitle
					maxSubScore 	:= subScore
					maxSubTitle  	:= subTitle
					maxSubWords	:= RTrim(w, ", ")
					mains[maintitle].subScore	:= subScore
					mains[maintitle].subTitle 	:= subTitle
					mains[maintitle].subwords 	:= w ", subclassifier: " data.sub.Count()
					;~ If mainScore {
					;~ }
				}

			}

			;~ If !printthis
			;~ SciTEOutput(tp)
			;~ printthis := true

		}


		  ; bessten Score sichern
		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			For mTitle, mData in mains {

				mData.subScore := !mData.haskey("subScore") ? 0 : mData.subScore
				thisScore := mData.mainScore + mData.subScore

				If (thisScore > bestscore.main)  { 				;|| (thisScore <= bestscore.main && bestscore.sub = 0)
					bestscore.titles[thisScore] := {"mainScore"	: mData.mainScore
															, 	"subScore" 	: mData.subScore
															, 	"main123C"	: mData.main123C
															, 	"words" 		: mData.words
															, 	"mainTitle"  	: mTitle
															, 	"subTitle"    	: mData.subTitle}
					bestscore.maxMainTitle	:= mTitle
					bestscore.maxSubTitle	:= mData.subTitle
					If (bestscore.sub = 0) {
						bestWithoutSub := bestscore.titles.MaxIndex()
						tmp := bestscore.titles.Delete(bestWithoutSub)
						bestscore.titles[thisScore-1] := tmp
						bestscore.mainsub := thisScore
					}
					bestscore.max 	:= thisScore
					bestscore.main	:= mData.mainScore
					bestscore.sub 	:= mData.subScore
				}

			}

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

		static removers := [{rx: "[,;:!\/\(\)\?]"    	, rp: " "          	}	, {rx: "Dr.(\pL)"                 	, rp: " $1"       	}	, {rx: "\d+\-\d+"          	, rp: " "          	}
				        		, 	{rx: "\d\d\.\d\d\.\d*"	, rp: " "          	}	, {rx: "\d+\."                    	, rp: " "          	}	, {rx: "\s\pL(\s|$)"         	, rp: " "          	}							;, {rx: "(\pL)\s+(\d+)\s"	, rp: "$1\s+$2"	}
		        				, 	{rx:"(\pL\.)(\pL)"        	, rp: "$1.\s*$2"	}	, {rx:"\-(\d+)"                	, rp: "\-*\s*$1"  	}
								, 	{rx:"(\d+)\-"             	, rp: "$1\-*\s*"	}	, {rx:"([^\pL])\+([^\pL])"	, rp: "$1 $2"  	}]

		item   	:= " " item " "

		If this.rmwords
			For each, rxrpl in this.stopwords
				item := RegExReplace(item, "i)\s" rxrpl "\s", " ")

		For each, remove in removers
			item := RegExReplace(item, remove.rx, remove.rp)

		if rmdoubles
			item 	:= RegExReplace(item, "(?<!\pL|\-|\-\s)\pL{1,3}\.*(\s|$)", " ")               ; entfernt 1-3 Zeichenwörter


		item  	:= RegExReplace(item, "\s{2,}", " ")
		item  	:= Trim(item)

	return item
	}

	getMatchItems(Items)                                                                   	{

		static replacers := {"(Magnetresonanztomo[ag]ra[fph]+ie|MRT*)\s"                                    	: ""
								, 	"((craniale|kraniale)\s+Magnetresonanztomo[ag]ra[fph]+ie|cMRT*)\s" 	: ""
								, 	"(Computertomographie|CT)\s"                                                          	: ""
								, 	"(Halswirbels(ä|ae)ule|HWS|HWK)\s"                                                  	: ""
								,	"(Brustwirbels(ä|ae)ule|BWS|BWK)\s"                                                    	: ""
								, 	"(Lendenwirbels(ä|ae)ule|LWS|LWK)\s"                                              	: ""
								, 	"(Abdomen|Bauch|Oberbauch|Mittelbauch|Unterbauch)\s"                	: ""
								, 	"(Kopf|Schädel|Cranium|Neurocranium)\s"                                         	: ""
								, 	"(Hüftgelenke*\s|Hüfte\s|Hüft\-)"                                                         	: ""
								, 	"(OP\s|Operation*\s|OP\-)"                                                                  	: ""
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
		Items := RegExReplace(Items, "\s", " ")

	  ; Zahlen als Wort darstellen
		If RegExMatch(Items, "O)\s(?<number>\d+)\s+(?<word>[\pL\-\+°§$']+)", item) {
			; 105 - einhundertundfünf, 1153 - eintausendeinhundertdreiundfünfzig, 13678 - dreizehntausendsechshundertachtundsiebzig
			d1 := SubStr(item.number, -1)
			Firstnumbs := d1 = 0 ? "" : d1 <= 12 ? numbers.twelves[d1] :
		}

		For each, item in StrSplit(Items, A_Space) {
			item := " " item " "
			For rxstring, replacestring in replacers
				item := RegExReplace(item, "\s" rxstring , (replacestring ? replacestring : rxstring) )
			item := RegExReplace(item, "\\s$", "")
			If Trim(item)
				retItems.Push(Trim(item))
		}

	return retItems
	}

	cleanup(txt)                                                                               	{ ;-- Dokumententext vorbereiten

		static replacers := {"((c|k)raniale\s+Magnetresonanztomo[ag]ra[fph]+ie|cMRT*)\s" 	: "cMRT"
								,	"(Magnetresonanztomo[ag]ra[fph]+ie|MRT)\s"                                    	: "MRT"
								, 	"(Computertomogra(ph|f)ie|CT)\s"                                                        	: "CT"
								, 	"(Halswirbels(ä|ae)ule|HWS|HWK)\s"                                                  	: "HWS"
								,	"(Brustwirbels(ä|ae)ule|BWS|BWK)\s"                                                    	: "BWS"
								, 	"(Lendenwirbels(ä|ae)ule|LWS|LWK)\s"                                              	: "LWS"
								, 	"(Abdomen|Bauch|Oberbauch|Mittelbauch|Unterbauch)\s"                	: "Abdomen"
								, 	"(Kopf|Sch(ä|ae)del|Kalotten*|Cranium|Neurocranium)\s"                   	: "Schädel"
								, 	"(Hüftgelenke*\s|Hüfte\s|Hüft\-)"                                                         	: "Hüfte"
								, 	"(OP\s|Operation*\s|OP\-)"                                                                  	: "OP"
								, 	"(TEP|Totalendoprothese|Endoprothese)\s"                                          	: "TEP"
								, 	"(Schultergelenk*|Schulter)\s"                                                              	: "Schulter"
								, 	"(Thorax|Brustkorb)\s"                                                                        	: "Thorax"
								, 	"(Dialysezentrum|Dialysepartner|Nephrologe|Nephrologische*r*s*)\s"   	: "Nephrologie"
								, 	"(Kardiologische Praxis|Kardiolog(en|in|isches|e))"                              	: "Kardiologie"
								, 	"(nuklearmedizin(ische\sPraxis)*|Nuklearmedizin\.)"                            	: "Nuklearmedizin"
								, 	"(Radiologie|Radiologische Praxis|Radiologen|MVZ\s*Diagnostikum|Radiologisches)\s"                                   	: "Radiologie"
								, 	"(Daumen|Finger)\s"                                                                          	: "Finger"
								, 	"(Krankenkasse|Krankenversicherung|Gesundheitskasse|Kasse)"          	: "Krankenkasse"
								, 	"(DRV|Deutsche\sRentenversicherung)\s"                                              	: "DRV"
								, 	"(OSG|USG|(oberes|unteres)\s+Sprunggelenk)\s"                                	: "OSG"
								, 	"(Lungenheilkunde|Lungen-\s+und|Bronchialheilkunde|Lungenarztpraxis|Pneumolog(ie|e|in|en))"                   	: "Pulmologie"
								, 	"(Ebenen|Ebn(\.|\s)|Eb(\.|\s))\s"                                                           	: "Ebenen"
								, 	"(re(\.|\s)|rechts)\s"                                                                              	: "rechts"
								, 	"(li(\.|\s)|links)\s"                                                                                 	: "links"
								, 	"Rö\s"                                                                                                	: "Röntgen"}

		txt := " " txt " "

		If this.rmwords
			For each, rxrplStr in this.stopwords
				txt := RegExReplace(txt, "i)\s" rxrplStr "(\s|$)", " ")

		For rxstring, replacestring in replacers
			txt := RegExReplace(txt, "i)\s" rxstring , replacestring " ")

     ; andere Ersetzungen/Fehlerkorrekturen
		txt := RegExReplace(txt, "([a-zß\d][\)\.]*)([A-ZÄÖÜ][a-zäöüß]+)", "$1 $2")          ; Leerzeichen zwischen einem Klein- und Großbuchstaben einfügen
		txt := RegExReplace(txt, "[^\pL\d\-\.\$%&°§]", " ")                                       	; alle Zeichen außer die Zeichen in der eckigen Klammer entfernen
		txt := RegExReplace(txt, "\-\s+", "-")                                                           		; an Bindestrichen angehängte Leerzeichen entfernen
		txt := RegExReplace(txt, "\s(Dr|jur|Str|Nr)[,;]", " $1.")                                 		; OCR Zeichenfehler korrigieren
		txt := RegExReplace(txt, "\s[a-zäöüß]\s", " ")                                                		; OCR Zeichenfehler korrigieren
		txt := RegExReplace(txt, "Unferschrift", "Unterschrift")
		txt := RegExReplace(txt, "i)(D|d)igitate", "$1igitale")
		txt := RegExReplace(txt, "Magnetresonanztomo[ag]ra[fph]+ie", "Magnetresonanztomographie")
		txt := RegExReplace(txt, "\s{2,}", " ")


		SciTEOutput(txt)

	return txt
	}

	spokennumber(number, uppercase:=true, separator:=" ")           	{ ;-- Darstellung ganzer Zahlen als Zahlwörter (deutsche Sprache)

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

	RegExTransformGerman(str)                                                          	{ ;-- für Stringmatching Algorithmen
		str   	:= StrReplace(str, "ä", "(ä|ae)")
		str   	:= StrReplace(str, "ö", "(ö|oe)")
		str   	:= StrReplace(str, "ü", "(ü|ue)")
		str   	:= StrReplace(str, "ß", "(ß|ss)")
	return str
	}

	optParser(tag, default:="", isbool:=false)                                     	{ ;-- Stringoptionen parsen

		static rxbool := "true|1|false|0"

		If !RegExMatch(this.options, "i)" tag "\s*=\s*(?<option>" (isbool ? rxbool:".+?") ")(\s|$)", p_)
			return default

	return isbool && p_option ~="(true|1)" ? true : isbool && p_option ~="(false|0)" ? false :  p_option
	}

	URLDownloadToVar(url)                                                             	{ ;-- lädt Textdateien aus dem Internet
		hObject:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
		hObject.Open("GET",url)
		hObject.Send()
	return hObject.ResponseText
	}

}





