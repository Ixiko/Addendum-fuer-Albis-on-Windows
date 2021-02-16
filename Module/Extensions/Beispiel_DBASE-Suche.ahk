
;	Beispielskript für die Verwendung von Addendum_DBASE.ahk
;

	#NoEnv
	SetBatchLines, -1

	; Skriptstartzeit
		starttime	:= A_TickCount

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Objekte / Variablen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		Search  	:= Array()
		Search[1] 	:= {	"dbasefile"		: "BEFUND"
							,	"output"     	: "Patienten mit Typ I Diabetes"
							,	"showPatID"  	: true
							,	"text"          	: "Diabetes Typ I:     `t"
							,	"fields"      	: {"DATUM"   	: "rx:20[12]\d\d{4}"
													,	"KUERZEL" 	: "dia"
													, 	"INHALT"   	: "rx:\{E10\."
													,	"QUARTAL"	: ""}}
		Search[2] 	:= {	"dbasefile"		: "BEFUND"
							,	"output"     	: "Patienten Carcinomdiagnosen"
							,	"showPatID"  	: false
							,	"text"          	: "Carcinompat:     `t"
							,	"fields"      	: {"DATUM"   	: "rx:20[12]\d\d{4}"
													,	"KUERZEL" 	: "dia"
													, 	"INHALT"   	: "rx:\{C\d+\."
													,	"QUARTAL"	: ""}}
		Search[3] 	:= {	"dbasefile"		: "BEFUND"
							,	"output"     	: "Patienten m. Hypertonie"
							,	"showPatID"  	: false
							,	"text"          	: "Hypertoniepat:     `t"
							,	"fields"      	: {"DATUM"   	: "rx:20[12]\d\d{4}"
													,	"KUERZEL" 	: "dia"
													, 	"INHALT"   	: "rx:\{I1\d\."
													,	"QUARTAL"	: ""}}
		Search[4] 	:= {	"dbasefile"		: "BEFUND"
							,	"output"     	: "Patienten m. Hypertonie und Diabetes"
							,	"showPatID"  	: false
							,	"text"          	: "Hypertonie + Diab:`t"
							,	"fields"      	: {"DATUM"   	: "rx:20[12]\d\d{4}"
													,	"KUERZEL" 	: "dia"
													, 	"INHALT"   	: "rx+:\{I1\d\.|\{E1\d\."
													,	"QUARTAL"	: ""}}

		nr         	:= 4
		With      	:= Object()
		cSearch		:= Object()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabedatei anlegen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		filename 	:= A_Temp "\" search[nr].output ".txt"
		savefile 	:= FileOpen(filename, "w", "UTF-8")

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; neues Datenbankobjekt anlegen. new DBASE(Datenbankpfad [, debug])
	; Datenbankpfad	- vollständiger Pfad zur Datenbank in der Form M:\albiswin\db\BEFUND.dbf
	; debug                	- Ausgabe von Werten zur Kontrolle des Ablaufes, Voreinstellung: keine Ausgabe
	;-------------------------------------------------------------------------------------------------------------------------------------------
		befund   	:= new DBASE("C:\tmp\" search[nr].DBASEfile ".dbf", true)
		befund := ""

		ExitApp
	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Lesezugriff einrichten. Rückgabewert ist die Position des Dateizeigers
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res        	:= befund.Open()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Suche starten. objekt.Search(StringObjekt [, exportPfad])
	; StringObjekt	- Objekt mit den Feldnamen als key und einem RegExString als value
	; exportPfad  	- optional ein Pfad - dann wird aber kein Array mit den extrahierten Zeilen zurückgegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		matches 	:= befund.Search(Search[nr].fields)

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbankzugriff kann geschlossen werden
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res         	:= befund.Close()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 1.Zeit - Lesedauer aus der Datenbankdatei
	;-------------------------------------------------------------------------------------------------------------------------------------------
		duration1	:= A_TickCount - starttime

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; zählt die Anzahl der Patienten im matches Array
	;-------------------------------------------------------------------------------------------------------------------------------------------
		For field, val in Search[nr].fields
			If RegExMatch(val, "^\s*rx(?<mode>.)\:", s) {
				cSearch[field] := {"mode":smode, "strings":StrSplit(RegExReplace(val, "^\s*rx.\:"), "|")}
				SciTEOutput( " field [" field "]: " cSearch[field].Mode ", "  cSearch[field].Count())
			}

		For index, m in matches {

			If !With.haskey(PATNR := Trim(m.PATNR)) {
				With[PATNR] := Object()
				With[PATNR].matchcount := 0
			}

		; im Falle mehrerer Konditionen - im Moment nur !und-Verknüpfungen!
			If (cSearch.Count() > 0) {
				For mfield, obj in cSearch
					For midx, mxStr in cSearch[mfield].strings
						If RegExMatch(m[mfield], "i)" mxStr) {

							If !IsObject(With[PATNR][mfield])
								With[PATNR][mfield] := Object()

							If !With[PATNR][mfield].haskey(midx)
								With[PATNR][mfield][midx] := 1
							else
								With[PATNR][mfield][midx] += 1

							If (With[PATNR][mfield].Count() = (MaxHits := cSearch[mfield].strings.Count()) ) {

								hits := 0
								If (cSearch[mfield].mode ="+") {

									Loop % (With[PATNR][mfield].Count() - 1)
										If (With[PATNR][mfield][1] = With[PATNR][mfield][A_Index + 1])
											hits ++
									If (hits+1 = MaxHits)
										With[PATNR].matchcount += 1

								}

							}

						}

			}
			else
				With[PATNR].matchcount += 1

			FormatTime, datum	, % m.Datum, dd.MM.yyyy
			FormatTime, monat	, % m.Datum, MM
			FormatTime, jahr   	, % m.Datum, yy
			Quartal	:= SubStr("00" Floor((monat - 1)/3) + 1, -1) jahr
			recNr	:= SubStr("0000000" m.RecordNr, -6)
			PatID	:= SubStr("000000" PATNR, -4)

			savefile.WriteLine(	recNr            	": ["
									.	Quartal         	"] "
									.	datum             	"  `t"
									.	PatID             	"  `t"
									.	m.INHALT     	"  `t")

		}
		savefile.Close()

		hitCount := 0
		For PATNR, matchcount in With
			If (With[PATNR].matchcount > 0)
				hitCount ++

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 2.Zeit - Dauer des Zählens
	;-------------------------------------------------------------------------------------------------------------------------------------------
		duration2	:= A_TickCount - starttime
		objSize1 	:= befund.GetCapacity()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabe in Scite
	;-------------------------------------------------------------------------------------------------------------------------------------------
		t := " records count:     `t"   	befund.record                                                 	"`n"
		t .= " entrys found:        `t" 	matches.Count()                                            	"`n"
		t .= " Duration search:    `t" 	Round(duration1 / 1000, 2) 	" seconds"           	"`n"
		t .= " Duration all:         `t" 	Round(duration2 / 1000, 2) 	" seconds"            	"`n"
		t .= " records/second:    `t" 	Round(recCount / (duration1/1000))               	"`n"
		t .= " ---------------------------------------------------------------------------------"     	"`n"
		t .= " Suchstring (RegEx):`t"  	befund.SearchRegExStr                                     	"`n"
		t .= " Suchstring (Länge):`t"  	befund.lastrxlen                                                	"`n"
		t .= " " search[nr].Text         	hitCount                                                       	"`n"
		t .= " ---------------------------------------------------------------------------------"     	"`n"

		If search[nr].showPatID
			For PatID, index in With.matchcount
				t .= SubStr(" " PatID, -6) ": " index "x`n"

		SciTEOutput(RTrim(t, "`n"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbank Objekt freigegeben. Damit werden alle gespeicherten Daten gelöscht und Speicher freigegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		ObjRelease(befund)
		SciTEOutput(" Anzahl der Elemente vor Freigabe: " objSize1 ", danach: " (befund.GetCapacity() = "" ? 0 : befund.GetCapacity()))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; extrahierte Daten ansehen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		;Run, % filename

ExitApp



#Include %A_ScriptDir%\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
