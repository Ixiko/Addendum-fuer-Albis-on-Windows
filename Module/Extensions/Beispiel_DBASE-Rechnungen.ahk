;	Beispielskript für die Verwendung von Addendum_DBASE.ahk
;

	#NoEnv
	#MaxMem 4096
	SetBatchLines, -1

	; Skriptstartzeit
		starttime	:= A_TickCount

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Objekte / Variablen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)        ; Addendum-Verzeichnis
		RegRead, PathAlbis, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
		basedir    	:= PathAlbis "\db"

		Search  	:= Array()
		Search[1] 	:= {	"dbasefile"		: "Rechnung"
							,	"output"     	: "Gesamtsumme"
							,	"showPatID"  	: true
							,	"text"          	: "Summe ausstehend:     `t"
							,	"fields"      	: {"ANDATUM"   	: "rx:(2020|2021)\d\d\d\d"}}

		nr         	:= 1
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
		rechnung   	:= new DBASE(basedir "\" search[nr].DBASEfile ".dbf", true)

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Lesezugriff einrichten. Rückgabewert ist die Position des Dateizeigers
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res        	:= rechnung.OpenDBF()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Suche starten. objekt.Search(StringObjekt [, exportPfad])
	; StringObjekt	- Objekt mit den Feldnamen als key und einem RegExString als value
	; exportPfad  	- optional ein Pfad - dann wird aber kein Array mit den extrahierten Zeilen zurückgegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		matches 	:= rechnung.Search(Search[nr].fields, 0, 1)

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbankzugriff kann geschlossen werden
	;-------------------------------------------------------------------------------------------------------------------------------------------
		res         	:= rechnung.CloseDBF()

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; 1.Zeit - Lesedauer aus der Datenbankdatei
	;-------------------------------------------------------------------------------------------------------------------------------------------
		summe := 0
		p := "PATNR  `tBetrag in EUR`n"
		For idx, data in matches {
			summe += data.BETRAGEUR
			p .= (idx = 1 ? "" : "`n") SubStr("     "  data.PATNR, -4) "  `t" data.BETRAGEUR
		}

		duration2	:= A_TickCount - starttime
	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Ausgabe in Scite
	;-------------------------------------------------------------------------------------------------------------------------------------------
		t := " Rechnungen gesamt:    `t"  	rechnung.records                                      	"`n"
		t .= " Rechnungen mit Datum:`t" 	matches.Count()                                        	"`n"
		t .= " Gesamtsumme:           `t" 	Round(summe,2) " EUR"                             	"`n"
		t .= " Duration all:               `t"  	Round(duration2 / 1000, 2) 	" seconds"       	"`n"
		t .= " records/second:          `t"  	Round(wzimmer.records/ (duration2/1000))  	"`n"
		t .= " ---------------------------------------------------------------------------------"      	"`n"
		t .= " Suchstring (RegEx):`t"  	rechnung.SearchRegExStr                                   	"`n"
		t .= " Suchstring (Länge):`t"  	rechnung.lastrxlen                                             	"`n"
		t .= " ---------------------------------------------------------------------------------"
		SciTEOutput(p "`n" RTrim(t, "`n"))

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; Datenbank Objekt freigegeben. Damit werden alle gespeicherten Daten gelöscht und Speicher freigegeben
	;-------------------------------------------------------------------------------------------------------------------------------------------
		rechnung := ""

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; extrahierte Daten ansehen
	;-------------------------------------------------------------------------------------------------------------------------------------------
		savefile.Write(p "`n`n" RTrim(t, "`n"))
		savefile.Close()

		Run, % filename

ExitApp



#Include %A_ScriptDir%\..\..\Include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
